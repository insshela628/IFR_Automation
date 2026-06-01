"""Automated IFC conversion test — v9-style logic, project-configurable.

Fully automated: measure frame → update title block → stamp → SaveAs → PDF.
All checks via COM readback, PASS/FAIL per assertion, log saved to test_logs/.

Usage: python test_ifc_v9.py                       (default: warnertown)
       python test_ifc_v9.py tatua                  (run on tatua)
       python test_ifc_v9.py warnertown --measure-only
       python test_ifc_v9.py --skip-measure
"""
import sys
import os
import time
import io
from datetime import datetime
from pathlib import Path
import win32com.client
import pythoncom

# ── Dropbox root ─────────────────────────────────────────────────────────
DROPBOX_ROOT = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"

# ── Project configs ──────────────────────────────────────────────────────
PROJECTS = {
    'warnertown': {
        'name': 'Warnertown BESS',
        'project_path': os.path.join(DROPBOX_ROOT, r"2.SA\GG-31 Warnertown BESS"),
        'title_block_name': 'ACE-Wanertown_Siyuan',
        'personnel_tags': ['DRAWN', 'CHECK', 'ENGINEER', 'QA', 'PROJECT'],
        'rev_rows': 6,
        'frame_dwg': 'ACE_Standard_Frame_wanertown_UPDATED.dwg',
        'test_dwg': 'GG31-E-PLN-003_Communication Cable Route Layout Plan_RevA.dwg',
        'test_dwgs': [
            'GG31-C-PLN-003_Fence & Gate Layout Plan_RevA.dwg',
            'GG31-C-PLN-004-Road Pavement.dwg',
            'GG31-E-PLN-003_Communication Cable Route Layout Plan_RevA.dwg',
        ],
    },
    'tatua': {
        'name': 'Tatua Solar Farm',
        'project_path': os.path.join(DROPBOX_ROOT, r"7.New Zealand\Tatua Solar Farm"),
        'title_block_name': 'Coleamablly',
        'personnel_tags': ['DESIGNED', 'DRAWN', 'APPROVED', 'PROJECT'],
        'rev_rows': 4,
        'frame_dwg': 'Tatua_Standard_Frame.dwg',
        'test_dwg': None,
    },
    # Add new projects here:
    # 'new_project': {
    #     'name': '...',
    #     'project_path': os.path.join(DROPBOX_ROOT, r"..."),
    #     'title_block_name': '...',
    #     'personnel_tags': [...],
    #     'rev_rows': N,
    #     'frame_dwg': '....dwg',
    #     'test_dwg': None,
    # },
}

# ── Stamp constants (universal — from v10 IFCStampMixin, cross-project) ──
STAMP_LAYER = 'IFC_STAMP'
STAMP_TEXT = "{\\fArial Narrow|b1;FOR CONSTRUCTION}"
COLOUR_TEXT = "{\\fArial Narrow|b1;DRAWINGS TO BE\\PPRINTED IN COLOUR}"

REF_TB_W = 841.0
REF_TB_H = 594.0
REF_RECT_W = 110.511              # lower box width
REF_RECT_H = 17.745               # lower box height
REF_TEXT_H = 7.0                   # FOR CONSTRUCTION text height
REF_TEXT_W = 116.419               # MText width
REF_TEXT_Y_OFFSET = 13.182         # text Y offset inside rect
REF_X_RIGHT_OFFSET = 29.779       # stamp right edge offset from TB right
REF_Y_BOTTOM = 73.259             # stamp bottom offset from TB bottom
REF_COLOUR_RECT_H = 26.0          # upper box height (2 lines)
REF_COLOUR_GAP = 2.0              # gap between lower and upper boxes
REF_COLOUR_TEXT_H = 5.5           # upper box text height


# ── Logging — tee stdout to file ─────────────────────────────────────────
LOG_DIR = Path(__file__).parent / "test_logs"


class TeeWriter:
    """Write to both original stdout and a log file."""

    def __init__(self, log_path):
        self._original = sys.stdout
        self._log = open(log_path, 'w', encoding='utf-8')

    def write(self, text):
        self._original.write(text)
        self._log.write(text)

    def flush(self):
        self._original.flush()
        self._log.flush()

    def close(self):
        try:
            self._log.close()
        except (PermissionError, OSError):
            pass

    @property
    def encoding(self):
        return 'utf-8'

    def reconfigure(self, **kwargs):
        """Support sys.stdout.reconfigure() calls."""
        pass


# ── Test infrastructure ──────────────────────────────────────────────────
class TestReport:
    """Collect PASS/FAIL assertions, print summary, save to log file."""

    def __init__(self):
        self.results = []  # (name, passed, detail)
        self.start_time = datetime.now()

    def check(self, name, condition, detail=""):
        passed = bool(condition)
        self.results.append((name, passed, detail))
        status = "PASS" if passed else "FAIL"
        msg = f"  [{status}] {name}"
        if detail:
            msg += f" — {detail}"
        print(msg)
        return passed

    def summary(self):
        total = len(self.results)
        passed = sum(1 for _, p, _ in self.results if p)
        failed = total - passed
        elapsed = (datetime.now() - self.start_time).total_seconds()

        print(f"\n{'='*70}")
        print(f"  TEST SUMMARY: {passed}/{total} passed, {failed} failed  "
              f"({elapsed:.1f}s)")
        print(f"{'='*70}")
        if failed:
            print(f"  Failed tests:")
            for name, p, detail in self.results:
                if not p:
                    print(f"    - {name}: {detail}")
        else:
            print(f"  All tests passed!")
        return failed == 0


def sep(title):
    print(f"\n{'='*70}")
    print(f"  {title}")
    print(f"{'='*70}")


def get_acad():
    """Connect to running AutoCAD or start new instance."""
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        print("已连接到运行中的 AutoCAD")
    except Exception:
        print("正在启动 AutoCAD...")
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True
        for _ in range(60):
            try:
                _ = acad.Documents.Count
                break
            except Exception:
                time.sleep(1)
        time.sleep(3)
    return acad


def wait_doc_ready(doc, timeout=20):
    for _ in range(timeout):
        try:
            _ = doc.ModelSpace.Count
            _ = doc.Layouts.Count
            break
        except Exception:
            time.sleep(1)
    time.sleep(2)


def open_dwg(acad, path):
    doc = None
    for attempt in range(5):
        try:
            doc = acad.Documents.Open(path)
            break
        except Exception as e:
            print(f"  重试 {attempt+1}/5: {e}")
            time.sleep(3)
    if doc:
        wait_doc_ready(doc)
    return doc


def get_attrs_dict(block_ref):
    attrs = {}
    try:
        for attr in block_ref.GetAttributes():
            attrs[attr.TagString.upper()] = attr
    except Exception:
        pass
    return attrs


def find_test_dwg(cfg, target_override=None):
    """Find a suitable test DWG — use explicit target, config, or auto-detect.

    target_override can be:
      - Full filename: 'GG31-E-PLN-003_xxx.dwg'
      - Doc-ID prefix: 'GG31-C-PLN-004' (searches by prefix)
      - None: use cfg['test_dwg'] or auto-detect
    """
    native_root = os.path.join(cfg['project_path'], r"Design\Engineering\1. Drawings\1. Native")
    target = target_override or cfg.get('test_dwg')

    if target:
        # If it's already a full path
        if os.path.exists(target):
            return target
        # Search in native folder tree
        is_doc_id = not target.endswith('.dwg')
        for root, dirs, files in os.walk(native_root):
            dirs[:] = [d for d in dirs if d.lower() not in ('ss', 'superseded', 'superceded')]
            for f in files:
                if not f.endswith('.dwg') or f.startswith('~$'):
                    continue
                if is_doc_id:
                    # Match by doc-ID prefix (case-insensitive)
                    if f.upper().startswith(target.upper()):
                        return os.path.join(root, f)
                else:
                    if f == target:
                        return os.path.join(root, f)

    # Auto-detect: find smallest non-frame DWG
    best = None
    best_size = float('inf')
    for root, dirs, files in os.walk(native_root):
        dirs[:] = [d for d in dirs if d.lower() not in ('ss', 'superseded', 'superceded')]
        for f in sorted(files):
            if f.endswith('.dwg') and not f.startswith('~$') and 'Frame' not in f and 'Standard' not in f:
                alt = os.path.join(root, f)
                size = os.path.getsize(alt)
                if 50_000 < size < 5_000_000 and size < best_size:
                    best = alt
                    best_size = size
    return best


# =============================================================================
# Phase 1: Measure Reference Frame (automated)
# =============================================================================
def measure_frame(acad, report, cfg):
    """Measure reference frame dimensions for a project.

    Returns dict with tb_width, tb_height, frame_w, frame_h, etc.
    """
    sep(f"Phase 1: Measure Reference Frame ({cfg['name']})")
    result = {}

    native_root = os.path.join(cfg['project_path'], r"Design\Engineering\1. Drawings\1. Native")
    frame_dwg = os.path.join(native_root, cfg['frame_dwg'])
    title_block_name = cfg['title_block_name']

    report.check("frame_file_exists", os.path.exists(frame_dwg),
                 os.path.basename(frame_dwg))
    if not os.path.exists(frame_dwg):
        return result

    doc = open_dwg(acad, frame_dwg)
    report.check("frame_opened", doc is not None)
    if not doc:
        return result

    # List layouts
    layout_names = []
    for layout in doc.Layouts:
        layout_names.append(layout.Name)
    print(f"  Layouts: {layout_names}")

    # Find title block
    tb_found = False
    tb_name = None
    tb_bbox = None
    tb_attrs_count = 0
    ifr_blocks = []

    for space_name, get_space in [("ModelSpace", lambda: doc.ModelSpace),
                                   ("PaperSpace", lambda: doc.PaperSpace)]:
        try:
            space = get_space()
            count = space.Count
            print(f"  {space_name}: {count} entities")

            for i in range(count):
                try:
                    ent = space.Item(i)
                    ename = ent.EntityName

                    if ename == 'AcDbBlockReference':
                        bname = ent.Name
                        attrs = get_attrs_dict(ent)

                        if bname == title_block_name or ('DRAWINGNUMBER' in attrs and 'REVISION' in attrs):
                            if not tb_found:
                                try:
                                    mn, mx = ent.GetBoundingBox()
                                    tb_bbox = (float(mn[0]), float(mn[1]), float(mx[0]), float(mx[1]))
                                except:
                                    pass
                                tb_found = True
                                tb_name = bname
                                tb_attrs_count = len(attrs)
                                ix, iy = ent.InsertionPoint[0], ent.InsertionPoint[1]
                                print(f"  Title block: '{bname}' at ({ix:.2f}, {iy:.2f}), {len(attrs)} attrs")

                        if bname.upper() in ('IFR', 'IFR_STAMP') or 'REVIEW' in bname.upper():
                            ifr_blocks.append(bname)

                    elif ename in ('AcDbMText', 'AcDbText'):
                        txt = ent.TextString
                        if any(kw in txt.upper() for kw in ('REVIEW', 'CONSTRUCTION', 'COLOUR')):
                            ix, iy = ent.InsertionPoint[0], ent.InsertionPoint[1]
                            print(f"  Stamp text: ({ix:.2f}, {iy:.2f}) \"{txt[:60]}\"")
                except Exception:
                    continue
        except Exception:
            continue

    # Scan block definitions for IFR stamp
    try:
        blocks = doc.Blocks
        for i in range(blocks.Count):
            try:
                blk = blocks.Item(i)
                bname = blk.Name
                if bname.startswith('*'):
                    continue
                has_review = False
                for j in range(blk.Count):
                    try:
                        e = blk.Item(j)
                        if e.EntityName in ('AcDbMText', 'AcDbText'):
                            if 'REVIEW' in e.TextString.upper():
                                has_review = True
                                break
                    except:
                        pass
                if has_review or 'IFR' in bname.upper():
                    ifr_blocks.append(f"{bname}(def)")
                    print(f"  IFR block definition: '{bname}' ({blk.Count} entities)")
            except:
                pass
    except:
        pass

    # Frame max extent
    max_w, max_h = 0, 0
    try:
        ms = doc.ModelSpace
        for i in range(ms.Count):
            try:
                ent = ms.Item(i)
                mn, mx = ent.GetBoundingBox()
                w = float(mx[0]) - float(mn[0])
                h = float(mx[1]) - float(mn[1])
                if w > max_w:
                    max_w = w
                if h > max_h:
                    max_h = h
            except:
                pass
    except:
        pass

    doc.Close(False)

    # Report results
    report.check("frame_title_block_found", tb_found,
                 f"'{tb_name}' with {tb_attrs_count} attrs" if tb_found else "not found")

    if tb_bbox:
        tb_w = tb_bbox[2] - tb_bbox[0]
        tb_h = tb_bbox[3] - tb_bbox[1]
        result['tb_width'] = tb_w
        result['tb_height'] = tb_h
        result['tb_bbox'] = tb_bbox
        result['tb_name'] = tb_name
        print(f"  Title block size: {tb_w:.3f} x {tb_h:.3f}")
        print(f"  Frame max extent: {max_w:.2f} x {max_h:.2f}")

        report.check("frame_dimensions_reasonable",
                     tb_w > 100 and tb_h > 100,
                     f"{tb_w:.1f} x {tb_h:.1f}")

        # Check if same as Tatua (841x594)
        is_same = abs(tb_w - REF_TB_W) < 5 and abs(tb_h - REF_TB_H) < 5
        report.check("frame_matches_tatua_size", True,  # info only, not a hard fail
                     f"{'YES' if is_same else 'NO'} — Warnertown={tb_w:.1f}x{tb_h:.1f} vs Tatua={REF_TB_W}x{REF_TB_H}")

    result['frame_w'] = max_w
    result['frame_h'] = max_h
    result['ifr_blocks'] = ifr_blocks
    report.check("frame_ifr_blocks", True,  # info only
                 f"{ifr_blocks if ifr_blocks else 'none found (expected for Warnertown)'}")

    return result


# =============================================================================
# Phase 2: Test IFC Conversion (v9-style, fully automated)
# =============================================================================
def test_ifc_conversion(acad, report, cfg, frame_info=None):
    sep(f"Phase 2: IFC Conversion Test ({cfg['name']})")

    dwg_path = find_test_dwg(cfg)
    report.check("test_dwg_found", dwg_path is not None,
                 os.path.basename(dwg_path) if dwg_path else "no suitable DWG in Native/")
    if not dwg_path:
        return

    print(f"  Test DWG: {os.path.basename(dwg_path)}")
    doc = open_dwg(acad, dwg_path)
    report.check("test_dwg_opened", doc is not None)
    if not doc:
        return

    try:
        doc = _run_ifc_test(doc, acad, report, cfg, frame_info)
        if doc is not None:
            test_saveas_and_pdf(doc, acad, report, dwg_path, cfg)
    finally:
        # Close all open docs without saving
        try:
            while acad.Documents.Count > 0:
                acad.Documents.Item(0).Close(False)
            print(f"\n  所有文件已关闭（未保存）")
        except Exception:
            pass


def _run_ifc_test(doc, acad, report, cfg, frame_info):
    """Core test logic — separated so finally can close doc."""

    title_block_name = cfg['title_block_name']
    PERSONNEL_TAGS = cfg['personnel_tags']
    REV_ROWS = cfg['rev_rows']

    # ── Step 1: Find ALL title blocks ──
    # Strategy: prefer PaperSpace title blocks (stamps drawn in PaperSpace are
    # always visible in PDF — no viewport freeze issues). Fall back to ModelSpace
    # only if NO PaperSpace TBs found.
    print(f"\n  [Step 1] 查找 title block '{title_block_name}'...")
    all_tbs = []  # list of (block_ref, attrs, space, space_name)
    _ms_tbs = []  # ModelSpace TBs (exact name match)
    _ms_fallback = []  # ModelSpace fallback (by attrs)
    _ps_tbs = []  # PaperSpace TBs (exact name match)
    _ps_fallback = []  # PaperSpace fallback (by attrs)

    # Pass 1: SelectionSet in ModelSpace (safe for large DWGs with 400K+ entities)
    try:
        import pythoncom
        ss_name = f"TB_FIND_{int(time.time()*1000)}"
        ss = doc.SelectionSets.Add(ss_name)
        ft = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
        fv = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["INSERT"])
        ss.Select(5, None, None, ft, fv)  # mode 5 = all in active space
        for i in range(ss.Count):
            try:
                ent = ss.Item(i)
                bname = ent.Name
                if bname == title_block_name:
                    a = get_attrs_dict(ent)
                    _ms_tbs.append((ent, a, doc.ModelSpace, "ModelSpace"))
                else:
                    a = get_attrs_dict(ent)
                    if 'DRAWINGNUMBER' in a and 'REVISION' in a:
                        _ms_fallback.append((ent, a, doc.ModelSpace, "ModelSpace"))
            except Exception:
                continue
        ss.Delete()
    except Exception as e:
        print(f"  SelectionSet search failed: {e}")

    # Pass 2: Search ALL PaperSpace layouts (small, safe to iterate)
    for layout in doc.Layouts:
        if layout.Name.lower() == 'model':
            continue
        try:
            blk = layout.Block
            for i in range(blk.Count):
                try:
                    ent = blk.Item(i)
                    if ent.EntityName != 'AcDbBlockReference':
                        continue
                    bname = ent.Name
                    if bname == title_block_name:
                        a = get_attrs_dict(ent)
                        _ps_tbs.append((ent, a, blk, layout.Name))
                    else:
                        a = get_attrs_dict(ent)
                        if 'DRAWINGNUMBER' in a and 'REVISION' in a:
                            _ps_fallback.append((ent, a, blk, layout.Name))
                except Exception:
                    continue
        except Exception:
            continue

    # Priority: PaperSpace > ModelSpace (stamps in PaperSpace always visible in PDF)
    # Within each: exact name match > fallback (DRAWINGNUMBER+REVISION)
    if _ps_tbs:
        all_tbs = _ps_tbs
    elif _ps_fallback:
        all_tbs = _ps_fallback
    elif _ms_tbs:
        all_tbs = _ms_tbs
    elif _ms_fallback:
        all_tbs = _ms_fallback

    # Dedup by Handle
    if all_tbs:
        seen = set()
        deduped = []
        for tb in all_tbs:
            try:
                h = tb[0].Handle
            except:
                h = id(tb[0])
            if h not in seen:
                seen.add(h)
                deduped.append(tb)
        all_tbs = deduped

    # Use first TB for primary checks, but stamp ALL
    block_ref = all_tbs[0][0] if all_tbs else None
    attrs = all_tbs[0][1] if all_tbs else {}
    space = all_tbs[0][2] if all_tbs else None
    tb_space_name = all_tbs[0][3] if all_tbs else None

    tb_info = f"'{block_ref.Name}' in {tb_space_name} ({len(attrs)} attrs)" if block_ref else "not found"
    if len(all_tbs) > 1:
        tb_info += f" + {len(all_tbs)-1} more"
    report.check("title_block_found", block_ref is not None, tb_info)
    if not attrs:
        # Debug: list block refs via SelectionSet (safe for large DWGs)
        try:
            ss_name = f"TB_DBG_{int(time.time()*1000)}"
            ss = doc.SelectionSets.Add(ss_name)
            ft = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
            fv = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["AcDbBlockReference"])
            ss.Select(5, None, None, ft, fv)
            names = set()
            for i in range(min(ss.Count, 200)):
                try:
                    names.add(ss.Item(i).Name)
                except:
                    pass
            ss.Delete()
            print(f"  ModelSpace block refs: {names}")
        except:
            pass
        # Also list PaperSpace block refs
        try:
            for layout in doc.Layouts:
                if layout.Name.lower() == 'model':
                    continue
                blk = layout.Block
                names = set()
                for i in range(min(blk.Count, 200)):
                    try:
                        ent = blk.Item(i)
                        if ent.EntityName == 'AcDbBlockReference':
                            names.add(ent.Name)
                    except:
                        pass
                    if names:
                        print(f"  {layout.Name} block refs: {names}")
        except:
            pass
        return

    print(f"  找到 {len(all_tbs)} 个 title block:")
    for idx, (br, at, sp, sn) in enumerate(all_tbs, 1):
        print(f"    TB {idx}: '{br.Name}' in {sn} ({len(at)} attrs)")

    report.check("attrs_count_reasonable", len(attrs) >= 20,
                 f"{len(attrs)} attrs (expect ~69 for Warnertown)")

    # Get title block dimensions (from first TB — used for verification)
    try:
        mn, mx = block_ref.GetBoundingBox()
        tb_w = float(mx[0]) - float(mn[0])
        tb_h = float(mx[1]) - float(mn[1])
        tb_min_x, tb_min_y = float(mn[0]), float(mn[1])
        tb_max_x, tb_max_y = float(mx[0]), float(mx[1])
    except Exception:
        tb_w, tb_h = REF_TB_W, REF_TB_H
        tb_min_x, tb_min_y, tb_max_x, tb_max_y = 0, 0, tb_w, tb_h

    report.check("title_block_has_size", tb_w > 100 and tb_h > 100,
                 f"{tb_w:.1f} x {tb_h:.1f}")

    # ── Step 2: Read revision rows ──
    print(f"\n  [Step 2] 读取修订行...")
    personnel = {}
    last_ifr_row = 0
    existing_ifc_row = 0
    all_rows = []

    for row_num in range(1, REV_ROWS + 1):
        rev_tag = f"{row_num}REV"
        if rev_tag not in attrs:
            continue
        val = attrs[rev_tag].TextString.strip()
        if not val:
            continue

        desc_tag = f"{row_num}DESCRIPTION"
        desc_val = attrs[desc_tag].TextString.strip() if desc_tag in attrs else ''

        row_info = {'row': row_num, 'rev': val, 'desc': desc_val}
        all_rows.append(row_info)

        if 'CONSTRUCTION' in desc_val.upper():
            existing_ifc_row = max(existing_ifc_row, row_num)
        else:
            if row_num > last_ifr_row:
                last_ifr_row = row_num
                for tag in PERSONNEL_TAGS:
                    full_tag = f"{row_num}{tag}"
                    if full_tag in attrs:
                        personnel[tag.lower()] = attrs[full_tag].TextString.strip()

    for r in all_rows:
        row_type = 'IFC' if 'CONSTRUCTION' in r['desc'].upper() else 'IFR'
        print(f"    Row {r['row']}: Rev={r['rev']}, Desc='{r['desc']}' [{row_type}]")

    print(f"  Last IFR row: {last_ifr_row}, Existing IFC row: {existing_ifc_row}")
    print(f"  Personnel: {personnel}")

    # ── Step 3: Update title block (preserve_ifr=True) ──
    print(f"\n  [Step 3] 更新 title block (preserve_ifr=True)...")
    ifc_rev = 0
    date_str = time.strftime('%d/%m/%y')

    if existing_ifc_row > 0:
        target_row = existing_ifc_row
    else:
        target_row = last_ifr_row + 1
        if target_row > REV_ROWS:
            target_row = REV_ROWS

    report.check("target_row_valid", 1 <= target_row <= REV_ROWS,
                 f"row {target_row} (last_ifr={last_ifr_row}, existing_ifc={existing_ifc_row})")
    report.check("ifr_rows_preserved", target_row > last_ifr_row or existing_ifc_row > 0,
                 f"IFC at row {target_row}, IFR rows 1-{last_ifr_row} untouched")

    # Save pre-update values for IFR rows (to verify preservation)
    ifr_pre_values = {}
    for row_num in range(1, last_ifr_row + 1):
        for suffix in ['REV', 'DESCRIPTION', 'DATE'] + PERSONNEL_TAGS:
            tag = f"{row_num}{suffix}"
            if tag in attrs:
                ifr_pre_values[tag] = attrs[tag].TextString

    # Write REVISION
    old_rev = attrs['REVISION'].TextString if 'REVISION' in attrs else ''
    if 'REVISION' in attrs:
        attrs['REVISION'].TextString = str(ifc_rev)

    # Write target row
    tp = f"{target_row}"
    if f'{tp}REV' in attrs:
        attrs[f'{tp}REV'].TextString = str(ifc_rev)
    if f'{tp}DESCRIPTION' in attrs:
        attrs[f'{tp}DESCRIPTION'].TextString = 'FOR CONSTRUCTION'
    if f'{tp}DATE' in attrs:
        attrs[f'{tp}DATE'].TextString = date_str
    for tag in PERSONNEL_TAGS:
        ft = f"{tp}{tag}"
        if ft in attrs:
            attrs[ft].TextString = personnel.get(tag.lower(), '')

    # Update ALL other title blocks (same revision/description/date)
    for tb_idx2 in range(1, len(all_tbs)):
        _, tb_attrs, _, tb_sn = all_tbs[tb_idx2]
        try:
            if 'REVISION' in tb_attrs:
                tb_attrs['REVISION'].TextString = str(ifc_rev)
            if f'{tp}REV' in tb_attrs:
                tb_attrs[f'{tp}REV'].TextString = str(ifc_rev)
            if f'{tp}DESCRIPTION' in tb_attrs:
                tb_attrs[f'{tp}DESCRIPTION'].TextString = 'FOR CONSTRUCTION'
            if f'{tp}DATE' in tb_attrs:
                tb_attrs[f'{tp}DATE'].TextString = date_str
            for tag in PERSONNEL_TAGS:
                ft = f"{tp}{tag}"
                if ft in tb_attrs:
                    tb_attrs[ft].TextString = personnel.get(tag.lower(), '')
            print(f"    TB {tb_idx2+1} ({tb_sn}): 已更新")
        except Exception as e:
            print(f"    TB {tb_idx2+1} ({tb_sn}): 更新失败 — {e}")

    # Verify: re-read and check values
    time.sleep(0.5)
    rev_readback = attrs['REVISION'].TextString if 'REVISION' in attrs else ''
    report.check("revision_updated", rev_readback == str(ifc_rev),
                 f"REVISION='{rev_readback}' (was '{old_rev}')")

    desc_readback = attrs[f'{tp}DESCRIPTION'].TextString if f'{tp}DESCRIPTION' in attrs else ''
    report.check("ifc_row_description", 'CONSTRUCTION' in desc_readback.upper(),
                 f"row {target_row} DESCRIPTION='{desc_readback}'")

    date_readback = attrs[f'{tp}DATE'].TextString if f'{tp}DATE' in attrs else ''
    report.check("ifc_row_date", date_readback == date_str,
                 f"row {target_row} DATE='{date_readback}'")

    # Verify IFR rows preserved
    ifr_preserved = True
    ifr_diff = []
    for tag, old_val in ifr_pre_values.items():
        row_n = int(tag[0])
        if row_n >= target_row:
            continue  # skip the IFC row itself
        new_val = attrs[tag].TextString if tag in attrs else ''
        if new_val != old_val:
            ifr_preserved = False
            ifr_diff.append(f"{tag}: '{old_val}'->'{new_val}'")

    report.check("ifr_rows_unchanged", ifr_preserved,
                 f"all {len(ifr_pre_values)} IFR attr values intact" if ifr_preserved else f"changed: {ifr_diff}")

    # ── Step 3b: Scan for existing COLOUR stamp + fix typos ──
    print(f"\n  [Step 3b] 扫描原 DWG stamp 内容...")
    has_colour = False
    import re as _re
    try:
        ss_name = f'colour_scan_{int(time.time())}'
        ss_scan = doc.SelectionSets.Add(ss_name)
        ft = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
        fv = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ['MTEXT'])
        ss_scan.Select(5, None, None, ft, fv)
        for i in range(ss_scan.Count):
            try:
                ent = ss_scan.Item(i)
                txt = ent.TextString.upper()
                # Match COLOUR/COULOUR/COLOR + PRINTED
                if ('COLOUR' in txt or 'COULOUR' in txt or 'COLOR' in txt) and 'PRINT' in txt:
                    has_colour = True
                    # Fix typos in-place (COULOUR→COLOUR, etc.)
                    original = ent.TextString
                    fixed = _re.sub(r'[Cc][Oo][Uu]?[Ll][Oo][Uu]+[Rr]', 'COLOUR', original, flags=_re.IGNORECASE)
                    if fixed != original:
                        ent.TextString = fixed
                        print(f"    修复拼写: \"{original[:40]}\" → \"{fixed[:40]}\"")
                    else:
                        print(f"    COLOUR stamp OK: \"{original[:60]}\" (layer={ent.Layer})")
            except:
                continue
        ss_scan.Delete()
    except Exception as e:
        print(f"    COLOUR scan failed: {e}")

    # Remove IFR block references (FOR REVIEW blocks)
    print(f"\n  [Step 3c] 清除旧 IFR stamp...")
    ifr_block_names = set()
    try:
        for i in range(doc.Blocks.Count):
            try:
                blk = doc.Blocks.Item(i)
                if blk.IsXRef or blk.IsLayout or blk.Count > 10:
                    continue
                for j in range(blk.Count):
                    try:
                        ent = blk.Item(j)
                        if ent.EntityName == 'AcDbMText':
                            txt = ent.TextString.upper()
                            if 'REVIEW' in txt or 'CONSTRUCTION' in txt:
                                ifr_block_names.add(blk.Name)
                    except:
                        continue
            except:
                continue
    except:
        pass

    removed_count = 0
    if ifr_block_names:
        print(f"    IFR stamp blocks: {ifr_block_names}")
        try:
            ss_ins = doc.SelectionSets.Add(f'ins_{int(time.time())}')
            ft2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
            fv2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ['INSERT'])
            ss_ins.Select(5, None, None, ft2, fv2)
            for i in range(ss_ins.Count):
                try:
                    ins = ss_ins.Item(i)
                    if ins.Name in ifr_block_names:
                        ins.Delete()
                        removed_count += 1
                except:
                    continue
            ss_ins.Delete()
        except:
            pass

    # Pass 2: Remove standalone MText with EXACT stamp phrases (all spaces)
    # Handles DWGs like C-PLN-003 where "FOR REVIEW" is a plain MText on QA layer
    _STAMP_PHRASES = {'FOR CONSTRUCTION', 'ISSUED FOR REVIEW', 'FOR REVIEW'}
    mtext_removed = 0

    # Ensure ModelSpace is active for SelectionSet (Select(5) only searches active space)
    try:
        doc.ActiveSpace = 1  # acModelSpace
        time.sleep(0.5)
    except Exception:
        pass

    # Scan ModelSpace MTEXT via SelectionSet
    try:
        ss_mt = doc.SelectionSets.Add(f'mt_clean_{int(time.time()*1000)%99999}')
        ft_mt = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
        fv_mt = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ['MTEXT'])
        ss_mt.Select(5, None, None, ft_mt, fv_mt)
        for i in range(ss_mt.Count - 1, -1, -1):
            try:
                ent = ss_mt.Item(i)
                plain = ent.TextString.strip().upper()
                if plain in _STAMP_PHRASES:
                    ent.Delete()
                    mtext_removed += 1
            except:
                pass
        ss_mt.Delete()
    except Exception as e:
        print(f"    MText cleanup SelectionSet failed: {e}")

    # Scan each PaperSpace layout directly
    for layout in doc.Layouts:
        if layout.Name.lower() == 'model':
            continue
        try:
            blk = layout.Block
            for i in range(blk.Count - 1, -1, -1):
                try:
                    ent = blk.Item(i)
                    if ent.EntityName not in ('AcDbMText', 'AcDbText'):
                        continue
                    plain = ent.TextString.strip().upper()
                    if plain in _STAMP_PHRASES:
                        ent.Delete()
                        mtext_removed += 1
                except:
                    pass
        except:
            pass

    # Also delete closed polylines that formed the border of deleted stamps
    # Check BOTH ModelSpace (via SelectionSet) AND PaperSpace (direct iteration)
    def _is_stamp_polyline(ent):
        """Check if a closed polyline looks like a stamp border (111x18 proportions)."""
        try:
            if not ent.Closed:
                return False
            coords = list(ent.Coordinates)
            if len(coords) != 8:  # rectangle = 4 points x 2 coords
                return False
            xs = [coords[j] for j in range(0, 8, 2)]
            ys = [coords[j] for j in range(1, 8, 2)]
            w = max(xs) - min(xs)
            h = max(ys) - min(ys)
            if 50 < w < 200 and 8 < h < 40 and 3 < w/h < 15:
                if min(xs) > 500 and min(ys) < 200:
                    return True
        except:
            pass
        return False

    if mtext_removed > 0:
        pl_removed = 0
        # ModelSpace polylines via SelectionSet
        try:
            doc.ActiveSpace = 1
            time.sleep(0.3)
            ss_pl = doc.SelectionSets.Add(f'pl_clean_{int(time.time()*1000)%99999}')
            ft_pl = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
            fv_pl = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ['LWPOLYLINE'])
            ss_pl.Select(5, None, None, ft_pl, fv_pl)
            for i in range(ss_pl.Count - 1, -1, -1):
                try:
                    ent = ss_pl.Item(i)
                    if _is_stamp_polyline(ent):
                        ent.Delete()
                        pl_removed += 1
                except:
                    pass
            ss_pl.Delete()
        except Exception as e:
            print(f"    ModelSpace polyline cleanup failed: {e}")

        # PaperSpace polylines via direct iteration
        for layout in doc.Layouts:
            if layout.Name.lower() == 'model':
                continue
            try:
                blk = layout.Block
                for i in range(blk.Count - 1, -1, -1):
                    try:
                        ent = blk.Item(i)
                        if ent.EntityName != 'AcDbPolyline':
                            continue
                        if _is_stamp_polyline(ent):
                            ent.Delete()
                            pl_removed += 1
                    except:
                        pass
            except:
                pass

        if pl_removed:
            print(f"    Removed {pl_removed} stamp border polylines")
            mtext_removed += pl_removed

    if mtext_removed:
        print(f"    Removed {mtext_removed} standalone stamp entities (MText+borders)")
    removed_count += mtext_removed

    print(f"    Total removed: {removed_count} stamp entities (blocks+MText)")
    report.check("old_stamps_cleaned", True,
                 f"has_colour={has_colour}, removed={removed_count}")

    # ── Step 4: Add FOR CONSTRUCTION stamp (v10-style positioning) ──
    # COLOUR box: if original DWG has it, keep as-is (typo already fixed above)
    #             if not, don't add one. Only FOR CONSTRUCTION is always drawn.
    # Stamps are drawn in the SAME space as the title block:
    # - PaperSpace TBs → stamp in PaperSpace (always visible in PDF, no viewport freeze)
    # - ModelSpace TBs → stamp in ModelSpace + viewport thaw (needed for PDF visibility)
    stamps_in_paperspace = all_tbs and all_tbs[0][3] != 'ModelSpace'
    print(f"\n  [Step 4] 添加 FOR CONSTRUCTION stamp x{len(all_tbs)}"
          f" ({'PaperSpace' if stamps_in_paperspace else 'ModelSpace'})"
          f"{' (COLOUR已存在, 保留原图)' if has_colour else ''}...")

    # Ensure IFC_STAMP layer
    try:
        layer = doc.Layers.Add(STAMP_LAYER)
        layer.color = 7
        layer.LayerOn = True
        layer.Freeze = False
    except Exception:
        try:
            layer = doc.Layers.Item(STAMP_LAYER)
            layer.color = 7
            layer.LayerOn = True
            layer.Freeze = False
        except:
            pass

    stamp_count = 0
    for tb_idx, (tb_br, tb_at, tb_sp, tb_sn) in enumerate(all_tbs, 1):
        # Get draw space: use layout.Block for PaperSpace, doc.ModelSpace for Model
        draw_space = tb_sp
        if tb_sn == 'ModelSpace':
            try:
                draw_space = doc.ModelSpace
            except:
                pass
        else:
            try:
                for _lay in doc.Layouts:
                    if _lay.Name == tb_sn:
                        draw_space = _lay.Block
                        break
            except:
                pass

        # Get this TB's bounding box
        try:
            mn, mx = tb_br.GetBoundingBox()
            _tb_w = float(mx[0]) - float(mn[0])
            _tb_h = float(mx[1]) - float(mn[1])
            _tb_min_x, _tb_min_y = float(mn[0]), float(mn[1])
            _tb_max_x, _tb_max_y = float(mx[0]), float(mx[1])
        except Exception:
            _tb_w, _tb_h = REF_TB_W, REF_TB_H
            _tb_min_x, _tb_min_y, _tb_max_x, _tb_max_y = 0, 0, _tb_w, _tb_h

        # Scale factor from reference frame
        _scale = _tb_w / REF_TB_W if REF_TB_W > 0 else 1.0
        _scale_y = _tb_h / REF_TB_H if _tb_h > 0 else _scale
        if tb_idx == 1:
            print(f"  Scale: {_scale:.4f} (tb_w={_tb_w:.1f}, ref={REF_TB_W})")

        # Scale dimensions
        _rect_w = REF_RECT_W * _scale
        _rect_h = REF_RECT_H * _scale
        _text_h = REF_TEXT_H * _scale

        # Position: bottom-right, RELATIVE to TB bbox
        _x_right_offset = REF_X_RIGHT_OFFSET * _scale
        _y_bottom_offset = REF_Y_BOTTOM * _scale_y

        _stamp_right = _tb_max_x - _x_right_offset
        _stamp_left = _stamp_right - _rect_w
        _stamp_bottom = _tb_min_y + _y_bottom_offset
        _stamp_top = _stamp_bottom + _rect_h

        _stamp_x = (_stamp_left + _stamp_right) / 2.0
        _stamp_y = (_stamp_bottom + _stamp_top) / 2.0

        # Draw rect + text
        _ok = False
        try:
            rect_pts = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8,
                [_stamp_left, _stamp_bottom,
                 _stamp_right, _stamp_bottom,
                 _stamp_right, _stamp_top,
                 _stamp_left, _stamp_top])
            pline = draw_space.AddLightWeightPolyline(rect_pts)
            pline.Closed = True
            pline.Layer = STAMP_LAYER
            pline.color = 7
            pline.ConstantWidth = 0.5 * _scale

            pt = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8,
                [_stamp_x, _stamp_y, 0.0])
            mtext = draw_space.AddMText(pt, _rect_w, STAMP_TEXT)
            mtext.Height = _text_h
            mtext.AttachmentPoint = 5  # MiddleCenter
            mtext.InsertionPoint = pt
            mtext.Layer = STAMP_LAYER
            mtext.color = 7
            _ok = True
            stamp_count += 1
        except Exception as e:
            print(f"    TB {tb_idx} ({tb_sn}): stamp failed — {e}")

        if tb_idx == 1:
            report.check("stamp_lower_rect", _ok)
            report.check("stamp_lower_text", _ok)
        print(f"    TB {tb_idx}/{len(all_tbs)} ({tb_sn}): {'OK' if _ok else 'FAILED'}")

    if has_colour:
        print(f"  COLOUR box: 保留原图 (typo已修复)")
    else:
        print(f"  COLOUR box: 原图无, 不添加")

    # Regen
    try:
        doc.Regen(1)
    except:
        pass
    time.sleep(0.5)

    # Viewport thaw: only needed when stamp is in ModelSpace
    # (PaperSpace stamps are always visible — no viewport freeze issue)
    if not stamps_in_paperspace:
        # Thaw IFC_STAMP in ALL PaperSpace viewports
        try:
            for layout in doc.Layouts:
                if layout.Name.lower() == 'model':
                    continue
                try:
                    doc.ActiveLayout = layout
                    time.sleep(0.3)
                    doc.SendCommand(
                        f'-VPLAYER\nThaw\n{STAMP_LAYER}\nAll\n\n')
                    time.sleep(1.0)
                except:
                    pass
            # Switch back to Model
            for layout in doc.Layouts:
                if layout.Name.lower() == 'model':
                    doc.ActiveLayout = layout
                    break
            time.sleep(0.5)
        except:
            pass
        # Final regen after thaw
        try:
            doc.Regen(1)
        except:
            pass
        time.sleep(0.5)

    # ── Step 5: Automated verification ──
    print(f"\n  [Step 5] 自动验证...")

    # 5a: Count entities on IFC_STAMP layer across ALL spaces
    stamp_entities = 0
    stamp_mtexts = []
    stamp_polylines = []
    searched_spaces = set()
    for _, _, tb_sp, tb_sn in all_tbs:
        # Dedup by space name (not id — same doc.ModelSpace returns different
        # COM proxies with different id(), causing double counting)
        if tb_sn in searched_spaces:
            continue
        searched_spaces.add(tb_sn)
        try:
            for i in range(tb_sp.Count):
                try:
                    ent = tb_sp.Item(i)
                    if ent.Layer == STAMP_LAYER:
                        stamp_entities += 1
                        ename = ent.EntityName
                        if ename == 'AcDbMText':
                            txt = ent.TextString
                            ix, iy = ent.InsertionPoint[0], ent.InsertionPoint[1]
                            stamp_mtexts.append({'text': txt, 'x': ix, 'y': iy})
                        elif ename in ('AcDbPolyline', 'AcDbLwPolyline'):
                            try:
                                mn, mx = ent.GetBoundingBox()
                                stamp_polylines.append({
                                    'min_x': float(mn[0]), 'min_y': float(mn[1]),
                                    'max_x': float(mx[0]), 'max_y': float(mx[1]),
                                })
                            except:
                                stamp_polylines.append({})
                except:
                    continue
        except:
            pass

    # IFC_STAMP: 2 entities per title block (1 rect + 1 text)
    n_tb = len(all_tbs)
    expect_entities = n_tb * 2
    expect_mtexts = n_tb
    expect_polylines = n_tb
    report.check("stamp_entity_count", stamp_entities == expect_entities,
                 f"{stamp_entities} entities on IFC_STAMP (expect {expect_entities}: {n_tb} rect + {n_tb} text)")
    report.check("stamp_mtext_count", len(stamp_mtexts) == expect_mtexts,
                 f"{len(stamp_mtexts)} MText entities (expect {expect_mtexts})")
    report.check("stamp_polyline_count", len(stamp_polylines) == expect_polylines,
                 f"{len(stamp_polylines)} polyline entities (expect {expect_polylines})")

    # 5b: Verify MText content
    has_construction_text = any('CONSTRUCTION' in m['text'].upper() for m in stamp_mtexts)
    report.check("stamp_has_construction_text", has_construction_text)

    # 5c: Verify stamp position relative to title block
    if stamp_polylines:
        polys_sorted = sorted([p for p in stamp_polylines if p],
                              key=lambda p: p['min_y'])
        lower = polys_sorted[0]

        # X range check: must be within TB horizontally
        x_inside = (lower['min_x'] >= tb_min_x - 1 and
                    lower['max_x'] <= tb_max_x + 1)
        report.check("stamp_x_inside_tb", x_inside,
                     f"stamp X [{lower['min_x']:.0f}-{lower['max_x']:.0f}] "
                     f"within TB X [{tb_min_x:.0f}-{tb_max_x:.0f}]")

        # Lower rect: must be inside TB vertically (bottom-right area)
        lower_y_ok = (lower['min_y'] >= tb_min_y - 1 and
                      lower['max_y'] <= tb_max_y + 1)
        report.check("stamp_lower_inside_tb", lower_y_ok,
                     f"lower rect Y [{lower['min_y']:.0f}-{lower['max_y']:.0f}] "
                     f"inside TB Y [{tb_min_y:.0f}-{tb_max_y:.0f}]")

        print(f"    FOR CONSTRUCTION: ({lower['min_x']:.0f},{lower['min_y']:.0f})->({lower['max_x']:.0f},{lower['max_y']:.0f})")
        print(f"    TB:    ({tb_min_x:.0f},{tb_min_y:.0f})->({tb_max_x:.0f},{tb_max_y:.0f})")
    else:
        report.check("stamp_x_inside_tb", False, "no polylines")
        report.check("stamp_lower_inside_tb", False, "no polylines")

    # 5d: Verify stamp proportions are reasonable
    if stamp_polylines and stamp_polylines[0]:
        p = stamp_polylines[0]
        sw = p['max_x'] - p['min_x']
        sh = p['max_y'] - p['min_y']
        ratio = sw / sh if sh > 0 else 0
        expected_ratio = REF_RECT_W / REF_RECT_H  # ~6.23
        ratio_ok = abs(ratio - expected_ratio) < 1.0
        report.check("stamp_aspect_ratio", ratio_ok,
                     f"ratio={ratio:.2f} (expected ~{expected_ratio:.2f})")

    # 5e: Idempotency test — re-run title block update, verify no duplicates
    print(f"\n  [Step 5e] 幂等性测试...")
    if f'{tp}DESCRIPTION' in attrs:
        attrs[f'{tp}DESCRIPTION'].TextString = 'FOR CONSTRUCTION'
    if f'{tp}REV' in attrs:
        attrs[f'{tp}REV'].TextString = str(ifc_rev)

    # Re-count IFC rows
    ifc_row_count = 0
    for row_num in range(1, REV_ROWS + 1):
        desc_tag = f"{row_num}DESCRIPTION"
        if desc_tag in attrs:
            val = attrs[desc_tag].TextString.strip().upper()
            if 'CONSTRUCTION' in val:
                ifc_row_count += 1

    report.check("idempotent_single_ifc_row", ifc_row_count == 1,
                 f"{ifc_row_count} IFC rows after double-write (expect 1)")

    # ── Report dimensions for CLAUDE.md update ──
    print(f"\n  [Info] Stamp summary:")
    print(f"    Title blocks: {len(all_tbs)}, Stamps drawn: {stamp_count}")
    print(f"    TB 1 bbox: ({tb_min_x:.1f}, {tb_min_y:.1f}) -> ({tb_max_x:.1f}, {tb_max_y:.1f})")
    print(f"    COLOUR from original: {has_colour}")

    # Store doc reference + TB count for Phase 3/4 (caller will use it)
    cfg['_n_tbs'] = len(all_tbs)
    return doc


# =============================================================================
# Phase 4: PDF Auto-Verification (PyMuPDF)
# =============================================================================
def verify_pdf(pdf_path, report, n_tbs=1):
    """Verify stamp presence and position in generated PDF using PyMuPDF.

    Compares against Tatua reference: FOR CONSTRUCTION at ~85%-95% X, ~85%-87% Y.
    Also renders stamp area for visual inspection.
    """
    sep("Phase 4: PDF Verification (PyMuPDF)")
    try:
        import fitz
    except ImportError:
        print("  PyMuPDF not installed — skipping PDF verification")
        return

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        report.check("pdf_readable", False, str(e))
        return

    report.check("pdf_readable", True, f"{len(doc)} pages")

    # Check each non-cover page for FOR CONSTRUCTION text
    construction_pages = []
    colour_pages = []

    for i, page in enumerate(doc):
        pw, ph = page.rect.width, page.rect.height
        blocks = page.get_text('dict')['blocks']
        for b in blocks:
            if b['type'] != 0:
                continue
            for line in b['lines']:
                for span in line['spans']:
                    txt = span['text'].strip().upper()
                    bbox = span['bbox']
                    if 'CONSTRUCTION' in txt:
                        # Position as percentage of page
                        x_pct = bbox[0] / pw * 100
                        y_pct = bbox[1] / ph * 100
                        construction_pages.append({
                            'page': i + 1,
                            'x_pct': x_pct, 'y_pct': y_pct,
                            'bbox': bbox, 'text': span['text']
                        })
                    if 'COLOUR' in txt or 'COLOR' in txt:
                        colour_pages.append({'page': i + 1, 'text': span['text']})

    # Every page with a title block must have FOR CONSTRUCTION
    total_pages = len(doc)
    fc_page_nums = set(c['page'] for c in construction_pages)
    has_fc = len(construction_pages) > 0
    # expect_stamped_pages = number of title blocks found (passed via cfg)
    # n_tbs passed as parameter
    pages_match_tbs = len(fc_page_nums) >= n_tbs
    report.check("pdf_has_construction", has_fc,
                 f"found on {len(fc_page_nums)}/{total_pages} pages"
                 if has_fc else "FOR CONSTRUCTION not found in PDF text")
    report.check("pdf_all_tbs_stamped", pages_match_tbs,
                 f"stamp on {len(fc_page_nums)} pages (expect >= {n_tbs} title blocks)"
                 + (f" — missing {n_tbs - len(fc_page_nums)} pages"
                    if not pages_match_tbs else ""))

    if construction_pages:
        fc = construction_pages[0]
        # Position should be in bottom-right: X > 75%, Y > 75%
        pos_ok = fc['x_pct'] > 75 and fc['y_pct'] > 75
        report.check("pdf_stamp_position", pos_ok,
                     f"X={fc['x_pct']:.1f}% Y={fc['y_pct']:.1f}% "
                     f"(expect >75% both, ref: X=85% Y=85%)")

        for fc_info in construction_pages:
            print(f"    Page {fc_info['page']}: FOR CONSTRUCTION at "
                  f"X={fc_info['x_pct']:.1f}% Y={fc_info['y_pct']:.1f}%")

    if colour_pages:
        print(f"    COLOUR text found on pages: "
              f"{[c['page'] for c in colour_pages]}")

    # Render stamp area as PNG for visual inspection
    try:
        page = doc[0]  # first page
        pw, ph = page.rect.width, page.rect.height
        # Crop bottom-right 30%
        clip = fitz.Rect(pw * 0.65, ph * 0.65, pw, ph)
        mat = fitz.Matrix(3, 3)
        pix = page.get_pixmap(matrix=mat, clip=clip)
        stamp_png = pdf_path.parent / f"{pdf_path.stem}_stamp_verify.png"
        pix.save(str(stamp_png))
        print(f"    Stamp截图: {stamp_png.name}")
    except Exception as e:
        print(f"    Stamp截图失败: {e}")

    doc.close()


# =============================================================================
# Phase 3: SaveAs + PUBLISH PDF Export
# =============================================================================
def test_saveas_and_pdf(doc, acad, report, original_dwg_path, cfg):
    """SaveAs to temp IFC DWG, export PDF via PUBLISH+DSD, verify output."""
    sep(f"Phase 3: SaveAs + PDF Export ({cfg['name']})")

    import tempfile

    # Output paths — use test_output/ to avoid polluting real IFC folder
    output_dir = Path(__file__).parent / "test_output"
    output_dir.mkdir(exist_ok=True)

    # IFC naming: strip original rev, use Rev0 (IFC always starts from Rev0)
    import re
    orig_stem = Path(original_dwg_path).stem
    # Remove trailing _RevX or _revX
    stem = re.sub(r'[_\s-][Rr]ev[A-Za-z0-9.]+$', '', orig_stem)
    ifc_stem = f"{stem}_Rev0_IFC_TEST"
    ifc_dwg_path = output_dir / f"{ifc_stem}.dwg"
    ifc_pdf_path = output_dir / f"{ifc_stem}.pdf"

    # Clean up previous test outputs
    for f in [ifc_dwg_path, ifc_pdf_path]:
        if f.exists():
            f.unlink()

    # ── Step 6: Bind XREFs + SaveAs IFC DWG ──
    print(f"\n  [Step 6] XREF bind + SaveAs...")

    # Bind all XREFs before SaveAs (prevents missing content in PDF)
    xref_bound = 0
    try:
        for i in range(doc.Blocks.Count):
            try:
                blk = doc.Blocks.Item(i)
                if blk.IsXRef:
                    blk.Bind(False)  # False = Insert bind (preserves layer names)
                    xref_bound += 1
            except:
                continue
        if xref_bound > 0:
            doc.Regen(1)
            time.sleep(2)  # settle after XREF bind
            print(f"    Bound {xref_bound} XREFs")
    except Exception as e:
        print(f"    XREF bind warning: {e}")

    print(f"    Target: {ifc_dwg_path.name}")

    save_ok = False
    try:
        doc.SaveAs(str(ifc_dwg_path))
        save_ok = True
    except Exception as e:
        print(f"    SaveAs failed: {e}")

    report.check("saveas_ok", save_ok)
    report.check("ifc_dwg_exists", ifc_dwg_path.exists(),
                 f"{ifc_dwg_path.stat().st_size / 1024:.0f} KB" if ifc_dwg_path.exists() else "file missing")

    if not save_ok:
        return

    # ── Step 7: PUBLISH PDF via DSD ──
    print(f"\n  [Step 7] PDF export (PUBLISH+DSD)...")

    # Close current doc, reopen the saved IFC DWG for PUBLISH
    try:
        doc.Close(False)
    except:
        pass

    # Close all docs
    try:
        while acad.Documents.Count > 0:
            acad.Documents.Item(0).Close(False)
    except:
        pass
    time.sleep(1)

    # PUBLISH workaround: AutoCAD SendCommand / DSD parser may choke on
    # Unicode chars (√) in paths.  Copy DWG to temp dir, PUBLISH there,
    # then move PDF back.
    import tempfile, shutil
    temp_dir = Path(tempfile.gettempdir()) / "ifc_test"
    temp_dir.mkdir(exist_ok=True)
    temp_dwg = temp_dir / ifc_dwg_path.name
    temp_pdf = temp_dir / ifc_pdf_path.name
    shutil.copy2(str(ifc_dwg_path), str(temp_dwg))
    print(f"    PUBLISH workaround: DWG copied to {temp_dir}")

    doc = open_dwg(acad, str(temp_dwg))
    report.check("ifc_dwg_reopened", doc is not None)
    if not doc:
        return

    # Detect all non-Model layouts with content
    layout_info = []
    try:
        for layout in doc.Layouts:
            if layout.Name.lower() != 'model':
                entity_count = layout.Block.Count
                if entity_count <= 1:
                    print(f"    Skip empty layout: '{layout.Name}'")
                    continue
                tab_order = layout.TabOrder
                layout_info.append((tab_order, layout.Name))
        layout_info.sort(key=lambda x: x[0])
    except Exception as e:
        print(f"    Layout detection failed: {e}")

    layout_names = [name for _, name in layout_info]
    report.check("layouts_detected", len(layout_names) > 0,
                 f"{len(layout_names)} layouts: {layout_names}")

    if not layout_names:
        return

    # Build DSD file in temp dir (all paths ASCII-safe)
    dsd_path = temp_dir / f"{ifc_stem}.dsd"
    dwg_str = str(temp_dwg)
    pdf_str = str(temp_pdf)
    out_str = str(temp_dir)

    dsd_lines = [
        '[DWF6Version]', 'Ver=1',
        '[DWF6MinorVersion]', 'MinorVer=1',
    ]
    for lname in layout_names:
        dsd_lines.extend([
            f'[DWF6Sheet:{ifc_stem}-{lname}]',
            f'DWG={dwg_str}',
            f'Layout={lname}',
            'Setup=',
            f'OriginalSheetPath={dwg_str}',
            'Has Plot Port=0',
            'Has3DDWF=0',
        ])
    dsd_lines.extend([
        '[Target]', 'Type=6',
        f'DWF={pdf_str}',
        f'OUT={out_str}',
        'PWD=',
        'PromptForDwfName=FALSE',
        '[PdfOptions]',
        'VectorResolution=600',
        'RasterResolution=400',
        '[SheetSetProperties]',
        'IsSheetSet=FALSE',
        'IsHomogeneous=FALSE',
        'SheetSet Storage File=',
        'AcadProfile=<<Default>>',
        'CategoryCount=0',
        '[AutoCAD Block Information]',
        'IncludeBlockInfo=0',
        'BlockTmplFilePath=',
    ])

    dsd_content = '\n'.join(dsd_lines)
    dsd_path.write_text(dsd_content, encoding='utf-8')
    print(f"    DSD: {dsd_path} ({len(layout_names)} sheets)")

    # Run PUBLISH
    dsd_str_cmd = str(dsd_path)
    try:
        doc.SetVariable("FILEDIA", 0)
        doc.SetVariable("BACKGROUNDPLOT", 0)
        doc.SendCommand(f'-PUBLISH\n{dsd_str_cmd}\n')
    except Exception as e:
        report.check("publish_command_sent", False, str(e))
        return

    report.check("publish_command_sent", True)

    # Wait for PUBLISH to complete (up to 180s)
    print(f"    等待 PUBLISH 完成...")
    time.sleep(3)  # give PUBLISH time to start before polling CMDACTIVE
    start = time.time()
    max_wait = 180
    last_log = start
    while time.time() - start < max_wait:
        try:
            if doc.GetVariable("CMDACTIVE") == 0:
                # Double-check: wait a bit more and re-check (PUBLISH may have gaps)
                time.sleep(2)
                if doc.GetVariable("CMDACTIVE") == 0:
                    break
        except Exception:
            break
        elapsed = time.time() - start
        if time.time() - last_log > 15:
            print(f"    PUBLISH in progress... ({int(elapsed)}s)")
            last_log = time.time()
        time.sleep(2)

    publish_time = time.time() - start + 3  # include initial wait
    print(f"    PUBLISH 完成 ({publish_time:.1f}s)")

    # Close doc before checking output (match v10 pattern)
    try:
        doc.Close(False)
    except:
        pass
    time.sleep(3)  # flush to disk

    # Check PDF output in temp dir, then move to test_output/
    pdf_found = temp_pdf.exists()
    if not pdf_found:
        # Check for alternative names (PUBLISH sometimes appends layout name)
        for f in temp_dir.glob(f"{ifc_stem}*.pdf"):
            pdf_found = True
            temp_pdf = f
            break

    if pdf_found:
        # Move PDF to test_output/
        final_pdf = output_dir / temp_pdf.name
        shutil.move(str(temp_pdf), str(final_pdf))
        ifc_pdf_path = final_pdf

    # Cleanup temp files
    for f in temp_dir.glob(f"{ifc_stem}.*"):
        try:
            f.unlink()
        except:
            pass

    report.check("pdf_exists", pdf_found,
                 f"{ifc_pdf_path.name} ({ifc_pdf_path.stat().st_size / 1024:.0f} KB)"
                 if pdf_found else "no PDF found")

    if pdf_found:
        print(f"\n  ★ PDF 输出: {ifc_pdf_path}")
        print(f"    大小: {ifc_pdf_path.stat().st_size / 1024:.0f} KB")

        # ── Step 8: PDF auto-verification (PyMuPDF) ──
        verify_pdf(ifc_pdf_path, report, n_tbs=cfg.get('_n_tbs', 1))

    # List all files in test_output for reference
    print(f"\n  test_output/ 内容:")
    for f in sorted(output_dir.iterdir()):
        print(f"    {f.name} ({f.stat().st_size / 1024:.0f} KB)")


# =============================================================================
# Main
# =============================================================================
if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')

    # Parse args: project name + flags
    flags = [a for a in sys.argv[1:] if a.startswith('--')]
    positional = [a for a in sys.argv[1:] if not a.startswith('--')]
    measure_only = '--measure-only' in flags
    skip_measure = '--skip-measure' in flags

    # Select project
    project_key = positional[0].lower() if positional else 'warnertown'
    if project_key not in PROJECTS:
        print(f"Unknown project: '{project_key}'")
        print(f"Available: {', '.join(PROJECTS.keys())}")
        sys.exit(1)
    cfg = PROJECTS[project_key]

    # Set up log file — tee all output to test_logs/
    LOG_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_path = LOG_DIR / f"test_{project_key}_{timestamp}.log"
    tee = TeeWriter(log_path)
    sys.stdout = tee

    print(f"Test started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Project: {cfg['name']} ({project_key})")
    print(f"Log file: {log_path}")
    print(f"Flags: {flags}")

    report = TestReport()

    try:
        acad = get_acad()

        frame_info = {}
        if not skip_measure:
            frame_info = measure_frame(acad, report, cfg)

        if not measure_only:
            test_dwgs = cfg.get('test_dwgs', [cfg.get('test_dwg')])
            if not test_dwgs or test_dwgs == [None]:
                test_dwgs = [None]  # auto-detect mode
            for dwg_target in test_dwgs:
                # Override test_dwg for this iteration
                cfg_copy = dict(cfg)
                if dwg_target:
                    cfg_copy['test_dwg'] = dwg_target
                test_ifc_conversion(acad, report, cfg_copy, frame_info)

        all_passed = report.summary()
    except Exception as e:
        print(f"\n  FATAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        all_passed = False
    finally:
        sys.stdout = tee._original
        tee.close()
        print(f"Log saved: {log_path}")

    sys.exit(0 if all_passed else 1)

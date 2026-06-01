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
        'test_dwg': None,  # auto-detect smallest DWG
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

# ── Stamp constants (universal — from Tatua reference, cross-project) ────
STAMP_LAYER = 'IFC_STAMP'
STAMP_TEXT = "{\\fArial Narrow|b1;FOR CONSTRUCTION}"
COLOUR_TEXT = "{\\fArial Narrow|b1;DRAWINGS TO BE\\PPRINTED IN COLOUR}"

REF_TB_W = 841.0
REF_TB_H = 594.0
REF_STAMP_X = 141.0
REF_STAMP_Y = 589.176
REF_RECT_W = 110.511
REF_RECT_H = 17.745
REF_TEXT_Y_OFFSET = 13.182
REF_TEXT_H = 7.0
REF_TEXT_W = 116.419


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
        self._log.close()

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


def find_test_dwg(cfg):
    """Find a suitable test DWG — use config test_dwg or auto-detect smallest."""
    native_root = os.path.join(cfg['project_path'], r"Design\Engineering\1. Drawings\1. Native")
    if cfg.get('test_dwg') and os.path.exists(cfg['test_dwg']):
        return cfg['test_dwg']
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

    # ── Step 1: Find title block ──
    print(f"\n  [Step 1] 查找 title block '{title_block_name}'...")
    block_ref = None
    attrs = {}
    space = None
    tb_space_name = None

    for sname, get_space in [("ModelSpace", lambda: doc.ModelSpace),
                              ("PaperSpace", lambda: doc.PaperSpace)]:
        try:
            s = get_space()
            count = s.Count
            for i in range(count):
                try:
                    ent = s.Item(i)
                    if ent.EntityName != 'AcDbBlockReference':
                        continue
                    bname = ent.Name
                    if bname == title_block_name:
                        block_ref = ent
                        attrs = get_attrs_dict(ent)
                        space = s
                        tb_space_name = sname
                        break
                    a = get_attrs_dict(ent)
                    if 'DRAWINGNUMBER' in a and 'REVISION' in a:
                        block_ref = ent
                        attrs = a
                        space = s
                        tb_space_name = sname
                        break
                except Exception:
                    continue
            if block_ref:
                break
        except Exception:
            continue

    report.check("title_block_found", block_ref is not None,
                 f"'{block_ref.Name}' in {tb_space_name} ({len(attrs)} attrs)" if block_ref else "not found")
    if not attrs:
        # Debug: list all block refs
        try:
            ms = doc.ModelSpace
            names = []
            for i in range(min(ms.Count, 200)):
                ent = ms.Item(i)
                try:
                    if ent.EntityName == 'AcDbBlockReference':
                        names.append(ent.Name)
                except:
                    pass
            print(f"  Block refs found: {set(names)}")
        except:
            pass
        return

    report.check("attrs_count_reasonable", len(attrs) >= 20,
                 f"{len(attrs)} attrs (expect ~69 for Warnertown)")

    # Get title block dimensions
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

    # ── Step 4: Add stamp (v9-style direct COM) ──
    print(f"\n  [Step 4] 添加 stamp (v9 style)...")

    # Scale factor from reference frame
    scale = tb_w / REF_TB_W if REF_TB_W > 0 else 1.0
    print(f"  Scale: {scale:.4f} (tb_w={tb_w:.1f}, ref={REF_TB_W})")

    # Scale dimensions (size only, not position)
    rect_w = REF_RECT_W * scale
    rect_h = REF_RECT_H * scale
    text_y_offset = REF_TEXT_Y_OFFSET * scale
    text_h = REF_TEXT_H * scale
    text_w = REF_TEXT_W * scale

    # Position RELATIVE to title block bbox (not absolute from origin)
    # In Tatua reference (841x594): stamp_x=141 is offset from LEFT edge,
    # stamp_y=589.176 is offset from BOTTOM edge
    stamp_x = tb_min_x + REF_STAMP_X * scale
    stamp_y = tb_min_y + REF_STAMP_Y * scale

    rect_half_w = rect_w / 2.0
    rect_left = stamp_x - rect_half_w
    rect_right = stamp_x + rect_half_w
    rect_bottom = stamp_y - text_y_offset
    rect_top = rect_bottom + rect_h

    # Ensure IFC_STAMP layer
    try:
        layer = doc.Layers.Add(STAMP_LAYER)
        layer.color = 7
    except Exception:
        try:
            layer = doc.Layers.Item(STAMP_LAYER)
            layer.color = 7
        except:
            pass

    draw_space = space

    # Lower stamp: FOR CONSTRUCTION
    lower_rect_ok = False
    lower_text_ok = False
    try:
        rect_pts = win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            [rect_left, rect_bottom,
             rect_right, rect_bottom,
             rect_right, rect_top,
             rect_left, rect_top])
        pline = draw_space.AddLightWeightPolyline(rect_pts)
        pline.Closed = True
        pline.Layer = STAMP_LAYER
        pline.color = 7
        lower_rect_ok = True
    except Exception as e:
        print(f"  Lower rect failed: {e}")

    try:
        pt = win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            [stamp_x, stamp_y, 0.0])
        mtext = draw_space.AddMText(pt, text_w, STAMP_TEXT)
        mtext.Height = text_h
        mtext.AttachmentPoint = 2  # TopCenter
        mtext.Layer = STAMP_LAYER
        mtext.color = 7
        lower_text_ok = True
    except Exception as e:
        print(f"  Lower text failed: {e}")

    report.check("stamp_lower_rect", lower_rect_ok)
    report.check("stamp_lower_text", lower_text_ok)

    # Upper stamp: PRINTED IN COLOUR
    upper_rect_ok = False
    upper_text_ok = False
    upper_rect_bottom = rect_top
    upper_rect_h = rect_h * 1.15
    upper_rect_top = upper_rect_bottom + upper_rect_h
    upper_text_y = upper_rect_top - (upper_rect_h - text_y_offset) / 2

    try:
        rect_pts2 = win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            [rect_left, upper_rect_bottom,
             rect_right, upper_rect_bottom,
             rect_right, upper_rect_top,
             rect_left, upper_rect_top])
        pline2 = draw_space.AddLightWeightPolyline(rect_pts2)
        pline2.Closed = True
        pline2.Layer = STAMP_LAYER
        pline2.color = 7
        upper_rect_ok = True
    except Exception as e:
        print(f"  Upper rect failed: {e}")

    try:
        pt2 = win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            [stamp_x, upper_text_y, 0.0])
        mtext2 = draw_space.AddMText(pt2, text_w, COLOUR_TEXT)
        mtext2.Height = text_h
        mtext2.AttachmentPoint = 2
        mtext2.Layer = STAMP_LAYER
        mtext2.color = 7
        upper_text_ok = True
    except Exception as e:
        print(f"  Upper text failed: {e}")

    report.check("stamp_upper_rect", upper_rect_ok)
    report.check("stamp_upper_text", upper_text_ok)

    try:
        doc.Regen(1)
    except:
        pass
    time.sleep(0.5)

    # ── Step 5: Automated verification ──
    print(f"\n  [Step 5] 自动验证...")

    # 5a: Count entities on IFC_STAMP layer
    stamp_entities = 0
    stamp_mtexts = []
    stamp_polylines = []
    try:
        for i in range(draw_space.Count):
            try:
                ent = draw_space.Item(i)
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

    report.check("stamp_entity_count", stamp_entities == 4,
                 f"{stamp_entities} entities on IFC_STAMP (expect 4: 2 rects + 2 texts)")
    report.check("stamp_mtext_count", len(stamp_mtexts) == 2,
                 f"{len(stamp_mtexts)} MText entities")
    report.check("stamp_polyline_count", len(stamp_polylines) == 2,
                 f"{len(stamp_polylines)} polyline entities")

    # 5b: Verify MText content
    has_construction = any('CONSTRUCTION' in m['text'].upper() for m in stamp_mtexts)
    has_colour = any('COLOUR' in m['text'].upper() for m in stamp_mtexts)
    report.check("stamp_has_construction_text", has_construction)
    report.check("stamp_has_colour_text", has_colour)

    # 5c: Verify stamp position relative to title block
    # Design: lower rect (FOR CONSTRUCTION) sits at frame top edge,
    #         upper rect (COLOUR) extends ABOVE frame (by design, same as Tatua reference)
    # Check: X must be within TB, lower rect Y must touch/be near TB top edge
    if len(stamp_polylines) >= 2:
        # Sort by min_y to identify lower vs upper
        polys_sorted = sorted([p for p in stamp_polylines if p],
                              key=lambda p: p['min_y'])
        lower = polys_sorted[0]
        upper = polys_sorted[1]

        # X range check: both must be within TB horizontally
        x_inside = (lower['min_x'] >= tb_min_x - 1 and
                    lower['max_x'] <= tb_max_x + 1)
        report.check("stamp_x_inside_tb", x_inside,
                     f"stamp X [{lower['min_x']:.0f}-{lower['max_x']:.0f}] "
                     f"within TB X [{tb_min_x:.0f}-{tb_max_x:.0f}]")

        # Lower rect: bottom should be inside TB, top near TB top edge
        lower_y_ok = (lower['min_y'] >= tb_min_y - 1 and
                      lower['max_y'] <= tb_max_y + rect_h * 0.5)  # small tolerance
        report.check("stamp_lower_position", lower_y_ok,
                     f"lower rect Y [{lower['min_y']:.0f}-{lower['max_y']:.0f}] "
                     f"near TB top {tb_max_y:.0f}")

        # Upper rect: sits above lower, may extend beyond TB top (by design)
        upper_stacked = abs(upper['min_y'] - lower['max_y']) < 2
        report.check("stamp_upper_stacked", upper_stacked,
                     f"upper bottom={upper['min_y']:.1f} ≈ lower top={lower['max_y']:.1f}")

        # Log for reference
        print(f"    Lower: ({lower['min_x']:.0f},{lower['min_y']:.0f})->({lower['max_x']:.0f},{lower['max_y']:.0f})")
        print(f"    Upper: ({upper['min_x']:.0f},{upper['min_y']:.0f})->({upper['max_x']:.0f},{upper['max_y']:.0f})")
        print(f"    TB:    ({tb_min_x:.0f},{tb_min_y:.0f})->({tb_max_x:.0f},{tb_max_y:.0f})")
    elif stamp_polylines:
        report.check("stamp_x_inside_tb", False, "expected 2 polylines")
        report.check("stamp_lower_position", False, "expected 2 polylines")
        report.check("stamp_upper_stacked", False, "expected 2 polylines")
    else:
        report.check("stamp_x_inside_tb", False, "no polylines")
        report.check("stamp_lower_position", False, "no polylines")
        report.check("stamp_upper_stacked", False, "no polylines")

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
    print(f"\n  [Info] Stamp dimensions used:")
    print(f"    Scale factor: {scale:.4f}")
    print(f"    TB bbox: ({tb_min_x:.1f}, {tb_min_y:.1f}) -> ({tb_max_x:.1f}, {tb_max_y:.1f})")
    print(f"    Stamp center: ({stamp_x:.2f}, {stamp_y:.2f})")
    print(f"    Lower rect: ({rect_left:.2f}, {rect_bottom:.2f}) -> ({rect_right:.2f}, {rect_top:.2f})")
    print(f"    Upper rect: ({rect_left:.2f}, {upper_rect_bottom:.2f}) -> ({rect_right:.2f}, {upper_rect_top:.2f})")
    print(f"    Text height: {text_h:.2f}")

    # Store doc reference for Phase 3 (caller will use it)
    return doc


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

    stem = Path(original_dwg_path).stem
    ifc_dwg_path = output_dir / f"{stem}_IFC_TEST.dwg"
    ifc_pdf_path = output_dir / f"{stem}_IFC_TEST.pdf"

    # Clean up previous test outputs
    for f in [ifc_dwg_path, ifc_pdf_path]:
        if f.exists():
            f.unlink()

    # ── Step 6: SaveAs IFC DWG ──
    print(f"\n  [Step 6] SaveAs...")
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

    doc = open_dwg(acad, str(ifc_dwg_path))
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

    # Build DSD file
    dsd_path = output_dir / f"{stem}_IFC_TEST.dsd"
    dwg_str = str(ifc_dwg_path).replace('\\', '/')
    pdf_str = str(ifc_pdf_path).replace('\\', '/')

    dsd_lines = [
        '[DWF6Version]', 'Ver=1',
        '[DWF6MinorVersion]', 'MinorVer=1',
    ]
    for lname in layout_names:
        dsd_lines.extend([
            f'[DWF6Sheet:{stem}_IFC_TEST-{lname}]',
            f'DWG={dwg_str}',
            f'Layout={lname}',
            'Setup=',
            f'OriginalSheetPath={dwg_str}',
            'Has Plot Port=0',
            'Has3DDWF=0',
        ])
    dsd_lines.extend([
        '[Target]',
        f'Type=6',
        f'DWF={pdf_str}',
        f'OUT={str(output_dir).replace(chr(92), "/")}/',
        'PWD=',
        '[PdfOptions]',
        'VectorResolution=600',
        'RasterResolution=400',
        '[SheetSetProperties]',
        'IsSheetSet=FALSE',
        'IsHomogeneous=FALSE',
        'SheetSet Name=',
        'NoOfCopies=1',
        'PlotStampOn=FALSE',
        'JobID=0',
        'SelectionSetName=',
        'AcadProfile=',
        'CategoryName=',
        'LogFilePath=',
        f'IncludeLayer=TRUE',
        f'LineMerge=FALSE',
        f'CurrentPrecision=',
        'PromptForDwfName=FALSE',
    ])

    dsd_content = '\n'.join(dsd_lines) + '\n'
    dsd_path.write_text(dsd_content, encoding='utf-8')
    print(f"    DSD: {dsd_path.name} ({len(layout_names)} sheets)")

    # Run PUBLISH
    dsd_str_cmd = str(dsd_path).replace('\\', '/')
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

    # Check PDF output
    # PUBLISH may output with a slightly different name — check for any PDF
    time.sleep(2)  # settle time
    pdf_found = ifc_pdf_path.exists()
    if not pdf_found:
        # Check for alternative names (PUBLISH sometimes appends layout name)
        for f in output_dir.glob(f"{stem}*.pdf"):
            pdf_found = True
            ifc_pdf_path = f
            break

    report.check("pdf_exists", pdf_found,
                 f"{ifc_pdf_path.name} ({ifc_pdf_path.stat().st_size / 1024:.0f} KB)"
                 if pdf_found else "no PDF found in test_output/")

    if pdf_found:
        print(f"\n  ★ PDF 输出: {ifc_pdf_path}")
        print(f"    大小: {ifc_pdf_path.stat().st_size / 1024:.0f} KB")
        print(f"    请打开检查: stamp位置、图框修订行、内容完整性")

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
            test_ifc_conversion(acad, report, cfg, frame_info)

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

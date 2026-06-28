"""
asbuilt_revclean_lms.py  —  bespoke NAWMA / LMS (50023-) AS BUILT title-block rev cleanup.

RECONSTRUCTED 2026-06-10 after the original was lost in a machine restart. Faithful to the
recovered spec in memories `lms-asbuilt-revclean-com-stability` + `lms-titleblock-baked-static-
rev-rows` and the API used by the session inspectors. NOT part of the IFR/IFC automation — this
is project-specific: it edits the EXISTING AS BUILT title block (keep only the AS BUILT row at a
given rev, strip the other revision rows, optionally remove revision clouds), then archives the
deliverable native into its per-drawing AS BUILT folder and copies a friendly-named PDF to the
client delivery folder.

Usage:
  python "asbuilt_revclean_lms.py" --target ga001 --dry-run        # inspect plan, no writes
  python "asbuilt_revclean_lms.py" --target ga001 --apply          # edit + save + PDF + archive
  python "asbuilt_revclean_lms.py" --target ca001_trench --apply --no-pdf
  python "asbuilt_revclean_lms.py" --target wave2 --apply           # all Wave-2 drawings

COM stability rules baked in (these crashed acad.exe until fixed):
  * search title blocks in PaperSpace layouts ONLY; never iterate ModelSpace entity-by-entity.
  * remove revision clouds with native -LAYDEL; only fall back to per-entity COM .Delete() for
    the few residuals -LAYDEL leaves on exploded -exp files.
  * grab acad.ActiveDocument after Open() (late-bound Open() return doesn't resolve).
  * SAVE attribute/block-def edits BEFORE the cloud step.
"""
import argparse
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

try:
    import win32com.client  # noqa
except Exception:
    win32com = None

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ----------------------------------------------------------------------------- paths / consts
NATIVE = Path(r"C:\Users\jilin\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\LMS"
              r"\Design\Engineering\1. Drawings\1. Native")
CLIENT_PDF_DIR = NATIVE.parent / "5. As Built" / "3. As Built Client"

PROJ_NO = "50023"
PROJ_DESC = "ADAMS RD, ULEYBURY, SA 5114"

MAX_ROWS = 8
# per-row title-block attribute tags are f"{n}{TAG}" (n = 1..MAX_ROWS) plus f"{n}PROJECT".
PERSONNEL_TAGS = ["REV", "DATE", "DESCRIPTION", "DESIGNED", "DRAWN", "CHECK", "APPROVED"]
ROW_TAGS = PERSONNEL_TAGS + ["PROJECT"]

CONSTR_KEYS = ("FOR CONSTRUCTION", "ISSUED FOR CONSTRUCTION", "FOR REVIEW",
               "UPDATE CONSTRUCTION", "UPDATE")

# dedicated revision-cloud layers seen on LMS drawings
CLOUD_LAYER_NAMES = {"cloud", "revision", "revisions", "revcloud", "rev cloud"}

DEFAULT_TB = {"riverina_tellhow", "lms_ace", "coleamablly"}


def _f(*parts):
    return NATIVE.joinpath(*parts)


def _lp(p):
    r"""Windows long-path (\\?\ prefix) form so files with path>=260 chars are visible."""
    p = str(p)
    return p if p.startswith("\\\\?\\") else "\\\\?\\" + p


def _exists(p):
    """MAX_PATH-safe existence check (plain Path.exists() returns False at >=260 chars)."""
    return os.path.exists(str(p)) or os.path.exists(_lp(p))


# ----------------------------------------------------------------------------- targets
# Each target: src (deliverable native to edit — the -exp where it exists), ab_dir (per-drawing
# AS BUILT folder to archive into), tb_blocks, new_rev, remove_clouds, draw_no, client_name.
TARGETS = {
    # ---- already completed earlier (kept for regression-testing the reconstruction) ----
    "ga001": {
        "label": "GA-001 SITE PLAN",
        "src": _f("50023-GA-001_Civil Site Layout Plan(Existing)", "Rev.1 - AB",
                  "50023-GA-001-Rev1-Site Plan_AB-exp.dwg"),
        "ab_dir": _f("50023-GA-001_Civil Site Layout Plan(Existing)", "Rev.1 - AB"),
        "tb_blocks": DEFAULT_TB, "new_rev": "8", "remove_clouds": False,
        "draw_no": "50023-GA-001",
        "client_name": "50023-GA-001 REV 8 Site Plan_AS BUILT.pdf",
    },
    # ---- Wave 2 (CSV target revisions) ----
    "ca001_trench": {
        "label": "CA-001 TRENCH ALIGNMENT LAYOUT PLAN",
        "src": _f("50023-CA-001_Trench Alignment Layout Plan(with coordination& cross section)(Existing)",
                  "50023-CA-001_REV1 Trench Alignment Layout Plan_AS BUILT-exp.dwg"),
        "ab_dir": _f("50023-CA-001_Trench Alignment Layout Plan(with coordination& cross section)(Existing)",
                     "Rev.1 - AB"),
        "tb_blocks": DEFAULT_TB, "new_rev": "6", "remove_clouds": False,
        "draw_no": "50023-CA-001",
        "client_name": "50023-CA-001 REV 6 Trench Alignment Layout Plan_AS BUILT.pdf",
    },
    "ca011": {
        "label": "CA-011 HV RMU CABLE PIT CIVIL LAYOUT",
        "src": _f("50023-CA-011_Cable Pit Civil Layout(Existing)",
                  "50023-CA-011 REV 1 - HV RMU CABLE PIT CIVIL LAYOUT_AB-exp.dwg"),
        "ab_dir": _f("50023-CA-011_Cable Pit Civil Layout(Existing)", "Rev.1 - AB"),
        "tb_blocks": DEFAULT_TB, "new_rev": "2", "remove_clouds": False,
        "draw_no": "50023-CA-011",
        "client_name": "50023-CA-011 REV 2 HV RMU Cable Pit Civil Layout_AS BUILT.pdf",
    },
    "ea013": {
        "label": "EA-013 CONTROL BUILDING",
        "src": _f("50023-EA-013_Control Building",
                  "50023-EA-013 Rev1 - CONTROL BUILDING_AB.dwg"),
        "ab_dir": _f("50023-EA-013_Control Building", "Rev.1 - AB"),
        "tb_blocks": DEFAULT_TB, "new_rev": "2", "remove_clouds": False,
        "draw_no": "50023-EA-013",
        "client_name": "50023-EA-013 REV 2 Control Building_AS BUILT.pdf",
    },
    "ea300": {
        "label": "EA-300 HV RMU ENCLOSURE GENERAL ARRANGEMENT",
        "src": _f("50023-EA-300_(New) HV RMU Enclosure General Arrangement", "Rev.1 - AB",
                  "50023-EA-300_HV RMU Enclosure General Arrangement_IFC_R2_AB.dwg"),
        "ab_dir": _f("50023-EA-300_(New) HV RMU Enclosure General Arrangement", "Rev.1 - AB"),
        "tb_blocks": DEFAULT_TB, "new_rev": "3", "remove_clouds": False,
        "draw_no": "50023-EA-300",
        "client_name": "50023-EA-300 REV 3 HV RMU Enclosure General Arrangement_AS BUILT.pdf",
    },
    "ea301": {
        "label": "EA-301 IP4 GENERAL ARRANGEMENT (multi-sheet)",
        "src": _f("50023-EA-301_(New) HV Interface Panel 4 General Arrangement", "Rev.1 - AB",
                  "50023-EA-301 Rev.2 - Interface Panel 4 (IP4) - General Arrangement_AB.dwg"),
        "ab_dir": _f("50023-EA-301_(New) HV Interface Panel 4 General Arrangement", "Rev.1 - AB"),
        "tb_blocks": DEFAULT_TB, "new_rev": "3", "remove_clouds": False,
        "draw_no": "50023-EA-301",
        "client_name": "50023-EA-301 REV 3 Interface Panel 4 (IP4) General Arrangement_AS BUILT.pdf",
    },
    "el002": {
        "label": "EL-002 PROTECTION KEY DIAGRAM",
        "src": _f("50023-EL-002_Protection Key Diagram (Existing)",
                  "50023-EL-002_Rev 1 - PROPOSED BESS_PROTECTION KEY DIAGRAM_As Built.dwg"),
        "ab_dir": _f("50023-EL-002_Protection Key Diagram (Existing)", "Rev.1 - AB"),
        "tb_blocks": DEFAULT_TB, "new_rev": "5", "remove_clouds": False,
        "draw_no": "50023-EL-002",
        "client_name": "50023-EL-002 REV 5 Protection Key Diagram_AS BUILT.pdf",
    },
}

WAVE2 = ["ca001_trench", "ca011", "ea013", "ea300", "ea301", "el002"]


# ----------------------------------------------------------------------------- COM helpers
def com_retry(fn, tries=5, wait=2.0):
    """Retry a COM call through transient 'busy/rejected' errors. Short-circuit on the
    'No database' error (retrying that one hangs for 20+ min — never retry it)."""
    last = None
    for i in range(tries):
        try:
            return fn()
        except Exception as e:
            last = e
            msg = str(e)
            if "No database" in msg or "no current database" in msg.lower():
                raise
            time.sleep(wait)
    raise last


def get_acad():
    if win32com is None:
        raise RuntimeError("pywin32 not available")
    acad = win32com.client.Dispatch("AutoCAD.Application")
    try:
        acad.Visible = True
    except Exception:
        pass
    return acad


def _same_file(a: Path, b) -> bool:
    try:
        return os.path.normcase(os.path.abspath(str(a))) == os.path.normcase(os.path.abspath(str(b)))
    except Exception:
        return False


def find_open_doc(acad, path: Path):
    """Return the open Document for `path`, opening it if needed. Uses ActiveDocument after
    Open() because the late-bound Open() return value does not resolve."""
    path = Path(path)
    for d in list(acad.Documents):
        try:
            if _same_file(path, d.FullName) or Path(d.FullName).name.lower() == path.name.lower():
                return d
        except Exception:
            continue
    com_retry(lambda: acad.Documents.Open(str(path)))
    time.sleep(1.5)
    return acad.ActiveDocument


def safe_set(attr_obj, value):
    """Set an attribute's TextString with retry; never retry the 'No database' error."""
    try:
        com_retry(lambda: setattr(attr_obj, "TextString", str(value)))
        return True
    except Exception as e:
        print(f"        safe_set failed ({repr(e)[:50]})")
        return False


# ----------------------------------------------------------------------------- title blocks
def _iter_paperspace_blockrefs(doc):
    """Yield (layout_name, block_ref) for every block reference in every PAPER layout."""
    for layout in doc.Layouts:
        try:
            if layout.ModelType:        # skip Model tab
                continue
        except Exception:
            if layout.Name.lower() == "model":
                continue
        block = layout.Block
        try:
            n = block.Count
        except Exception:
            continue
        for i in range(n):
            try:
                ent = block.Item(i)
            except Exception:
                continue
            try:
                if ent.ObjectName == "AcDbBlockReference" and ent.HasAttributes:
                    yield layout.Name, ent
            except Exception:
                continue


def _block_name(ent):
    for attr in ("EffectiveName", "Name"):
        try:
            v = getattr(ent, attr)
            if v:
                return str(v)
        except Exception:
            continue
    return ""


def collect_titleblocks(doc, tb_blocks):
    """Return [(layout_name, ent, attr_objs, text)] for title-block inserts matching tb_blocks.
    attr_objs: {TAG: attribute COM object}; text: {TAG: TextString}. PaperSpace only."""
    want = {b.lower() for b in tb_blocks}
    out = []
    for lname, ent in _iter_paperspace_blockrefs(doc):
        name = _block_name(ent)
        if name.lower() not in want:
            continue
        try:
            attrs = ent.GetAttributes()
        except Exception:
            continue
        attr_objs, text = {}, {}
        for a in attrs:
            try:
                tag = str(a.TagString)
                attr_objs[tag] = a
                text[tag] = str(a.TextString)
            except Exception:
                continue
        out.append((lname, ent, attr_objs, text))
    return out


def find_asbuilt_row(text):
    """Row index (1..MAX_ROWS) whose DESCRIPTION says AS BUILT, else None."""
    for n in range(1, MAX_ROWS + 1):
        desc = text.get(f"{n}DESCRIPTION", "")
        if "AS BUILT" in desc.upper().replace("-", " ").replace("_", " "):
            return n
    return None


# ----------------------------------------------------------------------------- QA / identity
def _identity_wants(cfg):
    return {
        "PROJECTNUMBER": PROJ_NO,
        "DRAWINGNUMBER": cfg["draw_no"],
        "PROJECTDESCRIPTION": PROJ_DESC,
    }


def qa_repair_identity(doc, cfg, dry_run):
    """Fill BLANK identity fields only (never overwrite). Returns (repaired, flagged)."""
    wants = _identity_wants(cfg)
    repaired = flagged = 0
    for sname, ent, attr_objs, text in collect_titleblocks(doc, cfg["tb_blocks"]):
        checks = list(wants.items()) + [("1PROJECT", PROJ_NO)]
        for tag, want in checks:
            if tag not in attr_objs:
                continue
            if text.get(tag, "").strip() != "":
                continue
            flagged += 1
            print(f"        QA: [{sname}/{_block_name(ent)}] {tag} BLANK -> {want!r}")
            if not dry_run and safe_set(attr_objs[tag], want):
                repaired += 1
                try:
                    ent.Update()
                except Exception:
                    pass
    if flagged == 0:
        print("        QA: all identity fields populated — OK")
    return repaired, flagged


# ----------------------------------------------------------------------------- rev-clean core
def plan_titleblock_edit(text, new_rev):
    """Compute the desired writes for ONE title block: relocate the AS BUILT row to row 1 with
    new_rev, clear rows 2..MAX_ROWS, set standalone REVISION. Returns dict {TAG: value}."""
    ab = find_asbuilt_row(text)
    writes = {}
    if ab is None:
        return writes, None
    # row-1 = AS BUILT row's values, with the new rev
    src = {t: text.get(f"{ab}{t}", "") for t in ROW_TAGS}
    writes["1REV"] = new_rev
    writes["1DATE"] = src.get("DATE", "")
    writes["1DESCRIPTION"] = "AS BUILT"
    writes["1DESIGNED"] = src.get("DESIGNED", "")
    writes["1DRAWN"] = src.get("DRAWN", "")
    writes["1CHECK"] = src.get("CHECK", "")
    writes["1APPROVED"] = src.get("APPROVED", "")
    writes["1PROJECT"] = src.get("PROJECT", "") or PROJ_NO
    # clear every other row
    for n in range(2, MAX_ROWS + 1):
        for t in ROW_TAGS:
            writes[f"{n}{t}"] = ""
    # standalone big REV box
    writes["REVISION"] = new_rev
    return writes, ab


def apply_titleblock_edit(doc, cfg, dry_run):
    """Apply the rev-clean to every matching title block. Returns count of TBs edited."""
    new_rev = cfg["new_rev"]
    edited = 0
    tbs = collect_titleblocks(doc, cfg["tb_blocks"])
    print(f"  title blocks found: {len(tbs)}")
    for sname, ent, attr_objs, text in tbs:
        writes, ab = plan_titleblock_edit(text, new_rev)
        if ab is None:
            print(f"   [{sname}/{_block_name(ent)}] no AS BUILT row found — skipped")
            continue
        print(f"   [{sname}/{_block_name(ent)}] AS BUILT row={ab} -> row1 REV {new_rev}, clear rows 2..{MAX_ROWS}")
        if dry_run:
            edited += 1
            continue
        n_ok = 0
        for tag, val in writes.items():
            if tag in attr_objs and attr_objs[tag].TextString != str(val):
                if safe_set(attr_objs[tag], val):
                    n_ok += 1
        try:
            ent.Update()
        except Exception:
            pass
        print(f"        wrote {n_ok} attribute(s)")
        edited += 1
    return edited


def fix_titleblock_def_statics(doc, cfg, dry_run):
    """Riverina_tellhow carries a baked-in static rev-0 row (plain AcDbText in the block DEF,
    not attributes) + a static REV digit. Blank those so the relocated AS BUILT row doesn't
    overlap them. No-op for blocks without baked statics (e.g. LMS_ACE). Returns count blanked."""
    if "riverina_tellhow" not in {b.lower() for b in cfg["tb_blocks"]}:
        return 0
    try:
        blk = doc.Blocks.Item("Riverina_tellhow")
    except Exception:
        return 0
    # find Y of a FOR-CONSTRUCTION/REVIEW static -> blank statics sharing that Y row
    constr_y = None
    statics = []
    try:
        cnt = blk.Count
    except Exception:
        return 0
    for i in range(cnt):
        try:
            e = blk.Item(i)
        except Exception:
            continue
        on = ""
        try:
            on = e.ObjectName
        except Exception:
            pass
        if on not in ("AcDbText", "AcDbMText"):
            continue
        try:
            s = str(e.TextString)
            y = float(e.InsertionPoint[1])
        except Exception:
            continue
        statics.append((e, s, y))
        if any(k in s.upper() for k in CONSTR_KEYS):
            constr_y = y
    if constr_y is None:
        return 0
    blanked = 0
    for e, s, y in statics:
        if "AS BUILT" in s.upper():
            continue  # keep the AS BUILT stamp box
        same_row = abs(y - constr_y) < 2.0
        rev_digit = (len(s.strip()) <= 2 and s.strip().isalnum() and abs(y - constr_y) < 50.0)
        if same_row or rev_digit:
            print(f"        baked static blank: {s!r} @y={y:.1f}")
            if not dry_run:
                try:
                    com_retry(lambda e=e: setattr(e, "TextString", ""))
                    blanked += 1
                except Exception as ex:
                    print(f"          failed: {repr(ex)[:40]}")
    return blanked


# ----------------------------------------------------------------------------- clouds
def cloud_layer_set(doc):
    out = set()
    try:
        for lyr in doc.Layers:
            if str(lyr.Name).strip().lower() in CLOUD_LAYER_NAMES:
                out.add(str(lyr.Name))
    except Exception:
        pass
    return out


def collect_cloud_entities(doc, cloud_layers):
    """(space_label, entity) for entities on cloud layers — PaperSpace via direct iteration."""
    if not cloud_layers:
        return []
    low = {c.lower() for c in cloud_layers}
    hits = []
    # scan each paper layout's block directly (PaperSpace only — never crawl ModelSpace)
    for layout in doc.Layouts:
        try:
            blk = layout.Block
            for i in range(blk.Count):
                e = blk.Item(i)
                try:
                    if str(e.Layer).lower() in low:
                        hits.append((layout.Name, e))
                except Exception:
                    continue
        except Exception:
            continue
    return hits


def count_cloud_entities(doc, cloud_layers):
    return len(collect_cloud_entities(doc, cloud_layers))


def erase_clouds_laydel(doc, cloud_layers):
    """Delete each cloud layer + all its contents across all spaces via native -LAYDEL."""
    for ln in sorted(cloud_layers):
        try:
            doc.SendCommand(f'_.-LAYDEL\n_Name\n{ln}\n\n_Yes\n')
            time.sleep(1.5)
            print(f"        -LAYDEL {ln}")
        except Exception as e:
            print(f"        -LAYDEL {ln} failed: {repr(e)[:40]}")


def erase_clouds_fallback(doc, cloud_layers):
    """Direct COM erase of residual cloud entities (-LAYDEL silently no-ops on exploded -exp)."""
    hits = collect_cloud_entities(doc, cloud_layers)
    if not hits:
        return 0
    print(f"        -LAYDEL left {len(hits)} residual cloud entities — direct erase")
    erased = 0
    for sname, e in hits:
        try:
            e.Delete()
            erased += 1
        except Exception:
            pass  # SelectionSet dup double-report -> 'already deleted' is harmless
    for ln in sorted(cloud_layers):
        try:
            doc.Layers.Item(ln).Delete()
        except Exception:
            pass
    try:
        doc.Regen(1)
    except Exception:
        pass
    return erased


# ----------------------------------------------------------------------------- save / archive
def save_doc(doc, label=""):
    for strat in ("Save", "QSAVE"):
        try:
            if strat == "Save":
                com_retry(lambda: doc.Save())
            else:
                com_retry(lambda: doc.SendCommand("_.QSAVE\n"))
            time.sleep(1.0)
            if label:
                print(f"  saved ({label})")
            return True
        except Exception as e:
            print(f"  save via {strat} failed: {repr(e)[:50]}")
    return False


def close_doc(doc):
    try:
        com_retry(lambda: doc.Close(True))
        time.sleep(2)
        return True
    except Exception as e:
        print(f"  close failed: {repr(e)[:60]}")
        return False


def archive_into_ab(src: Path, ab_dir: Path):
    try:
        ab_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    dest = ab_dir / src.name
    try:
        if _same_file(src, dest):
            return dest, "already in AB folder"
    except Exception:
        pass
    for attempt in range(6):
        try:
            if dest.exists():
                dest.unlink()
            shutil.move(str(src), str(dest))
            return dest, "moved into AB folder"
        except Exception as e:
            if attempt == 5:
                return src, f"move FAILED ({repr(e)[:40]}) — left in place"
            time.sleep(2.5)
    return src, "move FAILED — left in place"


# ----------------------------------------------------------------------------- PDF
def publish_pdf(doc, src: Path, pdf_path: Path):
    """Plot every paper layout to a temp PDF then merge (robust, in-our-control). Uses each
    layout's own page setup. Requires PyMuPDF (fitz) for the merge."""
    try:
        import fitz
    except Exception:
        print("  publish_pdf: PyMuPDF not available — skipping PDF")
        return False
    try:
        doc.SetVariable("BACKGROUNDPLOT", 0)
    except Exception:
        pass
    layouts = []
    for layout in doc.Layouts:
        try:
            if not layout.ModelType:
                layouts.append(layout)
        except Exception:
            if layout.Name.lower() != "model":
                layouts.append(layout)
    # AutoCAD orders layouts by TabOrder
    try:
        layouts.sort(key=lambda L: L.TabOrder)
    except Exception:
        pass
    tmp = Path(tempfile.mkdtemp(prefix="abrc_"))
    parts = []
    for idx, layout in enumerate(layouts):
        out = tmp / f"p{idx:02d}.pdf"
        try:
            com_retry(lambda: setattr(doc, "ActiveLayout", layout))
            ok = False
            try:
                ok = doc.Plot.PlotToFile(str(out))            # use layout's page setup
            except Exception:
                ok = doc.Plot.PlotToFile(str(out), "DWG To PDF.pc3")
            time.sleep(1.0)
            if out.exists():
                parts.append(out)
        except Exception as e:
            print(f"  plot layout {layout.Name} failed: {repr(e)[:50]}")
    if not parts:
        return False
    merged = fitz.open()
    for p in parts:
        try:
            merged.insert_pdf(fitz.open(str(p)))
        except Exception:
            pass
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    merged.save(str(pdf_path))
    merged.close()
    try:
        shutil.rmtree(tmp, ignore_errors=True)
    except Exception:
        pass
    return pdf_path.exists()


# ----------------------------------------------------------------------------- run one target
def run(key, dry_run=True, do_pdf=True, acad=None):
    cfg = TARGETS[key]
    src = Path(cfg["src"])
    print("=" * 90)
    print(f"TARGET {key}: {cfg['label']}   -> REV {cfg['new_rev']}")
    print(f"  src: {src}")
    if not _exists(src):
        print("  !! SOURCE NOT FOUND — skipping")
        return False
    if acad is None:
        acad = get_acad()
    doc = find_open_doc(acad, src)
    print(f"  doc: {doc.Name}")

    # ---- rev-clean the title block ----
    apply_titleblock_edit(doc, cfg, dry_run)
    # ---- baked statics (Riverina only) ----
    nb = fix_titleblock_def_statics(doc, cfg, dry_run)
    if nb:
        print(f"  baked statics blanked: {nb}")
    # ---- identity QA ----
    print("  identity-field QA + repair:")
    qa_repair_identity(doc, cfg, dry_run)

    cloud_layers = cloud_layer_set(doc) if cfg.get("remove_clouds", False) else set()
    if cfg.get("remove_clouds", False):
        print(f"  cloud layers: {sorted(cloud_layers)}")

    if dry_run:
        print("  DRY-RUN: no save, no changes written.")
        return True

    # ---- save reliable edits FIRST ----
    if not save_doc(doc, "after title-block + QA edits"):
        return False

    # ---- clouds (after save) ----
    if cloud_layers:
        print("  erasing revision clouds (-LAYDEL):")
        erase_clouds_laydel(doc, cloud_layers)
        save_doc(doc, "after cloud removal")
        if collect_cloud_entities(doc, cloud_layers):
            erase_clouds_fallback(doc, cloud_layers)
            save_doc(doc, "after fallback cloud erase")
        print(f"  cloud entities AFTER removal: {count_cloud_entities(doc, cloud_layers)}")

    # ---- publish PDF + deliver to client (PDF lives ONLY in client folder) ----
    if do_pdf:
        pdf_path = src.with_suffix(".pdf")
        ok = publish_pdf(doc, src, pdf_path)
        print(f"  PDF: {'OK' if ok else 'FAILED'} -> {pdf_path.name}")
        if ok and pdf_path.exists():
            try:
                CLIENT_PDF_DIR.mkdir(parents=True, exist_ok=True)
                client_pdf = CLIENT_PDF_DIR / cfg.get("client_name", pdf_path.name)
                shutil.copy2(str(pdf_path), str(client_pdf))
                print(f"  delivered client PDF -> {client_pdf.name}")
                pdf_path.unlink()    # don't leave a PDF in the native folder
            except Exception as e:
                print(f"  !! client PDF copy failed: {repr(e)[:50]}")

    # ---- archive deliverable native into the AS BUILT folder ----
    if close_doc(doc):
        dest, note = archive_into_ab(src, Path(cfg["ab_dir"]))
        print(f"  archive: {note} -> {dest}")
    else:
        print(f"  archive skipped (doc still open); native left at {src}")
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--target", required=True,
                    help="one of: " + ", ".join(TARGETS) + ", wave2, all")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--no-pdf", action="store_true")
    args = ap.parse_args()
    if not args.apply:
        args.dry_run = True
    dry = not args.apply

    if args.target == "wave2":
        keys = list(WAVE2)
    elif args.target == "all":
        keys = list(TARGETS)
    else:
        keys = [args.target]

    acad = get_acad()
    for k in keys:
        if k not in TARGETS:
            print(f"unknown target {k}"); continue
        try:
            run(k, dry_run=dry, do_pdf=not args.no_pdf, acad=acad)
        except Exception as e:
            print(f"!! {k} errored: {repr(e)[:80]}")
    print("=" * 90)
    print("DONE")


if __name__ == "__main__":
    main()

"""AS BUILT regression suite — consolidated, project-parametrized.

Replaces the per-scenario probes test_ab_dryrun.py + test_ab_issue_fixes.py.
One project registry, one set of checks, run any project (or `all`).

Checks (no AutoCAD needed) run by default:
  - scan        : scan_native_for_ab() finds IFC sources, single vs multi-page
  - dry_batch   : batch_convert(dry_run) plans output without converting
  - qa_existing : _qa_validate_ab_pdf() on existing AB / issue PDFs

Opt-in (needs AutoCAD COM), add --live:
  - convert     : real batch_convert() of the project's issue doc-IDs, then
                  QA the new PDFs and compare old vs new.

Usage:
  python test_ab_suite.py                # default project (lms), no-COM checks
  python test_ab_suite.py cole2          # one project, no-COM checks
  python test_ab_suite.py all            # every project, no-COM checks
  python test_ab_suite.py warnertown --live   # add real conversion (AutoCAD)
"""
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, str(Path(__file__).parent))

from ifr_automation_v10 import AsBuiltManager

_DROPBOX = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"

# Project registry. issue_doc_ids / issue_dir are optional (drive the --live phase
# and the qa_existing check on a project's known-problem PDFs).
PROJECTS = {
    "lms": {
        "path": rf"{_DROPBOX}\2.SA\LMS",
    },
    "cole2": {
        "path": rf"{_DROPBOX}\1.NSW\Coleambally #2",
    },
    "warnertown": {
        "path": rf"{_DROPBOX}\2.SA\GG-31 Warnertown BESS",
        "issue_doc_ids": [
            "GG31-E-PLN-004",   # broken export (phantom page)
            "GG31-C-PLN-001",   # print in colour, no frame
            "GG31-C-PLN-002",   # stamp overlap (double COLOUR)
            "GG31-C-PLN-005",   # stamp overlap (double COLOUR)
            "GG31-E-BLD-001",   # stamp overlap (FOR CONSTRUCTION leftover)
        ],
        "issue_dir": (
            r"Design\Engineering\1. Drawings\5. As Built"
            r"\3. As Built Client\Issue Targeted\Stamp Overlap"
        ),
    },
}


def _hdr(title):
    print("=" * 70)
    print(title)
    print("=" * 70)


# ---------------------------------------------------------------- no-COM checks
def scan(cfg):
    _hdr(f"scan — {cfg['path']}")
    mgr = AsBuiltManager(cfg["path"], dry_run=True)
    print(f"Native root: {mgr.native_root}")
    print(f"AB output:   {mgr.ab_output}")
    print(f"Existing AB PDFs: {mgr._scan_existing_ab()}\n")

    items = mgr.scan_native_for_ab()
    single = multi = 0
    for it in items:
        ifc = it["ifc_source"]
        n = len(ifc["dwg_paths"])
        is_multi = ifc["is_multi_page"]
        ab_rev = (it["existing_ab_rev"] or 0) + 1
        kind = f"MULTI ({n} pages)" if is_multi else "SINGLE"
        print(f"  {it['doc_id']:20s} {kind:18s} IFC Rev{ifc['ifc_rev']} -> AB Rev{ab_rev}")
        multi, single = (multi + 1, single) if is_multi else (multi, single + 1)
    print(f"\nSummary: {single} single + {multi} multi = {len(items)} total\n")


def dry_batch(cfg):
    _hdr(f"dry_batch — {cfg['path']}")
    mgr = AsBuiltManager(cfg["path"], dry_run=True)
    for r in mgr.batch_convert():
        status = "OK" if r["success"] else "FAIL"
        rev = r.get("ab_rev", "?")
        pdf = Path(r["pdf_path"]).name if r.get("pdf_path") else "N/A"
        print(f"  [{status}] {r['doc_id']} Rev{rev} -> {pdf}")
    print()


def qa_existing(cfg):
    """QA the project's existing issue PDFs (no AutoCAD). Confirms QA catches problems."""
    _hdr(f"qa_existing — {cfg['path']}")
    issue_dir = cfg.get("issue_dir")
    if not issue_dir:
        print("  (no issue_dir configured for this project — skipping)\n")
        return
    d = Path(cfg["path"]) / issue_dir
    pdfs = sorted(d.glob("*.pdf")) if d.exists() else []
    if not pdfs:
        print(f"  No PDFs found in {d}\n")
        return
    mgr = AsBuiltManager(cfg["path"], dry_run=True)
    for pdf in pdfs:
        doc_id = pdf.stem.split(" REV ")[0] if " REV " in pdf.stem else pdf.stem
        warnings = mgr._qa_validate_ab_pdf(pdf, doc_id)
        print(f"  QA: {pdf.name}")
        for w in (warnings or []):
            print(f"    WARNING: {w}")
        if not warnings:
            print("    PASS (0 warnings)")
    print()


# ------------------------------------------------------------- opt-in COM check
def live_convert(cfg):
    _hdr(f"live_convert (AutoCAD) — {cfg['path']}")
    doc_ids = cfg.get("issue_doc_ids")
    if not doc_ids:
        print("  (no issue_doc_ids configured — nothing to convert)\n")
        return
    mgr = AsBuiltManager(cfg["path"])
    results = mgr.batch_convert(doc_ids=doc_ids, force_doc_ids={d.upper() for d in doc_ids})

    print("\n--- QA on new outputs ---")
    all_pass = True
    qa_mgr = AsBuiltManager(cfg["path"], dry_run=True)
    for r in results:
        doc_id = r.get("doc_id", "unknown")
        pdf = r.get("pdf_path")
        if not r.get("success") or not pdf or not Path(pdf).exists():
            print(f"  {doc_id}: CONVERSION FAILED / NO PDF — {'; '.join(r.get('errors', []))}")
            all_pass = False
            continue
        warnings = qa_mgr._qa_validate_ab_pdf(Path(pdf), doc_id)
        print(f"  QA: {Path(pdf).name}")
        for w in (warnings or []):
            print(f"    WARNING: {w}")
            all_pass = False
        if not warnings:
            print("    PASS (0 warnings)")
    print("\n  " + ("ALL QA CHECKS PASSED" if all_pass else "SOME QA CHECKS FAILED — review above") + "\n")


NO_COM = [scan, dry_batch, qa_existing]


def run_project(name, live):
    cfg = PROJECTS[name]
    _hdr(f"PROJECT: {name}")
    for check in NO_COM:
        check(cfg)
    if live:
        live_convert(cfg)


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    live = "--live" in sys.argv
    target = (args[0].lower() if args else "lms")

    names = list(PROJECTS) if target == "all" else [target]
    bad = [n for n in names if n not in PROJECTS]
    if bad:
        print(f"Unknown project(s): {bad}. Known: {list(PROJECTS)} or 'all'.")
        sys.exit(2)

    for n in names:
        run_project(n, live)

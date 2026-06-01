"""Test AS BUILT conversion fixes for 5 Warnertown BESS issue doc-IDs.

Tests Fix 1-5 from 2026-05-28:
  1. _scan_has_colour two-pass (PaperSpace COLOUR detection)
  2. _remove_ifc_stamp IFC block + PaperSpace expansion
  3. _ensure_colour_has_border (missing COLOUR border)
  4. _publish_single_pdf phantom layout filtering
  5. _qa_validate_ab_pdf automated QA

Usage: python test_ab_issue_fixes.py [--dry-run]
"""
import sys
import os
import time
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from ifr_automation_v10 import AsBuiltManager

PROJECT_PATH = Path(
    r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"
    r"\2.SA\GG-31 Warnertown BESS"
)

ISSUE_DOC_IDS = [
    "GG31-E-PLN-004",   # broken export (phantom page)
    "GG31-C-PLN-001",   # print in color no frame
    "GG31-C-PLN-002",   # stamp overlap (double COLOUR)
    "GG31-C-PLN-005",   # stamp overlap (double COLOUR)
    "GG31-E-BLD-001",   # stamp overlap (FOR CONSTRUCTION leftover)
]

ISSUE_TARGETED_DIR = (
    PROJECT_PATH / "Design" / "Engineering" / "1. Drawings"
    / "5. As Built" / "3. As Built Client" / "Issue Targeted" / "Stamp Overlap"
)


def run_qa_on_existing_issues():
    """Run QA validation on existing issue PDFs (no AutoCAD needed)."""
    print("=" * 60)
    print("Phase 1: QA validation on EXISTING issue PDFs")
    print("=" * 60)

    mgr = AsBuiltManager(PROJECT_PATH, dry_run=True)

    issue_pdfs = list(ISSUE_TARGETED_DIR.glob("*.pdf"))
    if not issue_pdfs:
        print(f"  No PDFs found in {ISSUE_TARGETED_DIR}")
        return {}

    qa_results = {}
    for pdf_path in sorted(issue_pdfs):
        doc_id = pdf_path.stem.split(" REV ")[0] if " REV " in pdf_path.stem else pdf_path.stem
        print(f"\
  QA: {pdf_path.name}")
        warnings = mgr._qa_validate_ab_pdf(pdf_path, doc_id)
        qa_results[doc_id] = warnings
        if warnings:
            for w in warnings:
                print(f"    WARNING: {w}")
        else:
            print(f"    PASS (0 warnings)")

    return qa_results


def run_conversion_test(dry_run=False):
    """Run actual AS BUILT conversion on issue doc-IDs."""
    print("\
" + "=" * 60)
    print(f"Phase 2: AS BUILT conversion ({'DRY RUN' if dry_run else 'LIVE'})")
    print("=" * 60)

    if dry_run:
        print("  Dry run — skipping actual conversion")
        return []

    mgr = AsBuiltManager(PROJECT_PATH)
    force_set = {d.upper() for d in ISSUE_DOC_IDS}

    results = mgr.batch_convert(
        doc_ids=ISSUE_DOC_IDS,
        force_doc_ids=force_set
    )

    return results


def run_qa_on_new_outputs(results):
    """Run QA on newly generated PDFs."""
    print("\
" + "=" * 60)
    print("Phase 3: QA validation on NEW outputs")
    print("=" * 60)

    if not results:
        print("  No results to validate")
        return

    mgr = AsBuiltManager(PROJECT_PATH, dry_run=True)
    all_pass = True

    for r in results:
        doc_id = r.get('doc_id', 'unknown')
        pdf_path = r.get('pdf_path')
        success = r.get('success', False)

        if not success:
            print(f"\
  {doc_id}: CONVERSION FAILED — {'; '.join(r.get('errors', []))}")
            all_pass = False
            continue

        if not pdf_path or not Path(pdf_path).exists():
            print(f"\
  {doc_id}: NO PDF OUTPUT")
            all_pass = False
            continue

        print(f"\
  QA: {Path(pdf_path).name}")
        warnings = mgr._qa_validate_ab_pdf(Path(pdf_path), doc_id)

        if r.get('qa_warnings'):
            print(f"    (inline QA already caught: {r['qa_warnings']})")

        if warnings:
            for w in warnings:
                print(f"    WARNING: {w}")
            all_pass = False
        else:
            print(f"    PASS (0 warnings)")

    print("\
" + "=" * 60)
    if all_pass:
        print("ALL QA CHECKS PASSED")
    else:
        print("SOME QA CHECKS FAILED — review warnings above")
    print("=" * 60)


def compare_old_vs_new():
    """Compare old issue PDFs vs new outputs (page count, file size)."""
    print("\
" + "=" * 60)
    print("Phase 4: Old vs New comparison")
    print("=" * 60)

    ab_output = (
        PROJECT_PATH / "Design" / "Engineering" / "1. Drawings"
        / "5. As Built" / "3. As Built Client"
    )

    for doc_id in ISSUE_DOC_IDS:
        short_id = doc_id.replace("GG31-", "")
        old_pdfs = list(ISSUE_TARGETED_DIR.glob(f"*{short_id}*"))
        new_pdfs = list(ab_output.glob(f"*{short_id}*AS BUILT*"))

        old_pdf = old_pdfs[0] if old_pdfs else None
        new_pdf = new_pdfs[0] if new_pdfs else None

        print(f"\
  {doc_id}:")
        if old_pdf:
            print(f"    OLD: {old_pdf.name} ({old_pdf.stat().st_size:,} bytes)")
        else:
            print(f"    OLD: (not in Issue Targeted)")
        if new_pdf:
            print(f"    NEW: {new_pdf.name} ({new_pdf.stat().st_size:,} bytes)")
        else:
            print(f"    NEW: (not generated yet)")


if __name__ == "__main__":
    dry_run = "--dry-run" in sys.argv

    print("AS BUILT Issue Fix Verification Test")
    print(f"Project: {PROJECT_PATH}")
    print(f"Issue doc-IDs: {ISSUE_DOC_IDS}")
    print()

    # Phase 1: QA on old issue PDFs (confirms QA catches the problems)
    qa_old = run_qa_on_existing_issues()

    if dry_run:
        print("\
  --dry-run: skipping conversion. Remove --dry-run to run full test.")
        sys.exit(0)

    # Phase 2: Run actual conversion
    results = run_conversion_test(dry_run=False)

    # Phase 3: QA on new outputs
    if results:
        run_qa_on_new_outputs(results)

    # Phase 4: Compare
    compare_old_vs_new()


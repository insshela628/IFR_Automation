"""Test _mirror_reports with Excel doc-ID fallback.

Validates:
1. LMS: files with prefix mismatch but valid Excel doc-ID → collected
2. LMS: third-party files (not in Excel) → skipped
3. Warnertown: zero behavior change (old == new)
4. Full _mirror_reports dry-run on both projects
"""
import sys, io, re
from pathlib import Path
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, str(Path(__file__).parent))

from ifr_automation_v10 import (
    _extract_doc_id_standalone, DeliverableManager, IFRAutomation,
    ConfigManager, ProjectValidator, ProjectValidation,
)

DROPBOX = Path("C:/Users/ACE/GREEN GOLD ENERGY Dropbox/Projects/Project (EPC)")
LMS_PATH = DROPBOX / "2.SA/LMS"
WTN_PATH = DROPBOX / "2.SA/GG-31 Warnertown BESS"

passed = 0
failed = 0


def check(label, condition, detail=""):
    global passed, failed
    if condition:
        passed += 1
        print(f"  ✓ {label}")
    else:
        failed += 1
        print(f"  ✗ {label}  {detail}")


def _collect_from_source(auto, source_dir, excel_doc_ids):
    """Simulate _mirror_reports collection phase for a single source dir."""
    results = []
    if not source_dir.exists():
        return results
    for pdf_file in source_dir.rglob("*.pdf"):
        if pdf_file.name.startswith("~$"):
            continue
        try:
            rel_parts = pdf_file.relative_to(source_dir).parts[:-1]
            in_ss = any(
                p.lower() in ('ss', 'superseded', 'superceded', '_export', 'data')
                for p in rel_parts
            )
            if in_ss:
                in_excluded = True
            else:
                in_appendix = any(
                    p.lower().startswith(('appendix', 'stk'))
                    for p in rel_parts
                )
                if in_appendix:
                    _doc_id_pat = re.match(
                        r'^(GG\d{2}-[A-Z]-[A-Z]{3}-\d{3})',
                        pdf_file.name, re.IGNORECASE
                    )
                    in_excluded = not bool(_doc_id_pat)
                else:
                    in_excluded = False
        except Exception:
            in_excluded = False

        ok = auto._should_collect_as_report(
            pdf_file.name, True, in_excluded, rel_parts, excel_doc_ids
        )
        if ok:
            results.append(pdf_file.name)
    return sorted(results)


def _collect_old(auto, source_dir):
    """Old behavior: no Excel fallback."""
    return _collect_from_source(auto, source_dir, None)


def _collect_new(auto, source_dir, excel_doc_ids):
    """New behavior: with Excel fallback."""
    return _collect_from_source(auto, source_dir, excel_doc_ids)


# =========================================================================
# Test 1: LMS — Excel doc-ID preload
# =========================================================================
print("\n[Test 1] LMS — Excel doc-ID preload")
dm_lms = DeliverableManager(LMS_PATH, dry_run=True)
lms_ids = dm_lms.preload_doc_ids()
check("preload returns non-empty set", len(lms_ids) > 0, f"got {len(lms_ids)}")
check("50023-RE-305 in Excel", "50023-RE-305" in lms_ids)
check("50023-KE-301 in Excel", "50023-KE-301" in lms_ids)
check("50023-LE-301 in Excel", "50023-LE-301" in lms_ids)
check("50023-LE-304 in Excel", "50023-LE-304" in lms_ids)
# Third-party should NOT be in Excel
check("17111-REP-001 NOT in Excel", "17111-REP-001" not in lms_ids)
check("17190-REP-001 NOT in Excel", "17190-REP-001" not in lms_ids)


# =========================================================================
# Test 2: LMS — Phase 4b collects prefix-mismatched deliverables
# =========================================================================
print("\n[Test 2] LMS — Phase 4b Excel fallback")
config = ConfigManager()
auto_lms = IFRAutomation(
    root_path=str(LMS_PATH.parent), config=config, dry_run=True,
    excel_doc_ids=lms_ids,
)
lms_elec = LMS_PATH / "Design/Engineering/2. Calcs & Reports/Reports/Electrical"

old_lms = _collect_old(auto_lms, lms_elec)
new_lms = _collect_new(auto_lms, lms_elec, lms_ids)

added = set(new_lms) - set(old_lms)
removed = set(old_lms) - set(new_lms)

check("new collects MORE than old", len(new_lms) > len(old_lms),
      f"old={len(old_lms)} new={len(new_lms)}")
check("nothing REMOVED from old", len(removed) == 0,
      f"removed: {removed}")

# Specific files that should now be collected
expected_new = [
    "50023-RE-305-Earthing Design Report_Rev0.pdf",
    "50023-LE-301 Rev 2 - BESS RMU BESS Feeder Protection Cause and Effect.pdf",
    "50023-LE-304 Rev 0 - BESS Inverter Settings.pdf",
]
for fn in expected_new:
    check(f"NEW: {fn[:50]}...", fn in new_lms)

# Third-party files should NOT be collected even with Excel
third_party = [
    "17111-REP-001r1 - NAWMA Power Station HV Earthing Report_LMS EDIT 4May2017.pdf",
    "17190-REP-001r2 - NAWMA Solar Installation Protection Report.pdf",
    "Thermal - Conductivity test report January 2024.pdf",
]
for fn in third_party:
    check(f"SKIP: {fn[:50]}...", fn not in new_lms)

print(f"\n  LMS collection: {len(old_lms)} → {len(new_lms)} (+{len(added)})")
if added:
    for a in sorted(added):
        print(f"    + {a}")


# =========================================================================
# Test 3: Warnertown — zero behavior change
# =========================================================================
print("\n[Test 3] Warnertown — zero behavior change")
dm_wtn = DeliverableManager(WTN_PATH, dry_run=True)
wtn_ids = dm_wtn.preload_doc_ids()
check("Warnertown preload works", len(wtn_ids) > 0, f"got {len(wtn_ids)}")

auto_wtn = IFRAutomation(
    root_path=str(WTN_PATH.parent), config=config, dry_run=True,
    excel_doc_ids=wtn_ids,
)
wtn_elec = WTN_PATH / "Design/Engineering/2. Calcs & Reports/Reports/Electrical"

old_wtn = _collect_old(auto_wtn, wtn_elec)
new_wtn = _collect_new(auto_wtn, wtn_elec, wtn_ids)

check("Warnertown IDENTICAL", old_wtn == new_wtn,
      f"old={len(old_wtn)} new={len(new_wtn)}")

if old_wtn != new_wtn:
    wtn_added = set(new_wtn) - set(old_wtn)
    wtn_removed = set(old_wtn) - set(new_wtn)
    if wtn_added:
        print(f"  ADDED: {wtn_added}")
    if wtn_removed:
        print(f"  REMOVED: {wtn_removed}")


# =========================================================================
# Test 4: Full _mirror_reports dry-run (LMS)
# =========================================================================
print("\n[Test 4] LMS — full _mirror_reports dry-run")
auto_lms_full = IFRAutomation(
    root_path=str(LMS_PATH.parent), config=config, dry_run=True,
    excel_doc_ids=lms_ids,
)
result = auto_lms_full._mirror_reports(LMS_PATH)
check("_mirror_reports returns dict", isinstance(result, dict))
check("has 'copied' key", 'copied' in result)
check("has 'skipped' key", 'skipped' in result)
total = result['copied'] + result['skipped']
check("total > 0", total > 0, f"copied={result['copied']} skipped={result['skipped']}")
print(f"  Mirror result: copied={result['copied']}, skipped={result['skipped']}")


# =========================================================================
# Test 5: Full _mirror_reports dry-run (Warnertown)
# =========================================================================
print("\n[Test 5] Warnertown — full _mirror_reports dry-run")
auto_wtn_full = IFRAutomation(
    root_path=str(WTN_PATH.parent), config=config, dry_run=True,
    excel_doc_ids=wtn_ids,
)
result_wtn = auto_wtn_full._mirror_reports(WTN_PATH)
check("_mirror_reports returns dict", isinstance(result_wtn, dict))
total_wtn = result_wtn['copied'] + result_wtn['skipped']
check("total > 0", total_wtn > 0,
      f"copied={result_wtn['copied']} skipped={result_wtn['skipped']}")
print(f"  Mirror result: copied={result_wtn['copied']}, skipped={result_wtn['skipped']}")


# =========================================================================
# Test 6: No Excel → graceful fallback (same as old behavior)
# =========================================================================
print("\n[Test 6] No Excel → graceful fallback")
auto_no_excel = IFRAutomation(
    root_path=str(LMS_PATH.parent), config=config, dry_run=True,
    # no excel_doc_ids → defaults to empty set
)
no_excel_result = _collect_from_source(auto_no_excel, lms_elec, set())
check("no Excel = same as old behavior", sorted(no_excel_result) == sorted(old_lms),
      f"no_excel={len(no_excel_result)} old={len(old_lms)}")


# =========================================================================
# Summary
# =========================================================================
print(f"\n{'='*60}")
print(f"  Results: {passed} passed, {failed} failed")
print(f"{'='*60}")
sys.exit(1 if failed else 0)

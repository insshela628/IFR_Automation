#!/usr/bin/env python3
"""
Full dry-run test of /ifr pipeline for LMS + Warnertown (control).

Tests:
  1. Excel doc-ID preload (Phase 4b support)
  2. _should_collect_as_report() old vs new behavior comparison
  3. Full pipeline dry-run via PipelineOrchestrator
  4. Warnertown as control group (zero behavior change)

Output:
  - test_output/ifr_dryrun_lms_YYYYMMDD.log   (detailed)
  - test_output/ifr_dryrun_summary.txt          (key numbers)
  - stdout summary
"""

import sys
import os
import io
import re
import logging
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from contextlib import redirect_stdout

# Fix Unicode output on Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Add script directory to path for imports
SCRIPT_DIR = Path(r"D:\1. SOP\SOP_Stage 2 IFR Sync√\V6√")
sys.path.insert(0, str(SCRIPT_DIR))

from ifr_automation_v10 import (
    ConfigManager, IFRAutomation, DeliverableManager,
    PipelineOrchestrator, ProjectValidator, ProjectValidation,
    _extract_doc_id_standalone
)

# ===========================================================================
# Constants
# ===========================================================================
DROPBOX_ROOT = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)")
LMS_PATH = DROPBOX_ROOT / "2.SA" / "LMS"
WARNERTOWN_PATH = DROPBOX_ROOT / "2.SA" / "GG-31 Warnertown BESS"
OUTPUT_DIR = SCRIPT_DIR / "test_output"
TODAY = datetime.now().strftime("%Y%m%d")
LOG_FILE = OUTPUT_DIR / f"ifr_dryrun_lms_{TODAY}.log"
SUMMARY_FILE = OUTPUT_DIR / "ifr_dryrun_summary.txt"


class DualWriter:
    """Write to both a file and an in-memory buffer."""
    def __init__(self, filepath):
        self.file = open(filepath, 'w', encoding='utf-8')
        self.buffer = io.StringIO()

    def write(self, text):
        self.file.write(text)
        self.buffer.write(text)

    def flush(self):
        self.file.flush()
        self.buffer.flush()

    def close(self):
        self.file.close()

    def getvalue(self):
        return self.buffer.getvalue()


def log_section(log, title, char="="):
    line = char * 72
    log.write(f"\n{line}\n  {title}\n{line}\n")


def test_excel_preload(log, project_path, project_name):
    """Test DeliverableManager.preload_doc_ids()."""
    log_section(log, f"TEST 1: Excel doc-ID preload — {project_name}")
    try:
        dm = DeliverableManager(project_path, dry_run=True)
        excel_path = dm.find_deliverable_excel()
        log.write(f"  Excel path: {excel_path}\n")

        doc_ids = dm.preload_doc_ids()
        log.write(f"  Total doc-IDs loaded: {len(doc_ids)}\n")

        if doc_ids:
            samples = sorted(doc_ids)[:20]
            log.write(f"  Sample doc-IDs (first 20):\n")
            for did in samples:
                log.write(f"    {did}\n")
        return doc_ids
    except Exception as e:
        log.write(f"  ERROR: {e}\n")
        import traceback
        log.write(traceback.format_exc())
        return set()


def test_report_collection(log, project_path, project_name, excel_doc_ids):
    """Test _should_collect_as_report for every PDF in Reports/Electrical.

    Runs twice: once WITHOUT excel_doc_ids (old behavior), once WITH (new).
    Returns (old_collected, new_collected, diff).
    """
    log_section(log, f"TEST 2: Report collection comparison — {project_name}")

    config = ConfigManager()
    # Create two IFRAutomation instances: old (no excel_doc_ids) and new (with)
    automation_old = IFRAutomation(
        root_path=str(project_path.parent),
        config=config,
        dry_run=True,
        interactive=True,
        excel_doc_ids=set()  # old behavior
    )
    automation_new = IFRAutomation(
        root_path=str(project_path.parent),
        config=config,
        dry_run=True,
        interactive=True,
        excel_doc_ids=excel_doc_ids  # new behavior
    )

    # Scan all report source directories
    report_sources = [
        project_path / "Design/Engineering/2. Calcs & Reports/Reports/Electrical",
        project_path / "Design/Engineering/2. Calcs & Reports/Reports/Civil & Structure",
    ]

    old_collected = []
    new_collected = []
    diff_files = []  # files collected by new but not old
    total_files = 0
    skip_reasons = defaultdict(list)

    for source_dir in report_sources:
        if not source_dir.exists():
            log.write(f"\n  Source dir not found: {source_dir}\n")
            continue

        source_label = source_dir.name
        log.write(f"\n  --- Source: {source_label} ---\n")
        is_report_source = True

        try:
            pdf_files = sorted(source_dir.rglob("*.pdf"))
        except Exception as e:
            log.write(f"  ERROR scanning: {e}\n")
            continue

        for pdf_file in pdf_files:
            total_files += 1
            try:
                rel_parts = pdf_file.relative_to(source_dir).parts[:-1]
                in_ss = any(
                    part.lower() in ('ss', 'superseded', '_export', 'data')
                    for part in rel_parts
                )
                if in_ss:
                    in_excluded_subdir = True
                else:
                    in_appendix = any(
                        part.lower().startswith(('appendix', 'stk'))
                        for part in rel_parts
                    )
                    if in_appendix:
                        _doc_id_pat = re.match(
                            r'^(GG\d{2}-[A-Z]-[A-Z]{3}-\d{3})',
                            pdf_file.name, re.IGNORECASE
                        )
                        in_excluded_subdir = not bool(_doc_id_pat)
                    else:
                        in_excluded_subdir = False
            except Exception:
                rel_parts = ()
                in_excluded_subdir = False

            # Old behavior (no excel fallback)
            old_result = automation_old._should_collect_as_report(
                pdf_file.name, is_report_source,
                in_excluded_subdir, rel_parts,
                set()  # empty — old behavior
            )
            # New behavior (with excel fallback)
            new_result = automation_new._should_collect_as_report(
                pdf_file.name, is_report_source,
                in_excluded_subdir, rel_parts,
                excel_doc_ids
            )

            # Determine relative path for readable logging
            try:
                rel_display = str(pdf_file.relative_to(source_dir))
            except ValueError:
                rel_display = pdf_file.name

            doc_id = _extract_doc_id_standalone(pdf_file.name)

            old_status = "COLLECT" if old_result else "SKIP"
            new_status = "COLLECT" if new_result else "SKIP"
            changed = " *** NEW ***" if new_result and not old_result else ""

            if old_result:
                old_collected.append(pdf_file)
            if new_result:
                new_collected.append(pdf_file)
            if new_result and not old_result:
                diff_files.append(pdf_file)
                in_excel = doc_id and doc_id.upper() in excel_doc_ids
                reason = f"doc-ID={doc_id}, in_excel={in_excel}"
                skip_reasons["newly_collected"].append((rel_display, reason))

            # Determine skip reason for logging
            if not new_result:
                if in_excluded_subdir:
                    skip_reason = "excluded_subdir"
                elif not is_report_source:
                    skip_reason = "not_report_source"
                elif rel_parts:
                    parent_name = rel_parts[0]
                    if '_' in parent_name:
                        parent_prefix = parent_name.split('_')[0]
                        if not pdf_file.name.startswith(parent_prefix):
                            in_excel = doc_id and doc_id.upper() in excel_doc_ids
                            skip_reason = f"prefix_mismatch(parent={parent_prefix}, doc_id={doc_id}, in_excel={in_excel})"
                        else:
                            skip_reason = "unknown"
                    else:
                        skip_reason = "non_standard_folder"
                else:
                    skip_reason = "unknown"
            else:
                skip_reason = ""

            log.write(f"    [{old_status}->{new_status}]{changed}  {rel_display}")
            if skip_reason:
                log.write(f"  ({skip_reason})")
            log.write("\n")

    log.write(f"\n  --- Summary for {project_name} ---\n")
    log.write(f"  Total PDFs scanned: {total_files}\n")
    log.write(f"  Old behavior collected: {len(old_collected)}\n")
    log.write(f"  New behavior collected: {len(new_collected)}\n")
    log.write(f"  Newly collected (Phase 4b): {len(diff_files)}\n")

    if diff_files:
        log.write(f"\n  Files NOW collected (were previously skipped):\n")
        for f in diff_files:
            doc_id = _extract_doc_id_standalone(f.name)
            in_excel = doc_id and doc_id.upper() in excel_doc_ids
            log.write(f"    + {f.name}  (doc-ID={doc_id}, in_excel={in_excel})\n")

    return old_collected, new_collected, diff_files


def test_pipeline_dryrun(log, project_path, project_name, region="2.SA"):
    """Run PipelineOrchestrator.run_pipeline() in dry-run mode."""
    log_section(log, f"TEST 3: Full pipeline dry-run — {project_name}")

    config = ConfigManager()
    validator = ProjectValidator(config)

    log.write(f"  Validating project: {project_path}\n")
    validation = validator.validate_project(project_path, region=region)
    log.write(f"  Valid: {validation.valid}\n")
    log.write(f"  Drawings: {validation.drawings_count}\n")
    log.write(f"  Reports: {validation.reports_count}\n")
    log.write(f"  IFR(Client) exists: {validation.has_existing_ifr_client}\n")
    log.write(f"  IFC(Client) exists: {validation.has_existing_ifc_client}\n")

    if not validation.valid:
        log.write(f"  WARNING: project not valid, recommended={validation.recommended_action}\n")
        log.write(f"  Message: {validation.warning_message}\n")
        # Still try to run pipeline — it may proceed with warnings

    orchestrator = PipelineOrchestrator(config, dry_run=True)

    # Capture stdout from the pipeline run
    pipeline_output = io.StringIO()
    result = None
    error = None

    old_stdout = sys.stdout
    sys.stdout = io.TextIOWrapper(
        io.BytesIO(), encoding='utf-8', write_through=True
    )
    # Actually, simpler: redirect with a StringIO
    sys.stdout = pipeline_output

    try:
        result = orchestrator.run_pipeline(project_path, validation)
    except Exception as e:
        error = e
        import traceback
        pipeline_output.write(f"\nEXCEPTION: {e}\n{traceback.format_exc()}\n")
    finally:
        sys.stdout = old_stdout

    captured = pipeline_output.getvalue()
    log.write(f"\n  --- Pipeline stdout ---\n")
    for line in captured.splitlines():
        log.write(f"  | {line}\n")

    if result:
        log.write(f"\n  --- Pipeline result ---\n")
        log.write(f"  Success: {result.get('success')}\n")
        for stage in ['ifr_sync', 'version_mgmt', 'sharepoint_sync', 'deliverable']:
            stage_result = result.get(stage)
            log.write(f"  {stage}: {stage_result}\n")
    elif error:
        log.write(f"\n  Pipeline FAILED with exception: {error}\n")

    return result


def main():
    print(f"IFR Pipeline Dry-Run Test — {TODAY}")
    print(f"Log: {LOG_FILE}")
    print()

    log = DualWriter(str(LOG_FILE))
    log.write(f"IFR Pipeline Dry-Run Test\n")
    log.write(f"Date: {datetime.now().isoformat()}\n")
    log.write(f"Script: {__file__}\n")
    log.write(f"ifr_automation: {SCRIPT_DIR / 'ifr_automation_v10.py'}\n\n")

    summary_lines = []
    pass_count = 0
    fail_count = 0

    # ===================================================================
    # LMS Tests
    # ===================================================================
    log_section(log, "LMS PROJECT", "=")
    log.write(f"  Path: {LMS_PATH}\n")
    log.write(f"  Exists: {LMS_PATH.exists()}\n")

    if not LMS_PATH.exists():
        log.write("  FATAL: LMS project path does not exist!\n")
        summary_lines.append("LMS: FATAL — path not found")
        fail_count += 1
    else:
        # Test 1: Excel preload
        lms_doc_ids = test_excel_preload(log, LMS_PATH, "LMS")
        if lms_doc_ids:
            summary_lines.append(f"LMS Excel preload: {len(lms_doc_ids)} doc-IDs loaded — PASS")
            pass_count += 1
        else:
            summary_lines.append("LMS Excel preload: 0 doc-IDs — FAIL (or no Excel)")
            fail_count += 1

        # Test 2: Report collection comparison
        old_collected, new_collected, diff_files = test_report_collection(
            log, LMS_PATH, "LMS", lms_doc_ids
        )

        if len(diff_files) > 0:
            summary_lines.append(
                f"LMS Phase 4b: {len(diff_files)} newly collected files "
                f"(old={len(old_collected)}, new={len(new_collected)}) — PASS (improvement detected)"
            )
            pass_count += 1
        elif len(old_collected) == len(new_collected):
            summary_lines.append(
                f"LMS Phase 4b: no change (old={len(old_collected)}, new={len(new_collected)}) — "
                f"INFO (no prefix-mismatched files with Excel doc-IDs)"
            )
            pass_count += 1  # still a pass — no regression
        else:
            summary_lines.append(
                f"LMS Phase 4b: UNEXPECTED (old={len(old_collected)}, new={len(new_collected)}) — CHECK"
            )
            fail_count += 1

        # Test 3: Full pipeline dry-run
        lms_result = test_pipeline_dryrun(log, LMS_PATH, "LMS")
        if lms_result and lms_result.get('success'):
            summary_lines.append("LMS full pipeline dry-run: SUCCESS — PASS")
            pass_count += 1

            # Extract stage details
            ifr = lms_result.get('ifr_sync', {})
            if isinstance(ifr, dict):
                summary_lines.append(
                    f"  IFR sync: drawings={ifr.get('drawings','-')}, "
                    f"reports={ifr.get('reports','-')}"
                )
            vm = lms_result.get('version_mgmt', {})
            if isinstance(vm, dict):
                summary_lines.append(
                    f"  Version mgmt: scanned={vm.get('scanned','-')}, "
                    f"moved={vm.get('moved','-')}"
                )
        elif lms_result:
            summary_lines.append(f"LMS full pipeline dry-run: result={lms_result.get('success')} — CHECK")
            fail_count += 1
        else:
            summary_lines.append("LMS full pipeline dry-run: EXCEPTION — FAIL")
            fail_count += 1

    # ===================================================================
    # Warnertown Tests (control group)
    # ===================================================================
    log_section(log, "WARNERTOWN PROJECT (CONTROL GROUP)", "=")
    log.write(f"  Path: {WARNERTOWN_PATH}\n")
    log.write(f"  Exists: {WARNERTOWN_PATH.exists()}\n")

    if not WARNERTOWN_PATH.exists():
        log.write("  FATAL: Warnertown project path does not exist!\n")
        summary_lines.append("Warnertown: FATAL — path not found")
        fail_count += 1
    else:
        # Test 1: Excel preload
        wt_doc_ids = test_excel_preload(log, WARNERTOWN_PATH, "Warnertown")
        if wt_doc_ids:
            summary_lines.append(f"Warnertown Excel preload: {len(wt_doc_ids)} doc-IDs — PASS")
            pass_count += 1
        else:
            summary_lines.append("Warnertown Excel preload: 0 doc-IDs — WARN")
            # Not a fail — may not have deliverable Excel
            pass_count += 1

        # Test 2: Report collection — should show ZERO diff (no behavior change)
        wt_old, wt_new, wt_diff = test_report_collection(
            log, WARNERTOWN_PATH, "Warnertown", wt_doc_ids
        )

        if len(wt_diff) == 0 and len(wt_old) == len(wt_new):
            summary_lines.append(
                f"Warnertown Phase 4b: IDENTICAL (old={len(wt_old)}, new={len(wt_new)}) — PASS"
            )
            pass_count += 1
        else:
            summary_lines.append(
                f"Warnertown Phase 4b: CHANGED (old={len(wt_old)}, new={len(wt_new)}, "
                f"diff={len(wt_diff)}) — FAIL (should be zero diff!)"
            )
            fail_count += 1

        # Test 3: Full pipeline dry-run
        wt_result = test_pipeline_dryrun(log, WARNERTOWN_PATH, "Warnertown")
        if wt_result and wt_result.get('success'):
            summary_lines.append("Warnertown full pipeline dry-run: SUCCESS — PASS")
            pass_count += 1
        elif wt_result:
            summary_lines.append(f"Warnertown pipeline: result={wt_result.get('success')} — CHECK")
            fail_count += 1
        else:
            summary_lines.append("Warnertown full pipeline dry-run: EXCEPTION — FAIL")
            fail_count += 1

    # ===================================================================
    # Final Summary
    # ===================================================================
    log_section(log, "FINAL SUMMARY")
    total = pass_count + fail_count
    log.write(f"  PASS: {pass_count}/{total}\n")
    log.write(f"  FAIL: {fail_count}/{total}\n\n")
    for line in summary_lines:
        log.write(f"  {line}\n")

    log.close()

    # Write summary file
    with open(str(SUMMARY_FILE), 'w', encoding='utf-8') as f:
        f.write(f"IFR Pipeline Dry-Run Summary — {TODAY}\n")
        f.write(f"{'=' * 50}\n")
        f.write(f"PASS: {pass_count}/{total}  |  FAIL: {fail_count}/{total}\n\n")
        for line in summary_lines:
            f.write(f"{line}\n")

    # Print to stdout
    print("=" * 60)
    print(f"  IFR Pipeline Dry-Run Summary — {TODAY}")
    print("=" * 60)
    print(f"  PASS: {pass_count}/{total}  |  FAIL: {fail_count}/{total}")
    print()
    for line in summary_lines:
        print(f"  {line}")
    print()
    print(f"  Detailed log: {LOG_FILE}")
    print(f"  Summary file: {SUMMARY_FILE}")
    print("=" * 60)

    return 0 if fail_count == 0 else 1


if __name__ == "__main__":
    sys.exit(main())

"""Dry-run test for PanelIFCManager.
Tests scan_and_group, IFC rev detection, and batch_convert_panel (dry-run).
Does NOT require AutoCAD.
"""
import sys
import os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(__file__))

from pathlib import Path
from ifr_automation_v10 import PanelIFCManager

SOURCE = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"
              r"\7.New Zealand\Tatua Solar Farm\1. Drawings\1. Native"
              r"\STS1&2 Panel Design& SLD")

def sep(title):
    print(f"\n{'='*70}")
    print(f"  {title}")
    print(f"{'='*70}")

# --- Test 1: Scan and Group ---
sep("Test 1: scan_and_group()")
mgr = PanelIFCManager(SOURCE, dry_run=True)
groups = mgr.scan_and_group()
print(f"  Found {len(groups)} document groups:\n")
for doc_no, pages in groups.items():
    print(f"  [{doc_no}] {len(pages)} pages:")
    for p in pages:
        page_type = ""
        if p['page'] == '00':
            page_type = " (Cover)"
        elif p['page'].startswith('0') and not p['page'].isdigit():
            page_type = " (TOC/Appendix)"
        print(f"    page={p['page']:4s} {p['filename']}{page_type}")

# --- Test 2: IFC output detection ---
sep("Test 2: IFC output path detection")
print(f"  IFC output: {mgr.ifc_output}")
print(f"  Exists: {mgr.ifc_output.exists()}")

# --- Test 3: Existing IFC rev detection ---
sep("Test 3: existing IFC rev detection")
for doc_no in groups:
    rev = mgr._get_existing_ifc_rev(doc_no)
    print(f"  {doc_no}: next IFC Rev = {rev}")

# --- Test 4: Batch convert dry-run ---
sep("Test 4: batch_convert_panel() dry-run")
results = mgr.batch_convert_panel()

sep("Test 5: Results summary")
for r in results:
    pdf_name = Path(r['pdf_result']['pdf_path']).name if r.get('pdf_result') else 'N/A'
    print(f"  {r['doc_no']}: {r['pages']} pages, IFC Rev{r['ifc_rev']}")
    print(f"    PDF -> {pdf_name}")
    print(f"    DWGs updated: {r['dwgs_updated']}, failed: {r['dwgs_failed']}")

sep("All tests complete")

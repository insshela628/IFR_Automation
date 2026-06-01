"""Dry-run test for IFC automation pipeline.
Tests scan_native_folders, convert_to_ifc (dry-run), scan_ifc_folder,
and update_ifc_deliverables logic WITHOUT requiring AutoCAD.
"""
import sys
import os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(__file__))

from pathlib import Path
from ifr_automation_v10 import IFCManager, DeliverableManager, ConfigManager

PROJECT = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS")

def sep(title):
    print(f"\n{'='*70}")
    print(f"  {title}")
    print(f"{'='*70}")

# --- Test 1: Scan Native folders ---
sep("Test 1: scan_native_folders()")
mgr = IFCManager(PROJECT, dry_run=True)
scan = mgr.scan_native_folders()
print(f"  Found {len(scan)} doc-IDs with IFR DWGs:\n")
for info in scan:
    ifc_status = f"existing IFC Rev{info['existing_ifc_rev']}" if info['existing_ifc_rev'] is not None else "no IFC yet"
    print(f"  {info['doc_id']:30s} IFR=Rev{info['latest_ifr_rev']:3s}  {ifc_status}")
    print(f"    DWG: {info['latest_ifr_dwg'].name}")

# --- Test 2: Dry-run convert_to_ifc ---
sep("Test 2: convert_to_ifc() dry-run")
results = []
for info in scan[:5]:  # First 5 only
    r = mgr.convert_to_ifc(info)
    results.append(r)
    print(f"  {r['doc_id']:30s} -> IFC Rev{r['ifc_rev']}")
    print(f"    DWG -> {Path(r['dwg_path']).name if r['dwg_path'] else 'N/A'}")
    print(f"    PDF -> {Path(r['pdf_path']).name if r['pdf_path'] else 'N/A'}")
    print(f"    Success: {r['success']}, Errors: {r['errors']}")

# --- Test 3: scan_ifc_folder (DeliverableManager) ---
sep("Test 3: scan_ifc_folder() - existing IFC PDFs")
dm = DeliverableManager(PROJECT, dry_run=True)
ifc_info = dm.scan_ifc_folder()
print(f"  Found {len(ifc_info)} doc-IDs in IFC(Client):\n")
for doc_id, info in sorted(ifc_info.items()):
    print(f"  {doc_id:30s} Rev{info['revision']}  <- {info['filename']}")

# --- Test 4: Deliverable Excel scan ---
sep("Test 4: find_deliverable_excel()")
excel_path = dm.find_deliverable_excel()
if excel_path:
    print(f"  Found: {excel_path.name}")
    print(f"  Path:  {excel_path}")
else:
    print("  NOT FOUND")

# --- Test 5: Simulate update_ifc_deliverables ---
sep("Test 5: update_ifc_deliverables() dry-run simulation")
# Build fake conversion results matching real IFC PDFs
fake_results = []
for doc_id, info in ifc_info.items():
    fake_results.append({
        'success': True,
        'doc_id': doc_id,
        'ifc_rev': info['revision'],
        'dwg_path': 'fake.dwg',
        'pdf_path': 'fake.pdf',
        'errors': [],
    })

if fake_results and excel_path:
    print(f"  Simulating update for {len(fake_results)} IFC doc-IDs...")
    # Use dry_run IFCManager so it won't save
    dry_mgr = IFCManager(PROJECT, dry_run=True)
    summary = dry_mgr.update_ifc_deliverables(fake_results)
    print(f"\n  Result: updated={summary['updated']}, skipped={summary['skipped']}")
    if summary['errors']:
        print(f"  Errors: {summary['errors']}")
else:
    print("  Skipped (no IFC files or no Excel)")

sep("All tests complete")

"""Quick test: call IFCManager.convert_to_ifc() on C-PLN-003 RevA.
Reproduces the exact bot flow (ApprovedIFCManager → IFCManager).
"""
import sys, os, time, io
from pathlib import Path

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Add V6 to path
sys.path.insert(0, os.path.dirname(__file__))

import pythoncom
pythoncom.CoInitialize()

from ifr_automation_v10 import IFCManager

PROJECT_PATH = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS\Design\Engineering")
NATIVE_ROOT = PROJECT_PATH / "1. Drawings" / "1. Native"
DWG_FOLDER = NATIVE_ROOT / "GG31-C-PLN-003_Fence & Gate Layout Plan_Rev"
DWG_FILE = DWG_FOLDER / "GG31-C-PLN-003_Fence & Gate Layout Plan_RevA.dwg"
TEST_OUTPUT = Path(__file__).parent / "test_output"

print(f"DWG exists: {DWG_FILE.exists()}")
print(f"DWG: {DWG_FILE}")

# Create IFCManager same as ApprovedIFCManager does
ifc_mgr = IFCManager(
    project_path=PROJECT_PATH.parent,  # GG-31 Warnertown BESS
    title_block_name='ACE-Wanertown_Siyuan',
    dry_run=False,
)

# Override output to test_output
ifc_mgr._ifc_output_override = TEST_OUTPUT

dwg_info = {
    'doc_id': 'GG31-C-PLN-003',
    'description': 'Fence & Gate Layout Plan',
    'folder': DWG_FOLDER,
    'latest_ifr_dwg': DWG_FILE,
    'latest_ifr_rev': 'A',
    'existing_ifc_rev': None,  # First conversion (Rev0)
}

print(f"\n{'='*60}")
print(f"  Calling IFCManager.convert_to_ifc()")
print(f"{'='*60}\n")

result = ifc_mgr.convert_to_ifc(dwg_info)

print(f"\n{'='*60}")
print(f"  Result:")
print(f"{'='*60}")
for k, v in result.items():
    print(f"  {k}: {v}")

# Close AutoCAD docs
try:
    import win32com.client
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    while acad.Documents.Count > 0:
        acad.Documents.Item(0).Close(False)
    print("\nAll docs closed.")
except:
    pass

pythoncom.CoUninitialize()

# Tool-native DLV IFC update; reports/schedules keep their approved rev (scan regex can't read "_RevA-Approved")
import sys, openpyxl
from pathlib import Path
sys.path.insert(0, r"D:\1.dev\work\L03-pipeline\ifr-automation")
import ifr_automation_v10 as v
P = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\Waterloo")
ifc = v.DeliverableManager(P, dry_run=False).scan_ifc_folder()
for k, x in ifc.items():
    if not x['filename'].endswith('_Rev0_IFC.pdf'): x['revision'] = None
n = v.IFCTransmittalManager(P, dry_run=False).update_deliverable_ifc_status(ifc)
print("updated rows:", n)
dm = v.DeliverableManager(P, dry_run=True); xp = dm.find_deliverable_excel(); ws = openpyxl.load_workbook(xp).active
for r in (10,12,15,19,23,27,28,29,35,37,46,50,51,52,54,55,57,58,62,64,67,72,73,75,97):
    print(r, [ws.cell(r, c).value for c in (2, 12, 13, 14)])
print(xp.name, sorted(p.name for p in (xp.parent/"SS").iterdir())[-2:])

import sys, openpyxl
from pathlib import Path
sys.path.insert(0, r"D:\1.dev\work\L03-pipeline\ifr-automation")
import ifr_automation_v10 as v
P = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\Waterloo")
dm = v.DeliverableManager(P, dry_run=True)
ifc = dm.scan_ifc_folder()
for k, x in sorted(ifc.items()): print("SCAN", k, x)
xp = dm.find_deliverable_excel(); print("DLV", xp)
wb = openpyxl.load_workbook(xp); ws = wb.active; L = dm.detect_layout(ws)
print("sheet", ws.title, "cols id/rev/date/status", L.doc_id_col, L.rev_col, L.date_col, L.status_col, "hdr", L.header_row)
ids = {}
for r in range(L.header_row + 1, ws.max_row + 1):
    c = ws.cell(r, L.doc_id_col).value
    if c: ids[str(c).strip()] = (r, ws.cell(r, L.rev_col).value, ws.cell(r, L.date_col).value, ws.cell(r, L.status_col).value)
for k in sorted(ifc): print("MATCH" if k in ids else "NO-ROW", k, ids.get(k))
for k, x in sorted(ids.items()):
    if k not in ifc and x[3] and "pprov" in str(x[3]): print("APPROVED-BUT-NOT-SCANNED", k, x)

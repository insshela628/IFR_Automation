# GG38-TSMT-02 (IFC): same builder as TSMT-01, template = TSMT-01, files = 4. IFC(Client) top level
import openpyxl, re, datetime
from copy import copy
from pathlib import Path
T = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\Waterloo\Design\Engineering\8. Deliverables\Transmittal")
IFC = T.parents[1] / "1. Drawings/4. IFC(Client)"
src, out = T/"GG38-TSMT-01.xlsx", T/"GG38-TSMT-02.xlsx"
assert not out.exists()
wb = openpyxl.load_workbook(src); ws = wb.active
desc = {ws.cell(r, 3).value: ws.cell(r, 7).value for r in range(14, ws.max_row + 1) if ws.cell(r, 3).value}
proto = {r: [copy(ws.cell(r, c)._style) for c in range(1, 8)] for r in (14, 15)}
h14, h15 = ws.row_dimensions[14].height, ws.row_dimensions[15].height
for r in range(14, ws.max_row + 30):
    for c in range(1, 8): ws.cell(r, c).value = None
ws["E4"] = "No: GG38-TSMT-02"; ws["E5"] = "Date: " + datetime.date.today().strftime("%d/%m/%Y")
ws["C10"] = "ISSUED FOR CONSTRUCTION (IFC)"
files = sorted(f.name for f in IFC.glob("*_Rev0_IFC.pdf")); assert len(files) == 16
pat = re.compile(r"^(GG38-[A-Z]-[A-Z]{3}-\d{3}(?:-\d)?)_.+_Rev0_IFC\.pdf$")
for i, f in enumerate(files):
    d = pat.match(f).group(1); r = 14 + 2*i
    for rr, key in ((r, 14), (r+1, 15)):
        for c in range(1, 8): ws.cell(rr, c)._style = copy(proto[key][c-1])
    ws.row_dimensions[r].height = h14; ws.row_dimensions[r+1].height = h15
    ws.cell(r, 2, i+1); ws.cell(r, 3, d); ws.cell(r, 4, 1); ws.cell(r, 5, f); ws.cell(r, 6, "New"); ws.cell(r, 7, desc[d])
wb.save(out)
for r in range(4, 14 + 2*len(files), 2): print(r, [ws.cell(r, c).value for c in range(2, 8)])

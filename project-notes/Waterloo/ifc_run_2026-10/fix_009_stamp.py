# 009 PLAN layout: IFC_STAMP grazes note 18 -> move down 5 units (backup SS\*.pre-fix.dwg already taken by fix_ifc4)
import sys, time
src = open("fix_ifc4.py", encoding="utf-8").read().split("FIX = {")[0]
exec(src)
dwg = short(JOBS[4][2]); pdf = next(OUT.glob("GG38-C-PLN-009_*_Rev0_IFC.pdf"))
assert (dwg.parent / "SS" / (dwg.stem + ".pre-fix.dwg")).exists()
while a.Documents.Count > 0: a.Documents.Item(0).Close(False)
doc = ifc._com_retry(lambda: a.Documents.Open(str(dwg), False)); time.sleep(3)
done = []
for ln, blk in paper(doc):
    if ln != "PLAN": continue
    for e in [blk.Item(i) for i in range(blk.Count)]:
        if e.Layer == "IFC_STAMP": e.Move(pt(0, 0), pt(0, -5)); done.append(f"{ln}:{e.ObjectName}@{e.Handle}")
assert done
doc.Save(); time.sleep(1); doc.Close(False); time.sleep(1)
print("moved", done, flush=True)
print("republish", ifc._publish_single_pdf(a, dwg, pdf), flush=True)

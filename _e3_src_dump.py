"""Read-only dump of E-PLN-003 SOURCE IFC title block per sheet: revision
attributes (tag/value/X/Y) + standalone Text/MText in the TB area. Plans the
loose->attribute fix. Junction open; no save."""
import os, time, subprocess
import win32com.client

SRC = (r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\1.NSW"
       r"\NSW 153 - Coleambally #2\Design\Engineering\1. Drawings\1. Native"
       r"\NSW153-E-PLN-003_Communication Cable Route Layout Plan_Rev"
       r"\NSW153-E-PLN-003_Communication Cable Route Layout Plan_RevD_IFC.dwg")

def log(*a): print(*a, flush=True)
parent, fname = os.path.dirname(SRC), os.path.basename(SRC)
junc = os.path.join(os.environ.get("TEMP", r"C:\Temp"), "_e3src")
try:
    if os.path.exists(junc): subprocess.run(f'rmdir "{junc}"', shell=True)
    subprocess.run(f'mklink /J "{junc}" "{parent}"', shell=True, check=True)
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = False
    doc = acad.Documents.Open(os.path.join(junc, fname), True)
    for _ in range(60):
        try: _ = doc.Layouts.Count; _ = doc.ModelSpace.Count; break
        except Exception: time.sleep(0.5)
    time.sleep(2)
    log(f"opened layouts={doc.Layouts.Count}")
    for L in doc.Layouts:
        if L.Name.lower() == "model": continue
        blk = L.Block
        # title block attributes
        tb_rows = {}
        for i in range(blk.Count):
            try:
                e = blk.Item(i)
                if e.EntityName == "AcDbBlockReference" and e.HasAttributes:
                    for a in e.GetAttributes():
                        t = a.TagString.upper()
                        ip = a.InsertionPoint
                        tb_rows[t] = (a.TextString, round(float(ip[0]),1), round(float(ip[1]),1))
            except Exception: continue
        log(f"=== {L.Name} ATTRIBUTES (revision tags) ===")
        for t in sorted(tb_rows):
            if t[0].isdigit() and any(t.endswith(s) for s in
               ("REV","DATE","DESCRIPTION","DESIGNED","DRAWN","APPROVED","CHECK","PROJECT")):
                v,x,y = tb_rows[t]
                log(f"   {t:14} = {v!r:30} x={x} y={y}")
        # standalone text in TB area (y<60)
        log(f"--- {L.Name} STANDALONE Text/MText (y<60) ---")
        loose=[]
        for i in range(blk.Count):
            try:
                e = blk.Item(i)
                if e.EntityName in ("AcDbText","AcDbMText"):
                    ip=e.InsertionPoint; x,y=round(float(ip[0]),1),round(float(ip[1]),1)
                    if y<60: loose.append((y,x,e.Layer,e.EntityName,e.TextString))
            except Exception: continue
        for y,x,lay,en,t in sorted(loose):
            log(f"   y={y:6} x={x:7} {lay:9} {en[4:]:5} {t!r}")
    doc.Close(False)
    log("closed")
except Exception as e:
    log(f"ERROR: {e}")
finally:
    try: subprocess.run(f'rmdir "{junc}"', shell=True)
    except Exception: pass
    log("DONE")

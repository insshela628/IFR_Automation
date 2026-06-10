"""Repair E-PLN-003 SOURCE title block: write the loose C/D revision rows into
proper attribute rows 3/4, delete the loose revision MText. Then the bot's
rolling produces a clean aligned B,C,D,AS BUILT. Backup already taken.
Opens read-WRITE via junction; saves in place."""
import os, time, subprocess
import win32com.client

SRC = (r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\1.NSW"
       r"\NSW 153 - Coleambally #2\Design\Engineering\1. Drawings\1. Native"
       r"\NSW153-E-PLN-003_Communication Cable Route Layout Plan_Rev"
       r"\NSW153-E-PLN-003_Communication Cable Route Layout Plan_RevD_IFC.dwg")

C_ROW = {'REV':'C','DATE':'21/03/25','DESIGNED':'ACE','DRAWN':'MG','APPROVED':'AW',
         'PROJECT':'NSW153','DESCRIPTION':'MATCH TO AUXILIARY BLOCK DIAGRAM'}
D_ROW = {'REV':'D','DATE':'28/03/25','DESIGNED':'ACE','DRAWN':'MG','APPROVED':'AW',
         'PROJECT':'NSW153','DESCRIPTION':'ISSUED FOR CONSTRUCTION'}

def log(*a): print(*a, flush=True)
parent, fname = os.path.dirname(SRC), os.path.basename(SRC)
junc = os.path.join(os.environ.get("TEMP", r"C:\Temp"), "_e3fix")
try:
    if os.path.exists(junc): subprocess.run(f'rmdir "{junc}"', shell=True)
    subprocess.run(f'mklink /J "{junc}" "{parent}"', shell=True, check=True)
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = False
    doc = acad.Documents.Open(os.path.join(junc, fname), False)  # read-WRITE
    for _ in range(60):
        try: _ = doc.Layouts.Count; _ = doc.ModelSpace.Count; break
        except Exception: time.sleep(0.5)
    time.sleep(2)
    log(f"opened RW layouts={doc.Layouts.Count}")

    total_set = total_del = 0
    for L in doc.Layouts:
        if L.Name.lower() == "model": continue
        blk = L.Block
        # 1) fill attribute rows 3 (C) and 4 (D)
        set_n = 0
        for i in range(blk.Count):
            try:
                e = blk.Item(i)
                if e.EntityName == "AcDbBlockReference" and e.HasAttributes:
                    for a in e.GetAttributes():
                        t = a.TagString.upper()
                        for rownum, row in (('3', C_ROW), ('4', D_ROW)):
                            if t.startswith(rownum):
                                suf = t[len(rownum):]
                                if suf in row:
                                    a.TextString = row[suf]
                                    a.Update()
                                    set_n += 1
            except Exception as ex:
                log(f"   attr set err: {ex}")
        # 2) delete loose revision MText in the band (y28-39, x175-295)
        victims = []
        for i in range(blk.Count):
            try:
                e = blk.Item(i)
                if e.EntityName in ("AcDbText", "AcDbMText"):
                    ip = e.InsertionPoint
                    x, y = float(ip[0]), float(ip[1])
                    if 28 <= y <= 39 and 175 <= x <= 295:
                        victims.append(e)
            except Exception: continue
        for e in reversed(victims):
            try: e.Delete()
            except Exception: pass
        log(f"--- {L.Name}: set {set_n} attrs, deleted {len(victims)} loose MText ---")
        total_set += set_n; total_del += len(victims)

    log(f"TOTAL: set {total_set} attrs, deleted {total_del} loose MText")
    doc.Save()
    log("SAVED")
    doc.Close(False)
    log("closed")
except Exception as e:
    log(f"ERROR: {e}")
finally:
    try: subprocess.run(f'rmdir "{junc}"', shell=True)
    except Exception: pass
    log("DONE")

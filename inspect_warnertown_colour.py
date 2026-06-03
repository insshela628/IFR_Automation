"""
Read-only inspector: list stamp-zone entities of a Warnertown DWG to see exactly
what the surviving OLD COLOUR stamp is (type/layer/bbox/text). Reuses the proven
AsBuiltManager COM helpers (_com_retry, _find_all_title_blocks).

Opens ONE DWG read-only per run (opening 2 back-to-back gets COM-rejected).

Usage:
    python inspect_warnertown_colour.py SLD-001 ab       # the produced AB DWG
    python inspect_warnertown_colour.py SLD-001 source   # the source IFC DWG
"""
import sys, time
from pathlib import Path
import win32com.client, pythoncom

sys.path.insert(0, r"D:\1. SOP\SOP_Stage 2 IFR Sync√\V6√")
from ifr_automation_v10 import AsBuiltManager

NATIVE = Path(
    r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"
    r"\2.SA\GG-31 Warnertown BESS\Design\Engineering\1. Drawings\1. Native")
PROJECT = (r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"
           r"\2.SA\GG-31 Warnertown BESS")


def find_dwg(docid, which):
    docid = docid.upper()
    for folder in NATIVE.iterdir():
        if not folder.is_dir() or docid not in folder.name.upper():
            continue
        if which == 'source':
            ifcs = sorted(folder.glob("*_IFC.dwg"))
            if ifcs:
                return ifcs[-1]
        else:
            for sub in folder.glob("Rev*AB"):
                abs_ = sorted(sub.glob("*_AB.dwg"))
                if abs_:
                    return abs_[-1]
    return None


def main():
    docid = sys.argv[1] if len(sys.argv) > 1 else "SLD-001"
    which = sys.argv[2] if len(sys.argv) > 2 else "ab"
    path = find_dwg(docid, which)
    if not path:
        print(f"No {which} DWG for {docid}"); return
    print(f"INSPECT [{which}]: {path.name}", flush=True)

    pythoncom.CoInitialize()
    mgr = AsBuiltManager(PROJECT, dry_run=True)
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True
    mgr._acad = acad

    doc = mgr._com_retry(lambda: acad.Documents.Open(str(path), True))
    for _ in range(20):
        try:
            _ = doc.ModelSpace.Count; _ = doc.Layouts.Count; break
        except Exception:
            time.sleep(1)
    time.sleep(2)

    tbs = mgr._find_all_title_blocks(doc)
    print(f"  title blocks found: {len(tbs)}", flush=True)

    for idx, tb in enumerate(tbs):
        block_ref, attrs, space, layout_name = tb
        try:
            mn, mx = block_ref.GetBoundingBox()
            l, b, r, t = float(mn[0]), float(mn[1]), float(mx[0]), float(mx[1])
        except Exception as e:
            print(f"  TB#{idx} bbox err: {e}"); continue
        w, h = r - l, t - b
        zl, zr = l + w * 0.50, r + 10
        zb, zt = b - 10, b + h * 0.50   # bottom 50% (catch upper COLOUR box)
        print(f"\n  TB#{idx} [{layout_name}] bbox=({l:.0f},{b:.0f})-({r:.0f},{t:.0f}) "
              f"w={w:.0f} h={h:.0f}  zone x[{zl:.0f},{zr:.0f}] y[{zb:.0f},{zt:.0f}]",
              flush=True)
        try:
            n = space.Count
        except Exception:
            continue
        for i in range(n):
            try:
                e = win32com.client.Dispatch(space.Item(i))
                en = e.EntityName
                bmn, bmx = e.GetBoundingBox()
                ex0, ey0, ex1, ey1 = float(bmn[0]), float(bmn[1]), float(bmx[0]), float(bmx[1])
                cx, cy = (ex0 + ex1) / 2, (ey0 + ey1) / 2
                if not (zl <= cx <= zr and zb <= cy <= zt):
                    continue
                extra = ""
                if en in ('AcDbMText', 'AcDbText'):
                    raw = mgr._com_retry(lambda ent=e: ent.TextString) or ''
                    extra = " text='%s'" % mgr._strip_mtext_formatting(raw)[:42]
                elif en in ('AcDbPolyline', 'AcDbLwPolyline', 'AcDb2dPolyline'):
                    try:
                        extra = " closed=%s cw=%.2f" % (e.Closed, e.ConstantWidth)
                    except Exception:
                        pass
                elif en == 'AcDbBlockReference':
                    extra = " block='%s'" % e.Name
                print(f"    {en:<20} L='{e.Layer}' ({ex0:.0f},{ey0:.0f})-({ex1:.0f},{ey1:.0f}) "
                      f"{ex1-ex0:.0f}x{ey1-ey0:.0f}{extra}", flush=True)
            except Exception:
                continue

    try:
        doc.Close(False)
    except Exception:
        pass
    pythoncom.CoUninitialize()
    print("DONE", flush=True)


if __name__ == "__main__":
    main()

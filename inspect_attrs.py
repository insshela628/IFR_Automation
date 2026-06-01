"""List all title block attributes from a Tatua panel DWG.
Run with AutoCAD open.
"""
import win32com.client, time, sys

DWG = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\7.New Zealand\Tatua Solar Farm\1. Drawings\1. Native\STS1&2 Panel Design& SLD\TSF-EN-ELE-DRG-10-0A.dwg"

try:
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
except Exception:
    print("请先打开 AutoCAD")
    sys.exit(1)

doc = None
for _ in range(5):
    try:
        doc = acad.Documents.Open(DWG)
        break
    except:
        time.sleep(3)
if not doc:
    print("无法打开文件")
    sys.exit(1)

for _ in range(20):
    try:
        _ = doc.ModelSpace.Count
        break
    except:
        time.sleep(1)
time.sleep(2)

print(f"文件: {doc.Name}\n")

for space_name in ('ModelSpace', 'PaperSpace'):
    space = getattr(doc, space_name)
    for i in range(space.Count):
        ent = space.Item(i)
        try:
            if ent.EntityName != 'AcDbBlockReference':
                continue
            attrs = {}
            try:
                for a in ent.GetAttributes():
                    attrs[a.TagString] = a.TextString
            except:
                continue
            if len(attrs) < 3:
                continue
            print(f"[{space_name}] Block '{ent.Name}' — {len(attrs)} attrs:")
            for k, v in sorted(attrs.items()):
                print(f"  {k:30s} = '{v}'")
            print()
        except:
            continue

doc.Close(False)
print("Done.")

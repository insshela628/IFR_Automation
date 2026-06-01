"""Quick inspect of the 'IFR' block definition contents in Tatua_Standard_Frame.dwg"""
import win32com.client
import time

FRAME_DWG = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\7.New Zealand\Tatua Solar Farm\1. Drawings\1. Native\Tatua_Standard_Frame.dwg"

try:
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
except:
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True
    for _ in range(60):
        try:
            _ = acad.Documents
            break
        except:
            time.sleep(1)

# Check if already open
doc = None
try:
    doc = acad.ActiveDocument
    if 'Standard_Frame' not in doc.Name:
        doc = acad.Documents.Open(FRAME_DWG)
        time.sleep(3)
except:
    doc = acad.Documents.Open(FRAME_DWG)
    time.sleep(3)

print(f"Working with: {doc.Name}")

# Inspect ALL block definitions in detail
blocks = doc.Blocks
for i in range(blocks.Count):
    blk = blocks.Item(i)
    bname = blk.Name
    if bname.startswith('*'):
        continue

    ent_count = blk.Count
    # Show IFR block and any block with <10 entities (likely stamp-like)
    if bname in ('IFR', 'IFC', 'HAY', 'Frame_NEW') or ent_count <= 5:
        print(f"\n{'='*60}")
        print(f"Block '{bname}' - {ent_count} entities")
        print(f"  Origin: ({blk.Origin[0]:.2f}, {blk.Origin[1]:.2f}, {blk.Origin[2]:.2f})")
        print(f"{'='*60}")

        for j in range(ent_count):
            try:
                e = blk.Item(j)
                en = e.EntityName
                print(f"\n  [{j}] {en}")

                if en in ('AcDbMText', 'AcDbText'):
                    print(f"      Text: \"{e.TextString}\"")
                    print(f"      InsertionPoint: ({e.InsertionPoint[0]:.3f}, {e.InsertionPoint[1]:.3f})")
                    print(f"      Height: {e.Height:.3f}")
                    if en == 'AcDbMText':
                        try:
                            print(f"      AttachmentPoint: {e.AttachmentPoint}")
                            print(f"      Width: {e.Width:.3f}")
                            print(f"      Rotation: {e.Rotation:.3f}")
                        except:
                            pass
                    try:
                        print(f"      StyleName: {e.StyleName}")
                    except:
                        pass
                    try:
                        mn, mx = e.GetBoundingBox()
                        print(f"      BBox: ({float(mn[0]):.2f},{float(mn[1]):.2f})->({float(mx[0]):.2f},{float(mx[1]):.2f})")
                    except:
                        pass

                elif en == 'AcDbAttributeDefinition':
                    print(f"      Tag: '{e.TagString}'")
                    print(f"      Prompt: '{e.PromptString}'")
                    print(f"      Default: '{e.TextString}'")
                    print(f"      InsertionPoint: ({e.InsertionPoint[0]:.3f}, {e.InsertionPoint[1]:.3f})")
                    print(f"      Height: {e.Height:.3f}")

                elif en in ('AcDbPolyline', 'AcDbLwPolyline', 'AcDb2dPolyline'):
                    try:
                        mn, mx = e.GetBoundingBox()
                        print(f"      BBox: ({float(mn[0]):.2f},{float(mn[1]):.2f})->({float(mx[0]):.2f},{float(mx[1]):.2f})")
                        print(f"      Closed: {e.Closed}")
                        # Try to get vertices
                        try:
                            coords = e.Coordinates
                            n_pts = len(coords) // 2
                            print(f"      Vertices ({n_pts}):")
                            for k in range(n_pts):
                                print(f"        ({coords[k*2]:.3f}, {coords[k*2+1]:.3f})")
                        except Exception as ex:
                            print(f"      Coords error: {ex}")
                    except:
                        pass

                elif en == 'AcDbLine':
                    print(f"      Start: ({e.StartPoint[0]:.3f}, {e.StartPoint[1]:.3f})")
                    print(f"      End: ({e.EndPoint[0]:.3f}, {e.EndPoint[1]:.3f})")

                elif en == 'AcDbArc' or en == 'AcDbCircle':
                    print(f"      Center: ({e.Center[0]:.3f}, {e.Center[1]:.3f})")
                    print(f"      Radius: {e.Radius:.3f}")

                else:
                    try:
                        mn, mx = e.GetBoundingBox()
                        print(f"      BBox: ({float(mn[0]):.2f},{float(mn[1]):.2f})->({float(mx[0]):.2f},{float(mx[1]):.2f})")
                    except:
                        pass

            except Exception as ex:
                print(f"  [{j}] ERROR: {ex}")

# Also show the IFR block reference in ModelSpace
print(f"\n{'='*60}")
print("IFR block reference in ModelSpace:")
print(f"{'='*60}")
ms = doc.ModelSpace
for i in range(ms.Count):
    ent = ms.Item(i)
    if ent.EntityName == 'AcDbBlockReference' and ent.Name == 'IFR':
        print(f"  InsertionPoint: ({ent.InsertionPoint[0]:.3f}, {ent.InsertionPoint[1]:.3f})")
        print(f"  XScaleFactor: {ent.XScaleFactor:.3f}")
        print(f"  YScaleFactor: {ent.YScaleFactor:.3f}")
        print(f"  Rotation: {ent.Rotation:.6f}")
        try:
            mn, mx = ent.GetBoundingBox()
            print(f"  BBox: ({float(mn[0]):.2f},{float(mn[1]):.2f})->({float(mx[0]):.2f},{float(mx[1]):.2f})")
        except:
            pass

print("\nDone.")
input("Press Enter to close...")
doc.Close(False)

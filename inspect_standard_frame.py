"""Inspect Tatua_Standard_Frame.dwg to find the FOR CONSTRUCTION stamp structure.
Run with AutoCAD open.
"""
import win32com.client
import time

FRAME_DWG = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\7.New Zealand\Tatua Solar Farm\1. Drawings\1. Native\Tatua_Standard_Frame.dwg"

print("=" * 60)
print("Inspecting Tatua_Standard_Frame.dwg")
print("=" * 60)

try:
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    print("已连接到运行中的 AutoCAD")
except Exception:
    print("正在启动 AutoCAD（可能需要30秒）...")
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True
    # Wait for AutoCAD to be fully ready
    for _ in range(60):
        try:
            _ = acad.Documents.Count
            break
        except Exception:
            time.sleep(1)
    time.sleep(3)

print(f"Opening {FRAME_DWG}...")
doc = None
for _retry in range(5):
    try:
        doc = acad.Documents.Open(FRAME_DWG)
        break
    except Exception as e:
        print(f"  重试 {_retry+1}/5: {e}")
        time.sleep(3)

if doc is None:
    print("无法打开文件，退出")
    exit(1)

# Wait for document to fully load
for _wait in range(30):
    try:
        _ = doc.ModelSpace.Count
        break
    except Exception:
        time.sleep(1)
time.sleep(2)

print(f"\nLayouts:")
for layout in doc.Layouts:
    print(f"  '{layout.Name}' (TabOrder={layout.TabOrder})")

# Search both spaces for anything related to stamp
for space_name, get_space in [("PaperSpace", lambda: doc.PaperSpace),
                               ("ModelSpace", lambda: doc.ModelSpace)]:
    try:
        space = get_space()
        count = space.Count
        print(f"\n{'='*60}")
        print(f"{space_name}: {count} entities")
        print(f"{'='*60}")

        for i in range(count):
            try:
                ent = space.Item(i)
                ename = ent.EntityName

                # Print ALL entities with details
                if ename == 'AcDbBlockReference':
                    bname = ent.Name
                    ix, iy = ent.InsertionPoint[0], ent.InsertionPoint[1]
                    attrs = {}
                    try:
                        for a in ent.GetAttributes():
                            attrs[a.TagString] = a.TextString
                    except:
                        pass
                    try:
                        mn, mx = ent.GetBoundingBox()
                        bbox = f"bbox=({float(mn[0]):.1f},{float(mn[1]):.1f})->({float(mx[0]):.1f},{float(mx[1]):.1f})"
                    except:
                        bbox = "no bbox"
                    print(f"\n  [{i}] BlockRef '{bname}' at ({ix:.2f},{iy:.2f}) {bbox}")
                    if attrs:
                        print(f"      Attrs ({len(attrs)}):")
                        for k, v in sorted(attrs.items()):
                            print(f"        {k} = '{v}'")

                elif ename in ('AcDbMText', 'AcDbText'):
                    txt = ent.TextString
                    ix, iy = ent.InsertionPoint[0], ent.InsertionPoint[1]
                    h = ent.Height
                    try:
                        mn, mx = ent.GetBoundingBox()
                        bbox = f"bbox=({float(mn[0]):.1f},{float(mn[1]):.1f})->({float(mx[0]):.1f},{float(mx[1]):.1f})"
                    except:
                        bbox = "no bbox"
                    print(f"\n  [{i}] {ename} at ({ix:.2f},{iy:.2f}) h={h:.2f} {bbox}")
                    print(f"      Text: \"{txt[:100]}\"")
                    if ename == 'AcDbMText':
                        try:
                            print(f"      AttachPt: {ent.AttachmentPoint}")
                            print(f"      Width: {ent.Width:.2f}")
                        except:
                            pass

                elif ename == 'AcDbPolyline' or ename == 'AcDbLwPolyline' or ename == 'AcDb2dPolyline':
                    try:
                        mn, mx = ent.GetBoundingBox()
                        bbox = f"bbox=({float(mn[0]):.1f},{float(mn[1]):.1f})->({float(mx[0]):.1f},{float(mx[1]):.1f})"
                        closed = getattr(ent, 'Closed', 'N/A')
                        print(f"\n  [{i}] {ename} {bbox} Closed={closed}")
                    except:
                        print(f"\n  [{i}] {ename} (no bbox)")

                elif ename == 'AcDbLine':
                    try:
                        sx, sy = ent.StartPoint[0], ent.StartPoint[1]
                        ex, ey = ent.EndPoint[0], ent.EndPoint[1]
                        print(f"  [{i}] Line ({sx:.2f},{sy:.2f})->({ex:.2f},{ey:.2f})")
                    except:
                        print(f"  [{i}] Line (no coords)")

                else:
                    try:
                        mn, mx = ent.GetBoundingBox()
                        bbox = f"bbox=({float(mn[0]):.1f},{float(mn[1]):.1f})->({float(mx[0]):.1f},{float(mx[1]):.1f})"
                    except:
                        bbox = ""
                    print(f"  [{i}] {ename} {bbox}")

            except Exception as e:
                print(f"  [{i}] ERROR: {e}")

    except Exception as e:
        print(f"  {space_name} error: {e}")

# Also check block definitions for stamp-related blocks
print(f"\n{'='*60}")
print("Block Definitions:")
print(f"{'='*60}")
try:
    blocks = doc.Blocks
    for i in range(blocks.Count):
        try:
            blk = blocks.Item(i)
            bname = blk.Name
            if bname.startswith('*'):
                continue  # skip anonymous blocks
            ent_count = blk.Count
            # Check if block contains "CONSTRUCTION" related text
            has_construction = False
            for j in range(ent_count):
                try:
                    e = blk.Item(j)
                    if e.EntityName in ('AcDbMText', 'AcDbText'):
                        if 'CONSTRUCTION' in e.TextString.upper():
                            has_construction = True
                    elif e.EntityName == 'AcDbAttributeDefinition':
                        if 'CONSTRUCTION' in e.TagString.upper() or 'STAMP' in e.TagString.upper():
                            has_construction = True
                except:
                    pass

            if has_construction or 'STAMP' in bname.upper() or 'CONSTRUCTION' in bname.upper() or 'IFC' in bname.upper():
                print(f"\n  *** '{bname}' ({ent_count} entities) ***")
                for j in range(ent_count):
                    try:
                        e = blk.Item(j)
                        en = e.EntityName
                        if en in ('AcDbMText', 'AcDbText'):
                            print(f"    [{j}] {en}: \"{e.TextString[:80]}\" h={e.Height:.2f}")
                        elif en == 'AcDbAttributeDefinition':
                            print(f"    [{j}] AttrDef: tag='{e.TagString}' prompt='{e.PromptString}' default='{e.TextString}'")
                        elif en in ('AcDbPolyline', 'AcDbLwPolyline', 'AcDbLine'):
                            try:
                                mn, mx = e.GetBoundingBox()
                                print(f"    [{j}] {en} bbox=({float(mn[0]):.1f},{float(mn[1]):.1f})->({float(mx[0]):.1f},{float(mx[1]):.1f})")
                            except:
                                print(f"    [{j}] {en}")
                        else:
                            print(f"    [{j}] {en}")
                    except:
                        pass
            elif ent_count > 0:
                print(f"  '{bname}' ({ent_count} entities)")
        except:
            pass
except Exception as e:
    print(f"  Error reading blocks: {e}")

print(f"\n{'='*60}")
print("Done. Close without saving.")
print(f"{'='*60}")
input("Press Enter to close...")
doc.Close(False)

"""Inspect ACE_Standard_Frame_wanertown_UPDATED.dwg + PLN-004 to find:
1. Title block name and attributes
2. IFR stamp block structure and exact dimensions
3. 'DRAWINGS TO BE PRINTED IN COLOUR' box geometry

Run with AutoCAD open.
"""
import win32com.client
import time
import sys

FRAME_DWG = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS\Design\Engineering\1. Drawings\1. Native\ACE_Standard_Frame_wanertown_UPDATED.dwg"

PLN004_DWG = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS\Design\Engineering\1. Drawings\1. Native\GG31-C-PLN-004_External Road (site entrance) &Internal Track & Hardstands Layout Plan_Rev\GG31-C-PLN-004-Road Pavement.dwg"

try:
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    print("已连接到运行中的 AutoCAD")
except Exception:
    print("正在启动 AutoCAD...")
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True
    for _ in range(60):
        try:
            _ = acad.Documents.Count
            break
        except Exception:
            time.sleep(1)
    time.sleep(3)


def inspect_dwg(acad, dwg_path, label):
    """Open a DWG and inspect all block references and stamp-related blocks."""
    print(f"\n{'='*70}")
    print(f"  {label}")
    print(f"  {dwg_path}")
    print(f"{'='*70}")

    doc = None
    for _retry in range(5):
        try:
            doc = acad.Documents.Open(dwg_path)
            break
        except Exception as e:
            print(f"  重试 {_retry+1}/5: {e}")
            time.sleep(3)

    if doc is None:
        print("  *** 无法打开文件，跳过 ***")
        return

    for _wait in range(30):
        try:
            _ = doc.ModelSpace.Count
            break
        except Exception:
            time.sleep(1)
    time.sleep(2)

    # --- Layouts ---
    print(f"\nLayouts:")
    try:
        for layout in doc.Layouts:
            print(f"  '{layout.Name}' (TabOrder={layout.TabOrder})")
    except Exception as e:
        print(f"  Error: {e}")

    # --- Block references in all spaces ---
    for space_name in ('ModelSpace',):
        try:
            space = getattr(doc, space_name)
            print(f"\n--- {space_name}: {space.Count} entities ---")
            for i in range(space.Count):
                try:
                    ent = space.Item(i)
                    if ent.EntityName != 'AcDbBlockReference':
                        continue
                    bname = ent.Name
                    attrs = {}
                    try:
                        for a in ent.GetAttributes():
                            attrs[a.TagString] = a.TextString
                    except:
                        pass

                    if len(attrs) >= 3:
                        try:
                            mn, mx = ent.GetBoundingBox()
                            w = float(mx[0]) - float(mn[0])
                            h = float(mx[1]) - float(mn[1])
                            bbox = f"size={w:.1f}x{h:.1f}"
                        except:
                            bbox = ""
                        print(f"\n  ★ Block '{bname}' ({len(attrs)} attrs) {bbox}")
                        print(f"    InsertionPoint: ({ent.InsertionPoint[0]:.2f}, {ent.InsertionPoint[1]:.2f})")
                        for k, v in sorted(attrs.items()):
                            print(f"    {k:30s} = '{v}'")
                except:
                    continue
        except Exception as e:
            print(f"  {space_name} error: {e}")

    # Also check PaperSpace layouts
    try:
        for layout in doc.Layouts:
            if layout.Name.lower() == 'model':
                continue
            block = layout.Block
            print(f"\n--- PaperSpace '{layout.Name}': {block.Count} entities ---")
            for i in range(block.Count):
                try:
                    ent = block.Item(i)
                    if ent.EntityName != 'AcDbBlockReference':
                        continue
                    bname = ent.Name
                    attrs = {}
                    try:
                        for a in ent.GetAttributes():
                            attrs[a.TagString] = a.TextString
                    except:
                        pass
                    if len(attrs) >= 3:
                        try:
                            mn, mx = ent.GetBoundingBox()
                            w = float(mx[0]) - float(mn[0])
                            h = float(mx[1]) - float(mn[1])
                            bbox = f"size={w:.1f}x{h:.1f}"
                        except:
                            bbox = ""
                        print(f"\n  ★ Block '{bname}' ({len(attrs)} attrs) {bbox}")
                        for k, v in sorted(attrs.items()):
                            print(f"    {k:30s} = '{v}'")
                except:
                    continue
    except Exception as e:
        print(f"  PaperSpace error: {e}")

    # --- Stamp-related block definitions ---
    print(f"\n--- Block Definitions (stamp-related) ---")
    try:
        blocks = doc.Blocks
        for i in range(blocks.Count):
            try:
                blk = blocks.Item(i)
                bname = blk.Name
                if bname.startswith('*'):
                    continue
                ent_count = blk.Count
                # Check for stamp-related content
                has_stamp = False
                keywords = ('IFR', 'IFC', 'STAMP', 'CONSTRUCTION', 'REVIEW', 'COLOUR', 'COLOR', 'PRINTED')
                bname_upper = bname.upper()
                for kw in keywords:
                    if kw in bname_upper:
                        has_stamp = True
                        break
                if not has_stamp and ent_count <= 10:
                    for j in range(ent_count):
                        try:
                            e = blk.Item(j)
                            if e.EntityName in ('AcDbMText', 'AcDbText'):
                                txt = e.TextString.upper()
                                for kw in keywords:
                                    if kw in txt:
                                        has_stamp = True
                                        break
                        except:
                            pass
                        if has_stamp:
                            break

                if has_stamp:
                    print(f"\n  *** Block '{bname}' ({ent_count} entities) ***")
                    for j in range(ent_count):
                        try:
                            e = blk.Item(j)
                            en = e.EntityName
                            if en in ('AcDbMText', 'AcDbText'):
                                try:
                                    mn, mx = e.GetBoundingBox()
                                    w = float(mx[0]) - float(mn[0])
                                    h = float(mx[1]) - float(mn[1])
                                    bbox = f"size={w:.1f}x{h:.1f}"
                                except:
                                    bbox = ""
                                print(f"    [{j}] {en}: \"{e.TextString[:80]}\" h={e.Height:.2f} {bbox}")
                                print(f"         pos=({e.InsertionPoint[0]:.3f}, {e.InsertionPoint[1]:.3f})")
                            elif en in ('AcDbPolyline', 'AcDbLwPolyline'):
                                try:
                                    mn, mx = e.GetBoundingBox()
                                    w = float(mx[0]) - float(mn[0])
                                    h = float(mx[1]) - float(mn[1])
                                    print(f"    [{j}] {en}: size={w:.3f}x{h:.3f} Closed={e.Closed}")
                                    print(f"         bbox=({float(mn[0]):.3f},{float(mn[1]):.3f})->({float(mx[0]):.3f},{float(mx[1]):.3f})")
                                except:
                                    print(f"    [{j}] {en}")
                            elif en == 'AcDbLine':
                                print(f"    [{j}] Line ({e.StartPoint[0]:.3f},{e.StartPoint[1]:.3f})->({e.EndPoint[0]:.3f},{e.EndPoint[1]:.3f})")
                            else:
                                print(f"    [{j}] {en}")
                        except Exception as ex:
                            print(f"    [{j}] ERROR: {ex}")
            except:
                pass
    except Exception as e:
        print(f"  Error: {e}")

    doc.Close(False)
    time.sleep(2)
    print(f"\n  [已关闭 {label}]")


# --- Run inspections ---
inspect_dwg(acad, FRAME_DWG, "ACE_Standard_Frame_wanertown_UPDATED.dwg")
inspect_dwg(acad, PLN004_DWG, "GG31-C-PLN-004 (Road Pavement)")

print(f"\n{'='*70}")
print("全部完成。请把输出复制给我。")
print(f"{'='*70}")
input("按 Enter 退出...")

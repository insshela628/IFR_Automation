"""Inspect IFC DWG stamp status — check what stamps exist and where.
Also checks PLN-004 lock files.

Run with AutoCAD open.
"""
import win32com.client
import time
import sys
from pathlib import Path

IFC_CLIENT = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS\Design\Engineering\1. Drawings\4. IFC(Client)")

PLN004_FOLDER = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS\Design\Engineering\1. Drawings\1. Native\GG31-C-PLN-004_External Road (site entrance) &Internal Track & Hardstands Layout Plan_Rev")

# ── Check PLN-004 lock files ──
print("=" * 70)
print("  PLN-004 Lock File Check")
print("=" * 70)
for ext in ('.dwl', '.dwl2'):
    for f in PLN004_FOLDER.glob(f"*{ext}"):
        print(f"  LOCK: {f.name}")
        try:
            content = f.read_text(errors='replace')[:200]
            print(f"    内容: {content}")
        except Exception as e:
            print(f"    读取失败: {e}")
        try:
            f.unlink()
            print(f"    ✓ 已删除")
        except PermissionError:
            print(f"    ✗ 无法删除 (文件被占用 — 请关闭 AutoCAD 中的该文件)")
        except Exception as e:
            print(f"    ✗ 删除失败: {e}")

# ── Find E-PLN-003 IFC DWG ──
print(f"\n{'='*70}")
print("  E-PLN-003 IFC DWG Stamp Inspection")
print(f"{'='*70}")

ifc_dwgs = list(IFC_CLIENT.glob("GG31-E-PLN-003*.dwg"))
if not ifc_dwgs:
    # Check in native folder
    native_003 = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS\Design\Engineering\1. Drawings\1. Native\GG31-E-PLN-003_Communication Cable Route Layout Plan_Rev")
    ifc_dwgs = list(native_003.glob("*IFC*.dwg"))

if not ifc_dwgs:
    print("  未找到 E-PLN-003 IFC DWG")
    print(f"  搜索路径: {IFC_CLIENT}")
    # List what's actually there
    print(f"\n  IFC(Client) 内容:")
    for f in sorted(IFC_CLIENT.iterdir()):
        if f.is_file() and not f.name.startswith('~'):
            print(f"    {f.name}")
else:
    dwg_path = ifc_dwgs[0]
    print(f"  DWG: {dwg_path.name}")

    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        print("  已连接到 AutoCAD")
    except:
        print("  请先打开 AutoCAD")
        sys.exit(1)

    doc = None
    for _retry in range(5):
        try:
            doc = acad.Documents.Open(str(dwg_path))
            break
        except Exception as e:
            print(f"  重试 {_retry+1}/5: {e}")
            time.sleep(3)

    if not doc:
        print("  无法打开 DWG")
        sys.exit(1)

    for _wait in range(30):
        try:
            _ = doc.ModelSpace.Count
            break
        except:
            time.sleep(1)
    time.sleep(3)
    # Extra wait for Layouts (COM often rejects first call)
    for _wait in range(10):
        try:
            _ = doc.Layouts.Count
            break
        except:
            time.sleep(2)

    # ── Check Layouts ──
    print(f"\n  Layouts:")
    for _retry in range(5):
        try:
            for layout in doc.Layouts:
                ent_count = layout.Block.Count if layout.Name.lower() != 'model' else '(skip count)'
                print(f"    '{layout.Name}' — {ent_count} entities")
            break
        except Exception as e:
            if _retry < 4:
                time.sleep(2)
            else:
                print(f"    Layouts error: {e}")

    # ── Check IFC_STAMP layer entities ──
    print(f"\n  IFC_STAMP layer entities:")
    stamp_count = 0
    try:
        import pythoncom
        for space_name in ('ModelSpace',):
            space = getattr(doc, space_name)
            # Use SelectionSet to find IFC_STAMP entities
            ss_name = f"_InspStamp_{int(time.time()*1000)%1000000}"
            try:
                ss = doc.SelectionSets.Add(ss_name)
                filter_type = win32com.client.VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_I2, [8])  # layer
                filter_data = win32com.client.VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["IFC_STAMP"])
                ss.Select(5)
                for i in range(ss.Count):
                    ent = ss.Item(i)
                    if ent.Layer == "IFC_STAMP":
                        stamp_count += 1
                        en = ent.EntityName
                        try:
                            mn, mx = ent.GetBoundingBox()
                            w = float(mx[0]) - float(mn[0])
                            h = float(mx[1]) - float(mn[1])
                            bbox = f"size={w:.1f}x{h:.1f} at ({float(mn[0]):.1f},{float(mn[1]):.1f})"
                        except:
                            bbox = ""
                        if en in ('AcDbMText', 'AcDbText'):
                            txt = ent.TextString[:60]
                            print(f"    [{stamp_count}] {en}: \"{txt}\" h={ent.Height:.1f} {bbox}")
                        elif 'Polyline' in en:
                            print(f"    [{stamp_count}] {en}: {bbox} Closed={getattr(ent, 'Closed', '?')}")
                        else:
                            print(f"    [{stamp_count}] {en}: {bbox}")
                ss.Delete()
            except Exception as e:
                print(f"    SelectionSet error: {e}")
                try:
                    doc.SelectionSets.Item(ss_name).Delete()
                except:
                    pass
    except Exception as e:
        print(f"    Error: {e}")

    # ── Also check PaperSpace layouts ──
    print(f"\n  PaperSpace stamp entities:")
    ps_stamp = 0
    time.sleep(2)
    try:
        for layout in doc.Layouts:
            if layout.Name.lower() == 'model':
                continue
            block = layout.Block
            for i in range(block.Count):
                try:
                    ent = block.Item(i)
                    if ent.Layer == "IFC_STAMP":
                        ps_stamp += 1
                        en = ent.EntityName
                        if en in ('AcDbMText', 'AcDbText'):
                            print(f"    [{layout.Name}] {en}: \"{ent.TextString[:60]}\" h={ent.Height:.1f}")
                        else:
                            try:
                                mn, mx = ent.GetBoundingBox()
                                w = float(mx[0]) - float(mn[0])
                                h = float(mx[1]) - float(mn[1])
                                print(f"    [{layout.Name}] {en}: {w:.1f}x{h:.1f}")
                            except:
                                print(f"    [{layout.Name}] {en}")
                except:
                    continue
    except Exception as e:
        print(f"    Error: {e}")

    if stamp_count == 0 and ps_stamp == 0:
        print("    ⚠ 没有找到 IFC_STAMP 图层实体!")

    # ── Check ALL title blocks ──
    print(f"\n  Title blocks (all spaces):")
    tb_names = ('ACE-Wanertown_Siyuan',)
    for space_name in ('ModelSpace',):
        space = getattr(doc, space_name)
        for i in range(space.Count):
            try:
                ent = space.Item(i)
                if ent.EntityName != 'AcDbBlockReference':
                    continue
                bname = ent.Name
                attrs = {}
                try:
                    for a in ent.GetAttributes():
                        attrs[a.TagString.upper()] = a.TextString
                except:
                    pass
                if bname in tb_names or len(attrs) > 10:
                    rev = attrs.get('REVISION', '?')
                    desc1 = attrs.get('1DESCRIPTION', '?')
                    desc2 = attrs.get('2DESCRIPTION', '?')
                    print(f"    [{space_name}] '{bname}' ({len(attrs)} attrs)")
                    print(f"      REVISION={rev}, 1DESC={desc1}, 2DESC={desc2}")
            except:
                continue

    # PaperSpace title blocks
    try:
        for layout in doc.Layouts:
            if layout.Name.lower() == 'model':
                continue
            block = layout.Block
            for i in range(block.Count):
                try:
                    ent = block.Item(i)
                    if ent.EntityName != 'AcDbBlockReference':
                        continue
                    bname = ent.Name
                    attrs = {}
                    try:
                        for a in ent.GetAttributes():
                            attrs[a.TagString.upper()] = a.TextString
                    except:
                        pass
                    if bname in tb_names or len(attrs) > 10:
                        rev = attrs.get('REVISION', '?')
                        desc1 = attrs.get('1DESCRIPTION', '?')
                        desc2 = attrs.get('2DESCRIPTION', '?')
                        print(f"    [{layout.Name}] '{bname}' ({len(attrs)} attrs)")
                        print(f"      REVISION={rev}, 1DESC={desc1}, 2DESC={desc2}")
                except:
                    continue
    except Exception as e:
        print(f"    PaperSpace error: {e}")

    # ── Check for remaining IFR/legacy stamp MText ──
    print(f"\n  Legacy stamp MText (not on IFC_STAMP layer):")
    legacy = 0
    keywords = ('FOR CONSTRUCTION', 'FOR REVIEW', 'ISSUED FOR REVIEW',
                'DRAWINGS TO BE PRINTED IN COLOUR')
    try:
        import pythoncom
        ss_name2 = f"_InspLegacy_{int(time.time()*1000)%1000000}"
        ss2 = doc.SelectionSets.Add(ss_name2)
        filter_type = win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
        filter_data = win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["MTEXT"])
        ss2.Select(5)
        for i in range(ss2.Count):
            try:
                ent = ss2.Item(i)
                if ent.Layer == "IFC_STAMP":
                    continue
                txt = ent.TextString.upper()
                # Strip MText formatting
                import re
                plain = re.sub(r'\\[A-Za-z][^;]*;', '', txt)
                plain = re.sub(r'[{}]', '', plain).strip()
                for kw in keywords:
                    if kw in plain:
                        legacy += 1
                        print(f"    [{legacy}] Layer='{ent.Layer}' Text=\"{ent.TextString[:60]}\"")
                        break
            except:
                continue
        ss2.Delete()
    except Exception as e:
        print(f"    Error: {e}")
    if legacy == 0:
        print("    (无)")

    print(f"\n  总结: ModelSpace stamps={stamp_count}, PaperSpace stamps={ps_stamp}, Legacy={legacy}")

    doc.Close(False)
    print("  [已关闭]")

print(f"\n{'='*70}")
print("完成。请把输出复制给我。")
print(f"{'='*70}")

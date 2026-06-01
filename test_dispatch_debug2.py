"""Debug: replicate IFCManager flow step by step on C-PLN-003"""
import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import win32com.client
import pythoncom

acad = win32com.client.GetActiveObject("AutoCAD.Application")

# Close any open docs first
while acad.Documents.Count > 0:
    acad.Documents.Item(0).Close(False)
    time.sleep(0.5)

dwg = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS\Design\Engineering\1. Drawings\1. Native\GG31-C-PLN-003_Fence & Gate Layout Plan_Rev\GG31-C-PLN-003_Fence & Gate Layout Plan_RevA.dwg"
doc = acad.Documents.Open(dwg)
time.sleep(3)

# Step 1: Find title blocks (like _find_all_title_blocks)
print("Step 1: Find title blocks via SelectionSet...")
all_tbs = []

# Pass 1: ModelSpace SelectionSet
try:
    doc.ActiveSpace = 1  # acModelSpace
except:
    pass

ss_name = f"_TBSearch_{int(time.time()*1000) % 1000000}"
try:
    ss = doc.SelectionSets.Add(ss_name)
    ft = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
    fv = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["INSERT"])
    ss.Select(5, None, None, ft, fv)
    ms = doc.ModelSpace
    for i in range(ss.Count):
        entity = ss.Item(i)
        try:
            bname = entity.Name
            attrs_raw = entity.GetAttributes()
            attr_tags = set()
            for a in attrs_raw:
                attr_tags.add(a.TagString.upper())
            if 'DRAWINGNUMBER' in attr_tags and 'REVISION' in attr_tags:
                all_tbs.append((entity, attrs_raw, ms, 'Model'))
                print(f"  Found: '{bname}' in Model ({len(attr_tags)} attrs)")
        except:
            continue
    ss.Delete()
except Exception as e:
    print(f"  Pass 1 failed: {e}")

# Pass 2: PaperSpace
for layout in doc.Layouts:
    if layout.Name.lower() == 'model':
        continue
    blk = layout.Block
    for i in range(blk.Count):
        try:
            ent = blk.Item(i)
            if ent.EntityName != 'AcDbBlockReference':
                continue
            attrs_raw = ent.GetAttributes()
            attr_tags = set()
            for a in attrs_raw:
                attr_tags.add(a.TagString.upper())
            if 'DRAWINGNUMBER' in attr_tags and 'REVISION' in attr_tags:
                all_tbs.append((ent, attrs_raw, blk, layout.Name))
                print(f"  Found: '{ent.Name}' in {layout.Name} ({len(attr_tags)} attrs)")
        except:
            continue

print(f"\nTotal: {len(all_tbs)} title blocks")

# Step 2: Test if we can still set TextString on the first TB's attrs
print("\nStep 2: Test attr access...")
if all_tbs:
    entity, attrs_raw, space, sname = all_tbs[0]
    for a in attrs_raw:
        try:
            tag = a.TagString
            if tag.upper() == 'REVISION':
                old = a.TextString
                print(f"  Read REVISION: '{old}' — OK")
                a.TextString = 'TEST'
                new = a.TextString
                print(f"  Write REVISION: '{new}' — OK")
                a.TextString = old  # restore
                print(f"  Restored REVISION: '{old}'")
                break
        except Exception as e:
            print(f"  FAILED attr access: {e}")
            break

# Step 3: Run a SelectionSet (simulating _scan_has_colour or _remove_ifc_stamp)
print("\nStep 3: Run intermediate SelectionSet (simulating _scan_has_colour)...")
try:
    ss2 = doc.SelectionSets.Add(f'scan_{int(time.time())}')
    ft2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
    fv2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["MTEXT"])
    ss2.Select(5, None, None, ft2, fv2)
    print(f"  Found {ss2.Count} MTEXT entities")
    ss2.Delete()
except Exception as e:
    print(f"  Failed: {e}")

# Step 4: Now test if attrs are still accessible
print("\nStep 4: Re-test attr access AFTER intermediate SelectionSet...")
if all_tbs:
    entity, attrs_raw, space, sname = all_tbs[0]
    for a in attrs_raw:
        try:
            tag = a.TagString
            if tag.upper() == 'REVISION':
                old = a.TextString
                print(f"  Read REVISION: '{old}' — OK")
                a.TextString = 'TEST2'
                new = a.TextString
                print(f"  Write REVISION: '{new}' — OK")
                a.TextString = old  # restore
                break
        except Exception as e:
            print(f"  FAILED: {e}")
            break

# Step 5: Test AddLightWeightPolyline on the stored space
print("\nStep 5: AddLightWeightPolyline on stored space...")
if all_tbs:
    entity, attrs_raw, space, sname = all_tbs[0]
    print(f"  Space type: {type(space)}, from: {sname}")
    pts = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [100.0, 100.0, 200.0, 100.0, 200.0, 200.0, 100.0, 200.0])
    try:
        pline = space.AddLightWeightPolyline(pts)
        print(f"  OK! Polyline created")
        pline.Delete()
    except Exception as e:
        print(f"  FAILED: {e}")

    # Try fresh ModelSpace
    print("\nStep 5b: Fresh doc.ModelSpace...")
    try:
        ms_fresh = doc.ModelSpace
        pline = ms_fresh.AddLightWeightPolyline(pts)
        print(f"  OK! Polyline created on fresh ms")
        pline.Delete()
    except Exception as e:
        print(f"  FAILED: {e}")

doc.Close(False)
print("\nDone.")

"""Debug script: test FOR CONSTRUCTION stamp placement in AutoCAD.
Run with AutoCAD open and a DWG loaded (or it will open the test DWG).

Tests:
  1. Find title block in PaperSpace/ModelSpace
  2. Print GetBoundingBox() coordinates
  3. Try AddMText with array.array (preferred)
  4. Try AddMText with VARIANT
  5. Try SendCommand -MTEXT
  6. Verify stamp was created by re-scanning entities
"""
import win32com.client
import pythoncom
import array
import time

DWG = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\7.New Zealand\Tatua Solar Farm\1. Drawings\1. Native\STS1&2 Panel Design& SLD\TSF-EN-ELE-DIA-01-00.dwg"

print("=" * 60)
print("FOR CONSTRUCTION Stamp Debug Tool")
print("=" * 60)

# Connect to AutoCAD
print("\n[1] Connecting to AutoCAD...")
try:
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    print(f"  Connected to running AutoCAD")
except:
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True
    for _ in range(60):
        try:
            _ = acad.Documents
            break
        except:
            time.sleep(1)
    print(f"  Started new AutoCAD instance")

# Open DWG if needed
print(f"\n[2] Opening DWG...")
doc = None
try:
    doc = acad.ActiveDocument
    print(f"  Active doc: {doc.Name}")
    if 'DIA-01-00' not in doc.Name:
        print(f"  Not our test file, opening target...")
        doc = acad.Documents.Open(DWG)
        time.sleep(5)
except:
    doc = acad.Documents.Open(DWG)
    time.sleep(5)
print(f"  Working with: {doc.Name}")

# List layouts
print(f"\n[3] Layouts:")
for layout in doc.Layouts:
    print(f"  '{layout.Name}' (TabOrder={layout.TabOrder})")

# Search for title block
print(f"\n[4] Searching for title block...")
title_block = None
title_space = None
title_space_name = None

for space_name, get_space in [("PaperSpace", lambda: doc.PaperSpace),
                               ("ModelSpace", lambda: doc.ModelSpace)]:
    try:
        space = get_space()
        count = space.Count
        print(f"\n  --- {space_name}: {count} entities ---")

        blocks_found = []
        for i in range(count):
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

                blocks_found.append((bname, len(attrs)))

                # Check if this is a title block
                has_dwgnum = 'DRAWINGNUMBER' in attrs
                has_rev = 'REVISION' in attrs
                rev_tags = [k for k in attrs if 'REV' in k]

                if (has_dwgnum and has_rev) or (rev_tags and len(attrs) >= 5):
                    print(f"\n  *** TITLE BLOCK FOUND in {space_name} ***")
                    print(f"  Block: '{bname}'")
                    print(f"  InsertionPoint: ({ent.InsertionPoint[0]:.2f}, {ent.InsertionPoint[1]:.2f})")
                    print(f"  Attributes ({len(attrs)}):")
                    for k, v in sorted(attrs.items()):
                        print(f"    {k} = '{v}'")

                    # BoundingBox
                    try:
                        min_pt, max_pt = ent.GetBoundingBox()
                        bx1, by1 = float(min_pt[0]), float(min_pt[1])
                        bx2, by2 = float(max_pt[0]), float(max_pt[1])
                        print(f"  BoundingBox: ({bx1:.2f}, {by1:.2f}) -> ({bx2:.2f}, {by2:.2f})")
                        print(f"  Size: {bx2-bx1:.1f} x {by2-by1:.1f}")

                        title_block = ent
                        title_space = space
                        title_space_name = space_name
                    except Exception as e:
                        print(f"  BoundingBox ERROR: {e}")

            except Exception as e:
                pass

        if blocks_found:
            print(f"\n  All blocks in {space_name}:")
            for bname, ac in blocks_found:
                print(f"    '{bname}' ({ac} attrs)")

    except Exception as e:
        print(f"  {space_name} error: {e}")

if title_block is None:
    print("\n*** ERROR: No title block found! ***")
    input("Press Enter to exit...")
    exit(1)

# Calculate stamp position
min_pt, max_pt = title_block.GetBoundingBox()
stamp_x = float(max_pt[0])
stamp_y = float(max_pt[1]) + 5
print(f"\n[5] Stamp target position: ({stamp_x:.2f}, {stamp_y:.2f})")
print(f"    (title block top-right + 5mm up)")

# Ensure PaperSpace layout is active
print(f"\n[6] Ensuring PaperSpace layout is active...")
for layout in doc.Layouts:
    if layout.Name.lower() != 'model':
        doc.ActiveLayout = layout
        print(f"  Active layout set to: '{layout.Name}'")
        break
time.sleep(0.5)

# Re-acquire space reference
if title_space_name == "PaperSpace":
    title_space = doc.PaperSpace
    print(f"  Re-acquired PaperSpace reference (count={title_space.Count})")

stamp_text = "{\\fArial Narrow|b1;FOR CONSTRUCTION}"
stamp_height = 9.5

# Count entities before
before_count = title_space.Count
print(f"\n  Entities before stamp: {before_count}")

# --- Test 1: array.array ---
print(f"\n[7] Test 1: COM AddMText with array.array...")
try:
    pt = array.array('d', [stamp_x, stamp_y, 0.0])
    print(f"  Point: array.array('d', [{stamp_x}, {stamp_y}, 0.0])")
    mtext = title_space.AddMText(pt, 300.0, stamp_text)
    mtext.Height = stamp_height
    mtext.AttachmentPoint = 9  # BottomRight
    doc.Regen(1)
    after_count = title_space.Count
    print(f"  *** SUCCESS! Entities after: {after_count} (was {before_count})")
    print(f"  MText InsertionPoint: ({mtext.InsertionPoint[0]:.2f}, {mtext.InsertionPoint[1]:.2f})")
    print(f"  MText Height: {mtext.Height}")
    print(f"  MText TextString: {mtext.TextString}")

    # Delete it so we can test the next approach
    mtext.Delete()
    doc.Regen(1)
    print(f"  (deleted for next test)")

except Exception as e:
    print(f"  FAILED: {e}")
    print(f"  Error type: {type(e).__name__}")

# --- Test 2: VARIANT ---
print(f"\n[8] Test 2: COM AddMText with VARIANT...")
try:
    pt = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [stamp_x, stamp_y, 0.0])
    print(f"  Point: VARIANT(VT_ARRAY|VT_R8, [{stamp_x}, {stamp_y}, 0.0])")
    mtext = title_space.AddMText(pt, 300.0, stamp_text)
    mtext.Height = stamp_height
    mtext.AttachmentPoint = 9
    doc.Regen(1)
    after_count = title_space.Count
    print(f"  *** SUCCESS! Entities after: {after_count}")
    print(f"  MText InsertionPoint: ({mtext.InsertionPoint[0]:.2f}, {mtext.InsertionPoint[1]:.2f})")

    mtext.Delete()
    doc.Regen(1)
    print(f"  (deleted for next test)")

except Exception as e:
    print(f"  FAILED: {e}")
    print(f"  Error type: {type(e).__name__}")

# --- Test 3: SendCommand ---
print(f"\n[9] Test 3: SendCommand -MTEXT...")
try:
    opp_x = stamp_x - 300
    opp_y = stamp_y - stamp_height
    cmd = (
        f'-MTEXT\n'
        f'{stamp_x},{stamp_y}\n'
        f'J\nBR\n'
        f'H\n{stamp_height}\n'
        f'{opp_x},{opp_y}\n'
        f'{stamp_text}\n'
        f'\n'
    )
    print(f"  Command:\n{cmd}")
    doc.SendCommand(cmd)

    for _ in range(15):
        try:
            if doc.GetVariable("CMDACTIVE") == 0:
                break
        except:
            break
        time.sleep(0.5)

    doc.Regen(1)
    after_count = title_space.Count
    print(f"  Entities after: {after_count} (was {before_count})")
    if after_count > before_count:
        print(f"  *** SUCCESS! New entity added.")
    else:
        print(f"  WARNING: No new entity detected!")

except Exception as e:
    print(f"  FAILED: {e}")

# --- Scan for any MText with "CONSTRUCTION" ---
print(f"\n[10] Scanning for FOR CONSTRUCTION MText...")
for sname, get_sp in [("PaperSpace", lambda: doc.PaperSpace),
                       ("ModelSpace", lambda: doc.ModelSpace)]:
    try:
        sp = get_sp()
        for i in range(sp.Count):
            try:
                ent = sp.Item(i)
                if ent.EntityName in ('AcDbMText', 'AcDbText'):
                    txt = ent.TextString
                    if 'CONSTRUCTION' in txt.upper():
                        print(f"  FOUND in {sname}[{i}]: '{txt}'")
                        print(f"    Position: ({ent.InsertionPoint[0]:.2f}, {ent.InsertionPoint[1]:.2f})")
                        print(f"    Height: {ent.Height}")
            except:
                pass
    except:
        pass

print(f"\n{'='*60}")
print("Done. Check AutoCAD to see if stamp appeared.")
print("Close without saving: doc.Close(False)")
print(f"{'='*60}")
input("Press Enter to close document...")
doc.Close(False)

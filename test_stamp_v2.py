"""Test the new IFC stamp (rectangle + MText) matching IFR block style.
Run with AutoCAD open.
"""
import win32com.client
import pythoncom
import time

DWG = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\7.New Zealand\Tatua Solar Farm\1. Drawings\1. Native\STS1&2 Panel Design& SLD\TSF-EN-ELE-DIA-01-00.dwg"

# Stamp constants (from Tatua_Standard_Frame.dwg 'IFR' block)
STAMP_X = 141.0
STAMP_Y = 589.176
RECT_W = 110.511
RECT_H = 17.745
TEXT_Y_OFFSET = 13.182
TEXT_H = 7.0
TEXT_W = 116.419

print("=" * 60)
print("IFC Stamp Test v2 (rectangle + MText)")
print("=" * 60)

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

doc = None
try:
    doc = acad.ActiveDocument
    if 'DIA-01' not in doc.Name:
        doc = acad.Documents.Open(DWG)
        time.sleep(5)
except:
    doc = acad.Documents.Open(DWG)
    time.sleep(5)
print(f"Working with: {doc.Name}")

# Switch to PaperSpace layout
for layout in doc.Layouts:
    if layout.Name.lower() != 'model':
        doc.ActiveLayout = layout
        print(f"Active layout: '{layout.Name}'")
        break
time.sleep(0.5)
space = doc.PaperSpace
print(f"PaperSpace entities before: {space.Count}")

# Calculate rect corners
rect_half_w = RECT_W / 2.0
rect_left   = STAMP_X - rect_half_w
rect_right  = STAMP_X + rect_half_w
rect_bottom = STAMP_Y - TEXT_Y_OFFSET
rect_top    = rect_bottom + RECT_H

print(f"\nStamp layout:")
print(f"  MText insertion: ({STAMP_X}, {STAMP_Y})")
print(f"  Rectangle: ({rect_left:.2f}, {rect_bottom:.2f}) -> ({rect_right:.2f}, {rect_top:.2f})")
print(f"  Rect size: {RECT_W:.1f} x {RECT_H:.1f}")

# Create IFC_STAMP layer
try:
    doc.Layers.Add("IFC_STAMP")
    print("  Created layer 'IFC_STAMP'")
except:
    print("  Layer 'IFC_STAMP' exists")

# Add rectangle
print("\nAdding rectangle...")
try:
    rect_pts = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [rect_left, rect_bottom,
         rect_right, rect_bottom,
         rect_right, rect_top,
         rect_left, rect_top])
    pline = space.AddLightWeightPolyline(rect_pts)
    pline.Closed = True
    pline.Layer = "IFC_STAMP"
    print(f"  Rectangle OK!")
    mn, mx = pline.GetBoundingBox()
    print(f"  BBox: ({float(mn[0]):.2f}, {float(mn[1]):.2f}) -> ({float(mx[0]):.2f}, {float(mx[1]):.2f})")
except Exception as e:
    print(f"  Rectangle FAILED: {e}")

# Add MText
print("\nAdding MText...")
stamp_text = "{\\fArial Narrow|b1;FOR CONSTRUCTION}"
try:
    pt = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [STAMP_X, STAMP_Y, 0.0])
    mtext = space.AddMText(pt, TEXT_W, stamp_text)
    mtext.Height = TEXT_H
    mtext.AttachmentPoint = 2  # TopCenter
    mtext.Layer = "IFC_STAMP"
    doc.Regen(1)
    print(f"  MText OK!")
    print(f"  InsertionPoint: ({mtext.InsertionPoint[0]:.2f}, {mtext.InsertionPoint[1]:.2f})")
    print(f"  Height: {mtext.Height}")
    print(f"  AttachmentPoint: {mtext.AttachmentPoint}")
    mn, mx = mtext.GetBoundingBox()
    print(f"  BBox: ({float(mn[0]):.2f}, {float(mn[1]):.2f}) -> ({float(mx[0]):.2f}, {float(mx[1]):.2f})")
except Exception as e:
    print(f"  MText FAILED: {e}")

print(f"\nPaperSpace entities after: {space.Count}")
print(f"\n{'='*60}")
print("Check AutoCAD now - you should see the rectangle + text stamp.")
print("Close without saving: doc.Close(False)")
print(f"{'='*60}")
input("Press Enter to close...")
doc.Close(False)

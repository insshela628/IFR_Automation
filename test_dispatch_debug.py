"""Debug: test COM dispatch on C-PLN-003 ModelSpace"""
import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import win32com.client
import pythoncom

acad = win32com.client.GetActiveObject("AutoCAD.Application")
print(f"AutoCAD type: {type(acad)}")
print(f"Docs: {acad.Documents.Count}")

# Close any open docs
while acad.Documents.Count > 0:
    acad.Documents.Item(0).Close(False)
    time.sleep(1)

# Open C-PLN-003
dwg = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS\Design\Engineering\1. Drawings\1. Native\GG31-C-PLN-003_Fence & Gate Layout Plan_Rev\GG31-C-PLN-003_Fence & Gate Layout Plan_RevA.dwg"
doc = acad.Documents.Open(dwg)
time.sleep(3)

ms = doc.ModelSpace
print(f"ModelSpace type: {type(ms)}")
print(f"ModelSpace count: {ms.Count}")

# Test: create a simple LightWeightPolyline
pts = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
    [100.0, 100.0, 200.0, 100.0, 200.0, 200.0, 100.0, 200.0])

print("\nTest 1: Direct ModelSpace.AddLightWeightPolyline...")
try:
    pline = ms.AddLightWeightPolyline(pts)
    print(f"  OK! type={type(pline)}")
    pline.Delete()
except Exception as e:
    print(f"  FAILED: {e}")

print("\nTest 2: ModelSpace via Dispatch()...")
try:
    ms2 = win32com.client.Dispatch(ms)
    pline = ms2.AddLightWeightPolyline(pts)
    print(f"  OK! type={type(pline)}")
    pline.Delete()
except Exception as e:
    print(f"  FAILED: {e}")

print("\nTest 3: acad wrapped with Dispatch...")
try:
    acad2 = win32com.client.Dispatch(acad)
    ms3 = acad2.ActiveDocument.ModelSpace
    print(f"  ms3 type: {type(ms3)}")
    pline = ms3.AddLightWeightPolyline(pts)
    print(f"  OK! type={type(pline)}")
    pline.Delete()
except Exception as e:
    print(f"  FAILED: {e}")

print("\nTest 4: layout.Block direct...")
try:
    for layout in doc.Layouts:
        if layout.Name.lower() != 'model':
            blk = layout.Block
            print(f"  Block type: {type(blk)}, name={layout.Name}")
            pline = blk.AddLightWeightPolyline(pts)
            print(f"  OK! type={type(pline)}")
            pline.Delete()
            break
except Exception as e:
    print(f"  FAILED: {e}")

print("\nTest 5: layout.Block with acad wrapped by Dispatch...")
try:
    acad3 = win32com.client.Dispatch(acad)
    doc3 = acad3.ActiveDocument
    for layout in doc3.Layouts:
        if layout.Name.lower() != 'model':
            blk = layout.Block
            print(f"  Block type: {type(blk)}, name={layout.Name}")
            pline = blk.AddLightWeightPolyline(pts)
            print(f"  OK! type={type(pline)}")
            pline.Delete()
            break
except Exception as e:
    print(f"  FAILED: {e}")

doc.Close(False)
print("\nDone.")

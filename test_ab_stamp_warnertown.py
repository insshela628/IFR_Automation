"""Live test: AS BUILT stamp on Warnertown standard frame.

Opens ACE_Standard_Frame_wanertown_UPDATED by MG.dwg,
finds title block, draws AS BUILT stamp with proportional ratios,
and prints coordinates for visual verification.
"""

import sys
sys.path.insert(0, r"D:\1. SOP\SOP_Stage 2 IFR Sync√\V6√")

import time
import win32com.client
import pythoncom
from pathlib import Path

DWG_PATH = (
    r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"
    r"\2.SA\GG-31 Warnertown BESS\Design\Engineering\1. Drawings\1. Native"
    r"\ACE_Standard_Frame_wanertown_UPDATED by MG.dwg"
)

# Same ratios as AsBuiltManager._AB_*
_AB_RECT_W_RATIO = 435.1 / 4205.0
_AB_RECT_H_RATIO = 69.9 / 2970.0
_AB_TEXT_H_RATIO = 27.56 / 2970.0
_AB_CW_RATIO = 7.874 / 4205.0
_AB_X_RIGHT_OFFSET_RATIO = 153.0 / 4205.0
_AB_Y_BOTTOM_OFFSET_RATIO = 366.0 / 2970.0
_AB_COLOUR_RECT_H_RATIO = 109.2 / 2970.0
_AB_COLOUR_GAP_RATIO = 2.0 / 2970.0

STAMP_LAYER = "IFC_STAMP"
STAMP_TEXT = "{\\fArial Narrow|b1;AS BUILT}"
COLOUR_TEXT = "{\\fArial Narrow|b1;DRAWINGS TO BE PRINTED IN COLOUR}"

TB_NAMES = {"ACE-Wanertown_Siyuan", "Coleamablly", "Riverina_tellhow",
            "ACE_TitleBlock_CHINT", "GPA"}


def connect_autocad():
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception:
        acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True
    return acad


def open_dwg(acad, path):
    for doc in acad.Documents:
        if doc.FullName.lower() == path.lower():
            doc.Activate()
            return doc
    doc = acad.Documents.Open(path)
    time.sleep(3)
    return doc


def find_title_block(doc):
    """Find title block INSERT in ModelSpace — iterate directly (small DWG)."""
    ms = doc.ModelSpace
    count = ms.Count
    print(f"  ModelSpace entities: {count}")
    for i in range(count):
        try:
            entity = ms.Item(i)
            if entity.EntityName == 'AcDbBlockReference':
                bname = entity.Name
                if bname in TB_NAMES:
                    return entity
        except Exception:
            continue
    # Fallback: list all block references
    print("  No exact match. Listing all INSERTs:")
    for i in range(count):
        try:
            entity = ms.Item(i)
            if entity.EntityName == 'AcDbBlockReference':
                print(f"    {entity.Name}")
        except Exception:
            continue
    return None


def draw_stamp(doc, block_ref):
    """Draw AS BUILT stamp + COLOUR box, return coordinates for inspection."""
    min_pt, max_pt = block_ref.GetBoundingBox()
    tb_left = float(min_pt[0])
    tb_bottom = float(min_pt[1])
    tb_right = float(max_pt[0])
    tb_top = float(max_pt[1])
    tb_w = tb_right - tb_left
    tb_h = tb_top - tb_bottom

    print(f"\n  Title Block: {block_ref.Name}")
    print(f"  TB BBox: ({tb_left:.1f}, {tb_bottom:.1f}) -> ({tb_right:.1f}, {tb_top:.1f})")
    print(f"  TB Size: {tb_w:.1f} x {tb_h:.1f}")

    rect_w = tb_w * _AB_RECT_W_RATIO
    rect_h = tb_h * _AB_RECT_H_RATIO
    text_h = tb_h * _AB_TEXT_H_RATIO
    cw = tb_w * _AB_CW_RATIO

    stamp_right = tb_right - tb_w * _AB_X_RIGHT_OFFSET_RATIO
    stamp_left = stamp_right - rect_w
    stamp_bottom = tb_bottom + tb_h * _AB_Y_BOTTOM_OFFSET_RATIO
    stamp_top = stamp_bottom + rect_h

    print(f"\n  AS BUILT box:")
    print(f"    Position: ({stamp_left:.1f}, {stamp_bottom:.1f}) -> ({stamp_right:.1f}, {stamp_top:.1f})")
    print(f"    Size: {rect_w:.1f} x {rect_h:.1f}")
    print(f"    CW: {cw:.2f}, TextH: {text_h:.2f}")

    # Ensure layer
    try:
        layer = doc.Layers.Add(STAMP_LAYER)
        layer.color = 7
        layer.LayerOn = True
        layer.Freeze = False
    except Exception:
        try:
            layer = doc.Layers.Item(STAMP_LAYER)
            layer.color = 7
        except Exception:
            pass

    ms = doc.ModelSpace

    # AS BUILT rect
    rect_pts = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [stamp_left, stamp_bottom,
         stamp_right, stamp_bottom,
         stamp_right, stamp_top,
         stamp_left, stamp_top])
    pline = ms.AddLightWeightPolyline(rect_pts)
    pline.Closed = True
    pline.Layer = STAMP_LAYER
    pline.color = 7
    pline.ConstantWidth = cw

    center_pt = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [(stamp_left + stamp_right) / 2.0, (stamp_bottom + stamp_top) / 2.0, 0.0])
    mtext = ms.AddMText(center_pt, rect_w, STAMP_TEXT)
    mtext.Height = text_h
    mtext.AttachmentPoint = 5  # middle center
    mtext.InsertionPoint = center_pt
    mtext.Layer = STAMP_LAYER
    mtext.color = 7

    # COLOUR box (above AS BUILT)
    colour_rect_h = tb_h * _AB_COLOUR_RECT_H_RATIO
    colour_gap = tb_h * _AB_COLOUR_GAP_RATIO
    colour_bottom = stamp_top + colour_gap
    colour_top = colour_bottom + colour_rect_h

    print(f"\n  COLOUR box:")
    print(f"    Position: ({stamp_left:.1f}, {colour_bottom:.1f}) -> ({stamp_right:.1f}, {colour_top:.1f})")
    print(f"    Size: {rect_w:.1f} x {colour_rect_h:.1f}")

    rect_pts2 = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [stamp_left, colour_bottom,
         stamp_right, colour_bottom,
         stamp_right, colour_top,
         stamp_left, colour_top])
    pline2 = ms.AddLightWeightPolyline(rect_pts2)
    pline2.Closed = True
    pline2.Layer = STAMP_LAYER
    pline2.color = 7
    pline2.ConstantWidth = cw

    colour_center = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [(stamp_left + stamp_right) / 2.0, (colour_bottom + colour_top) / 2.0, 0.0])
    mtext2 = ms.AddMText(colour_center, rect_w, COLOUR_TEXT)
    mtext2.Height = text_h
    mtext2.AttachmentPoint = 5
    mtext2.InsertionPoint = colour_center
    mtext2.Layer = STAMP_LAYER
    mtext2.color = 7

    print(f"\n  [OK] Stamp drawn successfully. Check AutoCAD visually.")
    print(f"    Zoom to ({stamp_left-50:.0f},{stamp_bottom-50:.0f}) - ({stamp_right+50:.0f},{colour_top+50:.0f})")
    return True


if __name__ == '__main__':
    print("=" * 60)
    print("  AS BUILT Stamp Test — Warnertown Standard Frame")
    print("=" * 60)

    if not Path(DWG_PATH).exists():
        print(f"\n  ERROR: DWG not found: {DWG_PATH}")
        sys.exit(1)

    print(f"\n  Connecting to AutoCAD...")
    acad = connect_autocad()
    print(f"  AutoCAD version: {acad.Version}")

    print(f"\n  Opening: {Path(DWG_PATH).name}")
    doc = open_dwg(acad, DWG_PATH)
    print(f"  Document opened: {doc.Name}")

    print(f"\n  Finding title block...")
    tb = find_title_block(doc)
    if tb is None:
        print("  ERROR: No title block found!")
        sys.exit(1)

    draw_stamp(doc, tb)

    print(f"\n{'='*60}")
    print("  Done. Do NOT save this file — visual check only.")
    print("  Close without saving after inspecting stamp position.")
    print(f"{'='*60}")

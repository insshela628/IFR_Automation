"""Fix stamp visibility + test COLOUR overlap detection.

Opens bot-output IFC DWGs, fixes stamp, re-exports PDF, verifies.
Key change: PUBLISH from same open doc (no close/reopen).

Usage: python fix_and_test_stamp.py
"""
import sys
import os
import time
import shutil
import tempfile
import re
from pathlib import Path

import win32com.client
import pythoncom

TEST_DIR = Path(os.path.dirname(os.path.abspath(__file__))) / 'test_output'

DWGS = [
    'GG31-E-PLN-003_Communication_Cable_Route_Layout_Plan_Rev0_IFC.dwg',
    'GG31-C-PLN-004_External_Road_site_entrance_Internal_Track_Hardstands_Layout_Plan_Rev0_IFC.dwg',
]

STAMP_LAYER = 'IFC_STAMP'
STAMP_TEXT = r'{\fArial Narrow|b1;FOR CONSTRUCTION}'

REF_TB_W = 841.0
REF_TB_H = 594.0
REF_RECT_W = 110.511
REF_RECT_H = 17.745
REF_TEXT_H = 7.0
REF_X_RIGHT_OFFSET = 29.779
REF_Y_BOTTOM = 73.259
REF_COLOUR_RECT_H = 26.0
REF_COLOUR_GAP = 2.0


def strip_mtext(text):
    if not text:
        return ''
    s = re.sub(r'\{\\f[^;]*;', '', text)
    s = re.sub(r'\{\\H[^;]*;', '', s)
    s = re.sub(r'\{\\[A-Za-z][^;]*;', '', s)
    s = s.replace('}', '')
    s = re.sub(r'\\[PpLlOo]', '', s)
    s = re.sub(r'\\[Aa]\d', '', s)
    return s.strip()


def check_colour_overlap(doc, stamp_left, stamp_right, colour_bottom, colour_top):
    """Check if COLOUR stamp area overlaps with existing entities."""
    skip_layers = {STAMP_LAYER, 'DEFPOINTS', 'ASHADE'}
    # Expand search area
    expand = 15.0
    try:
        ss_name = f'ov_{int(time.time()*1000)%999999}'
        ss = doc.SelectionSets.Add(ss_name)
        pt1 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
                                       [stamp_left - expand, colour_bottom - expand, 0.0])
        pt2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
                                       [stamp_right + expand, colour_top + expand, 0.0])
        ss.Select(1, pt1, pt2)
        for i in range(ss.Count):
            try:
                ent = ss.Item(i)
                if ent.Layer.upper() in skip_layers:
                    continue
                if ent.EntityName == 'AcDbViewport':
                    continue
                ename = ent.EntityName
                info = f"{ename} on layer={ent.Layer}"
                if ename in ('AcDbMText', 'AcDbText'):
                    try:
                        info += f' "{ent.TextString[:30]}"'
                    except:
                        pass
                print(f"      Overlap: {info}")
                ss.Delete()
                return True
            except:
                continue
        ss.Delete()
    except Exception as e:
        print(f"    Overlap check error: {e}")
    return False


def process_dwg(acad, dwg_path):
    """Fix stamp in DWG and re-export PDF."""
    dwg_name = dwg_path.name
    print(f"\n{'='*70}")
    print(f"FILE: {dwg_name}")
    print(f"{'='*70}")

    # Open DWG
    doc = acad.Documents.Open(str(dwg_path))
    for _ in range(20):
        try:
            _ = doc.ModelSpace.Count
            _ = doc.Layouts.Count
            break
        except:
            time.sleep(1)
    time.sleep(2)

    # ── 1. Find title block ──
    tb_ref = None
    tb_candidates = ['ACE-Wanertown_Siyuan', 'Coleamablly']
    ss = doc.SelectionSets.Add(f'tb_{int(time.time()*1000)%999999}')
    ft = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
    fv = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ['INSERT'])
    ss.Select(5, None, None, ft, fv)
    for i in range(ss.Count):
        try:
            ent = ss.Item(i)
            if any(ent.Name.upper() == c.upper() for c in tb_candidates):
                tb_ref = ent
                break
        except:
            continue
    ss.Delete()

    if not tb_ref:
        print("  ERROR: Title block not found")
        doc.Close(False)
        return

    min_pt, max_pt = tb_ref.GetBoundingBox()
    tb_min_x, tb_min_y = float(min_pt[0]), float(min_pt[1])
    tb_max_x, tb_max_y = float(max_pt[0]), float(max_pt[1])
    tb_w = tb_max_x - tb_min_x
    tb_h = tb_max_y - tb_min_y
    print(f"  TB: {tb_w:.0f}x{tb_h:.0f} at ({tb_min_x:.1f},{tb_min_y:.1f})->({tb_max_x:.1f},{tb_max_y:.1f})")

    # ── 2. Scan for COLOUR (before any removal) ──
    has_colour = False
    try:
        ss2 = doc.SelectionSets.Add(f'col_{int(time.time()*1000)%999999}')
        ft2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
        fv2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ['MTEXT'])
        ss2.Select(5, None, None, ft2, fv2)
        for i in range(ss2.Count):
            try:
                ent = ss2.Item(i)
                if ent.Layer.upper() == STAMP_LAYER:
                    continue  # Skip our own entities
                raw = ent.TextString or ''
                txt = strip_mtext(raw).upper()
                if ('COLOUR' in txt or 'COULOUR' in txt or 'COLOR' in txt) and 'PRINT' in txt:
                    has_colour = True
                    fixed = re.sub(r'[Cc][Oo][Uu]?[Ll][Oo][Uu]+[Rr]', 'COLOUR',
                                   raw, flags=re.IGNORECASE)
                    if fixed != raw:
                        ent.TextString = fixed
                        print(f"    COLOUR typo fixed")
                    else:
                        print(f"    COLOUR found on layer={ent.Layer}")
            except:
                continue
        ss2.Delete()
    except Exception as e:
        print(f"    COLOUR scan error: {e}")
    print(f"  has_colour: {has_colour}")

    # ── 3. Remove ALL IFC_STAMP entities ──
    removed = 0
    try:
        ss3 = doc.SelectionSets.Add(f'rm_{int(time.time()*1000)%999999}')
        ft3 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [8])
        fv3 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [STAMP_LAYER])
        ss3.Select(5, None, None, ft3, fv3)
        for i in range(ss3.Count - 1, -1, -1):
            try:
                ss3.Item(i).Delete()
                removed += 1
            except:
                pass
        ss3.Delete()
    except:
        pass
    print(f"  Removed {removed} old IFC_STAMP entities")

    # ── 4. Calculate stamp position ──
    scale = tb_w / REF_TB_W
    scale_y = tb_h / REF_TB_H if tb_h > 0 else scale
    rect_w = REF_RECT_W * scale
    rect_h = REF_RECT_H * scale
    text_h = REF_TEXT_H * scale

    stamp_right = tb_max_x - REF_X_RIGHT_OFFSET * scale
    stamp_left = stamp_right - rect_w
    stamp_bottom = tb_min_y + REF_Y_BOTTOM * scale_y
    stamp_top = stamp_bottom + rect_h
    center_x = (stamp_left + stamp_right) / 2.0
    center_y = (stamp_bottom + stamp_top) / 2.0

    print(f"  Stamp: ({stamp_left:.1f},{stamp_bottom:.1f})->({stamp_right:.1f},{stamp_top:.1f})")

    # ── 5. Ensure layer ──
    try:
        layer = doc.Layers.Add(STAMP_LAYER)
        layer.color = 7
        layer.LayerOn = True
        layer.Freeze = False
    except:
        try:
            layer = doc.Layers.Item(STAMP_LAYER)
            layer.color = 7
            layer.LayerOn = True
            layer.Freeze = False
        except:
            pass

    space = doc.ModelSpace

    # ── 6. Draw FOR CONSTRUCTION ──
    rect_pts = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [stamp_left, stamp_bottom, stamp_right, stamp_bottom,
         stamp_right, stamp_top, stamp_left, stamp_top])
    pline = space.AddLightWeightPolyline(rect_pts)
    pline.Closed = True
    pline.Layer = STAMP_LAYER
    pline.color = 7
    pline.ConstantWidth = 0.5 * scale

    center_pt = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [center_x, center_y, 0.0])
    mtext = space.AddMText(center_pt, rect_w, STAMP_TEXT)
    mtext.Height = text_h
    mtext.AttachmentPoint = 5
    mtext.InsertionPoint = center_pt
    mtext.Layer = STAMP_LAYER
    mtext.color = 7
    print("  FOR CONSTRUCTION drawn (rect + MText)")

    # ── 7. COLOUR decision ──
    colour_gap = REF_COLOUR_GAP * scale
    colour_bottom = stamp_top + colour_gap
    colour_top = colour_bottom + REF_COLOUR_RECT_H * scale

    if has_colour:
        print("  COLOUR: original exists, keeping (skip draw)")
    else:
        overlap = check_colour_overlap(doc, stamp_left, stamp_right, colour_bottom, colour_top)
        if overlap:
            print("  COLOUR: overlap detected, skip draw")
        else:
            print("  COLOUR: no original, no overlap -- NOT drawing (Warnertown default)")
            # Note: For Warnertown, most DWGs have existing COLOUR or overlap
            # If a DWG truly needs COLOUR, it should be added here

    # ── 8. Viewport thaw via VPLAYER ──
    print("  Thawing IFC_STAMP in viewports...")
    for layout in doc.Layouts:
        if layout.Name.lower() == 'model':
            continue
        try:
            doc.ActiveLayout = layout
            time.sleep(1)
            doc.SendCommand(f'-VPLAYER\nThaw\n{STAMP_LAYER}\nAll\n\n')
            time.sleep(2)  # Wait for VPLAYER to complete
            print(f"    VPLAYER thaw: {layout.Name}")
        except Exception as e:
            print(f"    VPLAYER failed for {layout.Name}: {e}")

    # Switch back to Model for good measure
    try:
        for layout in doc.Layouts:
            if layout.Name.lower() == 'model':
                doc.ActiveLayout = layout
                break
    except:
        pass

    # Regen AFTER thaw
    try:
        doc.Regen(1)
    except:
        pass
    time.sleep(2)

    # ── 9. SaveAs to temp (Unicode path safety) ──
    temp_dir = Path(tempfile.gettempdir()) / "ifc_stamp_fix"
    temp_dir.mkdir(parents=True, exist_ok=True)
    temp_dwg = temp_dir / dwg_name
    # Clean old temp files
    for f in temp_dir.iterdir():
        try:
            f.unlink()
        except:
            pass

    doc.SaveAs(str(temp_dwg))
    print(f"  SaveAs: {temp_dwg}")
    time.sleep(2)

    # ── 10. PUBLISH PDF from same open doc ──
    layout_names = []
    for layout in doc.Layouts:
        if layout.Name.lower() != 'model':
            try:
                if layout.Block.Count > 1:
                    layout_names.append((layout.TabOrder, layout.Name))
            except:
                pass
    layout_names.sort()
    print(f"  Layouts: {[n for _, n in layout_names]}")

    if not layout_names:
        print("  ERROR: No layouts for PDF")
        doc.Close(False)
        return

    stem = temp_dwg.stem
    temp_pdf = temp_dir / f"{stem}.pdf"
    dwg_native = str(temp_dwg).replace('/', '\\')
    out_native = str(temp_dir).replace('/', '\\')
    pdf_native = str(temp_pdf).replace('/', '\\')

    dsd_lines = ['[DWF6Version]', 'Ver=1', '[DWF6MinorVersion]', 'MinorVer=1']
    for _, lname in layout_names:
        dsd_lines.extend([
            f'[DWF6Sheet:{stem}-{lname}]',
            f'DWG={dwg_native}',
            f'Layout={lname}',
            'Setup=',
            f'OriginalSheetPath={dwg_native}',
            'Has Plot Port=0',
            'Has3DDWF=0',
        ])
    dsd_lines.extend([
        '[Target]', 'Type=6',
        f'DWF={pdf_native}',
        f'OUT={out_native}',
        'PWD=',
        'PromptForDwfName=FALSE',
        '[PdfOptions]', 'VectorResolution=600', 'RasterResolution=400',
        '[SheetSetProperties]',
        'IsSheetSet=FALSE', 'IsHomogeneous=FALSE',
        'SheetSet Storage File=',
        'AcadProfile=<<Default>>',
        'CategoryCount=0',
        '[AutoCAD Block Information]',
        'IncludeBlockInfo=0',
        'BlockTmplFilePath=',
    ])
    dsd_path = temp_dir / f"{stem}.dsd"
    with open(str(dsd_path), 'w', encoding='utf-8') as f:
        f.write('\n'.join(dsd_lines))

    dsd_native = str(dsd_path).replace('/', '\\')
    print(f"  PUBLISH...")
    doc.SendCommand(f'-PUBLISH\n{dsd_native}\n')

    # Wait for PDF
    for _ in range(45):
        time.sleep(2)
        if temp_pdf.exists() and temp_pdf.stat().st_size > 0:
            break
    else:
        print("  ERROR: PDF not generated after 90s")
        doc.Close(False)
        return

    time.sleep(3)
    pdf_size = temp_pdf.stat().st_size / 1024
    print(f"  PDF generated: {pdf_size:.0f} KB")

    # Copy PDF to test_output
    final_pdf = TEST_DIR / f"{stem}.pdf"
    shutil.copy2(str(temp_pdf), str(final_pdf))
    print(f"  Copied to: {final_pdf.name}")

    doc.Close(False)
    time.sleep(3)

    # ── 11. PDF verification ──
    try:
        import fitz
        pdf = fitz.open(str(final_pdf))
        print(f"\n  PDF Verification ({pdf.page_count} pages):")
        all_ok = True
        for i, page in enumerate(pdf):
            text = page.get_text()
            has_fc = 'FOR CONSTRUCTION' in text
            has_col = ('COLOUR' in text or 'COLOR' in text) and 'PRINT' in text
            status = 'OK' if has_fc else 'MISSING'
            if not has_fc:
                all_ok = False
            print(f"    Page {i+1}: FOR CONSTRUCTION={status}, COLOUR={'yes' if has_col else 'no'}")

            # Stamp screenshot
            pw, ph = page.rect.width, page.rect.height
            clip = fitz.Rect(pw * 0.6, ph * 0.6, pw, ph)
            pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=clip)
            png = TEST_DIR / f"{stem}_p{i+1}_stamp.png"
            pix.save(str(png))

        pdf.close()
        print(f"\n  RESULT: {'PASS' if all_ok else 'FAIL'}")
    except ImportError:
        print("  PyMuPDF not installed, skip verification")


def main():
    pythoncom.CoInitialize()
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True

    for dwg_name in DWGS:
        dwg_path = TEST_DIR / dwg_name
        if not dwg_path.exists():
            print(f"\nSKIP: {dwg_name} not found")
            continue

        try:
            process_dwg(acad, dwg_path)
        except Exception as e:
            print(f"\nERROR: {e}")
            import traceback
            traceback.print_exc()

        # Close all docs between files
        try:
            while acad.Documents.Count > 0:
                try:
                    acad.Documents.Item(0).Close(False)
                except:
                    break
                time.sleep(1)
        except:
            pass
        time.sleep(5)

    print(f"\n{'='*70}")
    print("ALL DONE")


if __name__ == '__main__':
    main()

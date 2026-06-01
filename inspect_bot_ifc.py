"""Inspect bot-output IFC DWGs to diagnose stamp issues.

Opens each DWG in AutoCAD COM, checks IFC_STAMP layer entities,
COLOUR presence, and stamp positions.
"""
import sys
import os
import time
import win32com.client
import pythoncom

TEST_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_output')

DWGS = [
    'GG31-E-PLN-003_Communication_Cable_Route_Layout_Plan_Rev0_IFC.dwg',
    'GG31-C-PLN-004_External_Road_site_entrance_Internal_Track_Hardstands_Layout_Plan_Rev0_IFC.dwg',
]

def inspect_dwg(acad, dwg_path):
    print(f"\n{'='*70}")
    print(f"FILE: {os.path.basename(dwg_path)}")
    print(f"{'='*70}")

    doc = acad.Documents.Open(dwg_path)
    # Wait for load
    for _ in range(20):
        try:
            _ = doc.ModelSpace.Count
            _ = doc.Layouts.Count
            break
        except:
            time.sleep(1)
    time.sleep(2)

    # === 1. Find title block ===
    tb_name_candidates = ['ACE-Wanertown_Siyuan', 'Coleamablly']
    tb_found = None
    tb_ref = None

    ss_name = f'insp_{int(time.time()*1000)%999999}'
    ss = doc.SelectionSets.Add(ss_name)
    ft = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
    fv = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ['INSERT'])
    ss.Select(5, None, None, ft, fv)
    for i in range(ss.Count):
        try:
            ent = ss.Item(i)
            for cand in tb_name_candidates:
                if ent.Name.upper() == cand.upper():
                    tb_found = ent.Name
                    tb_ref = ent
                    break
            if tb_found:
                break
        except:
            continue
    ss.Delete()

    if tb_ref:
        min_pt, max_pt = tb_ref.GetBoundingBox()
        tb_min_x, tb_min_y = float(min_pt[0]), float(min_pt[1])
        tb_max_x, tb_max_y = float(max_pt[0]), float(max_pt[1])
        tb_w = tb_max_x - tb_min_x
        tb_h = tb_max_y - tb_min_y
        print(f"\n  Title block: {tb_found}")
        print(f"  TB bbox: ({tb_min_x:.1f},{tb_min_y:.1f}) -> ({tb_max_x:.1f},{tb_max_y:.1f})")
        print(f"  TB size: {tb_w:.1f} x {tb_h:.1f}")
    else:
        print(f"\n  Title block: NOT FOUND")

    # === 2. Check IFC_STAMP layer entities ===
    print(f"\n  --- IFC_STAMP layer entities ---")
    stamp_count = 0
    stamp_mtexts = []
    stamp_plines = []

    for layout in doc.Layouts:
        block = layout.Block
        for i in range(block.Count):
            try:
                ent = block.Item(i)
                if ent.Layer.upper() == 'IFC_STAMP':
                    stamp_count += 1
                    ename = ent.EntityName
                    if ename == 'AcDbMText':
                        txt = ent.TextString[:80]
                        ip = ent.InsertionPoint
                        h = ent.Height
                        c = ent.color
                        ap = ent.AttachmentPoint
                        stamp_mtexts.append({
                            'text': txt, 'x': float(ip[0]), 'y': float(ip[1]),
                            'height': h, 'color': c, 'attach': ap, 'layout': layout.Name
                        })
                    elif 'Polyline' in ename:
                        try:
                            coords = list(ent.Coordinates)
                            xs = coords[0::2]
                            ys = coords[1::2]
                            stamp_plines.append({
                                'min_x': min(xs), 'min_y': min(ys),
                                'max_x': max(xs), 'max_y': max(ys),
                                'width': ent.ConstantWidth, 'color': ent.color,
                                'layout': layout.Name
                            })
                        except:
                            stamp_plines.append({'error': str(ent.EntityName), 'layout': layout.Name})
            except:
                continue

    print(f"  Total entities on IFC_STAMP: {stamp_count}")
    for mt in stamp_mtexts:
        print(f"    MText: \"{mt['text']}\"")
        print(f"      pos=({mt['x']:.1f},{mt['y']:.1f}), h={mt['height']:.2f}, "
              f"color={mt['color']}, attach={mt['attach']}, layout={mt['layout']}")
    for pl in stamp_plines:
        if 'error' in pl:
            print(f"    Polyline: error reading - {pl['error']} (layout={pl['layout']})")
        else:
            print(f"    Polyline: ({pl['min_x']:.1f},{pl['min_y']:.1f})->({pl['max_x']:.1f},{pl['max_y']:.1f}), "
                  f"width={pl['width']:.2f}, color={pl['color']}, layout={pl['layout']}")

    # === 3. Scan ALL MTEXT for COLOUR / FOR CONSTRUCTION / FOR REVIEW ===
    print(f"\n  --- Stamp-related MTEXT (all layers) ---")
    ss2_name = f'scan_{int(time.time()*1000)%999999}'
    ss2 = doc.SelectionSets.Add(ss2_name)
    ft2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
    fv2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ['MTEXT'])
    ss2.Select(5, None, None, ft2, fv2)
    for i in range(ss2.Count):
        try:
            ent = ss2.Item(i)
            txt = ent.TextString
            txt_upper = txt.upper()
            if any(kw in txt_upper for kw in ('COLOUR', 'COLOR', 'CONSTRUCTION', 'REVIEW', 'PRINTED')):
                ip = ent.InsertionPoint
                print(f"    Layer={ent.Layer}, color={ent.color}, h={ent.Height:.2f}")
                print(f"      pos=({float(ip[0]):.1f},{float(ip[1]):.1f}), text=\"{txt[:80]}\"")
        except:
            continue
    ss2.Delete()

    # === 4. Check IFR block references still present ===
    print(f"\n  --- IFR block references ---")
    ss3_name = f'ifr_{int(time.time()*1000)%999999}'
    ss3 = doc.SelectionSets.Add(ss3_name)
    ft3 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
    fv3 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ['INSERT'])
    ss3.Select(5, None, None, ft3, fv3)
    ifr_found = 0
    for i in range(ss3.Count):
        try:
            ent = ss3.Item(i)
            name = ent.Name
            if 'IFR' in name.upper() or 'REVIEW' in name.upper():
                ip = ent.InsertionPoint
                print(f"    Block: {name} at ({float(ip[0]):.1f},{float(ip[1]):.1f})")
                ifr_found += 1
        except:
            continue
    ss3.Delete()
    if ifr_found == 0:
        print(f"    (none)")

    # === 5. Check viewport freeze on IFC_STAMP ===
    print(f"\n  --- Viewport IFC_STAMP freeze status ---")
    for layout in doc.Layouts:
        if layout.Name.lower() == 'model':
            continue
        block = layout.Block
        for i in range(block.Count):
            try:
                ent = block.Item(i)
                if ent.EntityName == 'AcDbViewport':
                    try:
                        frozen = list(ent.GetFrozenLayers() or [])
                        if 'IFC_STAMP' in frozen:
                            print(f"    {layout.Name}: IFC_STAMP is FROZEN in viewport!")
                        else:
                            print(f"    {layout.Name}: IFC_STAMP is thawed (OK)")
                    except:
                        print(f"    {layout.Name}: cannot read frozen layers")
            except:
                continue

    doc.Close(False)
    print(f"\n  Done: {os.path.basename(dwg_path)}")


if __name__ == '__main__':
    pythoncom.CoInitialize()
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True

    for dwg_name in DWGS:
        dwg_path = os.path.join(TEST_DIR, dwg_name)
        if os.path.exists(dwg_path):
            try:
                inspect_dwg(acad, dwg_path)
            except Exception as e:
                print(f"\nERROR inspecting {dwg_name}: {e}")
        else:
            print(f"\nNOT FOUND: {dwg_path}")

    print(f"\n{'='*70}")
    print("Inspection complete.")

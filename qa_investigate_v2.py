"""
QA Agent B - DWG Investigation v3
- Wait for AutoCAD to be ready
- GPA-A1 title block confirmed: inspect its attributes
- Check layouts for stamp entities
"""
import pythoncom
import win32com.client
import sys
import traceback
import os
import time

pythoncom.CoInitialize()

RESULT_PATH = r"D:\1. SOP\QA_Result_v3.txt"
OUT = []

def p(msg=""):
    OUT.append(str(msg))
    print(msg, flush=True)

def save():
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(OUT))

try:
    p("=== QA Agent B v3: NSW153-C-PLN-002 Equipment Scope ===")

    dwg_path = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\1.NSW\Coleambally #2\Design\Engineering\1. Drawings\1. Native\NSW153-C-PLN-002_Site Coordinates design (Equipment Scope)_Rev\NSW153-C-PLN-002_Site Coordinates design (Equipment Scope)_RevE_IFC.dwg"

    # Connect to AutoCAD
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        p("[INFO] Connected to existing AutoCAD.")
    except Exception:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        p("[INFO] Launched new AutoCAD.")

    acad.Visible = True

    # Wait until AutoCAD is not busy
    for attempt in range(30):
        try:
            _ = acad.Version
            p(f"[INFO] AutoCAD ready (attempt {attempt+1}). Version: {acad.Version}")
            break
        except:
            p(f"[INFO] Waiting for AutoCAD... (attempt {attempt+1})")
            time.sleep(2)

    p(f"  Open documents count: {acad.Documents.Count}")

    # Find the Equipment Scope doc by name
    doc = None
    for i in range(acad.Documents.Count):
        d = acad.Documents.Item(i)
        dname = d.Name.lower()
        if "equipment scope" in dname or ("pln-002" in dname and "fencing" not in dname):
            doc = d
            p(f"[INFO] Found open doc: {d.Name}")
            break

    if doc is None:
        p("[INFO] Opening target DWG via Documents.Open (read-only)...")
        try:
            doc = acad.Documents.Open(dwg_path, True)
            p(f"[INFO] Opened (read-only): {doc.Name}")
            time.sleep(5)
        except Exception as ex1:
            p(f"[WARN] Read-only open failed: {ex1}")
            try:
                doc = acad.Documents.Open(dwg_path, False)
                p(f"[INFO] Opened (writable): {doc.Name}")
                time.sleep(5)
            except Exception as ex2:
                p(f"[WARN] Open failed: {ex2}")

    if doc is None:
        raise RuntimeError("Could not find or open the Equipment Scope DWG.")

    # Re-fetch the document reference directly from the Documents collection
    # (COM proxy may be stale after Open)
    p("  Re-fetching doc reference from Documents collection...")
    for i in range(acad.Documents.Count):
        d = acad.Documents.Item(i)
        dname = d.Name.lower()
        if "equipment scope" in dname or ("pln-002" in dname and "fencing" not in dname):
            doc = d
            p(f"[INFO] Fresh doc ref: {d.Name}")
            break

    # Make it active
    try:
        doc.Activate()
        time.sleep(2)
    except Exception as ex:
        p(f"[WARN] Activate failed: {ex}")

    p(f"Active document: {doc.Name}")
    p()

    # ===== STEP 1: Block Definitions - find GPA-A1 and inspect =====
    p("=== STEP 1: GPA-A1 Block Definition inspection ===")
    blk_defs = doc.Blocks
    gpa_block = None
    for i in range(blk_defs.Count):
        try:
            blk = blk_defs.Item(i)
            if blk.Name == "GPA-A1":
                gpa_block = blk
                p(f"  Found GPA-A1 definition at index {i}")
                p(f"  Entity count in definition: {blk.Count}")
                break
        except:
            pass

    if gpa_block is None:
        p("  GPA-A1 block definition NOT FOUND. Searching all blocks for attributes...")
        for i in range(blk_defs.Count):
            try:
                blk = blk_defs.Item(i)
                bname = blk.Name
                if bname.startswith("*") or bname.startswith("A$C"):
                    continue
                for j in range(min(blk.Count, 200)):
                    try:
                        ent = blk.Item(j)
                        if ent.EntityName == "AcDbAttributeDefinition":
                            p(f"    Block {bname!r}: AttributeDefinition tag={ent.TagString!r}")
                    except:
                        pass
            except:
                pass
    else:
        p(f"  Inspecting entities inside GPA-A1 definition:")
        for j in range(gpa_block.Count):
            try:
                ent = gpa_block.Item(j)
                ename = ent.EntityName
                if ename == "AcDbAttributeDefinition":
                    p(f"    ATTR DEF: tag={ent.TagString!r}  default={ent.TextString!r}  layer={ent.Layer!r}")
                elif ename == "AcDbText":
                    p(f"    TEXT: {ent.TextString!r}  layer={ent.Layer!r}")
                elif ename == "AcDbMText":
                    p(f"    MTEXT: {ent.TextString[:80]!r}  layer={ent.Layer!r}")
            except:
                pass
    p()

    # ===== STEP 2: Search for GPA-A1 insert in ALL spaces =====
    p("=== STEP 2: Find GPA-A1 block insert (ModelSpace + Layouts) ===")

    gpa_insert = None

    def scan_space(space_obj, space_name):
        global gpa_insert
        try:
            count = space_obj.Count
            for ei in range(count):
                try:
                    ent = space_obj.Item(ei)
                    ename = ent.EntityName
                    if ename in ("AcDbBlockReference", "AcDbMInsertBlock"):
                        if ent.Name == "GPA-A1":
                            try:
                                ins = ent.InsertionPoint
                                pos = f"({ins[0]:.1f},{ins[1]:.1f})"
                            except:
                                pos = "N/A"
                            p(f"  GPA-A1 INSERT in {space_name!r}: pos={pos}")
                            if gpa_insert is None:
                                gpa_insert = ent
                except:
                    pass
        except Exception as ex:
            p(f"  scan_space({space_name}) error: {ex}")

    # ModelSpace
    scan_space(doc.ModelSpace, "*Model_Space")

    # All Layouts/PaperSpace
    try:
        layouts = doc.Layouts
        p(f"  Total layouts: {layouts.Count}")
        for li in range(layouts.Count):
            try:
                layout = layouts.Item(li)
                lname = layout.Name
                ps = layout.Block
                scan_space(ps, lname)
            except Exception as ex:
                p(f"  Layout[{li}] error: {ex}")
    except Exception as ex:
        p(f"  Layouts collection error: {ex}")

    p()

    # ===== STEP 3: Attributes of the found GPA-A1 insert =====
    p("=== STEP 3: GPA-A1 Insert Attributes ===")
    if gpa_insert is not None:
        try:
            attribs = gpa_insert.GetAttributes()
            p(f"  Attribute count: {len(attribs)}")
            for att in attribs:
                p(f"    TAG={att.TagString!r:45s}  VALUE={att.TextString!r}")
        except Exception as ex:
            p(f"  ERROR reading attributes: {ex}")
            traceback.print_exc()
    else:
        p("  GPA-A1 insert NOT FOUND in any space. Cannot read attributes.")
    p()

    # ===== STEP 4: Stamp entities scan across all spaces =====
    p("=== STEP 4: Stamp / AS BUILT / Colour / IFC entities ===")
    stamp_kw = ["colour", "color", "as built", "as-built", "for construction", "stamp", "issued", "ifc"]
    found_stamps = []

    def scan_for_stamps(space_obj, space_name):
        try:
            for ei in range(space_obj.Count):
                try:
                    ent = space_obj.Item(ei)
                    ename = ent.EntityName
                    if ename in ("AcDbMText", "AcDbText"):
                        txt = ent.TextString
                        layer = ent.Layer
                        try:
                            ins = ent.InsertionPoint
                            pos = f"({ins[0]:.1f},{ins[1]:.1f})"
                        except:
                            pos = "N/A"
                        if any(kw in txt.lower() for kw in stamp_kw):
                            found_stamps.append((ename, txt[:150], layer, pos, space_name))
                    elif ename in ("AcDbBlockReference",):
                        bname = ent.Name
                        bname_lower = bname.lower()
                        if any(kw in bname_lower for kw in ["stamp", "colour", "color", "asbuilt", "as_built"]):
                            try:
                                ins = ent.InsertionPoint
                                pos = f"({ins[0]:.1f},{ins[1]:.1f})"
                            except:
                                pos = "N/A"
                            found_stamps.append((ename, bname, "BLOCK", pos, space_name))
                except:
                    pass
        except Exception as ex:
            p(f"  scan_for_stamps({space_name}) error: {ex}")

    scan_for_stamps(doc.ModelSpace, "*Model_Space")

    try:
        for li in range(doc.Layouts.Count):
            try:
                layout = doc.Layouts.Item(li)
                scan_for_stamps(layout.Block, layout.Name)
            except:
                pass
    except:
        pass

    if found_stamps:
        for s in found_stamps:
            p(f"  [{s[0]}] Space={s[4]!r} Layer={s[2]!r} Pos={s[3]}")
            p(f"    Content: {s[1]!r}")
    else:
        p("  No stamp/AS BUILT/colour entities found.")
    p()

    # ===== STEP 5: IFC_STAMP or stamp layers =====
    p("=== STEP 5: Entities on stamp-related layers ===")
    stamp_layer_kw = ["stamp", "ifc_stamp", "as_built", "colour", "color"]
    found_stamp_layer = []

    def scan_for_stamp_layers(space_obj, space_name):
        try:
            for ei in range(space_obj.Count):
                try:
                    ent = space_obj.Item(ei)
                    layer = ent.Layer
                    if any(kw in layer.lower() for kw in stamp_layer_kw):
                        ename = ent.EntityName
                        try:
                            ins = ent.InsertionPoint
                            pos = f"({ins[0]:.1f},{ins[1]:.1f})"
                        except:
                            pos = "N/A"
                        found_stamp_layer.append((ename, layer, pos, space_name))
                        p(f"  [{ename}] Layer={layer!r} Space={space_name!r} Pos={pos}")
                except:
                    pass
        except:
            pass

    scan_for_stamp_layers(doc.ModelSpace, "*Model_Space")
    try:
        for li in range(doc.Layouts.Count):
            try:
                layout = doc.Layouts.Item(li)
                scan_for_stamp_layers(layout.Block, layout.Name)
            except:
                pass
    except:
        pass

    if not found_stamp_layer:
        p("  None found.")
    p()

    # ===== STEP 6: Locked layers =====
    p("=== STEP 6: Locked Layers ===")
    locked = []
    try:
        for j in range(doc.Layers.Count):
            lyr = doc.Layers.Item(j)
            if lyr.Lock:
                locked.append(lyr.Name)
                p(f"  LOCKED: {lyr.Name!r}")
    except Exception as ex:
        p(f"  ERROR: {ex}")
    if not locked:
        p("  No locked layers.")
    p()

    # ===== FINAL SUMMARY =====
    p("=" * 70)
    p("FINAL SUMMARY")
    title_block_name = "GPA-A1" if (gpa_insert is not None or gpa_block is not None) else "NOT FOUND"
    p(f"1. Title block block name: {title_block_name!r}")
    p(f"   GPA-A1 definition found: {gpa_block is not None}")
    p(f"   GPA-A1 insert found: {gpa_insert is not None}")
    if gpa_insert is not None:
        try:
            attribs = gpa_insert.GetAttributes()
            p(f"2. Attribute tags + values ({len(attribs)} total):")
            for att in attribs:
                p(f"   {att.TagString!r:45s} = {att.TextString!r}")
        except Exception as ex:
            p(f"2. ERROR: {ex}")
    else:
        p("2. Cannot read attributes - insert not found")
    p(f"3. Stamp entities: {len(found_stamps)}")
    for s in found_stamps:
        p(f"   {s[0]} | {s[4]} | Layer={s[2]} | {s[1][:80]!r}")
    p(f"4. Stamp-layer entities: {len(found_stamp_layer)}")
    p(f"5. Locked layers: {locked}")
    p("=" * 70)

except Exception as e:
    p(f"\n[FATAL ERROR] {e}")
    traceback.print_exc()
finally:
    save()
    print(f"\n[Results saved to: {RESULT_PATH}]")
    pythoncom.CoUninitialize()
    print("[DONE] DWG NOT closed/saved.")

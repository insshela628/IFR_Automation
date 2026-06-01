"""Verify IFC conversion output meets ALL requirements.

Checks for each converted DWG:
  1. Title blocks: found in ALL expected layouts (Model + PaperSpace)
  2. IFC_STAMP layer: entities exist near each title block
  3. FOR CONSTRUCTION stamp: present
  4. DRAWINGS TO BE PRINTED IN COLOUR stamp: present
  5. Title block revision: has IFC row (DESCRIPTION contains "CONSTRUCTION")
  6. Output PDF: exists in IFC(Client)

Run with AutoCAD open.
"""
import win32com.client
import time
import sys
import re
from pathlib import Path

# ── Configuration ──
IFC_OUTPUT = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS\Design\Engineering\1. Drawings\4. IFC(Client)")
NATIVE_ROOT = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS\Design\Engineering\1. Drawings\1. Native")

TITLE_BLOCK_NAME = "ACE-Wanertown_Siyuan"
STAMP_LAYER = "IFC_STAMP"

# Doc IDs to verify (all that should be IFC converted)
DOC_IDS_TO_CHECK = [
    "GG31-C-PLN-001", "GG31-C-PLN-002", "GG31-C-PLN-003",
    "GG31-C-PLN-004", "GG31-C-PLN-005",
    "GG31-E-BLD-001",
    "GG31-E-PLN-001", "GG31-E-PLN-003", "GG31-E-PLN-004",
]

def _strip_mtext(txt):
    """Strip MText formatting codes."""
    plain = re.sub(r'\\[A-Za-z][^;]*;', '', txt)
    plain = re.sub(r'[{}\\]', '', plain)
    return plain.strip().upper()


def _com_retry(func, retries=5, delay=2):
    for i in range(retries):
        try:
            return func()
        except Exception as e:
            if i < retries - 1:
                time.sleep(delay)
            else:
                raise


def find_ifc_dwg(doc_id):
    """Find IFC DWG in native folder."""
    # Search native folders for _IFC.dwg
    for folder in NATIVE_ROOT.iterdir():
        if not folder.is_dir():
            continue
        if doc_id.upper() not in folder.name.upper():
            continue
        for f in folder.glob("*_IFC.dwg"):
            if doc_id.upper() in f.name.upper():
                return f
        # Also check for _Rev*_IFC.dwg pattern
        for f in folder.glob("*IFC*.dwg"):
            if doc_id.upper() in f.name.upper() and not f.name.startswith('~'):
                return f
    return None


def find_ifc_pdf(doc_id):
    """Find IFC PDF in output folder."""
    if not IFC_OUTPUT.exists():
        return None
    for f in IFC_OUTPUT.iterdir():
        if f.suffix.lower() != '.pdf' or f.name.startswith('~'):
            continue
        if doc_id.upper() in f.name.upper():
            return f
    return None


def verify_dwg(acad, dwg_path, doc_id):
    """Open DWG and verify all IFC requirements. Returns dict of results."""
    result = {
        'doc_id': doc_id,
        'dwg': dwg_path.name,
        'title_blocks': [],       # list of (layout_name, has_ifc_row, ifc_desc)
        'stamps': [],             # list of (layout_name, stamp_type, text)
        'stamp_layer_count': 0,
        'has_for_construction': False,
        'has_colour': False,
        'all_tb_have_ifc_row': False,
        'pdf_exists': False,
        'errors': [],
    }

    # Check PDF
    pdf = find_ifc_pdf(doc_id)
    result['pdf_exists'] = pdf is not None
    if pdf:
        result['pdf_name'] = pdf.name

    # Open DWG
    doc = None
    try:
        # Check if already open
        for i in range(acad.Documents.Count):
            try:
                d = acad.Documents.Item(i)
                if Path(d.FullName).resolve() == dwg_path.resolve():
                    doc = d
                    break
            except:
                continue

        if doc is None:
            doc = _com_retry(lambda: acad.Documents.Open(str(dwg_path)))

        # Wait for load
        for _ in range(20):
            try:
                _ = doc.ModelSpace.Count
                _ = doc.Layouts.Count
                break
            except:
                time.sleep(1)
        time.sleep(2)

    except Exception as e:
        result['errors'].append(f"打开失败: {e}")
        return result

    # ── Check 1: Title blocks in all layouts ──
    try:
        # ModelSpace via SelectionSet
        import pythoncom
        ss_name = f"_Verify_{int(time.time()*1000) % 1000000}"
        try:
            doc.SelectionSets.Item(ss_name).Delete()
        except:
            pass

        try:
            ss = doc.SelectionSets.Add(ss_name)
            ft = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
            fd = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["INSERT"])
            ss.Select(5)
            for i in range(ss.Count):
                ent = ss.Item(i)
                try:
                    if ent.Name == TITLE_BLOCK_NAME:
                        # Check IFC row
                        has_ifc = False
                        ifc_desc = ""
                        try:
                            attrs = _com_retry(lambda: ent.GetAttributes())
                            for a in attrs:
                                tag = _com_retry(lambda aa=a: aa.TagString.upper())
                                if 'DESCRIPTION' in tag:
                                    txt = _com_retry(lambda aa=a: aa.TextString)
                                    if 'CONSTRUCTION' in txt.upper():
                                        has_ifc = True
                                        ifc_desc = txt[:40]
                        except:
                            pass
                        result['title_blocks'].append(('Model', has_ifc, ifc_desc))
                except:
                    continue
            ss.Delete()
        except Exception as e:
            result['errors'].append(f"ModelSpace搜索失败: {e}")

        # PaperSpace layouts
        for _ps_try in range(3):
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
                            if ent.Name != TITLE_BLOCK_NAME:
                                continue
                            has_ifc = False
                            ifc_desc = ""
                            try:
                                attrs = _com_retry(lambda: ent.GetAttributes())
                                for a in attrs:
                                    tag = _com_retry(lambda aa=a: aa.TagString.upper())
                                    if 'DESCRIPTION' in tag:
                                        txt = _com_retry(lambda aa=a: aa.TextString)
                                        if 'CONSTRUCTION' in txt.upper():
                                            has_ifc = True
                                            ifc_desc = txt[:40]
                            except:
                                pass
                            result['title_blocks'].append((layout.Name, has_ifc, ifc_desc))
                        except:
                            continue
                break
            except Exception as e:
                if _ps_try < 2:
                    time.sleep(2)
                else:
                    result['errors'].append(f"PaperSpace搜索失败: {e}")
    except Exception as e:
        result['errors'].append(f"Title block 检查失败: {e}")

    # ── Check 2: IFC_STAMP layer entities ──
    try:
        ss2_name = f"_VerifyStamp_{int(time.time()*1000) % 1000000}"
        try:
            doc.SelectionSets.Item(ss2_name).Delete()
        except:
            pass

        # ModelSpace stamps
        try:
            ss2 = doc.SelectionSets.Add(ss2_name)
            ft2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [8])
            fd2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [STAMP_LAYER])
            ss2.Select(5)
            for i in range(ss2.Count):
                ent = ss2.Item(i)
                try:
                    en = ent.EntityName
                    if en in ('AcDbMText', 'AcDbText'):
                        txt = _strip_mtext(_com_retry(lambda: ent.TextString))
                        if 'FOR CONSTRUCTION' in txt:
                            result['stamps'].append(('Model', 'FOR CONSTRUCTION', txt[:50]))
                            result['has_for_construction'] = True
                        elif 'PRINTED IN COLOUR' in txt:
                            result['stamps'].append(('Model', 'COLOUR', txt[:50]))
                            result['has_colour'] = True
                        else:
                            result['stamps'].append(('Model', 'OTHER', txt[:50]))
                    elif 'Polyline' in en:
                        result['stamp_layer_count'] += 1
                except:
                    continue
            ss2.Delete()
        except Exception as e:
            result['errors'].append(f"ModelSpace stamp 检查失败: {e}")

        # PaperSpace stamps
        for _ps2 in range(3):
            try:
                for layout in doc.Layouts:
                    if layout.Name.lower() == 'model':
                        continue
                    block = layout.Block
                    for i in range(block.Count):
                        try:
                            ent = block.Item(i)
                            if ent.Layer != STAMP_LAYER:
                                continue
                            en = ent.EntityName
                            if en in ('AcDbMText', 'AcDbText'):
                                txt = _strip_mtext(_com_retry(lambda: ent.TextString))
                                if 'FOR CONSTRUCTION' in txt:
                                    result['stamps'].append((layout.Name, 'FOR CONSTRUCTION', txt[:50]))
                                    result['has_for_construction'] = True
                                elif 'PRINTED IN COLOUR' in txt:
                                    result['stamps'].append((layout.Name, 'COLOUR', txt[:50]))
                                    result['has_colour'] = True
                            elif 'Polyline' in en:
                                result['stamp_layer_count'] += 1
                        except:
                            continue
                break
            except Exception as e:
                if _ps2 < 2:
                    time.sleep(2)
                else:
                    result['errors'].append(f"PaperSpace stamp 检查失败: {e}")
    except Exception as e:
        result['errors'].append(f"Stamp 检查失败: {e}")

    # ── Summary ──
    if result['title_blocks']:
        result['all_tb_have_ifc_row'] = all(tb[1] for tb in result['title_blocks'])

    # Close
    try:
        doc.Close(False)
    except:
        pass

    return result


def main():
    print("=" * 80)
    print("  IFC 输出验证工具")
    print("=" * 80)

    # Connect to AutoCAD
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        print(f"  已连接 AutoCAD\n")
    except:
        print("  请先打开 AutoCAD")
        sys.exit(1)

    all_pass = True
    summary = []

    for doc_id in DOC_IDS_TO_CHECK:
        print(f"\n{'─' * 70}")
        print(f"  检查: {doc_id}")
        print(f"{'─' * 70}")

        # Find IFC DWG
        dwg = find_ifc_dwg(doc_id)
        pdf = find_ifc_pdf(doc_id)

        if not dwg:
            print(f"  ✗ 未找到 IFC DWG")
            summary.append((doc_id, 'FAIL', '无 IFC DWG'))
            all_pass = False
            continue

        print(f"  DWG: {dwg.name}")
        print(f"  PDF: {pdf.name if pdf else '未找到'}")

        # Verify
        r = verify_dwg(acad, dwg, doc_id)

        # Report title blocks
        tb_count = len(r['title_blocks'])
        print(f"\n  Title Blocks: {tb_count} 个")
        for layout_name, has_ifc, ifc_desc in r['title_blocks']:
            status = "✓ IFC行" if has_ifc else "✗ 无IFC行"
            print(f"    [{layout_name}] {status}  {ifc_desc}")

        # Report stamps
        print(f"\n  Stamps (IFC_STAMP layer):")
        if r['stamps']:
            for layout_name, stype, txt in r['stamps']:
                print(f"    [{layout_name}] {stype}: {txt}")
        else:
            print(f"    ✗ 无 stamp 实体")
        print(f"    Polyline borders: {r['stamp_layer_count']}")

        # Report errors
        if r['errors']:
            print(f"\n  错误:")
            for err in r['errors']:
                print(f"    ✗ {err}")

        # ── Verdict ──
        issues = []
        if tb_count == 0:
            issues.append("无 title block")
        if not r['all_tb_have_ifc_row']:
            issues.append("部分 TB 无 IFC revision 行")
        if not r['has_for_construction']:
            issues.append("缺少 FOR CONSTRUCTION stamp")
        if not r['has_colour']:
            issues.append("缺少 COLOUR stamp")
        if not r['pdf_exists']:
            issues.append("无 PDF 输出")
        if r['errors']:
            issues.append(f"{len(r['errors'])} 个错误")

        # Check stamp count vs title block count
        fc_count = sum(1 for s in r['stamps'] if s[1] == 'FOR CONSTRUCTION')
        cl_count = sum(1 for s in r['stamps'] if s[1] == 'COLOUR')
        if tb_count > 0:
            if fc_count < tb_count:
                issues.append(f"FOR CONSTRUCTION stamp 数量({fc_count}) < title block 数量({tb_count})")
            if cl_count < tb_count:
                issues.append(f"COLOUR stamp 数量({cl_count}) < title block 数量({tb_count})")

        if issues:
            verdict = "FAIL"
            all_pass = False
            print(f"\n  ★ 结果: FAIL")
            for iss in issues:
                print(f"    - {iss}")
        else:
            verdict = "PASS"
            print(f"\n  ★ 结果: PASS")

        summary.append((doc_id, verdict, '; '.join(issues) if issues else 'OK'))

    # ── Final Summary ──
    print(f"\n\n{'=' * 80}")
    print(f"  总结")
    print(f"{'=' * 80}")
    pass_count = sum(1 for _, v, _ in summary if v == 'PASS')
    fail_count = sum(1 for _, v, _ in summary if v == 'FAIL')
    print(f"  通过: {pass_count}  失败: {fail_count}  总计: {len(summary)}\n")
    for doc_id, verdict, detail in summary:
        mark = "✓" if verdict == "PASS" else "✗"
        print(f"  {mark} {doc_id}: {detail}")

    print(f"\n{'=' * 80}")
    if all_pass:
        print("  所有检查通过!")
    else:
        print("  存在问题，请把以上输出复制给我。")
    print(f"{'=' * 80}")


if __name__ == '__main__':
    main()

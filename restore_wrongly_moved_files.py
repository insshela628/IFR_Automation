#!/usr/bin/env python3
"""
复原被错误移动到 SS/Superceded 文件夹的文件。

问题原因: version_manager v4 / ifr_automation v8-v9 的 process_folder()
只保留每个 rev_type (IFR/IFC) 中 mtime 最新的单个文件，
导致最新的 .docx 和唯一的辅助 .xlsx 被移到 SS。

此脚本将被错误移动的文件移回到正确的位置。
运行前会显示所有计划的移动操作，需确认后才执行。
"""

import shutil
from pathlib import Path

PROJECT_ROOT = Path(
    r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"
    r"\2.SA\GG-31 Warnertown BESS\Design\Engineering"
    r"\2. Calcs & Reports\Reports"
)

# (source_in_ss, destination_parent) — source 是 SS/Superceded 中的完整路径
RESTORE_LIST = []

# ── Civil & Structure ──────────────────────────────────────────────────
CIVIL = PROJECT_ROOT / "Civil & Structure"

# RPT-001: 最新 IFC docx + 辅助计算
rpt001 = CIVIL / "GG31-C-RPT-001_MV Power Station (MVPS) Foundation Report_Rev"
RESTORE_LIST.extend([
    (rpt001 / "Superceded" / "GG31-C-RPT-001_MV Power Station (MVPS) Foundation Report_Rev0_IFC.docx",
     rpt001),
    (rpt001 / "Superceded" / "Calculation.xlsx",
     rpt001),
])

# RPT-003: 最新 IFC docx + 辅助计算
rpt003 = CIVIL / "GG31-C-RPT-003_HVSB Platform & Footing Report_Rev"
RESTORE_LIST.extend([
    (rpt003 / "Superceded" / "GG31-C-RPT-003_HVSB Foundation Design Report_Rev0_IFC.docx",
     rpt003),
    (rpt003 / "Superceded" / "Calculation.xlsx",
     rpt003),
])

# RPT-004: 最新 IFC docx + 辅助计算
rpt004 = CIVIL / "GG31-C-RPT-004_Stormwater Drainage Report_Rev"
RESTORE_LIST.extend([
    (rpt004 / "Superceded" / "GG31-C-RPT-004_Stormwater Drainage Report_Rev0_IFC.docx",
     rpt004),
    (rpt004 / "Superceded" / "Warnertown stormwater calculation spreadsheet.xlsx",
     rpt004),
])

# RPT-006: 最新 IFC docx
rpt006 = CIVIL / "GG31-C-RPT-006_BESS Foundation Design Report_Rev"
RESTORE_LIST.append(
    (rpt006 / "Superceded" / "GG31-C-RPT-006_BESS Foundation Design Report_Rev0_IFC.docx",
     rpt006),
)

# RPT-007: 最新 IFC docx
rpt007 = CIVIL / "GG31-C-RPT-007_Smart Control Cabinet (SCC) Design Report_Rev"
RESTORE_LIST.append(
    (rpt007 / "Superceded" / "GG31-C-RPT-007_Smart Control Cabinet (SCC) Design Report_Rev0_IFC.docx",
     rpt007),
)

# RPT-010: 最新 IFC docx
rpt010 = CIVIL / "GG31-C-RPT-010_Light & CCTV Pole Foundation Design Report_Rev"
RESTORE_LIST.append(
    (rpt010 / "Superceded" / "GG31-C-RPT-010_Light Pole Report_Rev0_IFC.docx",
     rpt010),
)

# ── Electrical ─────────────────────────────────────────────────────────
ELEC = PROJECT_ROOT / "Electrical"

# E-RPT-001: 最新 docx + 辅助计算文件
erpt001 = ELEC / "GG31-E-RPT-001_Earthing Design Report_Rev"
RESTORE_LIST.extend([
    (erpt001 / "SS" / "GG31-E-RPT-001_Earthing Design Report_Rev.docx",
     erpt001),
    (erpt001 / "SS" / "Decrement factor.xlsx",
     erpt001),
    (erpt001 / "SS" / "Grounding Mesh IEEE 80-2000.xls",
     erpt001),
    (erpt001 / "SS" / "Soil resistivity test.xlsx",
     erpt001),
    (erpt001 / "SS" / "5.png",
     erpt001),
    (erpt001 / "SS" / "Roseworthy Point 4 Results_1_1_1.emf",
     erpt001),
    (erpt001 / "SS" / "App B GG31-C-PLN-001_Civil Site Plan_RevB.pdf",
     erpt001),
])

# E-RPT-002: 最新 docx
erpt002 = ELEC / "GG31-E-RPT-002_Lightning Risk Assessment Report_RevB"
RESTORE_LIST.append(
    (erpt002 / "SS" / "GG31-E-RPT-002_Lightning Risk Assessment Report_RevB.docx",
     erpt002),
)

# E-RPT-003: 最新 docx
erpt003 = ELEC / "GG31-E-RPT-003_DC Cable Calculation Report_RevC"
RESTORE_LIST.append(
    (erpt003 / "SS" / "GG31-E-RPT-003_DC Cable Calculation Report_RevC.docx",
     erpt003),
)


def main():
    print("=" * 70)
    print("  复原被错误移动的文件")
    print("=" * 70)

    valid = []
    missing = []

    for src, dest_parent in RESTORE_LIST:
        dest = dest_parent / src.name
        if src.exists():
            if dest.exists():
                print(f"  [跳过] 目标已存在: {dest.name}")
                print(f"         in {dest_parent}")
            else:
                valid.append((src, dest))
                print(f"  [移动] {src.name}")
                print(f"         从 {src.parent.name}/ -> {dest_parent.name}/")
        else:
            missing.append(src)
            print(f"  [缺失] {src.name}")
            print(f"         {src}")

    print(f"\n  汇总: {len(valid)} 个文件待移动, {len(missing)} 个文件不存在")

    if not valid:
        print("\n  没有需要移动的文件。")
        return

    confirm = input("\n  确认执行? (y/N): ").strip().lower()
    if confirm != 'y':
        print("  已取消。")
        return

    moved = 0
    errors = 0
    for src, dest in valid:
        try:
            shutil.move(str(src), str(dest))
            print(f"  [OK] {src.name}")
            moved += 1
        except Exception as e:
            print(f"  [错误] {src.name}: {e}")
            errors += 1

    print(f"\n  完成: {moved} 个已移动, {errors} 个失败")


if __name__ == '__main__':
    main()

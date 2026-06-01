"""Dry-run test for ApprovedIFCManager — verify scanning, archiving, mapping logic."""
import sys
from pathlib import Path

# Add script dir to path
sys.path.insert(0, str(Path(__file__).parent))

from ifr_automation_v10 import ApprovedIFCManager

# Warnertown project path
PROJECT_PATH = Path(
    r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"
    r"\2.SA\GG-31 Warnertown BESS"
)


def main():
    print(f"Project: {PROJECT_PATH.name}")
    print(f"Exists: {PROJECT_PATH.exists()}\n")

    try:
        mgr = ApprovedIFCManager(PROJECT_PATH, dry_run=True)
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        return

    print(f"Sharepoint IFR: {mgr.sharepoint_ifr}")
    print(f"Native root:    {mgr.native_root}")
    print(f"IFC output:     {mgr.ifc_output}")
    print()

    # Test scan
    print("=" * 60)
    print("STEP 1: Scan approved files")
    print("=" * 60)
    approved = mgr.scan_approved_files()
    for item in approved:
        archive_flag = "[TO ARCHIVE]" if item['needs_archive'] else "[ARCHIVED]"
        print(f"  {archive_flag} {item['doc_id']} | {item['category']} | {item['filename']}")
    print(f"\nTotal: {len(approved)} approved files")
    print(f"  To archive: {sum(1 for a in approved if a['needs_archive'])}")
    print(f"  Already archived: {sum(1 for a in approved if not a['needs_archive'])}")

    # Test native folder mapping
    print("\n" + "=" * 60)
    print("STEP 2: Map to Native folders")
    print("=" * 60)
    unique_doc_ids = sorted({a['doc_id'] for a in approved})
    mapped = []
    not_found = []

    for doc_id in unique_doc_ids:
        folder = mgr._find_native_folder(doc_id)
        if not folder:
            not_found.append(doc_id)
            print(f"  {doc_id}: NOT FOUND (report/schedule?)")
            continue

        dwg_info = mgr._scan_native_dwgs(folder, doc_id)
        if not dwg_info:
            not_found.append(doc_id)
            print(f"  {doc_id}: folder={folder.name} but NO IFR DWG")
            continue

        mapped.append(dwg_info)
        ifc_rev = dwg_info['existing_ifc_rev']
        ifc_str = f"IFC Rev{ifc_rev}" if ifc_rev is not None else "No IFC"
        print(f"  {doc_id}: {dwg_info['latest_ifr_dwg'].name}")
        print(f"    Folder: {folder.name}")
        print(f"    IFR Rev: {dwg_info['latest_ifr_rev']}")
        print(f"    {ifc_str}")
        print(f"    IFR DWGs: {len(dwg_info['all_ifr_dwgs'])}, "
              f"IFC DWGs: {len(dwg_info['all_ifc_dwgs'])}, "
              f"Other: {len(dwg_info['other_dwgs'])}")

        # Check incremental
        skip = mgr._check_incremental(dwg_info)
        if skip:
            print(f"    INCREMENTAL: would SKIP (source unchanged)")

        # Check cleanup
        keep = {dwg_info['latest_ifr_dwg']}
        for ifc_path, _ in dwg_info.get('all_ifc_dwgs', []):
            keep.add(ifc_path)
        to_clean = []
        for p, _ in dwg_info['all_ifr_dwgs']:
            if p not in keep:
                to_clean.append(p.name)
        for p in dwg_info['other_dwgs']:
            if p not in keep:
                to_clean.append(p.name)
        if to_clean:
            print(f"    CLEANUP: would move {len(to_clean)} DWGs to SS/")
            for name in to_clean:
                print(f"      - {name}")

    print(f"\nMapped: {len(mapped)}, Not found: {len(not_found)}")
    if not_found:
        print(f"  Not found doc-IDs: {', '.join(not_found)}")

    # IFC filename normalization (dry-run)
    print("\n" + "=" * 60)
    print("STEP 2.5: Normalize IFC filenames (DRY-RUN)")
    print("=" * 60)
    renamed = mgr.normalize_ifc_filenames()
    if renamed:
        for old_name, new_name in renamed:
            print(f"  {old_name} -> {new_name}")
    else:
        print("  所有 IFC 文件名已规范")

    # Full batch dry-run
    print("\n" + "=" * 60)
    print("STEP 3: Full batch_convert_approved (DRY-RUN)")
    print("=" * 60)
    result = mgr.batch_convert_approved()
    print(f"\nResult summary:")
    print(f"  Renamed IFC: {len(result.get('renamed_ifc', []))}")
    print(f"  Approved files: {len(result['approved_files'])}")
    print(f"  Archived: {result['archived_count']}")
    print(f"  Native mapped: {len(result['native_mapped'])}")
    print(f"  Native not found: {len(result['native_not_found'])}")
    print(f"  Converted: {len(result['converted'])}")
    print(f"  Skipped: {len(result['skipped'])}")
    print(f"  Errors: {len(result['errors'])}")
    if result['errors']:
        for err in result['errors']:
            print(f"    - {err}")


if __name__ == '__main__':
    main()

"""Dry-run test for AsBuiltManager - verify scanning logic without AutoCAD."""

import sys, os
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, r"D:\1. SOP\SOP_Stage 2 IFR Sync√\V6√")

from pathlib import Path
from ifr_automation_v10 import AsBuiltManager

LMS_PATH = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\LMS"
COLE2_PATH = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\1.NSW\Coleambally #2"


def test_scan(proj_path=LMS_PATH):
    """Test scanning logic: find all IFC sources in Native folders."""
    mgr = AsBuiltManager(proj_path, dry_run=True)
    print(f"Native root: {mgr.native_root}")
    print(f"AB output:   {mgr.ab_output}")
    print()

    # Test existing AB scan
    existing = mgr._scan_existing_ab()
    print(f"Existing AB PDFs: {existing}")
    print()

    # Test full scan
    scan = mgr.scan_native_for_ab()
    print(f"Found {len(scan)} doc-IDs with IFC sources:\n")

    single_count = 0
    multi_count = 0
    for item in scan:
        ifc = item['ifc_source']
        is_multi = ifc['is_multi_page']
        n_pages = len(ifc['dwg_paths'])
        ab_rev = (item['existing_ab_rev'] or 0) + 1
        type_label = f"MULTI ({n_pages} pages)" if is_multi else "SINGLE"

        print(f"  {item['doc_id']:20s} {type_label:20s} IFC Rev{ifc['ifc_rev']}  -> AB Rev{ab_rev}")
        print(f"    Source: {ifc['source_dir'].name}/")
        for dwg in ifc['dwg_paths'][:3]:
            print(f"      {dwg.name}")
        if n_pages > 3:
            print(f"      ... and {n_pages - 3} more")
        print()

        if is_multi:
            multi_count += 1
        else:
            single_count += 1

    print(f"\nSummary: {single_count} single-DWG + {multi_count} multi-page = {len(scan)} total")


def test_dry_run_batch(proj_path=LMS_PATH):
    """Test batch_convert in dry-run mode."""
    mgr = AsBuiltManager(proj_path, dry_run=True)
    results = mgr.batch_convert()

    print(f"\nDry-run results:")
    for r in results:
        status = "OK" if r['success'] else "FAIL"
        rev = r.get('ab_rev', '?')
        pdf = Path(r['pdf_path']).name if r.get('pdf_path') else 'N/A'
        print(f"  [{status}] {r['doc_id']} Rev{rev} -> {pdf}")


if __name__ == '__main__':
    project = sys.argv[1] if len(sys.argv) > 1 else 'lms'
    paths = {'lms': LMS_PATH, 'cole2': COLE2_PATH}
    proj_path = paths.get(project.lower(), project)
    print("=" * 70)
    print(f"AS BUILT Manager - Dry Run Test ({project})")
    print("=" * 70)
    print()
    test_scan(proj_path)
    print()
    print("=" * 70)
    print("Batch Convert (dry-run)")
    print("=" * 70)
    test_dry_run_batch(proj_path)

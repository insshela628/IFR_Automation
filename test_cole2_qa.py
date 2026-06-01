"""QA validation: compare Issue Targeted vs re-exported AS BUILT PDFs for Coleambally2."""
import sys, os
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')

import fitz  # PyMuPDF
from pathlib import Path

AB_ROOT = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"
               r"\1.NSW\Coleambally #2\Design\Engineering\1. Drawings\5. As Built")
ISSUE_DIR = AB_ROOT / "Issue Targeted"
CLIENT_DIR = AB_ROOT / "3. As Built Client"

STAMP_PHRASES = ["AS BUILT", "FOR CONSTRUCTION", "PRINTED IN COLOUR",
                 "PRINTED IN COLOR", "DRAWINGS TO BE", "FOR REVIEW"]


def analyze_pdf(pdf_path):
    """Analyze a single PDF for stamp issues."""
    result = {"path": pdf_path.name, "pages": 0, "stamps": {}, "issues": []}
    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        result["issues"].append(f"Cannot open: {e}")
        return result

    result["pages"] = len(doc)
    if len(doc) > 1:
        last_page = doc[-1]
        if last_page.get_text().strip() == "":
            result["issues"].append(f"Phantom blank page (page {len(doc)})")

    for phrase in STAMP_PHRASES:
        count = 0
        locations = []
        for page_idx, page in enumerate(doc):
            instances = page.search_for(phrase)
            count += len(instances)
            for inst in instances:
                pw, ph = page.rect.width, page.rect.height
                rx = inst.x0 / pw
                ry = inst.y0 / ph
                locations.append((page_idx + 1, f"({rx:.2f},{ry:.2f})",
                                  f"{inst.width:.0f}x{inst.height:.0f}"))
        if count > 0:
            result["stamps"][phrase] = {"count": count, "locations": locations}

    if "FOR CONSTRUCTION" in result["stamps"]:
        result["issues"].append(
            f"FOR CONSTRUCTION leftover ({result['stamps']['FOR CONSTRUCTION']['count']}x)")
    if "FOR REVIEW" in result["stamps"]:
        result["issues"].append(
            f"FOR REVIEW leftover ({result['stamps']['FOR REVIEW']['count']}x)")

    as_built = result["stamps"].get("AS BUILT", {})
    colour = result["stamps"].get("PRINTED IN COLOUR", {}) or result["stamps"].get("PRINTED IN COLOR", {})

    if not as_built:
        result["issues"].append("MISSING: AS BUILT stamp not found")
    elif as_built["count"] > 1:
        result["issues"].append(f"DUPLICATE: AS BUILT stamp x{as_built['count']}")

    if colour and colour["count"] > 1:
        result["issues"].append(f"DUPLICATE: COLOUR stamp x{colour['count']}")

    # Check stamp size consistency (AS BUILT vs COLOUR box width)
    if as_built and colour:
        ab_locs = as_built["locations"]
        cl_locs = colour["locations"]
        if ab_locs and cl_locs:
            ab_w = float(ab_locs[0][2].split("x")[0])
            cl_w = float(cl_locs[0][2].split("x")[0])
            if ab_w > 0 and cl_w > 0:
                ratio = ab_w / cl_w
                if ratio < 0.7 or ratio > 1.3:
                    result["issues"].append(
                        f"SIZE MISMATCH: AS BUILT width={ab_w:.0f} vs COLOUR width={cl_w:.0f} "
                        f"(ratio={ratio:.2f}, expected ~1.0)")

    doc.close()
    return result


def main():
    print("=" * 80)
    print("Coleambally2 AS BUILT QA — Issue Targeted vs Re-exported")
    print("=" * 80)

    # Collect Issue Targeted PDFs
    issue_pdfs = sorted(ISSUE_DIR.glob("*.pdf")) if ISSUE_DIR.exists() else []
    client_pdfs = sorted(CLIENT_DIR.glob("*.pdf")) if CLIENT_DIR.exists() else []

    print(f"\nIssue Targeted: {len(issue_pdfs)} PDFs")
    print(f"3. As Built Client: {len(client_pdfs)} PDFs\n")

    # Build doc-ID lookup for client PDFs
    def extract_docid(name):
        import re
        m = re.match(r'(NSW153-[A-Z]-[A-Z]+-\d+)', name)
        return m.group(1) if m else None

    client_by_docid = {}
    for p in client_pdfs:
        did = extract_docid(p.name)
        if did:
            client_by_docid.setdefault(did, []).append(p)

    # Analyze each Issue Targeted PDF and its re-exported counterpart
    print("-" * 80)
    print(f"{'Doc-ID':<22} {'Source':<10} {'Pgs':>3} {'AS BUILT':>8} {'FC':>3} "
          f"{'COLOUR':>6} {'Issues'}")
    print("-" * 80)

    all_issues = []

    for pdf in issue_pdfs:
        did = extract_docid(pdf.name)
        r = analyze_pdf(pdf)
        ab_count = r["stamps"].get("AS BUILT", {}).get("count", 0)
        fc_count = r["stamps"].get("FOR CONSTRUCTION", {}).get("count", 0)
        cl_count = (r["stamps"].get("PRINTED IN COLOUR", {}).get("count", 0) +
                    r["stamps"].get("PRINTED IN COLOR", {}).get("count", 0))
        issues_str = "; ".join(r["issues"]) if r["issues"] else "OK"
        print(f"{did or '???':<22} {'ISSUE':<10} {r['pages']:>3} {ab_count:>8} "
              f"{fc_count:>3} {cl_count:>6} {issues_str}")
        if r["issues"]:
            all_issues.append((did, "ISSUE", r["issues"]))

        # Check re-exported version
        if did and did in client_by_docid:
            for cp in client_by_docid[did]:
                cr = analyze_pdf(cp)
                ab2 = cr["stamps"].get("AS BUILT", {}).get("count", 0)
                fc2 = cr["stamps"].get("FOR CONSTRUCTION", {}).get("count", 0)
                cl2 = (cr["stamps"].get("PRINTED IN COLOUR", {}).get("count", 0) +
                       cr["stamps"].get("PRINTED IN COLOR", {}).get("count", 0))
                ci_str = "; ".join(cr["issues"]) if cr["issues"] else "OK"
                print(f"{'':<22} {'FIXED':<10} {cr['pages']:>3} {ab2:>8} "
                      f"{fc2:>3} {cl2:>6} {ci_str}")
                if cr["issues"]:
                    all_issues.append((did, "FIXED", cr["issues"]))

    # Also check client PDFs that are NOT in Issue Targeted
    print("\n" + "-" * 80)
    print("Other 3. As Built Client PDFs (not in Issue Targeted):")
    print("-" * 80)
    issue_docids = {extract_docid(p.name) for p in issue_pdfs}
    for pdf in client_pdfs:
        did = extract_docid(pdf.name)
        if did in issue_docids:
            continue
        r = analyze_pdf(pdf)
        ab_count = r["stamps"].get("AS BUILT", {}).get("count", 0)
        fc_count = r["stamps"].get("FOR CONSTRUCTION", {}).get("count", 0)
        cl_count = (r["stamps"].get("PRINTED IN COLOUR", {}).get("count", 0) +
                    r["stamps"].get("PRINTED IN COLOR", {}).get("count", 0))
        issues_str = "; ".join(r["issues"]) if r["issues"] else "OK"
        print(f"{did or '???':<22} {'CLIENT':<10} {r['pages']:>3} {ab_count:>8} "
              f"{fc_count:>3} {cl_count:>6} {issues_str}")
        if r["issues"]:
            all_issues.append((did, "CLIENT", r["issues"]))

    # Summary
    print("\n" + "=" * 80)
    print("SUMMARY")
    print("=" * 80)
    if all_issues:
        print(f"\n{len(all_issues)} doc-IDs with remaining issues:\n")
        for did, src, issues in all_issues:
            for iss in issues:
                print(f"  [{src}] {did}: {iss}")
    else:
        print("\nAll PDFs passed QA checks.")


if __name__ == "__main__":
    main()

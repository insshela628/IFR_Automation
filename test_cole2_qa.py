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

# Stamp box detection via get_drawings() (gold standard: BLD-003, SLD-001)
_RECT_ZONE_X  = 0.60   # stamp rects start at x > 60% of page
_RECT_ZONE_Y  = 0.60   # stamp rects start at y > 60% of page
_RECT_MIN_W   = 0.08   # min width = 8% of page width (~190pt on A1)
_RECT_MIN_H   = 0.02   # min height = 2% of page height (~34pt on A1)
_ALIGN_TOL    = 20     # left/right edge alignment tolerance in PDF points


def _stamp_rects(page):
    """Return stamp-box rects in the stamp zone (filled OR thin-stroked)."""
    pw, ph = page.rect.width, page.rect.height
    out = []
    for p in page.get_drawings():
        r = p["rect"]
        if (r.x0 > pw * _RECT_ZONE_X and r.y0 > ph * _RECT_ZONE_Y
                and r.width  > pw * _RECT_MIN_W
                and r.height > ph * _RECT_MIN_H
                and r.width  < pw * 0.30
                and r.height < ph * 0.10):
            out.append(r)
    return out


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

    # Position check: AS BUILT stamp must be in bottom-right quadrant (x>50%, y>50%)
    if as_built:
        for page_idx, page in enumerate(doc):
            pw, ph = page.rect.width, page.rect.height
            words = page.get_text("words")
            ab_words = [w for w in words if "BUILT" in w[4].upper()]
            for w in ab_words:
                cx = (w[0] + w[2]) / 2
                cy = (w[1] + w[3]) / 2
                if cx < pw * 0.5 or cy < ph * 0.5:
                    result["issues"].append(
                        f"POSITION: Page {page_idx+1} AS BUILT not in bottom-right "
                        f"(x={cx/pw:.0%}, y={cy/ph:.0%})")

    # RECT alignment check: COLOUR and AS BUILT boxes must share same left edge (±20pt)
    # Uses get_drawings() to find actual stamp rectangle borders (not text bounding boxes).
    for page_idx, page in enumerate(doc):
        rects = _stamp_rects(page)
        if len(rects) >= 2:
            rects_sorted = sorted(rects, key=lambda r: r.y0)  # top to bottom
            lefts  = [r.x0 for r in rects_sorted]
            rights = [r.x0 + r.width for r in rects_sorted]
            l_spread = max(lefts)  - min(lefts)
            r_spread = max(rights) - min(rights)
            if l_spread > _ALIGN_TOL or r_spread > _ALIGN_TOL:
                result["issues"].append(
                    f"MISALIGN: Page {page_idx+1} stamp boxes not aligned "
                    f"(left spread={l_spread:.0f}pt, right spread={r_spread:.0f}pt, "
                    f"tol={_ALIGN_TOL}pt)")
        elif len(rects) == 1:
            # Only one stamp box visible — likely has_colour=True but COLOUR is thin/missing
            pass  # not an error; AS BUILT box exists

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

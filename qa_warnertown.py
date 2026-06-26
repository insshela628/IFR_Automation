"""
Read-only AS BUILT QA for GG-31 Warnertown — PyMuPDF (fitz) only, never drives AutoCAD.

Reuses the proven cole2 QA core (analyze_pdf + _stamp_rects from test_cole2_qa)
and adds per-box FILL COLOUR + line-weight reporting so we can verify the
cross-project stamp standard on Warnertown's own frame:
  - 2 stamp boxes per page (AS BUILT lower + COLOUR upper)
  - both RED (fill ~= (1,0,0)) and same weight  -> gold BLD-003/SLD-001/E-PLN-002
  - left/right edge spread <= 20pt (aligned)
  - exactly 1 AS BUILT + 1 COLOUR text, no FOR CONSTRUCTION / FOR REVIEW leftover
  - AS BUILT in bottom-right quadrant; no phantom blank page

Usage:
    python qa_warnertown.py                 # QA the whole AB output folder
    python qa_warnertown.py PLN-002         # filter by doc-id substring
"""
import sys, re
from pathlib import Path
import fitz

sys.path.insert(0, str(Path(__file__).resolve().parent))  # 仓自指: 自动定位, 迁移无需改
from test_cole2_qa import analyze_pdf, _stamp_rects  # reuse verified core

AB_DIR = Path(
    r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"
    r"\2.SA\GG-31 Warnertown BESS\Design\Engineering"
    r"\1. Drawings\5. As Built\3. As Built Client")


def _box_styles(page):
    """Per stamp-box: (fill_rgb, stroke_rgb, width, is_filled). For colour/weight QA."""
    pw, ph = page.rect.width, page.rect.height
    out = []
    for p in page.get_drawings():
        r = p["rect"]
        if (r.x0 > pw * 0.60 and r.y0 > ph * 0.60
                and r.width > pw * 0.08 and r.height > ph * 0.02
                and r.width < pw * 0.30 and r.height < ph * 0.10):
            out.append({
                "y0": round(r.y0, 1),
                "fill": tuple(round(c, 2) for c in p["fill"]) if p.get("fill") else None,
                "stroke": tuple(round(c, 2) for c in p["color"]) if p.get("color") else None,
                "width": round(p.get("width") or 0, 2),
            })
    return sorted(out, key=lambda d: d["y0"])


def _is_red(rgb):
    return rgb is not None and rgb[0] > 0.6 and rgb[1] < 0.4 and rgb[2] < 0.4


def qa_one(pdf):
    r = analyze_pdf(pdf)
    # analyze_pdf sums stamp text across ALL pages, so a legit multi-page PDF
    # (1 stamp per page) trips its "DUPLICATE" flag. Drop those cross-page
    # false positives; we re-check duplicates PER PAGE below (skill: a single
    # source DWG can have many PaperSpace layouts → legitimately multi-page).
    r["issues"] = [i for i in r["issues"] if not i.startswith("DUPLICATE:")]
    doc = fitz.open(str(pdf))
    style_issues = []
    for i, page in enumerate(doc):
        # Per-page duplicate check (the real defect: 2+ stamps on the SAME page)
        ab_on_page = len(page.search_for("AS BUILT"))
        col_on_page = (len(page.search_for("PRINTED IN COLOUR")) +
                       len(page.search_for("PRINTED IN COLOR")))
        if ab_on_page > 1:
            style_issues.append(f"P{i+1}: DUPLICATE AS BUILT x{ab_on_page}")
        if col_on_page > 1:
            style_issues.append(f"P{i+1}: DUPLICATE COLOUR x{col_on_page}")

        boxes = _box_styles(page)
        # The standard is "the two boxes must MATCH EACH OTHER" (same colour +
        # same weight). It is NOT "must be red": a layout plotted with a
        # monochrome CTB renders the WHOLE drawing — stamp included — black/thin,
        # which is a legitimate fallback. So flag INCONSISTENCY between the two
        # boxes, not absence of red. (A surviving old stamp shows up here as
        # one red/thick + one black/thin → mismatch → caught.)
        if len(boxes) >= 2:
            fills = {b["fill"] for b in boxes}
            widths = {b["width"] for b in boxes}
            if len(fills) > 1 or len(widths) > 1:
                style_issues.append(
                    f"P{i+1}: stamp boxes don't match each other "
                    f"(fills={fills}, widths={widths}) → likely old stamp survived "
                    f"(styles={boxes})")
            # If the page IS colour-capable (has red elsewhere) but the stamp is
            # black, the stamp lost its colour — a real defect (not monochrome plot).
            page_has_red = any(
                _is_red(g.get("fill")) or _is_red(g.get("color"))
                for g in page.get_drawings())
            stamp_red = any(_is_red(b["fill"]) or _is_red(b["stroke"]) for b in boxes)
            if page_has_red and not stamp_red:
                style_issues.append(
                    f"P{i+1}: colour drawing but stamp is black (lost colour)")
    doc.close()
    r["issues"].extend(style_issues)
    return r


def main():
    filt = sys.argv[1].lower() if len(sys.argv) > 1 else ""
    pdfs = sorted(AB_DIR.glob("*.pdf")) if AB_DIR.exists() else []
    if filt:
        pdfs = [p for p in pdfs if filt in p.name.lower()]
    print("=" * 84)
    print(f"WARNERTOWN AS BUILT QA — {len(pdfs)} PDF(s) in 3. As Built Client")
    print("=" * 84)
    print(f"{'Doc-ID':<24} {'Pgs':>3} {'AB':>3} {'FC':>3} {'COL':>3}  Result")
    print("-" * 84)
    npass = 0
    fails = []
    for pdf in pdfs:
        m = re.match(r'(GG31-[A-Z]-[A-Z]+-\d+)', pdf.name)
        did = m.group(1) if m else pdf.name[:22]
        r = qa_one(pdf)
        ab = r["stamps"].get("AS BUILT", {}).get("count", 0)
        fc = r["stamps"].get("FOR CONSTRUCTION", {}).get("count", 0)
        col = (r["stamps"].get("PRINTED IN COLOUR", {}).get("count", 0) +
               r["stamps"].get("PRINTED IN COLOR", {}).get("count", 0))
        ok = not r["issues"]
        npass += ok
        print(f"{did:<24} {r['pages']:>3} {ab:>3} {fc:>3} {col:>3}  "
              f"{'PASS' if ok else 'FAIL: ' + '; '.join(r['issues'])}")
        if not ok:
            fails.append((did, r["issues"]))
    print("-" * 84)
    print(f"=== QA DONE: {npass}/{len(pdfs)} PASS, {len(fails)} FAIL ===")
    for did, iss in fails:
        print(f"  FAIL {did}: {'; '.join(iss)}")


if __name__ == "__main__":
    main()

"""
Visual QA script for AS BUILT PDF stamps using PyMuPDF (fitz).
Checks each PDF against reference stamp standards for GG-31 Warnertown.
"""

import fitz  # PyMuPDF
import sys
import os

PDF_DIR = r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)\2.SA\GG-31 Warnertown BESS\Design\Engineering\1. Drawings\5. As Built\3. As Built Client"

FILES = [
    ("GG31-C-PLN-001 REV 2 Civil Site Layout Plan_AS BUILT.pdf",         2),
    ("GG31-C-PLN-002 REV 1 Site Coordinates design Civil Scope_AS BUILT.pdf", 3),
    ("GG31-C-PLN-005 REV 1 Trench Alignment Layout Plan_AS BUILT.pdf",   3),
    ("GG31-E-BLD-001 REV 1 Auxiliary Cable Block Diagram_AS BUILT.pdf",  1),
    ("GG31-E-PLN-004 REV 1 HV Cable Route Layout Plan_AS BUILT.pdf",     1),
]

SEPARATOR = "=" * 80

def check_pdf(filepath, expected_pages):
    filename = os.path.basename(filepath)
    issues = []
    info_lines = []

    # --- Open ---
    try:
        doc = fitz.open(filepath)
    except Exception as e:
        return False, [f"CANNOT OPEN: {e}"], []

    actual_pages = doc.page_count

    # --- Page count ---
    if actual_pages != expected_pages:
        issues.append(f"WRONG PAGE COUNT: expected {expected_pages}, got {actual_pages}")
    else:
        info_lines.append(f"Page count: {actual_pages} (correct)")

    for page_num in range(actual_pages):
        page = doc[page_num]
        pw = page.rect.width
        ph = page.rect.height
        label = f"Page {page_num + 1}"

        # --- Extract all text blocks with bounding boxes ---
        blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]

        # Collect spans with text
        all_spans = []
        for block in blocks:
            if block.get("type") != 0:  # type 0 = text
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span.get("text", "").strip()
                    if text:
                        all_spans.append({
                            "text": text,
                            "bbox": span["bbox"],   # (x0, y0, x1, y1) — PDF coords, y increases downward in fitz
                            "font": span.get("font", ""),
                            "size": span.get("size", 0),
                        })

        # --- Search for key strings (case-insensitive) ---
        def find_spans(keyword):
            kw = keyword.upper()
            return [s for s in all_spans if kw in s["text"].upper()]

        as_built_spans  = find_spans("AS BUILT") + [s for s in all_spans if "AS-BUILT" in s["text"].upper()]
        # Deduplicate (a span could match both "AS BUILT" and "AS-BUILT" variants — unlikely but safe)
        seen_ids = set()
        unique_as_built = []
        for s in as_built_spans:
            sid = id(s)
            if sid not in seen_ids:
                seen_ids.add(sid)
                unique_as_built.append(s)
        as_built_spans = unique_as_built

        for_construction_spans = find_spans("FOR CONSTRUCTION")
        colour_spans            = find_spans("DRAWINGS TO BE PRINTED IN COLOUR")

        # --- AS BUILT presence ---
        if not as_built_spans:
            issues.append(f"{label}: MISSING AS BUILT stamp — no 'AS BUILT' text found")
        else:
            # --- Duplicate check ---
            if len(as_built_spans) > 1:
                issues.append(
                    f"{label}: DUPLICATE AS BUILT stamps — found {len(as_built_spans)} instances: "
                    + "; ".join(f"'{s['text']}' @ bbox{s['bbox']}" for s in as_built_spans)
                )
            else:
                info_lines.append(f"{label}: AS BUILT stamp found (1 instance) — OK")

            # --- Position check: bottom-right quadrant ---
            for s in as_built_spans:
                x0, y0, x1, y1 = s["bbox"]
                # In fitz/PyMuPDF, y=0 is TOP of page, y increases downward.
                # Bottom-right means x > pw*0.5 AND y > ph*0.5
                cx = (x0 + x1) / 2
                cy = (y0 + y1) / 2
                in_right  = cx > pw * 0.5
                in_bottom = cy > ph * 0.5
                pos_str = f"centre=({cx:.1f},{cy:.1f}), page=({pw:.0f}x{ph:.0f})"
                if not in_right or not in_bottom:
                    issues.append(
                        f"{label}: AS BUILT stamp NOT in bottom-right quadrant — "
                        f"{pos_str}  [in_right={in_right}, in_bottom={in_bottom}]"
                    )
                else:
                    info_lines.append(f"{label}: AS BUILT position OK — {pos_str}")

                # --- Font info (informational) ---
                info_lines.append(f"{label}: AS BUILT font='{s['font']}', size={s['size']:.1f}")

        # --- FOR CONSTRUCTION leftover ---
        if for_construction_spans:
            for s in for_construction_spans:
                issues.append(
                    f"{label}: FOR CONSTRUCTION text STILL PRESENT — "
                    f"'{s['text']}' @ bbox{s['bbox']} (font={s['font']})"
                )
        else:
            info_lines.append(f"{label}: FOR CONSTRUCTION — NOT present (correct)")

        # --- COLOUR box (optional, just note) ---
        if colour_spans:
            info_lines.append(f"{label}: 'DRAWINGS TO BE PRINTED IN COLOUR' found (optional, expected)")
        else:
            info_lines.append(f"{label}: 'DRAWINGS TO BE PRINTED IN COLOUR' NOT found (may be overlapped — OK)")

    doc.close()
    passed = len(issues) == 0
    return passed, issues, info_lines


def main():
    print(SEPARATOR)
    print("AS BUILT STAMP QA REPORT")
    print(f"PDF directory: {PDF_DIR}")
    print(SEPARATOR)

    overall_pass = True

    for filename, expected_pages in FILES:
        filepath = os.path.join(PDF_DIR, filename)
        print(f"\nFILE: {filename}")
        print(f"  Expected pages: {expected_pages}")

        if not os.path.exists(filepath):
            print(f"  [ERROR] File not found: {filepath}")
            overall_pass = False
            continue

        passed, issues, info_lines = check_pdf(filepath, expected_pages)

        for line in info_lines:
            print(f"  [INFO] {line}")

        if issues:
            overall_pass = False
            print(f"  [RESULT] *** FAIL ***")
            for issue in issues:
                print(f"    [ISSUE] {issue}")
        else:
            print(f"  [RESULT] PASS")

    print()
    print(SEPARATOR)
    if overall_pass:
        print("OVERALL: ALL FILES PASS")
    else:
        print("OVERALL: ONE OR MORE FILES FAILED — see issues above")
    print(SEPARATOR)


if __name__ == "__main__":
    main()

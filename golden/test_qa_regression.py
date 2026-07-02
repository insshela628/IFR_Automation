# -*- coding: utf-8 -*-
"""QA golden-master regression gate.

Runs the REAL `_qa_validate_ab_pdf` over the SAME live PDFs and asserts the
per-PDF warnings + verdict are byte-identical to the frozen golden snapshot
(`qa_verdicts.json`). OFFLINE, read-only — never opens AutoCAD, never mutates.

WHY this lives OUTSIDE the QA module: the gate must treat QA as a black box.
If it lived inside QA, one edit could change both the behaviour and the check
that should catch it. Keeping it separate makes it a true independent net for
①(structured fault objects)/②/③ — those must NOT shift any verdict.

Pass  -> exit 0, "GATE PASS".
Fail  -> exit 1, prints each drifted PDF (added/removed warnings, verdict flip,
         and whether the PDF itself changed size = input drift vs logic drift).

Run:  python golden/test_qa_regression.py
Re-baseline (only after deliberately accepting a new golden): gen_qa_golden.py
"""
import sys, json
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from gen_qa_golden import collect, GOLDEN  # same source → same fidelity


def main():
    if not GOLDEN.exists():
        print("NO GOLDEN baseline — run gen_qa_golden.py first.")
        return 2
    golden = json.loads(GOLDEN.read_text(encoding="utf-8"))
    current = collect()

    drift = []
    # PDFs missing / newly appeared
    for name in sorted(set(golden) - set(current)):
        drift.append((name, "MISSING (in golden, not on disk)", None))
    for name in sorted(set(current) - set(golden)):
        drift.append((name, "NEW (on disk, not in golden)", None))

    for name in sorted(set(golden) & set(current)):
        g, c = golden[name], current[name]
        gw, cw = set(g["warnings"]), set(c["warnings"])
        if g["verdict"] != c["verdict"] or gw != cw:
            note = []
            if g["verdict"] != c["verdict"]:
                note.append(f"verdict {g['verdict']} -> {c['verdict']}")
            for w in sorted(cw - gw):
                note.append(f"+ {w}")
            for w in sorted(gw - cw):
                note.append(f"- {w}")
            input_changed = g.get("size") != c.get("size")
            tag = "INPUT CHANGED (PDF reconverted)" if input_changed \
                else "LOGIC DRIFT (same PDF, different QA result)"
            drift.append((name, tag, note))

    if not drift:
        n = len(current)
        print(f"GATE PASS — {n} PDFs, all verdicts + warnings match golden.")
        return 0

    print(f"GATE FAIL — {len(drift)} PDF(s) drifted from golden:\n")
    for name, tag, note in drift:
        print(f"  [{tag}] {name}")
        for line in (note or []):
            print(f"       {line}")
    print("\nIf this drift is INTENDED, re-baseline with gen_qa_golden.py.")
    return 1


if __name__ == "__main__":
    sys.exit(main())

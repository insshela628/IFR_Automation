# Coleambally #2 — Punchlist Action Log (Issue Tracker → issue IDs **C2-031 … C2-046**, plus C2-065)

Source: `5. As Built\Coleambally_2_Issue_Tracker (16th July).xlsx` (sheet "Issue Tracker", data rows 3+).
Compiled: 2026-07-21. Support-doc root: `...\Handover punchlist - July 2026\Cole2 Punchlist_Support documents\`.
Note: this .md lives in the Obsidian vault (`L03-pipeline\ifr-automation\project-notes\Coleambally2\`), NOT in Dropbox.

## ⚠ 2026-07-21 scope correction
- **Range is C2-031 … C2-046** (tracker rows 33–48), plus C2-065. Earlier draft mis-scoped to C2-029…C2-044 (rows 31–46) — off by two at BOTH ends.
- Removed erroneous empty folders C2-029 / C2-030 (rows 31–32 = "Appendix T Spare Parts" commercial notes, no drawing).
- Added the two missed drawing rows:
  - **C2-045** = `NSW153-GAD-002 Array & DC Combiner…` (no C-fill).
  - **C2-046** = `NSW153-GAD-003 REV 1 HVSB Earth Connection…` (yellow) — this is the detail dwg that E-BLD-002 note 4 cross-references.
- All yellow rows in range carry Column L "engineering team to update, do not need [to touch titleblock]" → engineering updates native content; **titleblock + Rev bump + export is AI's job** (uniform, per user).
- C2-065 evidence corrected by user to `ENG3-Coleambally signed - R1.pdf` (signed doc), replacing my internal procurement xlsx.
- Note-defect status: designers' `Rev1_AB` native updates already CLEARED the flagged SHALL notes on E-BLD-002 & E-PLN-002 (verified read-only). To be re-confirmed visually at render/QA.

## Legend of the tracker highlighting (what the colours mean)
- **Yellow cell (Column C, drawing name)** = drawing is in scope; standard treatment = AS BUILT titleblock refresh + PDF export. Content edit deferred to engineering team (per each row's Column L).
- **Red cell (Column L, internal comment)** = Shela has given a *specific* content instruction in Column L to be executed now.
- Column L note "do not need to edit the titleblock" is an instruction to the **engineer only**; per Shela, AI updates the titleblock uniformly (revision governance is being centralised, not left to the drafter).

## Key finding
Every AS BUILT drawing is **already at REV 1 / AS BUILT with a correct two-box stamp**. The EE comments are **content/narrative** defects, not stamp defects. Three defect classes dominate:
1. **Future/imperative tense surviving into AS BUILT** ("SHALL", "TO BE CONFIRMED ON SITE", "TO RETICULATE") — an AS BUILT states installed fact.
2. **Temporary construction elements still shown** (laydown, site offices, temp buildings).
3. **Stale base map** (not the latest as-built site layout).

## Per-row disposition

| Issue | Drawing | Colour | Column L instruction | Disposition | Owner |
|---|---|---|---|---|---|
| C2-029 | Spare Parts (doc) | — | "Yes." | N/A — commercial note, no drawing | — |
| C2-030 | Spare Parts (doc) | — | "Free issued equipment." | N/A — commercial note, no drawing | — |
| C2-031 | Easement (doc) | Yellow | collect doc | **DONE** — executed deed + agreement Rev2 + title search copied to `C2-031\` | Filing |
| C2-032 | C-PLN-001 Civil Site Layout | Yellow | notes "confirmed on site" → reword | TB refresh + export; content reword | Engineer + AI-TB |
| C2-033 | C-PLN-001 | — | Water tank / APZ — "Adam to check with Michael" | Design check, not a text edit | Engineer |
| C2-034 | C-PLN-001 | — | "Delete site laydown buildings" | Geometry delete | Engineer |
| C2-035 | C-PLN-002 Civil Scope | Yellow | notes "confirmed on site" → reword | TB refresh + export; content reword | Engineer + AI-TB |
| C2-036 | C-PLN-002 Equipment Scope | Yellow | notes "confirmed on site" → reword | TB refresh + export; content reword | Engineer + AI-TB |
| **C2-037** | **C-PLN-003 Fence & Gate** | **Red** | "update the corresponding note — delete it or reword to AS BUILT stage" | **Needs note ID** — fence notes are fabrication specs; no explicit "confirmed on site" note on the sheet. Ambiguous which note. | Engineer/Shela to point |
| C2-038 | C-PLN-004 Ext Road/Track/Hardstand | Yellow | "shows temporary construction buildings" | Geometry delete + TB refresh | Engineer + AI-TB |
| **C2-039** | **C-PLN-005 Trench Alignment** | **Red** | "update the base map per the latest site layout base map" | **Base-map replacement** = major geometry redraw, not a text edit | Engineer |
| C2-040 | C-PLN-008 HVSB Foundation | Yellow | "Notes sheet 1 are missing" | Add missing notes sheet | Engineer + AI-TB |
| **C2-041** | **E-BLD-002 Earthing Block Diagram** | **Red** | "Update narration re notes 2 & 3 (SHALL not presenting AS BUILT stage)" | **AI-DOABLE text edit** (specs below) | AI |
| C2-042 | E-PLN-001 DC Cable Route | Yellow | "still contain temporary buildings" | Geometry delete + TB refresh | Engineer + AI-TB |
| **C2-043** | **E-PLN-002 Earth Cable Route** | **Red** | "Update narration re notes 4 & 5 (SHALL not AS BUILT; not sure if 4 ok)" | **AI-DOABLE text edit** (specs below) | AI |
| C2-044 | E-PLN-008 CCTV Layout | Yellow | notes "confirmed on site" → reword | TB refresh + export; content reword | Engineer + AI-TB |
| C2-065 | HV switchgear datasheets | — | "Existed. Shela to check." | **DONE** — `NSW153-E-DTS-008 …RevB.xlsx` copied to `C2-065\` | Filing |

## Exact AI note-edit specifications (the two clean, AI-doable red rows)

### C2-041 — E-BLD-002 Earthing Block Diagram (source: `1. Native\…E-BLD-002…\Rev.1 - AB\…Rev1_AB.dwg`)
- Note 2: `FINAL RETICULATION OF EARTHING CABLES SHALL BE COORDINATED ON SITE.`
  → `FINAL RETICULATION OF EARTHING CABLES WAS COORDINATED ON SITE.`
- Note 3: `ALL EARTH CABLES SHALL HAVE NON-CORROSIVE LABELS TIED ON BOTH ENDS OF CABLE.`
  → `ALL EARTH CABLES HAVE NON-CORROSIVE LABELS TIED ON BOTH ENDS OF CABLE.`

### C2-043 — E-PLN-002 Earth Cable Route Layout Plan (source: `…E-PLN-002…\Rev.1 - AB\…Rev1_AB.dwg`)
- Note 4: `EARTH CABLE TO ARRAY PIERS TO RETICULATE BUNDLED WITH DC CABLES TO ARRAY.`
  → `EARTH CABLE TO ARRAY PIERS RETICULATED, BUNDLED WITH DC CABLES TO ARRAY.`  *(addresses "not sure if 4 is ok": "TO RETICULATE" is future intent → past fact)*
- Note 5: `FINAL RETICULATION OF EARTHING CABLES SHALL BE COORDINATED ON SITE.`
  → `FINAL RETICULATION OF EARTHING CABLES WAS COORDINATED ON SITE.`

## ⚠ 2026-07-21 — BLOCKER: two conflicting AS BUILT masters (edits PAUSED pending drafter)

Read-only COM/PDF forensics before editing surfaced a version conflict that stops the note edits above. **No DWG was modified.**

**The conflict (per doc-folder `Rev.1 - AB\`):**
| Drawing | Issued to client (06-09, complete) | Newer un-issued "master" (07-20) |
|---|---|---|
| E-BLD-002 | `…RevE_AB.dwg` — **5 notes** (incl. #1 AS/NZS 3000/5033 compliance, #4 GAD-004/003 detail refs, #2&3 the SHALL notes EE flagged) | `…Rev1_AB.dwg` — **only 1 note** (old note 5 "see E-PLN-002"). Notes 1–4 **deleted**, not reworded. |
| E-PLN-002 | `…RevD_AB.dwg` — complete | `…Rev1_AB.dwg` — **will not open** via COM (`<unknown>.Open`; likely broken / missing XREF / needs RECOVER). |

**Corroborating facts:**
- Client-facing PDFs (`5. As Built\3. As Built Client\`) are dated **2026-06-09**, plotted from `RevE_AB`/`RevD_AB`. The 07-20 `Rev1_AB` files were **never issued**.
- The EE punchlist (16 July) notes-2&3 wording matches the **06-09 `RevE_AB`** exactly (5-note block), NOT `Rev1_AB`.
- AutoCAD was found with `E-BLD-002_RevE_AB.dwg` **open + unsaved** (locked ~17h overnight, idle). Its unsaved state is NOT note edits (notes still read "SHALL") — incidental dirty flag. Left untouched.

**Assessment:** the 07-20 `Rev1_AB` batch looks like a defective/abandoned iteration (one drawing gutted its notes incl. the standards-compliance note; the other won't open). The authoritative, complete content is the issued **06-09 `RevE_AB`/`RevD_AB`**, which needs exactly the note-2&3 reword the punchlist requested.

**DECISION (2026-07-21):** User paused to confirm with the drafter whether the `Rev1_AB` note-stripping was deliberate (e.g. earthing notes consolidated onto E-PLN-002) before any edit. **Question for the drafter:** Is `Rev1_AB` (07-20) intentional, or should it be superseded and Rev 2 built from the complete 06-09 `RevE_AB`/`RevD_AB`? Also: why won't `E-PLN-002_Rev1_AB.dwg` open?

> ⚠ The edit-source paths in the two specs above (`…Rev1_AB.dwg`) are **superseded by this finding** — do NOT edit `Rev1_AB` until the drafter question is resolved. If the answer is "supersede Rev1_AB", the correct base is `RevE_AB` / `RevD_AB`.

## Pending engineering actions (geometry — cannot be automated as text)
- C2-034 / C2-038 / C2-042: remove temporary construction buildings / laydown from base map.
- C2-039 (C-PLN-005): replace base map with the latest as-built site layout.
- C2-037 (C-PLN-003): identify the specific note to reword/delete (none obviously matches on the current sheet).
- C2-033: water-tank / APZ encroachment design check.
- C2-040 (C-PLN-008): add the missing Notes sheet 1.

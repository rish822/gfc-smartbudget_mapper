# Smart Budget ↔ GFC Mapper

This codebase auto-enriches GFC (Good For Construction) line items with data from a Smart Budget (BOQ), and flags every item's PR-readiness per the Smart Budgeter spec (GF5.1).

---

## What it does

Takes two inputs per project:
1. **GFC file** — design team's item-level spreadsheet (wildly varied structure across sheets and projects)
2. **Smart Budget / BOQ** — approved budget with quantities, rates, and makes

Runs a 7-layer pipeline and outputs an enriched workbook where every GFC line item is linked to its BOQ counterpart, normalized to the Category Master taxonomy, and flagged as 🟢 Auto-Matched / 🟡 Suggested / 🔴 Not in BOQ.

---

## Quick start

```bash
pip install -r requirements.txt

# Copy input files into data/
cp /path/to/GFC.xlsx data/
cp /path/to/BOQ.xlsx data/

# Run
python run.py --project hosteller
python run.py --project quantiphi
python run.py --list               # see all configured projects
python run.py --project hosteller --sheets-only  # fast L1 audit only
```

Output lands in `output/<ProjectName>_Enriched.xlsx`.

---

## Adding a new project

**Step 1 — Drop files in `data/`**
```
data/MyClient_GFC.xlsx
data/MyClient_BOQ.xlsx
data/Category_master_for_GFC.xlsx   ← shared, already there
```

**Step 2 — Add a config block in `config/projects.yaml`**
```yaml
myclient:
  name: "My Client"
  gfc_path: "data/MyClient_GFC.xlsx"
  boq_path: "data/MyClient_BOQ.xlsx"
  boq_loader: "hosteller"          # or "quantiphi" — see below
  output_path: "output/MyClient_Enriched.xlsx"
  notes: "Brief description"
```

**Step 3 — Choose or create a BOQ loader**

| BOQ format | Loader | Description |
|---|---|---|
| Multi-sheet (one sheet per category) | `"hosteller"` | Hosteller, most project-specific BOQs |
| Single sheet, section codes (A/B/C) | `"quantiphi"` | Quantiphi, typical fit-out summary BOQ |
| Need new format | Create `engine/boq_loaders/<name>.py` | See below |

**If neither loader fits**, copy `engine/boq_loaders/hosteller.py` as a starting point and implement:
```python
def load_<name>_boq(path: str) -> list[dict]:
    """Returns list of dicts, each with keys:
       BOQ No, Section, Category, Sub-category, Item, Description,
       Specification, Make/Brand, Config (size/spec), UoM,
       BCS Qty, BCS Rate, BCS Amount, Space/Area
    """
```
Then set `boq_loader: "<name>"` in the config and register the loader name in `run.py`'s `load_boq()` function.

**Step 4**
```bash
python run.py --project myclient
```

---

## Project structure

```
smart_budget_mapper/
├── CLAUDE.md                     ← you are here
├── run.py                        ← unified CLI entry point
├── requirements.txt
├── config/
│   └── projects.yaml             ← all project configs, paths, thresholds
├── engine/
│   ├── gfc_mapping_engine.py     ← the 7-layer pipeline (core)
│   └── boq_loaders/
│       ├── quantiphi.py          ← single-sheet BOQ loader
│       ├── hosteller.py          ← multi-sheet BOQ loader
│       └── synthetic.py         ← synthetic budget (AVEVA POC)
├── runners/
│   └── output_builder.py         ← shared workbook builder (all 8 output sheets)
├── data/                         ← input files (not committed)
└── output/                       ← generated enriched workbooks
```

---

## The 7-layer pipeline (`engine/gfc_mapping_engine.py`)

Each GFC row passes through all 7 layers in order. Layers 1–4 are extraction/normalization; Layers 5–7 are matching and enrichment.

| Layer | Function | What it does |
|---|---|---|
| **L1** | `classify_sheet()` | Sheet name → Category + sub-hint. Driven by `SHEET_RULES` list — token-based, not exact match |
| **L2** | `map_columns()` / `detect_header_row()` | Finds the header row (scans first 12 rows) and maps column headers to canonical field names via `_NORMALIZED_ALIASES` |
| **L3** | `classify_row()` | Determines whether a row is an `item`, `subrow`, `section`, `total`, `note`, or `empty`. Tries description → product_name → areas in order for the primary description cell |
| **L4** | `normalize_status()` / `normalize_uom()` / `parse_bcs_rate()` | Status: 50+ raw strings → Pending / Client Approved / Excluded. UoM: SQFT/NOS/RFT/etc. Rate: parses "85rs/sqft", "2.5L", "3,75,000" |
| **L5** | `score_match()` + `best_match()` | 9-signal scoring (0–100): Category (hard filter +30), Description similarity (Jaccard+seq max, +40), Sub-cat alignment (±10), Brand (+8), UoM (+5), Item-keyword (+7), Pax/Side/Seater disambiguators (±3–6), Word-overlap hard floor |
| **L6** | `enrich_from_master()` | Finds the best Category Master triple (Category / Sub-category / Item) using stemmed item-keyword match. Result fed back into L5 as sub-cat signal |
| **L7** | `explode_multi_floor()` | Detects "1st Floor / 2nd Floor" column pattern and creates one row per floor |

---

## Key constants to tune

All in `engine/gfc_mapping_engine.py` — edit here first when something doesn't match.

| Constant | Purpose | Tune when… |
|---|---|---|
| `SHEET_RULES` | Token → Category mapping (top of file) | A new GFC sheet type isn't being classified |
| `SKIP_TOKENS` | Sheet names to skip entirely | A junk sheet is generating spurious rows |
| `_NORMALIZED_ALIASES` | Column name → canonical field | A new GFC column header isn't being recognized |
| `MATCH_THRESHOLDS` | `{"auto": 75, "suggested": 55}` | Auto-match rate is too low/high |
| `STOPWORDS` | Words ignored in token overlap | A common non-semantic word is giving false matches |

Override thresholds at runtime without editing code:
```bash
python run.py --project hosteller --auto-threshold 70 --suggest-threshold 50
```

---

## Scoring algorithm (L5)

```
score = 0
+ 30  if category matches (HARD FILTER — returns 0 immediately if mismatch)
       HARD FLOOR: if no meaningful token overlap → return 0
+ 0-40 description similarity = max(SequenceMatcher ratio, token Jaccard) × 40
        uses composite text: description + sub_hint + areas + product_name (all stemmed)
±10   sub-category alignment: +10 if gfc_master_subcat matches budget sub-cat, -10 if mismatch
+ 0-8  brand: +8 if brand/make overlap (substring either direction)
+ 0-5  UoM: +5 if exact match
+ 0-7  item-keyword: +7 if budget Item text appears in composite description
+ 0-3  Pax disambiguator: +3 if "N Pax" matches, -5 if different
+ 0-4  Side disambiguator: +4 if single/double matches, -6 if different
+ 0-3  Seater disambiguator: +3 if "N seater" matches, -5 if different
```

Thresholds: ≥75 → 🟢 Auto-Matched | 55–74 → 🟡 Suggested | <55 → 🔴 Not in BOQ

---

## PR readiness logic (GF5.1)

A GFC row is `pr_ready=True` when ALL of:
1. `match_status == "🟢 Matched"` (auto only, not suggested)
2. `norm_status == "Client Approved"`
3. `raw_qty_design or raw_qty_ops` is not empty
4. `norm_uom` is not None

PR blockers are written to the `PR Blocker` column in the output workbook.

---

## Output workbook (8 sheets)

| Sheet | Purpose |
|---|---|
| 1. Run Summary | Headline metrics + legend |
| 2. Enriched GFC Items | Main deliverable — every GFC row enriched and color-coded |
| 3. Score Breakdown | Forensic — full scoring breakdown per row (for debugging matches) |
| 4. L1 Sheet Audit | Which sheets were classified, skipped, or unclassified |
| 5. BOQ Coverage Gaps | BOQ items with no GFC link (Civil/MEP expected; Furniture here = scope drop) |
| 6. New Scope (GF6.5) | GFC categories absent from BOQ — flag to PM for Variation pricing |
| 7. BOQ Parsed | The BOQ as loaded — for reference |
| 8. Category Master | The canonical taxonomy — Category / Sub-category / Item |

---

## Known limitations and common fixes

**"0 items extracted from a sheet"**
→ Check Sheet 4 (L1 Sheet Audit). If UNCLASSIFIED, add a rule to `SHEET_RULES`.
→ If MAPPED but 0 items: run with `--sheets-only` first, then inspect that sheet's header manually. The `detect_header_row` scans 12 rows — if the header is deeper, increase `max_scan`.

**"Wrong match — GFC item pointing to unrelated BOQ item"**
→ Open Sheet 3 (Score Breakdown) for that row. The `subcat:` signal usually explains it.
→ If L6 master enrichment is picking the wrong sub-cat: check the raw `description` and `sub_hint` — usually a sheet-level sub_hint is contaminating rows that belong to a different sub-cat.

**"PR-ready count is 0"**
→ Most GFC sheets in Indian fit-out projects don't have a Status column, so `norm_status` defaults to Pending. This is a data quality issue, not a code issue. The BOQ's `On site Qty > 0` could serve as a proxy — consider adding a status inference step to the BOQ loader.

**"BOQ rates look wrong (showing AMOUNT not UNIT RATE)"**
→ The BOQ loader is picking up the wrong column. BOQs often have multiple "rate"-like columns (BCS, P/O Rate, CSC Rate). The loader uses `first-match-wins` on column names — `unit rate` must be detected before lookalike columns. Check `load_<name>_boq.py` column detection logic.

**"New GFC has a sheet structure we've never seen"**
→ Run `--sheets-only` first to see L1 results, then inspect the raw sheet. Add aliases to `_NORMALIZED_ALIASES` for any new column names, and a rule to `SHEET_RULES` for the sheet name. Most new GFCs are handled by these two additions alone.

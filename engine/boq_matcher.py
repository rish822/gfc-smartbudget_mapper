"""
Engine 2 — Classified GFC Row → BOQ / Smart Budget Matching

Receives a GFC row that has already been classified by Engine 1
(category + sub-category confirmed) and finds the best matching
BOQ line item using an 11-signal scoring algorithm.

Match status thresholds (tunable via MATCH_THRESHOLDS):
  🟢 Auto-matched  score ≥ 75   no human review needed
  🟡 Suggested     score 55–74  confirm/reject in GFC Connect
  🔴 Not in BOQ    score < 55   new scope or unmatched item
"""
from __future__ import annotations

import re
from difflib import SequenceMatcher
from typing import Optional


# ── Tunable thresholds (also exported for GUI / output_builder) ───────────────
MATCH_THRESHOLDS: dict[str, int] = {"auto": 75, "suggested": 55}


# ── Text helpers (self-contained — no imports from gfc_mapping_engine) ────────

def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9 ]", " ", (s or "").lower()).strip()


def _light_stem(tok: str) -> str:
    t = tok.lower()
    for suf in ("ings", "ing", "ies", "ed", "es", "s"):
        if len(t) > len(suf) + 2 and t.endswith(suf):
            return t[:-3] + "y" if suf == "ies" else t[:-len(suf)]
    return t


_NOISE = [
    re.compile(r"\(\s*[\d'\"\.,xX\s\-mmft]+\s*\)"),
    re.compile(r"\d+['\"]?-?\d*['\"]?\s*[xX]\s*\d+['\"]?-?\d*['\"]?"
               r"(\s*[xX]\s*\d+['\"]?-?\d*['\"]?)?"),
    re.compile(r"\d+\s*mm\s*[xX]\s*\d+\s*mm"),
    re.compile(r"note\s*\d+", re.I),
    re.compile(r"\b\d{3,}\b"),
]

_STOPWORDS = {"the", "a", "an", "and", "or", "of", "in", "for", "to",
              "with", "by", "on", "at", "is", "it", "as", "be", "x"}


def _stem_for_match(s: str) -> str:
    if not s:
        return ""
    out = s
    for p in _NOISE:
        out = p.sub(" ", out)
    out = re.sub(r"[()]", " ", out)
    return re.sub(r"\s+", " ", out).strip()


def _tokens(s: str) -> set:
    return {_light_stem(t)
            for t in re.split(r"[\s,/()\-]+", _norm(s))
            if t and len(t) >= 3 and t not in _STOPWORDS}


def _similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, _norm(a or ""), _norm(b or "")).ratio()


def _composite_gfc(row: dict) -> str:
    parts = [
        row.get("description") or "",
        row.get("subcategory") or "",     # Engine 1 result — strong signal
        row.get("sub_hint") or "",
        row.get("areas") or "",
        row.get("product_name") or "",
    ]
    return " ".join(str(p) for p in parts if p)


# ── Scoring ───────────────────────────────────────────────────────────────────

def score_match(gfc_row: dict, budget_item: dict) -> tuple[int, dict]:
    """
    Engine 2 — 11-signal scoring (0–100).

    gfc_row keys expected:
      category, subcategory (from Engine 1), sub_hint, description, areas,
      product_name, brand, uom, size, finish_code, master_item

    Signals:
      1.  Category hard filter          +30 / immediate 0
          Word-overlap hard floor              immediate 0
      2.  Description similarity        +0–40
      3.  Sub-category alignment        +10 / -10
      4.  Brand match                   +8
      5.  UoM match                     +5
      6.  Item-keyword                  +7
      7.  Pax disambiguator             +3 / -5
      8.  Single/Double disambiguator   +4 / -6
      9.  Seater-count disambiguator    +3 / -5
     10.  Size / config dimensions      +0–6 / -3
     11.  Finish code / product code    +10
    """
    breakdown: dict = {}
    score = 0

    # ── 1. Category hard filter ───────────────────────────────
    gfc_cat = (gfc_row.get("category") or "").upper().strip()
    bud_cat = (budget_item.get("Category") or "").upper().strip()
    if not gfc_cat:
        return (0, {"category": "HARD FAIL — GFC row has no category"})
    if not bud_cat:
        return (0, {"category": "HARD FAIL — BOQ item has no category"})
    if gfc_cat != bud_cat:
        return (0, {"category": f"HARD FAIL — {gfc_cat} ≠ {bud_cat}"})
    score += 30
    breakdown["category"] = "+30 (match)"

    # ── 2. Description similarity ─────────────────────────────
    gfc_text = _stem_for_match(_composite_gfc(gfc_row))
    bud_text = _stem_for_match(" ".join(filter(None, [
        budget_item.get("Description") or "",
        budget_item.get("Sub-category") or "",
        budget_item.get("Item") or "",
    ])))
    gfc_tok = _tokens(gfc_text)
    bud_tok = _tokens(bud_text)
    overlap = gfc_tok & bud_tok
    if not overlap:
        return (0, {"overlap": (
            f"HARD FAIL — no shared tokens "
            f"(gfc={sorted(gfc_tok)[:4]} vs bud={sorted(bud_tok)[:4]})"
        )})
    seq  = _similarity(gfc_text, bud_text)
    jacc = len(overlap) / len(gfc_tok | bud_tok)
    dp   = int(max(seq, jacc) * 40)
    score += dp
    breakdown["description"] = (
        f"+{dp} (seq={seq:.2f}/jacc={jacc:.2f}, shared={sorted(overlap)})"
    )

    # ── 3. Sub-category alignment ─────────────────────────────
    # Engine 1 now guarantees the GFC sub-category is valid for the category
    # (or blank), so it is a trustworthy signal.  Exact match is rewarded
    # strongly; an explicit mismatch is penalised; a blank GFC sub-category is
    # neutral (Engine 1 deliberately left it blank — do not punish).
    gfc_sub = (gfc_row.get("subcategory") or "").upper()
    bud_sub = (budget_item.get("Sub-category") or "").upper()
    if gfc_sub and bud_sub:
        if gfc_sub == bud_sub:
            score += 14
            breakdown["subcat"] = f"+14 (exact '{gfc_sub}')"
        elif gfc_sub in bud_sub or bud_sub in gfc_sub:
            score += 10
            breakdown["subcat"] = f"+10 ('{gfc_sub}' ⊂ '{bud_sub}')"
        else:
            score -= 10
            breakdown["subcat"] = f"-10 ('{gfc_sub}' ≠ '{bud_sub}')"
    elif not gfc_sub:
        # Fall back to sheet-level hint only as a soft tie-breaker, no penalty
        hint = (gfc_row.get("sub_hint") or "").upper()
        if hint and bud_sub and (hint == bud_sub or hint in bud_sub or bud_sub in hint):
            score += 4
            breakdown["subcat"] = f"+4 (sheet-hint '{hint}')"
        else:
            breakdown["subcat"] = "0 (GFC sub-cat blank — neutral)"
    else:
        breakdown["subcat"] = "0 (BOQ sub-cat missing)"

    # ── 4. Brand ──────────────────────────────────────────────
    gfc_brand = _norm(gfc_row.get("brand") or "")
    bud_brand = _norm(budget_item.get("Make/Brand") or "")
    if gfc_brand and bud_brand:
        if gfc_brand in bud_brand or bud_brand in gfc_brand:
            score += 8;  breakdown["brand"] = "+8"
        else:
            breakdown["brand"] = "0 (different)"
    else:
        breakdown["brand"] = "0 (missing)"

    # ── 5. UoM ────────────────────────────────────────────────
    if (gfc_row.get("uom") or "") and (budget_item.get("UoM") or ""):
        if gfc_row["uom"].upper() == budget_item["UoM"].upper():
            score += 5;  breakdown["uom"] = "+5"
        else:
            breakdown["uom"] = "0 (mismatch)"
    else:
        breakdown["uom"] = "0"

    # ── 6. Item-keyword ───────────────────────────────────────
    item_kw = (budget_item.get("Item") or "").lower()
    if item_kw and item_kw in _norm(gfc_text):
        score += 7;  breakdown["item_kw"] = f"+7 ('{item_kw}')"
    else:
        breakdown["item_kw"] = "0"

    # ── 7. Pax ────────────────────────────────────────────────
    gp = set(re.findall(r"\b(\d+)\s*pax\b", _norm(gfc_text)))
    bp = set(re.findall(r"\b(\d+)\s*pax\b", _norm(bud_text)))
    if gp and bp:
        if gp == bp:
            score += 3;   breakdown["pax"] = f"+3 ({gp})"
        else:
            score -= 5;   breakdown["pax"] = f"-5 ({gp}≠{bp})"

    # ── 8. Single/Double ──────────────────────────────────────
    def _side(s):
        t = _norm(s)
        if re.search(r"\b(double|twin|two)\b", t):  return "double"
        if re.search(r"\b(single|one)\b",      t):  return "single"
        return None
    gs, bs = _side(gfc_text), _side(bud_text)
    if gs and bs:
        if gs == bs:
            score += 4;   breakdown["side"] = f"+4 (both {gs})"
        else:
            score -= 6;   breakdown["side"] = f"-6 ({gs}≠{bs})"

    # ── 9. Seater ─────────────────────────────────────────────
    gse = set(re.findall(r"\b(\d+)\s*seater\b", _norm(gfc_text)))
    bse = set(re.findall(r"\b(\d+)\s*seater\b", _norm(bud_text)))
    if gse and bse:
        if gse == bse:
            score += 3;   breakdown["seater"] = f"+3 ({gse})"
        else:
            score -= 5;   breakdown["seater"] = f"-5 ({gse}≠{bse})"

    # ── 10. Size / config ─────────────────────────────────────
    gfc_sz = _norm(gfc_row.get("size") or "")
    bud_cf = _norm(budget_item.get("Config (size/spec)") or "")
    if gfc_sz and bud_cf:
        g_dims = set(re.findall(r"\d+[\d.'\"]*", gfc_sz))
        b_dims = set(re.findall(r"\d+[\d.'\"]*", bud_cf))
        if g_dims and b_dims:
            shared = g_dims & b_dims
            if shared:
                pts = min(6, len(shared) * 2)
                score += pts;  breakdown["size"] = f"+{pts} ({sorted(shared)})"
            else:
                score -= 3;    breakdown["size"] = f"-3 ({sorted(g_dims)}≠{sorted(b_dims)})"
        else:
            breakdown["size"] = "0 (no dims)"
    else:
        breakdown["size"] = "0"

    # ── 11. Finish code ───────────────────────────────────────
    code = _norm(gfc_row.get("finish_code") or "").strip()
    if code and len(code) >= 3:
        bud_spec = _norm(budget_item.get("Specification") or
                         budget_item.get("Config (size/spec)") or "")
        bud_desc = _norm(budget_item.get("Description") or "")
        if code in bud_spec or code in bud_desc:
            score += 10;   breakdown["finish_code"] = f"+10 ('{code}')"
        else:
            breakdown["finish_code"] = "0 (not found)"
    else:
        breakdown["finish_code"] = "0"

    return (max(0, min(score, 100)), breakdown)


# ── Best-match selector ───────────────────────────────────────────────────────

def best_match(
    gfc_row:             dict,
    budget_for_category: list,
    category_exists:     bool,
) -> tuple[Optional[dict], int, str, dict]:
    """
    Find the best-matching BOQ item for a classified GFC row.

    Parameters
    ----------
    gfc_row              Dict with at minimum: category, subcategory, description,
                         areas, product_name, brand, uom, size, finish_code.
    budget_for_category  BOQ items pre-filtered to gfc_row's category.
    category_exists      True if ANY BOQ items exist for this category.

    Returns
    -------
    (best_item | None, score, status_emoji_string, breakdown_dict)
    """
    cat = (gfc_row.get("category") or "").upper()

    if not category_exists:
        return (
            None, 0, "🔴 Not in BOQ",
            {"category": f"NEW SCOPE — '{cat}' not found in BOQ at all"},
        )

    best_item, best_score, best_bd = None, 0, {}
    for item in budget_for_category:
        sc, bd = score_match(gfc_row, item)
        if sc > best_score:
            best_score, best_item, best_bd = sc, item, bd

    if best_score >= MATCH_THRESHOLDS["auto"]:
        status = "🟢 Matched"
    elif best_score >= MATCH_THRESHOLDS["suggested"]:
        status = "🟡 Suggested"
    else:
        status    = "🔴 Not in BOQ"
        best_item = None

    return (best_item, best_score, status, best_bd)

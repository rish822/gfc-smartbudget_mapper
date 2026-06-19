"""
Streamlit front-end for the Smart Budget ↔ GFC Mapper.

Run with:
    pip install streamlit pandas openpyxl
    streamlit run app.py

Modes:
  • Upload a GFC only            → Engine 1: Category / Sub-category mapping
  • Upload a GFC + a BOQ/Budget  → Engine 1 + Engine 2: BOQ match + score/status

Vision (Signal 5):
  Enable in the sidebar to use claude-haiku-4-5-20251001 vision for MEDIUM/LOW confidence
  items that have a product render embedded in the GFC sheet.  Requires an
  Anthropic API key (set ANTHROPIC_API_KEY in Streamlit secrets or paste below).
"""
import os
import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

# Make the engine package importable regardless of where streamlit is launched
ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from engine.gfc_mapping_engine import process_gfc, load_category_master, MATCH_THRESHOLDS
from engine.boq_loaders.auto import load_boq
from engine.image_classifier import IMAGE_THRESHOLD, MAX_IMAGE_CALLS

DEFAULT_MASTER = ROOT / "data" / "Category_master_for_GFC.xlsx"


# ── helpers ────────────────────────────────────────────────────────────────────

def _save_temp(uploaded_file) -> str:
    """Persist an uploaded file to a temp path (openpyxl needs a real path)."""
    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path


def _tier(conf: int) -> str:
    return "🟢 HIGH" if conf >= 75 else ("🟡 MEDIUM" if conf >= 50 else "🔴 LOW")


def run_engine(
    gfc_path: str,
    master_path: str | None,
    boq_path: str | None,
    vision_enabled: bool = False,
    vision_api_key: str | None = None,
) -> pd.DataFrame:
    """Core: run the GFC mapping pipeline and return a tidy DataFrame."""
    master = load_category_master(master_path) if master_path else []
    budget = load_boq(boq_path) if boq_path else []
    run_match = bool(budget)

    results = process_gfc(
        gfc_path, budget, master,
        run_match=run_match,
        vision_enabled=vision_enabled,
        vision_api_key=vision_api_key or None,
    )

    rows = []
    for el in results:
        row = {
            "GFC Sheet":     el.gfc_sheet,
            "Row":           el.gfc_row_idx,
            "GFC Line Item": el.raw_description,
            "Area/Location": el.raw_areas,
            "Category":      el.master_category or el.detected_category or "",
            "Sub-Category":  el.master_subcategory or "(blank — review)",
            "Confidence":    el.classification_confidence,
            "Tier":          _tier(el.classification_confidence),
            "Method":        el.classification_method,
        }
        if vision_enabled:
            row["Vision Signal"] = (
                f"👁 {el.vision_subcategory} ({el.vision_confidence}%): {el.vision_reason}"
                if el.vision_subcategory else ""
            )
        if run_match:
            row.update({
                "Match Status":      el.match_status,
                "Score":             el.match_score,
                "Matched BOQ Item":  el.matched_description or "",
                "BOQ Make":          el.matched_brand or "",
                "BOQ Rate":          el.matched_bcs_rate or "",
            })
        rows.append(row)
    return pd.DataFrame(rows)


# ── UI ──────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="GFC Line Item Mapper", layout="wide")
st.title("GFC Line Item → Category / Sub-category Mapper")
st.caption("Engine 1 classifies every GFC line item to the Category Master taxonomy. "
           "Add a BOQ to also run Engine 2 (BOQ matching + score).")

with st.sidebar:
    st.header("Inputs")
    gfc_file = st.file_uploader("GFC Excel  (required)", type=["xlsx", "xls"])
    master_file = st.file_uploader("Category Master  (optional)", type=["xlsx", "xls"])
    boq_file = st.file_uploader("BOQ / Smart Budget  (optional → runs Engine 2)",
                                type=["xlsx", "xls"])

    st.divider()
    st.subheader("Match thresholds (Engine 2)")
    auto_th = st.slider("🟢 Auto-match ≥", 50, 95, MATCH_THRESHOLDS["auto"])
    sugg_th = st.slider("🟡 Suggested ≥", 30, 74, MATCH_THRESHOLDS["suggested"])

    st.divider()
    st.subheader("👁 Vision (Signal 5)")
    st.caption(
        f"Uses claude-haiku-4-5-20251001 to read product renders embedded in the GFC. "
        f"Only triggered when Engine 1 confidence < {IMAGE_THRESHOLD}. "
        f"Capped at {MAX_IMAGE_CALLS} API calls per run."
    )
    vision_on = st.toggle("Enable Vision", value=False)
    if vision_on:
        # Try Streamlit secrets first (set in Community Cloud dashboard)
        _secret_key = st.secrets.get("ANTHROPIC_API_KEY", "") if hasattr(st, "secrets") else ""
        if _secret_key:
            st.success("API key loaded from Streamlit secrets.")
            vision_key = _secret_key
        else:
            vision_key = st.text_input(
                "Anthropic API key",
                type="password",
                placeholder="sk-ant-…  (or set ANTHROPIC_API_KEY in secrets)",
                help="Required for Vision. Not stored — only used during this run.",
            )
    else:
        vision_key = ""

if not gfc_file:
    st.info("⬅️  Upload a GFC Excel file to begin.")
    st.stop()

# Preview every sheet so the user sees the multi-sheet structure
with st.expander("Preview uploaded GFC (per sheet)", expanded=False):
    try:
        xls = pd.ExcelFile(gfc_file)
        for sh in xls.sheet_names[:8]:
            st.markdown(f"**{sh}**")
            st.dataframe(pd.read_excel(xls, sheet_name=sh, nrows=4), use_container_width=True)
        if len(xls.sheet_names) > 8:
            st.caption(f"… and {len(xls.sheet_names) - 8} more sheets")
    except Exception as e:
        st.warning(f"Could not preview: {e}")

if st.button("▶  Run Mapping", type="primary"):
    MATCH_THRESHOLDS["auto"] = auto_th
    MATCH_THRESHOLDS["suggested"] = sugg_th

    if vision_on and not vision_key:
        st.error("Vision is enabled but no API key was provided. "
                 "Paste your Anthropic API key in the sidebar or disable Vision.")
        st.stop()

    gfc_path = _save_temp(gfc_file)
    master_path = _save_temp(master_file) if master_file else (
        str(DEFAULT_MASTER) if DEFAULT_MASTER.exists() else None)
    boq_path = _save_temp(boq_file) if boq_file else None

    spinner_msg = "Classifying line items…"
    if vision_on:
        spinner_msg = "Classifying line items (Vision enabled — reading product renders)…"

    try:
        with st.spinner(spinner_msg):
            result = run_engine(
                gfc_path, master_path, boq_path,
                vision_enabled=vision_on,
                vision_api_key=vision_key if vision_on else None,
            )
    finally:
        for p in (gfc_path, boq_path):
            if p and os.path.exists(p):
                os.remove(p)
        if master_file and master_path and os.path.exists(master_path):
            os.remove(master_path)

    if result.empty:
        st.error("No mappable line items found. Check that the GFC has recognisable "
                 "sheet names and column headers.")
        st.stop()

    mode = "Engine 1 + Engine 2 (BOQ match)" if boq_path else "Engine 1 only (classification)"
    if vision_on:
        mode += " + Vision"
    st.success(f"Done — {len(result)} line items mapped.  Mode: {mode}")

    # Vision summary (only shown when vision was on)
    if vision_on and "Vision Signal" in result.columns:
        vision_hits = result["Vision Signal"].astype(bool).sum()
        st.info(f"👁 Vision enriched {vision_hits} / {len(result)} rows "
                f"(items with Engine 1 confidence < {IMAGE_THRESHOLD} that had an embedded image).")

    # ── summary metrics ──────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Line items", len(result))
    c2.metric("🟢 High conf.", int((result["Confidence"] >= 75).sum()))
    c3.metric("🟡 Medium",     int(((result["Confidence"] >= 50) & (result["Confidence"] < 75)).sum()))
    c4.metric("🔴 Low / blank", int((result["Confidence"] < 50).sum()))

    if "Match Status" in result.columns:
        a = int(result["Match Status"].astype(str).str.contains("🟢").sum())
        s = int(result["Match Status"].astype(str).str.contains("🟡").sum())
        n = len(result) - a - s
        m1, m2, m3 = st.columns(3)
        m1.metric("🟢 Auto-matched", a)
        m2.metric("🟡 Suggested", s)
        m3.metric("🔴 Not in BOQ", n)

    # ── filters ──────────────────────────────────────────────────────────────
    fcol1, fcol2 = st.columns(2)
    cats = ["(all)"] + sorted(result["Category"].unique().tolist())
    pick_cat = fcol1.selectbox("Filter by Category", cats)
    sheets = ["(all)"] + sorted(result["GFC Sheet"].unique().tolist())
    pick_sheet = fcol2.selectbox("Filter by GFC Sheet", sheets)

    view = result
    if pick_cat != "(all)":
        view = view[view["Category"] == pick_cat]
    if pick_sheet != "(all)":
        view = view[view["GFC Sheet"] == pick_sheet]

    st.dataframe(view, use_container_width=True, hide_index=True)

    # ── downloads ──────────────────────────────────────────────────────────────
    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇  Download mapping (CSV)",
        result.to_csv(index=False).encode("utf-8"),
        file_name="gfc_category_subcategory_mapping.csv",
        mime="text/csv",
    )
    import io
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        result.to_excel(xw, index=False, sheet_name="GFC Mapping")
    d2.download_button(
        "⬇  Download mapping (Excel)",
        buf.getvalue(),
        file_name="gfc_category_subcategory_mapping.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

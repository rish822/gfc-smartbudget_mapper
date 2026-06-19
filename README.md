# Smart Budget ↔ GFC Mapper

Auto-classifies every GFC (Good For Construction) line item to a **Category /
Sub-category** from the Category Master, and — when a BOQ / Smart Budget is
supplied — matches each item to its BOQ counterpart with a confidence score.

- **Engine 1** (`engine/gfc_classifier.py`) — GFC line item → Category + Sub-category.
  Sub-categories are strictly constrained to the chosen category (no cross-category leakage).
- **Engine 2** (`engine/boq_matcher.py`) — classified item → best BOQ match → 🟢 Auto / 🟡 Suggested / 🔴 Not in BOQ.

## Try it (web)

A Streamlit app: upload a GFC (and optionally a BOQ / Category Master), run, and
download the mapping as CSV or Excel.

### Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

### Deploy a public link (Streamlit Community Cloud — free)
1. Push this repo to GitHub (see below).
2. Go to **https://share.streamlit.io** → sign in with GitHub.
3. **New app** → pick this repo, branch `main`, main file `app.py` → **Deploy**.
4. Share the generated `https://<your-app>.streamlit.app` link.

## Project layout
```
app.py                       Streamlit web app
gui.py                       Desktop (tkinter) app — alternative to the web app
requirements.txt
engine/
  gfc_classifier.py          Engine 1
  boq_matcher.py             Engine 2
  gfc_mapping_engine.py      Orchestrator (sheet/header/row parsing + both engines)
  boq_loaders/auto.py        Auto-detecting BOQ loader
runners/output_builder.py    8-sheet enriched Excel report
```

> The Category Master is optional — without it, Engine 1 uses its built-in
> taxonomy. `data/` and `output/` are git-ignored so client files stay private.

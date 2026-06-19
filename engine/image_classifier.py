"""
Engine 1 — Signal 5: Image-based sub-category classifier

Extracts embedded product renders from GFC xlsx files, then calls
claude-haiku-4-5-20251001 vision to propose a sub-category when text
signals yield MEDIUM (< 75) or LOW (< 50) confidence.

Design for token efficiency:
  • Only triggered when classification_confidence < IMAGE_THRESHOLD (default 75)
  • Takes the leftmost-column image per row (primary render, not option swatches)
  • Compresses/resizes to ≤ MAX_PX before encoding (saves ~60–80 % input tokens)
  • Uses claude-haiku-4-5-20251001 with max_tokens=120 (tiny prompt + tiny response)
  • Per-run hard cap: MAX_IMAGE_CALLS API calls (prevents runaway cost)
  • In-process cache: same (sheet, row) is never analysed twice
"""
from __future__ import annotations

import base64
import json
import re
import zipfile
from io import BytesIO
from typing import Optional
from xml.etree import ElementTree as ET

try:
    from PIL import Image as PILImage
    _PIL = True
except ImportError:
    _PIL = False

# ── Tunable constants ─────────────────────────────────────────────────────────
IMAGE_THRESHOLD  = 75    # only call vision below this Engine 1 confidence
MAX_IMAGE_CALLS  = 50    # hard cap per process_gfc() run
MAX_PX           = 512   # resize images to this max dimension (saves tokens)
MAX_KB_RAW       = 300   # skip resize only if image already ≤ this size

# XML namespaces used in xlsx drawing files
_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_WB  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


# ── xlsx image extraction ─────────────────────────────────────────────────────

def _sheet_id_map(xlsx_path: str) -> dict[str, int]:
    """Return {sheet_name: sheetId} by reading xl/workbook.xml."""
    result = {}
    try:
        with zipfile.ZipFile(xlsx_path) as z:
            root = ET.fromstring(z.read("xl/workbook.xml"))
            for sheet in root.iter(f"{{{_WB}}}sheet"):
                name = sheet.get("name", "")
                sid  = sheet.get("sheetId", "")
                if name and sid:
                    result[name] = int(sid)
    except Exception:
        pass
    return result


def get_sheet_images(xlsx_path: str, sheet_id: int) -> dict[int, tuple[bytes, str]]:
    """
    Extract embedded images from one sheet, keyed by 1-indexed Excel row.

    When a row has multiple images (e.g. Render col + Product Image col),
    only the leftmost column image is kept — that is typically the 3-D render
    rather than a swatch / laminate option.

    Returns {row: (image_bytes, ext)} where ext is "png" / "jpg" etc.
    Returns {} when the sheet has no drawing or parsing fails.
    """
    # row → (lowest_col, image_bytes, ext)
    best: dict[int, tuple[int, bytes, str]] = {}

    try:
        with zipfile.ZipFile(xlsx_path) as z:
            names = set(z.namelist())

            # 1. Find drawing relationship for this sheet
            rels_path = f"xl/worksheets/_rels/sheet{sheet_id}.xml.rels"
            if rels_path not in names:
                return {}

            drawing_file: Optional[str] = None
            for rel in ET.fromstring(z.read(rels_path)):
                if "drawing" in rel.get("Type", "").lower():
                    tgt = rel.get("Target", "")
                    # Target is like "../drawings/drawing1.xml"
                    drawing_file = "xl/drawings/" + tgt.split("/")[-1]
                    break

            if not drawing_file or drawing_file not in names:
                return {}

            # 2. Drawing → rId → media path
            rels_d = drawing_file.replace("xl/drawings/", "xl/drawings/_rels/") + ".rels"
            if rels_d not in names:
                return {}

            rid_to_media: dict[str, str] = {}
            for rel in ET.fromstring(z.read(rels_d)):
                if "image" in rel.get("Type", "").lower():
                    rid = rel.get("Id", "")
                    tgt = rel.get("Target", "").lstrip("./")
                    # "../media/image1.png" → "xl/media/image1.png"
                    media = "xl/media/" + tgt.split("/")[-1]
                    if media in names:
                        rid_to_media[rid] = media

            if not rid_to_media:
                return {}

            # 3. Parse anchors → row, col, rId
            for anchor in ET.fromstring(z.read(drawing_file)):
                from_elem = anchor.find(f"{{{_XDR}}}from")
                if from_elem is None:
                    continue
                row_el = from_elem.find(f"{{{_XDR}}}row")
                col_el = from_elem.find(f"{{{_XDR}}}col")
                if row_el is None:
                    continue

                row_0 = int(row_el.text or 0)
                col_0 = int(col_el.text or 0) if col_el is not None else 9999
                row_1 = row_0 + 1  # → 1-indexed Excel row

                blip = anchor.find(f".//{{{_A}}}blip")
                if blip is None:
                    continue
                rid = blip.get(f"{{{_R}}}embed", "")
                if rid not in rid_to_media:
                    continue

                # Keep only the leftmost-column image per row
                ext = rid_to_media[rid].rsplit(".", 1)[-1].lower()
                if ext not in ("png", "jpg", "jpeg", "gif", "webp"):
                    ext = "png"
                if row_1 not in best or col_0 < best[row_1][0]:
                    best[row_1] = (col_0, z.read(rid_to_media[rid]), ext)

    except Exception:
        pass

    return {row: (data[1], data[2]) for row, data in best.items()}


def extract_all_images(
    xlsx_path: str, sheet_names: list[str]
) -> dict[str, dict[int, tuple[bytes, str]]]:
    """
    Extract images for a list of sheet names.
    Returns {sheet_name: {row_1indexed: (image_bytes, ext)}}.
    Silently skips sheets with no embedded images.
    """
    sid_map = _sheet_id_map(xlsx_path)
    result: dict[str, dict[int, tuple[bytes, str]]] = {}
    for name in sheet_names:
        sid = sid_map.get(name)
        if sid is None:
            continue
        imgs = get_sheet_images(xlsx_path, sid)
        if imgs:
            result[name] = imgs
    return result


# ── Image compression ─────────────────────────────────────────────────────────

def _compress(image_bytes: bytes, ext: str) -> tuple[bytes, str]:
    """
    Resize to MAX_PX (if PIL available) and return (bytes, ext).
    Falls back to raw bytes if PIL unavailable or the image is small.
    """
    if not _PIL or len(image_bytes) <= MAX_KB_RAW * 1024:
        return image_bytes, ext

    try:
        img = PILImage.open(BytesIO(image_bytes)).convert("RGB")
        img.thumbnail((MAX_PX, MAX_PX), PILImage.LANCZOS)
        out = BytesIO()
        img.save(out, format="JPEG", optimize=True, quality=75)
        return out.getvalue(), "jpg"
    except Exception:
        return image_bytes, ext


# ── Claude claude-haiku-4-5-20251001 vision call ──────────────────────────────────────────────────

_VISION_PROMPT = (
    "This is a product render / 3-D image from a GFC (Good For Construction) sheet.\n"
    "The item's category is already confirmed as '{category}'.\n"
    "GFC description: \"{description}\".\n"
    "Based ONLY on what you see in the image, which sub-category fits best?\n"
    "Valid options: {options}\n"
    "Reply in JSON only (no markdown): "
    '{"sub": "<exact option>", "conf": <0-100>, "why": "<≤6 words>"}'
)


def classify_image(
    image_bytes: bytes,
    category: str,
    valid_subcats: set,
    description: str = "",
    image_ext: str = "png",
    api_key: Optional[str] = None,
) -> tuple[Optional[str], int, str]:
    """
    Ask claude-haiku-4-5-20251001 vision to pick a sub-category from the render image.

    Returns (subcategory | None, confidence 0-100, one-line reason).
    Returns (None, 0, <error>) on any failure so callers can fall back gracefully.
    """
    if not valid_subcats:
        return None, 0, "no valid sub-cats"

    try:
        import anthropic  # lazy — only needed when vision enabled
    except ImportError:
        return None, 0, "anthropic package not installed — run: pip install anthropic"

    try:
        kb, ext = _compress(image_bytes, image_ext)

        b64        = base64.standard_b64encode(kb).decode()
        media_type = "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"

        options = " | ".join(sorted(valid_subcats))
        prompt  = _VISION_PROMPT.format(
            category    = category,
            description = description or "—",
            options     = options,
        )

        client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()
        resp   = client.messages.create(
            model      = "claude-haiku-4-5-20251001",
            max_tokens = 120,
            messages   = [{
                "role": "user",
                "content": [
                    {"type": "image",
                     "source": {"type": "base64",
                                "media_type": media_type,
                                "data": b64}},
                    {"type": "text", "text": prompt},
                ],
            }],
        )

        raw = resp.content[0].text.strip()
        raw = re.sub(r"```(?:json)?|```", "", raw).strip()
        parsed = json.loads(raw)

        sub  = str(parsed.get("sub", "")).strip().upper()
        conf = max(0, min(100, int(parsed.get("conf", 0))))
        why  = str(parsed.get("why", ""))

        # Exact match (case-insensitive)
        upper_map = {s.upper(): s for s in valid_subcats}
        if sub in upper_map:
            return upper_map[sub], conf, why

        # Partial match — tolerate slight wording differences
        for canonical_upper, canonical in upper_map.items():
            if canonical_upper in sub or sub in canonical_upper:
                return canonical, max(0, conf - 10), why + " (partial)"

        return None, 0, f"vision sub '{sub}' not in valid set"

    except json.JSONDecodeError:
        return None, 0, "vision: JSON parse failed"
    except Exception as exc:
        return None, 0, f"vision: {type(exc).__name__}: {exc}"

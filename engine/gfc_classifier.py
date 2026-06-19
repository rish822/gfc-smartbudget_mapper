"""
Engine 1 — GFC Line Item → Category Master Classification

Maps every extracted GFC row to (Category, Sub-category, Item) from the
Category Master before any BOQ matching runs.

CORE PRINCIPLE (enforced strictly):
  Once the CATEGORY of a line item is determined, the SUB-CATEGORY is only
  ever drawn from that category's own sub-categories in the Master.  A
  sub-category from a different category can NEVER be assigned.  If no valid
  sub-category is found, the result is left blank (category-only) rather than
  borrowing a sub-category from elsewhere.

Pipeline per line item:
  STEP 0  Category overrides   — material/context rules that switch the category
  STEP 1  Determine valid sub-categories for the (possibly overridden) category
  STEP 2  Four signals propose a sub-category — ALL constrained to that category
            Signal 2 uses DESCRIPTION + PRODUCT_NAME only (no sub_hint) to
            prevent the sheet hint name from contaminating keyword matches.
  STEP 3  Hard guard: drop any proposed sub-category not valid for the category
  STEP 4  Special area-based hint for ACOUSTIC and PAINT categories

Four signals (highest → lowest priority):
  Signal 1  Row-level explicit hint   ("Furniture Type" / "Wall/Ceiling" col) conf 92
  Signal 2  Phrase keyword match      (category-scoped phrase list)    conf 65–90
  Signal 3  Master item-name match    (stemmed overlap vs that category) conf 40–82
  Signal 4  Sheet-level sub_hint      (from SHEET_RULES L1)            conf 65

Confidence tiers:
  HIGH    ≥ 75   sub-category confirmed
  MEDIUM  50–74  likely correct, spot-check
  LOW     < 50   needs human review / left blank
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional


# ── Shared text helpers ───────────────────────────────────────────────────────

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]", " ", (s or "").lower())).strip()


def _light_stem(tok: str) -> str:
    """Strip common inflection endings (longest-first, no duplicates)."""
    t = tok.lower()
    for suf in ("ings", "ing", "ies", "ed", "es", "s"):
        if len(t) > len(suf) + 2 and t.endswith(suf):
            return t[:-3] + "y" if suf == "ies" else t[:-len(suf)]
    return t


# ══════════════════════════════════════════════════════════════════════════════
# AUTHORITATIVE TAXONOMY — valid sub-categories per category (from Category Master)
# Updated to match "Final Category Master.xlsx"
# ══════════════════════════════════════════════════════════════════════════════
MASTER_SUBCATS: dict[str, set] = {
    "ACOUSTIC":             {"CEILING", "LIGHTS", "PARTITION", "WALL SOLUTIONS"},
    "CEILING":              {"DECORATIVE CEILING", "GYPSUM CEILING", "METAL CEILING", "WOOD CEILING"},
    "CIVIL":                {"COUNTER TOPS", "DADO TILE", "DEMOLITION AND DISMANTLING",
                             "FABRICATION", "MASONARY WORKS", "PLUMBING",
                             "SANITARY FIXTURES", "WET WORKS"},
    "DECORATIVES":          {"ACOUSTIC WALL SOLUTIONS", "ARTIFACTS AND ACCESSORIES", "BLIND",
                             "GRAPHICS", "PANELLING", "SIGNAGES AND BRANDING",
                             "SKIRTING AND PROFILE", "UTILITIES", "WALL CLADDING"},
    "DESIGN & MAINTENANCE": {"DESIGN", "MAINTENANACE"},
    "ELECTRICAL AND NETWORKING": {
        "CABLE TRAY & RACEWAYS", "CIRCUIT WIRING", "EARTHING SYSTEM",
        "ELECTRICAL STANDARD", "LT CABLES", "NETWORKING", "POINT WIRING",
        "POWER OUTLETS & WIRING", "POWER PANELS & DBS", "SUBMAINS",
        "SUBMAINS & POINT WIRING", "TERMINATION",
        "WORKSTATION SWITCHSOCKET WITH BACKBOX"},
    "FIRE AND SECURITY WORK": {
        "ACCESS CONTROL SYSTEM", "CCTV SYSTEM", "FIRE ALARM SYSTEM",
        "FIRE EXTINGUISHER AND CLEAN AGENT", "FIRE SIGNAGES", "PA SYSTEM",
        "SPRINKLER SYSTEMS"},
    "FLOORING":             {"FLOOR LEVELLING", "RAISED FLOORING", "RESILIENT FLOORING",
                             "TEXTILE FLOORING", "TILE FLOORING"},
    "FURNITURE":            {"CHAIRS", "CUSTOM FURNITURE", "DESIGN & ACCESSORIES",
                             "LOOSE FURNITURE", "MODULAR FURNITURE"},
    "HVAC WORK":            {
        "AC UNIT", "AHU", "ATC CABLE", "AUTO SEQUENTIAL PANEL", "CHILLER",
        "DIGITAL THERMOSTAT", "DUCTABLE_LOW SIDE", "DUCTABLE_LOW SIDE_INSULATION",
        "ELECTRICAL WORK", "HIGH SIDE WORKS", "HVAC STANDARD", "INSTALLATION",
        "LOW SIDE WORKS"},
    "LIGHTING":             {"ACOUSTIC LIGHTS", "AMBIENT LIGHTS", "ARCHITECTURAL LIGHTS",
                             "DECORATIVE LIGHTS"},
    "PAINT":                {"CEILING PAINT", "DUCO PAINT", "DUCT PAINT",
                             "TEXTURE PAINT", "WALL PAINT"},
    "PARTITIONS AND DOORS": {"DOORS", "GLASS PARTITION", "PARTITION", "WINDOWS"},
    "SURFACE AND FINISHES": {"WOODEN"},
}


# ══════════════════════════════════════════════════════════════════════════════
# STEP 0 — CATEGORY OVERRIDES
# Each entry: (trigger_phrase, also_any_of | None, new_category, new_subcategory)
#   also_any_of = None         → unconditional: trigger phrase alone switches category
#   also_any_of = (a, b, ...)   → only switch if trigger AND any of these also appear
# Checked against the composite text BEFORE the four signals run.
# ══════════════════════════════════════════════════════════════════════════════
_FURNITURE_NOUNS = (
    "reception table", "conference table", "book shelf", "bookshelf", "credenza",
    "table top", "side table", "study table", "training table", "desk", "wardrobe",
    "cabinet", "shelf", "pantry", "vanity", "modesty",
)

_CATEGORY_OVERRIDES: list[tuple[str, Optional[tuple], str, str]] = [
    # ── Wall tiles → CIVIL / DADO TILE  (floor tiles stay FLOORING / TILE FLOORING) ──
    ("wall tiles",     None, "CIVIL", "DADO TILE"),
    ("wall tile",      None, "CIVIL", "DADO TILE"),
    ("dado tile",      None, "CIVIL", "DADO TILE"),
    ("dado",           None, "CIVIL", "DADO TILE"),
    ("ceramic dado",   None, "CIVIL", "DADO TILE"),
    ("porcelain dado", None, "CIVIL", "DADO TILE"),
    ("vitrified dado", None, "CIVIL", "DADO TILE"),

    # ── Doors by material ──────────────────────────────────────────────────
    ("glass door",     None, "PARTITIONS AND DOORS", "DOORS"),
    ("aluminium door", None, "PARTITIONS AND DOORS", "DOORS"),
    ("aluminum door",  None, "PARTITIONS AND DOORS", "DOORS"),
    ("toughened glass door", None, "PARTITIONS AND DOORS", "DOORS"),
    ("washroom door",  None, "PARTITIONS AND DOORS", "DOORS"),
    ("toilet door",    None, "PARTITIONS AND DOORS", "DOORS"),
    ("wc door",        None, "PARTITIONS AND DOORS", "DOORS"),

    # ── Solid-surface finishes ─────────────────────────────────────────────
    ("corian",         None, "SURFACE AND FINISHES", "WOODEN"),
    ("hi macs",        None, "SURFACE AND FINISHES", "WOODEN"),
    ("solid surface",  None, "SURFACE AND FINISHES", "WOODEN"),

    # ── Laminate / Fluted panel / Veneer APPLIED ON a furniture piece ──────
    ("laminate",   _FURNITURE_NOUNS, "SURFACE AND FINISHES", "WOODEN"),
    ("laminated",  _FURNITURE_NOUNS, "SURFACE AND FINISHES", "WOODEN"),
    ("fluted",     _FURNITURE_NOUNS, "SURFACE AND FINISHES", "WOODEN"),
    ("panelling",  _FURNITURE_NOUNS, "SURFACE AND FINISHES", "WOODEN"),
    ("paneling",   _FURNITURE_NOUNS, "SURFACE AND FINISHES", "WOODEN"),
    ("veneer",     _FURNITURE_NOUNS, "SURFACE AND FINISHES", "WOODEN"),
]


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2, Signal 2 — CATEGORY-SCOPED phrase → sub-category lookup.
# KEY DESIGN: These phrases are matched against DESCRIPTION + PRODUCT_NAME only
# (sub_hint / sheet name is EXCLUDED from Signal 2).  This prevents the sheet
# hint value (e.g. "METAL CEILING", "PARTITION") from accidentally triggering a
# keyword match that overrides the sheet hint in the wrong direction.
# ══════════════════════════════════════════════════════════════════════════════
_SUBCAT_PHRASES_BY_CAT: dict[str, list[tuple[str, str]]] = {

    "FURNITURE": [
        # MODULAR FURNITURE — workstations, meeting/conference tables, pedestals
        ("workstation partition",   "MODULAR FURNITURE"),
        ("linear workstation",      "MODULAR FURNITURE"),
        ("l-shape workstation",     "MODULAR FURNITURE"),
        ("l shape workstation",     "MODULAR FURNITURE"),
        ("workstation",             "MODULAR FURNITURE"),
        ("height adjustable",       "MODULAR FURNITURE"),
        ("meeting table",           "MODULAR FURNITURE"),
        ("meeting room table",      "MODULAR FURNITURE"),
        ("conference table",        "MODULAR FURNITURE"),
        ("cafeteria table",         "MODULAR FURNITURE"),
        ("cafe table",              "MODULAR FURNITURE"),
        ("reception table",         "MODULAR FURNITURE"),
        ("phone booth",             "MODULAR FURNITURE"),
        ("folding table",           "MODULAR FURNITURE"),
        ("high table",              "MODULAR FURNITURE"),
        ("high stand table",        "MODULAR FURNITURE"),
        ("3 drawer pedestal",       "MODULAR FURNITURE"),
        ("pedestal drawers",        "MODULAR FURNITURE"),
        ("pedestal body",           "MODULAR FURNITURE"),
        ("pedestal",                "MODULAR FURNITURE"),
        # CUSTOM FURNITURE  (incl. carpentry/wooden/flush doors)
        ("reception desk",          "CUSTOM FURNITURE"),
        ("host desk",               "CUSTOM FURNITURE"),
        ("cabin table",             "CUSTOM FURNITURE"),
        ("credenza",                "CUSTOM FURNITURE"),
        ("book shelf",              "CUSTOM FURNITURE"),
        ("bookshelf",               "CUSTOM FURNITURE"),
        ("wardrobe",                "CUSTOM FURNITURE"),
        ("pantry",                  "CUSTOM FURNITURE"),
        ("booth seating",           "CUSTOM FURNITURE"),
        ("step seating",            "CUSTOM FURNITURE"),
        ("ledge above chair",       "CUSTOM FURNITURE"),
        ("ledge seating",           "CUSTOM FURNITURE"),
        ("ledge",                   "CUSTOM FURNITURE"),
        ("custom seating",          "CUSTOM FURNITURE"),
        ("custom storage",          "CUSTOM FURNITURE"),
        ("low height storage",      "CUSTOM FURNITURE"),
        ("over head storage",       "CUSTOM FURNITURE"),
        ("overhead storage",        "CUSTOM FURNITURE"),
        ("under counter",           "CUSTOM FURNITURE"),
        ("janitor",                 "CUSTOM FURNITURE"),
        ("tv unit",                 "CUSTOM FURNITURE"),
        ("tvunit",                  "CUSTOM FURNITURE"),
        ("planter box",             "CUSTOM FURNITURE"),
        ("puja unit",               "CUSTOM FURNITURE"),
        ("wet bar",                 "CUSTOM FURNITURE"),
        ("washroom vanity",         "CUSTOM FURNITURE"),
        ("vanity unit",             "CUSTOM FURNITURE"),
        ("vanity",                  "CUSTOM FURNITURE"),
        ("open shelf",              "CUSTOM FURNITURE"),
        ("designer shelf",          "CUSTOM FURNITURE"),
        ("study table",             "CUSTOM FURNITURE"),
        ("training table",          "CUSTOM FURNITURE"),
        ("dummy boxing",            "CUSTOM FURNITURE"),
        ("modesty",                 "CUSTOM FURNITURE"),
        ("shutter",                 "CUSTOM FURNITURE"),
        # Carpentry doors (wooden / flush / solid) — stay in FURNITURE
        ("solid flush door",        "CUSTOM FURNITURE"),
        ("flush door",              "CUSTOM FURNITURE"),
        ("wooden door",             "CUSTOM FURNITURE"),
        ("sliding flush door",      "CUSTOM FURNITURE"),
        ("vision panel",            "CUSTOM FURNITURE"),
        ("door",                    "CUSTOM FURNITURE"),
        # LOOSE FURNITURE
        ("lounge chair",            "LOOSE FURNITURE"),
        ("cafe chair",              "LOOSE FURNITURE"),
        ("dining chair",            "LOOSE FURNITURE"),
        ("dining table",            "LOOSE FURNITURE"),
        ("bar stool",               "LOOSE FURNITURE"),
        ("centre table",            "LOOSE FURNITURE"),
        ("center table",            "LOOSE FURNITURE"),
        ("coffee table",            "LOOSE FURNITURE"),
        ("console table",           "LOOSE FURNITURE"),
        ("display table",           "LOOSE FURNITURE"),
        ("sofa",                    "LOOSE FURNITURE"),
        ("pouffe",                  "LOOSE FURNITURE"),
        ("ottomon",                 "LOOSE FURNITURE"),
        ("ottoman",                 "LOOSE FURNITURE"),
        ("lounger",                 "LOOSE FURNITURE"),
        ("bench",                   "LOOSE FURNITURE"),
        ("pod seating",             "LOOSE FURNITURE"),
        ("break out",               "LOOSE FURNITURE"),
        ("outdoor furniture",       "LOOSE FURNITURE"),
        ("seating stool",           "LOOSE FURNITURE"),
        # CHAIRS (after LOOSE FURNITURE so "cafe chair" is caught first)
        ("high back chair",         "CHAIRS"),
        ("medium back chair",       "CHAIRS"),
        ("executive chair",         "CHAIRS"),
        ("md cabin chair",          "CHAIRS"),
        ("cabin hb chair",          "CHAIRS"),
        ("task chair",              "CHAIRS"),
        ("office chair",            "CHAIRS"),
        ("visitor chair",           "CHAIRS"),
        ("training chair",          "CHAIRS"),
        ("gaming chair",            "CHAIRS"),
        ("mesh chair",              "CHAIRS"),
        ("wings mesh",              "CHAIRS"),
        ("cafeteria chair",         "CHAIRS"),
        ("high chair",              "CHAIRS"),
        ("chair",                   "CHAIRS"),
    ],

    "FLOORING": [
        # TILE FLOORING — floor tiles only (wall tiles handled by category override)
        ("floor tiles",             "TILE FLOORING"),
        ("floor tile",              "TILE FLOORING"),
        ("vitrified tile",          "TILE FLOORING"),
        ("ceramic tile",            "TILE FLOORING"),
        ("porcelain tile",          "TILE FLOORING"),
        ("marble flooring",         "TILE FLOORING"),
        ("marble",                  "TILE FLOORING"),
        ("mosaic tile",             "TILE FLOORING"),
        ("subway tile",             "TILE FLOORING"),
        ("border tile",             "TILE FLOORING"),
        ("terrazo",                 "TILE FLOORING"),
        ("designer tile",           "TILE FLOORING"),
        ("tile",                    "TILE FLOORING"),
        # TEXTILE FLOORING
        ("carpet design",           "TEXTILE FLOORING"),
        ("carpet",                  "TEXTILE FLOORING"),
        ("rug",                     "TEXTILE FLOORING"),
        ("flocked flooring",        "TEXTILE FLOORING"),
        # RAISED FLOORING
        ("raised floor",            "RAISED FLOORING"),
        ("raised flooring",         "RAISED FLOORING"),
        # FLOOR LEVELLING
        ("floor levelling",         "FLOOR LEVELLING"),
        ("floor leveling",          "FLOOR LEVELLING"),
        ("ips",                     "FLOOR LEVELLING"),
        ("pcc",                     "FLOOR LEVELLING"),
        # RESILIENT FLOORING
        ("laminated wooden flooring","RESILIENT FLOORING"),
        ("wooden flooring",         "RESILIENT FLOORING"),
        ("rubber flooring",         "RESILIENT FLOORING"),
        ("linoleum",                "RESILIENT FLOORING"),
        ("anti static flooring",    "RESILIENT FLOORING"),
        ("antistatic flooring",     "RESILIENT FLOORING"),
        ("antistatic",              "RESILIENT FLOORING"),
        ("anti static",             "RESILIENT FLOORING"),
        ("spc flooring",            "RESILIENT FLOORING"),
        ("spc",                     "RESILIENT FLOORING"),
        ("lvt",                     "RESILIENT FLOORING"),
        ("vinyl",                   "RESILIENT FLOORING"),
    ],

    "CEILING": [
        ("laminated gypsum ceiling","GYPSUM CEILING"),
        ("gypsum false ceiling",    "GYPSUM CEILING"),
        ("gypsum cove ceiling",     "GYPSUM CEILING"),
        ("gypsum ceiling",          "GYPSUM CEILING"),
        ("gypsum",                  "GYPSUM CEILING"),
        ("open cell ceiling",       "METAL CEILING"),
        ("open cell",               "METAL CEILING"),
        ("metal baffle",            "METAL CEILING"),
        ("linear ceiling",          "METAL CEILING"),
        ("axiom trim",              "METAL CEILING"),
        ("modular tile ceiling",    "METAL CEILING"),
        ("metal ceiling",           "METAL CEILING"),
        ("grid ceiling",            "METAL CEILING"),
        # Wooden baffles/open cell → WOOD CEILING
        ("wooden open cell",        "WOOD CEILING"),
        ("wooden baffle",           "WOOD CEILING"),
        ("baffle ceiling",          "WOOD CEILING"),   # default baffle = wooden
        ("wood ceiling",            "WOOD CEILING"),   # explicit wood ceiling type
        # Decorative ceiling finishes
        ("wooden ceiling",          "DECORATIVE CEILING"),
        ("stretch ceiling",         "DECORATIVE CEILING"),
        ("sunlight ceiling",        "DECORATIVE CEILING"),
        ("cove ceiling",            "DECORATIVE CEILING"),
        ("decorative ceiling",      "DECORATIVE CEILING"),
    ],

    "PARTITIONS AND DOORS": [
        ("fixed glass partition",   "GLASS PARTITION"),
        ("above glass partition",   "GLASS PARTITION"),
        ("glass partition",         "GLASS PARTITION"),
        ("acoustic partition",      "PARTITION"),
        ("gypsum partition",        "PARTITION"),
        ("phenolic partition",      "PARTITION"),
        ("phenolic",                "PARTITION"),
        ("partition",               "PARTITION"),
        # Doors — longer/more-specific phrases first to beat "partition" in product names
        ("toilet flush door",       "DOORS"),
        ("washroom door",           "DOORS"),   # "Male & Female Washroom Door" etc.
        ("toilet door",             "DOORS"),
        ("wc door",                 "DOORS"),
        ("double leaf glass door",  "DOORS"),
        ("glass door",              "DOORS"),
        ("wooden door",             "DOORS"),
        ("panel door",              "DOORS"),
        ("flush door",              "DOORS"),
        ("double leaf",             "DOORS"),
        ("single leaf",             "DOORS"),
        ("door handle",             "DOORS"),
        ("door",                    "DOORS"),
        ("aluminium window",        "WINDOWS"),
        ("upvc window",             "WINDOWS"),
        ("glass window",            "WINDOWS"),
        ("wooden window",           "WINDOWS"),
        ("window",                  "WINDOWS"),
    ],

    "DECORATIVES": [
        ("honeycomb blind",         "BLIND"),
        ("roller blind",            "BLIND"),
        ("venetian blind",          "BLIND"),
        ("zebra blind",             "BLIND"),
        ("vertical blind",          "BLIND"),
        ("blind",                   "BLIND"),
        ("wallpaper",               "GRAPHICS"),
        ("glass film",              "GRAPHICS"),
        ("glass board",             "GRAPHICS"),
        ("wall framing",            "GRAPHICS"),
        ("pin up board",            "GRAPHICS"),
        ("pin up",                  "GRAPHICS"),
        ("decal",                   "GRAPHICS"),
        ("canvas frame",            "GRAPHICS"),
        ("frame",                   "GRAPHICS"),
        ("canvas",                  "GRAPHICS"),
        ("acoustic panel",          "ACOUSTIC WALL SOLUTIONS"),
        ("acoustic wall",           "ACOUSTIC WALL SOLUTIONS"),
        ("movable wall",            "ACOUSTIC WALL SOLUTIONS"),
        ("grooved panel",           "ACOUSTIC WALL SOLUTIONS"),
        ("moulded acoustic",        "ACOUSTIC WALL SOLUTIONS"),
        ("3d mdf panelling",        "PANELLING"),
        ("fluted panelling",        "PANELLING"),
        ("fluted panel",            "PANELLING"),
        ("laminate panelling",      "PANELLING"),
        ("veneer panelling",        "PANELLING"),
        ("decorative panelling",    "PANELLING"),
        ("custom panels",           "PANELLING"),
        ("pvc louvers",             "PANELLING"),
        ("charcoal boards",         "PANELLING"),
        ("pu stone panel",          "PANELLING"),
        ("wall panelling",          "PANELLING"),
        ("wall paneling",           "PANELLING"),
        ("panelling",               "PANELLING"),
        ("paneling",                "PANELLING"),
        ("fluted",                  "PANELLING"),
        ("metal wall cladding",     "WALL CLADDING"),
        ("wall cladding",           "WALL CLADDING"),
        ("signage",                 "SIGNAGES AND BRANDING"),
        ("branding",                "SIGNAGES AND BRANDING"),
        ("logo",                    "SIGNAGES AND BRANDING"),
        ("transition profile",      "SKIRTING AND PROFILE"),
        ("chair guard",             "SKIRTING AND PROFILE"),
        ("corner guard",            "SKIRTING AND PROFILE"),
        ("skirting",                "SKIRTING AND PROFILE"),
        ("transition",              "SKIRTING AND PROFILE"),
        ("washroom mirror",         "UTILITIES"),
        ("hardware",                "UTILITIES"),
        ("mirror",                  "UTILITIES"),
        ("artifact",                "ARTIFACTS AND ACCESSORIES"),
    ],

    "SURFACE AND FINISHES": [
        ("laminates",               "WOODEN"),
        ("laminate",                "WOODEN"),
        ("veneer",                  "WOODEN"),
        ("corian",                  "WOODEN"),
        ("hi macs",                 "WOODEN"),
        ("wooden",                  "WOODEN"),
    ],

    "CIVIL": [
        # SANITARY FIXTURES (always → CIVIL)
        ("water closet",            "SANITARY FIXTURES"),
        ("wash basin",              "SANITARY FIXTURES"),
        ("washbasin",               "SANITARY FIXTURES"),
        ("urinal sensor",           "SANITARY FIXTURES"),
        ("urinal senser",           "SANITARY FIXTURES"),
        ("urinal",                  "SANITARY FIXTURES"),
        ("pillar cock",             "SANITARY FIXTURES"),
        ("bib cock",                "SANITARY FIXTURES"),
        ("angle cock",              "SANITARY FIXTURES"),
        ("health faucet",           "SANITARY FIXTURES"),
        ("flush valve",             "SANITARY FIXTURES"),
        ("flush plate",             "SANITARY FIXTURES"),
        ("faucet",                  "SANITARY FIXTURES"),
        ("shower head",             "SANITARY FIXTURES"),
        ("shower arm",              "SANITARY FIXTURES"),
        ("shower",                  "SANITARY FIXTURES"),
        ("cistern",                 "SANITARY FIXTURES"),
        ("soap dispenser",          "SANITARY FIXTURES"),
        ("paper dispenser",         "SANITARY FIXTURES"),
        ("tissue paper holder",     "SANITARY FIXTURES"),
        ("toilet paper holder",     "SANITARY FIXTURES"),
        ("hand dryer",              "SANITARY FIXTURES"),
        ("towel warmer",            "SANITARY FIXTURES"),
        ("bottle trap",             "SANITARY FIXTURES"),
        ("waste coupling",          "SANITARY FIXTURES"),
        ("floor trap",              "SANITARY FIXTURES"),
        ("sink tap",                "SANITARY FIXTURES"),
        ("sink",                    "SANITARY FIXTURES"),
        ("diverter",                "SANITARY FIXTURES"),
        ("spout",                   "SANITARY FIXTURES"),
        ("wc",                      "SANITARY FIXTURES"),
        ("sanitary",                "SANITARY FIXTURES"),
        # DADO TILE
        ("dado tile",               "DADO TILE"),
        ("ceramic dado",            "DADO TILE"),
        ("porcelain dado",          "DADO TILE"),
        ("vitrified dado",          "DADO TILE"),
        ("wall tile",               "DADO TILE"),
        # COUNTER TOPS
        ("granite counter",         "COUNTER TOPS"),
        ("quartz counter",          "COUNTER TOPS"),
        ("counter top",             "COUNTER TOPS"),
        # FABRICATION
        ("fabricated structure",    "FABRICATION"),
        ("ms fabrication",          "FABRICATION"),
        ("fabrication",             "FABRICATION"),
        # WET WORKS
        ("waterproofing",           "WET WORKS"),
        ("water proofing",          "WET WORKS"),
        ("epoxy flooring",          "WET WORKS"),
        ("screeding",               "WET WORKS"),
        ("brick bat",               "WET WORKS"),
        ("brick bat coba",          "WET WORKS"),
        # MASONARY WORKS
        ("brick wall",              "MASONARY WORKS"),
        ("siporex block",           "MASONARY WORKS"),
        ("wall plaster",            "MASONARY WORKS"),
        ("wall punning",            "MASONARY WORKS"),
        ("ceiling plaster",         "MASONARY WORKS"),
        ("exposed brick",           "MASONARY WORKS"),
        ("platform",                "MASONARY WORKS"),
        # DEMOLITION AND DISMANTLING
        ("demolition",              "DEMOLITION AND DISMANTLING"),
        ("dismantling",             "DEMOLITION AND DISMANTLING"),
        ("dismantlement",           "DEMOLITION AND DISMANTLING"),
        ("debris removal",          "DEMOLITION AND DISMANTLING"),
        ("core cutting",            "DEMOLITION AND DISMANTLING"),
        # PLUMBING
        ("internal plumbing",       "PLUMBING"),
        ("plumbing",                "PLUMBING"),
    ],

    "PAINT": [
        ("ceiling paint",           "CEILING PAINT"),
        ("wall paint",              "WALL PAINT"),
        ("duco paint",              "DUCO PAINT"),
        ("duct paint",              "DUCT PAINT"),
        ("texture paint",           "TEXTURE PAINT"),
    ],

    "LIGHTING": [
        # ACOUSTIC LIGHTS
        ("acoustic linear light",   "ACOUSTIC LIGHTS"),
        ("acoustic moulded light",  "ACOUSTIC LIGHTS"),
        ("acoustic pendant light",  "ACOUSTIC LIGHTS"),
        ("acoustic light",          "ACOUSTIC LIGHTS"),
        # DECORATIVE LIGHTS — pendant/chandelier/floor lamps/decorative named
        ("linear suspension light", "DECORATIVE LIGHTS"),
        ("linear suspended light",  "DECORATIVE LIGHTS"),
        ("linear suspended",        "DECORATIVE LIGHTS"),
        ("decorative linear light", "DECORATIVE LIGHTS"),
        ("rectangular linear light","DECORATIVE LIGHTS"),
        ("suspended cylindrical",   "DECORATIVE LIGHTS"),
        ("chandelier",              "DECORATIVE LIGHTS"),
        ("pendant light",           "DECORATIVE LIGHTS"),
        ("floor lamp",              "DECORATIVE LIGHTS"),
        ("table lamp",              "DECORATIVE LIGHTS"),
        ("wall sconce",             "DECORATIVE LIGHTS"),
        ("wall light",              "DECORATIVE LIGHTS"),
        ("outdoor lighting",        "DECORATIVE LIGHTS"),
        ("decorative hanging",      "DECORATIVE LIGHTS"),
        ("decorative light",        "DECORATIVE LIGHTS"),
        # AMBIENT LIGHTS — functional general-purpose lighting
        ("concealed led",           "AMBIENT LIGHTS"),
        ("surface led panel",       "AMBIENT LIGHTS"),
        ("led panel light",         "AMBIENT LIGHTS"),
        ("led downlight",           "AMBIENT LIGHTS"),
        ("led track light",         "AMBIENT LIGHTS"),
        ("led strip light",         "AMBIENT LIGHTS"),
        ("led flexible",            "AMBIENT LIGHTS"),
        ("led linear light",        "AMBIENT LIGHTS"),
        ("led linear profile",      "AMBIENT LIGHTS"),  # recessed LED linear profile
        ("linear profile light",    "AMBIENT LIGHTS"),
        ("led profile light",       "AMBIENT LIGHTS"),
        ("profile light",           "AMBIENT LIGHTS"),
        ("cob light",               "AMBIENT LIGHTS"),
        ("hanging spot",            "AMBIENT LIGHTS"),
        ("spot light",              "AMBIENT LIGHTS"),
        ("led spot light",          "AMBIENT LIGHTS"),
        ("recessed led",            "AMBIENT LIGHTS"),
        ("led surface light",       "AMBIENT LIGHTS"),
        ("led recessed",            "AMBIENT LIGHTS"),
        ("recess mounted",          "AMBIENT LIGHTS"),
        ("surface mounted",         "AMBIENT LIGHTS"),
        ("panel light",             "AMBIENT LIGHTS"),
        ("down light",              "AMBIENT LIGHTS"),
        ("cove light",              "AMBIENT LIGHTS"),  # cove = indirect ambient, not decorative
        # ARCHITECTURAL LIGHTS — custom/designer/magnetic profiles
        ("magnetic light",          "ARCHITECTURAL LIGHTS"),
        ("led magnetic",            "ARCHITECTURAL LIGHTS"),
        ("led cabinet",             "ARCHITECTURAL LIGHTS"),
        ("led staircase",           "ARCHITECTURAL LIGHTS"),
        ("custom made profile",     "ARCHITECTURAL LIGHTS"),
        ("customised linear",       "ARCHITECTURAL LIGHTS"),
    ],

    "ACOUSTIC": [
        ("moulded acoustic ceiling","CEILING"),
        ("acoustic baffle",         "CEILING"),
        ("acoustic tile",           "CEILING"),
        ("acoustic grid",           "CEILING"),
        ("acoustic ceiling",        "CEILING"),
        ("acoustic linear light",   "LIGHTS"),
        ("acoustic pendant",        "LIGHTS"),
        ("privacy screen",          "PARTITION"),
        ("workstation partition",   "PARTITION"),
        ("customized plain panel",  "WALL SOLUTIONS"),
        ("grooved panel",           "WALL SOLUTIONS"),
        ("moulded acoustic",        "WALL SOLUTIONS"),
        ("movable wall",            "WALL SOLUTIONS"),
        ("plain panel",             "WALL SOLUTIONS"),
        ("print panel",             "WALL SOLUTIONS"),
        ("fluted panel",            "WALL SOLUTIONS"),
    ],
}


# ══════════════════════════════════════════════════════════════════════════════
# Row-level hint normaliser  (e.g. "Furniture Type" column, "Wall/Ceiling/Slab" column)
# ══════════════════════════════════════════════════════════════════════════════
_ROW_HINT_NORMALISE: list[tuple[str, str]] = [
    # FURNITURE
    ("carpent",             "CUSTOM FURNITURE"),
    ("custom",              "CUSTOM FURNITURE"),
    ("modular",             "MODULAR FURNITURE"),
    ("loose",               "LOOSE FURNITURE"),
    ("chair",               "CHAIRS"),
    # FLOORING
    ("tile flooring",       "TILE FLOORING"),
    ("textile",             "TEXTILE FLOORING"),
    ("resilient",           "RESILIENT FLOORING"),
    ("raised floor",        "RAISED FLOORING"),
    # CEILING
    ("gypsum ceiling",      "GYPSUM CEILING"),
    ("metal ceiling",       "METAL CEILING"),
    ("decorative ceiling",  "DECORATIVE CEILING"),
    ("wood ceiling",        "WOOD CEILING"),
    # PARTITIONS AND DOORS
    ("glass partition",     "GLASS PARTITION"),
    ("partition",           "PARTITION"),
    ("doors",               "DOORS"),
    ("door",                "DOORS"),
    ("windows",             "WINDOWS"),
    # DECORATIVES
    ("skirting",            "SKIRTING AND PROFILE"),
    ("blind",               "BLIND"),
    ("graphics",            "GRAPHICS"),
    ("panelling",           "PANELLING"),
    ("signage",             "SIGNAGES AND BRANDING"),
    ("branding",            "SIGNAGES AND BRANDING"),
    # SURFACE AND FINISHES
    ("wooden",              "WOODEN"),
    ("laminate",            "WOODEN"),
    ("veneer",              "WOODEN"),
    # CIVIL
    ("sanitary",            "SANITARY FIXTURES"),
    ("fabrication",         "FABRICATION"),
    # PAINT — "Wall/Ceiling/Slab" column values
    ("ceiling paint",       "CEILING PAINT"),
    ("wall paint",          "WALL PAINT"),
    ("duco",                "DUCO PAINT"),
    ("duct",                "DUCT PAINT"),
    ("texture",             "TEXTURE PAINT"),
    ("ceiling",             "CEILING PAINT"),   # e.g. column C = "Ceiling"
    ("wall",                "WALL PAINT"),       # e.g. column C = "Wall"
    ("slab",                "WALL PAINT"),       # slab/floor paint treated as wall
]


# ── Result dataclass ──────────────────────────────────────────────────────────

@dataclass
class ClassificationResult:
    """Output of Engine 1 for a single GFC line item."""
    category:    Optional[str] = None
    subcategory: Optional[str] = None
    item:        Optional[str] = None
    confidence:  int = 0        # 0–100
    method:      str = ""       # winning signal type
    signals:     list = field(default_factory=list)  # audit trail


# ── Internal helpers ──────────────────────────────────────────────────────────

def _composite(description: str, areas: str, product_name: str, sub_hint: str) -> str:
    """Full search text incl. area — used for category context only."""
    return _norm(" ".join(p for p in (description, areas, product_name, sub_hint) if p))


_LOC_WORDS = (
    "area", "areas", "room", "lobby", "near", "behind", "floor", "cabin",
    "zone", "passage", "entrance", "washroom", "secretary", "mandir",
    "reception area", "medical", "toilet",
)


def _strip_locations(text: str) -> str:
    """Remove parenthetical groups that are purely location context."""
    def _repl(m):
        inner = m.group(1).lower()
        return " " if any(w in inner for w in _LOC_WORDS) else m.group(0)
    return re.sub(r"\(([^)]*)\)", _repl, text or "")


def _kw_text(description: str, product_name: str) -> str:
    """Text for Signal 2 keyword matching: description + product name ONLY.
    Sub_hint is deliberately excluded to prevent the sheet hint value
    (e.g. 'METAL CEILING', 'PARTITION') from contaminating keyword matches."""
    raw = " ".join(p for p in (description, product_name) if p)
    return _norm(_strip_locations(raw))


def _subcat_text(description: str, product_name: str, sub_hint: str) -> str:
    """Full sub-category detection text incl. sub_hint — used for override
    detection (STEP 0) and master item matching (Signal 3)."""
    raw = " ".join(p for p in (description, product_name, sub_hint) if p)
    return _norm(_strip_locations(raw))


def _valid_subs_for(category: str, master_for_category: list) -> set:
    live = {(m["Sub-category"] or "").upper() for m in master_for_category if m.get("Sub-category")}
    if live:
        return live
    return MASTER_SUBCATS.get((category or "").upper(), set())


def _signal_keyword(text: str, category: str, valid_subs: set) -> tuple[Optional[str], int]:
    """Signal 2 — phrase keyword match, category-scoped, description-only text."""
    phrases = _SUBCAT_PHRASES_BY_CAT.get((category or "").upper(), [])
    best_sub, best_len = None, 0
    for phrase, sub in phrases:
        if phrase in text and sub.upper() in valid_subs and len(phrase) > best_len:
            best_sub, best_len = sub, len(phrase)
    if best_sub is None:
        return None, 0
    conf = min(90, 63 + best_len)
    return best_sub, conf


def _signal_master_items(
    text: str,
    candidates: list,
    sub_hint: str,
) -> tuple[Optional[str], Optional[str], int]:
    """Signal 3 — master item-name similarity (uses full text incl. sub_hint)."""
    desc_tokens = {_light_stem(t) for t in text.split() if len(t) >= 3}
    best, best_pts = None, 0

    for m in candidates:
        pts  = 0
        msub = (m["Sub-category"] or "").upper()
        mitm = (m["Item"] or "").lower()

        mitm_stemmed = " ".join(_light_stem(t) for t in mitm.split() if len(t) >= 3)
        text_stemmed = " ".join(_light_stem(t) for t in text.split() if len(t) >= 3)
        if mitm_stemmed and mitm_stemmed in text_stemmed:
            pts += 20

        item_tokens = {_light_stem(t) for t in mitm.split() if len(t) >= 3}
        pts += len(item_tokens & desc_tokens) * 2

        if sub_hint and (msub == sub_hint.upper()
                         or sub_hint.upper() in msub
                         or msub in sub_hint.upper()):
            pts += 5

        if pts > best_pts:
            best_pts, best = pts, m

    if best is None or best_pts < 4:
        return None, None, 0
    conf = min(82, 34 + best_pts * 2)
    return best["Sub-category"], best["Item"], conf


def _signal_row_hint(raw_hint: str, valid_subs: set) -> tuple[Optional[str], int]:
    """Signal 1 — explicit sub-cat hint from a GFC column (e.g. 'Furniture Type',
    'Wall/Ceiling/Slab').  Only fires if the resulting sub-category is valid for
    the current category."""
    if not raw_hint:
        return None, 0
    r = _norm(raw_hint)
    for fragment, sub in _ROW_HINT_NORMALISE:
        if fragment in r and sub.upper() in valid_subs:
            return sub, 92
    upper = raw_hint.strip().upper()
    if upper in valid_subs:
        return upper, 92
    return None, 0


def _apply_category_override(composite: str) -> Optional[tuple[str, str, str]]:
    """STEP 0 — return (trigger, new_cat, new_sub) if an override fires, else None."""
    for trigger, also_any, new_cat, new_sub in _CATEGORY_OVERRIDES:
        if trigger not in composite:
            continue
        if also_any is None or any(extra in composite for extra in also_any):
            return (trigger, new_cat, new_sub)
    return None


# ── Public API ────────────────────────────────────────────────────────────────

def classify_line_item(
    category:            Optional[str],
    sub_hint:            Optional[str],
    description:         str,
    areas:               str,
    product_name:        str,
    row_subcat_hint:     str,
    master_for_category: list,
    master_by_cat:       Optional[dict] = None,
) -> ClassificationResult:
    """
    Engine 1 — classify a single GFC line item against the Category Master.

    KEY CHANGE: Signal 2 (keyword) now uses description + product_name ONLY
    (no sub_hint) to prevent the sheet hint from contaminating keyword matches.
    """
    result = ClassificationResult(category=category)

    if not category:
        result.method  = "no_category"
        result.signals = ["L1 did not assign a category — cannot classify"]
        return result

    # kw_text: description + product_name only — for Signal 2 keyword matching.
    # full_text: includes sub_hint — for category override (STEP 0) and Signal 3.
    kw_text   = _kw_text(description, product_name)
    full_text  = _subcat_text(description, product_name, sub_hint or "")
    signals   = []
    override_sub = None

    # ── STEP 0: Category override (checks full_text incl. sub_hint) ───────
    ov = _apply_category_override(full_text)
    if ov and master_by_cat is not None:
        trigger, new_cat, new_sub = ov
        if new_cat.upper() != (category or "").upper():
            signals.append(f"CATEGORY_OVERRIDE:'{trigger}'→{category}→{new_cat}")
            category            = new_cat
            master_for_category = master_by_cat.get(new_cat.upper(), [])
        override_sub = new_sub
        result.category = category

    # ── STEP 1: valid sub-categories for this category ────────────────────
    valid_subs = _valid_subs_for(category, master_for_category)

    # ── STEP 2: four category-scoped signals ──────────────────────────────
    s1_sub, s1_conf = _signal_row_hint(row_subcat_hint, valid_subs)
    if s1_sub:
        signals.append(f"S1_row_hint→'{s1_sub}'({s1_conf})")

    # Signal 2 uses kw_text (no sub_hint) to avoid circular contamination
    s2_sub, s2_conf = _signal_keyword(kw_text, category, valid_subs)
    if s2_sub:
        signals.append(f"S2_keyword→'{s2_sub}'({s2_conf})")

    s3_sub, s3_item, s3_conf = _signal_master_items(
        full_text, master_for_category, sub_hint or "")
    if s3_sub and s3_sub.upper() not in valid_subs:
        s3_sub, s3_conf = None, 0
    if s3_sub:
        signals.append(f"S3_item→'{s3_item}':'{s3_sub}'({s3_conf})")

    s4_sub, s4_conf = None, 0
    if sub_hint and sub_hint.upper() in valid_subs:
        s4_sub, s4_conf = sub_hint.upper(), 65
        signals.append(f"S4_sheet_hint→'{sub_hint}'({s4_conf})")

    ov_conf = 85 if override_sub and override_sub.upper() in valid_subs else 0
    if override_sub and ov_conf:
        signals.append(f"S0_override→'{override_sub}'({ov_conf})")

    # ── Combine: rank by confidence ───────────────────────────────────────
    candidates = [
        (s1_sub,      s1_conf, "row_hint"),
        (s2_sub,      s2_conf, "keyword"),
        (s3_sub,      s3_conf, "item_match"),
        (s4_sub,      s4_conf, "sheet_hint"),
        (override_sub, ov_conf, "category_override"),
    ]
    candidates = [(s, c, m) for s, c, m in candidates if s and c > 0]

    if not candidates:
        # ── STEP 4a: ACOUSTIC — use area column for ceiling vs wall distinction ──
        if (category or "").upper() == "ACOUSTIC" and areas:
            area_n = _norm(areas)
            if "ceiling" in area_n:
                result.category   = category
                result.subcategory = "CEILING"
                result.confidence = 60
                result.method     = "area_hint"
                result.signals    = signals + [f"ACOUSTIC area_hint→CEILING (area='{areas}')"]
                return result
            elif any(w in area_n for w in ("wall", "column", "partition", "pillar")):
                result.category   = category
                result.subcategory = "WALL SOLUTIONS"
                result.confidence = 55
                result.method     = "area_hint"
                result.signals    = signals + [f"ACOUSTIC area_hint→WALL SOLUTIONS (area='{areas}')"]
                return result

        result.category   = category
        result.method     = "category_only"
        result.confidence = 28
        result.signals    = signals + [f"no valid sub-category for {category} — left blank"]
        return result

    candidates.sort(key=lambda x: -x[1])
    primary_sub, primary_conf, primary_method = candidates[0]

    # Agreement / conflict adjustment
    if len(candidates) >= 2:
        second_sub, second_conf, _ = candidates[1]
        if second_sub.upper() == primary_sub.upper():
            primary_conf = min(100, primary_conf + 10)
            signals.append("agreement_boost:+10")
        elif second_conf >= 50:
            primary_conf = max(0, primary_conf - 8)
            signals.append(f"conflict_penalty:-8(vs '{second_sub}')")

    # ── STEP 3: hard guard ────────────────────────────────────────────────
    if primary_sub.upper() not in valid_subs:
        result.category   = category
        result.subcategory = None
        result.method     = "category_only(guarded)"
        result.confidence = 28
        result.signals    = signals + [
            f"GUARD: '{primary_sub}' not valid for {category} — dropped"]
        return result

    result.category    = category
    result.subcategory = primary_sub
    result.confidence  = primary_conf
    result.method      = primary_method
    result.item        = s3_item
    result.signals     = signals
    return result

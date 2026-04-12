"""
Address Corrector
-----------------
Reads a CSV or Excel file containing address fields, applies standardisation
corrections, and writes an output file with both the original and corrected
columns side-by-side.

Usage:
    python address_corrector.py <input_file> <output_file>

Supported formats:  .csv  .xlsx  .xls
"""

import os
import re
import sys
from difflib import get_close_matches

import pandas as pd
import pycountry
import geonamescache as _gnc
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────────────────────
# COLUMN NAME DETECTION
# ──────────────────────────────────────────────────────────────────────────────

# ──────────────────────────────────────────────────────────────────────────────
# COLUMN DETECTION  — keyword-based, not a fixed list
# Works with any column name that contains relevant keywords or digits
# ──────────────────────────────────────────────────────────────────────────────

# Keywords that must appear somewhere in the column name (after normalisation)
# Listed in priority order for each field; first hit wins.
_FIELD_KEYWORDS = {
    # Address lines — distinguished by trailing digit (1 / 2 / 3)
    "address_line_1": [
        r"\baddr(?:ess)?\b.*\b1\b",
        r"\bstreet\b.*\b1\b",
        r"\bline\b.*\b1\b",
        r"\baddress1\b",
        r"\baddr1\b",
        r"\badd1\b",
        r"\bhouse\s*(?:no\.?|number|num)?\b",
        # Catch bare "address" / "street address" only when no digit variant exists (resolved later)
        r"\b(?:mailing|billing|delivery|residential|primary|house|full)?\s*addr(?:ess)?\b",
        r"\bstreet\s+addr(?:ess)?\b",
        r"\bstreet\b",
    ],
    "address_line_2": [
        r"\baddr(?:ess)?\b.*\b2\b",
        r"\bstreet\b.*\b2\b",
        r"\bline\b.*\b2\b",
        r"\baddress2\b",
        r"\baddr2\b",
        r"\badd2\b",
        r"\b(?:suite|apartment|unit|apt|floor|building)\b",
        r"\bsubpremise\b",
    ],
    "address_line_3": [
        r"\baddr(?:ess)?\b.*\b3\b",
        r"\bstreet\b.*\b3\b",
        r"\bline\b.*\b3\b",
        r"\baddress3\b",
        r"\baddr3\b",
        r"\badd3\b",
        r"\badditional\s*addr(?:ess)?\b",
        r"\bextra\s*addr(?:ess)?\b",
        r"\bsupplementary\s*addr(?:ess)?\b",
        r"\baddr(?:ess)?\s*(?:cont(?:inued)?|cont\.?)\b",
    ],
    "city": [
        r"\bcity\b",
        r"\btown\b",
        r"\bsuburb\b",
        r"\blocality\b",
        r"\bmunicipality\b",
        r"\btownship\b",
        r"\bvillage\b",
    ],
    "state": [
        r"\bstate\b",
        r"\bprovince\b",
        r"\bregion\b",
        r"\bcounty\b",
        r"\bterritory\b",
        r"\bprefecture\b",
    ],
    "country": [
        r"\bcountry\b",
        r"\bnation\b",
        r"\bnationality\b",
    ],
    "postal_code": [
        r"\bpostal\b",
        r"\bpostcode\b",
        r"\bzip\b",
        r"\bpin\s*code\b",
        r"\bpin\b",
        r"\bpost\s*code\b",
    ],
}


def _normalise_col_name(name):
    """Lower-case, collapse spaces/underscores/hyphens."""
    return re.sub(r"[\s_\-/]+", " ", name.lower().strip())


def _detect_columns(df_cols):
    """
    Map each address field to the best matching column name (or None).

    Strategy:
      1. For each column, test patterns in priority order.
      2. When multiple columns match the same field, pick the best one
         (highest priority pattern index = lower number wins).
      3. Each column can only be assigned to one field.
    """
    normalised = {c: _normalise_col_name(c) for c in df_cols}

    # candidates[field] = list of (priority, original_col_name)
    candidates = {f: [] for f in _FIELD_KEYWORDS}

    for orig_col, norm in normalised.items():
        for field, patterns in _FIELD_KEYWORDS.items():
            for priority, pat in enumerate(patterns):
                if re.search(pat, norm):
                    candidates[field].append((priority, orig_col))
                    break  # only record lowest-priority match per field per column

    # Resolve: assign each column to at most one field, prefer highest priority
    assigned_cols = set()
    col_map = {}

    for field in _FIELD_KEYWORDS:
        best = sorted(candidates[field], key=lambda x: x[0])
        matched = None
        for _, col in best:
            if col not in assigned_cols:
                matched = col
                assigned_cols.add(col)
                break
        col_map[field] = matched

    return col_map


# ──────────────────────────────────────────────────────────────────────────────
# LOOKUP TABLES
# ──────────────────────────────────────────────────────────────────────────────

# US states (full name → 2-letter abbreviation)
US_STATES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT", "delaware": "DE",
    "florida": "FL", "georgia": "GA", "hawaii": "HI", "idaho": "ID",
    "illinois": "IL", "indiana": "IN", "iowa": "IA", "kansas": "KS",
    "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD",
    "massachusetts": "MA", "michigan": "MI", "minnesota": "MN", "mississippi": "MS",
    "missouri": "MO", "montana": "MT", "nebraska": "NE", "nevada": "NV",
    "new hampshire": "NH", "new jersey": "NJ", "new mexico": "NM", "new york": "NY",
    "north carolina": "NC", "north dakota": "ND", "ohio": "OH", "oklahoma": "OK",
    "oregon": "OR", "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC",
    "south dakota": "SD", "tennessee": "TN", "texas": "TX", "utah": "UT",
    "vermont": "VT", "virginia": "VA", "washington": "WA", "west virginia": "WV",
    "wisconsin": "WI", "wyoming": "WY", "district of columbia": "DC",
    "puerto rico": "PR", "guam": "GU", "virgin islands": "VI",
    "american samoa": "AS", "northern mariana islands": "MP",
}

# Valid 2-letter US state codes (for uppercase normalisation)
US_STATE_CODES = set(US_STATES.values())

# Canadian provinces/territories (full name → 2-letter abbreviation)
CA_PROVINCES = {
    "alberta": "AB", "british columbia": "BC", "manitoba": "MB",
    "new brunswick": "NB", "newfoundland and labrador": "NL",
    "newfoundland": "NL", "labrador": "NL",
    "northwest territories": "NT", "nova scotia": "NS", "nunavut": "NU",
    "ontario": "ON", "prince edward island": "PE", "quebec": "QC",
    "québec": "QC", "saskatchewan": "SK", "yukon": "YT",
}

CA_PROVINCE_CODES = set(CA_PROVINCES.values())

# Australian states/territories
AU_STATES = {
    "new south wales": "NSW", "victoria": "VIC", "queensland": "QLD",
    "south australia": "SA", "western australia": "WA", "tasmania": "TAS",
    "northern territory": "NT", "australian capital territory": "ACT",
}

AU_STATE_CODES = set(AU_STATES.values())

# All known 2-letter region codes (do not title-case these)
ALL_REGION_CODES = US_STATE_CODES | CA_PROVINCE_CODES | AU_STATE_CODES | {
    # Indian states abbreviations kept as entered, just uppercase
    "AP", "AR", "AS", "BR", "CG", "GA", "GJ", "HR", "HP", "JH", "KA",
    "KL", "MP", "MH", "MN", "ML", "MZ", "NL", "OD", "PB", "RJ", "SK",
    "TN", "TS", "TR", "UP", "UK", "WB", "AN", "CH", "DN", "DD", "DL",
    "JK", "LA", "LD", "PY",
}

# Country aliases → standardised name
# Manual alias map — covers abbreviations, native-language names, and
# common alternate spellings that pycountry cannot resolve on its own.
# pycountry handles the remaining 200+ countries automatically.
COUNTRY_MAP = {
    # ── United States ──────────────────────────────────────────────────────
    "usa": "United States", "u.s.a": "United States", "u.s.a.": "United States",
    "us": "United States", "u.s.": "United States", "u.s": "United States",
    "united states of america": "United States", "america": "United States",
    "the united states": "United States", "the us": "United States",
    # ── United Kingdom ─────────────────────────────────────────────────────
    "uk": "United Kingdom", "u.k.": "United Kingdom", "u.k": "United Kingdom",
    "great britain": "United Kingdom", "gb": "United Kingdom", "g.b.": "United Kingdom",
    "england": "United Kingdom", "britain": "United Kingdom",
    "scotland": "United Kingdom", "wales": "United Kingdom",
    "northern ireland": "United Kingdom",
    # ── UAE ────────────────────────────────────────────────────────────────
    "uae": "United Arab Emirates", "u.a.e.": "United Arab Emirates",
    "u.a.e": "United Arab Emirates", "emirates": "United Arab Emirates",
    # ── Germany ────────────────────────────────────────────────────────────
    "deutschland": "Germany", "allemagne": "Germany", "almanya": "Germany",
    # ── France ─────────────────────────────────────────────────────────────
    "frankreich": "France", "frankrig": "France",
    # ── Italy ──────────────────────────────────────────────────────────────
    "italia": "Italy", "italie": "Italy", "italien": "Italy",
    # ── Spain ──────────────────────────────────────────────────────────────
    "espana": "Spain", "españa": "Spain", "espagne": "Spain", "spanien": "Spain",
    # ── Netherlands ────────────────────────────────────────────────────────
    "the netherlands": "Netherlands", "holland": "Netherlands",
    "nederland": "Netherlands", "niederlande": "Netherlands",
    # ── Belgium ────────────────────────────────────────────────────────────
    "belgique": "Belgium", "belgien": "Belgium", "belgië": "Belgium",
    # ── Switzerland ────────────────────────────────────────────────────────
    "schweiz": "Switzerland", "suisse": "Switzerland", "svizzera": "Switzerland",
    "confederazione svizzera": "Switzerland",
    # ── Sweden ─────────────────────────────────────────────────────────────
    "sverige": "Sweden", "schweden": "Sweden",
    # ── Norway ─────────────────────────────────────────────────────────────
    "norge": "Norway", "norwegen": "Norway",
    # ── Denmark ────────────────────────────────────────────────────────────
    "danmark": "Denmark", "dänemark": "Denmark",
    # ── Finland ────────────────────────────────────────────────────────────
    "suomi": "Finland", "finnland": "Finland",
    # ── Austria ────────────────────────────────────────────────────────────
    "österreich": "Austria", "osterreich": "Austria", "autriche": "Austria",
    # ── Poland ─────────────────────────────────────────────────────────────
    "polska": "Poland", "pologne": "Poland", "polen": "Poland",
    # ── Czech Republic ─────────────────────────────────────────────────────
    "czechia": "Czech Republic", "ceska republika": "Czech Republic",
    "česká republika": "Czech Republic", "tschechien": "Czech Republic",
    # ── Hungary ────────────────────────────────────────────────────────────
    "magyarország": "Hungary", "magyarorszag": "Hungary", "hongrie": "Hungary",
    # ── Greece ─────────────────────────────────────────────────────────────
    "hellas": "Greece", "ellada": "Greece", "griechenland": "Greece",
    # ── Russia ─────────────────────────────────────────────────────────────
    "russian federation": "Russia", "rossiya": "Russia", "russland": "Russia",
    # ── Turkey ─────────────────────────────────────────────────────────────
    "türkiye": "Turkey", "turkiye": "Turkey", "türkei": "Turkey",
    # ── Saudi Arabia ───────────────────────────────────────────────────────
    "ksa": "Saudi Arabia", "kingdom of saudi arabia": "Saudi Arabia",
    "al-mamlaka al-arabiyya as-saudiyya": "Saudi Arabia",
    # ── China ──────────────────────────────────────────────────────────────
    "prc": "China", "people's republic of china": "China",
    "peoples republic of china": "China", "zhongguo": "China",
    # ── South Korea ────────────────────────────────────────────────────────
    "korea": "South Korea", "republic of korea": "South Korea",
    "south korea": "South Korea", "hanguk": "South Korea",
    # ── North Korea ────────────────────────────────────────────────────────
    "dprk": "North Korea", "north korea": "North Korea",
    # ── Taiwan ─────────────────────────────────────────────────────────────
    "taiwan": "Taiwan", "roc": "Taiwan", "republic of china": "Taiwan",
    # ── Hong Kong ──────────────────────────────────────────────────────────
    "hong kong sar": "Hong Kong", "hksar": "Hong Kong",
    # ── Brazil ─────────────────────────────────────────────────────────────
    "brasil": "Brazil", "brésil": "Brazil",
    # ── Mexico ─────────────────────────────────────────────────────────────
    "méxico": "Mexico", "mejico": "Mexico",
    # ── Australia ──────────────────────────────────────────────────────────
    "oz": "Australia", "aussie": "Australia",
    # ── New Zealand ────────────────────────────────────────────────────────
    "aotearoa": "New Zealand",
    # ── South Africa ───────────────────────────────────────────────────────
    "rsa": "South Africa", "suid-afrika": "South Africa",
    # ── India ──────────────────────────────────────────────────────────────
    "bharat": "India", "hindustan": "India",
    # ── Pakistan ───────────────────────────────────────────────────────────
    "pak": "Pakistan",
    # ── Iran ───────────────────────────────────────────────────────────────
    "persia": "Iran", "islamic republic of iran": "Iran",
    # ── Iraq ───────────────────────────────────────────────────────────────
    "al-iraq": "Iraq",
    # ── Egypt ──────────────────────────────────────────────────────────────
    "misr": "Egypt", "arab republic of egypt": "Egypt",
    # ── Morocco ────────────────────────────────────────────────────────────
    "maroc": "Morocco", "marruecos": "Morocco", "marokko": "Morocco",
    # ── Algeria ────────────────────────────────────────────────────────────
    "algérie": "Algeria", "algerie": "Algeria",
    # ── Tunisia ────────────────────────────────────────────────────────────
    "tunisie": "Tunisia",
    # ── Ethiopia ───────────────────────────────────────────────────────────
    "abyssinia": "Ethiopia",
    # ── Ivory Coast ────────────────────────────────────────────────────────
    "ivory coast": "Côte d'Ivoire", "cote d'ivoire": "Côte d'Ivoire",
    "cote divoire": "Côte d'Ivoire",
    # ── Democratic Republic of Congo ───────────────────────────────────────
    "drc": "Congo, The Democratic Republic of the",
    "dr congo": "Congo, The Democratic Republic of the",
    "democratic republic of congo": "Congo, The Democratic Republic of the",
    "zaire": "Congo, The Democratic Republic of the",
    # ── Republic of Congo ──────────────────────────────────────────────────
    "congo": "Congo",
    # ── Vietnam ────────────────────────────────────────────────────────────
    "viet nam": "Vietnam",
    # ── Myanmar ────────────────────────────────────────────────────────────
    "burma": "Myanmar",
    # ── Sri Lanka ──────────────────────────────────────────────────────────
    "ceylon": "Sri Lanka",
    # ── Cambodia ───────────────────────────────────────────────────────────
    "kampuchea": "Cambodia",
    # ── Bosnia ─────────────────────────────────────────────────────────────
    "bosnia": "Bosnia and Herzegovina", "bih": "Bosnia and Herzegovina",
    # ── Macedonia ──────────────────────────────────────────────────────────
    "north macedonia": "North Macedonia", "macedonia": "North Macedonia",
    # ── Kosovo ─────────────────────────────────────────────────────────────
    "kosovo": "Kosovo",
    # ── Palestine ──────────────────────────────────────────────────────────
    "palestine": "Palestine, State of",
    # ── Vatican ────────────────────────────────────────────────────────────
    "vatican": "Holy See", "vatican city": "Holy See",
    # ── Macau ──────────────────────────────────────────────────────────────
    "macau": "Macao", "macao sar": "Macao",
}

# ──────────────────────────────────────────────────────────────────────────────
# STREET TYPE ABBREVIATIONS  (USPS standard — always uppercase)
# ──────────────────────────────────────────────────────────────────────────────

STREET_TYPES = {
    "alley": "ALY", "aly": "ALY",
    "avenue": "AVE", "ave": "AVE", "av": "AVE",
    "boulevard": "BLVD", "blvd": "BLVD",
    "bypass": "BYP", "byp": "BYP",
    "causeway": "CSWY", "cswy": "CSWY",
    "circle": "CIR", "cir": "CIR",
    "close": "CL", "cl": "CL",
    "court": "CT", "ct": "CT",
    "cove": "CV", "cv": "CV",
    "crescent": "CRES", "cres": "CRES",
    "crossing": "XING", "xing": "XING",
    "drive": "DR", "dr": "DR",
    "estate": "EST", "est": "EST",
    "expressway": "EXPY", "expy": "EXPY",
    "extension": "EXT", "ext": "EXT",
    "freeway": "FWY", "fwy": "FWY",
    "grove": "GRV", "grv": "GRV",
    "heights": "HTS", "hts": "HTS",
    "highway": "HWY", "hwy": "HWY",
    "hill": "HL", "hl": "HL",
    "hollow": "HOLW", "holw": "HOLW",
    "junction": "JCT", "jct": "JCT",
    "lake": "LK", "lk": "LK",
    "lane": "LN", "ln": "LN",
    "loop": "LOOP",
    "manor": "MNR", "mnr": "MNR",
    "meadow": "MDW", "mdw": "MDW",
    "mount": "MT", "mt": "MT",
    "motorway": "MTWY", "mtwy": "MTWY",
    "park": "PARK",
    "parkway": "PKWY", "pkwy": "PKWY", "pky": "PKWY",
    "place": "PL", "pl": "PL",
    "plaza": "PLZ", "plz": "PLZ",
    "point": "PT", "pt": "PT",
    "ridge": "RDG", "rdg": "RDG",
    "road": "RD", "rd": "RD",
    "row": "ROW",
    "run": "RUN",
    "square": "SQ", "sq": "SQ",
    "street": "ST", "str": "ST",
    "terrace": "TER", "ter": "TER", "tce": "TER",
    "trail": "TRL", "trl": "TRL",
    "turnpike": "TPKE", "tpke": "TPKE",
    "valley": "VLY", "vly": "VLY",
    "vista": "VIS", "vis": "VIS",
    "walk": "WALK",
    "way": "WAY",
}

# Directional prefixes/suffixes
DIRECTIONALS = {
    "north": "N", "south": "S", "east": "E", "west": "W",
    "northeast": "NE", "northwest": "NW", "southeast": "SE", "southwest": "SW",
    "n.": "N", "s.": "S", "e.": "E", "w.": "W",
    "n": "N", "s": "S", "e": "E", "w": "W",
    "ne": "NE", "nw": "NW", "se": "SE", "sw": "SW",
    "n.e.": "NE", "n.w.": "NW", "s.e.": "SE", "s.w.": "SW",
}

# Secondary unit designators
UNIT_TYPES = {
    "apartment": "APT", "apt": "APT",
    "suite": "STE", "ste": "STE", "suit": "STE",
    "unit": "UNIT", "unt": "UNIT",
    "room": "RM", "rm": "RM",
    "floor": "FL", "fl": "FL", "flr": "FL",
    "building": "BLDG", "bldg": "BLDG", "bld": "BLDG",
    "department": "DEPT", "dept": "DEPT",
    "po box": "PO BOX", "p.o. box": "PO BOX", "p.o box": "PO BOX",
    "po. box": "PO BOX", "pobox": "PO BOX", "post office box": "PO BOX",
    "p o box": "PO BOX", "p.o. box": "PO BOX",
}

# ──────────────────────────────────────────────────────────────────────────────
# CORRECTION FUNCTIONS
# ──────────────────────────────────────────────────────────────────────────────

def _clean(val):
    """Strip and collapse internal whitespace."""
    if pd.isna(val) or str(val).strip() in ("", "nan", "NaN", "None", "none"):
        return ""
    return re.sub(r"\s+", " ", str(val).strip())


def _normalise_po_box(val):
    """Standardise PO Box variations to 'PO BOX <number>'."""
    pattern = re.compile(
        r"\b(p\.?\s*o\.?\s*box|post\s+office\s+box|pobox)\b\s*(\d*)",
        re.IGNORECASE,
    )
    return pattern.sub(lambda m: f"PO BOX {m.group(2)}".strip(), val)


def correct_address_line(val):
    val = _clean(val)
    if not val:
        return ""

    val = _normalise_po_box(val)

    words = val.split()
    result = []
    i = 0
    while i < len(words):
        w = words[i]
        wl = w.lower().rstrip(".,")

        # Try two-word combinations first (e.g. "po box", "north east")
        two_word = None
        if i + 1 < len(words):
            two_word_key = (wl + " " + words[i + 1].lower().rstrip(".,"))
            if two_word_key in UNIT_TYPES:
                result.append(UNIT_TYPES[two_word_key])
                i += 2
                continue
            if two_word_key in DIRECTIONALS:
                result.append(DIRECTIONALS[two_word_key])
                i += 2
                continue

        # Single word lookups
        if wl in STREET_TYPES:
            result.append(STREET_TYPES[wl])
        elif wl in UNIT_TYPES:
            result.append(UNIT_TYPES[wl])
        elif wl in DIRECTIONALS:
            result.append(DIRECTIONALS[wl])
        elif re.match(r"^\d+$", w):           # pure number
            result.append(w)
        elif re.match(r"^\d+(st|nd|rd|th)$", w, re.IGNORECASE):   # ordinal
            result.append(w.lower())
        elif re.match(r"^[a-zA-Z]\d+[a-zA-Z]?$", w):              # unit codes e.g. A2, B12C
            result.append(w.upper())
        elif re.match(r"^\d+[-–]\d+$", w):    # range e.g. 12-14
            result.append(w)
        else:
            result.append(w.title())
        i += 1

    return " ".join(result).upper()


def correct_city(val):
    val = _clean(val)
    if not val:
        return ""

    lower = val.lower()

    # 1. Exact match in known cities → use canonical casing
    if lower in _ALL_CITIES:
        return _ALL_CITIES[lower].upper()

    # 2. Fuzzy match for misspellings — only for inputs ≥ 4 chars
    if len(val) >= 4:
        from difflib import SequenceMatcher
        candidates = get_close_matches(lower, _ALL_CITY_NAMES, n=5, cutoff=0.78)
        if candidates:
            import math
            def _score(c):
                sim = SequenceMatcher(None, lower, c).ratio()
                pop = _CITY_POPULATION.get(c, 0)
                pop_bonus = math.log10(pop + 1) * 0.045
                return sim + pop_bonus
            best = max(candidates, key=_score)
            if SequenceMatcher(None, lower, best).ratio() >= 0.78:
                return _ALL_CITIES[best].upper()

    # 3. Fallback — all caps
    return val.upper()


def correct_state(val, country_hint=""):
    val = _clean(val)
    if not val:
        return ""
    lower = val.lower()

    # 1. Exact match → abbreviation (already uppercase by definition)
    if lower in US_STATES:
        return US_STATES[lower]
    if lower in CA_PROVINCES:
        return CA_PROVINCES[lower]
    if lower in AU_STATES:
        return AU_STATES[lower]

    # 2. Short code (2–4 alpha chars) → uppercase
    if len(val) <= 4 and val.replace(" ", "").isalpha():
        return val.upper()

    # 3. Fuzzy match against all known state / province full names
    if len(val) > 4:
        hits = get_close_matches(lower, _ALL_STATE_NAMES, n=1, cutoff=0.78)
        if hits:
            return _STATE_FUZZY_INDEX[hits[0]]

    # 4. Fallback — all caps
    return val.upper()


def _pycountry_name(c):
    """Return the preferred display name for a pycountry country object."""
    return getattr(c, "common_name", None) or c.name


# Build a flat lookup: lowercase name variant → pycountry object
# Used for difflib fuzzy matching (covers misspellings)
_COUNTRY_NAME_INDEX: dict = {}
for _c in pycountry.countries:
    for _attr in ("name", "common_name", "official_name", "alpha_2", "alpha_3"):
        _v = getattr(_c, _attr, None)
        if _v:
            _COUNTRY_NAME_INDEX[_v.lower()] = _c
_ALL_COUNTRY_NAMES = list(_COUNTRY_NAME_INDEX.keys())

# ── State / province fuzzy index ──────────────────────────────────────────────
# Maps lowercase full-name → canonical abbreviation for every known region.
_STATE_FUZZY_INDEX: dict[str, str] = {}
_STATE_FUZZY_INDEX.update({k: v for k, v in US_STATES.items()})
_STATE_FUZZY_INDEX.update({k: v for k, v in CA_PROVINCES.items()})
_STATE_FUZZY_INDEX.update({k: v for k, v in AU_STATES.items()})
# Also index Indian states via pycountry subdivisions
for _sub in pycountry.subdivisions.get(country_code="IN") or []:
    _STATE_FUZZY_INDEX[_sub.name.lower()] = _sub.code.split("-")[-1]
_ALL_STATE_NAMES = list(_STATE_FUZZY_INDEX.keys())

# ── City fuzzy index ──────────────────────────────────────────────────────────
# Only cities with population > 50 000 to keep matching fast and accurate.
_gc = _gnc.GeonamesCache()
_ALL_CITIES: dict[str, str] = {}         # lowercase → canonical Title Case name
_CITY_POPULATION: dict[str, int] = {}    # lowercase → population (for tiebreaking)
for _city in _gc.get_cities().values():
    if _city["population"] > 50_000:
        _name = _city["name"]
        _pop  = _city["population"]
        _key  = _name.lower()
        # Keep the higher-population entry when names collide (e.g. multiple "Springfield")
        if _pop > _CITY_POPULATION.get(_key, 0):
            _ALL_CITIES[_key]      = _name
            _CITY_POPULATION[_key] = _pop
        # Also index name without " City" suffix  ("New York City" → "new york")
        _alt = _name.replace(" City", "").replace(" city", "").strip()
        if _alt and _alt.lower() != _key:
            if _pop > _CITY_POPULATION.get(_alt.lower(), 0):
                _ALL_CITIES[_alt.lower()]      = _name
                _CITY_POPULATION[_alt.lower()] = _pop

# Explicit common aliases not in geonamescache
_CITY_ALIASES = {
    "new york":    "New York City",
    "nyc":         "New York City",
    "la":          "Los Angeles",
    "sf":          "San Francisco",
    "dc":          "Washington",
    "philly":      "Philadelphia",
    "vegas":       "Las Vegas",
    "nola":        "New Orleans",
    "chi":         "Chicago",
}
_ALL_CITIES.update(_CITY_ALIASES)
_ALL_CITY_NAMES = list(_ALL_CITIES.keys())


def correct_country(val):
    val = _clean(val)
    if not val:
        return ""

    lookup  = val.lower().strip(" .")
    no_dots = lookup.replace(".", "")
    stripped = no_dots.replace(" ", "")

    # 1. Manual alias map — fastest; handles abbreviations & native names
    if lookup in COUNTRY_MAP:
        return COUNTRY_MAP[lookup].upper()
    if no_dots in COUNTRY_MAP:
        return COUNTRY_MAP[no_dots].upper()

    # 2. ISO alpha-2 code  (e.g. "DE", "JP", "au")
    if len(stripped) == 2 and stripped.isalpha():
        c = pycountry.countries.get(alpha_2=stripped.upper())
        if c:
            return _pycountry_name(c).upper()

    # 3. ISO alpha-3 code  (e.g. "DEU", "aus", "GBR")
    if len(stripped) == 3 and stripped.isalpha():
        c = pycountry.countries.get(alpha_3=stripped.upper())
        if c:
            return _pycountry_name(c).upper()

    # 4. Exact ISO name / common_name / official_name match
    for attr, query in [
        ("name",          val.title()),
        ("common_name",   val.title()),
        ("official_name", val.title()),
        ("name",          val.upper()),
    ]:
        c = pycountry.countries.get(**{attr: query})
        if c:
            return _pycountry_name(c).upper()

    # 5. pycountry token search
    try:
        results = pycountry.countries.search_fuzzy(val)
        if results:
            return _pycountry_name(results[0]).upper()
    except LookupError:
        pass

    # 6. difflib edit-distance fuzzy match over all 249 country names
    hits = get_close_matches(lookup, _ALL_COUNTRY_NAMES, n=1, cutoff=0.72)
    if hits:
        return _pycountry_name(_COUNTRY_NAME_INDEX[hits[0]]).upper()

    # 7. Fallback — all caps
    return val.upper()


def correct_postal_code(val, country_hint=""):
    val = _clean(val)
    if not val:
        return ""

    # Remove surrounding quotes/spaces
    val = val.strip("'\"")
    stripped = val.replace(" ", "")

    # Pure numeric postal code — remove spaces, keep leading zeros
    if stripped.isdigit():
        return stripped

    # US ZIP+4 format  e.g. "12345 - 6789" → "12345-6789"
    zip4 = re.match(r"^(\d{5})\s*[-–]\s*(\d{4})$", val)
    if zip4:
        return f"{zip4.group(1)}-{zip4.group(2)}"

    # Canadian postal code  A1A1A1 or A1A 1A1 → A1A 1A1
    ca_match = re.match(r"^([A-Za-z]\d[A-Za-z])\s*(\d[A-Za-z]\d)$", stripped)
    if ca_match:
        return f"{ca_match.group(1).upper()} {ca_match.group(2).upper()}"

    # UK postcode — ensure space before last 3 chars if missing
    uk_match = re.match(
        r"^([A-Za-z]{1,2}\d[A-Za-z\d]?)\s*(\d[A-Za-z]{2})$", stripped
    )
    if uk_match:
        return f"{uk_match.group(1).upper()} {uk_match.group(2).upper()}"

    # Alphanumeric — uppercase
    return val.upper()


# ──────────────────────────────────────────────────────────────────────────────
# CORRECTION DISPATCH
# ──────────────────────────────────────────────────────────────────────────────

CORRECTORS = {
    "address_line_1": correct_address_line,
    "address_line_2": correct_address_line,
    "address_line_3": correct_address_line,
    "city":           correct_city,
    "state":          correct_state,
    "country":        correct_country,
    "postal_code":    correct_postal_code,
}

DISPLAY_LABELS = {
    "address_line_1": "Address Line 1",
    "address_line_2": "Address Line 2",
    "address_line_3": "Address Line 3",
    "city":           "City",
    "state":          "State",
    "country":        "Country",
    "postal_code":    "Postal Code",
}

# ──────────────────────────────────────────────────────────────────────────────
# EXCEL OUTPUT FORMATTING
# ──────────────────────────────────────────────────────────────────────────────

_HDR_FILL_ORIG = PatternFill("solid", fgColor="1F4E79")   # dark navy – original
_HDR_FILL_CORR = PatternFill("solid", fgColor="375623")   # dark green – corrected
_HDR_FONT      = Font(bold=True, color="FFFFFF", name="Arial", size=10)
_DATA_FONT     = Font(name="Arial", size=10)
_CHANGED_FILL  = PatternFill("solid", fgColor="E2EFDA")   # light green – cell changed
_ALT_FILL      = PatternFill("solid", fgColor="F2F2F2")   # light grey – alt rows
_BORDER_SIDE   = Side(style="thin", color="BFBFBF")
_CELL_BORDER   = Border(
    left=_BORDER_SIDE, right=_BORDER_SIDE,
    top=_BORDER_SIDE,  bottom=_BORDER_SIDE,
)


def _write_excel(result_df, orig_cols, corrected_col_map, output_path, col_map):
    wb = Workbook()
    ws = wb.active
    ws.title = "Corrected Addresses"

    all_cols = list(result_df.columns)
    orig_set  = set(orig_cols)
    corr_set  = set(corrected_col_map.values())

    # ── Header row ──────────────────────────────────────────────────────────
    for c_idx, col_name in enumerate(all_cols, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.font      = _HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = _CELL_BORDER
        if col_name in corr_set:
            cell.fill = _HDR_FILL_CORR
        else:
            cell.fill = _HDR_FILL_ORIG

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    # ── Data rows ───────────────────────────────────────────────────────────
    # Build reverse map: corrected_col → original_col for change detection
    rev_corr = {v: k for k, v in corrected_col_map.items()}

    for r_idx, row in enumerate(result_df.itertuples(index=False), start=2):
        alt = (r_idx % 2 == 0)
        for c_idx, col_name in enumerate(all_cols, start=1):
            val = getattr(row, col_name.replace(" ", "_").replace("/", "_")
                          .replace("(", "").replace(")", ""), None)
            # Fallback using positional index
            val = result_df.iloc[r_idx - 2, c_idx - 1]

            cell = ws.cell(row=r_idx, column=c_idx, value=val if val != "" else None)
            cell.font      = _DATA_FONT
            cell.alignment = Alignment(vertical="center")
            cell.border    = _CELL_BORDER

            # Highlight changed corrected cells
            if col_name in corr_set:
                orig_col = rev_corr[col_name]
                orig_val = result_df.iloc[r_idx - 2][orig_col] if orig_col else ""
                if str(val or "").strip() != str(orig_val or "").strip():
                    cell.fill = _CHANGED_FILL
                elif alt:
                    cell.fill = _ALT_FILL
            elif alt:
                cell.fill = _ALT_FILL

    # ── Column widths ────────────────────────────────────────────────────────
    for c_idx, col_name in enumerate(all_cols, start=1):
        col_vals = result_df.iloc[:, c_idx - 1].astype(str)
        max_len  = max(col_vals.str.len().max(), len(col_name))
        ws.column_dimensions[get_column_letter(c_idx)].width = min(max_len + 4, 40)

    # ── Summary sheet ────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Address Correction Summary"
    ws2["A1"].font = Font(bold=True, name="Arial", size=13)
    ws2.merge_cells("A1:C1")

    ws2["A3"] = "Total Rows Processed"
    ws2["B3"] = len(result_df)
    ws2["A4"] = "Columns Detected & Corrected"
    ws2["B4"] = len(corrected_col_map)

    ws2["A6"] = "Field"
    ws2["B6"] = "Original Column"
    ws2["C6"] = "Corrections Made"
    for cell in (ws2["A6"], ws2["B6"], ws2["C6"]):
        cell.font = Font(bold=True, name="Arial", size=10)
        cell.fill = _HDR_FILL_ORIG
        cell.font = _HDR_FONT

    row_num = 7
    for field, orig_col in col_map.items():
        if orig_col is None:
            continue
        corr_col = corrected_col_map.get(f"Corrected {DISPLAY_LABELS[field]}")
        if corr_col is None:
            continue
        changed = (
            result_df[orig_col].astype(str).str.strip()
            != result_df[corr_col].astype(str).str.strip()
        ).sum()
        ws2.cell(row=row_num, column=1, value=DISPLAY_LABELS[field]).font = _DATA_FONT
        ws2.cell(row=row_num, column=2, value=orig_col).font              = _DATA_FONT
        ws2.cell(row=row_num, column=3, value=int(changed)).font          = _DATA_FONT
        row_num += 1

    for col in ("A", "B", "C"):
        ws2.column_dimensions[col].width = 28

    wb.save(output_path)


# ──────────────────────────────────────────────────────────────────────────────
# MAIN PROCESS
# ──────────────────────────────────────────────────────────────────────────────

def process_file(input_path, output_path):
    ext = os.path.splitext(input_path)[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(input_path, dtype=str)
    elif ext == ".csv":
        df = pd.read_csv(input_path, dtype=str)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Use .csv, .xls, or .xlsx")

    df = df.fillna("")

    col_map = _detect_columns(df.columns)

    # Print detected columns
    print("\nDetected columns:")
    for field, col in col_map.items():
        status = f'"{col}"' if col else "NOT FOUND"
        print(f"  {DISPLAY_LABELS[field]:20s} -> {status}")

    # Apply corrections
    result = df.copy()
    corrected_col_map = {}  # "Corrected X" → series

    for field, orig_col in col_map.items():
        if orig_col is None:
            continue
        label = f"Corrected {DISPLAY_LABELS[field]}"
        result[label] = df[orig_col].apply(CORRECTORS[field])
        corrected_col_map[label] = label

    # Write output
    out_ext = os.path.splitext(output_path)[1].lower()
    if out_ext == ".csv":
        result.to_csv(output_path, index=False)
        print(f"\nDone. Output written to: {output_path}")
    else:
        _write_excel(result, list(df.columns), corrected_col_map, output_path, col_map)
        print(f"\nDone. Output written to: {output_path}")

    # Print summary
    print(f"Rows processed     : {len(df)}")
    print(f"Corrected columns  : {len(corrected_col_map)}")
    total_changes = 0
    for field, orig_col in col_map.items():
        if orig_col is None:
            continue
        label = f"Corrected {DISPLAY_LABELS[field]}"
        changed = (
            result[orig_col].astype(str).str.strip()
            != result[label].astype(str).str.strip()
        ).sum()
        total_changes += changed
        if changed:
            print(f"  {DISPLAY_LABELS[field]:20s}: {changed} cell(s) corrected")
    print(f"Total corrections  : {total_changes}")

    return result


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        print("Example:")
        print("  python address_corrector.py addresses.csv   corrected.xlsx")
        print("  python address_corrector.py addresses.xlsx  corrected.xlsx")
        sys.exit(1)

    input_file  = sys.argv[1]
    output_file = sys.argv[2]

    if not os.path.exists(input_file):
        print(f"Error: Input file not found: {input_file}")
        sys.exit(1)

    process_file(input_file, output_file)

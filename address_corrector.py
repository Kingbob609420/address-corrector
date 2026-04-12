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
COUNTRY_MAP = {
    # United States
    "usa": "United States", "u.s.a": "United States", "u.s.a.": "United States",
    "us": "United States", "u.s.": "United States", "u.s": "United States",
    "united states of america": "United States", "america": "United States",
    "the united states": "United States",
    # United Kingdom
    "uk": "United Kingdom", "u.k.": "United Kingdom", "u.k": "United Kingdom",
    "great britain": "United Kingdom", "gb": "United Kingdom", "g.b.": "United Kingdom",
    "england": "United Kingdom", "britain": "United Kingdom",
    # UAE
    "uae": "United Arab Emirates", "u.a.e.": "United Arab Emirates",
    "u.a.e": "United Arab Emirates", "emirates": "United Arab Emirates",
    # Australia
    "aus": "Australia", "au": "Australia",
    # Canada
    "can": "Canada", "ca": "Canada",
    # Germany
    "germany": "Germany", "deutschland": "Germany", "de": "Germany",
    # France
    "fr": "France",
    # Italy
    "it": "Italy", "italia": "Italy",
    # Spain
    "es": "Spain", "espana": "Spain", "españa": "Spain",
    # Netherlands
    "nl": "Netherlands", "the netherlands": "Netherlands", "holland": "Netherlands",
    # Belgium
    "be": "Belgium", "belgique": "Belgium",
    # Switzerland
    "ch": "Switzerland", "schweiz": "Switzerland", "suisse": "Switzerland",
    # Sweden
    "se": "Sweden", "sverige": "Sweden",
    # Norway
    "no": "Norway", "norge": "Norway",
    # Denmark
    "dk": "Denmark", "danmark": "Denmark",
    # Finland
    "fi": "Finland", "suomi": "Finland",
    # Portugal
    "pt": "Portugal",
    # Ireland
    "ie": "Ireland", "republic of ireland": "Ireland", "eire": "Ireland",
    # New Zealand
    "nz": "New Zealand",
    # South Africa
    "sa": "South Africa", "rsa": "South Africa",
    # Singapore
    "sg": "Singapore",
    # India
    "in": "India", "ind": "India",
    # China
    "cn": "China", "prc": "China", "people's republic of china": "China",
    "peoples republic of china": "China",
    # Japan
    "jp": "Japan",
    # South Korea
    "kr": "South Korea", "korea": "South Korea", "republic of korea": "South Korea",
    # Brazil
    "br": "Brazil", "brasil": "Brazil",
    # Mexico
    "mx": "Mexico", "mex": "Mexico",
    # Argentina
    "ar": "Argentina",
    # Chile
    "cl": "Chile",
    # Colombia
    "co": "Colombia",
    # Russia
    "ru": "Russia", "russian federation": "Russia",
    # Turkey
    "tr": "Turkey", "türkiye": "Turkey", "turkiye": "Turkey",
    # Saudi Arabia
    "ksa": "Saudi Arabia", "kingdom of saudi arabia": "Saudi Arabia",
    # Israel
    "il": "Israel",
    # Poland
    "pl": "Poland", "polska": "Poland",
    # Czech Republic
    "cz": "Czech Republic", "czechia": "Czech Republic",
    # Hungary
    "hu": "Hungary", "magyarország": "Hungary",
    # Greece
    "gr": "Greece", "hellas": "Greece",
    # Romania
    "ro": "Romania",
    # Ukraine
    "ua": "Ukraine",
    # Pakistan
    "pk": "Pakistan",
    # Bangladesh
    "bd": "Bangladesh",
    # Indonesia
    "id": "Indonesia",
    # Malaysia
    "my": "Malaysia",
    # Philippines
    "ph": "Philippines",
    # Thailand
    "th": "Thailand",
    # Vietnam
    "vn": "Vietnam",
    # Egypt
    "eg": "Egypt",
    # Nigeria
    "ng": "Nigeria",
    # Kenya
    "ke": "Kenya",
    # Ghana
    "gh": "Ghana",
    # Morocco
    "ma": "Morocco",
    # Hong Kong
    "hk": "Hong Kong",
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

    return " ".join(result)


def correct_city(val):
    val = _clean(val)
    if not val:
        return ""
    # Title case but keep conjunctions lowercase (e.g. "Port-au-Prince")
    parts = re.split(r"([-/])", val)
    return "".join(p.title() if not p in ("-", "/") else p for p in parts)


def correct_state(val, country_hint=""):
    val = _clean(val)
    if not val:
        return ""
    lower = val.lower()

    # Try US state full name → abbreviation
    if lower in US_STATES:
        return US_STATES[lower]

    # Try Canadian province full name → abbreviation
    if lower in CA_PROVINCES:
        return CA_PROVINCES[lower]

    # Try Australian state full name → abbreviation
    if lower in AU_STATES:
        return AU_STATES[lower]

    # Pure alphabetic short codes (2–4 chars) → uppercase  (e.g. "ny", "nsw", "qld", "on")
    if len(val) <= 4 and val.replace(" ", "").isalpha():
        return val.upper()

    # Anything longer that looks like an abbreviation (ALL CAPS already) → keep uppercase
    if val.isupper():
        return val

    # Default: Title Case (works for any language / region name)
    return val.title()


def correct_country(val):
    val = _clean(val)
    if not val:
        return ""
    lookup = val.lower().strip(" .")
    if lookup in COUNTRY_MAP:
        return COUNTRY_MAP[lookup]
    # Try without dots  (u.s.a → usa)
    no_dots = lookup.replace(".", "")
    if no_dots in COUNTRY_MAP:
        return COUNTRY_MAP[no_dots]
    # 2-letter ISO country code → uppercase  (any country)
    if len(val.replace(".", "").replace(" ", "")) == 2 and val.replace(".", "").isalpha():
        return val.replace(".", "").upper()
    # 3-letter ISO code → uppercase
    if len(val.replace(".", "").replace(" ", "")) == 3 and val.replace(".", "").isalpha():
        return val.replace(".", "").upper()
    # Everything else → Title Case  (works for any country name in any language)
    return val.title()


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

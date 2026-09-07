"""
matcher.py -- Supplier matching (Phase 2).

Implements TWO matchers, each with its own normalization:

1. SENSITIVE-SUPPLIER check  (SensitiveMatcher)
   * Uses normalize() -- STRIPS legal suffixes (Inc, Ltd, GmbH, SAS, ...).
   * GENEROUS matching -- false negative is the bad outcome.  Threshold 85.
   * Two outcomes: SENSITIVE vs NOT SENSITIVE.

2. LEGAL-ENTITY (LE) matcher  (LEMatcher)
   * Uses normalize_le() -- KEEPS legal forms and EXPANDS abbreviations so
     that VLM output variations still match the reference.
     Expansions grounded in the actual LE_List.xlsx data:
       Ltd→Limited, AB→Aktiebolag, Corp→Corporation, KK→Kabushiki Kaisha,
       Pvt→Private, Ste→Societe, Hiz→Hizmetleri, STI→Sirketi, Ctr→Center,
       z.o.o.→spzoo, d.o.o.→doo.
     Tokens kept as-is (they distinguish LEs): GmbH, SAS, SA, NV, BV, AG,
       SIA, OY, AS, Aps, SRL, SARL, LDA, LLC, KG, Kft, Sdn, Bhd, Pte, Pty,
       Inc, Co, SL, SPA.
     NOTHING is stripped -- every token matters.
   * CJK whitespace collapse (OCR inserts random spaces into CJK strings).
   * TIGHT matching -- wrong LE = wrong Kofax routing.  Threshold 93.
   * token_sort_ratio ONLY (no partial_ratio).
   * /-separated aliases split and flattened.

NON-LATIN SCRIPT HANDLING
  Both normalize() and normalize_le() previously had a real gap: their
  final step is a Latin-only regex ([^a-z0-9&\\s] -> space). For a name
  written in a script with NO Latin characters at all -- Thai, Arabic,
  Devanagari, Cyrillic, and (for normalize(), which never had ANY
  non-Latin handling) even CJK -- this silently stripped the entire name
  to an empty string. An empty normalized key can never match anything,
  so any invoice where the VLM failed to translate a non-English name
  (a real, observed failure mode -- see TRANSLATE-ON-MISS below) would
  always come back "no match", with no chance to even try.
  Fixed with _raw_fallback(): if the Latin-normalized result is empty but
  the original text had real content, fall back to a lightly-cleaned RAW
  version (lowercased, whitespace-collapsed, script characters kept
  as-is) instead of returning "". This at least allows an EXACT match
  against another occurrence of the same non-Latin name, and gives the
  translate-on-miss fallback below a non-empty starting point.
  Also added: explicit Vietnamese đ/Đ -> d/D handling (NFKD does not
  decompose this letter the way it does ordinary accented vowels).

TRANSLATE-ON-MISS
  Both SensitiveMatcher.check() and LEMatcher.match() now accept an
  optional `translator` argument (any object with a .translate(text)
  method, e.g. vietnam_renamer.GoogleTranslateTranslator). Default is
  None, so existing calls are unaffected.
  When a translator IS given:
    1. Match the ORIGINAL extracted text first, as always.
    2. If that comes back "no_match" AND the original text contains
       non-ASCII characters, translate it and match the TRANSLATED text.
    3. Return whichever attempt succeeded (translation is only ever a
       fallback, never tried first) -- with translated/translated_query
       fields added to the result so callers/logs can see which path
       found it.
  This covers the case where the VLM was supposed to translate a non-
  English name in-prompt but didn't (an observed failure mode with the
  local VLM), independent of whether the name has a genuine non-English
  reference entry or not -- the retry is free if it isn't needed (only
  fires on a real miss with non-ASCII input), and doesn't change behavior
  for any name that already matches in its original form.
"""

import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import Optional, List, Dict

import pandas as pd
from rapidfuzz import fuzz, process


# --- Config ----------------------------------------------------------------
SENSITIVE_THRESHOLD = 90
LE_THRESHOLD = 85

# partial_ratio guard for sensitive matcher -- both names must be at least
# this long before partial_ratio is allowed to contribute.
_MIN_LEN_FOR_PARTIAL = 6


def _has_special_chars(text: Optional[str]) -> bool:
    """True if `text` has any non-ASCII character -- the signal used to
    decide whether a failed match is worth retrying via translation
    (translate-on-miss only fires for non-ASCII input; retranslating
    already-plain-ASCII text that didn't match wouldn't change anything)."""
    return any(ord(c) > 127 for c in (text or ""))


def _raw_fallback(original: str) -> str:
    """Used when the Latin-only normalization regex strips a name to
    nothing (a script with no Latin characters at all -- Thai, Arabic,
    Devanagari, Cyrillic, and for normalize() specifically, CJK too).
    Instead of returning '', fall back to a lightly-cleaned raw form:
    lowercased, whitespace-collapsed, script characters kept as-is. This
    at least allows an exact match against another occurrence of the
    same non-Latin name, rather than guaranteeing "no_match" outright."""
    return re.sub(r"\s+", " ", original.strip().lower())


# ---------------------------------------------------------------------------
# Shared abbreviation expansions (used by BOTH normalizers)
# ---------------------------------------------------------------------------
# Word abbreviations that carry identifying meaning.  Legal-form tokens
# (Ltd, GmbH, SAS ...) are deliberately NOT here -- they are handled
# differently by each normalizer.
_ABBREVIATIONS = {
    "svcs": "services", "svc": "services", "serv": "services",
    "corp": "corporation", "corpn": "corporation",
    "intl": "international", "intnl": "international", "internatl": "international",
    "natl": "national",
    "tech": "technology", "techs": "technologies", "technol": "technology",
    "sys": "systems", "syst": "systems",
    "comm": "communications", "comms": "communications",
    "commun": "communications",
    "elec": "electronics", "elect": "electronics", "electr": "electronics",
    "mfg": "manufacturing", "mfrs": "manufacturers", "mfr": "manufacturer",
    "div": "division", "divn": "division",
    "prod": "products", "prods": "products", "prd": "products",
    "sol": "solutions", "soln": "solutions", "solns": "solutions",
    "engg": "engineering", "engr": "engineering", "eng": "engineering",
    "ind": "industries", "inds": "industries", "indl": "industrial",
    "assoc": "associates", "assocs": "associates", "assn": "association",
    "grp": "group",
    "dist": "distribution", "distrib": "distribution",
    "mkt": "marketing", "mktg": "marketing", "mark": "marketing",
    "consult": "consulting", "cons": "consulting",
    "mgmt": "management", "mgt": "management",
    "sci": "scientific", "lab": "laboratories", "labs": "laboratories",
    "pharma": "pharmaceuticals", "pharm": "pharmaceuticals",
    "ctr": "center",
    "rnd": "research and development",
    "&": "and",
}

# --- Sensitive-only: legal suffixes REMOVED during normalization ----------
_LEGAL_SUFFIXES = {
    "inc", "incorporated", "llc", "ltd", "limited", "co", "company",
    "gmbh", "ag", "sa", "sas", "sl", "srl", "spa", "bv", "nv", "plc",
    "pte", "pvt", "pty", "kg", "ab", "oy", "as", "aps", "sdn", "bhd",
    "lp", "llp", "sau", "sac", "sarl", "kk", "kabushiki", "ltda", "ooo",
    "sia", "oyj", "kft", "doo", "gk", "ug", "sro", "spzoo", "aktiebolag",
    "aktiengesellschaft", "sociedad", "anonima", "societe", "anonyme",
}

# --- LE-only: legal-form EXPANSIONS ----------------------------------------
_LE_LEGAL_EXPANSIONS = {
    "ltd":    "limited",

    # Evidence: LE400 has both "Intel Sweden AB" and "Intel Sweden Aktiebolag"
    "ab":     "aktiebolag",

    # Evidence: LE828/831 use "Pvt", LE827/841/842 use "Private"
    "pvt":    "private",

    # Evidence: LE330 alias "STE Intel Corporation SAS Toulouse"
    "ste":    "societe",

    # Evidence: LE571 "Intel Teknoloji Hiz LTD STI" is alias of
    # "Intel Teknoloji Hizmetleri Limited Sirketi"
    "hiz":    "hizmetleri",
    "sti":    "sirketi",
}

# CJK Unicode ranges for detecting CJK-heavy strings.
_CJK_RE = re.compile(
    r'[\u4e00-\u9fff'     # CJK Unified Ideographs
    r'\u3400-\u4dbf'      # CJK Unified Ideographs Extension A
    r'\uf900-\ufaff'      # CJK Compatibility Ideographs
    r'\u3000-\u303f'      # CJK Symbols and Punctuation
    r'\u3040-\u309f'      # Hiragana
    r'\u30a0-\u30ff'      # Katakana
    r'\uac00-\ud7af]'     # Hangul
)

# Known CJK bill-to names → canonical English mapping.
_CJK_CANONICAL = {
    "美商英特爾亞太科技有限公司台灣分公司": "Intel Microelectronics Asia LLC Taiwan Branch",
    "英特尔创新科技股份有限公司": "Intel Innovation Technology Ltd",
    # Add more CJK -> English mappings here as needed
}


# ---------------------------------------------------------------------------
# normalize() — for SENSITIVE matcher (strips legal suffixes)
# ---------------------------------------------------------------------------
def normalize(name: Optional[str]) -> str:
    """Canonicalize a company name for SENSITIVE matching:
      - Vietnamese đ/Đ -> d/D explicitly (NFKD doesn't decompose this the
        way it does ordinary accented vowels)
      - unicode-normalize + strip accents (é -> e)
      - lowercase
      - drop punctuation -> spaces (but keep & to expand to 'and')
      - EXPAND abbreviations (svcs->services, corp->corporation, ...)
      - REMOVE standalone legal suffixes (inc, ltd, gmbh, sas, ...)
      - collapse whitespace
      - if the result is EMPTY but the input had real content (a script
        with no Latin characters at all -- CJK, Thai, Arabic, Devanagari,
        Cyrillic, ...), fall back to a lightly-cleaned raw form instead
        of returning '' (see _raw_fallback)
    Returns '' for empty/na input.
    """
    if not name or str(name).strip().lower() in ("", "nan", "none"):
        return ""
    original = str(name)
    s = original.replace("\u0110", "D").replace("\u0111", "d")  # Đ / đ
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()
    s = s.replace("&", " & ")
    s = re.sub(r"[^a-z0-9&\s]", " ", s)

    out = []
    for t in s.split():
        if not t:
            continue
        t = _ABBREVIATIONS.get(t, t)
        if t in _LEGAL_SUFFIXES:
            continue
        out.append(t)
    result = " ".join(out)
    if not result and original.strip():
        return _raw_fallback(original)
    return result


# ---------------------------------------------------------------------------
# normalize_le() — for LE matcher (KEEPS legal forms, EXPANDS abbreviations)
# ---------------------------------------------------------------------------
def normalize_le(name: Optional[str]) -> str:
    """Canonicalize a company name for LE matching.

    Key difference from normalize(): legal-form tokens (GmbH, SAS, SA, NV,
    Ltd, ...) are KEPT and EXPANDED where the VLM might use a different form.
    Nothing is stripped — every token matters for distinguishing Intel LEs.

    Steps:
      - Vietnamese đ/Đ -> d/D (NFKD doesn't decompose this letter)
      - CJK: collapse whitespace, check _CJK_CANONICAL for known mapping
      - unicode-normalize + strip accents
      - pre-process dotted abbreviations (K.K., z.o.o., d.o.o., R&D)
      - lowercase, punctuation -> spaces
      - expand general abbreviations (corp->corporation, tech->technology, ...)
      - expand LE legal-form abbreviations (ltd->limited, ab->aktiebolag, ...)
      - collapse whitespace
      - if the result is EMPTY but the input had real content (a non-Latin,
        non-CJK script -- Thai, Arabic, Devanagari, Cyrillic, ...), fall
        back to a lightly-cleaned raw form instead of returning ''
    Returns '' for empty/na input.
    """
    if not name or str(name).strip().lower() in ("", "nan", "none"):
        return ""
    original = str(name)
    s = original.replace("\u0110", "D").replace("\u0111", "d")  # Đ / đ

    # --- CJK handling: collapse whitespace and try canonical lookup ----------
    cjk_chars = _CJK_RE.findall(s)
    if len(cjk_chars) >= 3:
        collapsed = re.sub(r'\s+', '', s)
        canonical = _CJK_CANONICAL.get(collapsed)
        if canonical:
            s = canonical  # continue normalizing the English form
        else:
            return collapsed  # no known mapping — return collapsed CJK as-is

    # --- Standard Latin normalization ----------------------------------------
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()

    # --- Pre-process dotted/spaced abbreviations BEFORE punctuation strip ----
    # These get mangled by the general [^a-z0-9&\s] -> " " rule.
    s = re.sub(r'\bk\.?\s*k\.?\b', 'kabushiki kaisha', s)    # K.K. / K.K / KK
    s = re.sub(r'\bz\.?\s*o\.?\s*o\.?\b', 'spzoo', s)        # z.o.o. / z.o.o / zoo
    s = re.sub(r'\bd\.?\s*o\.?\s*o\.?\b', 'doo', s)           # d.o.o. / d.o.o
    s = re.sub(r'\br\s*&\s*d\b', 'research and development', s)  # R&D / R & D
    # Dotted legal forms — collapse dots so "S.A." stays "sa" not "s a"
    s = re.sub(r'\bs\.\s*p\.\s*a\.?\b', 'spa', s)             # S.P.A. / S.P.A
    s = re.sub(r'\bs\.\s*a\.\s*s\.?\b', 'sas', s)             # S.A.S. (rare but safe)
    s = re.sub(r'\bs\.\s*a\.?\b', 'sa', s)                    # S.A. / S.A
    s = re.sub(r'\bs\.\s*l\.?\b', 'sl', s)                    # S.L. / S.L
    s = re.sub(r'\bb\.\s*v\.?\b', 'bv', s)                    # B.V. / B.V
    s = re.sub(r'\bn\.\s*v\.?\b', 'nv', s)                    # N.V.

    s = s.replace("&", " & ")
    s = re.sub(r"[^a-z0-9&\s]", " ", s)

    out = []
    for t in s.split():
        if not t:
            continue
        # 1) General abbreviations first (corp->corporation, tech->technology ...)
        t = _ABBREVIATIONS.get(t, t)
        # 2) LE-specific legal-form expansions (ltd->limited, ab->aktiebolag ...)
        t = _LE_LEGAL_EXPANSIONS.get(t, t)
        out.append(t)
    result = " ".join(out)
    if not result and original.strip():
        # A non-Latin, non-CJK script (Thai, Arabic, Devanagari, Cyrillic,
        # ...) has no Latin characters for the regex above to keep, and
        # isn't covered by the CJK branch either -- fall back to a raw
        # cleaned form instead of returning '' (see _raw_fallback).
        return _raw_fallback(original)
    return result


# ---------------------------------------------------------------------------
# Stage 0: build the SENSITIVE reference JSON (run offline)
# ---------------------------------------------------------------------------
def build_sensitive_reference(xlsx_path: str,
                              out_path: str = "sensitive_reference.json",
                              sheet: str = "Sensitive supplier list",
                              name_column: str = "Supplier Remit to Name",
                              header_row: int = 1) -> dict:
    """Read the sensitive-supplier Excel, dedupe to a set of normalized names,
    and write a JSON reference."""
    df = pd.read_excel(xlsx_path, sheet_name=sheet, header=header_row)
    raw_names = df[name_column].dropna().astype(str).str.strip()
    raw_names = raw_names[raw_names.ne("")]

    norm_to_original: Dict[str, str] = {}
    skipped = 0
    for original in raw_names:
        key = normalize(original)
        if not key:
            skipped += 1
            continue
        norm_to_original.setdefault(key, original)

    ref = {
        "source_file": str(xlsx_path),
        "threshold": SENSITIVE_THRESHOLD,
        "count_raw_rows": int(len(raw_names)),
        "count_unique_normalized": len(norm_to_original),
        "skipped_unnormalizable": skipped,
        "names": norm_to_original,
    }
    Path(out_path).write_text(json.dumps(ref, ensure_ascii=False, indent=2),
                              encoding="utf-8")
    print(f"Wrote {out_path}: {len(raw_names)} rows -> "
          f"{len(norm_to_original)} unique normalized names ({skipped} skipped).")
    return ref


# ---------------------------------------------------------------------------
# Runtime SENSITIVE matcher
# ---------------------------------------------------------------------------
class SensitiveMatcher:
    def __init__(self, reference_path: str = "sensitive_reference.json",
                 threshold: int = SENSITIVE_THRESHOLD):
        ref = json.loads(Path(reference_path).read_text(encoding="utf-8"))
        self.threshold = threshold
        self.norm_to_original: Dict[str, str] = ref["names"]
        self._norm_keys: List[str] = list(self.norm_to_original.keys())

    def check(self, supplier_name: Optional[str],
             translator=None) -> dict:
        """Return {is_sensitive, matched_name, score, method, translated,
        translated_query} for one name.

        `translator`: optional object with a .translate(text) method (e.g.
        vietnam_renamer.GoogleTranslateTranslator). If given, and the
        ORIGINAL name doesn't match, AND the original name has non-ASCII
        characters, it's translated and matched again. This covers the
        case where the VLM was supposed to translate a non-English name
        in-prompt but didn't -- an observed real failure mode, independent
        of whether this specific name happens to need it. The original is
        always tried first; translation is only ever a fallback.
        """
        result = self._check_once(supplier_name)
        if (result["method"] == "no_match" and translator is not None
                and _has_special_chars(supplier_name)):
            translated = translator.translate(supplier_name)
            retry = self._check_once(translated)
            retry["translated"] = True
            retry["translated_query"] = translated
            return retry
        result["translated"] = False
        result["translated_query"] = None
        return result

    def _check_once(self, supplier_name: Optional[str]) -> dict:
        """The original single-pass matching logic (Return
        {is_sensitive, matched_name, score, method} for one name), used
        by check() for both the original-text attempt and the translated
        retry."""
        key = normalize(supplier_name)
        if not key:
            return {"is_sensitive": False, "matched_name": None,
                    "score": 0, "method": "empty"}

        # 1) Normalized exact
        if key in self.norm_to_original:
            return {"is_sensitive": True,
                    "matched_name": self.norm_to_original[key],
                    "score": 100, "method": "exact"}

        # 2) Fuzzy — token_sort_ratio + guarded partial_ratio
        best_key = None
        best_score = 0
        key_len = len(key)
        for nk in self._norm_keys:
            base = fuzz.token_sort_ratio(key, nk)
            if min(key_len, len(nk)) >= _MIN_LEN_FOR_PARTIAL:
                base = max(base, fuzz.partial_ratio(key, nk))
            if base > best_score:
                best_score, best_key = base, nk

        if best_score >= self.threshold:
            return {"is_sensitive": True,
                    "matched_name": self.norm_to_original.get(best_key),
                    "score": int(best_score), "method": "fuzzy"}

        return {"is_sensitive": False,
                "matched_name": self.norm_to_original.get(best_key),
                "score": int(best_score), "method": "no_match"}


# ---------------------------------------------------------------------------
# Stage 0-LE: build the LE reference JSON from LE_List.xlsx (run offline)
# ---------------------------------------------------------------------------
_LE_JUNK_PATTERNS = [
    "do not scan",
    "send it to",
    "@",
]


def _is_junk_le_name(name: str) -> bool:
    """Return True if the Bill to Name is an instruction row, not a real entity."""
    low = name.strip().lower()
    return any(pat in low for pat in _LE_JUNK_PATTERNS)


def build_le_reference(xlsx_path: str,
                       out_path: str = "le_reference.json",
                       sheet: str = "LEs",
                       ) -> dict:
    """Read LE_List.xlsx, clean, split /-separated aliases, flatten, and write
    a JSON reference for the LE matcher.

    Uses normalize_le() (keeps legal forms) instead of normalize().
    Prints warnings for any collisions (different LE codes normalizing to the
    same string).
    """
    df = pd.read_excel(xlsx_path, sheet_name=sheet)

    def _find_col(test_fn, label):
        matches = [c for c in df.columns if test_fn(c)]
        if not matches:
            raise ValueError(
                f"Could not find '{label}' column in sheet '{sheet}'. "
                f"Columns found: {list(df.columns)}"
            )
        return matches[0]

    col_le = _find_col(lambda c: str(c).strip().upper() == "LEGAL ENTITY (LE)", "LE")
    col_bill = _find_col(lambda c: "bill" in str(c).lower() and "name" in str(c).lower(),
                         "Bill to Name")
    col_site = _find_col(lambda c: "source" in str(c).lower() and "site" in str(c).lower(),
                         "Source Site")
    col_country = _find_col(lambda c: str(c).strip().lower() == "country", "Country")

    entries: List[dict] = []
    norm_to_indices: Dict[str, List[int]] = {}
    skipped_blank = 0
    skipped_junk = 0

    for _, row in df.iterrows():
        le_code = str(row[col_le]).strip() if pd.notna(row[col_le]) else ""
        bill_raw = str(row[col_bill]).strip() if pd.notna(row[col_bill]) else ""
        site = str(row[col_site]).strip() if pd.notna(row[col_site]) else ""
        country = str(row[col_country]).strip() if pd.notna(row[col_country]) else ""

        if not le_code or not bill_raw or bill_raw.lower() in ("nan", "none", ""):
            skipped_blank += 1
            continue

        if _is_junk_le_name(bill_raw):
            skipped_junk += 1
            continue

        # Split on "/" for aliases, trim each
        aliases = [a.strip() for a in bill_raw.split("/") if a.strip()]

        for alias in aliases:
            norm = normalize_le(alias)
            if not norm:
                skipped_blank += 1
                continue
            idx = len(entries)
            entries.append({
                "normalized": norm,
                "original": alias,
                "le": le_code,
                "source_site": site,
                "country": country,
            })
            norm_to_indices.setdefault(norm, []).append(idx)

    # --- Collision check: warn if different LE codes share a normalized key --
    collisions = []
    for norm_key, indices in norm_to_indices.items():
        le_codes = set(entries[i]["le"] for i in indices)
        if len(le_codes) > 1:
            details = [(entries[i]["le"], entries[i]["original"]) for i in indices]
            collisions.append((norm_key, details))
            print(f"  WARNING: COLLISION on '{norm_key}':")
            for lc, orig in details:
                print(f"    {lc}: {orig}")

    ref = {
        "source_file": str(xlsx_path),
        "threshold": LE_THRESHOLD,
        "count_entries": len(entries),
        "skipped_blank": skipped_blank,
        "skipped_junk": skipped_junk,
        "collisions": len(collisions),
        "entries": entries,
        "norm_to_indices": norm_to_indices,
    }
    Path(out_path).write_text(json.dumps(ref, ensure_ascii=False, indent=2),
                              encoding="utf-8")
    print(f"\nWrote {out_path}: {len(entries)} flattened aliases "
          f"({skipped_blank} blank, {skipped_junk} junk skipped, "
          f"{len(collisions)} collisions).")
    return ref


# ---------------------------------------------------------------------------
# Runtime LE matcher
# ---------------------------------------------------------------------------
class LEMatcher:
    """Match a bill-to name against the LE reference.

    Design: TIGHT threshold — wrong LE = wrong Kofax inbox.
    Uses normalize_le() (keeps legal forms) and token_sort_ratio ONLY.
    """

    def __init__(self, reference_path: str = "le_reference.json",
                 threshold: int = LE_THRESHOLD):
        ref = json.loads(Path(reference_path).read_text(encoding="utf-8"))
        self.threshold = threshold
        self.entries: List[dict] = ref["entries"]
        self.norm_to_indices: Dict[str, List[int]] = ref["norm_to_indices"]
        self._norm_keys: List[str] = list(self.norm_to_indices.keys())

    def match(self, bill_to_name: Optional[str],
             translator=None) -> dict:
        """Return the best LE match for a bill-to name.

        `translator`: optional object with a .translate(text) method (e.g.
        vietnam_renamer.GoogleTranslateTranslator). If given, and the
        ORIGINAL name doesn't match, AND the original name has non-ASCII
        characters, it's translated and matched again -- same fallback
        pattern as SensitiveMatcher.check(). The original is always tried
        first; translation is only ever a fallback.

        Returns:
          {
            "matched": True/False,
            "matched_le_name": str or None,
            "le": str or None,
            "source_site": str or None,
            "country": str or None,
            "score": int,
            "method": "exact" | "fuzzy" | "no_match" | "empty",
            "translated": bool,
            "translated_query": str or None,
          }
        """
        result = self._match_once(bill_to_name)
        if (result["method"] == "no_match" and translator is not None
                and _has_special_chars(bill_to_name)):
            translated = translator.translate(bill_to_name)
            retry = self._match_once(translated)
            retry["translated"] = True
            retry["translated_query"] = translated
            return retry
        result["translated"] = False
        result["translated_query"] = None
        return result

    def _match_once(self, bill_to_name: Optional[str]) -> dict:
        """The original single-pass matching logic, used by match() for
        both the original-text attempt and the translated retry."""
        key = normalize_le(bill_to_name)
        if not key:
            return self._result(matched=False, method="empty")

        # 1) Normalized exact
        if key in self.norm_to_indices:
            idx = self.norm_to_indices[key][0]
            e = self.entries[idx]
            return self._result(matched=True, entry=e, score=100, method="exact")

        # 2) Fuzzy — token_sort_ratio ONLY (tight, no partial_ratio).
        best_key = None
        best_score = 0
        for nk in self._norm_keys:
            score = fuzz.token_sort_ratio(key, nk)
            if score > best_score:
                best_score, best_key = score, nk

        if best_score >= self.threshold and best_key is not None:
            idx = self.norm_to_indices[best_key][0]
            e = self.entries[idx]
            return self._result(matched=True, entry=e,
                                score=int(best_score), method="fuzzy")

        # No match — return best near-miss for review/logging
        near_entry = None
        if best_key is not None:
            idx = self.norm_to_indices[best_key][0]
            near_entry = self.entries[idx]
        return self._result(matched=False, entry=near_entry,
                            score=int(best_score), method="no_match")

    @staticmethod
    def _result(matched: bool, entry: Optional[dict] = None,
                score: int = 0, method: str = "empty") -> dict:
        if entry:
            return {
                "matched": matched,
                "matched_le_name": entry["original"],
                "le": entry["le"],
                "source_site": entry["source_site"],
                "country": entry["country"],
                "score": score,
                "method": method,
            }
        return {
            "matched": matched,
            "matched_le_name": None,
            "le": None,
            "source_site": None,
            "country": None,
            "score": score,
            "method": method,
        }


# ---------------------------------------------------------------------------
#  build references, or test names
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import pathlib
    le_xlsx_file = pathlib.Path(r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\GenAI-OCR-P2P-Protyping\Supporting Documents\Legal Entity Lists.xlsx")
    le_json_file = pathlib.Path(r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\GenAI-OCR-P2P-Protyping\Supporting Documents\le_reference.json")

    sensitive_xlsx_file = pathlib.Path(r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\GenAI-OCR-P2P-Protyping\Supporting Documents\Sensitive Supplier list_210_4423384668390299001.xlsx")
    senstive_output_json_file = pathlib.Path(r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\GenAI-OCR-P2P-Protyping\Supporting Documents\sensitive_reference.json")


    build_sensitive_reference(xlsx_path=sensitive_xlsx_file,
                              out_path=senstive_output_json_file)

 
    build_le_reference(xlsx_path=le_xlsx_file,
                       out_path=le_json_file)

    # le = LEMatcher(reference_path=le_json_file)
    se = SensitiveMatcher(reference_path=senstive_output_json_file)

    res = se.check('daewon semiconductor packaging industries')
    print(res)

            
            
    # tests = [   
    #         ("Intel Corporation",                   "exact - common bill-to"),
    #         ("Intel GmbH",                          "exact - German entity"),
    #         ("McAfee Norway AS",                    "exact - non-Intel LE"),
    #         ("Lantiq Latvia SIA",                   "exact - non-Intel LE"),
    #         # Alias matches (should hit via flattened /-split)
    #         ("STE Intel Corporation SAS Toulouse",  "exact - /-alias of LE330"),
    #         ("SAS Intel Corporation",               "exact - /-alias of LE330"),
    #         ("Intel Corporation SAS Toulouse", "LE330"),
    #         # Fuzzy matches (minor variations, should still match)
    #         ("INTEL CORPORATION",                   "case variation"),
    #         ("Intel Corp.",                         "abbreviation - Corp -> Corporation"),
    #         ("Intel Corp",                          "abbreviation no dot"),
    #         ("Intel  GmbH",                         "extra whitespace"),
    #         # Near-misses that must NOT cross-match (tight threshold)
    #         ("Intel Germany Services GmbH",         "must NOT match plain 'Intel GmbH'"),
    #         ("Intel Magdeburg GmbH",                "must NOT match plain 'Intel GmbH'"),
    #         # Completely unknown names
    #         ("Totally Random Company",              "should NOT match any LE"),
    #         ("Expeditors International France",     "should NOT match any LE"),
    #         # Edge cases
    #         ("",                                    "empty input"),
    #         (None,                                  "None input"),
    #         ]


    # for name, note in tests:
    #     r = le.match(name)
    #     status = f"{r['le']}" if r["matched"] else "NO MATCH"
    #     print(f"  [{status:>10}] {str(name):<45} score={r['score']:3d} "
    #           f"method={r['method']:<8} site={r['source_site']}  ({note})")
    # print()
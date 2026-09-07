import json
import os
import re
import unicodedata
from pathlib import Path
from typing import Optional, List, Dict, Tuple

import pandas as pd
from rapidfuzz import fuzz


VN_THRESHOLD = 80  

# Legal-suffix EXPANSIONS applied during normalization.
_VN_LEGAL_EXPANSIONS = {
    "ltd": "limited",
    "co": "company",
    "corp": "corporation",
    "corpn": "corporation",
    "inc": "incorporated",
    "pte": "private",
    "pvt": "private",     
    "sdn": "sendirian",
    "bhd": "berhad",
}

# Junk values seen in the 3rd (alt-name) column -- ignore these.
_ALT_JUNK = {".", "-", "", "nan", "none", "n/a"}

# Excel artifact pattern (cell-reference leaks like "+B33" glued onto text).
_ARTIFACT_RE = re.compile(r"\+[A-Z]\d+")


# ---------------------------------------------------------------------------
# Translation backends (pluggable)
# ---------------------------------------------------------------------------
class Translator:
    """Interface every translation backend implements. Swap backends
    without touching build or matching code -- only .translate() needs
    to work."""

    def translate(self, text: str) -> str:
        raise NotImplementedError


class GoogleTranslateTranslator(Translator):
    """Translator backed by deep_translator's GoogleTranslator (auto-detect
    source -> English). Used for BOTH build-time (translating the Excel
    reference) and runtime (translating a live VLM-extracted supplier
    name) in this module.

    pip install deep-translator
    """

    def __init__(self, target: str = "en"):
        from deep_translator import GoogleTranslator as _GT
        self._target = target
        self._GT = _GT

    def translate(self, text: str) -> str:
        return self._GT(source="auto", target=self._target).translate(text)

class QwenTranslator(Translator):
    """Translator backed by the local Qwen VLM server (see qwen_llm.py).
    Keeps translation on your own hardware -- no data leaves the network,
    no corporate-proxy dependency, no reliance on Google's free translate
    endpoint. Used for BOTH build-time and runtime translation, same as
    GoogleTranslateTranslator.
    """

    def __init__(self, context: str = "company name"):
        from qwen_llm import translate_to_english as _translate
        self._translate = _translate
        self._context = context

    def translate(self, text: str) -> str:
        return self._translate(text, context=self._context)

# ---------------------------------------------------------------------------
# Normalization -- Vietnamese-aware
# ---------------------------------------------------------------------------
def normalize_vn(name: Optional[str]) -> str:
    """Normalize a Vietnamese/English supplier name for matching.
      - dd/DD explicitly FIRST: Vietnamese "d with stroke" (\u0110/\u0111) is
        its own letter, not a base+combining-mark pair, so NFKD does NOT
        strip its stroke the way it does the tone/vowel marks.
      - NFKD + drop combining marks (handles all Vietnamese tone/vowel marks)
      - lowercase, punctuation -> spaces
      - EXPAND legal suffixes (ltd->limited, co->company, corp->corporation,
        ...) -- safe to expand (not just strip or keep-as-is) because this
        supplier list has no repeated names, so there's no risk of two
        different suppliers colliding the way LE codes could.
      - collapse whitespace
    Returns '' for empty/na input.
    """
    if not name or str(name).strip().lower() in ("", "nan", "none"):
        return ""
    s = str(name)
    s = s.replace("\u0110", "D").replace("\u0111", "d")  # Đ / đ
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    words = [_VN_LEGAL_EXPANSIONS.get(w, w) for w in s.split()]
    s = " ".join(words)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _has_special_chars(text: str) -> bool:
    """True if `text` has any non-ASCII character -- the signal, both at
    build time and runtime, for 'this is Vietnamese script -- translate
    it' vs 'this is already plain text -- use/match it directly'."""
    return any(ord(c) > 127 for c in text)


def _clean_artifact(text: str) -> str:
    """Strip known Excel artifacts and
    collapse whitespace."""
    text = _ARTIFACT_RE.sub("", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _is_junk_alt(text: Optional[str]) -> bool:
    if not text:
        return True
    return text.strip().lower() in _ALT_JUNK or len(text.strip()) <= 2


# ---------------------------------------------------------------------------
# Stage 0: build the Vietnam supplier reference JSON (run offline)
# ---------------------------------------------------------------------------
def build_vietnam_reference(xlsx_path: str,
                            translator: Translator,
                            out_path: str = "vietnam_reference.json",
                            sheet: str = "LE763-Supplier List") -> dict:
    """Read LE763_Supplier_List.xlsx, translate Vietnamese-script names,
    normalize, and write a JSON reference for VietnamSupplierMatcher.

    Columns expected: a Supplier ID column, a Supplier Name column, and
    optionally a 3rd column carrying an alternate name (language not
    assumed -- used generically as an extra alias if present and not junk).
    """
    df = pd.read_excel(xlsx_path, sheet_name=sheet)
    col_id = next(c for c in df.columns if "supplier" in str(c).lower() and "id" in str(c).lower())
    col_name = next(c for c in df.columns if "supplier" in str(c).lower() and "name" in str(c).lower())
    other_cols = [c for c in df.columns if c not in (col_id, col_name)]
    col_alt = other_cols[0] if other_cols else None

    rows = []
    for _, row in df.iterrows():
        if pd.isna(row[col_id]) or pd.isna(row[col_name]):
            continue
        sid = str(row[col_id]).strip()
        raw_name = str(row[col_name]).strip()
        alt_raw = str(row[col_alt]).strip() if col_alt and pd.notna(row[col_alt]) else None
        rows.append((sid, raw_name, alt_raw))

    return _build_reference_from_rows(rows, str(xlsx_path), translator, out_path)


def build_vietnam_reference_from_list(rows: List[Tuple[str, str, Optional[str]]],
                                      translator: Translator,
                                      out_path: str = "vietnam_reference.json") -> dict:
    """Same as build_vietnam_reference but takes an in-memory list of
    (supplier_id, raw_name, alt_name_or_None) tuples -- for local testing
    without an Excel file."""
    return _build_reference_from_rows(rows, "inline_list", translator, out_path)


def _build_reference_from_rows(rows, source_label: str,
                               translator: Translator, out_path: str) -> dict:
    suppliers: List[dict] = []        # one record per supplier: id/VN name/EN name
    alias_entries: List[dict] = []    # flattened aliases, for fuzzy scanning
    norm_to_indices: Dict[str, List[int]] = {}
    translated_count = 0

    for sid, raw_name, alt_raw in rows:
        sid = sid.strip()
        raw_name = _clean_artifact(raw_name)
        if not sid or not raw_name or raw_name.lower() in ("nan", "none", ""):
            continue

        if _is_junk_alt(alt_raw):
            alt_raw = None
        else:
            alt_raw = _clean_artifact(alt_raw)

        # Primary name: special (non-ASCII) chars -> Vietnamese; else English.
        vietnamese_name, english_name = None, None
        if _has_special_chars(raw_name):
            vietnamese_name = raw_name
        else:
            english_name = raw_name

        # Fold in the alt-name column the same way, filling whichever
        # side is still missing.
        if alt_raw:
            if _has_special_chars(alt_raw):
                vietnamese_name = vietnamese_name or alt_raw
            else:
                english_name = english_name or alt_raw

        # Translate to fill in whichever side is still missing.
        if vietnamese_name and not english_name:
            english_name = translator.translate(vietnamese_name)
            translated_count += 1
        # (English-only supplier with no Vietnamese form -- fine as-is.)

        norm_vn = normalize_vn(vietnamese_name) if vietnamese_name else ""
        norm_en = normalize_vn(english_name) if english_name else ""

        suppliers.append({
            "supplier_id": sid,
            "vietnamese_name": vietnamese_name,
            "english_name": english_name,
            "normalized_vietnamese": norm_vn,
            "normalized_english": norm_en,
        })

        # Alias pool: every raw form we saw, never drop one just because
        # the VN/EN classification above was ambiguous.
        alias_candidates = [raw_name]
        if alt_raw and alt_raw not in alias_candidates:
            alias_candidates.append(alt_raw)
        if english_name and english_name not in alias_candidates:
            alias_candidates.append(english_name)
        if vietnamese_name and vietnamese_name not in alias_candidates:
            alias_candidates.append(vietnamese_name)

        for alias in alias_candidates:
            norm = normalize_vn(alias)
            if not norm:
                continue
            idx = len(alias_entries)
            alias_entries.append({"supplier_id": sid, "alias": alias, "normalized": norm})
            norm_to_indices.setdefault(norm, []).append(idx)

    # Collision check: same normalized alias mapping to different supplier
    # IDs -- unresolvable by name alone.
    collisions = []
    for norm_key, idxs in norm_to_indices.items():
        sids = sorted(set(alias_entries[i]["supplier_id"] for i in idxs))
        if len(sids) > 1:
            collisions.append((norm_key, sids))

    ref = {
        "source": source_label,
        "threshold": VN_THRESHOLD,
        "count_suppliers": len(suppliers),
        "count_aliases": len(alias_entries),
        "translated_count": translated_count,
        "collisions": collisions,
        "suppliers": suppliers,      # {supplier_id, vietnamese_name, english_name, ...}
        "entries": alias_entries,    # flattened for fuzzy matching
        "norm_to_indices": norm_to_indices,
    }
    Path(out_path).write_text(json.dumps(ref, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Wrote {out_path}: {len(suppliers)} suppliers, {len(alias_entries)} match aliases "
          f"({translated_count} translated via {type(translator).__name__})")
    if collisions:
        print(f"\n  WARNING: {len(collisions)} normalized name(s) match MULTIPLE "
              f"different supplier IDs -- cannot be disambiguated by name alone:")
        for norm_key, sids in collisions:
            print(f"    '{norm_key}' -> {sids}")
    print()
    return ref


# ---------------------------------------------------------------------------
# Runtime matcher
# ---------------------------------------------------------------------------
class VietnamSupplierMatcher:
    """Match a VLM-extracted supplier_name against the Vietnam reference.

    token_set_ratio (not token_sort_ratio) because an alias-pool entry can
    be a subset or superset of the query's tokens -- token_set_ratio
    scores those near 100 regardless of extra tokens on either side.
    """

    def __init__(self, reference_path: str = "vietnam_reference.json",
                 threshold: int = VN_THRESHOLD):
        ref = json.loads(Path(reference_path).read_text(encoding="utf-8"))
        self.threshold = threshold
        self.entries: List[dict] = ref["entries"]
        self.norm_to_indices: Dict[str, List[int]] = ref["norm_to_indices"]
        self._norm_keys: List[str] = list(self.norm_to_indices.keys())
        self.suppliers_by_id: Dict[str, dict] = {s["supplier_id"]: s for s in ref["suppliers"]}

    def match(self, supplier_name: Optional[str]) -> dict:
        """Direct match -- no translation. Return
        {matched, supplier_id, matched_alias, score, method}."""
        key = normalize_vn(supplier_name)
        if not key:
            return self._result(False, method="empty")

        if key in self.norm_to_indices:
            e = self.entries[self.norm_to_indices[key][0]]
            return self._result(True, e, 100, "exact")

        best_key, best_score = None, 0
        for nk in self._norm_keys:
            score = fuzz.token_set_ratio(key, nk)
            if score > best_score:
                best_score, best_key = score, nk

        if best_key is not None and best_score >= self.threshold:
            e = self.entries[self.norm_to_indices[best_key][0]]
            return self._result(True, e, int(best_score), "fuzzy")

        near = self.entries[self.norm_to_indices[best_key][0]] if best_key else None
        return self._result(False, near, int(best_score), "no_match")

    def match_with_translation(self, supplier_name: Optional[str],
                               translator: Optional[Translator] = None) -> dict:
        """Runtime entry point. Translates ONLY when `supplier_name`
        contains non-ASCII (Vietnamese-script) characters -- plain-ASCII
        input is matched directly, no translation call spent. Pass a
        GoogleTranslateTranslator() instance."""
        if not supplier_name:
            r = self.match(supplier_name)
            r.update(translated=False, translated_query=None)
            return r

        if _has_special_chars(supplier_name) and translator is not None:
            try:
                translated_query = translator.translate(supplier_name)
            except Exception:
                translated_query = supplier_name  # fail open: match on original
            r = self.match(translated_query)
            r["translated"] = True
            r["translated_query"] = translated_query
            return r

        r = self.match(supplier_name)
        r["translated"] = False
        r["translated_query"] = None
        return r

    @staticmethod
    def _result(matched, entry=None, score=0, method="empty") -> dict:
        if entry:
            return {"matched": matched,
                    "supplier_id": entry["supplier_id"] if matched else None,
                    "matched_alias": entry["alias"], "score": score, "method": method}
        return {"matched": matched, "supplier_id": None, "matched_alias": None,
                "score": score, "method": method}


# ---------------------------------------------------------------------------
# File renaming
# ---------------------------------------------------------------------------
def rename_vietnam_files(pdf_path: str, supplier_id: Optional[str],
                         po_number: Optional[str],
                         invoice_number: Optional[str]) -> dict:
    """Rename the PDF and any same-stem XML in the same folder to:
        "ESD{supplier_id} PO{po_number} INV{invoice_number}"

    If supplier_id, po_number, or invoice_number is missing, does NOT
    rename -- returns a reason so the invoice can be flagged for review.
    """
    missing = [n for n, v in
              [("supplier_id", supplier_id), ("po_number", po_number),
               ("invoice_number", invoice_number)] if not v]
    if missing:
        return {"renamed": False, "new_pdf_path": None, "new_xml_path": None,
                "reason": f"missing required field(s): {', '.join(missing)}"}

    pdf_path = Path(pdf_path)
    new_stem = f"ESD{supplier_id} PO{po_number} INV{invoice_number}"
    new_pdf_path = pdf_path.with_name(new_stem + pdf_path.suffix)

    candidate, n = new_pdf_path, 2
    while candidate.exists():
        candidate = pdf_path.with_name(f"{new_stem}_{n}{pdf_path.suffix}")
        n += 1
    new_pdf_path = candidate

    pdf_path.rename(new_pdf_path)

    xml_path = pdf_path.with_suffix(".xml")
    new_xml_path = None
    if xml_path.exists():
        new_xml_path = new_pdf_path.with_suffix(".xml")
        xml_path.rename(new_xml_path)

    return {"renamed": True, "new_pdf_path": str(new_pdf_path),
            "new_xml_path": str(new_xml_path) if new_xml_path else None,
            "reason": None}


def process_vietnam_invoice(pdf_path: str, supplier_name: Optional[str],
                            po_numbers: List[str], invoice_number: Optional[str],
                            matcher: VietnamSupplierMatcher,
                            translator: Optional[Translator] = None) -> dict:
    """Convenience wrapper for main.py: match supplier (translating only
    if needed) -> rename files. Only call when the invoice's LE == 'LE763'.
    `translator` should be a GoogleTranslateTranslator() instance."""
    m = matcher.match_with_translation(supplier_name, translator=translator)
    po_number = po_numbers[0] if po_numbers else None
    rn = rename_vietnam_files(pdf_path, m["supplier_id"], po_number, invoice_number)
    return {"supplier_match": m, "po_number_used": po_number, **rn}


# ---------------------------------------------------------------------------
# Direct run: build from the real Excel (MockTranslator stands in for
# GoogleTranslateTranslator here since this sandbox has no network access
# to translate.google.com) + test matching
# ---------------------------------------------------------------------------
if __name__ == "__main__":

    import pathlib
    XLSX_PATH = pathlib.Path(r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\GenAI-OCR-P2P-Protyping\Supporting Documents\LE763 Supplier List.xlsx")

    translator = GoogleTranslateTranslator()
    build_vietnam_reference(XLSX_PATH, translator)

    print("[2] Loading matcher ...\n")
    matcher = VietnamSupplierMatcher()

    print("[3] Inspect the 5 supplier records that had Vietnamese script:\n")
    for sid in ["1000609384", "1000010820", "1000081356", "1000690978", "1000147137"]:
        s = matcher.suppliers_by_id[sid]
        print(f"  {sid}: VN={s['vietnamese_name']!r}")
        print(f"             EN={s['english_name']!r}")

    print("\n[4] Runtime matching -- match_with_translation() ...\n")
    test_cases = [
        # (query, expect_translation_call, expected_supplier_id, note)
        ("A Chi Son Joint Stock Company", False, "1000052774", "plain ASCII -> match directly"),
        ("FPT IS Company Limited", False, "1000058320", "plain ASCII -> match directly"),
        ("VNPT Th\u00e0nh Ph\u1ed1 H\u1ed3 Ch\u00ed Minh", True, "1000010820", "special chars -> translate first"),
        ("Eastern Region Military-Civilian Hospital", False, "1000081356", "English translation, ASCII -> direct match"),
        ("B\u1ec7nh Vi\u1ec7n Qu\u00e2n D\u00e2n Y Mi\u1ec1n \u0110\u00f4ng", True, "1000081356", "special chars -> translate first"),
        ("MARS Insurance Brokerage Co., Ltd", False, "1000147137", "English translation, ASCII -> direct match"),
        ("Totally Unrelated Company XYZ", False, None, "should not match"),
    ]

    print(f"  {'QUERY':<46} {'TRANSLATED?':<12} {'EXPECT':<12} {'GOT':<12} {'SCORE':>5}  RESULT")
    print(f"  {'-'*46} {'-'*12} {'-'*12} {'-'*12} {'-'*5}  ------")
    for query, expect_translated, expected, note in test_cases:
        r = matcher.match_with_translation(query, translator=translator)
        got = r["supplier_id"] or "NONE"
        exp = expected or "NONE"
        translated_ok = r["translated"] == expect_translated
        match_ok = (got == exp)
        result = "correct" if (translated_ok and match_ok) else "CHECK"
        print(f"  {query:<46} {str(r['translated']):<12} {exp:<12} {got:<12} "
              f"{r['score']:>5}  {result}   ({note})")
        if r["translated"]:
            print(f"      -> translated to: {r['translated_query']!r}")

    # --- Rename demo (temp files only) ---
    print("\n[5] Rename demo (temp files, nothing real is touched) ...\n")
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pdf = tmp / "invoice_scan_0001.pdf"
        xml = tmp / "invoice_scan_0001.xml"
        pdf.write_text("dummy pdf")
        xml.write_text("<dummy/>")

        result = process_vietnam_invoice(
            pdf_path=str(pdf),
            supplier_name="B\u1ec7nh Vi\u1ec7n Qu\u00e2n D\u00e2n Y Mi\u1ec1n \u0110\u00f4ng",
            po_numbers=["7002936777"],
            invoice_number="00009531",
            matcher=matcher,
            translator=translator,
        )
        print(f"  supplier match: {result['supplier_match']}")
        print(f"  renamed:        {result['renamed']}")
        print(f"  new PDF:        {result['new_pdf_path']}")
        print(f"  new XML:        {result['new_xml_path']}")
    print()

import base64
import json
import logging
import re
import time
from io import BytesIO
from pathlib import Path
from typing import Optional, List, Union

from pdf2image import convert_from_path
from pydantic import BaseModel, Field
from langchain_core.messages import HumanMessage

from qwen_llm import invoke_for_schema

# --- Logging ---------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("signals")

# --- Config ------------------------------------------------------------
POPPLER_PATH = r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\GenAI-OCR-P2P-Protyping\prereq-lin\Release-26.02.0-0\poppler-26.02.0\Library\bin"

# No rate-limit pacing needed against a local server; kept as a knob in
# case you ever point this at a shared/slow AI PC.
PER_PAGE_DELAY_SEC = 0.0

_NULL_SENTINELS = {"none", "null", "n/a", "na", "nil", ""}



# ---------------------------------------------------------------------------
# Schema
# ---------------------------------------------------------------------------
class PageSignal(BaseModel):
    supplier_name_on_page: Optional[str] = Field(
        default=None,
        description = (
            "The supplier/issuer company name on THIS page (the seller/'From' "
            "party tied to the logo -- NEVER the Intel 'Bill To' customer, "
            "regardless of how prominent it looks). Prefer a legally-suffixed name "
            "(Inc/Ltd/LP/GmbH/S.A./Co./LLC/Pvt) over a bare brand if both appear "
            "for the same company. For supporting/justification documents, use the "
            "exporter/shipper name instead. Translate non-English/non-Latin names "
            "to English. Null if no clear issuer header is present."
            )
    )
    # Union accepts a bare number -- models sometimes emit an all-digit invoice
    # number as a JSON number, which a strict str schema would reject.
    invoice_number: Union[str, int, None] = Field(
        default=None,
        description=(
            "The invoice / debit-memo / credit-memo document number on this "
            "page (NOT a PO number, order number, or date). none for a "
            "supporting document or if none is present."
        )
    )

    def model_post_init(self, __context) -> None:
        if isinstance(self.invoice_number, int):
            object.__setattr__(self, "invoice_number", str(self.invoice_number))

        if isinstance(self.invoice_number, str) and \
           self.invoice_number.strip().lower() in _NULL_SENTINELS:
            object.__setattr__(self, "invoice_number", None)


# ---------------------------------------------------------------------------
# Prompt (shared by the structured path and the plain-JSON fallback)
# ---------------------------------------------------------------------------
# _SIGNAL_INSTRUCTIONS = (
#     "You are a precise document-extraction system reading ONE page from a "
#     "multi-document PDF that may contain several concatenated invoices plus "
#     "supporting material (shipping docs, customs paperwork, annexures, "
#     "justification, proof documents). The page may be in any language, Latin or non-Latin.\n"
#     "\n"
#     "Extract exactly two fields from THIS page only:\n"
#     "  - supplier_name_on_page\n"
#     "  - invoice_number\n"
#     "\n"
#     "=== STEP 0 -- DOCUMENT TYPE TRIAGE (do this FIRST) ===\n"
#     "Decide which of these THIS page is:\n"
#     "  (a) INVOICE PAGE -- an invoice, debit memo, or credit memo itself.\n"
#     "  (b) SUPPORTING DOCUMENT -- anything that is NOT itself an invoice: "
#     "shipping document, air waybill, bill of lading, customs declaration, "
#     "customs clearance certificate ('justificatif de dedouanement' or "
#     "local-language equivalent), packing list, delivery note, certificate "
#     "of origin, annexure, attachment, justification letter, statement, "
#     "remittance advice, or a supplier-of-supplier's own invoice included "
#     "as backup.\n"
#     "\n"
#     "  STRONG SIGNAL -- printed heading text: if the page carries a "
#     "heading, title, stamp, or watermark reading any of the following (or "
#     "its local-language equivalent), classify it as a SUPPORTING DOCUMENT "
#     "(b), even if the rest of the page is laid out like an invoice:\n"
#     "      'Annexure', 'Annex', 'Appendix', 'Attachment', 'Enclosure', "
#     "'Exhibit', 'Schedule', 'Supporting Document', 'Supporting "
#     "Documentation', 'Backup', 'Proof', 'Proof of Delivery', 'POD', "
#     "'Proof of Payment', 'Justification', 'Justification Document', "
#     "'Justificatif', 'Supplementary', 'For Reference Only', "
#     "'Reference Copy', 'Copy - Not an Invoice'.\n"
#     "  This heading rule OVERRIDES invoice-like appearance: a page that "
#     "looks exactly like an invoice but is headed 'Annexure' or "
#     "'Supporting Document' is still (b), and its invoice_number is still "
#     "null.\n"
#     "\n"
#     "--- IF (b) SUPPORTING DOCUMENT ---\n"
#     "This is the MOST IMPORTANT rule in this prompt. Follow it exactly:\n"
#     "\n"
#     "  1. invoice_number = null. ALWAYS. NO EXCEPTIONS.\n"
#     "     A supporting document NEVER has an invoice number, even when it "
#     "displays a prominent reference number.\n"
#     "     Supporting documents are full of numbers that LOOK like invoice "
#     "numbers. You must return null for ALL of them, including: waybill "
#     "numbers, AWB numbers, tracking numbers, shipment numbers, customs "
#     "declaration numbers, MRN numbers, registration numbers, certificate "
#     "numbers, packing list numbers, delivery note numbers, document "
#     "reference numbers, order numbers, account numbers, and any number "
#     "labeled 'No.', 'Ref', 'Reference', 'Document No', or similar.\n"
#     "     Even if the number is labeled with the word 'Invoice' (e.g. an "
#     "'Invoice Reference' field on a customs form that cites a separate "
#     "invoice), it is still null -- that number belongs to a DIFFERENT "
#     "document, not to this page.\n"
#     "     Do not reason your way to an exception. If the page is a "
#     "supporting document, invoice_number is null.\n"
#     "\n"
#     "  2. supplier_name_on_page = the EXPORTER or SHIPPER name printed on "
#     "the page, exactly as printed (apply the Step 4 language rule below), "
#     "or null if no exporter/shipper name is identifiable.\n"
#     "     Do NOT apply Step 2's Intel-exclusion check or Step 3's "
#     "legal-name reconciliation to the shipper name -- those rules are for "
#     "invoice pages only.\n"
#     "\n"
#     "  3. SKIP Steps 1, 2, 3 and FIELD 2 entirely. Apply Step 4 to the "
#     "name, then output. You are done.\n"
#     "\n"
#     "--- IF (a) INVOICE PAGE ---\n"
#     "Continue to Step 1 below.\n"
#     "\n"
#     "=== FIELD 1 -- supplier_name_on_page (INVOICE PAGES ONLY) ===\n"
#     "The company that ISSUED this page.\n"
#     "\n"
#     "  Step 1 -- Find the issuer candidate:\n"
#     "  * The issuer/seller/'From' party is the text block visually tied to "
#     "the LOGO graphic at the top of the page (below, beside, above, or "
#     "parallel to it).\n"
#     "  * For some Asian-language (e.g. Chinese, Japanese) invoices, if no "
#     "clear top/logo header is present, check near a stamp/seal or at the "
#     "bottom instead.\n"
#     "\n"
#     "  Step 2 -- Exclude the customer, categorically:\n"
#     "  * Every invoice in this dataset is addressed TO an Intel entity. Any "
#     "company name containing 'Intel' is the customer/buyer, NEVER the "
#     "supplier -- exclude it regardless of font size, prominence, or "
#     "position on the page.\n"
#     "  * Also ignore: any 'Bill To' / 'Sold To' / 'Ship To' / 'to' / "
#     "'Buyer' address block, bank names, and freight carriers listed as a "
#     "line item.\n"
#     "  * The customer/Bill-To block is a separate address block, typically "
#     "introduced by an explicit label, positioned below or beside the "
#     "header -- NOT attached to the logo. Use that spatial distinction to "
#     "tell it apart from the issuer block in Step 1.\n"
#     "\n"
#     "  Step 3 -- Resolve the final name (stop at the first rule that "
#     "applies, then commit -- do not re-evaluate):\n"
#     "  1. If a legally-suffixed name (Inc, Ltd, LP, GmbH, S.A., Co., LLC, "
#     "Pvt, etc.) appears anywhere on the page and is clearly the same "
#     "company as the logo brand, use that legal name (e.g. logo says "
#     "'Bloomberg', footer says 'Bloomberg Finance LP' -> use 'Bloomberg "
#     "Finance LP').\n"
#     "  2. Otherwise, use the logo/header brand text as printed.\n"
#     "\n"
#     "  Step 4 -- Language (applies to BOTH the invoice path and the "
#     "supporting-document path in Step 0):\n"
#     "  * If the resolved name is written in a NON-LATIN script (Chinese, "
#     "Japanese, Korean, Cyrillic, Arabic, Thai, etc.) OR in Vietnamese "
#     "(Latin script with diacritics -- e.g. Cong ty TNHH, Cong ty Co phan), "
#     "transliterate/translate it to English so it can be matched downstream. "
#     "Legal-entity suffixes (GmbH, S.A.R.L., B.V., etc.) still stay in their "
#     "original language even when the rest of the name is translated.\n"
#     "  * If the name is already in LATIN script and is English or another "
#     "language you are NOT told to translate (Polish, French, German, "
#     "Spanish, etc.), return it EXACTLY AS PRINTED -- do NOT translate it. "
#     "Preserve legal-suffix words and all accented/special characters "
#     "verbatim.\n"
#     "\n"
#     "  Step 5 -- If no clear issuer/logo header is present on this "
#     "invoice page at all, return null.\n"
#     "\n"
#     "=== FIELD 2 -- invoice_number (INVOICE PAGES ONLY) ===\n"
#     "  * If Step 0 classified this page as a SUPPORTING DOCUMENT, this is "
#     "already null -- do not re-derive it, do not reconsider.\n"
#     "  * On an invoice page: the invoice / debit-memo / credit-memo number, "
#     "usually labeled 'Invoice No', 'Invoice #', 'Document No', 'Debit "
#     "Memo', 'Credit Memo', or the local-language equivalent.\n"
#     "  * This is NOT a Purchase Order (PO) number, order number, account "
#     "number, customer number, tracking number, or date.\n"
#     "  * If this invoice page has no identifiable invoice number, return "
#     "null.\n"
#     "\n"
#     "=== OUTPUT ===\n"
#     "Reason internally, but output ONLY the final answer -- once a field "
#     "resolves, commit to it rather than re-evaluating. Copy names and "
#     "numbers verbatim as printed, keeping accents and non-Latin characters "
#     "exactly (apply the Step 4 language rule only to the returned value, "
#     "not your internal reasoning).\n"
#     "\n"
#     "Before you output: if you classified this page as a supporting "
#     "document in Step 0, verify that invoice_number is null. If you are "
#     "about to return a non-null invoice_number for a supporting document, "
#     "you have made an error -- correct it to null."
# )

# _SIGNAL_INSTRUCTIONS = (
# "You are a precise document-extraction system reading ONE page from a "
# "multi-document PDF. The PDF may contain several concatenated invoices "
# "and supporting material such as shipping documents, customs paperwork, "
# "annexures, justification documents, or proof documents. The page may be "
# "in any language and may use Latin or non-Latin scripts.\n"
# "\n"

# "Extract exactly two fields from THIS PAGE ONLY:\n"
# "  - supplier_name_on_page\n"
# "  - invoice_number\n"
# "\n"

# "=== STEP 0 — CLASSIFY THE PAGE ===\n"
# "\n"
# "First classify THIS page as exactly one of:\n"
# "  (a) INVOICE PAGE — an invoice, debit memo, or credit memo itself.\n"
# "  (b) SUPPORTING DOCUMENT — anything that is not itself an invoice, "
# "including shipping documents, air waybills, bills of lading, customs "
# "declarations, customs-clearance certificates, packing lists, delivery "
# "notes, certificates of origin, annexures, attachments, justification "
# "letters, statements, remittance advice, or a supplier-of-supplier's "
# "invoice included only as backup.\n"
# "\n"

# "A supporting-document heading overrides invoice-like visual appearance. "
# "If the page has a heading, title, stamp, or watermark reading any of "
# "the following, or a clear local-language equivalent, classify it as "
# "SUPPORTING DOCUMENT even if the page otherwise looks like an invoice:\n"
# "  Annexure, Annex, Appendix, Attachment, Enclosure, Exhibit, Schedule, "
# "Supporting Document, Supporting Documentation, Backup, Proof, "
# "Proof of Delivery, POD, Proof of Payment, Justification, "
# "Justification Document, Justificatif, Supplementary, "
# "For Reference Only, Reference Copy, Copy - Not an Invoice.\n"
# "\n"

# "Once the page type is determined, follow only that branch. Do not "
# "override the classification later.\n"
# "\n"

# "=== IF THE PAGE IS A SUPPORTING DOCUMENT ===\n"
# "\n"

# "supplier_name_on_page:\n"
# "  Return the EXPORTER or SHIPPER name printed on this page, applying "
# "the language rules below. Return null if no identifiable exporter or "
# "shipper name is present.\n"
# "\n"

# "invoice_number:\n"
# "  Always return null for a supporting document.\n"
# "  Do not extract or infer any number from the page as an invoice number. "
# "This includes waybill numbers, AWB numbers, tracking numbers, shipment "
# "numbers, customs-declaration numbers, MRN numbers, registration "
# "numbers, certificate numbers, packing-list numbers, delivery-note "
# "numbers, document-reference numbers, order numbers, account numbers, "
# "or numbers labeled No., Ref, Reference, Document No., or similar.\n"
# "  Even if a supporting document contains a number labeled 'Invoice' or "
# "'Invoice Reference', return null because that number refers to a "
# "different document.\n"
# "\n"

# "Do not apply the invoice-page supplier rules to supporting documents. "
# "In particular, do not apply the Intel-customer exclusion or legal-name "
# "reconciliation rules. Apply only the supporting-document exporter/"
# "shipper rule and the language rules, then stop.\n"
# "\n"

# "=== IF THE PAGE IS AN INVOICE PAGE ===\n"
# "\n"

# "FIELD 1 — supplier_name_on_page\n"
# "\n"

# "The supplier is the company that ISSUED this page.\n"
# "\n"

# "1. Find the issuer candidate:\n"
# "   - Prefer the text block visually associated with the LOGO graphic "
# "     at the top of the page. The company name may be below, beside, "
# "     above, or parallel to the logo.\n"
# "   - For some Asian-language invoices, including Chinese or Japanese "
# "     invoices, if there is no clear top/logo header, also check near "
# "     a stamp/seal or at the bottom of the page.\n"
# "\n"

# "2. Exclude the customer:\n"
# "   - Every invoice in this dataset is addressed to an Intel entity. "
# "     Any company name containing 'Intel' is the customer/buyer and "
# "     must never be selected as the supplier.\n"
# "   - Ignore names appearing in Bill To, Sold To, Ship To, to, Buyer, "
# "     or equivalent customer/address blocks.\n"
# "   - Ignore bank names and freight carriers listed as line items.\n"
# "   - The customer/Bill-To block is normally a separate address block "
# "     introduced by an explicit label and positioned below or beside "
# "     the header, rather than being attached to the logo. Use this "
# "     spatial distinction when identifying the issuer.\n"
# "\n"

# "3. Resolve the supplier name:\n"
# "   - If a legally suffixed company name such as Inc, Ltd, LP, GmbH, "
# "     S.A., Co., LLC, Pvt., etc. appears anywhere on the page and is "
# "     clearly the same company as the logo brand, use that legal name.\n"
# "     Example: if the logo says 'Bloomberg' and the footer says "
# "     'Bloomberg Finance LP', return 'Bloomberg Finance LP'.\n"
# "   - Otherwise, use the logo/header brand text as printed.\n"
# "   - If no clear issuer/logo header is present on the invoice page, "
# "     return null.\n"
# "\n"

# "FIELD 2 — invoice_number\n"
# "\n"

# "Extract the invoice, debit-memo, or credit-memo number belonging to "
# "THIS invoice page.\n"
# "\n"

# "Valid examples include values labeled:\n"
# "  Invoice No, Invoice #, Document No, Debit Memo, Credit Memo, "
# "or the equivalent local-language label.\n"
# "\n"

# "Do NOT use a Purchase Order number, order number, account number, "
# "customer number, tracking number, date, or other unrelated reference "
# "number as the invoice number.\n"
# "\n"

# "If the invoice page has no identifiable invoice/debit-memo/credit-memo "
# "number, return null.\n"
# "\n"

# "=== LANGUAGE / NAME NORMALIZATION ===\n"
# "\n"

# "Apply these rules to supplier_name_on_page for both invoice pages and "
# "supporting documents:\n"
# "\n"

# "1. NON-LATIN SCRIPT OR VIETNAMESE:\n"
# "   If the resolved supplier name is written in a non-Latin script "
# "(including Chinese, Japanese, Korean, Cyrillic, Arabic, Thai, etc.) "
# "or is Vietnamese, transliterate/translate it to English so it can be "
# "matched downstream.\n"
# "\n"

# "2. LEGAL SUFFIXES:\n"
# "   Preserve legal-entity suffixes such as GmbH, S.A.R.L., B.V., etc. "
# "in their original form even when the rest of the company name is "
# "translated or transliterated.\n"
# "\n"

# "3. OTHER LATIN-SCRIPT LANGUAGES:\n"
# "   If the supplier name is already written in Latin script and is "
# "English or another language you are not specifically instructed to "
# "translate, such as Polish, French, German, or Spanish, return it "
# "EXACTLY AS PRINTED. Preserve legal-suffix words and all accented or "
# "special characters.\n"
# "\n"

# "=== OUTPUT ===\n"
# "\n"

# "Output ONLY the final answer containing exactly these two fields:\n"
# "\n"
# "{\n"
# "  \"supplier_name_on_page\": \"...\" or null,\n"
# "  \"invoice_number\": \"...\" or null\n"
# "}\n"
# "\n"

# "Do not include explanations, reasoning, document classification, "
# "confidence scores, or additional fields.\n"
# "\n"

# "Copy invoice numbers exactly as printed. Do not translate, normalize, "
# "reformat, or infer invoice numbers.\n"
# "\n"

# "For supplier names, apply the language/normalization rules above. "
# "For Latin-script names that do not require translation, preserve the "
# "printed spelling, accents, and special characters.\n"

# )

_SIGNAL_INSTRUCTIONS = (
"You are a precise document-extraction system reading ONE page from a "
"multi-document PDF. The PDF may contain several concatenated invoices "
"and supporting material such as shipping documents, customs paperwork, "
"annexures, justification documents, or proof documents. The page may be "
"in any language and may use Latin or non-Latin scripts.\n"
"\n"

"Extract exactly two fields from THIS PAGE ONLY:\n"
"  - supplier_name_on_page\n"
"  - invoice_number\n"
"\n"

"=== STEP 0 — CLASSIFY THE PAGE ===\n"
"\n"

"First classify THIS page as exactly one of:\n"
"  (a) INVOICE PAGE — an invoice, debit memo, or credit memo itself.\n"
"  (b) SUPPORTING DOCUMENT — anything that is not itself an invoice, "
"including shipping documents, air waybills, bills of lading, customs "
"declarations, customs-clearance certificates, packing lists, delivery "
"notes, certificates of origin, annexures, attachments, justification "
"letters, statements, remittance advice, or a supplier-of-supplier's "
"invoice included only as backup.\n"
"\n"

"A supporting-document heading overrides invoice-like visual appearance. "
"If the page has a heading, title, stamp, or watermark reading any of "
"the following, or a clear local-language equivalent, classify it as "
"SUPPORTING DOCUMENT even if the page otherwise looks like an invoice:\n"
"  Annexure, Annex, Appendix, Attachment, Enclosure, Exhibit, Schedule, "
"Supporting Document, Supporting Documentation, Backup, Proof, "
"Proof of Delivery, POD, Proof of Payment, Justification, "
"Justification Document, Justificatif, Supplementary, "
"For Reference Only, Reference Copy, Copy - Not an Invoice.\n"
"\n"

"Once the page type is determined, follow only that branch. Do not "
"override the classification later.\n"
"\n"

"=== IF THE PAGE IS A SUPPORTING DOCUMENT ===\n"
"\n"

"supplier_name_on_page:\n"
"  Return the EXPORTER or SHIPPER name printed on this page, applying "
"the language and company-name normalization rules below. Return null "
"if no identifiable exporter or shipper name is present.\n"
"\n"

"invoice_number:\n"
"  Always return null for a supporting document.\n"
"  Do not extract or infer any number from the page as an invoice number. "
"This includes waybill numbers, AWB numbers, tracking numbers, shipment "
"numbers, customs-declaration numbers, MRN numbers, registration "
"numbers, certificate numbers, packing-list numbers, delivery-note "
"numbers, document-reference numbers, order numbers, account numbers, "
"or numbers labeled No., Ref, Reference, Document No., or similar.\n"
"  Even if a supporting document contains a number labeled 'Invoice' or "
"'Invoice Reference', return null because that number refers to a "
"different document.\n"
"\n"

"Do not apply the invoice-page supplier rules to supporting documents. "
"In particular, do not apply the Intel-customer exclusion or legal-name "
"reconciliation rules. Apply only the exporter/shipper rule and the "
"language and company-name normalization rules, then stop.\n"
"\n"

"=== IF THE PAGE IS AN INVOICE PAGE ===\n"
"\n"

"FIELD 1 — supplier_name_on_page\n"
"\n"

"The supplier is the company that ISSUED this page.\n"
"\n"

"1. Find the issuer candidate:\n"
"   - Prefer the text block visually associated with the LOGO graphic "
"     at the top of the page. The company name may be below, beside, "
"     above, or parallel to the logo.\n"
"   - For some Asian-language invoices, including Chinese or Japanese "
"     invoices, if there is no clear top/logo header, also check near "
"     a stamp/seal or at the bottom of the page.\n"
"\n"

"2. Exclude the customer:\n"
"   - Every invoice in this dataset is addressed to an Intel entity. "
"     Any company name containing 'Intel' is the customer/buyer and "
"     must never be selected as the supplier.\n"
"   - Ignore names appearing in Bill To, Sold To, Ship To, to, Buyer, "
"     or equivalent customer/address blocks.\n"
"   - Ignore bank names and freight carriers listed as line items.\n"
"   - The customer/Bill-To block is normally a separate address block "
"     introduced by an explicit label and positioned below or beside "
"     the header, rather than being attached to the logo. Use this "
"     spatial distinction when identifying the issuer.\n"
"\n"

"3. Resolve the supplier name:\n"
"   - If a legally suffixed company name such as Inc, Ltd, LP, GmbH, "
"     S.A., Co., LLC, Pvt., etc. appears anywhere on the page and is "
"     clearly the same company as the logo brand, use that legal name.\n"
"     Example: if the logo says 'Bloomberg' and the footer says "
"     'Bloomberg Finance LP', return 'Bloomberg Finance LP'.\n"
"   - Otherwise, use the logo/header brand text as printed.\n"
"   - If no clear issuer/logo header is present on the invoice page, "
"     return null.\n"
"\n"

"FIELD 2 — invoice_number\n"
"\n"

"Extract the invoice, debit-memo, or credit-memo number belonging to "
"THIS invoice page.\n"
"\n"

"Valid examples include values labeled:\n"
"  Invoice No, Invoice #, Document No, Debit Memo, Credit Memo, "
"or the equivalent local-language label.\n"
"\n"

"Do NOT use a Purchase Order number, order number, account number, "
"customer number, tracking number, date, or other unrelated reference "
"number as the invoice number.\n"
"\n"

"If the invoice page has no identifiable invoice/debit-memo/credit-memo "
"number, return null.\n"
"\n"

"=== LANGUAGE / COMPANY-NAME NORMALIZATION ===\n"
"\n"

"Apply these rules to supplier_name_on_page for both invoice pages and "
"supporting documents.\n"
"\n"

"1. WHEN TO TRANSLATE OR TRANSLITERATE:\n"
"   - If the resolved company name is written in a NON-LATIN script "
"     such as Chinese, Japanese, Korean, Cyrillic, Arabic, Thai, etc., "
"     translate/transliterate it into natural English for downstream "
"     matching.\n"
"   - If the company name is Vietnamese, translate/transliterate it "
"     into natural English for downstream matching, even though "
"     Vietnamese uses Latin script with diacritics.\n"
"   - If the company name is already written in Latin script and is "
"     English or another language that is not specifically instructed "
"     to be translated, such as Polish, French, German, or Spanish, "
"     preserve it EXACTLY AS PRINTED. Do not translate it.\n"
"\n"

"2. TREAT IT AS A COMPANY NAME:\n"
"   - Treat the resolved text as a COMPANY NAME, not as a literal "
"     sentence.\n"
"   - Translate the main business/name meaning naturally rather than "
"     translating each individual word mechanically.\n"
"   - Reorder words when necessary to produce normal English "
"     company-name order.\n"
"   - The translated result should sound like a natural English "
"     company name, not like a word-for-word machine translation.\n"
"\n"

"3. BRANDS AND PROPER NAMES:\n"
"   - Keep established brand names and proper names when there is no "
"     clear English equivalent.\n"
"   - Do not translate a brand or proper name merely because it appears "
"     in a non-English company name.\n"
"   - Do not invent information, expand abbreviations without evidence, "
"     or add words that are not supported by the printed company name.\n"
"\n"

"4. LEGAL ENTITY FORMS:\n"
"   - For a translated or transliterated company name, place the legal "
"     form at the end when this reflects the natural English company-"
"     name structure.\n"
"   - When the original legal form has a clear English equivalent, a "
"     standard English form may be used, such as Co., Ltd., JSC, or LLC.\n"
"   - Preserve legal forms that should remain in their original form, "
"     such as GmbH, S.A.R.L., B.V., etc., rather than replacing them "
"     with an unrelated English legal form.\n"
"   - Do not change the legal form merely to make the company name "
"     sound more English.\n"
"\n"

"5. PRESERVE LATIN-SCRIPT NAMES:\n"
"   - For Latin-script names that do not require translation, preserve "
"     the printed spelling, legal suffix, accents, punctuation, and "
"     special characters exactly.\n"
"\n"

"=== OUTPUT ===\n"
"\n"

"Output ONLY the final answer containing exactly these two fields:\n"
"\n"

"{\n"
"  \"supplier_name_on_page\": \"...\" or null,\n"
"  \"invoice_number\": \"...\" or null\n"
"}\n"
"\n"

"Do not include explanations, reasoning, document classification, "
"confidence scores, or additional fields.\n"
"\n"

"Copy invoice numbers exactly as printed. Do not translate, normalize, "
"reformat, or infer invoice numbers.\n"
"\n"

"For supplier names, apply the language and company-name normalization "
"rules above.\n"
)


_JSON_TAIL = (
    "\n\nRespond with ONLY a single JSON object, no prose and no code fences, "
    "using exactly these two keys:\n"
    '{"supplier_name_on_page": <string or null>, '
    '"invoice_number": <string or null>}\n'
    'Use JSON null (not the string "null") when a field is absent.'
)


# ---------------------------------------------------------------------------
# PDF -> images
# ---------------------------------------------------------------------------
def pdf_to_base64_images(pdf_path: str, dpi: int = 200) -> List[str]:
    log.info("Rasterizing PDF pages to images (dpi=%d)...", dpi)
    pages = convert_from_path(pdf_path, dpi=dpi, poppler_path=POPPLER_PATH)
    b64_images = []
    for page in pages:
        buf = BytesIO()
        page.save(buf, format="PNG")
        b64_images.append(base64.b64encode(buf.getvalue()).decode("utf-8"))
    log.info("Rasterized %d page(s).", len(b64_images))
    return b64_images


# ---------------------------------------------------------------------------
# Message building + page signal
# ---------------------------------------------------------------------------
def _build_message(img_b64: str) -> HumanMessage:
    return HumanMessage(
        content=[
            {"type": "text", "text": _SIGNAL_INSTRUCTIONS + _JSON_TAIL},
            {"type": "image_url",
             "image_url": {"url": f"data:image/png;base64,{img_b64}"}},
        ]
    )


def _get_page_signal(img_b64: str) -> Optional[PageSignal]:
    """One page: prompt for JSON directly and validate against PageSignal.

    (No tool-calling primary path -- see qwen_llm.py's module docstring for
    why. This is the same JSON-prompt-and-validate approach the Groq
    version used only as a fallback, now used directly.)
    """
    return invoke_for_schema(_build_message(img_b64), PageSignal, required_key="supplier_name_on_page")


# ---------------------------------------------------------------------------
# Signal pass
# ---------------------------------------------------------------------------
def get_page_signals(images: List[str],
                     delay_between_pages: float = PER_PAGE_DELAY_SEC) -> List[Optional[PageSignal]]:
    signals: List[Optional[PageSignal]] = []

    log.info("Signal pass: scanning %d page(s) with the local VLM...", len(images))
    for idx, img_b64 in enumerate(images):
        if idx > 0 and delay_between_pages > 0:
            time.sleep(delay_between_pages)

        log.info("  page %d/%d: requesting signal...", idx + 1, len(images))
        sig = _get_page_signal(img_b64)

        if sig is not None:
            log.info("  page %d/%d: supplier=%r  inv_no=%r",
                     idx + 1, len(images), sig.supplier_name_on_page, sig.invoice_number)
        else:
            log.warning("  page %d/%d: signal FAILED -- marking page as unknown",
                        idx + 1, len(images))

        signals.append(sig)

    ok = sum(1 for s in signals if s is not None)
    log.info("Signal pass complete: %d/%d pages read, %d failed.",
             ok, len(images), len(images) - ok)
    return signals


def get_signals_for_pdf(pdf_path: str) -> List[Optional[PageSignal]]:
    """Convenience: PDF path -> list of PageSignal (one per page, None if failed)."""
    images = pdf_to_base64_images(pdf_path)
    return get_page_signals(images)


# ---------------------------------------------------------------------------
# Run standalone
# ---------------------------------------------------------------------------
if __name__ == "__main__":

    pdf_path = r"C:\Users\vmodalax\OneDrive - Intel Corporation\Desktop\GenAI-OCR-P2P-Protyping\sample-invoices\00010836.pdf"

    log.info("=" * 70)
    log.info("SIGNAL MODULE: %s", Path(pdf_path).name)
    log.info("=" * 70)

    signals = get_signals_for_pdf(pdf_path)

    summary = [
        {
            "page": i + 1,
            "supplier_name_on_page": s.supplier_name_on_page if s else None,
            "invoice_number": s.invoice_number if s else None,
            "read_ok": s is not None,
        }
        for i, s in enumerate(signals)
    ]
    print(json.dumps(summary, indent=2, ensure_ascii=False))
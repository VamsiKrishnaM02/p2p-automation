import logging
import re
import time
from typing import Optional, List

from pydantic import BaseModel, Field, field_validator
from langchain_core.messages import HumanMessage

from qwen_llm import invoke_for_schema

log = logging.getLogger("extraction")

# --- Config ------------------------------------------------------------
# No rate-limit pacing needed against a local server; kept as a knob in
# case you ever point this at a shared/slow AI PC.
PER_PAGE_DELAY_SEC = 0.0

# Valid PO: exactly 10 digits with one of these prefixes. Regex is the source
# of truth -- the model returns candidates, this filters them.
PO_PATTERN = re.compile(r'\b(350\d{7}|300\d{7}|5906\d{6}|4506\d{6}|4501\d{6}|4502\d{6}|700\d{7})\b')

_NULL_SENTINELS = {"none", "null", "n/a", "na", "nil", ""}

# ---------------------------------------------------------------------------
# Schema -- ONLY the two fields we still need (supplier/invoice_number reused)
# ---------------------------------------------------------------------------

class ExtractedFields(BaseModel):
    bill_to_client_name: Optional[str] = Field(
        default=None,
        description=(
            "The name of the client/customer being BILLED on this invoice -- "
            "an organisation/entity name found under a heading like 'Client', "
            "'Bill To', 'Bill To Client', 'Billed To', 'Buyer', 'Sold To','to', or "
            "the local-language equivalent. This is an ENTITY NAME, not a "
            "street address. It is NEVER the supplier/vendor name tied to "
            "the logo/header at the top of the page, regardless of "
            "prominence. If a 'Buyer' or 'Sold To' heading is present, "
            "prefer the name under THAT heading over a generic 'Bill To' / "
            "'Client' / 'Billed To' block when they disagree -- it reflects "
            "the actual transacting entity. Never use 'Ship To' / "
            "'Deliver To' under any circumstance. The customer is usually "
            "an Intel entity (e.g. 'Intel Corporation', 'Intel Corporation "
            "SAS', 'Intel Products'), but in rare cases it is McAfee (e.g. "
            "'McAfee LLC', 'McAfee Ireland', 'McAfee Co.') -- read and "
            "return the name exactly as printed rather than defaulting to "
            "Intel. If the printed name is non-English or non-Latin script, "
            "return its ENGLISH translation/transliteration; Null if not "
            "present on this page."
        )
    )
    po_numbers: List[str] = Field(
        default_factory=list,
        description=(
            "All Purchase Order (PO) numbers on this invoice. A valid PO is "
            "EXACTLY 10 digits starting with one of these prefixes: 350 (+7 "
            "digits), 300 (+7), 5906 (+6), 4506 (+6), 4501 (+6), 4502 (+6), "
            "or 700 (+7). Check both the header area and individual line "
            "items -- different line items can reference different PO "
            "numbers, and ALL of them must be returned. Look near labels "
            "like 'PO', 'P.O.', 'PO Number', 'Purchase Order' -- but it may "
            "also be unlabeled, so check anywhere on the page. Do NOT "
            "return invoice numbers, tax IDs, phone numbers, or other digit "
            "sequences that happen to match the length. Return all matches; "
            "empty list if none."
        )
    )

    @field_validator("po_numbers", mode="after")
    @classmethod
    def filter_valid_pos(cls, v: List[str]) -> List[str]:
        # The model may return a PO wrapped with formatting -- "PO#3500510823",
        # "PO: 3500510823", "3500510823 (PO)". Use findall to EXTRACT every
        # valid PO from inside each returned string, rather than fullmatch
        # (which rejects anything that isn't the bare number). Dedupe, keep order.
        out = []
        for x in v:
            for match in PO_PATTERN.findall(str(x)):
                if match not in out:
                    out.append(match)
        return out

    @field_validator("bill_to_client_name", mode="after")
    @classmethod
    def null_sentinel_to_none(cls, v: Optional[str]) -> Optional[str]:
        if v is not None and v.strip().lower() in _NULL_SENTINELS:
            return None
        return v

# _EXTRACT_INSTRUCTIONS = (
#     "You are a precise invoice-extraction system reading ONE page of a "
#     "single invoice. The page may be in any language, Latin or non-Latin. "
#     "Extract exactly two things:\n"
#     "\n"
#     "1. bill_to_client_name -- the entity being BILLED (the customer):\n"
#     "  Step 1 -- Find it under a 'Client' / 'Bill To' / 'Billed To' / "
#     "'Buyer' / 'Sold To'/ 'To'/  heading (or the local-language equivalent). It "
#     "is an entity NAME, not an address -- if the block has a company "
#     "name plus an attention line or department, return only the "
#     "company entity name.\n"
#     "  * The customer is usually an Intel entity, but in rare cases it is "
#     "McAfee (e.g. McAfee LLC, McAfee Ireland, McAfee Co., etc.) -- read "
#     "and return the name exactly as printed rather than defaulting to "
#     "Intel.\n"
#     "  Step 2 -- Exclude the supplier, categorically: the supplier/vendor "
#     "name tied to the logo/header at the top of the page is NEVER the "
#     "bill_to entity, regardless of prominence or position -- exclude it.\n"
#     "  Step 3 -- Multiple headings / multiple Intel entity names: invoices "
#     "sometimes show more than one Intel-related name under different "
#     "headings on the same page (e.g. a generic 'Bill To' AP address "
#     "plus a separate 'Buyer' or 'Sold To' block naming the actual "
#     "transacting entity). Apply this priority order:\n"
#     "    (a) If a 'Buyer' or 'Sold To' heading is present, use the name "
#     "under THAT heading -- it reflects the actual transacting legal "
#     "entity and takes priority over 'Bill To' / 'Client' / 'Billed To' "
#     "when they disagree.\n"
#     "    (b) Otherwise, use 'Bill To' / 'Client' / 'Billed To'.\n"
#     "    (c) NEVER use 'Ship To' / 'Deliver To' / consignee-style "
#     "headings for this field under any circumstance, even if it is the "
#     "only Intel name present on the page.\n"
#     "  Step 4 -- Language: If the name is written in a NON-English (Chinese, Japanese,"
#     " Korean, Cyrillic, Arabic, Thai, Vietnamese etc.), transliterate/translate it to"
#     "English so it can be matched downstream "
#     "Legal-entity suffixes (Sp. z o.o., GmbH, S.A.R.L., B.V., etc.) must be preserved verbatim,"
#     "never translated into English words -- they are part of the"
#     "company's registered legal name, not descriptive text. \n"
#     "  Step 5 -- If no such heading appears on this page, return null.\n"
#     "\n"
#     "2. po_numbers -- every Purchase Order number on the page:\n"
#     "  * A valid PO number is EXACTLY 10 digits, starting with one of "
#     "these prefixes: 350 (+7 digits), 300 (+7), 5906 (+6), 4506 (+6), "
#     "4501 (+6), 4502 (+6), or 700 (+7).\n"
#     "  * Look both in the header area (often labeled 'PO Number', "
#     "'Purchase Order', 'Order No', 'PO#', or 'P.O.') and in individual "
#     "line items -- different line items can reference different PO "
#     "numbers, and ALL of them must be returned.\n"
#     "  * A PO number may also appear unlabeled -- check anywhere on the "
#     "page, not just near an explicit label.\n"
#     "  * Return every match as a string in a list; empty list if none.\n"
#     "  * Do NOT return invoice numbers, tax IDs, or phone numbers, even if "
#     "they happen to be 10 digits.\n"
#     "\n"
#     "Reason internally, but output ONLY the final answer -- once a field "
#     "resolves against the rules above, commit to it rather than "
#     "re-evaluating."
# )

# _EXTRACT_INSTRUCTIONS = (
#     "You are a precise invoice-extraction system reading ONE page of a "
#     "single invoice. The page may be in any language, Latin or non-Latin. "
#     "Extract exactly two things:\n"
#     "\n"
#     "1. bill_to_client_name -- the entity being BILLED (the customer).\n"
#     "\n"
#     "  KEY FACT: every invoice in this dataset is billed to an Intel "
#     "entity (occasionally McAfee, which is part of the same group). The "
#     "customer is ALWAYS one of these. A non-Intel, non-McAfee company on "
#     "the page is the supplier, a carrier, a bank, or a third party -- "
#     "never the answer.\n"
#     "\n"
#     "  Step 1 -- Scan the WHOLE page for any company name containing "
#     "'Intel' (e.g. Intel Corporation, Intel Corporation SAS, Intel "
#     "Products Vietnam, Intel GmbH) or 'McAfee'. Search everywhere, not "
#     "just near obvious headings -- the customer block may be unlabeled, "
#     "in a footer, in a table cell, or in a different language.\n"
#     "\n"
#     "  Step 2 -- If one or more Intel/McAfee names are found, the answer "
#     "is one of them. Choose using this priority:\n"
#     "    (a) the one under a 'Buyer' or 'Sold To' heading;\n"
#     "    (b) else the one under 'Bill To' / 'Billed To' / 'Client' / 'To';\n"
#     "    (c) else the most complete/specific Intel name on the page.\n"
#     "    NEVER pick a name under 'Ship To' / 'Deliver To' / consignee "
#     "headings if any other Intel/McAfee name exists on the page.\n"
#     "    Return the FULL entity name as printed (e.g. 'Intel Corporation "
#     "SAS', not just 'Intel').\n"
#     "\n"
#     "  Step 3 -- ONLY if NO Intel and NO McAfee name appears anywhere on "
#     "the page: fall back to the entity named under 'Bill To' / 'Sold To' / "
#     "'Buyer' / 'Client' / 'To'. If there is no such block either, return "
#     "null.\n"
#     "\n"
#     "  Always exclude: the supplier/vendor name tied to the logo or header "
#     "at the top of the page -- it is never the customer, however "
#     "prominent. Return the company entity NAME only, not the street "
#     "address or attention/department line.\n"
#     "\n"
#     "  Language: if the name is non-English or non-Latin script (Chinese, "
#     "Japanese, Korean, Cyrillic, Arabic, Thai, Vietnamese, etc.), "
#     "translate/transliterate it to English. Keep legal-entity suffixes "
#     "(Sp. z o.o., GmbH, S.A.R.L., B.V., SAS) verbatim -- they are part of "
#     "the registered name, not descriptive text.\n"
#     "  * BUT: if the page already prints an English version of the name -- "
#     "in brackets, on the line below, or beside the local-language name "
#     "(e.g. 'CONG TY TNHH INTEL PRODUCTS VIETNAM (Intel Products Vietnam "
#     "Co., Ltd.)') -- use that printed English name EXACTLY as it appears. "
#     "Do not translate it yourself when the invoice has already done so.\n"
#     "\n"
#     "2. po_numbers -- every Purchase Order number on the page:\n"
#     "  * A valid PO is EXACTLY 10 digits with one of these prefixes: "
#     "350, 300, 700 (+7 digits each), or 5906, 4506, 4501, 4502 (+6 "
#     "digits each).\n"
#     "  * Check the header area AND individual line items -- different "
#     "line items can carry different POs, and ALL must be returned. A PO "
#     "may also be unlabeled, so check anywhere on the page.\n"
#     "  * Return every match as a string in a list; empty list if none.\n"
#     "  * Do NOT return invoice numbers, tax IDs, account numbers, or "
#     "phone numbers, even if they are 10 digits.\n"
#     "\n"
#     "Reason internally, but output ONLY the final answer -- once a field "
#     "resolves, commit to it rather than re-evaluating."
# )

_EXTRACT_INSTRUCTIONS = (
"You are a precise invoice-extraction system. Read ONE PAGE of a "
"single invoice and extract only:\n"
"1. bill_to_client_name — the customer being billed.\n"
"2. po_numbers — every distinct Purchase Order number on THIS PAGE.\n"
"The page may use any language or script.\n"
"\n"

"=== 1. BILL-TO CUSTOMER ===\n"
"\n"
"Every invoice in this dataset is billed to an Intel entity, with "
"occasional McAfee invoices. Therefore, the customer should be an "
"Intel or McAfee entity. Do not select the supplier/vendor, carrier, "
"bank, or other third party.\n"
"\n"

"Scan the ENTIRE PAGE for Intel or McAfee company names, including "
"headers, address blocks, footers, tables, and unlabeled areas.\n"
"\n"

"If Intel/McAfee entities are found, select using this priority:\n"
"1. Buyer or Sold To\n"
"2. Bill To, Billed To, or Client\n"
"3. Most complete/specific Intel or McAfee entity, if clearly the "
"same customer\n"
"\n"

"Prefer the billing/customer entity over shipping information. Do not "
"select an entity under Ship To, Deliver To, or Consignee when another "
"Intel/McAfee entity exists in a billing/customer context.\n"
"\n"

"A standalone 'To' may be used only when its surrounding context "
"clearly identifies it as the billing/customer entity.\n"
"\n"

"Return the FULL entity name, not only the brand. Exclude addresses, "
"contact names, departments, phone numbers, and attention lines.\n"
"\n"

"ONLY if no Intel or McAfee entity is identifiable anywhere on the "
"page, fall back to the entity under Bill To, Billed To, Invoice To, "
"Sold To, Buyer, or Client. Otherwise return null.\n"
"\n"

"=== CUSTOMER NAME NORMALIZATION ===\n"
"\n"
"If the invoice already provides an English version of the customer "
"name in parentheses, brackets, beside, or below the local-language "
"name, use that English version EXACTLY as printed.\n"
"\n"

"Otherwise, for non-Latin or Vietnamese names, translate/transliterate "
"to natural English. Treat the text as a COMPANY NAME, not a literal "
"sentence. Translate the business/name meaning naturally and reorder "
"words into normal English company-name order. Keep established "
"brands and proper names when there is no clear English equivalent. "
"Do not invent or add unsupported information.\n"
"\n"

"Preserve legal forms such as GmbH, S.A.R.L., B.V., SAS, and Sp. z o.o. "
"when they are part of the registered name. When a translated legal "
"form has a clear English equivalent, a standard form such as "
"Co., Ltd., JSC, or LLC may be used. Do not replace a legal form with "
"an unrelated form.\n"
"\n"
"For other Latin-script names that do not require translation, preserve "
"the printed spelling, accents, punctuation, and legal suffix exactly.\n"
"\n"

"=== 2. PO NUMBERS ===\n"
"\n"
"Extract EVERY DISTINCT valid PO number appearing anywhere on THIS PAGE. "
"Check the header, summary sections, tables, and individual line items. "
"Different line items may contain different POs.\n"
"\n"

"A valid PO is EXACTLY 10 digits and matches one of these patterns:\n"
"350XXXXXXX, 300XXXXXXX, 700XXXXXXX, 5906XXXXXX, 4506XXXXXX, "
"4501XXXXXX, or 4502XXXXXX, where X is a digit.\n"
"\n"

"A PO may be unlabeled. However, do not return a number as a PO when "
"there is stronger evidence that it is an invoice number, tax ID, "
"account number, customer number, phone number, tracking number, "
"or another identifier.\n"
"\n"

"Return each PO only once, preserving first-appearance order. If no "
"valid POs are found, return an empty list.\n"
"\n"

"=== OUTPUT ===\n"
"\n"

"Return ONLY valid JSON with exactly these fields:\n"
"{\n"
"  \"bill_to_client_name\": \"...\" or null,\n"
"  \"po_numbers\": [\"...\", \"...\"]\n"
"}\n"
"\n"

"Do not output explanations, reasoning, confidence, or additional fields."

)

_JSON_TAIL = (
    "\n\nRespond with ONLY a single JSON object, no code fences, using exactly "
    "these keys:\n"
    '{"bill_to_client_name": <string or null>, "po_numbers": [<strings>]}\n'
    'Use JSON null when bill_to is absent, and [] when there are no PO numbers.'
)


# ---------------------------------------------------------------------------
# Message building + page extraction
# ---------------------------------------------------------------------------
def _build_message(img_b64: str) -> HumanMessage:
    return HumanMessage(
        content=[
            {"type": "text", "text": _EXTRACT_INSTRUCTIONS + _JSON_TAIL},
            {"type": "image_url",
             "image_url": {"url": f"data:image/png;base64,{img_b64}"}},
        ]
    )


def _extract_page(img_b64: str) -> Optional[ExtractedFields]:
    """One page: prompt for JSON directly and validate against ExtractedFields.

    (No tool-calling primary path here -- the local server doesn't implement
    OpenAI-style function calling, so with_structured_output() would fail on
    every call and just waste a round trip before falling back. This is the
    same JSON-prompt-and-validate approach the Groq version used only as a
    fallback, now used directly. See qwen_llm.invoke_for_schema for the
    retry/parsing logic.)
    """
    return invoke_for_schema(_build_message(img_b64), ExtractedFields, required_key="po_numbers")


# ---------------------------------------------------------------------------
# Merge helper -- never overwrite a found value with a blank; union POs
# ---------------------------------------------------------------------------
def _merge(acc: dict, new: ExtractedFields) -> dict:
    if new.bill_to_client_name and not acc.get("bill_to_client_name"):
        acc["bill_to_client_name"] = new.bill_to_client_name
    for po in new.po_numbers:
        if po not in acc["po_numbers"]:
            acc["po_numbers"].append(po)
    return acc


def _complete(acc: dict) -> bool:
    """Stop once bill_to is found AND at least one valid PO is found.
    (Caller still stops when pages are exhausted; empty PO then is valid.)"""
    return bool(acc.get("bill_to_client_name")) and len(acc["po_numbers"]) > 0


# ---------------------------------------------------------------------------
# Public: extract one invoice
# ---------------------------------------------------------------------------
def extract_invoice(images: List[str],
                    main_pages: List[int],
                    supplier_name: Optional[str],
                    invoice_number: Optional[str],
                    delay_between_pages: float = PER_PAGE_DELAY_SEC) -> dict:
    """
    Extract bill_to + po_numbers for one invoice by looping its MAIN pages.
    supplier_name and invoice_number are REUSED from the signal/segmentation
    step (not re-extracted) and simply included in the returned record.

    `images`     : full list of page images (base64) for the whole PDF.
    `main_pages` : 0-based indices of THIS invoice's main pages.

    Returns a dict:
      {supplier_name, invoice_number, bill_to_client_name, po_numbers,
       needs_manual_review}
    """
    acc = {"bill_to_client_name": None, "po_numbers": []}

    if not main_pages:
        log.info("  invoice has no main pages -- nothing to extract.")
        return {
            "supplier_name": supplier_name,
            "invoice_number": invoice_number,
            "bill_to_client_name": None,
            "po_numbers": [],
            "needs_manual_review": True,
            "review_reasons": [{
                "code": "no_main_pages",
                "message": "This invoice has no main (body) pages to extract from.",
            }],
        }

    for n, page_idx in enumerate(main_pages, start=1):
        if n > 1 and delay_between_pages > 0:
            time.sleep(delay_between_pages)
        log.info("  extracting from main page %d (%d/%d of this invoice)...",
                 page_idx + 1, n, len(main_pages))

        fields = _extract_page(images[page_idx])
        if fields is not None:
            acc = _merge(acc, fields)
            log.info("    so far: bill_to=%r  po=%s",
                     acc["bill_to_client_name"], acc["po_numbers"])
        else:
            log.warning("    page %d extraction failed", page_idx + 1)

        if _complete(acc):
            log.info("  bill_to + PO found -- stopping page loop.")
            break

    # bill_to missing after all pages => flag for review (PO may legitimately
    # be empty, so empty PO alone does NOT trigger review).
    needs_review = not bool(acc.get("bill_to_client_name"))
    reasons = []
    if needs_review:
        checked = [p + 1 for p in main_pages]
        reasons.append({
            "code": "bill_to_not_found",
            "message": f"No 'Bill To' / client name found on any main page "
                      f"(checked page(s) {checked}).",
        })

    return {
        "supplier_name": supplier_name,
        "invoice_number": invoice_number,
        "bill_to_client_name": acc["bill_to_client_name"],
        "po_numbers": acc["po_numbers"],
        "needs_manual_review": needs_review,
        "review_reasons": reasons,
    }

"""
tags.py -- Business-rule tags for a processed invoice: TELECOM, UTILITY, DD.

Tags feed into batch_id.py later, to build the Batch ID suffix
(e.g. "LE420.LW11-TELCO-SENSITIVE"). Stored as a plain list under
invoice["tags"], e.g. invoice["tags"] = ["TELECOM", "UTILITY", "DD"].
Tags are INDEPENDENT -- any combination can apply to the same invoice.

=============================================================================
UTILITY/TELECOM KEYWORD LIST -- PORTED FROM THE OLD SYSTEM (tess_embedding.py)
=============================================================================
The old system ran Tesseract OCR across the full invoice page and searched
that raw OCR text for ~156 keywords/phrases: a 9-word English base list plus
a 21-concept list repeated across 7 languages. We do NOT have full-page OCR
text in this pipeline (a deliberate choice -- see below), so the same word
list is searched against a narrower set of fields instead.

DECISIONS MADE (all confirmed explicitly, not assumed):

1. SEARCH SCOPE -- supplier_name + bill_to_client_name + email subject/body.
   NOT full OCR page text (that would need a new OCR extraction step, which
   was explicitly declined for now). PRACTICAL CONSEQUENCE: the old list's
   long, multi-word descriptive phrases (e.g. the German
   "Wartung von Telekommunikationsgeraeten einschliesslich Nebenstellenanlagen")
   were written to match an invoice's printed line-item text, and are
   unlikely to ever appear verbatim in a supplier name, bill-to name, or
   email subject/body. Only the SHORT single/couple-word items (the 9-word
   English base list, "Telekommunikation" alone, etc.) have a realistic
   chance of matching in this narrower scope. The full phrase lists are
   still included for the (rarer) cases where they DO appear -- e.g. an
   email subject that literally reads "Rechnung - Elektrizitaetsversorgung".

2. CONCEPT LIST -- kept FULL (all 21 concepts), matching the old system
   exactly, NOT filtered down to the ~7 that are genuinely utility/telecom-
   related. This is a deliberate choice with a known, accepted side effect:
   concepts 12-20 (national income tax, VAT, customs violations, tool/
   office-equipment rental, employee lunch vouchers, customer discounts,
   withholding tax, vehicle leasing) are NOT utility-related at all, so an
   invoice merely mentioning e.g. VAT in a language whose phrase list is
   below WILL get tagged UTILITY. See the worked example in this file's
   __main__ block, which demonstrates this concretely rather than leaving
   it as an abstract warning.

3. LANGUAGE COVERAGE -- two of the old system's seven language lists were
   MISLABELED, discovered by reading the actual text:
     - "l3_taiwan" is not Chinese -- it is TWI (a Ghanaian language).
       Kept below as UTILITY_PHRASES_TWI, still labeled honestly.
     - "l6_romania" is not Romanian -- it is ROMANI (the Romani people's
       language, unrelated to Romanian despite the similar name).
       Kept below as UTILITY_PHRASES_ROMANI, still labeled honestly.
   Both Taiwan and Romania are REAL countries in our LE list (LE791/794/796
   Taiwan, LE474 Romania), so neither mislabeled list actually covered its
   claimed country. Per instruction: KEEP both mislabeled lists as extra
   coverage AND add genuine Traditional Chinese (Taiwan) and genuine
   Romanian translations of the same 21-concept list on top -- see
   UTILITY_PHRASES_CHINESE_TW and UTILITY_PHRASES_ROMANIAN. These two new
   lists are original translations produced for this file, not carried
   over from anywhere, and are worth a native-speaker spot-check before
   depending on them heavily in production, the same way any new
   translation would be.
=============================================================================

DD (Direct Debit) detection is UNCHANGED by any of the above -- it still
searches only the email subject/body (see is_direct_debit / derive_tags),
since it's about how the EMAIL communicates payment terms, not about the
invoice's own content.
"""

import re
from pathlib import Path
from typing import List, Optional

import email_reader


# ---------------------------------------------------------------------------
# TELECOM -- our own curated keyword list (separate from the ported Utility
# list below, though the two overlap on "telecom"/"telco" by design -- an
# invoice can pick up BOTH tags, which is fine since tags are independent).
# ---------------------------------------------------------------------------
TELECOM_KEYWORDS = {
    "telecom", "telecoms", "telecommunication", "telecommunications",
    "telco", "mobile", "wireless", "broadband", "cellular",
}

TAG_TELECOM = "TELECOM"
TAG_UTILITY = "UTILITY"
TAG_DD = "DD"


# ---------------------------------------------------------------------------
# UTILITY -- ported from the old system's tess_embedding.py (see module
# docstring for exactly what changed and why).
# ---------------------------------------------------------------------------

# English base list (9 words in the original, +1 here) -- searched as short
# substrings. "gas" added: every OTHER language's phrase list covers gas
# supply (concept #7) as part of its translated phrase, but the English
# base list never got the bare word -- found via testing, looks like an
# oversight in the source list rather than an intentional exclusion.
UTILITY_KEYWORDS_BASE = {
    "utility", "water", "electricity", "power", "internet",
    "broadband", "broad band", "telecom", "telco", "gas",
}

# 21-concept list, Korean. Item order matches the module docstring's
# numbered gloss (1=PBX maintenance ... 21=vehicle leasing benefits).
UTILITY_PHRASES_KOREAN = [
    'PBX를 포함한 통신 장비 유지관리', '통신', '임대', '부동산 서비스',
    '부동산 또는 건물 임대', '상하수도 서비스', '천연가스 공급', '전기 서비스',
    '전화 서비스', '무선 음성 또는 데이터 서비스 또는 휴대폰', '청구 서비스',
    '국민소득세', '소득세 이외의 세금, 재산세', '부가가치세', '관세',
    '공구 및 일반 기계 또는 장비 임대 또는 리스', '사무 장비 및 소모품 임대 또는 리스',
    '점심 식사 쿠폰, 교통비 등 직원 복리후생', '고객 할인', '원천징수세',
    '차량 리스 혜택',
]

# 21-concept list, Japanese.
UTILITY_PHRASES_JAPANESE = [
    'PBXを含む通信機器の保守', '通信', 'リース', '不動産サービス',
    '不動産または建物のリースおよびレンタル', '水道・下水道事業', '天然ガス供給',
    '電気事業', '電話サービス', '無線音声・データサービスまたは携帯電話',
    '課金サービス', '国民所得税', '所得税以外の税金、固定資産税',
    '付加価値税（VAT）', '関税違反', '工具および一般機械・設備のレンタルまたはリース',
    '事務機器および備品のレンタルまたはリース', '昼食券、交通費などの従業員福利厚生',
    '顧客割引', '源泉徴収税', '福利厚生車両リース',
]

# 21-concept list, TWI (Ghanaian language) -- originally mislabeled
# "l3_taiwan" in the source system. Taiwan is a real LE country; this list
# does NOT cover it. Kept per instruction as extra coverage, alongside the
# genuine Chinese list below.
UTILITY_PHRASES_TWI = [
    'Maintenance Telecom Mfiri a PBX ka ho', 'Telecom so nkitahodi', 'Wɔagye atom',
    'Adan ne afie ho adwumayɛ', 'Lease & rental of Agyapadeɛ anaa ɔdan',
    'Nsu ne nsu a ɛkɔ nsu mu a wɔde fa nsu mu', 'Abɔde mu mframa a wɔde ma',
    'Nneɛma a wɔde anyinam ahoɔden di dwuma ', 'telefon so dwumadi',
    'Wireless nne anaa data dwumadie anaa Mobile phone', 'Nnwuma a wɔde tua ho ka',
    'Ɔman Sika a Wonya Tow', 'Tow a ɛnyɛ sika a wonya fi mu tow. Agyapade Ho Towtua',
    'Tow a Wɔde Ka Ho VAT', 'Amanneɛbɔ ho mfomso ahorow.',
    'Nnwinnade ne mfiri anaa nnwinnade a wɔde di dwuma nyinaa a wɔbɛtɔ anaa wɔagye',
    'Office nnwinnade ne nneɛma a wɔde tua ho ka anaasɛ wɔde ma afoforo',
    'Adwumayɛfoɔ mfasoɔ te sɛ awia aduane nkrataa, akwantuo',
    'Adetɔfo Sika a Wɔbɛsan Atua', 'Tow a Wɔde Sie',
    'Mfaso a Ɛwɔ Kar a Wɔde Ma Wɔn Ho',
]

# 21-concept list, German.
UTILITY_PHRASES_GERMAN = [
    'Wartung von Telekommunikationsgeräten einschließlich Nebenstellenanlagen',
    'Telekommunikation', 'Leasing', 'Immobiliendienstleistungen',
    'Leasing und Vermietung von Immobilien oder Gebäuden',
    'Wasser- und Abwasserversorgung', 'Erdgasversorgung', 'Stromversorgung',
    'Telefondienste', 'Drahtlose Sprach- oder Datendienste oder Mobiltelefonie',
    'Abrechnungsdienste', 'Einkommensteuer',
    'Steuern außer der Einkommensteuer. Grundsteuer', 'Mehrwertsteuer',
    'Zolldelikte', 'Miete oder Leasing von Werkzeugen und allgemeinen Maschinen oder Geräten',
    'Miete oder Leasing von Büroausstattung und -bedarf',
    'Mitarbeiterleistungen wie Essensgutscheine, Transport', 'Kundenrabatte',
    'Quellensteuer', 'Vergünstigungen Fahrzeugleasing',
]

# 21-concept list, Czech.
UTILITY_PHRASES_CZECH = [
    'Údržba telekomunikačních zařízení včetně ústředen', 'Telekomunikace',
    'Pronájem', 'Realitní služby', 'Pronájem nemovitostí nebo budov',
    'Vodovody a kanalizace', 'Dodávky zemního plynu', 'Elektřina',
    'Telefonní služby', 'Bezdrátové hlasové nebo datové služby nebo mobilní telefon',
    'Fakturační služby', 'Daň z příjmu',
    'Daně jiné než daň z příjmu. Daň z nemovitosti', 'Daň z přidané hodnoty DPH',
    'Celní přestupky.', 'Pronájem nebo leasing nástrojů a běžných strojů nebo zařízení',
    'Pronájem nebo leasing kancelářského vybavení a potřeb',
    'Zaměstnanecké výhody, jako jsou obědové poukázky, doprava', 'Slevy pro zákazníky',
    'Srážková daň', 'Výhody Leasing vozidel',
]

# 21-concept list, ROMANI (the Romani people's language) -- originally
# mislabeled "l6_romania" in the source system. Romania is a real LE
# country; this list does NOT cover it. Kept per instruction as extra
# coverage, alongside the genuine Romanian list below.
UTILITY_PHRASES_ROMANI = [
    'Telekomunikaciaqo arakhipnasqo arakhipnasqo aparaturǎ inkluziv PBX', 'Telekom',
    'Lesed', 'Servisura pe imobilie', 'Leasing & renting of Proprietat or building',
    'Pani thaj kanalizaciake utilitetura', 'O furniziripe e naturalno gazosko',
    'Elektrikane utilitètură ', 'telefonosko serviso',
    'Bi-telara glasoske vaj datake servisura vaj mobilno telefono',
    'Servisura vaś fakturacia', 'Nacionalno taksa pe pokinipe',
    'Takse aver katar o takse pe inkomesti. Taksa pe Barvalipe',
    'O TVA vaś o thodino valuto', 'Doganesqe ćhinadimata.',
    'Renta vaj lease e Alavengo thaj e generalno mašinengo vaj e aparaturengo',
    'Ofisoske aparatura thaj materialura te kines vaj te les',
    'Beneficije e butjarnenge sar kuponura vash habe, transporto',
    'Rebatură le kliènturenqe', 'O ćhinavipen e takseqo',
    'Le laćhipenata vaś o ćhivipen e maśinenqo',
]

# 21-concept list, Spanish.
UTILITY_PHRASES_SPANISH = [
    'Mantenimiento de equipos de telecomunicaciones, incluyendo centralitas telefónicas (PBX)',
    'Telecomunicaciones', 'Arrendamiento', 'Servicios inmobiliarios',
    'Arrendamiento y alquiler de inmuebles', 'Servicios de agua y alcantarillado',
    'Suministro de gas natural', 'Servicios de electricidad', 'Servicio telefónico',
    'Servicios de voz o datos inalámbricos o telefonía móvil', 'Servicios de facturación',
    'Impuesto Nacional sobre la Renta',
    'Impuestos distintos del impuesto sobre la renta. Impuesto sobre Bienes Inmuebles',
    'Impuesto al Valor Agregado (IVA)', 'Infracciones aduaneras',
    'Arrendamiento o alquiler de herramientas y maquinaria o equipo en general',
    'Arrendamiento o alquiler de equipos y suministros de oficina',
    'Beneficios para empleados como vales de almuerzo y transporte',
    'Descuentos para clientes', 'Retención de impuestos',
    'Beneficios de arrendamiento de vehículos',
]

# 21-concept list, TRADITIONAL CHINESE (Taiwan business register) -- NEW,
# added to give Taiwan genuine coverage (the mislabeled Twi list above never
# did). Original translation for this file; worth a native-speaker
# spot-check before leaning on it heavily.
UTILITY_PHRASES_CHINESE_TW = [
    '包含專用交換機（PBX）的電信設備維護', '電信', '租賃', '不動產服務',
    '不動產或建築物的租賃', '自來水及污水服務', '天然氣供應', '電力服務',
    '電話服務', '無線語音或數據服務或行動電話', '帳單服務', '國民所得稅',
    '所得稅以外之稅捐，財產稅', '加值稅（營業稅）', '海關違規',
    '工具及一般機械或設備租賃', '辦公設備及用品租賃', '員工福利，如午餐券、交通補助',
    '客戶折扣', '扣繳稅款', '車輛租賃福利',
]

# 21-concept list, ROMANIAN -- NEW, added to give Romania genuine coverage
# (the mislabeled Romani list above never did). Original translation for
# this file; worth a native-speaker spot-check before leaning on it heavily.
UTILITY_PHRASES_ROMANIAN = [
    'Întreținerea echipamentelor de telecomunicații, inclusiv centrale telefonice (PBX)',
    'Telecomunicații', 'Leasing', 'Servicii imobiliare',
    'Leasing și închiriere de proprietăți sau clădiri', 'Servicii de apă și canalizare',
    'Furnizare de gaze naturale', 'Servicii de electricitate', 'Servicii telefonice',
    'Servicii de voce sau date wireless sau telefonie mobilă', 'Servicii de facturare',
    'Impozitul național pe venit',
    'Impozite altele decât impozitul pe venit. Impozitul pe proprietate',
    'Taxa pe valoarea adăugată (TVA)', 'Infracțiuni vamale',
    'Închirierea sau leasingul uneltelor și utilajelor sau echipamentelor generale',
    'Închirierea sau leasingul echipamentelor și consumabilelor de birou',
    'Beneficii pentru angajați, cum ar fi tichete de masă, transport',
    'Reduceri pentru clienți', 'Impozitul reținut la sursă', 'Beneficii de leasing auto',
]

# Everything combined into one set for scanning. Lowercased once here so
# _matches_any() doesn't repeatedly lowercase the same ~198 strings on
# every call.
UTILITY_ALL_TERMS = {t.lower() for t in (
    UTILITY_KEYWORDS_BASE
    | set(UTILITY_PHRASES_KOREAN) | set(UTILITY_PHRASES_JAPANESE)
    | set(UTILITY_PHRASES_TWI) | set(UTILITY_PHRASES_GERMAN)
    | set(UTILITY_PHRASES_CZECH) | set(UTILITY_PHRASES_ROMANI)
    | set(UTILITY_PHRASES_SPANISH) | set(UTILITY_PHRASES_CHINESE_TW)
    | set(UTILITY_PHRASES_ROMANIAN)
)}

DD_PHRASES = {
    "direct debit", "DD"
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _normalize(text: str) -> str:
    """Lowercase and collapse whitespace. Works uniformly across scripts:
    CJK has no letter case so .lower() is a harmless no-op there, and plain
    substring containment ("in") doesn't require word boundaries or spaces,
    so it works the same regardless of whether the source script inserts
    spaces between words (Korean does, Chinese/Japanese generally don't)."""
    return re.sub(r"\s+", " ", text.strip().lower())


def _matches_any(text: str, terms) -> bool:
    """True if any term (already-normalized) appears as a substring of the
    normalized text."""
    if not text:
        return False
    norm = _normalize(text)
    return any(term in norm for term in terms)


def _combined_text(*fields: Optional[str]) -> str:
    """Concatenate every available text field into one blob for keyword/
    phrase searching. Empty/None fields contribute nothing."""
    return " ".join(f for f in fields if f)


# ---------------------------------------------------------------------------
# Individual tag rules
# ---------------------------------------------------------------------------
def is_telecom(text: Optional[str]) -> bool:
    """True if `text` (already the combined supplier/bill-to/email blob)
    contains a telecom keyword."""
    return _matches_any(text or "", TELECOM_KEYWORDS)


def is_utility(text: Optional[str]) -> bool:
    """True if `text` (already the combined supplier/bill-to/email blob)
    contains any of the ~198 ported utility terms/phrases, in any of the
    9 languages covered. See the module docstring for what's included and
    the known tax/VAT/benefits false-positive tradeoff this carries."""
    return _matches_any(text or "", UTILITY_ALL_TERMS)


def is_direct_debit(subject: Optional[str] = None,
                    body: Optional[str] = None) -> bool:
    """True if any DD phrase appears in the subject or body text. Scope is
    deliberately narrower than telecom/utility -- DD is about how the EMAIL
    communicates payment terms, not the invoice's own content."""
    text = _combined_text(subject, body)
    return _matches_any(text, DD_PHRASES)


# ---------------------------------------------------------------------------
# Combined entry point
# ---------------------------------------------------------------------------
def derive_tags(supplier_name: Optional[str] = None,
                bill_to_name: Optional[str] = None,
                source_folder: Optional[str] = None) -> List[str]:
    """Compute every tag that applies to one invoice.

    `supplier_name`, `bill_to_name` -- from the invoice's own extracted
    fields. Both feed TELECOM and UTILITY matching.

    `source_folder` -- a path to (or inside) the invoice's originating
    email folder. Subject/body from there ALSO feed TELECOM and UTILITY
    matching (in addition to DD, which only ever uses subject/body). Uses
    email_reader.load_email_metadata(), which walks up to 2 levels looking
    for email_metadata.json, so this works whether the invoice PDF sits
    flat in the message folder or one level down in a per-invoice
    subfolder. Optional: pass None for invoices that didn't come from an
    email at all (a loose-PDF-folder run) -- TELECOM/UTILITY then rely on
    supplier_name/bill_to_name alone, and DD is skipped, not an error.

    Returns tags in a fixed order (TELECOM, UTILITY, DD) when present, so
    output is deterministic regardless of which rules fired. Any
    combination can appear together.
    """
    subject = body = None
    if source_folder:
        meta = email_reader.load_email_metadata(source_folder)
        if meta:
            subject = meta.get("subject")
            body = meta.get("body")

    combined = _combined_text(supplier_name, bill_to_name, subject, body)

    tags: List[str] = []
    if is_telecom(combined):
        tags.append(TAG_TELECOM)
        print("Tagged Telecom")
    if is_utility(combined):
        tags.append(TAG_UTILITY)
        print("Tagged Utility")
    if is_direct_debit(subject, body):
        tags.append(TAG_DD)
        print("Tagged Direct Debit")

    return tags


# ---------------------------------------------------------------------------
# Direct run: sanity checks, including a worked example of the accepted
# tax/VAT false-positive tradeoff from porting the full 21-concept list.
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print(f"Total utility terms loaded: {len(UTILITY_ALL_TERMS)}\n")

    print("=== Straightforward matches ===\n")
    cases = [
        ("Vodafone Telecom Ireland", None, None),
        ("City Water & Power Authority", None, None),
        ("Global Tasks Asesores, S.L.", None, None),
        ("Some German Supplier GmbH", "Rechnung - Telekommunikation", None),
        ("台灣電信股份有限公司", None, None),          # contains 電信 (telecom, Chinese)
        ("SC Exemplu SRL", "Factura - Servicii de electricitate", None),  # Romanian
    ]
    for supplier, bill_to, _ in cases:
        combined = _combined_text(supplier, bill_to)
        tags = []
        if is_telecom(combined): tags.append("TELECOM")
        if is_utility(combined): tags.append("UTILITY")
        print(f"  {str(supplier):<35} {str(bill_to or ''):<35} -> {tags}")

    print("\n=== Accepted tradeoff: full-list porting causes non-utility "
          "false positives ===\n")
    # A German invoice's bill-to block mentioning VAT (Mehrwertsteuer,
    # concept #14) gets wrongly tagged UTILITY -- this is concept #14 from
    # the German list, nothing to do with an actual utility bill. This is
    # the DIRECT, concrete consequence of "full 21-concept list as-is" --
    # shown here so it's not just an abstract warning in a docstring.
    vat_example = "Mehrwertsteuer 19% ausgewiesen"
    print(f"  German invoice text: {vat_example!r}")
    print(f"  is_utility() -> {is_utility(vat_example)}  "
          f"(True = false positive, matched VAT concept #14, not an actual utility)")

    print("\n=== DD unaffected by any of the above (subject/body only) ===\n")
    print("  ", is_direct_debit(subject="Your Direct Debit is due"))
    print("  ", is_direct_debit(subject="Mehrwertsteuer 19%"), "(no DD phrase -> False)")

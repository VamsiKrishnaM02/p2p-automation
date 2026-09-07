"""
qwen_llm.py

Shared LangChain client + retry/parsing helpers for the local Qwen VLM
server, used in place of ChatGroq across extraction.py and signals.py.

The local server (main.py) is a plain OpenAI-compatible /v1/chat/completions
endpoint -- it does NOT implement tool/function calling (no `tools` field
handling), so LangChain's with_structured_output() can't work against it:
it would send a `tools` payload the server silently ignores, get back a
response with no tool_calls, and fail every time.

Instead we always prompt for JSON directly in the message text and
validate the result against a Pydantic schema client-side. This is exactly
the "fallback" path both original Groq-based modules already had -- it's
just promoted here to the only path, since the "primary" structured path
would never succeed against this server anyway.
"""

import json
import re
import time
import logging
from typing import Optional, Type, TypeVar

from openai import RateLimitError, InternalServerError, APIConnectionError
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage
from pydantic import BaseModel, ValidationError

log = logging.getLogger("qwen_llm")

# --- Server config -----------------------------------------------------
# Confirm this matches your colleague's actual server IP (client.py and
# client_setup.ps1 disagreed on this earlier -- verify before trusting it).
AI_PC_IP = "10.211.180.60"
BASE_URL = f"http://{AI_PC_IP}:8000/v1"
MODEL = "Qwen/Qwen3.5-4B"

# Same error categories the Groq setup retried on transient failures.
# The openai SDK ships these under the same names.
TRANSIENT_ERRORS = (RateLimitError, InternalServerError, APIConnectionError)

T = TypeVar("T", bound=BaseModel)


def get_llm(max_tokens: int = 4096, temperature: float = 0.0) -> ChatOpenAI:
    """Equivalent of the  ChatGroq(model=..., reasoning_format='hidden') constructor."""
    return ChatOpenAI(
        model=MODEL,
        base_url=BASE_URL,
        api_key="EMPTY",  # server enforces no auth; SDK requires a non-empty string
        temperature=temperature,
        max_tokens=max_tokens,
        # Qwen's thinking mode is the nearest equivalent to Groq's
        # reasoning_format='hidden' -- turn it off so we get a direct answer
        # instead of a <think>...</think> block ahead of the JSON.
        extra_body={"chat_template_kwargs": {"enable_thinking": False}},
    )


def extract_json_object(text: Optional[str], required_key: Optional[str] = None) -> Optional[dict]:
    """Pull a JSON object out of a plain-text response, robust to models that
    emit reasoning first or wrap the answer in code fences.

    Identical logic to what both original modules had duplicated --
    centralized here. If `required_key` is given, prefers the last
    candidate object that contains that key (the committed final answer,
    not an illustrative example shown mid-reasoning).
    """
    if not text:
        return None

    cleaned = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL)
    cleaned = re.sub(r"```(?:json)?", "", cleaned).replace("`", "").strip()

    candidates = []
    depth = 0
    start = None
    for i, ch in enumerate(cleaned):
        if ch == "{":
            if depth == 0:
                start = i
            depth += 1
        elif ch == "}":
            if depth > 0:
                depth -= 1
                if depth == 0 and start is not None:
                    try:
                        candidates.append(json.loads(cleaned[start:i + 1]))
                    except json.JSONDecodeError:
                        pass
                    start = None

    if not candidates:
        return None
    if required_key:
        for cand in reversed(candidates):
            if isinstance(cand, dict) and required_key in cand:
                return cand
    return candidates[-1] if isinstance(candidates[-1], dict) else None


def invoke_for_schema(
    message,
    schema: Type[T],
    required_key: Optional[str] = None,
    retries: int = 3,
) -> Optional[T]:
    """Send one LangChain message to the local VLM, parse + validate the JSON
    response against `schema`.

    Retries on transient connection errors only -- a malformed JSON shape
    won't fix itself by repeating an identical request, so validation
    failures return None immediately rather than retrying.
    """
    llm = get_llm()
    raw = None
    for attempt in range(1, retries + 1):
        try:
            resp = llm.invoke([message])
            raw = resp.content if isinstance(resp.content, str) else str(resp.content)
            break
        except TRANSIENT_ERRORS as e:
            log.warning("attempt %d/%d transient error (%s: %s)", attempt, retries, type(e).__name__, e)
            time.sleep(2.0 * attempt)
        except Exception as e:
            log.warning("call failed (%s: %s)", type(e).__name__, e)
            return None

    if raw is not None:
        log.info("RAW response:\n%s", raw)

    data = extract_json_object(raw, required_key=required_key)
    if data is None:
        return None
    try:
        return schema(**data)
    except ValidationError as e:
        log.warning("validation failed: %s", e)
        return None


# ---------------------------------------------------------------------------
# Plain-text translation (no image, no JSON schema) -- used for the
# translate-and-retry step in the sensitive/LE matchers.
# ---------------------------------------------------------------------------
_TRANSLATION_PREFIX_RE = re.compile(
    r"^(translation|translated( text| name)?|english( translation)?)\s*[:\-]\s*",
    re.IGNORECASE,
)


def _clean_translation_output(raw: Optional[str], fallback: str) -> str:
    """Strip whatever wrapper text a model adds despite being told not to --
    a leading 'Translation:' label, surrounding quotes, or stray whitespace.
    Falls back to the original text if nothing usable is left."""
    if not raw:
        return fallback

    cleaned = re.sub(r"<think>.*?</think>", "", raw, flags=re.DOTALL).strip()
    cleaned = _TRANSLATION_PREFIX_RE.sub("", cleaned).strip()

    if len(cleaned) >= 2 and cleaned[0] in "\"'" and cleaned[-1] == cleaned[0]:
        cleaned = cleaned[1:-1].strip()

    # Some models echo a trailing period on a short name that didn't have one.
    if cleaned.endswith(".") and not fallback.endswith("."):
        cleaned = cleaned.rstrip(".").strip()

    return cleaned or fallback


def translate_to_english(text: str, context: str = "company name", retries: int = 2) -> str:
    """
    Translate `text` to English using the local VLM, text-only (no image).
    Returns ONLY the translated string -- no explanation, no quotes, no JSON.

    Legal-entity suffixes (GmbH, S.A.R.L., B.V., Ltd, etc.) are instructed
    to stay untranslated, matching the extraction-prompt behavior elsewhere
    in this pipeline, since the LE matcher depends on those exact tokens.

    On any failure (transient error, empty response, etc.) this returns
    `text` unchanged rather than raising -- a failed translation should
    mean "no better than what we started with" for a caller doing
    translate-and-retry matching, not an exception that kills the batch.
    """
    if not text or not text.strip():
        return text

    prompt = (
        "Convert this company name into a natural English company name.\n\n"
        "Rules:\n"
        "- Treat it as a company name, not a literal sentence.\n"
        "- Translate the main business/name meaning naturally.\n"
        "- Reorder words into normal English company-name order.\n"
        "- Keep brand/proper names when there is no clear English equivalent.\n"
        "- Put the legal form at the end using a standard English form such as Co., Ltd., JSC, LLC.\n"
        "- Do not invent information.\n\n"
        
        "Example:\n"
        "CÔNG TY TNHH MỘT THÀNH VIÊN KỸ THUẬT LƯU TRỮ Á CHÂU\n"
        "-> Asia Storage Technology Co., Ltd.\n\n"

        "FINAL CHECK:\n"
        "The output must contain NO text in the source language. "
        "Do not leave Vietnamese diacritics, Chinese/Japanese/Korean characters, "
        "Cyrillic, Arabic, Thai, Devanagari, or accented characters from the source. "
        "Translate or transliterate them before answering.\n\n"

        "Return ONLY the final English company name. "
        "No quotation marks, explanation, or labels.\n\n"

        f"Company name: {text}"
    )

    llm = get_llm(max_tokens=128, temperature=0.0)
    raw = None
    for attempt in range(1, retries + 1):
        try:
            resp = llm.invoke([HumanMessage(content=prompt)])
            raw = resp.content if isinstance(resp.content, str) else str(resp.content)
            break
        except TRANSIENT_ERRORS as e:
            log.warning("translate attempt %d/%d transient error (%s: %s)",
                        attempt, retries, type(e).__name__, e)
            time.sleep(2.0 * attempt)
        except Exception as e:
            log.warning("translate call failed for %r: %s", text, e)
            return text

    result = _clean_translation_output(raw, fallback=text)
    log.info("translated %r -> %r", text, result)
    return result
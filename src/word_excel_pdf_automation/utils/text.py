from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher


_NON_ALNUM = re.compile(r"[^a-z0-9]+")


def normalize_for_match(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.lower().strip()
    return _NON_ALNUM.sub(" ", value).strip()


def normalize_series(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    return str(value).strip()


def sanitize_filename(value: str, fallback: str = "document") -> str:
    value = normalize_for_match(value)
    value = value.replace(" ", "_")
    value = re.sub(r"[^a-z0-9_\-]+", "", value)
    value = re.sub(r"_+", "_", value).strip("._ ")
    return value or fallback


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, normalize_for_match(a), normalize_for_match(b)).ratio()


def dedupe_preserve_order(items: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result

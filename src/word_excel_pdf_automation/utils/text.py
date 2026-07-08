from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path


_NON_ALNUM = re.compile(r"[^a-z0-9]+")
_WINDOWS_INVALID_FILENAME_CHARS = re.compile(r"[<>:\"/\\|?*\x00-\x1f]")
_WINDOWS_RESERVED_NAMES = {
    "con",
    "prn",
    "aux",
    "nul",
    *{f"com{i}" for i in range(1, 10)},
    *{f"lpt{i}" for i in range(1, 10)},
}


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


def sanitize_windows_filename(value: str, fallback: str = "document") -> str:
    value = (value or "").replace("\r", " ").replace("\n", " ").strip()
    value = _WINDOWS_INVALID_FILENAME_CHARS.sub("", value)
    value = re.sub(r"\s+", " ", value).strip(" .")
    if not value:
        return fallback

    stem = value.split(".")[0].rstrip(" .").lower()
    if stem in _WINDOWS_RESERVED_NAMES:
        value = f"{value}_"
    return value or fallback


def build_output_stem(template_path: str | Path, serie: str) -> str:
    template_stem = Path(template_path).stem
    cleaned_series = sanitize_windows_filename(normalize_series(serie), fallback="serie")
    marker = "__SERIE_PLACEHOLDER__"

    replaced = re.sub(r"\[serie\]", marker, template_stem, flags=re.IGNORECASE)
    replaced = re.sub(r"\bserie\b", marker, replaced, flags=re.IGNORECASE)
    replaced = replaced.replace(marker, cleaned_series)
    if replaced == template_stem:
        replaced = f"{template_stem} - {cleaned_series}" if cleaned_series else template_stem

    return sanitize_windows_filename(replaced, fallback=cleaned_series or "document")


def build_output_filename(template_path: str | Path, serie: str, extension: str = "") -> str:
    stem = build_output_stem(template_path, serie)
    suffix = extension.lstrip(".")
    return f"{stem}.{suffix}" if suffix else stem


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

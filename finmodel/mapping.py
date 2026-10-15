"""Explicit header aliases; never fuzzy-match financial definitions."""

from .config import METADATA, NUMERIC_FIELDS
from .schemas import ModelError


def header_key(value: str) -> str:
    return value.strip().lower().replace(" ", "_").replace("-", "_")


def alias_lookup(aliases: dict) -> dict[str, str]:
    lookup = {}
    for canonical, names in aliases.items():
        if canonical not in METADATA + NUMERIC_FIELDS:
            raise ModelError(f"Unknown canonical field in aliases: {canonical}")
        if not isinstance(names, list) or not all(isinstance(n, str) for n in names):
            raise ModelError(f"Aliases for {canonical} must be a list of strings")
        for name in [canonical, *names]:
            key = header_key(name)
            if key in lookup and lookup[key] != canonical:
                raise ModelError(f"Ambiguous alias {name!r}: {lookup[key]} / {canonical}")
            lookup[key] = canonical
    return lookup


def map_headers(headers: list, lookup: dict) -> tuple[dict, list]:
    mapped, unknown = {}, []
    for index, raw in enumerate(headers):
        if not isinstance(raw, str) or not raw.strip():
            raise ModelError(f"Empty/non-text column header at column {index + 1}")
        canonical = lookup.get(header_key(raw))
        if canonical is None:
            unknown.append(raw)
        elif canonical in mapped.values():
            raise ModelError(f"Multiple columns map to {canonical}; remove the ambiguity")
        else:
            mapped[index] = canonical
    missing = set(METADATA) - set(mapped.values())
    if missing:
        raise ModelError(f"Missing required columns: {', '.join(sorted(missing))}")
    return mapped, unknown

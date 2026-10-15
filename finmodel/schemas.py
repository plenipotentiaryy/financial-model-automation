"""Data contracts; exact normalized values remain Decimal until Excel export."""

from dataclasses import asdict, dataclass, field
from datetime import date, time
from decimal import Decimal
import math
from typing import Any


class ModelError(Exception):
    """Expected input, configuration or export failure."""


@dataclass
class Issue:
    code: str
    message: str
    severity: str = "warning"
    company_id: str | None = None
    metric: str | None = None
    source: str | None = None


@dataclass
class Source:
    source_file: str
    source_sheet: str | None
    source_row: int
    source_column: str
    source_locator: str
    original_value: Any
    original_unit: str | None = None
    candidate_id: str | None = None
    revision: int | None = None
    source_sha256: str | None = None


@dataclass
class RawRow:
    values: dict[str, Any]
    sources: dict[str, Source]


@dataclass
class Observation:
    company_id: str
    metric: str
    fiscal_year: int
    period_type: str
    scope: str
    currency: str
    canonical_value: Decimal | None
    canonical_unit: str
    original_unit: str
    value_status: str
    measurement_basis: str
    sources: list[Source]
    derivation: dict | None = None


@dataclass
class Company:
    company_id: str
    company_name: str = ""
    currency: str = ""
    scope: str = ""
    years: set[int] = field(default_factory=set)
    observations: dict[tuple[str, int], Observation] = field(default_factory=dict)
    issues: list[Issue] = field(default_factory=list)
    metric_statuses: dict[str, dict[str, str]] = field(default_factory=dict)
    source_files: set[str] = field(default_factory=set)
    review_audit: dict | None = None

    @property
    def latest_year(self) -> int | None:
        return max(self.years) if self.years else None

    def period_year(self, period: str) -> int | None:
        if not self.years:
            return None
        if period == "start":
            return min(self.years)
        if period == "previous":
            return self.latest_year - 1
        if period == "latest":
            return self.latest_year
        raise ModelError(f"Unknown fiscal period role: {period}")

    @property
    def blocked(self) -> bool:
        return any(i.severity == "error" for i in self.issues)

    def value(self, metric: str, year: int | None = None) -> Decimal | None:
        obs = self.observations.get((metric, year if year is not None else self.latest_year))
        return obs.canonical_value if obs and obs.value_status in {"reported", "derived_approximate"} else None


@dataclass
class Dataset:
    companies: dict[str, Company] = field(default_factory=dict)
    issues: list[Issue] = field(default_factory=list)


def serializable(value: Any) -> Any:
    if hasattr(value, "__dataclass_fields__"):
        return serializable(asdict(value))
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (date, time)):
        return value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)
    if isinstance(value, dict):
        return {str(k): serializable(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [serializable(v) for v in value]
    return value

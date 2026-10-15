"""Decimal-safe normalization and company/period compatibility enforcement."""

import re
from decimal import Decimal, InvalidOperation, localcontext

from .config import BALANCES, NUMERIC_FIELDS, RATES
from .schemas import Company, Dataset, Issue, ModelError, Observation

SCALES = {"units": Decimal(1), "thousands": Decimal(1000), "millions": Decimal(1000000)}
STATES = {"": "missing", "missing": "missing", "n/a": "not_applicable", "not applicable": "not_applicable", "not_applicable": "not_applicable", "unavailable": "unavailable", "invalid": "invalid"}


def parse_number(raw) -> tuple[Decimal | None, str]:
    if raw is None:
        return None, "missing"
    if isinstance(raw, bool) or not isinstance(raw, (str, int, float, Decimal)):
        raise ModelError(f"Invalid numeric input: {raw!r}")
    text = str(raw).strip()
    if text.lower() in STATES:
        return None, STATES[text.lower()]
    if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?", text):
        raise ModelError(f"Invalid number {raw!r}; use a decimal point and no grouping separators")
    try:
        value = Decimal(text)
    except InvalidOperation as exc:
        raise ModelError(f"Invalid number {raw!r}") from exc
    if not value.is_finite() or abs(value.adjusted()) > 300:
        raise ModelError(f"Nonfinite/out-of-range number: {raw!r}")
    return value, "reported"


def normalize(rows, initial_issues=None) -> Dataset:
    dataset = Dataset(issues=list(initial_issues or []))
    for row in rows:
        data = row.values
        raw_id = data.get("company_id")
        if raw_id is None or isinstance(raw_id, (dict, list, bool)) or not str(raw_id).strip():
            dataset.issues.append(Issue("INVALID_IDENTITY", "Missing/invalid company_id", "error", source=next(iter(row.sources.values())).source_file))
            continue
        company_id = str(raw_id).strip()
        company = dataset.companies.setdefault(company_id, Company(company_id))
        source_file = next(iter(row.sources.values())).source_file
        company.source_files.add(source_file)

        def error(code, message, metric=None):
            company.issues.append(Issue(code, message, "error", company_id, metric, source_file))

        metadata = {}
        for field in ("company_name", "period_type", "scope", "currency", "monetary_unit", "fiscal_year"):
            raw = data.get(field)
            if raw is None or isinstance(raw, (dict, list, bool)) or not str(raw).strip() or (field != "company_name" and str(raw).startswith("=")):
                error("INVALID_METADATA", f"Missing/invalid {field}")
            else:
                metadata[field] = str(raw).strip()
        if len(metadata) != 6:
            continue
        name = metadata["company_name"]
        if company.company_name and company.company_name != name:
            error("INVALID_IDENTITY", "Conflicting company names for the same company_id")
        company.company_name = company.company_name or name
        year_text = metadata["fiscal_year"].upper()
        if not re.fullmatch(r"(?:FY\s*)?\d{4}", year_text):
            error("INVALID_PERIOD", f"Invalid fiscal year: {year_text}")
            continue
        year = int(re.sub(r"^FY\s*", "", year_text))
        company.years.add(year)
        period = metadata["period_type"].upper()
        scope = metadata["scope"].lower()
        currency = metadata["currency"].upper()
        unit = metadata["monetary_unit"].lower()
        if period != "FY":
            error("INCOMPATIBLE_PERIODS", f"MVP accepts FY rows, received {period}")
        if scope not in {"consolidated", "standalone"}:
            error("INVALID_SCOPE", f"Unsupported scope: {scope}")
        if not re.fullmatch(r"[A-Z]{3}", currency):
            error("INVALID_CURRENCY", f"Currency must be a three-letter code: {currency}")
        if company.currency and company.currency != currency:
            error("CURRENCY_MISMATCH", f"Cannot mix {company.currency} and {currency}")
        if company.scope and company.scope != scope:
            error("SCOPE_MISMATCH", f"Cannot mix {company.scope} and {scope}")
        company.currency = company.currency or currency
        company.scope = company.scope or scope
        if unit not in SCALES:
            error("INVALID_UNIT", f"Unsupported monetary_unit: {unit}")
            continue
        for field in NUMERIC_FIELDS:
            if field not in data:
                continue
            raw = data[field]
            try:
                value, status = parse_number(raw)
            except ModelError as exc:
                error("INVALID_NUMBER", f"{field} FY{year}: {exc}", field)
                value, status = None, "invalid"
            if status == "invalid" and not any(i.metric == field and i.code == "INVALID_NUMBER" for i in company.issues):
                error("INVALID_NUMBER", f"{field} FY{year} explicitly marked invalid", field)
            is_money = field not in RATES | {"equity_beta", "shares_outstanding", "share_price"}
            if value is not None and is_money:
                with localcontext() as context:
                    context.prec = max(40, len(value.as_tuple().digits) + 10)
                    value *= SCALES[unit]
            canonical_unit = currency if is_money else ("ratio" if field in RATES | {"equity_beta"} else ("count" if field == "shares_outstanding" else f"{currency}/share"))
            original_unit = unit if is_money else canonical_unit
            row.sources[field].original_unit = original_unit
            basis = "fiscal_year_end" if field in BALANCES else ("annual_flow" if is_money else "explicit_assumption")
            if field == "monthly_recurring_revenue":
                basis = "monthly_run_rate_at_fiscal_year_end"
            observation = Observation(company_id, field, year, period, scope, currency, value, canonical_unit, original_unit, status, basis, [row.sources[field]])
            key = (field, year)
            previous = company.observations.get(key)
            if previous is not None:
                signature = lambda o: (o.canonical_value, o.value_status, o.period_type, o.scope, o.currency)
                if signature(previous) != signature(observation):
                    error("AMBIGUOUS_SOURCE", f"Conflicting duplicate {field} FY{year}", field)
                else:
                    # Each source retains its own original value and monetary scale.
                    previous.sources.extend(observation.sources)
            else:
                company.observations[key] = observation
    return dataset

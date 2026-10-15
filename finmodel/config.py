"""Small editable configuration with no executable configuration expressions."""

from copy import deepcopy
from pathlib import Path

import yaml

from .schemas import ModelError

METADATA = ["company_id", "company_name", "fiscal_year", "period_type", "scope", "currency", "monetary_unit"]
CORE_FIELDS = ["revenue", "ebitda", "total_debt", "cash_and_equivalents"]
BALANCES = {"total_debt", "cash_and_equivalents", "total_assets", "total_equity", "current_assets", "current_liabilities", "inventory", "accounts_receivable", "accounts_payable", "market_cap", "market_value_debt"}
RATES = {"pre_tax_cost_of_debt", "risk_free_rate", "equity_risk_premium", "marginal_tax_rate"}
OPTIONAL_FIELDS = ["cogs", "gross_profit", "ebit", "net_income", "interest_expense", "income_tax_expense", "cash_flow_from_operations", "capex", "dividends_paid", "total_assets", "total_equity", "current_assets", "current_liabilities", "inventory", "accounts_receivable", "accounts_payable", "market_cap", "share_price", "shares_outstanding", "pre_tax_cost_of_debt", "risk_free_rate", "equity_beta", "equity_risk_premium", "marginal_tax_rate", "monthly_recurring_revenue", "market_value_debt"]
NUMERIC_FIELDS = CORE_FIELDS + OPTIONAL_FIELDS

ALIASES = {k: [k] for k in METADATA + NUMERIC_FIELDS}
ALIASES["revenue"] += ["sales", "turnover"]
ALIASES["company_name"] += ["name"]
ALIASES["fiscal_year"] += ["year", "fy"]
ALIASES["cash_and_equivalents"] += ["cash", "cash_equivalents"]
ALIASES["total_debt"] += ["debt"]

ASSUMPTION_FIELDS = ["market_value_debt", "pre_tax_cost_of_debt", "risk_free_rate", "equity_beta", "equity_risk_premium", "marginal_tax_rate"]
FINANCIAL_FIELDS = [field for field in NUMERIC_FIELDS if field not in ASSUMPTION_FIELDS]
TEMPLATE_MAP = {
    "version": 2,
    "metadata": {"company_name": "Inputs!B2", "company_id": "Inputs!B5", "latest_fiscal_year": "Inputs!C3", "previous_fiscal_year": "Inputs!B3", "first_fiscal_year": "Inputs!D3", "currency": "Inputs!C4", "scope": "Inputs!C5"},
    "financials": {field: {"previous": f"Inputs!B{row}", "latest": f"Inputs!C{row}"} for row, field in enumerate(FINANCIAL_FIELDS, 7)},
    "assumptions": {field: {"latest": f"Assumptions!B{row}"} for row, field in enumerate(ASSUMPTION_FIELDS, 5)},
    "regions": {"checks": "Checks!A5:F1004", "source_map": "Source_Map!A5:Q10004", "metric_statuses": "Report!D7:E34", "run_status": "Report!B38:B38"},
}
TEMPLATE_MAP["financials"]["revenue"]["start"] = "Inputs!D7"
TEMPLATE_MAP["financials"]["ebitda"]["start"] = "Inputs!D8"

from .catalog import build_catalog

METRICS = build_catalog(TEMPLATE_MAP)
METRIC_ORDER = tuple(METRICS)
REPORT_END_ROW = 6 + len(METRICS)


def load_config(path: str | Path | None, default: dict) -> dict:
    if path is None:
        return deepcopy(default)
    try:
        value = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
    except (OSError, yaml.YAMLError) as exc:
        raise ModelError(f"Cannot read configuration {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise ModelError(f"Configuration must be a mapping: {path}")
    return value

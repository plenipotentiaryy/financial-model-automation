"""Prerequisite checks shared with the catalog's Excel guards; no KPI engine."""

from .config import METRICS, METRIC_ORDER
from .schemas import Issue


def rule_failure(company, rule):
    kind, fields, code = rule[0], rule[1:-1], rule[-1]
    values = [company.value(r.split(":")[0], company.period_year(r.split(":")[1])) for r in fields]
    if kind == "year_span":
        invalid = company.period_year("start") >= company.latest_year
    elif any(value is None for value in values):
        return None
    elif kind == "positive":
        invalid = values[0] <= 0
    elif kind == "nonzero":
        invalid = values[0] == 0
    elif kind == "nonnegative":
        invalid = values[0] < 0
    elif kind == "fraction":
        invalid = not 0 <= values[0] <= 1
    elif kind == "sum_positive":
        invalid = sum(values) <= 0
    elif kind == "at_most":
        invalid = values[0] > values[1]
    else:
        raise ValueError(f"Unknown validation rule: {kind}")
    if invalid:
        descriptions = {"positive": "must be positive", "nonzero": "must be nonzero", "nonnegative": "must be nonnegative", "fraction": "must be between 0 and 1", "sum_positive": "must have a positive sum", "at_most": "first value must not exceed second", "year_span": "CAGR requires distinct first and latest fiscal years"}
        return code, f"{', '.join(fields) or 'Fiscal years'}: {descriptions[kind]}."
    return None


def validate_company(company, metrics=None):
    metrics = metrics or METRICS
    company.metric_statuses.clear()
    company.issues = [i for i in company.issues if i.metric not in METRIC_ORDER or i.severity == "error"]
    for metric_id in METRIC_ORDER:
        definition = metrics[metric_id]
        missing, states = [], []
        for requirement in definition["requires"]:
            field, period = requirement.split(":")
            target = company.period_year(period)
            observation = company.observations.get((field, target))
            if company.value(field, target) is None:
                missing.append(f"{field} FY{target}")
                states.append(observation.value_status if observation else "missing")
        if company.blocked:
            status, message = "BLOCKED", "Company has blocking data errors; no workbook is produced."
        elif missing:
            status = "NOT_APPLICABLE" if "not_applicable" in states else "MISSING_INPUTS"
            message = "Missing/unavailable inputs: " + ", ".join(missing)
        else:
            status, message = "OK", "Required inputs validated."
            for rule in definition["rules"]:
                failure = rule_failure(company, rule)
                if failure:
                    status, message = failure
                    break
        company.metric_statuses[metric_id] = {"status": status, "message": message}
        if status not in {"OK", "BLOCKED"}:
            company.issues.append(Issue(status, message, company_id=company.company_id, metric=metric_id))
    return company

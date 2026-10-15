"""Populate only whitelisted template cells; no financial arithmetic here."""

from decimal import Decimal, localcontext
from copy import copy

from .config import METRIC_ORDER
from .schemas import ModelError
from .template import load_template, request_recalculation, set_value, split_address


def _excel_number(value):
    if value is None:
        return None
    # Excel stores 15 significant decimal digits; reject silent precision loss.
    with localcontext() as context:
        context.prec = max(40, len(value.as_tuple().digits) + 10)
        normalized = value.normalize()
    if len(normalized.as_tuple().digits) > 15 or (value != 0 and not Decimal("1e-300") <= abs(value) <= Decimal("1e300")):
        raise ModelError(f"Value exceeds Excel's safe 15-digit precision/range: {value}")
    return int(value) if value == value.to_integral_value() else float(value)


def populate(template, output, company, mapping, records, calculation_status):
    wb = load_template(template, mapping)

    def write(address, value):
        sheet, target, _ = split_address(address)
        set_value(wb[sheet][target], value)

    for address in mapping["metadata"].values():
        write(address, None)
    for kind in ("financials", "assumptions"):
        for periods in mapping[kind].values():
            for address in periods.values():
                write(address, None)
    for address in mapping["regions"].values():
        sheet, _, (x1, y1, x2, y2) = split_address(address)
        ws = wb[sheet]
        for row in ws.iter_rows(min_row=y1, max_row=min(y2, ws.max_row), min_col=x1, max_col=min(x2, ws.max_column)):
            for cell in row:
                cell.value = None
    year = company.latest_year
    metadata = {"company_id": company.company_id, "company_name": company.company_name, "latest_fiscal_year": year, "previous_fiscal_year": year - 1, "first_fiscal_year": company.period_year("start"), "currency": company.currency, "scope": company.scope}
    for field, address in mapping["metadata"].items():
        write(address, metadata[field])
    for kind in ("financials", "assumptions"):
        for metric, periods in mapping[kind].items():
            for period, address in periods.items():
                value = company.value(metric, company.period_year(period))
                write(address, _excel_number(value))

    def write_region(name, matrix):
        sheet, _, (x1, y1, x2, y2) = split_address(mapping["regions"][name])
        if len(matrix) > y2 - y1 + 1 or any(len(row) > x2 - x1 + 1 for row in matrix):
            raise ModelError(f"Template region {name} has insufficient capacity")
        ws = wb[sheet]
        previous_max_row = ws.max_row
        for row_index, values in enumerate(matrix, y1):
            for column_index, value in enumerate(values, x1):
                cell = ws.cell(row_index, column_index)
                set_value(cell, value)
                # Existing template formats are immutable. For newly appended
                # audit rows, copy the first allowlisted row's template style.
                if row_index > previous_max_row and name in {"checks", "source_map"}:
                    cell._style = copy(ws.cell(y1, column_index)._style)
            if row_index > previous_max_row and name in {"checks", "source_map"}:
                ws.row_dimensions[row_index].height = ws.row_dimensions[y1].height

    issues = [[i.severity, i.code, i.metric, i.company_id, i.message, i.source] for i in company.issues]
    if calculation_status != "RECALCULATED":
        issues.append(["warning", "FORMULA_NOT_RECALCULATED", None, company.company_id, "Open in Excel and recalculate, or rerun with --recalculate.", None])
    write_region("checks", issues or [["info", "OK", None, company.company_id, "No input issues found.", None]])
    columns = ["source_file", "source_sheet", "source_row", "source_column", "source_locator", "metric", "fiscal_year", "period_type", "scope", "currency", "original_value", "original_unit", "canonical_value", "canonical_unit", "value_status", "destination_cell", "dependent_metrics"]
    # Audit values remain exact text; typed numerical inputs reside on Inputs.
    write_region("source_map", [[str(r[key]) if key == "original_value" and r.get(key) is not None else r.get(key) for key in columns] for r in records])
    write_region("metric_statuses", [[company.metric_statuses[key]["status"], company.metric_statuses[key]["message"]] for key in METRIC_ORDER])
    write_region("run_status", [[calculation_status]])
    request_recalculation(wb)
    wb.save(output)

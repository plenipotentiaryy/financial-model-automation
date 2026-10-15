"""Versioned metric definitions and Excel expressions, never a Python KPI engine."""


def build_catalog(mapping):
    metrics = {}

    def ref(requirement):
        field, period = requirement.split(":")
        for group in ("financials", "assumptions"):
            if field in mapping[group]:
                return mapping[group][field][period]
        raise KeyError(requirement)

    def add(key, label, description, requirements, unit, expression, rules=(), limitations="", depends=()):
        requirements = list(dict.fromkeys(requirements))
        guards = [f"ISNUMBER({ref(r)})" for r in requirements]
        for rule in rules:
            kind, fields = rule[0], rule[1:-1]
            if kind == "positive":
                guards.append(f"{ref(fields[0])}>0")
            elif kind == "nonzero":
                guards.append(f"{ref(fields[0])}<>0")
            elif kind == "nonnegative":
                guards.append(f"{ref(fields[0])}>=0")
            elif kind == "fraction":
                guards.extend([f"{ref(fields[0])}>=0", f"{ref(fields[0])}<=1"])
            elif kind == "sum_positive":
                guards.append(f"SUM({','.join(ref(f) for f in fields)})>0")
            elif kind == "at_most":
                guards.append(f"{ref(fields[0])}<={ref(fields[1])}")
            elif kind == "year_span":
                first = mapping["metadata"]["first_fiscal_year"]
                last = mapping["metadata"]["latest_fiscal_year"]
                guards.extend([f"ISNUMBER({first})", f"ISNUMBER({last})", f"{last}>{first}"])
            else:
                raise ValueError(f"Unknown validation rule: {kind}")
        guards.extend(f"ISNUMBER({metrics[d]['cell']})" for d in depends)
        metrics[key] = {
            "label": label, "formula": description, "requires": requirements,
            "unit": unit, "cell": f"Calculations!B{len(metrics) + 2}",
            "report_cell": f"Report!B{len(metrics) + 7}",
            "implementation": "implemented", "limitations": limitations,
            "rules": [list(rule) for rule in rules], "depends_on": list(depends),
            "excel_formula": f'=IF(AND({",".join(guards)}),{expression},"")',
        }

    def positive(field):
        return ("positive", field, "INVALID_DENOMINATOR")

    def nonnegative(field):
        return ("nonnegative", field, "INVALID_INPUT")

    def pair(field):
        return [f"{field}:previous", f"{field}:latest"]

    def avg(field):
        return f"AVERAGE({ref(field + ':previous')},{ref(field + ':latest')})"

    rev, ebitda, debt, cash = [ref(f + ":latest") for f in ("revenue", "ebitda", "total_debt", "cash_and_equivalents")]
    prev_rev = ref("revenue:previous")
    add("revenue_growth_yoy", "Revenue growth YoY", "Revenue FY / Revenue FY-1 - 1", pair("revenue"), "%", f"{rev}/{prev_rev}-1", [("nonzero", "revenue:previous", "INVALID_DENOMINATOR")], "Consecutive comparable fiscal years; nonzero prior revenue.")
    add("ebitda_margin", "EBITDA margin", "EBITDA / Revenue", ["ebitda:latest", "revenue:latest"], "%", f"{ebitda}/{rev}", [("nonzero", "revenue:latest", "INVALID_DENOMINATOR")], "Same FY, scope and currency; nonzero revenue.")
    add("net_debt", "Net debt", "Total debt - Cash and equivalents", ["total_debt:latest", "cash_and_equivalents:latest"], "currency", f"{debt}-{cash}", limitations="Both balances at the same fiscal year end; negative net debt is allowed.")
    add("net_debt_to_ebitda", "Net debt / EBITDA", "Net debt / EBITDA", ["total_debt:latest", "cash_and_equivalents:latest", "ebitda:latest"], "x", f"B4/{ebitda}", [positive("ebitda:latest")], "Positive annual EBITDA required.", ["net_debt"])
    mrr = ref("monthly_recurring_revenue:latest")
    add("arr", "ARR", "Explicit fiscal-year-end MRR x 12", ["monthly_recurring_revenue:latest"], "currency", f"{mrr}*12", [nonnegative("monthly_recurring_revenue:latest")], "Annualized recurring run rate, not recognized annual revenue. Input MRR explicitly excludes nonrecurring revenue.")
    wacc_fields = ["market_cap", "market_value_debt", "pre_tax_cost_of_debt", "risk_free_rate", "equity_beta", "equity_risk_premium", "marginal_tax_rate"]
    e, d, kd, rf, beta, erp, tax = [ref(f + ":latest") for f in wacc_fields]
    wacc_rules = [nonnegative("market_cap:latest"), nonnegative("market_value_debt:latest"), ("sum_positive", "market_cap:latest", "market_value_debt:latest", "INVALID_DENOMINATOR"), nonnegative("pre_tax_cost_of_debt:latest"), nonnegative("equity_risk_premium:latest"), ("fraction", "marginal_tax_rate:latest", "INVALID_ASSUMPTION")]
    add("wacc", "WACC", "E/(E+D)*(Rf+beta*ERP) + D/(E+D)*Kd*(1-tax)", [f + ":latest" for f in wacc_fields], "%", f"{e}/SUM({e},{d})*({rf}+{beta}*{erp})+{d}/SUM({e},{d})*{kd}*(1-{tax})", wacc_rules, "Two-component CAPM model; explicit market equity/debt at FY end. Rates are decimal fractions; no book-debt substitution or inferred assumptions.")
    for key, label, numerator in [("gross_margin", "Gross margin", "gross_profit"), ("ebit_margin", "EBIT margin", "ebit"), ("net_profit_margin", "Net profit margin", "net_income")]:
        add(key, label, f"{numerator} / Revenue", [numerator + ":latest", "revenue:latest"], "%", f"{ref(numerator + ':latest')}/{rev}", [positive("revenue:latest")], "Reported numerator and positive annual revenue; no inferred numerator.")
    cfo, capex = ref("cash_flow_from_operations:latest"), ref("capex:latest")
    fcf_req = ["cash_flow_from_operations:latest", "capex:latest"]
    fcf_rules = [nonnegative("capex:latest")]
    add("free_cash_flow", "Free cash flow", "Operating cash flow - positive capex spending", fcf_req, "currency", f"{cfo}-{capex}", fcf_rules, "CFO less capex spending; this is not FCFF or FCFE.")
    fcf_cell = metrics["free_cash_flow"]["cell"]
    add("fcf_margin", "FCF margin", "Free cash flow / Revenue", fcf_req + ["revenue:latest"], "%", f"{fcf_cell}/{rev}", fcf_rules + [positive("revenue:latest")], "Same FY and CFO-minus-capex convention.", ["free_cash_flow"])
    ca, cl, inventory = [ref(f + ":latest") for f in ("current_assets", "current_liabilities", "inventory")]
    liquidity_rules = [nonnegative("current_assets:latest"), positive("current_liabilities:latest")]
    add("current_ratio", "Current ratio", "Current assets / Current liabilities", ["current_assets:latest", "current_liabilities:latest"], "x", f"{ca}/{cl}", liquidity_rules, "FY-end balances; current liabilities must be positive.")
    add("quick_ratio", "Quick ratio", "(Current assets - Inventory) / Current liabilities", ["current_assets:latest", "inventory:latest", "current_liabilities:latest"], "x", f"({ca}-{inventory})/{cl}", liquidity_rules + [nonnegative("inventory:latest"), ("at_most", "inventory:latest", "current_assets:latest", "INVALID_INPUT")], "Inventory-only exclusion convention; other illiquid current assets are not separately removed.")
    add("interest_coverage", "Interest coverage", "EBIT / Interest expense", ["ebit:latest", "interest_expense:latest"], "x", f"{ref('ebit:latest')}/{ref('interest_expense:latest')}", [positive("interest_expense:latest")], "Positive interest expense; negative EBIT remains a negative coverage result.")
    for key, label, field in [("roa", "Return on assets", "total_assets"), ("roe", "Return on equity", "total_equity")]:
        requirements = ["net_income:latest", *pair(field)]
        rules = [nonnegative(r) for r in pair(field)] + [("sum_positive", *pair(field), "INVALID_DENOMINATOR")]
        add(key, label, f"Net income / Average opening and closing {field}", requirements, "%", f"{ref('net_income:latest')}/{avg(field)}", rules, "Uses FY-1 and FY-end balances. Negative balance bases are excluded; average denominator must be positive.")
    ev_req = ["market_cap:latest", "total_debt:latest", "cash_and_equivalents:latest"]
    ev_rules = [nonnegative(r) for r in ev_req]
    add("enterprise_value", "Enterprise value (simplified)", "Market capitalization + reported debt - Cash", ev_req, "currency", f"{e}+{debt}-{cash}", ev_rules, "Reported debt is a proxy here. Excludes preferred stock and noncontrolling interests; market cap and balances must be at FY end.")
    ev_cell = metrics["enterprise_value"]["cell"]
    for key, label, denominator in [("ev_to_revenue", "EV / Revenue", "revenue"), ("ev_to_ebitda", "EV / EBITDA", "ebitda")]:
        add(key, label, f"Simplified enterprise value / Annual {denominator}", ev_req + [denominator + ":latest"], "x", f"{ev_cell}/{ref(denominator + ':latest')}", ev_rules + [positive(denominator + ":latest")], "Inherits simplified EV convention; denominator must be positive.", ["enterprise_value"])
    add("price_to_earnings", "P/E", "Market capitalization / Annual net income", ["market_cap:latest", "net_income:latest"], "x", f"{e}/{ref('net_income:latest')}", [nonnegative("market_cap:latest"), positive("net_income:latest")], "Loss-making and zero-profit periods have no meaningful P/E.")
    first, last = mapping["metadata"]["first_fiscal_year"], mapping["metadata"]["latest_fiscal_year"]
    for key, label, field in [("revenue_cagr", "Revenue CAGR", "revenue"), ("ebitda_cagr", "EBITDA CAGR", "ebitda")]:
        add(key, label, f"({field} latest / {field} first)^(1 / elapsed fiscal years) - 1", [field + ":start", field + ":latest"], "%", f"({ref(field + ':latest')}/{ref(field + ':start')})^(1/({last}-{first}))-1", [positive(field + ":start"), nonnegative(field + ":latest"), ("year_span", "INCOMPATIBLE_PERIODS")], "First and last presented FY, not count of rows; positive start and nonnegative end. Missing/invalid endpoints are not skipped.")
    prev_debt, prev_cash, prev_ebitda = [ref(f + ":previous") for f in ("total_debt", "cash_and_equivalents", "ebitda")]
    add("net_leverage_trend", "Net leverage change YoY", "Latest net debt / EBITDA - Previous FY net debt / EBITDA", [*pair("total_debt"), *pair("cash_and_equivalents"), *pair("ebitda")], "x", f"B5-({prev_debt}-{prev_cash})/{prev_ebitda}", [positive("ebitda:previous"), positive("ebitda:latest")], "Change in turns, not percent; positive means increased leverage. Compares consecutive FYs.", ["net_debt_to_ebitda"])
    add("cash_conversion", "Cash conversion (CFO / EBITDA)", "Operating cash flow / EBITDA", ["cash_flow_from_operations:latest", "ebitda:latest"], "%", f"{cfo}/{ebitda}", [positive("ebitda:latest")], "Chosen cash conversion definition is CFO/EBITDA; excludes capex. A negative CFO gives negative conversion.")
    cycle_requirements, cycle_rules = [], []
    for key, label, balance, flow in [("dso", "Days sales outstanding", "accounts_receivable", "revenue"), ("dio", "Days inventory outstanding", "inventory", "cogs"), ("dpo", "Days payable outstanding", "accounts_payable", "cogs")]:
        requirements = [*pair(balance), flow + ":latest"]
        rules = [nonnegative(r) for r in pair(balance)] + [positive(flow + ":latest")]
        add(key, label, f"Average {balance} / Annual {flow} x 365", requirements, "days", f"{avg(balance)}/{ref(flow + ':latest')}*365", rules, "365-day convention and opening/closing average. Revenue proxies credit sales for DSO; COGS proxies credit purchases for DPO; seasonal businesses may need finer data.")
        cycle_requirements.extend(requirements)
        cycle_rules.extend(rules)
    add("cash_conversion_cycle", "Cash conversion cycle (CCC)", "DSO + DIO - DPO", cycle_requirements, "days", f"{metrics['dso']['cell']}+{metrics['dio']['cell']}-{metrics['dpo']['cell']}", list(dict.fromkeys(cycle_rules)), "Inherits DSO/DIO/DPO conventions. Negative CCC is valid. This is a duration, separate from CFO/EBITDA.", ["dso", "dio", "dpo"])
    return metrics

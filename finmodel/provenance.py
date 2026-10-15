"""Lineage for every observation, including missing and unmapped optional fields."""


def source_records(company, mapping, metrics):
    destinations = {}
    for kind in ("financials", "assumptions"):
        for metric, periods in mapping[kind].items():
            for period, destination in periods.items():
                year = company.period_year(period)
                destinations.setdefault((metric, year), []).append(destination)
    records = []
    for key, observation in sorted(company.observations.items()):
        dependent = []
        for metric_id, definition in metrics.items():
            if definition["implementation"] != "implemented":
                continue
            for requirement in definition["requires"]:
                field, _, period = requirement.partition(":")
                year = company.period_year(period)
                if (field, year) == key:
                    dependent.append(metric_id)
        for source in observation.sources:
            records.append({
                "source_file": source.source_file, "source_sheet": source.source_sheet,
                "source_row": source.source_row, "source_column": source.source_column,
                "source_locator": source.source_locator, "metric": observation.metric,
                "fiscal_year": observation.fiscal_year, "period_type": observation.period_type,
                "scope": observation.scope, "currency": observation.currency,
                "original_value": source.original_value,
                "original_unit": source.original_unit,
                "canonical_value": str(observation.canonical_value) if observation.canonical_value is not None else None,
                "canonical_unit": observation.canonical_unit, "value_status": observation.value_status,
                "destination_cell": ", ".join(destinations[key]) if key in destinations else None,
                "dependent_metrics": ", ".join(dict.fromkeys(dependent)),
                "measurement_basis": observation.measurement_basis,
                "candidate_id": source.candidate_id, "revision": source.revision,
                "source_sha256": source.source_sha256, "derivation": observation.derivation,
            })
    return records

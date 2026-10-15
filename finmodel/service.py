"""Pipeline orchestration and per-company failure isolation."""

import os
import re
import tempfile
from pathlib import Path
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from .config import ALIASES, METRICS, TEMPLATE_MAP, load_config
from .ingest import read_file
from .mapping import alias_lookup
from .normalize import normalize
from .outputs import ensure_available, publish_reports, write_reports
from .provenance import source_records
from .recalculate import recalculate
from .schemas import Issue, ModelError
from .validate import validate_company
from .workbook import populate


def configurations(aliases=None, template_map=None, metric_definitions=None):
    def choose(path, name):
        local = Path("config") / name
        return path if path is not None else (local if local.is_file() else None)

    config = {
        "aliases": load_config(choose(aliases, "aliases.yaml"), ALIASES),
        "template_map": load_config(choose(template_map, "template_map.yaml"), TEMPLATE_MAP),
        "metric_definitions": load_config(choose(metric_definitions, "metric_definitions.yaml"), METRICS),
    }
    alias_lookup(config["aliases"])
    if config["template_map"].get("version") != 2:
        raise ModelError("Template map v2 is required for all 28 metrics. Generate a new v2 template/config with init in a separate directory, then migrate custom mappings.")
    definitions = config["metric_definitions"]
    for metric_id, definition in definitions.items():
        if not isinstance(definition, dict) or not isinstance(definition.get("requires"), list) or not all(isinstance(item, str) for item in definition["requires"]):
            raise ModelError(f"Invalid dependency list for metric {metric_id}")
    for metric_id, baseline in METRICS.items():
        definition = definitions.get(metric_id)
        if not isinstance(definition, dict) or any(key not in definition for key in ("label", "formula", "requires", "unit", "implementation", "limitations")):
            raise ModelError(f"Invalid/missing metric definition: {metric_id}")
        for key in ("requires", "formula", "implementation", "cell", "report_cell", "unit", "rules", "depends_on", "excel_formula"):
            if definition.get(key) != baseline.get(key):
                raise ModelError(f"Changing {metric_id}.{key} requires a versioned template/code change")
    return config


def load_dataset(paths, config, sheet=None):
    rows, issues = [], []
    for path in paths:
        try:
            imported, messages = read_file(path, config["aliases"], sheet)
            rows.extend(imported)
            issues.extend(messages)
        except (ModelError, BadZipFile, InvalidFileException) as exc:
            issues.append(Issue("INPUT_ERROR", str(exc), "error", source=str(path.resolve())))
    dataset = normalize(rows, issues)
    for company in dataset.companies.values():
        for issue in dataset.issues:
            if issue.source in company.source_files:
                company.issues.append(issue)
        validate_company(company, config["metric_definitions"])
    return dataset


def report_path(output: Path):
    return output.with_name(output.stem + ".reports")


def safe_name(company):
    import hashlib

    slug = re.sub(r"[^A-Za-z0-9_-]+", "_", company.company_id).strip("_-")[:64] or "company"
    suffix = hashlib.sha256(company.company_id.encode()).hexdigest()[:10]
    return f"{slug}_{suffix}_{company.latest_year}.xlsx"


def build_company(company, template, output, config, *, overwrite=False, calculate=False):
    if Path(output).is_symlink():
        raise ModelError(f"Refusing to replace a symbolic link: {output}")
    output, template = Path(output).resolve(), Path(template).resolve()
    if output.suffix.lower() != ".xlsx":
        raise ModelError("Output must have an .xlsx extension")
    if output == template or str(output) in company.source_files:
        raise ModelError("Output may not overwrite its template or source data")
    reports = report_path(output)
    ensure_available([output, reports], overwrite)
    output.parent.mkdir(parents=True, exist_ok=True)
    records = source_records(company, config["template_map"], config["metric_definitions"])
    calculation = {"status": "NOT_ATTEMPTED" if company.blocked else "FORMULA_NOT_RECALCULATED", "engine": None}
    failure = None
    with tempfile.TemporaryDirectory(prefix=".finmodel-", dir=output.parent) as temporary:
        temporary = Path(temporary)
        staged = temporary / output.name
        if company.blocked:
            outcome = "FAILED"
        else:
            try:
                if not template.is_file():
                    raise ModelError(f"Template not found: {template}")
                verify_source_hashes(company)
                populate(template, staged, company, config["template_map"], records, "RECALCULATED" if calculate else "FORMULA_NOT_RECALCULATED")
                if calculate:
                    calculation = recalculate(staged)
                verify_source_hashes(company)
                outcome = "SUCCESS_WITH_WARNINGS" if any(i.severity == "warning" for i in company.issues) or not calculate else "SUCCESS"
            except (ModelError, OSError, ValueError, TypeError, KeyError, IndexError, SyntaxError, BadZipFile, InvalidFileException) as exc:
                failure = Issue("RECALCULATION_FAILED" if calculate and staged.exists() else "EXPORT_FAILED", str(exc), "error", company.company_id)
                calculation = {"status": "FAILED" if calculate else "FORMULA_NOT_RECALCULATED", "engine": None}
                outcome = "FAILED"
        staged_reports = temporary / "reports"
        manifest = write_reports(staged_reports, company, records, outcome, calculation, config, template if template.is_file() else None, output if outcome != "FAILED" else None, [failure] if failure else None)
        publish_reports(staged_reports, reports)
        if outcome != "FAILED":
            os.replace(staged, output)
    return manifest


def verify_source_hashes(company):
    """Recheck evidence at build boundaries, including after a slow engine call."""
    from hashlib import sha256
    seen = set()
    for obs in company.observations.values():
        for source in obs.sources:
            if source.source_sha256 and (source.source_file, source.source_sha256) not in seen:
                seen.add((source.source_file, source.source_sha256))
                if sha256(Path(source.source_file).read_bytes()).hexdigest() != source.source_sha256:
                    raise ModelError(f"SOURCE_CHANGED: {source.source_file}")

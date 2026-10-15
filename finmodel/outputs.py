"""Machine-readable reports, readable validation and reproducibility hashes."""

import csv
import hashlib
import json
import shutil
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path

from . import __version__
from .recalculate import detect_engine
from .schemas import Issue, ModelError, serializable


def json_text(value):
    return json.dumps(serializable(value), indent=2, ensure_ascii=False, allow_nan=False) + "\n"


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def configuration_hash(config):
    return hashlib.sha256(json.dumps(config, sort_keys=True, ensure_ascii=False).encode()).hexdigest()


def ensure_available(paths, overwrite=False):
    for path in paths:
        if path.exists() and not overwrite:
            raise ModelError(f"Output already exists: {path}. Use --overwrite explicitly.")
        if path.is_symlink():
            raise ModelError(f"Refusing to replace a symbolic link: {path}")


def write_reports(directory, company, records, outcome, calculation, config, template=None, output=None, extra_issues=None):
    directory.mkdir(parents=True, exist_ok=True)
    issues = [*company.issues, *(extra_issues or [])]
    if calculation["status"] == "FORMULA_NOT_RECALCULATED":
        issues.append(Issue("FORMULA_NOT_RECALCULATED", "Formula caches are not fresh. Open in Excel and recalculate, or use --recalculate.", company_id=company.company_id))
    validation = {"company_id": company.company_id, "fiscal_year": company.latest_year, "outcome": outcome, "metric_statuses": company.metric_statuses, "issues": issues, "calculation": calculation}
    (directory / "validation.json").write_text(json_text(validation), encoding="utf-8")
    lines = [f"Company: {company.company_name} ({company.company_id})", f"Fiscal year: {company.latest_year}", f"Outcome: {outcome}", f"Calculation: {calculation['status']}", "", "Metric input readiness:"]
    for metric, entry in company.metric_statuses.items():
        lines.append(f"  {metric}: {entry['status']} — {entry['message']}")
    lines += ["", "Issues:"] + [f"  [{i.severity}] {i.code}: {i.message}" for i in issues]
    (directory / "validation.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    # CSV exported for audit can be opened in spreadsheets: neutralize text
    # injection here too, while exact raw strings remain in JSON/XLSX text cells.
    columns = list(records[0]) if records else ["source_file", "metric", "destination_cell"]
    with (directory / "source_map.csv").open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=columns)
        writer.writeheader()
        for record in records:
            safe = {}
            for key, value in record.items():
                if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
                    # Numeric negative values are not malicious text.
                    try:
                        Decimal(value)
                    except InvalidOperation:
                        value = "'" + value
                safe[key] = value
            writer.writerow(safe)
    manifest = {
        "schema_version": 1, "application_version": __version__,
        "run_at_utc": datetime.now(timezone.utc).isoformat(),
        "company_id": company.company_id, "fiscal_year": company.latest_year,
        "outcome": outcome, "output_file": str(output.resolve()) if output else None,
        "input_files": [{"path": path, "sha256": digest(path)} for path in sorted(company.source_files)],
        "template": {"path": str(template.resolve()), "sha256": digest(template)} if template else None,
        "configuration_sha256": {name: configuration_hash(value) for name, value in config.items()},
        "calculation": calculation, "available_calculation_engine": detect_engine(),
        "metric_statuses": company.metric_statuses,
        "counts": {"reported_observations": sum(o.value_status == "reported" for o in company.observations.values()), "derived_observations": sum(o.value_status == "derived_approximate" for o in company.observations.values()), "missing_observations": sum(o.value_status not in {"reported", "derived_approximate"} for o in company.observations.values()), "metrics_ready": sum(s["status"] == "OK" for s in company.metric_statuses.values()), "metrics_total": len(company.metric_statuses), "errors": sum(i.severity == "error" for i in issues), "warnings": sum(i.severity == "warning" for i in issues)},
        "observations": list(company.observations.values()),
        "issues": issues,
    }
    if company.review_audit is not None:
        manifest["input_mode"] = "approved_candidates"
        manifest["review_audit_sha256"] = configuration_hash(company.review_audit)
        (directory / "review_audit.json").write_text(json_text(company.review_audit))
    else:
        manifest["input_mode"] = "prepared_data"
    manifest["derivations"] = [o.derivation for o in company.observations.values() if o.derivation]
    (directory / "run_manifest.json").write_text(json_text(manifest), encoding="utf-8")
    return manifest


def publish_reports(staged: Path, destination: Path):
    # Only application-owned report directories can be replaced. Never delete
    # arbitrary user directories passed as an output path.
    if destination.exists():
        expected = {"validation.json", "validation.txt", "source_map.csv", "run_manifest.json", "review_audit.json"}
        if any(p.name not in expected or not p.is_file() for p in destination.iterdir()):
            raise ModelError(f"Report directory contains unrelated files: {destination}")
        for path in destination.iterdir():
            path.unlink()
        destination.rmdir()
    shutil.move(str(staged), str(destination))


def write_unresolved_reports(directory, paths, issues, config):
    """An unreadable file has no company identity; do not invent one."""
    directory.mkdir(parents=True, exist_ok=True)
    calculation = {"status": "NOT_ATTEMPTED", "engine": None}
    validation = {"company_id": None, "fiscal_year": None, "outcome": "FAILED", "metric_statuses": {}, "issues": issues, "calculation": calculation}
    (directory / "validation.json").write_text(json_text(validation), encoding="utf-8")
    (directory / "validation.txt").write_text("Import failed; company identity could not be established.\n" + "\n".join(f"[{i.severity}] {i.code}: {i.message}" for i in issues) + "\n", encoding="utf-8")
    (directory / "source_map.csv").write_text("source_file,metric,destination_cell\n", encoding="utf-8")
    manifest = {**validation, "schema_version": 1, "application_version": __version__, "run_at_utc": datetime.now(timezone.utc).isoformat(), "output_file": None, "input_files": [{"path": str(p.resolve()), "sha256": digest(p) if p.is_file() else None} for p in paths], "configuration_sha256": {key: configuration_hash(value) for key, value in config.items()}, "counts": {"reported_observations": 0, "missing_observations": 0, "errors": sum(i.severity == "error" for i in issues), "warnings": sum(i.severity == "warning" for i in issues)}}
    (directory / "run_manifest.json").write_text(json_text(manifest), encoding="utf-8")
    return manifest

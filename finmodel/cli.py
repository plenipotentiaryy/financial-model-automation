"""Public CLI: init, validate, build, batch and explain."""

import argparse
import sys
import tempfile
from pathlib import Path

from .outputs import ensure_available, json_text, publish_reports, write_reports, write_unresolved_reports
from .provenance import source_records
from .samples import initialize
from .schemas import ModelError
from .service import build_company, configurations, load_dataset, report_path, safe_name


def positive_budget(value):
    from decimal import Decimal, InvalidOperation
    try:
        amount = Decimal(value)
        if not amount.is_finite() or amount <= 0:
            raise InvalidOperation
    except InvalidOperation:
        raise argparse.ArgumentTypeError("cost limit must be a positive finite USD amount") from None
    return str(amount)


def parser():
    root = argparse.ArgumentParser(prog="finmodel", description="Offline financial data to traceable Excel models")
    commands = root.add_subparsers(dest="command", required=True)
    init = commands.add_parser("init", help="Generate example data, configurations and trusted template")
    init.add_argument("--directory", type=Path, default=Path("."))
    init.add_argument("--overwrite", action="store_true")
    for name in ("validate", "build", "batch", "explain"):
        command = commands.add_parser(name)
        command.add_argument("--aliases", type=Path)
        command.add_argument("--template-map", type=Path)
        command.add_argument("--metric-definitions", type=Path)
        if name == "explain":
            command.add_argument("metric_id")
            continue
        command.add_argument("--sheet", help="XLSX worksheet; mandatory for multiple nonempty sheets")
        command.add_argument("--overwrite", action="store_true")
        if name == "batch":
            command.add_argument("--input-dir", required=True, type=Path)
        else:
            if name == "build":
                source = command.add_mutually_exclusive_group(required=True)
                source.add_argument("--input", type=Path, help="Explicit prepared-data import (no Jev gate)")
                source.add_argument("--run", type=Path, help="Only approved candidates from a review run")
            else:
                command.add_argument("--input", required=True, type=Path)
            command.add_argument("--company", help="Exact company_id selection")
        if name in {"build", "batch"}:
            command.add_argument("--template", type=Path, default=Path("samples/template.xlsx"))
            command.add_argument("--recalculate", action="store_true", help="Require actual LibreOffice recalculation")
        if name == "build":
            command.add_argument("--output", type=Path, required=True)
            command.add_argument("--derive-ebitda", type=Path, help="Explicit component evidence JSON; prepared-data mode only")
        else:
            command.add_argument("--output-dir", type=Path, default=Path("outputs/validation" if name == "validate" else "outputs/batch"))
    extract = commands.add_parser("extract", help="Extract evidence-bound candidates; never approves values")
    extract.add_argument("--input", type=Path, required=True, nargs="+")
    extract.add_argument("--run-dir", type=Path, required=True)
    extract.add_argument("--sheet")
    extract.add_argument("--context", type=Path, help="JSON company/period/unit context; claims checked by reviewer")
    verify = commands.add_parser("verify", help="Verify with official Jev API; abstain on errors")
    verify.add_argument("--run", required=True, type=Path)
    verify.add_argument("--provider", choices=["jev"], required=True)
    verify.add_argument("--max-cost-usd", required=True, type=positive_budget)
    review = commands.add_parser("review", help="Serve local evidence review at 127.0.0.1")
    review.add_argument("--run", required=True, type=Path)
    review.add_argument("--port", type=int, default=8765)
    review.add_argument("--no-open", action="store_true")
    return root


def print_result(manifest):
    counts = manifest["counts"]
    print(f"{manifest['company_id']}: {manifest['outcome']} | metrics_ready={counts.get('metrics_ready', 0)}/{counts.get('metrics_total', 0)}, reported={counts['reported_observations']}, missing/unavailable={counts['missing_observations']}, errors={counts['errors']}, warnings={counts['warnings']} | {manifest['calculation']['status']}")
    if manifest["output_file"]:
        print(f"Workbook: {manifest['output_file']}")


def run(args):
    if args.command == "init":
        names = initialize(args.directory.resolve(), args.overwrite)
        print(f"Initialized {len(names)} files in {args.directory.resolve()}")
        return 0
    if args.command in {"extract", "verify", "review"}:
        import json
        from .review.store import Store
        if args.command == "extract":
            from .review.extract import extract_run
            context = json.loads(args.context.read_text()) if args.context else {}
            print(json_text(extract_run(args.input, args.run_dir, context, args.sheet)))
            return 0
        store = Store(args.run)
        if args.command == "verify":
            from .review.jev import verify_run
            result = verify_run(store, args.max_cost_usd)
            (store.path.parent / "verification.json").write_text(json_text(result))
            print(json_text(result))
            return 1 if result["status"] in {"NOT_RUN", "COMPLETED_WITH_ERRORS"} else 0
        if not 1024 <= args.port <= 65535:
            raise ModelError("Review port must be between 1024 and 65535")
        from .review.ui import create_app
        import uvicorn
        import threading
        import webbrowser
        import subprocess
        url = f"http://127.0.0.1:{args.port}"
        def open_page():
            if sys.platform == "darwin":
                subprocess.run(["open", "-a", "Safari", url], check=False)
            else:
                webbrowser.open(url)
        if not args.no_open:
            threading.Timer(1, open_page).start()
        print(f"Review: {url} (Ctrl+C to stop)")
        uvicorn.run(create_app(store, args.port), host="127.0.0.1", port=args.port)
        return 0
    config = configurations(args.aliases, args.template_map, args.metric_definitions)
    if args.command == "explain":
        metric = config["metric_definitions"].get(args.metric_id)
        if metric is None:
            raise ModelError(f"Unknown metric {args.metric_id}; available: {', '.join(config['metric_definitions'])}")
        print(json_text({"metric_id": args.metric_id, **metric}), end="")
        return 0
    if args.command == "build" and args.run:
        if args.derive_ebitda:
            raise ModelError("Derivation bundles are prepared-data inputs; not allowed to bypass approval in --run mode")
        from .review.store import Store
        from .review.gate import export_companies
        store = Store(args.run)
        companies = list(export_companies(store, config["metric_definitions"], args.company).values())
        if args.company:
            companies = [c for c in companies if c.company_id == args.company]
        if len(companies) != 1:
            raise ModelError("Select exactly one company with --company")
        manifest = build_company(companies[0], args.template, args.output, config, overwrite=args.overwrite, calculate=args.recalculate)
        print_result(manifest)
        return 1 if manifest["outcome"] == "FAILED" else 0
    if args.command == "batch":
        if not args.input_dir.is_dir():
            raise ModelError(f"Input directory not found: {args.input_dir}")
        paths = sorted(p for p in args.input_dir.iterdir() if p.is_file() and p.suffix.lower() in {".csv", ".json", ".xlsx"} and not p.name.startswith("~$"))
        if not paths:
            raise ModelError("Input directory has no CSV/JSON/XLSX files")
        output_root = args.output_dir.resolve()
        # Running into the input directory could ingest previous output workbooks.
        if output_root == args.input_dir.resolve():
            raise ModelError("Batch output directory must differ from input directory")
    else:
        paths = [args.input]
    dataset = load_dataset(paths, config, args.sheet)
    companies = list(dataset.companies.values())
    if getattr(args, "company", None):
        companies = [c for c in companies if c.company_id == args.company]
        if not companies:
            raise ModelError(f"Company {args.company!r} not found")
    if args.command == "build":
        if not companies and dataset.issues:
            if args.output.suffix.lower() != ".xlsx":
                raise ModelError("Output must have an .xlsx extension")
            destination = report_path(args.output.resolve())
            ensure_available([args.output, destination], args.overwrite)
            destination.parent.mkdir(parents=True, exist_ok=True)
            with tempfile.TemporaryDirectory(prefix=".finmodel-", dir=destination.parent) as temporary:
                staged = Path(temporary) / "reports"
                manifest = write_unresolved_reports(staged, paths, dataset.issues, config)
                publish_reports(staged, destination)
            print_result(manifest)
            print(f"Reports: {destination}")
            return 1
        if len(companies) != 1:
            detail = "; ".join(i.message for i in dataset.issues if i.severity == "error")
            raise ModelError(detail or "build requires exactly one company; use --company or batch")
        if args.derive_ebitda:
            import json
            from .derivations import reconstruct_ebitda
            reconstruct_ebitda(companies[0], json.loads(args.derive_ebitda.read_text()))
            companies[0].source_files.add(str(args.derive_ebitda.resolve()))
        manifest = build_company(companies[0], args.template, args.output, config, overwrite=args.overwrite, calculate=args.recalculate)
        print_result(manifest)
        print(f"Reports: {args.output.resolve().with_name(args.output.stem + '.reports')}")
        return 1 if manifest["outcome"] == "FAILED" else 0
    output_root = args.output_dir.resolve()
    summary_file = output_root / ("batch_summary.json" if args.command == "batch" else "validation_summary.json")
    ensure_available([summary_file], args.overwrite)
    output_root.mkdir(parents=True, exist_ok=True)
    results = []
    for company in sorted(companies, key=lambda c: c.company_id):
        try:
            if args.command == "batch":
                manifest = build_company(company, args.template, output_root / safe_name(company), config, overwrite=args.overwrite, calculate=args.recalculate)
            else:
                destination = output_root / (Path(safe_name(company)).stem + ".reports")
                ensure_available([destination], args.overwrite)
                records = source_records(company, config["template_map"], config["metric_definitions"])
                with tempfile.TemporaryDirectory(prefix=".finmodel-", dir=output_root) as temporary:
                    staged = Path(temporary) / "reports"
                    manifest = write_reports(staged, company, records, "FAILED" if company.blocked else "VALIDATED", {"status": "NOT_ATTEMPTED", "engine": None}, config)
                    publish_reports(staged, destination)
            print_result(manifest)
            results.append({"company_id": company.company_id, "outcome": manifest["outcome"], "output_file": manifest["output_file"], "calculation": manifest["calculation"]})
        except (ModelError, OSError, ValueError) as exc:
            results.append({"company_id": company.company_id, "outcome": "FAILED", "error": str(exc)})
            print(f"{company.company_id}: FAILED — {exc}", file=sys.stderr)
    failed = any(r["outcome"] == "FAILED" for r in results) or any(i.severity == "error" for i in dataset.issues) or not companies
    summary = {"outcome": "FAILED" if failed else "SUCCESS", "companies": results, "input_issues": dataset.issues}
    summary_file.write_text(json_text(summary), encoding="utf-8")
    for issue in dataset.issues:
        print(f"[{issue.severity}] {issue.code}: {issue.message}", file=sys.stderr)
    print(f"Summary: {summary_file}")
    return 1 if failed else 0


def main(argv=None):
    args = parser().parse_args(argv)
    try:
        return run(args)
    except (ModelError, OSError, ValueError, TypeError, KeyError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

"""Mandatory deterministic checks and the sole approved-candidate export boundary."""
from dataclasses import asdict
from decimal import Decimal, localcontext
import hashlib
import re

from ..config import NUMERIC_FIELDS, METADATA, RATES, METRICS
from ..normalize import SCALES, parse_number, normalize
from ..schemas import ModelError, RawRow, Source
from ..validate import validate_company
from .contracts import ReviewDecision

CANDIDATE_SCALES = SCALES | {'billions': Decimal(1000000000)}


def number(candidate):
    text = candidate.raw_value.strip().replace('−', '-')
    percent = text.endswith('%')
    if percent:
        text = text[:-1]
    if text.startswith('(') and text.endswith(')'):
        text = '-' + text[1:-1].strip()
    if ',' in text:
        if not re.fullmatch(r'[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?', text):
            raise ModelError('Ambiguous grouping separators')
        text = text.replace(',', '')
    value, status = parse_number(text)
    if percent and candidate.metric not in RATES:
        raise ModelError('Percentage cannot populate a money/count input')
    if percent and value is not None:
        value /= 100
    is_money = candidate.metric not in RATES | {'equity_beta', 'shares_outstanding', 'share_price'}
    if value is not None and is_money:
        if candidate.monetary_unit not in CANDIDATE_SCALES:
            raise ModelError('Missing/unsupported monetary scale')
        with localcontext() as ctx:
            ctx.prec = max(40, len(value.as_tuple().digits) + 12)
            value *= CANDIDATE_SCALES[candidate.monetary_unit]
    return value, status


def inspect_candidate(c):
    errors = []
    from pathlib import Path
    try:
        if hashlib.sha256(Path(c.source_file).read_bytes()).hexdigest() != c.source_sha256:
            errors.append('SOURCE_CHANGED')
    except OSError:
        errors.append('SOURCE_MISSING')
    if not c.locator or not c.excerpt:
        errors.append('MISSING_SOURCE_CONTEXT')
    if 'OCR_REQUIRED' in c.extraction_flags:
        errors.append('OCR_REQUIRED')
    if not c.metric or c.metric not in NUMERIC_FIELDS:
        errors.append('UNKNOWN_METRIC')
    if not c.company_id or not c.company_name:
        errors.append('MISSING_COMPANY')
    if isinstance(c.fiscal_year, bool) or not isinstance(c.fiscal_year, int) or not 1900 <= c.fiscal_year <= 2200 or c.period_type != 'FY':
        errors.append('INVALID_PERIOD')
    if c.scope not in {'consolidated', 'standalone'}:
        errors.append('INVALID_SCOPE')
    if not isinstance(c.currency, str) or not re.fullmatch('[A-Z]{3}', c.currency):
        errors.append('INVALID_CURRENCY')
    if c.monetary_unit not in CANDIDATE_SCALES:
        errors.append('INVALID_SCALE')
    try:
        value, status = number(c)
        c.normalized_value = str(value) if value is not None else None
        c.value_status = status
        if value is None:
            errors.append('NO_NUMERIC_VALUE')
    except (ModelError, TypeError, ValueError):
        c.normalized_value, c.value_status = None, 'invalid'
        errors.append('INVALID_NUMBER_OR_DIMENSION')
    # These labels must never be silently promoted to a clean D&A / EBITDA input.
    label = c.raw_label.casefold()
    if 'depreciation' in label and 'other' in label:
        errors.append('MIXED_DA_NOT_SUPPORTED')
    return list(dict.fromkeys(errors))


def conflicts(c, rows):
    result = []
    for row in rows:
        other = row['candidate']
        if other.candidate_id == c.candidate_id or row['state'] == 'rejected':
            continue
        if other.company_id != c.company_id:
            continue
        inspect_candidate(other)
        incompatible = (other.currency != c.currency or other.scope != c.scope or other.company_name != c.company_name)
        duplicate = (other.metric == c.metric and other.fiscal_year == c.fiscal_year and other.normalized_value != c.normalized_value)
        if incompatible or duplicate:
            result.append(other.candidate_id)
    return result


def decide(store, cid, revision, action, reason, changes=None):
    if not isinstance(reason, str) or not reason.strip():
        raise ModelError('A review reason is required')
    if action == 'correct':
        return store.revise(cid, revision, changes, reason).payload()
    if action not in {'accept', 'reject', 'select'}:
        raise ModelError('Unknown review action')
    row = store.get(cid)
    c = row['candidate']
    if c.revision != revision:
        raise ModelError('Stale revision')
    errors = inspect_candidate(c)
    if action != 'reject' and errors:
        raise ModelError('Correct source/mapping before acceptance: ' + ', '.join(errors))
    related = conflicts(c, store.list()) if action != 'reject' else []
    if related and action != 'select':
        raise ModelError('Conflicting sources; explicitly select one and reject alternatives')
    decision = ReviewDecision(cid, revision, action, reason)
    with store.connect() as db:
        db.execute('BEGIN IMMEDIATE')
        # Select is atomic and rejects every current conflicting alternative.
        if action == 'select':
            for other_id in related:
                other = store.get(other_id)['candidate']
                store.set_state(other_id, other.revision, 'rejected', f'Superseded by selected {cid}: {reason}', asdict(decision), db)
        store.set_state(cid, revision, 'rejected' if action == 'reject' else 'accepted_manual', reason, asdict(decision), db)
    return {'candidate_id': cid, 'action': action}


def export_companies(store, metrics=None, company_id=None):
    rows = [r for r in store.list() if company_id is None or r["candidate"].company_id == company_id]
    approved = [r['candidate'] for r in rows if r['state'] in {'accepted_manual', 'accepted_auto'}]
    if not approved:
        raise ModelError('No approved candidates; verify or review the run first')
    raw_rows = []
    for c in approved:
        errors = inspect_candidate(c)
        if errors:
            raise ModelError(f'{c.candidate_id}: approval invalid: {", ".join(errors)}')
        # Unresolved competitors also block previously accepted values after re-extraction.
        if conflicts(c, rows):
            raise ModelError(f'{c.candidate_id}: unresolved conflicting sources')
        metadata = {k: getattr(c, k) for k in METADATA}
        metadata['monetary_unit'] = 'units'  # c.normalized_value is already absolute.
        location = str(c.locator)
        source = Source(c.source_file, c.locator.get('sheet'), c.locator.get('row', c.locator.get('page', 0)), str(c.locator.get('column', '')), location, c.raw_value, c.monetary_unit)
        source.candidate_id, source.revision, source.source_sha256 = c.candidate_id, c.revision, c.source_sha256
        source.source_locator += f"; candidate={c.candidate_id}@{c.revision}; sha256={c.source_sha256}"
        raw_rows.append(RawRow(metadata | {c.metric: c.normalized_value}, {k: source for k in [*METADATA, c.metric]}))
    dataset = normalize(raw_rows)
    for company in dataset.companies.values():
        company_rows = [r for r in rows if r['candidate'].company_id == company.company_id]
        ids = {r['candidate'].candidate_id for r in company_rows}
        history = [event for event in store.history() if event['payload'].get('candidate_id', event['payload'].get('before', {}).get('candidate_id')) in ids]
        company.review_audit = {"schema_version": 1, "candidates": [r["candidate"].payload() | {"state": r["state"], "reason": r["reason"]} for r in company_rows], "history": history}
        for obs in company.observations.values():
            for source in obs.sources:
                original = next(c for c in approved if c.candidate_id == source.candidate_id)
                source.original_unit = original.monetary_unit
        validate_company(company, metrics or METRICS)
        if company.blocked:
            raise ModelError('; '.join(i.message for i in company.issues if i.severity == 'error'))
    return dataset.companies


def snapshot(store):
    rows = store.list()
    data = []
    for row in rows:
        c = row['candidate']
        errors = inspect_candidate(c)
        data.append(c.payload() | {'state': row['state'], 'reason': row['reason'], 'errors': errors, 'conflicts': conflicts(c, rows)})
    coverage = {}
    try:
        for company in export_companies(store).values():
            coverage[company.company_id] = company.metric_statuses
    except ModelError as exc:
        coverage['blocked'] = str(exc)
    return {'candidates': data, 'coverage': coverage, 'history': store.history()}

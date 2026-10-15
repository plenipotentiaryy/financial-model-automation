"""Opt-in source recipes, not an alternative KPI calculation engine."""
from decimal import Decimal
from pathlib import Path
from hashlib import sha256

from .schemas import Observation, Source, Issue, ModelError
from .normalize import parse_number, SCALES
from .validate import validate_company
from .config import METRICS

EBITDA_RECIPE = 'operating-ebitda-separate-da-v1'
EBITDA_NOTE = 'Approximate reconstructed operating EBITDA: operating EBIT + separately disclosed depreciation + intangible amortization; rounded source components. Not reported EBITDA; no extra ROU add-back.'


def component(data):
    required = {'value', 'unit', 'source_file', 'source_sha256', 'locator', 'label'}
    if not required <= set(data):
        raise ModelError('Incomplete derivation component evidence')
    path = Path(data['source_file']).resolve()
    if sha256(path.read_bytes()).hexdigest() != data['source_sha256']:
        raise ModelError('Derivation component source hash changed')
    if not data['locator'] or not data['label'] or 'other' in data['label'].casefold() or 'rou' in data['label'].casefold():
        raise ModelError('Separate pure depreciation/amortization evidence required; mixed/ROU lines are excluded')
    scales = SCALES | {'billions': Decimal(1000000000)}
    value, status = parse_number(data['value'])
    if status != 'reported' or value < 0 or data['unit'] not in scales:
        raise ModelError('Invalid derivation component number/unit')
    return value * scales[data['unit']], Source(str(path), None, 0, '', str(data['locator']), data['value'], data['unit'], source_sha256=data['source_sha256'])


def reconstruct_ebitda(company, bundle):
    """Explicit invocation required. Preserve any selected reported EBITDA."""
    if bundle.get('recipe') != EBITDA_RECIPE or bundle.get('company_id') != company.company_id or bundle.get('currency') != company.currency or bundle.get('scope') != company.scope:
        raise ModelError('EBITDA recipe identity/currency/scope mismatch')
    for item in bundle['periods']:
        year = item['fiscal_year']
        if year not in company.years:
            raise ModelError('EBITDA recipe year absent from company inputs')
        if company.value('ebitda', year) is not None:
            continue
        ebit = company.observations.get(('ebit', year))
        if ebit is None or company.value('ebit', year) is None:
            raise ModelError('Operating EBIT is required before reconstruction')
        dep, dep_source = component(item['depreciation'])
        amort, amort_source = component(item['intangible_amortization'])
        # Exact arithmetic on rounded published inputs; source rounding is not erased.
        value = ebit.canonical_value + dep + amort
        lineage = {'recipe': EBITDA_RECIPE, 'fiscal_year': year, 'note': EBITDA_NOTE,
                   'components': {'ebit': str(ebit.canonical_value), 'ebit_sources': [s.__dict__ for s in ebit.sources],
                                  'depreciation': item['depreciation'], 'intangible_amortization': item['intangible_amortization']}}
        company.observations[('ebitda', year)] = Observation(company.company_id, 'ebitda', year, 'FY', company.scope, company.currency,
            value, company.currency, 'units', 'derived_approximate', EBITDA_NOTE, [*ebit.sources, dep_source, amort_source], lineage)
        company.source_files.update([dep_source.source_file, amort_source.source_file])
    annotate_derivations(company)


def annotate_derivations(company):
    validate_company(company)
    company.issues = [i for i in company.issues if i.code != "DERIVED_APPROXIMATE"]
    for key, observation in company.observations.items():
        if not observation.derivation:
            continue
        company.issues.append(Issue('DERIVED_APPROXIMATE', observation.measurement_basis, company_id=company.company_id, metric=observation.metric))
        for metric, definition in METRICS.items():
            if any((field, company.period_year(period)) == key for field, period in (req.split(':') for req in definition['requires'])):
                if observation.measurement_basis not in company.metric_statuses[metric]['message']:
                    company.metric_statuses[metric]['message'] += ' ' + observation.measurement_basis

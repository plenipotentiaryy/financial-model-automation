"""Pure source recipes and separate independent Microsoft arithmetic checks."""
import copy
import hashlib
import json
from decimal import Decimal

import pytest

from finmodel.derivations import reconstruct_ebitda, EBITDA_RECIPE
from finmodel.service import configurations, load_dataset
from finmodel.schemas import ModelError
from research.microsoft_fy2025.replay_v03 import expected, bundle
from research.microsoft_fy2025.replay import extract


def setup_company(tmp_path, reported=None):
    row=dict(company_id='a',company_name='A',fiscal_year=2025,period_type='FY',scope='consolidated',currency='USD',monetary_unit='millions',ebit=100)
    if reported is not None:row['ebitda']=reported
    path=tmp_path/'data.json';path.write_text(json.dumps([row]))
    source=tmp_path/'note.txt';source.write_text('Depreciation 20.0 million; intangible amortization 3.0 million')
    component=dict(value='20.0',unit='millions',source_file=str(source),source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),locator='line 1',label='Depreciation')
    recipe=dict(recipe=EBITDA_RECIPE,company_id='a',scope='consolidated',currency='USD',periods=[dict(fiscal_year=2025,depreciation=component,intangible_amortization=component|{'value':'3.0','label':'Intangible amortization'})])
    return load_dataset([path],configurations()).companies['a'],recipe


def test_explicit_recipe_and_lineage(tmp_path):
    company,recipe=setup_company(tmp_path)
    assert company.value('ebitda') is None
    reconstruct_ebitda(company,recipe)
    assert company.value('ebitda')==Decimal('123000000')
    obs=company.observations[('ebitda',2025)]
    assert obs.value_status=='derived_approximate' and len(obs.sources)==3
    assert obs.derivation['recipe']==EBITDA_RECIPE
    assert 'Approximate reconstructed' in company.metric_statuses['ebitda_margin']['message']


def test_reported_ebitda_has_priority_including_zero(tmp_path):
    company,recipe=setup_company(tmp_path,0)
    reconstruct_ebitda(company,recipe)
    assert company.value('ebitda')==0
    assert company.observations[('ebitda',2025)].derivation is None


@pytest.mark.parametrize('bad', ['mixed','rou','hash','currency','year','unit','negative'])
def test_invalid_recipe_rejected(tmp_path,bad):
    company,recipe=setup_company(tmp_path);part=recipe['periods'][0]['depreciation']
    if bad=='mixed':part['label']='Depreciation, amortization and other'
    if bad=='rou':part['label']='ROU amortization'
    if bad=='hash':part['source_sha256']='wrong'
    if bad=='currency':recipe['currency']='PLN'
    if bad=='year':recipe['periods'][0]['fiscal_year']=2024
    if bad=='unit':part['unit']='unknown'
    if bad=='negative':part['value']='-1'
    with pytest.raises(ModelError):reconstruct_ebitda(company,recipe)


def test_three_case_expected_coverage_and_components():
    rows,*_=extract()
    assert [len(expected(p,rows)) for p in ('strict','ebitda','ebitda_market')]==[17,22,26]
    assert expected('ebitda_market',rows)['enterprise_value']==Decimal('3710654940000')
    assert bundle()['periods'][0]['depreciation']['value']=='22.0'
    assert bundle()['periods'][0]['intangible_amortization']['value']=='6.0'

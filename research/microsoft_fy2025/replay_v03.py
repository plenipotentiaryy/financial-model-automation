"""Three reproducible historical cases, independently audited after real recalculation.

No API/LLM responses are used. Source adapters are deliberately case-specific.
"""
import argparse
import hashlib
import json
import math
import sys
import shutil
from decimal import Decimal as D, localcontext
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from research.microsoft_fy2025 import replay
from finmodel.config import METRICS, METRIC_ORDER
from finmodel.derivations import reconstruct_ebitda, annotate_derivations, EBITDA_RECIPE, EBITDA_NOTE
from finmodel.outputs import ensure_available, json_text
from finmodel.schemas import Observation, Source
from finmodel.service import configurations, load_dataset, build_company
from openpyxl import load_workbook

CASE = Path(__file__).resolve().parent
V03 = CASE / 'v03'
OUTPUT = ROOT / 'outputs/microsoft_v03'
PRICE_FILE = CASE / 'raw/price_evidence.json'


def bundle():
    lines, locations = replay.archive_lines()
    result = {'recipe': EBITDA_RECIPE, 'company_id': 'msft', 'currency': 'USD', 'scope': 'consolidated', 'periods': []}
    for year, depreciation, amortization in [(2025,'22.0','6.0'),(2024,'15.2','4.8'),(2023,'11.0','2.5')]:
        item = {'fiscal_year': year}
        for field, value, line in [('depreciation',depreciation,1386),('intangible_amortization',amortization,1476)]:
            location = locations[line]
            assert value in lines[line] and str(year) in lines[line]
            item[field] = {'value': value, 'unit':'billions', 'label': field, 'source_file': location['file'],
                           'source_sha256': hashlib.sha256((ROOT/location['file']).read_bytes()).hexdigest(),
                           'locator': location, 'url': replay.URL, 'published_rounding': '0.1 billion USD'}
        result['periods'].append(item)
    return result


def add_market_cap(company):
    evidence = json.loads(PRICE_FILE.read_text())
    assert evidence['date']=='2025-06-30' and evidence['closing_price']=='497.41'
    shares = company.observations[('shares_outstanding',2025)]
    price = D(evidence['closing_price'])
    source = Source(str(PRICE_FILE), None, 0, '', evidence['source_location'], evidence['closing_price'], 'USD/share', source_sha256=hashlib.sha256(PRICE_FILE.read_bytes()).hexdigest())
    company.observations[('share_price',2025)] = Observation('msft','share_price',2025,'FY','consolidated','USD',price,'USD/share','USD/share','reported','fiscal_year_end',[source])
    note = 'Approximate historical market cap: unadjusted USD 497.41 close on 2025-06-30 x 7,434 million FY-end shares; share count is rounded. Not weighted-average shares.'
    lineage = {'recipe':'close-times-fy-end-shares-v1','fiscal_year':2025,'note':note,
               'components': {'share_price': evidence, 'shares_outstanding': str(shares.canonical_value), 'share_sources':[s.__dict__ for s in shares.sources]}}
    company.observations[('market_cap',2025)] = Observation('msft','market_cap',2025,'FY','consolidated','USD',price*shares.canonical_value,'USD','units','derived_approximate',note,[source,*shares.sources],lineage)
    company.source_files.add(str(PRICE_FILE))
    annotate_derivations(company)


def expected(profile, rows):
    """Independent arithmetic oracle; never fills any Excel cell."""
    result = replay.expected_metrics(rows)
    if profile != 'strict':
        # Hard-coded independent published-source values in millions.
        current, previous, first = D(156528), D(129433), D(102023)
        result.update(ebitda_margin=current/D(281724), net_debt_to_ebitda=D(12909)/current,
                      ebitda_cagr=(current/first).sqrt()-1,
                      net_leverage_trend=D(12909)/current-D(33315)/previous,
                      cash_conversion=D(136162)/current)
    if profile == 'ebitda_market':
        cap = D('497.41') * D(7434000000)
        ev = cap + D(12909000000)
        result.update(enterprise_value=ev,ev_to_revenue=ev/D(281724000000),
                      ev_to_ebitda=ev/D(156528000000),price_to_earnings=cap/D(101832000000))
    return result


def main(overwrite=False):
    V03.mkdir(exist_ok=True);OUTPUT.mkdir(parents=True,exist_ok=True)
    config=configurations(ROOT/'config/aliases.yaml',ROOT/'config/template_map.yaml',ROOT/'config/metric_definitions.yaml')
    rows,ledger,gaps,reconciliations=replay.extract()
    recipe=bundle()
    (V03/'ebitda_components.json').write_text(json_text(recipe))
    results={}
    for profile,count in [('strict',17),('ebitda',22),('ebitda_market',26)]:
        output=OUTPUT/f'{profile}.xlsx';template=OUTPUT/f'{profile}_template.xlsx'
        ensure_available([output,template,output.with_suffix('.reports'),V03/f'{profile}.xlsx'],overwrite)
        company=load_dataset([CASE/'financials.csv'],config).companies['msft']
        if profile!='strict':reconstruct_ebitda(company,recipe)
        if profile=='ebitda_market':add_market_cap(company)
        replay.case_template(ledger,template)
        wb=load_workbook(template);report=wb['Report']
        for row in wb['Filing_Sources'].iter_rows(min_row=4):
            obs = company.observations.get((row[1].value, row[0].value))
            if obs and obs.derivation:
                row[2].value = str(obs.canonical_value)
                row[3].value = 'absolute USD'
                row[4].value = 'derived_approximate'
                row[5].value = obs.derivation['recipe']
                row[6].value = 'See Derivations sheet for components and exact evidence'
                row[7].value = '; '.join(s.source_locator for s in obs.sources)
                row[8].value = obs.measurement_basis

        report['A44']=f'Microsoft FY2025 | {profile} | {count}/28 available | Historical June 30, 2025 inputs'
        report['A45']='Strict reported inputs; unavailable values stay blank.' if profile=='strict' else 'Approximate reconstructed operating EBITDA; rounded Notes 6/9 components. See Derivations.'
        report['A46']='Borrowing debt excludes leases; cash excludes short-term investments; CCC uses coarse proxies.'
        if profile!='strict':
            ws=wb.create_sheet('Derivations');ws.append(['Recipe','FY','Value (absolute USD)','Source / rounding / caveat'])
            for obs in company.observations.values():
                if obs.derivation:
                    ws.append([obs.derivation['recipe'],obs.fiscal_year,str(obs.canonical_value),json.dumps(obs.derivation,ensure_ascii=False)])
            for col,width in [('A',42),('B',10),('C',24),('D',120)]:ws.column_dimensions[col].width=width
            from openpyxl.styles import Alignment
            for row in ws:
                for cell in row:cell.alignment=Alignment(wrap_text=True,vertical='top')
                ws.row_dimensions[row[0].row].height=120 if row[0].row>1 else 28
            ws.freeze_panes='A2'
        # Source-specific interpretation plus explicit inherited derivation caveat.
        for metric,status in company.metric_statuses.items():
            if status['status']=='OK':
                status['message']=replay.notes_for(metric)+' '+status['message'].replace('Required inputs validated.','')
                if 'Approximate reconstructed' in status['message']:
                    status['message'] = 'Uses approximate reconstructed operating EBITDA (rounded Notes 6/9). Not reported EBITDA; see Derivations.'
                if profile=='ebitda_market' and any(req.startswith('market_cap:') for req in METRICS[metric]['requires']):
                    status['message'] += ' Historical close x rounded FY-end shares; simplified EV/PE convention.'
                report.row_dimensions[int(METRICS[metric]['report_cell'].split('!')[1][1:])].height=78

        wb.save(template)
        manifest=build_company(company,template,output,config,overwrite=overwrite,calculate=True)
        assert manifest['outcome']!='FAILED',manifest
        actual=load_workbook(output,data_only=True);formulas=load_workbook(output,data_only=False);baseline=load_workbook(template,data_only=False)
        for ws in baseline:
            for row in ws:
                for cell in row:
                    if cell.data_type=='f':assert formulas[ws.title][cell.coordinate].value==cell.value
        with localcontext() as ctx:
            ctx.prec=40;oracle=expected(profile,rows)
        audit=[]
        for metric in METRIC_ORDER:
            definition=METRICS[metric];sheet,cell=definition['cell'].split('!');rs,rc=definition['report_cell'].split('!')
            value,shown=actual[sheet][cell].value,actual[rs][rc].value
            status=company.metric_statuses[metric]
            if metric in oracle:
                assert status['status']=='OK' and math.isclose(value,float(oracle[metric]),rel_tol=1e-12,abs_tol=1e-8),(profile,metric,value,oracle[metric])
                assert value==shown
            else:
                assert status['status']=='MISSING_INPUTS' and value in (None,'') and shown in (None,'')
            audit.append({'metric_id':metric,'value':value if value!='' else None,'status':status['status'],'note':status['message'],
                          'expected_independent':str(oracle[metric]) if metric in oracle else None,'verification':'PASS' if metric in oracle else 'PASS_UNAVAILABLE'})
        assert len(oracle)==count and len(audit)==28
        results[profile]={'computed':count,'unavailable':28-count,'all_28_verified':True,'calculation':manifest['calculation'],
                          'workbook':str(output.relative_to(ROOT)),'sha256':hashlib.sha256(output.read_bytes()).hexdigest(),
                          'formula_preservation':'PASS','audit':audit,'derivations':manifest['derivations']}
        shutil.copyfile(output, V03/f'{profile}.xlsx')
        print(profile,count,'/28',manifest['calculation']['status'])
    # Recalculate an edited copy; check dependent Report outputs, not Python-only values.
    from finmodel.recalculate import recalculate
    edited=OUTPUT/'edited_input_check.xlsx';wb=load_workbook(OUTPUT/'ebitda_market.xlsx')
    wb['Inputs']['C7']=300000000000;wb.save(edited);recalculate(edited)
    check=load_workbook(edited,data_only=True)
    assert math.isclose(check['Report']['B7'].value,300000/D(245122)-1,rel_tol=1e-12)
    assert math.isclose(check['Report']['B8'].value,156528/D(300000),rel_tol=1e-12)
    evidence={'schema_version':1,'basis':'Historical FY2025, not current valuation', 'source_reconciliations':reconciliations,
              'variants':results,'edited_input_recalculation':'PASS','jev_live':{'status':'NOT_RUN','reason':'No TYPESAFE_API_KEY configured; no fixture responses used for Microsoft cases'}}
    (V03/'verification.json').write_text(json_text(evidence))
    return evidence


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--overwrite',action='store_true')
    main(parser.parse_args().overwrite)

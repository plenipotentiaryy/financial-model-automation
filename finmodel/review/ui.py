"""Loopback-only review service. No arbitrary file access, no remote CORS."""
import secrets
import json
from ..config import NUMERIC_FIELDS
from pathlib import Path

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse

from ..schemas import ModelError
from .gate import snapshot, decide

HTML = '''<!doctype html><html><head><meta charset="utf-8"><title>Financial evidence review</title>
<style>body{font:15px/1.5 system-ui;margin:32px;color:#24364b;background:#f5f7fa}h1{margin-bottom:0}button,select,input,textarea{font:inherit;padding:8px;margin:4px}button{cursor:pointer;background:#143b5b;color:white;border:0;border-radius:5px}article{background:white;padding:20px;margin:18px 0;border:1px solid #ccd6df;border-radius:8px}.grid{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:20px}.grid>div{min-width:0;overflow-wrap:anywhere}table{border-collapse:collapse;width:100%}td{padding:6px;border-bottom:1px solid #ddd}#coverage{max-height:320px;overflow:auto}@media(max-width:800px){.grid{grid-template-columns:1fr}}pre{white-space:pre-wrap;overflow-wrap:anywhere;max-height:340px;overflow:auto;background:#eef2f6;padding:12px}label{display:block}textarea{width:90%}.bad{color:#a42c28}#notice{white-space:pre-wrap}a{color:#1462a1}</style></head><body>
<h1>Financial evidence review</h1><p>Only approved values enter Excel. Source errors require correction. Every decision is retained locally.</p>
<p id="summary"></p><details><summary>Metric coverage</summary><div id="coverage"></div></details><p id="notice" role="alert"></p><button id="refresh">Refresh</button><div id="candidates"></div>
<script nonce="__TOKEN__">
const token='__TOKEN__';const metrics=__METRICS__;
const el=(tag,text)=>{const x=document.createElement(tag);if(text!==undefined)x.textContent=text;return x};
async function load(){const data=await (await fetch('/api/run')).json();const approved=data.candidates.filter(c=>c.state.startsWith('accepted')&&!c.errors.length&&!c.conflicts.length).length;document.querySelector('#summary').textContent=`${approved} approved / ${data.candidates.length} candidates. `+(data.coverage.blocked||Object.entries(data.coverage).map(([name,items])=>name+': '+Object.values(items).filter(x=>x.status==='OK').length+'/28 metrics ready').join('; '));const cover=document.querySelector('#coverage');cover.replaceChildren();for(const [name,items] of Object.entries(data.coverage)){if(name==='blocked'){cover.append(el('p',items));continue}cover.append(el('h3',name));const table=el('table');for(const [metric,status] of Object.entries(items)){const row=el('tr');row.append(el('td',metric),el('td',status.status),el('td',status.message));table.append(row)}cover.append(table)}const root=document.querySelector('#candidates');root.replaceChildren();for(const c of data.candidates){
const box=el('article');box.append(el('h2',`${c.raw_label || '(unlabelled)'}: ${c.raw_value} → ${c.metric || 'unknown'} / FY${c.fiscal_year || '?'}`));
box.append(el('p',`${c.state}: ${c.reason}`));const errors=el('p',[...c.errors,...c.conflicts.map(x=>'conflict '+x)].join(' · '));errors.className='bad';box.append(errors);
const grid=el('div');grid.className='grid';const evidence=el('div');evidence.append(el('strong',`${c.source_file} ${JSON.stringify(c.locator)}`));evidence.append(el('pre',c.excerpt));const link=el('a','Open source');link.href='/api/source/'+c.candidate_id;link.target='_blank';link.rel='noopener';evidence.append(link);grid.append(evidence);
const form=el('div');form.append(el('pre',JSON.stringify({id:c.candidate_id,revision:c.revision,normalized:c.normalized_value,company:c.company_name,currency:c.currency,scale:c.monetary_unit,scope:c.scope},null,2)));
const audit=data.history.filter(e=>e.payload.candidate_id===c.candidate_id);const detail=el('details');detail.append(el('summary','Checks and decision history'),el('pre',JSON.stringify(audit,null,2)));form.append(detail);
const reason=el('textarea');reason.placeholder='Reason for decision (required)';form.append(reason);
const edits=el('details');edits.append(el('summary','Correct metric or context'));const fields={};for(const [key,title,choices] of [['metric','Financial input',metrics],['company_id','Company identifier'],['company_name','Company name'],['fiscal_year','Fiscal year'],['period_type','Period',['FY']],['currency','Currency'],['monetary_unit','Monetary scale',['units','thousands','millions','billions']],['scope','Perimeter',['consolidated','standalone']]]){const label=el('label',title);const input=el(choices?'select':'input');if(choices){input.append(el('option',''));for(const choice of choices){const option=el('option',choice);option.value=choice;input.append(option)}}input.value=c[key]??'';fields[key]=input;label.append(input);edits.append(label)}form.append(edits);
for(const [action,title] of [['accept','Accept'],['reject','Reject'],['correct','Correct mapping'],['select','Select; reject conflicts']]){const b=el('button',title);b.onclick=async()=>{try{let changes=null;if(action==='correct'){changes={};for(const [key,input] of Object.entries(fields)){const value=key==='fiscal_year'?Number(input.value):input.value;if(value!==c[key])changes[key]=value}}const res=await fetch('/api/decision',{method:'POST',headers:{'Content-Type':'application/json','X-Review-Token':token},body:JSON.stringify({candidate_id:c.candidate_id,revision:c.revision,action,reason:reason.value,changes})});const out=await res.json();if(!res.ok)throw Error(out.detail);document.querySelector('#notice').textContent='Saved';await load()}catch(e){document.querySelector('#notice').textContent=e.message}};form.append(b)}grid.append(form);box.append(grid);root.append(box)}}
document.querySelector('#refresh').onclick=load;load();
</script></body></html>'''


def create_app(store, port=8765, token=None):
    token = token or secrets.token_urlsafe(32)
    origin = f'http://127.0.0.1:{port}'
    app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)

    @app.middleware('http')
    async def local_only(request: Request, call_next):
        if request.headers.get('host') != f'127.0.0.1:{port}':
            return JSONResponse({'detail': 'Invalid local host'}, status_code=403)
        if request.method not in {'GET', 'HEAD'}:
            if request.headers.get('origin') != origin or not secrets.compare_digest(request.headers.get('x-review-token', ''), token):
                return JSONResponse({'detail': 'Origin/session token required'}, status_code=403)
        response = await call_next(request)
        response.headers['Cache-Control'] = 'no-store'
        response.headers['X-Content-Type-Options'] = 'nosniff'
        response.headers['Content-Security-Policy'] = f"default-src 'self'; script-src 'nonce-{token}'; style-src 'unsafe-inline'; frame-ancestors 'none'"
        return response

    @app.get('/')
    def index():
        return HTMLResponse(HTML.replace('__TOKEN__', token).replace('__METRICS__', json.dumps(NUMERIC_FIELDS)))

    @app.get('/api/run')
    def run():
        return snapshot(store)

    @app.get('/api/source/{cid}')
    def source(cid: str):
        try:
            c = store.get(cid)['candidate']
            from .gate import inspect_candidate
            if any(e.startswith('SOURCE_') for e in inspect_candidate(c)):
                raise ModelError('Source changed or missing')
            path = Path(c.source_file)
            return FileResponse(path, media_type='application/pdf' if path.suffix.lower() == '.pdf' else 'application/octet-stream', filename=path.name)
        except ModelError as exc:
            raise HTTPException(409, str(exc)) from exc

    @app.post('/api/decision')
    async def decision(request: Request):
        try:
            data = await request.json()
            return decide(store, data['candidate_id'], data['revision'], data['action'], data['reason'], data.get('changes'))
        except (ModelError, KeyError, ValueError, TypeError) as exc:
            raise HTTPException(409, str(exc)) from exc

    return app

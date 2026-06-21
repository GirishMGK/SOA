"""
Local web UI for the SOA / RPS extraction & TOC-TOD tool.

Run:
    pip install -r requirements.txt
    python app.py            # then open http://127.0.0.1:5000

Two modes:
  * SOA - Statement of Account: per-loan working-paper workbooks + a portfolio
    exception report, with TOC/TOD validation checks.
  * RPS - Repayment Schedule: ONE combined workbook (all schedules stacked with
    Agreement No + a one-row-per-loan details sheet).

Upload individual PDFs, a whole folder, or a .zip. Files stream through one by
one with a live progress bar. Everything runs locally; nothing is persisted.
"""
import io
import os
import json
import time
import uuid
import shutil
import zipfile
import tempfile

from flask import (Flask, request, render_template_string, send_file,
                   Response, jsonify, abort)

from extract_soa import extract_loan, write_workbook, write_portfolio, to_num
from extract_rps import extract_rps, write_rps_workbook
from reconcile import classify, reconcile_jobs, write_reconciliation_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500 MB per request
VERSION = "4.0"


@app.after_request
def _no_cache(resp):
    # never let the browser serve a stale page/script
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    return resp

JOBS = {}
MAX_JOBS = 6


def _purge_old_jobs():
    while len(JOBS) > MAX_JOBS:
        oldest = min(JOBS, key=lambda k: JOBS[k]["created"])
        try:
            shutil.rmtree(JOBS[oldest]["dir"], ignore_errors=True)
        finally:
            JOBS.pop(oldest, None)


def _collect_pdfs(files, dest):
    pdfs = []
    for up in files:
        if not up or not up.filename:
            continue
        name = os.path.basename(up.filename.replace("\\", "/"))
        lower = name.lower()
        if lower.endswith(".zip"):
            zpath = os.path.join(dest, name)
            up.save(zpath)
            try:
                with zipfile.ZipFile(zpath) as z:
                    for member in z.namelist():
                        if member.lower().endswith(".pdf") and not member.startswith("__MACOSX"):
                            target = os.path.join(dest, f"{len(pdfs):03d}_{os.path.basename(member)}")
                            with z.open(member) as src, open(target, "wb") as out:
                                shutil.copyfileobj(src, out)
                            pdfs.append(target)
            except zipfile.BadZipFile:
                pass
            finally:
                os.remove(zpath)
        elif lower.endswith(".pdf"):
            target = os.path.join(dest, f"{len(pdfs):03d}_{name}")
            up.save(target)
            pdfs.append(target)
    return pdfs


PAGE = r"""
<!doctype html>
<html><head><meta charset="utf-8"><title>SOA / RPS Extractor (build __VER__)</title>
<style>
  body{font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:#f4f6f9;color:#1f2d3d;margin:0;padding:32px}
  .wrap{max-width:980px;margin:0 auto}
  .card{background:#fff;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.08);padding:24px;margin-bottom:18px}
  h1{color:#1F4E78;margin:0 0 2px;font-size:22px}
  .ver{font-size:12px;color:#8a93a0;margin:0 0 18px}
  h2{font-size:16px;margin:0 0 4px}
  .h-soa{color:#1F4E78}.h-rps{color:#1d6b32}.h-recon{color:#7a4f00}
  p.hint{color:#6b7280;margin:0 0 14px;font-size:13px}
  .drop{border:2px dashed #c3ccd6;border-radius:10px;padding:22px;text-align:center;background:#fafbfc;cursor:pointer}
  .drop.drag{border-color:#1F4E78;background:#eef4fb}
  input[type=file]{display:none}
  .files{margin-top:10px;font-size:13px;color:#374151}
  button{margin-top:14px;color:#fff;border:0;border-radius:8px;padding:11px 18px;font-size:14px;cursor:pointer}
  .b-soa{background:#1F4E78}.b-rps{background:#1d6b32}.b-recon{background:#7a4f00}
  button:hover{filter:brightness(1.08)} button:disabled{opacity:.5;cursor:not-allowed}
  .btn,a.btn{background:#1F4E78;color:#fff;border-radius:8px;padding:11px 18px;font-size:14px;text-decoration:none;display:inline-block;margin-top:6px}
  .btn.sec{background:#e8edf3;color:#1F4E78}
  #panel{display:none}
  .counts{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:14px}
  .pill{flex:1;min-width:110px;border-radius:10px;padding:12px 14px;text-align:center}
  .pill b{display:block;font-size:24px;line-height:1.1}.pill span{font-size:12px;color:#5b6675}
  .p-tot{background:#eef2f7}.p-clean{background:#dff3e6}.p-exc{background:#fff4d6}.p-fail{background:#fde2e4}
  .barwrap{background:#e8edf3;border-radius:20px;height:14px;overflow:hidden;margin-bottom:6px}
  .bar{height:100%;width:0;background:#1F4E78;transition:width .25s}
  .barlbl{font-size:12px;color:#5b6675;margin-bottom:14px}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th,td{padding:8px 10px;border-bottom:1px solid #eef0f3;text-align:left}
  th{background:#1F4E78;color:#fff;position:sticky;top:0}
  .tag{padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600}
  .t-clean{background:#c6efce;color:#1d6b32}.t-exc{background:#ffeb9c;color:#7a5b00}
  .t-fail{background:#ffc7ce;color:#7a1f2b}.t-rev{background:#ffeb9c;color:#7a5b00}.t-ok{background:#ddf3e6;color:#1d6b32}
  .scroll{max-height:420px;overflow:auto;border:1px solid #eef0f3;border-radius:8px}
  .summary{font-size:13px;color:#374151;margin:6px 0 12px}.dlrow{display:flex;gap:10px;flex-wrap:wrap}
  a.dl{color:#1F4E78;font-weight:600;text-decoration:none}
</style></head>
<body><div class="wrap">
  <div class="card">
    <h1>L&amp;T Finance Document Extractor</h1>
    <p class="ver">Build __VER__ &middot; separate fields for SOA and RPS &mdash; if this line is missing, you are running an old copy</p>

    <h2 class="h-soa">1 &middot; SOA &mdash; Statement of Account</h2>
    <p class="hint">Per-loan TOC/TOD workbooks + portfolio exception report. Upload SOA PDFs / folder-zip.</p>
    <label class="drop" data-for="soa"><input type="file" id="soa-input" accept=".pdf,.zip" multiple>
      <div><b>Click or drop SOA PDFs here</b></div><div class="files" id="soa-files"></div></label>
    <button class="b-soa" id="soa-go" disabled>Extract SOA</button>
  </div>

  <div class="card">
    <h2 class="h-rps">2 &middot; RPS &mdash; Repayment Schedule</h2>
    <p class="hint">Combined workbook: Loan_Details (one row per RPS) + Repayment_Schedule (all rows, tagged with Agreement No). No SOA sheets.</p>
    <label class="drop" data-for="rps"><input type="file" id="rps-input" accept=".pdf,.zip" multiple>
      <div><b>Click or drop RPS PDFs here</b></div><div class="files" id="rps-files"></div></label>
    <button class="b-rps" id="rps-go" disabled>Extract RPS</button>
  </div>

  <div class="card">
    <h2 class="h-recon">3 &middot; Reconcile &mdash; SOA vs RPS</h2>
    <p class="hint">Upload an SOA and its matching RPS together; matches by Agreement No and flags actual-vs-scheduled deviations.</p>
    <label class="drop" data-for="recon"><input type="file" id="recon-input" accept=".pdf,.zip" multiple>
      <div><b>Click or drop SOA + RPS PDFs here</b></div><div class="files" id="recon-files"></div></label>
    <button class="b-recon" id="recon-go" disabled>Reconcile</button>
  </div>

  <div class="card" id="panel">
    <div class="counts" id="counts"></div>
    <div class="barwrap"><div class="bar" id="bar"></div></div>
    <div class="barlbl" id="barlbl">Waiting…</div>
    <div class="summary" id="summary"></div>
    <div class="dlrow" id="dlrow" style="display:none"></div>
    <div class="scroll"><table><thead><tr id="thead"></tr></thead><tbody id="rows"></tbody></table></div>
  </div>
</div>
<script>
const panel=document.getElementById('panel'),rows=document.getElementById('rows');
const fmt=n=>n==null?'':Number(n).toLocaleString('en-IN');
function pills(defs){document.getElementById('counts').innerHTML=defs.map(d=>
  '<div class="pill '+d.cls+'"><b id="'+d.id+'">0</b><span>'+d.label+'</span></div>').join('');}
function thead(cols){document.getElementById('thead').innerHTML=cols.map(c=>'<th>'+c+'</th>').join('');}

// wire each upload field + button
['soa','rps','recon'].forEach(mode=>{
  const inp=document.getElementById(mode+'-input'),go=document.getElementById(mode+'-go'),
        fl=document.getElementById(mode+'-files'),drop=inp.closest('.drop');
  let chosen=[];
  function set(list){chosen=[...list].filter(f=>/\.(pdf|zip)$/i.test(f.name));
    fl.textContent=chosen.length?chosen.length+' file(s) selected':'';go.disabled=chosen.length===0;}
  inp.addEventListener('change',e=>set(e.target.files));
  ['dragover','dragenter'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('drag');}));
  ['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('drag');}));
  drop.addEventListener('drop',ev=>{ev.preventDefault();drop.classList.remove('drag');set(ev.dataTransfer.files);});
  go.addEventListener('click',()=>run(mode,chosen));
});

let tot,clean,exc,fail,done;
async function run(mode,chosen){
  if(!chosen.length)return;
  panel.style.display='block'; panel.scrollIntoView({behavior:'smooth'}); rows.innerHTML='';
  tot=clean=exc=fail=done=0;
  document.getElementById('summary').textContent=''; document.getElementById('dlrow').style.display='none';
  if(mode==='soa'){
    pills([{id:'c-tot',cls:'p-tot',label:'Uploaded'},{id:'c-clean',cls:'p-clean',label:'Clean'},
           {id:'c-exc',cls:'p-exc',label:'With exceptions'},{id:'c-fail',cls:'p-fail',label:'Failed'}]);
    thead(['#','File','Result','Stage','Exceptions','Parse','Sanctioned','Download']);
  }else if(mode==='rps'){
    pills([{id:'c-tot',cls:'p-tot',label:'Uploaded'},{id:'c-clean',cls:'p-clean',label:'Parsed'},
           {id:'c-fail',cls:'p-fail',label:'Failed'}]);
    thead(['#','File','Agreement No','Instalments','Status']);
  }else{
    pills([{id:'c-tot',cls:'p-tot',label:'Uploaded'},{id:'c-clean',cls:'p-clean',label:'SOA'},
           {id:'c-exc',cls:'p-exc',label:'RPS'},{id:'c-fail',cls:'p-fail',label:'Failed'}]);
    thead(['#','File','Type','Agreement No','Status']);
  }
  const fd=new FormData(); fd.append('mode',mode); chosen.forEach(f=>fd.append('pdfs',f));
  document.getElementById('barlbl').textContent='Uploading…';
  let job;
  try{job=await(await fetch('/upload',{method:'POST',body:fd})).json();}
  catch(err){document.getElementById('barlbl').textContent='Upload failed: '+err;return;}
  tot=job.total; document.getElementById('c-tot').textContent=tot;
  if(!tot){document.getElementById('barlbl').textContent='No PDF files found.';return;}
  const es=new EventSource('/process_stream/'+job.job_id);
  es.onmessage=ev=>{
    const d=JSON.parse(ev.data);
    if(d.type==='progress'){
      done++;
      if(!d.ok)fail++;
      else if(mode==='soa'){if(d.exceptions>0)exc++;else clean++;}
      else if(mode==='recon'){if(d.doctype==='rps')exc++;else if(d.doctype==='soa')clean++;else fail++;}
      else clean++;
      if(document.getElementById('c-clean'))document.getElementById('c-clean').textContent=clean;
      if(document.getElementById('c-exc'))document.getElementById('c-exc').textContent=exc;
      document.getElementById('c-fail').textContent=fail;
      const pct=Math.round(done/tot*100);
      document.getElementById('bar').style.width=pct+'%';
      document.getElementById('barlbl').textContent='Processed '+done+' / '+tot+' ('+pct+'%)';
      if(mode==='soa'){
        let rt,res; if(!d.ok){rt='t-fail';res='FAILED';}
        else if(d.exceptions>0){rt='t-exc';res='EXCEPTION';}else{rt='t-clean';res='CLEAN';}
        const pt=d.parse==='REVIEW'?'t-rev':'t-ok';
        const dl=d.ok?'<a class="dl" href="/download/'+job.job_id+'/'+d.index+'">Excel</a>':'';
        rows.insertAdjacentHTML('beforeend','<tr><td>'+(d.index+1)+'</td><td>'+d.file+'</td>'+
          '<td><span class="tag '+rt+'">'+res+'</span></td><td>'+(d.stage||'')+'</td>'+
          '<td>'+(d.ok?d.exceptions:('— '+(d.error||'')))+'</td>'+
          '<td>'+(d.ok?'<span class="tag '+pt+'">'+d.parse+'</span>':'')+'</td>'+
          '<td>'+(d.sanctioned!=null?fmt(d.sanctioned):'')+'</td><td>'+dl+'</td></tr>');
      }else if(mode==='rps'){
        const rt=d.ok?'t-ok':'t-fail', st=d.ok?'OK':('FAILED — '+(d.error||''));
        rows.insertAdjacentHTML('beforeend','<tr><td>'+(d.index+1)+'</td><td>'+d.file+'</td>'+
          '<td>'+(d.agreement||'')+'</td><td>'+(d.ok?d.instalments:'')+'</td>'+
          '<td><span class="tag '+rt+'">'+st+'</span></td></tr>');
      }else{
        const ok=d.ok&&d.doctype!=='unknown', rt=ok?'t-ok':'t-fail';
        const st=ok?(d.doctype.toUpperCase()+' parsed'):('FAILED — '+(d.error||'unrecognised'));
        rows.insertAdjacentHTML('beforeend','<tr><td>'+(d.index+1)+'</td><td>'+d.file+'</td>'+
          '<td>'+(d.doctype||'').toUpperCase()+'</td><td>'+(d.agreement||'')+'</td>'+
          '<td><span class="tag '+rt+'">'+st+'</span></td></tr>');
      }
    }else if(d.type==='done'){
      es.close(); const s=d.summary, dlrow=document.getElementById('dlrow');
      if(mode==='soa'){
        document.getElementById('barlbl').textContent='Done — processed '+s.total+' file(s).';
        document.getElementById('summary').innerHTML='<b>Portfolio health:</b> '+s.clean+
          ' clean · '+s.exceptions+' with exceptions · '+s.failed+' failed · NPA: '+s.npa+
          ' · Total sanctioned: ₹'+fmt(s.total_sanctioned);
        dlrow.innerHTML=(s.parsed>=2?'<a class="btn" href="/download_portfolio/'+job.job_id+'">Portfolio report</a>':'')+
          '<a class="btn sec" href="/download_zip/'+job.job_id+'?filter=all">Download all (zip)</a>'+
          '<a class="btn sec" href="/download_zip/'+job.job_id+'?filter=exceptions">Exceptions only (zip)</a>';
        dlrow.style.display='flex';
      }else if(mode==='rps'){
        document.getElementById('barlbl').textContent='Done — '+s.parsed+' schedule(s) extracted.';
        document.getElementById('summary').innerHTML='<b>Combined:</b> '+s.parsed+' agreement(s) · '+
          s.rows+' schedule rows · '+s.failed+' failed';
        dlrow.innerHTML='<a class="btn" href="/download_rps/'+job.job_id+'">Download combined workbook</a>';
        dlrow.style.display='flex';
      }else{
        document.getElementById('barlbl').textContent='Done — '+s.matched+' agreement(s) reconciled.';
        document.getElementById('summary').innerHTML='<b>Reconciliation:</b> '+s.matched+
          ' matched ('+s.reconciled+' reconciled, '+s.exceptions+' with deviations) · '+
          s.soa_only+' SOA-only · '+s.rps_only+' RPS-only';
        if(s.matched>0){dlrow.innerHTML='<a class="btn" href="/download_recon/'+job.job_id+'">Download reconciliation workbook</a>';dlrow.style.display='flex';}
      }
    }
  };
  es.onerror=()=>{document.getElementById('barlbl').textContent='Connection lost during processing.';es.close();};
}
</script></body></html>
"""


@app.route("/")
def index():
    return render_template_string(PAGE.replace("__VER__", VERSION))


@app.route("/upload", methods=["POST"])
def upload():
    uploads = request.files.getlist("pdfs")
    if not uploads:
        return jsonify(error="No files received."), 400
    mode = request.form.get("mode", "soa")
    _purge_old_jobs()
    job_id = uuid.uuid4().hex
    job_dir = tempfile.mkdtemp(prefix="job_" + job_id + "_")
    pdfs = _collect_pdfs(uploads, job_dir)
    JOBS[job_id] = {"dir": job_dir, "pdfs": pdfs, "results": [], "mode": mode,
                    "created": time.time(), "portfolio": None, "rps": None,
                    "recon": None, "loans": []}
    return jsonify(job_id=job_id, total=len(pdfs), mode=mode)


@app.route("/process_stream/<job_id>")
def process_stream(job_id):
    job = JOBS.get(job_id)
    if not job:
        abort(404)

    def gen_soa():
        results = job["results"]
        for i, pdf_path in enumerate(job["pdfs"]):
            display = os.path.basename(pdf_path)[4:]
            rec = {"index": i, "file": display, "ok": False, "exceptions": 0,
                   "stage": "", "parse": "", "sanctioned": None, "error": "", "xlsx": None}
            try:
                if classify(pdf_path) == "rps":
                    raise ValueError("This is a Repayment Schedule (RPS) - use the RPS field")
                d = extract_loan(pdf_path)
                xlsx = os.path.join(job["dir"], f"{i:03d}_{os.path.splitext(display)[0]}.xlsx")
                write_workbook(xlsx, d["master"], d["fin_rows"], d["recv"], d["disb"], d["txns"],
                               d["checks"], d["dpd_rows"], part_pay=d["part_pay"], bounce=d["bounce"],
                               charges=d["charges"], amort=d["amort"], bounce_grid=d["bounce_grid"],
                               quality=d["quality"])
                rec.update(ok=True, exceptions=d["summary"]["Exceptions"],
                           stage=d["summary"]["Current Stage"], parse=d["quality"]["status"],
                           sanctioned=to_num(d["master"].get("Loan Amount (Rs)")), xlsx=xlsx, data=d)
            except Exception as e:
                rec["error"] = str(e)[:120]
            results.append(rec)
            pl = {k: rec[k] for k in ("index", "file", "ok", "exceptions", "stage", "parse", "sanctioned", "error")}
            pl["type"] = "progress"
            yield f"data: {json.dumps(pl)}\n\n"
        good = [r for r in results if r["ok"]]
        if len(good) >= 2:
            ppath = os.path.join(job["dir"], "Portfolio_Exception_Report.xlsx")
            write_portfolio(ppath, [r["data"] for r in good]); job["portfolio"] = ppath
        summary = {"total": len(results), "parsed": len(good),
                   "failed": sum(1 for r in results if not r["ok"]),
                   "clean": sum(1 for r in good if r["exceptions"] == 0),
                   "exceptions": sum(1 for r in good if r["exceptions"] > 0),
                   "npa": sum(1 for r in good if "NPA" in (r["stage"] or "")),
                   "total_sanctioned": round(sum(r["sanctioned"] or 0 for r in good), 2)}
        yield f"data: {json.dumps({'type': 'done', 'summary': summary})}\n\n"

    def gen_rps():
        results, loans = job["results"], job["loans"]
        for i, pdf_path in enumerate(job["pdfs"]):
            display = os.path.basename(pdf_path)[4:]
            rec = {"index": i, "file": display, "ok": False, "agreement": "",
                   "instalments": 0, "error": ""}
            try:
                if classify(pdf_path) == "soa":
                    raise ValueError("This is a Statement of Account (SOA) - use the SOA field")
                ln = extract_rps(pdf_path)
                loans.append(ln)
                rec.update(ok=True, agreement=ln["master"].get("Agreement No"),
                           instalments=len(ln["schedule"]))
            except Exception as e:
                rec["error"] = str(e)[:120]
            results.append(rec)
            pl = {**rec, "type": "progress"}
            yield f"data: {json.dumps(pl)}\n\n"
        if loans:
            rpath = os.path.join(job["dir"], "RPS_Combined.xlsx")
            write_rps_workbook(rpath, loans); job["rps"] = rpath
        summary = {"parsed": len(loans), "failed": sum(1 for r in results if not r["ok"]),
                   "rows": sum(len(l["schedule"]) for l in loans)}
        yield f"data: {json.dumps({'type': 'done', 'summary': summary})}\n\n"

    def gen_recon():
        results, soas, rpss = job["results"], [], []
        for i, pdf_path in enumerate(job["pdfs"]):
            display = os.path.basename(pdf_path)[4:]
            rec = {"index": i, "file": display, "ok": False, "doctype": "unknown",
                   "agreement": "", "error": ""}
            try:
                kind = classify(pdf_path)
                rec["doctype"] = kind
                if kind == "soa":
                    d = extract_loan(pdf_path); soas.append(d)
                    rec.update(ok=True, agreement=d["master"].get("Agreement No"))
                elif kind == "rps":
                    d = extract_rps(pdf_path); rpss.append(d)
                    rec.update(ok=True, agreement=d["master"].get("Agreement No"))
                else:
                    rec["error"] = "not a recognised SOA or RPS"
            except Exception as e:
                rec["error"] = str(e)[:120]
            results.append(rec)
            yield f"data: {json.dumps({**rec, 'type': 'progress'})}\n\n"

        pairs, soa_only, rps_only = reconcile_jobs(soas, rpss)
        if pairs:
            rpath = os.path.join(job["dir"], "SOA_RPS_Reconciliation.xlsx")
            write_reconciliation_workbook(rpath, pairs, soa_only, rps_only)
            job["recon"] = rpath
        summary = {"matched": len(pairs),
                   "reconciled": sum(1 for p in pairs if p["summary"]["result"] == "RECONCILED"),
                   "exceptions": sum(1 for p in pairs if p["summary"]["result"] == "EXCEPTION"),
                   "soa_only": len(soa_only), "rps_only": len(rps_only)}
        yield f"data: {json.dumps({'type': 'done', 'summary': summary})}\n\n"

    gen = {"rps": gen_rps, "recon": gen_recon}.get(job["mode"], gen_soa)
    return Response(gen(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/download/<job_id>/<int:idx>")
def download_one(job_id, idx):
    job = JOBS.get(job_id)
    if not job or idx >= len(job["results"]):
        abort(404)
    rec = job["results"][idx]
    if not rec.get("xlsx") or not os.path.exists(rec["xlsx"]):
        abort(404)
    return send_file(rec["xlsx"], as_attachment=True,
                     download_name=os.path.splitext(rec["file"])[0] + ".xlsx")


@app.route("/download_portfolio/<job_id>")
def download_portfolio(job_id):
    job = JOBS.get(job_id)
    if not job or not job.get("portfolio"):
        abort(404)
    return send_file(job["portfolio"], as_attachment=True,
                     download_name="Portfolio_Exception_Report.xlsx")


@app.route("/download_rps/<job_id>")
def download_rps(job_id):
    job = JOBS.get(job_id)
    if not job or not job.get("rps"):
        abort(404)
    return send_file(job["rps"], as_attachment=True, download_name="RPS_Combined.xlsx")


@app.route("/download_recon/<job_id>")
def download_recon(job_id):
    job = JOBS.get(job_id)
    if not job or not job.get("recon"):
        abort(404)
    return send_file(job["recon"], as_attachment=True,
                     download_name="SOA_RPS_Reconciliation.xlsx")


@app.route("/download_zip/<job_id>")
def download_zip(job_id):
    job = JOBS.get(job_id)
    if not job:
        abort(404)
    only_exc = request.args.get("filter") == "exceptions"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        if job.get("portfolio") and not only_exc:
            z.write(job["portfolio"], "Portfolio_Exception_Report.xlsx")
        for r in job["results"]:
            if not r.get("xlsx"):
                continue
            if only_exc and r.get("exceptions", 0) == 0:
                continue
            z.write(r["xlsx"], os.path.join("loan_details", os.path.splitext(r["file"])[0] + ".xlsx"))
    buf.seek(0)
    name = "SOA_exceptions.zip" if only_exc else "SOA_TOC_TOD_results.zip"
    return send_file(buf, as_attachment=True, mimetype="application/zip", download_name=name)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n  SOA / RPS web UI (build "+VERSION+")  ->  http://127.0.0.1:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)

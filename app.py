"""
Local web UI for the SOA TOC/TOD tool.

Run:
    pip install -r requirements.txt
    python app.py            # then open http://127.0.0.1:5000

Upload SOA PDFs (individually, as a folder, or as a .zip). Files are processed
one-by-one with a live progress bar and an on-screen results table; download any
single workbook, the portfolio report, or a zip of all / only-exception loans.
Everything runs locally and nothing is persisted beyond the running session.
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

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500 MB per request

# In-memory job registry (local, single-user). Each job owns a temp dir.
JOBS = {}
MAX_JOBS = 6


def _purge_old_jobs():
    """Keep only the most recent jobs to bound disk/memory use."""
    while len(JOBS) > MAX_JOBS:
        oldest = min(JOBS, key=lambda k: JOBS[k]["created"])
        try:
            shutil.rmtree(JOBS[oldest]["dir"], ignore_errors=True)
        finally:
            JOBS.pop(oldest, None)


def _collect_pdfs(files, dest):
    """Save uploaded files into dest, expanding any .zip; return list of PDF paths."""
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
<html><head><meta charset="utf-8"><title>SOA TOC/TOD Extractor</title>
<style>
  body{font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:#f4f6f9;
       color:#1f2d3d;margin:0;padding:32px}
  .wrap{max-width:980px;margin:0 auto}
  .card{background:#fff;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.08);padding:28px;margin-bottom:20px}
  h1{color:#1F4E78;margin:0 0 4px;font-size:22px}
  p.sub{color:#6b7280;margin:0 0 20px;font-size:14px}
  .drop{border:2px dashed #c3ccd6;border-radius:10px;padding:30px;text-align:center;background:#fafbfc;cursor:pointer}
  .drop.drag{border-color:#1F4E78;background:#eef4fb}
  input[type=file]{display:none}
  .btnrow{margin-top:16px;display:flex;gap:10px;flex-wrap:wrap}
  button,.btn{background:#1F4E78;color:#fff;border:0;border-radius:8px;padding:11px 18px;font-size:14px;cursor:pointer;text-decoration:none;display:inline-block}
  button.sec,.btn.sec{background:#e8edf3;color:#1F4E78}
  button:hover{filter:brightness(1.07)}
  button:disabled{opacity:.5;cursor:not-allowed}
  .files{margin-top:12px;font-size:13px;color:#374151}
  /* progress */
  #panel{display:none}
  .counts{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:14px}
  .pill{flex:1;min-width:120px;border-radius:10px;padding:12px 14px;text-align:center}
  .pill b{display:block;font-size:24px;line-height:1.1}
  .pill span{font-size:12px;color:#5b6675}
  .p-tot{background:#eef2f7}.p-clean{background:#dff3e6}.p-exc{background:#fff4d6}.p-fail{background:#fde2e4}
  .barwrap{background:#e8edf3;border-radius:20px;height:14px;overflow:hidden;margin-bottom:6px}
  .bar{height:100%;width:0;background:#1F4E78;transition:width .25s}
  .barlbl{font-size:12px;color:#5b6675;margin-bottom:16px}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th,td{padding:8px 10px;border-bottom:1px solid #eef0f3;text-align:left}
  th{background:#1F4E78;color:#fff;position:sticky;top:0}
  .tag{padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600}
  .t-clean{background:#c6efce;color:#1d6b32}.t-exc{background:#ffeb9c;color:#7a5b00}
  .t-fail{background:#ffc7ce;color:#7a1f2b}.t-rev{background:#ffeb9c;color:#7a5b00}.t-ok{background:#ddf3e6;color:#1d6b32}
  .scroll{max-height:420px;overflow:auto;border:1px solid #eef0f3;border-radius:8px}
  .summary{font-size:13px;color:#374151;margin:6px 0 16px}
  a.dl{color:#1F4E78;font-weight:600;text-decoration:none}
</style></head>
<body><div class="wrap">
  <div class="card">
    <h1>SOA TOC/TOD Extractor</h1>
    <p class="sub">L&amp;T Finance NBFC Statement-of-Account &rarr; audit-ready Excel, with TOC/TOD validation</p>
    <label class="drop" id="drop">
      <input type="file" id="pdfs" accept=".pdf,.zip" multiple>
      <div id="label"><b>Click to choose</b> or drop SOA PDFs / a .zip here</div>
      <div class="files" id="files"></div>
    </label>
    <input type="file" id="folder" webkitdirectory directory multiple>
    <div class="btnrow">
      <button type="button" id="start" disabled>Extract &amp; Validate</button>
      <button type="button" class="sec" id="pickFolder">Select a folder instead</button>
      <button type="button" class="sec" id="reset">Reset</button>
    </div>
  </div>

  <div class="card" id="panel">
    <div class="counts">
      <div class="pill p-tot"><b id="c-tot">0</b><span>Uploaded</span></div>
      <div class="pill p-clean"><b id="c-clean">0</b><span>Clean</span></div>
      <div class="pill p-exc"><b id="c-exc">0</b><span>With exceptions</span></div>
      <div class="pill p-fail"><b id="c-fail">0</b><span>Failed to parse</span></div>
    </div>
    <div class="barwrap"><div class="bar" id="bar"></div></div>
    <div class="barlbl" id="barlbl">Waiting…</div>
    <div class="summary" id="summary"></div>
    <div class="btnrow" id="dlrow" style="display:none">
      <a class="btn" id="dl-portfolio">Portfolio report</a>
      <a class="btn sec" id="dl-all">Download all (zip)</a>
      <a class="btn sec" id="dl-exc">Exceptions only (zip)</a>
    </div>
    <div class="scroll"><table>
      <thead><tr><th>#</th><th>File</th><th>Result</th><th>Stage</th><th>Exceptions</th>
                 <th>Parse</th><th>Sanctioned</th><th>Download</th></tr></thead>
      <tbody id="rows"></tbody>
    </table></div>
  </div>
</div>
<script>
const inp=document.getElementById('pdfs'),folder=document.getElementById('folder'),
      drop=document.getElementById('drop'),filesDiv=document.getElementById('files'),
      startBtn=document.getElementById('start'),panel=document.getElementById('panel'),
      rows=document.getElementById('rows');
let chosen=[];

function setFiles(list){
  chosen=[...list].filter(f=>/\.(pdf|zip)$/i.test(f.name));
  filesDiv.innerHTML = chosen.length ? chosen.length+' file(s) selected' : '';
  startBtn.disabled = chosen.length===0;
}
inp.addEventListener('change',e=>setFiles(e.target.files));
folder.addEventListener('change',e=>setFiles(e.target.files));
document.getElementById('pickFolder').addEventListener('click',()=>folder.click());
document.getElementById('reset').addEventListener('click',()=>{location.reload();});
['dragover','dragenter'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('drag');}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('drag');}));
drop.addEventListener('drop',ev=>{ev.preventDefault();drop.classList.remove('drag');setFiles(ev.dataTransfer.files);});

let tot=0,clean=0,exc=0,fail=0,done=0;
function fmt(n){return n==null?'':Number(n).toLocaleString('en-IN');}

startBtn.addEventListener('click',async()=>{
  if(!chosen.length)return;
  startBtn.disabled=true; panel.style.display='block';
  rows.innerHTML=''; tot=clean=exc=fail=done=0;
  const fd=new FormData(); chosen.forEach(f=>fd.append('pdfs',f));
  document.getElementById('barlbl').textContent='Uploading…';
  let job;
  try{ const r=await fetch('/upload',{method:'POST',body:fd}); job=await r.json(); }
  catch(err){ document.getElementById('barlbl').textContent='Upload failed: '+err; return; }
  if(job.error){ document.getElementById('barlbl').textContent=job.error; return; }
  tot=job.total; document.getElementById('c-tot').textContent=tot;
  if(tot===0){ document.getElementById('barlbl').textContent='No PDF files found in the upload.'; return; }

  const es=new EventSource('/process_stream/'+job.job_id);
  es.onmessage=ev=>{
    const d=JSON.parse(ev.data);
    if(d.type==='progress'){
      done++;
      if(!d.ok) fail++; else if(d.exceptions>0) exc++; else clean++;
      document.getElementById('c-clean').textContent=clean;
      document.getElementById('c-exc').textContent=exc;
      document.getElementById('c-fail').textContent=fail;
      const pct=Math.round(done/tot*100);
      document.getElementById('bar').style.width=pct+'%';
      document.getElementById('barlbl').textContent='Processed '+done+' / '+tot+' ('+pct+'%)';
      let resTag,res;
      if(!d.ok){resTag='t-fail';res='FAILED';}
      else if(d.exceptions>0){resTag='t-exc';res='EXCEPTION';}
      else {resTag='t-clean';res='CLEAN';}
      const parseTag=d.parse==='REVIEW'?'t-rev':'t-ok';
      const dl=d.ok?'<a class="dl" href="/download/'+job.job_id+'/'+d.index+'">Excel</a>':'';
      rows.insertAdjacentHTML('beforeend',
        '<tr><td>'+(d.index+1)+'</td><td>'+d.file+'</td>'+
        '<td><span class="tag '+resTag+'">'+res+'</span></td>'+
        '<td>'+(d.stage||'')+'</td><td>'+(d.ok?d.exceptions:('—  '+(d.error||'')))+'</td>'+
        '<td>'+(d.ok?'<span class="tag '+parseTag+'">'+d.parse+'</span>':'')+'</td>'+
        '<td>'+(d.sanctioned!=null?fmt(d.sanctioned):'')+'</td><td>'+dl+'</td></tr>');
    } else if(d.type==='done'){
      es.close();
      const s=d.summary;
      document.getElementById('barlbl').textContent='Done — processed '+s.total+' file(s).';
      document.getElementById('summary').innerHTML=
        '<b>Portfolio health:</b> '+s.clean+' clean · '+s.exceptions+' with exceptions · '+
        s.failed+' failed · NPA: '+s.npa+' · Total sanctioned: ₹'+fmt(s.total_sanctioned);
      const dlrow=document.getElementById('dlrow'); dlrow.style.display='flex';
      document.getElementById('dl-portfolio').href='/download_portfolio/'+job.job_id;
      document.getElementById('dl-all').href='/download_zip/'+job.job_id+'?filter=all';
      document.getElementById('dl-exc').href='/download_zip/'+job.job_id+'?filter=exceptions';
      if(s.parsed<2) document.getElementById('dl-portfolio').style.display='none';
    }
  };
  es.onerror=()=>{ document.getElementById('barlbl').textContent='Connection lost during processing.'; es.close(); };
});
</script>
</body></html>
"""


@app.route("/")
def index():
    return render_template_string(PAGE)


@app.route("/upload", methods=["POST"])
def upload():
    uploads = request.files.getlist("pdfs")
    if not uploads:
        return jsonify(error="No files received."), 400
    _purge_old_jobs()
    job_id = uuid.uuid4().hex
    job_dir = tempfile.mkdtemp(prefix="soa_" + job_id + "_")
    pdfs = _collect_pdfs(uploads, job_dir)
    JOBS[job_id] = {"dir": job_dir, "pdfs": pdfs, "results": [],
                    "created": time.time(), "portfolio": None}
    return jsonify(job_id=job_id, total=len(pdfs))


@app.route("/process_stream/<job_id>")
def process_stream(job_id):
    job = JOBS.get(job_id)
    if not job:
        abort(404)

    def gen():
        results = job["results"]
        for i, pdf_path in enumerate(job["pdfs"]):
            display = os.path.basename(pdf_path)[4:]  # strip "NNN_" prefix
            rec = {"index": i, "file": display, "ok": False, "exceptions": 0,
                   "stage": "", "parse": "", "sanctioned": None, "error": "",
                   "xlsx": None, "audit": "FAILED"}
            try:
                d = extract_loan(pdf_path)
                xlsx_path = os.path.join(job["dir"], f"{i:03d}_{os.path.splitext(display)[0]}.xlsx")
                write_workbook(xlsx_path, d["master"], d["fin_rows"], d["recv"], d["disb"],
                               d["txns"], d["checks"], d["dpd_rows"], part_pay=d["part_pay"],
                               bounce=d["bounce"], charges=d["charges"], amort=d["amort"],
                               bounce_grid=d["bounce_grid"], quality=d["quality"])
                rec.update(
                    ok=True, exceptions=d["summary"]["Exceptions"],
                    stage=d["summary"]["Current Stage"], parse=d["quality"]["status"],
                    sanctioned=to_num(d["master"].get("Loan Amount (Rs)")),
                    xlsx=xlsx_path, data=d,
                    audit="EXCEPTION" if d["summary"]["Exceptions"] else "CLEAN")
            except Exception as e:
                rec["error"] = str(e)[:120]
            results.append(rec)
            payload = {k: rec[k] for k in
                       ("index", "file", "ok", "exceptions", "stage", "parse", "sanctioned", "error")}
            payload["type"] = "progress"
            yield f"data: {json.dumps(payload)}\n\n"

        # Build portfolio report for the successful loans
        good = [r for r in results if r["ok"]]
        if len(good) >= 2:
            ppath = os.path.join(job["dir"], "Portfolio_Exception_Report.xlsx")
            write_portfolio(ppath, [r["data"] for r in good])
            job["portfolio"] = ppath

        npa = sum(1 for r in good if "NPA" in (r["stage"] or ""))
        summary = {
            "total": len(results),
            "parsed": len(good),
            "failed": sum(1 for r in results if not r["ok"]),
            "clean": sum(1 for r in good if r["exceptions"] == 0),
            "exceptions": sum(1 for r in good if r["exceptions"] > 0),
            "npa": npa,
            "total_sanctioned": round(sum(r["sanctioned"] or 0 for r in good), 2),
        }
        yield f"data: {json.dumps({'type': 'done', 'summary': summary})}\n\n"

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
            if only_exc and r["exceptions"] == 0:
                continue
            z.write(r["xlsx"], os.path.join("loan_details",
                    os.path.splitext(r["file"])[0] + ".xlsx"))
    buf.seek(0)
    name = "SOA_exceptions.zip" if only_exc else "SOA_TOC_TOD_results.zip"
    return send_file(buf, as_attachment=True, mimetype="application/zip", download_name=name)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n  SOA TOC/TOD web UI  ->  http://127.0.0.1:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)

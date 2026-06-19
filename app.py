"""
Local web UI for the SOA TOC/TOD tool.

Run:
    pip install flask pdfplumber openpyxl
    python app.py            # then open http://127.0.0.1:5000

Upload one SOA PDF to get its working-paper workbook; upload several to get a
single Portfolio Exception Report (per-loan books are bundled into a zip).
"""
import io
import os
import zipfile
import tempfile

from flask import Flask, request, render_template_string, send_file, abort

from extract_soa import extract_loan, write_workbook, write_portfolio

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB per request

PAGE = """
<!doctype html>
<title>SOA TOC/TOD Extractor</title>
<style>
  body{font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:#f4f6f9;
       color:#1f2d3d;margin:0;padding:40px}
  .card{max-width:680px;margin:0 auto;background:#fff;border-radius:12px;
        box-shadow:0 4px 20px rgba(0,0,0,.08);padding:32px}
  h1{color:#1F4E78;margin:0 0 4px;font-size:22px}
  p.sub{color:#6b7280;margin:0 0 24px;font-size:14px}
  .drop{border:2px dashed #c3ccd6;border-radius:10px;padding:34px;text-align:center;
        background:#fafbfc;cursor:pointer}
  .drop.drag{border-color:#1F4E78;background:#eef4fb}
  input[type=file]{display:none}
  button{margin-top:20px;background:#1F4E78;color:#fff;border:0;border-radius:8px;
         padding:12px 22px;font-size:15px;cursor:pointer;width:100%}
  button:hover{background:#163a5c}
  .files{margin-top:14px;font-size:13px;color:#374151;text-align:left}
  .note{margin-top:22px;font-size:12px;color:#9097a1;line-height:1.5}
  .err{background:#FFC7CE;color:#7a1f2b;padding:10px 14px;border-radius:8px;
       margin-bottom:18px;font-size:14px}
</style>
<div class="card">
  <h1>SOA TOC/TOD Extractor</h1>
  <p class="sub">L&amp;T Finance NBFC Statement-of-Account &rarr; audit-ready Excel</p>
  {% if error %}<div class="err">{{ error }}</div>{% endif %}
  <form method="post" action="/process" enctype="multipart/form-data" id="f">
    <label class="drop" id="drop">
      <input type="file" name="pdfs" id="pdfs" accept=".pdf" multiple>
      <div id="label"><b>Click to choose</b> or drop SOA PDF(s) here</div>
      <div class="files" id="files"></div>
    </label>
    <button type="submit">Extract &amp; Validate</button>
  </form>
  <div class="note">
    1 PDF &rarr; full working-paper workbook (master, ledger, charges,
    amortization, DPD, TOC/TOD, parse-quality).<br>
    Multiple PDFs &rarr; Portfolio Exception Report + per-loan workbooks (zip).
  </div>
</div>
<script>
  const inp=document.getElementById('pdfs'),drop=document.getElementById('drop'),
        files=document.getElementById('files');
  function show(){files.innerHTML=[...inp.files].map(f=>'&bull; '+f.name).join('<br>')||'';}
  inp.addEventListener('change',show);
  ['dragover','dragenter'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('drag');}));
  ['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('drag');}));
  drop.addEventListener('drop',ev=>{inp.files=ev.dataTransfer.files;show();});
</script>
"""


@app.route("/")
def index():
    return render_template_string(PAGE, error=None)


@app.route("/process", methods=["POST"])
def process_uploads():
    uploads = [f for f in request.files.getlist("pdfs") if f and f.filename]
    if not uploads:
        return render_template_string(PAGE, error="Please choose at least one PDF.")

    with tempfile.TemporaryDirectory() as tmp:
        results, detail_paths = [], []
        for up in uploads:
            if not up.filename.lower().endswith(".pdf"):
                continue
            pdf_path = os.path.join(tmp, os.path.basename(up.filename))
            up.save(pdf_path)
            try:
                d = extract_loan(pdf_path)
            except Exception as e:
                return render_template_string(
                    PAGE, error=f"Failed to parse {up.filename}: {e}")
            xlsx_path = os.path.join(tmp, os.path.splitext(os.path.basename(up.filename))[0] + ".xlsx")
            write_workbook(xlsx_path, d["master"], d["fin_rows"], d["recv"], d["disb"],
                           d["txns"], d["checks"], d["dpd_rows"], part_pay=d["part_pay"],
                           bounce=d["bounce"], charges=d["charges"], amort=d["amort"],
                           bounce_grid=d["bounce_grid"], quality=d["quality"])
            results.append(d)
            detail_paths.append(xlsx_path)

        if not results:
            return render_template_string(PAGE, error="No valid PDF files found.")

        # Single loan -> return its workbook directly
        if len(results) == 1:
            return send_file(detail_paths[0], as_attachment=True,
                             download_name=os.path.basename(detail_paths[0]))

        # Multiple loans -> portfolio report + per-loan books, zipped
        report_path = os.path.join(tmp, "Portfolio_Exception_Report.xlsx")
        write_portfolio(report_path, results)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(report_path, "Portfolio_Exception_Report.xlsx")
            for p in detail_paths:
                z.write(p, os.path.join("loan_details", os.path.basename(p)))
        buf.seek(0)
        return send_file(buf, as_attachment=True, mimetype="application/zip",
                         download_name="SOA_TOC_TOD_results.zip")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n  SOA TOC/TOD web UI  ->  http://127.0.0.1:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)

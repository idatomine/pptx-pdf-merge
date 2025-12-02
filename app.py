#!/usr/bin/env python3
import subprocess
from pathlib import Path
import tempfile
from PyPDF2 import PdfMerger
from flask import Flask, request, send_file, render_template_string

SOFFICE_CMD = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

app = Flask(__name__)

HTML_FORM = """
<!doctype html>
<html>
  <head>
    <meta charset="utf-8">
    <title>PPTX → merged PDF</title>
    <style>
      body {
        font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        margin: 40px;
      }
      h1 {
        margin-bottom: 0.5rem;
      }
      button {
        padding: 0.5rem 1.2rem;
        font-size: 1rem;
        border-radius: 6px;
        border: none;
        cursor: pointer;
      }
      button:disabled {
        opacity: 0.6;
        cursor: not-allowed;
      }
      #loading-overlay {
        position: fixed;
        inset: 0;
        background: rgba(255, 255, 255, 0.8);
        display: none;
        align-items: center;
        justify-content: center;
        z-index: 9999;
        flex-direction: column;
        gap: 12px;
        font-size: 1.1rem;
      }
      .spinner {
        width: 36px;
        height: 36px;
        border-radius: 50%;
        border: 4px solid #ccc;
        border-top-color: #444;
        animation: spin 0.8s linear infinite;
      }
      @keyframes spin {
        to { transform: rotate(360deg); }
      }
    </style>
  </head>
  <body>
    <h1>🔥Merge PPTX til én PDF🔥</h1>
    <p>Velg PPTX-filene dine og trykk <strong>Merge!</strong> Nettleseren vil automatisk laste ned den ferdige PDF-en.</p>

    <form id="merge-form" enctype="multipart/form-data">
      <p><strong>Filer:</strong></p>
      <input type="file" name="files" multiple accept=".pptx">
      <br><br>
      <button type="submit" id="submit-btn">Merge!</button>
    </form>

    <div id="loading-overlay">
      <div class="spinner"></div>
      <div>Jobber med å konvertere og merge… 🤖</div>
    </div>

    <script>
      const form = document.getElementById('merge-form');
      const loading = document.getElementById('loading-overlay');
      const submitBtn = document.getElementById('submit-btn');
      const fileInput = form.querySelector('input[type="file"]');

      form.addEventListener('submit', async function (event) {
        event.preventDefault();  // Vi håndterer submit selv med fetch()

        if (!fileInput.files.length) {
          alert('Velg minst én PPTX-fil først.');
          return;
        }

        loading.style.display = 'flex';
        submitBtn.disabled = true;

        const formData = new FormData(form);

        try {
          const response = await fetch('/', {
            method: 'POST',
            body: formData
          });

          if (!response.ok) {
            const text = await response.text();
            throw new Error(text || 'Serverfeil');
          }

          // Få PDF-en som blob
          const blob = await response.blob();
          const url = window.URL.createObjectURL(blob);

          // Lag en midlertidig <a>-link for å trigge nedlasting
          const a = document.createElement('a');
          a.href = url;
          a.download = 'merged_presentations.pdf';
          document.body.appendChild(a);
          a.click();
          a.remove();
          window.URL.revokeObjectURL(url);

        } catch (err) {
          console.error(err);
          alert('Noe gikk galt under merging: ' + err.message);
        } finally {
          // Uansett suksess/feil: skjul loader, enable knapp
          loading.style.display = 'none';
          submitBtn.disabled = false;
        }
      });
    </script>
  </body>
</html>
"""

def convert_pptx_to_pdf(pptx_path: Path, pdf_dir: Path) -> Path:
    pdf_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        SOFFICE_CMD,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(pdf_dir),
        str(pptx_path),
    ]
    subprocess.run(cmd, check=True)
    pdf_path = pdf_dir / f"{pptx_path.stem}.pdf"
    if not pdf_path.exists():
        raise FileNotFoundError(f"Forventet PDF ikke funnet: {pdf_path}")
    return pdf_path

def merge_pdfs(pdf_files, output_path: Path):
    merger = PdfMerger()
    for pdf in pdf_files:
        merger.append(str(pdf))
    merger.write(str(output_path))
    merger.close()

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template_string(HTML_FORM)

    uploaded_files = request.files.getlist("files")
    if not uploaded_files or uploaded_files[0].filename == "":
        return "Ingen filer valgt", 400

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        pptx_dir = tmpdir / "pptx"
        pdf_dir = tmpdir / "pdf"
        pptx_dir.mkdir()
        pdf_dir.mkdir()

        pdf_files = []
        for f in uploaded_files:
            pptx_path = pptx_dir / f.filename
            f.save(pptx_path)
            pdf_path = convert_pptx_to_pdf(pptx_path, pdf_dir)
            pdf_files.append(pdf_path)

        merged_path = tmpdir / "merged.pdf"
        merge_pdfs(pdf_files, merged_path)

        # Flask svarer fortsatt med ren PDF – men nå leser JS det som blob
        return send_file(
            merged_path,
            as_attachment=True,
            download_name="merged_presentations.pdf",
        )

if __name__ == "__main__":
    app.run(debug=True)

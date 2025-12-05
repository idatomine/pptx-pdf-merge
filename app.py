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
<title>PPTX/PDF Merger</title>

<style>
  body { font-family: system-ui, sans-serif; margin: 40px; }
  h1 { margin-bottom: 0.5rem; }

  #instructions {
    font-size: 0.9rem;
    margin: 4px 0 14px 0;
    color: #666;
  }

  ul#file-list {
    padding: 0;
    list-style: none;
    width: 400px;
    margin-top: 8px;
  }
  ul#file-list li {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 8px 10px;
    background: #fafafa;
    border-radius: 6px;
    margin-bottom: 6px;
    cursor: grab;
    border: 1px solid #ddd;
    transition: background 0.15s;
  }
  ul#file-list li:hover { background: #f0f0f0; }
  ul#file-list li.dragging {
    opacity: 0.35;
    border-style: dashed;
  }
  .file-index {
    font-weight: 600;
    width: 30px;
    text-align: center;
    color: #444;
  }

  button {
    padding: 0.5rem 1.4rem;
    margin-top: 12px;
    font-size: 1rem;
    border-radius: 6px;
    border: none;
    background: #0078ff;
    color: white;
    cursor: pointer;
  }

  #loading-overlay {
    position: fixed; inset: 0; background: rgba(255,255,255,0.8);
    display: none; align-items: center; justify-content: center;
    z-index: 9999; flex-direction: column;
    gap: 12px; font-size: 1.1rem;
  }
  .spinner {
    width: 36px; height: 36px;
    border-radius: 50%;
    border: 4px solid #ccc;
    border-top-color: #444;
    animation: spin 0.8s linear infinite;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
</style>

</head>
<body>

<h1>🔥Merge PPTX & PDF🔥</h1>
<p>Upload multiple files, <strong>drag to reorder</strong>, then merge!</p>
<p id="instructions">Dra filene opp/ned for å justere rekkefølgen 🌸</p>

<input type="file" id="file-input" name="files" multiple accept=".pptx,.pdf">
<ul id="file-list"></ul>

<button id="submit-btn" style="display:none;">Merge!</button>

<div id="loading-overlay">
  <div class="spinner"></div>
  <div>Processing… Please wait 🤖</div>
</div>

<script>
const fileInput = document.getElementById("file-input");
const fileList = document.getElementById("file-list");
const submitBtn = document.getElementById("submit-btn");
let fileArray = [];

fileInput.addEventListener("change", () => {
  const newFiles = Array.from(fileInput.files);

  newFiles.forEach(f => {
    if (!fileArray.some(existing => existing.name === f.name && existing.size === f.size)) {
      fileArray.push(f);
    }
  });

  fileInput.value = "";
  renderList();
  submitBtn.style.display = "inline-block";
});

function renderList() {
  fileList.innerHTML = "";
  fileArray.forEach((file, index) => {
    const li = document.createElement("li");
    li.draggable = true;
    li.dataset.index = index;

    const idx = document.createElement("span");
    idx.className = "file-index";
    idx.textContent = index + 1;

    const name = document.createElement("span");
    name.textContent = file.name;

    li.appendChild(idx);
    li.appendChild(name);

    addDragHandlers(li);
    fileList.appendChild(li);
  });
}

function updateIndexes() {
  [...fileList.children].forEach((li, i) => {
    li.dataset.index = i;
    li.querySelector(".file-index").textContent = i + 1;
  });
}

function addDragHandlers(el) {
  el.addEventListener("dragstart", (e) => {
    el.classList.add("dragging");
  });

  el.addEventListener("dragend", () => {
    el.classList.remove("dragging");
    updateIndexes();
  });

  el.addEventListener("dragover", (e) => {
    e.preventDefault();
    const dragging = document.querySelector(".dragging");
    const siblings = [...fileList.querySelectorAll("li:not(.dragging)")];
    const nextSibling = siblings.find(s => e.clientY <= s.offsetTop + s.offsetHeight / 2);
    fileList.insertBefore(dragging, nextSibling);
  });
}

submitBtn.addEventListener("click", async () => {
  const reordered = [...fileList.children].map(li => fileArray[li.dataset.index]);

  const formData = new FormData();
  reordered.forEach(f => formData.append("files", f));

  document.getElementById("loading-overlay").style.display = "flex";
  submitBtn.disabled = true;

  try {
    const response = await fetch("/", { method: "POST", body: formData });
    if (!response.ok) throw new Error(await response.text());

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "merged.pdf";
    a.click();
    window.URL.revokeObjectURL(url);
  } catch (err) {
    alert("Error: " + err.message);
  } finally {
    document.getElementById("loading-overlay").style.display = "none";
    submitBtn.disabled = false;
  }
});
</script>

</body>
</html>
"""



def convert_pptx_to_pdf(pptx_path: Path, pdf_dir: Path) -> Path:
    """
    Convert PPTX → PDF using LibreOffice.
    """
    pdf_dir.mkdir(parents=True, exist_ok=True)
    cmd = [SOFFICE_CMD, "--headless", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(pptx_path)]
    subprocess.run(cmd, check=True)
    pdf_path = pdf_dir / f"{pptx_path.stem}.pdf"
    if not pdf_path.exists():
        raise FileNotFoundError(f"Pdf missing: {pdf_path}")
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
        return "No files uploaded", 400

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        work_dir = tmpdir / "upload"
        pdf_dir = tmpdir / "pdf"
        work_dir.mkdir()
        pdf_dir.mkdir()

        pdf_files = []

        for f in uploaded_files:
            filename = f.filename.lower()
            temp_path = work_dir / f.filename
            f.save(temp_path)

            if filename.endswith(".pptx"):
                pdf_files.append(convert_pptx_to_pdf(temp_path, pdf_dir))
            elif filename.endswith(".pdf"):
                pdf_files.append(temp_path)  # Already PDF
            else:
                return f"Unsupported file type: {filename}", 400

        merged_path = tmpdir / "merged.pdf"
        merge_pdfs(pdf_files, merged_path)

        return send_file(merged_path, as_attachment=True, download_name="merged.pdf")

if __name__ == "__main__":
    app.run(debug=True)

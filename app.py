#!/usr/bin/env python3
import subprocess
from pathlib import Path
from io import BytesIO
import tempfile
from PyPDF2 import PdfMerger
from flask import Flask, request, send_file, render_template_string

SOFFICE_CMD = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

app = Flask(__name__)

HTML_FORM = """
<!doctype html>
<html lang="no">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PPTX/PDF Merger</title>

<style>
  :root {
    --accent: #6d5efc;
    --accent-2: #b06dfc;
    --bg-card: rgba(255, 255, 255, 0.92);
    --text: #2b2b3a;
    --muted: #7a7a8c;
    --line: #e7e5f2;
  }

  * { box-sizing: border-box; }

  body {
    font-family: "Segoe UI", system-ui, -apple-system, sans-serif;
    margin: 0;
    min-height: 100vh;
    color: var(--text);
    background: linear-gradient(135deg, #f6d5f7 0%, #c9d6ff 50%, #d4f4ec 100%);
    display: flex;
    justify-content: center;
    align-items: flex-start;
    padding: 48px 16px;
  }

  .card {
    width: 100%;
    max-width: 560px;
    background: var(--bg-card);
    backdrop-filter: blur(8px);
    border-radius: 20px;
    box-shadow: 0 18px 50px rgba(80, 60, 160, 0.18);
    padding: 32px 32px 36px;
  }

  h1 {
    margin: 0 0 4px;
    font-size: 1.7rem;
    background: linear-gradient(90deg, var(--accent), var(--accent-2));
    -webkit-background-clip: text;
    background-clip: text;
    color: transparent;
  }
  .subtitle { margin: 0 0 22px; color: var(--muted); font-size: 0.95rem; }

  /* ---- drop zone ---- */
  .dropzone {
    border: 2px dashed #cfc9ec;
    border-radius: 14px;
    padding: 28px 18px;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.15s, background 0.15s, transform 0.1s;
    background: rgba(255, 255, 255, 0.5);
  }
  .dropzone:hover { border-color: var(--accent); background: rgba(109, 94, 252, 0.06); }
  .dropzone.dragover {
    border-color: var(--accent);
    background: rgba(109, 94, 252, 0.12);
    transform: scale(1.01);
  }
  .dropzone .icon { font-size: 2rem; }
  .dropzone .big { font-weight: 600; margin-top: 6px; }
  .dropzone .small { color: var(--muted); font-size: 0.85rem; margin-top: 2px; }
  #file-input { display: none; }

  /* ---- file list ---- */
  ul#file-list { padding: 0; list-style: none; margin: 20px 0 0; }
  ul#file-list li {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 10px 12px;
    background: #fff;
    border-radius: 12px;
    margin-bottom: 8px;
    border: 1px solid var(--line);
    box-shadow: 0 2px 6px rgba(80, 60, 160, 0.05);
    transition: box-shadow 0.15s, transform 0.1s, opacity 0.1s;
  }
  ul#file-list li.dragging {
    opacity: 0.4;
    border-style: dashed;
    box-shadow: 0 8px 20px rgba(80, 60, 160, 0.18);
  }

  .handle {
    cursor: grab;
    color: #b9b4d6;
    font-size: 1.1rem;
    user-select: none;
    padding: 0 2px;
    line-height: 1;
  }
  .handle:active { cursor: grabbing; }

  .file-index {
    font-weight: 700;
    min-width: 24px;
    height: 24px;
    display: flex;
    align-items: center;
    justify-content: center;
    border-radius: 50%;
    background: linear-gradient(135deg, var(--accent), var(--accent-2));
    color: #fff;
    font-size: 0.8rem;
  }

  .file-meta { flex: 1; min-width: 0; }
  .file-name {
    font-size: 0.92rem;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .file-type { font-size: 0.72rem; color: var(--muted); text-transform: uppercase; }

  .row-actions { display: flex; gap: 4px; }
  .icon-btn {
    border: none;
    background: #f3f1fb;
    color: #6a6585;
    width: 30px;
    height: 30px;
    border-radius: 8px;
    cursor: pointer;
    font-size: 0.9rem;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: background 0.12s, color 0.12s, transform 0.08s;
    padding: 0;
  }
  .icon-btn:hover { background: var(--accent); color: #fff; }
  .icon-btn:active { transform: scale(0.9); }
  .icon-btn:disabled { opacity: 0.3; cursor: default; background: #f3f1fb; color: #6a6585; }
  .icon-btn.danger:hover { background: #ff5d73; }

  /* ---- merge button ---- */
  #submit-btn {
    width: 100%;
    margin-top: 18px;
    padding: 0.85rem 1.4rem;
    font-size: 1.05rem;
    font-weight: 600;
    border-radius: 12px;
    border: none;
    color: white;
    cursor: pointer;
    background: linear-gradient(90deg, var(--accent), var(--accent-2));
    box-shadow: 0 8px 20px rgba(109, 94, 252, 0.35);
    transition: transform 0.1s, box-shadow 0.15s, opacity 0.15s;
  }
  #submit-btn:hover { transform: translateY(-1px); box-shadow: 0 12px 26px rgba(109, 94, 252, 0.45); }
  #submit-btn:disabled { opacity: 0.6; cursor: default; transform: none; }

  /* ---- loading overlay ---- */
  #loading-overlay {
    position: fixed; inset: 0; background: rgba(255,255,255,0.78);
    backdrop-filter: blur(3px);
    display: none; align-items: center; justify-content: center;
    z-index: 9999; flex-direction: column;
    gap: 14px; font-size: 1.05rem; color: var(--text);
  }
  .spinner {
    width: 40px; height: 40px;
    border-radius: 50%;
    border: 4px solid #e0dcf5;
    border-top-color: var(--accent);
    animation: spin 0.8s linear infinite;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* ---- preview modal ---- */
  #preview-modal {
    position: fixed; inset: 0; z-index: 10000;
    display: none;
    background: rgba(40, 30, 70, 0.55);
    backdrop-filter: blur(2px);
    align-items: center; justify-content: center;
    padding: 24px;
  }
  .preview-panel {
    background: #fff;
    border-radius: 16px;
    width: min(900px, 100%);
    height: min(88vh, 100%);
    display: flex;
    flex-direction: column;
    overflow: hidden;
    box-shadow: 0 24px 60px rgba(40, 30, 70, 0.4);
  }
  .preview-header {
    display: flex; align-items: center; gap: 10px;
    padding: 14px 18px;
    border-bottom: 1px solid var(--line);
  }
  .preview-header .title {
    flex: 1; font-weight: 600; font-size: 0.95rem;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }
  .preview-body { flex: 1; position: relative; background: #f4f3fa; }
  .preview-body iframe { width: 100%; height: 100%; border: none; }
  .preview-loading {
    position: absolute; inset: 0;
    display: flex; flex-direction: column; gap: 12px;
    align-items: center; justify-content: center;
    color: var(--muted);
  }
</style>

</head>
<body>

<div class="card">
  <h1>🔥 Merge PPTX &amp; PDF 🔥</h1>
  <p class="subtitle">Last opp filer, dra eller bruk pilene for å sortere, forhåndsvis, og merge! 🌸</p>

  <label class="dropzone" id="dropzone" for="file-input">
    <div class="icon">📂</div>
    <div class="big">Klikk for å velge filer</div>
    <div class="small">eller dra dem hit · .pptx og .pdf</div>
  </label>
  <input type="file" id="file-input" name="files" multiple accept=".pptx,.pdf">

  <ul id="file-list"></ul>

  <button id="submit-btn" style="display:none;">✨ Merge til én PDF</button>
</div>

<div id="loading-overlay">
  <div class="spinner"></div>
  <div>Behandler… vent litt 🤖</div>
</div>

<div id="preview-modal">
  <div class="preview-panel">
    <div class="preview-header">
      <span class="title" id="preview-title">Forhåndsvisning</span>
      <button class="icon-btn" id="preview-close" title="Lukk">✕</button>
    </div>
    <div class="preview-body">
      <div class="preview-loading" id="preview-loading">
        <div class="spinner"></div>
        <div>Lager forhåndsvisning…</div>
      </div>
      <iframe id="preview-frame" style="display:none;"></iframe>
    </div>
  </div>
</div>

<script>
const fileInput = document.getElementById("file-input");
const fileList = document.getElementById("file-list");
const submitBtn = document.getElementById("submit-btn");
const dropzone = document.getElementById("dropzone");
let fileArray = [];
let previewUrl = null;

/* ---------- adding files ---------- */
function addFiles(files) {
  Array.from(files).forEach(f => {
    const name = f.name.toLowerCase();
    if (!name.endsWith(".pptx") && !name.endsWith(".pdf")) return;
    if (!fileArray.some(e => e.name === f.name && e.size === f.size)) {
      fileArray.push(f);
    }
  });
  renderList();
  submitBtn.style.display = fileArray.length ? "block" : "none";
}

fileInput.addEventListener("change", () => {
  addFiles(fileInput.files);
  fileInput.value = "";
});

["dragenter", "dragover"].forEach(ev =>
  dropzone.addEventListener(ev, e => { e.preventDefault(); dropzone.classList.add("dragover"); })
);
["dragleave", "drop"].forEach(ev =>
  dropzone.addEventListener(ev, e => { e.preventDefault(); dropzone.classList.remove("dragover"); })
);
dropzone.addEventListener("drop", e => {
  if (e.dataTransfer && e.dataTransfer.files) addFiles(e.dataTransfer.files);
});

/* ---------- rendering ---------- */
function renderList() {
  fileList.innerHTML = "";
  fileArray.forEach((file, index) => {
    const li = document.createElement("li");
    li.draggable = true;

    const handle = document.createElement("span");
    handle.className = "handle";
    handle.textContent = "⠿";
    handle.title = "Dra for å sortere";

    const idx = document.createElement("span");
    idx.className = "file-index";
    idx.textContent = index + 1;

    const meta = document.createElement("div");
    meta.className = "file-meta";
    const name = document.createElement("div");
    name.className = "file-name";
    name.textContent = file.name;
    name.title = file.name;
    const type = document.createElement("div");
    type.className = "file-type";
    type.textContent = file.name.split(".").pop();
    meta.appendChild(name);
    meta.appendChild(type);

    const actions = document.createElement("div");
    actions.className = "row-actions";

    const upBtn = iconButton("▲", "Flytt opp", () => move(index, index - 1));
    upBtn.disabled = index === 0;
    const downBtn = iconButton("▼", "Flytt ned", () => move(index, index + 1));
    downBtn.disabled = index === fileArray.length - 1;
    const previewBtn = iconButton("👁", "Forhåndsvis", () => openPreview(file));
    const removeBtn = iconButton("✕", "Fjern", () => removeFile(index));
    removeBtn.classList.add("danger");

    actions.append(upBtn, downBtn, previewBtn, removeBtn);
    li.append(handle, idx, meta, actions);
    addDragHandlers(li, index);
    fileList.appendChild(li);
  });
}

function iconButton(label, title, onClick) {
  const b = document.createElement("button");
  b.className = "icon-btn";
  b.type = "button";
  b.textContent = label;
  b.title = title;
  b.addEventListener("click", onClick);
  return b;
}

/* ---------- reordering ---------- */
function move(from, to) {
  if (to < 0 || to >= fileArray.length) return;
  const [item] = fileArray.splice(from, 1);
  fileArray.splice(to, 0, item);
  renderList();
}

function removeFile(index) {
  fileArray.splice(index, 1);
  renderList();
  submitBtn.style.display = fileArray.length ? "block" : "none";
}

let dragIndex = null;
function addDragHandlers(el, index) {
  el.addEventListener("dragstart", () => { dragIndex = index; el.classList.add("dragging"); });
  el.addEventListener("dragend", () => el.classList.remove("dragging"));
  el.addEventListener("dragover", e => {
    e.preventDefault();
    const rows = [...fileList.querySelectorAll("li")];
    const overIndex = rows.indexOf(el);
    if (overIndex === dragIndex || dragIndex === null) return;
    move(dragIndex, overIndex);
    dragIndex = overIndex;
  });
}

/* ---------- preview ---------- */
const previewModal = document.getElementById("preview-modal");
const previewFrame = document.getElementById("preview-frame");
const previewLoading = document.getElementById("preview-loading");
const previewTitle = document.getElementById("preview-title");

async function openPreview(file) {
  previewTitle.textContent = file.name;
  previewFrame.style.display = "none";
  previewLoading.style.display = "flex";
  previewModal.style.display = "flex";

  try {
    const fd = new FormData();
    fd.append("file", file);
    const res = await fetch("/preview", { method: "POST", body: fd });
    if (!res.ok) throw new Error(await res.text());
    const blob = await res.blob();
    if (previewUrl) URL.revokeObjectURL(previewUrl);
    previewUrl = URL.createObjectURL(blob);
    previewFrame.src = previewUrl;
    previewFrame.style.display = "block";
    previewLoading.style.display = "none";
  } catch (err) {
    previewLoading.innerHTML = "<div>😢 Kunne ikke lage forhåndsvisning:<br>" + err.message + "</div>";
  }
}

function closePreview() {
  previewModal.style.display = "none";
  previewFrame.src = "about:blank";
  if (previewUrl) { URL.revokeObjectURL(previewUrl); previewUrl = null; }
  previewLoading.innerHTML = '<div class="spinner"></div><div>Lager forhåndsvisning…</div>';
}
document.getElementById("preview-close").addEventListener("click", closePreview);
previewModal.addEventListener("click", e => { if (e.target === previewModal) closePreview(); });
document.addEventListener("keydown", e => { if (e.key === "Escape") closePreview(); });

/* ---------- merge ---------- */
submitBtn.addEventListener("click", async () => {
  const formData = new FormData();
  fileArray.forEach(f => formData.append("files", f));

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

@app.route("/preview", methods=["POST"])
def preview():
    """Convert a single uploaded file to PDF and return it inline for preview."""
    f = request.files.get("file")
    if not f or f.filename == "":
        return "No file uploaded", 400

    filename = f.filename.lower()
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        src = tmpdir / f.filename
        f.save(src)

        if filename.endswith(".pptx"):
            pdf_path = convert_pptx_to_pdf(src, tmpdir / "pdf")
        elif filename.endswith(".pdf"):
            pdf_path = src
        else:
            return f"Unsupported file type: {filename}", 400

        # Read into memory so the file survives the temp-dir cleanup.
        data = pdf_path.read_bytes()

    return send_file(
        BytesIO(data),
        mimetype="application/pdf",
        as_attachment=False,
        download_name="preview.pdf",
    )

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
    app.run(debug=True, port=5001)

#!/usr/bin/env python3
import subprocess
from pathlib import Path
import re
from PyPDF2 import PdfMerger


# ---- Konfig ----

# Rotmappe = mappen der dette scriptet ligger
BASE_DIR = Path(__file__).resolve().parent

# Her ligger PPTX-filene dine
PPTX_DIR = BASE_DIR / "presentations"

# Midlertidig mappe for PDF-ene
PDF_DIR = BASE_DIR / "_pdf_temp"

# Output-fil (samlet PDF)
OUTPUT_PDF = BASE_DIR / "merged_presentations.pdf"

# Sti til LibreOffice-kommandolinje
SOFFICE_CMD = "/Applications/LibreOffice.app/Contents/MacOS/soffice"


# ---- Hjelpefunksjoner ----

def natural_sort_key(path: Path):
    """
    Sorterer etter tallet i starten av filnavnet: '1 foo.pptx', '2 foo.pptx', ..., '10 foo.pptx'
    Så rekkefølgen blir 1,2,3,...,10 (ikke 1,10,2,...)
    """
    m = re.match(r"(\d+)", path.name)
    if m:
        return int(m.group(1))
    return 999999  # filer uten tall havner til slutt


def convert_pptx_to_pdf(pptx_path: Path, pdf_dir: Path) -> Path:
    """
    Konverterer én PPTX til PDF via LibreOffice (soffice).
    Returnerer sti til generert PDF.
    """
    pdf_dir.mkdir(parents=True, exist_ok=True)

    cmd = [
        SOFFICE_CMD,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(pdf_dir),
        str(pptx_path),
    ]

    print(f"Konverterer: {pptx_path.name} -> PDF ...")
    subprocess.run(cmd, check=True)

    pdf_path = pdf_dir / f"{pptx_path.stem}.pdf"
    if not pdf_path.exists():
        raise FileNotFoundError(f"Forventet PDF ikke funnet: {pdf_path}")
    return pdf_path


def merge_pdfs(pdf_files, output_path: Path):
    """
    Merger en liste med PDF-filer til én PDF.
    """
    merger = PdfMerger()
    for pdf in pdf_files:
        print(f"Legger til i merge: {pdf.name}")
        merger.append(str(pdf))
    merger.write(str(output_path))
    merger.close()


def main():
    if not PPTX_DIR.is_dir():
        print(f"Fant ikke mappen med presentasjoner: {PPTX_DIR}")
        return

    pptx_files = sorted(PPTX_DIR.glob("*.pptx"), key=natural_sort_key)
    if not pptx_files:
        print(f"Ingen .pptx-filer funnet i {PPTX_DIR}")
        return

    print(f"Fant {len(pptx_files)} PPTX-filer:")
    for p in pptx_files:
        print("  -", p.name)

    PDF_DIR.mkdir(exist_ok=True)
    pdf_files = []

    try:
        # 1) PPTX -> PDF
        for pptx in pptx_files:
            pdf_path = convert_pptx_to_pdf(pptx, PDF_DIR)
            pdf_files.append(pdf_path)

        # 2) Merge alle PDF-ene til én fil
        print(f"\nMerger {len(pdf_files)} PDF-er til: {OUTPUT_PDF.name}")
        merge_pdfs(pdf_files, OUTPUT_PDF)

        print("\n✅ Ferdig!")
        print("Samlet PDF:", OUTPUT_PDF)

    except subprocess.CalledProcessError as e:
        print("\n❌ Feil ved konvertering med LibreOffice (soffice).")
        print("Sjekk at LibreOffice er installert og at 'soffice' fungerer i terminalen.")
        print("Detaljer:", e)


if __name__ == "__main__":
    main()

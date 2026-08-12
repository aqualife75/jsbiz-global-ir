#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cross-platform visual QA renderer for the built deck.

Fallback chain (auto-detected):
  1. Windows + MS Office : PowerPoint COM (true rendering - most accurate)
  2. LibreOffice present : soffice -> PDF -> PyMuPDF page PNGs
  3. Neither             : exits with instructions (text-QA only via markitdown)

Usage:
    python render_qa.py deck.pptx -o qa_png [--pdf] [--slides 2 4 7]
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile

PS_TEMPLATE = r"""
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open("{src}", $true, $false, $false)
{exports}
{pdf}
$pres.Close(); $ppt.Quit()
Write-Output "COM-OK"
"""


def find_soffice():
    cand = [shutil.which("soffice"),
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]
    return next((c for c in cand if c and os.path.exists(c)), None)


def render_com(src, out, slides, want_pdf):
    exports = []
    if slides:
        for i in slides:
            exports.append(f'$pres.Slides.Item({i}).Export("{out}\\slide-{i:02d}.png", "PNG", 1600, 900)')
    else:
        exports.append('$i = 1; foreach ($s in $pres.Slides) { '
                       f'$s.Export("{out}\\slide-$(\'{{0:d2}}\' -f $i).png", "PNG", 1600, 900); $i++ }}')
    pdf = f'$pres.SaveAs("{os.path.splitext(src)[0]}.pdf", 32)' if want_pdf else ""
    script = PS_TEMPLATE.format(src=src, exports="\n".join(exports), pdf=pdf)
    ps1 = os.path.join(tempfile.gettempdir(), "jsbiz_render_qa.ps1")
    with open(ps1, "w", encoding="utf-8-sig") as f:
        f.write(script)
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1],
                       capture_output=True, text=True, timeout=300)
    return "COM-OK" in (r.stdout or "")


def render_soffice(soffice, src, out, want_pdf):
    import fitz  # PyMuPDF renders the intermediate PDF
    with tempfile.TemporaryDirectory() as td:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                        "--outdir", td, src], check=True, timeout=300,
                       capture_output=True)
        pdfs = glob.glob(os.path.join(td, "*.pdf"))
        if not pdfs:
            return False
        doc = fitz.open(pdfs[0])
        for pno in range(len(doc)):
            doc[pno].get_pixmap(dpi=130).save(os.path.join(out, f"slide-{pno + 1:02d}.png"))
        if want_pdf:
            shutil.copy(pdfs[0], os.path.splitext(src)[0] + ".pdf")
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out", default="qa_png")
    ap.add_argument("--pdf", action="store_true", help="also export a PDF next to the pptx")
    ap.add_argument("--slides", type=int, nargs="*", help="re-render only these slide numbers")
    args = ap.parse_args()
    src = os.path.abspath(args.pptx)
    out = os.path.abspath(args.out)
    os.makedirs(out, exist_ok=True)

    if sys.platform == "win32":
        try:
            if render_com(src, out, args.slides, args.pdf):
                print(f"rendered via PowerPoint COM -> {out}")
                return
        except Exception as e:  # noqa: BLE001
            print(f"PowerPoint COM failed ({e}); trying LibreOffice...")

    soffice = find_soffice()
    if soffice:
        if render_soffice(soffice, src, out, args.pdf):
            print(f"rendered via LibreOffice -> {out}")
            return

    print("No renderer available (no MS Office COM, no LibreOffice).\n"
          "Fallback: run `markitdown deck.pptx` for text QA and open the deck\n"
          "in PowerPoint/Keynote manually for visual checks.")
    sys.exit(2)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extract reusable image assets from a source IR deck (PDF or PPTX).

- PDF : PyMuPDF embedded-image extraction (deduped by xref, >= min size)
- PPTX: ppt/media/* from the zip archive (>= min size)
- Both: labeled contact sheet for vision-based identification + manifest.json

Usage:
    python extract_assets.py deck.pdf  -o assets_out [--min 120] [--render-pages 110]
    python extract_assets.py deck.pptx -o assets_out
"""
import argparse
import io
import json
import os
import zipfile

from PIL import Image, ImageDraw


def extract_pdf(src, out, min_px):
    import fitz  # PyMuPDF
    doc = fitz.open(src)
    seen, items = set(), []
    for pno in range(len(doc)):
        for info in doc[pno].get_images(full=True):
            xref = info[0]
            if xref in seen:
                continue
            seen.add(xref)
            try:
                pix = fitz.Pixmap(doc, xref)
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                if pix.width < min_px or pix.height < min_px:
                    continue
                fn = f"p{pno + 1:02d}_x{xref}_{pix.width}x{pix.height}.png"
                pix.save(os.path.join(out, fn))
                items.append({"file": fn, "page": pno + 1, "w": pix.width, "h": pix.height})
            except Exception as e:  # noqa: BLE001 - skip unreadable images, keep going
                items.append({"xref": xref, "page": pno + 1, "error": str(e)[:80]})
    return items


def render_pdf_pages(src, out, dpi):
    import fitz
    doc = fitz.open(src)
    pages = []
    for pno in range(len(doc)):
        pix = doc[pno].get_pixmap(dpi=dpi)
        fn = f"page-{pno + 1:02d}.png"
        pix.save(os.path.join(out, fn))
        pages.append(fn)
    return pages


def extract_pptx(src, out, min_px):
    items = []
    with zipfile.ZipFile(src) as z:
        for name in z.namelist():
            if not name.startswith("ppt/media/"):
                continue
            data = z.read(name)
            base = os.path.basename(name)
            try:
                img = Image.open(io.BytesIO(data))
                if img.width < min_px or img.height < min_px:
                    continue
                fn = f"{os.path.splitext(base)[0]}_{img.width}x{img.height}.png"
                img.convert("RGBA").save(os.path.join(out, fn))
                items.append({"file": fn, "media": base, "w": img.width, "h": img.height})
            except Exception:  # noqa: BLE001 - wmf/emf etc: keep raw copy for reference
                raw = os.path.join(out, base)
                with open(raw, "wb") as f:
                    f.write(data)
                items.append({"file": base, "media": base, "note": "unconverted (wmf/emf?)"})
    return items


def contact_sheet(out, items, cols=6):
    files = [i["file"] for i in items if "error" not in i and i["file"].endswith(".png")]
    if not files:
        return None
    cell_w, cell_h, th_w, th_h = 270, 220, 240, 180
    rows = (len(files) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * cell_w, rows * cell_h), "white")
    d = ImageDraw.Draw(sheet)
    for idx, f in enumerate(files):
        im = Image.open(os.path.join(out, f)).convert("RGB")
        im.thumbnail((th_w, th_h))
        x, y = (idx % cols) * cell_w + 10, (idx // cols) * cell_h + 5
        sheet.paste(im, (x, y))
        d.text((x, y + th_h + 5), f[:34], fill="black")
    path = os.path.join(out, "_contact_sheet.png")
    sheet.save(path)
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--min", type=int, default=120, help="min px on both sides")
    ap.add_argument("--render-pages", type=int, metavar="DPI", default=0,
                    help="PDF only: also render full pages at this DPI")
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)

    ext = os.path.splitext(args.src)[1].lower()
    if ext == ".pdf":
        items = extract_pdf(args.src, args.out, args.min)
        pages = render_pdf_pages(args.src, args.out, args.render_pages) if args.render_pages else []
    elif ext == ".pptx":
        items = extract_pptx(args.src, args.out, args.min)
        pages = []
    else:
        raise SystemExit("source must be .pdf or .pptx")

    sheet = contact_sheet(args.out, items)
    manifest = {"source": os.path.abspath(args.src), "items": items,
                "pages": pages, "contact_sheet": sheet}
    with open(os.path.join(args.out, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    ok = len([i for i in items if "error" not in i])
    print(f"extracted {ok} images -> {args.out}")
    if sheet:
        print(f"contact sheet: {sheet}  (view this to identify assets)")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Extract everything useful from an assignment PDF.

Writes:
  <out>/task_text.md        text of every page, under "## Page N" headings
  <out>/pages/page-*.png    a rendered image of every page (for figures, tables, scans)

Prints a short report: page count, characters per page, pages that look scanned
or contain figures, and any fillable form fields.

Works with whichever PDF library is installed: PyMuPDF (fitz), pdfplumber, pypdf,
or the pdftotext/pdftoppm command line tools.

Usage:
  python3 extract_pdf.py task.pdf --out work/ [--dpi 130] [--max-pages 40]
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

Extraction = tuple[list[str], list[int], list[str], int]  # texts, figure pages, form fields, page count

SCANNED_THRESHOLD = 40  # fewer characters than this on a page = probably an image/scan


def extract_with_pymupdf(pdf_path: Path, pages_dir: Path, dpi: int, max_pages: int) -> Extraction:
    import fitz  # PyMuPDF

    doc = fitz.open(pdf_path)
    texts, figure_pages, fields = [], [], []
    zoom = dpi / 72
    for index, page in enumerate(doc):
        if index >= max_pages:
            break
        texts.append(page.get_text("text"))
        if page.get_images(full=True) or page.get_drawings():
            figure_pages.append(index + 1)
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        pix.save(str(pages_dir / f"page-{index + 1:02d}.png"))
        for widget in page.widgets() or []:
            fields.append(f"p{index + 1}: {widget.field_name} ({widget.field_type_string})")
    return texts, figure_pages, fields, len(doc)


def extract_with_pdfplumber(pdf_path: Path, pages_dir: Path, dpi: int, max_pages: int) -> Extraction:
    import pdfplumber

    texts, figure_pages, render_failed = [], [], False
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for index, page in enumerate(pdf.pages):
            if index >= max_pages:
                break
            texts.append(page.extract_text() or "")
            if page.images or page.rects or page.curves:
                figure_pages.append(index + 1)
            if not render_failed:
                try:
                    page.to_image(resolution=dpi).save(str(pages_dir / f"page-{index + 1:02d}.png"))
                except Exception:  # rendering needs extra deps; fall back to pdftoppm below
                    render_failed = True
    if render_failed:
        clear_page_images(pages_dir)
        render_with_pdftoppm(pdf_path, pages_dir, dpi, max_pages)
    return texts, figure_pages, [], total


def extract_with_pypdf(pdf_path: Path, pages_dir: Path, dpi: int, max_pages: int) -> Extraction:
    from pypdf import PdfReader

    reader = PdfReader(pdf_path)
    texts, figure_pages = [], []
    for index, page in enumerate(reader.pages):
        if index >= max_pages:
            break
        texts.append(page.extract_text() or "")
        try:
            if page.images:
                figure_pages.append(index + 1)
        except Exception:
            pass
    fields = [f"{name}" for name in (reader.get_fields() or {})]
    render_with_pdftoppm(pdf_path, pages_dir, dpi, max_pages)
    return texts, figure_pages, fields, len(reader.pages)


def extract_with_cli(pdf_path: Path, pages_dir: Path, dpi: int, max_pages: int) -> Extraction:
    if not shutil.which("pdftotext"):
        raise RuntimeError("no PDF library or pdftotext available")
    raw = subprocess.run(
        ["pdftotext", "-layout", "-l", str(max_pages), str(pdf_path), "-"],
        capture_output=True, text=True, check=True,
    ).stdout
    texts = raw.split("\f")
    if texts and not texts[-1].strip():
        texts = texts[:-1]
    render_with_pdftoppm(pdf_path, pages_dir, dpi, max_pages)
    return texts, [], [], len(texts)


def clear_page_images(pages_dir: Path) -> None:
    """Remove page images from an earlier task so they can't be mistaken for this one."""
    for old in pages_dir.glob("page*.png"):
        old.unlink()


def render_with_pdftoppm(pdf_path: Path, pages_dir: Path, dpi: int, max_pages: int) -> None:
    if not shutil.which("pdftoppm"):
        print("  (pdftoppm not found, page images skipped)", file=sys.stderr)
        return
    subprocess.run(
        ["pdftoppm", "-png", "-r", str(dpi), "-l", str(max_pages), str(pdf_path), str(pages_dir / "page")],
        check=False,
    )


BACKENDS = [
    ("PyMuPDF", extract_with_pymupdf),
    ("pdfplumber", extract_with_pdfplumber),
    ("pypdf", extract_with_pypdf),
    ("pdftotext", extract_with_cli),
]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--out", type=Path, default=Path("work"))
    parser.add_argument("--dpi", type=int, default=130)
    parser.add_argument("--max-pages", type=int, default=40)
    args = parser.parse_args()

    if not args.pdf.is_file():
        sys.exit(f"File not found: {args.pdf}")

    pages_dir = args.out / "pages"
    pages_dir.mkdir(parents=True, exist_ok=True)
    clear_page_images(pages_dir)

    result, used, errors = None, None, []
    for name, backend in BACKENDS:
        try:
            result = backend(args.pdf, pages_dir, args.dpi, args.max_pages)
            used = name
            break
        except ImportError:
            errors.append(f"{name}: not installed")
        except Exception as error:
            errors.append(f"{name}: {error}")
            clear_page_images(pages_dir)
    if result is None:
        sys.exit("Could not read the PDF.\n" + "\n".join(errors) +
                 "\nTry: pip install pymupdf  (or read the PDF shown in the conversation)")

    texts, figure_pages, fields, total_pages = result
    lines = [f"# Task text ({args.pdf.name})", ""]
    for number, text in enumerate(texts, start=1):
        lines += [f"## Page {number}", "", text.strip(), ""]
    text_file = args.out / "task_text.md"
    text_file.write_text("\n".join(lines), encoding="utf-8")

    scanned = [n for n, t in enumerate(texts, start=1) if len(t.strip()) < SCANNED_THRESHOLD]
    images = sorted(pages_dir.glob("page*.png"))

    print(f"Backend: {used}")
    limited = f" (only first {args.max_pages} processed)" if total_pages > args.max_pages else ""
    print(f"Pages: {total_pages}{limited}")
    print("Characters per page: " + ", ".join(f"p{n}={len(t.strip())}" for n, t in enumerate(texts, start=1)))
    print(f"Text written to: {text_file}")
    print(f"Page images: {len(images)} in {pages_dir}")
    if scanned:
        print(f"LOOKS SCANNED / IMAGE ONLY (read these page images and transcribe): {scanned}")
    if figure_pages:
        print(f"Pages with figures, drawings or tables (look at the image): {figure_pages}")
    if fields:
        print("Fillable form fields: " + "; ".join(fields[:30]))


if __name__ == "__main__":
    main()

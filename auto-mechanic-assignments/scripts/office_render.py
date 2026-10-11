#!/usr/bin/env python3
"""Render a .docx with LibreOffice so you can see how it really looks.

  python3 office_render.py Assignment.docx --out work/preview

writes Assignment.pdf plus page-1.png, page-2.png ... into work/preview. Open the images and look
at every page before handing the file over: cover on one page, no table split badly, no heading
alone at the bottom of a page, figures where they belong.

build_docx.py uses the same module to put real page numbers into the table of contents.
Everything here fails soft: without LibreOffice the functions return None or an empty list.
"""

from __future__ import annotations

import argparse
import glob
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

MAC_SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
TIMEOUT_SECONDS = 120


def load_pymupdf():
    try:
        import pymupdf
        return pymupdf
    except ImportError:
        import fitz  # older PyMuPDF releases
        return fitz


def find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    return MAC_SOFFICE if Path(MAC_SOFFICE).exists() else None


def sandbox_runner():
    """Claude.ai's sandbox blocks the socket LibreOffice needs; the public docx skill ships a workaround."""
    if platform.system() != "Linux":
        return None
    for folder in glob.glob("/mnt/skills/*/docx/scripts"):
        sys.path.insert(0, folder)
        try:
            from office.soffice import run_soffice
            return run_soffice
        except Exception:
            sys.path.remove(folder)
    return None


def available() -> bool:
    return sandbox_runner() is not None or find_soffice() is not None


def to_pdf(docx: Path, out_dir: Path) -> Path | None:
    """Convert with LibreOffice; None when it is missing or the conversion failed."""
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf = out_dir / f"{docx.stem}.pdf"
    if pdf.exists():
        pdf.unlink()  # never report an old render as the new one
    args = ["--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx)]
    try:
        runner = sandbox_runner()
        if runner is not None:
            runner(args, capture_output=True, timeout=TIMEOUT_SECONDS)
        else:
            soffice = find_soffice()
            if soffice is None:
                return None
            env = dict(os.environ)
            if platform.system() == "Linux":
                env.setdefault("SAL_USE_VCLPLUGIN", "svp")
            with tempfile.TemporaryDirectory(prefix="lo_profile_") as profile:
                subprocess.run([soffice, f"-env:UserInstallation={Path(profile).as_uri()}", *args],
                               capture_output=True, timeout=TIMEOUT_SECONDS, env=env)
    except (OSError, subprocess.SubprocessError) as error:
        print(f"NOTE: LibreOffice failed ({error})", file=sys.stderr)
        return None
    return pdf if pdf.is_file() and pdf.stat().st_size > 0 else None


def pdf_outline(pdf: Path) -> list[tuple[int, str, int]]:
    """(level, heading text, page number) from the PDF bookmarks LibreOffice writes for headings."""
    try:
        with load_pymupdf().open(str(pdf)) as doc:
            return [(level, title, page) for level, title, page in doc.get_toc()]
    except ImportError:
        pass
    try:
        from pypdf import PdfReader
    except ImportError:
        return []
    reader = PdfReader(str(pdf))
    entries: list[tuple[int, str, int]] = []

    def walk(items, level: int) -> None:
        for item in items:
            if isinstance(item, list):
                walk(item, level + 1)
            else:
                entries.append((level, item.title, reader.get_destination_page_number(item) + 1))
    walk(reader.outline, 1)
    return entries


def pdf_page_count(pdf: Path) -> int:
    try:
        with load_pymupdf().open(str(pdf)) as doc:
            return doc.page_count
    except ImportError:
        from pypdf import PdfReader
        return len(PdfReader(str(pdf)).pages)


def normalize(text: str) -> str:
    text = text.replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    return re.sub(r"\s+", " ", text).strip().lower()


def heading_pages(docx: Path, headings: list[str]) -> dict[int, int] | None:
    """Page number of every heading (by position in the list), or None if LibreOffice is not available."""
    with tempfile.TemporaryDirectory() as temp:
        pdf = to_pdf(docx, Path(temp))
        if pdf is None:
            return None
        outline = [(normalize(title), page) for _, title, page in pdf_outline(pdf)]
    pages: dict[int, int] = {}
    cursor = 0
    for index, heading in enumerate(headings):
        target = normalize(heading)
        for position in range(cursor, len(outline)):
            if outline[position][0] == target:
                pages[index] = outline[position][1]
                cursor = position + 1
                break
    return pages or None


def page_images(pdf: Path, out_dir: Path, dpi: int = 70) -> list[Path]:
    for old in out_dir.glob("page-*.png"):
        old.unlink()
    try:
        with load_pymupdf().open(str(pdf)) as doc:
            paths = []
            for number, page in enumerate(doc, start=1):
                path = out_dir / f"page-{number}.png"
                page.get_pixmap(dpi=dpi).save(str(path))
                paths.append(path)
            return paths
    except ImportError:
        pass
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(out_dir / "page")], check=False)
        found = sorted(out_dir.glob("page-*.png"), key=lambda p: int(re.findall(r"\d+", p.stem)[-1]))
        renamed = []
        for number, path in enumerate(found, start=1):
            target = out_dir / f"page-{number}.png"
            path.rename(target)
            renamed.append(target)
        return renamed
    return []


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path, default=Path("preview"), help="folder for the PDF and page images")
    parser.add_argument("--dpi", type=int, default=70, help="image resolution (70 is enough to judge layout)")
    args = parser.parse_args()
    if not args.docx.is_file():
        sys.exit(f"File not found: {args.docx}")
    if not available():
        sys.exit("LibreOffice is not available here, so no preview. Check the structure with check_text.py instead.")
    pdf = to_pdf(args.docx.resolve(), args.out.resolve())
    if pdf is None:
        sys.exit(f"LibreOffice could not convert {args.docx.name}. Rebuild it with build_docx.py and try again.")
    images = page_images(pdf, args.out.resolve(), args.dpi)
    print(f"PDF: {pdf}")
    print(f"Pages: {pdf_page_count(pdf)}")
    for image in images:
        print(f"  {image}")


if __name__ == "__main__":
    main()

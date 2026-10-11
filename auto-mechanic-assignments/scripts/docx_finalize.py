"""Make a saved .docx look like an ordinary Word file instead of a script's output.

python-docx starts every file from the same template, so every file it makes carries the same
edit-session ids (rsids), the same bibliography GUID, a preview thumbnail of an empty page, a Word 2010
styles part, statistics that say 0 words, and its own XML and zip formatting. Each of those is a
fingerprint a checker can match. finalize() removes them and writes the package the way Word does.

Nothing is invented: editing time and revision history start at zero and fill in for real when the
student opens, edits and saves the file in Word.
"""

from __future__ import annotations

import datetime as dt
import math
import os
import re
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree

WORD_DECLARATION = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
WORD_ZIP_DATE = (1980, 1, 1, 0, 0, 0)  # Word stores every part with this date
TEMPLATE_PARTS = ("docProps/thumbnail", "customXml/", "word/stylesWithEffects")
TEMPLATE_TARGETS = ("docProps/thumbnail", "customXml/", "stylesWithEffects")
IMAGE_TYPES = {"png", "jpeg", "jpg", "gif", "bmp", "tif", "tiff", "emf", "wmf"}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
STYLE_REFERENCE = re.compile(rb'<w:(?:pStyle|rStyle|tblStyle|numStyle|styleLink|numStyleLink) w:val="([^"]+)"')
STYLE_LINKS = ("basedOn", "link", "next")
CHARS_PER_LINE = 90
LINES_PER_PAGE = 46

CORE = ('<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:creator>{author}</dc:creator><cp:lastModifiedBy>{author}</cp:lastModifiedBy>'
        '<cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">{time}</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">{time}</dcterms:modified></cp:coreProperties>')
APP = ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
       'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
       '<Template>Normal.dotm</Template><TotalTime>0</TotalTime><Pages>{pages}</Pages><Words>{words}</Words>'
       '<Characters>{characters}</Characters><Application>Microsoft Office Word</Application>'
       '<DocSecurity>0</DocSecurity><Lines>{lines}</Lines><Paragraphs>{paragraphs}</Paragraphs>'
       '<ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate>'
       '<CharactersWithSpaces>{characters_with_spaces}</CharactersWithSpaces><SharedDoc>false</SharedDoc>'
       '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>')


@dataclass(frozen=True)
class Stats:
    words: int
    characters: int
    characters_with_spaces: int
    paragraphs: int
    lines: int
    pages: int


def document_stats(paragraph_texts: list[str], pages: int | None, extra_pages: int) -> Stats:
    """Counts the way Word reports them in File > Info (pages estimated when no render is available)."""
    texts = [text for text in paragraph_texts if text.strip()]
    lines = sum(max(1, math.ceil(len(text) / CHARS_PER_LINE)) for text in texts)
    return Stats(
        words=sum(len(text.split()) for text in texts),
        characters=sum(len(re.sub(r"\s", "", text)) for text in texts),
        characters_with_spaces=sum(len(text) for text in texts),
        paragraphs=len(texts),
        lines=lines,
        pages=pages or max(1, math.ceil(lines / LINES_PER_PAGE)) + extra_pages,
    )


def word_time(moment: dt.datetime) -> str:
    """Word writes UTC and whole minutes."""
    return moment.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:00Z")


def used_styles(parts: dict[str, bytes]) -> set[str]:
    """Style ids the document, headers, footers and lists actually point at."""
    return {ref.decode() for name, data in parts.items() if name != "word/styles.xml"
            for ref in STYLE_REFERENCE.findall(data)}


def prune_styles(data: bytes, used: set[str]) -> bytes:
    """Word saves only the styles a document uses; python-docx's template carries 100 unused table styles."""
    root = etree.fromstring(data, etree.XMLParser(resolve_entities=False))
    styles = {style.get(W + "styleId"): style for style in root.iter(W + "style")}
    keep = {sid for sid, style in styles.items() if style.get(W + "default") == "1"} | (used & styles.keys())
    pending = list(keep)
    while pending:
        current = pending.pop()
        for tag in STYLE_LINKS:
            link = styles[current].find(W + tag)
            target = None if link is None else link.get(W + "val")
            if target in styles and target not in keep:
                keep.add(target)
                pending.append(target)
    for sid, style in styles.items():
        if sid not in keep:
            root.remove(style)
    return etree.tostring(root, encoding="UTF-8", standalone=True)


def clean_settings(data: bytes) -> bytes:
    """Drop the template's fixed document id and Mac Word 2011 leftovers; Word writes its own on save."""
    data = re.sub(rb"<w:rsids>.*?</w:rsids>", b"", data, flags=re.DOTALL)
    data = re.sub(rb"<w14:docId [^>]*/>", b"", data)
    data = re.sub(rb"<w:zoom [^>]*/>", b'<w:zoom w:percent="100"/>', data)
    data = re.sub(rb'(<w:themeFontLang [^>]*?) w:eastAsia="[^"]*"', rb"\1", data)
    for setting in (b"<w:savePreviewPicture/>", b"<w:doNotAutoCompressPictures/>"):
        data = data.replace(setting, b"")
    return data


def strip_template_traces(name: str, data: bytes, kept: set[str]) -> bytes:
    if name.endswith(".rels"):
        targets = "|".join(re.escape(target) for target in TEMPLATE_TARGETS)
        data = re.sub(rb'<Relationship (?![^>]*TargetMode="External")[^>]*Target="[^"]*(?:' + targets.encode()
                      + rb')[^"]*"[^>]*/>', b"", data)
    if name == "[Content_Types].xml":
        data = re.sub(rb'<Override PartName="/(?:customXml/|word/stylesWithEffects|docProps/thumbnail)[^"]*"[^>]*/>',
                      b"", data)
        used = {part.rsplit(".", 1)[-1].lower() for part in kept}
        for extension in IMAGE_TYPES - used:
            data = re.sub(rb'<Default Extension="' + extension.encode() + rb'"[^>]*/>', b"", data)
    if name == "word/settings.xml":
        data = clean_settings(data)
    if name.startswith("word/") and name.endswith(".xml"):
        data = re.sub(rb' w:rsid[A-Za-z]*="[0-9A-Fa-f]{8}"', b"", data)  # on paragraphs, runs, sections
        data = re.sub(rb'<w:rsid w:val="[0-9A-Fa-f]{8}"/>', b"", data)  # inside style definitions
    if name == "word/document.xml":
        counter = iter(range(1, 10_000))
        data = re.sub(rb'<pic:cNvPr id="(\d+)" name="[^"]*"',
                      lambda m: b'<pic:cNvPr id="%s" name="Picture %d"' % (m.group(1), next(counter)), data)
    return data


def with_word_declaration(data: bytes) -> bytes:
    """Word's declaration, and no indentation between tags (the template's theme and font table are indented)."""
    body = re.sub(rb"^\s*<\?xml[^>]*\?>\s*", b"", data)
    return WORD_DECLARATION + re.sub(rb">\s*\n\s*<", b"><", body)


def finalize(path: Path, author: str, stats: Stats, now: dt.datetime | None = None,
             keep_styles: frozenset[str] = frozenset()) -> None:
    """Rewrite the package in place: template traces out, honest properties in, Word's formatting."""
    with zipfile.ZipFile(path) as source:
        names = [name for name in source.namelist() if not name.startswith(TEMPLATE_PARTS)]
        parts = {name: source.read(name) for name in names}
    if "word/styles.xml" in parts:
        parts["word/styles.xml"] = prune_styles(parts["word/styles.xml"], used_styles(parts) | keep_styles)
    moment = now or dt.datetime.now(dt.timezone.utc)
    parts["docProps/core.xml"] = CORE.format(author=escape(author), time=word_time(moment)).encode()
    parts["docProps/app.xml"] = APP.format(**stats.__dict__).encode()
    order = ["[Content_Types].xml"] + [name for name in names if name != "[Content_Types].xml"]
    handle, temp_name = tempfile.mkstemp(prefix=f".{path.stem}.", suffix=".tmp", dir=path.parent)
    os.close(handle)
    temp = Path(temp_name)
    temp.chmod(path.stat().st_mode & 0o777)  # mkstemp makes it private; keep the saved file's permissions
    try:
        with zipfile.ZipFile(temp, "w") as target:
            for name in order:
                data = strip_template_traces(name, parts[name], set(names))
                if name.endswith((".xml", ".rels")):
                    data = with_word_declaration(data)
                info = zipfile.ZipInfo(name, date_time=WORD_ZIP_DATE)
                info.compress_type = zipfile.ZIP_DEFLATED
                info.create_system = 0
                target.writestr(info, data)
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)

#!/usr/bin/env python3
"""Check a finished assignment for anything that gives away AI writing.

Hard failures (exit code 1):
  - any em dash, en dash, horizontal bar or "--" anywhere in the file, including the
    file properties and image names (every XML part of a .docx is scanned)
  - file metadata that still says "python-docx"
  - word count outside --min-words / --max-words (when given)

Warnings (fix them unless there is a good reason):
  - AI-typical phrases (delve, plays a crucial role, it's important to note, ...)
  - words that are used far too often (crucial, ensure, furthermore, ...)
  - American and British spelling mixed in one document (tire / tyre, color / colour)
  - every sentence roughly the same length, every paragraph the same size
  - many paragraphs opening with the same word or a stiff transition word
  - lots of lists of exactly three things
  - spaced hyphens used as a dash, exclamation marks, emoji, many semicolons
  - leftover [placeholders] (listed so they can be mentioned to the student)

The word count works like Microsoft Word's (every space-separated token with a letter or
digit), and leaves out the cover page, name/class lines, table of contents and references.

Usage:
  python3 check_text.py Assignment.docx [--min-words 800 --max-words 1000]
  python3 check_text.py answers.md
"""

from __future__ import annotations

import argparse
import re
import statistics
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

DASHES = {chr(0x2014): "em dash", chr(0x2013): "en dash", chr(0x2015): "horizontal bar", chr(0x2012): "figure dash"}
DASH_CLASS = "[" + "".join(DASHES) + "]"
DOUBLE_HYPHEN = re.compile(r"(?<=\w)--(?=\w)|(?:^|\s)--(?:\s|$)")

BANNED_PHRASES = [
    "delve", "delves", "delving", "tapestry", "testament to", "in today's", "fast-paced", "ever-evolving",
    "ever-changing", "it's important to note", "it is important to note", "it's worth noting",
    "it is worth noting", "it should be noted", "plays a crucial role", "plays a vital role",
    "plays a pivotal role", "plays a key role", "plays an important role", "pivotal", "seamless",
    "seamlessly", "realm", "landscape", "navigate the", "navigating the", "unlock the", "unleash",
    "game-changer", "game changer", "harness the power", "a myriad of", "myriad", "multifaceted", "embark",
    "elevate", "showcase", "showcases", "underscore", "underscores", "in conclusion", "to sum up",
    "in summary", "not only", "whether you're", "let's dive", "dive into", "deep dive", "let's take a closer look",
    "first and foremost", "rest assured", "meticulous", "meticulously", "intricate", "intricacies",
    "bustling", "vibrant", "nestled", "paramount", "in essence", "i hope this helps", "as an ai",
    "symphony", "beacon", "cutting-edge", "state-of-the-art", "revolutionize", "revolutionise", "boasts",
    "in the world of", "stands as", "serves as a testament", "key takeaway", "a crucial aspect",
    "an essential aspect", "a wide range of", "a wide array of", "when it comes to", "the realm of",
    "transformative", "holistic", "synergy", "in the ever", "it's not just", "it is not just",
]

OVERUSED_WORDS = [
    "crucial", "vital", "essential", "ensure", "ensures", "ensuring", "enhance", "enhances", "enhanced",
    "optimal", "optimize", "optimise", "robust", "leverage", "utilize", "utilise", "utilizing", "facilitate",
    "comprehensive", "furthermore", "moreover", "additionally", "overall", "ultimately", "significantly",
    "numerous", "notably", "various", "efficient", "efficiently", "effectively", "however", "thus",
    "therefore", "hence",
]

STIFF_OPENERS = {
    "furthermore", "moreover", "additionally", "however", "overall", "ultimately", "thus", "therefore",
    "hence", "consequently", "notably", "importantly", "subsequently",
}

US_UK_PAIRS = [
    ("tire", "tyre"), ("tires", "tyres"), ("color", "colour"), ("aluminum", "aluminium"),
    ("liter", "litre"), ("liters", "litres"), ("center", "centre"), ("analyze", "analyse"),
    ("behavior", "behaviour"), ("organize", "organise"), ("realize", "realise"), ("recognize", "recognise"),
    ("gray", "grey"), ("labor", "labour"), ("favorite", "favourite"),
    ("defense", "defence"), ("fueled", "fuelled"),
    ("traveled", "travelled"), ("labeled", "labelled"), ("modeling", "modelling"), ("vapor", "vapour"),
]

WORDS_PER_ALLOWED_USE = 400
WORDS_PER_TRIPLE = 250
MIN_SENTENCE_SPREAD = 4.5
EMOJI = re.compile("[\U0001F300-\U0001FAFF\U00002600-\U000027BF\U0001F000-\U0001F2FF]")
PLACEHOLDER = re.compile(r"\[(?![A-Z0-9 &.,'/]{2,12}\])(?!\d)[^\]\n]{2,60}\]")  # skips [AFDC], [1]
SENTENCE_SPLIT = re.compile(r"(?<=[.!?])\s+(?=[A-Z0-9\"'(])")
WORD = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ0-9][A-Za-zÀ-ÖØ-öø-ÿ0-9'’\-/]*(?:[.,]\d+)*")
TRIPLE = re.compile(r"\b[\w'-]+(?: [\w'-]+)?, [\w'-]+(?: [\w'-]+)?,? (?:and|or) [\w'-]+\b")
REFERENCES = re.compile(r"^(references|works cited|bibliography|sources|reference list)\s*:?$", re.IGNORECASE)
NOT_COUNTED = {"cover", "toc", "refs", "picture"}  # left out of the word count, like teachers do
PICTURE_STYLE = "Picture Placeholder"
PAGEBREAKS = {"\\pagebreak", "<!-- pagebreak -->", "[pagebreak]"}  # same markers as build_docx.py
PICTURE_PATH = re.compile(r'^!\[(.*?)\]\(\s*(?:placeholder|photo|picture|image)?\s*(?:"([^"]*)")?\s*\)', re.IGNORECASE)
COVER_STYLES = {"Student Details", "Assignment Title", "Kicker", "Title", "Subtitle"}
MD_LIST = re.compile(r"^[-*+•]\s+")
MD_NUMBERED = re.compile(r"^(\d{1,3}[.)]|[a-zA-Z][.)])\s")


@dataclass(frozen=True)
class Block:
    text: str
    kind: str  # prose, heading, list, table, caption, cover, toc, refs, picture (placeholder box)


@dataclass
class Result:
    failures: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    info: list[str] = field(default_factory=list)


def count_words(text: str) -> int:
    """Count like Word: whitespace-separated tokens that contain a letter or digit."""
    return sum(1 for token in text.split() if re.search(r"[^\W_]", token))


def snippet(text: str, start: int, end: int, width: int = 40) -> str:
    left = max(0, start - width)
    return "..." + text[left:end + width].replace("\n", " ") + "..."


# ------------------------------------------------------------------ loading

class Sections:
    """Tracks whether we are inside the references section (and its sub-headings)."""

    def __init__(self) -> None:
        self.references_level = 0

    def heading(self, text: str, level: int) -> None:
        if REFERENCES.match(text.strip()):
            self.references_level = level
        elif self.references_level and level <= self.references_level:
            self.references_level = 0

    @property
    def in_references(self) -> bool:
        return self.references_level > 0


def scan_docx_parts(path: Path, result: Result) -> None:
    """Find dashes hiding in file properties, image names, comments and other XML parts."""
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.endswith((".xml", ".rels")):
                continue
            xml = archive.read(name).decode("utf-8", errors="ignore")
            text_checked_elsewhere = name.startswith(("word/document", "word/header", "word/footer"))
            attributes = " ".join(re.findall(r'="([^"]*)"', xml))
            searchable = attributes if text_checked_elsewhere else re.sub(r"<[^>]+>", " ", xml)
            if not text_checked_elsewhere:
                searchable += " " + " ".join(re.findall(r'="([^"]*)"', xml))
            for match in re.finditer(DASH_CLASS, searchable):
                result.failures.append(f"dash hidden in {name}: {snippet(searchable, match.start(), match.end(), 30)}")


TEMPLATE_IDS = ("00B47730", "EF278816-EC6F-A645-907D-7F25AECB1D4A")  # same in every python-docx file
FIXED_DOC_ID = '<w14:docId w14:val="24062061"/>'  # python-docx template's document id
GENERATORS = re.compile(r"\b(?:python-docx|libreoffice|openoffice|claude|anthropic|openai|chatgpt|gpt-\d)\b",
                        re.IGNORECASE)
AUTHOR_FIELDS = re.compile(r"<(dc:creator|cp:lastModifiedBy)>.*?</\1>", re.DOTALL)  # the student's own name


def scan_file_traces(path: Path, result: Result) -> None:
    """Things that show the file was made by a script rather than saved by Word."""
    problems: list[str] = []
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        parts = {name: archive.read(name).decode("utf-8", errors="ignore") for name in names
                 if name.endswith((".xml", ".rels"))}
    joined = "".join(parts.values())
    if any(marker in joined for marker in TEMPLATE_IDS):  # Word writes rsids too, but never these ones
        problems.append("template edit ids (rsids) or bibliography id from python-docx")
    if any(name.startswith("docProps/thumbnail") or "stylesWithEffects" in name for name in names):
        problems.append("leftover template parts (thumbnail or stylesWithEffects)")
    if FIXED_DOC_ID in parts.get("word/settings.xml", ""):
        problems.append("the template's fixed document id in the settings")
    if 'w:styleId="LightShading' in parts.get("word/styles.xml", ""):
        problems.append("unused template styles (the 100 table styles python-docx ships)")
    if any(re.search(r">\s*\n\s*<", re.sub(r"^<\?xml[^>]*\?>\s*", "", text)) for text in parts.values()):
        problems.append("indented XML (Word writes every part on one line)")
    if any(text.startswith("<?xml version='1.0'") for text in parts.values()):
        problems.append("XML written by a script (single-quoted declaration)")
    app = parts.get("docProps/app.xml", "")
    has_text = re.search(r"<w:t(?: [^>]*)?>[^<]*\w", parts.get("word/document.xml", ""))
    if "<Application>Microsoft Office Word</Application>" not in app or (has_text and "<Words>0</Words>" in app):
        problems.append("file properties do not match a Word document (application or word count)")
    metadata = " ".join(AUTHOR_FIELDS.sub("", text) for name, text in parts.items() if name.startswith("docProps/"))
    metadata += " ".join(re.findall(r'(?:descr|title|name)="([^"]*)"', parts.get("word/document.xml", "")))
    generator = GENERATORS.search(metadata)
    if generator:
        problems.append(f"generator name in the properties or image details: {generator.group(0)!r}")
    if re.search(r'<pic:cNvPr [^>]*name="[^"]+\.(?:png|jpe?g|gif|bmp|tiff?)"', parts.get("word/document.xml", "")):
        problems.append("image file names in the picture properties")
    for problem in problems:
        result.failures.append(f"file traces: {problem}. Rebuild with build_docx.py.")


def load_docx(path: Path, result: Result) -> list[Block]:
    try:
        from docx import Document
        from docx.oxml.ns import qn
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except ImportError:
        sys.exit("python-docx is missing. Install it with: pip install python-docx")
    doc = Document(str(path))
    author = (doc.core_properties.author or "") + (doc.core_properties.last_modified_by or "")
    if "python-docx" in author.lower():
        result.failures.append("Metadata author is 'python-docx'. Rebuild with build_docx.py or clear the author.")
    scan_docx_parts(path, result)
    scan_file_traces(path, result)

    in_cover = any(p.style is not None and p.style.name == "Title" for p in doc.paragraphs)
    sections = Sections()
    blocks: list[Block] = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:tbl"):
            table = Table(child, doc)
            cells = [c.text.strip() for row in table.rows for c in row.cells if c.text.strip()]
            styles = {p.style.name for row in table.rows for c in row.cells for p in c.paragraphs}
            if not cells:
                continue
            if in_cover or styles <= COVER_STYLES:  # cover details table, worksheet info bar
                kind = "cover"
            elif styles == {PICTURE_STYLE}:
                kind = "picture"
            else:
                kind = "refs" if sections.in_references else "table"
            blocks.append(Block(" ".join(cells), kind))
            continue
        if child.tag != qn("w:p"):
            continue
        paragraph = Paragraph(child, doc)
        text, style = paragraph.text.strip(), paragraph.style.name if paragraph.style is not None else ""
        if in_cover:
            if text:
                blocks.append(Block(text, "cover"))
            if any(br.get(qn("w:type")) == "page" for br in child.iter(qn("w:br"))):
                in_cover = False
            continue
        if not text:
            continue
        heading = re.match(r"^Heading (\d)$", style)
        if style in COVER_STYLES:
            kind = "cover"
        elif style.lower().startswith("toc"):
            kind = "toc"
        elif heading:
            sections.heading(text, int(heading.group(1)))
            kind = "refs" if sections.in_references else "heading"
        elif sections.in_references:
            kind = "refs"
        elif style == "Caption":
            kind = "caption"
        elif style.startswith("List") or paragraph.paragraph_format.first_line_indent is not None:
            kind = "list"
        else:
            kind = "prose"
        blocks.append(Block(text, kind))
    for section in doc.sections:
        for part in (section.header, section.footer):
            blocks += [Block(p.text.strip(), "cover") for p in part.paragraphs if p.text.strip()]
    return blocks


def clean_markdown(text: str) -> str:
    text = re.sub(r"\*+|`", "", text)
    return re.sub(r"(?<![\w_])_(?=[^\W_])|(?<=[^\W_])_(?![\w_])", "", text)  # italic markers, not ___ blanks


def load_markdown(path: Path) -> list[Block]:
    lines = path.read_text(encoding="utf-8-sig").splitlines()
    blocks: list[Block] = []
    if lines and lines[0].strip() == "---":
        for index in range(1, len(lines)):
            if lines[index].strip() == "---":
                blocks += [Block(line, "cover") for line in lines[1:index] if line.strip()]
                lines = lines[index + 1:]
                break
    sections = Sections()
    for line in lines:
        text = line.strip()
        if not text or text in {"---", "***"} or text.lower() in PAGEBREAKS:
            continue
        clean = clean_markdown(text)
        heading = re.match(r"^(#{1,6})\s+(.*)$", text)
        image = re.match(r"^!\[(.*?)\]", text)
        if heading:
            title = clean_markdown(heading.group(2)).strip()
            sections.heading(title, len(heading.group(1)))
            blocks.append(Block(title, "refs" if sections.in_references else "heading"))
        elif text.startswith("|"):
            if not re.match(r"^\|?[\s:\-|]+\|?$", text):
                blocks.append(Block(clean.replace("|", " "), "refs" if sections.in_references else "table"))
        elif image:
            if image.group(1).strip():
                blocks.append(Block(clean_markdown(image.group(1)), "caption"))
            picture = PICTURE_PATH.match(text)
            if picture:
                blocks.append(Block(picture.group(2) or picture.group(1) or "picture", "picture"))
        elif sections.in_references:
            blocks.append(Block(MD_LIST.sub("", clean), "refs"))
        elif MD_LIST.match(text):
            blocks.append(Block(MD_LIST.sub("", clean), "list"))
        elif MD_NUMBERED.match(text):
            blocks.append(Block(clean, "list"))  # Word keeps the typed "1." so it is counted too
        else:
            blocks.append(Block(clean.lstrip("> "), "prose"))
    return blocks


# ------------------------------------------------------------------- checks

def check_dashes(blocks: list[Block], result: Result) -> None:
    for block in blocks:
        for match in re.finditer(DASH_CLASS, block.text):
            label = DASHES[match.group()]
            result.failures.append(f"{label}: {snippet(block.text, match.start(), match.end())}")
        for match in DOUBLE_HYPHEN.finditer(block.text):
            result.failures.append(f"double hyphen used as dash: {snippet(block.text, match.start(), match.end())}")
        if block.kind in {"prose", "list"}:
            for match in re.finditer(r"[A-Za-zÀ-ÿ,)]\s-\s[A-Za-zÀ-ÿ(]", block.text):
                result.warnings.append(
                    f"spaced hyphen reads like a dash, rewrite: {snippet(block.text, match.start(), match.end())}")


def count_phrase(text: str, phrase: str) -> int:
    pattern = r"(?<![\wÀ-ÿ])" + re.escape(phrase) + (r"(?![\wÀ-ÿ])" if phrase[-1].isalnum() else "")
    return len(re.findall(pattern, text, flags=re.IGNORECASE))


def check_phrases(body: str, word_count: int, result: Result) -> None:
    for phrase in BANNED_PHRASES:
        hits = count_phrase(body, phrase)
        if hits:
            result.warnings.append(f"AI-sounding phrase '{phrase}' x{hits}: rewrite in plain words")
    allowed = max(1, round(word_count / WORDS_PER_ALLOWED_USE))
    for word in OVERUSED_WORDS:
        hits = count_phrase(body, word)
        if hits > allowed:
            result.warnings.append(f"overused '{word}' x{hits} (keep it at {allowed} or fewer)")
    triples = TRIPLE.findall(body)
    if len(triples) > max(3, word_count // WORDS_PER_TRIPLE):
        result.warnings.append(f"{len(triples)} lists of exactly three things (e.g. '{triples[0]}'). "
                               "Fine for real parts lists, but break up the rest into two or four.")


def check_spelling_variant(body: str, result: Result) -> None:
    mixed = []
    for us, uk in US_UK_PAIRS:
        us_hits, uk_hits = count_phrase(body, us), count_phrase(body, uk)
        if us_hits and uk_hits:
            mixed.append(f"{us} x{us_hits} / {uk} x{uk_hits}")
    us_total = sum(count_phrase(body, us) for us, _ in US_UK_PAIRS)
    uk_total = sum(count_phrase(body, uk) for _, uk in US_UK_PAIRS)
    if mixed:
        result.warnings.append("mixed US/UK spelling, pick one: " + "; ".join(mixed))
    elif us_total and uk_total:
        result.warnings.append(f"both US ({us_total}) and UK ({uk_total}) spellings appear, pick one variant")


def check_rhythm(prose: list[str], result: Result) -> None:
    sentences = [s for text in prose for s in SENTENCE_SPLIT.split(text) if len(WORD.findall(s)) >= 3]
    lengths = [len(WORD.findall(s)) for s in sentences]
    if len(lengths) >= 12:
        spread = statistics.pstdev(lengths)
        mean = statistics.mean(lengths)
        result.info.append(f"sentences: {len(lengths)}, average {mean:.1f} words, spread {spread:.1f}")
        if spread < MIN_SENTENCE_SPREAD:
            result.warnings.append(
                f"sentence lengths are too even (spread {spread:.1f}). Mix short punchy sentences with longer ones.")
        if mean > 24:
            result.warnings.append(f"sentences are long on average ({mean:.1f} words). Students write shorter.")

    long_paragraphs = [len(WORD.findall(p)) for p in prose if len(WORD.findall(p)) >= 25]
    if len(long_paragraphs) >= 6:
        ratio = statistics.pstdev(long_paragraphs) / statistics.mean(long_paragraphs)
        if ratio < 0.22:
            result.warnings.append("paragraphs are all about the same length. Vary them a bit.")

    openers = [WORD.findall(p)[0].lower() for p in prose if len(WORD.findall(p)) >= 8]
    if len(openers) >= 6:
        word, hits = Counter(openers).most_common(1)[0]
        if hits >= 3 and hits / len(openers) > 0.3:
            result.warnings.append(f"{hits} of {len(openers)} paragraphs start with '{word}'. Vary the openings.")
        stiff = [w for w in openers if w in STIFF_OPENERS]
        if len(stiff) > 2:
            result.warnings.append(f"too many paragraphs open with a stiff transition: {Counter(stiff).most_common()}")


def check_misc(body_blocks: list[Block], all_blocks: list[Block], body: str, word_count: int,
               result: Result) -> None:
    if EMOJI.search(body):
        result.warnings.append("emoji found: remove them")
    exclamations = body.count("!")
    if exclamations:
        result.warnings.append(f"{exclamations} exclamation mark(s): school writing rarely needs them")
    semicolons = body.count(";")
    if semicolons > max(2, word_count // 300):
        result.warnings.append(f"{semicolons} semicolons: students rarely use them, split into two sentences")
    placeholders = sorted({m.group(0) for b in all_blocks if b.kind != "refs"  # APA uses [descriptions]
                           for m in PLACEHOLDER.finditer(b.text)})
    if placeholders:
        result.info.append("placeholders to mention to the student: " + ", ".join(placeholders))
    pictures = sum(1 for b in all_blocks if b.kind == "picture")
    if pictures:
        result.info.append(f"picture placeholders for the student to fill in: {pictures}")
    lists = sum(1 for b in body_blocks if b.kind == "list")
    prose = sum(1 for b in body_blocks if b.kind == "prose")
    if prose + lists:
        result.info.append(f"paragraphs: {prose} prose, {lists} list items")


def section_counts(blocks: list[Block]) -> list[tuple[str, int]]:
    """Words per section, so a draft can be compared with its word budget section by section."""
    sections: list[tuple[str, int]] = []
    for block in blocks:
        if block.kind == "heading":
            sections.append((block.text[:40], 0))
        elif block.kind not in NOT_COUNTED:
            if not sections:
                sections.append(("(before first heading)", 0))
            title, words = sections[-1]
            sections[-1] = (title, words + count_words(block.text))
    return sections


def run_checks(blocks: list[Block], min_words: int | None, max_words: int | None) -> tuple[Result, int]:
    result = Result()
    check_dashes(blocks, result)
    body_blocks = [b for b in blocks if b.kind not in NOT_COUNTED]
    body = "\n".join(b.text for b in body_blocks).replace("’", "'").replace("‘", "'")
    word_count = count_words(body)
    prose = [b.text for b in body_blocks if b.kind == "prose"]
    check_phrases(body, word_count, result)
    check_spelling_variant(body, result)
    check_rhythm(prose, result)
    check_misc(body_blocks, blocks, body, word_count, result)
    sections = section_counts(blocks)
    if len(sections) > 1:
        result.info.append("words per section: " + " | ".join(f"{title} {words}" for title, words in sections))
    if min_words and word_count < min_words:
        result.failures.append(f"too short: {word_count} words, task asks for at least {min_words}")
    if max_words and word_count > max_words:
        result.failures.append(f"too long: {word_count} words, task allows at most {max_words}")
    return result, word_count


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("file", type=Path, help=".docx, .md or .txt")
    parser.add_argument("--min-words", type=int)
    parser.add_argument("--max-words", type=int)
    args = parser.parse_args()

    if not args.file.is_file():
        sys.exit(f"File not found: {args.file}")
    pre = Result()
    blocks = load_docx(args.file, pre) if args.file.suffix.lower() == ".docx" else load_markdown(args.file)
    result, word_count = run_checks(blocks, args.min_words, args.max_words)
    failures = pre.failures + result.failures

    print(f"Checked: {args.file.name}  |  words (Word-style, without cover, contents, references): {word_count}")
    for line in result.info:
        print(f"  info: {line}")
    for line in result.warnings:
        print(f"  WARN: {line}")
    for line in failures:
        print(f"  FAIL: {line}")
    if failures:
        print(f"RESULT: FAIL ({len(failures)} problem(s), {len(result.warnings)} warning(s))")
        sys.exit(1)
    print(f"RESULT: PASS ({len(result.warnings)} warning(s))")


if __name__ == "__main__":
    main()

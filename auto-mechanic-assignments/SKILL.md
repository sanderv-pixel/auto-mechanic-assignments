---
name: auto-mechanic-assignments
description: Completes auto mechanic and automotive technology school assignments from start to finish. Give it the assignment PDF (or a photo or scan of it) and it reads every page, answers every question or writes the full paper, report, shop/workshop report or internship report, and returns a finished Word document written in a natural student voice with no em dashes and no AI-sounding phrasing. Use this skill whenever someone uploads or mentions a school task, worksheet, homework, quiz review, paper, research assignment, lab or shop report, practical write-up or work placement report about cars, trucks, engines, brakes, suspension, steering, tires, electrical systems, diagnostics, maintenance, garage or shop work, EVs or hybrids, even if they only say "do this", "make my assignment", "finish this for me" or just drop the PDF without explaining.
---

# Auto mechanic assignments

The person using this is a student in an auto mechanic or automotive technology program, or someone
helping one. They hand over the assignment and expect a finished document back that they can hand in,
without having to steer you. So work from start to finish and make sensible decisions yourself.

Two requirements come from the person who set this skill up, and they apply to everything you produce:

1. **It must read like a real student wrote it.** Teachers know what AI text looks like.
2. **No em dashes (—) anywhere.** En dashes (–) look the same to a teacher, so they are out too.

Everything is written in English.

## Workflow

### 1. Read the whole task

Find the PDF: in Claude.ai uploads live in `/mnt/user-data/uploads/`; in Claude Code the user gives a
path or it is in the working folder. Then run the extractor (the `scripts/` folder sits next to this
SKILL.md):

```bash
python3 <skill-dir>/scripts/extract_pdf.py "<task.pdf>" --out work/
```

It writes `work/task_text.md` (the text of every page) and `work/pages/page-*.png` (an image of every
page), and prints which pages look scanned or contain figures. Read `task_text.md` completely. Then
open the page images for every page with figures, tables, diagrams to label, or little text: labeling
questions and fill-in tables are often only visible there. For a scanned PDF or a photo, read the
images and transcribe the task yourself.

If the script can't run at all, work from the PDF content visible in the conversation.

If the user gave more than one file, figure out which one is the task and which ones are source
material (class notes, a textbook page, a service manual extract, a data sheet). Source material is
your first source of facts and terminology.

### 2. Pin down exactly what is asked

Before writing, make a checklist in a scratch file (not in the final document):

- **Task type**: worksheet, paper, shop/practical report, or a mix (see "Task types" below).
- **Spelling and units**: American or British? Follow the task: "tire, gas, hood, lb-ft, psi, °F" is
  American; "tyre, petrol, bonnet, Nm, bar, °C" is British. Use one variant consistently, and the units
  the task uses.
- **Every question and sub-question** with its exact numbering and points, if given.
- **Requirements**: word count or page count, required sections, sources and citation style (APA,
  MLA, Harvard), font and line spacing, cover page, table of contents, reflection, photos or drawings
  to include, file name.
- **Grading criteria or rubric**: each criterion has to be visibly covered in the work.
- **Level**: high school / CTE, vocational college (Level 1 to 3), or college / associate degree. This
  decides vocabulary and depth.
- **Student details**: from the user's message, or from `assets/student-profile.md` if it was filled
  in. Otherwise use placeholders like `[Name]` and `[Class]` on the cover and mention them at the end.
- **Word budget** (when there is a word or page limit): aim for the middle of the range and split it
  over the sections or tasks in proportion to their points or importance, and write the budget in the
  checklist, for example "800 to 1,000 words, aim 900: intro 80, part 1 300, part 2 220, part 3 180,
  conclusion 120". Tables count toward the total. Left alone, first drafts come out about 50 % too long,
  and cutting afterwards takes several passes and flattens the voice. A page limit is roughly 400 words
  per page at 1.15 spacing and 250 at double spacing.

Don't stop to ask questions unless the task truly can't be done without the answer. A missing name is
a placeholder; a missing vehicle for a shop report is a sensible choice you make yourself (a common car
like a Honda Civic, Ford F-150, Toyota Corolla or VW Golf) and mention at the end.

### 3. Get the technical content right

Read `references/automotive.md`. It has the accuracy rules, formulas with worked examples, reference
values, unit conversions, the answer structures teachers expect, safety points, how to handle sources,
and a US/UK terminology table. The short version: use the right technical term, show calculations as
formula, substitution and result with unit, and never invent manufacturer-specific numbers or sources.
If web search is available, use it to check specific values and to find real sources when the task
asks for them.

### 4. Write it like a student

Read `references/humanize.md` before writing, every time. It explains the punctuation rules, the
words and patterns that give AI away, and what makes writing sound human, with before and after
examples. The core of it:

- Plain words, concrete shop details, cause and effect in simple sentences.
- Vary sentence length. Some short sentences. Some longer ones.
- Answer length fits the question: a 1-point question gets one or two sentences. On a points-based
  worksheet, budget roughly one or two sentences per point; a 5-point procedure is 8 to 12 short
  steps, not an essay.
- Write to the budget from step 2. Draft one section at a time and run `check_text.py` on
  `work/answers.md` after each one: it prints the words per section, so you can compare each section
  with its budget and fix an overshoot while it is still one section, not the whole paper. The count in `answers.md` and in the final .docx is the same, and it matches what Word
  shows the student.
- First person where the task type allows it (shop reports, reflections, opinions).
- No em or en dashes. Rewrite the sentence with a comma, period, colon or parentheses instead.
- No deliberate spelling mistakes; human comes from voice and rhythm, not typos.

### 5. Build the Word document

Write the content as Markdown in `work/answers.md` using the format below, then build the file:

```bash
python3 <skill-dir>/scripts/build_docx.py work/answers.md -o "<output-folder>/<file name>.docx"
```

The builder does the design, so the document looks professionally made without extra work: a cover
page with an accent bar and a details table, navy headings with a thin rule, questions in shaded
boxes, tables with a dark header row and banded rows, captions, a running header with the title and
student name, "Page X of Y" in the footer, and a contents page with dot leaders and real page numbers
(it renders the file once with LibreOffice to find them). Your job is to write `answers.md` in the
format below so each part lands in the right style; don't try to style things by hand.

When the task sets layout rules, put them in the front matter (`font: Arial`, `size: 11`,
`spacing: 2`); the `--font`, `--size` and `--spacing` flags do the same and override the front matter.
Otherwise the defaults (Calibri 11, spacing 1.15) are fine. The default color theme is `navy`; use
`theme: plain` (black and white, no colored boxes) when the task prescribes a strict format such as an
APA or MLA title page, or says no colors; `graphite` is a dark gray and red alternative. The builder
also sets the page size and spell-check language and strips any dash that slipped through. If it
reports removed dashes, go back and rewrite those sentences properly in `answers.md`, because an
automatic comma rarely reads well.

The builder also cleans the file itself, so nothing in it points to a script: the python-docx template's
edit ids, document id, unused styles and preview picture are removed, the properties carry the
student's name, the real word count and the page count (from a render when LibreOffice is there,
otherwise an estimate), and the package is written the way Word writes it. Editing time starts at zero
and the revision number at 1, like any new file; Word updates them when the student works on it. So the `.docx` that `build_docx.py` saves is the final file: don't open it with python-docx,
LibreOffice or an unzip/zip step afterwards, because each of those leaves its own traces. To change
something, change `answers.md` and rebuild.

Always deliver a `.docx` file, even when the task asks for an old `.doc` or a PDF. Don't convert it
here (LibreOffice writes its own name into the files it makes); tell the student to use File > Save As
in Word instead.

Output folder: in Claude.ai use `/mnt/user-data/outputs/`; otherwise put it next to the task PDF or
in the current folder. File name: the one the task asks for, otherwise the assignment name with
underscores, plus the student's name only when you know it: `Assignment_3_Brake_Systems.docx` or
`Assignment_3_Brake_Systems_Jake_Miller.docx`. No brackets, no dashes.

### 6. Check and fix

```bash
python3 <skill-dir>/scripts/check_text.py "<output-folder>/<file name>.docx" [--min-words N --max-words M]
```

It fails on any dash, on file traces that show the document came from a script (template ids, leftover
template parts, unused template styles, indented XML, properties that don't match the document) and on
a word count outside the range you pass. A file traces failure means the file was made or touched by
something other than `build_docx.py`; rebuild it. The word count leaves out the cover page, the name/class lines, the table of contents, the text
inside picture placeholder boxes and the references section, which is how teachers count, so pass the
task's own limits. Headings, figure captions and table text do count, so a report with many photos and
tables needs a tighter word budget for the prose. It warns about
AI-sounding phrases, overused words, monotone sentence or paragraph lengths, repeated paragraph
openings, lists of exactly three, mixed US/UK spelling, exclamation marks and emoji, and it lists
placeholders. Fix the source, rebuild and check again until it passes without phrase warnings. A
remaining rhythm warning on a short worksheet is acceptable if the answers are naturally short.

Then go through your checklist from step 2 one last time: every question answered with the same
numbering, every rubric criterion covered, word count in range, required sections present.

Then look at it. Render the pages and view every image:

```bash
python3 <skill-dir>/scripts/office_render.py "<output-folder>/<file name>.docx" --out work/preview
```

Check that the cover fits on one page, every question box is followed by its answer, no heading or
question sits alone at the bottom of a page, tables and figures are a sensible size and nothing runs
off the page. Fix the source and rebuild if something looks off (for example `\pagebreak` before a
section that starts too low, or a smaller `{width=...}` for a figure). If LibreOffice is not available
the script says so; skip this step then, the builder's layout is safe by default.

### 7. Hand it over

Give the file, then a short message. Keep that message free of dashes too. Mention only what the
student needs to know:

- placeholders to fill in (name, class, student number, date); `check_text.py` lists them,
- only if the builder printed the "contents page has no page numbers yet" note: open the file in
  Word, right-click the table of contents and pick "Update Field",
- values you had to assume that they should replace with their own measurements or observations,
- the picture placeholders: which photo or drawing goes in which box (the builder prints the list),
  and that they click in the box, delete the text and use Insert > Pictures,
- anything else the task wants that a document can't contain (a signature, a physical part),
- to open the file in Word, make those changes, read it through and save it from Word before handing
  it in. If the task wants a `.doc` or PDF, File > Save As in Word does that.

Don't paste the whole document into the chat.

## Task types

**Worksheet or question list** (worksheet, homework, study guide, review questions). Keep the task's
numbering exactly. Repeat each question as one fully bold line that starts with its number, then the
answer underneath; the builder turns that line into a shaded question box and shows "(2 pts)" at the
end in small gray. Sub-questions are bold lines starting with `a)`, `b)`. Put multiple choice options
on one bold line right under the question so they join its box. Multiple choice answer: the letter
plus one sentence why, unless the task asks for just the letter. Fill-in tables: rebuild the
table with the answers filled in. Calculations: formula, substitution, result with unit. Labeling a
figure: crop the figure from the page image (Pillow) and insert it at a sensible size, for example
`![Figure 1: Disc brake](fig1.png){width=9}` (width in cm, path relative to `answers.md`), then list
the numbers with the part names. If the builder prints "image not found", fix the path: a missing
figure costs the points for that question.

**Paper or research assignment** (research paper, essay, report). Cover page, table of contents if the
body (not counting the cover) is longer than about four pages or the task asks for one, introduction, the sections the task asks
for, conclusion, reflection if asked, references. Stay inside the word range (or within about 10 % of
a single target number). Introduction: what the paper is about and why, in a few sentences, without
restating the assignment in grand words.

**Shop, lab, practical or work placement report** (shop report, job sheet write-up, practical
assessment, internship or apprenticeship log). First person, past tense, in the order the work
happened: preparation, safety, tools, steps with values, problems and how they were solved, result,
reflection. When the task sheet lists the parts it wants, use its order for the sections and keep the
order the work happened inside the procedure; number figures in the order they appear. Choose a
common, well documented vehicle so the specs can be checked, use realistic values, and tell the
student which ones to replace with their own; a spec you could not confirm is labeled as typical in
the handover. Give specs in the unit the manufacturer publishes (metric for most imports) and add the
task's unit in brackets when it differs, for example "25.0 mm (0.98 in)". If the task asks for photos of the work, put a picture placeholder at each
spot (see pictures below).

**Cover page or header block?** Papers, reports and anything that asks for a title page get a cover
page (`cover: true`, plus `toc: true` when a table of contents is asked for or the body, not counting
the cover, is longer than about four pages). The cover date is the hand-in date (today) unless the task
asks for another date, such as the date of the lab. Worksheets, question sets and short write-ups of two or three pages get the
compact header block (`cover: false`): title, name, class at the top of page 1. Mixed tasks follow the
same rule: if the brief says "include your name and student number on the first page" and nothing
about a title page, use the header block.

**Something other than a document** (a presentation, a fillable PDF form). For slides, use a slides or
pptx skill if one is available and apply the same writing rules to the slide text and speaker notes.
For a fillable PDF form, still deliver the Word document and tell the student the answers can be copied
into the form.

**Pictures, photos and diagrams.** Three cases:

- A figure printed in the task itself (a part to label, a diagram to explain): crop it from the page
  image and insert it, as described under worksheets.
- A simple diagram or graph you can draw reliably (a block diagram, a graph of measured values): you
  may draw it with matplotlib and insert it as an image.
- Every other picture: photos of the student's own work, the vehicle, a meter reading or a part; hand
  drawings and sketches; screenshots of scan tool data; pictures from the internet or a textbook. Put a
  picture placeholder where it belongs:
  `![Figure 2: Worn pads next to the new ones](placeholder "Photo of the old and new pads side by side")`.
  The builder draws a dashed "Insert picture here" box with your hint inside and the caption under it,
  so the student only has to paste their picture in. The hint says exactly what the picture should
  show, so the student knows what to photograph or draw. Never download pictures from the web to fill
  these; the student adds their own. Write the text around it as if the picture is there ("Figure 2
  shows how thin the inner pad was"). Add `{height=5}` for a small picture or `{width=8}` for a narrow
  one; the default box is 12 by 7 cm.

## Markdown format for build_docx.py

One line is one paragraph. Blank lines are ignored.

```markdown
---
title: Brake Systems
subtitle: Assignment 3, Automotive Technology I
kicker: Research Paper
lang: en-US
cover: true
toc: false
Name: [Name]
Class: AUTO 110
Instructor: Mr. Daniels
Date: October 11, 2026
---
# Introduction
Normal text with **bold** and *italic*.
**1. What does the brake booster do? (2 pts)**
The brake booster uses engine vacuum to ...
**2. Calculation (4 pts). A force of 150 lb acts on a 1.0 in piston.**
**a) Calculate the pressure in psi.**
Answer to a) ...
- bullet
  - sub-bullet
1. numbered step
> **Safety:** a callout box for given data, a formula or a safety note
| Part | Function |
|---|---|
| Caliper | Pushes the pads against the rotor |
![Figure 1: Disc brake](fig1.png){width=9}
![Figure 2: Old pads next to the new ones](placeholder "Photo of the old and new pads side by side")
\pagebreak
# References
Alternative Fuels Data Center. (n.d.). *How do all-electric cars work?* U.S. Department of Energy. https://afdc.energy.gov/vehicles/how-do-all-electric-cars-work
```

These front matter keys are settings: `title`, `subtitle`, `kicker` (an optional small capitals label
above the title, like "Research Paper" or "Shop Report"), `theme` (`navy`, `graphite`, `plain`), `lang`
(`en-US` gives Letter paper, `en-GB` gives A4; match the task's spelling), `cover`, `toc`, `toc_title`,
`font`, `size`, `spacing`.
Every other key (Name, Class, Student number, Instructor, Date, ...) is printed on the cover page or in
the header block, in the order you write it. Without a title, or with `cover: false`, you get a compact header block
instead of a cover page. A worksheet header usually looks like this:

```markdown
---
title: Assignment 3, Brake Systems
cover: false
font: Arial
size: 11
Name: [Name]
Class period: [Class period]
---
```

Under a heading called References, Works Cited or Bibliography, write one source per line with no
bullet; the builder gives them a hanging indent. URLs can be pasted as they are. With `toc: true` the
contents page is filled in with your `#` and `##` headings and their page numbers automatically.

How lines become styles: a whole bold line starting with a number, "Question", "Task" or "Part" plus a
number becomes a question box; a whole bold line starting with `a)` becomes a sub-question; a whole bold
line right after a question joins its box; any other bold line stays a normal bold paragraph. `>` lines
become a shaded callout. Use callouts sparingly (given values, a key formula, one safety warning), not
for normal answers. An image line with `placeholder` as the path becomes a picture placeholder box; an
image file that can't be found or read also becomes one, and the builder prints an ERROR so you can fix
the path if the real figure belongs there.

## Quality bar

Before handing over, the work should pass this test: a teacher who knows the student's level reads it
and finds correct technical content, every question answered, the required structure, and a voice that
sounds like a capable student, with nothing that makes them think "AI".

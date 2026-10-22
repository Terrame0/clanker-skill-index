---
name: gost-report
description: Use when writing, formatting, or converting a Russian university report — отчёт по лабораторной работе, курсовая, ВКР, ГОСТ 7.32 — into a .docx with Times New Roman 14, 1.5 spacing, margins 25/10/20/20 mm, uppercase numbered sections, titled tables and figures. Covers the pandoc markdown→docx pipeline and PlantUML / IDEF0 diagram rendering, with helper scripts for the formatting that pandoc cannot express.
---

# GOST-formatted reports

Turn a markdown draft into a `.docx` that passes for a university lab report,
then hand the markdown to `pandoc` with a reference document that carries the
formatting.

## Formatting rules to enforce

| Parameter | Value |
|---|---|
| Page | A4, one side |
| Margins | left 25 mm, right 10 mm, top/bottom 20 mm |
| Font | Times New Roman, 14 pt |
| Line spacing | 1.5 |
| Paragraph | justified, first-line indent 1.25 cm |
| Page numbers | bottom centre; title page counted but not printed |
| Section heading | `1 НАЗВАНИЕ` — arabic number without dot, uppercase, left with indent |
| Subsection heading | `1.1 Название` — lowercase except first letter, no trailing dot |
| New section | starts on a new page |
| Figure | caption below, centred, 12 pt, `Рисунок N — Название` |
| Table | `Таблица N — Название` above, 9-12 pt body |
| Code listing | monospace 11 pt, single spacing, flush left, no indent |
| Lists | dash `—` or `а)`, `б)`, `в)` |

A headline must not be the last line on a page; keep it with at least three
lines of the following text.

## Pipeline

```
report.md ──pandoc + reference_gost.docx──▶ out.docx ──gost_postprocess──▶ report.docx
*.puml  ──plantuml──▶ *.png                        (source paths named on the command line)
*.idef0 ──schematic|decompose|focus──▶ *.svg ──resvg──▶ *.png
```

1. Write the report as markdown (`#` = section, `##` = subsection). Prefix
   every numbered heading with its intended number (`# 1. Цель работы`,
   `## 4.1. Организационная структура`). The post-processor strips the literal
   number and installs a native Word numbered-list field, so deleting a section
   renumbers the rest automatically; level-1 headings are uppercased. A heading
   **without** a leading number is left unnumbered — that is how structural
   elements and appendices are written, and the post-processor centres and
   restyles them (see [Structural elements](#structural-elements)). Do **not**
   number the title page, a preamble, a structural element, or an appendix.

2. Render the diagrams, naming every source file explicitly:

   ```nu
   scripts/render_diagrams.nu diagrams/org.puml diagrams/usecase.puml diagrams/model.idef0 'diagrams/model.idef0#decompose=Заселение'
   ```

   The script takes a list of `.puml` and `.idef0` specs and writes each output
   next to its source. It never scans a directory, so a scratch model left in
   `diagrams/` is ignored unless you name it. The script handles two source
   types:

   - **`.puml` → `.png`** via PlantUML. Structural diagrams (component, use case)
     start with `left to right direction`; activity diagrams (`start`/`stop`)
     must not — that keyword makes plantuml reject the file. Renders at
     `PLANTUML_DPI` (default 300) because plantuml's native PNG is ~89 DPI and
     pixelates once scaled to the page width; set `PLANTUML_DPI=600` for a
     larger image.
   - **`.idef0` → `.svg` + `.png`** via IDEF0-SVG-GOST-wrapped, then `resvg`.
     Renders at `IDEF0_DPI` (default 300). The `.svg` is kept next to the model
     so the PNG can be re-rendered at another DPI without re-running the
     renderer. An IDEF0 spec may carry a level directive `#<verb>[=<process>]`:

     | Spec | Verb | Output |
     |---|---|---|
     | `model.idef0` | `schematic` (every leaf) | `model.svg` |
     | `model.idef0#focus` | `focus "<root>"` (A0 context) | `model-a0.svg` |
     | `model.idef0#decompose` | `decompose "<root>"` (top level) | `model-decompose.svg` |
     | `model.idef0#decompose=<name>` | `decompose "<name>"` (level 2+) | `model-decompose-<name>.svg` |
     | `model.idef0#focus=<name>` | `focus "<name>"` (one box) | `model-focus-<name>.svg` |

     `<name>` is a process name exactly as `toc` prints it; quote the whole
     spec because it contains `#`, `=` and often spaces. See
     [IDEF0 models are multi-level](#idef0-models-are-multi-level) for the
     one-file-per-project rule. The model itself is a small DSL, one statement
     per line:

     ```
     Обработка заявок receives Заявки клиентов
     Обработка заявок respects Регламент обработки
     Обработка заявок requires Специалисты
     Обработка заявок produces Обработанные заявки
     ```

     Predicates are `receives` (input, left), `respects` (control, top),
     `requires` (mechanism, bottom), `produces` (output, right), and
     `is composed of` (sub-function). Every statement is `Noun Verb Noun`;
     blank lines and lines starting with `#` are ignored. Nouns are free-form
     (any text without `;`), including lowercase Latin words, though their first
     letter is uppercased in the label — see the gotchas.

   Embed the PNG in the markdown with the caption as the alt text (pandoc turns
   it into a centred caption under the image):

   ```markdown
   ![Рисунок 1 — Организационная структура объекта](diagrams/org.png){ width=15cm }
   ```

   Pick the width from the image's aspect ratio `h/w` (height ÷ width of the
   rendered PNG). The text column is 17.5 cm wide and 25.7 cm tall, so a figure
   can never exceed 17.5 cm, and its rendered height `width × h/w` must leave
   room for the heading and caption above and below — budget about 22 cm.

   | Aspect `h/w` | Width | Example |
   |---|---|---|
   | ≤ 0.4 (wide landscape) | 17 cm | IDEF0 context diagram |
   | 0.4–1.0 | 15 cm | component / architecture |
   | > 1.0 (portrait) | `min(15, 22 / (h/w))` cm | tall use-case diagram |

   A portrait figure taller than ~22 cm at 15 cm (i.e. any `h/w > 1.5`) will not
   fit alongside other content and gets pushed down, leaving a gap. Size it with
   the formula and put a `PageBreak` div in front so it starts at the top of a
   page. Do **not** add a separate `Рисунок …` paragraph after the image — it
   duplicates the caption.

3. Title tables with pandoc's native caption syntax — a `Table:` line directly
   above the pipe table. Pandoc emits a centred `TableCaption` paragraph, and
   the post-processor keeps it with the table:

   ```markdown
   Table: Таблица 1 — Кадровая структура объекта

   | Сотрудник | Деятельность |
   |---|---|
   ```

   Number table titles by hand (`Таблица N — …`), the same way figures are
   numbered. A `::: {custom-style="Caption"}` div renders the same style, but
   the native form marks the paragraph as a table caption and is preferred. Do
   **not** add a separate caption paragraph after the table — the title goes
   above it.

   Number the objects inside an appendix with the appendix letter as prefix —
   `Рисунок А.3`, `Таблица А.1`, `(В.1)`. The heading `# Приложение А. Название`
   is rewritten into a centred `ПРИЛОЖЕНИЕ А` line (own page) plus a bold centred
   title; a status in parentheses (`# Приложение А (обязательное). Название`) is
   emitted as its own centred line. Use only allowed letters: Cyrillic without
   `Ё З Й О Ч Ъ Ы Ь`, or Latin without `I O`.

5. Build the `.docx`:

   ```nu
   scripts/build.nu report.md
   scripts/build.nu report.md out.docx --titlepage titlepages/kamchatgtu-blank.docx
   scripts/build.nu report.md out.docx --toc
   ```

   `--toc` (`-T`) inserts a `СОДЕРЖАНИЕ` page carrying a native Word TOC field.
   The field is stored with `w:dirty="true"`, so the reader's editor fills it in
   on open; the entries then pick up the `TOC1`-`TOC3` styles, which carry the
   GOST dot-leader right tab. Without `--toc` no contents page is produced. When
   a title page is combined with `--toc`, the body's first section still starts a
   new page (the title page's `pageBreakBefore` is deliberately not skipped).

   The script fetches pandoc's default reference, rewrites its styles and page
   setup to the rules above, runs pandoc, then post-processes heading case,
   page breaks, page numbers, the bibliography list, abbreviations, appendix
   headings and formula numbers. Output defaults to `report.docx` (same stem as
   the input).

   The title page is a separate file passed with `--titlepage`. Two kinds are
   accepted:

   - a `.docx` — its paragraphs are spliced in **verbatim** after pandoc, with
     the source document's section (margins), default table style, and paragraph
     spacing merged in. This reproduces the source one-to-one and is the
     reliable option; the university's blank form is kept at
     `~/documents/titlepages/kamchatgtu-blank.docx`. Crop a title out of a real
     report with
     `scripts/extract_titlepage.py src.docx out.docx --until "Задание 1."`.
   - a `.md` — a markdown fragment with `custom-style` divs, concatenated in
     front of the report before pandoc. Handy for a quick title, but exact
     line breaks and spacing are not guaranteed. Ready fragments live next to
     the blank form in `~/documents/titlepages/` (`kamchatgtu-lab.md`,
     `kamchatgtu-it-eu.md`) and can be copied into the project and filled in.

   For the markdown form, keep the fragment outside the read-only skill tree
   (a `titlepages/` directory in the project or in `~`) and write it with the
   dedicated paragraph styles so it centers without raw XML:

   ```markdown
   ::: {custom-style="TitleCenterBold"}
   «КАМЧАТСКИЙ ГОСУДАРСТВЕННЫЙ ТЕХНИЧЕСКИЙ УНИВЕРСИТЕТ»
   :::
   ```

   Available title styles: `TitleCenter`, `TitleCenterBold`, `TitleLeft`,
   `TitleRight`. Use `&#160;` on its own line inside a `TitleCenter` div for a
   blank line; an empty div disappears.

   To force a table (or any block) onto a fresh page when page-flow heuristics
   let it split, drop a `PageBreak` paragraph in front of it:

   ```markdown
   ::: {custom-style="PageBreak"}
   &#160;
   :::
   ```

   This is a hard `pageBreakBefore`, honored by every renderer — unlike the
   automatic row-keeping below.

## IDEF0 models are multi-level

Keep **one `.idef0` file per project**, not one per figure. The file is a flat
list of statements in a single namespace, and every level of the model is
derived from it by a different command. Splitting it into `context.idef0`,
`decomp.idef0`, `a2.idef0`, … duplicates the statements, and the copies then
drift out of sync — the classic agent failure. There is no per-level file to
write; render the level you need from the one model:

| Figure | Spec |
|---|---|
| A0 context (one box) | `model.idef0#focus` |
| top-level breakdown (`A1`–`A7`) | `model.idef0#decompose` |
| a level-2 diagram | `model.idef0#decompose=Заселение` |
| a single process | `model.idef0#focus=Заселение` |
| full flat schematic (every leaf) | `model.idef0` |

`toc` prints the process names to use after `=`. A "context file" is not a
separate artefact: it is just a model that happens to contain a single box, and
`#focus` produces the same A0 picture from the full model. Two projects that
share process names — an AS-IS and a TO-BE model of the same system — each keep
their own file, because names are global within a file and merging them would
fuse the two decompositions.

## Importing an existing .docx

When the source is a `.docx` rather than markdown — a previous report, a
teammate's draft — convert it to GOST markdown first, then run the normal
[Pipeline](#pipeline). There is no script for this; it is one pandoc call plus
a cleanup pass:

```bash
pandoc src.docx -t gfm --wrap=none --extract-media=. -o report.md
```

`--extract-media=.` writes the embedded images to `./media/`; without it the
output references files that do not exist. Then fix, in order:

1. **Title page.** Drop the imported title block from the markdown and keep the
   original title page instead — pass it through `--titlepage`, either an
   already-cropped `.docx` or one cut out of `src.docx` with
   `scripts/extract_titlepage.py`. gfm turns the two-column «Выполнил / Принял»
   layout table into raw `<table>` HTML, which does not render.
2. **Figures.** gfm emits each figure as a `<figure>` block: the `<img>` carries
   the size in its `style` (`width:5.90551in` is 15 cm) and the `<figcaption>`
   repeats the alt text, so a naive import shows the caption twice. Replace the
   whole block with a single image line:

   ```markdown
   ![Рисунок 1 — Организационная структура объекта](media/rId17.png){ width=15cm }
   ```

   Keep the source width or re-derive it from the image's aspect ratio.
3. **Tables.** gfm puts each `Таблица N – …` caption on its own paragraph
   *after* the table. Move it above and rewrite it as a native
   `Table: Таблица N — …` line. A source that used an en dash (`Таблица N – …`)
   keeps it through the round-trip, but the GOST convention is an em dash.
4. **Numbering.** Word heading numbers are a `numPr` field, not text, so gfm
   returns bare headings. Prefix every section heading with its number
   (`# 5. Отчёт об обследовании`) or the whole document stays unnumbered.
5. **Cross-references.** Inserting captions and renumbering shifts table and
   figure numbers. Renumber to a single sequence, then grep the prose for
   `таблиц` and `рисун` to fix the references.
6. **Contents.** A `СОДЕРЖАНИЕ` page comes back as a bare paragraph where the TOC
   field was. Delete it and pass `--toc` to regenerate the field.

Do not try to carry over the source document's margins, fonts or styles — the
GOST reference replaces them wholesale.

## Structural elements

`ВВЕДЕНИЕ`, `ЗАКЛЮЧЕНИЕ`, `СОДЕРЖАНИЕ`, `СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ`,
`ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ`, `РЕФЕРАТ` and `ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ` are
structural elements: they carry no section number, are centred, uppercased, and
start on a new page. Write them as a plain `#` heading **without** a number:

```markdown
# Введение

# Список использованных источников
```

The post-processor recognises exactly that set, switches the heading to
`Heading1Center`, and uppercases it. A structural element written **with** a
number (`# 6. Заключение`) is treated as an ordinary numbered section — left
aligned, with a number — so drop the number when you mean the structural form.

## Список использованных источников

Write the entries as a normal markdown ordered list directly under the
structural heading. The post-processor detects that heading and rewrites the
whole contiguous list: each item becomes a `SourceList` paragraph with the
number as literal text (`1. `), left-aligned with a 1.25 cm first-line indent.
This is what makes the list reorderable and keeps the numbers out of a Word
field.

```markdown
# Список использованных источников

1. Иванов И. И. Методы обработки данных // Вестник науки. — 2020. — Т. 5, № 3. — С. 12–25.
2. Петров П. П. Основы проектирования информационных систем. — М.: Наука, 2019. — 320 с.
3. ГОСТ 7.32-2017. Отчёт о научно-исследовательской работе. — М.: Стандартинформ, 2017. — 32 с.
```

Order the entries by first citation in the text; cite them as `[1]`, `[1]—[4]`,
`в работе [9]`. The number is typed by hand, not cross-referenced — this is the
one place the skill does not automate Word fields. Description templates per
source type (article, book, conference, patent, URL, standard) are in
ГОСТ 7.32-2017 Appendix Е.

## Перечень сокращений и обозначений

Write each entry on its own line as `СОКРАЩЕНИЕ — расшифровка` with an em dash.
Any paragraph starting with an uppercase token followed by `—` is switched to
the `Abbrev` style: no first-line indent, flush left, so the dash column lines
up. The list is only required when there are more than three designations.

```markdown
# Перечень сокращений и обозначений

ПО — программное обеспечение

АСУ — автоматизированная система управления
```

## Формулы

Put each numbered formula in a `Formula` div, with the number at the end of the
line. The post-processor moves the number to a right-aligned tab at the margin,
so it sits on the same line as the equation. Explanations go in a `Where` div,
one symbol per line.

```markdown
::: {custom-style="Formula"}
$P = U \cdot I$ (1)
:::

::: {custom-style="Where"}
где $P$ — мощность, Вт;

$U$ — напряжение, В;
:::
```

The `Formula` style centres the equation via a centre tab and right-aligns the
number; the `Where` style gives `где` a first-line indent with a hanging
explanation. Numbers and cross-references (`в формуле (1)`) are typed by hand.

## What the scripts do

- `scripts/gost_reference.py` — takes pandoc's default `reference.docx` and
  rewrites `word/styles.xml` (Normal, BodyText, FirstParagraph, Compact,
  Heading1-3, Caption, TableCaption, TableText, SourceCode/Verbatim,
  TitleCenter/TitleLeft/TitleRight, PageBreak, Title/Author/Date) plus the
  feature styles (Heading1Center, AppendixTitle, SourceList, Abbrev, Formula,
  Where, TOCHeading, TOC1-3) and `word/sectPr` page setup. Tables get
  `tblBorders`; the document and Normal style are pinned to `ru-RU` so
  spell-check does not flag every word. Called by `build.nu`; call directly only
  to produce a reusable `reference_gost.docx`.
- `scripts/gost_postprocess.py` — after pandoc, in order: rewrites the
  bibliography ordered list under `СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ` into
  `SourceList` paragraphs with literal numbers; switches `СОКРАЩЕНИЕ — …`
  paragraphs to `Abbrev`; centres the recognised structural-element headings
  (`Heading1Center`, uppercased); splits `# Приложение А. Название` into a
  centred `ПРИЛОЖЕНИЕ А` heading plus an `AppendixTitle`; right-aligns hand-typed
  formula numbers in `Formula` paragraphs via a tab. Then: strips the literal
  number from `Heading1`-`Heading3`, adds a `numPr` pointing at a heading
  numbering definition it appends to `word/numbering.xml`, uppercases `Heading1`,
  puts `<w:pageBreakBefore/>` on every section heading (skipped for the first one
  when a title page is present and no `--toc`), switches table cell paragraphs
  from `Compact` to `TableText`, forces code paragraphs (`SourceCode`/`Verbatim`)
  to flush left with single spacing, keeps table rows from splitting and a table
  up to `KEEP_TABLE_MAX_ROWS` rows together with its caption, adds a centred
  page-number footer to the body section, and — with `--titlepage <file.docx>` —
  splices in the title page's body, section properties and styles.
- `scripts/extract_titlepage.py` — crops the leading body of a report `.docx`
  into a standalone title-page `.docx` (`--until <text>` or `--count <n>`).
- `scripts/build.nu` — end-to-end `.md → .docx`. Runs pandoc with
  `--no-highlight`, so code listings stay black monospace (pandoc's default
  syntax highlighting colours them, which is wrong for a printed report).
- `scripts/render_diagrams.nu` — takes a list of `.puml`/`.idef0` specs and
  renders each to a sibling output: `.puml → .png` at `PLANTUML_DPI` (default
  300), `.idef0 → .svg + .png` at `IDEF0_DPI` (default 300) via
  `schematic`/`decompose`/`focus` then `resvg`. An `.idef0` spec may carry a
  level directive `#<verb>[=<process>]` (`#focus` for A0, `#decompose=<name>`
  for a level-2 diagram); the output name grows the directive suffix so levels
  do not collide. It does not glob a directory; every source must be named.

All scripts pull `pandoc`, `python3`, `plantuml`, `graphviz`, `resvg`, and the
`schematic`/`decompose`/`focus` binaries through `nix shell` when the tools are
missing, so they work on a bare NixOS host. The IDEF0 binaries come from the
[IDEF0-SVG-GOST-wrapped](https://github.com/Terrame0/IDEF0-SVG-GOST-wrapped)
flake, not nixpkgs.

## Gotchas

- **`pandoc -t gfm` moves a table caption below its table.** A caption that was
  correctly above the table in the source `.docx` comes back as a plain
  paragraph *after* it, so the round-trip does not preserve position. Move it up
  and restore the `Table:` prefix, or the rebuild emits an unnumbered caption
  under the table.
- **There is no safe default width for a portrait figure.** The 15 cm default is
  calibrated for landscape images; a tall use-case diagram (`h/w` ≈ 1.9) renders
  about 28 cm high and overflows a 25.7 cm text column. Size it with
  `width ≤ 22 / (h/w)` cm and give it a `PageBreak`, or the renderer clips or
  floats it.
- **Table cells use `Compact` by default**, which inherits the 1.25 cm
  first-line indent and 1.5 spacing; in narrow columns this breaks words
  mid-syllable (`Сопроти\вл ение`). The post-processor rewrites cell
  paragraphs to `TableText` (12 pt, single spacing, no indent). If you change
  that mapping, keep the `ind firstLine=0`.
- **The Table style has no borders** out of the box. `gost_reference.py`
  injects `tblBorders`; without them the table prints as guide lines only.
- **Pandoc wraps a figure in a single-cell `FigureTable`.** When the image
  cannot be fetched it puts the alt text in that cell, and a naive
  "every `Compact` inside `w:tbl` becomes `TableText`" rule then mis-styles the
  caption cell. Both table passes skip tables whose `tblStyle` is
  `FigureTable`; keep that guard when adding table rewriting.
- **The contents page is an unpopulated field until the editor refreshes it.**
  Pandoc stores the TOC with `w:dirty="true"`, so Word and LibreOffice fill it
  in on open (Word prompts), but OnlyOffice can show it empty until you press
  `F9`. The `TOC1`-`TOC3` styles are in place, so the entries get the GOST dot
  leaders once populated. Do not try to pre-render page numbers — page layout is
  the renderer's job.
- **A `.docx` title page plus `--toc` used to leave page 2 blank.** The title
  page's `nextPage` section break already opens a new page, and the `TOCHeading`
  style carried its own `pageBreakBefore`, so `СОДЕРЖАНИЕ` skipped to page 3 and
  page 2 stayed empty. The `TOCHeading` style no longer sets a page break: the
  TOC is always the first block after the title, so the section break alone is
  enough. If you reintroduce a break there, expect the empty page back.
- **ГОСТ 7.32 contradicts itself on the bibliography dot.** Clause 6.16 says to
  number entries "арабскими цифрами с точкой" (`1.`), while the worked example
  in Appendix Д shows no dot (`1`). The post-processor follows the clause text
  and emits `1.`; change `convert_source_list` if your department wants the
  example's form.
- **A structural element written with a number is not centred.** Centring keys
  off an unnumbered `Heading1` whose text matches the fixed set, so
  `# 6. Заключение` stays an ordinary left-aligned numbered section. Write
  `# Заключение` to get the centred structural form.
- **The abbreviation pass is a text heuristic.** Any paragraph whose text starts
  with an uppercase token followed by an em dash (`ПО — …`) is switched to
  `Abbrev`, wherever it appears. A body sentence beginning that way is caught
  too; keep such sentences from opening with a bare uppercase token and a dash.
- **The bibliography pass keys off the heading text, not a marker.** It matches
  `СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ` (case-insensitive, so a numbered
  `# 2 Список использованных источников` works) and rewrites the contiguous
  list that follows. A list elsewhere in the document is left alone; a
  bibliography not under that heading is not converted.
- **Images with a width in braces** need the `{ width=15cm }` attribute form;
  a bare `![alt](path)` scales to natural size and can overflow the page.
- **PlantUML's default PNG is ~89 DPI and pixelates at page width.**
  `render_diagrams.nu` passes `-Sdpi=300`; the `scale` directive does not help
  because it only applies to SVG output. Raise `PLANTUML_DPI` if a figure is
  still soft.
- **IDEF0 box numbers follow model order, not the picture.** Children are
  numbered `A1`, `A2`, … by their order in the `.idef0` file, and a child of `A2`
  becomes `A21`, `A22`, …, but the layout engine reorders boxes to reduce line
  crossings, so `A2` can appear left of `A1`. Every level has its own picture —
  `decompose "<name>"` — while the A0 context diagram (`focus "<root>"`) is
  unaffected because it has exactly one box.
- **A mechanism required by several processes is drawn once.** The renderer
  merges bottom-entering mechanisms that share a name into a single trunk with
  branch arrows, so seven `X requires PMS` lines produce one `PMS` label, not
  seven. This is the GOST look; give the mechanisms distinct names if you want
  them drawn separately. Only external mechanisms (bottom) are merged, not the
  `produces`/`requires` internal arrows.
- **The DSL parser uppercases the first letter of every Latin word.** Lowercase
  Latin is accepted, but the label is normalised: `requires Навык gost-report`
  renders as `Навык Gost-report`. There is no way to keep an all-lowercase Latin
  word in a figure.
- **A noun must not contain a connective word.** The verb is identified by
  matching one of the five connectives, so a noun that itself contains
  `receives`, `respects`, `requires`, `produces`, or `is composed of` mis-splits:
  the parser takes the last such word as the verb. Keep such words out of nouns,
  or write the noun in Cyrillic.
- **A named `.idef0` that fails to parse aborts the whole render and leaves a
  0-byte `.svg`.** The IDEF0 verb exits with a Ruby stack trace, nushell's `o>`
  redirect has already created an empty `.svg`, and the pipeline stops there —
  `resvg` never runs, so no `.png` is written and the script exits non-zero.
  IDEF0 diagrams listed after the bad one are not rendered; PlantUML ones are
  unaffected because they render first. Because the script takes an explicit
  file list rather than globbing, only a file you name can trigger this.
- **`Times New Roman` is not installed on Linux.** The figure's font stack falls
  back to `Liberation Serif`, which is metric-compatible and covers Cyrillic.
  `resvg` resolves this through system fonts, so the diagram inherits the
  desktop's font configuration.
- **Do not reach for PlantUML, Graphviz, or d2 for IDEF0.** None can put ports
  on all four sides of the box. IDEF0-SVG is the only CLI renderer that does;
  its `schematic`, `decompose`, and `focus` entry points are the ones to use.
- **A `SourceCode` style in `reference.docx` is not enough to keep code flush
  left.** Pandoc's default `SourceCode` does not set `jc`, so once `Normal` is
  justified the code inherits `both` and stretches each line. The post-processor
  therefore injects the alignment directly into every code paragraph; do not
  remove that step in favour of the style alone.
- **Pandoc regenerates `word/numbering.xml` and discards the reference's
  copy.** Do not put the heading numbering definition in `gost_reference.py`;
  append it to the pandoc output in `gost_postprocess.py`, picking ids above
  the ones pandoc already used.
- **Only headings with a leading number get auto-numbering.** The
  post-processor keys off the literal prefix (`1`, `4.1`), so an appendix or a
  preamble heading written without one stays unnumbered — which is what you
  want, but also means a section you forget to prefix silently loses its
  number.
- **A page break must be `<w:pageBreakBefore/>` on the heading, not a
  `<w:br w:type="page"/>` run.** With auto-numbering the number is rendered
  before the runs, so a break run lands between the number and the heading text
  and sends them to different pages (and the number miscounts).
- **A `.docx` title page only renders faithfully if its styles come along.**
  The form's two-column "Выполнил / Принял" block is a layout table that needs
  the source's default table style, and its vertical rhythm comes from the
  source's `docDefaults` spacing. The post-processor merges the missing styles
  and bakes the source spacing into title paragraphs; without that the columns
  collapse and the title spills onto a second page.
- **A run whose text starts or ends with a space needs `xml:space="preserve"`**
  or Word and OnlyOffice collapse the space on open (`от университета` renders
  as `отуниверситета`, and a save then bakes the loss in). Title-page forms
  split text across runs at arbitrary points, so the leading space often lands
  inside the run. Pandoc marks its own runs, but a spliced title page does not
  pass through pandoc; `preserve_run_spaces` in the post-processor tags any
  `<w:t>` whose text has edge whitespace. Match real `<w:t>` elements only —
  a loose `<w:t[^>]*>` also matches `<w:tbl>`, `<w:tc>` and friends.
- **A `.md` title page gets no section break of its own.** Unlike the `.docx`
  form, the fragment is just concatenated, so the title is separated only by
  the `pageBreakBefore` on the report's first `Heading1`. If the body starts
  with plain prose instead, the prose lands on the title page. End the fragment
  with a `PageBreak` div (`&#160;` inside) only when the body does **not**
  start with a heading: with one, the div's break and the heading's own stack
  and leave a blank page between.
- **A table that does not fit the remaining space is moved whole to the next
  page, not split.** The post-processor stamps `<w:cantSplit/>` on every row
  and `<w:keepNext/>` on all but the last row of a table when it has at most
  `KEEP_TABLE_MAX_ROWS` (15) rows, and keeps the caption paragraph with the
  table. A longer table is allowed to split, and pandoc's header row repeats on
  each page. Raise or lower the constant to change the cutoff.
- **`cantSplit` on a row does not keep the whole table together.** It only
  stops a single row from breaking mid-height. Whole-table retention needs
  `keepNext` on the preceding rows as well.
- **OnlyOffice does not honor `keepNext` on table cells.** Rows still split
  across pages there while Word and LibreOffice keep them together — ONLYOFFICE
  issue #3641, a confirmed bug (fix not merged). When a table must not split and
  the output will be opened in OnlyOffice, put a `PageBreak` paragraph in front
  of it instead of trusting the automatic keep. The `keepNext`/`cantSplit`
  stamps are still worth adding for Word and LibreOffice.
- **`<w:pPr>` children must stay in schema order** — `pStyle`, then `keepNext`,
  then `keepLines`, `pageBreakBefore`, `numPr`, `spacing`, `ind`, `jc`. Word
  and OnlyOffice silently drop out-of-order properties, so a `keepNext` inserted
  before `pStyle` looks right in LibreOffice (lenient) but does nothing in Word.
  Insert `keepNext` right after `pStyle`, not at the start of `pPr`.
- **An empty `custom-style` div vanishes before the docx is written.** For a
  blank line in a markdown title page, put `&#160;` inside the div rather than
  leaving the body empty; pandoc drops a div whose only content is whitespace.
- **A document left in `en-US` gets every Russian word underlined by
  spell-check.** Pandoc's default `docDefaults` declares `en-US`; Word and
  OnlyOffice then flag the whole report, which looks like an error in a
  printed screenshot. `gost_reference.py` sets `<w:lang w:val="ru-RU">` in both
  `docDefaults` and `Normal`; do not drop it.
- **Page numbers live in a footer part, not in `document.xml`.** The
  post-processor adds `word/footer1.xml`, its relationship and content type,
  and a `<w:footerReference>` on the last section only. Because the title-page
  section keeps none, the title is counted but not printed. The reference must
  be the **first** child of `<w:sectPr>` — before `pgSz` — or Word ignores it.
- The layout is tuned for this university's guide; different faculties may
  vary margins or fonts. Recheck `MARGINS`, `FONT`, `HALF_PT`, `LINE` at the
  top of `gost_reference.py`.

## Suggested document structure

```
Title page
СОДЕРЖАНИЕ            (build --toc; unnumbered structural element)
<UNNUMBERED preamble: role distribution / состав отчёта>
1 ЦЕЛЬ РАБОТЫ
2 ПОСТАНОВКА ЗАДАЧИ
…
ЗАКЛЮЧЕНИЕ            (unnumbered structural element)
СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ   (unnumbered; ordered list follows)
Приложение А …        (unnumbered; split into ПРИЛОЖЕНИЕ А + title)
```

Structural elements (`СОДЕРЖАНИЕ`, `ЗАКЛЮЧЕНИЕ`, `СПИСОК …`) are written without
a number so the post-processor centres and uppercases them; see
[Structural elements](#structural-elements).

Keep prose in the body and tabular data in pipe tables. Attribute parts to
their authors with a short italic line under the heading when the report is a
group effort.

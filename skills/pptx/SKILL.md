---
name: pptx
description: Use when building, restyling, or reviewing a slide deck (.pptx) — conference talks, defence slides, internal presentations, technical overviews. Covers the python-pptx deck builder, the 16:9 grid and theme model, code slides with syntax highlighting and ligatures, official logos and Nerd Font icons, speaker notes, and the render-audit loop that catches overflow and font substitution before the talk.
---

# Presentation decks

A deck is a background for speech, not a document. Every slide carries one
claim, and the speaker notes carry the argument. This skill is a working
`python-pptx` builder, a rendering loop that shows what the audience will
actually see, and the rules that keep the result consistent.

## What is in the skill

| Path | Purpose |
|---|---|
| `scripts/deck.py` | The builder: grid, themes, slides, code highlighting, tables, logos, notes |
| `scripts/new_deck.nu` | Scaffold a project from `template/` (or `--refresh` to update `deck.py`) |
| `scripts/render.nu` | Deck → PDF → page PNGs → contact sheets, plus embedded font list |
| `scripts/audit.py` | Out-of-bounds shapes, text overflow, fonts requested, slides without notes |
| `template/` | A seven-slide worked example that builds and passes audit |

Start a project:

```bash
nu <skill>/scripts/new_deck.nu ~/talks/my-talk
cd ~/talks/my-talk
./build.sh
nu <skill>/scripts/render.nu deck.pptx
```

## The grid

16:9, 13.333 × 7.5 in. Everything is in inches; no `Inches(...)` in slide code.

| Element | Value |
|---|---|
| Slide margins | 0.72 all round |
| Content top | 1.62 |
| Content bottom | 6.75 |
| Title baseline | title box at 0.5, 26 pt, auto-shrunk to fit |
| Rule under title | y = 1.26, 0.9 pt, `#D5DBE2` |
| Accent bar | 0.12 in wide, full height, left edge, theme accent |
| Key line | 5.7, italic 13 pt, theme accent |
| Page number | bottom right, 9 pt, muted, `3 / 18` |

`deck.M`, `deck.CW`, `deck.top`, `deck.bottom`, `deck.W`, `deck.H` expose the
grid so slide code does not repeat constants.

## Theme model

One theme per topic, four colours, derived from the subject's own brand
palette when it has one:

| Field | Role |
|---|---|
| `accent` | bar, title, header row of tables, key line |
| `tint` | section-slide background, callout fills |
| `edge` | borders of tinted boxes |
| `secondary` | large display numbers on section slides |

Ready themes: `NEUTRAL`, `NIX`, `GLSL`, `SCSS`, `NUSHELL`, `PYTHON`, `CPP`,
`RUST`. Build another with `Theme.hex(accent, tint, edge, secondary)`.

Do not invent a palette. Take the accent from the project's own brand page and
record the code point in the deck repository:

| Project | Accent | Other |
|---|---|---|
| Nix | `#5277C3` | light `#7EBAE4` |
| Sass | `#CC6699` (`$sl-color--hopbush`) | on white, black variants provided |
| OpenGL | `#5586A4` | trademark of HPE, attribution required |
| Nushell | `#1C9438` | logo is white-on-dark, needs a plate |

## Slide anatomy

```
accent bar | title ....................... [mark]
           | rule -----------------------------------
           | content
           | (key line, italic)
           |                                    n / total
```

- **Title** names the claim, not the topic: «Store: каждая сборка в своей
  папке», not «Архитектура Nix». A title that only labels a topic wastes the
  most-read line on the slide.
- **Key line** at the bottom restates the takeaway in one sentence. If the
  slide needs two takeaways, it is two slides.
- **Marks** (a small glyph in the top-right) identify the section without a
  repeated colour legend. Optional; only with a font that has the glyph.
- **Speaker notes** on every slide. The slide shows, the notes argue.

## Slide types

**Title.** `deck.title_slide(...)` with 1–2 lines, optional subtitle, year, and
an optional `logo_row`. No icons, no agenda, no spoilers. The subtitle says
what the talk is about, not what the deck contains.

**Section divider.** `deck.section(slide, name, theme, number="02", subtitle=...)`.
Full-bleed tint, giant number, name, one factual line, logo or glyph. Use one
before each block so the audience knows where it is. If the block is not
written yet, ship the divider alone — an empty divider is honest, an invented
slide is not.

**Content.** `header` + `bullets` + `key`. Bullets are lowercase after the
bullet unless the word is a proper noun (`Nix`, `WHNF`, `Элко`); the helper
`lowercase=True` does this. Max six items, one line each where possible.

**Code.** `deck.code(...)` with a language key. Rules in the section below.

**Table.** `deck.table(data, col_fracs=...)`. Header row is the accent colour,
body white, 0.6 pt borders. The table must answer one question; if it is a
transcript of the talk, cut it.

**Comparison.** Two or three `box` calls side by side, same size, left-aligned
text, different tint. Never centre multi-line prose in a box.

**Closing.** One sentence, large, plus at most four short lines. Not a summary
of the deck — the one thing to remember.

## Code slides

The tokenizer in `deck.py` colours keywords, types, builtins, strings, numbers
and comments per language. Keys: `python`, `nix`, `glsl`, `nu`, `scss`, `cpp`,
`js`, `ts`, `rust`, `go`, `bash`, `sql`.

Rules that matter more than colour:

- **Comment markers are language-specific, not universal.** In Nix `//` is the
  attribute-set merge operator and only `#` starts a comment; in C-family
  languages it is the opposite. A highlighter that treats `//` as a comment
  everywhere paints an operator grey and italic. `Syntax.line_comment` exists
  for this reason.
- **Never disable word wrap on prose.** Code frames set `word_wrap = False`
  deliberately, because wrapping code invents line breaks that change meaning.
  `audit.py` then reports horizontal overflow instead.
- **Size the box to the code.** The builder computes height from the line
  count; a fixed-height box leaves two inches of empty grey under a four-line
  snippet. If you must cap the height, pass `height=`.
- **Do not duplicate the result.** If the code already ends with
  `# => 3`, do not add a box with `3` next to it.
- **Show the smallest program that makes the point.** Slides are not files.

## Ligatures

JetBrains Mono (and its Nerd Font patch) ships 152 base ligatures in the
`calt` feature, so they are one switch in the editor, not two.

| Written | Shown | Use |
|---|---|---|
| `<=>` | ⟺ | equivalence, substitution |
| `=>` | ⇒ | result of an expression |
| `->` | → | mapping, transition |
| `::` | ∷ | type annotation |
| `//` | ∕∕ | operator, not a comment |
| `>=`, `<=`, `!=`, `===` | ≥, ≤, ≠, ≡ | comparisons |

Keep the source ASCII (`<=>`, not `⇔`): the ligature gives the arrow where it
is supported, and the literal operator still reads correctly where it is not.

Two traps: `==` and `=` become nearly indistinguishable, so never rely on
`===` versus `==` at a glance; and `//` looks like a comment marker, so the
colour must carry the distinction.

## Assets

Logos come from the project, not from a search engine. Check the brand page
for colour codes, clear space, allowed variants, and a licence; record all of
it in the deck repository.

- **Square marks** (avatar, seal) go on section slides; **wordmarks** go in the
  header and are sized by width, not forced into a square.
- **Keep the aspect ratio.** `deck.logo(...)` fits inside a box without
  stretching.
- **Dark-background variants.** A white-on-transparent logo disappears on a
  light slide. Put it on its own dark rounded plate, or use the variant the
  project ships.
- **Attribution.** Trademarked marks (OpenGL is HPE's) need a credit line: in
  speaker notes or the last slide, small but present.
- **Nerd Font icons** are glyphs from the Private Use Area, coloured like text.
  Only codepoints in the BMP (`U+E000`–`U+FFFF`) render reliably; a glyph from
  plane 15 (`U+F01A7`) comes out as a wrong character in LibreOffice. Check
  each glyph in the render before using it.

## Fonts

Default is a serif heading face plus a sans body and a mono code face. A single
mono family everywhere (JetBrainsMono NFP) reads well for technical talks and
adds icon glyphs, but it needs the font installed on the presenting machine.

Verify what the PDF actually embeds, not what you asked for:

```bash
nu scripts/render.nu deck.pptx     # prints the embedded font list
```

A name that comes back as a different family means the renderer substituted it
silently. If the deck travels, either embed the fonts or carry them.

## Speaker notes

Notes are not a transcript. One note per slide, 2–5 sentences: what to say
first, what the audience should look at, the caveat you would give if asked.
Write them while the slide is fresh; the deck is not done until every slide has
one.

## Build and review loop

```bash
./build.sh                                  # python-pptx → deck.pptx
nu <skill>/scripts/render.nu deck.pptx      # → deck.render/sheet-1.png, page-*.png
python3 <skill>/scripts/audit.py deck.pptx  # bounds, overflow, notes, fonts
```

Then look at the contact sheet first — it shows rhythm, empty slides, and
layouts that repeat too much. Zoom into individual pages only after the sheet
looks right:

```bash
pdftoppm -f 7 -l 7 -r 110 -png deck.render/deck.pdf /tmp/slide7
```

Repeat until the sheet reads as one deck. The audit catches geometry; only the
image catches taste.

## API reference

```python
from deck import Deck, Theme, NIX, NEUTRAL

deck = Deck(head="Georgia", body="Calibri", code="DejaVu Sans Mono")

deck.title_slide(["Line one", "line two"], subtitle="...", year="2026", entries=[(path, label)])
deck.section(slide, "Nix", NIX, number="03", subtitle="...", logo=path, mark="nixos", notes="...")
deck.header(slide, "Claim as a title", NIX, mark="nixos")
deck.title(slide, "Claim", NIX)                     # title without the bar
deck.bullets(slide, ["one", "two"], top=1.7, size=17, lowercase=True)
deck.text(slide, "Multi\nline prose", 0.72, 3.0, 6.0, size=18, align=PP_ALIGN.CENTER)
deck.key(slide, "One sentence takeaway", NIX)
deck.box(slide, "Title\n\nbody line", left, top, width, height, theme=NIX,
         fill=NIX.tint, valign="top")
deck.chevrons(slide, ["a", "b", "c"], NIX)
deck.code(slide, "Nix", source, "nix", left=0.72, top=deck.top, width=7.0, size=12)
deck.table(slide, data, col_fracs=[0.2, 0.4, 0.4], theme=NIX)
deck.logo(slide, path, left, top, max_w, max_h)
deck.logo_row(slide, [(path, "Nix"), (path, "GLSL")], top=3.1, height=0.8)
deck.glyph(slide, "cube", left, top, width, height, NIX.accent, size=28)
deck.notes(slide, "what to say")
deck.properties(title="...", author="...")
deck.save("deck.pptx", thumbnail=True, page_numbers=True)
```

Geometry arguments are inches. `Deck` converts to EMU internally, so mixing
`deck.CW` (inches) with literals never produces a half-inch-wide text box.

## Gotchas

- **python-pptx does not shrink text.** A 40-character title at 26 pt overflows
  a 12-inch box, and PowerPoint will not fix it. `deck.title` measures with
  `PIL.ImageFont` and steps down to 15 pt; text added through `deck.text` does
  not, so pass a size that fits.
- **Shapes carry a default shadow.** `shadow.inherit = False` is set on every
  shape the builder creates. A raw `add_shape` in slide code reintroduces the
  drop shadow and breaks the flat look.
- **`word_wrap = False` is for code only.** With wrap on, a long token is
  broken mid-word; with wrap off in prose, the text runs off the slide.
  `audit.py` reports both, but it cannot tell which one you meant.
- **Centred multi-line prose reads badly.** Centring is for titles, numbers and
  single-line callouts. In boxes with three or more lines, left-align.
- **A table needs explicit column widths.** python-pptx divides evenly
  otherwise, which gives a two-word column the same width as a sentence column.
  Pass `col_fracs`.
- **The default slide size is 4:3.** Always set `slide_width` and
  `slide_height`; `Deck` does, a bare `Presentation()` does not.
- **`notes_slide` is created on first access.** Counting notes on a deck built
  elsewhere can create empty ones; read before writing.
- **LibreOffice is the preview, not the deliverable.** It renders close to
  PowerPoint but not identically: check the final file in the tool the venue
  uses, at least once, on the machine that will present.
- **Icons and ligatures are font features, not images.** A deck that looks
  perfect locally shows boxes and raw operators on a machine without the font.
  This is the most common way a good deck fails at the podium.
- **Do not invent slide content for a section whose text is not written.**
  Ship the divider and a note. Placeholder filler is what makes a talk feel
  hollow, and it survives into the final version because nobody re-reads it.

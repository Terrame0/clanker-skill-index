"""A small, opinionated deck builder on top of python-pptx.

The grid, type scale and colour model are tuned for 13.333 x 7.5 in (16:9).
Everything else is a thin wrapper so a deck is one function per slide.

    from deck import Deck, Theme, PYTHON, NIX

    deck = Deck(head="Georgia", body="Calibri", code="DejaVu Sans Mono")
    slide = deck.slide()
    deck.title(slide, "Заголовок", PYTHON)
    deck.bullets(slide, ["первый пункт", "второй пункт"])
    deck.key(slide, "одна мысль внизу слайда", PYTHON)
    deck.save("deck.pptx")
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import zipfile
from dataclasses import dataclass
from functools import lru_cache

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Length, Pt

W = Inches(13.333)
H = Inches(7.5)
M = Inches(0.72)
CONTENT_TOP = Inches(1.62)
TEXT_BOTTOM = Inches(6.75)
TEXT_WIDTH = W - 2 * M

BASE = {
    "text": "#262B33",
    "muted": "#6B7280",
    "white": "#FFFFFF",
    "border": "#D5DBE2",
    "code_bg": "#F7F8FA",
    "warn": "#B42318",
    "warn_tint": "#FDECEA",
    "ok": "#2E7D5B",
    "ok_tint": "#EAF6F0",
}


EMU_PER_INCH = 914400


def _inches(value) -> float:
    if isinstance(value, Length):
        return value / EMU_PER_INCH
    return float(value)


def _emu(value) -> int:
    """Accept inches as a plain number, a Length, or raw EMU."""
    if isinstance(value, Length):
        return int(value)
    if isinstance(value, int) and abs(value) > 10000:
        return value
    return int(round(float(value) * EMU_PER_INCH))


def rgb(value: str) -> RGBColor:
    value = value.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


TEXT = rgb(BASE["text"])
MUTED = rgb(BASE["muted"])
WHITE = rgb(BASE["white"])
BORDER = rgb(BASE["border"])
CODE_BG = rgb(BASE["code_bg"])
WARN = rgb(BASE["warn"])
WARN_TINT = rgb(BASE["warn_tint"])
OK = rgb(BASE["ok"])
OK_TINT = rgb(BASE["ok_tint"])

TOKEN_COLORS = {
    "plain": rgb("#24292E"),
    "keyword": rgb("#0B4F8A"),
    "type": rgb("#1F6F8B"),
    "builtin": rgb("#795E26"),
    "string": rgb("#A31515"),
    "comment": rgb("#6A737D"),
    "number": rgb("#098658"),
    "pp": rgb("#6D3AA8"),
    "var": rgb("#6D3AA8"),
    "punct": rgb("#24292E"),
}


@dataclass(frozen=True)
class Theme:
    accent: RGBColor
    tint: RGBColor
    edge: RGBColor
    secondary: RGBColor | None = None

    @classmethod
    def hex(cls, accent, tint, edge, secondary=None):
        return cls(rgb(accent), rgb(tint), rgb(edge), rgb(secondary) if secondary else None)


NEUTRAL = Theme.hex("#1F4E79", "#ECF3F9", "#C9DCEC", "#9FC0DC")
NIX = Theme.hex("#5277C3", "#EEF2FB", "#C7D3EE", "#7EBAE4")
GLSL = Theme.hex("#5586A4", "#EEF3F7", "#C6D6E1", "#9ABCCE")
SCSS = Theme.hex("#CC6699", "#FCEFF5", "#F2CCDD", "#E8A8C4")
NUSHELL = Theme.hex("#1C9438", "#F2F8E9", "#CFE3B6", "#8FD49A")
PYTHON = Theme.hex("#2E7D5B", "#EAF6F0", "#C5E3D5", "#9ED0B7")
CPP = Theme.hex("#1F4E79", "#ECF3F9", "#C9DCEC", "#9FC0DC")
RUST = Theme.hex("#B7410E", "#FBEEE7", "#EBC9B8", "#D99B7C")


@dataclass(frozen=True)
class Syntax:
    keywords: frozenset
    types: frozenset = frozenset()
    builtins: frozenset = frozenset()
    line_comment: str = "//"
    block_comment: tuple | None = None
    ident: str = r"[A-Za-z_]\w*"
    extra: str = ""


def _kw(text):
    return frozenset(text.split())


PY = Syntax(
    keywords=_kw(
        "and as assert async await break class continue def del elif else except "
        "False finally for from global if import in is lambda None not or pass "
        "raise return True try while with yield match case"
    ),
    types=_kw("int str float bool bytes list dict set tuple object"),
    builtins=_kw("self print open len range type enumerate zip map filter super"),
    line_comment="#",
)

NIX_SYNTAX = Syntax(
    keywords=_kw("let in with rec inherit if then else assert or import true false null"),
    builtins=_kw("builtins import throw abort map filter foldl head tail toString toJSON"),
    types=_kw("pkgs lib self stdenv"),
    line_comment="#",
    ident=r"[A-Za-z_][\w-]*",
)

GLSL_SYNTAX = Syntax(
    keywords=_kw(
        "const uniform buffer shared layout in out inout float double int uint "
        "void bool true false if else for while switch case default break "
        "continue return struct discard"
    ),
    types=_kw("vec2 vec3 vec4 ivec2 ivec3 ivec4 uvec2 mat2 mat3 mat4 sampler2D"),
    builtins=_kw("texture normalize clamp dot cross mix round length distance"),
    line_comment="//",
    block_comment=("/*", "*/"),
)

NU_SYNTAX = Syntax(
    keywords=_kw(
        "let mut const def export module use if else for in while loop match "
        "return try catch do where reduce each par-each echo print error"
    ),
    types=_kw("int float string bool list table record nothing any path duration"),
    builtins=_kw("open save ls get select first last length columns sort-by uniq"),
    line_comment="#",
    ident=r"[A-Za-z_][\w-]*",
    extra=r"|\$[A-Za-z_][\w-]*|--[A-Za-z][\w-]*",
)

SCSS_SYNTAX = Syntax(
    keywords=_kw(
        "@mixin @include @function @return @if @else @each @for @while @use "
        "@forward @extend @content @media @import"
    ),
    line_comment="//",
    block_comment=("/*", "*/"),
    ident=r"[A-Za-z_-][\w-]*",
    extra=r"|\$[A-Za-z_][\w-]*|@[A-Za-z-]+|#[0-9a-fA-F]{3,8}\b",
)

CPP_SYNTAX = Syntax(
    keywords=_kw(
        "auto break case class concept const constexpr consteval continue "
        "default delete do else enum explicit export extern false float for "
        "friend if inline int long namespace new noexcept nullptr operator "
        "override private protected public requires return short sizeof static "
        "static_assert struct switch template this throw true try typedef "
        "typename union unsigned using virtual void volatile while"
    ),
    types=_kw("std string vector size_t int8_t uint32_t FILE"),
    builtins=_kw("cout cin printf fopen fclose views ranges"),
    line_comment="//",
    block_comment=("/*", "*/"),
)

JS_SYNTAX = Syntax(
    keywords=_kw(
        "async await break case catch class const continue debugger default "
        "delete do else export extends finally for from function if import in "
        "instanceof let new of return static super switch this throw try "
        "typeof var void while with yield true false null undefined"
    ),
    builtins=_kw("console document window Promise Array Object JSON fetch"),
    line_comment="//",
    block_comment=("/*", "*/"),
)

RUST_SYNTAX = Syntax(
    keywords=_kw(
        "as async await break const continue crate dyn else enum extern false "
        "fn for if impl in let loop match mod move mut pub ref return self Self "
        "static struct super trait true type unsafe use where while"
    ),
    types=_kw("i8 i16 i32 i64 u8 u16 u32 u64 usize f32 f64 bool char str String Vec Option Result Box"),
    builtins=_kw("println format vec Some None Ok Err"),
    line_comment="//",
    block_comment=("/*", "*/"),
)

GO_SYNTAX = Syntax(
    keywords=_kw(
        "break case chan const continue default defer else fallthrough for func "
        "go goto if import interface map package range return select struct "
        "switch type var nil true false"
    ),
    types=_kw("int int8 int16 int32 int64 uint uint8 uint16 uint32 uint64 "
              "float32 float64 string bool byte rune error"),
    builtins=_kw("fmt make new len cap append copy panic recover"),
    line_comment="//",
    block_comment=("/*", "*/"),
)

BASH_SYNTAX = Syntax(
    keywords=_kw(
        "if then else elif fi for while do done case esac function return "
        "local export readonly in"
    ),
    builtins=_kw("echo printf cd ls grep sed awk cat mkdir rm cp mv source set unset"),
    line_comment="#",
    ident=r"[A-Za-z_][\w-]*",
    extra=r"|\$[A-Za-z_][\w-]*|\$\{[^}]*\}",
)

SQL_SYNTAX = Syntax(
    keywords=_kw(
        "select from where join left right inner outer on group by order having "
        "insert into values update set delete create table primary key foreign "
        "references not null as distinct limit offset union all case when then "
        "else end and or in is like"
    ),
    types=_kw("int integer bigint text varchar boolean timestamp date numeric serial"),
    builtins=_kw("count sum avg min max coalesce now"),
    line_comment="--",
    block_comment=("/*", "*/"),
)

SYNTAX = {
    "python": PY,
    "nix": NIX_SYNTAX,
    "glsl": GLSL_SYNTAX,
    "nu": NU_SYNTAX,
    "scss": SCSS_SYNTAX,
    "cpp": CPP_SYNTAX,
    "js": JS_SYNTAX,
    "ts": JS_SYNTAX,
    "rust": RUST_SYNTAX,
    "go": GO_SYNTAX,
    "bash": BASH_SYNTAX,
    "sql": SQL_SYNTAX,
}

GLYPHS = {
    "terminal": "\uf489",
    "shell": "\uf489",
    "cube": "\uf1b2",
    "chip": "\uf2db",
    "nixos": "\ue843",
    "sass": "\ue603",
    "python": "\ue73c",
    "rust": "\ue7a8",
    "go": "\ue627",
    "js": "\ue74e",
    "db": "\uf1c0",
    "git": "\ue702",
    "docker": "\uf308",
}

PROPER_NOUNS = {
    "Nix", "Nushell", "GLSL", "SCSS", "Sass", "Python", "Rust", "Go",
    "JavaScript", "TypeScript", "Java", "Kotlin", "Swift", "Haskell", "Lua",
    "C++", "C#", "SQL", "Linux", "NixOS", "Store", "WHNF", "REST", "API",
    "JSON", "CSS", "HTML", "OpenGL", "Vulkan", "WebGL", "GPU", "CPU",
    "Элко", "Элко Долстра",
}


@lru_cache(maxsize=None)
def font_path(name: str, bold: bool = False, italic: bool = False) -> str | None:
    style = "Bold" if bold else "Regular"
    if italic:
        style = "Bold Italic" if bold else "Italic"
    query = f"{name}:style={style}"
    try:
        out = subprocess.run(
            ["fc-match", "-f", "%{file}", query],
            capture_output=True, text=True, check=False, timeout=10,
        ).stdout.strip()
    except (OSError, subprocess.SubprocessError):
        return None
    return out or None


def text_width(text: str, size: float, font: str, bold=False, italic=False) -> float:
    path = font_path(font, bold, italic)
    if path is None:
        return 0.0
    from PIL import ImageFont
    probe = ImageFont.truetype(path, 100)
    return probe.getlength(text) * size / 100


def fit_size(text, max_emu, start, minimum, font, bold=False, italic=False):
    limit = max_emu / 12700
    size = float(start)
    while size > minimum and text_width(text, size, font, bold, italic) > limit:
        size -= 0.5
    return size


def lower_first(text: str) -> str:
    first = text.split()[0] if text.split() else ""
    if not first:
        return text
    if first.rstrip(",.") in PROPER_NOUNS:
        return text
    head, sep, tail = text.partition(" ")
    if head[0].isupper() and (not head[1:] or head[1:].islower()):
        return head[0].lower() + head[1:] + sep + tail
    return text


def tokenize(line: str, syntax: Syntax):
    parts = []
    if syntax.block_comment:
        open_, close = syntax.block_comment
        block = re.escape(open_) + r".*?" + re.escape(close)
    else:
        block = None
    pattern = re.compile(
        r'("(?:\\.|[^"\\])*"'
        r"|'(?:\\.|[^'\\])*'"
        + (r"|" + block if block else "")
        + r"|" + re.escape(syntax.line_comment) + r".*"
        + r"|\b\d+(?:\.\d+)?\b"
        + r"|" + syntax.ident + syntax.extra
        + r"|\s+|.)"
    )
    for tok in pattern.findall(line):
        if syntax.block_comment and tok.startswith(syntax.block_comment[0]):
            kind = "comment"
        elif tok.startswith(syntax.line_comment):
            kind = "comment"
        elif tok.startswith('"') or tok.startswith("'"):
            kind = "string"
        elif syntax.extra and tok.startswith(("$", "@")) and not tok.startswith("@media"):
            kind = "var" if tok.startswith("$") else "keyword"
        elif re.fullmatch(r"\d+(?:\.\d+)?", tok):
            kind = "number"
        elif re.fullmatch(syntax.ident, tok):
            if tok in syntax.keywords:
                kind = "keyword"
            elif tok in syntax.types:
                kind = "type"
            elif tok in syntax.builtins:
                kind = "builtin"
            else:
                kind = "plain"
        else:
            kind = "punct"
        parts.append((tok, kind))
    return parts


class Deck:
    """A deck on a fixed inch grid; every geometry argument is in inches."""

    def __init__(
        self,
        head: str = "Georgia",
        body: str = "Calibri",
        code: str = "DejaVu Sans Mono",
        glyph: str | None = None,
        width=W,
        height=H,
        margin=M,
    ):
        self.prs = Presentation()
        self.prs.slide_width = _emu(width)
        self.prs.slide_height = _emu(height)
        self.W = width / EMU_PER_INCH
        self.H = height / EMU_PER_INCH
        self.M = margin / EMU_PER_INCH
        self.CW = self.W - 2 * self.M
        self.top = CONTENT_TOP / EMU_PER_INCH
        self.bottom = TEXT_BOTTOM / EMU_PER_INCH
        self.head = head
        self.body = body
        self.code_font = code
        self.glyph_font = glyph or code

    def slide(self):
        return self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def style(self, run, size, color=TEXT, bold=False, italic=False, font=None):
        run.font.size = Pt(size)
        run.font.name = font or self.body
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        return run

    def bar(self, slide, theme: Theme, width=0.12):
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0,
                                     _emu(width), _emu(self.H))
        bar.fill.solid()
        bar.fill.fore_color.rgb = theme.accent
        bar.line.fill.background()
        bar.shadow.inherit = False
        return bar

    def glyph(self, slide, name_or_char, left, top, width, height, color, size=26):
        char = GLYPHS.get(name_or_char, name_or_char)
        box = slide.shapes.add_textbox(_emu(left), _emu(top), _emu(width), _emu(height))
        tf = box.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        run = p.add_run()
        run.text = char
        self.style(run, size, color, font=self.glyph_font)
        return box

    def title(self, slide, text, theme=NEUTRAL, size=26, mark=None, width=None):
        width = self.CW if width is None else width
        if mark:
            mark_w = 1.15 if len(text) < 30 else 0.75
            self.glyph(slide, mark, self.W - self.M - mark_w, 0.42,
                       mark_w, 0.6, theme.accent, size=28)
            width = width - mark_w - 0.25
        size = fit_size(text, _emu(width), size, 15, self.head, bold=True)
        box = slide.shapes.add_textbox(_emu(self.M), _emu(0.5), _emu(width), _emu(0.7))
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = text
        self.style(run, size, theme.accent, bold=True, font=self.head)
        rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, _emu(self.M), _emu(1.26),
                                      _emu(self.CW), Pt(0.9))
        rule.fill.solid()
        rule.fill.fore_color.rgb = BORDER
        rule.line.fill.background()
        rule.shadow.inherit = False
        return box

    def header(self, slide, text, theme=NEUTRAL, mark=None):
        self.bar(slide, theme)
        return self.title(slide, text, theme, mark=mark)

    def bullets(self, slide, items, left=None, top=None, width=None, height=3.4,
                size=15, space=12, color=TEXT, lowercase=False, bullet="• "):
        left = self.M if left is None else left
        top = self.top if top is None else top
        width = self.CW if width is None else width
        box = slide.shapes.add_textbox(_emu(left), _emu(top), _emu(width), _emu(height))
        tf = box.text_frame
        tf.word_wrap = True
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(space)
            p.line_spacing = 1.08
            if isinstance(item, tuple):
                lead, rest = item
                first = p.add_run()
                first.text = bullet + (lower_first(lead) if lowercase else lead)
                self.style(first, size, color)
                tail = p.add_run()
                tail.text = rest
                self.style(tail, size, color)
                continue
            text = lower_first(item) if lowercase else item
            run = p.add_run()
            run.text = bullet + text
            self.style(run, size, color)
        return box

    def key(self, slide, text, theme=NEUTRAL, top=5.7, size=13, color=None):
        box = slide.shapes.add_textbox(_emu(self.M), _emu(top), _emu(self.CW), _emu(0.55))
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = text
        self.style(run, size, color or theme.accent, italic=True)
        return box

    def box(self, slide, text, left, top, width, height, theme=NEUTRAL,
            fill=None, line=None, color=None, size=13, bold=False,
            align=PP_ALIGN.LEFT, radius=0.1, valign="middle"):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, _emu(left), _emu(top),
                                       _emu(width), _emu(height))
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill or WHITE
        shape.line.color.rgb = line or BORDER
        shape.line.width = Pt(0.9)
        shape.shadow.inherit = False
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.14)
        tf.margin_top = tf.margin_bottom = Inches(0.06)
        tf.vertical_anchor = MSO_ANCHOR.TOP if valign == "top" else MSO_ANCHOR.MIDDLE
        for i, line_text in enumerate(str(text).split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = 1.15
            run = p.add_run()
            run.text = line_text
            self.style(run, size, color or theme.accent, bold=bold)
        return shape

    def chevrons(self, slide, steps, theme=NEUTRAL, top=2.0, height=0.95,
                 gap=0.2, size=12):
        width = (self.CW - gap * (len(steps) - 1)) / len(steps)
        for i, step in enumerate(steps):
            x = self.M + i * (width + gap)
            shape = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, _emu(x), _emu(top),
                                           _emu(width), _emu(height))
            shape.fill.solid()
            shape.fill.fore_color.rgb = theme.accent if i < len(steps) - 1 else MUTED
            shape.line.fill.background()
            shape.shadow.inherit = False
            tf = shape.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.18)
            tf.margin_right = Inches(0.1)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = step
            self.style(run, size, WHITE, bold=True)

    def code(self, slide, label, source, lang="python", left=None, top=None,
             width=None, height=None, size=10, label_size=10.5, theme=NEUTRAL,
             bg=None, radius=0.03):
        left = self.M if left is None else left
        top = self.top if top is None else top
        width = self.CW if width is None else width
        tag = slide.shapes.add_textbox(_emu(left), _emu(top), _emu(width), _emu(0.26))
        p = tag.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = label
        self.style(run, label_size, theme.accent, bold=True)

        syntax = SYNTAX.get(lang, PY)
        lines = source.split("\n")
        line_h = size * 1.2 * 0.95 / 72
        needed = 0.3 + 0.26 + len(lines) * line_h
        box_h = needed if height is None else min(_inches(height) - 0.3, needed)

        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, _emu(left),
                                       _emu(top + 0.3), _emu(width), _emu(box_h))
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg or CODE_BG
        shape.line.color.rgb = BORDER
        shape.line.width = Pt(0.8)
        shape.shadow.inherit = False

        pad = 0.14
        frame = slide.shapes.add_textbox(_emu(left + pad), _emu(top + 0.3 + 0.08),
                                         _emu(width - 2 * pad), _emu(box_h - 0.14))
        tf = frame.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            p.line_spacing = 0.95
            tokens = tokenize(line, syntax) if line else [(" ", "plain")]
            for text, kind in tokens:
                run = p.add_run()
                run.text = text
                self.style(run, size, TOKEN_COLORS.get(kind, TEXT),
                           bold=(kind == "keyword"),
                           italic=(kind == "comment"),
                           font=self.code_font)
        return shape

    def table(self, slide, data, col_fracs=None, left=None, top=None, width=None,
              font_size=11, row_height=0.5, theme=NEUTRAL,
              header=True, first_col_bold=False):
        left = self.M if left is None else left
        top = self.top if top is None else top
        width = self.CW if width is None else width
        rows, cols = len(data), len(data[0])
        col_fracs = col_fracs or [1 / cols] * cols
        shape = slide.shapes.add_table(rows, cols, _emu(left), _emu(top),
                                       _emu(width), _emu(row_height))
        table = shape.table
        table.first_row = False
        table.horz_banding = False

        for r in range(rows):
            for c in range(cols):
                cell = table.cell(r, c)
                self._cell_border(cell)
                cell.fill.solid()
                head_row = header and r == 0
                cell.fill.fore_color.rgb = theme.accent if head_row else WHITE
                cell.margin_left = cell.margin_right = Inches(0.08)
                cell.margin_top = cell.margin_bottom = Inches(0.03)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = str(data[r][c])
                self.style(run, font_size, WHITE if head_row else TEXT,
                           bold=head_row or (first_col_bold and c == 0))

        total = sum(col_fracs)
        for c, frac in enumerate(col_fracs):
            table.columns[c].width = _emu(width * frac / total)
        for r in range(rows):
            table.rows[r].height = _emu(row_height)
        return shape

    @staticmethod
    def _cell_border(cell):
        tc_pr = cell._tc.get_or_add_tcPr()
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            for element in tc_pr.findall(qn(tag)):
                tc_pr.remove(element)
        for tag in ("a:lnB", "a:lnT", "a:lnR", "a:lnL"):
            ln = tc_pr.makeelement(qn(tag), {
                "w": str(int(Pt(0.6))), "cap": "flat", "cmpd": "sng", "algn": "ctr",
            })
            fill = ln.makeelement(qn("a:solidFill"), {})
            fill.append(ln.makeelement(qn("a:srgbClr"), {"val": "D5DBE2"}))
            ln.append(fill)
            tc_pr.insert(0, ln)

    def logo(self, slide, path, left, top, max_w, max_h):
        from PIL import Image
        with Image.open(path) as im:
            w, h = im.size
        scale = min(_emu(max_w) / w, _emu(max_h) / h)
        width, height = int(w * scale), int(h * scale)
        return slide.shapes.add_picture(
            path,
            _emu(left) + (_emu(max_w) - width) // 2,
            _emu(top) + (_emu(max_h) - height) // 2,
            width=width,
            height=height,
        )

    def logo_row(self, slide, entries, top, height, gap=0.3):
        from PIL import Image
        widths = []
        for path, _ in entries:
            with Image.open(path) as im:
                w, h = im.size
            widths.append(int(_emu(height) * w / h))
        cell = max(widths) + _emu(gap) * 2
        x0 = int((_emu(self.W) - cell * len(entries)) / 2)
        for i, ((path, label), width) in enumerate(zip(entries, widths)):
            x = int(x0 + i * cell + (cell - width) / 2)
            slide.shapes.add_picture(path, x, _emu(top), width=width, height=_emu(height))
            box = slide.shapes.add_textbox(int(x0 + i * cell), _emu(top + height + 0.1),
                                           int(cell), _emu(0.32))
            p = box.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = label
            self.style(run, 13, MUTED)

    def section(self, slide, name, theme, number=None, subtitle=None,
                mark=None, logo=None, notes=None):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, _emu(self.W), _emu(self.H))
        bg.fill.solid()
        bg.fill.fore_color.rgb = theme.tint
        bg.line.fill.background()
        bg.shadow.inherit = False
        self.bar(slide, theme)

        if number:
            box = slide.shapes.add_textbox(_emu(self.M), _emu(0.8), _emu(3.0), _emu(1.4))
            p = box.text_frame.paragraphs[0]
            run = p.add_run()
            run.text = str(number)
            self.style(run, 88, theme.secondary or theme.accent, bold=True, font=self.head)

        if logo:
            self.logo(slide, logo, 6.9, 1.7, 5.7, 3.4)
        elif mark:
            self.glyph(slide, mark, 7.2, 1.9, 4.6, 3.2, theme.accent, size=150)

        name_top = 2.6 if number else 2.2
        box = slide.shapes.add_textbox(_emu(self.M), _emu(name_top), _emu(6.2), _emu(1.1))
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = name
        self.style(run, 50, theme.accent, bold=True, font=self.head)

        rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, _emu(self.M),
                                      _emu(name_top + 1.15), _emu(2.2), Pt(2.2))
        rule.fill.solid()
        rule.fill.fore_color.rgb = theme.accent
        rule.line.fill.background()
        rule.shadow.inherit = False

        if subtitle:
            box = slide.shapes.add_textbox(_emu(self.M), _emu(name_top + 1.45),
                                           _emu(6.0), _emu(1.4))
            tf = box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.line_spacing = 1.15
            run = p.add_run()
            run.text = subtitle
            self.style(run, 19, TEXT)

        if notes:
            self.notes(slide, notes)

    def title_slide(self, title, subtitle=None, year=None, entries=None, theme=NEUTRAL):
        slide = self.slide()
        self.bar(slide, theme)
        box = slide.shapes.add_textbox(_emu(self.M), _emu(2.0), _emu(self.CW), _emu(1.9))
        tf = box.text_frame
        tf.word_wrap = True
        lines = title if isinstance(title, (list, tuple)) else [title]
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            p.space_before = Pt(8) if i else Pt(0)
            run = p.add_run()
            run.text = line
            self.style(run, 32, theme.accent, bold=True, font=self.head)

        if subtitle:
            box = slide.shapes.add_textbox(_emu(self.M), _emu(3.95), _emu(self.CW), _emu(0.5))
            p = box.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = subtitle
            self.style(run, 18, MUTED, font=self.head)

        if entries:
            self.logo_row(slide, entries, 3.1, 0.8)

        if year:
            box = slide.shapes.add_textbox(_emu(self.M), _emu(5.3), _emu(self.CW), _emu(0.4))
            p = box.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = str(year)
            self.style(run, 13, MUTED)
        return slide

    def text(self, slide, text, left, top, width, height=0.8, size=18,
             color=TEXT, bold=False, italic=False, font=None,
             align=PP_ALIGN.LEFT, line_spacing=None, anchor=None):
        box = slide.shapes.add_textbox(_emu(left), _emu(top), _emu(width), _emu(height))
        tf = box.text_frame
        tf.word_wrap = True
        for i, line_text in enumerate(str(text).split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            if line_spacing:
                p.line_spacing = line_spacing
            run = p.add_run()
            run.text = line_text
            self.style(run, size, color, bold=bold, italic=italic, font=font)
        if anchor:
            tf.vertical_anchor = anchor
        return box

    def notes(self, slide, text):
        slide.notes_slide.notes_text_frame.text = text

    def page_numbers(self, start=1):
        total = len(self.prs.slides)
        for i, slide in enumerate(self.prs.slides, start):
            box = slide.shapes.add_textbox(_emu(self.W - self.M - 0.6),
                                           _emu(self.H - 0.45), _emu(0.6), _emu(0.3))
            p = box.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT
            run = p.add_run()
            run.text = f"{i} / {total}"
            self.style(run, 9, MUTED)

    def properties(self, title="", subject="", author="", keywords="", comments="", category=""):
        cp = self.prs.core_properties
        cp.title = title
        cp.subject = subject
        cp.author = author
        cp.last_modified_by = author
        cp.keywords = keywords
        cp.comments = comments
        cp.category = category

    def save(self, path, thumbnail=True, page_numbers=False):
        if page_numbers:
            self.page_numbers()
        self.prs.save(path)
        if thumbnail:
            try:
                self._thumbnail(path)
            except Exception:
                pass
        return path

    def _thumbnail(self, path, width=1280, height=720):
        from PIL import Image, ImageDraw, ImageFont
        img = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(img)
        draw.rectangle([0, 0, 12, height], fill=(0x1F, 0x4E, 0x79))

        line = ""
        if len(self.prs.slides):
            first = self.prs.slides[0]
            texts = [sh.text_frame.text.strip() for sh in first.shapes
                     if sh.has_text_frame and sh.text_frame.text.strip()]
            line = texts[0].split("\n")[0] if texts else ""

        probe = font_path(self.head, bold=True)
        font = ImageFont.truetype(probe, 46) if probe else ImageFont.load_default()
        if line:
            box = draw.textbbox((0, 0), line, font=font)
            draw.text(((width - (box[2] - box[0])) / 2 - box[0], 300),
                      line, font=font, fill=(0x1F, 0x4E, 0x79))

        tmp = path + ".thumb.jpg"
        img.save(tmp, "JPEG", quality=88)
        try:
            with zipfile.ZipFile(path) as src, \
                    zipfile.ZipFile(path + ".tmp", "w", zipfile.ZIP_DEFLATED) as dst:
                for item in src.infolist():
                    data = src.read(item.filename)
                    if item.filename == "docProps/thumbnail.jpeg":
                        data = open(tmp, "rb").read()
                    dst.writestr(item, data)
            shutil.move(path + ".tmp", path)
        finally:
            os.remove(tmp)

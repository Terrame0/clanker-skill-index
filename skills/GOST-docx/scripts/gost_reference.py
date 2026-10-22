#!/usr/bin/env python3
"""Build a pandoc reference.docx matching the university formatting rules.

Rules (from the formatting guide, section 3.2):
  A4, margins: left 25 mm, right 10 mm, top/bottom 20 mm
  Times New Roman, 14 pt, line spacing 1.5, justified
  first-line indent 1.25 cm, section headings uppercase

Usage:
  python3 gost_reference.py <out.docx> <base_reference.docx>

Fetch the base reference with:
  pandoc --print-default-data-file reference.docx > base.docx
"""
import re
import sys
import zipfile

TWIP_PER_MM = 1440 / 25.4
MARGINS = {
    "top": round(20 * TWIP_PER_MM),
    "bottom": round(20 * TWIP_PER_MM),
    "left": round(25 * TWIP_PER_MM),
    "right": round(10 * TWIP_PER_MM),
}
FONT = "Times New Roman"
HALF_PT = 28
LINE = 360
FIRST_INDENT = round(12.5 * TWIP_PER_MM)
LANG = '<w:lang w:val="ru-RU" w:eastAsia="ru-RU"/>'
SECT_PR = (
    "<w:sectPr>"
    f'<w:pgSz w:w="11906" w:h="16838"/>'
    f'<w:pgMar w:top="{MARGINS["top"]}" w:right="{MARGINS["right"]}"'
    f' w:bottom="{MARGINS["bottom"]}" w:left="{MARGINS["left"]}"'
    ' w:header="720" w:footer="720" w:gutter="0"/>'
    "</w:sectPr>"
)

RUN_FONT = (
    f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}"'
    f' w:eastAsia="{FONT}" w:cs="{FONT}"/>'
    f'<w:sz w:val="{HALF_PT}"/><w:szCs w:val="{HALF_PT}"/>'
    + LANG
)
RUN_BOLD = (
    f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}"'
    f' w:eastAsia="{FONT}" w:cs="{FONT}"/>'
    f'<w:b/><w:bCs/><w:sz w:val="{HALF_PT}"/><w:szCs w:val="{HALF_PT}"/>'
    + LANG
)
BODY_PPR = (
    f'<w:spacing w:before="0" w:after="0" w:line="{LINE}" w:lineRule="auto"/>'
    f'<w:ind w:firstLine="{FIRST_INDENT}"/>'
    "<w:jc w:val=\"both\"/>"
)

WHERE_HANG = round(10 * TWIP_PER_MM)
TEXT_WIDTH = 11906 - MARGINS["left"] - MARGINS["right"]
FORMULA_TABS = (
    f'<w:tabs><w:tab w:val="center" w:pos="{TEXT_WIDTH // 2}"/>'
    f'<w:tab w:val="right" w:pos="{TEXT_WIDTH}"/></w:tabs>'
)

CODE_FONT = "Courier New"
CODE_HALF_PT = 22
CODE_PPR = (
    '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
    '<w:ind w:left="0" w:right="0" w:firstLine="0"/>'
    '<w:jc w:val="left"/>'
)
CODE_RPR = (
    f'<w:rFonts w:ascii="{CODE_FONT}" w:hAnsi="{CODE_FONT}"'
    f' w:eastAsia="{CODE_FONT}" w:cs="{CODE_FONT}"/>'
    f'<w:sz w:val="{CODE_HALF_PT}"/><w:szCs w:val="{CODE_HALF_PT}"/>'
)


def insert_styles(xml, new_styles):
    """Insert styles after docDefaults/latentStyles, as CT_Styles requires."""
    patterns = (
        r"<w:latentStyles\b[^>]*/>",
        r"</w:latentStyles>",
        r"</w:docDefaults>",
    )
    for pattern in patterns:
        match = re.search(pattern, xml)
        if match:
            return xml[: match.end()] + new_styles + xml[match.end():]
    return re.sub(
        r"(<w:styles\b[^>]*>)", lambda m: m.group(1) + new_styles, xml, count=1
    )


def replace_style(xml, style_id, new_style):
    pattern = re.compile(
        r'<w:style [^>]*w:styleId="' + re.escape(style_id) + r'".*?</w:style>', re.DOTALL
    )
    if not pattern.search(xml):
        return xml, False
    return pattern.sub(new_style, xml, count=1), True


def style(style_id, name, ppr="", rpr="", based=None, next_=None, custom=False):
    parts = [f'<w:style w:type="paragraph"{" w:customStyle=\"1\"" if custom else ""}'
             f' w:styleId="{style_id}">']
    parts.append(f'<w:name w:val="{name}"/>')
    if based:
        parts.append(f'<w:basedOn w:val="{based}"/>')
    if next_:
        parts.append(f'<w:next w:val="{next_}"/>')
    parts.append("<w:qFormat/>")
    if ppr:
        parts.append(f"<w:pPr>{ppr}</w:pPr>")
    if rpr:
        parts.append(f"<w:rPr>{rpr}</w:rPr>")
    parts.append("</w:style>")
    return "".join(parts)


TBLPR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
]


def insert_tblpr_child(inner, tag, fragment):
    """Insert a child into a tblPr body at the position CT_TblPr prescribes."""
    index = TBLPR_ORDER.index(tag)
    position = len(inner)
    for child in re.finditer(r"<w:([A-Za-z0-9]+)\b", inner):
        name = child.group(1)
        if name in TBLPR_ORDER and TBLPR_ORDER.index(name) > index:
            position = child.start()
            break
    return inner[:position] + fragment + inner[position:]


def patch_table_borders(xml):
    borders = (
        "<w:tblBorders>"
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        "</w:tblBorders>"
    )

    def patch(match):
        block = match.group(0)
        if "<w:tblBorders>" in block:
            return block
        tblpr = re.search(r"<w:tblPr>(.*?)</w:tblPr>", block, re.DOTALL)
        if not tblpr:
            return block
        inner = insert_tblpr_child(tblpr.group(1), "tblBorders", borders)
        return (
            block[: tblpr.start()]
            + "<w:tblPr>"
            + inner
            + "</w:tblPr>"
            + block[tblpr.end():]
        )

    return re.sub(
        r'<w:style [^>]*w:styleId="Table".*?</w:style>', patch, xml, flags=re.DOTALL
    )


def patch_doc_lang(xml):
    def patch(match):
        block = match.group(0)
        if "<w:lang " in block:
            return re.sub(r"<w:lang [^>]*/>", LANG, block, count=1)
        return block.replace("<w:rPr>", "<w:rPr>" + LANG, 1)

    return re.sub(
        r"<w:rPrDefault>.*?</w:rPrDefault>", patch, xml, count=1, flags=re.DOTALL
    )


def build_styles(xml):
    normal = style("Normal", "Normal", ppr=BODY_PPR, rpr=RUN_FONT)
    body = style(
        "BodyText", "Body Text",
        ppr=f'<w:spacing w:before="0" w:after="0" w:line="{LINE}" w:lineRule="auto"/>'
            f'<w:ind w:firstLine="{FIRST_INDENT}"/><w:jc w:val="both"/>',
        rpr=RUN_FONT, based="Normal",
    )
    first = style(
        "FirstParagraph", "First Paragraph",
        ppr=f'<w:ind w:firstLine="{FIRST_INDENT}"/><w:jc w:val="both"/>',
        rpr=RUN_FONT, based="BodyText", next_="BodyText", custom=True,
    )
    compact = style(
        "Compact", "Compact",
        ppr=f'<w:spacing w:before="0" w:after="0" w:line="{LINE}" w:lineRule="auto"/>'
            f'<w:jc w:val="both"/>',
        rpr=RUN_FONT, based="BodyText",
    )
    h1 = style(
        "Heading1", "heading 1",
        ppr=f'<w:keepNext/><w:keepLines/>'
            f'<w:spacing w:before="0" w:after="{LINE}" w:line="{LINE}" w:lineRule="auto"/>'
            f'<w:ind w:firstLine="{FIRST_INDENT}"/><w:jc w:val="left"/>'
            '<w:outlineLvl w:val="0"/>',
        rpr=RUN_BOLD,
        based="Normal", next_="BodyText",
    )
    h2 = style(
        "Heading2", "heading 2",
        ppr=f'<w:keepNext/><w:keepLines/>'
            f'<w:spacing w:before="{LINE}" w:after="{LINE}" w:line="{LINE}" w:lineRule="auto"/>'
            f'<w:ind w:firstLine="{FIRST_INDENT}"/><w:jc w:val="left"/>'
            '<w:outlineLvl w:val="1"/>',
        rpr=RUN_BOLD,
        based="Normal", next_="BodyText",
    )
    h3 = style(
        "Heading3", "heading 3",
        ppr='<w:keepNext/>'
            f'<w:spacing w:before="{LINE}" w:after="{LINE}" w:line="{LINE}" w:lineRule="auto"/>'
            f'<w:ind w:firstLine="{FIRST_INDENT}"/><w:jc w:val="left"/>'
            '<w:outlineLvl w:val="2"/>',
        rpr=RUN_BOLD,
        based="Normal", next_="BodyText",
    )
    caption = style(
        "Caption", "Caption",
        ppr=f'<w:spacing w:before="120" w:after="120" w:line="{LINE}" w:lineRule="auto"/>'
            '<w:jc w:val="center"/>',
        rpr=f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:cs="{FONT}"/>'
            '<w:sz w:val="24"/><w:szCs w:val="24"/>',
        based="Normal",
    )
    tabletext = style(
        "TableText", "Table Text",
        ppr='<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:ind w:left="0" w:right="0" w:firstLine="0"/>'
            '<w:jc w:val="left"/>',
        rpr=f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:cs="{FONT}"/>'
            '<w:sz w:val="24"/><w:szCs w:val="24"/>',
        based="Normal", custom=True,
    )
    tablecaption = style(
        "TableCaption", "Table Caption",
        ppr='<w:keepNext/>', based="Caption", custom=True,
    )
    h1c = style(
        "Heading1Center", "heading 1 centered",
        ppr='<w:keepNext/><w:keepLines/><w:pageBreakBefore/>'
            f'<w:spacing w:before="0" w:after="{LINE}" w:line="{LINE}" w:lineRule="auto"/>'
            '<w:ind w:firstLine="0"/><w:jc w:val="center"/>'
            '<w:outlineLvl w:val="0"/>',
        rpr=RUN_BOLD,
        based="Heading1", next_="BodyText", custom=True,
    )
    apptitle = style(
        "AppendixTitle", "Appendix Title",
        ppr='<w:keepNext/><w:keepLines/>'
            f'<w:spacing w:before="0" w:after="{LINE}" w:line="{LINE}" w:lineRule="auto"/>'
            '<w:ind w:firstLine="0"/><w:jc w:val="center"/>'
            '<w:outlineLvl w:val="9"/>',
        rpr=RUN_BOLD,
        based="Normal", next_="BodyText", custom=True,
    )
    sourcelist = style(
        "SourceList", "Source List",
        ppr=f'<w:ind w:left="0" w:right="0" w:firstLine="{FIRST_INDENT}"/>'
            '<w:jc w:val="left"/>',
        rpr=RUN_FONT, based="Normal", custom=True,
    )
    abbrev = style(
        "Abbrev", "Abbreviation",
        ppr='<w:ind w:left="0" w:right="0" w:firstLine="0"/>'
            '<w:jc w:val="left"/>',
        rpr=RUN_FONT, based="Normal", custom=True,
    )
    formula = style(
        "Formula", "Formula",
        ppr=FORMULA_TABS
            + '<w:spacing w:before="240" w:after="240"'
            f' w:line="{LINE}" w:lineRule="auto"/>'
            '<w:ind w:left="0" w:right="0" w:firstLine="0"/>'
            '<w:jc w:val="left"/>',
        rpr=RUN_FONT, based="Normal", custom=True,
    )
    where = style(
        "Where", "Where",
        ppr=f'<w:ind w:left="{FIRST_INDENT + WHERE_HANG}" w:hanging="{WHERE_HANG}"/>'
            '<w:jc w:val="both"/>',
        rpr=RUN_FONT, based="Normal", custom=True,
    )
    toc_heading = style(
        "TOCHeading", "TOC Heading",
        ppr='<w:keepNext/>'
            f'<w:spacing w:before="0" w:after="{LINE}" w:line="{LINE}" w:lineRule="auto"/>'
            '<w:ind w:firstLine="0"/><w:jc w:val="center"/>'
            '<w:outlineLvl w:val="9"/>',
        rpr=RUN_BOLD,
        based="Normal", next_="BodyText",
    )
    styles = [normal, body, first, compact, h1, h2, h3, caption, tabletext,
              tablecaption, h1c, apptitle, sourcelist, abbrev, formula, where,
              toc_heading]
    for lvl in range(1, 4):
        styles.append(style(
            f"TOC{lvl}", f"toc {lvl}",
            ppr=f'<w:tabs><w:tab w:val="right" w:leader="dot"'
                f' w:pos="{TEXT_WIDTH}"/></w:tabs>'
                f'<w:spacing w:before="0" w:after="0" w:line="{LINE}" w:lineRule="auto"/>'
                f'<w:ind w:left="{FIRST_INDENT * (lvl - 1)}" w:firstLine="0"/>'
                '<w:jc w:val="left"/>',
            rpr=RUN_FONT, based="Normal",
        ))
    for sid in ("SourceCode", "Verbatim"):
        styles.append(style(
            sid, sid if sid == "SourceCode" else "Verbatim",
            ppr=CODE_PPR, rpr=CODE_RPR, based="Normal", custom=True,
        ))
    for sid, name in (("Heading4", "heading 4"), ("Heading5", "heading 5"),
                      ("Heading6", "heading 6")):
        lvl = int(sid[-1]) - 1
        styles.append(style(
            sid, name,
            ppr=f'<w:keepNext/><w:outlineLvl w:val="{lvl}"/>',
            rpr=RUN_BOLD, based="Normal", next_="BodyText",
        ))
    for sid, name in (("Title", "Title"), ("Author", "Author"), ("Date", "Date")):
        styles.append(style(
            sid, name, ppr='<w:jc w:val="center"/>',
            rpr=RUN_BOLD, based="Normal",
        ))
    for sid, name, jc, bold in (
        ("TitleCenter", "Title Center", "center", False),
        ("TitleCenterBold", "Title Center Bold", "center", True),
        ("TitleLeft", "Title Left", "left", False),
        ("TitleRight", "Title Right", "right", False),
    ):
        styles.append(style(
            sid, name,
            ppr=f'<w:spacing w:before="0" w:after="0" w:line="{LINE}" w:lineRule="auto"/>'
                f'<w:ind w:firstLine="0"/><w:jc w:val="{jc}"/>',
            rpr=RUN_BOLD if bold else RUN_FONT,
            based="Normal", custom=True,
        ))
    styles.append(style(
        "PageBreak", "Page Break",
        ppr='<w:pageBreakBefore/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:ind w:left="0" w:right="0" w:firstLine="0"/>',
        rpr=f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}"'
            f' w:eastAsia="{FONT}" w:cs="{FONT}"/>'
            '<w:sz w:val="4"/><w:szCs w:val="4"/>',
        based="Normal", custom=True,
    ))
    new = "".join(styles)
    ids = re.findall(r'w:styleId="([^"]+)"', new)
    for sid in ids:
        xml = re.sub(
            r'<w:style [^>]*w:styleId="' + re.escape(sid) + r'".*?</w:style>',
            "", xml, flags=re.DOTALL,
        )
    xml = insert_styles(xml, new)
    xml = patch_table_borders(xml)
    xml = patch_doc_lang(xml)
    return xml


def main():
    out = sys.argv[1]
    base = sys.argv[2] if len(sys.argv) > 2 else None
    if not base:
        print("provide a base reference.docx", file=sys.stderr)
        sys.exit(2)
    zin = zipfile.ZipFile(base)
    styles = build_styles(zin.read("word/styles.xml").decode("utf-8"))
    document = zin.read("word/document.xml").decode("utf-8")
    document = re.sub(r"<w:sectPr>.*?</w:sectPr>", SECT_PR, document, flags=re.DOTALL)
    if "<w:sectPr" not in document:
        document = document.replace("</w:body>", SECT_PR + "</w:body>")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/styles.xml":
                data = styles.encode("utf-8")
            elif item.filename == "word/document.xml":
                data = document.encode("utf-8")
            zout.writestr(item, data)


if __name__ == "__main__":
    main()

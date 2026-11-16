"""Render README.md to README.pdf using reportlab directly.

Handles the Markdown subset used in the README: headings, paragraphs,
fenced code blocks, pipe tables, bullet lists, blockquotes, horizontal
rules, and inline `code` / **bold** / [links](url).
No native dependencies.
"""
import re
import html

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
    Preformatted, HRFlowable, ListFlowable, ListItem,
)

SRC, OUT = "README.md", "README.pdf"

INK = colors.HexColor("#1a1a1a")
DARK = colors.HexColor("#14202e")
BLUE = colors.HexColor("#2d6cdf")
GREY = colors.HexColor("#cfd6de")
HEADBG = colors.HexColor("#eef2f7")
ZEBRA = colors.HexColor("#f8fafc")
CODEBG = colors.HexColor("#f6f8fa")

styles = getSampleStyleSheet()
BODY = ParagraphStyle("body", parent=styles["Normal"], fontName="Helvetica",
                      fontSize=10.5, leading=15.5, textColor=INK, spaceAfter=4)
H1 = ParagraphStyle("h1", parent=BODY, fontName="Helvetica-Bold", fontSize=21,
                    leading=25, textColor=DARK, spaceBefore=6, spaceAfter=6)
H2 = ParagraphStyle("h2", parent=BODY, fontName="Helvetica-Bold", fontSize=14,
                    leading=18, textColor=DARK, spaceBefore=16, spaceAfter=4)
H3 = ParagraphStyle("h3", parent=BODY, fontName="Helvetica-Bold", fontSize=11.5,
                    leading=15, textColor=colors.HexColor("#24303d"),
                    spaceBefore=10, spaceAfter=3)
CELL = ParagraphStyle("cell", parent=BODY, fontSize=9, leading=12, spaceAfter=0)
CELLH = ParagraphStyle("cellh", parent=CELL, fontName="Helvetica-Bold", textColor=DARK)
CODE = ParagraphStyle("code", parent=styles["Code"], fontName="Courier",
                      fontSize=8.6, leading=11.5, textColor=colors.HexColor("#1a2530"),
                      backColor=CODEBG, borderPadding=(7, 7, 7, 7),
                      spaceBefore=4, spaceAfter=8, leftIndent=2, rightIndent=2)
QUOTE = ParagraphStyle("quote", parent=BODY, leftIndent=10, borderPadding=(6, 8, 6, 8),
                       backColor=colors.HexColor("#fff9e9"),
                       textColor=colors.HexColor("#4a3d13"), spaceBefore=6, spaceAfter=6)


def inline(text):
    """Markdown inline -> reportlab mini-HTML."""
    # protect code spans first
    codes = []
    def stash(m):
        codes.append(m.group(1))
        return f"\x00{len(codes)-1}\x00"
    text = re.sub(r"`([^`]+)`", stash, text)
    text = html.escape(text, quote=False)
    text = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"\[([^\]]+)\]\(([^)]+)\)",
                  lambda m: f'<link href="{m.group(2)}" color="#2d6cdf">{m.group(1)}</link>',
                  text)
    def restore(m):
        c = html.escape(codes[int(m.group(1))], quote=False)
        return (f'<font face="Courier" size="9" backColor="#f2f4f7" '
                f'color="#b0246b">{c}</font>')
    return re.sub(r"\x00(\d+)\x00", restore, text)


def make_table(rows):
    header, body = rows[0], rows[1:]
    data = [[Paragraph(inline(c), CELLH) for c in header]]
    for r in body:
        data.append([Paragraph(inline(c), CELL) for c in r])
    ncol = len(header)
    avail = A4[0] - 2 * 1.9 * cm
    tbl = Table(data, colWidths=[avail / ncol] * ncol, repeatRows=1, hAlign="LEFT")
    st = [
        ("GRID", (0, 0), (-1, -1), 0.5, GREY),
        ("BACKGROUND", (0, 0), (-1, 0), HEADBG),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    for i in range(2, len(data), 2):
        st.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    tbl.setStyle(TableStyle(st))
    return tbl


def parse(md):
    lines = md.split("\n")
    story, i = [], 0
    while i < len(lines):
        line = lines[i]
        # fenced code
        if line.strip().startswith("```"):
            i += 1
            buf = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                buf.append(lines[i]); i += 1
            i += 1
            story.append(Preformatted("\n".join(buf), CODE))
            continue
        # table
        if line.strip().startswith("|") and i + 1 < len(lines) and re.match(r"^\s*\|[\s:|-]+\|\s*$", lines[i + 1]):
            block = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                block.append(lines[i]); i += 1
            rows = []
            for r, raw in enumerate(block):
                if r == 1:
                    continue  # separator
                cells = [c.strip() for c in raw.strip().strip("|").split("|")]
                rows.append(cells)
            story.append(make_table(rows))
            story.append(Spacer(1, 6))
            continue
        # headings
        if line.startswith("### "):
            story.append(Paragraph(inline(line[4:]), H3)); i += 1; continue
        if line.startswith("## "):
            story.append(Paragraph(inline(line[3:]), H2)); i += 1; continue
        if line.startswith("# "):
            story.append(Paragraph(inline(line[2:]), H1)); i += 1; continue
        # hr
        if re.match(r"^\s*---+\s*$", line):
            story.append(HRFlowable(width="100%", thickness=0.7,
                                    color=colors.HexColor("#d5dbe2"),
                                    spaceBefore=10, spaceAfter=10)); i += 1; continue
        # blockquote
        if line.startswith(">"):
            buf = []
            while i < len(lines) and lines[i].startswith(">"):
                buf.append(lines[i].lstrip(">").strip()); i += 1
            story.append(Paragraph(inline(" ".join(buf)), QUOTE)); continue
        # list (bulleted or numbered)
        if re.match(r"^\s*([-*]|\d+\.)\s+", line):
            items = []
            while i < len(lines) and re.match(r"^\s*([-*]|\d+\.)\s+", lines[i]):
                txt = re.sub(r"^\s*([-*]|\d+\.)\s+", "", lines[i])
                items.append(ListItem(Paragraph(inline(txt), BODY), leftIndent=14))
                i += 1
            story.append(ListFlowable(items, bulletType="bullet",
                                      start="•", leftIndent=12, spaceAfter=6))
            continue
        # blank
        if not line.strip():
            i += 1; continue
        # paragraph (gather until blank / block start)
        buf = [line]
        i += 1
        while i < len(lines) and lines[i].strip() and not re.match(
            r"^(#|\||>|```|\s*([-*]|\d+\.)\s|\s*---+\s*$)", lines[i]
        ):
            buf.append(lines[i]); i += 1
        story.append(Paragraph(inline(" ".join(buf)), BODY))
    return story


def main():
    with open(SRC, encoding="utf-8") as f:
        md = f.read()
    doc = BaseDocTemplate(OUT, pagesize=A4,
                          leftMargin=1.9 * cm, rightMargin=1.9 * cm,
                          topMargin=1.7 * cm, bottomMargin=1.7 * cm,
                          title="Questionnaire Long to Wide Converter")
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="main")

    def footer(canvas, d):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#8894a3"))
        canvas.drawRightString(A4[0] - 1.9 * cm, 1.0 * cm, f"{d.page}")
        canvas.drawString(1.9 * cm, 1.0 * cm, "Questionnaire Long → Wide Converter")
        canvas.restoreState()

    doc.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=footer)])
    doc.build(parse(md))
    print("PDF written ->", OUT)


if __name__ == "__main__":
    main()

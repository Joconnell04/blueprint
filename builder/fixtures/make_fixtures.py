#!/usr/bin/env python3
"""make_fixtures.py: write the small xlsx and pptx files the extractor tests read.

Usage: python3 builder/fixtures/make_fixtures.py [out_dir]

The test suite generates into a temp directory, so nothing binary belongs in the
repository. Running this by hand with no argument writes into this folder; those
files are scratch, not fixtures to commit.

Three fixtures, all generated so nothing binary has to be committed:

  house.xlsx  a workbook built to the house rules: Arial 10, blue inputs, black
              formulas, grey italic notes, dashed label divider, medium navy
              title rule, gridlines off, currency/percent/multiple/count formats,
              a Read Me cover and a frozen header row.
  plain.xlsx  the same numbers with no styling at all: Calibri 11, no colour, no
              borders, gridlines on, General formats, no cover tab.
  deck.pptx   960x540 deck in the Harbor shape: Lato theme fonts, the Harbor
              colour scheme, a dark blue title slide, sky blue dividers, white
              content slides with a logo, a page number, a kicker, an action
              headline, three 2.25 pt roundRect cards and a source line.
  measure.pptx  a deck whose content measure and type weights are known: every
              body shape between x=60 and x=900 with one parked in the left
              margin and one running past the right edge, headlines in Montserrat
              Light, kicker and labels in Montserrat Medium, card headers in
              Montserrat SemiBold, body in Montserrat.

Every value here is a fixture value, not a house default; the tests assert that
the extractor recovers exactly these.
"""

import os
import struct
import sys
import zlib

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu, Pt

# ------------------------------------------------------------------ xlsx house

ARIAL = "Arial"
BLUE = "0000FF"
BLACK = "000000"
GREY = "595959"
NAVY = "1F3864"

CURRENCY = "$#,##0;($#,##0)"
PERCENT = "0.0%"
MULTIPLE = '0.0"x"'
COUNT = "#,##0"

LABEL_WIDTH = 24
VALUE_WIDTH = 12
NOTE_WIDTH = 40

THIN = Side(style="thin", color=BLACK)
DASHED = Side(style="dashed", color=BLACK)
MED_NAVY = Side(style="medium", color=NAVY)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LABEL_BOX = Border(left=THIN, right=DASHED, top=THIN, bottom=THIN)
TITLE_RULE = Border(top=THIN, bottom=MED_NAVY)
TOTAL_BOX = Border(left=THIN, right=THIN, top=MED_NAVY, bottom=THIN)
TOTAL_LABEL = Border(left=THIN, right=DASHED, top=MED_NAVY, bottom=THIN)


def _font(size=10, bold=False, italic=False, color=BLACK, name=ARIAL):
    return Font(name=name, size=size, bold=bold, italic=italic, color=color)


def _title(ws, row, col, width, text):
    for c in range(col, col + width):
        cell = ws.cell(row, c)
        cell.font = _font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = TITLE_RULE
    ws.cell(row, col).value = text


def _note(ws, row, col, text):
    cell = ws.cell(row, col)
    cell.value = text
    cell.font = _font(italic=True, color=GREY)


def _widths(ws, label_col, value_cols, note_col=None):
    ws.column_dimensions[label_col].width = LABEL_WIDTH
    for col in value_cols:
        ws.column_dimensions[col].width = VALUE_WIDTH
    if note_col:
        ws.column_dimensions[note_col].width = NOTE_WIDTH


def make_house_xlsx(path):
    """A four tab workbook in the house format."""
    wb = Workbook()
    wb.remove(wb.active)

    # ---- Read Me cover: the one place the single size rule relaxes.
    ws = wb.create_sheet("Read Me")
    ws.sheet_view.showGridLines = False
    ws["B2"] = "Fieldnote market model"
    ws["B2"].font = _font(size=16, bold=True, color=NAVY)
    ws["B3"] = "v1, internal draft"
    ws["B3"].font = _font(italic=True, color=GREY)
    for col in ("B", "C", "D", "E"):
        ws["%s3" % col].border = Border(bottom=MED_NAVY)
    ws["B5"] = "This workbook sizes the market and ramps revenue off one penetration input."
    ws["B5"].font = _font()
    _title(ws, 7, 2, 3, "Contents")
    for i, (tab, holds) in enumerate(
        (("Assumptions", "Inputs"), ("Model", "The build"), ("Sources & Appendix", "Refs"))
    ):
        ws.cell(8 + i, 2).value = tab
        ws.cell(8 + i, 2).font = _font()
        ws.cell(8 + i, 2).border = LABEL_BOX
        ws.cell(8 + i, 3).value = holds
        ws.cell(8 + i, 3).font = _font()
        ws.cell(8 + i, 3).border = BOX
    _widths(ws, "B", ("C",))

    # ---- Assumptions: every literal is an input, so every value cell is blue.
    ws = wb.create_sheet("Assumptions")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B6"
    _title(ws, 2, 2, 3, "Market frame")
    rows = [
        ("Base year", 2026, COUNT, "The model starts here"),
        ("US small businesses (M)", 33.2, COUNT, "SBA 2024 small business profile"),
        ("Serviceable share", 0.42, PERCENT, "Firms with 5 to 200 employees"),
        ("Price per seat", 480, CURRENCY, "List price, blended discount below"),
        ("Discount", 0.18, PERCENT, "Management assumption; no published figure"),
    ]
    for i, (label, value, nf, note) in enumerate(rows):
        r = 3 + i
        lab = ws.cell(r, 2)
        lab.value = label
        lab.font = _font()
        lab.border = LABEL_BOX
        val = ws.cell(r, 3)
        val.value = value
        val.font = _font(color=BLUE)
        val.border = BOX
        val.number_format = nf
        val.alignment = Alignment(horizontal="center", vertical="center")
        _note(ws, r, 4, note)
    _widths(ws, "B", ("C",), "D")

    # ---- Model: a table with a bold header row, formulas in black, a total row.
    ws = wb.create_sheet("Model")
    ws.sheet_view.showGridLines = False
    _title(ws, 2, 2, 4, "Revenue by segment")
    headers = ["Segment", "Seats", "Price", "Revenue"]
    for j, text in enumerate(headers):
        cell = ws.cell(3, 2 + j)
        cell.value = text
        cell.font = _font(bold=True)
        cell.border = LABEL_BOX if j == 0 else BOX
        cell.alignment = Alignment(horizontal="center", vertical="center")
    data = [("SMB", 1200, 480), ("Mid market", 640, 520), ("Public sector", 310, 450)]
    for i, (seg, seats, price) in enumerate(data):
        r = 4 + i
        lab = ws.cell(r, 2)
        lab.value = seg
        lab.font = _font()
        lab.border = LABEL_BOX
        for j, (value, nf) in enumerate(((seats, COUNT), (price, CURRENCY))):
            cell = ws.cell(r, 3 + j)
            cell.value = value
            cell.font = _font(color=BLUE)
            cell.border = BOX
            cell.number_format = nf
            cell.alignment = Alignment(horizontal="center", vertical="center")
        out = ws.cell(r, 5)
        out.value = "=C%d*D%d" % (r, r)
        out.font = _font(color=BLACK)
        out.border = BOX
        out.number_format = CURRENCY
        out.alignment = Alignment(horizontal="center", vertical="center")
    lab = ws.cell(7, 2)
    lab.value = "Total"
    lab.font = _font(bold=True)
    lab.border = TOTAL_LABEL
    for col in (3, 5):
        cell = ws.cell(7, col)
        cell.value = "=SUM(%s4:%s6)" % (chr(64 + col), chr(64 + col))
        cell.font = _font(bold=True, color=BLACK)
        cell.border = TOTAL_BOX
        cell.number_format = COUNT if col == 3 else CURRENCY
        cell.alignment = Alignment(horizontal="center", vertical="center")
    _note(ws, 8, 2, "Revenue is seats times price; see R12-3")

    _title(ws, 10, 2, 3, "Coverage check")
    checks = [("Seats per account", 3.4, MULTIPLE), ("Revenue per account", 2.1, MULTIPLE),
              ("Accounts reached", 640, COUNT)]
    for i, (label, value, nf) in enumerate(checks):
        r = 11 + i
        lab = ws.cell(r, 2)
        lab.value = label
        lab.font = _font()
        lab.border = LABEL_BOX
        cell = ws.cell(r, 3)
        cell.value = value
        cell.font = _font(color=BLACK)
        cell.border = BOX
        cell.number_format = nf
        cell.alignment = Alignment(horizontal="center", vertical="center")
    _widths(ws, "B", ("C", "D", "E"))

    # ---- Sources & Appendix
    ws = wb.create_sheet("Sources & Appendix")
    ws.sheet_view.showGridLines = False
    _title(ws, 2, 2, 4, "Sources")
    for j, text in enumerate(("Ref", "Source", "Key figures", "Link")):
        cell = ws.cell(3, 2 + j)
        cell.value = text
        cell.font = _font(bold=True)
        cell.border = LABEL_BOX if j == 0 else BOX
    for i, row in enumerate(
        (("R12-3", "SBA small business profile", "33.2M firms", "sba.gov"),
         ("R12-4", "Vendor price list", "$480 per seat", "example.com"))
    ):
        for j, text in enumerate(row):
            cell = ws.cell(4 + i, 2 + j)
            cell.value = text
            cell.font = _font()
            cell.border = LABEL_BOX if j == 0 else BOX
    _widths(ws, "B", ("C", "D"), "E")

    wb.save(path)
    return path


def make_plain_xlsx(path):
    """The same numbers with nothing on them: no colour, no borders, no formats."""
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Data")
    ws["A1"] = "Market frame"
    rows = [
        ("Base year", 2026),
        ("US small businesses", 33.2),
        ("Serviceable share", 0.42),
        ("Price per seat", 480),
        ("Discount", 0.18),
    ]
    for i, (label, value) in enumerate(rows):
        ws.cell(2 + i, 1).value = label
        ws.cell(2 + i, 2).value = value
    ws["A8"] = "Segment"
    ws["B8"] = "Seats"
    ws["C8"] = "Revenue"
    for i, (seg, seats) in enumerate((("SMB", 1200), ("Mid market", 640), ("Public sector", 310))):
        r = 9 + i
        ws.cell(r, 1).value = seg
        ws.cell(r, 2).value = seats
        ws.cell(r, 3).value = "=B%d*480" % r
    wb.save(path)
    return path


# ------------------------------------------------------------------ pptx deck

HARBOR_DARK = "1E3A5F"
HARBOR_SKY = "2A9D8F"
ST_SLATE = "264653"
ST_PALE = "E3EEF2"
ST_INK = "1F2937"
ST_GREY = "666666"
ST_ORANGE = "E9C46A"
WHITE = "FFFFFF"
LATO = "Lato"

SLIDE_W = 960
SLIDE_H = 540

THEME_COLOURS = [
    ("dk1", ST_INK), ("lt1", WHITE), ("dk2", ST_SLATE), ("lt2", ST_PALE),
    ("accent1", HARBOR_DARK), ("accent2", HARBOR_SKY), ("accent3", ST_ORANGE),
    ("accent4", "6B7280"), ("accent5", ST_PALE), ("accent6", ST_SLATE),
    ("hlink", HARBOR_SKY), ("folHlink", HARBOR_DARK),
]


def tiny_png(path, size=8, rgb=(26, 66, 138)):
    """An n x n solid PNG, written by hand so the fixtures need no image library."""
    raw = b"".join(b"\x00" + bytes(rgb) * size for _ in range(size))

    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(raw, 9))
    png += chunk(b"IEND", b"")
    with open(path, "wb") as handle:
        handle.write(png)
    return path


def _set_theme(prs):
    """Rewrite the theme part's colour scheme and fonts; python-pptx has no API for it."""
    part = prs.slide_masters[0].part.part_related_by(RT.THEME)
    xml = part.blob.decode("utf-8")
    clr = "".join(
        '<a:%s><a:srgbClr val="%s"/></a:%s>' % (name, value, name) for name, value in THEME_COLOURS
    )
    start = xml.index("<a:clrScheme")
    end = xml.index("</a:clrScheme>") + len("</a:clrScheme>")
    xml = xml[:start] + '<a:clrScheme name="Harbor">' + clr + "</a:clrScheme>" + xml[end:]
    start = xml.index("<a:fontScheme")
    end = xml.index("</a:fontScheme>") + len("</a:fontScheme>")
    fonts = (
        '<a:fontScheme name="Lato">'
        '<a:majorFont><a:latin typeface="Lato"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Lato"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
        "</a:fontScheme>"
    )
    xml = xml[:start] + fonts + xml[end:]
    part._blob = xml.encode("utf-8")


def _blank(prs, bg=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    if bg:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(bg)
    return slide


def _text(slide, x, y, w, h, text, size, color, bold=False, font=LATO, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    frame = box.text_frame
    frame.word_wrap = True
    para = frame.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = font
    run.font.color.rgb = RGBColor.from_string(color)
    return box


def _card(slide, x, y, w, h, header, sub, body):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
    shape.adjustments[0] = 0.08                      # prstGeom adj = 8000
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(WHITE)
    shape.line.color.rgb = RGBColor.from_string(HARBOR_DARK)
    shape.line.width = Pt(2.25)
    frame = shape.text_frame
    frame.word_wrap = True
    frame.margin_left = Pt(14)
    frame.margin_right = Pt(14)
    frame.margin_top = Pt(13)
    for i, (text, size, colour, bold) in enumerate(
        ((header, 17, HARBOR_DARK, True), (sub, 14, HARBOR_SKY, True), (body, 12, ST_GREY, False))
    ):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        run = para.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = LATO
        run.font.color.rgb = RGBColor.from_string(colour)
    return shape


def _page_number(slide, number, colour=ST_GREY):
    return _text(slide, 888, 25, 40, 14, str(number), 8, colour, align=PP_ALIGN.RIGHT)


def _content_slide(prs, logo, number, kicker, headline, cards, source):
    slide = _blank(prs, WHITE)
    slide.shapes.add_picture(logo, Pt(29), Pt(21), Pt(55), Pt(32))
    _page_number(slide, number)
    _text(slide, 32, 73, 500, 20, kicker, 16, ST_INK)
    _text(slide, 32, 108, 868, 70, headline, 28, ST_INK)
    for i, (header, sub, body) in enumerate(cards):
        _card(slide, 32 + i * 300, 197, 268, 210, header, sub, body)
    _text(slide, 32, 500, 400, 16, source, 10, ST_GREY)
    return slide


def make_deck_pptx(path, logo_path=None):
    """Five slides: title, divider, two content slides with cards, a second divider."""
    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W * 12700)
    prs.slide_height = Emu(SLIDE_H * 12700)
    _set_theme(prs)
    logo = logo_path or tiny_png(os.path.join(os.path.dirname(os.path.abspath(path)), "logo.png"))

    title = _blank(prs, HARBOR_DARK)
    _text(title, 30, 60, 500, 30, "Harbor", 20, HARBOR_SKY, bold=True)
    _text(title, 30, 330, 700, 60, "Donor growth: the next twelve months", 33, WHITE, bold=True)
    _text(title, 30, 400, 400, 20, "Working draft", 14, WHITE)

    divider = _blank(prs, HARBOR_SKY)
    _text(divider, 30, 81, 800, 100, "Where growth comes from", 66, WHITE, bold=True)
    _page_number(divider, 2, WHITE)

    _content_slide(
        prs, logo, 3,
        "Strategic Rationale | Why Referrals",
        "Referrals already carry a third of new partners, and the ask is the only step nobody owns",
        [("Open networks", "Reach", "Partners sit inside networks the team cannot buy into."),
         ("Warm entry", "Trust", "A referred partner arrives already believing the case."),
         ("Low cost", "Economics", "No media spend sits behind a referred partner.")],
        "Source: partner CRM, 2025",
    )

    _content_slide(
        prs, logo, 4,
        "Objective | The Ask Moment",
        "Asking at the moment of highest conviction is what turns a good meeting into a referral",
        [("Name the moment", "Timing", "The ask lands right after the partner says yes."),
         ("Make it small", "Scope", "One name, not a list, keeps the ask answerable."),
         ("Close the loop", "Follow up", "Report back on the referral so the next ask is easier.")],
        "Source: partner interviews, n=18",
    )

    divider2 = _blank(prs, HARBOR_SKY)
    _text(divider2, 30, 81, 800, 100, "Appendix", 66, WHITE, bold=True)
    _page_number(divider2, 5, WHITE)

    prs.save(path)
    return path


# ------------------------------------------------- a deck with a known content measure
#
# Q18 and Q20. The measure is the union of the body shapes, trimmed: every slide puts
# its content between x=60 and x=900, parks a legend in the left margin at x=30 and
# runs one band out to x=930, so a measure taken from the extremes reads 30/900 and
# only a trimmed one reads 60/840. The weights live in the face names, never in a bold
# flag, Montserrat Light headlines the builder kept writing bold.

MONTSERRAT = "Montserrat"
MONTSERRAT_LIGHT = "Montserrat Light"
MONTSERRAT_MEDIUM = "Montserrat Medium"
MONTSERRAT_SEMI = "Montserrat SemiBold"
INK = "111111"
COPPER = "D9A066"

# What the extractor has to read back: x and w from the trimmed body shapes, y from
# under the headline (100 + 60), h from there to the source line at 486.
MEASURE = {"x": 60, "y": 160, "w": 840, "h": 326}

# (x, y, w, h) per body shape on every content slide. The first is the legend in the
# margin and the last is the band that overruns; the eight between them are the measure.
MEASURE_BODY = [
    (30, 200, 24, 120), (60, 180, 260, 100), (340, 180, 260, 100), (620, 180, 280, 100),
    (60, 430, 200, 20), (340, 430, 260, 20), (60, 440, 870, 30),
]
MEASURE_CARDS = [(60, 300, 260, 120), (340, 300, 260, 120), (620, 300, 280, 120)]


def _weighted_card(slide, box, header, body):
    """A card whose header and body name their own Montserrat face, not a bold flag."""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, *[Pt(v) for v in box])
    shape.adjustments[0] = 0.04
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string("222222")
    shape.line.color.rgb = RGBColor.from_string("333333")
    shape.line.width = Pt(1)
    frame = shape.text_frame
    frame.word_wrap = True
    for i, (text, size, colour, font) in enumerate(
        ((header, 15, COPPER, MONTSERRAT_SEMI), (body, 12, WHITE, MONTSERRAT))
    ):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        run = para.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.name = font
        run.font.color.rgb = RGBColor.from_string(colour)
    return shape


def make_measure_pptx(path, logo_path=None):
    """A title slide and three content slides on one measure, weights in the face names."""
    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W * 12700)
    prs.slide_height = Emu(SLIDE_H * 12700)
    _set_theme(prs)
    logo = logo_path or tiny_png(os.path.join(os.path.dirname(os.path.abspath(path)), "logo.png"))

    title = _blank(prs, INK)
    _text(title, 60, 300, 700, 60, "The measure holds", 44, WHITE, font=MONTSERRAT_LIGHT)
    _text(title, 60, 380, 400, 20, "Working draft", 14, WHITE, font=MONTSERRAT)

    headlines = ("Every column starts on the same edge, so the deck reads as one grid",
                 "The measure is what the body shapes agree on, not what the margin allows",
                 "A legend in the margin is furniture and never widens the measure")
    for number, headline in enumerate(headlines, start=2):
        slide = _blank(prs, INK)
        slide.shapes.add_picture(logo, Pt(29), Pt(21), Pt(55), Pt(32))
        _text(slide, 888, 25, 40, 14, str(number), 9, "919191",
              font=MONTSERRAT, align=PP_ALIGN.RIGHT)
        _text(slide, 60, 60, 500, 20, "Measure | Slide %d" % number, 11, "919191",
              font=MONTSERRAT_MEDIUM)
        _text(slide, 60, 100, 840, 60, headline, 30, WHITE, font=MONTSERRAT_LIGHT)
        for i, box in enumerate(MEASURE_BODY):
            font = MONTSERRAT_MEDIUM if box[3] <= 30 else MONTSERRAT
            _text(slide, box[0], box[1], box[2], box[3],
                  "Body shape %d on the measure" % (i + 1), 12, WHITE, font=font)
        for i, box in enumerate(MEASURE_CARDS):
            _weighted_card(slide, box, "Column %d" % (i + 1),
                           "The card sits on the measure like every other body shape.")
        _text(slide, 60, 486, 400, 16, "Source: the fixture, by construction", 9,
              "7A7A7A", font=MONTSERRAT)

    prs.save(path)
    return path


# ---------------------------------------------------------------------- main


def build_all(out_dir):
    """Write every fixture into out_dir and return the paths."""
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    logo = tiny_png(os.path.join(out_dir, "logo.png"))
    return {
        "house_xlsx": make_house_xlsx(os.path.join(out_dir, "house.xlsx")),
        "plain_xlsx": make_plain_xlsx(os.path.join(out_dir, "plain.xlsx")),
        "deck_pptx": make_deck_pptx(os.path.join(out_dir, "deck.pptx"), logo),
        "measure_pptx": make_measure_pptx(os.path.join(out_dir, "measure.pptx"), logo),
        "logo_png": logo,
    }


def main(argv):
    out_dir = argv[1] if len(argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    made = build_all(out_dir)
    for name in sorted(made):
        sys.stdout.write("%s: %s\n" % (name, made[name]))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))

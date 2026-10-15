#!/usr/bin/env python3
"""make_fixtures.py: generate the test decks that import_pptx.py is scanned against.

Usage: python3 builder/fixtures-deck/make_fixtures.py [outdir]

Writes two .pptx files (960x540 pt, the 32x18 / 30 pt grid from the design spec):

  harbor_deck.pptx  a Harbor shaped deck: title slide, a card slide with a
                   kicker, an action title, three roundRect cards, a logo, a
                   page number and a source line, a sky-blue section divider,
                   and a slide with a takeaway bar.
  plain_deck.pptx  a plain deck: bullets, a native chart, a table, a picture,
                   a stat, a quote, a dashed placeholder and one deliberately
                   ambiguous little text box.

Every shape is placed on exact grid coordinates (x = (col-1)*30, y = (row-1)*30)
so the tests can assert snapping without tolerance. python-pptx only, no PIL:
the logo PNG is written by hand.
"""

import json
import os
import struct
import sys
import zlib

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

CELL = 30.0
SLIDE_W, SLIDE_H = 960.0, 540.0

DARK = RGBColor(0x1A, 0x42, 0x8A)
SKY = RGBColor(0x00, 0x9A, 0xDE)
PALE = RGBColor(0xDC, 0xEB, 0xF7)
INK = RGBColor(0x1F, 0x29, 0x37)
GREY = RGBColor(0x66, 0x66, 0x66)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


# ------------------------------------------------------------------ helpers


def png(path, w=64, h=40, rgb=(0x1A, 0x42, 0x8A)):
    """A valid truecolour PNG of one flat colour, written without PIL."""
    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    blob = (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw))
            + chunk(b"IEND", b""))
    with open(path, "wb") as fh:
        fh.write(blob)
    return path


def deck():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(SLIDE_W), Pt(SLIDE_H)
    return prs


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def box(slide, x, y, w, h):
    tb = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tf = tb.text_frame
    tf.word_wrap = True
    return tb


def write(shape, lines, size=12, bold=False, color=INK, align=None, anchor=None):
    """lines: [str] or [(text, size, bold, color, level)] - one paragraph each."""
    tf = shape.text_frame
    tf.word_wrap = True
    if anchor is not None:
        tf.vertical_anchor = anchor
    for i, line in enumerate(lines):
        spec = line if isinstance(line, tuple) else (line, size, bold, color, 0)
        text, sz, bd, col, level = (list(spec) + [size, bold, color, 0])[:5]
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.level = level
        if align is not None:
            p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(sz)
        run.font.bold = bool(bd)
        run.font.color.rgb = col
    return shape


def page_number(slide, n):
    tb = box(slide, 870, 18, 60, 18)
    write(tb, [str(n)], size=8, color=GREY, align=PP_ALIGN.RIGHT)
    return tb


def source_line(slide, text):
    tb = box(slide, 60, 510, 500, 18)
    write(tb, [text], size=8, color=GREY)
    return tb


def card(slide, x, y, w, h, header, sub, body):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = WHITE
    sh.line.color.rgb = DARK
    sh.line.width = Pt(2.25)
    sh.text_frame.vertical_anchor = MSO_ANCHOR.TOP
    write(sh, [(header, 17, True, DARK, 0), (sub, 14, True, SKY, 0), (body, 12, False, GREY, 0)])
    return sh


# ------------------------------------------------------------ harbor fixture


def harbor_deck(path, logo):
    prs = deck()

    # 1. title slide: a real title placeholder, so the headline rule takes the placeholder branch.
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text_frame.paragraphs[0].add_run().text = "Referral playbook for the fall ask"
    s.placeholders[1].text_frame.paragraphs[0].add_run().text = "Harbor | Fall 2026"

    # 2. card slide: furniture + three cards in the card box (cols 3, 13, 23 at 8 cols, rows 7 to 16).
    s = blank(prs)
    s.shapes.add_picture(logo, Pt(30), Pt(21), Pt(55), Pt(32))
    page_number(s, 2)
    write(box(s, 60, 73, 600, 20), ["Growth | Referrals"], size=16, color=INK)
    write(box(s, 60, 100, 840, 50), ["Referrals stall at the ask, not at the intro"], size=28, color=INK)
    card(s, 60, 180, 240, 300, "Open networks", "Reach", "Partners hold the relationships we never see.")
    card(s, 360, 180, 240, 300, "Warm intros", "Trust", "An intro from a peer converts four times better.")
    card(s, 660, 180, 240, 300, "Referral loop", "Repeat", "Every closed gift should return one new name.")
    source_line(s, "Source: Pilot data, 2026")

    # 3. divider: solid sky background, one huge line, page number only.
    s = blank(prs)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = SKY
    write(box(s, 30, 81, 870, 90), ["Where referrals stall"], size=66, color=WHITE)
    page_number(s, 3)

    # 4. content slide inside that section, with a takeaway bar on the bottom band.
    s = blank(prs)
    page_number(s, 4)
    write(box(s, 60, 73, 600, 20), ["Growth | Fixes"], size=16, color=INK)
    write(box(s, 60, 100, 840, 50), ["Three fixes close the referral loop"], size=28, color=INK)
    write(box(s, 60, 180, 420, 240), [
        "Partners ask late, after the gift has already closed.",
        "The follow-up never names the person who made the intro.",
    ], size=12, color=GREY)
    bar = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(60), Pt(480), Pt(840), Pt(30))
    bar.fill.solid()
    bar.fill.fore_color.rgb = PALE
    bar.line.fill.background()
    write(bar, ["Fix the ask and the loop closes itself"], size=13, color=INK)

    # 5. the second divider mechanic: a full-bleed rectangle instead of a slide background.
    s = blank(prs)
    back = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(0), Pt(0), Pt(SLIDE_W), Pt(SLIDE_H))
    back.fill.solid()
    back.fill.fore_color.rgb = RGBColor(0x32, 0x3E, 0x48)
    back.line.fill.background()
    write(box(s, 30, 81, 870, 90), ["How we fix it"], size=66, color=WHITE)
    page_number(s, 5)

    # 6. the third: the colour lives on the layout, and the text in its title placeholder.
    layout = prs.slide_layouts[5]
    layout.background.fill.solid()
    layout.background.fill.fore_color.rgb = DARK
    s = prs.slides.add_slide(layout)
    s.shapes.title.text_frame.paragraphs[0].add_run().text = "What it costs"
    page_number(s, 6)

    prs.save(path)
    return path


# ------------------------------------------------------------ plain fixture


def plain_deck(path, photo):
    prs = deck()

    def headline(slide, text):
        write(box(slide, 60, 60, 840, 50), [text], size=28, color=INK)

    # 1. bullets: literal bullet characters plus one indented level.
    s = blank(prs)
    headline(s, "What the pilot showed")
    write(box(s, 60, 180, 510, 240), [
        ("• Partners forget to ask", 14, False, INK, 0),
        ("• No follow-up after the intro", 14, False, INK, 0),
        ("• Referrals arrive without context", 14, False, INK, 0),
        ("Only one in five gets a reply", 14, False, GREY, 1),
    ])

    # 2. native chart.
    s = blank(prs)
    headline(s, "Referrals by quarter")
    data = CategoryChartData()
    data.categories = ["Q1", "Q2", "Q3", "Q4"]
    data.add_series("Referrals", (12, 19, 24, 31))
    s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Pt(60), Pt(180), Pt(540), Pt(300), data)

    # 3. table.
    s = blank(prs)
    headline(s, "Channels compared")
    table = s.shapes.add_table(4, 3, Pt(60), Pt(180), Pt(840), Pt(270)).table
    for c, label in enumerate(["Channel", "Donors", "Cost"]):
        table.cell(0, c).text = label
    for r, row in enumerate([["Referral", "128", "$41"], ["Event", "64", "$180"], ["Paid", "22", "$620"]], start=1):
        for c, val in enumerate(row):
            table.cell(r, c).text = val

    # 4. picture and a stat.
    s = blank(prs)
    headline(s, "One number carries the case")
    s.shapes.add_picture(photo, Pt(540), Pt(180), Pt(300), Pt(210))
    write(box(s, 60, 180, 240, 150), [("62%", 44, True, INK, 0), ("of donors give again", 12, False, GREY, 0)])

    # 5. quote.
    s = blank(prs)
    headline(s, "Partners say it plainly")
    write(box(s, 150, 240, 660, 120), ["“They ask, we deliver. Nobody closes the loop.”"],
          size=24, color=INK, align=PP_ALIGN.CENTER)
    write(box(s, 150, 390, 660, 30), ["— Pilot partner"], size=12, color=GREY, align=PP_ALIGN.CENTER)

    # 6. dashed placeholder and one ambiguous little box.
    s = blank(prs)
    headline(s, "The rollout still needs a shape")
    ph = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(60), Pt(180), Pt(420), Pt(240))
    ph.fill.background()
    ph.line.color.rgb = GREY
    ph.line.dash_style = 4  # MSO_LINE_DASH_STYLE.DASH
    write(ph, ["[PLACEHOLDER] Timeline of the pilot, quarter by quarter"], size=12, color=GREY)
    write(box(s, 690, 300, 120, 30), ["Q3 push"], size=11, color=GREY)

    prs.save(path)
    return path


def build_all(outdir):
    os.makedirs(outdir, exist_ok=True)
    logo = png(os.path.join(outdir, "logo.png"), 64, 40, (0x1A, 0x42, 0x8A))
    photo = png(os.path.join(outdir, "photo.png"), 96, 64, (0x99, 0xAA, 0xBB))
    return [harbor_deck(os.path.join(outdir, "harbor_deck.pptx"), logo),
            plain_deck(os.path.join(outdir, "plain_deck.pptx"), photo)]


def main(argv):
    outdir = argv[1] if len(argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    try:
        files = build_all(outdir)
    except Exception as exc:
        print(json.dumps({"ok": False, "errors": ["%s: %s" % (type(exc).__name__, exc)]}))
        return 1
    print(json.dumps({"ok": True, "files": files}))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))

#!/usr/bin/env python3
"""import_pptx.py: scan a .pptx into the Blueprint deck block model.

Usage: python3 builder/import_pptx.py <input.pptx> <out.json>

Writes a plan fragment (slides, blocks, warnings) to out.json and prints exactly
one JSON line on stdout. Exit 0 on success, 1 with {"ok": false, "errors": [...]}
on failure. Python 3.9, python-pptx only.

Roles (docs/superpowers/specs/2026-09-16-blueprint-design.md section 9,
docs/research/deck-planning.md sections 4 to 6):

  furniture (never a block)  title placeholder or largest top text -> the slide
                             headline; a short line above it -> the kicker;
                             page numbers and the logo picture -> dropped.
  blocks                     cards (repeated rounded rectangles), stat, chart,
                             table, image, quote, takeaway, placeholder, notes
                             (the source line), bullets, text.
  slides                     a solid accent background with one huge line and
                             nothing else is a divider; its text is the section,
                             and following slides inherit it.

Geometry is snapped to the 30 pt grid on 32 x 18 cells (cells are derived from
the real slide size, so a 4:3 deck still lands on 32 x 18); the point geometry
is preserved per block in `shape`. Every block carries a confidence in 0 to 1
and the signals that drove it.
"""

import json
import math
import os
import re
import sys

try:
    from pptx import Presentation
    from pptx.enum.dml import MSO_FILL
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.enum.text import PP_ALIGN
except ImportError:  # the one install line the server echoes back to the user
    print(json.dumps({"ok": False, "errors": [
        "python-pptx is not installed. Run: python3 -m pip install python-pptx"]}))
    sys.exit(1)

COLS, ROWS = 32, 18
EMU_PT = 12700.0
SAMPLE_CAP = 200

BULLET = "•‣▪●◦·⁃"
MARKER_RE = re.compile(r"^\s*([" + BULLET + r"]|[-–]\s|\d{1,2}[.)]\s)")
PAGENUM_RE = re.compile(r"^(page\s*)?\d{1,3}$", re.I)
STAT_RE = re.compile(r"^[+\-]?[$€£]?\d[\d,.\s]*\s*(%|x|×|k|m|bn|b|pp|\+)?$", re.I)
SOURCE_RE = re.compile(r"^\s*sources?\s*[:—-]?", re.I)
QUOTE_OPEN = "“„«\"'"
QUOTE_CLOSE = "”“»\"'"
PLACEHOLDER_RE = re.compile(r"\[\s*placeholder", re.I)
DASH_RE = re.compile(r'prstDash[^/>]*val="(?!solid)')
BUCHAR_RE = re.compile(r"<a:bu(Char|AutoNum)")

# A placeholder run usually carries no explicit size; the master's list style does.
PH_SIZE = {"TITLE": 40.0, "CENTER_TITLE": 40.0, "SUBTITLE": 24.0, "BODY": 18.0,
           "OBJECT": 18.0, "VERTICAL_TITLE": 40.0, "VERTICAL_BODY": 18.0}
BODY_SIZE = 18.0

CHART_KIND = {"BAR": "bar", "COLUMN": "bar", "LINE": "line", "PIE": "pie",
              "DOUGHNUT": "pie", "AREA": "line", "XY": "line", "SCATTER": "line",
              "RADAR": "line", "BUBBLE": "line", "STOCK": "line"}


def pt(v):
    return 0.0 if v is None else float(v) / EMU_PT


def snap(v, cell):
    """Round half up; Python's round() is banker's rounding and would wobble on .5 cells."""
    return int(math.floor(float(v) / cell + 0.5))


def one_line(text, cap=SAMPLE_CAP):
    return " ".join(str(text or "").split())[:cap]


def words(text, n=5):
    return " ".join(str(text or "").split()[:n]).rstrip(" ,.;:-–—")


def near_white(rgb):
    return rgb is not None and min(rgb) >= 240


# --------------------------------------------------------------- shape facts


class Sh(object):
    """One shape with everything the rules read, resolved once."""

    def __init__(self, shape, W, H):
        self.s = shape
        self.x, self.y = pt(shape.left), pt(shape.top)
        self.w, self.h = pt(shape.width), pt(shape.height)
        self.kind = self._kind()
        self.is_pic = self.kind == "picture"
        self.is_group = self.kind == "group"
        self.has_chart = bool(getattr(shape, "has_chart", False))
        self.has_table = bool(getattr(shape, "has_table", False))
        self.ph = self._ph()
        self.paras = self._paras()
        self.text = " / ".join(p["text"] for p in self.paras)
        self.max_size = max([p["size"] for p in self.paras if p["size"]] or [0.0])
        self.centred = any(p["align"] == PP_ALIGN.CENTER for p in self.paras)
        self.levels = max([p["level"] for p in self.paras] or [0])
        self.xml = self._xml()
        self.fill = self._fill()
        self.dashed = self._dashed()
        self.rounded = self._rounded()
        self.name = getattr(shape, "name", "") or ""
        self.alt = self._alt()
        self.W, self.H = W, H

    # -- identity

    def _kind(self):
        try:
            t = self.s.shape_type
        except Exception:
            return "other"
        if t == MSO_SHAPE_TYPE.PICTURE or t == MSO_SHAPE_TYPE.LINKED_PICTURE:
            return "picture"
        if t == MSO_SHAPE_TYPE.GROUP:
            return "group"
        if t == MSO_SHAPE_TYPE.AUTO_SHAPE:
            return "auto"
        if t == MSO_SHAPE_TYPE.TEXT_BOX:
            return "textbox"
        if t == MSO_SHAPE_TYPE.PLACEHOLDER:
            try:  # a picture placeholder holds an image and should read as one
                return "picture" if self.s.image is not None else "placeholder"
            except Exception:
                return "placeholder"
        return "other"

    def _ph(self):
        try:
            if self.s.is_placeholder:
                return str(self.s.placeholder_format.type).split(" ")[0]
        except Exception:
            pass
        return None

    def _paras(self):
        out = []
        if not getattr(self.s, "has_text_frame", False):
            return out
        for p in self.s.text_frame.paragraphs:
            text = " ".join(("".join(r.text for r in p.runs) or (p.text or "")).split())
            if not text:
                continue
            sizes = [r.font.size.pt for r in p.runs if r.font.size is not None]
            if not sizes and p.font.size is not None:
                sizes = [p.font.size.pt]
            out.append({"text": text, "size": max(sizes) if sizes else None,
                        "bold": any(bool(r.font.bold) for r in p.runs),
                        "level": p.level or 0, "align": p.alignment})
        return out

    def _xml(self):
        try:
            return self.s._element.xml
        except Exception:
            return ""

    def _fill(self):
        try:
            f = self.s.fill
            if f.type == MSO_FILL.SOLID:
                return tuple(bytes.fromhex(str(f.fore_color.rgb)))
        except Exception:
            pass
        return None

    def _dashed(self):
        try:
            d = self.s.line.dash_style
            if d is not None and "SOLID" not in str(d):
                return True
        except Exception:
            pass
        return bool(DASH_RE.search(self.xml))

    def _rounded(self):
        if self.kind != "auto":
            return False
        try:
            return "ROUND" in str(self.s.auto_shape_type)
        except Exception:
            return False

    def _alt(self):
        try:
            return self.s._element._nvXxPr.cNvPr.get("descr") or ""
        except Exception:
            return ""

    # -- derived

    @property
    def has_text(self):
        return bool(self.paras)

    @property
    def area(self):
        return self.w * self.h

    def size_or(self, default=BODY_SIZE):
        if self.max_size:
            return self.max_size
        return PH_SIZE.get(self.ph or "", default)

    def covers(self):
        return self.area / max(1.0, self.W * self.H)

    def bullets(self):
        marked = [p for p in self.paras if MARKER_RE.match(p["text"])]
        sigs = []
        if len(self.paras) >= 2 and len(marked) >= 2:
            sigs.append("%d lines start with a bullet character" % len(marked))
        if len(self.paras) >= 2 and self.levels > 0:
            sigs.append("indented paragraph levels")
        if BUCHAR_RE.search(self.xml):
            sigs.append("bullet formatting in the shape")
        return sigs


# ------------------------------------------------------------ furniture rules


def is_page_number(sh):
    return (sh.has_text and len(sh.paras) == 1 and PAGENUM_RE.match(sh.text.strip()) is not None
            and sh.size_or(11.0) <= 14.0 and sh.w <= 0.2 * sh.W
            and (sh.x > 0.72 * sh.W or sh.y < 0.10 * sh.H or sh.y > 0.85 * sh.H))


def is_logo(sh):
    if not sh.is_pic:
        return False
    if "logo" in (sh.name + " " + sh.alt).lower():
        return True
    return (sh.w <= 0.18 * sh.W and sh.h <= 0.12 * sh.H and sh.y <= 0.14 * sh.H
            and (sh.x <= 0.25 * sh.W or sh.x + sh.w >= 0.75 * sh.W))


def is_backdrop(sh):
    """A full-bleed solid rectangle behind everything. A full-bleed picture is content, not furniture."""
    return not sh.has_text and not sh.is_pic and sh.covers() >= 0.9 and sh.fill is not None


def background_of(part):
    """The solid background colour set on one slide or layout, or None when it inherits."""
    try:
        fill = part.background.fill
        if fill.type == MSO_FILL.SOLID:
            return tuple(bytes.fromhex(str(fill.fore_color.rgb)))
    except Exception:
        pass
    try:
        m = re.search(r"<p:bg>.*?<a:solidFill>.*?<a:srgbClr val=\"([0-9A-Fa-f]{6})\"", part._element.xml, re.S)
        return tuple(bytes.fromhex(m.group(1))) if m else None
    except Exception:
        return None


def slide_background(slide):
    """A divider's colour sits on the slide, or on the layout it was built from."""
    rgb = background_of(slide)
    if rgb is None:
        try:
            rgb = background_of(slide.slide_layout)
        except Exception:
            rgb = None
    return rgb


# --------------------------------------------------------------- block rules


def title_from(text, n=6):
    t = words(text, n)
    return t if t else "Text"


def classify(sh):
    """(type, title, content, confidence, signals) for one content shape."""
    sigs = []

    if sh.has_chart:
        kind, series, cats, ctitle = chart_facts(sh)
        sigs.append("native %s chart" % kind)
        if ctitle:
            sigs.append("chart title")
        content = "%s chart: %d series over %d categories" % (kind, series, cats)
        return "chart", ctitle or "Chart", content, 0.95, sigs, {"chartKind": kind}

    if sh.has_table:
        try:
            tbl = sh.s.table
            nrows, ncols = len(tbl.rows), len(tbl.columns)
            heads = [one_line(tbl.cell(0, c).text, 24) for c in range(min(ncols, 4))]
        except Exception:
            nrows, ncols, heads = 0, 0, []
        sigs.append("native table, %d rows x %d columns" % (nrows, ncols))
        content = "%dx%d table: %s" % (nrows, ncols, " / ".join([h for h in heads if h]) or "unlabelled")
        return "table", "Table", content, 0.95, sigs, {}

    if sh.is_pic:
        sigs.append("picture, %d x %d pt" % (round(sh.w), round(sh.h)))
        label = one_line(sh.alt or sh.name, 60)
        return "image", label or "Image", "image: %s" % (label or "no alt text"), 0.9, sigs, {}

    if sh.is_group:
        sigs.append("grouped shapes; roles inside a group are not read")
        return "text", title_from(sh.text) if sh.has_text else "Group", \
            "grouped shapes: %s" % (one_line(sh.text, 80) or "no text"), 0.35, sigs, {}

    if not sh.has_text:
        if sh.dashed:
            sigs.append("dashed outline, no text")
            return "placeholder", "Placeholder", "empty dashed placeholder", 0.7, sigs, {}
        return None

    text, first = sh.text, sh.paras[0]["text"]
    size = sh.size_or()

    # notes: the tiny source line at the bottom left
    if SOURCE_RE.match(first) and size <= 11.0 and sh.y >= 0.85 * sh.H:
        sigs.append("tiny text on the bottom band starting \"Source\"")
        if sh.x <= 0.3 * sh.W:
            sigs.append("bottom-left corner")
        return "notes", "Source", "source line: %s" % one_line(re.sub(SOURCE_RE, "", first), 90), \
            0.92 if sh.x <= 0.3 * sh.W else 0.8, sigs, {}

    # takeaway: a filled bar across the bottom of the slide
    if (sh.fill is not None and not near_white(sh.fill) and sh.w >= 0.7 * sh.W
            and sh.y >= 0.72 * sh.H and sh.h <= 3.2 * (sh.H / ROWS)):
        sigs.append("filled bar %d%% of the slide width" % round(100 * sh.w / sh.W))
        sigs.append("sits on the bottom band")
        return "takeaway", title_from(text), "takeaway bar: %s" % one_line(text, 90), 0.85, sigs, {}

    # placeholder: literal marker or a dashed outline
    if PLACEHOLDER_RE.search(text) or sh.dashed:
        conf = 0.75
        if PLACEHOLDER_RE.search(text):
            sigs.append("text marked [PLACEHOLDER]")
            conf = 0.9
        if sh.dashed:
            sigs.append("dashed outline")
            conf = min(0.95, conf + 0.1)
        label = one_line(PLACEHOLDER_RE.sub("", text).lstrip("] ").strip(), 80)
        return "placeholder", title_from(label or text), "placeholder: %s" % (label or "unlabelled"), conf, sigs, {}

    # quote: quoted text, usually centred
    stripped = text.strip()
    if stripped[0] in QUOTE_OPEN and len(stripped) <= 300:
        sigs.append("opens and closes with quote marks" if stripped[-1] in QUOTE_CLOSE else "opens with a quote mark")
        conf = 0.6
        if sh.centred:
            sigs.append("centred")
            conf += 0.15
        if stripped[-1] in QUOTE_CLOSE:
            conf += 0.1
        return "quote", title_from(stripped.strip(QUOTE_OPEN + QUOTE_CLOSE)), \
            "quote: %s" % one_line(stripped, 90), round(min(conf, 0.9), 2), sigs, {}

    # stat: a big number with a small label
    if STAT_RE.match(first) and (sh.paras[0]["size"] or size) >= 24.0:
        value = first
        label = sh.paras[1]["text"] if len(sh.paras) > 1 else ""
        sigs.append("number at %dpt" % round(sh.paras[0]["size"] or size))
        conf = 0.6
        if label:
            sigs.append("smaller label under the number")
            conf = 0.85
        else:
            sigs.append("no label in the same box")
        return "stat", one_line(label or value, 40), \
            "stat %s%s" % (value, (" — " + one_line(label, 60)) if label else ""), conf, sigs, \
            {"count": 1}

    # bullets vs prose
    blt = sh.bullets()
    if blt:
        sigs.extend(blt)
        items = [re.sub(MARKER_RE, "", p["text"]) for p in sh.paras]
        return "bullets", title_from(items[0]), \
            "%d bullets: %s" % (len(items), one_line(" / ".join(items[:3]), 110)), \
            0.85 if len(blt) > 1 else 0.75, sigs, {}

    lead = sh.paras[0]
    heading = (len(sh.paras) > 1 and len(lead["text"]) <= 60
               and (lead["bold"] or (lead["size"] or 0) > (sh.paras[1]["size"] or 0)))
    if heading:
        sigs.append("first line reads as a block title")
    sigs.append("%d paragraph%s of prose" % (len(sh.paras), "" if len(sh.paras) == 1 else "s"))

    conf = 0.7
    tiny = sh.area <= 0.04 * sh.W * sh.H
    short = len(text) <= 28
    if tiny and short and not heading:
        sigs.append("small floating box, %d words, no sentence" % len(text.split()))
        conf = 0.45
    elif len(sh.paras) == 1 and len(text) <= 40 and not heading:
        sigs.append("one short line; could be a label")
        conf = 0.55
    body = " / ".join(p["text"] for p in sh.paras[1:]) if heading else text
    return "text", lead["text"] if heading else title_from(text), \
        "%d paragraph%s: %s" % (len(sh.paras), "" if len(sh.paras) == 1 else "s", one_line(body, 110)), \
        conf, sigs, {}


def chart_facts(sh):
    kind, series, cats, title = "bar", 0, 0, ""
    try:
        ch = sh.s.chart
        name = str(ch.chart_type).split(" ")[0]
        for key, val in CHART_KIND.items():
            if name.startswith(key):
                kind = val
                break
        series = len(ch.plots[0].series) if ch.plots else 0
        cats = len(list(ch.plots[0].categories)) if ch.plots else 0
        if ch.has_title:
            title = one_line(ch.chart_title.text_frame.text, 60)
    except Exception:
        pass
    return kind, series, cats, title


def cluster_cards(cands):
    """Repeated rounded rectangles of a similar size, in a row or a grid -> one cards block."""
    groups = []
    used = set()
    for i, a in enumerate(cands):
        if i in used:
            continue
        group = [a]
        used.add(i)
        for j, b in enumerate(cands):
            if j in used:
                continue
            if (abs(a.w - b.w) <= 0.15 * max(a.w, b.w) and abs(a.h - b.h) <= 0.15 * max(a.h, b.h)
                    and (abs(a.y - b.y) <= 12 or abs(a.x - b.x) <= 12)):
                group.append(b)
                used.add(j)
        if len(group) >= 2:
            group.sort(key=lambda s: (round(s.y), round(s.x)))
            groups.append(group)
    return groups


def cards_block(group):
    heads = [g.paras[0]["text"] if g.paras else "" for g in group]
    rows = len(set(round(g.y / 12) for g in group))
    sigs = ["%d rounded rectangles of the same size" % len(group),
            "laid out in a %s" % ("row" if rows == 1 else "%d-row grid" % rows)]
    depth = max(len(g.paras) for g in group)
    if depth >= 2:
        sigs.append("header%s per card" % (" and body" if depth >= 3 else " and subhead"))
    conf = 0.85 if len(group) >= 3 else 0.7
    content = "%d cards: %s" % (len(group), one_line(" / ".join(h for h in heads if h), 120))
    return {"type": "cards", "title": "Cards", "content": content, "confidence": conf,
            "signals": sigs, "count": len(group),
            "text": one_line(" / ".join(g.text for g in group))}


# ------------------------------------------------------------------- slides


def geometry(sh, cw, ch):
    col = max(1, min(COLS, snap(sh.x, cw) + 1))
    row = max(1, min(ROWS, snap(sh.y, ch) + 1))
    cols = max(1, min(snap(sh.w, cw), COLS - col + 1))
    rows = max(1, min(snap(sh.h, ch), ROWS - row + 1))
    return {"col": col, "row": row, "cols": cols, "rows": rows,
            "shape": {"x": round(sh.x, 1), "y": round(sh.y, 1), "w": round(sh.w, 1), "h": round(sh.h, 1)}}


def span(group, cw, ch):
    class Box(object):
        pass
    b = Box()
    b.x = min(g.x for g in group)
    b.y = min(g.y for g in group)
    b.w = max(g.x + g.w for g in group) - b.x
    b.h = max(g.y + g.h for g in group) - b.y
    return geometry(b, cw, ch)


def scan_slide(slide, index, W, H, state, warnings):
    cw, ch = W / COLS, H / ROWS
    shapes = [Sh(s, W, H) for s in slide.shapes]
    out = {"index": index, "name": "", "story": "", "isDivider": False, "blocks": []}

    furniture = [s for s in shapes if is_page_number(s) or is_logo(s) or is_backdrop(s)]
    content = [s for s in shapes if s not in furniture]
    texts = [s for s in content if s.has_text]

    # divider: a solid accent behind one huge line and nothing else
    bg = slide_background(slide)
    if bg is None:
        backdrops = [s for s in furniture if is_backdrop(s) and s.fill is not None]
        bg = backdrops[0].fill if backdrops else None
    heavy = [s for s in content if s.has_chart or s.has_table or s.is_pic]
    # Anything else on the slide is empty by definition: only text shapes reach `texts`.
    if (bg is not None and not near_white(bg) and len(texts) == 1 and not heavy
            and texts[0].size_or() >= 30.0):
        section = one_line(texts[0].text, 90)
        state["section"] = section
        out.update({"isDivider": True, "section": section, "name": words(section) or "Slide %d" % index})
        return out

    # furniture: headline and kicker come off the slide, they are never blocks
    headline = next((s for s in texts if s.ph in ("TITLE", "CENTER_TITLE") and s.has_text), None)
    if headline is not None:
        head_sig = "title placeholder"
    else:
        band = [s for s in texts if s.y <= 0.28 * H and s.size_or() >= 15.0
                and not STAT_RE.match(s.paras[0]["text"]) and len(s.text) <= 220]
        headline = max(band, key=lambda s: (s.size_or(), -s.y, s.area)) if band else None
        head_sig = "largest text on the top band"
    kicker = None
    if headline is not None:
        above = [s for s in texts if s is not headline and len(s.paras) == 1 and len(s.text) <= 80
                 and s.y < headline.y and s.y + s.h <= headline.y + 8 and s.size_or() < headline.size_or()]
        kicker = max(above, key=lambda s: s.y) if above else None

    if headline is not None:
        out["headline"] = one_line(headline.text, 200)
        out["headlineSignal"] = head_sig
    if kicker is not None:
        out["kicker"] = one_line(kicker.text, 90)
    if state.get("section"):
        out["section"] = state["section"]
    out["name"] = words(out.get("headline", "")) or "Slide %d" % index

    rest = [s for s in content if s is not headline and s is not kicker]

    # cards first: a group of rounded rectangles is one block, not n text blocks
    taken = set()
    blocks = []
    card_cands = [s for s in rest if s.rounded and s.has_text
                  and not (s.fill and s.w >= 0.7 * W and s.y >= 0.72 * H)]
    for group in cluster_cards(card_cands):
        blk = cards_block(group)
        blk.update(span(group, cw, ch))
        blocks.append(blk)
        for g in group:
            taken.add(id(g))

    empties = 0
    for sh in rest:
        if id(sh) in taken:
            continue
        got = classify(sh)
        if got is None:
            empties += 1
            continue
        btype, title, content_line, conf, sigs, extra = got
        if sh.rounded and btype == "text":
            sigs.append("lone rounded rectangle, not a repeated card")
            conf = min(conf, 0.55)
        blk = {"type": btype, "title": one_line(title, 80) or btype.title(),
               "content": one_line(content_line, 140), "text": one_line(sh.text),
               "confidence": round(float(conf), 2), "signals": sigs}
        blk.update(extra)
        blk.update(geometry(sh, cw, ch))
        blocks.append(blk)

    if empties:
        warnings.append("slide %d: %d empty decorative shape%s ignored" % (index, empties, "" if empties == 1 else "s"))

    blocks.sort(key=lambda b: (b["row"], b["col"]))
    for n, b in enumerate(blocks, 1):
        b["id"] = "sld_%d_%d" % (index, n)
    out["blocks"] = blocks
    return out


def scan(path):
    warnings = []
    prs = Presentation(path)
    W, H = pt(prs.slide_width), pt(prs.slide_height)
    if W <= 0 or H <= 0:
        raise ValueError("slide size is zero")
    if abs(W / H - 16.0 / 9.0) > 0.02:
        warnings.append("slide is %dx%d pt, not 16:9; the %dx%d grid is stretched to fit"
                        % (round(W), round(H), COLS, ROWS))
    state, slides = {}, []
    for i, slide in enumerate(prs.slides, 1):
        slides.append(scan_slide(slide, i, W, H, state, warnings))
    seen = {}
    for s in slides:
        seen.setdefault(s["name"], []).append(s["index"])
    for name, idx in seen.items():
        if len(idx) > 1:
            warnings.append("slides %s share the name \"%s\"" % (", ".join(str(i) for i in idx), name))
    return {"source": os.path.abspath(path),
            "slideSize": {"w": round(W, 1), "h": round(H, 1), "cell": round(W / COLS, 2),
                          "cols": COLS, "rows": ROWS},
            "slides": slides, "warnings": warnings}


def main(argv):
    if len(argv) != 3:
        print(json.dumps({"ok": False, "errors": ["usage: import_pptx.py <input.pptx> <out.json>"]}))
        return 1
    src, out = argv[1], argv[2]
    try:
        plan = scan(src)
        with open(out, "w") as fh:
            json.dump(plan, fh, indent=2, default=str)
    except Exception as exc:
        print(json.dumps({"ok": False, "errors": ["%s: %s" % (type(exc).__name__, exc)]}))
        return 1
    blocks = [b for s in plan["slides"] for b in s["blocks"]]
    print(json.dumps({
        "ok": True,
        "slides": len(plan["slides"]),
        "blocks": len(blocks),
        "lowConfidence": len([b for b in blocks if b["confidence"] < 0.6]),
        "dividers": len([s for s in plan["slides"] if s["isDivider"]]),
        "warnings": plan["warnings"],
    }))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))

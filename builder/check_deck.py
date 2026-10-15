#!/usr/bin/env python3
"""Deck check after build (Blueprint deck mode).

Usage:
    python3 builder/check_deck.py <deck.pptx> [<out.json>] [--spec <spec.json>]

Reads the built deck with python-pptx, the sibling `<name>.snapshot.json` when
build_pptx.py wrote one (it carries the block id, role and slide kind of every
shape), and the build spec when given (it carries the style, so font floors are
the ones the deck was built with).

Checks, one finding per problem:

| kind | severity | what it means |
|---|---|---|
| `off-slide` | error | a shape sits outside the slide rectangle |
| `overlap` | error | two content shapes cover more than 10% of the smaller one |
| `missing-headline` | error | a content slide has no headline |
| `overflow` | warning | the text needs more lines than its box holds |
| `font-floor` | warning | a role was shrunk to its floor to fit. A role the build set
  once for the whole deck (the snapshot's `typeScale`) is reported once, with `slide: 0`
  and `scope: "deck"` |
| `empty` | warning | a box, card or placeholder with no content |
| `widow` | warning | a paragraph whose last line is one word |
| `diagram-edge` | error | a diagram edge names a node that does not exist |
| `diagram-overlap` | error | two node boxes in one diagram cover each other |
| `diagram-font` | warning | diagram labels are at the readable floor |

Consistency, the findings a reader would call "this deck was not proofread" (Q7):

| kind | severity | what it means |
|---|---|---|
| `kicker-drift` | warning | a slide's kicker does not follow the style's own pattern |
| `headline-case` | warning | a headline is title case, all caps, or ends in a full stop |
| `not-parallel` | warning | card bodies in one row do not end alike, or their line counts differ by more than one |
| `empty-band` | warning | more than 4 empty rows under the last block on a content slide. Strict only: balanced fills the band |
| `source-missing` | warning | fewer than half the slides that show figures carry a source. One finding, `slide: 0`, `scope: "deck"` |
| `bullet-count` | warning | more than 7 bullets in one block |
| `duplicate-headline` | warning | two consecutive slides carry the same headline |

Output (stdout, and the out.json path when given):

    { "ok": true, "slides": n, "shapes": n,
      "findings": [ { "slide", "blockId"?, "role"?, "kind", "severity", "message",
                      "scope"? } ],
      "summary": { "errors": n, "warnings": n } }

Exit status is 0 whenever the check ran, whatever it found. Exit 1 only when the
checker itself crashed; stdout then carries {"ok": false, "error": ...}.

Python 3.9, python-pptx only.
"""

import json
import os
import re
import sys

from pptx import Presentation

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import build_pptx  # noqa: E402
import fit  # noqa: E402

# ------------------------------------------------------------------ constants

SEVERITY = {
    "off-slide": "error",
    "overlap": "error",
    "missing-headline": "error",
    "empty-slide": "error",
    "diagram-edge": "error",
    "diagram-overlap": "error",
    "overflow": "warning",
    "font-floor": "warning",
    "empty": "warning",
    "widow": "warning",
    "diagram-font": "warning",
    "kicker-drift": "warning",
    "headline-case": "warning",
    "not-parallel": "warning",
    "empty-band": "warning",
    "source-missing": "warning",
    "bullet-count": "warning",
    "duplicate-headline": "warning",
}
KIND_ORDER = ["off-slide", "overlap", "missing-headline", "empty-slide", "diagram-edge",
              "diagram-overlap", "overflow", "font-floor", "empty", "widow",
              "diagram-font", "headline-case", "duplicate-headline", "kicker-drift",
              "not-parallel", "bullet-count", "empty-band", "source-missing"]

# A snapshot role that is not itself a type-scale key.
ROLE_TO_SCALE = {
    "card": "cardHeader",
    "stat": "statValue",
    "source": "caption",
    "table": "tableCell",
    "visual": "placeholder",
    "hairline": None,
    "notes": None,
    "wordmark": "wordmark",
    # A rating mark, a rule and the format's own marks are ink, not text.
    "tableRating": None,
    "rule": None,
    "chrome": None,
    # A template placeholder is the template's type, not ours (Q8).
    "templateBody": None,
    "templatePicture": None,
    # Diagram shapes are checked as a diagram, never as loose text.
    "diagramNode": None,
    "diagramLabel": None,
    "diagramEdge": None,
    "diagramAxis": None,
}

# A connector or an axis rule legitimately crosses the boxes it joins.
DIAGRAM_ROLES = ("diagramNode", "diagramLabel", "diagramEdge", "diagramAxis")
DIAGRAM_PASS_THROUGH = ("diagramEdge", "diagramAxis", "diagramLabel")
# A rating mark is drawn inside the table cell it belongs to, and the format's
# own chrome (dots, a hairline, a wordmark) sits wherever the format puts it.
PASS_THROUGH = ("hairline", "tableRating", "chrome")

# What dresses a slide rather than says anything on it. A folded headline or
# kicker block leaves its id on the furniture it filled, so the id alone does
# not make a shape content, and a rule or a mark never does (F-5).
FURNITURE_ROLES = ("headline", "kicker", "source", "pageNumber", "hairline", "rule",
                   "statRule", "tableRating", "chrome", "wordmark", "notes",
                   "dividerTitle", "dividerSubtitle", "coverWordmark", "coverTitle",
                   "coverSubtitle", "coverDate", "agendaEyebrow", "agendaStatement",
                   "agendaList")

# Consistency (Q7).
MAX_BULLETS = 7
EMPTY_BAND_ROWS = 4
TITLE_CASE_SHARE = 0.5
CARD_ROW_SLACK = 1.0        # points: cards in one row share a top edge
NUMBER_IN_TEXT = re.compile(
    r"\d+(?:\.\d+)?\s*%|[$\u00a3\u20ac\u00a5]\s?\d|\b\d{2,}\b|\b\d+(?:\.\d+)?\s*x\b",
    re.I)
ACRONYM = re.compile(r"^[A-Z0-9&/.-]{1,4}$")
SECTION_MARK = "QQSECTIONQQ"
TOPIC_MARK = "QQTOPICQQ"

EDGE = 0.5              # a shape may sit this many points past the slide edge
OVERLAP_SHARE = 0.10    # of the smaller shape's area
FULL_BLEED = 0.90       # a shape this big is a background, not content
HEIGHT_SLACK = 2.0      # points of measurement slop before text counts as overflow
EMU_PER_PT = 12700.0


def pt(value):
    return float(value) / EMU_PER_PT if value is not None else 0.0


def rect_of(shape):
    return (pt(shape.left), pt(shape.top), pt(shape.width), pt(shape.height))


def overlap_area(a, b):
    dx = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
    dy = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    if dx <= 0 or dy <= 0:
        return 0.0
    return dx * dy


# -------------------------------------------------------------------- reading


class Shape(object):
    """One drawn shape, flattened to what the checks need."""

    def __init__(self, shape, index):
        self.raw = shape
        self.index = index
        self.rect = rect_of(shape)
        self.name = str(getattr(shape, "name", "") or "")
        self.is_chart = bool(getattr(shape, "has_chart", False))
        self.is_table = bool(getattr(shape, "has_table", False))
        try:
            self.is_picture = "PICTURE" in str(shape.shape_type)
        except Exception:  # noqa: BLE001 - some shapes have no type
            self.is_picture = False
        self.paragraphs = []
        self.sizes = []
        self.text = ""
        self.inset_x = 0.0
        self.inset_y = 0.0
        self.blockId = None
        self.role = None
        self.snapshot_kind = None
        self.cell_text = ""
        if self.is_table:
            try:
                self.cell_text = " ".join(cell.text for row in shape.table.rows
                                          for cell in row.cells)
            except Exception:  # noqa: BLE001 - a table we cannot read is not a finding
                self.cell_text = ""
        if getattr(shape, "has_text_frame", False):
            frame = shape.text_frame
            self.inset_x = pt(frame.margin_left)
            self.inset_y = pt(frame.margin_top)
            for para in frame.paragraphs:
                text = "".join(run.text for run in para.runs)
                size = None
                for run in para.runs:
                    if run.font.size is not None:
                        size = max(size or 0.0, run.font.size.pt)
                if size:
                    self.sizes.append(size)
                self.paragraphs.append((text, size))
            self.text = frame.text or ""

    @property
    def area(self):
        return max(self.rect[2], 0.0) * max(self.rect[3], 0.0)

    @property
    def size(self):
        return max(self.sizes) if self.sizes else None

    def label(self):
        if self.blockId:
            return "block '%s'" % self.blockId
        if self.role:
            return "the %s" % self.role
        return "shape %d" % self.index


def kicker_form(pattern):
    """(regex for the full pattern, the separators it uses).

    A kicker is on pattern when it matches the filled pattern, or when it
    carries none of the pattern's separators: `fill_pattern` takes the dangling
    separator with an empty field, so a slide with no section reads "The prize",
    not "| The prize". A kicker that uses a different separator is drift.
    """
    text = build_pptx.fill_pattern(pattern, SECTION_MARK, TOPIC_MARK)
    body = re.escape(text).replace(SECTION_MARK, ".+").replace(TOPIC_MARK, ".+")
    seps = "".join(sorted(set(c for c in text if c in build_pptx.SEPARATORS)))
    return (re.compile("^" + body + "$", re.I) if text.strip() else None, seps)


def sentence_case(text):
    """Is this headline sentence case: one leading capital and nothing shouted.

    A word with its own capitals is left alone, so "AI" and "MetaDesign" are
    not title case; a majority of capitalised ordinary words is.
    """
    words = [w for w in re.split(r"\s+", str(text or "").strip()) if w]
    if not words:
        return True
    letters = [w for w in words if re.search(r"[A-Za-z]", w)]
    if len(letters) >= 3 and all(w.upper() == w for w in letters):
        return False
    capped = counted = 0
    for word in words[1:]:
        core = word.strip("\"'\u201c\u201d\u2018\u2019()[],.;:\u2014-")
        if not core or not core[0].isalpha():
            continue
        if ACRONYM.match(core) or any(c.isupper() for c in core[1:]):
            continue
        counted += 1
        if core[0].isupper():
            capped += 1
    if counted < 3:
        return True
    return capped <= TITLE_CASE_SHARE * counted


def wrapped_lines(shape, skip_title=False):
    """How many rendered lines a shape's paragraphs come to."""
    width = max(shape.rect[2] - 2 * shape.inset_x, 12.0)
    paras = [(t, s) for t, s in shape.paragraphs if t.strip() and s]
    if skip_title and len(paras) > 1 and paras[0][1] > paras[1][1]:
        paras = paras[1:]
    total = 0
    for text, size in paras:
        total += len(build_pptx.wrap_lines(text, build_pptx.cpl(width, size, 0.5)))
    return total


def live_paragraphs(shape, skip_title=False):
    """The paragraphs that carry copy, the block title dropped when asked."""
    paras = [(t, s) for t, s in shape.paragraphs if t.strip()]
    if (skip_title and len(paras) > 1 and paras[0][1] and paras[1][1]
            and paras[0][1] > paras[1][1]):
        paras = paras[1:]
    return paras


class Facts(object):
    """What the deck-wide consistency checks remember about one slide."""

    def __init__(self, number, kind):
        self.number = number
        self.kind = kind
        self.headline = ""
        self.kicker = ""
        self.has_source = False
        self.figures = False
        self.block_bottom = 0.0
        self.last_block = None


def read_snapshot(deck_path):
    path = build_pptx.snapshot_path(deck_path)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r") as handle:
            return json.load(handle)
    except Exception:  # noqa: BLE001 - a broken snapshot just means less context
        return None


def attach_snapshot(shapes, entries):
    """Match drawn shapes to snapshot entries by geometry, then by text."""
    left = list(entries or [])
    for shape in shapes:
        best = None
        for entry in left:
            if entry.get("kind") == "notes":
                continue
            dx = abs(float(entry.get("x", 0)) - shape.rect[0])
            dy = abs(float(entry.get("y", 0)) - shape.rect[1])
            dw = abs(float(entry.get("w", 0)) - shape.rect[2])
            dh = abs(float(entry.get("h", 0)) - shape.rect[3])
            if max(dx, dy, dw, dh) <= 1.0:
                best = entry
                break
        if best is None:
            flat = " ".join(shape.text.split())[:build_pptx.TEXT_PREVIEW]
            for entry in left:
                if flat and entry.get("text") == flat:
                    best = entry
                    break
        if best is not None:
            shape.blockId = best.get("blockId")
            shape.role = best.get("role")
            shape.snapshot_kind = best.get("kind")
            left.remove(best)


# ------------------------------------------------------------------- checking


class Checker(object):

    def __init__(self, deck_path, spec=None):
        self.prs = Presentation(deck_path)
        self.spec = spec or {}
        self.snapshot = read_snapshot(deck_path)
        self.warnings = []
        self.style, _ = build_pptx.make_style(self.spec, self.warnings)
        self.slide_w = pt(self.prs.slide_width)
        self.slide_h = pt(self.prs.slide_height)
        self.findings = []
        self.shape_count = 0
        # Roles the build set once for the whole deck: their floor is a deck
        # fact, so it is reported once, not on every slide that shows it.
        scale = (self.snapshot or {}).get("typeScale") or {}
        self.deck_sizes = dict((str(k), float(v)) for k, v in scale.items()
                               if isinstance(v, (int, float)))
        self._by_id = None
        self.fidelity = str((self.snapshot or {}).get("fidelity")
                            or fit.fidelity_of(self.spec))
        grid = (self.snapshot or {}).get("grid") or self.spec.get("grid") or {}
        self.rows = int(grid.get("rows") or build_pptx.GRID_ROWS)
        self.cell_h = self.slide_h / float(max(self.rows, 1))
        self.kicker_full, self.kicker_seps = kicker_form(self.style.kicker_pattern())
        self.facts = []

    def add(self, slide, kind, message, block_id=None, role=None, scope=None):
        finding = {"slide": slide, "kind": kind,
                   "severity": SEVERITY.get(kind, "warning"), "message": message}
        if block_id:
            finding["blockId"] = block_id
        if role:
            finding["role"] = role
        if scope:
            # A finding about the deck rather than a slide: slide 0, scope "deck".
            finding["scope"] = scope
        self.findings.append(finding)

    def spec_slide(self, index):
        """The spec slide this built slide came from, matched by its id.

        The builder inserts a generated cover and agenda, so the deck's
        position is not the spec's position. A generated slide has no spec
        slide of its own and gets an empty one, which skips every spec-based
        check rather than reading the wrong slide's blocks.
        """
        slides = self.spec.get("slides") or []
        if not slides:
            return {}
        snap = self.snapshot_slide(index)
        if not snap:
            return slides[index] if index < len(slides) else {}
        if self._by_id is None:
            self._by_id = {}
            for one in slides:
                key = str(one.get("id") or "")
                if key and key not in self._by_id:
                    self._by_id[key] = one
        key = str(snap.get("id") or "")
        if key:
            return self._by_id.get(key) or {}
        if len(slides) == len(self.snapshot.get("slides") or []):
            return slides[index] if index < len(slides) else {}
        return {}

    def snapshot_slide(self, index):
        if not self.snapshot:
            return {}
        slides = self.snapshot.get("slides") or []
        return slides[index] if index < len(slides) else {}

    def run(self):
        for i, slide in enumerate(self.prs.slides):
            number = i + 1
            snap = self.snapshot_slide(i)
            shapes = [Shape(s, n) for n, s in enumerate(slide.shapes)]
            attach_snapshot(shapes, snap.get("shapes"))
            self.shape_count += len(shapes)
            content = [s for s in shapes
                       if s.area > 1 and s.area < self.slide_w * self.slide_h * FULL_BLEED]
            self.check_off_slide(number, shapes)
            self.check_overlap(number, content)
            self.check_headline(number, i, shapes, snap)
            self.check_empty_slide(number, i, shapes, snap)
            self.check_diagrams(number, i, shapes)
            for shape in shapes:
                self.check_text(number, shape)
                self.check_bullets(number, shape)
            facts = self.read_facts(number, i, shapes, snap)
            self.facts.append(facts)
            self.check_kicker(facts)
            self.check_headline_shape(facts)
            self.check_cards(number, shapes)
            self.check_empty_band(facts)
        self.check_deck_type()
        self.check_sources()
        self.check_repeat_headlines()
        return self.result()

    # ------------------------------------------------------- consistency (Q7)

    def read_facts(self, number, index, shapes, snap):
        """One slide's facts for the consistency checks."""
        kind = snap.get("kind") or build_pptx.slide_kind(self.spec_slide(index))
        facts = Facts(number, kind)
        facts.headline = str(snap.get("headline") or "")
        for shape in shapes:
            if shape.role == "headline" and not facts.headline:
                facts.headline = shape.text
            if shape.role == "kicker":
                facts.kicker = " ".join(shape.text.split())
            if shape.role == "source":
                facts.has_source = True
            if not shape.blockId or shape.role in ("source", "pageNumber"):
                continue
            if NUMBER_IN_TEXT.search(shape.text) or NUMBER_IN_TEXT.search(shape.cell_text):
                facts.figures = True
            bottom = shape.rect[1] + shape.rect[3]
            if bottom > facts.block_bottom:
                facts.block_bottom = bottom
                facts.last_block = shape.blockId
        return facts

    def check_kicker(self, facts):
        """A kicker that does not follow the style's own pattern."""
        text = facts.kicker.strip()
        if facts.kind != "content" or not text:
            return
        if self.kicker_full is not None and self.kicker_full.match(text):
            return
        if not self.kicker_seps:
            return                       # a pattern with no separator takes anything
        if not any(ch in text for ch in build_pptx.SEPARATORS):
            return                       # one field, and its separator went with the other
        self.add(facts.number, "kicker-drift",
                 "the kicker reads '%s' where the style's pattern is '%s'; write it "
                 "to the pattern or leave it to the slide's section and name"
                 % (facts.kicker[:40], self.style.kicker_pattern()))

    def check_headline_shape(self, facts):
        """A headline in title case, in caps, or closed with a full stop."""
        text = " ".join(str(facts.headline or "").split())
        if not text or facts.kind not in ("content", "divider", "agenda"):
            return
        if text.endswith(".") and not text.endswith(".."):
            self.add(facts.number, "headline-case",
                     "the headline ends with a full stop; a slide headline is a "
                     "statement, not a sentence, so drop the stop")
        if not sentence_case(text):
            self.add(facts.number, "headline-case",
                     "the headline is not sentence case; capitalise the first word "
                     "and the proper nouns only")

    def check_cards(self, number, shapes):
        """Card bodies in one row that are not parallel."""
        rows = {}
        for shape in shapes:
            if shape.role != "card" or not shape.blockId:
                continue
            rows.setdefault((shape.blockId, round(shape.rect[1] / CARD_ROW_SLACK)),
                            []).append(shape)
        for key in sorted(rows, key=lambda k: (str(k[0]), k[1])):
            group = rows[key]
            if len(group) < 2:
                continue
            block_id = key[0]
            stops = set(" ".join(s.text.split()).rstrip().endswith(".") for s in group)
            if len(stops) > 1:
                self.add(number, "not-parallel",
                         "the cards in block '%s' do not close alike: some end in a "
                         "full stop and some do not; pick one and use it in every card"
                         % block_id, block_id, "card")
            counts = [wrapped_lines(s) for s in group]
            if counts and max(counts) - min(counts) > 1:
                self.add(number, "not-parallel",
                         "the cards in block '%s' run %d to %d lines; cut the longest "
                         "back so the row reads as one set"
                         % (block_id, min(counts), max(counts)), block_id, "card")

    def check_empty_band(self, facts):
        """A dead band under the last block. Balanced fills it, so strict only."""
        if facts.kind != "content" or self.fidelity != "strict":
            return
        if not facts.block_bottom:
            return
        floor = min(fit.BODY_BOTTOM, self.rows) * self.cell_h
        empty = (floor - facts.block_bottom) / self.cell_h
        if empty <= EMPTY_BAND_ROWS:
            return
        self.add(facts.number, "empty-band",
                 "%d empty rows sit under block '%s'; grow the last block or move the "
                 "band down, or build this deck on balanced fidelity"
                 % (int(empty), facts.last_block or "?"), facts.last_block)

    def check_bullets(self, number, shape):
        """More bullets in one block than a reader takes in."""
        if shape.role != "bullets":
            return
        count = len(live_paragraphs(shape, skip_title=True))
        if count > MAX_BULLETS:
            self.add(number, "bullet-count",
                     "%s carries %d bullets; cut to %d or split the block in two"
                     % (shape.label(), count, MAX_BULLETS), shape.blockId, shape.role)

    def check_sources(self):
        """A deck that shows figures and mostly does not say where they came from."""
        shown = [f for f in self.facts if f.kind == "content" and f.figures]
        if len(shown) < 2:
            return
        missing = [f.number for f in shown if not f.has_source]
        if len(shown) - len(missing) >= len(shown) / 2.0:
            return
        self.add(0, "source-missing",
                 "%d of the %d slides that show figures carry no source line "
                 "(slides %s); add a source to each or take the figure out"
                 % (len(missing), len(shown),
                    ", ".join(str(n) for n in missing[:8])),
                 scope="deck")

    def check_repeat_headlines(self):
        """Two consecutive slides saying the same thing."""
        previous = None
        for facts in self.facts:
            key = " ".join(str(facts.headline or "").lower().split())
            if key and previous and key == previous[0]:
                self.add(facts.number, "duplicate-headline",
                         "this headline repeats slide %d word for word; say what is "
                         "new on this slide or merge the two" % previous[1])
            previous = (key, facts.number)

    def check_deck_type(self):
        """The deck-wide sizes: one finding per role that sits at its floor.

        A role the build sets once for the whole deck is at the floor on every
        slide or on none, so it is a fact about the deck. The finding carries
        `slide: 0` and `scope: "deck"`, the only finding that does, so a view
        that lists slides can leave it out and a reader still sees it once.
        """
        for role in sorted(self.deck_sizes):
            spec = self.style.role(role)
            floor = float(spec.get("min", 0) or 0)
            default = float(spec.get("size", 0) or 0)
            if floor and default > floor and self.deck_sizes[role] <= floor:
                self.add(0, "font-floor",
                         "the deck-wide %s size sits at the %g pt floor" % (role, floor),
                         None, role, scope="deck")

    def check_diagrams(self, number, index, shapes):
        """Diagram findings: a broken edge, overlapping nodes, unreadable labels."""
        for block in self.spec_diagrams(index):
            content = block.get("content")
            content = content if isinstance(content, dict) else {}
            known = set()
            for i, node in enumerate(content.get("nodes") or []):
                if isinstance(node, dict) and str(node.get("id") or "").strip():
                    known.add(str(node.get("id")).strip())
                else:
                    known.add("n%d" % (i + 1))
            missing = []
            for edge in content.get("edges") or []:
                if not isinstance(edge, dict):
                    continue
                for end in ("from", "to"):
                    name = str(edge.get(end) or "").strip()
                    if name not in known:
                        missing.append(name or "?")
            if missing:
                self.add(number, "diagram-edge",
                         "block '%s' has an edge to a node that does not exist: %s"
                         % (block.get("id"), ", ".join(sorted(set(missing)))),
                         block.get("id"), "diagramNode")

        nodes = {}
        for shape in shapes:
            if shape.role != "diagramNode":
                continue
            nodes.setdefault(shape.blockId, []).append(shape)
        for block_id, boxes in nodes.items():
            for i, a in enumerate(boxes):
                for b in boxes[i + 1:]:
                    area = overlap_area(a.rect, b.rect)
                    if area <= 0.5:
                        continue
                    self.add(number, "diagram-overlap",
                             "two nodes in block '%s' overlap by %d points"
                             % (block_id, round(area)), block_id, "diagramNode")
            floor = float(self.style.role("diagramNode").get("min", 8))
            default = float(self.style.role("diagramNode").get("size", 12))
            sizes = [b.size for b in boxes if b.size]
            if sizes and default > floor and min(sizes) <= floor:
                self.add(number, "diagram-font",
                         "the labels in block '%s' are at the %g pt floor"
                         % (block_id, floor), block_id, "diagramNode")

    def spec_diagrams(self, index):
        """The diagram blocks the spec asked for, when a spec was given."""
        out = []
        for block in (self.spec_slide(index) or {}).get("blocks") or []:
            if str(block.get("type") or "").strip().lower() == "diagram":
                out.append(block)
        return out

    def check_off_slide(self, number, shapes):
        for shape in shapes:
            x, y, w, h = shape.rect
            if x < -EDGE or y < -EDGE or x + w > self.slide_w + EDGE or y + h > self.slide_h + EDGE:
                self.add(number, "off-slide",
                         "%s sits at %d,%d %dx%d, outside %dx%d"
                         % (shape.label(), round(x), round(y), round(w), round(h),
                            round(self.slide_w), round(self.slide_h)),
                         shape.blockId, shape.role)

    def check_overlap(self, number, shapes):
        for i, a in enumerate(shapes):
            for b in shapes[i + 1:]:
                if a.role in PASS_THROUGH or b.role in PASS_THROUGH:
                    continue
                if a.role in DIAGRAM_PASS_THROUGH or b.role in DIAGRAM_PASS_THROUGH:
                    continue
                if a.role == "diagramNode" and b.role == "diagramNode":
                    continue   # check_diagrams reports node overlap per diagram
                area = overlap_area(a.rect, b.rect)
                if area <= 0:
                    continue
                smaller = min(a.area, b.area)
                if smaller <= 0 or area / smaller <= OVERLAP_SHARE:
                    continue
                self.add(number, "overlap",
                         "%s and %s overlap by %d%% of the smaller"
                         % (a.label(), b.label(), round(100.0 * area / smaller)),
                         a.blockId or b.blockId, a.role or b.role)

    def check_headline(self, number, index, shapes, snap):
        kind = snap.get("kind") or build_pptx.slide_kind(self.spec_slide(index))
        if kind != "content":
            return
        if any(s.role == "headline" for s in shapes):
            return
        floor = float(self.style.role("headline").get("min", 24))
        top = self.slide_h / 3.0
        for shape in shapes:
            if shape.size and shape.size >= floor and shape.rect[1] < top and shape.text.strip():
                return
        self.add(number, "missing-headline", "this content slide has no headline")

    def check_empty_slide(self, number, index, shapes, snap):
        """A content slide left with nothing on it but its furniture (F-5).

        The duplicate pass drops a block that only repeated the headline, and a
        slide whose one block went that way keeps its kicker, its headline and
        its page number and says nothing else. The builder warns as it drops;
        the hole it leaves in the deck is an error here.
        """
        if not self.snapshot:
            return          # without the snapshot no shape carries a block id
        kind = snap.get("kind") or build_pptx.slide_kind(self.spec_slide(index))
        if kind != "content":
            return
        for shape in shapes:
            if shape.blockId and shape.role not in FURNITURE_ROLES:
                return
        name = str(snap.get("name") or "").strip()
        self.add(number, "empty-slide",
                 "slide %d%s holds nothing but furniture; add a block or drop the slide"
                 % (number, " '%s'" % name if name else ""))

    def check_text(self, number, shape):
        if shape.is_chart or shape.is_table or shape.is_picture:
            return
        if shape.role in DIAGRAM_ROLES:
            return          # check_diagrams owns every diagram shape
        if not shape.paragraphs:
            return
        text = shape.text.strip()
        role = shape.role
        scale_key = ROLE_TO_SCALE.get(role, role) if role else None
        if scale_key is None and role in ROLE_TO_SCALE:
            return
        spec = self.style.role(scale_key) if scale_key else None

        if not text:
            self.add(number, "empty", "%s has no text" % shape.label(),
                     shape.blockId, role)
            return
        bracketed = [t for t, _ in shape.paragraphs
                     if t.strip().startswith("[") and t.strip().endswith("]")]
        if bracketed and shape.snapshot_kind not in (None, "placeholder"):
            self.add(number, "empty", "%s only holds placeholder text" % shape.label(),
                     shape.blockId, role)
        if shape.snapshot_kind == "placeholder" and text in ("[]", "[ ]"):
            self.add(number, "empty", "%s is an unlabelled placeholder" % shape.label(),
                     shape.blockId, role)

        width = max(shape.rect[2] - 2 * shape.inset_x, 12.0)
        height = max(shape.rect[3] - 2 * shape.inset_y, 8.0)
        used = 0.0
        lines_total = 0
        for para_text, size in shape.paragraphs:
            if not para_text.strip() or not size:
                continue
            k = float(spec.get("k", 0.5)) if spec else 0.5
            lead = float(spec.get("leading", 1.25)) if spec else 1.25
            lines = build_pptx.wrap_lines(para_text, build_pptx.cpl(width, size, k))
            lines_total += len(lines)
            used += len(lines) * size * lead
            if len(lines) > 1 and len(lines[-1].split()) == 1:
                self.add(number, "widow",
                         "%s ends on the single word '%s'" % (shape.label(), lines[-1]),
                         shape.blockId, role)
        if used > height + HEIGHT_SLACK and lines_total:
            per_line = used / float(lines_total)
            self.add(number, "overflow",
                     "%s needs %d lines and the box holds %d"
                     % (shape.label(), lines_total, int(height / max(per_line, 1.0))),
                     shape.blockId, role)
        # Only a role that was shrunk counts: a caption whose default is its floor is fine.
        # A role the deck sets once is reported once, by check_deck_type.
        if (spec and shape.size is not None and scale_key not in self.deck_sizes
                and float(spec.get("size", 0)) > float(spec.get("min", 0))
                and shape.size <= float(spec.get("min", 0))):
            self.add(number, "font-floor",
                     "%s is at the %s floor of %g pt" % (shape.label(), scale_key,
                                                         float(spec.get("min"))),
                     shape.blockId, role)

    def result(self):
        order = dict((kind, i) for i, kind in enumerate(KIND_ORDER))
        findings = sorted(self.findings,
                          key=lambda f: (f["slide"], order.get(f["kind"], 99),
                                         str(f.get("blockId") or "")))
        errors = sum(1 for f in findings if f["severity"] == "error")
        result = {
            "ok": True,
            "slides": len(list(self.prs.slides)),
            "shapes": self.shape_count,
            "findings": findings,
            "summary": {"errors": errors, "warnings": len(findings) - errors},
        }
        if self.warnings:
            result["note"] = "; ".join(self.warnings)
        return result


def check_deck(deck_path, spec_path=None):
    spec = None
    if spec_path:
        with open(spec_path, "r") as handle:
            spec = json.load(handle)
    return Checker(deck_path, spec).run()


USAGE = "usage: python3 builder/check_deck.py <deck.pptx> [<out.json>] [--spec <spec.json>]"


def main(argv):
    spec_path = None
    args = []
    rest = list(argv[1:])
    while rest:
        item = rest.pop(0)
        if item in ("-h", "--help"):
            print(USAGE)
            return 0
        if item == "--spec":
            spec_path = rest.pop(0) if rest else None
        elif item.startswith("--spec="):
            spec_path = item.split("=", 1)[1]
        else:
            args.append(item)
    if not args:
        print(USAGE)
        return 2
    out_path = args[1] if len(args) > 1 else None
    try:
        result = check_deck(args[0], spec_path)
    except Exception as exc:  # noqa: BLE001 - one JSON object on stdout, always
        print(json.dumps({"ok": False, "error": "%s: %s" % (type(exc).__name__, exc)}))
        return 1
    text = json.dumps(result, indent=2)
    if out_path:
        with open(out_path, "w") as handle:
            handle.write(text + "\n")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))

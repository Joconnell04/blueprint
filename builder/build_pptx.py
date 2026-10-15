#!/usr/bin/env python3
"""build_pptx.py: build a .pptx deck from a Blueprint deck spec JSON.

Usage: python3 builder/build_pptx.py <spec.json>

Writes the deck, a sibling `<name>.snapshot.json`, a `<name>.pdf` when the spec
asks for it and LibreOffice is installed, and slide PNGs under `<name>.slides/`.
Prints exactly one JSON object on stdout, exits 0 on success or 1 with
{"ok": false, "errors": [...]} on failure. All logging goes to stderr.
Python 3.9 compatible, python-pptx only (`python3 -m pip install python-pptx`).

Contract: builder/SPEC-deck.md. Geometry, type ranges and furniture follow
docs/research/deck-planning.md.
"""

import copy
import glob
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_MARK
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import diagram_layout  # noqa: E402
import fit  # noqa: E402

# --------------------------------------------------------------- constants

SLIDE_W = 960.0
SLIDE_H = 540.0
GRID_COLS = 32
GRID_ROWS = 18
CELL = 30.0

EMU_PER_PT = 12700
HEX_RE = re.compile(r"^#?([0-9A-Fa-f]{6})$")

STYLES_DIR = os.path.join(HERE, "styles")

# Patched by the tests to simulate a machine without LibreOffice.
SOFFICE_PATHS = [
    "/opt/homebrew/bin/soffice",
    "/usr/local/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]
PDFTOPPM_PATHS = ["/opt/homebrew/bin/pdftoppm", "/usr/local/bin/pdftoppm"]
SOFFICE_TIMEOUT = 240
PDFTOPPM_TIMEOUT = 180
THUMB_DPI = 96

TEXT_PREVIEW = 80
CARD_WRAP_AT = 5          # more cards than this wrap to a second row
MIN_SIDE = 6.0            # a block never shrinks below this after clamping
KICKER_PATTERN = "{section} | {topic}"

# Fidelity: what content plainly does not suit its slot looks like (Q3, loose).
MIN_TABLE_COL_W = 64.0    # a table column narrower than this cannot hold a word
MIN_CARD_W = 150.0        # a card narrower than this holds a header and nothing
MAX_STAT_WORDS = 3        # a "stat" of more words is a sentence, not a number
MAX_STAT_CHARS = 20

# Roles the deck sets once for the whole deck (Q4). A group is fitted together,
# because the roles in it share a box: a card is a header over a body.
DECK_TYPE_GROUPS = (
    ("headline", ("headline",)),
    ("text", ("body",)),
    ("bullets", ("bullets",)),
    ("cards", ("cardHeader", "cardSub", "cardBody")),
    ("table", ("tableHeader", "tableCell")),
    ("stat", ("statValue", "statLabel")),
)
SEPARATORS = "|/\u00b7\u2013\u2014"   # a kicker separator the pattern may use

# Furniture a block is never allowed to sit on. The guard reads the drawn
# rectangles, so a style with its own geometry reserves its own bands.
GUARDED_ROLES = ("kicker", "headline", "hairline", "source", "pageNumber", "wordmark",
                 "coverTitle", "coverSubtitle", "coverDate", "agendaEyebrow",
                 "agendaStatement", "dividerTitle", "dividerSubtitle", "chrome",
                 "templateBody", "templatePicture")
FURNITURE_TYPES = ("headline", "kicker")

# Q15: the marks a format repeats on every slide, which are neither content nor
# a type role: the three dots top left, the hairline under the kicker row, a
# wordmark, a logo. House and harbor define none; the extractor fills them in.
CHROME_KINDS = ("dots", "line", "wordmark", "logo")
CHROME_ON = ("content",)       # the slide kinds an entry draws on by default
CHROME_DOTS = 3
CHROME_ROLE = "chrome"

# Q15: the divider, cover and agenda slots a style with its own furniture gets
# from its own measure instead of from house's 60/840.
DERIVED_SLOTS = ("dividerTitle", "dividerSubtitle", "coverWordmark", "coverTitle",
                 "coverSubtitle", "coverDate", "agendaEyebrow", "agendaStatement",
                 "agendaList")
SLOT_GAP = 18.0                # between a derived title and its subtitle

# Q15: a source line is one line. Over it, the detail in brackets goes first and
# then the trailing clauses, and the full text goes to the notes.
SOURCE_REASON = "source"
PARENTHETICAL = re.compile(r"\s*[(\[][^()\[\]]*[)\]]")
CLAUSE_SPLIT = ";"
SENTENCE_SPLIT = re.compile(r"(?<=[.!?])\s+")

# Q15: a block paragraph that repeats the headline is redundancy, not content.
DUPLICATE_OVERLAP = 0.8
DUPLICATE_REASON = "duplicate"
NOTES_REASON = "moved to notes"
PARALLEL_REASON = "parallel"
TEXT_LIKE = ("text", "bullets")

# Q15: chart discipline. Labels on a chart the reader can count, a legend only
# when there is more than one series to tell apart.
CHART_LABEL_POINTS = 8
CHART_SORTED_KINDS = ("bar", "column")
TIME_WORDS = ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct",
              "nov", "dec", "q1", "q2", "q3", "q4", "h1", "h2", "fy", "week", "wk",
              "day", "month", "quarter", "year", "phase", "sprint", "mon", "tue",
              "wed", "thu", "fri", "sat", "sun")
TIME_NUMBER = re.compile(r"^(19|20)\d{2}$")
NUMBER_UNITS = (("%", "%"), ("$", "$"), ("£", "£"), ("€", "€"))
UNIT_SUFFIX = re.compile(r"^([-+]?[\d,]*\.?\d+)\s*([%xX])?$")
SCALE_SUFFIX = {"k": 1000.0, "m": 1000000.0, "b": 1000000000.0}

# Where a block type's text sits in the type scale.
ROLE_FOR_TYPE = {
    "headline": "headline",
    "kicker": "kicker",
    "text": "body",
    "bullets": "bullets",
    "cards": "cardBody",
    "stat": "statValue",
    "chart": "caption",
    "table": "tableCell",
    "image": "placeholder",
    "quote": "quote",
    "takeaway": "takeaway",
    "placeholder": "placeholder",
    "diagram": "diagramNode",
    "notes": None,
}

CHART_KINDS = {
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "stacked": XL_CHART_TYPE.COLUMN_STACKED,
}

NO_STYLE_NO_GRID = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"

# The type scale. Sizes, floors, ceilings and line caps come from
# docs/research/deck-planning.md section 7; a style's `typeScale` overrides any key.
DEFAULT_TYPE_SCALE = {
    "coverTitle": {"size": 36, "min": 33, "max": 44, "maxLines": 2, "bold": True,
                   "font": "display", "color": "ink", "leading": 1.15, "k": 0.48},
    "coverSubtitle": {"size": 16, "min": 13, "max": 20, "maxLines": 2, "color": "muted"},
    "coverDate": {"size": 12, "min": 10, "max": 14, "maxLines": 1, "color": "muted"},
    "wordmark": {"size": 22, "min": 14, "max": 34, "maxLines": 1, "bold": True,
                 "font": "display", "color": "accent2"},
    "agendaEyebrow": {"size": 13, "min": 11, "max": 14, "maxLines": 1, "bold": True,
                      "font": "display", "color": "ink", "caps": True, "k": 0.55},
    "agendaStatement": {"size": 40, "min": 32, "max": 40, "maxLines": 3, "bold": True,
                        "font": "display", "color": "accent", "leading": 1.15, "k": 0.48},
    "dividerTitle": {"size": 66, "min": 54, "max": 72, "maxLines": 2, "bold": True,
                     "font": "display", "color": "ink", "leading": 1.1, "k": 0.48},
    "dividerSubtitle": {"size": 21, "min": 16, "max": 24, "maxLines": 3, "color": "muted"},
    "kicker": {"size": 12, "min": 11, "max": 16, "maxLines": 1, "color": "muted", "k": 0.52},
    # Q20: the headline's weight is the style's call, not the builder's. house
    # and plain say bold in their own token files; a style that says nothing
    # gets the family as it comes, which is what a light display face wants.
    "headline": {"size": 28, "min": 24, "max": 32, "maxLines": 2, "bold": False,
                 "font": "display", "color": "ink", "leading": 1.15, "k": 0.48},
    "statement": {"size": 36, "min": 28, "max": 40, "maxLines": 3, "bold": True,
                  "font": "display", "color": "ink", "leading": 1.15, "k": 0.48},
    "blockTitle": {"size": 14, "min": 12, "max": 16, "maxLines": 2, "bold": True, "color": "ink"},
    "body": {"size": 12, "min": 11, "max": 14, "maxLines": 6, "color": "muted"},
    "bullets": {"size": 14, "min": 12, "max": 16, "maxLines": 8, "color": "muted", "leading": 1.3},
    "cardHeader": {"size": 17, "min": 15, "max": 18, "maxLines": 2, "bold": True,
                   "font": "display", "color": "accent", "k": 0.48},
    "cardSub": {"size": 14, "min": 12, "max": 14, "maxLines": 1, "bold": True, "color": "accent2"},
    "cardBody": {"size": 12, "min": 11, "max": 14, "maxLines": 6, "color": "muted"},
    "statValue": {"size": 44, "min": 32, "max": 54, "maxLines": 1, "bold": True,
                  "font": "display", "color": "accent", "k": 0.6, "leading": 1.1},
    "statLabel": {"size": 12, "min": 11, "max": 13, "maxLines": 2, "color": "muted"},
    "quote": {"size": 24, "min": 20, "max": 28, "maxLines": 4, "font": "display", "color": "ink",
              "k": 0.48},
    "attribution": {"size": 12, "min": 11, "max": 14, "maxLines": 2, "color": "muted"},
    "tableHeader": {"size": 11, "min": 10, "max": 12, "maxLines": 1, "bold": True, "color": "ink"},
    "tableCell": {"size": 10, "min": 9, "max": 12, "maxLines": 1, "color": "ink", "k": 0.55},
    "caption": {"size": 8, "min": 8, "max": 10, "maxLines": 1, "color": "muted"},
    "pageNumber": {"size": 8, "min": 8, "max": 10, "maxLines": 1, "font": "mono",
                   "color": "muted", "k": 0.6},
    "takeaway": {"size": 13, "min": 11, "max": 16, "maxLines": 2, "bold": True, "color": "ink"},
    "placeholder": {"size": 16, "min": 11, "max": 16, "maxLines": 3, "italic": True,
                    "color": "muted"},
    "diagramNode": {"size": 12, "min": 8, "max": 16, "maxLines": 2, "color": "ink",
                    "leading": 1.2, "align": "center"},
    "diagramLabel": {"size": 9, "min": 8, "max": 12, "maxLines": 2, "bold": True,
                     "caps": True, "color": "muted", "k": 0.55},
}

ROLE_DEFAULTS = {"size": 12, "min": 9, "max": 40, "maxLines": 99, "bold": False,
                 "italic": False, "caps": False, "font": "body", "weight": "",
                 "color": "ink", "leading": 1.25, "k": 0.5, "align": "left"}

# Q10: a style's own type scale wins. When a style names a size and no floor,
# the floor is a readability limit and never the house floor: a caption, a
# label or a number reads at 8 pt, anything carrying a sentence at 9 pt.
CAPTION_FLOOR = 8.0
BODY_FLOOR = 9.0
CAPTION_ROLES = ("caption", "pageNumber", "attribution", "statLabel",
                 "tableHeader", "tableCell", "diagramLabel")

# Q19: a display role is a size, not a sentence. Letting a 66 pt divider title
# step to the 9 pt readability floor is not a smaller title, it is a different
# slide, so a display role a style gave no floor of its own stops at 85% of the
# size it asked for and the overflow is reported instead.
DISPLAY_ROLES = ("headline", "dividerTitle", "coverTitle", "statValue",
                 "agendaStatement")
DISPLAY_FLOOR = 0.85

# Q20: `typeScale.<role>.weight` names a face rather than a bold flag. The face
# is written when the family has it installed ("Montserrat SemiBold"), and the
# family plus the bold attribute when it does not.
# Regular and bold are the family's own RIBBI faces: a renderer reaches them
# through the bold attribute, and there is no family called "Montserrat Regular"
# to write. Only the weights outside that set ship as families of their own.
WEIGHT_FACES = {"light": "Light", "medium": "Medium", "semibold": "SemiBold"}
WEIGHT_RIBBI = {"regular": False, "bold": True}
WEIGHT_BOLD = ("semibold",)

# F-1: `kicker.case` recases a derived kicker so it matches the format's own
# convention. An explicit kicker in the payload is left exactly as written.
KICKER_CASES = ("title", "sentence", "caps", "none")
# The words Title Case leaves lower inside a phrase. First and last always rise.
TITLE_SMALL = ("a", "an", "and", "as", "at", "but", "by", "for", "from", "in",
               "into", "nor", "of", "on", "onto", "or", "per", "the", "to",
               "v", "via", "vs", "with")

# Q23: `table.labelColumn` is how the first column reads when it names the rows.
LABEL_COLUMN_DEFAULT = {"case": "none", "color": "ink", "bold": True}
LABEL_COLUMN_CASES = ("caps", "none")

# Q22: `stat` `layout: "grid"` sets number and label pairs in two columns with a
# hairline under each pair, the way a deck lists supporting figures.
STAT_LAYOUTS = ("band", "grid")
STAT_GRID_ITEMS = 4        # this many figures in a tall block picks grid on its own
STAT_GRID_ROWS = 6         # "tall": more than this many grid rows
STAT_GRID_COLS = 2         # pairs across
STAT_GRID_VALUE = 0.6      # the grid number is statValue x this
STAT_GRID_GUTTER = 30.0    # between the two columns
STAT_GRID_PAD = 10.0       # between a number and its label, and above a hairline

# F-5: `stat.rule` says which side of a figure its rule sits on. House rules
# under the value; a format that heads each figure with a rule says "above".
STAT_RULES = ("above", "below", "none")
STAT_RULE_DEFAULT = "below"
STAT_RULE_GAP = 6.0        # between a rule above a figure and the figure

# An accent below this contrast ratio against the slide background cannot
# carry a column head or a rule, so ink does instead (Q14).
ACCENT_CONTRAST = 3.0

# The neutral scaffold every style inherits. builder/styles/house_deck.json is the
# shipped house style; a skill's tokens merge over house, a template's over house.
BASE_TOKENS = {
    "name": "Base",
    "palette": {
        "bg": "FFFFFF", "ink": "1F2937", "muted": "6B7280", "accent": "3B5BDB",
        "accent2": "3B5BDB", "line": "D1D5DB", "cardFill": "FFFFFF",
        "cardBorder": "D1D5DB", "takeawayFill": "F3F4F6", "inverse": "FFFFFF",
        "series": ["3B5BDB", "1F2937", "9CA3AF", "A5B4FC", "6B7280", "C7D2FE"],
    },
    "fonts": {"display": "Arial", "body": "Arial", "mono": "Menlo", "fallback": "Arial"},
    "card": {"shape": "roundRect", "radiusAdj": 4000, "borderPt": 1.0,
             "borderColor": "cardBorder", "fill": "cardFill", "insetPt": 14, "gutterPt": 30},
    "table": {"headerFill": "F3F4F6", "ruleColor": "line", "rowHeightPt": 22},
    "chart": {"legend": "bottom", "gridlines": True},
    "diagrams": "shapes",
    "diagram": {"gutterPt": 18, "insetPt": 6, "nodeMinW": 54, "nodeMaxW": 260,
                "nodeMinH": 26, "nodeMaxH": 130, "maxNodeLines": 2, "edgePt": 1.25},
    "kicker": {"pattern": "{section} | {topic}"},
    "divider": {"numbered": True},
    "chrome": [],
    "typeScale": {},
    "furniture": {
        "kicker": {"x": 60, "y": 42, "w": 600, "h": 22},
        "headline": {"x": 60, "y": 66, "w": 840, "h": 84},
        "hairline": None,
        "body": {"x": 60, "y": 180, "w": 840, "h": 300},
        "takeaway": {"x": 60, "y": 480, "w": 840, "h": 30},
        "source": {"x": 60, "y": 514, "w": 600, "h": 14},
        "pageNumber": {"x": 866, "y": 514, "w": 64, "h": 14, "align": "right"},
        "dividerTitle": {"x": 60, "y": 180, "w": 840, "h": 150},
        "dividerSubtitle": {"x": 60, "y": 340, "w": 700, "h": 60},
        "coverWordmark": {"x": 60, "y": 60, "w": 500, "h": 40},
        "coverTitle": {"x": 60, "y": 330, "w": 720, "h": 96},
        "coverSubtitle": {"x": 60, "y": 432, "w": 720, "h": 28},
        "coverDate": {"x": 60, "y": 468, "w": 400, "h": 20},
        "agendaEyebrow": {"x": 60, "y": 60, "w": 400, "h": 20},
        "agendaStatement": {"x": 60, "y": 96, "w": 780, "h": 120},
    },
    "backgrounds": {
        "content": {"fill": "FFFFFF"},
        "title": {"fill": "1F2937", "ink": "FFFFFF", "muted": "D1D5DB", "accent": "FFFFFF",
                  "accent2": "FFFFFF"},
        "divider": {"fill": "1F2937", "ink": "FFFFFF", "muted": "D1D5DB", "accent": "FFFFFF",
                    "accent2": "FFFFFF"},
        "agenda": {"fill": "1F2937", "ink": "FFFFFF", "muted": "D1D5DB", "accent": "FFFFFF",
                   "accent2": "FFFFFF"},
    },
    "copy": {"wordmark": "", "agendaEyebrow": "Today", "emptyText": "[Add content]"},
    "rules": {"kicker": True, "pageNumber": True, "source": True,
              "pageNumberOnTitle": False, "cardsOnTitle": False},
}

# plain: Arial, charcoal ink, one accent, no furniture beyond headline and page number.
PLAIN_TOKENS = {
    "name": "Plain",
    "palette": {"accent": "3B5BDB", "accent2": "1F2937", "cardBorder": "D1D5DB",
                "takeawayFill": "F3F4F6"},
    "fonts": {"display": "Arial", "body": "Arial", "mono": "Menlo", "fallback": "Arial"},
    "card": {"shape": "rect", "radiusAdj": 0, "borderPt": 1.0, "insetPt": 12},
    "typeScale": {"cardHeader": {"color": "ink"}, "cardSub": {"color": "ink"},
                  "headline": {"size": 28, "bold": True}},
    "furniture": {"hairline": None},
    "rules": {"kicker": False, "pageNumber": True, "source": False},
}


def log(msg):
    sys.stderr.write(str(msg) + "\n")


def norm(text):
    return str(text if text is not None else "").strip().lower()


def as_int(value, default):
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def as_float(value, default):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def emu(points):
    return Emu(int(round(float(points) * EMU_PER_PT)))


def merge(base, over):
    """Deep merge `over` onto a copy of `base`. Lists and scalars replace."""
    out = copy.deepcopy(base)
    if not isinstance(over, dict):
        return out
    for key, value in over.items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = merge(out[key], value)
        else:
            out[key] = copy.deepcopy(value)
    return out


def preview(text):
    flat = re.sub(r"\s+", " ", str(text or "")).strip()
    return flat[:TEXT_PREVIEW]


def cap_first(text):
    """Capitalise the first character of a headline and change nothing else.

    A first word that carries its own capital ("iPhone", "eBay") is left alone.
    """
    out = str(text if text is not None else "")
    if not out.strip():
        return out
    stripped = out.lstrip()
    lead = len(out) - len(stripped)
    head = stripped.split()[0]
    if any(ch.isupper() for ch in head[1:]):
        return out
    return out[:lead] + stripped[0].upper() + stripped[1:]


def title_case(text):
    """Title Case one phrase: every word rises but the small joining words.

    A word that already carries its own capitals ("AI", "GTM", "McKinsey") is
    left alone, so an acronym never comes back as "Ai".
    """
    words = str(text or "").split()
    out = []
    for i, word in enumerate(words):
        if any(ch.isupper() for ch in word[1:]):
            out.append(word)
            continue
        low = re.sub(r"[^a-z]", "", word.lower())
        if low in TITLE_SMALL and 0 < i < len(words) - 1:
            out.append(word.lower())
        else:
            out.append(cap_first(word.lower()))
    return " ".join(out)


def sentence_case(text):
    """Sentence case one phrase: the first word rises, the rest fall.

    Only a plainly capitalised word falls; an acronym or a mixed-case name
    keeps its shape, because lowercasing those loses the reader a fact.
    """
    words = str(text or "").split()
    out = []
    for i, word in enumerate(words):
        if any(ch.isupper() for ch in word[1:]):
            out.append(word)
        elif i == 0:
            out.append(cap_first(word))
        else:
            out.append(word[:1].lower() + word[1:])
    return " ".join(out)


def apply_case(case, text):
    """Recase a derived kicker per `kicker.case`, field by field (F-1).

    The pattern's separators split the line, so "{section} | {topic}" cases each
    side on its own and the separator itself is untouched.
    """
    case = norm(case)
    if case in ("", "none") or not str(text or "").strip():
        return text
    if case == "caps":
        return str(text).upper()
    fn = title_case if case == "title" else sentence_case
    parts = re.split(r"([%s])" % re.escape(SEPARATORS), str(text))
    return "".join(part if i % 2 else _case_field(fn, part)
                   for i, part in enumerate(parts))


def _case_field(fn, part):
    lead = part[:len(part) - len(part.lstrip())]
    tail = part[len(part.rstrip()):]
    return lead + fn(part.strip()) + tail


def fill_pattern(pattern, section, topic):
    """`kicker.pattern` with the slide's section and name, dangling separators cut."""
    text = str(pattern or KICKER_PATTERN)
    text = text.replace("{section}", str(section or "")).replace("{topic}", str(topic or ""))
    sep = re.escape(SEPARATORS)
    text = re.sub(r"^[\s%s]+" % sep, "", text)
    text = re.sub(r"[\s%s]+$" % sep, "", text)
    text = re.sub(r"([%s])[\s%s]+([%s])" % (sep, sep, sep), r"\1", text)
    return " ".join(text.split())


# ----------------------------------------------------------------- fonts


class FontBook:
    """Which font families this Mac actually has, so a missing family can fall back."""

    DIRS = [
        os.path.expanduser("~/Library/Fonts"),
        "/Library/Fonts",
        "/System/Library/Fonts",
        "/System/Library/Fonts/Supplemental",
    ]

    def __init__(self):
        self.stems = set()
        for folder in self.DIRS:
            try:
                names = os.listdir(folder)
            except OSError:
                continue
            for name in names:
                stem = os.path.splitext(name)[0]
                self.stems.add(self.key(stem))

    # Weight and slant words a family name may carry. "Montserrat Light" is the
    # Montserrat family; the face may or may not be installed.
    STYLE_WORDS = ("thin", "hairline", "extralight", "ultralight", "light", "regular",
                   "book", "normal", "text", "medium", "demibold", "demi", "semibold",
                   "bold", "extrabold", "ultrabold", "black", "heavy", "italic",
                   "oblique")

    @staticmethod
    def key(name):
        return re.sub(r"[^a-z0-9]", "", str(name or "").lower())

    def has(self, family):
        want = self.key(family)
        if not want:
            return False
        if want in self.stems:
            return True
        return any(stem.startswith(want) for stem in self.stems)

    @classmethod
    def base_family(cls, family):
        """'Montserrat Light Italic' -> 'Montserrat'. The words a weight is written in."""
        words = str(family or "").replace("-", " ").split()
        while len(words) > 1 and cls.key(words[-1]) in cls.STYLE_WORDS:
            words.pop()
        return " ".join(words)

    def resolve(self, family):
        """(name to write, note) for one family, or (None, None) when nothing matches.

        The name PowerPoint expects is the installed face when there is one
        ("Montserrat SemiBold"), otherwise the base family ("Montserrat" for a missing
        Light face). Measurement always uses the base family.
        """
        family = str(family or "").strip()
        if not family:
            return (None, None)
        if self.has(family):
            return (family, None)
        base = self.base_family(family)
        if base and self.key(base) != self.key(family) and self.has(base):
            return (base, "font '%s' has no matching face; used '%s'" % (family, base))
        return (None, None)


FONTS = FontBook()


# ------------------------------------------------------------- derived tokens


def hex6(value, default="000000"):
    match = HEX_RE.match(str(value or ""))
    return match.group(1).upper() if match else str(default).upper()


def luminance(hex_value):
    """Relative luminance, 0 (black) to 1 (white)."""
    value = hex6(hex_value, "000000")
    red, green, blue = (int(value[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    return 0.2126 * red + 0.7152 * green + 0.0722 * blue


def is_dark(hex_value):
    return luminance(hex_value) < 0.5


def blend(front, back, amount):
    """`amount` of `front` mixed into `back`, as six hex digits."""
    a = hex6(front, "FFFFFF")
    b = hex6(back, "000000")
    out = ""
    for i in (0, 2, 4):
        one = int(a[i:i + 2], 16)
        two = int(b[i:i + 2], 16)
        out += "%02X" % int(round(two + (one - two) * max(min(float(amount), 1.0), 0.0)))
    return out


def wcag_luminance(hex_value):
    """Relative luminance the WCAG way: sRGB gamma removed first.

    `luminance` above is the cheap linear sum the light-or-dark test uses; a
    contrast ratio needs the real thing or a saturated blue reads as too dark.
    """
    value = hex6(hex_value, "000000")
    out = 0.0
    for weight, i in ((0.2126, 0), (0.7152, 2), (0.0722, 4)):
        channel = int(value[i:i + 2], 16) / 255.0
        linear = (channel / 12.92 if channel <= 0.03928
                  else ((channel + 0.055) / 1.055) ** 2.4)
        out += weight * linear
    return out


def contrast(front, back):
    """WCAG contrast ratio between two colours, 1 (none) to 21 (black on white)."""
    high = max(wcag_luminance(front), wcag_luminance(back))
    low = min(wcag_luminance(front), wcag_luminance(back))
    return (high + 0.05) / (low + 0.05)


def readability_floor(role, size):
    """The floor a role falls back to when its own style names none (Q10, Q19).

    Never above the size the style asked for: a floor is a limit, not a lift.
    A display role floors at 85% of its own size instead (Q19): a headline, a
    divider title or a figure is the shape of the slide, so it is worth
    reporting an overflow over setting it at a body size.
    """
    if role in DISPLAY_ROLES:
        return round(float(size) * DISPLAY_FLOOR, 2)
    floor = CAPTION_FLOOR if role in CAPTION_ROLES else BODY_FLOOR
    return min(floor, float(size))


def role_range(role, merged, own):
    """Settle a role's size and its range, the style's type scale winning (Q10).

    `own` is what the style's own token file said about this role, which is not
    what `merged` says: a partial token file inherits the house floors through
    the merge, and a house floor must never lift a size the style chose. When
    the style names a size, the floor is its own `min` or a readability floor,
    and the ceiling never sits below the size.
    """
    size = as_float(merged.get("size"), 12)
    own = own if isinstance(own, dict) else {}
    if "size" in own:
        size = as_float(own.get("size"), size)
        if "min" in own:
            floor = as_float(own.get("min"), size)
        else:
            floor = readability_floor(role, size)
        ceiling = (as_float(own.get("max"), size) if "max" in own
                   else max(as_float(merged.get("max"), size), size))
    else:
        floor = as_float(merged.get("min"), 9)
        ceiling = as_float(merged.get("max"), 72)
    ceiling = max(ceiling, floor)
    merged["min"] = floor
    merged["max"] = ceiling
    merged["size"] = min(max(size, floor), ceiling)
    return merged


# Every key a `backgrounds` entry may scope. A style that names the entry and
# leaves one of these out derives it from its own palette, never from house.
BACKGROUND_KEYS = ("ink", "muted", "accent", "accent2", "line", "inverse",
                   "cardFill", "cardBorder", "takeawayFill")
MUTED_BLEND = 0.65


def derive_backgrounds(tokens, given):
    """A background-scoped key the style left out comes from its own palette (Q16).

    House is a light style: its white divider accent and light grey muted have
    nothing to do with a style that names a black background. So for every
    slide kind the style itself scopes, an inherited key it never set is
    dropped, which lets `hex_of` fall through to the style's palette, and
    `muted` is derived as 65% of that kind's ink over that kind's fill.
    """
    scoped = (given.get("backgrounds") or {}) if isinstance(given, dict) else {}
    backgrounds = tokens.setdefault("backgrounds", {})
    palette = tokens.get("palette") or {}
    # A style that names its own `palette.bg` and scopes no content fill is
    # telling us what colour its slides are: the house fill is not a second
    # opinion. Everything else (dividers, the title) keeps its own scoping.
    own_bg = (given.get("palette") or {}).get("bg") if isinstance(given, dict) else None
    content = scoped.get("content") if isinstance(scoped.get("content"), dict) else {}
    if own_bg and "fill" not in content:
        backgrounds.setdefault("content", {})["fill"] = hex6(own_bg, "FFFFFF")
    for kind, own in scoped.items():
        if not isinstance(own, dict):
            continue
        entry = backgrounds.get(kind)
        if not isinstance(entry, dict):
            continue
        fill = hex6(own.get("fill") or entry.get("fill") or palette.get("bg"), "FFFFFF")
        ink = hex6(own.get("ink") or palette.get("ink"), "1F2937")
        for key in BACKGROUND_KEYS:
            if key in own or key not in entry:
                continue
            if key == "muted":
                entry[key] = blend(ink, fill, MUTED_BLEND)
            else:
                del entry[key]
    return tokens


def derive_tokens(tokens, given):
    """Fill what a partial token file leaves out from its own palette, not from house.

    A style that names a dark background must not inherit the house table header,
    takeaway bar or rule colours: white ink on a light grey fill is unreadable.
    `given` is the style's own token dict, so only the keys it never set move.
    """
    given = given if isinstance(given, dict) else {}
    palette = tokens.setdefault("palette", {})
    given_palette = given.get("palette") or {}
    bg = hex6(palette.get("bg"), "FFFFFF")
    ink = hex6(palette.get("ink"), "1F2937")

    # A palette that sets muted to the ink colour has no second voice: make one.
    if hex6(palette.get("muted"), "?") == ink:
        palette["muted"] = blend(ink, bg, 0.65)

    if is_dark(bg):
        card_fill = hex6(palette.get("cardFill"), bg)
        card_border = hex6(palette.get("cardBorder"), palette.get("muted") or ink)
        if "takeawayFill" not in given_palette:
            palette["takeawayFill"] = card_fill
        if "line" not in given_palette:
            palette["line"] = card_border
        if "table" not in given:
            # No header fill: a style that never named one gets Tufte defaults
            # (Q14), not a grey band inherited from a light house style.
            table = tokens.setdefault("table", {})
            table.pop("headerFill", None)
            table["ruleColor"] = card_border

    # A style that lays out its own furniture and never mentions the hairline
    # does not want the house one drawn through its headline.
    furniture = given.get("furniture")
    if isinstance(furniture, dict) and "hairline" not in furniture:
        tokens.setdefault("furniture", {})["hairline"] = None
    return derive_backgrounds(tokens, given)


def derive_furniture(furniture, own, scale, slide_h=SLIDE_H):
    """Divider, cover and agenda slots from the style's own measure (Q15).

    A style that lays out its own furniture has said where its slides start and
    stop, and house's 60/840 divider title is a fact about house. So every
    divider, cover and agenda slot such a style did not name is derived from
    the measure it did name: `furniture.body` when it has one, else the
    headline column. The divider title sets top-left in that band with its
    subtitle under it, where the hand-built deck puts it, and the cover group
    bottoms out. A style with no furniture of its own is house's, and nothing
    is derived.
    """
    own_f = own.get("furniture") if isinstance(own, dict) else None
    if not isinstance(own_f, dict):
        return furniture
    base = own_f.get("body")
    if not isinstance(base, dict):
        base = own_f.get("headline")
    if not isinstance(base, dict):
        return furniture
    x = as_float(base.get("x"), 60.0)
    w = max(as_float(base.get("w"), 840.0), 60.0)
    if own_f.get("body") is base:
        top = as_float(base.get("y"), 180.0)
        bottom = top + as_float(base.get("h"), 300.0)
    else:
        top = as_float(base.get("y"), 66.0) + as_float(base.get("h"), 84.0) + SLOT_GAP
        tail = own_f.get("source") if isinstance(own_f.get("source"), dict) else None
        bottom = as_float(tail.get("y"), slide_h - 40.0) if tail else slide_h - 40.0
    bottom = max(bottom, top + 60.0)

    def box(role, lines, fallback, leading=1.2):
        # A quarter of an em of slack under the last line, so a two-line title
        # keeps its descenders clear of whatever sits under it.
        spec = scale.get(role) or {}
        size = as_float(spec.get("size"), fallback)
        return max(lines * size * as_float(spec.get("leading"), leading) + size * 0.25,
                   16.0)

    head = own_f.get("headline") if isinstance(own_f.get("headline"), dict) else None
    tag = own_f.get("kicker") if isinstance(own_f.get("kicker"), dict) else None
    tag_y = as_float(tag.get("y"), 42.0) if tag else max(top - 120.0, 24.0)
    tag_h = as_float(tag.get("h"), 20.0) if tag else 20.0
    out = dict(furniture)

    title_h = box("dividerTitle", 2, 66.0, 1.1)
    sub_h = box("dividerSubtitle", 2, 21.0, 1.25)
    # A section title reads as the head of the band it opens, not as a caption
    # hung off its floor: the title takes the top of the measure and the
    # subtitle sits under it. Under the title the style named, when it named one,
    # so a half-derived pair never lands on itself.
    named = own_f.get("dividerTitle") if isinstance(own_f.get("dividerTitle"), dict) else None
    title_y = as_float(named.get("y"), top) if named else top
    title_box = as_float(named.get("h"), title_h) if named else title_h
    derived = {
        "dividerTitle": {"x": x, "y": top, "w": w, "h": title_h},
        "dividerSubtitle": {"x": x,
                            "y": min(title_y + title_box + SLOT_GAP,
                                     max(bottom - sub_h, title_y + title_box)),
                            "w": w, "h": sub_h},
        "coverWordmark": {"x": x, "y": tag_y, "w": min(w, 400.0), "h": tag_h},
        "agendaEyebrow": dict(tag) if tag else {"x": x, "y": tag_y, "w": min(w, 400.0),
                                                "h": tag_h},
    }
    date_h = box("coverDate", 1, 12.0)
    csub_h = box("coverSubtitle", 2, 16.0)
    ctitle_h = box("coverTitle", 2, 36.0, 1.15)
    derived["coverDate"] = {"x": x, "y": bottom - date_h, "w": min(w, 400.0), "h": date_h}
    derived["coverSubtitle"] = {"x": x, "y": bottom - date_h - 10.0 - csub_h,
                                "w": w, "h": csub_h}
    derived["coverTitle"] = {"x": x, "y": bottom - date_h - 10.0 - csub_h - SLOT_GAP - ctitle_h,
                             "w": w, "h": ctitle_h}
    state_h = box("agendaStatement", 2, 40.0, 1.15)
    if head:
        derived["agendaStatement"] = dict(head)
        list_top = as_float(head.get("y"), 66.0) + as_float(head.get("h"), 84.0) + SLOT_GAP
    else:
        derived["agendaStatement"] = {"x": x, "y": top, "w": w, "h": state_h}
        list_top = top + state_h + SLOT_GAP
    derived["agendaList"] = {"x": x, "y": list_top, "w": w,
                             "h": max(bottom - list_top, 60.0)}
    for name in DERIVED_SLOTS:
        if name in own_f:
            continue
        out[name] = derived[name]
    return out


# ----------------------------------------------------------------- style


class Style:
    """Palette, fonts, type scale, card style and furniture for one deck."""

    def __init__(self, tokens, mode, warnings, own=None):
        self.tokens = tokens
        # What the style itself said, which is not what the merge says: a
        # partial token file inherits the house floors, fonts and table fill,
        # and none of those may overrule the style's own choices (Q10, Q11,
        # Q14). A style with no token file of its own *is* the tokens.
        self.own = own if isinstance(own, dict) else tokens
        self.mode = mode
        self.warnings = warnings
        self.name = str(tokens.get("name") or mode.title())
        self.palette = tokens.get("palette") or {}
        self.card = tokens.get("card") or {}
        self.table = tokens.get("table") or {}
        self.chart = tokens.get("chart") or {}
        self.diagram = tokens.get("diagram") or {}
        self.copy = tokens.get("copy") or {}
        self.rules = tokens.get("rules") or {}
        self.backgrounds = tokens.get("backgrounds") or {}
        self._furniture = tokens.get("furniture") or {}
        self._scale = {}
        overrides = tokens.get("typeScale") or {}
        own_scale = self.own.get("typeScale") or {}
        for role, spec in DEFAULT_TYPE_SCALE.items():
            merged = dict(ROLE_DEFAULTS)
            merged.update(spec)
            merged.update(overrides.get(role) or {})
            own_role = own_scale.get(role)
            if (isinstance(own_role, dict) and "bold" not in own_role
                    and "weight" not in own_role):
                # Q20: weight is the style's own call. A style that names a role
                # and not its weight takes the builder's default for it, never
                # the house file's: house says its headline is bold, and that is
                # a fact about house, not about every style that merges over it.
                merged["bold"] = bool(spec.get("bold", False))
                merged.pop("weight", None)
            self._scale[role] = role_range(role, merged, own_role)
        for role, spec in overrides.items():
            if role not in self._scale:
                merged = dict(ROLE_DEFAULTS)
                merged.update(spec)
                self._scale[role] = role_range(role, merged, own_scale.get(role))
        # Q15: a style with its own furniture derives the slots it never named
        # from its own measure, so a divider never lands on house's geometry.
        self._furniture = derive_furniture(self._furniture, self.own, self._scale)
        self._font_cache = {}

    # -- colours

    def hex_of(self, value, kind="content"):
        """Resolve a palette key or a literal hex, honouring the slide kind's overrides."""
        name = str(value or "")
        over = self.backgrounds.get(kind) or {}
        if name in over and name != "fill":
            name = str(over.get(name))
        elif name in self.palette:
            name = str(self.palette.get(name))
        match = HEX_RE.match(name)
        if match:
            return match.group(1).upper()
        match = HEX_RE.match(str(self.palette.get("ink") or "1F2937"))
        return match.group(1).upper() if match else "1F2937"

    def rgb(self, value, kind="content"):
        return RGBColor.from_string(self.hex_of(value, kind))

    def background(self, kind):
        entry = self.backgrounds.get(kind) or self.backgrounds.get("content") or {}
        return self.hex_of(entry.get("fill") or "bg", "content")

    # -- fonts

    def font_chain(self, slot):
        """The families to try for a slot, the style's own fonts first (Q11).

        Every role resolves from the style's own `fonts`, so a table cell, a
        stat label or a diagram node never falls through to a house family.
        A missing family tries the style's other faces before the fallback, so
        an uninstalled body font lands on the style's display font, not Arial.
        Only `mono` may be a mono face, and only when the style names one:
        otherwise a mono role is set in the body font.
        """
        own = self.own.get("fonts") or {}
        merged = self.tokens.get("fonts") or {}
        chain = []
        if slot == "mono":
            chain.append(own.get("mono"))     # never the house mono
        elif slot:
            chain.extend([own.get(slot), merged.get(slot)])
        chain.extend([own.get("body"), own.get("display"),
                      merged.get("body"), merged.get("display"),
                      own.get("fallback"), merged.get("fallback"), "Arial"])
        out = []
        for family in chain:
            name = str(family or "").strip()
            if name and name not in out:
                out.append(name)
        return out

    def font(self, slot):
        """display | body | mono -> the family name to write into the deck.

        A weight suffix resolves to its base family first ("Montserrat Light" ->
        "Montserrat"), so an installed family is never thrown away.
        """
        slot = str(slot or "body")
        if slot in self._font_cache:
            return self._font_cache[slot]
        chain = self.font_chain(slot)
        wanted = chain[0] if chain else "Arial"
        chosen = None
        note = None
        for family in chain:
            chosen, note = FONTS.resolve(family)
            if chosen is not None:
                if family != wanted:
                    note = "font '%s' is not installed; using '%s'" % (wanted, chosen)
                break
        if chosen is None:
            chosen = "Arial"
            note = "font '%s' is not installed; using 'Arial'" % wanted
        if note and note not in self.warnings:
            self.warnings.append(note)
        self._font_cache[slot] = chosen
        return chosen

    def face(self, spec, bold=None):
        """(typeface to write, bold attribute) for a role spec (Q20).

        `typeScale.<role>.weight` ("light", "regular", "medium", "semibold",
        "bold") names a face on the role's own family: when the family has that
        face installed it is written by name, the way a designer sets it
        ("Montserrat Light", "Montserrat SemiBold"), and the bold attribute stays off.
        When it is not installed the family is written with bold on for a
        semibold or bolder weight and off for the rest, which is the nearest
        thing the renderer can synthesise.

        With no `weight` the role keeps `bold`, and the family is whatever the
        style's own `fonts` resolved to. A `bold` argument is a run overruling
        the role (an emphasised lead sentence, a cell the table sets in the
        medium weight): the base family carries the attribute, because bold on
        top of a Light face is not a bold word.
        """
        spec = spec if isinstance(spec, dict) else {}
        family = self.font(spec.get("font") or "body")
        base = FontBook.base_family(family) or family
        if bold is not None:
            return (base, bool(bold))
        want = bool(spec.get("bold"))
        weight = norm(spec.get("weight"))
        if not weight:
            return (family, want)
        if weight in WEIGHT_RIBBI:
            return (base, WEIGHT_RIBBI[weight])
        suffix = WEIGHT_FACES.get(weight)
        if suffix is None:
            note = "unknown weight '%s'; used the family" % spec.get("weight")
            if note not in self.warnings:
                self.warnings.append(note)
            return (family, want)
        named = ("%s %s" % (base, suffix)).strip()
        if FONTS.has(named):
            return (named, False)
        return (base, weight in WEIGHT_BOLD)

    def faces(self):
        """{role: typeface} for every role, for the result's `typeScale` (Q20)."""
        return dict((role, self.face(spec)[0]) for role, spec in self._scale.items())

    # -- roles and furniture

    def role(self, name):
        spec = self._scale.get(name)
        if spec is None:
            spec = dict(ROLE_DEFAULTS)
            self._scale[name] = spec
        return spec

    def slot(self, name, kind="content"):
        """A furniture rectangle, per slide kind when the style splits them."""
        per_kind = self._furniture.get(kind)
        if isinstance(per_kind, dict) and name in per_kind:
            entry = per_kind.get(name)
        else:
            entry = self._furniture.get(name)
        if not isinstance(entry, dict):
            return None
        return {
            "x": as_float(entry.get("x"), 60.0),
            "y": as_float(entry.get("y"), 60.0),
            "w": as_float(entry.get("w"), 300.0),
            "h": as_float(entry.get("h"), 30.0),
            "align": str(entry.get("align") or "left"),
        }

    def enabled(self, rule, default=True):
        value = self.rules.get(rule, default)
        return bool(value)

    def chrome_for(self, kind):
        """`chrome`: the marks this format repeats, for one slide kind (Q15).

        An entry draws on content slides unless it names its own `kinds`, so a
        format that carries its dots through the dividers says so once.
        """
        raw = self.tokens.get("chrome")
        if not isinstance(raw, list):
            return []
        out = []
        for entry in raw:
            if not isinstance(entry, dict):
                continue
            want = entry.get("kinds")
            kinds = tuple(norm(k) for k in want) if isinstance(want, list) and want \
                else CHROME_ON
            if kind in kinds:
                out.append(entry)
        return out

    def kicker_pattern(self):
        """`kicker.pattern`: "{section} | {topic}", "{section}" or "{topic}"."""
        block = self.tokens.get("kicker")
        if isinstance(block, dict) and block.get("pattern"):
            return str(block.get("pattern"))
        if isinstance(block, str) and block.strip():
            return block
        return KICKER_PATTERN

    def kicker_case(self):
        """`kicker.case`: "title", "sentence", "caps" or "none" (F-1).

        Only a derived kicker is recased. A kicker the payload wrote is the
        writer's wording and is set as given.
        """
        block = self.tokens.get("kicker")
        if isinstance(block, dict):
            name = norm(block.get("case"))
            if name in KICKER_CASES:
                return name
            if name:
                note = "unknown kicker case '%s'; left as written" % block.get("case")
                if note not in self.warnings:
                    self.warnings.append(note)
        return "none"

    def divider_numbered(self):
        """`divider.numbered`: dividers read "03 Competitive analysis"."""
        block = self.tokens.get("divider")
        if isinstance(block, dict):
            return bool(block.get("numbered", True))
        return True

    def stat_rule(self):
        """`stat.rule`: "above", "below" (the default and house) or "none" (F-5).

        The rule is the same mark either side of the figure, so one token
        settles both the hairline under an emphasised stat and the hairlines of
        a figure grid: a deck that heads each figure with a rule does it
        everywhere or nowhere.
        """
        block = self.tokens.get("stat")
        if isinstance(block, dict):
            name = norm(block.get("rule"))
            if name in STAT_RULES:
                return name
            if name:
                note = ("unknown stat rule '%s'; drew it under the figure"
                        % block.get("rule"))
                if note not in self.warnings:
                    self.warnings.append(note)
        return STAT_RULE_DEFAULT

    def diagrams_mode(self):
        """`diagrams`: "shapes" draws the diagram, "placeholder" labels the gap."""
        value = self.tokens.get("diagrams")
        if isinstance(value, dict):
            value = value.get("mode")
        mode = norm(value) or "shapes"
        return "placeholder" if mode.startswith("placeholder") else "shapes"

    def diagram_opts(self):
        """The `diagram` token block over the type scale, for diagram_layout."""
        node = self.role("diagramNode")
        caption = self.role("caption")
        opts = {
            "bodyPt": as_float(node.get("size"), 12),
            "floorPt": as_float(node.get("min"), 8),
            "k": as_float(node.get("k"), 0.5),
            "leading": as_float(node.get("leading"), 1.2),
            "maxNodeLines": as_int(node.get("maxLines"), 2),
            "axisPt": as_float(caption.get("size"), 8),
            "insetPt": min(as_float(self.card.get("insetPt"), 14), 8.0),
            "gutterPt": as_float(self.card.get("gutterPt"), 30.0) * 0.6,
        }
        block = self.tokens.get("diagram")
        if isinstance(block, dict):
            for key, value in block.items():
                opts[key] = value
        return opts

    def series_colors(self):
        raw = self.palette.get("series")
        if isinstance(raw, list) and raw:
            return [self.hex_of(c) for c in raw]
        return [self.hex_of("accent"), self.hex_of("muted"), self.hex_of("accent2")]

    # -- tables (Q14)

    def table_header_fill(self):
        """`table.headerFill`, and only when the style itself names one.

        Tufte defaults: a column head is caps in the accent over a rule, so the
        band behind it is non-data ink that a style has to ask for.
        """
        own = self.own.get("table")
        value = own.get("headerFill") if isinstance(own, dict) else None
        return str(value) if value else None

    def label_column(self):
        """`table.labelColumn`: how the first column reads when it names the rows (Q23).

        `{case: "caps"|"none", color: "accent"|"ink"|"muted", size, bold}`, over
        the default of plain ink in the family's medium weight. A format whose
        row labels are copper caps in the gutter says so here rather than in the
        builder. `size` is a point size; leaving it out keeps the cell size.
        """
        out = dict(LABEL_COLUMN_DEFAULT)
        given = self.table.get("labelColumn")
        if isinstance(given, dict):
            for key in ("case", "color", "size", "bold", "weight"):
                if key in given:
                    out[key] = given.get(key)
        case = norm(out.get("case"))
        if case and case not in LABEL_COLUMN_CASES:
            note = "unknown table label case '%s'; left as written" % out.get("case")
            if note not in self.warnings:
                self.warnings.append(note)
            case = "none"
        out["case"] = case or "none"
        return out

    def label_column_spec(self, cell_spec, kind="content"):
        """(cell spec, size) for the first column of a label-column table (Q23)."""
        token = self.label_column()
        spec = dict(cell_spec)
        spec["caps"] = token["case"] == "caps"
        spec["bold"] = bool(token.get("bold"))
        if token.get("weight"):
            spec["weight"] = token.get("weight")
        colour = str(token.get("color") or "ink")
        spec["color"] = self.accent_key(kind) if colour == "accent" else colour
        size = as_float(token.get("size"), as_float(cell_spec.get("size"), 10))
        spec["size"] = size
        return (spec, size)

    def accent_key(self, kind="content"):
        """`accent`, or `ink` when the accent has too little contrast to read."""
        if contrast(self.hex_of("accent", kind), self.background(kind)) >= ACCENT_CONTRAST:
            return "accent"
        return "ink"


def load_tokens_file(path, warnings):
    try:
        with open(path, "r") as handle:
            return json.load(handle)
    except Exception as exc:  # noqa: BLE001 - any read problem falls back
        warnings.append("could not read style tokens %s: %s" % (os.path.basename(path), exc))
        return {}


def house_tokens(warnings):
    path = os.path.join(STYLES_DIR, "house_deck.json")
    if not os.path.exists(path):
        warnings.append("builder/styles/house_deck.json is missing; used the base tokens")
        return copy.deepcopy(BASE_TOKENS)
    return merge(BASE_TOKENS, load_tokens_file(path, warnings))


def make_style(spec, warnings):
    """Build the Style for the spec's style block. Returns (style, template_path)."""
    raw = spec.get("style") or {}
    mode = norm(raw.get("mode")) or "house"
    if mode not in ("plain", "house", "skill", "template"):
        warnings.append("unknown style mode '%s'; used house" % mode)
        mode = "house"

    if mode == "plain":
        return Style(merge(BASE_TOKENS, PLAIN_TOKENS), "plain", warnings), None

    tokens = house_tokens(warnings)
    if mode == "skill":
        extra = {}
        path = raw.get("tokensPath")
        if path:
            extra = load_tokens_file(path, warnings)
        inline = raw.get("tokens")
        if isinstance(inline, dict):
            extra = merge(extra, inline)
        if not extra:
            warnings.append("skill style has no tokens; built as house")
            return Style(tokens, "house", warnings), None
        return Style(derive_tokens(merge(tokens, extra), extra), "skill", warnings,
                     own=extra), None

    if mode == "house":
        inline = raw.get("tokens")
        if isinstance(inline, dict):
            return Style(derive_tokens(merge(tokens, inline), inline), "house", warnings,
                         own=inline), None
        return Style(tokens, "house", warnings), None

    template_path = raw.get("templatePath")
    if not template_path or not os.path.exists(str(template_path)):
        warnings.append("template not found; built as house")
        return Style(tokens, "house", warnings), None
    inline = raw.get("tokens")
    own = None
    if isinstance(inline, dict):
        tokens = derive_tokens(merge(tokens, inline), inline)
        own = inline
    return Style(tokens, "template", warnings, own=own), str(template_path)


# ------------------------------------------------------------- text fitting


def cpl(width, size, k):
    """Characters per line at this width and size (k is the average glyph width in ems)."""
    return max(int(math.floor(width / max(size * k, 0.1))), 1)


def wrap_lines(text, chars):
    """Greedy word wrap, one entry per rendered line. Explicit breaks split paragraphs."""
    lines = []
    for para in str(text if text is not None else "").split("\n"):
        words = para.split()
        if not words:
            lines.append("")
            continue
        current = words[0]
        for word in words[1:]:
            if len(current) + 1 + len(word) <= chars:
                current = current + " " + word
            else:
                lines.append(current)
                current = word
        lines.append(current)
    return lines or [""]


def step_down(size, floor):
    step = 2.0 if size > 14 else 1.0
    return max(float(floor), float(size) - step)


class Para:
    """One paragraph queued for a text frame, with the role that sizes it."""

    def __init__(self, text, role_name, spec, level=0, space_after=0.0, align=None,
                 color=None, bullet=False, runs=None):
        self.text = "" if text is None else str(text)
        self.role_name = role_name
        self.spec = spec
        self.level = level
        self.space_after = space_after
        self.align = align or spec.get("align") or "left"
        self.color = color or spec.get("color")
        self.bullet = bullet
        self.runs = runs
        self.size = as_float(spec.get("size"), 12)
        self.lines = [self.text]
        self.widow = False

    @property
    def indent(self):
        """Left margin of the text: a bullet hangs one step to the left of its copy."""
        return 18.0 * (self.level + (1 if self.bullet else 0))

    def measure(self, width):
        chars = cpl(max(width - self.indent, 24.0), self.size, as_float(self.spec.get("k"), 0.5))
        self.lines = wrap_lines(self.text, chars)
        return self.lines

    def height(self):
        lead = as_float(self.spec.get("leading"), 1.25)
        return len(self.lines) * self.size * lead + self.space_after


def fit_paragraphs(paras, width, height):
    """Step sizes down within each role's range until the stack fits. Returns (overflow, widow).

    Paragraphs sharing a role step together so a list keeps one size. At the floor
    the fit stops and the caller reports overflow (deck-planning section 7).
    """
    live = [p for p in paras if p.text.strip() or p.runs]
    if not live:
        return False, False
    guard = 0
    while guard < 60:
        guard += 1
        total = 0.0
        over_lines = False
        for para in live:
            para.measure(width)
            total += para.height()
            if len(para.lines) > as_int(para.spec.get("maxLines"), 99):
                over_lines = True
        if total <= height and not over_lines:
            break
        movable = [p for p in live if p.size > as_float(p.spec.get("min"), 9)]
        if not movable:
            return True, _widow_flag(live, width)
        target = max(movable, key=lambda p: p.size - as_float(p.spec.get("min"), 9))
        new_size = step_down(target.size, as_float(target.spec.get("min"), 9))
        for para in live:
            if para.role_name == target.role_name:
                para.size = new_size
    else:
        return True, _widow_flag(live, width)

    total = sum(p.height() for p in live)
    return total > height, _widow_flag(live, width)


def _widow_flag(paras, width):
    """Pull a one-word last line up a step when that helps; otherwise flag it."""
    widow = False
    for para in paras:
        if len(para.lines) < 2 or not para.lines[-1].strip():
            continue
        if len(para.lines[-1].split()) != 1:
            continue
        floor = as_float(para.spec.get("min"), 9)
        smaller = step_down(para.size, floor)
        if smaller < para.size:
            before = len(para.lines)
            keep = para.size
            para.size = smaller
            trial = para.measure(width)
            if len(trial) < before or (trial and len(trial[-1].split()) > 1):
                continue
            para.size = keep
            para.measure(width)
        widow = True
    return widow


# ------------------------------------------------------------------ canvas


class Ctx:
    """Per-build state: style, geometry, warnings and the snapshot being assembled."""

    def __init__(self, spec, style, warnings):
        size = spec.get("slideSize") or {}
        self.slide_w = as_float(size.get("w"), SLIDE_W)
        self.slide_h = as_float(size.get("h"), SLIDE_H)
        grid = spec.get("grid") or {}
        self.cols = max(as_int(grid.get("cols"), GRID_COLS), 1)
        self.rows = max(as_int(grid.get("rows"), GRID_ROWS), 1)
        self.cell = as_float(grid.get("cell"), CELL)
        self.cell_w = self.cell
        self.cell_h = self.cell
        if abs(self.cols * self.cell - self.slide_w) > 0.5:
            self.cell_w = self.slide_w / self.cols
        if abs(self.rows * self.cell - self.slide_h) > 0.5:
            self.cell_h = self.slide_h / self.rows
        self.style = style
        self.warnings = warnings
        self.shapes = []
        self.kind = "content"
        self.where = ""
        self.headline = ""
        self.ordinal = None
        self.seq = 0
        self.number = 0
        self.fidelity = fit.fidelity_of(spec, self.warn)
        self.notes_from_story = spec.get("notesFromStory", True) is not False
        self.adjusted = []
        self.notes_extra = []     # what the slide cut and the notes keep (Q15)
        self.type_scale = {}
        self.type_scale_floored = []
        # The style's column module, from `furniture.body` (Q9). One module for
        # the deck, so every slide's blocks land on the same few left edges.
        self.module = fit.module_of(style.slot("body", "content"), self.cell_w)

    def warn(self, message):
        self.warnings.append(("%s: %s" % (self.where, message)) if self.where else str(message))

    def rect(self, block):
        """Grid cells to points, clamped to the slide with a warning."""
        col = as_int(block.get("col"), 1)
        row = as_int(block.get("row"), 1)
        cols = max(as_int(block.get("cols"), 1), 1)
        rows = max(as_int(block.get("rows"), 1), 1)
        x = (col - 1) * self.cell_w
        y = (row - 1) * self.cell_h
        w = cols * self.cell_w
        h = rows * self.cell_h
        clamped = False
        if x < 0:
            w, x, clamped = w + x, 0.0, True
        if y < 0:
            h, y, clamped = h + y, 0.0, True
        if x > self.slide_w - MIN_SIDE:
            x, clamped = self.slide_w - MIN_SIDE, True
        if y > self.slide_h - MIN_SIDE:
            y, clamped = self.slide_h - MIN_SIDE, True
        if x + w > self.slide_w:
            w, clamped = self.slide_w - x, True
        if y + h > self.slide_h:
            h, clamped = self.slide_h - y, True
        w = max(w, MIN_SIDE)
        h = max(h, MIN_SIDE)
        if clamped:
            self.warn("block '%s' runs off the slide; clamped to %d x %d"
                      % (block.get("id") or block.get("type") or "?",
                         int(self.slide_w), int(self.slide_h)))
        return (x, y, w, h)

    def record(self, kind, role, rect, text="", font_pt=None, block_id=None):
        self.seq += 1
        entry = {
            "id": "s%d-%d" % (len(self.shapes) + 1, self.seq),
            "blockId": block_id,
            "role": role,
            "x": round(rect[0], 2), "y": round(rect[1], 2),
            "w": round(rect[2], 2), "h": round(rect[3], 2),
            "kind": kind,
            "text": preview(text),
        }
        if font_pt is not None:
            entry["fontPt"] = round(float(font_pt), 1)
        self.shapes.append(entry)
        return entry


def set_background(slide, hex_color):
    try:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(hex_color)
    except Exception as exc:  # noqa: BLE001 - a master background still renders
        log("could not set the slide background: %s" % exc)


def style_shape(shape, fill_hex, line_hex, line_pt, dashed=False):
    # Drop the theme style reference: PowerPoint honours our empty effectLst, LibreOffice
    # keeps drawing the referenced shadow. Without it the shape is only what spPr says.
    element = shape._element
    for node in element.findall(qn("p:style")):
        element.remove(node)
    if fill_hex:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(fill_hex)
    else:
        shape.fill.background()
    if line_hex and line_pt:
        shape.line.color.rgb = RGBColor.from_string(line_hex)
        shape.line.width = Pt(float(line_pt))
        if dashed:
            try:
                shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH
            except Exception:  # noqa: BLE001 - solid is an acceptable fallback
                pass
    else:
        shape.line.fill.background()
    try:
        shape.shadow.inherit = False
    except Exception:  # noqa: BLE001 - shadow control is cosmetic
        pass


def add_box(ctx, slide, rect, shape_name="rect", fill=None, border=None, border_pt=0.0,
            dashed=False, radius_adj=0):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if shape_name == "roundRect" else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, emu(rect[0]), emu(rect[1]), emu(rect[2]), emu(rect[3]))
    if shape_name == "roundRect":
        try:
            shape.adjustments[0] = max(float(radius_adj), 0.0) / 100000.0
        except Exception:  # noqa: BLE001 - the default fillet is fine
            pass
    style_shape(shape, fill, border, border_pt, dashed)
    shape.text_frame.word_wrap = True
    return shape


ALIGNMENTS = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT,
              "justify": PP_ALIGN.JUSTIFY}
ANCHORS = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}


def set_bullet(paragraph, hex_color, font_name):
    pPr = paragraph._p.get_or_add_pPr()
    for tag in ("a:buClr", "a:buFont", "a:buChar", "a:buNone"):
        for node in pPr.findall(qn(tag)):
            pPr.remove(node)
    clr = pPr.makeelement(qn("a:buClr"), {})
    srgb = pPr.makeelement(qn("a:srgbClr"), {"val": hex_color})
    clr.append(srgb)
    pPr.append(clr)
    pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": font_name}))
    pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•"}))


def write_paragraphs(ctx, frame, paras, kind, inset=0.0, anchor="top", inset_y=None,
                     wrap=True):
    """Write fitted paragraphs into a text frame."""
    style = ctx.style
    if inset_y is None:
        inset_y = inset
    frame.word_wrap = bool(wrap)
    try:
        frame.auto_size = MSO_AUTO_SIZE.NONE
    except Exception:  # noqa: BLE001 - not every frame exposes auto_size
        pass
    frame.margin_left = emu(inset)
    frame.margin_right = emu(inset)
    frame.margin_top = emu(inset_y)
    frame.margin_bottom = emu(inset_y)
    frame.vertical_anchor = ANCHORS.get(anchor, MSO_ANCHOR.TOP)
    live = [p for p in paras if p.text.strip() or p.runs]
    if not live:
        live = paras[:1]
    first = True
    for para in live:
        p = frame.paragraphs[0] if first else frame.add_paragraph()
        first = False
        p.alignment = ALIGNMENTS.get(str(para.align), PP_ALIGN.LEFT)
        if para.bullet:
            set_bullet(p, style.hex_of(para.color, kind), style.font(para.spec.get("font")))
        if para.level or para.bullet:
            if para.level:
                p.level = min(int(para.level), 8)
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(int(para.indent * EMU_PER_PT)))
            pPr.set("indent", str(int((-18.0 if para.bullet else 0.0) * EMU_PER_PT)))
        try:
            p.line_spacing = as_float(para.spec.get("leading"), 1.25)
            p.space_after = Pt(para.space_after)
        except Exception:  # noqa: BLE001 - spacing is cosmetic
            pass
        runs = para.runs or [{"text": para.text}]
        for spec in runs:
            run = p.add_run()
            text = str(spec.get("text", ""))
            if para.spec.get("caps"):
                text = text.upper()
            run.text = text
            font = run.font
            font.size = Pt(para.size)
            # The role's weight names the face; a run may still ask for bold,
            # and then the base family carries it (Q20).
            name, weighted = style.face(para.spec, bold=spec.get("bold"))
            font.name = name
            font.bold = weighted
            font.italic = bool(spec.get("italic", para.spec.get("italic")))
            font.color.rgb = style.rgb(spec.get("color") or para.color, kind)


def add_text(ctx, slide, rect, paras, role, block_id=None, inset=0.0, anchor="top",
             shape_kind="text", where=None, shrink=False):
    """Fit, draw and record one text box. Returns the fitted point size.

    `shrink` trims a top-anchored box to the height of the text it holds. Every
    furniture slot uses it, so an oversized headline slot never reserves a band
    the body then collides with.
    """
    overflow, widow = fit_paragraphs(paras, rect[2] - 2 * inset, rect[3] - 2 * inset)
    if shrink and anchor == "top":
        live_now = [p for p in paras if p.text.strip() or p.runs]
        used = sum(p.height() for p in live_now) if live_now else 0.0
        if used > 0:
            rect = (rect[0], rect[1], rect[2],
                    max(min(rect[3], used + 2 * inset + 2.0), MIN_SIDE))
    box = slide.shapes.add_textbox(emu(rect[0]), emu(rect[1]), emu(rect[2]), emu(rect[3]))
    write_paragraphs(ctx, box.text_frame, paras, ctx.kind, inset=inset, anchor=anchor)
    live = [p for p in paras if p.text.strip() or p.runs] or paras[:1]
    size = live[0].size if live else None
    text = " ".join(p.text for p in live).strip()
    ctx.record(shape_kind, role, rect, text, size, block_id)
    label = where or ("block '%s'" % block_id if block_id else role)
    if overflow:
        lines = sum(len(p.lines) for p in live)
        ctx.warn("%s overflows at the %s floor (%d lines)" % (label, role, lines))
    if widow:
        ctx.warn("%s ends on a one-word line" % label)
    return size


def para_for(ctx, text, role_name, **kwargs):
    return Para(text, role_name, ctx.style.role(role_name), **kwargs)


# -------------------------------------------------------------- emphasis (Q13)

EMPH_BAR_PT = 4.0          # the accent bar down an emphasised text block
EMPH_BAR_GAP = 10.0        # and the air between the bar and the copy
EMPH_BORDER_PT = 0.75      # added to an emphasised card's border
EMPH_RULE_PT = 1.5         # the accent rule under an emphasised table header
HAIRLINE_PT = 0.75
SENTENCE_END = re.compile(r"(.+?[.!?:])(\s+)(.*)$", re.S)


def emphasised(block):
    """`emphasis: true`: the plan says this is the block that carries the slide."""
    return bool(block.get("emphasis")) if isinstance(block, dict) else False


def text_rect(block, rect):
    """Where an emphasised text or bullets block's copy sits.

    The accent bar takes the left edge, so the copy starts inside it. The
    measure and the deck-wide type pass use this too: the box the reader sees
    is the box the sizes were chosen for.
    """
    if not emphasised(block) or norm(block.get("type")) not in ("text", "bullets"):
        return rect
    gap = EMPH_BAR_PT + EMPH_BAR_GAP
    return (rect[0] + gap, rect[1], max(rect[2] - gap, MIN_SIDE), rect[3])


def draw_rule(ctx, slide, rect, color="accent", role="hairline", block_id=None,
              width_pt=HAIRLINE_PT):
    """One filled rectangle: a hairline, a rule or a bar. No border, no text."""
    where = (rect[0], rect[1], max(rect[2], 0.5), max(rect[3], width_pt))
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, emu(where[0]), emu(where[1]),
                                   emu(where[2]), emu(where[3]))
    style_shape(shape, ctx.style.hex_of(color, ctx.kind), None, 0)
    ctx.record("rule", role, where, "", None, block_id)
    return shape


def bold_lead(paras, roles=("body", "bullets")):
    """Bold the first sentence of the first paragraph that carries copy.

    The text is left alone, so the measurement is the same: only the runs change.
    """
    for para in paras:
        if para.role_name not in roles or not para.text.strip():
            continue
        if para.runs:
            runs = [dict(run) for run in para.runs]
            runs[0]["bold"] = True
            para.runs = runs
            return para
        match = SENTENCE_END.match(para.text)
        if match:
            para.runs = [{"text": match.group(1), "bold": True},
                         {"text": match.group(2) + match.group(3)}]
        else:
            para.runs = [{"text": para.text, "bold": True}]
        return para
    return None


# ------------------------------------------------------------- block writers


def content_of(block):
    raw = block.get("content")
    if isinstance(raw, dict):
        return raw
    if isinstance(raw, str):
        return {"text": raw}
    if isinstance(raw, list):
        return {"items": raw}
    return {}


def empty_para(ctx, block, role_name):
    label = block.get("title") or block.get("role") or block.get("type") or "content"
    spec = dict(ctx.style.role(role_name))
    spec["italic"] = True
    spec["color"] = "muted"
    return Para("[%s]" % str(label), role_name, spec)


def title_paras(ctx, block, paras):
    """Prepend a block title when the plan gave one."""
    if block.get("title") and block.get("type") not in ("cards", "stat"):
        head = para_for(ctx, block.get("title"), "blockTitle", space_after=4.0)
        return [head] + paras
    return paras


def text_paras(ctx, block, warn=True):
    """The paragraphs of a `text` block, shared by the writer and the measure."""
    body = content_of(block)
    role = "body"
    paras = []
    raw = body.get("paragraphs")
    if raw is None and body.get("runs"):
        raw = [{"runs": body.get("runs")}]
    if raw is None and body.get("text"):
        raw = [body.get("text")]
    for item in raw or []:
        if isinstance(item, dict):
            runs = item.get("runs")
            text = item.get("text") or " ".join(str(r.get("text", "")) for r in (runs or []))
            paras.append(para_for(ctx, text, role, space_after=6.0, runs=runs))
        else:
            paras.append(para_for(ctx, item, role, space_after=6.0))
    if not paras:
        paras = [empty_para(ctx, block, role)]
        if warn:
            ctx.warn("block '%s' has no text" % block.get("id"))
    return title_paras(ctx, block, paras)


def emph_bar(ctx, slide, block, rect):
    """The accent bar down the left edge of an emphasised text or bullets block."""
    draw_rule(ctx, slide, (rect[0], rect[1], EMPH_BAR_PT, rect[3]),
              ctx.style.accent_key(ctx.kind), block_id=block.get("id"),
              width_pt=EMPH_BAR_PT)


def write_text_block(ctx, slide, block, rect):
    paras = text_paras(ctx, block)
    if emphasised(block):
        bold_lead(paras)
        emph_bar(ctx, slide, block, rect)
    add_text(ctx, slide, text_rect(block, rect), paras, "body", block.get("id"))


def flatten_bullets(ctx, items, level, out):
    for item in items or []:
        if isinstance(item, dict):
            text = item.get("text", "")
            out.append(para_for(ctx, text, "bullets", level=level, space_after=4.0, bullet=True))
            flatten_bullets(ctx, item.get("children"), level + 1, out)
        else:
            out.append(para_for(ctx, item, "bullets", level=level, space_after=4.0, bullet=True))


def bullet_paras(ctx, block, warn=True):
    body = content_of(block)
    paras = []
    flatten_bullets(ctx, body.get("items"), 0, paras)
    if not paras:
        paras = [empty_para(ctx, block, "bullets")]
        if warn:
            ctx.warn("block '%s' has no items" % block.get("id"))
    return title_paras(ctx, block, paras)


def write_bullets_block(ctx, slide, block, rect):
    paras = bullet_paras(ctx, block)
    if emphasised(block):
        bold_lead(paras)
        emph_bar(ctx, slide, block, rect)
    add_text(ctx, slide, text_rect(block, rect), paras, "bullets", block.get("id"))


def tile_rects(rect, count, gutter, wrap_at=CARD_WRAP_AT):
    """Lay `count` tiles across the rect, wrapping to a second row past wrap_at."""
    count = max(int(count), 1)
    cols = count if count <= wrap_at - 1 else int(math.ceil(count / 2.0))
    rows = int(math.ceil(count / float(cols)))
    x, y, w, h = rect
    tile_w = (w - gutter * (cols - 1)) / float(cols)
    tile_h = (h - gutter * (rows - 1)) / float(rows)
    out = []
    for i in range(count):
        r, c = divmod(i, cols)
        in_row = min(cols, count - r * cols)
        offset = (w - (in_row * tile_w + (in_row - 1) * gutter)) / 2.0 if in_row < cols else 0.0
        out.append((x + offset + c * (tile_w + gutter), y + r * (tile_h + gutter), tile_w, tile_h))
    return out


def card_items(block):
    items = content_of(block).get("items") or []
    return [i if isinstance(i, dict) else {"body": i} for i in items]


def card_paras(ctx, block, item, warn=True):
    paras = []
    if item.get("header"):
        paras.append(para_for(ctx, item.get("header"), "cardHeader", space_after=2.0))
    if item.get("sub"):
        paras.append(para_for(ctx, item.get("sub"), "cardSub", space_after=6.0))
    text = item.get("body")
    if isinstance(text, list):
        for line in text:
            paras.append(para_for(ctx, line, "cardBody", space_after=4.0))
    elif text:
        paras.append(para_for(ctx, text, "cardBody", space_after=4.0))
    if not paras:
        paras = [empty_para(ctx, block, "cardBody")]
        if warn:
            ctx.warn("block '%s' has an empty card" % block.get("id"))
    return paras


def stat_paras(ctx, block, item, warn=True):
    paras = [para_for(ctx, item.get("value"), "statValue", space_after=4.0)]
    if item.get("label"):
        paras.append(para_for(ctx, item.get("label"), "statLabel"))
    if not str(item.get("value") or "").strip():
        paras = [empty_para(ctx, block, "statValue")]
        if warn:
            ctx.warn("block '%s' has a stat with no value" % block.get("id"))
    return paras


def write_cards_block(ctx, slide, block, rect):
    style = ctx.style
    items = card_items(block)
    if not items:
        items = [{"header": "", "body": ""}]
        ctx.warn("block '%s' has no cards" % block.get("id"))
    gutter = as_float(style.card.get("gutterPt"), 30.0)
    inset = as_float(style.card.get("insetPt"), 14.0)
    emph = emphasised(block)
    border = style.card.get("borderColor") or "cardBorder"
    border_pt = as_float(style.card.get("borderPt"), 1.0)
    if emph:
        border = style.accent_key(ctx.kind)
        border_pt += EMPH_BORDER_PT
    for rect_i, item in zip(tile_rects(rect, len(items), gutter), items):
        shape = add_box(
            ctx, slide, rect_i,
            shape_name=str(style.card.get("shape") or "roundRect"),
            fill=style.hex_of(style.card.get("fill") or "cardFill", ctx.kind),
            border=style.hex_of(border, ctx.kind),
            border_pt=border_pt,
            radius_adj=as_float(style.card.get("radiusAdj"), 4000),
        )
        paras = card_paras(ctx, block, item)
        if emph and paras and paras[0].role_name == "cardHeader":
            paras[0].color = style.accent_key(ctx.kind)
        label = "block '%s' card '%s'" % (block.get("id"), preview(item.get("header") or ""))
        overflow, widow = fit_paragraphs(paras, rect_i[2] - 2 * inset, rect_i[3] - 2 * inset)
        write_paragraphs(ctx, shape.text_frame, paras, ctx.kind, inset=inset)
        head = paras[0]
        ctx.record("card", "card", rect_i,
                   " ".join(p.text for p in paras), head.size, block.get("id"))
        if overflow:
            ctx.warn("%s overflows at the card floor" % label)
        if widow:
            ctx.warn("%s ends on a one-word line" % label)


def stat_items(block):
    items = content_of(block).get("items") or []
    return [i if isinstance(i, dict) else {"value": i, "label": ""} for i in items]


def stat_layout(ctx, block, rect):
    """"band" or "grid" for a stat block (Q22).

    `layout` on the block settles it. Otherwise four or more figures in a block
    more than six rows deep are supporting figures, not the slide's headline
    numbers, and they read as a two-column figure list.
    """
    named = norm(block.get("layout") or content_of(block).get("layout"))
    if named in STAT_LAYOUTS:
        return named
    if named:
        ctx.warn("block '%s' has unknown stat layout '%s'; used the band"
                 % (block.get("id"), block.get("layout")
                    or content_of(block).get("layout")))
        return "band"
    tall = rect[3] > STAT_GRID_ROWS * ctx.cell_h + 0.5
    return "grid" if (len(stat_items(block)) >= STAT_GRID_ITEMS and tall) else "band"


def stat_grid_roles(ctx):
    """(number spec, label spec) for a figure grid: the accent at 0.6 x statValue.

    The label spec is the live `body` role, not a copy, so a grid label is set
    at the deck-wide body size like every other sentence on the slide (Q4).
    """
    value = dict(ctx.style.role("statValue"))
    size = round(as_float(value.get("size"), 44) * STAT_GRID_VALUE, 1)
    value.update({"size": size, "min": size, "max": size, "maxLines": 1,
                  "color": ctx.style.accent_key(ctx.kind)})
    return (value, ctx.style.role("body"))


def stat_grid_number_w(ctx, block, pair_w):
    """The gutter the numbers sit in: the widest of them, so every label lines up."""
    spec = stat_grid_roles(ctx)[0]
    size = as_float(spec.get("size"), 26)
    k = as_float(spec.get("k"), 0.6)
    widest = 0.0
    for item in stat_items(block):
        widest = max(widest, len(str(item.get("value") or "")) * size * k)
    return min(max(widest + STAT_GRID_PAD, pair_w * 0.18), pair_w * 0.45)


def stat_grid_rects(rect, count, cols=STAT_GRID_COLS):
    """One rect per number-and-label pair, in reading order, two to a row."""
    count = max(int(count), 1)
    cols = max(min(int(cols), count), 1)
    rows = int(math.ceil(count / float(cols)))
    x, y, w, h = rect
    pair_w = (w - STAT_GRID_GUTTER * (cols - 1)) / float(cols)
    pair_h = h / float(rows)
    out = []
    for i in range(count):
        r, c = divmod(i, cols)
        out.append((x + c * (pair_w + STAT_GRID_GUTTER), y + r * pair_h, pair_w, pair_h))
    return out


def stat_grid_label(ctx, block, item, warn=True):
    """The label half of one grid pair: body copy beside the figure."""
    text = str(item.get("label") or "").strip()
    if not text:
        if warn:
            ctx.warn("block '%s' has a figure with no label" % block.get("id"))
        return [empty_para(ctx, block, "body")]
    return [Para(text, "body", stat_grid_roles(ctx)[1])]


def write_stat_grid(ctx, slide, block, rect, items):
    """Number and label pairs in two columns, a hairline per pair (Q22, F-5).

    `stat.rule` says which side of the pair the hairline sits on, and the pair
    moves out of its way: a rule above heads the figure, a rule below closes it.
    """
    number_spec, _ = stat_grid_roles(ctx)
    rects = stat_grid_rects(rect, len(items))
    pair_w = rects[0][2] if rects else rect[2]
    number_w = stat_grid_number_w(ctx, block, pair_w)
    line = ctx.style.hex_of("line", ctx.kind)
    rule = ctx.style.stat_rule()
    for pair, item in zip(rects, items):
        x, y, w, h = pair
        body_y = y + STAT_GRID_PAD if rule == "above" else y
        body_h = max(h - STAT_GRID_PAD, MIN_SIDE)
        number = Para(str(item.get("value") or ""), "statValue", number_spec)
        add_text(ctx, slide, (x, body_y, number_w, body_h), [number], "statValue",
                 block.get("id"), anchor="middle", shape_kind="stat",
                 where="block '%s' figure" % block.get("id"))
        add_text(ctx, slide, (x + number_w, body_y, max(w - number_w, MIN_SIDE), body_h),
                 stat_grid_label(ctx, block, item), "statLabel", block.get("id"),
                 anchor="middle", shape_kind="stat",
                 where="block '%s' figure label" % block.get("id"))
        if rule == "none":
            continue
        rule_y = y if rule == "above" else y + h - HAIRLINE_PT
        draw_rule(ctx, slide, (x, rule_y, w, HAIRLINE_PT), line,
                  role="statRule", block_id=block.get("id"))


def write_stat_block(ctx, slide, block, rect):
    items = stat_items(block)
    if not items:
        items = [{"value": "", "label": ""}]
        ctx.warn("block '%s' has no stats" % block.get("id"))
    first = len(ctx.shapes)
    layout = stat_layout(ctx, block, rect)
    rule = ctx.style.stat_rule()
    if layout == "grid":
        write_stat_grid(ctx, slide, block, rect, items)
    else:
        gutter = as_float(ctx.style.card.get("gutterPt"), 30.0)
        emph = emphasised(block)
        for rect_i, item in zip(tile_rects(rect, len(items), gutter), items):
            paras = stat_paras(ctx, block, item)
            drawn = rect_i
            if emph and rule == "above":
                # The rule heads the figure, so the figure starts under it.
                drawn = (rect_i[0], rect_i[1] + STAT_RULE_GAP, rect_i[2],
                         max(rect_i[3] - STAT_RULE_GAP, MIN_SIDE))
            add_text(ctx, slide, drawn, paras, "stat", block.get("id"),
                     shape_kind="stat", where="block '%s' stat" % block.get("id"))
            if not emph or not paras or rule == "none":
                continue
            # The value already carries the accent: the hairline beside it is
            # what tells the reader which figure the slide is about.
            if rule == "above":
                at = rect_i[1]
            else:
                value = paras[0]
                drop = (len(value.lines) * value.size
                        * as_float(value.spec.get("leading"), 1.25))
                at = rect_i[1] + drop + 1.25
            draw_rule(ctx, slide, (rect_i[0], at, rect_i[2], HAIRLINE_PT),
                      ctx.style.accent_key(ctx.kind), block_id=block.get("id"))
    for entry in ctx.shapes[first:]:
        entry["layout"] = layout
        if entry.get("kind") == "rule":
            # Which side the rule went, so the canvas and the check read the
            # same slide the builder drew.
            entry["rulePosition"] = rule


def write_quote_block(ctx, slide, block, rect):
    body = content_of(block)
    text = body.get("text") or ""
    paras = []
    if text:
        paras.append(para_for(ctx, "“%s”" % str(text).strip('"“”'),
                              "quote", space_after=10.0))
    else:
        paras.append(empty_para(ctx, block, "quote"))
        ctx.warn("block '%s' has no quote" % block.get("id"))
    if body.get("attribution"):
        paras.append(para_for(ctx, "— %s" % body.get("attribution"), "attribution"))
    add_text(ctx, slide, rect, paras, "quote", block.get("id"), shape_kind="quote")


def write_takeaway_block(ctx, slide, block, rect):
    style = ctx.style
    emph = emphasised(block)
    fill = style.palette.get("takeawayFill") or "takeawayFill"
    if emph:
        fill = style.accent_key(ctx.kind)
    shape = add_box(ctx, slide, rect, shape_name="roundRect",
                    fill=style.hex_of(fill, ctx.kind),
                    border=None, border_pt=0,
                    radius_adj=as_float(style.card.get("radiusAdj"), 4000))
    body = content_of(block)
    text = body.get("text") or block.get("title") or ""
    if not str(text).strip():
        para = empty_para(ctx, block, "takeaway")
        ctx.warn("block '%s' has no takeaway text" % block.get("id"))
    else:
        para = para_for(ctx, text, "takeaway")
    if emph:
        # Inverse: the accent carries the bar, so the ink is the slide's own
        # background colour, whatever kind of slide this is.
        para.color = style.background(ctx.kind)
    inset, inset_y = 14.0, 4.0
    overflow, _ = fit_paragraphs([para], rect[2] - 2 * inset, rect[3] - 2 * inset_y)
    write_paragraphs(ctx, shape.text_frame, [para], ctx.kind, inset=inset, anchor="middle",
                     inset_y=inset_y)
    ctx.record("takeaway", "takeaway", rect, para.text, para.size, block.get("id"))
    if overflow:
        ctx.warn("block '%s' overflows the takeaway bar" % block.get("id"))


def write_placeholder_block(ctx, slide, block, rect, label=None, warn=True):
    style = ctx.style
    text = label or content_of(block).get("label") or block.get("title") or "Visual to come"
    shape = add_box(ctx, slide, rect, shape_name="roundRect", fill=None,
                    border=style.hex_of(style.card.get("borderColor") or "cardBorder", ctx.kind),
                    border_pt=as_float(style.card.get("borderPt"), 1.0), dashed=True,
                    radius_adj=as_float(style.card.get("radiusAdj"), 4000))
    para = para_for(ctx, "[%s]" % str(text), "placeholder", align="center")
    overflow, _ = fit_paragraphs([para], rect[2] - 24, rect[3] - 12)
    write_paragraphs(ctx, shape.text_frame, [para], ctx.kind, inset=12.0, anchor="middle",
                     inset_y=6.0)
    if overflow:
        ctx.warn("block '%s' label does not fit its placeholder" % block.get("id"))
    ctx.record("placeholder", block.get("role") or "placeholder", rect, para.text, para.size,
               block.get("id"))
    if warn and not str(text).strip():
        ctx.warn("block '%s' is an unlabelled placeholder" % block.get("id"))


def write_image_block(ctx, slide, block, rect):
    body = content_of(block)
    path = body.get("path")
    if not path or not os.path.exists(str(path)):
        if path:
            ctx.warn("block '%s' image not found: %s" % (block.get("id"), path))
        write_placeholder_block(ctx, slide, block, rect,
                                label=body.get("placeholder") or body.get("label") or "Image",
                                warn=False)
        return
    try:
        pic = slide.shapes.add_picture(str(path), emu(rect[0]), emu(rect[1]))
    except Exception as exc:  # noqa: BLE001 - an unreadable image becomes a placeholder
        ctx.warn("block '%s' image could not be placed: %s" % (block.get("id"), exc))
        write_placeholder_block(ctx, slide, block, rect, label="Image", warn=False)
        return
    native_w = pic.width / float(EMU_PER_PT)
    native_h = pic.height / float(EMU_PER_PT)
    scale = min(rect[2] / native_w, rect[3] / native_h) if native_w and native_h else 1.0
    draw_w, draw_h = native_w * scale, native_h * scale
    pic.left = emu(rect[0] + (rect[2] - draw_w) / 2.0)
    pic.top = emu(rect[1] + (rect[3] - draw_h) / 2.0)
    pic.width = emu(draw_w)
    pic.height = emu(draw_h)
    ctx.record("image", block.get("role") or "visual",
               (rect[0] + (rect[2] - draw_w) / 2.0, rect[1] + (rect[3] - draw_h) / 2.0,
                draw_w, draw_h), os.path.basename(str(path)), None, block.get("id"))


def emphasis_series(block, names):
    """Which series an emphasised chart puts in the accent: the named one, else the first."""
    want = content_of(block).get("emphasisSeries")
    if isinstance(want, bool):
        want = None
    if isinstance(want, int):
        return want if 0 <= want < len(names) else 0
    if want is not None:
        for i, name in enumerate(names):
            if norm(name) == norm(want):
                return i
    return 0


def parse_value(value):
    """(number, unit) from a chart value: "7.7%" -> (7.7, "%"), "$60K" -> (60000, "$").

    The unit is what the axis is counting, and the values are where it comes
    from: a spec that writes its figures the way the deck says them keeps the
    format instead of throwing it away at the chart's edge (Q15).
    """
    if value is None:
        return (None, "")
    if isinstance(value, bool):
        return (None, "")
    if isinstance(value, (int, float)):
        return (float(value), "")
    text = str(value).strip()
    unit = ""
    for mark, name in NUMBER_UNITS:
        if mark in text:
            unit = name
            text = text.replace(mark, "")
    text = text.strip()
    scale = 1.0
    if len(text) > 1 and text[-1].lower() in SCALE_SUFFIX and text[-2].isdigit():
        scale = SCALE_SUFFIX[text[-1].lower()]
        text = text[:-1]
    match = UNIT_SUFFIX.match(text.strip())
    if not match:
        return (as_float(value, None), unit)
    if match.group(2) and match.group(2).lower() == "x":
        unit = "x"
    number = as_float(match.group(1).replace(",", ""), None)
    return ((number * scale) if number is not None else None, unit)


def axis_format(unit, values):
    """The axis number format the values imply, or None for a plain count."""
    live = [v for v in values if v is not None]
    core = "#,##0"
    if any(abs(v - round(v)) > 1e-6 for v in live):
        core = "#,##0.0"
    if unit in ("%", "x"):
        return '%s"%s"' % (core, unit)
    if unit:
        return '"%s"%s' % (unit, core)
    return None


def looks_like_time(categories):
    """True when the categories are a sequence the reader expects in order."""
    if not categories:
        return False
    hits = 0
    for one in categories:
        text = norm(one)
        if TIME_NUMBER.match(text.strip()):
            hits += 1
            continue
        if any(word in TIME_WORDS for word in WORD_RE.findall(text)):
            hits += 1
    return hits * 2 >= len(categories)


def chart_series(body, categories):
    """[(name, [numbers])] and the unit the values carry."""
    out = []
    unit = ""
    for entry in body.get("series") or []:
        if isinstance(entry, dict):
            name = str(entry.get("name") or "Series")
            values = entry.get("values") or []
        else:
            name, values = "Series", entry
        numbers = []
        for value in values:
            number, found = parse_value(value)
            numbers.append(number)
            if found and not unit:
                unit = found
        while len(numbers) < len(categories):
            numbers.append(None)
        out.append((name, numbers[:len(categories)]))
    return (out, unit)


def sort_categories(kind, categories, series):
    """Rank a single-series bar chart, largest first as the reader sees it.

    PowerPoint plots the first category at the bottom of a bar chart and at the
    left of a column chart, so the data order is reversed for one and not the
    other. A time series keeps its own order: a year out of place is a lie.
    """
    if kind not in CHART_SORTED_KINDS or len(series) != 1:
        return (categories, series, False)
    if looks_like_time(categories):
        return (categories, series, False)
    values = series[0][1]
    order = sorted(range(len(categories)),
                   key=lambda i: (values[i] is None, -(values[i] or 0.0)))
    if kind == "bar":
        order = list(reversed(order))
    if order == list(range(len(categories))):
        return (categories, series, False)
    return ([categories[i] for i in order],
            [(series[0][0], [values[i] for i in order])], True)


def chart_title_of(ctx, block, body):
    """The chart's own title, dropped when it repeats the title above it (Q15)."""
    title = str(body.get("title") or "").strip()
    if not title:
        return ""
    above = str(block.get("title") or "").strip() or ctx.headline
    if above and repeats(title, above):
        ctx.warn("block '%s' chart title repeats the title above it; dropped"
                 % block.get("id"))
        return ""
    return title


def write_chart_block(ctx, slide, block, rect):
    style = ctx.style
    body = content_of(block)
    kind = norm(body.get("kind")) or "column"
    categories = [str(c) for c in (body.get("categories") or [])]
    if not categories or not (body.get("series") or []):
        ctx.warn("block '%s' has no chart data" % block.get("id"))
        write_placeholder_block(ctx, slide, block, rect, label="Chart", warn=False)
        return
    chart_type = CHART_KINDS.get(kind)
    if chart_type is None:
        ctx.warn("block '%s' has unknown chart kind '%s'; drew a column chart"
                 % (block.get("id"), kind))
        chart_type = XL_CHART_TYPE.COLUMN_CLUSTERED
        kind = "column"
    series, unit = chart_series(body, categories)
    categories, series, ranked = sort_categories(kind, categories, series)
    data = CategoryChartData()
    data.categories = categories
    names = []
    flat = []
    for name, numbers in series:
        names.append(name)
        flat.extend(numbers)
        data.add_series(name, numbers)
    frame = slide.shapes.add_chart(chart_type, emu(rect[0]), emu(rect[1]),
                                   emu(rect[2]), emu(rect[3]), data)
    chart = frame.chart
    colors = style.series_colors()
    want = content_of(block).get("emphasisSeries")
    picked = want is not None and not isinstance(want, bool)
    lead = emphasis_series(block, names) if (emphasised(block) or picked) else None
    if lead is not None:
        # One accent against muted others: the emphasised series is the point,
        # the rest are the context it stands out from (Q13).
        accent = style.hex_of(style.accent_key(ctx.kind), ctx.kind)
        muted = style.hex_of("muted", ctx.kind)
        colors = [accent if i == lead else muted for i in range(max(len(names), 1))]
    muted = style.hex_of("muted", ctx.kind)
    line_hex = style.hex_of("line", ctx.kind)
    caption = as_float(style.role("caption").get("size"), 8)
    number_format = axis_format(unit, flat)
    title = chart_title_of(ctx, block, body)
    try:
        chart.font.size = Pt(caption)
        chart.font.name = style.font("body")
        chart.font.color.rgb = RGBColor.from_string(muted)
        chart.has_title = bool(title)
        if title:
            chart.chart_title.text_frame.text = title
        # A legend tells two series apart; with one series it names the only
        # thing on the chart, so it goes (Knaflic, Zelazny).
        chart.has_legend = len(series) > 1
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
    except Exception as exc:  # noqa: BLE001 - chrome only
        log("chart chrome: %s" % exc)
    try:
        plot = chart.plots[0]
        # A chart the reader can count gets its numbers on the marks instead of
        # on an axis they have to walk back to.
        plot.has_data_labels = len(categories) * max(len(series), 1) <= CHART_LABEL_POINTS
        if plot.has_data_labels:
            labels = plot.data_labels
            labels.font.size = Pt(caption)
            labels.font.color.rgb = RGBColor.from_string(muted)
            if number_format:
                labels.number_format = number_format
                labels.number_format_is_linked = False
            if kind == "pie":
                labels.show_category_name = True
        if kind == "pie":
            for i, point in enumerate(plot.series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = RGBColor.from_string(colors[i % len(colors)])
        else:
            for i, one in enumerate(plot.series):
                color = RGBColor.from_string(colors[i % len(colors)])
                if kind == "line":
                    one.format.line.color.rgb = color
                    one.format.line.width = Pt(2.25)
                    one.smooth = False
                else:
                    one.format.fill.solid()
                    one.format.fill.fore_color.rgb = color
                    one.format.line.fill.background()
    except Exception as exc:  # noqa: BLE001 - labels and colour are cosmetic
        log("chart series: %s" % exc)
    try:
        value_axis = chart.value_axis
        # Gridlines run horizontally or not at all: on a bar chart the value
        # axis is the horizontal one, so its gridlines would be vertical rules.
        horizontal = kind in ("column", "line", "stacked")
        value_axis.has_major_gridlines = bool(style.chart.get("gridlines", True)) and horizontal
        if value_axis.has_major_gridlines:
            grid = value_axis.major_gridlines.format.line
            grid.color.rgb = RGBColor.from_string(line_hex)
            grid.width = Pt(0.75)
        if number_format:
            value_axis.tick_labels.number_format = number_format
            value_axis.tick_labels.number_format_is_linked = False
        for axis in (value_axis, chart.category_axis):
            axis.format.line.color.rgb = RGBColor.from_string(line_hex)
            axis.tick_labels.font.color.rgb = RGBColor.from_string(muted)
            axis.tick_labels.font.size = Pt(caption)
            try:
                axis.has_major_gridlines = axis.has_major_gridlines
                axis.major_tick_mark = XL_TICK_MARK.NONE
                axis.minor_tick_mark = XL_TICK_MARK.NONE
            except Exception:  # noqa: BLE001 - tick marks are cosmetic
                pass
        chart.category_axis.has_major_gridlines = False
    except Exception as exc:  # noqa: BLE001 - axes are cosmetic, pie has none
        log("chart axes: %s" % exc)
    entry = ctx.record("chart", block.get("role") or "visual", rect,
                       title or block.get("title") or kind, None, block.get("id"))
    if ranked:
        entry["sorted"] = "descending"
    if number_format:
        entry["numberFormat"] = number_format


def set_table_style(table, style_id):
    tbl = table._tbl
    props = tbl.find(qn("a:tblPr"))
    if props is None:
        return
    for node in props.findall(qn("a:tableStyleId")):
        props.remove(node)
    node = props.makeelement(qn("a:tableStyleId"), {})
    node.text = style_id
    props.append(node)


def cell_rule(cell, hex_color, width_pt):
    """A bottom hairline under a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for node in tcPr.findall(qn("a:lnB")):
        tcPr.remove(node)
    ln = tcPr.makeelement(qn("a:lnB"), {"w": str(int(width_pt * EMU_PER_PT)), "cap": "flat",
                                        "cmpd": "sng", "algn": "ctr"})
    fill = ln.makeelement(qn("a:solidFill"), {})
    fill.append(ln.makeelement(qn("a:srgbClr"), {"val": hex_color}))
    ln.append(fill)
    tcPr.append(ln)


# ------------------------------------------------------------------ tables (Q14)

TABLE_PAD_X = 6.0             # the cell margin, left and right
TABLE_PAD_Y = 3.0             # and top and bottom
TABLE_HEAD_RULE_PT = 1.0      # the accent rule under the column heads
TABLE_ROW_RULE_PT = 0.5       # the hairlines between the rows
RATING_STEPS = 5              # "3" reads as 3 of 5
LABEL_COL_WORDS = 5           # a label column's cells are no longer than this
# A cell with a bare empty run has no size for the renderer to set its line box
# from, so it falls back to its own default and the row grows past the height we
# set: a rating mark drawn over that row would then miss it. One sized
# non-breaking space keeps the row the height the table was measured at.
EMPTY_CELL = "\u00a0"

# A cell that reads as data: a number, a percent, a currency or an "x" multiple.
NUMERIC_CELL = re.compile(
    r"^[(+\-\u2212]?\s*[$\u00a3\u20ac\u00a5]?\s*\d[\d,\u202f ]*(?:\.\d+)?\s*"
    r"(?:%|x|\u00d7|k|m|bn)?\s*\)?$", re.I)
RATING_FRACTION = re.compile(r"^([0-5])\s*/\s*([1-9])$")
RATING_GLYPHS = {"\u25cf": 1.0, "\u2b24": 1.0, "\u25d0": 0.5, "\u25d1": 0.5,
                 "\u25d2": 0.5, "\u25d3": 0.5, "\u25cd": 0.5, "\u25cb": 0.0,
                 "\u25ef": 0.0, "\u25cc": 0.0}


def is_numeric_cell(value):
    """Does this cell hold a figure, so that it belongs on the right edge."""
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    text = str(value if value is not None else "").strip()
    return bool(text) and bool(NUMERIC_CELL.match(text))


def rating_of(value):
    """(filled, steps) for a rating cell: "3", "3/5" or a run of rating glyphs.

    None when the cell is not a rating at all.
    """
    text = str(value if value is not None else "").strip()
    if not text:
        return None
    match = RATING_FRACTION.match(text)
    if match:
        return (float(match.group(1)), int(match.group(2)))
    if len(text) == 1 and text.isdigit() and 1 <= int(text) <= RATING_STEPS:
        return (float(text), RATING_STEPS)
    glyphs = [c for c in text if not c.isspace()]
    if glyphs and all(c in RATING_GLYPHS for c in glyphs):
        return (sum(RATING_GLYPHS[c] for c in glyphs), len(glyphs))
    return None


def rating_marks(filled, steps):
    """One mark per step: 1 full, 0.5 half, 0 empty."""
    out = []
    left = float(filled)
    for _ in range(max(int(steps), 1)):
        out.append(1.0 if left >= 0.99 else (0.5 if left >= 0.49 else 0.0))
        left -= 1.0
    return out


def numeric_columns(headers, rows):
    """The columns that hold figures, so the whole column goes to the right edge.

    Alignment is a column decision: one figure in a column of phrases pulled to
    the right edge on its own reads as a mistake, so a column earns the right
    edge only when every cell in it is a figure.
    """
    n_cols = max(len(headers), max((len(r) for r in rows), default=0))
    out = set()
    for c in range(n_cols):
        values = [row[c] for row in rows if c < len(row)
                  and str(row[c] if row[c] is not None else "").strip()]
        if values and all(is_numeric_cell(v) for v in values):
            out.add(c)
    return out


def rating_columns(headers, rows):
    """Which columns read as ratings rather than as numbers.

    A whole column has to agree, because a lone digit is only a rating in a
    column of ratings: a "2" in a column of hours is two hours, not two of five.
    One real rating glyph or a fraction settles a column on its own; bare digits
    need at least two rows saying the same thing.
    """
    n_cols = max(len(headers), max((len(r) for r in rows), default=0))
    out = set()
    for c in range(n_cols):
        values = [row[c] for row in rows if c < len(row)
                  and str(row[c] if row[c] is not None else "").strip()]
        marks = [rating_of(v) for v in values]
        if not marks or any(mark is None for mark in marks):
            continue
        glyphs = any(not str(v).strip().isdigit() for v in values)
        if glyphs or len(marks) >= 2:
            out.add(c)
    return out


def label_column(headers, rows):
    """Does the first column read as a label column rather than as data.

    Short text on the left with figures or ratings to the right of it: the
    column is naming the rows, so it is set in the medium weight the family has.
    """
    first = [row[0] for row in rows if row]
    if len(first) < 2:
        return False
    for value in first:
        text = str(value if value is not None else "").strip()
        if not text or is_numeric_cell(value) or rating_of(value) is not None:
            return False
        if len(text.split()) > LABEL_COL_WORDS:
            return False
    if headers and str(headers[0] or "").strip() == "":
        return True
    for row in rows:
        for value in row[1:]:
            if is_numeric_cell(value) or rating_of(value) is not None:
                return True
    return False


def table_type(ctx, sizes=None):
    """(head spec, cell spec) for a table, on Tufte defaults.

    The column head is the deck's `cardSub` size in caps in the accent, never
    larger than the cell size: a column name is a label, not a second headline.
    An accent with too little contrast against the slide gives way to ink.
    """
    sizes = sizes or {}
    cell = dict(ctx.style.role("tableCell"))
    cell_size = as_float(sizes.get("tableCell"), as_float(cell.get("size"), 10))
    cell["size"] = cell_size
    head = dict(ctx.style.role("tableHeader"))
    sub = as_float(sizes.get("cardSub"),
                   as_float(ctx.style.role("cardSub").get("size"), 12))
    size = min(sub, cell_size)
    head.update({"size": size, "min": size, "max": size, "caps": True,
                 "color": ctx.style.accent_key(ctx.kind)})
    return (head, cell)


def table_rows_pt(ctx, block, rect, sizes=None):
    """(row heights, column width, total height) at the natural height (Q14).

    A row is as tall as its own tallest cell needs at the deck-wide table size
    plus the cell padding, so a table is the height of what is in it and the
    rows never stretch to fill the block.
    """
    headers, rows, n_cols, n_rows = table_grid(ctx, block)
    if not n_cols or not n_rows:
        return ([], 0.0, 0.0)
    head, cell = table_type(ctx, sizes)
    col_w = rect[2] / float(n_cols)
    inner = max(col_w - 2 * TABLE_PAD_X, 12.0)
    # A label column may carry its own size (Q23), so the first cell of each
    # data row is measured at the size it will really be set in.
    label = ctx.style.label_column_spec(cell, ctx.kind)[0] if label_column(headers,
                                                                          rows) else None
    heights = []
    for spec, values in [(head, headers)] + [(cell, row) for row in rows]:
        lead = as_float(spec.get("leading"), 1.25)
        lines = 1
        for i, value in enumerate(values):
            text = "" if value is None else str(value)
            if not text.strip() or rating_of(value) is not None:
                continue
            at = label if (label is not None and i == 0 and spec is not head) else spec
            size = as_float(at.get("size"), 10)
            k = as_float(at.get("k"), 0.5)
            lines = max(lines, len(wrap_lines(text, cpl(inner, size, k))))
        size = as_float(spec.get("size"), 10)
        if label is not None and spec is not head:
            size = max(size, as_float(label.get("size"), size))
        heights.append(lines * size * lead + 2 * TABLE_PAD_Y)
    return (heights, col_w, sum(heights))


def draw_rating(ctx, slide, rect, filled, steps, block_id=None):
    """A rating cell as filled and empty circles in the accent, sized to the row."""
    style = ctx.style
    marks = rating_marks(filled, steps)
    accent = style.hex_of(style.accent_key(ctx.kind), ctx.kind)
    back = style.background(ctx.kind)
    room = max(rect[2] - 2 * TABLE_PAD_X, 8.0)
    size = min(max(rect[3] - 2 * TABLE_PAD_Y - 1.0, 5.0), 12.0)
    gap = size * 0.45
    while len(marks) * size + (len(marks) - 1) * gap > room and size > 4.0:
        size -= 0.5
        gap = size * 0.45
    x = rect[0] + TABLE_PAD_X
    y = rect[1] + (rect[3] - size) / 2.0
    for mark in marks:
        shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, emu(x), emu(y), emu(size), emu(size))
        if mark >= 0.99:
            style_shape(shape, accent, accent, 0.75)
        elif mark >= 0.49:
            style_shape(shape, blend(accent, back, 0.5), accent, 0.75)
        else:
            style_shape(shape, None, accent, 0.75)
        ctx.record("rule", "tableRating", (x, y, size, size), "", None, block_id)
        x += size + gap


def write_table_block(ctx, slide, block, rect):
    """A native table on Tufte defaults: no fill, accent caps heads, hairlines."""
    style = ctx.style
    body = content_of(block)
    headers = [str(h) for h in (body.get("headers") or [])]
    rows = body.get("rows") or []
    if not headers and rows:
        headers = ["" for _ in rows[0]]
    if not headers:
        ctx.warn("block '%s' has no table data" % block.get("id"))
        write_placeholder_block(ctx, slide, block, rect, label="Table", warn=False)
        return
    n_rows = len(rows) + 1
    n_cols = max(len(headers), max((len(r) for r in rows), default=0))
    heights, col_w, total = table_rows_pt(ctx, block, rect)
    if total > rect[3] + 1.0:
        ctx.warn("block '%s' table needs %d pt and the block holds %d"
                 % (block.get("id"), round(total), round(rect[3])))
    frame = slide.shapes.add_table(n_rows, n_cols, emu(rect[0]), emu(rect[1]),
                                   emu(rect[2]), emu(total or rect[3]))
    table = frame.table
    set_table_style(table, NO_STYLE_NO_GRID)
    header_fill = style.table_header_fill()
    table.first_row = bool(header_fill)
    table.horz_banding = False
    for col in table.columns:
        col.width = emu(col_w)
    for i, row in enumerate(table.rows):
        row.height = emu(heights[i] if i < len(heights) else rect[3] / float(n_rows))
    fill_hex = style.hex_of(header_fill, ctx.kind) if header_fill else None
    accent = style.accent_key(ctx.kind)
    rule = style.hex_of(style.table.get("ruleColor") or "line", ctx.kind)
    head_spec, cell_spec = table_type(ctx)
    head_size = as_float(head_spec.get("size"), 10)
    cell_size = as_float(cell_spec.get("size"), 10)
    head_rule = EMPH_RULE_PT if emphasised(block) else TABLE_HEAD_RULE_PT
    # The table first, then the marks drawn inside it: a reader of the snapshot
    # meets the container before its contents.
    ctx.record("table", block.get("role") or "table",
               (rect[0], rect[1], rect[2], total or rect[3]),
               " | ".join(headers), cell_size, block.get("id"))
    for c in range(n_cols):
        cell = table.cell(0, c)
        paint_cell(cell, fill_hex)
        write_cell(ctx, cell, headers[c] if c < len(headers) else "", head_spec, head_size)
        cell_rule(cell, style.hex_of(accent, ctx.kind), head_rule)
    labels = label_column(headers, rows)
    label_spec, label_size = (style.label_column_spec(cell_spec, ctx.kind)
                              if labels else (cell_spec, cell_size))
    ratings = rating_columns(headers, rows)
    figures = numeric_columns(headers, rows)
    for r, row in enumerate(rows):
        for c in range(n_cols):
            cell = table.cell(r + 1, c)
            paint_cell(cell, None)
            value = row[c] if c < len(row) else ""
            rating = rating_of(value) if c in ratings else None
            spec = dict(label_spec if (labels and c == 0) else cell_spec)
            size = label_size if (labels and c == 0) else cell_size
            if rating is not None:
                write_cell(ctx, cell, "", spec, size)
                where = (rect[0] + c * col_w, rect[1] + sum(heights[:r + 1]),
                         col_w, heights[r + 1] if r + 1 < len(heights) else cell_size)
                draw_rating(ctx, slide, where, rating[0], rating[1], block.get("id"))
            else:
                write_cell(ctx, cell, value, spec, size, right=c in figures)
            # Hairlines go between the rows, so the last row closes with air.
            if r < len(rows) - 1:
                cell_rule(cell, rule, TABLE_ROW_RULE_PT)
    if n_rows > 10:
        ctx.warn("block '%s' has %d table rows; more than 9 belongs in an appendix"
                 % (block.get("id"), len(rows)))


def paint_cell(cell, fill_hex):
    """A named header fill, or no fill at all so the slide shows through."""
    if fill_hex:
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(fill_hex)
    else:
        cell.fill.background()


def write_cell(ctx, cell, value, spec, size, right=False):
    cell.margin_left = emu(TABLE_PAD_X)
    cell.margin_right = emu(TABLE_PAD_X)
    cell.margin_top = emu(TABLE_PAD_Y)
    cell.margin_bottom = emu(TABLE_PAD_Y)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame = cell.text_frame
    frame.word_wrap = True
    para = frame.paragraphs[0]
    para.alignment = PP_ALIGN.RIGHT if right else PP_ALIGN.LEFT
    run = para.add_run()
    text = "" if value is None else str(value)
    if text and spec.get("caps"):
        text = text.upper()
    run.text = text if text.strip() else EMPTY_CELL
    run.font.size = Pt(size)
    # The role's own font slot and weight: a table cell is the style's body
    # face, never a house family it never asked for (Q11, Q20).
    name, bold = ctx.style.face(spec)
    run.font.name = name
    run.font.bold = bold
    run.font.italic = bool(spec.get("italic"))
    run.font.color.rgb = ctx.style.rgb(spec.get("color"), ctx.kind)


def set_arrow(shape, hex_color, width_pt, dashed=False):
    """Colour a connector and put a filled triangle on its tail.

    python-pptx has no arrowhead API, so the a:tailEnd element goes on the
    line's XML directly; it is the last child a:ln takes.
    """
    element = shape._element
    for node in element.findall(qn("p:style")):
        element.remove(node)
    line = shape.line
    line.color.rgb = RGBColor.from_string(hex_color)
    line.width = Pt(float(width_pt))
    if dashed:
        try:
            line.dash_style = MSO_LINE_DASH_STYLE.DASH
        except Exception:  # noqa: BLE001 - solid still reads
            pass
    ln = line._get_or_add_ln()
    for tag in ("a:headEnd", "a:tailEnd"):
        for node in ln.findall(qn(tag)):
            ln.remove(node)
    ln.append(ln.makeelement(qn("a:tailEnd"),
                             {"type": "triangle", "w": "med", "len": "med"}))
    try:
        shape.shadow.inherit = False
    except Exception:  # noqa: BLE001 - shadow control is cosmetic
        pass


def draw_connector(ctx, slide, points):
    """One shape for the whole polyline: a connector for two points, else a freeform."""
    if len(points) <= 2:
        (x1, y1), (x2, y2) = points[0], points[-1]
        return slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                          emu(x1), emu(y1), emu(x2), emu(y2))
    path = slide.shapes.build_freeform(points[0][0], points[0][1], EMU_PER_PT)
    path.add_line_segments(points[1:], close=False)
    shape = path.convert_to_shape()
    try:
        shape.fill.background()
    except Exception:  # noqa: BLE001 - an unfilled path is the default
        pass
    return shape


def pinned(spec, size, **over):
    """A copy of a type-scale entry pinned to one size, so the fit cannot move it."""
    out = dict(spec)
    out["size"] = float(size)
    out["min"] = float(size)
    out["max"] = float(size)
    out.update(over)
    return out


def bounds_of(points):
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return (min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))


def write_diagram_block(ctx, slide, block, rect):
    """A native flow, cycle, timeline, tree or matrix from the style tokens.

    `diagrams: "placeholder"` draws the labelled dashed box the Harbor format
    mandates instead of shape art. Geometry comes from diagram_layout, so the
    canvas sketch and the deck agree.
    """
    style = ctx.style
    body = content_of(block)
    label = diagram_layout.summary(body)
    if style.diagrams_mode() == "placeholder":
        for note in diagram_layout.normalise(body)[4]:
            ctx.warn("block '%s' %s" % (block.get("id"), note))
        write_placeholder_block(ctx, slide, block, rect, label=label, warn=False)
        return
    plan = diagram_layout.plan(body, rect, style.diagram_opts())
    for note in plan["warnings"]:
        ctx.warn("block '%s' %s" % (block.get("id"), note))
    if not plan["nodes"]:
        write_placeholder_block(ctx, slide, block, rect, label=label, warn=False)
        return

    edge_hex = style.hex_of(style.diagram.get("edgeColor") or "muted", ctx.kind)
    edge_pt = as_float(style.diagram.get("edgePt"), 1.25)
    for axis in plan["axis"]:
        where = (axis["x"], axis["y"], axis["w"], max(as_float(axis.get("h"), 0.75), 0.75))
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, emu(where[0]), emu(where[1]),
                                      emu(where[2]), emu(where[3]))
        style_shape(line, style.hex_of("line", ctx.kind), None, 0)
        ctx.record("rule", "diagramAxis", where, "", None, block.get("id"))

    caption = style.role("caption")
    for edge in plan["edges"]:
        points = [(float(x), float(y)) for x, y in edge["points"]]
        shape = draw_connector(ctx, slide, points)
        set_arrow(shape, edge_hex, edge_pt)
        entry = ctx.record("edge", "diagramEdge", bounds_of(points),
                           edge.get("label") or "", None, block.get("id"))
        # the canvas draws the same polyline from the snapshot
        entry["points"] = [[round(px, 2), round(py, 2)] for px, py in points]
        text = str(edge.get("label") or "").strip()
        if not text:
            continue
        size = as_float(caption.get("size"), 8)
        width = max(len(text) * size * as_float(caption.get("k"), 0.5) * 1.2 + 10.0, 28.0)
        height = size * as_float(caption.get("leading"), 1.25) + 4.0
        spot = (edge["labelXY"][0] - width / 2.0, edge["labelXY"][1] - height / 2.0,
                width, height)
        para = Para(text, "caption", pinned(caption, size, align="center"))
        box = slide.shapes.add_textbox(emu(spot[0]), emu(spot[1]), emu(spot[2]), emu(spot[3]))
        # an edge label is one short line: it never wraps, it just sits on the midpoint
        write_paragraphs(ctx, box.text_frame, [para], ctx.kind, inset=2.0, anchor="middle",
                         inset_y=1.0, wrap=False)
        ctx.record("text", "diagramEdge", spot, text, size, block.get("id"))

    node_spec = style.role("diagramNode")
    label_spec = style.role("diagramLabel")
    inset = as_float(style.diagram_opts().get("insetPt"), 6.0)
    for node in plan["nodes"]:
        where = (node["x"], node["y"], node["w"], node["h"])
        size = as_float(node.get("fontPt"), as_float(node_spec.get("size"), 12))
        if node.get("shape") == "text":
            para = Para(node["label"], "diagramLabel",
                        pinned(label_spec, size, align=node.get("align") or "left"))
            box = slide.shapes.add_textbox(emu(where[0]), emu(where[1]),
                                           emu(where[2]), emu(where[3]))
            write_paragraphs(ctx, box.text_frame, [para], ctx.kind, inset=2.0,
                             anchor="middle", inset_y=1.0)
            ctx.record("node", "diagramLabel", where, node["label"], size, block.get("id"))
            continue
        shape = add_box(
            ctx, slide, where,
            shape_name=str(style.card.get("shape") or "roundRect"),
            fill=style.hex_of(style.card.get("fill") or "cardFill", ctx.kind),
            border=style.hex_of(style.card.get("borderColor") or "cardBorder", ctx.kind),
            border_pt=as_float(style.card.get("borderPt"), 1.0),
            radius_adj=as_float(style.card.get("radiusAdj"), 4000),
        )
        para = Para(node["label"], "diagramNode", pinned(node_spec, size, align="center"))
        write_paragraphs(ctx, shape.text_frame, [para], ctx.kind, inset=inset,
                         anchor="middle", inset_y=2.0)
        ctx.record("node", "diagramNode", where, node["label"], size, block.get("id"))


def write_notes_block(ctx, slide, block, rect):
    """Notes never draw; they land in the speaker notes and stay in the snapshot."""
    body = content_of(block)
    text = body.get("text") or ""
    if text:
        add_notes(slide, text)
    else:
        ctx.warn("block '%s' has no notes text" % block.get("id"))
    ctx.record("notes", "notes", rect, text, None, block.get("id"))


# ------------------------------------------------- measuring a block (Q3, Q4)


def stack_height(paras, width, sizes=None):
    """(height, lines, widest line, widow) for a paragraph stack at fixed sizes.

    `sizes` maps a role name to the size to measure it at; a role it does not
    name keeps the size the paragraph already carries. Nothing steps down here:
    the fit ladder is add_text's job, this is the question "does it fit".
    """
    total = 0.0
    lines = 0
    widest = 0.0
    widow = False
    for para in paras:
        if not (para.text.strip() or para.runs):
            continue
        if sizes and para.role_name in sizes:
            para.size = float(sizes[para.role_name])
        para.measure(width)
        total += para.height()
        lines += len(para.lines)
        if len(para.lines) > 1 and len(para.lines[-1].split()) == 1:
            widow = True
        k = as_float(para.spec.get("k"), 0.5)
        for line in para.lines:
            widest = max(widest, len(line) * para.size * k + para.indent)
    return (total, lines, widest, widow)


def role_sizes(style, names, key="min"):
    return dict((name, as_float(style.role(name).get(key), 12)) for name in names)


def over_lines(paras, style):
    """True when a stack needs more lines than its roles allow."""
    for para in paras:
        if not (para.text.strip() or para.runs):
            continue
        if len(para.lines) > as_int(para.spec.get("maxLines"), 99):
            return True
    return False


def block_paras(ctx, block, kind=None, rect=None):
    """Every paragraph stack a block draws: one per tile for cards and stats.

    `rect` is the rectangle the block will be drawn in, which a stat block needs
    to know whether it is a band or a figure grid (Q22).
    """
    kind = kind or norm(block.get("type"))
    if kind == "text":
        return [text_paras(ctx, block, warn=False)]
    if kind == "bullets":
        return [bullet_paras(ctx, block, warn=False)]
    if kind == "cards":
        return [card_paras(ctx, block, item, warn=False)
                for item in card_items(block)] or [[]]
    if kind == "stat":
        # A figure grid sets the number beside its label, so the label is the
        # only half that can wrap and the only half worth measuring (Q22).
        if stat_layout(ctx, block, rect or (0.0, 0.0, 0.0, 0.0)) == "grid":
            return [stat_grid_label(ctx, block, item, warn=False)
                    for item in stat_items(block)] or [[]]
        return [stat_paras(ctx, block, item, warn=False)
                for item in stat_items(block)] or [[]]
    if kind == "quote":
        body = content_of(block)
        paras = [para_for(ctx, "“%s”" % str(body.get("text") or "").strip('"“”'),
                          "quote", space_after=10.0)]
        if body.get("attribution"):
            paras.append(para_for(ctx, "— %s" % body.get("attribution"), "attribution"))
        return [paras]
    if kind == "takeaway":
        text = content_of(block).get("text") or block.get("title") or ""
        return [[para_for(ctx, text, "takeaway")]]
    return []


TILED = ("cards", "stat")
MEASURED = ("text", "bullets", "cards", "stat", "quote", "takeaway", "table")


def tile_shape(count):
    """(cols, rows) of tiles for a cards or stat block, the writer's own rule."""
    count = max(int(count), 1)
    cols = count if count <= CARD_WRAP_AT - 1 else int(math.ceil(count / 2.0))
    return (cols, int(math.ceil(count / float(cols))))


def block_insets(ctx, kind):
    """(x, y) inset the writer uses, so the measure sees the same box."""
    if kind == "cards":
        return (as_float(ctx.style.card.get("insetPt"), 14.0),) * 2
    if kind == "takeaway":
        return (14.0, 4.0)
    return (0.0, 0.0)


def table_need(ctx, block, rect):
    """(need rows, need cols, repick) for a native table in this rectangle."""
    headers, rows, n_cols, n_rows = table_grid(ctx, block)
    if not n_cols or not n_rows:
        return (0, 0, None)
    floors = {"tableCell": as_float(ctx.style.role("tableCell").get("min"), 9),
              "cardSub": as_float(ctx.style.role("cardSub").get("min"), 10)}
    deficit_h = table_rows_pt(ctx, block, rect, floors)[2] - rect[3]
    need_rows = int(math.ceil(deficit_h / ctx.cell_h)) if deficit_h > 0.5 else 0
    col_w = rect[2] / float(n_cols)
    deficit_w = MIN_TABLE_COL_W * n_cols - rect[2]
    need_cols = int(math.ceil(deficit_w / ctx.cell_w)) if deficit_w > 0.5 else 0
    repick = None
    if n_cols >= 5 and col_w < MIN_TABLE_COL_W:
        repick = {"archetype": "table", "slot": "table",
                  "reason": "a %d-column table does not read in %d pt; used the "
                            "full-width table slot" % (n_cols, int(round(col_w)))}
    return (need_rows, need_cols, repick)


def cards_repick(ctx, block, rect, count):
    """A cards block whose tiles are too narrow to hold a card."""
    cols, _ = tile_shape(count)
    gutter = as_float(ctx.style.card.get("gutterPt"), 30.0)
    tile_w = (rect[2] - gutter * (cols - 1)) / float(cols)
    if count < 4 or tile_w >= MIN_CARD_W:
        return None
    return {"archetype": "wide-cards", "slot": "cards",
            "reason": "%d cards leave %d pt each; used the wide card box"
                      % (count, int(round(tile_w)))}


def stat_repick(block, count):
    """A stat whose value is a sentence belongs in the statement slot."""
    for item in stat_items(block):
        value = str(item.get("value") or "").strip()
        if len(value.split()) > MAX_STAT_WORDS or len(value) > MAX_STAT_CHARS:
            return {"archetype": "statement" if count == 1 else "stat-band",
                    "slot": "statement" if count == 1 else "stat",
                    "reason": "a stat of %d words reads as a statement; used the %s slot"
                              % (len(value.split()),
                                 "statement" if count == 1 else "stat band")}
    return None


def block_need(ctx, block, rect, loose=False):
    """What a block needs and what it can give up, in whole grid cells.

    `need_*` is measured at the role floor, because a block only earns extra
    cells when the content will not fit even at the smallest size the role
    allows. `slack_*` is measured at the role's own size, because a neighbour
    only gives up cells it is not using.
    """
    kind = norm(block.get("type"))
    if kind not in MEASURED:
        return (0, 0, 0, 0, None)
    if kind == "table":
        need_rows, need_cols, repick = table_need(ctx, block, rect)
        return (need_rows, need_cols, 0, 0, repick if loose else None)

    stacks = block_paras(ctx, block, kind, rect)
    if not stacks or not any(stacks):
        return (0, 0, 0, 0, None)
    count = len(stacks)
    inset_x, _ = block_insets(ctx, kind)
    width, height, tile_cols, tile_rows = content_box(ctx, kind, text_rect(block, rect),
                                                      count, block=block)

    deficit = 0.0
    slack = None
    widest = 0.0
    names = set()
    for stack in stacks:
        for para in stack:
            names.add(para.role_name)
    floors = role_sizes(ctx.style, names, "min")
    defaults = role_sizes(ctx.style, names, "size")
    for stack in stacks:
        if not stack:
            continue
        used = stack_height(stack, width, floors)[0]
        if over_lines(stack, ctx.style):
            used = max(used, height + 1.0)
        deficit = max(deficit, used - height)
        used_now, lines_now, wide_now, _ = stack_height(stack, width, defaults)
        room = height - used_now
        if over_lines(stack, ctx.style) or lines_now > 1:
            wide_now = width           # a wrapped stack wants every point it has
        slack = room if slack is None else min(slack, room)
        widest = max(widest, wide_now)

    per_row = ctx.cell_h / float(tile_rows)
    need_rows = int(math.ceil(deficit / per_row)) if deficit > 0.5 else 0
    slack_rows = int(math.floor((slack or 0.0) / per_row)) if (slack or 0) > 0 else 0
    per_col = ctx.cell_w / float(tile_cols)
    spare_w = width - widest - 2 * inset_x
    slack_cols = int(math.floor(spare_w / per_col)) if spare_w > per_col else 0
    repick = None
    if loose and kind == "cards":
        repick = cards_repick(ctx, block, rect, count)
    if loose and kind == "stat":
        repick = stat_repick(block, count)
    return (need_rows, 0, slack_rows, min(slack_cols, fit.MAX_GROW), repick)


def has_source(bands):
    """Did this slide draw a source line: the last body row is then spoken for (F-3)."""
    return any(band[4] == "source" for band in bands)


def apply_fidelity(ctx, blocks, rects, bands):
    """Run the fidelity pass over one slide's rectangles. Returns the deviations.

    The rectangles come in already guarded, so the pass works from what the
    slide really looks like, and the furniture it must clear is the furniture
    that was actually drawn.
    """
    if ctx.fidelity == "strict" or not blocks:
        return []
    loose = ctx.fidelity == "loose"
    items = []
    for i, block in enumerate(blocks):
        kind = norm(block.get("type"))
        cell = fit.rect_to_cells(rects[i], ctx.cell_w, ctx.cell_h)
        if kind == "notes" or kind not in WRITERS:
            items.append(fit.Item(block, cell, movable=False))
            continue
        need_rows, need_cols, slack_rows, slack_cols, repick = block_need(
            ctx, block, rects[i], loose=loose)
        items.append(fit.Item(block, cell, need_rows=need_rows, need_cols=need_cols,
                              slack_rows=slack_rows, slack_cols=slack_cols,
                              repick=repick))
    band_cells = [fit.rect_to_band(b[:4], ctx.cell_w, ctx.cell_h) for b in bands]
    moves = fit.plan_slide(items, band_cells, ctx.cols, ctx.rows, ctx.fidelity,
                           module=ctx.module, cell_w=ctx.cell_w,
                           source=has_source(bands), warn=ctx.warn)
    for i, item in enumerate(items):
        # A block the pass left alone keeps the rectangle the guard gave it, to
        # the point: a trimmed block is a fraction off the row grid, and handing
        # it back its whole row would undo the trim.
        snapped = item.pt_rect()
        if not fit.same_cells(item.start, item.cells) or snapped:
            whole = fit.cells_to_rect(item.cells, ctx.cell_w, ctx.cell_h)
            y, h = whole[1], whole[3]
            if (item.cells["row"] == item.start["row"]
                    and item.cells["rows"] == item.start["rows"]):
                y, h = rects[i][1], rects[i][3]   # the guard's own trim stands
            # The module did not divide the grid evenly: the points win.
            x, w = snapped if snapped else (whole[0], whole[2])
            rects[i] = (x, y, w, h)
    for move in moves:
        ctx.adjusted.append(dict(move, slide=ctx.number))
    return moves


WRITERS = {
    "text": write_text_block,
    "bullets": write_bullets_block,
    "cards": write_cards_block,
    "stat": write_stat_block,
    "chart": write_chart_block,
    "table": write_table_block,
    "image": write_image_block,
    "quote": write_quote_block,
    "takeaway": write_takeaway_block,
    "placeholder": write_placeholder_block,
    "diagram": write_diagram_block,
    "notes": write_notes_block,
}


# --------------------------------------------- deck-wide type scale (Q4)


def pin_role(style, name, size):
    """Fix a role at one size for the whole deck, so the fit cannot move it."""
    spec = style.role(name)
    spec["size"] = float(size)
    spec["min"] = float(size)
    spec["max"] = float(size)


def effective_headline(slide_spec):
    """The headline a slide will draw, including one that arrives as a block."""
    text = furniture_text(slide_spec, "headline")
    if text:
        return text
    for block in slide_spec.get("blocks") or []:
        if norm(block.get("type")) == "headline":
            found = content_of(block).get("text") or block.get("title") or ""
            if str(found).strip():
                return str(found)
    return ""


def content_box(ctx, kind, rect, count, block=None):
    """(width, height, tile cols, tile rows) the text of a block is measured in."""
    inset_x, inset_y = block_insets(ctx, kind)
    if kind == "stat" and block is not None and stat_layout(ctx, block, rect) == "grid":
        # A pair is a number in a fixed gutter with its label beside it, so the
        # box the text is measured in is the label half (Q22).
        tile_cols = max(min(STAT_GRID_COLS, count), 1)
        tile_rows = int(math.ceil(count / float(tile_cols)))
        pair_w = (rect[2] - STAT_GRID_GUTTER * (tile_cols - 1)) / float(tile_cols)
        number_w = stat_grid_number_w(ctx, block, pair_w)
        box_w = pair_w - number_w
        box_h = rect[3] / float(tile_rows) - STAT_GRID_PAD
    elif kind in TILED:
        tile_cols, tile_rows = tile_shape(count)
        gutter = as_float(ctx.style.card.get("gutterPt"), 30.0)
        box_w = (rect[2] - gutter * (tile_cols - 1)) / float(tile_cols)
        box_h = (rect[3] - gutter * (tile_rows - 1)) / float(tile_rows)
    else:
        tile_cols, tile_rows = 1, 1
        box_w, box_h = rect[2], rect[3]
    return (max(box_w - 2 * inset_x, 24.0), max(box_h - 2 * inset_y, 8.0),
            tile_cols, tile_rows)


def table_grid(ctx, block):
    """(headers, rows, n_cols, n_rows) of a table block."""
    body = content_of(block)
    headers = [str(h) for h in (body.get("headers") or [])]
    rows = body.get("rows") or []
    n_cols = max(len(headers), max((len(r) for r in rows), default=0))
    n_rows = len(rows) + (1 if headers or rows else 0)
    return (headers, rows, n_cols, n_rows)


def table_fits(ctx, block, rect, sizes):
    """Does the whole table fit its block at these sizes.

    A cell may wrap: `maxLines` is a rule for prose, and a two-line row label is
    normal. What matters is that the rows it needs add up to the block's height.
    """
    headers, rows, n_cols, n_rows = table_grid(ctx, block)
    if not n_cols or not n_rows:
        return True
    return table_rows_pt(ctx, block, rect, sizes)[2] <= rect[3] + 1.0


def block_fits(ctx, block, rect, sizes, no_widows=False):
    """Does this block's content fit its rectangle with the roles at these sizes.

    `no_widows` also refuses a size that leaves a paragraph on a one-word line,
    which is how the widow rule survives a deck-wide size: the whole role takes
    one step together, or nobody does.
    """
    kind = norm(block.get("type"))
    if kind == "table":
        return table_fits(ctx, block, rect, sizes)
    stacks = block_paras(ctx, block, kind, rect)
    if not stacks:
        return True
    width, height, _, _ = content_box(ctx, kind, text_rect(block, rect), len(stacks),
                                      block=block)
    for stack in stacks:
        if not stack:
            continue
        used, _, _, widow = stack_height(stack, width, sizes)
        if used > height + 1.0 or over_lines(stack, ctx.style):
            return False
        if no_widows and widow:
            return False
    return True


TYPE_FOR_GROUP = {"text": "text", "bullets": "bullets", "cards": "cards",
                  "table": "table", "stat": "stat"}


def choose_group(ctx, roles, fits, clean=None):
    """Walk every role in a group down its range together; largest that fits wins.

    The roles in a group share a box (a card header sits over its body), so
    they step as one. When the winning size leaves a one-word last line
    somewhere, the group takes one more step together if that clears every
    widow in the deck: the widow rule from deck-planning section 7, applied to
    the deck instead of to one paragraph, so the size stays the same on every
    slide. Returns ({role: size}, [roles sitting at their floor]).
    """
    ladders = dict((role, fit.size_ladder(ctx.style.role(role))) for role in roles)
    steps = max(len(ladder) for ladder in ladders.values())

    def sizes_at(step):
        return dict((role, ladders[role][min(step, len(ladders[role]) - 1)])
                    for role in roles)

    chosen = None
    for step in range(steps):
        sizes = sizes_at(step)
        if not fits(sizes):
            continue
        chosen = sizes
        if clean is not None and not clean(sizes) and step + 1 < steps:
            relief = sizes_at(step + 1)
            # Never spend the floor on a widow: cut the copy instead
            # (deck-planning section 7, "cut before you shrink fonts").
            above_floor = all(relief[role] > ladders[role][-1] for role in roles)
            if above_floor and fits(relief) and clean(relief):
                chosen = relief
        break
    if chosen is None:
        chosen = dict((role, ladders[role][-1]) for role in roles)
    # At the floor only counts when the deck was stepped down to it: a role
    # whose style already sits at its floor was never shrunk.
    floored = [role for role in roles
               if len(ladders[role]) > 1 and chosen[role] <= ladders[role][-1]]
    return (chosen, floored)


def choose_headline_type(ctx, slides_spec):
    """One headline size for the whole deck: the largest that fits every slot.

    Every content headline has to fit its slot in no more lines than the role
    allows, so the tightest slot in the deck sets the size for all of them.
    """
    cases = []
    for slide_spec in slides_spec:
        if slide_kind(slide_spec) != "content":
            continue
        text = cap_first(effective_headline(slide_spec))
        slot = ctx.style.slot("headline", "content")
        if not text.strip() or not slot:
            continue
        cases.append((text, slot["w"], slot["h"]))
    if not cases:
        return {}
    # The tightest slot in the deck sets the size, so the smallest slot is the
    # one to look at first; a headline that cannot fit even at the floor is
    # reported as overflow rather than allowed to shrink every other slide.
    cases.sort(key=lambda case: (case[1] * case[2]))
    spec = ctx.style.role("headline")
    floor = as_float(spec.get("min"), 24)
    can = [case for case in cases
           if len(wrap_lines(case[0], cpl(case[1], floor,
                                         as_float(spec.get("k"), 0.5)))) <= max(
               as_int(spec.get("maxLines"), 2), 1)]
    if can:
        cases = can

    def measure(sizes):
        size = sizes["headline"]
        spec = ctx.style.role("headline")
        k = as_float(spec.get("k"), 0.5)
        lead = as_float(spec.get("leading"), 1.25)
        cap = as_int(spec.get("maxLines"), 2)
        fits = True
        widow = False
        for text, width, height in cases:
            lines = wrap_lines(text, cpl(width, size, k))
            if len(lines) > cap or len(lines) * size * lead > height + 1.0:
                fits = False
            if len(lines) > 1 and len(lines[-1].split()) == 1:
                widow = True
        return (fits, widow)

    sizes, floored = choose_group(ctx, ("headline",),
                                  lambda s: measure(s)[0],
                                  lambda s: not measure(s)[1])
    return {"sizes": sizes, "floored": floored}


def choose_block_type(ctx, plans):
    """One size per body role across the deck: text, bullets, cards, tables, stats."""
    out = {}
    floored = []
    for group, roles in DECK_TYPE_GROUPS:
        want = TYPE_FOR_GROUP.get(group)
        if not want:
            continue
        found = []
        for plan in plans:
            for block, rect in zip(plan["blocks"], plan["rects"]):
                if norm(block.get("type")) == want:
                    found.append((plan["kind"], block, rect))
        if not found:
            continue
        # A block that will not fit at the floor cannot be satisfied at any
        # size, so it must not drag the whole deck down: it is reported as
        # overflow instead. Only blocks that can fit set the deck-wide size.
        floors = role_sizes(ctx.style, roles, "min")
        cases = []
        for kind, block, rect in found:
            ctx.kind = kind
            if block_fits(ctx, block, rect, floors):
                cases.append((kind, block, rect))
        if not cases:
            out.update(floors)
            floored.extend(role for role in roles
                           if len(fit.size_ladder(ctx.style.role(role))) > 1)
            continue

        def fits(sizes, cases=cases, no_widows=False):
            # The groups already settled stay settled: a table head is sized
            # from the deck's cardSub, so the walk has to see it.
            merged = dict(out)
            merged.update(sizes)
            for kind, block, rect in cases:
                ctx.kind = kind
                if not block_fits(ctx, block, rect, merged, no_widows=no_widows):
                    return False
            return True

        sizes, at_floor = choose_group(ctx, roles, fits,
                                       lambda s, f=fits: f(s, no_widows=True))
        out.update(sizes)
        floored.extend(at_floor)
    return {"sizes": out, "floored": floored}


def apply_type_scale(ctx, chosen):
    """Pin the chosen sizes into the style and record them for the result."""
    sizes = chosen.get("sizes") or {}
    for role, size in sizes.items():
        pin_role(ctx.style, role, size)
        ctx.type_scale[role] = round(float(size), 1)
    for role in chosen.get("floored") or []:
        if role not in ctx.type_scale_floored:
            ctx.type_scale_floored.append(role)
    return sizes


# ------------------------------------------------------------------ slides


def add_notes(slide, text):
    try:
        frame = slide.notes_slide.notes_text_frame
        existing = frame.text
        frame.text = (existing + "\n" + str(text)).strip() if existing else str(text)
    except Exception as exc:  # noqa: BLE001 - notes are optional
        log("could not write speaker notes: %s" % exc)


def slide_kind(slide_spec):
    kind = norm(slide_spec.get("kind") or slide_spec.get("type"))
    if kind in ("cover", "title"):
        return "title"
    if kind in ("agenda", "contents"):
        return "agenda"
    if kind in ("divider", "section", "close", "thanks", "thankyou"):
        return "divider"
    if kind in ("content", "body", ""):
        pass
    if slide_spec.get("section") and not (slide_spec.get("blocks") or []):
        return "divider"
    return "content"


def is_closing(slide_spec):
    """A close or thank-you slide draws like a divider but is not a section."""
    return norm(slide_spec.get("kind") or slide_spec.get("type")) in ("close", "thanks",
                                                                      "thankyou")


def divider_numbers(slides_spec):
    """{slide index: ordinal} counting the section dividers in deck order."""
    out = {}
    count = 0
    for i, one in enumerate(slides_spec):
        if slide_kind(one) != "divider" or is_closing(one):
            continue
        count += 1
        out[i] = count
    return out


def furniture_text(slide_spec, *names):
    for name in names:
        value = slide_spec.get(name)
        if value:
            return str(value)
    return ""


def kicker_for(style, slide_spec):
    """The kicker line: the slide's own, else `kicker.pattern` over section and name.

    A derived kicker is recased to `kicker.case` (F-1); one the payload wrote
    is the writer's wording and goes in as given.
    """
    explicit = furniture_text(slide_spec, "kicker")
    if explicit:
        return explicit
    derived = fill_pattern(style.kicker_pattern(), slide_spec.get("section"),
                           slide_spec.get("name"))
    return apply_case(style.kicker_case(), derived)


def divider_title(style, slide_spec, ordinal):
    """"03 Competitive analysis" when `divider.numbered` is on, else the bare title."""
    title = cap_first(furniture_text(slide_spec, "headline", "section", "name"))
    if ordinal and style.divider_numbered() and title:
        return "%02d %s" % (int(ordinal), title)
    return title


# ------------------------------------------------- cover and agenda (Q5b)

CLAUSE_BREAKS = (": ", "; ", " - ", ", ")


def lead_clause(text, limit=70):
    """The first sentence of a note, cut back to its leading clause when long.

    A `story` or a brief `goal` is written for the planner, not for a slot:
    this takes the part that carries the point and leaves the enumeration.
    """
    flat = re.sub(r"\s+", " ", str(text or "")).strip()
    if not flat:
        return ""
    match = re.match(r"(.+?[.!?])(\s|$)", flat)
    out = (match.group(1) if match else flat).strip()
    if len(out) <= limit:
        return out
    for mark in CLAUSE_BREAKS:
        head = out.split(mark)[0]
        if head and len(head) <= limit:
            return head.strip()
    cut = out[:limit].rsplit(" ", 1)[0]
    return (cut or out[:limit]).rstrip(",;:") + "…"


def cover_slide(spec):
    """The cover the deck did not bring: the deck name over the brief's goal."""
    brief = spec.get("brief") or {}
    name = str(spec.get("name") or "Deck")
    out = {"id": "gen_cover", "kind": "title", "name": name, "headline": name,
           "generated": "cover", "pageNumber": False, "blocks": []}
    subtitle = lead_clause(brief.get("goal") or brief.get("takeaway"), 90)
    if subtitle:
        out["subtitle"] = subtitle
    stamp = spec.get("date") or brief.get("date")
    if stamp:
        out["date"] = str(stamp)
    return out


def agenda_sections(slides_spec):
    """[(title, one line)] per section divider: the first story under it."""
    out = []
    for i, one in enumerate(slides_spec):
        if slide_kind(one) != "divider" or is_closing(one):
            continue
        title = cap_first(furniture_text(one, "headline", "section", "name"))
        line = ""
        for later in slides_spec[i + 1:]:
            if slide_kind(later) == "divider":
                break
            if slide_kind(later) != "content":
                continue
            line = lead_clause(later.get("story") or later.get("message"), 110)
            if line:
                break
        out.append((title, line))
    return out


AGENDA_TOP = 9            # docs/research/deck-planning.md section 5, agenda
AGENDA_ROWS = 8
AGENDA_MAX = 6


def agenda_rows(style, count, cell_w=CELL, cell_h=CELL):
    """[(col, row, cols, rows)] per agenda entry, from `furniture.agendaList` (Q15).

    A style that laid out its own agenda list says where the entries go; the
    fallback is the agenda archetype's own rows 9, 11, 13 and 15.
    """
    slot = style.slot("agendaList", "agenda")
    if slot:
        step = slot["h"] / float(max(count, 1))
        out = []
        for i in range(count):
            cell = fit.rect_to_cells((slot["x"], slot["y"] + i * step, slot["w"], step),
                                     cell_w, cell_h)
            out.append((cell["col"], cell["row"], cell["cols"], cell["rows"]))
        return out
    step = max(AGENDA_ROWS // max(count, 1), 1)
    return [(2, AGENDA_TOP + i * step, 28, step) for i in range(count)]


def agenda_slide(spec, slides_spec, style, warn, cell_w=CELL, cell_h=CELL):
    """The agenda the deck did not bring: the dividers, one line each."""
    sections = agenda_sections(slides_spec)
    if not sections:
        return None
    if len(sections) > 4:
        warn("the agenda lists %d sections; this geometry reads four" % len(sections))
    brief = spec.get("brief") or {}
    statement = str(spec.get("agendaStatement") or "").strip()
    if not statement:
        statement = lead_clause(brief.get("goal"), 90) or str(spec.get("name") or "Agenda")
    shown = sections[:AGENDA_MAX]
    places = agenda_rows(style, len(shown), cell_w, cell_h)
    blocks = []
    for i, (title, line) in enumerate(shown):
        label = "%02d %s" % (i + 1, title) if style.divider_numbered() else title
        runs = [{"text": label, "bold": True, "color": "ink"}]
        if line:
            runs.append({"text": "   " + cap_first(line)})
        col, row, cols, rows = places[i]
        blocks.append({"id": "gen_agenda_%d" % (i + 1), "type": "text", "role": "text",
                       "col": col, "row": row, "cols": cols,
                       "rows": rows, "content": {"paragraphs": [{"runs": runs}]}})
    return {"id": "gen_agenda", "kind": "agenda", "name": "Agenda",
            "headline": statement, "generated": "agenda", "blocks": blocks}


def with_chrome(spec, slides_spec, style, warnings, cell_w=CELL, cell_h=CELL):
    """(slides, a cover was added).

    `cover` and `agenda` mean "the deck should have one", so true is satisfied
    by the deck's own title or agenda slide and only builds what is missing.
    Either flag set false keeps it off. Both default to on when the deck has no
    slide of that kind.
    """
    def warn(message):
        warnings.append(message)

    kinds = [slide_kind(one) for one in slides_spec]
    want_cover = spec.get("cover")
    want_agenda = spec.get("agenda")
    has_cover = "title" in kinds
    has_agenda = "agenda" in kinds
    if want_cover is None:
        want_cover = not has_cover
    if want_agenda is None:
        want_agenda = not has_agenda
    out = list(slides_spec)
    added_cover = False
    if want_cover and not has_cover:
        out.insert(0, cover_slide(spec))
        added_cover = True
    if want_agenda and not has_agenda:
        made = agenda_slide(spec, slides_spec, style, warn, cell_w, cell_h)
        if made is None:
            warn("no section dividers to list; no agenda was added")
        else:
            # After the cover, whether the deck brought its own or we made one.
            at = 0
            for i, one in enumerate(out):
                if slide_kind(one) == "title":
                    at = i + 1
                    break
            out.insert(at, made)
    return (out, added_cover)


def fold_furniture_blocks(ctx, slide_spec):
    """Move headline and kicker blocks into the slide's own fields.

    A `headline` block fills an empty slide headline and is dropped when the
    slide already has one, so a slide never draws two headlines. Same for
    `kicker`. Returns (slide spec copy, blocks to draw, {slot: block id}).
    """
    blocks = slide_spec.get("blocks") or []
    if not any(norm(b.get("type")) in FURNITURE_TYPES for b in blocks):
        return slide_spec, blocks, {}
    spec = dict(slide_spec)
    rest = []
    ids = {}
    for block in blocks:
        kind = norm(block.get("type"))
        if kind not in FURNITURE_TYPES:
            rest.append(block)
            continue
        text = content_of(block).get("text") or block.get("title") or ""
        if not str(spec.get(kind) or "").strip() and str(text).strip():
            spec[kind] = str(text)
            ids[kind] = block.get("id")
        else:
            ctx.warn("block '%s' repeats the slide %s; dropped"
                     % (block.get("id"), kind))
    return spec, rest, ids


def chrome_rect(entry):
    return (as_float(entry.get("x"), 0.0), as_float(entry.get("y"), 0.0),
            as_float(entry.get("w"), 8.0), as_float(entry.get("h"), 8.0))


def draw_chrome(ctx, slide):
    """The format's own marks, from `chrome` (Q15).

    Neither content nor a type role: the three circles a dark deck carries
    top left, the hairline under its kicker row, a wordmark, a logo. The marks
    are drawn before the furniture so a block is guarded against them, and are
    recorded at `role: "chrome"` so a preview can draw them too.
    """
    style = ctx.style
    for entry in style.chrome_for(ctx.kind):
        kind = norm(entry.get("kind"))
        x, y, w, h = chrome_rect(entry)
        colour = entry.get("color") or "muted"
        if kind == "dots":
            count = max(as_int(entry.get("count"), CHROME_DOTS), 1)
            size = min(h, w / float(count))
            gap = (w - size * count) / float(count - 1) if count > 1 else 0.0
            for i in range(count):
                at = (x + i * (size + gap), y, size, size)
                dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, emu(at[0]), emu(at[1]),
                                             emu(at[2]), emu(at[3]))
                fill = entry.get("fill")
                if fill:
                    style_shape(dot, style.hex_of(fill, ctx.kind), None, 0)
                else:
                    style_shape(dot, None, style.hex_of(colour, ctx.kind),
                                as_float(entry.get("linePt"), 1.0))
                ctx.record("rule", CHROME_ROLE, at)
        elif kind == "line":
            draw_rule(ctx, slide, (x, y, w, max(h, HAIRLINE_PT)), colour, role=CHROME_ROLE,
                      width_pt=max(h, HAIRLINE_PT))
        elif kind == "wordmark":
            text = str(entry.get("text") or style.copy.get("wordmark") or "").strip()
            if not text:
                continue
            para = para_for(ctx, text, "wordmark", align=entry.get("align") or "left",
                            color=colour)
            if entry.get("size"):
                para.size = as_float(entry.get("size"), para.size)
            add_text(ctx, slide, (x, y, w, h), [para], CHROME_ROLE)
        elif kind == "logo":
            path = str(entry.get("path") or "")
            if not path or not os.path.exists(path):
                ctx.warn("chrome logo not found: %s" % (path or "(no path)"))
                continue
            try:
                pic = slide.shapes.add_picture(path, emu(x), emu(y))
            except Exception as exc:  # noqa: BLE001 - chrome never stops a build
                ctx.warn("chrome logo could not be placed: %s" % exc)
                continue
            native_w = pic.width / float(EMU_PER_PT)
            native_h = pic.height / float(EMU_PER_PT)
            scale = min(w / native_w, h / native_h) if native_w and native_h else 1.0
            pic.width, pic.height = emu(native_w * scale), emu(native_h * scale)
            ctx.record("image", CHROME_ROLE,
                       (x, y, native_w * scale, native_h * scale),
                       os.path.basename(path))
        else:
            ctx.warn("unknown chrome kind '%s'; nothing drawn" % (entry.get("kind") or ""))


def source_text(source):
    """"Source: X", unless the writer already wrote the prefix."""
    text = str(source)
    return text if norm(text).startswith("source") else "Source: %s" % text


def shorten_source(ctx, text, width):
    """One line at the caption floor: drop the brackets, then the last clauses (Q15).

    A source that will not fit its slot is a source nobody reads. The detail in
    brackets goes first, then the clauses after each ";" from the back, and the
    full wording is kept for the notes so nothing is lost.
    """
    spec = ctx.style.role("caption")
    size = as_float(spec.get("min"), CAPTION_FLOOR)
    chars = cpl(max(width, 24.0), size, as_float(spec.get("k"), 0.5))

    def fits(value):
        return len(wrap_lines(value, chars)) <= 1

    if fits(text):
        return text
    flat = re.sub(PARENTHETICAL, "", text).strip()
    tries = [flat]
    for splitter in (CLAUSE_SPLIT, None):
        if splitter:
            parts = [p.strip() for p in flat.split(splitter) if p.strip()]
            join = splitter + " "
        else:
            parts = [p.strip() for p in SENTENCE_SPLIT.split(flat) if p.strip()]
            join = " "
        for keep in range(len(parts) - 1, 0, -1):
            tries.append(join.join(parts[:keep]))
    for candidate in tries:
        if candidate and fits(candidate):
            return candidate
    # Nothing left to drop: cut on a word, never in the middle of one.
    head = wrap_lines(flat or text, max(chars - 1, 8))[0].rstrip(" ,;:.")
    return (head + "\u2026") if head else text


WORD_RE = re.compile(r"[a-z0-9]+")
STOP_WORDS = frozenset(("a", "an", "and", "the", "of", "to", "in", "for", "on", "with",
                        "is", "are", "that", "it", "as", "at", "by", "from", "its",
                        "this", "their", "they", "be", "or"))


def token_set(text):
    """The words that carry the meaning of a line, lowercased."""
    return set(w for w in WORD_RE.findall(norm(text)) if w not in STOP_WORDS)


def repeats(text, other, threshold=DUPLICATE_OVERLAP):
    """True when one line says what the other already said (Q15)."""
    a, b = token_set(text), token_set(other)
    if not a or not b:
        return False
    return len(a & b) / float(max(len(a), len(b))) >= threshold


def block_first_paragraph(block):
    """(text, where) for the paragraph a block leads with, else (None, None)."""
    kind = norm(block.get("type"))
    body = content_of(block)
    if kind == "text":
        raw = body.get("paragraphs")
        if raw is None and body.get("runs"):
            return (" ".join(str(r.get("text", "")) for r in body.get("runs")), "whole")
        if raw is None:
            return ((str(body.get("text")), "whole") if body.get("text") else (None, None))
        if isinstance(raw, list) and raw:
            first = raw[0]
            if isinstance(first, dict):
                text = first.get("text") or " ".join(str(r.get("text", ""))
                                                     for r in (first.get("runs") or []))
            else:
                text = first
            return (str(text or ""), "paragraphs")
    elif kind == "bullets":
        items = body.get("items")
        if isinstance(items, list) and items:
            first = items[0]
            text = first.get("text") if isinstance(first, dict) else first
            return (str(text or ""), "items")
    elif kind == "takeaway":
        return ((str(body.get("text")), "whole") if body.get("text") else (None, None))
    return (None, None)


def without_first(block, where):
    """The block without its first paragraph, or None when nothing is left."""
    if where == "whole":
        return None
    body = dict(content_of(block))
    rest = list(body.get(where) or [])[1:]
    if not rest:
        return None
    body[where] = rest
    out = dict(block)
    out["content"] = body
    return out


def drop_duplicate_content(ctx, slide_spec, blocks):
    """Drop a first paragraph that repeats the headline; warn on twin titles (Q15).

    Redundancy costs the reader a second reading of the same sentence and costs
    the slide the rows it sits on (Mayer, and item 14 of the output-quality
    research). The paragraph goes, the block goes with it when nothing is left,
    and both are reported. Strict draws the plan as approved and keeps them.
    """
    titles = {}
    headline = furniture_text(slide_spec, "headline")
    out = []
    for block in blocks:
        title = norm(block.get("title"))
        if title:
            if title in titles:
                ctx.warn("blocks '%s' and '%s' carry the same title '%s'"
                         % (titles[title], block.get("id"), block.get("title")))
            else:
                titles[title] = block.get("id")
        text, where = block_first_paragraph(block)
        if (ctx.fidelity == "strict" or not headline or not text
                or not repeats(text, headline)):
            out.append(block)
            continue
        ctx.adjusted.append({"slide": ctx.number, "blockId": block.get("id"),
                             "reason": DUPLICATE_REASON,
                             "from": {"text": preview(text)}, "to": {"text": ""}})
        trimmed = without_first(block, where)
        if trimmed is None:
            ctx.warn("block '%s' repeats the headline and says nothing else; dropped"
                     % block.get("id"))
            continue
        ctx.warn("block '%s' opens by repeating the headline; that line was dropped"
                 % block.get("id"))
        out.append(trimmed)
    return out


def as_bullets(block):
    """A `text` block rewritten as `bullets`: its paragraphs become the items."""
    body = content_of(block)
    raw = body.get("paragraphs")
    if raw is None and body.get("runs"):
        raw = [{"runs": body.get("runs")}]
    if raw is None and body.get("text"):
        raw = [body.get("text")]
    items = []
    for item in raw or []:
        if isinstance(item, dict):
            text = item.get("text") or " ".join(str(r.get("text", ""))
                                                for r in (item.get("runs") or []))
        else:
            text = item
        if str(text).strip():
            items.append(str(text))
    out = dict(block)
    out["type"] = "bullets"
    out["content"] = {"items": items}
    return out


def harmonise_rows(ctx, blocks):
    """One treatment per row of parallel content (Q15).

    Blocks planned on one row read as one row, so they have to look like one:
    the same marker and the same title treatment. Bullets win, because the
    reader can see a marker and cannot see its absence, and a converted block
    keeps its title as a title rather than turning it into the first bullet.
    A row that is not all text-like is not a row of parallel copy and is left
    alone (a `text` rail beside a table is a rail).
    """
    rows = {}
    for i, block in enumerate(blocks):
        if norm(block.get("type")) == "notes":
            continue
        rows.setdefault(as_int(block.get("row"), 1), []).append(i)
    out = list(blocks)
    for _, group in sorted(rows.items()):
        kinds = [norm(blocks[i].get("type")) for i in group]
        if len(group) < 2 or any(k not in TEXT_LIKE for k in kinds):
            continue
        if "bullets" not in kinds or len(set(kinds)) < 2:
            continue
        for i in group:
            block = blocks[i]
            if norm(block.get("type")) == "bullets":
                continue
            out[i] = as_bullets(block)
            ctx.adjusted.append({"slide": ctx.number, "blockId": block.get("id"),
                                 "reason": PARALLEL_REASON,
                                 "from": {"type": "text"}, "to": {"type": "bullets"}})
            ctx.warn("block '%s' shares a row with bullets; drawn with the same marker"
                     % block.get("id"))
    return out


def draw_furniture(ctx, slide, slide_spec, number, ordinal=None, ids=None):
    """Draw the standing furniture for this slide kind, from the style tokens.

    Every slot is drawn at the height of the text it holds, so an oversized
    headline slot never reserves a band the body then collides with.
    """
    style = ctx.style
    kind = ctx.kind
    ids = ids or {}
    draw_chrome(ctx, slide)

    if kind == "title":
        wordmark = furniture_text(slide_spec, "wordmark") or str(style.copy.get("wordmark") or "")
        slot = style.slot("coverWordmark", kind)
        if wordmark and slot:
            add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]),
                     [para_for(ctx, wordmark, "wordmark")], "wordmark", shrink=True)
        slot = style.slot("coverTitle", kind)
        title = cap_first(furniture_text(slide_spec, "headline", "name"))
        if slot:
            add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]),
                     [para_for(ctx, title, "coverTitle")], "coverTitle",
                     block_id=ids.get("headline"), shrink=True)
        subtitle = furniture_text(slide_spec, "subtitle", "kicker", "message")
        slot = style.slot("coverSubtitle", kind)
        if subtitle and slot:
            add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]),
                     [para_for(ctx, subtitle, "coverSubtitle")], "coverSubtitle", shrink=True)
        stamp = furniture_text(slide_spec, "date")
        slot = style.slot("coverDate", kind)
        if stamp and slot:
            add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]),
                     [para_for(ctx, stamp, "coverDate")], "coverDate", shrink=True)
        return

    if kind == "divider":
        title = divider_title(style, slide_spec, ordinal)
        slot = style.slot("dividerTitle", kind)
        if slot:
            add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]),
                     [para_for(ctx, title, "dividerTitle")], "dividerTitle",
                     block_id=ids.get("headline"), shrink=True)
        subtitle = furniture_text(slide_spec, "subtitle", "message")
        slot = style.slot("dividerSubtitle", kind)
        if subtitle and slot:
            add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]),
                     [para_for(ctx, subtitle, "dividerSubtitle")], "dividerSubtitle",
                     shrink=True)
        draw_page_number(ctx, slide, slide_spec, number)
        return

    if kind == "agenda":
        eyebrow = furniture_text(slide_spec, "eyebrow") or str(style.copy.get("agendaEyebrow") or "")
        slot = style.slot("agendaEyebrow", kind)
        if eyebrow and slot:
            add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]),
                     [para_for(ctx, eyebrow, "agendaEyebrow")], "agendaEyebrow", shrink=True)
        slot = style.slot("agendaStatement", kind)
        statement = cap_first(furniture_text(slide_spec, "headline", "message", "name"))
        if slot:
            add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]),
                     [para_for(ctx, statement, "agendaStatement")], "agendaStatement",
                     block_id=ids.get("headline"), shrink=True)
        draw_page_number(ctx, slide, slide_spec, number)
        return

    # content
    kicker = kicker_for(style, slide_spec)
    if kicker:
        slot = style.slot("kicker", kind)
        if slot and style.enabled("kicker", True):
            add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]),
                     [para_for(ctx, kicker, "kicker")], "kicker",
                     block_id=ids.get("kicker"), shrink=True)
        elif furniture_text(slide_spec, "kicker") or ids.get("kicker"):
            who = ("block '%s'" % ids["kicker"]) if ids.get("kicker") else "'%s'" % preview(kicker)
            ctx.warn("this style has no kicker; %s dropped" % who)
    headline = cap_first(furniture_text(slide_spec, "headline"))
    slot = style.slot("headline", kind)
    if not headline:
        ctx.warn("slide has no headline")
    if slot and headline:
        add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]),
                 [para_for(ctx, headline, "headline")], "headline",
                 block_id=ids.get("headline"), shrink=True)
    rule = style.slot("hairline", kind)
    if rule:
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, emu(rule["x"]), emu(rule["y"]),
                                      emu(rule["w"]), emu(max(rule["h"], 0.75)))
        style_shape(line, style.hex_of("line", kind), None, 0)
        ctx.record("rule", "hairline", (rule["x"], rule["y"], rule["w"], rule["h"]))
    source = furniture_text(slide_spec, "source")
    if source and style.enabled("source", True):
        slot = style.slot("source", kind)
        if slot:
            text = source_text(source)
            short = shorten_source(ctx, text, slot["w"])
            if short != text:
                ctx.notes_extra.append(text)
                ctx.adjusted.append({"slide": ctx.number, "blockId": None,
                                     "reason": SOURCE_REASON,
                                     "from": {"text": text}, "to": {"text": short}})
                ctx.warn("the source line was shortened to one line; the full text is "
                         "in the notes")
            add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]),
                     [para_for(ctx, short, "caption")], "source", shrink=True)
    draw_page_number(ctx, slide, slide_spec, number)


def furniture_bands(ctx):
    """The furniture rectangles a block must clear, from what was just drawn."""
    out = []
    for entry in ctx.shapes:
        if entry.get("blockId") or entry.get("role") not in GUARDED_ROLES:
            continue
        out.append((entry["x"], entry["y"], entry["w"], entry["h"], entry["role"]))
    return out


def crosses(rect, band):
    return (rect[0] < band[0] + band[2] and rect[0] + rect[2] > band[0]
            and rect[1] < band[1] + band[3] and rect[1] + rect[3] > band[1])


def guard_rect(ctx, block, rect, bands, others=()):
    """Move or trim a block that sits on the furniture, and say so.

    A block that intrudes into the kicker, headline or hairline band drops to
    the next whole row below it. One that runs into the source or page-number
    line moves up when it can keep its height and nothing else is in the way,
    and is trimmed when it cannot. Rows stay whole, so columns on one slide
    keep their alignment.
    """
    if not bands:
        return rect
    cell = max(ctx.cell_h, 1.0)
    x, y, w, h = rect
    notes = []

    def clear(top, height):
        if any(crosses((x, top, w, height), band) for band in bands):
            return False
        return not any(crosses((x, top, w, height), other) for other in others)

    for _ in range(4):
        hit = None
        for band in bands:
            if crosses((x, y, w, h), band):
                hit = band
                break
        if hit is None:
            break
        if hit[1] + hit[3] / 2.0 <= y + h / 2.0:
            floor_y = math.ceil((hit[1] + hit[3]) / cell - 0.001) * cell
            if floor_y >= y + h - MIN_SIDE:
                break
            h = max(h - (floor_y - y), cell)
            y = floor_y
            notes.append(("moved", hit[4]))
            continue
        top_y = math.floor((hit[1] + 0.001) / cell) * cell
        lift = top_y - h
        if lift >= 0 and clear(lift, h):
            y = lift
            notes.append(("moved", hit[4]))
        elif top_y > y + MIN_SIDE:
            h = max(top_y - y, cell)
            notes.append(("trimmed", hit[4]))
        elif hit[1] > y + 2 * MIN_SIDE:
            h = hit[1] - y          # last resort: trim to the band, off the row grid
            notes.append(("trimmed", hit[4]))
        else:
            break
    if notes:
        what, where = notes[-1]
        row = int(round(y / cell)) + 1 if what == "moved" else int(round((y + h) / cell))
        ctx.warn("block '%s' ran into the %s; %s to row %d"
                 % (block.get("id") or block.get("type") or "?", where, what, row))
    return (x, y, w, h)


def draw_page_number(ctx, slide, slide_spec, number):
    style = ctx.style
    if not style.enabled("pageNumber", True):
        return
    value = slide_spec.get("pageNumber")
    if value is False:
        return
    if value is None:
        if ctx.kind == "title" and not style.enabled("pageNumberOnTitle", False):
            return
        value = number
    slot = style.slot("pageNumber", ctx.kind)
    if not slot:
        return
    para = para_for(ctx, str(value), "pageNumber", align=slot.get("align") or "right")
    add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]), [para], "pageNumber")


def pick_layout(prs, want_title):
    """A blank layout for our own drawing, or one with a title placeholder in template mode."""
    layouts = list(prs.slide_layouts)
    if not layouts:
        return None
    if want_title:
        for layout in layouts:
            names = [ph.placeholder_format.type for ph in layout.placeholders]
            if PP_PLACEHOLDER.TITLE in names or PP_PLACEHOLDER.CENTER_TITLE in names:
                return layout
    for layout in layouts:
        if norm(layout.name) == "blank":
            return layout
    return min(layouts, key=lambda lay: len(list(lay.placeholders)))


# ------------------------------------------------- template layouts (Q8)

TITLE_PH = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE,
            PP_PLACEHOLDER.VERTICAL_TITLE)
BODY_PH = (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.VERTICAL_BODY,
           PP_PLACEHOLDER.VERTICAL_OBJECT)
PICTURE_PH = (PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.BITMAP, PP_PLACEHOLDER.SLIDE_IMAGE)
CHROME_PH = (PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.HEADER,
             PP_PLACEHOLDER.SLIDE_NUMBER)
SECTION_WORDS = ("section", "divider", "chapter", "part")
LAYOUT_KINDS = ("titleSlide", "section", "titleOnly", "titleBody", "twoContent", "blank")


def layout_profile(layout):
    """What a template layout offers: its title, body and picture placeholders.

    `kind` is the shape a deck names it by, so a slide can ask for the layout
    it means rather than for the first one with a title on it.
    """
    title = None
    subtitle = None
    bodies = []
    pictures = []
    for shape in layout.placeholders:
        kind = shape.placeholder_format.type
        idx = shape.placeholder_format.idx
        if kind in CHROME_PH:
            continue
        if kind in TITLE_PH and title is None:
            title = idx
        elif kind == PP_PLACEHOLDER.SUBTITLE and subtitle is None:
            subtitle = idx
        elif kind in PICTURE_PH:
            pictures.append(idx)
        elif kind in BODY_PH:
            bodies.append(idx)
    name = norm(layout.name)
    if title is None:
        shape = "blank"
    elif subtitle is not None and not bodies:
        shape = "titleSlide"
    elif any(word in name for word in SECTION_WORDS):
        shape = "section"
    elif len(bodies) + len(pictures) >= 2:
        shape = "twoContent"
    elif bodies or pictures:
        shape = "titleBody"
    else:
        shape = "titleOnly"
    return {"layout": layout, "name": layout.name, "kind": shape, "title": title,
            "subtitle": subtitle, "bodies": bodies, "pictures": pictures,
            "slots": len(bodies) + len(pictures)}


def layout_profiles(prs):
    return [layout_profile(layout) for layout in prs.slide_layouts]


def wanted_layout(slide_spec, blocks):
    """What a slide needs from a layout: a title, this many bodies, a picture."""
    kind = slide_kind(slide_spec)
    text_like = [b for b in blocks if norm(b.get("type")) in TEXT_LIKE]
    pictures = [b for b in blocks if norm(b.get("type")) == "image"]
    if kind == "title":
        return {"kind": "titleSlide", "bodies": 0, "pictures": 0}
    if kind == "divider":
        return {"kind": "section", "bodies": 0, "pictures": 0}
    if kind == "agenda":
        return {"kind": "titleBody", "bodies": 0, "pictures": 0}
    if not blocks:
        return {"kind": "titleOnly", "bodies": 0, "pictures": 0}
    return {"kind": "twoContent" if len(text_like) > 1 else "titleBody",
            "bodies": min(len(text_like), 2), "pictures": min(len(pictures), 1)}


def score_layout(profile, want):
    """How well a layout matches what the slide asked for. Higher wins."""
    score = 0.0
    # A slide with a picture needs a picture placeholder more than it needs the
    # layout it would otherwise have taken.
    if want["pictures"]:
        score += 120.0 if profile["pictures"] else -60.0
    elif profile["pictures"]:
        score -= 20.0
    if profile["kind"] == want["kind"]:
        score += 100.0
    if profile["title"] is not None:
        score += 20.0
    have = len(profile["bodies"])
    score -= abs(have - want["bodies"]) * 12.0
    if want["bodies"] > have:
        score -= 8.0                      # a missing slot costs more than a spare
    if want["kind"] == "titleSlide" and profile["subtitle"] is not None:
        score += 15.0
    score -= profile["slots"] * 0.1       # ties go to the simpler layout
    return score


def choose_layout(profiles, slide_spec, blocks, fallback):
    """The template layout whose placeholders best match this slide (Q8)."""
    usable = [p for p in profiles if p["title"] is not None or p["slots"]]
    if not usable:
        return None
    want = wanted_layout(slide_spec, blocks)
    best = max(usable, key=lambda p: score_layout(p, want))
    if score_layout(best, want) <= 0 and fallback is not None:
        return None
    return best


def placeholder_at(slide, idx):
    for shape in slide.placeholders:
        if shape.placeholder_format.idx == idx:
            return shape
    return None


def placeholder_rect(shape):
    return (shape.left / float(EMU_PER_PT), shape.top / float(EMU_PER_PT),
            shape.width / float(EMU_PER_PT), shape.height / float(EMU_PER_PT))


def fill_placeholder_text(ctx, shape, paras, role):
    """Write paragraphs into a template placeholder, in the master's own face.

    The theme carries the family and colour; only the size and the wording are
    ours, so a client template comes out looking like the client's template.
    """
    frame = shape.text_frame
    frame.word_wrap = True
    size = as_float(ctx.style.role(role).get("size"), 12)
    first = True
    for para in paras:
        p = frame.paragraphs[0] if first else frame.add_paragraph()
        first = False
        p.text = ""
        run = p.add_run()
        run.text = para
        run.font.size = Pt(size)
    return size


def block_lines(block):
    """The block's copy as plain lines, for a template body placeholder."""
    kind = norm(block.get("type"))
    body = content_of(block)
    out = []
    if block.get("title"):
        out.append(str(block.get("title")))
    if kind == "bullets":
        stack = []
        flat_items(body.get("items"), stack)
        out.extend(stack)
    else:
        raw = body.get("paragraphs")
        if raw is None and body.get("runs"):
            raw = [{"runs": body.get("runs")}]
        if raw is None and body.get("text"):
            raw = [body.get("text")]
        for item in raw or []:
            if isinstance(item, dict):
                text = item.get("text") or " ".join(str(r.get("text", ""))
                                                    for r in (item.get("runs") or []))
            else:
                text = item
            if str(text).strip():
                out.append(str(text))
    return out or [str(block.get("title") or "")]


def flat_items(items, out):
    for item in items or []:
        if isinstance(item, dict):
            if str(item.get("text", "")).strip():
                out.append(str(item.get("text")))
            flat_items(item.get("children"), out)
        elif str(item).strip():
            out.append(str(item))


def fill_layout(ctx, slide, profile, slide_spec, blocks, ids):
    """Put the slide into the layout's placeholders. Returns (kept shapes, rest).

    The headline goes in the title, the first text or bullets blocks in the
    body placeholders, the first image in the picture placeholder. Every block
    that is left is drawn as usual, inside the body placeholder's rectangle.
    """
    keep = []
    rest = list(blocks)
    title_shape = None
    if profile["title"] is not None:
        shape = placeholder_at(slide, profile["title"])
        text = cap_first(furniture_text(slide_spec, "headline", "name"))
        if ctx.kind == "divider":
            text = divider_title(ctx.style, slide_spec, ctx.ordinal)
        if shape is not None and text:
            role = {"title": "coverTitle", "divider": "dividerTitle",
                    "agenda": "agendaStatement"}.get(ctx.kind, "headline")
            size = fill_placeholder_text(ctx, shape, [text], role)
            ctx.record("text", "headline", placeholder_rect(shape), text, size,
                       ids.get("headline"))
            keep.append(profile["title"])
            title_shape = shape
    if profile["subtitle"] is not None:
        shape = placeholder_at(slide, profile["subtitle"])
        text = furniture_text(slide_spec, "subtitle", "message")
        if shape is not None and text:
            role = "coverSubtitle" if ctx.kind == "title" else "dividerSubtitle"
            size = fill_placeholder_text(ctx, shape, [text], role)
            ctx.record("text", role, placeholder_rect(shape), text, size)
            keep.append(profile["subtitle"])
    if ctx.kind != "content":
        return (keep, rest, title_shape)
    for idx in profile["pictures"]:
        picture = next((b for b in rest if norm(b.get("type")) == "image"
                        and os.path.exists(str(content_of(b).get("path") or ""))), None)
        if picture is None:
            break
        shape = placeholder_at(slide, idx)
        if shape is None:
            break
        try:
            placed = shape.insert_picture(str(content_of(picture).get("path")))
        except Exception as exc:  # noqa: BLE001 - the block is drawn instead
            ctx.warn("template picture placeholder could not take block '%s': %s"
                     % (picture.get("id"), exc))
            break
        ctx.record("image", "templatePicture", placeholder_rect(placed),
                   os.path.basename(str(content_of(picture).get("path"))), None,
                   picture.get("id"))
        keep.append(idx)
        rest.remove(picture)
    for idx in profile["bodies"]:
        block = next((b for b in rest if norm(b.get("type")) in TEXT_LIKE), None)
        if block is None:
            break
        shape = placeholder_at(slide, idx)
        if shape is None:
            break
        lines = block_lines(block)
        role = "bullets" if norm(block.get("type")) == "bullets" else "body"
        size = fill_placeholder_text(ctx, shape, lines, role)
        ctx.record("text", "templateBody", placeholder_rect(shape), " ".join(lines),
                   size, block.get("id"))
        keep.append(idx)
        rest.remove(block)
    return (keep, rest, title_shape)


def strip_placeholders(slide, keep=()):
    keep = set(keep)
    for shape in list(slide.placeholders):
        if shape.placeholder_format.idx in keep:
            continue
        if shape.has_text_frame and shape.text_frame.text.strip():
            continue
        shape._element.getparent().remove(shape._element)


def plan_slide(ctx, prs, layout, slide_spec, number, template_mode, ordinal=None,
               page=None, profiles=None):
    """Add the slide, draw its furniture and settle every block rectangle.

    Blocks are not drawn here: the deck-wide type scale needs every slide's
    final geometry before the first body word is written (Q4), so the build
    plans all the slides, chooses the sizes, then draws. In template mode the
    layout is picked per slide from the template's own layouts (Q8).
    """
    ctx.kind = slide_kind(slide_spec)
    ctx.ordinal = ordinal
    profile = None
    if template_mode and profiles:
        profile = choose_layout(profiles, slide_spec, slide_spec.get("blocks") or [],
                                layout)
        if profile is not None:
            layout = profile["layout"]
    slide = prs.slides.add_slide(layout)
    ctx.shapes = []
    ctx.seq = 0
    ctx.number = number
    name = slide_spec.get("name") or slide_spec.get("id") or ("Slide %d" % number)
    ctx.where = "slide %d '%s'" % (number, name)
    ctx.headline = cap_first(furniture_text(slide_spec, "headline"))
    ctx.notes_extra = []
    mark = len(ctx.adjusted)
    slide_spec, blocks, ids = fold_furniture_blocks(ctx, slide_spec)
    if ctx.kind == "content":
        blocks = drop_duplicate_content(ctx, slide_spec, blocks)
        blocks = harmonise_rows(ctx, blocks)

    used_title = None
    keep = []
    if profile is not None:
        keep, blocks, used_title = fill_layout(ctx, slide, profile, slide_spec, blocks, ids)
    strip_placeholders(slide, keep=keep)
    if used_title is None:
        set_background(slide, ctx.style.background(ctx.kind))

    if used_title is None:
        draw_furniture(ctx, slide, slide_spec, number if page is None else page,
                       ordinal=ordinal, ids=ids)
    else:
        # The template drew the title: ours would be a second one. The marks the
        # style carries and the kicker still belong to the deck.
        draw_chrome(ctx, slide)
        if ctx.kind == "content":
            kicker = kicker_for(ctx.style, slide_spec)
            if kicker and ctx.style.enabled("kicker", True):
                slot = ctx.style.slot("kicker", ctx.kind)
                if slot:
                    add_text(ctx, slide, (slot["x"], slot["y"], slot["w"], slot["h"]),
                             [para_for(ctx, kicker, "kicker")], "kicker",
                             block_id=ids.get("kicker"), shrink=True)
        draw_page_number(ctx, slide, slide_spec, number if page is None else page)

    bands = furniture_bands(ctx)
    planned = [ctx.rect(block) for block in blocks]
    rects = list(planned)
    for i, block in enumerate(blocks):
        if norm(block.get("type")) == "notes":
            continue
        others = [r for j, r in enumerate(planned)
                  if j != i and norm(blocks[j].get("type")) != "notes"]
        rects[i] = guard_rect(ctx, block, planned[i], bands, others)
    adjusted = apply_fidelity(ctx, blocks, rects, bands)

    # The copy changes (a duplicate line, a harmonised marker, a shortened
    # source) are reported beside the geometry, but never badge a shape: they
    # carry text where a move carries cells.
    text_moves = [m for m in ctx.adjusted[mark:] if "text" in (m.get("from") or {})
                  or "type" in (m.get("from") or {})]
    return {"slide": slide, "spec": slide_spec, "number": number, "name": name,
            "kind": ctx.kind, "blocks": blocks, "rects": rects, "bands": bands,
            "shapes": ctx.shapes, "seq": ctx.seq, "where": ctx.where,
            "ordinal": ordinal, "adjusted": adjusted, "notes": list(ctx.notes_extra),
            "textAdjusted": text_moves}


def without_last(block):
    """(the block minus its last paragraph, that paragraph), or (None, None)."""
    kind = norm(block.get("type"))
    key = "items" if kind == "bullets" else "paragraphs"
    body = dict(content_of(block))
    items = list(body.get(key) or [])
    if len(items) < 2:
        return (None, None)
    last = items[-1]
    if isinstance(last, dict):
        text = last.get("text") or " ".join(str(r.get("text", ""))
                                            for r in (last.get("runs") or []))
    else:
        text = last
    body[key] = items[:-1]
    out = dict(block)
    out["content"] = body
    return (out, str(text or ""))


def move_overflow_to_notes(ctx, plans):
    """At the floor, the last paragraph goes to the notes, not to the bin (Q15).

    Cutting a sentence in half is the one thing a fit must never do. When a
    text or bullets block still overflows at the role's floor, its last
    paragraph moves to the speaker notes and the move is reported as
    `moved to notes`. Strict keeps the overflow, because strict draws the plan
    as approved and says so in the findings instead.
    """
    if ctx.fidelity == "strict":
        return
    title_floor = {"blockTitle": as_float(ctx.style.role("blockTitle").get("min"), 12)}
    for plan in plans:
        ctx.where = plan["where"]
        ctx.number = plan["number"]
        for i, block in enumerate(plan["blocks"]):
            kind = norm(block.get("type"))
            if kind not in TEXT_LIKE:
                continue
            rect = text_rect(block, plan["rects"][i])
            for _ in range(8):
                stacks = block_paras(ctx, block, kind=kind)
                paras = stacks[0] if stacks else []
                height = stack_height(paras, rect[2], title_floor)[0]
                if height <= rect[3] + 1.0 and not over_lines(paras, ctx.style):
                    break
                trimmed, moved = without_last(block)
                if trimmed is None:
                    break
                block = trimmed
                plan["blocks"][i] = block
                note = {"slide": plan["number"], "blockId": block.get("id"),
                        "reason": NOTES_REASON,
                        "from": {"text": preview(moved)}, "to": {"text": ""}}
                ctx.adjusted.append(note)
                plan.setdefault("textAdjusted", []).append(note)
                plan.setdefault("notes", []).append(moved)
                ctx.warn("block '%s' does not fit at the floor; its last paragraph "
                         "moved to the notes" % block.get("id"))


def draw_blocks(ctx, plan):
    """Draw one planned slide's blocks and return its snapshot entry."""
    ctx.kind = plan["kind"]
    ctx.shapes = plan["shapes"]
    ctx.seq = plan["seq"]
    ctx.where = plan["where"]
    ctx.number = plan["number"]
    slide = plan["slide"]
    slide_spec = plan["spec"]
    ctx.headline = cap_first(furniture_text(slide_spec, "headline"))
    by_block = dict((m["blockId"], m) for m in plan["adjusted"] if m.get("blockId"))

    for block, rect in zip(plan["blocks"], plan["rects"]):
        kind = norm(block.get("type"))
        writer = WRITERS.get(kind)
        if writer is None:
            ctx.warn("block '%s' has unknown type '%s'" % (block.get("id"), block.get("type")))
            continue
        first = len(ctx.shapes)
        writer(ctx, slide, block, rect)
        move = by_block.get(block.get("id"))
        for entry in ctx.shapes[first:]:
            if move:
                entry["adjusted"] = {"from": move["from"], "to": move["to"],
                                     "reason": move["reason"]}
                if move.get("snappedPt"):
                    entry["adjusted"]["snappedPt"] = move["snappedPt"]
            if emphasised(block):
                entry["emphasis"] = True

    if slide_spec.get("notes"):
        add_notes(slide, slide_spec.get("notes"))
    elif ctx.notes_from_story and str(slide_spec.get("story") or "").strip():
        add_notes(slide, str(slide_spec.get("story")).strip())
    # What the slide could not hold goes to the person presenting, never to the
    # bin: the full source line and any paragraph moved off the slide (Q15).
    for extra in plan.get("notes") or []:
        add_notes(slide, extra)

    entry = {
        "index": plan["number"],
        "id": slide_spec.get("id"),
        "name": plan["name"],
        "kind": ctx.kind,
        "headline": cap_first(furniture_text(slide_spec, "headline")),
        "shapes": list(ctx.shapes),
    }
    if slide_spec.get("section"):
        entry["section"] = str(slide_spec.get("section"))
    if plan["ordinal"] and ctx.kind == "divider":
        entry["ordinal"] = int(plan["ordinal"])
    moves = [dict(m) for m in plan["adjusted"]] + [dict(m) for m in plan.get("textAdjusted") or []]
    if moves:
        entry["adjusted"] = moves
    return entry


# -------------------------------------------------------------- template mode


def clear_slides(prs):
    id_list = prs.slides._sldIdLst
    for entry in list(id_list):
        rid = entry.get(qn("r:id"))
        try:
            prs.part.drop_rel(rid)
        except Exception:  # noqa: BLE001 - a missing rel is fine
            pass
        id_list.remove(entry)


def template_fonts(prs):
    """(major, minor) typefaces from the template theme, when readable."""
    try:
        for part in prs.part.package.iter_parts():
            if "theme" not in str(part.partname):
                continue
            blob = part.blob.decode("utf-8", "ignore")
            major = re.search(r"<a:majorFont>.*?typeface=\"([^\"]*)\"", blob, re.S)
            minor = re.search(r"<a:minorFont>.*?typeface=\"([^\"]*)\"", blob, re.S)
            return (major.group(1) if major else None, minor.group(1) if minor else None)
    except Exception as exc:  # noqa: BLE001 - theme reading is best effort
        log("could not read the template theme: %s" % exc)
    return (None, None)


THEME_COLORS = (("dk1", "ink"), ("lt1", "bg"), ("accent1", "accent"),
                ("accent2", "accent2"))


def template_colors(prs):
    """{palette key: hex} from the template theme's colour scheme (Q8)."""
    try:
        for part in prs.part.package.iter_parts():
            if "theme" not in str(part.partname):
                continue
            blob = part.blob.decode("utf-8", "ignore")
            scheme = re.search(r"<a:clrScheme.*?</a:clrScheme>", blob, re.S)
            if not scheme:
                continue
            text = scheme.group(0)
            out = {}
            for tag, key in THEME_COLORS:
                found = re.search(
                    r"<a:%s>(.*?)</a:%s>" % (tag, tag), text, re.S)
                if not found:
                    continue
                value = re.search(r"(?:srgbClr val|lastClr)=\"([0-9A-Fa-f]{6})\"",
                                  found.group(1))
                if value:
                    out[key] = value.group(1).upper()
            return out
    except Exception as exc:  # noqa: BLE001 - theme reading is best effort
        log("could not read the template colours: %s" % exc)
    return {}


def template_body_furniture(prs):
    """The body measure from the layout most content slides will take (Q8)."""
    profiles = layout_profiles(prs)
    for want in ("titleBody", "twoContent", "titleOnly"):
        found = next((p for p in profiles if p["kind"] == want), None)
        if found is not None:
            return template_furniture(found["layout"])
    found = next((p for p in profiles if p["title"] is not None), None)
    return template_furniture(found["layout"]) if found is not None else {}


def template_furniture(layout):
    """Headline (and body) geometry from the layout's placeholders."""
    out = {}
    for shape in layout.placeholders:
        kind = shape.placeholder_format.type
        rect = {"x": shape.left / float(EMU_PER_PT), "y": shape.top / float(EMU_PER_PT),
                "w": shape.width / float(EMU_PER_PT), "h": shape.height / float(EMU_PER_PT)}
        if kind in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE) and "headline" not in out:
            out["headline"] = rect
        elif kind in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT) and "body" not in out:
            out["body"] = rect
    return out


def open_presentation(spec, style, template_path, warnings):
    """A Presentation sized to the spec, from the template when template mode asked for one."""
    if template_path:
        try:
            prs = Presentation(template_path)
        except Exception as exc:  # noqa: BLE001 - fall back to a blank deck
            warnings.append("could not open the template (%s); built blank" % exc)
            prs = Presentation()
            template_path = None
        else:
            clear_slides(prs)
            major, minor = template_fonts(prs)
            fonts = dict(style.tokens.get("fonts") or {})
            if major:
                fonts["display"] = major
            if minor:
                fonts["body"] = minor
            style.tokens["fonts"] = fonts
            style._font_cache = {}
            found = template_body_furniture(prs)
            if found:
                style._furniture = merge(style._furniture, found)
            palette = template_colors(prs)
            if palette:
                # The master's theme wins: a client template comes out in the
                # client's colours, not in house's (Q8).
                merged = dict(style.tokens.get("palette") or {})
                merged.update(palette)
                merged["muted"] = blend(merged.get("ink", "1F2937"),
                                        merged.get("bg", "FFFFFF"), MUTED_BLEND)
                style.tokens["palette"] = merged
                style.palette = merged
                dressed = {"content": {"fill": merged.get("bg", "FFFFFF")}}
                style.tokens["backgrounds"] = dressed
                style.backgrounds = dressed
    else:
        prs = Presentation()

    size = spec.get("slideSize") or {}
    want_w = as_float(size.get("w"), SLIDE_W)
    want_h = as_float(size.get("h"), SLIDE_H)
    have_w = prs.slide_width / float(EMU_PER_PT)
    have_h = prs.slide_height / float(EMU_PER_PT)
    if template_path and (abs(have_w - want_w) > 1 or abs(have_h - want_h) > 1):
        warnings.append("template is %d x %d pt; kept it and spread the grid over it"
                        % (int(have_w), int(have_h)))
        spec = dict(spec)
        spec["slideSize"] = {"w": have_w, "h": have_h}
    else:
        prs.slide_width = emu(want_w)
        prs.slide_height = emu(want_h)
    return prs, spec


# ---------------------------------------------------------------- rendering


def find_tool(env_name, exe, extra_paths):
    override = os.environ.get(env_name)
    if override:
        return override if os.path.exists(override) else None
    found = shutil.which(exe)
    if found:
        return found
    for path in extra_paths:
        if os.path.exists(path):
            return path
    return None


def find_soffice():
    return find_tool("BLUEPRINT_SOFFICE", "soffice", SOFFICE_PATHS)


def find_pdftoppm():
    return find_tool("BLUEPRINT_PDFTOPPM", "pdftoppm", PDFTOPPM_PATHS)


def run_soffice(soffice, args, timeout):
    profile = os.path.join(tempfile.gettempdir(), "blueprint-lo-profile")
    cmd = [soffice, "-env:UserInstallation=file://" + profile, "--headless", "--norestore"]
    cmd.extend(args)
    return subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                          timeout=timeout, check=False)


def convert_pdf(soffice, pptx_path, out_dir, warnings):
    try:
        run_soffice(soffice, ["--convert-to", "pdf", "--outdir", out_dir, pptx_path],
                    SOFFICE_TIMEOUT)
    except subprocess.TimeoutExpired:
        warnings.append("LibreOffice timed out; no PDF written")
        return None
    except Exception as exc:  # noqa: BLE001 - never fail the build over a render
        warnings.append("LibreOffice failed (%s); no PDF written" % exc)
        return None
    base = os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
    made = os.path.join(out_dir, base)
    if not os.path.exists(made):
        warnings.append("LibreOffice produced no PDF")
        return None
    return made


def render_thumbnails(pdf_path, slides_dir, soffice, warnings):
    """Slide PNGs from the PDF: pdftoppm when present, else one page via soffice."""
    if os.path.isdir(slides_dir):
        shutil.rmtree(slides_dir, ignore_errors=True)
    os.makedirs(slides_dir)
    tool = find_pdftoppm()
    if tool:
        try:
            subprocess.run([tool, "-png", "-r", str(THUMB_DPI), pdf_path,
                            os.path.join(slides_dir, "slide")],
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                           timeout=PDFTOPPM_TIMEOUT, check=False)
        except Exception as exc:  # noqa: BLE001 - thumbnails are optional
            warnings.append("pdftoppm failed (%s); no slide PNGs" % exc)
        made = sorted(glob.glob(os.path.join(slides_dir, "slide*.png")))
        if made:
            return made
    try:
        run_soffice(soffice, ["--convert-to", "png", "--outdir", slides_dir, pdf_path],
                    SOFFICE_TIMEOUT)
    except Exception as exc:  # noqa: BLE001 - thumbnails are optional
        warnings.append("slide PNGs skipped (%s)" % exc)
        return []
    made = sorted(glob.glob(os.path.join(slides_dir, "*.png")))
    if made:
        warnings.append("pdftoppm is missing; rendered the first slide only")
    else:
        warnings.append("no PNG renderer available; slide PNGs skipped")
    return made


def snapshot_path(out_path):
    base = str(out_path)
    if base.lower().endswith(".pptx"):
        base = base[:-5]
    return base + ".snapshot.json"


def slides_dir_for(out_path):
    base = str(out_path)
    if base.lower().endswith(".pptx"):
        base = base[:-5]
    return base + ".slides"


# -------------------------------------------------------------------- build


def build(spec):
    warnings = []
    errors = []
    out_path = spec.get("outputPath")
    if not out_path:
        return {"ok": False, "path": None, "warnings": warnings,
                "errors": ["spec has no outputPath"], "slides": []}
    slides_spec = spec.get("slides") or []
    if not slides_spec:
        errors.append("spec has no slides")

    seen = {}
    for i, one in enumerate(slides_spec):
        for block in one.get("blocks") or []:
            bid = block.get("id")
            if not bid:
                errors.append("a block on slide %d has no id" % (i + 1))
            elif bid in seen:
                errors.append("duplicate block id '%s'" % bid)
            else:
                seen[bid] = True
    if errors:
        return {"ok": False, "path": None, "warnings": warnings, "errors": errors, "slides": []}

    style, template_path = make_style(spec, warnings)
    prs, spec = open_presentation(spec, style, template_path, warnings)
    ctx = Ctx(spec, style, warnings)
    slides_spec, cover_added = with_chrome(spec, slides_spec, style, warnings,
                                          ctx.cell_w, ctx.cell_h)
    template_mode = bool(template_path)
    layout_content = pick_layout(prs, template_mode)
    layout_blank = pick_layout(prs, False)
    profiles = layout_profiles(prs) if template_mode else None

    try:
        prs.core_properties.title = str(spec.get("name") or "Deck")
        prs.core_properties.author = "Blueprint"
        brief = spec.get("brief") or {}
        if brief.get("takeaway"):
            prs.core_properties.comments = str(brief.get("takeaway"))
    except Exception as exc:  # noqa: BLE001 - metadata only
        log("could not set core properties: %s" % exc)

    snapshot = {
        "name": spec.get("name"),
        "slideSize": {"w": ctx.slide_w, "h": ctx.slide_h},
        "grid": {"cols": ctx.cols, "rows": ctx.rows, "cell": ctx.cell},
        "style": {"mode": style.mode, "name": style.name},
        "fidelity": ctx.fidelity,
        "slides": [],
    }
    report = []
    ordinals = divider_numbers(slides_spec)

    # One headline size for the whole deck before any furniture is drawn, so the
    # bands the blocks are guarded against are the bands the reader will see.
    apply_type_scale(ctx, choose_headline_type(ctx, slides_spec))

    plans = []
    for i, slide_spec in enumerate(slides_spec):
        kind = slide_kind(slide_spec)
        layout = layout_content if (template_mode and kind == "content") else layout_blank
        plans.append(plan_slide(ctx, prs, layout, slide_spec, i + 1, template_mode,
                                ordinal=ordinals.get(i),
                                page=(i + 1) - (1 if cover_added else 0),
                                profiles=profiles))

    # Then one size per body role, now that every block rectangle is settled.
    apply_type_scale(ctx, choose_block_type(ctx, plans))
    # At the floor, a paragraph that will not fit moves to the notes (Q15).
    move_overflow_to_notes(ctx, plans)

    for plan in plans:
        entry = draw_blocks(ctx, plan)
        snapshot["slides"].append(entry)
        report.append({"index": plan["number"],
                       "title": entry["headline"] or entry["name"],
                       "shapes": len(entry["shapes"])})
    faces = style.faces()
    if ctx.type_scale:
        snapshot["typeScale"] = dict(ctx.type_scale)
        if ctx.type_scale_floored:
            snapshot["typeScaleFloored"] = list(ctx.type_scale_floored)
    snapshot["typeFaces"] = faces

    directory = os.path.dirname(os.path.abspath(str(out_path)))
    if directory and not os.path.isdir(directory):
        os.makedirs(directory)
    prs.save(out_path)
    snap_path = snapshot_path(out_path)
    with open(snap_path, "w") as handle:
        json.dump(snapshot, handle, separators=(",", ":"))

    result = {"ok": True, "path": str(out_path), "snapshotPath": snap_path,
              "warnings": warnings, "errors": [], "slides": report,
              "fidelity": ctx.fidelity}
    scale = dict(ctx.type_scale)
    if ctx.type_scale_floored:
        scale["atFloor"] = list(ctx.type_scale_floored)
    # Which typeface each role really landed on, so a weight that is not
    # installed is visible rather than silently synthesised (Q20).
    scale["faces"] = faces
    result["typeScale"] = scale
    if ctx.adjusted:
        result["adjusted"] = list(ctx.adjusted)

    want_pdf = bool(spec.get("exportPdf"))
    want_thumbs = spec.get("thumbnails", True) is not False
    if want_pdf or want_thumbs:
        soffice = find_soffice()
        if soffice is None:
            warnings.append("LibreOffice not found; no PDF or slide PNGs")
        else:
            work = tempfile.mkdtemp(prefix="blueprint_pdf_")
            try:
                pdf = convert_pdf(soffice, str(out_path), directory if want_pdf else work,
                                  warnings)
                if pdf and want_pdf:
                    result["pdfPath"] = pdf
                if pdf and want_thumbs:
                    thumbs = render_thumbnails(pdf, slides_dir_for(out_path), soffice, warnings)
                    if thumbs:
                        result["thumbnails"] = thumbs
            finally:
                shutil.rmtree(work, ignore_errors=True)

    log("wrote %s (%d slides) and %s" % (out_path, len(report), snap_path))
    return result


def main(argv):
    if len(argv) != 2:
        sys.stdout.write(json.dumps({"ok": False, "errors": ["usage: build_pptx.py <spec.json>"]}))
        return 1
    try:
        with open(argv[1], "r") as handle:
            spec = json.load(handle)
    except Exception as exc:  # noqa: BLE001 - report any read problem as an error
        sys.stdout.write(json.dumps({"ok": False, "errors": ["could not read spec: %s" % exc]}))
        return 1
    try:
        result = build(spec)
    except Exception as exc:  # noqa: BLE001 - one JSON object on stdout, always
        sys.stdout.write(json.dumps({"ok": False,
                                     "errors": ["%s: %s" % (type(exc).__name__, exc)]}))
        return 1
    sys.stdout.write(json.dumps(result))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))

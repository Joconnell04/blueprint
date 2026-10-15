#!/usr/bin/env python3
"""Workbook craft check after build: the layer above check_formulas.py.

Usage:
    python3 builder/check_workbook.py <workbook.xlsx> <snapshot.json> [<spec.json>] [--out <findings.json>]

`check_formulas.py` covers the formula layer: missing sheets, references past
the used area, empty references, #REF!, cycles and the seven Excel error values
after a recalculation. It answers "do the formulas compute". This module answers
the question above it, "is this a workbook a reader can use", and reports through
the same finding shape so the built view can list both.

Every kind comes from `docs/research/workbook-quality.md`, section "Spreadsheet
craft rules worth enforcing, and where they come from", which traces each rule
to ICAEW's Twenty Principles, the FAST standard, Tufte, Few, Panko or the house
`spreadsheet-format` skill.

| kind | code(s) | severity | what it means |
|---|---|---|---|
| `clipped` | `clipped` | warning | a cell's formatted text is wider than the space it has |
| `note-clipped` | `note-clipped` | warning | the same, for a note cell: the annotation is cut |
| `mixed-format` | `mixed-format` | warning | one column of a period table formats the same quantity two ways |
| `hardcode` | `hardcode` | warning | a numeric literal that is not a unit scalar sits inside a formula |
| `row-inconsistent` | `row-inconsistent` | warning | a formula's R1C1 shape differs from its row neighbours across a period series |
| `unused-input` | `unused-input` | warning | an input cell no formula reads |
| `dangling-output` | `dangling-output` | warning | an output cell whose dependencies reach no input |
| `ref-colour` | `circular`, `ref-error` | error | a reference cycle, a #REF!, a missing sheet or an external workbook |
| `ref-colour` | `off-sheet-uncoloured`, `link-uncoloured` | warning | a cross-sheet pass-through or a hyperlink drawn in the plain formula colour |
| `no-unit` | `no-unit` | warning | a numeric column whose format and header both carry no unit |
| `empty-band` | `empty-band` | warning | a run of empty rows inside the used range |
| `no-freeze` | `no-freeze` | warning | a table sheet with no freeze pane, or one that holds no columns |
| `no-cover` | `no-cover` | warning | no Read Me tab, or a Read Me that does not list every tab |
| `total-mismatch` | `total-mismatch` | error | a total row's value does not equal the rows above it |
| `magic-density` | `magic-density` | warning | too many of a sheet's formulas carry a hardcoded number |

Severity is `error` for reconciliation and #REF-class problems, `warning` for
everything else, and every finding carries a one-line `fix`.

Measurement. Nothing here is a copy of a builder constant; build_xlsx is
imported and asked. The text a reader sees comes from `build_xlsx.display`, the
builder's own number formatter. A cell is clipped when that text is longer, in
characters, than the space it has: its own column, plus its merge span, plus the
empty columns it spills over, which numbers never do. A wrapped cell and a
number under `General` are never reported: the first grows down instead of
across, and Excel shortens the second to fit rather than clipping it. The width
the column
should have instead comes from `build_xlsx.Ctx.fit` and `Ctx.final_widths`,
carrying the builder's padding, floors and caps, so the fix line names the
builder's own answer. Widths are in character units, the unit Excel stores them
in. The style (fonts, the input, formula and link colours, label and value
widths) comes from `build_xlsx.make_styler` over the spec when a spec is given,
so the check uses the style the workbook was built with.

Scope. Block geometry, block types and cell kinds come from the snapshot; checks
that need a block quietly skip a sheet the snapshot does not describe. The spec
is optional and only adds `totalRow`, `rowNotes` and the style.

Output (stdout, and the --out path when given):

    { "ok": true, "sheets": n, "cells": n, "formulas": n,
      "findings": [ { "sheet", "cell", "blockId"?, "code", "severity", "message",
                      "kind", "fix", "scope"? } ],
      "summary": { "errors": n, "warnings": n },
      "byKind": { "clipped": n, ... } }

A sheet-wide finding carries `"cell": ""` and `"scope": "sheet"`; the
workbook-wide one carries `"sheet": ""`, `"cell": ""` and `"scope": "workbook"`.

Exit status is 0 whenever the check ran, whatever it found. Exit 1 only when the
checker itself crashed; stdout then carries {"ok": false, "error": ...}.

Python 3.9, openpyxl only.
"""

import json
import os
import re
import sys

import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils import get_column_letter, column_index_from_string

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import build_xlsx                      # noqa: E402  the width model and the style
import check_formulas as cf            # noqa: E402  the workbook loader and ref helpers

# ------------------------------------------------------------------ constants

# Kinds in report order, so the built view can group without sorting.
KIND_ORDER = (
    "total-mismatch",
    "ref-colour",
    "row-inconsistent",
    "hardcode",
    "magic-density",
    "mixed-format",
    "no-unit",
    "clipped",
    "note-clipped",
    "unused-input",
    "dangling-output",
    "empty-band",
    "no-freeze",
    "no-cover",
)

CODE_KIND = {
    "clipped": "clipped",
    "note-clipped": "note-clipped",
    "mixed-format": "mixed-format",
    "hardcode": "hardcode",
    "row-inconsistent": "row-inconsistent",
    "unused-input": "unused-input",
    "dangling-output": "dangling-output",
    "circular": "ref-colour",
    "ref-error": "ref-colour",
    "off-sheet-uncoloured": "ref-colour",
    "link-uncoloured": "ref-colour",
    "no-unit": "no-unit",
    "empty-band": "empty-band",
    "no-freeze": "no-freeze",
    "no-cover": "no-cover",
    "total-mismatch": "total-mismatch",
    "magic-density": "magic-density",
}

# Only reconciliation and the #REF class are errors; craft problems are warnings.
ERROR_CODES = {"total-mismatch", "circular", "ref-error"}

# FAST and ICAEW principle 14: a literal inside a formula is a defect unless it
# is a unit scalar. The core five are 0, 1, 100, 12 and 1000; the research's own
# list adds the three that turn a number into its own unit (a million, a
# billion, weeks in a year), without which every `/1000000` that writes dollars
# as $M reads as a hardcode.
UNIT_SCALARS = (0.0, 1.0, 12.0, 52.0, 100.0, 1000.0, 1e6, 1e9)

# Thresholds, named so a style or a fidelity setting can move them later.
EMPTY_BAND_ROWS = 6          # a run of this many empty rows inside the used range
FREEZE_MIN_ROWS = 15         # a table sheet taller than this wants a freeze pane
MAGIC_DENSITY = 0.25         # share of a sheet's formulas carrying a hardcode
MAGIC_MIN_FORMULAS = 5       # below this a sheet is too small to have a density
ROW_MIN_FORMULAS = 3         # a period row needs this many formulas to have a shape
PERIOD_MIN_COLS = 2          # a header row needs this many periods to be a series
TOTAL_REL_TOLERANCE = 1e-6   # relative slack when reconciling a total
MAX_SPILL_COLS = 12          # how far text is allowed to spill over empty columns
DEFAULT_COL_WIDTH = 8.43     # Excel's own default when the sheet sets none

CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
ABS_CELL_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")
TOTAL_LABEL_RE = re.compile(r"^\s*(?:grand\s+)?(?:total|subtotal|sum)\b", re.I)
EXTERNAL_REF_RE = re.compile(r"^\s*\[")

# Period series. A header row parses when PERIOD_MIN_COLS of its cells are the
# same kind of period and at most one other cell is not a period at all.
YEAR_RE = re.compile(r"^(?:FY|CY)?\s*((?:19|20)\d{2})\s*(?:E|A|F|P|B|e|a|f|p|b)?$")
SHORT_YEAR_RE = re.compile(r"^(?:FY|CY)\s*'?(\d{2})\s*(?:E|A|F|P|B)?$", re.I)
QUARTER_RE = re.compile(
    r"^(?:Q([1-4])|([1-4])Q)\s*(?:[-/ ]?\s*(?:FY|CY)?\s*'?((?:19|20)?\d{2}))?"
    r"\s*(?:E|A|F|P|B)?$", re.I)
ISO_MONTH_RE = re.compile(r"^((?:19|20)\d{2})[-/](0?[1-9]|1[0-2])$")
MONTH_NAMES = (
    "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec",
    "january", "february", "march", "april", "june", "july", "august",
    "september", "october", "november", "december",
)
MONTH_RE = re.compile(
    r"^(%s)\.?\s*(?:[-/ ]?\s*'?((?:19|20)?\d{2}))?$" % "|".join(MONTH_NAMES), re.I)

# A unit in a header or a row label. Symbols first, then the words that name a
# unit rather than a quantity. Kept short on purpose: a longer list turns
# `no-unit` into noise.
UNIT_SYMBOL_RE = re.compile(r"[$€£%#]|\bx\b|\(\s*x\s*\)|\bbps\b", re.I)
UNIT_WORDS = (
    "count", "counts", "number", "no.", "qty", "quantity", "units", "unit",
    "seats", "seat", "accounts", "account", "customers", "employees", "people",
    "headcount", "fte", "hours", "hour", "days", "day", "weeks", "months",
    "month", "years", "year", "ratio", "multiple", "index", "score", "rate",
    "share", "mix", "margin", "per", "each", "bn", "mm", "000s",
)
UNIT_WORD_RE = re.compile(
    r"(?<![a-z])(%s)(?![a-z])" % "|".join(re.escape(w) for w in UNIT_WORDS), re.I)

FIXES = {
    "clipped": "Widen the column to %s characters, or wrap the cell and set a row height.",
    "note-clipped": "Merge the note across to the sheet's right edge, or shorten it to one line.",
    "mixed-format": "Give the whole column one format, or move the odd row to its own rowFormat.",
    "hardcode": "Move the literal into an inputs row and reference it, or add a note saying why.",
    "row-inconsistent": "Copy one formula across the period row, or split the odd cell out.",
    "unused-input": "Reference the input from a formula, or drop the row.",
    "dangling-output": "Point the output at an input chain, or colour its literals as inputs.",
    "circular": "Break the cycle: one of these cells needs a literal or a different reference.",
    "ref-error": "Repoint the reference at a cell that exists.",
    "off-sheet-uncoloured": "Colour the cross-sheet pass-through, or name the range it reads.",
    "link-uncoloured": "Draw the hyperlink in link blue with an underline, never in input blue.",
    "no-unit": "Put the unit in the header, or use a format that carries it.",
    "empty-band": "Delete the band, or move the block below it up.",
    "no-freeze": "Freeze at the first data cell of the sheet's widest block.",
    "no-cover": "Add a Read Me tab with a Contents row per tab.",
    "total-mismatch": "Repoint the total at the whole column, or fix the rows it misses.",
    "magic-density": "Move the literals into an inputs block and reference them.",
}


# -------------------------------------------------------------------- helpers


def a1(row, col):
    return "%s%d" % (get_column_letter(col), row)


def norm_colour(value):
    """An openpyxl font colour as six upper-case hex digits, or None.

    openpyxl reports ARGB ('00595959'), theme colours and indexed colours; only
    an explicit RGB is comparable with the style's own hex.
    """
    if value is None:
        return None
    rgb = getattr(value, "rgb", None)
    if not isinstance(rgb, str):
        return None
    text = rgb.upper()
    if len(text) == 8:
        text = text[2:]
    if len(text) != 6:
        return None
    return text


def cell_colour(cell):
    return norm_colour(getattr(cell.font, "color", None))


def period_tag(value):
    """'year', 'quarter' or 'month' when this header cell names a period, else None."""
    if value is None:
        return None
    if hasattr(value, "year") and hasattr(value, "month"):
        return "month"
    if isinstance(value, bool):
        return None
    if isinstance(value, float) and value == int(value):
        value = int(value)
    text = str(value).strip()
    if not text:
        return None
    if YEAR_RE.match(text) or SHORT_YEAR_RE.match(text):
        return "year"
    if QUARTER_RE.match(text):
        return "quarter"
    if ISO_MONTH_RE.match(text) or MONTH_RE.match(text):
        return "month"
    return None


def has_unit(text):
    if text is None:
        return False
    text = str(text)
    return bool(UNIT_SYMBOL_RE.search(text) or UNIT_WORD_RE.search(text))


def format_family(nf):
    """The quantity a number format describes, so two spellings of one quantity match."""
    if nf in (None, "", "General"):
        return "number"
    if "%" in nf:
        return "percent"
    if "$" in nf or "€" in nf or "£" in nf or "[$" in nf:
        return "currency"
    if '"x"' in nf or nf.rstrip('"').endswith("x"):
        return "multiple"
    if nf == "@":
        return "text"
    if re.search(r"[#0]", nf) is None and re.search(r"[ymdhs]", nf, re.I):
        return "date"
    return "number"


def is_bare_format(nf):
    """A format that carries no unit: General, text, or digits and separators only."""
    family = format_family(nf)
    return family in ("number", "text")


def to_r1c1(formula, row, col):
    """A formula's shape: every A1 reference rewritten relative to its own cell.

    Literals are kept verbatim, so `=C10*(1+$C$5)^1` and `...^2` come out
    different. That is deliberate: the `^1` to `^5` ramp is the copy
    inconsistency the check exists to find.
    """
    try:
        items = Tokenizer(formula).items
    except TokenizerError:
        return formula
    out = ["="] if formula.lstrip().startswith("=") else []
    for tok in items:
        if tok.type == "OPERAND" and tok.subtype == "RANGE":
            out.append(_ref_r1c1(tok.value, row, col))
        else:
            out.append(tok.value)
    return "".join(out)


def _ref_r1c1(raw, row, col):
    sheet, rest = cf.split_sheet(raw)
    parts = rest.split(":")
    converted = []
    for part in parts:
        m = ABS_CELL_RE.match(part)
        if not m:
            return raw           # a defined name, a whole column, or #REF!: leave it
        abs_col, letters, abs_row, digits = m.groups()
        r, c = int(digits), column_index_from_string(letters.upper())
        r_txt = "R%d" % r if abs_row else "R[%d]" % (r - row)
        c_txt = "C%d" % c if abs_col else "C[%d]" % (c - col)
        converted.append(r_txt + c_txt)
    prefix = ""
    if sheet:
        prefix = "'%s'!" % sheet
    return prefix + ":".join(converted)


def formula_refs(formula):
    """Every reference token in a formula as (raw, sheet or None, area or None)."""
    try:
        items = Tokenizer(formula).items
    except TokenizerError:
        return []
    out = []
    for tok in items:
        if tok.type != "OPERAND" or tok.subtype != "RANGE":
            continue
        sheet, rest = cf.split_sheet(tok.value)
        out.append((tok.value, sheet, cf.parse_area(rest.replace("$", ""))))
    return out


def formula_numbers(formula):
    """Every numeric literal in a formula, as floats."""
    try:
        items = Tokenizer(formula).items
    except TokenizerError:
        return []
    out = []
    for tok in items:
        if tok.type == "OPERAND" and tok.subtype == "NUMBER":
            try:
                out.append(float(tok.value))
            except ValueError:
                continue
    return out


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


# ------------------------------------------------------------- width model


class WidthModel:
    """The builder's fit model, driven through build_xlsx so nothing is copied.

    `need` is the width build_xlsx would ask for, `capped` is that width after
    its floor and 60-character cap, and `bold_allowance` is the model's own
    bold factor, measured on a probe rather than written down here.
    """

    def __init__(self, styler=None):
        self.styler = styler or build_xlsx.HouseStyler()

    def _ctx(self):
        return build_xlsx.Ctx(self.styler, {}, [], [])

    def need(self, value, fmt=None, bold=False, is_formula=False):
        ctx = self._ctx()
        ctx.fit(1, value, fmt, is_formula, bold)
        return ctx.fits.get(1, 0)

    def capped(self, value, fmt=None, bold=False, is_formula=False):
        ctx = self._ctx()
        ctx.fit(1, value, fmt, is_formula, bold)
        return ctx.final_widths().get(1, 0)

    def cap(self):
        """The model's own width cap, read by asking it for something absurd."""
        return self.capped("x" * 4000)

    def bold_allowance(self):
        """How much wider the model makes bold text. 1.0 when it wraps it instead."""
        probe = "x" * 1000
        plain = float(self.need(probe))
        bold = float(self.need(probe, bold=True))
        if plain <= 0 or bold <= 0:
            return 1.0
        return max(1.0, bold / plain)

    def display(self, value, nf):
        """The text a reader sees, from the builder's own formatter."""
        shown = build_xlsx.display(value, nf)
        return shown


# ------------------------------------------------------------------ snapshot


class Block:
    __slots__ = ("id", "type", "r1", "c1", "r2", "c2", "rows", "cols")

    def __init__(self, raw):
        self.id = raw.get("id")
        self.type = str(raw.get("type") or "")
        self.r1 = int(raw["row"])
        self.c1 = int(raw["col"])
        self.rows = int(raw.get("rows", 1))
        self.cols = int(raw.get("cols", 1))
        self.r2 = self.r1 + self.rows - 1
        self.c2 = self.c1 + self.cols - 1

    def holds(self, row, col):
        return self.r1 <= row <= self.r2 and self.c1 <= col <= self.c2


class Snapshot:
    """Blocks, cell kinds and widths per sheet, from the build snapshot."""

    def __init__(self, raw):
        self.raw = raw or {}
        self.blocks = {}      # sheet -> [Block]
        self.kinds = {}       # sheet -> {(row, col): kind}
        self.widths = {}      # sheet -> {col: width}
        for sheet in self.raw.get("sheets") or []:
            name = sheet.get("name")
            blocks = []
            for raw_block in sheet.get("blocks") or []:
                try:
                    blocks.append(Block(raw_block))
                except (KeyError, TypeError, ValueError):
                    continue
            self.blocks[name] = blocks
            kinds = {}
            for entry in sheet.get("cells") or []:
                m = CELL_RE.match(str(entry.get("a") or ""))
                if not m:
                    continue
                key = (int(m.group(2)), column_index_from_string(m.group(1).upper()))
                kinds[key] = entry.get("k")
            self.kinds[name] = kinds
            widths = {}
            for letter, width in (sheet.get("widths") or {}).items():
                try:
                    widths[column_index_from_string(str(letter).upper())] = float(width)
                except (ValueError, TypeError):
                    continue
            self.widths[name] = widths

    @property
    def empty(self):
        return not self.blocks

    def block_at(self, sheet, row, col):
        for block in self.blocks.get(sheet, []):
            if block.holds(row, col):
                return block
        return None

    def kind_at(self, sheet, row, col):
        return self.kinds.get(sheet, {}).get((row, col))


def load_json(path):
    with open(path, "r") as handle:
        return json.load(handle)


# --------------------------------------------------------------------- sheets


class SheetModel:
    """One sheet as the checker needs it: values, formulas, widths, merges, styles."""

    def __init__(self, book, ws, snapshot):
        self.name = ws.title
        self.ws = ws
        self.values = book.values[self.name]
        self.formulas = book.formulas[self.name]
        self.max_row, self.max_col = book.used[self.name]
        self.snap_widths = snapshot.widths.get(self.name, {})
        self.merges = []
        self.merge_of = {}
        for rng in ws.merged_cells.ranges:
            box = (rng.min_row, rng.min_col, rng.max_row, rng.max_col)
            self.merges.append(box)
            for r in range(box[0], box[2] + 1):
                for c in range(box[1], box[3] + 1):
                    self.merge_of[(r, c)] = box
        self.blocks = snapshot.blocks.get(self.name, [])
        self._kinds = snapshot.kinds.get(self.name, {})
        self._widths = {}

    # ------------------------------------------------------------- geometry

    def width_of(self, col):
        if col in self._widths:
            return self._widths[col]
        width = None
        dim = self.ws.column_dimensions.get(get_column_letter(col))
        if dim is not None and dim.width:
            width = float(dim.width)
        if width is None:
            width = self.snap_widths.get(col)
        if width is None:
            fmt = self.ws.sheet_format
            width = getattr(fmt, "defaultColWidth", None) or DEFAULT_COL_WIDTH
            width = float(width)
        self._widths[col] = width
        return width

    def is_empty(self, row, col):
        key = (row, col)
        if key in self.values or key in self.formulas:
            return False
        return key not in self.merge_of

    def row_is_empty(self, row):
        for col in range(1, self.max_col + 1):
            if (row, col) in self.values or (row, col) in self.formulas:
                return False
        return True

    def cell(self, row, col):
        return self.ws.cell(row, col)

    def kind(self, row, col):
        return self._kinds.get((row, col))

    def block_at(self, row, col):
        for block in self.blocks:
            if block.holds(row, col):
                return block
        return None

    def written(self):
        """Every written cell, row-major: (row, col, value, is_formula)."""
        keys = set(self.values) | set(self.formulas)
        for key in sorted(keys):
            row, col = key
            if key in self.formulas:
                yield row, col, self.formulas[key], True
            else:
                yield row, col, self.values[key], False

    # -------------------------------------------------------------- content

    def header_periods(self, block):
        """{col: tag} for a table block whose header row is a period series, else {}.

        The header row is the block's second row: build_xlsx puts the title on
        the block's own row and the column headers under it.
        """
        if block.type not in ("table", "calc"):
            return {}
        header_row = block.r1 + 1
        if header_row > block.r2:
            return {}
        cells = []
        for col in range(block.c1, block.c2 + 1):
            value = self.values.get((header_row, col))
            if value is None and (header_row, col) in self.formulas:
                value = None
            cells.append((col, value))
        if not cells:
            return {}
        _label_col, label_value = cells[0]
        if period_tag(label_value) is not None:
            return {}        # periods down the label column is a different shape
        rest = [(c, v) for c, v in cells[1:] if v not in (None, "")]
        if len(rest) < PERIOD_MIN_COLS:
            return {}
        tags = {}
        for col, value in rest:
            tag = period_tag(value)
            if tag:
                tags.setdefault(tag, []).append(col)
        if not tags:
            return {}
        tag = max(tags, key=lambda t: len(tags[t]))
        cols = tags[tag]
        if len(cols) < PERIOD_MIN_COLS or len(cols) < len(rest) - 1:
            return {}
        return dict((col, tag) for col in cols)

    def data_rows(self, block):
        """The block's data rows: below the header, above any total row."""
        if block.type in ("table", "calc"):
            first = block.r1 + 2
        elif block.type in ("inputs", "output"):
            first = block.r1 + 1
        else:
            return []
        return [r for r in range(first, block.r2 + 1)]

    def total_row_of(self, block, spec_block):
        """The block's total row, from the snapshot kind, the spec, or the label."""
        for row in range(block.r2, block.r1, -1):
            for col in range(block.c1, block.c2 + 1):
                if self.kind(row, col) == "total":
                    return row
        if spec_block and isinstance(spec_block.get("totalRow"), list):
            data = spec_block.get("data") or []
            row = block.r1 + 2 + len(data)
            if row <= block.r2:
                return row
        for row in range(block.r2, block.r1 + 1, -1):
            label = self.values.get((row, block.c1))
            if isinstance(label, str) and TOTAL_LABEL_RE.match(label):
                return row
        return None


# -------------------------------------------------------------------- checker


class Checker:
    def __init__(self, xlsx_path, snapshot_raw=None, spec=None):
        self.path = xlsx_path
        self.book = cf.Book(xlsx_path)
        self.wb = self.book.wb
        self.snapshot = Snapshot(snapshot_raw)
        self.spec = spec or {}
        self.notes = []
        self.styler = self._styler()
        self.model = WidthModel(self.styler)
        self.cached = self._cached_values()
        self.sheets = []
        for ws in self.wb.worksheets:
            self.sheets.append(SheetModel(self.book, ws, self.snapshot))
        self.by_name = dict((s.name, s) for s in self.sheets)
        self.spec_blocks = self._spec_blocks()
        self.findings = []
        self.seen = set()
        self.refs = {}            # (sheet, row, col) -> set of referenced cells
        self.referenced = set()   # every cell any formula reads
        self.hardcodes = {}       # sheet -> count
        self.formula_counts = {}  # sheet -> count

    # ------------------------------------------------------------ set-up

    def _styler(self):
        try:
            return build_xlsx.make_styler(self.spec, self.notes)
        except Exception as exc:                      # noqa: BLE001 - style is advisory
            self.notes.append("could not build the spec's styler (%s); used house defaults"
                              % exc)
            return build_xlsx.HouseStyler()

    def _cached_values(self):
        """Cached formula results, when the file carries any (Excel or LibreOffice saved it)."""
        out = {}
        try:
            wb = openpyxl.load_workbook(self.path, data_only=True)
        except Exception:                             # noqa: BLE001 - cached values are optional
            return out
        for ws in wb.worksheets:
            sheet = {}
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None and cell.value != "":
                        sheet[(cell.row, cell.column)] = cell.value
            out[ws.title] = sheet
        return out

    def _spec_blocks(self):
        out = {}
        for sheet in self.spec.get("sheets") or []:
            for block in sheet.get("blocks") or []:
                if block.get("id"):
                    out[block["id"]] = block
        return out

    def spec_block(self, block):
        if block is None or not block.id:
            return None
        return self.spec_blocks.get(block.id)

    # ----------------------------------------------------------- findings

    def add(self, sheet, row, col, code, message, block=None, scope=None, fix=None):
        cell = a1(row, col) if row and col else ""
        key = (sheet, cell, code, message)
        if key in self.seen:
            return
        self.seen.add(key)
        entry = {
            "sheet": sheet,
            "cell": cell,
            "code": code,
            "kind": CODE_KIND.get(code, code),
            "severity": "error" if code in ERROR_CODES else "warning",
            "message": message,
            "fix": fix or FIXES.get(code, ""),
        }
        block_id = None
        if block is not None:
            block_id = block.id
        elif sheet and row and col:
            found = self.snapshot.block_at(sheet, row, col)
            block_id = found.id if found else None
        if block_id:
            entry["blockId"] = block_id
        if scope:
            entry["scope"] = scope
        self.findings.append(entry)

    # ---------------------------------------------------------------- run

    def run(self):
        self.index_references()
        for sheet in self.sheets:
            self.check_clipped(sheet)
            self.check_mixed_formats(sheet)
            self.check_hardcodes(sheet)
            self.check_row_consistency(sheet)
            self.check_units(sheet)
            self.check_totals(sheet)
            self.check_empty_band(sheet)
            self.check_freeze(sheet)
            self.check_magic_density(sheet)
        self.check_references()
        self.check_inputs_and_outputs()
        self.check_cover()
        return self.result()

    def result(self):
        order = dict((kind, i) for i, kind in enumerate(KIND_ORDER))
        sheet_order = dict((s.name, i) for i, s in enumerate(self.sheets))

        def sort_key(f):
            m = CELL_RE.match(f["cell"]) if f["cell"] else None
            row = int(m.group(2)) if m else 0
            col = column_index_from_string(m.group(1).upper()) if m else 0
            return (sheet_order.get(f["sheet"], 99), row, col,
                    0 if f["severity"] == "error" else 1,
                    order.get(f["kind"], 99), f["code"])

        findings = sorted(self.findings, key=sort_key)
        errors = sum(1 for f in findings if f["severity"] == "error")
        by_kind = {}
        for f in findings:
            by_kind[f["kind"]] = by_kind.get(f["kind"], 0) + 1
        result = {
            "ok": True,
            "sheets": len(self.sheets),
            "cells": sum(len(s.values) + len(s.formulas) for s in self.sheets),
            "formulas": sum(len(s.formulas) for s in self.sheets),
            "findings": findings,
            "summary": {"errors": errors, "warnings": len(findings) - errors},
            "byKind": by_kind,
        }
        if self.notes:
            result["note"] = "; ".join(self.notes)
        return result

    # ------------------------------------------------------- reference index

    def index_references(self):
        """Every cell each formula reads, areas clamped to the target's used range."""
        for sheet in self.sheets:
            for key, formula in sheet.formulas.items():
                out = set()
                for raw, ref_sheet, area in formula_refs(formula):
                    target = sheet.name
                    if ref_sheet:
                        resolved = self.book.resolve_sheet(ref_sheet)
                        if resolved is None:
                            continue
                        target = resolved
                    if area is None:
                        named = self.book.names.get(raw.upper())
                        if not named:
                            continue
                        target, area = named
                        resolved = self.book.resolve_sheet(target) if target else None
                        target = resolved or sheet.name
                    model = self.by_name.get(target)
                    if model is None:
                        continue
                    r1, c1, r2, c2 = area
                    r2 = min(r2, max(model.max_row, r1))
                    c2 = min(c2, max(model.max_col, c1))
                    for r in range(r1, r2 + 1):
                        for c in range(c1, c2 + 1):
                            out.add((target, r, c))
                self.refs[(sheet.name, key[0], key[1])] = out
                self.referenced |= out

    # ------------------------------------------------------------- clipped

    def check_clipped(self, sheet):
        cap = self.model.cap()
        for row, col, value, is_formula in sheet.written():
            box = sheet.merge_of.get((row, col))
            if box and (row, col) != (box[0], box[1]):
                continue                     # a covered cell shows nothing of its own
            cell = sheet.cell(row, col)
            if cell.alignment is not None and cell.alignment.wrap_text:
                continue                     # a wrapped cell grows down, not across
            shown, numeric = self.shown_text(sheet, row, col, value, is_formula)
            if not shown:
                continue
            available = self.available_width(sheet, row, col, numeric, cell)
            if len(shown) <= available:
                continue
            kind = sheet.kind(row, col)
            code = "note-clipped" if kind == "note" else "clipped"
            want = self.model.capped(value, self.fmt_hint(sheet, row, col),
                                     bold=bool(cell.font is not None and cell.font.bold))
            preview = shown if len(shown) <= 42 else shown[:39] + "..."
            message = "'%s' needs %d characters and the cell shows %d" % (
                preview, len(shown), int(available))
            fix = FIXES[code]
            if code == "clipped":
                fix = (FIXES["clipped"] % want) if want < cap else (
                    "Wrap the cell and set a row height: it is past the %d character cap." % cap)
            self.add(sheet.name, row, col, code, message, fix=fix)

    def shown_text(self, sheet, row, col, value, is_formula):
        """(text the reader sees, is it a number). Formula cells need a cached value."""
        if is_formula:
            cached = self.cached.get(sheet.name, {}).get((row, col))
            if cached is None or isinstance(cached, str) and cached.startswith("="):
                return None, False
            value = cached
        if is_number(value):
            nf = sheet.cell(row, col).number_format
            shown = self.model.display(value, nf)
            if shown is None:
                return None, True
            return shown, True
        if isinstance(value, str):
            return value, False
        return None, False

    def fmt_hint(self, sheet, row, col):
        """The builder's format name for this cell, so the model measures it the same way."""
        nf = sheet.cell(row, col).number_format
        for name, spelling in (self.styler.formats or {}).items():
            if spelling == nf:
                return name
        return None

    def available_width(self, sheet, row, col, numeric, cell):
        box = sheet.merge_of.get((row, col))
        cols = [col]
        if box:
            cols = list(range(box[1], box[3] + 1))
        width = sum(sheet.width_of(c) for c in cols)
        if numeric:
            return width                     # a number that does not fit becomes ####
        horizontal = (cell.alignment.horizontal if cell.alignment is not None else None)
        right = cols[-1]
        left = cols[0]
        if horizontal in (None, "general", "left", "justify", "center", "centerContinuous"):
            c = right + 1
            while c <= right + MAX_SPILL_COLS and sheet.is_empty(row, c):
                width += sheet.width_of(c)
                c += 1
        if horizontal in ("right", "center", "centerContinuous"):
            c = left - 1
            while c >= 1 and c >= left - MAX_SPILL_COLS and sheet.is_empty(row, c):
                width += sheet.width_of(c)
                c -= 1
        return width

    # -------------------------------------------------------- mixed formats

    def check_mixed_formats(self, sheet):
        """One column of a period table formatting the same quantity two ways.

        Restricted to formats of the same family: a period table's column spans
        several quantities by design, so currency beside percent is correct and
        only `$#,##0` beside `$#,##0.0` is the defect.
        """
        for block in sheet.blocks:
            periods = sheet.header_periods(block)
            if not periods:
                continue
            spec_block = self.spec_block(block)
            total_row = sheet.total_row_of(block, spec_block)
            rows = [r for r in sheet.data_rows(block) if r != total_row]
            cols = sorted(periods)
            grid = {}                    # (row, col) -> number format, numbers only
            for row in rows:
                for col in cols:
                    key = (row, col)
                    if key in sheet.formulas:
                        grid[key] = sheet.cell(row, col).number_format
                    elif is_number(sheet.values.get(key)):
                        grid[key] = sheet.cell(row, col).number_format
            # The block's own convention decides which spelling is the odd one:
            # the format its family uses most across the whole period grid.
            tally = {}
            for (row, _col), nf in grid.items():
                tally.setdefault(format_family(nf), {}).setdefault(nf, []).append(row)
            dominant = {}
            for family, spellings in tally.items():
                ranked = sorted(spellings.items(), key=lambda kv: (-len(kv[1]), min(kv[1])))
                dominant[family] = ranked[0][0]
            for col in cols:
                families = {}
                for row in rows:
                    nf = grid.get((row, col))
                    if nf is None:
                        continue
                    families.setdefault(format_family(nf), {}).setdefault(nf, []).append(row)
                for family, spellings in sorted(families.items()):
                    if len(spellings) < 2:
                        continue        # one spelling in this column: nothing mixed
                    common = dominant.get(family)
                    for nf, hits in sorted(spellings.items()):
                        if nf == common:
                            continue
                        self.add(sheet.name, hits[0], col, "mixed-format",
                                 "%s is %s while the table formats %s as %s"
                                 % (a1(hits[0], col), nf or "General", family,
                                    common or "General"),
                                 block=block)

    # ------------------------------------------------------------ hardcodes

    def check_hardcodes(self, sheet):
        allowed = set(UNIT_SCALARS)
        count = 0
        for key, formula in sorted(sheet.formulas.items()):
            row, col = key
            self.formula_counts[sheet.name] = self.formula_counts.get(sheet.name, 0) + 1
            block = sheet.block_at(row, col)
            if block is not None and block.type not in ("calc", "output", "table"):
                continue
            literals = [n for n in formula_numbers(formula) if abs(n) not in allowed]
            if not literals:
                continue
            if self.row_has_note(sheet, row, block):
                continue                     # the research waives a noted row
            count += 1
            shown = ", ".join(_trim_number(n) for n in literals[:3])
            self.add(sheet.name, row, col, "hardcode",
                     "the formula carries the literal %s" % shown, block=block)
        self.hardcodes[sheet.name] = count

    def row_has_note(self, sheet, row, block):
        if block is not None:
            for col in range(block.c1, block.c2 + 2):
                if sheet.kind(row, col) == "note":
                    return True
            spec_block = self.spec_block(block)
            if spec_block:
                notes = spec_block.get("rowNotes")
                data_rows = sheet.data_rows(block)
                if isinstance(notes, list) and row in data_rows:
                    i = data_rows.index(row)
                    if i < len(notes) and notes[i]:
                        return True
        else:
            for col in range(1, sheet.max_col + 1):
                if sheet.kind(row, col) == "note":
                    return True
        return False

    def check_magic_density(self, sheet):
        formulas = self.formula_counts.get(sheet.name, 0)
        hits = self.hardcodes.get(sheet.name, 0)
        if formulas < MAGIC_MIN_FORMULAS or not hits:
            return
        density = hits / float(formulas)
        if density <= MAGIC_DENSITY:
            return
        self.add(sheet.name, 0, 0, "magic-density",
                 "%d of %d formulas on this sheet carry a hardcoded number (%d%%)"
                 % (hits, formulas, int(round(density * 100))),
                 scope="sheet")

    # ----------------------------------------------------- row consistency

    def check_row_consistency(self, sheet):
        """A formula whose shape differs from its neighbours across a period series.

        Scoped to blocks whose header row parses as a period series, which is
        what keeps it quiet: run over every row of a spec sheet and the same
        rule reports every table whose columns are different quantities.
        """
        for block in sheet.blocks:
            periods = sheet.header_periods(block)
            if len(periods) < ROW_MIN_FORMULAS:
                continue
            cols = sorted(periods)
            for row in sheet.data_rows(block):
                shape_of = {}
                for col in cols:
                    formula = sheet.formulas.get((row, col))
                    if formula:
                        shape_of[col] = to_r1c1(formula, row, col)
                written = [c for c in cols if c in shape_of]
                if len(written) < ROW_MIN_FORMULAS:
                    continue
                if self.is_base_column(cols, written, shape_of):
                    continue
                shapes = {}
                for col in written:
                    shapes.setdefault(shape_of[col], []).append(col)
                total = len(written)
                if len(shapes) < 2:
                    continue
                ranked = sorted(shapes.items(), key=lambda kv: (-len(kv[1]), kv[1][0]))
                if len(ranked[0][1]) * 2 > total:
                    for _shape, hits in ranked[1:]:
                        for col in hits:
                            self.add(sheet.name, row, col, "row-inconsistent",
                                     "%s does not copy the shape of the other %d cells in the row"
                                     % (a1(row, col), len(ranked[0][1])), block=block)
                else:
                    first = cols[0]
                    self.add(sheet.name, row, first, "row-inconsistent",
                             "%d different formula shapes across %s:%s"
                             % (len(shapes), a1(row, cols[0]), a1(row, cols[-1])),
                             block=block)

    @staticmethod
    def is_base_column(cols, written, shape_of):
        """True when the row is one copied formula behind a base column.

        `=$C$4` in the first period column followed by four copies of
        `=C5*(1+$C$3)` is the ordinary shape of a ramp, not an inconsistency.
        The waiver only covers the first period column, so an odd cell further
        along the row is still reported.
        """
        if not written or written[0] != cols[0] or len(written) < 3:
            return False
        rest = set(shape_of[c] for c in written[1:])
        return len(rest) == 1 and shape_of[written[0]] not in rest

    # ----------------------------------------------------------------- units

    def check_units(self, sheet):
        for block in sheet.blocks:
            if block.type not in ("table", "calc"):
                continue
            periods = sheet.header_periods(block)
            spec_block = self.spec_block(block)
            total_row = sheet.total_row_of(block, spec_block)
            rows = [r for r in sheet.data_rows(block) if r != total_row]
            header_row = block.r1 + 1
            if periods:
                # A period table's unit belongs to the row label.
                for row in rows:
                    cols = [c for c in sorted(periods) if self.numeric_cell(sheet, row, c)]
                    if not cols:
                        continue
                    if any(not is_bare_format(sheet.cell(row, c).number_format) for c in cols):
                        continue
                    label = sheet.values.get((row, block.c1))
                    if has_unit(label):
                        continue
                    self.add(sheet.name, row, block.c1, "no-unit",
                             "row '%s' is a bare number and its label names no unit"
                             % (label if label not in (None, "") else a1(row, block.c1)),
                             block=block)
                continue
            for col in range(block.c1 + 1, block.c2 + 1):
                cells = [r for r in rows if self.numeric_cell(sheet, r, col)]
                if not cells:
                    continue
                if any(not is_bare_format(sheet.cell(r, col).number_format) for r in cells):
                    continue
                header = sheet.values.get((header_row, col))
                if has_unit(header):
                    continue
                labels = [sheet.values.get((r, block.c1)) for r in cells]
                if any(has_unit(lab) for lab in labels):
                    continue
                self.add(sheet.name, header_row, col, "no-unit",
                         "column '%s' is a bare number and its header names no unit"
                         % (header if header not in (None, "") else get_column_letter(col)),
                         block=block)

    def numeric_cell(self, sheet, row, col):
        key = (row, col)
        if key in sheet.formulas:
            return True
        return is_number(sheet.values.get(key))

    # ---------------------------------------------------------------- totals

    def check_totals(self, sheet):
        """A total row against the rows above it, from cached values when present."""
        for block in sheet.blocks:
            if block.type not in ("table", "calc"):
                continue
            spec_block = self.spec_block(block)
            total_row = sheet.total_row_of(block, spec_block)
            if total_row is None:
                continue
            rows = [r for r in sheet.data_rows(block) if r < total_row]
            if not rows:
                continue
            for col in range(block.c1 + 1, block.c2 + 1):
                total = self.number_at(sheet, total_row, col)
                if total is None:
                    continue
                parts = [self.number_at(sheet, r, col) for r in rows]
                known = [p for p in parts if p is not None]
                if len(known) < 2 or len(known) != len([
                        r for r in rows if self.numeric_cell(sheet, r, col)]):
                    continue                 # a value we cannot read: say nothing
                expected = sum(known)
                slack = max(abs(expected), abs(total), 1.0) * TOTAL_REL_TOLERANCE
                if abs(expected - total) <= slack:
                    continue
                self.add(sheet.name, total_row, col, "total-mismatch",
                         "the total shows %s and the %d rows above it sum to %s"
                         % (_trim_number(total), len(known), _trim_number(expected)),
                         block=block)

    def number_at(self, sheet, row, col):
        key = (row, col)
        if key in sheet.formulas:
            cached = self.cached.get(sheet.name, {}).get(key)
            return float(cached) if is_number(cached) else None
        value = sheet.values.get(key)
        return float(value) if is_number(value) else None

    # ------------------------------------------------------------ empty band

    def check_empty_band(self, sheet):
        rows = sorted(set(r for r, _c in list(sheet.values) + list(sheet.formulas)))
        if len(rows) < 2:
            return
        first, last = rows[0], rows[-1]
        run_start = None
        for row in range(first, last + 2):
            empty = row <= last and sheet.row_is_empty(row)
            if empty:
                if run_start is None:
                    run_start = row
                continue
            if run_start is not None:
                length = row - run_start
                if length >= EMPTY_BAND_ROWS:
                    self.add(sheet.name, 0, 0, "empty-band",
                             "rows %d to %d are empty inside the used range"
                             % (run_start, row - 1), scope="sheet")
                run_start = None

    # ----------------------------------------------------------- freeze pane

    def check_freeze(self, sheet):
        tables = [b for b in sheet.blocks if b.type in ("table", "calc")]
        if not tables:
            return
        rows = sorted(set(r for r, _c in list(sheet.values) + list(sheet.formulas)))
        if not rows or rows[-1] - rows[0] + 1 <= FREEZE_MIN_ROWS:
            return
        freeze = sheet.ws.freeze_panes
        first = min(b.r1 + 2 for b in tables)
        widest = max(tables, key=lambda b: b.c2 - b.c1)
        want = a1(first, widest.c1 + 1)
        if not freeze:
            self.add(sheet.name, 0, 0, "no-freeze",
                     "a table sheet %d rows tall has no freeze pane; %s holds labels and header"
                     % (rows[-1] - rows[0] + 1, want), scope="sheet")
            return
        m = CELL_RE.match(str(freeze).replace("$", ""))
        if not m:
            return
        col = column_index_from_string(m.group(1).upper())
        row = int(m.group(2))
        if col <= 1 or row <= 1:
            self.add(sheet.name, 0, 0, "no-freeze",
                     "the freeze pane at %s holds %s; %s holds the labels and the header"
                     % (str(freeze).upper(),
                        "no columns" if col <= 1 else "no rows", want),
                     scope="sheet")

    # ------------------------------------------------------------ references

    def check_references(self):
        """Cycles, #REF-class references, and cross-sheet pulls nothing marks."""
        cycles = self.find_cycles()
        for cycle in cycles:
            names = ", ".join("%s!%s" % (s, a1(r, c)) for s, r, c in cycle)
            for sheet_name, row, col in cycle:
                self.add(sheet_name, row, col, "circular",
                         "this cell is in a reference cycle: %s" % names)
        formula_colour = (self.styler.formula_color or "").upper()
        link_colour = (getattr(self.styler, "link_color", None) or "").upper()
        marks_links = bool(link_colour) and link_colour != formula_colour
        # W14: the house style marks a link with an underline, not a colour, so the check follows the underline there.
        marks_underline = bool(getattr(self.styler, "link_underline", True)) and not marks_links
        for sheet in self.sheets:
            for key, formula in sorted(sheet.formulas.items()):
                row, col = key
                cell = sheet.cell(row, col)
                off_sheet = []
                for raw, ref_sheet, _area in formula_refs(formula):
                    if EXTERNAL_REF_RE.match(raw):
                        self.add(sheet.name, row, col, "ref-error",
                                 "the formula reads another workbook: %s" % raw)
                        continue
                    if "#REF!" in raw.upper():
                        self.add(sheet.name, row, col, "ref-error",
                                 "the formula carries a #REF!")
                        continue
                    if ref_sheet and self.book.resolve_sheet(ref_sheet) is None:
                        self.add(sheet.name, row, col, "ref-error",
                                 "the formula reads sheet '%s', which does not exist" % ref_sheet)
                        continue
                    if ref_sheet and self.book.resolve_sheet(ref_sheet) != sheet.name:
                        off_sheet.append(raw)
                if "#REF!" in str(formula).upper():
                    self.add(sheet.name, row, col, "ref-error",
                             "the formula carries a #REF!")
                if sheet.kind(row, col) == "readme":
                    continue            # the cover restates figures as prose
                colour = cell_colour(cell) or formula_colour
                underlined = bool(getattr(getattr(cell, "font", None), "underline", None))
                unmarked = (marks_links and colour != link_colour) or (marks_underline and not underlined)
                if cell.hyperlink is not None and unmarked:
                    self.add(sheet.name, row, col, "link-uncoloured",
                             "a hyperlink with no link colour or underline does not read as a link")
                if off_sheet and unmarked \
                        and self.is_pass_through(formula):
                    self.add(sheet.name, row, col, "off-sheet-uncoloured",
                             "%s is a bare pull from %s drawn in %s, so nothing marks it off-sheet"
                             % (a1(row, col), off_sheet[0], colour))
            if not (marks_links or marks_underline):
                continue
            for key, _value in sorted(sheet.values.items()):
                row, col = key
                cell = sheet.cell(row, col)
                if cell.hyperlink is None or sheet.kind(row, col) == "readme":
                    continue
                colour = cell_colour(cell) or formula_colour
                underlined = bool(getattr(getattr(cell, "font", None), "underline", None))
                if (marks_links and colour != link_colour) or (marks_underline and not underlined):
                    self.add(sheet.name, row, col, "link-uncoloured",
                             "a hyperlink with no link colour or underline does not read as a link")

    def is_pass_through(self, formula):
        """True when the formula is nothing but one reference: `='Inputs'!C3`."""
        try:
            items = [t for t in Tokenizer(formula).items if t.type != "WHITE-SPACE"]
        except TokenizerError:
            return False
        return len(items) == 1 and items[0].type == "OPERAND" and items[0].subtype == "RANGE"

    def find_cycles(self):
        """Every reference cycle among formula cells, each reported once."""
        graph = {}
        for key, targets in self.refs.items():
            graph[key] = set(t for t in targets if t in self.refs)
        colour = {}
        stack = []
        cycles = []
        seen_cycles = set()

        def visit(node):
            colour[node] = 1
            stack.append(node)
            for nxt in sorted(graph.get(node, ())):
                if colour.get(nxt, 0) == 0:
                    visit(nxt)
                elif colour.get(nxt) == 1:
                    start = stack.index(nxt)
                    cycle = tuple(sorted(stack[start:]))
                    if cycle not in seen_cycles:
                        seen_cycles.add(cycle)
                        cycles.append(cycle)
            stack.pop()
            colour[node] = 2

        limit = sys.getrecursionlimit()
        sys.setrecursionlimit(max(limit, len(graph) * 2 + 1000))
        try:
            for node in sorted(graph):
                if colour.get(node, 0) == 0:
                    visit(node)
        except RecursionError:                        # noqa: PERF203 - depth is the workbook's
            self.notes.append("the dependency graph is too deep to scan for cycles")
        finally:
            sys.setrecursionlimit(limit)
        return cycles

    # ----------------------------------------------------- inputs / outputs

    def check_inputs_and_outputs(self):
        inputs = set()
        for sheet in self.sheets:
            for key, value in sheet.values.items():
                row, col = key
                if not self.is_input_cell(sheet, row, col, value):
                    continue
                inputs.add((sheet.name, row, col))
                if (sheet.name, row, col) in self.referenced:
                    continue
                label = self.label_for(sheet, row, col)
                self.add(sheet.name, row, col, "unused-input",
                         "no formula reads this input%s" % (" ('%s')" % label if label else ""))
        for sheet in self.sheets:
            for block in sheet.blocks:
                if block.type != "output":
                    continue
                for row in sheet.data_rows(block):
                    for col in range(block.c1 + 1, block.c2 + 1):
                        if (row, col) not in sheet.formulas:
                            continue
                        if self.reaches_input((sheet.name, row, col), inputs):
                            continue
                        label = sheet.values.get((row, block.c1))
                        self.add(sheet.name, row, col, "dangling-output",
                                 "nothing this output depends on is an input%s"
                                 % (" ('%s')" % label if label else ""), block=block)

    def is_input_cell(self, sheet, row, col, value):
        """An input: the snapshot's own `input` kind, or a literal in the input colour."""
        kind = sheet.kind(row, col)
        if kind == "input":
            return True
        if kind is not None:
            return False
        if not is_number(value) and not (isinstance(value, str) and value.strip() == "TBD"):
            return False
        colour = cell_colour(sheet.cell(row, col))
        wanted = (self.styler.input_color or "").upper()
        if not colour or not wanted or colour != wanted:
            return False
        return True

    def label_for(self, sheet, row, col):
        block = sheet.block_at(row, col)
        first = block.c1 if block else 1
        for c in range(col - 1, first - 1, -1):
            value = sheet.values.get((row, c))
            if isinstance(value, str) and value.strip():
                return value
        return None

    def reaches_input(self, start, inputs):
        seen = set()
        stack = [start]
        while stack:
            node = stack.pop()
            if node in seen:
                continue
            seen.add(node)
            if node in inputs:
                return True
            for nxt in self.refs.get(node, ()):
                if nxt not in seen:
                    stack.append(nxt)
        return False

    # ------------------------------------------------------------- read me

    def check_cover(self):
        cover = None
        for sheet in self.sheets:
            if build_xlsx.norm(sheet.name).replace(" ", "") == "readme":
                cover = sheet
                break
        if cover is None:
            self.add("", 0, 0, "no-cover",
                     "the workbook has no Read Me tab, so nothing says what it is for",
                     scope="workbook")
            return
        text = " ".join(
            str(v).lower() for v in cover.values.values() if isinstance(v, str))
        missing = [s.name for s in self.sheets
                   if s is not cover and s.name.lower() not in text]
        if missing:
            self.add("", 0, 0, "no-cover",
                     "the Read Me has no Contents row for %s" % ", ".join(missing),
                     scope="workbook")


def _trim_number(value):
    """A float as the shortest text that still reads as the same number."""
    if value == int(value) and abs(value) < 1e15:
        return "{:,}".format(int(value))
    return "{:,.4f}".format(value).rstrip("0").rstrip(".")


# ------------------------------------------------------------------ entry point


def check_workbook(xlsx_path, snapshot_path=None, spec_path=None):
    """Run every craft check and return the result dict."""
    snapshot_raw = None
    notes = []
    if snapshot_path:
        if os.path.exists(snapshot_path):
            snapshot_raw = load_json(snapshot_path)
        else:
            notes.append("no snapshot at %s; block checks were skipped"
                         % os.path.basename(snapshot_path))
    spec = None
    if spec_path:
        if os.path.exists(spec_path):
            spec = load_json(spec_path)
        else:
            notes.append("no spec at %s; house defaults were used"
                         % os.path.basename(spec_path))
    checker = Checker(xlsx_path, snapshot_raw, spec)
    checker.notes = notes + checker.notes
    if snapshot_raw is not None and checker.snapshot.empty:
        checker.notes.append("the snapshot describes no blocks; block checks were skipped")
    return checker.run()


USAGE = ("usage: python3 builder/check_workbook.py <workbook.xlsx> <snapshot.json> "
         "[<spec.json>] [--out <findings.json>]")


def main(argv):
    out_path = None
    args = []
    rest = list(argv[1:])
    while rest:
        item = rest.pop(0)
        if item in ("-h", "--help"):
            print(USAGE)
            return 0
        if item == "--out":
            out_path = rest.pop(0) if rest else None
        elif item.startswith("--out="):
            out_path = item.split("=", 1)[1]
        else:
            args.append(item)
    if not args:
        print(USAGE)
        return 2
    try:
        result = check_workbook(args[0],
                                args[1] if len(args) > 1 else None,
                                args[2] if len(args) > 2 else None)
    except Exception as exc:                          # noqa: BLE001 - one JSON object, always
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

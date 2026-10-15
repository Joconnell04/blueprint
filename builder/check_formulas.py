#!/usr/bin/env python3
"""Formula check after build (roadmap D13).

Usage:
    python3 builder/check_formulas.py <workbook.xlsx> [<out.json>] [--engine auto|python|libreoffice]

Two stages run over every formula in the workbook:

1. Static. Each formula is tokenised with openpyxl's Tokenizer and every
   reference is checked: sheets that do not exist, ranges outside the sheet's
   used area, single-cell references to empty cells that feed arithmetic,
   self references and reference cycles, #REF!, unknown names, and formulas
   the tokenizer cannot parse.

2. Evaluation. With LibreOffice installed (soffice on PATH or in
   /Applications/LibreOffice.app) the workbook is recalculated headless into a
   temp copy and the cached values are read back with data_only=True; any
   error value (#DIV/0!, #VALUE!, #REF!, #NAME?, #N/A, #NUM!, #NULL!) is
   reported. Otherwise a small pure-Python engine evaluates the formulas in
   dependency order and reports the same error codes. Functions the engine
   does not know produce an "unsupported" warning, never a crash.

When the workbook has a sibling <name>.snapshot.json (written by
build_xlsx.py) each finding is also mapped to the block that owns the cell.

Output (stdout, and the out.json path when given):

    { "ok": true, "engine": "libreoffice" | "python", "evaluated": n,
      "findings": [ { "sheet", "cell", "blockId"?, "severity", "code",
                      "message", "formula", "propagated"? } ],
      "summary": { "errors": n, "warnings": n } }

Exit status is 0 whenever the check ran, whatever it found. Exit 1 only when
the checker itself crashed; stdout then carries {"ok": false, "error": ...}.

Python 3.9, openpyxl only.
"""

import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile

import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils import get_column_letter, column_index_from_string

# ------------------------------------------------------------------ constants

ERROR_CODES = {
    "#DIV/0!": "div0",
    "#VALUE!": "value",
    "#REF!": "ref",
    "#NAME?": "name",
    "#N/A": "na",
    "#NUM!": "num",
    "#NULL!": "null",
}
CODE_TO_ERROR = dict((v, k) for k, v in ERROR_CODES.items())

# Codes that are warnings; everything else is an error.
WARNING_CODES = {"empty-ref", "outside-used-area", "unsupported", "external-ref"}

# Functions whose direct range or cell arguments may legitimately be blank.
BLANK_TOLERANT = {
    "SUM", "AVERAGE", "MIN", "MAX", "MINA", "MAXA", "COUNT", "COUNTA",
    "COUNTBLANK", "SUMPRODUCT", "PRODUCT", "MEDIAN", "ISBLANK", "IF",
    "IFERROR", "AND", "OR", "N", "COUNTIF", "SUMIF", "AVERAGEIF",
    "SUMIFS", "COUNTIFS", "AVERAGEIFS", "INDEX", "MATCH", "VLOOKUP",
    "HLOOKUP", "XLOOKUP", "LOOKUP", "OFFSET", "CHOOSE", "TEXT", "CONCATENATE",
    "CONCAT", "TEXTJOIN", "LEN", "ISNUMBER", "ISTEXT", "ISERROR",
}

# Range expansion cap for dependency tracking (cells per range).
RANGE_EXPAND_CAP = 20000

CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")
RANGE_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+):\$?([A-Za-z]{1,3})\$?([0-9]+)$")
COL_RANGE_RE = re.compile(r"^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$")
ROW_RANGE_RE = re.compile(r"^\$?([0-9]+):\$?([0-9]+)$")
NAME_RE = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.]*$")

MAX_ROW = 1048576
MAX_COL = 16384

LIBREOFFICE_TIMEOUT = 120


# ------------------------------------------------------------------ helpers


def a1(row, col):
    return "%s%d" % (get_column_letter(col), row)


def split_sheet(ref):
    """Split 'Sheet name'!A1 into (sheet or None, 'A1'). Unescapes doubled quotes."""
    if ref.upper().endswith("#REF!"):
        base = ref[:-5]
        if not base:
            return None, "#REF!"
        if not base.endswith("!"):
            return None, ref
        sheet, rest = base[:-1], "#REF!"
    elif "!" not in ref:
        return None, ref
    else:
        sheet, rest = ref.rsplit("!", 1)
    if len(sheet) >= 2 and sheet[0] == "'" and sheet[-1] == "'":
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, rest


def parse_area(text):
    """Parse A1, A1:B2, A:A or 1:1 into (r1, c1, r2, c2) or None."""
    m = CELL_RE.match(text)
    if m:
        r = int(m.group(2))
        c = column_index_from_string(m.group(1).upper())
        return (r, c, r, c)
    m = RANGE_RE.match(text)
    if m:
        c1 = column_index_from_string(m.group(1).upper())
        c2 = column_index_from_string(m.group(3).upper())
        r1, r2 = int(m.group(2)), int(m.group(4))
        return (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))
    m = COL_RANGE_RE.match(text)
    if m:
        c1 = column_index_from_string(m.group(1).upper())
        c2 = column_index_from_string(m.group(2).upper())
        return (1, min(c1, c2), MAX_ROW, max(c1, c2))
    m = ROW_RANGE_RE.match(text)
    if m:
        r1, r2 = int(m.group(1)), int(m.group(2))
        return (min(r1, r2), 1, max(r1, r2), MAX_COL)
    return None


def area_text(area):
    r1, c1, r2, c2 = area
    if (r1, c1) == (r2, c2):
        return a1(r1, c1)
    return "%s:%s" % (a1(r1, c1), a1(r2, c2))


class Err(object):
    """An Excel error value flowing through the python engine.

    code is one of the ERROR_CODES values, or "unsupported" for a value the
    engine cannot compute. origin is the (sheet, cell) that produced it, so a
    finding can say whether the error started here or was inherited.
    """

    __slots__ = ("code", "origin", "detail")

    def __init__(self, code, origin=None, detail=""):
        self.code = code
        self.origin = origin
        self.detail = detail

    def __repr__(self):
        return "Err(%s)" % self.code


class RangeVal(object):
    """A rectangular block of cell values (list of rows) with its address."""

    __slots__ = ("sheet", "area", "rows")

    def __init__(self, sheet, area, rows):
        self.sheet = sheet
        self.area = area
        self.rows = rows

    def flat(self):
        for row in self.rows:
            for v in row:
                yield v

    @property
    def shape(self):
        return (len(self.rows), len(self.rows[0]) if self.rows else 0)


# ------------------------------------------------------------------ workbook


class Book(object):
    """The workbook as plain dicts: values, formulas, used areas, names."""

    def __init__(self, path):
        self.path = path
        self.wb = openpyxl.load_workbook(path, data_only=False)
        self.sheet_names = [ws.title for ws in self.wb.worksheets]
        self.lower_names = dict((n.lower(), n) for n in self.sheet_names)
        self.values = {}     # sheet -> {(row, col): literal value}
        self.formulas = {}   # sheet -> {(row, col): formula text}
        self.used = {}       # sheet -> (max_row, max_col) over non-empty cells
        for ws in self.wb.worksheets:
            vals, forms = {}, {}
            max_r = max_c = 0
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if v is None or v == "":
                        continue
                    key = (cell.row, cell.column)
                    if cell.data_type == "f" or (isinstance(v, str) and v.startswith("=")):
                        text = v if isinstance(v, str) else getattr(v, "text", str(v))
                        forms[key] = text
                    else:
                        vals[key] = v
                    if cell.row > max_r:
                        max_r = cell.row
                    if cell.column > max_c:
                        max_c = cell.column
            self.values[ws.title] = vals
            self.formulas[ws.title] = forms
            self.used[ws.title] = (max_r, max_c)
        self.names = self._defined_names()

    def _defined_names(self):
        """name (upper) -> (sheet, area) for names that point at one area."""
        out = {}
        dn = getattr(self.wb, "defined_names", None)
        items = []
        try:
            if hasattr(dn, "items"):
                items = list(dn.items())
            elif dn is not None:
                items = [(d.name, d) for d in dn.definedName]
        except Exception:
            items = []
        for ws in self.wb.worksheets:
            local = getattr(ws, "defined_names", None)
            try:
                if hasattr(local, "items"):
                    items.extend(local.items())
            except Exception:
                pass
        for name, defn in items:
            target = None
            try:
                for sheet, ref in defn.destinations:
                    area = parse_area(ref.replace("$", ""))
                    if area:
                        target = (sheet, area)
                    break
            except Exception:
                target = None
            out[str(name).upper()] = target
        return out

    def resolve_sheet(self, name):
        return self.lower_names.get(name.lower())

    def literal(self, sheet, row, col):
        return self.values[sheet].get((row, col))

    def is_empty(self, sheet, row, col):
        key = (row, col)
        return key not in self.values[sheet] and key not in self.formulas[sheet]

    def all_formulas(self):
        for sheet in self.sheet_names:
            for key in sorted(self.formulas[sheet]):
                yield sheet, key[0], key[1], self.formulas[sheet][key]


# ------------------------------------------------------------------ snapshot


def load_snapshot(xlsx_path):
    """Block extents from the sibling snapshot, or {} when there is none."""
    base, ext = os.path.splitext(xlsx_path)
    snap_path = base + ".snapshot.json"
    if not os.path.exists(snap_path):
        return {}
    try:
        with open(snap_path, "r") as handle:
            snap = json.load(handle)
    except Exception:
        return {}
    blocks = {}
    for sheet in snap.get("sheets", []):
        entries = []
        for b in sheet.get("blocks", []):
            try:
                r, c = int(b["row"]), int(b["col"])
                rows, cols = int(b.get("rows", 1)), int(b.get("cols", 1))
                entries.append((r, c, r + rows - 1, c + cols - 1, b.get("id")))
            except (KeyError, TypeError, ValueError):
                continue
        blocks[sheet.get("name")] = entries
    return blocks


def block_for(blocks, sheet, row, col):
    for r1, c1, r2, c2, bid in blocks.get(sheet, []):
        if r1 <= row <= r2 and c1 <= col <= c2:
            return bid
    return None


# ------------------------------------------------------------------ parser


class ParseError(Exception):
    pass


class Unsupported(Exception):
    """Raised while parsing or evaluating for constructs the engine skips."""


def tokens_of(formula):
    return [t for t in Tokenizer(formula).items if t.type != "WHITE-SPACE"]


class Parser(object):
    """Recursive descent over Tokenizer tokens; produces a tuple AST.

    Nodes:
        ("num", float) ("str", text) ("bool", b) ("err", "#DIV/0!")
        ("ref", raw_text)              a cell, range or name operand
        ("neg", node) ("pos", node) ("pct", node)
        ("bin", op, left, right)
        ("call", NAME, [args])         an arg may be None for an empty slot
    """

    PRECEDENCE = [
        ("=", "<>", "<", ">", "<=", ">="),
        ("&",),
        ("+", "-"),
        ("*", "/"),
        ("^",),
    ]

    def __init__(self, tokens):
        self.toks = tokens
        self.i = 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else None

    def take(self):
        tok = self.peek()
        if tok is None:
            raise ParseError("unexpected end of formula")
        self.i += 1
        return tok

    def parse(self):
        node = self.expr(0)
        if self.peek() is not None:
            raise ParseError("unexpected token %r" % self.peek().value)
        return node

    def expr(self, level):
        if level >= len(self.PRECEDENCE):
            return self.unary()
        ops = self.PRECEDENCE[level]
        left = self.expr(level + 1)
        while True:
            tok = self.peek()
            if tok is None or tok.type != "OPERATOR-INFIX" or tok.value not in ops:
                break
            self.take()
            if tok.value == "^":
                right = self.unary()
            else:
                right = self.expr(level + 1)
            left = ("bin", tok.value, left, right)
        return left

    def unary(self):
        tok = self.peek()
        if tok is not None and tok.type == "OPERATOR-PREFIX":
            self.take()
            operand = self.unary()
            return ("neg", operand) if tok.value == "-" else ("pos", operand)
        node = self.primary()
        while True:
            tok = self.peek()
            if tok is not None and tok.type == "OPERATOR-POSTFIX" and tok.value == "%":
                self.take()
                node = ("pct", node)
            else:
                break
        return node

    def primary(self):
        tok = self.take()
        if tok.type == "OPERAND":
            if tok.subtype == "NUMBER":
                return ("num", float(tok.value))
            if tok.subtype == "TEXT":
                return ("str", tok.value[1:-1].replace('""', '"'))
            if tok.subtype == "LOGICAL":
                return ("bool", tok.value.upper() == "TRUE")
            if tok.subtype == "ERROR":
                return ("err", tok.value)
            return ("ref", tok.value)
        if tok.type == "PAREN" and tok.subtype == "OPEN":
            node = self.expr(0)
            close = self.take()
            if close.type != "PAREN" or close.subtype != "CLOSE":
                raise ParseError("expected )")
            return node
        if tok.type == "FUNC" and tok.subtype == "OPEN":
            name = tok.value[:-1].upper()
            if name.startswith("_XLFN."):
                name = name[6:]
            args = []
            nxt = self.peek()
            if nxt is not None and nxt.type == "FUNC" and nxt.subtype == "CLOSE":
                self.take()
                return ("call", name, args)
            while True:
                nxt = self.peek()
                if nxt is not None and nxt.type == "SEP":
                    args.append(None)
                elif nxt is not None and nxt.type == "FUNC" and nxt.subtype == "CLOSE":
                    args.append(None)
                else:
                    args.append(self.expr(0))
                sep = self.take()
                if sep.type == "SEP":
                    continue
                if sep.type == "FUNC" and sep.subtype == "CLOSE":
                    return ("call", name, args)
                raise ParseError("expected , or ) in %s" % name)
        if tok.type == "ARRAY":
            raise Unsupported("array constant")
        if tok.type == "OPERATOR-INFIX" and tok.value in (":", ",", " "):
            raise Unsupported("range operator %r" % tok.value)
        raise ParseError("unexpected token %r" % tok.value)


def parse_formula(formula):
    """(ast, None) or (None, ("parse"|"unsupported", message))."""
    try:
        toks = tokens_of(formula)
    except TokenizerError as exc:
        return None, ("parse", "cannot tokenize: %s" % exc)
    if len(toks) == 1 and toks[0].type == "LITERAL":
        return ("str", formula), None
    try:
        return Parser(toks).parse(), None
    except Unsupported as exc:
        return None, ("unsupported", str(exc))
    except ParseError as exc:
        return None, ("parse", str(exc))
    except Exception as exc:  # defensive: never crash on one odd formula
        return None, ("parse", "cannot parse: %s" % exc)


# ------------------------------------------------------------------ references


class RefInfo(object):
    """One reference operand, classified."""

    __slots__ = ("raw", "kind", "sheet", "area", "explicit_sheet", "detail")

    def __init__(self, raw, kind, sheet=None, area=None, explicit_sheet=False, detail=""):
        self.raw = raw
        self.kind = kind          # area | name | ref-error | missing-sheet | external | unsupported
        self.sheet = sheet
        self.area = area
        self.explicit_sheet = explicit_sheet
        self.detail = detail


def classify_ref(book, raw, current_sheet):
    if raw.startswith("["):
        return RefInfo(raw, "external", detail="reference into another workbook")
    sheet_part, rest = split_sheet(raw)
    if sheet_part is not None and sheet_part.startswith("["):
        return RefInfo(raw, "external", detail="reference into another workbook")
    if "[" in rest and rest.endswith("]"):
        return RefInfo(raw, "unsupported", detail="structured table reference")
    if rest.upper() == "#REF!":
        return RefInfo(raw, "ref-error")
    sheet = current_sheet
    explicit = False
    if sheet_part is not None:
        explicit = True
        if ":" in sheet_part and book.resolve_sheet(sheet_part) is None:
            return RefInfo(raw, "unsupported", detail="3D reference across sheets")
        sheet = book.resolve_sheet(sheet_part)
        if sheet is None:
            return RefInfo(raw, "missing-sheet", detail=sheet_part)
    area = parse_area(rest)
    if area is not None:
        return RefInfo(raw, "area", sheet=sheet, area=area, explicit_sheet=explicit)
    if NAME_RE.match(rest):
        target = book.names.get(rest.upper(), "absent")
        if target == "absent":
            return RefInfo(raw, "name", detail=rest)
        if target is None:
            return RefInfo(raw, "unsupported", detail="defined name %s is not a plain range" % rest)
        tsheet, tarea = target
        tsheet = book.resolve_sheet(tsheet) if tsheet else sheet
        if tsheet is None:
            return RefInfo(raw, "missing-sheet", detail=str(target[0]))
        return RefInfo(raw, "area", sheet=tsheet, area=tarea, explicit_sheet=True)
    return RefInfo(raw, "unsupported", detail="reference %s" % raw)


def walk_refs(node, parent, out):
    """Collect (ref_raw, parent) pairs from an AST.

    parent is None for a bare reference (=A1), "op" when the reference is an
    operand of an operator, or the upper-case function name when it is a
    direct argument of a call.
    """
    if node is None:
        return
    kind = node[0]
    if kind == "ref":
        out.append((node[1], parent))
    elif kind == "err" and node[1].upper() == "#REF!":
        out.append((node[1], parent))
    elif kind in ("neg", "pos", "pct"):
        walk_refs(node[1], "op", out)
    elif kind == "bin":
        walk_refs(node[2], "op", out)
        walk_refs(node[3], "op", out)
    elif kind == "call":
        for arg in node[2]:
            walk_refs(arg, node[1], out)


# ------------------------------------------------------------------ checker


class Checker(object):
    def __init__(self, book, blocks):
        self.book = book
        self.blocks = blocks
        self.findings = []
        self.seen = set()
        self.static_errors = set()       # (sheet, row, col) with a static error
        self.cell_codes = {}             # (sheet, row, col) -> set of finding codes
        self.asts = {}                   # (sheet, row, col) -> ast or None
        self.deps = {}                   # (sheet, row, col) -> set of (sheet, row, col)
        self.formula_text = {}

    # -------------------------------------------------------------- findings

    def add(self, sheet, row, col, code, message, formula=None, propagated=None):
        key = (sheet, row, col, code)
        if key in self.seen:
            return
        self.seen.add(key)
        severity = "warning" if code in WARNING_CODES else "error"
        if severity == "error":
            self.static_errors.add((sheet, row, col))
        self.cell_codes.setdefault((sheet, row, col), set()).add(code)
        entry = {
            "sheet": sheet,
            "cell": a1(row, col),
            "severity": severity,
            "code": code,
            "message": message,
            "formula": formula if formula is not None else self.formula_text.get((sheet, row, col), ""),
        }
        bid = block_for(self.blocks, sheet, row, col)
        if bid:
            entry["blockId"] = bid
        if propagated:
            entry["propagated"] = True
        self.findings.append(entry)

    # ---------------------------------------------------------------- static

    def static(self):
        book = self.book
        for sheet, row, col, formula in book.all_formulas():
            key = (sheet, row, col)
            self.formula_text[key] = formula
            ast, problem = parse_formula(formula)
            self.asts[key] = ast
            self.deps[key] = set()
            if problem:
                self.add(sheet, row, col, problem[0], problem[1])
                continue
            refs = []
            walk_refs(ast, None, refs)
            for raw, enclosing in refs:
                info = classify_ref(book, raw, sheet)
                if info.kind == "missing-sheet":
                    self.add(sheet, row, col, "missing-sheet",
                             "references sheet '%s' which does not exist" % info.detail)
                elif info.kind == "ref-error":
                    self.add(sheet, row, col, "ref", "contains #REF! (a deleted or invalid reference)")
                elif info.kind == "name":
                    self.add(sheet, row, col, "name", "unknown name '%s'" % info.detail)
                elif info.kind == "external":
                    self.add(sheet, row, col, "external-ref", "%s: %s" % (raw, info.detail))
                elif info.kind == "unsupported":
                    self.add(sheet, row, col, "unsupported", "%s: %s" % (raw, info.detail))
                elif info.kind == "area":
                    self.check_area(sheet, row, col, info, enclosing)
        self.cycles()

    def check_area(self, sheet, row, col, info, enclosing):
        r1, c1, r2, c2 = info.area
        target = info.sheet
        max_r, max_c = self.book.used[target]
        whole = (r2 - r1 + 1 >= MAX_ROW) or (c2 - c1 + 1 >= MAX_COL)
        where = "" if target == sheet else "'%s'!" % target
        # Dependencies (capped) for cycle detection and evaluation order.
        rr2 = min(r2, max_r)
        cc2 = min(c2, max_c)
        if rr2 >= r1 and cc2 >= c1:
            count = (rr2 - r1 + 1) * (cc2 - c1 + 1)
            if count <= RANGE_EXPAND_CAP:
                forms = self.book.formulas[target]
                for r in range(r1, rr2 + 1):
                    for c in range(c1, cc2 + 1):
                        if (r, c) in forms:
                            self.deps[(sheet, row, col)].add((target, r, c))
            else:
                forms = self.book.formulas[target]
                for (r, c) in forms:
                    if r1 <= r <= rr2 and c1 <= c <= cc2:
                        self.deps[(sheet, row, col)].add((target, r, c))
        if whole:
            return
        if r1 > max_r or c1 > max_c:
            self.add(sheet, row, col, "outside-used-area",
                     "%s%s lies entirely outside the used area of '%s' (%s)"
                     % (where, area_text(info.area), target, a1(max(max_r, 1), max(max_c, 1))))
            return
        if r2 > max_r or c2 > max_c:
            self.add(sheet, row, col, "outside-used-area",
                     "%s%s extends past the used area of '%s' (%s)"
                     % (where, area_text(info.area), target, a1(max_r, max_c)))
            return
        if (r1, c1) == (r2, c2) and self.book.is_empty(target, r1, c1):
            if enclosing in BLANK_TOLERANT:
                return
            self.add(sheet, row, col, "empty-ref",
                     "%s%s is empty and feeds this formula" % (where, area_text(info.area)))

    def cycles(self):
        """Report every formula cell that sits on a reference cycle.

        Iterative Tarjan over the formula dependency graph. Components come
        out with dependencies first, which is also the evaluation order.
        """
        index, low = {}, {}
        on_stack, stack, sccs = set(), [], []
        counter = 0
        for root in self.deps:
            if root in index:
                continue
            index[root] = low[root] = counter
            counter += 1
            stack.append(root)
            on_stack.add(root)
            work = [(root, iter(self.deps[root]))]
            while work:
                v, it = work[-1]
                descended = False
                for w in it:
                    if w not in self.deps:
                        continue
                    if w not in index:
                        index[w] = low[w] = counter
                        counter += 1
                        stack.append(w)
                        on_stack.add(w)
                        work.append((w, iter(self.deps[w])))
                        descended = True
                        break
                    if w in on_stack:
                        low[v] = min(low[v], index[w])
                if descended:
                    continue
                work.pop()
                if work:
                    parent = work[-1][0]
                    low[parent] = min(low[parent], low[v])
                if low[v] == index[v]:
                    comp = []
                    while True:
                        w = stack.pop()
                        on_stack.discard(w)
                        comp.append(w)
                        if w == v:
                            break
                    sccs.append(comp)
        self.circular = set()
        for comp in sccs:
            if len(comp) > 1 or comp[0] in self.deps.get(comp[0], ()):
                self.circular.update(comp)
                home = comp[0][0]
                others = sorted(a1(w[1], w[2]) if w[0] == home else "'%s'!%s" % (w[0], a1(w[1], w[2]))
                                for w in comp)
                for v in comp:
                    if len(comp) == 1:
                        msg = "references itself"
                    else:
                        msg = "circular reference through %s" % ", ".join(others)
                    self.add(v[0], v[1], v[2], "circular", msg)
        self.sccs = sccs

    # ------------------------------------------------------------ evaluation

    def evaluate_python(self):
        engine = Engine(self.book, self.asts, self.circular)
        # Tarjan emits components in reverse topological order: dependencies first.
        order = [v for comp in self.sccs for v in comp]
        evaluated = 0
        for key in order:
            if key not in self.asts:
                continue
            if self.asts[key] is None:
                continue  # the static stage already reported the parse problem
            value = engine.cell_value(key)
            evaluated += 1
            if isinstance(value, Err):
                sheet, row, col = key
                if value.code == "unsupported":
                    already = self.cell_codes.get(key, set())
                    if value.origin == key and not (already & {"unsupported", "external-ref"}):
                        self.add(sheet, row, col, "unsupported", value.detail or "unsupported construct")
                    continue
                if key in self.static_errors:
                    continue  # the static stage already explained this cell
                propagated = value.origin is not None and value.origin != key
                msg = "evaluates to %s" % CODE_TO_ERROR.get(value.code, value.code)
                if value.detail and not propagated:
                    msg += ": " + value.detail
                if propagated:
                    o = value.origin
                    src = a1(o[1], o[2]) if o[0] == sheet else "'%s'!%s" % (o[0], a1(o[1], o[2]))
                    msg += " (inherited from %s)" % src
                self.add(sheet, row, col, value.code, msg, propagated=propagated)
        return evaluated

    def evaluate_libreoffice(self, soffice):
        tmp = tempfile.mkdtemp(prefix="wbcheck_")
        try:
            src = os.path.join(tmp, "in.xlsx")
            shutil.copyfile(self.book.path, src)
            outdir = os.path.join(tmp, "out")
            os.makedirs(outdir)
            profile = os.path.join(tempfile.gettempdir(), "workbook-planner-lo-profile")
            cmd = [
                soffice,
                "-env:UserInstallation=file://" + profile,
                "--headless", "--norestore", "--convert-to", "xlsx",
                "--outdir", outdir, src,
            ]
            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                           timeout=LIBREOFFICE_TIMEOUT, check=False)
            result = os.path.join(outdir, "in.xlsx")
            if not os.path.exists(result):
                raise RuntimeError("LibreOffice produced no output")
            wb = openpyxl.load_workbook(result, data_only=True)
            evaluated = 0
            for sheet, row, col, formula in self.book.all_formulas():
                if sheet not in wb.sheetnames:
                    continue
                cell = wb[sheet].cell(row=row, column=col)
                evaluated += 1
                v = cell.value
                if cell.data_type == "e" or (isinstance(v, str) and (v in ERROR_CODES or v.startswith("Err:"))):
                    code = ERROR_CODES.get(v, "error")
                    if (sheet, row, col) in self.static_errors:
                        continue
                    self.add(sheet, row, col, code, "evaluates to %s" % v)
            return evaluated
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


# ------------------------------------------------------------------ engine


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


class Engine(object):
    def __init__(self, book, asts, circular):
        self.book = book
        self.asts = asts
        self.circular = circular
        self.cache = {}
        self.in_progress = set()
        self.current = None

    # ----------------------------------------------------------- cell access

    def cell_value(self, key):
        if key in self.cache:
            return self.cache[key]
        sheet, row, col = key
        if key not in self.asts:
            return self.book.literal(sheet, row, col)
        if key in self.circular:
            # Excel shows 0 for cells in a circular chain; the static stage
            # already reported the cycle.
            self.cache[key] = 0
            return 0
        if key in self.in_progress:
            self.cache[key] = 0
            return 0
        ast = self.asts[key]
        if ast is None:
            value = Err("unsupported", key, "formula could not be parsed")
            self.cache[key] = value
            return value
        self.in_progress.add(key)
        saved = self.current
        self.current = key
        try:
            value = self.eval(ast)
        except Unsupported as exc:
            value = Err("unsupported", key, str(exc))
        except RecursionError:
            value = Err("unsupported", key, "dependency chain too deep to evaluate")
        except ZeroDivisionError:
            value = Err("div0", key)
        except (OverflowError, ValueError):
            value = Err("num", key)
        finally:
            self.current = saved
            self.in_progress.discard(key)
        if isinstance(value, RangeVal):
            value = self.scalar(value)
        self.cache[key] = value
        return value

    def err(self, code, detail=""):
        return Err(code, self.current, detail)

    # -------------------------------------------------------------- evaluate

    def eval(self, node):
        kind = node[0]
        if kind == "num":
            return node[1]
        if kind == "str":
            return node[1]
        if kind == "bool":
            return node[1]
        if kind == "err":
            return self.err(ERROR_CODES.get(node[1], "value"), "literal %s" % node[1])
        if kind == "ref":
            return self.ref(node[1])
        if kind == "neg":
            v = self.num(self.eval(node[1]))
            return v if isinstance(v, Err) else -v
        if kind == "pos":
            return self.eval(node[1])
        if kind == "pct":
            v = self.num(self.eval(node[1]))
            return v if isinstance(v, Err) else v / 100.0
        if kind == "bin":
            return self.binary(node[1], self.eval(node[2]), self.eval(node[3]))
        if kind == "call":
            return self.call(node[1], node[2])
        raise Unsupported("node %s" % kind)

    def ref(self, raw):
        sheet, _, _ = self.current
        info = classify_ref(self.book, raw, sheet)
        if info.kind == "missing-sheet":
            return self.err("ref", "sheet '%s' does not exist" % info.detail)
        if info.kind == "ref-error":
            return self.err("ref", "#REF!")
        if info.kind == "name":
            return self.err("name", "unknown name '%s'" % info.detail)
        if info.kind != "area":
            raise Unsupported(info.detail or raw)
        r1, c1, r2, c2 = info.area
        if (r1, c1) == (r2, c2):
            return self.cell_value((info.sheet, r1, c1))
        max_r, max_c = self.book.used[info.sheet]
        r2 = min(r2, max(max_r, r1))
        c2 = min(c2, max(max_c, c1))
        if (r2 - r1 + 1) * (c2 - c1 + 1) > RANGE_EXPAND_CAP:
            raise Unsupported("range %s is too large to evaluate" % raw)
        rows = []
        for r in range(r1, r2 + 1):
            rows.append([self.cell_value((info.sheet, r, c)) for c in range(c1, c2 + 1)])
        return RangeVal(info.sheet, (r1, c1, r2, c2), rows)

    def scalar(self, v):
        """A range where a single value is needed: 1x1 collapses, else unsupported."""
        if isinstance(v, RangeVal):
            if v.shape == (1, 1):
                return v.rows[0][0]
            raise Unsupported("range %s used where a single value is needed (array formula)"
                              % area_text(v.area))
        return v

    def num(self, v):
        """Coerce to a number the way arithmetic does; Err passes through."""
        v = self.scalar(v)
        if isinstance(v, Err):
            return v
        if v is None:
            return 0
        if isinstance(v, bool):
            return 1 if v else 0
        if is_number(v):
            return v
        if isinstance(v, str):
            try:
                return float(v.strip())
            except ValueError:
                return self.err("value", "text '%s' used in arithmetic" % v[:30])
        try:
            return float(v)
        except (TypeError, ValueError):
            return self.err("value", "cannot use %r in arithmetic" % (v,))

    def text(self, v):
        v = self.scalar(v)
        if isinstance(v, Err):
            return v
        if v is None:
            return ""
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if is_number(v):
            if float(v).is_integer():
                return str(int(v))
            return repr(float(v))
        return str(v)

    def boolean(self, v):
        v = self.scalar(v)
        if isinstance(v, Err):
            return v
        if v is None:
            return False
        if isinstance(v, bool):
            return v
        if is_number(v):
            return v != 0
        if isinstance(v, str):
            up = v.strip().upper()
            if up == "TRUE":
                return True
            if up == "FALSE":
                return False
            return self.err("value", "text '%s' used as a condition" % v[:30])
        return bool(v)

    def binary(self, op, a, b):
        if op == "&":
            ta, tb = self.text(a), self.text(b)
            if isinstance(ta, Err):
                return ta
            if isinstance(tb, Err):
                return tb
            return ta + tb
        if op in ("=", "<>", "<", ">", "<=", ">="):
            return self.compare(op, self.scalar(a), self.scalar(b))
        x, y = self.num(a), self.num(b)
        if isinstance(x, Err):
            return x
        if isinstance(y, Err):
            return y
        try:
            if op == "+":
                return x + y
            if op == "-":
                return x - y
            if op == "*":
                return x * y
            if op == "/":
                if y == 0:
                    return self.err("div0", "division by zero")
                return x / y
            if op == "^":
                if x == 0 and y < 0:
                    return self.err("div0", "zero raised to a negative power")
                result = math.pow(x, y)
                if isinstance(result, complex) or math.isnan(result):
                    return self.err("num")
                return result
        except OverflowError:
            return self.err("num", "overflow")
        except (ValueError, ZeroDivisionError):
            return self.err("num", "invalid power")
        raise Unsupported("operator %s" % op)

    def compare(self, op, a, b):
        if isinstance(a, Err):
            return a
        if isinstance(b, Err):
            return b

        def rank(v):
            if v is None:
                return (0, 0)
            if isinstance(v, bool):
                return (2, 1 if v else 0)
            if is_number(v):
                return (0, v)
            return (1, str(v).lower())

        ra, rb = rank(a), rank(b)
        if ra[0] == 0 and rb[0] == 1 and a is None:
            ra = (1, "")
        if rb[0] == 0 and ra[0] == 1 and b is None:
            rb = (1, "")
        if op == "=":
            return ra == rb
        if op == "<>":
            return ra != rb
        if op == "<":
            return ra < rb
        if op == ">":
            return ra > rb
        if op == "<=":
            return ra <= rb
        return ra >= rb

    # ------------------------------------------------------------- functions

    def numbers(self, args, include_text=False):
        """Numbers from mixed args the way SUM does; returns list or Err."""
        out = []
        for arg in args:
            if arg is None:
                continue
            v = self.eval(arg)
            if isinstance(v, RangeVal):
                for item in v.flat():
                    if isinstance(item, Err):
                        return item
                    if is_number(item):
                        out.append(item)
                    elif include_text and (isinstance(item, (str, bool))):
                        out.append(1 if item is True else 0)
            else:
                if isinstance(v, Err):
                    return v
                if v is None:
                    continue
                n = self.num(v)
                if isinstance(n, Err):
                    return n
                out.append(n)
        return out

    def call(self, name, args):
        fn = getattr(self, "fn_" + name.replace(".", "_"), None)
        if fn is None:
            raise Unsupported("function %s is not supported by the python engine" % name)
        return fn(args)

    def arg(self, args, i, default=None):
        if i < len(args) and args[i] is not None:
            return self.eval(args[i])
        return default

    def fn_SUM(self, args):
        nums = self.numbers(args)
        return nums if isinstance(nums, Err) else sum(nums)

    def fn_PRODUCT(self, args):
        nums = self.numbers(args)
        if isinstance(nums, Err):
            return nums
        out = 1
        for n in nums:
            out *= n
        return out

    def fn_AVERAGE(self, args):
        nums = self.numbers(args)
        if isinstance(nums, Err):
            return nums
        if not nums:
            return self.err("div0", "AVERAGE of no numbers")
        return sum(nums) / len(nums)

    def fn_MEDIAN(self, args):
        nums = self.numbers(args)
        if isinstance(nums, Err):
            return nums
        if not nums:
            return self.err("num", "MEDIAN of no numbers")
        s = sorted(nums)
        n = len(s)
        return s[n // 2] if n % 2 else (s[n // 2 - 1] + s[n // 2]) / 2.0

    def fn_MIN(self, args):
        nums = self.numbers(args)
        return nums if isinstance(nums, Err) else (min(nums) if nums else 0)

    def fn_MAX(self, args):
        nums = self.numbers(args)
        return nums if isinstance(nums, Err) else (max(nums) if nums else 0)

    def fn_MINA(self, args):
        nums = self.numbers(args, include_text=True)
        return nums if isinstance(nums, Err) else (min(nums) if nums else 0)

    def fn_MAXA(self, args):
        nums = self.numbers(args, include_text=True)
        return nums if isinstance(nums, Err) else (max(nums) if nums else 0)

    def fn_COUNT(self, args):
        count = 0
        for arg in args:
            if arg is None:
                continue
            v = self.eval(arg)
            items = list(v.flat()) if isinstance(v, RangeVal) else [v]
            for item in items:
                if is_number(item):
                    count += 1
                elif not isinstance(v, RangeVal) and isinstance(item, str):
                    try:
                        float(item)
                        count += 1
                    except ValueError:
                        pass
        return count

    def fn_COUNTA(self, args):
        count = 0
        for arg in args:
            if arg is None:
                continue
            v = self.eval(arg)
            items = list(v.flat()) if isinstance(v, RangeVal) else [v]
            count += sum(1 for item in items if item is not None and item != "")
        return count

    def fn_COUNTBLANK(self, args):
        count = 0
        for arg in args:
            if arg is None:
                continue
            v = self.eval(arg)
            items = list(v.flat()) if isinstance(v, RangeVal) else [v]
            count += sum(1 for item in items if item is None or item == "")
        return count

    def fn_IF(self, args):
        cond = self.boolean(self.arg(args, 0))
        if isinstance(cond, Err):
            return cond
        if cond:
            return self.scalar(self.arg(args, 1, 0)) if len(args) > 1 else True
        if len(args) > 2:
            return self.scalar(self.arg(args, 2, 0))
        return False

    def fn_IFERROR(self, args):
        v = self.scalar(self.arg(args, 0))
        if isinstance(v, Err) and v.code != "unsupported":
            return self.scalar(self.arg(args, 1, ""))
        return v

    def _logicals(self, args):
        out = []
        for arg in args:
            if arg is None:
                continue
            v = self.eval(arg)
            if isinstance(v, RangeVal):
                for item in v.flat():
                    if isinstance(item, Err):
                        return item
                    if isinstance(item, bool) or is_number(item):
                        out.append(bool(item))
            else:
                b = self.boolean(v)
                if isinstance(b, Err):
                    return b
                out.append(b)
        return out

    def fn_AND(self, args):
        vals = self._logicals(args)
        if isinstance(vals, Err):
            return vals
        if not vals:
            return self.err("value", "AND with no logical values")
        return all(vals)

    def fn_OR(self, args):
        vals = self._logicals(args)
        if isinstance(vals, Err):
            return vals
        if not vals:
            return self.err("value", "OR with no logical values")
        return any(vals)

    def fn_NOT(self, args):
        b = self.boolean(self.arg(args, 0))
        return b if isinstance(b, Err) else (not b)

    def fn_TRUE(self, args):
        return True

    def fn_FALSE(self, args):
        return False

    def fn_NA(self, args):
        return self.err("na", "NA()")

    def _one_num(self, args, i=0, default=0):
        return self.num(self.arg(args, i, default))

    def fn_ABS(self, args):
        x = self._one_num(args)
        return x if isinstance(x, Err) else abs(x)

    def fn_INT(self, args):
        x = self._one_num(args)
        return x if isinstance(x, Err) else math.floor(x)

    def fn_SQRT(self, args):
        x = self._one_num(args)
        if isinstance(x, Err):
            return x
        if x < 0:
            return self.err("num", "SQRT of a negative number")
        return math.sqrt(x)

    def fn_EXP(self, args):
        x = self._one_num(args)
        if isinstance(x, Err):
            return x
        try:
            return math.exp(x)
        except OverflowError:
            return self.err("num", "overflow")

    def fn_LN(self, args):
        x = self._one_num(args)
        if isinstance(x, Err):
            return x
        if x <= 0:
            return self.err("num", "LN of a non-positive number")
        return math.log(x)

    def fn_LOG10(self, args):
        x = self._one_num(args)
        if isinstance(x, Err):
            return x
        if x <= 0:
            return self.err("num", "LOG10 of a non-positive number")
        return math.log10(x)

    def fn_LOG(self, args):
        x = self._one_num(args)
        base = self._one_num(args, 1, 10)
        if isinstance(x, Err):
            return x
        if isinstance(base, Err):
            return base
        if x <= 0 or base <= 0 or base == 1:
            return self.err("num", "LOG of a non-positive number or invalid base")
        return math.log(x, base)

    def fn_PI(self, args):
        return math.pi

    def fn_POWER(self, args):
        return self.binary("^", self.arg(args, 0, 0), self.arg(args, 1, 0))

    def fn_MOD(self, args):
        x = self._one_num(args)
        d = self._one_num(args, 1)
        if isinstance(x, Err):
            return x
        if isinstance(d, Err):
            return d
        if d == 0:
            return self.err("div0", "MOD by zero")
        return x - d * math.floor(x / d)

    def _round(self, args, mode):
        x = self._one_num(args)
        n = self._one_num(args, 1, 0)
        if isinstance(x, Err):
            return x
        if isinstance(n, Err):
            return n
        digits = int(n)
        factor = 10.0 ** digits
        scaled = abs(x) * factor
        if mode == "half":
            rounded = math.floor(scaled + 0.5 + 1e-9)
        elif mode == "up":
            rounded = math.ceil(scaled - 1e-9)
        else:
            rounded = math.floor(scaled + 1e-9)
        result = rounded / factor
        return -result if x < 0 else result

    def fn_ROUND(self, args):
        return self._round(args, "half")

    def fn_ROUNDUP(self, args):
        return self._round(args, "up")

    def fn_ROUNDDOWN(self, args):
        return self._round(args, "down")

    def fn_SUMPRODUCT(self, args):
        arrays = []
        for arg in args:
            if arg is None:
                continue
            v = self.eval(arg)
            if isinstance(v, Err):
                return v
            if not isinstance(v, RangeVal):
                n = self.num(v)
                if isinstance(n, Err):
                    return n
                v = RangeVal(None, (1, 1, 1, 1), [[n]])
            arrays.append(v)
        if not arrays:
            return 0
        shape = arrays[0].shape
        for arr in arrays[1:]:
            if arr.shape != shape:
                return self.err("value", "SUMPRODUCT ranges have different sizes")
        total = 0
        for r in range(shape[0]):
            for c in range(shape[1]):
                prod = 1
                for arr in arrays:
                    item = arr.rows[r][c]
                    if isinstance(item, Err):
                        return item
                    prod *= item if is_number(item) else 0
                total += prod
        return total

    def fn_ISBLANK(self, args):
        v = self.scalar(self.arg(args, 0))
        return v is None

    def fn_ISNUMBER(self, args):
        v = self.scalar(self.arg(args, 0))
        return is_number(v)

    def fn_ISTEXT(self, args):
        v = self.scalar(self.arg(args, 0))
        return isinstance(v, str)

    def fn_ISERROR(self, args):
        v = self.scalar(self.arg(args, 0))
        return isinstance(v, Err) and v.code != "unsupported"

    def fn_N(self, args):
        v = self.scalar(self.arg(args, 0))
        if isinstance(v, Err):
            return v
        if is_number(v):
            return v
        if isinstance(v, bool):
            return 1 if v else 0
        return 0

    def fn_LEN(self, args):
        t = self.text(self.arg(args, 0))
        return t if isinstance(t, Err) else len(t)

    def fn_CONCATENATE(self, args):
        out = ""
        for arg in args:
            if arg is None:
                continue
            t = self.text(self.eval(arg))
            if isinstance(t, Err):
                return t
            out += t
        return out

    fn_CONCAT = fn_CONCATENATE


# ------------------------------------------------------------------ driver


def find_soffice():
    found = shutil.which("soffice")
    if found:
        return found
    mac = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if os.path.exists(mac):
        return mac
    return None


def check_workbook(path, engine="auto"):
    """Run both stages and return the result dict. engine: auto | python | libreoffice."""
    book = Book(path)
    blocks = load_snapshot(path)
    checker = Checker(book, blocks)
    checker.static()

    used_engine = "python"
    evaluated = 0
    note = None
    soffice = find_soffice() if engine in ("auto", "libreoffice") else None
    if engine == "libreoffice" and soffice is None:
        note = "LibreOffice not found; used the python engine"
    if soffice is not None:
        try:
            evaluated = checker.evaluate_libreoffice(soffice)
            used_engine = "libreoffice"
        except Exception as exc:
            note = "LibreOffice recalculation failed (%s); used the python engine" % exc
            soffice = None
    if soffice is None:
        evaluated = checker.evaluate_python()
        used_engine = "python"

    order = dict((name, i) for i, name in enumerate(book.sheet_names))

    def sort_key(f):
        m = CELL_RE.match(f["cell"])
        return (order.get(f["sheet"], 99), int(m.group(2)), column_index_from_string(m.group(1)),
                0 if f["severity"] == "error" else 1, f["code"])

    findings = sorted(checker.findings, key=sort_key)
    errors = sum(1 for f in findings if f["severity"] == "error")
    warnings = len(findings) - errors
    result = {
        "ok": True,
        "engine": used_engine,
        "evaluated": evaluated,
        "formulas": len(checker.asts),
        "findings": findings,
        "summary": {"errors": errors, "warnings": warnings},
    }
    if note:
        result["note"] = note
    return result


USAGE = "usage: python3 builder/check_formulas.py <workbook.xlsx> [<out.json>] [--engine auto|python|libreoffice]"


def main(argv):
    engine = "auto"
    args = []
    rest = list(argv[1:])
    while rest:
        item = rest.pop(0)
        if item in ("-h", "--help"):
            print(USAGE)
            return 0
        if item == "--engine":
            engine = rest.pop(0) if rest else ""
        elif item.startswith("--engine="):
            engine = item.split("=", 1)[1]
        else:
            args.append(item)
    if not args:
        print(USAGE)
        return 2
    if engine not in ("auto", "python", "libreoffice"):
        print(json.dumps({"ok": False, "error": "unknown engine %r" % engine}))
        return 1
    path = args[0]
    out_path = args[1] if len(args) > 1 else None
    try:
        result = check_workbook(path, engine)
    except Exception as exc:
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

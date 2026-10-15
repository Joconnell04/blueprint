#!/usr/bin/env python3
"""Tests for builder/check_workbook.py. Run: python3 builder/test_check_workbook.py

Every fixture is built with openpyxl directly and paired with a hand-written
snapshot, so the tests describe the finding rules rather than build_xlsx's
current output. The one exception is SampleWorkbookTests, which runs the
builder to get a snapshot for builder/sample_house.xlsx and skips when it
cannot.
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import build_xlsx                       # noqa: E402
import check_workbook as cw             # noqa: E402

STYLE = build_xlsx.PlainStyler()        # check_workbook's default when no spec is given
BLUE = Font(color=STYLE.input_color)
GREY_IT = Font(color=STYLE.note_color, italic=True)
CURRENCY = build_xlsx.FORMATS["currency"]
CURRENCY2 = build_xlsx.FORMATS["currency2"]
PERCENT = build_xlsx.FORMATS["percent"]
COUNT = build_xlsx.FORMATS["count"]


def widths(ws, mapping):
    for letter, width in mapping.items():
        ws.column_dimensions[letter].width = width


def save(wb, directory, name, snapshot):
    path = os.path.join(directory, name + ".xlsx")
    wb.save(path)
    snap_path = os.path.join(directory, name + ".snapshot.json")
    with open(snap_path, "w") as handle:
        json.dump(snapshot, handle)
    return path, snap_path


def block(bid, btype, row, col, rows, cols):
    return {"id": bid, "type": btype, "row": row, "col": col, "rows": rows, "cols": cols}


def define_name(wb, name, target):
    """A workbook-scoped defined name, across openpyxl 3.0 and 3.1."""
    defn = DefinedName(name, attr_text=target)
    try:
        wb.defined_names[name] = defn
    except TypeError:                   # openpyxl 3.0
        wb.defined_names.append(defn)


def read_me(wb, others):
    ws = wb.create_sheet("Read Me", 0)
    ws["A1"] = "Model cover"
    ws["A2"] = "Contents: " + ", ".join(others)
    ws["A3"] = "Blue cells are inputs. Change these."
    return {"name": "Read Me", "blocks": [], "cells": [], "widths": {}}


# ------------------------------------------------------------------- fixtures


def make_clean(directory):
    """A workbook that breaks none of the rules: the check must stay silent."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Model"
    widths(ws, {"B": 24, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14})

    ws["B2"] = "Inputs"
    ws["B3"], ws["C3"] = "Growth rate (%)", 0.10
    ws["B4"], ws["C4"] = "Base revenue ($M)", 100
    ws["C3"].number_format, ws["C3"].font = PERCENT, BLUE
    ws["C4"].number_format, ws["C4"].font = CURRENCY, BLUE

    ws["B6"] = "Revenue ramp"
    for col, year in zip("CDEFG", ("2026", "2027", "2028", "2029", "2030")):
        ws["%s7" % col] = year
    ws["B7"] = "Line"
    ws["B8"] = "Revenue ($M)"
    ws["C8"] = "=$C$4"
    for col, prev in (("D", "C"), ("E", "D"), ("F", "E"), ("G", "F")):
        ws["%s8" % col] = "=%s8*(1+$C$3)" % prev
    for col in "CDEFG":
        ws["%s8" % col].number_format = CURRENCY
    ws["B9"] = "Gross margin (%)"
    for col in "CDEFG":
        ws["%s9" % col] = 0.62
        ws["%s9" % col].number_format = PERCENT

    ws["B12"] = "Segments"
    ws["B13"], ws["C13"], ws["D13"] = "Segment", "Accounts", "Revenue ($M)"
    rows = (("Small", 100, 5), ("Mid", 200, 7), ("Total", 300, 12))
    for i, (label, accounts, revenue) in enumerate(rows):
        r = 14 + i
        ws.cell(r, 2).value = label
        ws.cell(r, 3).value = accounts
        ws.cell(r, 3).number_format = COUNT
        ws.cell(r, 4).value = revenue
        ws.cell(r, 4).number_format = CURRENCY

    ws["B18"] = "Outcome"
    ws["B19"], ws["C19"] = "Year 5 revenue ($M)", "=G8"
    ws["C19"].number_format = CURRENCY
    ws.freeze_panes = "C8"

    snapshot = {"sheets": [
        read_me(wb, ["Model"]),
        {"name": "Model", "widths": {}, "cells": [], "blocks": [
            block("b_in", "inputs", 2, 2, 3, 2),
            block("b_ramp", "calc", 6, 2, 4, 6),
            block("b_seg", "table", 12, 2, 5, 3),
            block("b_out", "output", 18, 2, 2, 2),
        ]},
    ]}
    return save(wb, directory, "clean", snapshot)


def make_period(directory):
    """A label column plus a period series. One row breaks the copy, nothing else."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ramp"
    widths(ws, {"B": 30, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14})

    ws["B2"] = "Inputs"
    pairs = (("Growth rate (%)", 0.10, PERCENT),
             ("Base revenue ($M)", 100, CURRENCY),
             ("Start penetration (%)", 0.01, PERCENT),
             ("Gross margin (%)", 0.70, PERCENT))
    for i, (label, value, fmt) in enumerate(pairs):
        r = 3 + i
        ws.cell(r, 2).value = label
        cell = ws.cell(r, 3)
        cell.value = value
        cell.number_format = fmt
        cell.font = BLUE

    ws["B8"] = "Revenue ramp"
    ws["B9"] = "Line"
    for col, year in zip("CDEFG", ("2026", "2027", "2028", "2029", "2030")):
        ws["%s9" % col] = year

    # A literal base and four copies: consistent.
    ws["B10"] = "Seats (count)"
    ws["C10"] = 1000
    for col, prev in (("D", "C"), ("E", "D"), ("F", "E"), ("G", "F")):
        ws["%s10" % col] = "=%s10*(1+$C$3)" % prev
    for col in "CDEFG":
        ws["%s10" % col].number_format = COUNT

    # A base formula and four copies: also consistent, the base is waived.
    ws["B11"] = "Revenue ($M)"
    ws["C11"] = "=$C$4"
    for col, prev in (("D", "C"), ("E", "D"), ("F", "E"), ("G", "F")):
        ws["%s11" % col] = "=%s11*(1+$C$3)" % prev
    for col in "CDEFG":
        ws["%s11" % col].number_format = CURRENCY

    # The violation: an exponent typed per column instead of a copied formula.
    ws["B12"] = "Penetration (%)"
    for i, col in enumerate("CDEFG", start=1):
        ws["%s12" % col] = "=$C$5*(1+$C$3)^%d" % i
        ws["%s12" % col].number_format = PERCENT
    ws["H12"] = "Ramp compounds off year one"
    ws["H12"].font = GREY_IT

    ws["B13"] = "Gross profit ($M)"
    for col in "CDEFG":
        ws["%s13" % col] = "=%s11*$C$6" % col
        ws["%s13" % col].number_format = CURRENCY

    # A spec table: its columns are different quantities, so the row rule must
    # not run over it. This is the shape that produced ten false positives.
    ws["B15"] = "Spec sheet"
    ws["B16"], ws["C16"], ws["D16"] = "Field", "Unit", "Low"
    ws["E16"], ws["F16"] = "High", "Source"
    ws["B17"], ws["C17"] = "Take rate", 0.02
    ws["C17"].number_format, ws["C17"].font = PERCENT, BLUE
    ws["D17"], ws["E17"], ws["F17"] = "=C17", "=C17*$C$3", "=SUM(C17:D17)"
    ws["B18"], ws["C18"], ws["F18"] = "Method", "n/a", "desk research"
    ws.freeze_panes = "C10"

    snapshot = {"sheets": [
        read_me(wb, ["Ramp"]),
        {"name": "Ramp", "widths": {}, "blocks": [
            block("b_in", "inputs", 2, 2, 5, 2),
            block("b_ramp", "calc", 8, 2, 6, 6),
            block("b_spec", "table", 15, 2, 4, 5),
        ], "cells": [{"a": "H12", "k": "note"}]},
    ]}
    return save(wb, directory, "period", snapshot)


def make_defects(directory):
    """One of every other kind, each on a cell a test can name."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    src = wb.create_sheet("Inputs")
    widths(src, {"B": 20, "C": 12})
    src["B2"] = "Inputs"
    src["B3"], src["C3"] = "Take rate (%)", 0.02
    src["C3"].number_format, src["C3"].font = PERCENT, BLUE
    src["B4"], src["C4"] = "Unused lever (%)", 0.50
    src["C4"].number_format, src["C4"].font = PERCENT, BLUE
    src["B5"], src["C5"] = "Named lever (%)", 0.30
    src["C5"].number_format, src["C5"].font = PERCENT, BLUE
    define_name(wb, "Named_lever", "'Inputs'!$C$5")
    define_name(wb, "Unused_lever", "'Inputs'!$C$4")

    ws = wb.create_sheet("Model")
    widths(ws, {"B": 14, "C": 8, "D": 8, "E": 8, "F": 8, "G": 8, "H": 10, "I": 8})

    ws["B2"] = "Problems"
    ws["B3"], ws["C3"], ws["D3"] = "Item", "Value", "Check"
    ws["B4"] = "Employers in the target region"          # clipped: 30 chars in a 14 column
    ws["C4"] = "='Inputs'!C3"                            # off-sheet pass-through, black
    ws["C4"].number_format = PERCENT
    ws["D4"] = "=C4*4.7"                                 # hardcode
    ws["D4"].number_format = COUNT
    ws["H4"] = "This note is much too long for the column it sits in"
    ws["H4"].font = GREY_IT
    ws["I4"] = 1                                         # blocks the note's spill
    ws["I4"].number_format = COUNT
    ws["B5"], ws["C5"], ws["D5"] = "Circular", "=D5+1", "=C5+1"
    ws["D5"].number_format = COUNT

    ws["B13"] = "Segments"
    ws["B14"], ws["C14"], ws["D14"] = "Seg", "Volume", "Revenue ($M)"
    for i, (label, volume, revenue) in enumerate(
            (("Small", 100, 5), ("Mid", 200, 7), ("Total", 250, 12))):
        r = 15 + i
        ws.cell(r, 2).value = label
        ws.cell(r, 3).value = volume
        ws.cell(r, 3).number_format = COUNT
        ws.cell(r, 4).value = revenue
        ws.cell(r, 4).number_format = CURRENCY

    ws["B19"] = "By year"
    ws["B20"] = "Line"
    for col, year in zip("CDEFG", ("2026", "2027", "2028", "2029", "2030")):
        ws["%s20" % col] = year
    ws["B21"] = "Volume ($M)"
    for i, col in enumerate("CDEFG"):
        ws["%s21" % col] = 10 * (i + 1)
        ws["%s21" % col].number_format = CURRENCY
    ws["B22"] = "Value ($M)"
    for i, col in enumerate("CDEFG"):
        ws["%s22" % col] = 10 * (i + 1) + 1
        ws["%s22" % col].number_format = CURRENCY2 if col == "C" else CURRENCY

    ws["B24"] = "Outcome"
    ws["B25"], ws["C25"] = "Dangling", "=C21+D21"
    ws["C25"].number_format = CURRENCY
    ws["B26"], ws["C26"] = "Broken sheet", "=Missing!A1"
    ws["B27"], ws["C27"] = "Broken ref", "=#REF!+1"
    ws["B29"] = "Source link"
    ws["B29"].hyperlink = "https://example.com"
    ws["B31"] = ("A paragraph much longer than its column that wraps, "
                 "so a reader still sees all of it")
    ws["B31"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["B33"], ws["C33"] = "External", "=[1]Sheet1!A1"
    ws["B34"], ws["C34"] = "By name", "=Named_lever*100"
    ws["H35"] = "Merged wide enough to fit"       # 25 chars over H:J = 26
    ws["H35"].font = GREY_IT
    ws.merge_cells(start_row=35, start_column=8, end_row=35, end_column=10)

    calc = wb.create_sheet("Calc")
    widths(calc, {"B": 14, "C": 14, "D": 14})
    calc["B2"] = "Magic"
    calc["B3"], calc["C3"], calc["D3"] = "Item", "Value ($M)", "Other ($M)"
    for i, (label, left, right) in enumerate(
            (("A", "=1+7", "=2*3"), ("B", "=1+8", "=100*4"), ("C", "=1+9", "=1000*5"))):
        r = 4 + i
        calc.cell(r, 2).value = label
        calc.cell(r, 3).value = left
        calc.cell(r, 4).value = right

    snapshot = {"sheets": [
        {"name": "Inputs", "widths": {}, "cells": [], "blocks": [
            block("b_src", "inputs", 2, 2, 4, 2),
        ]},
        {"name": "Model", "widths": {},
         "cells": [{"a": "H4", "k": "note"}, {"a": "H35", "k": "note"}], "blocks": [
            block("b_bad", "calc", 2, 2, 4, 3),
            block("b_tot", "table", 13, 2, 5, 3),
            block("b_period", "calc", 19, 2, 4, 6),
            block("b_out", "output", 24, 2, 2, 2),
        ]},
        {"name": "Calc", "widths": {}, "cells": [], "blocks": [
            block("b_magic", "calc", 2, 2, 5, 3),
        ]},
    ]}
    return save(wb, directory, "defects", snapshot)


class Base(unittest.TestCase):
    maker = None

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="check_workbook_test_")
        cls.path, cls.snap = cls.maker(cls.tmp)
        cls.result = cw.check_workbook(cls.path, cls.snap)
        cls.findings = cls.result["findings"]

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def of_kind(self, kind):
        return [f for f in self.findings if f["kind"] == kind]

    def of_code(self, code):
        return [f for f in self.findings if f["code"] == code]

    def cells(self, code):
        return sorted((f["sheet"], f["cell"]) for f in self.of_code(code))

    def one(self, code):
        hits = self.of_code(code)
        self.assertEqual(len(hits), 1,
                         "expected one %s, got %r" % (code, self.cells(code)))
        return hits[0]


# --------------------------------------------------------------------- helpers


class HelperTests(unittest.TestCase):
    def test_period_tag_years(self):
        for text in ("2026", "FY26", "FY2026", "2026E", "CY2027", 2026):
            self.assertEqual(cw.period_tag(text), "year", text)

    def test_period_tag_quarters_and_months(self):
        for text in ("Q1", "Q3 2026", "1Q24", "Q4'27"):
            self.assertEqual(cw.period_tag(text), "quarter", text)
        for text in ("Jan", "January", "Mar-26", "2026-03"):
            self.assertEqual(cw.period_tag(text), "month", text)

    def test_period_tag_rejects_labels(self):
        for text in ("Line", "Revenue ($M)", "", None, "Segment", "1.5", "2026 revenue"):
            self.assertIsNone(cw.period_tag(text), text)

    def test_to_r1c1_keeps_literals_and_absolutes(self):
        self.assertEqual(cw.to_r1c1("=C10*(1+$C$5)^2", 11, 4),
                         "=R[-1]C[-1]*(1+R5C3)^2")
        self.assertNotEqual(cw.to_r1c1("=C10*(1+$C$5)^1", 11, 4),
                            cw.to_r1c1("=C10*(1+$C$5)^2", 11, 4))

    def test_to_r1c1_normalises_a_copied_row(self):
        left = cw.to_r1c1("=C8*(1+$C$3)", 8, 4)
        right = cw.to_r1c1("=D8*(1+$C$3)", 8, 5)
        self.assertEqual(left, right)

    def test_to_r1c1_keeps_sheets_and_names(self):
        self.assertEqual(cw.to_r1c1("='Key Assumptions'!C3", 5, 3),
                         "='Key Assumptions'!R[-2]C[0]")
        self.assertEqual(cw.to_r1c1("=Take_rate*2", 5, 3), "=Take_rate*2")

    def test_format_family_groups_spellings(self):
        self.assertEqual(cw.format_family(CURRENCY), "currency")
        self.assertEqual(cw.format_family(CURRENCY2), "currency")
        self.assertEqual(cw.format_family(PERCENT), "percent")
        self.assertEqual(cw.format_family('0.0"x"'), "multiple")
        self.assertEqual(cw.format_family("General"), "number")
        self.assertEqual(cw.format_family(COUNT), "number")

    def test_bare_format(self):
        self.assertTrue(cw.is_bare_format("General"))
        self.assertTrue(cw.is_bare_format(COUNT))
        self.assertFalse(cw.is_bare_format(CURRENCY))
        self.assertFalse(cw.is_bare_format(PERCENT))

    def test_has_unit(self):
        for text in ("Revenue ($M)", "Penetration (%)", "Multiple (x)", "Seats",
                     "Accounts", "Hours per week", "Churn rate"):
            self.assertTrue(cw.has_unit(text), text)
        for text in ("Volume", "Widgets", "Employers in the target region", "Check", None):
            self.assertFalse(cw.has_unit(text), text)

    def test_formula_numbers_and_refs(self):
        self.assertEqual(cw.formula_numbers("=C4*4.7+1"), [4.7, 1.0])
        refs = cw.formula_refs("=SUM('Inputs'!C2:C3)+D1")
        self.assertEqual([r[1] for r in refs], ["Inputs", None])
        self.assertEqual(refs[0][2], (2, 3, 3, 3))


class WidthModelTests(unittest.TestCase):
    """The model must be build_xlsx's, not a copy of its numbers."""

    def setUp(self):
        self.model = cw.WidthModel(build_xlsx.HouseStyler())

    def test_need_matches_the_builder(self):
        ctx = build_xlsx.Ctx(build_xlsx.HouseStyler(), {}, [], [])
        ctx.fit(1, "Serviceable share of all revenue")
        self.assertEqual(self.model.need("Serviceable share of all revenue"),
                         ctx.fits[1])

    def test_cap_is_the_builders_cap(self):
        ctx = build_xlsx.Ctx(build_xlsx.HouseStyler(), {}, [], [])
        ctx.fit(1, "x" * 400)
        self.assertEqual(self.model.cap(), ctx.final_widths()[1])

    def test_bold_allowance_is_never_below_one(self):
        """The builder may widen bold text or wrap it; either way this stays sane."""
        self.assertGreaterEqual(self.model.bold_allowance(), 1.0)
        self.assertLess(self.model.bold_allowance(), 2.0)

    def test_display_is_the_builders_formatter(self):
        self.assertEqual(self.model.display(52000000, CURRENCY), "$52,000,000")
        self.assertEqual(self.model.display(0.421, PERCENT), "42.1%")
        self.assertIsNone(self.model.display("text", COUNT))


class TrimNumberTests(unittest.TestCase):
    def test_trim(self):
        self.assertEqual(cw._trim_number(300.0), "300")
        self.assertEqual(cw._trim_number(1234567.0), "1,234,567")
        self.assertEqual(cw._trim_number(4.7), "4.7")


# ----------------------------------------------------------------- clean file


class CleanTests(Base):
    maker = staticmethod(make_clean)

    def test_nothing_is_reported(self):
        self.assertEqual(self.findings, [], "a clean workbook must stay silent")

    def test_summary_is_zero(self):
        self.assertEqual(self.result["summary"], {"errors": 0, "warnings": 0})
        self.assertEqual(self.result["byKind"], {})

    def test_it_still_read_the_workbook(self):
        self.assertTrue(self.result["ok"])
        self.assertEqual(self.result["sheets"], 2)
        self.assertGreater(self.result["cells"], 30)
        self.assertGreater(self.result["formulas"], 4)

    def test_a_reconciling_total_is_not_reported(self):
        self.assertEqual(self.of_code("total-mismatch"), [])


# --------------------------------------------------------- period series file


class PeriodTests(Base):
    maker = staticmethod(make_period)

    def test_one_true_positive_and_no_noise(self):
        hits = self.of_kind("row-inconsistent")
        self.assertEqual(len(hits), 1, [(f["cell"], f["message"]) for f in hits])
        self.assertEqual(hits[0]["cell"], "C12")
        self.assertEqual(hits[0]["blockId"], "b_ramp")
        self.assertIn("C12:G12", hits[0]["message"])

    def test_it_is_the_only_finding_in_the_file(self):
        self.assertEqual([(f["sheet"], f["cell"], f["code"]) for f in self.findings],
                         [("Ramp", "C12", "row-inconsistent")])

    def test_a_literal_base_row_is_consistent(self):
        self.assertNotIn("C10", [f["cell"] for f in self.of_kind("row-inconsistent")])

    def test_a_base_formula_is_waived(self):
        self.assertNotIn("C11", [f["cell"] for f in self.of_kind("row-inconsistent")])

    def test_the_spec_table_is_out_of_scope(self):
        with open(self.snap) as handle:
            sheet = cw.Checker(self.path, json.load(handle)).sheets[1]
        spec = [b for b in sheet.blocks if b.id == "b_spec"][0]
        ramp = [b for b in sheet.blocks if b.id == "b_ramp"][0]
        self.assertEqual(sheet.header_periods(spec), {})
        self.assertEqual(len(sheet.header_periods(ramp)), 5)

    def test_a_noted_row_waives_its_hardcodes(self):
        self.assertEqual(self.of_code("hardcode"), [])

    def test_a_note_spills_over_empty_columns(self):
        self.assertEqual(self.of_code("note-clipped"), [])

    def test_severity_and_fix_are_set(self):
        finding = self.one("row-inconsistent")
        self.assertEqual(finding["severity"], "warning")
        self.assertTrue(finding["fix"].endswith("."))
        self.assertIn("period row", finding["fix"])


# ---------------------------------------------------------------- defect file


class DefectTests(Base):
    maker = staticmethod(make_defects)

    def test_clipped_label(self):
        finding = self.one("clipped")
        self.assertEqual((finding["sheet"], finding["cell"]), ("Model", "B4"))
        self.assertEqual(finding["blockId"], "b_bad")
        self.assertIn("Employers in the target region", finding["message"])
        self.assertIn("Widen the column", finding["fix"])

    def test_clipped_note(self):
        hits = self.of_code("note-clipped")
        self.assertEqual(len(hits), 1, [f["cell"] for f in hits])
        finding = hits[0]
        self.assertEqual((finding["sheet"], finding["cell"]), ("Model", "H4"))
        self.assertEqual(finding["kind"], "note-clipped")
        self.assertIn("Merge the note", finding["fix"])

    def test_hardcode(self):
        hits = [f for f in self.of_code("hardcode") if f["sheet"] == "Model"]
        self.assertEqual([f["cell"] for f in hits], ["D4"])
        self.assertIn("4.7", hits[0]["message"])

    def test_unit_scalars_are_not_hardcodes(self):
        for cell in ("C5", "D5", "C27"):
            self.assertNotIn(cell, [f["cell"] for f in self.of_code("hardcode")
                                    if f["sheet"] == "Model"])

    def test_magic_density(self):
        finding = self.one("magic-density")
        self.assertEqual(finding["sheet"], "Calc")
        self.assertEqual(finding["cell"], "")
        self.assertEqual(finding["scope"], "sheet")
        self.assertIn("6 of 6", finding["message"])

    def test_mixed_format(self):
        finding = self.one("mixed-format")
        self.assertEqual((finding["sheet"], finding["cell"]), ("Model", "C22"))
        self.assertIn(CURRENCY2, finding["message"])
        self.assertEqual(finding["blockId"], "b_period")

    def test_mixed_format_ignores_different_quantities(self):
        """A period column holding currency beside percent is correct, not mixed."""
        for cell in ("C21", "D21", "D22", "E22"):
            self.assertNotIn(cell, [f["cell"] for f in self.of_code("mixed-format")])

    def test_circular(self):
        hits = self.of_code("circular")
        self.assertEqual([f["cell"] for f in hits], ["C5", "D5"])
        for f in hits:
            self.assertEqual(f["severity"], "error")
            self.assertEqual(f["kind"], "ref-colour")
            self.assertIn("Model!C5", f["message"])

    def test_ref_error(self):
        hits = self.of_code("ref-error")
        self.assertEqual(sorted(f["cell"] for f in hits), ["C26", "C27", "C33"])
        for f in hits:
            self.assertEqual(f["severity"], "error")
        messages = " ".join(f["message"] for f in hits)
        self.assertIn("Missing", messages)
        self.assertIn("#REF!", messages)
        self.assertIn("another workbook", messages)

    def test_a_wrapped_cell_is_not_clipped(self):
        """A wrapped paragraph grows down, not across, so it shows its whole value."""
        self.assertNotIn(("Model", "B31"), self.cells("clipped"))

    def test_a_merge_wide_enough_is_not_clipped(self):
        self.assertNotIn(("Model", "H35"), self.cells("note-clipped"))
        self.assertNotIn(("Model", "H35"), self.cells("clipped"))

    def test_off_sheet_pass_through(self):
        finding = self.one("off-sheet-uncoloured")
        self.assertEqual((finding["sheet"], finding["cell"]), ("Model", "C4"))
        self.assertEqual(finding["severity"], "warning")
        self.assertIn("'Inputs'!C3", finding["message"])

    def test_uncoloured_hyperlink(self):
        finding = self.one("link-uncoloured")
        self.assertEqual((finding["sheet"], finding["cell"]), ("Model", "B29"))

    def test_no_unit(self):
        self.assertEqual(self.cells("no-unit"), [("Model", "C14"), ("Model", "D3")])

    def test_no_unit_skips_a_column_whose_format_carries_it(self):
        self.assertNotIn(("Model", "D14"), self.cells("no-unit"))

    def test_empty_band(self):
        finding = self.one("empty-band")
        self.assertEqual(finding["sheet"], "Model")
        self.assertEqual(finding["scope"], "sheet")
        self.assertIn("6 to 12", finding["message"])

    def test_no_freeze(self):
        finding = self.one("no-freeze")
        self.assertEqual(finding["sheet"], "Model")
        self.assertIn("C4", finding["message"])

    def test_total_mismatch(self):
        finding = self.one("total-mismatch")
        self.assertEqual((finding["sheet"], finding["cell"]), ("Model", "C17"))
        self.assertEqual(finding["severity"], "error")
        self.assertIn("250", finding["message"])
        self.assertIn("300", finding["message"])

    def test_a_reconciling_column_is_quiet(self):
        self.assertNotIn(("Model", "D17"), self.cells("total-mismatch"))

    def test_unused_input(self):
        finding = self.one("unused-input")
        self.assertEqual((finding["sheet"], finding["cell"]), ("Inputs", "C4"))
        self.assertIn("Unused lever", finding["message"])

    def test_a_defined_name_does_not_excuse_an_unused_input(self):
        """C4 carries a defined name and is still unused: the name is not a reader."""
        self.assertIn(("Inputs", "C4"), self.cells("unused-input"))

    def test_an_input_read_through_its_name_is_used(self):
        self.assertNotIn(("Inputs", "C5"), self.cells("unused-input"))

    def test_dangling_output(self):
        finding = self.one("dangling-output")
        self.assertEqual((finding["sheet"], finding["cell"]), ("Model", "C25"))
        self.assertEqual(finding["blockId"], "b_out")

    def test_no_cover(self):
        finding = self.one("no-cover")
        self.assertEqual(finding["sheet"], "")
        self.assertEqual(finding["cell"], "")
        self.assertEqual(finding["scope"], "workbook")
        self.assertIn("no Read Me tab", finding["message"])

    def test_errors_are_only_reconciliation_and_ref_class(self):
        codes = set(f["code"] for f in self.findings if f["severity"] == "error")
        self.assertEqual(codes, {"total-mismatch", "circular", "ref-error"})

    def test_every_finding_has_the_check_formulas_shape(self):
        for f in self.findings:
            self.assertEqual(set(f) - {"blockId", "scope"},
                             {"sheet", "cell", "code", "severity", "message", "kind", "fix"})
            self.assertIn(f["severity"], ("error", "warning"))
            self.assertIn(f["kind"], cw.KIND_ORDER)
            self.assertTrue(f["fix"])

    def test_findings_are_sorted_by_sheet_then_cell(self):
        order = dict((name, i) for i, name in enumerate(["Inputs", "Model", "Calc"]))
        seen = [order.get(f["sheet"], 99) for f in self.findings]
        self.assertEqual(seen, sorted(seen))
        rows = [int(f["cell"][1:]) for f in self.findings
                if f["sheet"] == "Model" and f["cell"]]
        self.assertEqual(rows, sorted(rows))

    def test_by_kind_counts_match_the_findings(self):
        counted = {}
        for f in self.findings:
            counted[f["kind"]] = counted.get(f["kind"], 0) + 1
        self.assertEqual(self.result["byKind"], counted)
        self.assertEqual(sum(counted.values()), len(self.findings))

    def test_every_documented_kind_is_reachable(self):
        """Between the three fixtures, each of the fourteen kinds fires."""
        seen = set(self.result["byKind"])
        tmp = tempfile.mkdtemp(prefix="check_workbook_kinds_")
        try:
            path, snap = make_period(tmp)
            seen |= set(cw.check_workbook(path, snap)["byKind"])
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
        self.assertEqual(sorted(seen), sorted(cw.KIND_ORDER))


# ------------------------------------------------------------- missing inputs


class MissingFileTests(unittest.TestCase):
    def test_a_missing_snapshot_notes_and_keeps_going(self):
        tmp = tempfile.mkdtemp(prefix="check_workbook_nosnap_")
        try:
            path, _snap = make_defects(tmp)
            result = cw.check_workbook(path, os.path.join(tmp, "absent.snapshot.json"))
            self.assertTrue(result["ok"])
            self.assertIn("no snapshot", result["note"])
            # Without the snapshot's cell kinds the note reads as an ordinary cell,
            # and without block types the block-scoped kinds cannot run.
            self.assertEqual(result["byKind"].get("clipped"), 2)
            self.assertNotIn("note-clipped", result["byKind"])
            self.assertNotIn("dangling-output", result["byKind"])
            self.assertNotIn("mixed-format", result["byKind"])
            self.assertEqual(result["byKind"].get("no-cover"), 1)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_no_snapshot_argument_at_all(self):
        tmp = tempfile.mkdtemp(prefix="check_workbook_bare_")
        try:
            path, _snap = make_clean(tmp)
            result = cw.check_workbook(path)
            self.assertTrue(result["ok"])
            self.assertEqual(result["summary"]["errors"], 0)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


# --------------------------------------------------------------- sample house


class SampleWorkbookTests(unittest.TestCase):
    """sample_spec.json built fresh, so the file and its snapshot agree.

    builder/sample_house.xlsx has no committed snapshot and is rebuilt by hand,
    so checking it against a snapshot from today's builder measures drift, not
    craft. The build goes to a temp directory: it never touches the sample.
    """

    @classmethod
    def build(cls, directory):
        spec_path = os.path.join(HERE, "sample_spec.json")
        if not os.path.exists(spec_path):
            return None, None, None
        with open(spec_path) as handle:
            spec = json.load(handle)
        spec["outputPath"] = os.path.join(directory, "sample_house.xlsx")
        scratch_spec = os.path.join(directory, "spec.json")
        with open(scratch_spec, "w") as handle:
            json.dump(spec, handle)
        proc = subprocess.run(
            [sys.executable, os.path.join(HERE, "build_xlsx.py"), scratch_spec],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=300)
        snap = os.path.join(directory, "sample_house.snapshot.json")
        if proc.returncode != 0 or not os.path.exists(snap):
            return None, None, None
        return spec["outputPath"], snap, scratch_spec

    def test_the_sample_builds_clean(self):
        tmp = tempfile.mkdtemp(prefix="check_workbook_sample_")
        try:
            path, snap, spec = self.build(tmp)
            if path is None:
                self.skipTest("build_xlsx.py could not build sample_spec.json")
            result = cw.check_workbook(path, snap, spec)
            self.assertTrue(result["ok"])
            self.assertEqual(result["sheets"], 5)
            self.assertEqual(result["summary"]["errors"], 0)
            for kind in ("clipped", "note-clipped", "total-mismatch", "no-cover",
                         "empty-band", "no-freeze", "mixed-format", "no-unit"):
                self.assertEqual(result["byKind"].get(kind, 0), 0,
                                 "%s on the sample: %r" % (kind, [
                                     (f["sheet"], f["cell"], f["message"])
                                     for f in result["findings"] if f["kind"] == kind]))
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_the_sample_needs_no_spec_to_check(self):
        tmp = tempfile.mkdtemp(prefix="check_workbook_sample_nospec_")
        try:
            path, snap, _spec = self.build(tmp)
            if path is None:
                self.skipTest("build_xlsx.py could not build sample_spec.json")
            result = cw.check_workbook(path, snap)
            self.assertTrue(result["ok"])
            self.assertEqual(result["summary"]["errors"], 0)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


# ------------------------------------------------------------------------ cli


class CliTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="check_workbook_cli_")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def run_cli(self, *args):
        return subprocess.run(
            [sys.executable, os.path.join(HERE, "check_workbook.py")] + list(args),
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)

    def test_cli_writes_json_and_exits_zero(self):
        path, snap = make_defects(self.tmp)
        out = os.path.join(self.tmp, "findings.json")
        proc = self.run_cli(path, snap, "--out", out)
        self.assertEqual(proc.returncode, 0, proc.stderr.decode())
        printed = json.loads(proc.stdout.decode())
        with open(out) as handle:
            written = json.load(handle)
        self.assertEqual(printed, written)
        self.assertTrue(printed["ok"])
        self.assertGreater(printed["summary"]["errors"], 0)

    def test_cli_takes_a_spec_positionally(self):
        path, snap = make_clean(self.tmp)
        spec = os.path.join(self.tmp, "spec.json")
        with open(spec, "w") as handle:
            json.dump({"formatting": "house", "sheets": []}, handle)
        proc = self.run_cli(path, snap, spec)
        self.assertEqual(proc.returncode, 0, proc.stderr.decode())
        self.assertTrue(json.loads(proc.stdout.decode())["ok"])

    def test_cli_crash_exits_one(self):
        proc = self.run_cli("/nonexistent/file.xlsx", "/nonexistent/snap.json")
        self.assertEqual(proc.returncode, 1)
        self.assertFalse(json.loads(proc.stdout.decode())["ok"])

    def test_cli_with_no_arguments_prints_usage(self):
        proc = self.run_cli()
        self.assertEqual(proc.returncode, 2)
        self.assertIn("usage:", proc.stdout.decode())

    def test_cli_help(self):
        proc = self.run_cli("--help")
        self.assertEqual(proc.returncode, 0)
        self.assertIn("snapshot.json", proc.stdout.decode())


if __name__ == "__main__":
    unittest.main()

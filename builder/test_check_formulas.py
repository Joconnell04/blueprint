#!/usr/bin/env python3
"""Tests for builder/check_formulas.py. Run: python3 builder/test_check_formulas.py

Uses the python engine throughout so the run stays fast and does not depend
on LibreOffice being installed.
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import check_formulas as cf  # noqa: E402


def make_fixture(directory):
    """A two-sheet workbook with one clean model and one of each defect."""
    wb = openpyxl.Workbook()
    inputs = wb.active
    inputs.title = "Inputs"
    inputs["B2"] = "Price"
    inputs["C2"] = 100
    inputs["B3"] = "Units"
    inputs["C3"] = 5
    inputs["B4"] = "Zero"
    inputs["C4"] = 0
    inputs["B5"] = "Label"
    inputs["C5"] = "ten"

    calc = wb.create_sheet("Calc")
    calc["A1"] = "Revenue"
    calc["B1"] = "='Inputs'!C2*'Inputs'!C3"           # clean cross-sheet: 500
    calc["B2"] = "=SUM(B1:B1)/'Inputs'!C4"              # #DIV/0!
    calc["B3"] = "=Missing!A1"                          # missing sheet
    calc["B4"] = "=D4*2"                                # D4 is empty inside the used area
    calc["B5"] = "=B6+1"                                # circular pair
    calc["B6"] = "=B5+1"
    calc["B7"] = "=XLOOKUP(1,B1:B2,B1:B2)"              # unsupported by the python engine
    calc["B8"] = "=IF(B1>100,ROUND(B1/3,2),MAX(B1,1))"  # clean: 166.67
    calc["B9"] = "=IFERROR(1/0,5)+AVERAGE(B1,B8)"       # clean: 5 + 333.335
    calc["B10"] = "=SUMPRODUCT(Inputs!C2:C3,Inputs!C2:C3)"  # clean: 10025
    calc["B11"] = "=B2*2"                               # inherits #DIV/0! from B2
    calc["B12"] = "='Inputs'!C5*2"                      # text in arithmetic: #VALUE!
    calc["B13"] = "=#REF!+1"                            # literal #REF!
    calc["B14"] = "=SUM(D1:D30)"                        # range past the used area
    calc["B15"] = "=-B1^2+B1%&\"x\""                    # clean: (-500)^2 + 5, then & "x"
    calc["B16"] = "=AND(B1>1,NOT(B1<1),OR(FALSE,TRUE))"  # clean: TRUE
    calc["B17"] = "=ROUNDUP(B1/7,0)-ROUNDDOWN(B1/7,0)+INT(-1.5)+ABS(-2)+MIN(B1,B8)+COUNT(B8:B10)+COUNTA(A1:A3)"
    calc["F18"] = "end"                                 # extends the used area to F18

    path = os.path.join(directory, "fixture.xlsx")
    wb.save(path)
    snapshot = {
        "sheets": [
            {"name": "Calc", "blocks": [
                {"id": "b_calc", "row": 1, "col": 1, "rows": 6, "cols": 2, "type": "calc"},
                {"id": "b_more", "row": 7, "col": 1, "rows": 12, "cols": 2, "type": "calc"},
            ]},
            {"name": "Inputs", "blocks": [
                {"id": "b_in", "row": 2, "col": 2, "rows": 4, "cols": 2, "type": "inputs"},
            ]},
        ]
    }
    with open(os.path.join(directory, "fixture.snapshot.json"), "w") as handle:
        json.dump(snapshot, handle)
    return path


class FixtureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="check_formulas_test_")
        cls.path = make_fixture(cls.tmp)
        cls.result = cf.check_workbook(cls.path, engine="python")
        cls.by_cell = {}
        for f in cls.result["findings"]:
            cls.by_cell.setdefault((f["sheet"], f["cell"]), []).append(f)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def codes(self, cell, sheet="Calc"):
        return sorted(f["code"] for f in self.by_cell.get((sheet, cell), []))

    def one(self, cell, code, sheet="Calc"):
        matches = [f for f in self.by_cell.get((sheet, cell), []) if f["code"] == code]
        self.assertEqual(len(matches), 1, "expected one %s on %s, got %r" % (code, cell, self.codes(cell, sheet)))
        return matches[0]

    def test_shape(self):
        r = self.result
        self.assertTrue(r["ok"])
        self.assertEqual(r["engine"], "python")
        self.assertEqual(r["evaluated"], 17)
        self.assertEqual(set(r["summary"]), {"errors", "warnings"})
        for f in r["findings"]:
            self.assertEqual(set(f) - {"blockId", "propagated"}, {"sheet", "cell", "severity", "code", "message", "formula"})
            self.assertIn(f["severity"], ("error", "warning"))
            self.assertTrue(f["formula"].startswith("="))

    def test_clean_cells_have_no_findings(self):
        for cell in ("B1", "B8", "B9", "B10", "B15", "B16", "B17"):
            self.assertEqual(self.codes(cell), [], "%s should be clean" % cell)

    def test_div0(self):
        f = self.one("B2", "div0")
        self.assertEqual(f["severity"], "error")
        self.assertNotIn("propagated", f)
        self.assertEqual(f["formula"], "=SUM(B1:B1)/'Inputs'!C4")

    def test_div0_propagates_to_dependents(self):
        f = self.one("B11", "div0")
        self.assertTrue(f["propagated"])
        self.assertIn("B2", f["message"])

    def test_missing_sheet(self):
        f = self.one("B3", "missing-sheet")
        self.assertEqual(f["severity"], "error")
        self.assertIn("Missing", f["message"])
        self.assertEqual(self.codes("B3"), ["missing-sheet"], "no second finding restating the same cell")

    def test_empty_ref(self):
        f = self.one("B4", "empty-ref")
        self.assertEqual(f["severity"], "warning")
        self.assertIn("D4", f["message"])

    def test_circular_pair(self):
        for cell in ("B5", "B6"):
            f = self.one(cell, "circular")
            self.assertEqual(f["severity"], "error")
            self.assertIn("B5", f["message"])
            self.assertIn("B6", f["message"])

    def test_unsupported_function(self):
        f = self.one("B7", "unsupported")
        self.assertEqual(f["severity"], "warning")
        self.assertIn("XLOOKUP", f["message"])

    def test_value_error(self):
        f = self.one("B12", "value")
        self.assertEqual(f["severity"], "error")

    def test_ref_literal(self):
        f = self.one("B13", "ref")
        self.assertEqual(f["severity"], "error")

    def test_outside_used_area(self):
        f = self.one("B14", "outside-used-area")
        self.assertEqual(f["severity"], "warning")

    def test_block_mapping(self):
        self.assertEqual(self.one("B2", "div0")["blockId"], "b_calc")
        self.assertEqual(self.one("B7", "unsupported")["blockId"], "b_more")

    def test_summary_counts(self):
        errors = sum(1 for f in self.result["findings"] if f["severity"] == "error")
        warnings = sum(1 for f in self.result["findings"] if f["severity"] == "warning")
        self.assertEqual(self.result["summary"], {"errors": errors, "warnings": warnings})
        self.assertEqual(errors, 7)   # div0 x2, missing-sheet, circular x2, value, ref
        self.assertEqual(warnings, 3)  # empty-ref, unsupported, outside-used-area

    def test_findings_sorted_by_sheet_then_cell(self):
        cells = [(f["sheet"], f["cell"]) for f in self.result["findings"]]
        rows = [int(c[1:]) for s, c in cells if s == "Calc"]
        self.assertEqual(rows, sorted(rows))


class EngineValueTests(unittest.TestCase):
    """The python engine agrees with Excel on the fixture's clean cells."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="check_formulas_engine_")
        path = make_fixture(cls.tmp)
        book = cf.Book(path)
        checker = cf.Checker(book, {})
        checker.static()
        cls.engine = cf.Engine(book, checker.asts, checker.circular)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def value(self, cell):
        area = cf.parse_area(cell)
        return self.engine.cell_value(("Calc", area[0], area[1]))

    def test_values(self):
        self.assertEqual(self.value("B1"), 500)
        self.assertAlmostEqual(self.value("B8"), 166.67)
        self.assertAlmostEqual(self.value("B9"), 5 + (500 + 166.67) / 2)
        self.assertEqual(self.value("B10"), 100 * 100 + 5 * 5)
        self.assertEqual(self.value("B15"), "250005x")   # negation binds before ^, % before +
        self.assertIs(self.value("B16"), True)
        # ROUNDUP(71.43)=72, ROUNDDOWN=71, INT(-1.5)=-2, ABS=2, MIN=166.67, COUNT(B8:B10)=3, COUNTA(A1:A3)=1
        self.assertAlmostEqual(self.value("B17"), 72 - 71 - 2 + 2 + 166.67 + 3 + 1)

    def test_errors_carry_origin(self):
        v = self.value("B11")
        self.assertIsInstance(v, cf.Err)
        self.assertEqual(v.code, "div0")
        self.assertEqual(v.origin, ("Calc", 2, 2))

    def test_circular_cells_read_as_zero(self):
        self.assertEqual(self.value("B5"), 0)


class SampleWorkbookTests(unittest.TestCase):
    def test_sample_house_has_no_errors(self):
        sample = os.path.join(HERE, "sample_house.xlsx")
        if not os.path.exists(sample):
            subprocess.run([sys.executable, os.path.join(HERE, "build_xlsx.py"),
                            os.path.join(HERE, "sample_spec.json")],
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, timeout=30)
        result = cf.check_workbook(sample, engine="python")
        self.assertTrue(result["ok"])
        self.assertEqual(result["engine"], "python")
        self.assertGreater(result["evaluated"], 30)
        errors = [f for f in result["findings"] if f["severity"] == "error"]
        self.assertEqual(errors, [])
        self.assertEqual(result["summary"]["errors"], 0)


class CliTests(unittest.TestCase):
    def test_cli_writes_json_and_exits_zero(self):
        tmp = tempfile.mkdtemp(prefix="check_formulas_cli_")
        try:
            path = make_fixture(tmp)
            out = os.path.join(tmp, "out.json")
            proc = subprocess.run(
                [sys.executable, os.path.join(HERE, "check_formulas.py"), path, out, "--engine", "python"],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
            self.assertEqual(proc.returncode, 0, proc.stderr.decode())
            printed = json.loads(proc.stdout.decode())
            with open(out) as handle:
                written = json.load(handle)
            self.assertEqual(printed, written)
            self.assertTrue(printed["ok"])
            self.assertEqual(printed["summary"]["errors"], 7)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_cli_crash_exits_one(self):
        proc = subprocess.run(
            [sys.executable, os.path.join(HERE, "check_formulas.py"), "/nonexistent/file.xlsx"],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
        self.assertEqual(proc.returncode, 1)
        self.assertFalse(json.loads(proc.stdout.decode())["ok"])


if __name__ == "__main__":
    unittest.main()

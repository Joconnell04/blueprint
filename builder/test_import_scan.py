#!/usr/bin/env python3
"""Tests for import_scan.py. Run: python3 builder/test_import_scan.py

Builds the corpus (make_corpus.py) into a temp dir, scans every workbook
through the CLI, scores each expected block by type with an error of boundary
(EoB) of at most one cell on every edge, and checks the scanner-only fields
the research doc calls out. sample_house.xlsx is the regression anchor: zero
unclassified regions and every spec block found with its type.
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import build_xlsx  # noqa: E402
import make_corpus  # noqa: E402

SCANNER = os.path.join(HERE, "import_scan.py")
SAMPLE_SPEC = os.path.join(HERE, "sample_spec.json")
SAMPLE_XLSX = os.path.join(HERE, "sample_house.xlsx")
EOB = 1

# Minimum recall per corpus workbook (matched expected blocks / expected blocks).
# All seven currently score 1.0 with 1.0 precision; the numbers are the floor.
MIN_SCORE = {
    "plain_financial.xlsx": 1.0,
    "data_dump.xlsx": 1.0,
    "side_by_side.xlsx": 1.0,
    "chart_heavy.xlsx": 1.0,
    "messy.xlsx": 1.0,
    "listobject.xlsx": 1.0,
    "hidden_frozen.xlsx": 1.0,
}


def run_scanner(src, out):
    proc = subprocess.run([sys.executable, SCANNER, src, out], capture_output=True, text=True)
    lines = [ln for ln in proc.stdout.splitlines() if ln.strip()]
    return proc.returncode, json.loads(lines[-1]) if lines else None, proc.stderr


def edges(b):
    return b["row"], b["col"], b["row"] + b["rows"] - 1, b["col"] + b["cols"] - 1


def within(a, b, tol=EOB):
    return all(abs(x - y) <= tol for x, y in zip(edges(a), edges(b)))


def match(expected, plan):
    """Pair expected blocks with detected blocks of the same sheet and type, EoB <= 1."""
    by_sheet = {t["name"]: list(t["blocks"]) for t in plan["tabs"]}
    hits, misses = [], []
    for e in expected:
        pool = [b for b in by_sheet.get(e["sheet"], []) if b["type"] == e["type"] and within(e, b)]
        if pool:
            b = pool[0]
            by_sheet[e["sheet"]].remove(b)
            hits.append((e, b))
        else:
            misses.append(e)
    return hits, misses


class CorpusTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="import_scan_")
        cls.expected = make_corpus.make(cls.tmp)
        cls.plans = {}
        for name in cls.expected:
            code, result, err = run_scanner(os.path.join(cls.tmp, name), os.path.join(cls.tmp, name + ".json"))
            if code != 0:
                raise RuntimeError("%s: %s %s" % (name, result, err))
            with open(os.path.join(cls.tmp, name + ".json")) as fh:
                cls.plans[name] = (result, json.load(fh))

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def check_fields(self, name, e, b, plan):
        where = "%s %s %s@%s" % (name, e["sheet"], e["type"], b["range"])
        self.assertGreaterEqual(b["confidence"], e.get("min_conf", 0.0), where)
        if "max_conf" in e:
            self.assertLessEqual(b["confidence"], e["max_conf"], where)
        for key in ("note", "labels", "header_rows", "hidden_rows", "hidden_cols", "subtotals", "dataRows"):
            if key in e:
                self.assertEqual(b.get(key), e[key], where + " " + key)
        if e.get("has_total"):
            self.assertIn("totalRow", b, where)
        if "sample_len" in e:
            self.assertEqual(len(b["sample"]), e["sample_len"], where)
        signals = b["detected"]["signals"]
        if e.get("listobject"):
            self.assertIn("listobject", signals, where)
        if e.get("bridged"):
            self.assertIn("bridged", signals, where)
        if "chart_kind" in e:
            self.assertEqual(b["chart"]["kind"], e["chart_kind"], where)
            self.assertIsNotNone(b["chart"]["sourceBlock"], where)
        if "series_cols" in e:
            self.assertEqual(b["chart"]["seriesCols"], e["series_cols"], where)
        if "source_sheet" in e:
            self.assertEqual(b["chart"]["sourceSheet"], e["source_sheet"], where)
        if "source_type" in e:
            src = [x for t in plan["tabs"] for x in t["blocks"] if x["id"] == b["chart"]["sourceBlock"]]
            self.assertEqual(src[0]["type"], e["source_type"], where)

    def test_corpus_scores(self):
        report = []
        for name, exp in self.expected.items():
            result, plan = self.plans[name]
            hits, misses = match(exp["blocks"], plan)
            detected = sum(len(t["blocks"]) for t in plan["tabs"])
            score = len(hits) / float(len(exp["blocks"]))
            precision = len(hits) / float(detected or 1)
            report.append("%-22s recall %.2f precision %.2f unclassified %d" % (name, score, precision, result["unclassified"]))
            self.assertGreaterEqual(score, MIN_SCORE[name], "%s missed %s" % (name, misses))
            self.assertEqual(precision, 1.0, "%s: %d detected, %d expected" % (name, detected, len(exp["blocks"])))
            self.assertEqual(result["unclassified"], exp["unclassified"], name)
            self.assertTrue(result["ok"])
            for e, b in hits:
                self.check_fields(name, e, b, plan)
        sys.stderr.write("\n" + "\n".join(report) + "\n")

    def test_messy_unclassified_reasons(self):
        _, plan = self.plans["messy.xlsx"]
        reasons = {u["range"]: u["reason"] for u in plan["tabs"][0]["unclassified"]}
        self.assertEqual(reasons, {"A1:A1": "lone-text", "B3:B3": "lone-text", "Z100:Z100": "stray-cell"})

    def test_chart_combo_warning(self):
        result, _ = self.plans["chart_heavy.xlsx"]
        self.assertTrue(any("combo chart" in w for w in result["warnings"]), result["warnings"])

    def test_block_ids_unique_and_ordered(self):
        for name, (_, plan) in self.plans.items():
            ids = [b["id"] for t in plan["tabs"] for b in t["blocks"]]
            self.assertEqual(len(ids), len(set(ids)), name)
            for t in plan["tabs"]:
                rows = [(b["row"], b["col"]) for b in t["blocks"]]
                self.assertEqual(rows, sorted(rows), name)


class SampleHouseTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not os.path.exists(SAMPLE_XLSX):
            subprocess.run([sys.executable, os.path.join(HERE, "build_xlsx.py"), SAMPLE_SPEC], check=True,
                           capture_output=True)
        cls.tmp = tempfile.mkdtemp(prefix="import_scan_house_")
        out = os.path.join(cls.tmp, "house.json")
        code, cls.result, err = run_scanner(SAMPLE_XLSX, out)
        if code != 0:
            raise RuntimeError("%s %s" % (cls.result, err))
        with open(out) as fh:
            cls.plan = json.load(fh)
        with open(SAMPLE_SPEC) as fh:
            cls.spec = json.load(fh)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def test_zero_unclassified(self):
        self.assertEqual(self.result["unclassified"], 0, [t["unclassified"] for t in self.plan["tabs"]])
        self.assertTrue(any(w.startswith("readme-like") for w in self.result["warnings"]))

    def test_every_spec_block_found_with_its_type(self):
        expected = []
        for sheet in self.spec["sheets"]:
            for block in sheet["blocks"]:
                g = build_xlsx.Geom(block, sheet["name"])
                expected.append({"sheet": sheet["name"], "type": block["type"], "title": block["title"],
                                 "row": g.row, "col": g.col, "rows": g.height, "cols": g.width})
        hits, misses = match(expected, self.plan)
        self.assertEqual(misses, [], "spec blocks the scanner missed or mistyped")
        for e, b in hits:
            self.assertEqual(b["title"], e["title"], b["range"])
            floor = {"text": 0.8, "chart": 0.9, "table": 0.75, "calc": 0.85, "inputs": 0.85, "output": 0.85}
            self.assertGreaterEqual(b["confidence"], floor[e["type"]], "%s %s" % (b["range"], b["detected"]))

    def test_house_notes_and_charts(self):
        tabs = {t["name"]: t for t in self.plan["tabs"]}
        market = [b for b in tabs["Assumptions"]["blocks"] if b["title"] == "Market frame"][0]
        self.assertEqual(market["note"], "Refresh when the SBA profile updates")
        self.assertEqual(market["rowNotes"][0], "The model starts here")
        self.assertEqual(market["rowLabels"], ["Base year", "US small businesses (M)", "Serviceable share", "Target segments"])
        self.assertIn("title_house", market["detected"]["signals"])
        tam = [b for b in tabs["TAM Build"]["blocks"] if b["type"] == "table"][0]
        self.assertEqual(tam["totalRow"][0], "Total")
        self.assertEqual(tam["columns"][0], "Segment")
        chart = [b for b in tabs["TAM Build"]["blocks"] if b["type"] == "chart"][0]
        self.assertEqual(chart["chart"]["sourceBlock"], tam["id"])
        self.assertEqual(chart["chart"]["seriesCols"], [4])
        self.assertEqual(chart["chart"]["categoriesCol"], 0)
        ramp_chart = [b for b in tabs["Revenue Ramp"]["blocks"] if b["type"] == "chart"][0]
        self.assertEqual(ramp_chart["chart"]["seriesCols"], [4, 5])
        self.assertEqual(ramp_chart["chart"]["kind"], "line")

    def test_ids_follow_sheet_index(self):
        for si, tab in enumerate(self.plan["tabs"]):
            for n, b in enumerate(tab["blocks"], 1):
                self.assertTrue(b["id"].startswith("imp_%d_" % si), b["id"])


class CliTest(unittest.TestCase):
    def test_missing_file_fails_cleanly(self):
        tmp = tempfile.mkdtemp(prefix="import_scan_cli_")
        try:
            code, result, _ = run_scanner(os.path.join(tmp, "nope.xlsx"), os.path.join(tmp, "out.json"))
            self.assertEqual(code, 1)
            self.assertFalse(result["ok"])
            self.assertTrue(result["errors"])
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_usage(self):
        proc = subprocess.run([sys.executable, SCANNER], capture_output=True, text=True)
        self.assertEqual(proc.returncode, 1)
        self.assertFalse(json.loads(proc.stdout)["ok"])


if __name__ == "__main__":
    unittest.main()

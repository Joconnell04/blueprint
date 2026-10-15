#!/usr/bin/env python3
"""Tests for build_xlsx.py. Run: python3 builder/test_build.py"""

import copy
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import build_xlsx  # noqa: E402

EM_DASH = "\u2014"


def sample_spec(out_path, mode, template_path=None):
    """A 3 sheet spec: inputs on one tab, a table plus chart on the next, outputs last."""
    spec = {
        "outputPath": out_path,
        "formatting": mode,
        "project": {"name": "Test Model", "brief": "A small model used by the tests."},
        "readme": {
            "purpose": "Sizes a made up market so the builder has something to chew on.",
            "contents": [
                {"tab": "Assumptions", "holds": "Inputs"},
                {"tab": "Model", "holds": "The build"},
                {"tab": "Outputs", "holds": "Results"},
            ],
            "rules": ["Blue cells are inputs.", "Black cells are formulas."],
        },
        "sheets": ([{"name": "Read Me", "blocks": []}] if mode in ("house", "skill") else []) + [
            {
                "name": "Assumptions",
                "blocks": [
                    {
                        "id": "b_in", "type": "inputs", "title": "Inputs",
                        "row": 2, "col": 2, "cols": 4,
                        "note": "Owned by finance",
                        "rows": [
                            {"label": "Growth", "value": 0.1, "format": "percent",
                             "note": "Analyst estimate"},
                            {"label": "Price", "value": 100, "format": "currency"},
                            {"label": "Blend",
                             "value": {"f": "={{b_in!Price}}*(1+{{b_in!Growth}})"},
                             "format": "currency"},
                        ],
                    }
                ],
            },
            {
                "name": "Model",
                "blocks": [
                    {
                        "id": "b_tbl", "type": "table", "title": "Revenue by segment",
                        "row": 2, "col": 2, "cols": 3,
                        "columns": ["Segment", "Year 1", "Year 2"],
                        "colFormats": ["text", "currency", "currency"],
                        "data": [
                            ["SMB", 120, {"f": "={{b_tbl!0,1}}*(1+{{b_in!Growth}})"}],
                            ["Enterprise", 300, {"f": "={{b_tbl!1,1}}*1.2"}],
                        ],
                        "totalRow": ["Total",
                                     {"f": "=SUM({{b_tbl!col:1}})"},
                                     {"f": "=SUM({{b_tbl!col:2}})"}],
                    },
                    {
                        "id": "b_chart", "type": "chart", "title": "Revenue by year",
                        "row": 8, "col": 2, "cols": 3, "rows": 12,
                        "chart": {"kind": "bar", "sourceBlock": "b_tbl",
                                  "categoriesCol": 0, "seriesCols": [1, 2],
                                  "yTitle": "USD M"},
                    },
                ],
            },
            {
                "name": "Outputs",
                "blocks": [
                    {
                        "id": "b_res", "type": "output", "title": "Headline",
                        "row": 2, "col": 2, "cols": 3,
                        "rows": [
                            {"label": "Total year 1",
                             "value": {"f": "={{b_tbl!total:1}}"}, "format": "currency"},
                            {"label": "Same cell, long hand",
                             "value": {"f": "={{sheet:b_tbl}}!C6"}, "format": "currency"},
                        ],
                    },
                    {
                        "id": "b_txt", "type": "text", "title": "Caveats",
                        "row": 6, "col": 2, "cols": 3,
                        "text": "Everything here is invented and should not leave the test suite.",
                    },
                ],
            },
        ],
    }
    if template_path:
        spec["templatePath"] = template_path
    return spec


def make_template(path):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Body copy"
    ws["A1"].font = build_xlsx.Font(name="Georgia", size=12)
    ws["A2"] = "A heading"
    ws["A2"].font = build_xlsx.Font(name="Trebuchet MS", size=14, bold=True)
    wb.save(path)


HEADER_FILL = "D9E1F2"


def make_style_template(path):
    """A template whose Style sheet defines Title and Header only; the rest falls back."""
    wb = Workbook()
    body = wb.active
    body.title = "Sample"
    body["A1"] = "Body copy"
    body["A1"].font = build_xlsx.Font(name="Georgia", size=12)
    style = wb.create_sheet("Style")
    style["A1"], style["B1"] = "Title", "Key Assumptions"
    style["B1"].font = build_xlsx.Font(name="Verdana", size=14, bold=True, color="1F3864")
    style["B1"].alignment = build_xlsx.Alignment(horizontal="center")
    style["A2"], style["B2"] = "Header", "Year 1"
    style["B2"].font = build_xlsx.Font(name="Georgia", size=12, bold=True)
    style["B2"].fill = build_xlsx.PatternFill(fill_type="solid", fgColor=HEADER_FILL)
    style["A3"], style["B3"] = "Note", "Analyst estimate"
    style["B3"].font = build_xlsx.Font(name="Georgia", size=10, italic=True, color="7F7F7F")
    style["A4"], style["B4"] = "Ignored", "not a role"
    style.column_dimensions["B"].width = 33
    style.column_dimensions["C"].width = 15
    style.column_dimensions["D"].width = 45
    wb.save(path)


SKILL_WITH_BLOCK = """---
name: test-format
description: "A spreadsheet format used by the builder tests."
format:
  font: Calibri
  size: 9
  input_color: 1F5F99
  label_divider: none
---

# Test format

Body text the builder never reads.
"""

SKILL_WITHOUT_BLOCK = """---
name: bare-format
description: "A spreadsheet format with no mechanical block."
---

# Bare format
"""


def write_text(path, text):
    with open(path, "w", encoding="utf-8") as handle:
        handle.write(text)


class BuildTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="workbook_planner_test_")
        cls.template = os.path.join(cls.tmp, "template.xlsx")
        make_template(cls.template)
        cls.results = {}
        cls.books = {}
        for mode in ("plain", "house", "template"):
            path = os.path.join(cls.tmp, mode + ".xlsx")
            result = build_xlsx.build(sample_spec(path, mode, cls.template))
            cls.results[mode] = result
            if result["ok"]:
                cls.books[mode] = load_workbook(path)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    # ------------------------------------------------------------- all modes

    def test_all_three_modes_build(self):
        for mode in ("plain", "house", "template"):
            result = self.results[mode]
            self.assertTrue(result["ok"], "%s failed: %s" % (mode, result["errors"]))
            self.assertEqual(result["errors"], [])
            self.assertTrue(os.path.exists(result["path"]))

    def test_titles_land_on_the_planned_cells(self):
        for mode, wb in self.books.items():
            self.assertEqual(wb["Assumptions"]["B2"].value, "Inputs", mode)
            self.assertEqual(wb["Assumptions"]["B3"].value, "Growth", mode)
            self.assertEqual(wb["Assumptions"]["C4"].value, 100, mode)
            self.assertEqual(wb["Model"]["B2"].value, "Revenue by segment", mode)
            self.assertEqual(wb["Model"]["B3"].value, "Segment", mode)
            self.assertEqual(wb["Model"]["B4"].value, "SMB", mode)
            self.assertEqual(wb["Model"]["B6"].value, "Total", mode)
            self.assertEqual(wb["Outputs"]["B2"].value, "Headline", mode)
            self.assertEqual(wb["Outputs"]["B6"].value, "Caveats", mode)
            self.assertTrue(str(wb["Outputs"]["B7"].value).startswith("Everything here"), mode)

    def test_templates_resolve_to_real_addresses(self):
        for mode, wb in self.books.items():
            self.assertEqual(wb["Assumptions"]["C5"].value, "=C4*(1+C3)", mode)
            self.assertEqual(wb["Model"]["D4"].value, "=C4*(1+'Assumptions'!C3)", mode)
            self.assertEqual(wb["Model"]["D5"].value, "=C5*1.2", mode)
            self.assertEqual(wb["Model"]["C6"].value, "=SUM(C4:C5)", mode)
            self.assertEqual(wb["Model"]["D6"].value, "=SUM(D4:D5)", mode)
            self.assertEqual(wb["Outputs"]["C3"].value, "='Model'!C6", mode)
            self.assertEqual(wb["Outputs"]["C4"].value, "='Model'!C6", mode)

    def test_no_curly_templates_survive(self):
        for mode, wb in self.books.items():
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        if isinstance(cell.value, str):
                            self.assertNotIn("{{", cell.value, "%s %s" % (mode, cell.coordinate))

    def test_number_formats(self):
        for mode, wb in self.books.items():
            self.assertEqual(wb["Assumptions"]["C3"].number_format, '0.0%;(0.0%);"-"', mode)
            self.assertEqual(wb["Assumptions"]["C4"].number_format, '$#,##0;($#,##0);"-"', mode)
            self.assertEqual(wb["Model"]["C4"].number_format, '$#,##0;($#,##0);"-"', mode)

    def test_chart_count_and_wiring(self):
        for mode, wb in self.books.items():
            self.assertEqual(len(wb["Model"]._charts), 1, mode)
            self.assertEqual(len(wb["Assumptions"]._charts), 0, mode)
            chart = wb["Model"]._charts[0]
            self.assertEqual(len(chart.series), 2, mode)
            refs = [s.val.numRef.f for s in chart.series]
            self.assertEqual(refs, ["'Model'!$C$4:$C$5", "'Model'!$D$4:$D$5"], mode)
            self.assertEqual(chart.series[0].tx.strRef.f, "'Model'!C3", mode)
        self.assertEqual(
            [s["charts"] for s in self.results["plain"]["sheets"]], [0, 1, 0])

    def test_reported_counts(self):
        report = {s["name"]: s for s in self.results["plain"]["sheets"]}
        self.assertEqual(sorted(report), ["Assumptions", "Model", "Outputs"])
        self.assertEqual(report["Model"]["formulas"], 4)
        self.assertEqual(report["Outputs"]["formulas"], 2)
        self.assertEqual(report["Assumptions"]["formulas"], 1)

    # ----------------------------------------------------------- house mode

    def test_house_read_me_is_first(self):
        wb = self.books["house"]
        self.assertEqual(wb.sheetnames[0], "Read Me")
        self.assertEqual(wb.sheetnames,
                         ["Read Me", "Assumptions", "Model", "Outputs"])
        rm = wb["Read Me"]
        self.assertEqual(rm["A1"].value, "Test Model")
        self.assertEqual(rm["A1"].font.name, "Arial")
        self.assertEqual(rm["A1"].font.sz, 16)
        self.assertTrue(rm["A1"].font.bold)
        self.assertEqual(rm["A1"].font.color.rgb[-6:], build_xlsx.NAVY)
        self.assertTrue(str(rm["A2"].value).startswith("v1, built "))
        self.assertTrue(rm["A2"].font.italic)
        self.assertEqual(rm["A2"].font.color.rgb[-6:], build_xlsx.GREY)
        headings = [rm.cell(r, 1).value for r in range(1, rm.max_row + 1)]
        for section in ("Purpose", "Contents", "Reading a Cell", "Rules"):
            self.assertIn(section, headings)

    def test_house_fonts_and_gridlines(self):
        wb = self.books["house"]
        for ws in wb.worksheets:
            self.assertFalse(ws.sheet_view.showGridLines, ws.title)
        cell = wb["Assumptions"]["B3"]
        self.assertEqual(cell.font.name, "Arial")
        self.assertEqual(cell.font.sz, 10)
        title = wb["Assumptions"]["B2"]
        self.assertTrue(title.font.bold)
        self.assertEqual(title.alignment.horizontal, "center")
        self.assertEqual(title.border.bottom.style, "medium")
        self.assertEqual(title.border.bottom.color.rgb[-6:], build_xlsx.NAVY)
        self.assertEqual(title.border.top.style, "thin")
        self.assertIn("B2:C2", [str(r) for r in wb["Assumptions"].merged_cells.ranges])

    def test_house_colours_and_dashed_divider(self):
        ws = self.books["house"]["Assumptions"]
        self.assertEqual(ws["C3"].font.color.rgb[-6:], build_xlsx.BLUE)   # literal input
        self.assertEqual(ws["C5"].font.color.rgb[-6:], build_xlsx.BLACK)  # formula
        note = ws["D3"]
        self.assertEqual(note.font.color.rgb[-6:], build_xlsx.GREY)
        self.assertTrue(note.font.italic)
        self.assertIsNone(note.border.left.style)
        self.assertFalse(note.alignment.wrap_text)
        self.assertEqual(ws["B3"].border.right.style, "dashed")
        self.assertEqual(ws["C3"].border.right.style, "thin")
        self.assertEqual(self.books["house"]["Model"]["B4"].border.right.style, "dashed")
        self.assertTrue(self.books["house"]["Model"]["B6"].font.bold)

    def test_house_column_widths(self):
        ws = self.books["house"]["Assumptions"]
        # content fit: label column at least the house minimum, values at least 12, the
        # note column sized to its longest note because no values live in it
        self.assertGreaterEqual(ws.column_dimensions["B"].width, 24)
        self.assertGreaterEqual(ws.column_dimensions["C"].width, 12)
        self.assertLessEqual(ws.column_dimensions["C"].width, 20)
        self.assertGreaterEqual(ws.column_dimensions["D"].width, len("Analyst estimate"))

    def test_house_replaces_long_dashes(self):
        path = os.path.join(self.tmp, "dash.xlsx")
        spec = sample_spec(path, "house")
        spec["sheets"][1]["blocks"][0]["rows"][2]["label"] = "Blend %s net" % EM_DASH
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        self.assertTrue(any("dash" in w for w in result["warnings"]))
        self.assertEqual(load_workbook(path)["Assumptions"]["B5"].value, "Blend, net")

    # ----------------------------------------------------------- plain mode

    def test_plain_defaults(self):
        wb = self.books["plain"]
        self.assertNotIn("Read Me", wb.sheetnames)
        ws = wb["Assumptions"]
        self.assertTrue(ws.sheet_view.showGridLines)
        self.assertEqual(ws["B3"].font.name, "Calibri")
        self.assertEqual(ws["B3"].font.sz, 11)
        self.assertTrue(ws["B2"].font.bold)
        self.assertIsNone(ws["B3"].border.left.style)
        self.assertGreaterEqual(ws.column_dimensions["B"].width, 22)
        model = wb["Model"]
        self.assertTrue(model["B3"].font.bold)
        self.assertEqual(model["B3"].border.bottom.style, "thin")
        self.assertEqual(model["C4"].border.left.style, "thin")

    # -------------------------------------------------------- template mode

    def test_template_fonts_come_from_the_template(self):
        ws = self.books["template"]["Assumptions"]
        self.assertEqual(ws["B3"].font.name, "Georgia")
        self.assertEqual(ws["B3"].font.sz, 12)
        self.assertEqual(ws["B2"].font.name, "Trebuchet MS")
        self.assertEqual(ws["B2"].font.sz, 14)
        self.assertNotIn("Read Me", self.books["template"].sheetnames)
        self.assertEqual(self.results["template"]["warnings"], [])

    def test_unreadable_template_falls_back_to_plain(self):
        path = os.path.join(self.tmp, "fallback.xlsx")
        spec = sample_spec(path, "template", os.path.join(self.tmp, "missing.xlsx"))
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        self.assertTrue(any("template" in w for w in result["warnings"]))
        self.assertEqual(load_workbook(path)["Assumptions"]["B3"].font.name, "Calibri")

    def test_template_style_sheet_clones_roles(self):
        template = os.path.join(self.tmp, "style_template.xlsx")
        make_style_template(template)
        path = os.path.join(self.tmp, "styled.xlsx")
        result = build_xlsx.build(sample_spec(path, "template", template))
        self.assertTrue(result["ok"], result["errors"])
        self.assertEqual(result["warnings"], [])
        wb = load_workbook(path)
        title = wb["Assumptions"]["B2"]
        self.assertEqual((title.font.name, title.font.sz, title.font.bold), ("Verdana", 14, True))
        self.assertIn("B2:C2", [str(r) for r in wb["Assumptions"].merged_cells.ranges])
        header = wb["Model"]["B3"]
        self.assertEqual(header.fill.fill_type, "solid")
        self.assertEqual(header.fill.fgColor.rgb[-6:], HEADER_FILL)
        self.assertTrue(header.font.bold)
        note = wb["Assumptions"]["D3"]
        self.assertTrue(note.font.italic)
        self.assertEqual(note.font.color.rgb[-6:], "7F7F7F")
        # no Label, Input or Formula rows: those fall back to the base font, plain borders
        label = wb["Assumptions"]["B3"]
        self.assertEqual((label.font.name, label.font.sz), ("Georgia", 12))
        self.assertIsNone(label.border.left.style)
        self.assertEqual(wb["Assumptions"]["C3"].font.color.rgb[-6:], build_xlsx.BLACK)  # not an Input role
        self.assertEqual(wb["Model"]["C4"].border.left.style, "thin")
        # widths from the Style sheet's B, C, D
        self.assertGreaterEqual(wb["Assumptions"].column_dimensions["B"].width, 33)
        self.assertGreaterEqual(wb["Assumptions"].column_dimensions["C"].width, 15)
        self.assertEqual(wb["Assumptions"].column_dimensions["D"].width, 45)

    def test_shipped_style_template_builds(self):
        template = os.path.join(HERE, "corpus", "style_template.xlsx")
        if not os.path.exists(template):
            self.skipTest("builder/corpus/style_template.xlsx not generated (python3 builder/make_corpus.py)")
        path = os.path.join(self.tmp, "corpus_styled.xlsx")
        result = build_xlsx.build(sample_spec(path, "template", template))
        self.assertTrue(result["ok"], result["errors"])
        self.assertEqual(result["warnings"], [])
        wb = load_workbook(path)
        self.assertEqual(wb["Model"]["B3"].fill.fgColor.rgb[-6:], "44546A")
        self.assertEqual(wb["Assumptions"]["C3"].font.color.rgb[-6:], "1F5F99")  # Input role
        self.assertEqual(wb["Assumptions"]["C5"].font.color.rgb[-6:], "000000")  # Formula role
        self.assertEqual(wb["Model"]["B6"].border.bottom.style, "double")        # Total role

    # ------------------------------------------------------------ skill mode

    def skill_spec(self, name, skill_md, text):
        skill_dir = os.path.join(self.tmp, name)
        os.makedirs(skill_dir, exist_ok=True)
        skill_path = os.path.join(skill_dir, "SKILL.md")
        write_text(skill_path, text)
        spec = sample_spec(os.path.join(self.tmp, name + ".xlsx"), "skill")
        spec["skill"] = {"path": skill_path}
        return spec

    def test_skill_mode_reads_the_format_block(self):
        spec = self.skill_spec("skill_block", "SKILL.md", SKILL_WITH_BLOCK)
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        self.assertEqual(result["warnings"], [])
        wb = load_workbook(spec["outputPath"])
        ws = wb["Assumptions"]
        self.assertEqual((ws["B3"].font.name, ws["B3"].font.sz), ("Calibri", 9))
        self.assertEqual(ws["C3"].font.color.rgb[-6:], "1F5F99")   # literal input
        self.assertEqual(ws["C5"].font.color.rgb[-6:], build_xlsx.BLACK)  # formula
        self.assertEqual(ws["B3"].border.right.style, "thin")       # no dashed divider
        self.assertEqual(wb["Model"]["B4"].border.right.style, "thin")
        # everything else is still the house format
        self.assertFalse(ws.sheet_view.showGridLines)
        self.assertEqual(ws["B2"].border.bottom.style, "medium")
        self.assertEqual(ws["B2"].border.bottom.color.rgb[-6:], build_xlsx.NAVY)
        self.assertEqual(wb.sheetnames[0], "Read Me")
        self.assertEqual((wb["Read Me"]["A1"].font.name, wb["Read Me"]["A1"].font.sz), ("Calibri", 16))
        self.assertEqual(ws["C3"].number_format, build_xlsx.FORMATS["percent"])

    def test_skill_mode_without_block_falls_back_to_house(self):
        spec = self.skill_spec("skill_bare", "SKILL.md", SKILL_WITHOUT_BLOCK)
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        self.assertEqual(len(result["warnings"]), 1)
        self.assertIn("bare-format", result["warnings"][0])
        self.assertIn("format:", result["warnings"][0])
        ws = load_workbook(spec["outputPath"])["Assumptions"]
        self.assertEqual((ws["B3"].font.name, ws["B3"].font.sz), ("Arial", 10))
        self.assertEqual(ws["C3"].font.color.rgb[-6:], build_xlsx.BLUE)
        self.assertEqual(ws["B3"].border.right.style, "dashed")

    def test_skill_mode_warns_on_bad_values_and_keeps_defaults(self):
        text = SKILL_WITH_BLOCK.replace("  size: 9\n", "  size: big\n  title_rule: dotted\n  extra: 1\n")
        spec = self.skill_spec("skill_bad", "SKILL.md", text)
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        joined = "\n".join(result["warnings"])
        self.assertIn("format.size", joined)
        self.assertIn("format.title_rule", joined)
        self.assertIn("format.extra", joined)
        ws = load_workbook(spec["outputPath"])["Assumptions"]
        self.assertEqual((ws["B3"].font.name, ws["B3"].font.sz), ("Calibri", 10))
        self.assertEqual(ws["B2"].border.bottom.style, "medium")

    def test_skill_mode_missing_file_uses_house(self):
        spec = sample_spec(os.path.join(self.tmp, "skill_missing.xlsx"), "skill")
        spec["skill"] = {"path": os.path.join(self.tmp, "nowhere", "SKILL.md")}
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        self.assertTrue(any("skill file could not be read" in w for w in result["warnings"]))
        ws = load_workbook(spec["outputPath"])["Assumptions"]
        self.assertEqual(ws["B3"].font.name, "Arial")

    def test_front_matter_parser(self):
        fm = build_xlsx.front_matter(SKILL_WITH_BLOCK)
        self.assertEqual(fm["name"], "test-format")
        self.assertEqual(fm["description"], "A spreadsheet format used by the builder tests.")
        self.assertEqual(fm["format"], {"font": "Calibri", "size": "9",
                                        "input_color": "1F5F99", "label_divider": "none"})
        self.assertEqual(build_xlsx.front_matter("# no front matter\n"), {})
        # a scalar `format: spreadsheet` tag is not a block
        self.assertEqual(build_xlsx.front_matter("---\nformat: spreadsheet\n---\n")["format"], "spreadsheet")
        cfg = build_xlsx.skill_config({"gridlines": "yes", "note_italic": "false", "input_colour": "#abcdef",
                                       "size": "10.5", "max_note_chars": "80"}, "t", [])
        self.assertEqual(cfg, {"gridlines": True, "note_italic": False, "input_color": "ABCDEF",
                               "size": 10.5, "max_note_chars": 80})

    # ---------------------------------------------------------------- errors

    def test_unresolved_template_fails_with_a_listed_error(self):
        path = os.path.join(self.tmp, "bad.xlsx")
        spec = sample_spec(path, "plain")
        spec["sheets"][1]["blocks"][0]["data"][0][2] = {"f": "={{b_in!Nope}}*2"}
        result = build_xlsx.build(spec)
        self.assertFalse(result["ok"])
        self.assertTrue(result["errors"])
        joined = " ".join(result["errors"])
        self.assertIn("b_in!Nope", joined)
        self.assertIn("Nope", joined)
        self.assertFalse(os.path.exists(path))

    def test_unknown_block_id_fails(self):
        spec = sample_spec(os.path.join(self.tmp, "bad2.xlsx"), "plain")
        spec["sheets"][2]["blocks"][0]["rows"][0]["value"] = {"f": "={{nosuch!total:1}}"}
        result = build_xlsx.build(spec)
        self.assertFalse(result["ok"])
        self.assertIn("nosuch", " ".join(result["errors"]))

    def test_formula_without_equals_fails(self):
        spec = sample_spec(os.path.join(self.tmp, "bad3.xlsx"), "plain")
        spec["sheets"][1]["blocks"][0]["data"][1][2] = {"f": "C5*1.2"}
        result = build_xlsx.build(spec)
        self.assertFalse(result["ok"])
        self.assertIn("does not start with", " ".join(result["errors"]))

    def test_overlapping_blocks_fail_and_name_both(self):
        spec = sample_spec(os.path.join(self.tmp, "bad4.xlsx"), "plain")
        spec["sheets"][2]["blocks"][1]["row"] = 3
        result = build_xlsx.build(spec)
        self.assertFalse(result["ok"])
        joined = " ".join(result["errors"])
        self.assertIn("overlap", joined)
        self.assertIn("b_res", joined)
        self.assertIn("b_txt", joined)

    def test_bad_sheet_name_fails(self):
        spec = sample_spec(os.path.join(self.tmp, "bad5.xlsx"), "plain")
        spec["sheets"][1]["name"] = "Model/One"
        result = build_xlsx.build(spec)
        self.assertFalse(result["ok"])
        self.assertIn("Model/One", " ".join(result["errors"]))

    def test_duplicate_sheet_names_fail(self):
        spec = sample_spec(os.path.join(self.tmp, "bad6.xlsx"), "plain")
        spec["sheets"][2]["name"] = "Model"
        result = build_xlsx.build(spec)
        self.assertFalse(result["ok"])
        self.assertIn("duplicate sheet name", " ".join(result["errors"]))

    # ------------------------------------------------------------------ cli

    def test_cli_prints_one_json_object_and_exits_zero(self):
        path = os.path.join(self.tmp, "cli.xlsx")
        spec_path = os.path.join(self.tmp, "cli_spec.json")
        with open(spec_path, "w") as handle:
            json.dump(sample_spec(path, "house"), handle)
        proc = subprocess.run(
            [sys.executable, os.path.join(HERE, "build_xlsx.py"), spec_path],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        self.assertEqual(proc.returncode, 0, proc.stderr.decode())
        payload = json.loads(proc.stdout.decode())
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["path"], path)
        self.assertEqual([s["name"] for s in payload["sheets"]],
                         ["Read Me", "Assumptions", "Model", "Outputs"])

    def test_cli_exits_one_on_a_bad_spec(self):
        spec = sample_spec(os.path.join(self.tmp, "cli_bad.xlsx"), "plain")
        spec["sheets"][1]["blocks"][0]["data"][0][2] = {"f": "={{b_in!Nope}}"}
        spec_path = os.path.join(self.tmp, "cli_bad_spec.json")
        with open(spec_path, "w") as handle:
            json.dump(spec, handle)
        proc = subprocess.run(
            [sys.executable, os.path.join(HERE, "build_xlsx.py"), spec_path],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        self.assertEqual(proc.returncode, 1)
        payload = json.loads(proc.stdout.decode())
        self.assertFalse(payload["ok"])
        self.assertTrue(payload["errors"])

    # --------------------------------------------------------------- snapshot

    def test_snapshot_written_next_to_the_workbook(self):
        for mode in ("plain", "house", "template"):
            result = self.results[mode]
            expected = os.path.join(self.tmp, mode + ".snapshot.json")
            self.assertEqual(result["snapshot"], expected, mode)
            self.assertTrue(os.path.exists(expected), mode)
            with open(expected) as handle:
                snap = json.load(handle)
            self.assertEqual([s["name"] for s in snap["sheets"]],
                             [s["name"] for s in result["sheets"]], mode)

    def snapshot_sheet(self, mode, name):
        with open(self.results[mode]["snapshot"]) as handle:
            snap = json.load(handle)
        sheet = [s for s in snap["sheets"] if s["name"] == name][0]
        cells = {c["a"]: c for c in sheet["cells"]}
        return sheet, cells

    def test_snapshot_kinds_at_known_cells(self):
        for mode in ("plain", "house"):
            sheet, cells = self.snapshot_sheet(mode, "Assumptions")
            self.assertEqual(cells["B2"], {"a": "B2", "v": "Inputs", "k": "title"}, mode)
            self.assertEqual(cells["B3"]["k"], "label", mode)
            self.assertEqual(cells["C3"], {"a": "C3", "v": 0.1, "k": "input",
                                           "fmt": '0.0%;(0.0%);"-"', "d": "10.0%"}, mode)
            self.assertEqual(cells["C4"], {"a": "C4", "v": 100, "k": "input",
                                           "fmt": '$#,##0;($#,##0);"-"', "d": "$100"}, mode)
            self.assertEqual(cells["C5"], {"a": "C5", "v": "=C4*(1+C3)", "k": "formula",
                                           "fmt": '$#,##0;($#,##0);"-"'}, mode)
            self.assertEqual(cells["D3"]["k"], "note", mode)   # row note
            self.assertEqual(cells["D2"]["k"], "note", mode)   # block note, same column
            self.assertEqual(sheet["maxRow"], 5, mode)
            self.assertEqual(sheet["maxCol"], 6, mode)
            self.assertIn("D3:F3", sheet["merges"], mode)      # merged to the sheet edge
            self.assertGreaterEqual(sheet["widths"]["B"], 22, mode)
            # cells arrive in row-major order
            addresses = [c["a"] for c in sheet["cells"]]
            self.assertEqual(addresses, ["B2", "D2", "B3", "C3", "D3", "B4", "C4", "B5", "C5"], mode)

            sheet, cells = self.snapshot_sheet(mode, "Model")
            self.assertEqual(cells["B3"]["k"], "header", mode)
            self.assertEqual(cells["B4"]["k"], "label", mode)   # the table's own label column
            self.assertEqual(cells["C4"]["k"], "input", mode)   # a literal in a table block
            self.assertEqual(cells["C4"]["d"], "$120", mode)
            self.assertEqual(cells["D4"]["k"], "formula", mode)
            self.assertEqual(cells["B6"]["k"], "total", mode)
            self.assertEqual(cells["C6"]["k"], "total", mode)
            self.assertNotIn("d", cells["C6"], mode)
            self.assertEqual(sheet["charts"], [
                {"blockId": "b_chart", "anchor": "B8", "kind": "bar", "series": 2}], mode)

            sheet, cells = self.snapshot_sheet(mode, "Outputs")
            self.assertEqual(cells["B7"]["k"], "text", mode)
        # house mode merges the title rule across the block, notes excluded
        self.assertIn("B2:C2", self.snapshot_sheet("house", "Assumptions")[0]["merges"])
        # the cover page is all readme kind
        sheet, cells = self.snapshot_sheet("house", "Read Me")
        self.assertEqual(cells["A1"]["v"], "Test Model")
        self.assertEqual({c["k"] for c in sheet["cells"]}, {"readme"})
        self.assertEqual(sheet["widths"], {"A": 30, "B": 45, "C": 14})   # 89: a portrait page
        self.assertIn("Purpose", [c["v"] for c in sheet["cells"]])

    def test_snapshot_block_extents_match_written_cells(self):
        sheet, _ = self.snapshot_sheet("house", "Assumptions")
        # title row 2 through row 5, B through D: label, value and the one note column
        self.assertEqual(sheet["blocks"], [
            {"id": "b_in", "row": 2, "col": 2, "rows": 4, "cols": 3, "type": "inputs",
             "adjusted": "fit"}])
        self.assertEqual(sheet["adjusted"], [
            {"blockId": "b_in", "from": 4, "to": 2, "reason": "fit"}])
        sheet, _ = self.snapshot_sheet("house", "Model")
        self.assertEqual(sheet["blocks"], [
            {"id": "b_tbl", "row": 2, "col": 2, "rows": 5, "cols": 3, "type": "table"},
            {"id": "b_chart", "row": 8, "col": 2, "rows": 12, "cols": 3, "type": "chart"}])
        sheet, _ = self.snapshot_sheet("house", "Outputs")
        self.assertEqual(sheet["blocks"], [
            {"id": "b_res", "row": 2, "col": 2, "rows": 3, "cols": 2, "type": "output",
             "adjusted": "fit"},
            {"id": "b_txt", "row": 6, "col": 2, "rows": 2, "cols": 3, "type": "text"}])
        # every recorded cell sits inside its sheet's block extents
        for name in ("Assumptions", "Model", "Outputs"):
            sheet, cells = self.snapshot_sheet("house", name)
            boxes = [(b["row"], b["col"], b["row"] + b["rows"] - 1, b["col"] + b["cols"] - 1)
                     for b in sheet["blocks"]]
            for c in sheet["cells"]:
                row = int("".join(ch for ch in c["a"] if ch.isdigit()))
                letters = "".join(ch for ch in c["a"] if ch.isalpha())
                colno = 0
                for ch in letters:
                    colno = colno * 26 + (ord(ch) - 64)
                self.assertTrue(any(r1 <= row <= r2 and c1 <= colno <= c2
                                    for r1, c1, r2, c2 in boxes), "%s %s" % (name, c["a"]))

    def test_snapshot_display_strings(self):
        d = build_xlsx.display
        F = build_xlsx.FORMATS
        self.assertEqual(d(1200, F["currency"]), "$1,200")
        self.assertEqual(d(-1200.4, F["currency"]), "($1,200)")
        self.assertEqual(d(1234.5, F["currency2"]), "$1,234.50")
        self.assertEqual(d(0.125, F["percent"]), "12.5%")
        self.assertEqual(d(-0.05, F["percent"]), "(5.0%)")
        self.assertEqual(d(2.44, F["multiple"]), "2.4x")
        self.assertEqual(d(-2.44, F["multiple"]), "(2.4x)")
        self.assertEqual(d(1234567, F["count"]), "1,234,567")
        self.assertEqual(d(-1234, F["count"]), "(1,234)")
        for key in ("currency", "currency2", "percent", "multiple", "count", "number"):
            self.assertEqual(d(0, F[key]), "-", key)          # three part: zero is a dash
        self.assertEqual(d(-0.0001, F["currency"]), "-")      # rounds to zero, shows the dash
        self.assertEqual(d(0.5, F["count"]), "1")          # half rounds up, as in Excel
        self.assertEqual(d(1234.567, F["number"]), "1,234.57")
        self.assertIsNone(d("text", F["currency"]))
        self.assertIsNone(d(True, F["count"]))
        self.assertIsNone(d(12, F["text"]))

    def test_snapshot_cap_sets_truncated(self):
        path = os.path.join(self.tmp, "big.xlsx")
        spec = sample_spec(path, "plain")
        ncols, nrows = 60, 100   # 6,000 data cells, over the 5,000 cap
        spec["sheets"].append({
            "name": "Big",
            "blocks": [{
                "id": "b_big", "type": "table", "title": "Big table",
                "row": 2, "col": 2, "cols": ncols,
                "columns": ["c%d" % j for j in range(ncols)],
                "data": [[i * ncols + j for j in range(ncols)] for i in range(nrows)],
            }],
        })
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        with open(result["snapshot"]) as handle:
            snap = json.load(handle)
        by_name = {s["name"]: s for s in snap["sheets"]}
        big = by_name["Big"]
        self.assertTrue(big.get("truncated"))
        self.assertEqual(len(big["cells"]), build_xlsx.SNAPSHOT_CELL_CAP)
        self.assertEqual(big["cells"][0]["a"], "B2")
        self.assertEqual(big["cells"][1]["a"], "B3")   # first header follows the title
        self.assertNotIn("truncated", by_name["Model"])
        self.assertLess(os.path.getsize(result["snapshot"]), 2 * 1024 * 1024)

    # ------------------------------------------- row notes, validation, setup

    def noted_table(self, extra=None):
        block = {
            "id": "b_n", "type": "table", "title": "Ramp", "row": 2, "col": 2, "cols": 3,
            "columns": ["Segment", "Year 1", "Year 2"],
            "colFormats": ["text", "currency", "currency"],
            "data": [["SMB", 1, 2], ["Enterprise", 3, 4], ["Mid", 5, 6]],
            "rowNotes": ["From CRM", None, "Estimate"],
            "totalRow": ["Total", {"f": "=SUM({{b_n!col:1}})"}, {"f": "=SUM({{b_n!col:2}})"}],
            "totalNote": "Sum check",
        }
        block.update(extra or {})
        return block

    def test_table_row_notes_land_and_grow_the_extent(self):
        for mode in ("plain", "house"):
            path = os.path.join(self.tmp, "rownotes_%s.xlsx" % mode)
            spec = sample_spec(path, mode)
            spec["sheets"].append({"name": "Noted", "blocks": [self.noted_table()]})
            result = build_xlsx.build(spec)
            self.assertTrue(result["ok"], result["errors"])
            ws = load_workbook(path)["Noted"]
            # notes sit in the column after the last data column, one per data row plus the total
            self.assertEqual(ws["E4"].value, "From CRM", mode)
            self.assertIsNone(ws["E5"].value, mode)
            self.assertEqual(ws["E6"].value, "Estimate", mode)
            self.assertEqual(ws["E7"].value, "Sum check", mode)
            self.assertEqual(ws["D7"].value, "=SUM(D4:D6)", mode)
            for a in ("E4", "E6", "E7"):
                cell = ws[a]
                self.assertTrue(cell.font.italic, "%s %s" % (mode, a))
                self.assertEqual(cell.font.color.rgb[-6:], build_xlsx.GREY, "%s %s" % (mode, a))
                self.assertIsNone(cell.border.left.style, "%s %s" % (mode, a))
                self.assertIsNone(cell.border.bottom.style, "%s %s" % (mode, a))
                self.assertFalse(cell.alignment.wrap_text, "%s %s" % (mode, a))
            self.assertEqual(ws["D4"].border.right.style, "thin", mode)  # data cells still boxed
            self.assertGreaterEqual(ws.column_dimensions["E"].width, len("Sum check"), mode)
            with open(result["snapshot"]) as handle:
                snap = json.load(handle)
            sheet = [s for s in snap["sheets"] if s["name"] == "Noted"][0]
            cells = {c["a"]: c for c in sheet["cells"]}
            self.assertEqual(cells["E4"], {"a": "E4", "v": "From CRM", "k": "note"}, mode)
            self.assertEqual(cells["E7"]["k"], "note", mode)
            self.assertNotIn("E5", cells, mode)
            # the extent grows by the note column: B..E, title row 2 through the total row 7
            self.assertEqual(sheet["blocks"], [
                {"id": "b_n", "row": 2, "col": 2, "rows": 6, "cols": 4, "type": "table"}], mode)
            if mode == "house":
                merges = [str(r) for r in ws.merged_cells.ranges]
                self.assertIn("B2:D2", merges)      # the title rule stops before the notes
                self.assertNotIn("B2:E2", merges)
        # without notes the same table keeps a 3 column extent
        path = os.path.join(self.tmp, "rownotes_none.xlsx")
        spec = sample_spec(path, "plain")
        spec["sheets"].append({"name": "Noted", "blocks": [
            self.noted_table({"rowNotes": [None, "", None], "totalNote": ""})]})
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        with open(result["snapshot"]) as handle:
            snap = json.load(handle)
        sheet = [s for s in snap["sheets"] if s["name"] == "Noted"][0]
        self.assertEqual(sheet["blocks"][0]["cols"], 3)
        self.assertIsNone(load_workbook(path)["Noted"]["E4"].value)
        # a neighbour sharing the table's rows moves the note past it, and both blocks use it (W4)
        path = os.path.join(self.tmp, "rownotes_band.xlsx")
        spec = sample_spec(path, "house")
        spec["sheets"].append({"name": "Noted", "blocks": [
            self.noted_table({"rowNotes": ["x" * 120, None, None]}),
            {"id": "b_side", "type": "text", "title": "Side", "row": 2, "col": 5, "cols": 2,
             "text": "t", "note": "beside"}]})
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        ws = load_workbook(path)["Noted"]
        self.assertEqual(ws["G4"].value, "x" * 120)     # right of b_side, which ends at F
        self.assertEqual(ws["G2"].value, "beside")      # the same column serves both blocks
        self.assertTrue(any("b_n r0 note" in w and "120 chars" in w for w in result["warnings"]))

    def test_data_validation_on_inputs_and_table_columns(self):
        path = os.path.join(self.tmp, "validation.xlsx")
        spec = sample_spec(path, "house")
        rows = spec["sheets"][1]["blocks"][0]["rows"]
        rows[0]["validation"] = {"min": 0, "max": 1}
        rows[1]["validation"] = {"type": "whole", "min": 1}
        rows.append({"label": "Scenario", "value": "Base",
                     "validation": {"list": ["Low", "Base", "High"]}})
        rows.append({"label": "Cap", "value": 5, "validation": {"type": "decimal", "max": 10}})
        rows.append({"label": "Broken", "value": 1, "validation": {"type": "nope"}})
        rows.append({"label": "Empty", "value": 1, "validation": {}})
        spec["sheets"].append({"name": "Noted", "blocks": [self.noted_table({
            "colValidation": [None, {"min": 0}, {"list": "={{b_n!col:1}}"}]})]})
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        self.assertTrue(any("row 5 validation" in w and "whole or decimal" in w for w in result["warnings"]))
        self.assertTrue(any("row 6 validation" in w and "needs a list" in w for w in result["warnings"]))
        wb = load_workbook(path)
        dvs = {str(dv.sqref): dv for dv in wb["Assumptions"].data_validations.dataValidation}
        self.assertEqual(sorted(dvs), ["C3", "C4", "C6", "C7"])
        self.assertEqual((dvs["C3"].type, dvs["C3"].operator, dvs["C3"].formula1, dvs["C3"].formula2),
                         ("decimal", "between", "0", "1"))
        self.assertEqual((dvs["C4"].type, dvs["C4"].operator, dvs["C4"].formula1),
                         ("whole", "greaterThanOrEqual", "1"))
        self.assertEqual((dvs["C6"].type, dvs["C6"].formula1), ("list", '"Low,Base,High"'))
        self.assertFalse(dvs["C6"].showDropDown)  # openpyxl's showDropDown hides the arrow, so it stays off
        self.assertEqual((dvs["C7"].type, dvs["C7"].operator, dvs["C7"].formula1),
                         ("decimal", "lessThanOrEqual", "10"))
        for dv in dvs.values():
            self.assertTrue(dv.allow_blank)
            self.assertTrue(dv.showErrorMessage)
            self.assertTrue(dv.error)
        dvs = {str(dv.sqref): dv for dv in wb["Noted"].data_validations.dataValidation}
        self.assertEqual(sorted(dvs), ["C4:C6", "D4:D6"])
        self.assertEqual((dvs["C4:C6"].type, dvs["C4:C6"].operator, dvs["C4:C6"].formula1),
                         ("decimal", "greaterThanOrEqual", "0"))
        self.assertEqual((dvs["D4:D6"].type, dvs["D4:D6"].formula1), ("list", "C4:C6"))
        with open(result["snapshot"]) as handle:
            snap = json.load(handle)
        by_name = {s["name"]: {c["a"]: c for c in s["cells"]} for s in snap["sheets"]}
        for a in ("C3", "C4", "C6", "C7"):
            self.assertTrue(by_name["Assumptions"][a].get("validation"), a)
        self.assertNotIn("validation", by_name["Assumptions"]["C5"])
        self.assertNotIn("validation", by_name["Assumptions"]["C8"])
        for a in ("C4", "D5", "C6"):
            self.assertTrue(by_name["Noted"][a].get("validation"), a)
        self.assertNotIn("validation", by_name["Noted"]["B4"])
        self.assertNotIn("validation", by_name["Noted"]["C7"])  # the total row is not validated

    def test_sheet_setup_freeze_print_and_zoom(self):
        path = os.path.join(self.tmp, "setup.xlsx")
        spec = sample_spec(path, "house")
        spec["sheets"][0].update({"printArea": "auto"})
        spec["sheets"][1].update({"freeze": "c3", "printArea": "auto", "landscape": True,
                                  "fitToWidth": True, "zoom": 90})
        spec["sheets"][2].update({"freeze": "B4", "printArea": "B2:E20", "landscape": False, "zoom": 85})
        spec["sheets"][3].update({"freeze": "nope", "printArea": "1:2", "landscape": "yes",
                                  "fitToWidth": 1, "zoom": 5})
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        wb = load_workbook(path)
        ws = wb["Assumptions"]
        self.assertEqual(ws.freeze_panes, "C3")
        self.assertEqual(ws.print_area, "'Assumptions'!$B$2:$F$5")   # the written range, note merge included
        self.assertEqual(ws.page_setup.orientation, "landscape")
        self.assertEqual((ws.page_setup.fitToWidth, ws.page_setup.fitToHeight), (1, 0))
        self.assertTrue(ws.sheet_properties.pageSetUpPr.fitToPage)
        self.assertEqual(ws.sheet_view.zoomScale, 90)
        ws = wb["Model"]
        self.assertEqual(ws.freeze_panes, "B4")
        self.assertEqual(ws.print_area, "'Model'!$B$2:$E$20")
        self.assertEqual(ws.page_setup.orientation, "portrait")   # false wins over the computed default
        self.assertEqual(ws.sheet_view.zoomScale, 85)
        self.assertEqual(wb["Read Me"].print_area, "'Read Me'!$A$1:$C$%d" % wb["Read Me"].max_row)
        # a bad value warns and the sheet falls back to the computed default (W6)
        ws = wb["Outputs"]
        self.assertEqual(ws.freeze_panes, "C3")                   # first pair row, label column held
        self.assertEqual(ws.print_area, "'Outputs'!$B$2:$D$7")
        self.assertIsNone(ws.sheet_view.zoomScale)
        for text in ("freeze 'nope'", "printArea '1:2'", "landscape must be", "fitToWidth must be", "zoom must be"):
            self.assertTrue(any("sheet 'Outputs'" in w and text in w for w in result["warnings"]), text)

    # --------------------------------------------------------- shipped files

    def test_shipped_sample_spec_builds(self):
        spec_path = os.path.join(HERE, "sample_spec.json")
        if not os.path.exists(spec_path):
            self.skipTest("sample_spec.json not present")
        with open(spec_path) as handle:
            spec = json.load(handle)
        for mode, name in (("house", "house.xlsx"), ("plain", "plain.xlsx")):
            local = copy.deepcopy(spec)
            local["formatting"] = mode
            local["outputPath"] = os.path.join(self.tmp, "sample_" + name)
            result = build_xlsx.build(local)
            self.assertTrue(result["ok"], "%s: %s" % (mode, result["errors"]))
            self.assertEqual(result["warnings"], [])
            wb = load_workbook(local["outputPath"])
            expected = ["Read Me", "Assumptions", "TAM Build", "Revenue Ramp", "Sources & Appendix"]
            self.assertEqual(wb.sheetnames, expected)
            self.assertEqual(sum(len(ws._charts) for ws in wb.worksheets), 2)


# --------------------------------------------------------------- quality rules
# W1 labels fit, W2 colour by cell, W3 per row formats, W4 one note column,
# W5 widths from the values, W6 computed sheet setup, W9 to W16 the rest.


def quality_spec(out_path, sheets, mode="house", readme=None, project=None, **extra):
    """A spec with nothing in it but the sheets a test needs, cover tab included.

    `extra` goes straight onto the spec, which is how a test passes `fidelity`, `comments`
    and `threads`.
    """
    spec = {
        "outputPath": out_path,
        "formatting": mode,
        "project": project or {"name": "Quality", "brief": "One brief."},
        "sheets": ([{"name": "Read Me", "blocks": []}] if mode in ("house", "skill") else []) + sheets,
    }
    if readme is not None:
        spec["readme"] = readme
    spec.update(extra)
    return spec


def pairs(rows, **kw):
    block = {"id": "b_p", "type": "inputs", "title": "Inputs", "row": 2, "col": 2, "cols": 4,
             "rows": rows}
    block.update(kw)
    return block


def table(columns, data, **kw):
    block = {"id": "b_t", "type": "table", "title": "Table", "row": 2, "col": 2,
             "cols": len(columns), "columns": columns, "data": data}
    block.update(kw)
    return block


def widths_of(ws):
    return dict((k, d.width) for k, d in ws.column_dimensions.items() if d.width)


def clipped(ws, exclude_wrapped=True):
    """Text cells whose value cannot be shown in their column, merge span and empty neighbours."""
    w = dict((get_column_letter(c), None) for c in range(1, (ws.max_column or 1) + 1))
    for letter in list(w):
        d = ws.column_dimensions.get(letter)
        w[letter] = d.width if d and d.width else 8.43
    merged = dict(((r.min_row, r.min_col), r.max_col) for r in ws.merged_cells.ranges)
    out = []
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not isinstance(v, str) or v.startswith("=") or not v.strip():
                continue
            end = merged.get((cell.row, cell.column), cell.column)
            span = sum(w[get_column_letter(x)] for x in range(cell.column, end + 1))
            if end == cell.column:
                x = cell.column + 1
                while x <= (ws.max_column or 1) and ws.cell(cell.row, x).value in (None, ""):
                    span += w[get_column_letter(x)]
                    x += 1
            if cell.alignment and cell.alignment.wrap_text and exclude_wrapped:
                lines = max(1, int((ws.row_dimensions[cell.row].height or 15.0) // 13))
                span = sum(w[get_column_letter(x)] for x in range(cell.column, end + 1)) * lines
            if len(v) > span + 0.5:
                out.append((cell.coordinate, len(v), round(span, 1)))
    return out


class QualityTests(unittest.TestCase):
    """The workbook-quality rules from docs/research/workbook-quality.md."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="workbook_quality_test_")

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def build(self, sheets, name=None, mode="house", readme=None, project=None, ok=True, **extra):
        """Build one spec and hand back (workbook, result, snapshot by sheet name)."""
        path = os.path.join(self.tmp, (name or self.id().rsplit(".", 1)[-1]) + ".xlsx")
        result = build_xlsx.build(quality_spec(path, sheets, mode, readme, project, **extra))
        self.assertEqual(result["ok"], ok, result["errors"])
        if not ok:
            return None, result, {}
        with open(result["snapshot"]) as handle:
            snap = json.load(handle)
        return (load_workbook(path), result,
                dict((s["name"], s) for s in snap["sheets"]))

    # ------------------------------------------------- W1: labels fit their column

    LONG = "Comparable EV / forward revenue multiple"       # 40 characters

    def test_w1_pair_label_column_fits_its_longest_label(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [pairs([
            {"label": "Growth", "value": 0.1, "format": "percent"},
            {"label": self.LONG, "value": 12, "format": "multiple"},
        ])]}])
        ws = wb["S"]
        self.assertEqual(ws["B3"].value, "Growth")
        self.assertGreaterEqual(ws.column_dimensions["B"].width, len(self.LONG))
        self.assertEqual(clipped(ws), [])

    def test_w1_label_column_never_below_the_style_minimum(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [pairs([
            {"label": "Rate", "value": 1}])]}])
        self.assertEqual(wb["S"].column_dimensions["B"].width,
                         build_xlsx.HOUSE_DEFAULTS["label_width"])

    def test_w1_label_column_caps_at_sixty_and_wraps_instead_of_clipping(self):
        label = "x" * 140
        wb, _, _ = self.build([{"name": "S", "blocks": [pairs([
            {"label": label, "value": 1}])]}])
        ws = wb["S"]
        self.assertEqual(ws.column_dimensions["B"].width, 60)
        self.assertTrue(ws["B3"].alignment.wrap_text)
        self.assertGreater(ws.row_dimensions[3].height, 15)
        self.assertEqual(clipped(ws), [])

    def test_w1_table_label_column_fits_its_labels(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Segment", "Year 1"], [[self.LONG, 120], ["SMB", 4]],
            colFormats=["text", "currency"])]}])
        ws = wb["S"]
        self.assertGreaterEqual(ws.column_dimensions["B"].width, len(self.LONG))
        self.assertEqual(clipped(ws), [])

    def test_w1_the_shipped_sample_clips_nothing(self):
        spec_path = os.path.join(HERE, "sample_spec.json")
        with open(spec_path) as handle:
            spec = json.load(handle)
        spec["outputPath"] = os.path.join(self.tmp, "sample_clip.xlsx")
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        wb = load_workbook(spec["outputPath"])
        for ws in wb.worksheets:
            self.assertEqual(clipped(ws), [], ws.title)

    # --------------------------------------------- W2: colour by cell, not by block

    def roles_book(self):
        return self.build([
            {"name": "Src", "blocks": [pairs([
                {"label": "Growth", "value": 0.1, "format": "percent"},
                {"label": "Open item", "value": "TBD", "format": "text"},
                {"label": "Scenario", "value": "Base",
                 "validation": {"list": ["Low", "Base", "High"]}},
                {"label": "Base year", "value": 2026, "format": "year"},
            ], id="b_in")]},
            {"name": "Calc", "blocks": [table(
                ["Line", "Count", "Share", "Note"],
                [["Micro", 4100, {"f": "=C4/1000"}, "from the file"],
                 ["Small", 1450, {"f": "={{b_in!Growth}}"}, "carried"]],
                id="b_c", type="calc", colFormats=["text", "count", "percent", "text"])]},
            {"name": "Out", "blocks": [pairs([
                {"label": "Pass through", "value": {"f": "={{b_c!0,1}}"}, "format": "count"},
                {"label": "Same sheet", "value": {"f": "=C3*2"}, "format": "count"},
            ], id="b_o", type="output")]},
        ], name="roles")

    def test_w2_a_literal_number_in_a_table_is_an_input(self):
        wb, _, snap = self.roles_book()
        cells = dict((c["a"], c) for c in snap["Calc"]["cells"])
        self.assertEqual(cells["C4"]["k"], "input")
        self.assertEqual(cells["C5"]["k"], "input")
        self.assertEqual(wb["Calc"]["C4"].font.color.rgb[-6:], build_xlsx.BLUE)

    def test_w2_a_formula_is_never_an_input(self):
        wb, _, snap = self.roles_book()
        cells = dict((c["a"], c) for c in snap["Calc"]["cells"])
        self.assertEqual(cells["D4"]["k"], "formula")
        self.assertEqual(wb["Calc"]["D4"].font.color.rgb[-6:], build_xlsx.BLACK)

    def test_w2_a_bare_cross_sheet_reference_is_a_link(self):
        wb, _, snap = self.roles_book()
        cells = dict((c["a"], c) for c in snap["Out"]["cells"])
        self.assertEqual(cells["C3"]["k"], "link")      # ='Calc'!C4
        self.assertEqual(cells["C4"]["k"], "formula")   # =C3*2, same sheet
        self.assertEqual(wb["Out"]["C3"].font.color.rgb[-6:], build_xlsx.BLACK)   # W14: the ink
        self.assertEqual(wb["Out"]["C3"].font.underline, "single")
        # and a cross-sheet reference inside a bigger formula is still a formula
        cells = dict((c["a"], c) for c in snap["Calc"]["cells"])
        self.assertEqual(cells["D5"]["k"], "link")      # ='Src'!C3 on its own
        self.assertEqual(wb["Calc"]["D5"].font.color.rgb[-6:], build_xlsx.BLACK)

    def test_w2_tbd_and_a_picklist_value_are_inputs(self):
        wb, _, snap = self.roles_book()
        cells = dict((c["a"], c) for c in snap["Src"]["cells"])
        self.assertEqual(cells["C4"]["k"], "input")     # TBD
        self.assertEqual(cells["C5"]["k"], "input")     # a dropdown value
        self.assertEqual(wb["Src"]["C5"].font.color.rgb[-6:], build_xlsx.BLUE)

    def test_w2_a_label_and_a_text_value_are_not_inputs(self):
        wb, _, snap = self.roles_book()
        cells = dict((c["a"], c) for c in snap["Calc"]["cells"])
        self.assertEqual(cells["B4"]["k"], "label")     # the table's own label column
        self.assertEqual(cells["E4"]["k"], "text")      # prose in a data column
        self.assertEqual(wb["Calc"]["E4"].font.color.rgb[-6:], build_xlsx.BLACK)

    def test_w2_a_year_is_not_an_input_unless_the_spec_says_so(self):
        _, _, snap = self.roles_book()
        cells = dict((c["a"], c) for c in snap["Src"]["cells"])
        self.assertEqual(cells["C6"], {"a": "C6", "v": "2026", "k": "text", "fmt": "@"})
        wb, _, snap = self.build([{"name": "S", "blocks": [pairs([
            {"label": "Base year", "value": 2026, "format": "year", "role": "input"}])]}],
            name="year_role")
        self.assertEqual(dict((c["a"], c) for c in snap["S"]["cells"])["C3"]["k"], "input")
        self.assertEqual(wb["S"]["C3"].font.color.rgb[-6:], build_xlsx.BLUE)

    def test_w2_cell_roles_override_one_cell(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [table(
            ["Line", "Value"], [["Typed", 5], ["Also typed", 6]],
            colFormats=["text", "count"], cellRoles={"0,1": "formula", "1,1": "note"})]}])
        cells = dict((c["a"], c) for c in snap["S"]["cells"])
        self.assertEqual(cells["C4"]["k"], "formula")
        self.assertEqual(cells["C5"]["k"], "note")
        self.assertEqual(wb["S"]["C4"].font.color.rgb[-6:], build_xlsx.BLACK)
        self.assertEqual(wb["S"]["C5"].font.color.rgb[-6:], build_xlsx.GREY)

    def test_w2_a_row_role_overrides_an_inputs_value(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [pairs([
            {"label": "Carried", "value": 5, "format": "count", "role": "note"},
            {"label": "Adjustable", "value": 6, "format": "count"},
        ])]}])
        cells = dict((c["a"], c) for c in snap["S"]["cells"])
        self.assertEqual(cells["C3"]["k"], "note")
        self.assertEqual(cells["C4"]["k"], "input")
        self.assertEqual(wb["S"]["C4"].font.color.rgb[-6:], build_xlsx.BLUE)

    def test_w2_bad_cell_roles_warn_and_change_nothing(self):
        _, result, snap = self.build([{"name": "S", "blocks": [table(
            ["Line", "Value"], [["Typed", 5]], colFormats=["text", "count"],
            cellRoles={"nope": "input", "0,1": "wizard"})]}])
        self.assertEqual(dict((c["a"], c) for c in snap["S"]["cells"])["C4"]["k"], "input")
        self.assertTrue(any("cellRoles" in w and "row,col" in w for w in result["warnings"]))
        self.assertTrue(any("cellRoles" in w and "wizard" in w for w in result["warnings"]))

    def test_w2_plain_mode_keeps_inputs_black_and_links_blue(self):
        wb, _, _ = self.build([
            {"name": "Src", "blocks": [pairs([{"label": "Growth", "value": 0.1}], id="b_in")]},
            {"name": "Out", "blocks": [pairs(
                [{"label": "Pass", "value": {"f": "={{b_in!Growth}}"}}], id="b_o", type="output")]},
        ], mode="plain", name="roles_plain")
        self.assertEqual(wb["Src"]["C3"].font.color.rgb[-6:], build_xlsx.BLACK)
        self.assertEqual(wb["Out"]["C3"].font.color.rgb[-6:], build_xlsx.LINK_BLUE)

    # -------------------------------------------- W3: a format per row on a table

    def ramp_table(self, **kw):
        block = table(["Line", "2026", "2027"],
                      [["Share of volume", 0.001, 0.003],
                       ["Accounts", 40, 120],
                       ["Price per seat", 39, 40]],
                      colFormats=["text", "percent", "percent"])
        block.update(kw)
        return block

    def test_w3_a_row_format_wins_over_the_column_format(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [
            self.ramp_table(rowFormats=["percent", "count", "currency"])]}])
        ws = wb["S"]
        self.assertEqual(ws["C4"].number_format, build_xlsx.FORMATS["percent"])
        self.assertEqual(ws["C5"].number_format, build_xlsx.FORMATS["count"])
        self.assertEqual(ws["D5"].number_format, build_xlsx.FORMATS["count"])
        self.assertEqual(ws["C6"].number_format, build_xlsx.FORMATS["currency"])
        self.assertEqual(ws["D6"].number_format, build_xlsx.FORMATS["currency"])

    def test_w3_a_row_format_leaves_the_text_label_column_alone(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [
            self.ramp_table(rowFormats=["count", "count", "count"])]}])
        ws = wb["S"]
        self.assertEqual(ws["B4"].number_format, "@")
        self.assertEqual(ws["B4"].value, "Share of volume")
        self.assertEqual(ws["C4"].number_format, build_xlsx.FORMATS["count"])

    def test_w3_a_null_row_format_falls_back_to_the_column(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [
            self.ramp_table(rowFormats=[None, "count"])]}])
        ws = wb["S"]
        self.assertEqual(ws["C4"].number_format, build_xlsx.FORMATS["percent"])
        self.assertEqual(ws["C5"].number_format, build_xlsx.FORMATS["count"])
        self.assertEqual(ws["C6"].number_format, build_xlsx.FORMATS["percent"])  # past the list

    def test_w3_a_per_cell_format_wins_over_the_row_and_the_column(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Line", "2026"], [["Mixed", {"v": 2.4, "format": "multiple"}]],
            colFormats=["text", "percent"], rowFormats=["currency"])]}])
        self.assertEqual(wb["S"]["C4"].number_format, build_xlsx.FORMATS["multiple"])

    def test_w3_row_formats_that_are_not_a_list_warn(self):
        wb, result, _ = self.build([{"name": "S", "blocks": [
            self.ramp_table(rowFormats="count")]}])
        self.assertTrue(any("rowFormats must be a list" in w for w in result["warnings"]))
        self.assertEqual(wb["S"]["C5"].number_format, build_xlsx.FORMATS["percent"])

    def test_w3_the_total_row_keeps_the_column_formats(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [self.ramp_table(
            rowFormats=["percent", "count", "currency"],
            totalRow=["Total", {"f": "=SUM(C4:C6)"}, {"f": "=SUM(D4:D6)"}])]}])
        self.assertEqual(wb["S"]["C7"].number_format, build_xlsx.FORMATS["percent"])

    # ------------------------------------------- W4: one note column per row band

    def test_w4_the_note_sits_next_to_the_block_with_no_gap(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [pairs([
            {"label": "Growth", "value": 0.1, "note": "Analyst estimate"}],
            note="Owned by finance")]}])
        ws = wb["S"]
        self.assertEqual(ws["D2"].value, "Owned by finance")   # right of label and value
        self.assertEqual(ws["D3"].value, "Analyst estimate")
        self.assertIsNone(ws["E3"].value)
        self.assertEqual(snap["S"]["blocks"][0]["cols"], 3)    # label, value, note

    def test_w4_one_note_column_serves_every_block_in_a_row_band(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [
            pairs([{"label": "Growth", "value": 0.1, "note": "left block"}], id="b_l"),
            dict(table(["Seg", "N"], [["SMB", 1]], id="b_r"), row=2, col=5,
                 note="right block"),
        ]}])
        ws = wb["S"]
        self.assertEqual(ws["G2"].value, "right block")    # b_r is F..G? no: E..F, so notes at G
        self.assertEqual(ws["G3"].value, "left block")

    def test_w4_notes_merge_to_the_sheets_last_column_plus_two(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [pairs([
            {"label": "Growth", "value": 0.1, "note": "Analyst estimate"}])]}])
        self.assertIn("D3:F3", snap["S"]["merges"])
        self.assertFalse(wb["S"]["D3"].alignment.wrap_text)     # a note has its merge, not wrapping

    def test_w4_the_title_rule_never_spans_the_note_column(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [table(
            ["Seg", "N"], [["SMB", 1]], rowNotes=["from the CRM"])]}])
        merges = snap["S"]["merges"]
        self.assertIn("B2:C2", merges)
        self.assertFalse([m for m in merges if m.startswith("B2:D")])

    def test_w4_a_note_never_writes_over_a_neighbour(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [
            pairs([{"label": "Growth", "value": 0.1, "note": "Analyst estimate"}], id="b_l"),
            dict(pairs([{"label": "Side", "value": 2}], id="b_r"), row=3, col=6, cols=2),
        ]}])
        ws = wb["S"]
        # the right block shares the rows, so the band's note column is right of it
        self.assertEqual(ws["H3"].value, "Analyst estimate")
        self.assertEqual(ws["F3"].value, "Inputs")      # the neighbour's own title row
        self.assertEqual(ws["F4"].value, "Side")
        self.assertEqual(ws["G4"].value, 2)
        self.assertIn("H3:J3", snap["S"]["merges"])
        for merge in snap["S"]["merges"]:
            first = merge.split(":")[0]
            self.assertIsNotNone(ws[first].value, merge)

    def test_w4_two_notes_in_one_cell_warn(self):
        _, result, _ = self.build([{"name": "S", "blocks": [
            pairs([{"label": "Growth", "value": 0.1, "note": "first"}], id="b_l"),
            dict(pairs([{"label": "Side", "value": 2, "note": "second"}], id="b_r"),
                 row=2, col=5, cols=2),
        ]}])
        self.assertTrue(any("a note is already in" in w for w in result["warnings"]))

    # ------------------------------------------------ W5: widths from the values

    def test_w5_a_numeric_column_is_as_wide_as_its_formatted_values(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Segment", "Share of firms, normalized"],
            [["SMB", 0.125], ["Enterprise", 0.875]],
            colFormats=["text", "percent"])]}])
        ws = wb["S"]
        self.assertLessEqual(ws.column_dimensions["C"].width, build_xlsx.HOUSE_DEFAULTS["value_cap"])
        self.assertGreaterEqual(ws.column_dimensions["C"].width, len("12.5%"))

    def test_w5_a_long_header_wraps_instead_of_widening_its_column(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Segment", "Share of firms, normalized"],
            [["SMB", 0.125]], colFormats=["text", "percent"])]}])
        ws = wb["S"]
        self.assertTrue(ws["C3"].alignment.wrap_text)
        self.assertGreater(ws.row_dimensions[3].height, 15)
        self.assertLessEqual(ws.column_dimensions["C"].width, 20)
        self.assertEqual(clipped(ws), [])

    def test_w5_a_wide_number_still_gets_the_room_it_needs(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Segment", "Volume"], [["SMB", 1234567890]],
            colFormats=["text", "count"])]}])
        self.assertGreaterEqual(wb["S"].column_dimensions["C"].width, len("1,234,567,890"))

    def test_w5_a_text_cell_at_the_cap_wraps_and_the_row_grows(self):
        prose = ("Small business profile, firms with employees, taken from the published "
                 "profile and the county business patterns tables")
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Ref", "Detail"], [["R1", prose]], colFormats=["text", "text"])]}])
        ws = wb["S"]
        self.assertEqual(ws.column_dimensions["C"].width, build_xlsx.HOUSE_DEFAULTS["label_cap"])
        self.assertTrue(ws["C4"].alignment.wrap_text)
        self.assertGreaterEqual(ws.row_dimensions[4].height, 2 * 13)
        self.assertEqual(clipped(ws), [])

    def test_w5_wrapping_survives_the_cell_styling(self):
        prose = "y" * 130
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Ref", "Detail"], [["R1", prose]], colFormats=["text", "text"])]}])
        cell = wb["S"]["C4"]
        self.assertTrue(cell.alignment.wrap_text)
        self.assertEqual(cell.border.right.style, "thin")   # still a bordered body cell
        self.assertEqual(cell.font.name, "Arial")

    def test_w5_a_header_does_not_drag_a_note_column_wider(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Seg", "N"], [["SMB", 1]], rowNotes=["Derived: annual customers x average spend"])]}])
        ws = wb["S"]
        self.assertLessEqual(ws.column_dimensions["D"].width,
                             build_xlsx.HOUSE_DEFAULTS["note_width"])

    # ----------------------------------------------- W6: sheet setup, computed

    def setup_book(self, **sheet_kw):
        sheet = {"name": "S", "blocks": [
            pairs([{"label": "Growth", "value": 0.1, "format": "percent"}], id="b_in"),
            dict(table(["Segment", "Year 1", "Year 2"],
                       [["SMB", 1, 2], ["Mid", 3, 4]], id="b_t"), row=6, col=2),
        ]}
        sheet.update(sheet_kw)
        return self.build([sheet], name="setup_%d" % len(sheet_kw))

    def test_w6_freeze_lands_on_the_first_data_cell_of_the_first_table(self):
        wb, _, _ = self.setup_book()
        self.assertEqual(wb["S"].freeze_panes, "C8")     # header row 7, label column B

    def test_w6_a_sheet_of_pairs_freezes_under_its_title(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [pairs(
            [{"label": "Growth", "value": 0.1}])]}])
        self.assertEqual(wb["S"].freeze_panes, "C3")

    def test_w6_print_area_is_the_used_range(self):
        wb, _, _ = self.setup_book()
        self.assertEqual(wb["S"].print_area, "'S'!$B$2:$D$9")

    def test_w6_one_page_wide_and_portrait_until_the_sheet_is_too_wide(self):
        wb, _, _ = self.setup_book()
        ws = wb["S"]
        self.assertEqual((ws.page_setup.fitToWidth, ws.page_setup.fitToHeight), (1, 0))
        self.assertTrue(ws.sheet_properties.pageSetUpPr.fitToPage)
        self.assertEqual(ws.page_setup.orientation, "portrait")

    def test_w6_a_wide_sheet_turns_landscape(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Segment"] + ["Year %d" % i for i in range(1, 9)],
            [["SMB"] + [1000000 * i for i in range(1, 9)]],
            colFormats=["text"] + ["currency"] * 8)]}])
        ws = wb["S"]
        self.assertEqual(ws.page_setup.orientation, "landscape")
        self.assertGreater(sum(widths_of(ws).values()), build_xlsx.PORTRAIT_WIDTH)

    def test_w6_a_long_sheet_repeats_its_title_rows(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Segment", "Value"], [["Row %d" % i, i] for i in range(60)],
            colFormats=["text", "count"])]}])
        self.assertEqual(wb["S"].print_title_rows, "$2:$3")
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Segment", "Value"], [["Row %d" % i, i] for i in range(3)],
            colFormats=["text", "count"])]}], name="short")
        self.assertFalse(wb["S"].print_title_rows)

    def test_w6_the_spec_still_wins_on_every_field(self):
        wb, _, _ = self.setup_book(freeze="B4", printArea="B2:C3", landscape=True,
                                   fitToWidth=False, zoom=80, titleRows="2:2")
        ws = wb["S"]
        self.assertEqual(ws.freeze_panes, "B4")
        self.assertEqual(ws.print_area, "'S'!$B$2:$C$3")
        self.assertEqual(ws.page_setup.orientation, "landscape")
        self.assertFalse(ws.sheet_properties.pageSetUpPr and ws.sheet_properties.pageSetUpPr.fitToPage)
        self.assertEqual(ws.sheet_view.zoomScale, 80)
        self.assertEqual(ws.print_title_rows, "$2:$2")

    def test_w6_the_cover_gets_a_print_area_and_no_freeze(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [pairs(
            [{"label": "Growth", "value": 0.1}])]}])
        rm = wb["Read Me"]
        self.assertIsNone(rm.freeze_panes)
        self.assertTrue(rm.print_area.startswith("'Read Me'!$A$1:$C$"))
        self.assertEqual(rm.page_setup.orientation, "portrait")

    # ------------------------------------------- W9: a named range per input cell

    def test_w9_every_input_cell_gets_a_defined_name(self):
        wb, _, _ = self.build([{"name": "Src", "blocks": [pairs([
            {"label": "Take rate", "value": 0.003, "format": "percent"},
            {"label": "Derived", "value": {"f": "=C3*2"}, "format": "percent"},
        ], id="b_in")]}])
        names = dict((k, v.attr_text) for k, v in wb.defined_names.items())
        self.assertEqual(names["Take_rate"], "'Src'!$C$3")
        self.assertNotIn("Derived", names)

    def test_w9_a_table_input_is_named_from_its_row_and_column(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Segment", "Accounts (K)"], [["Micro, 5 to 19 staff", 4100]],
            colFormats=["text", "count"])]}])
        self.assertIn("Micro_5_to_19_staff_Accounts_K", wb.defined_names)

    def test_w9_names_are_unique_and_legal(self):
        wb, _, _ = self.build([
            {"name": "One", "blocks": [pairs([{"label": "Growth", "value": 1},
                                              {"label": "2026", "value": 2},
                                              {"label": "A1", "value": 3}], id="b_1")]},
            {"name": "Two", "blocks": [pairs([{"label": "Growth", "value": 4}], id="b_2")]},
        ], name="names")
        names = dict((k, v.attr_text) for k, v in wb.defined_names.items())
        self.assertEqual(names["Growth"], "'One'!$C$3")
        self.assertEqual(names["Two_Growth"], "'Two'!$C$3")
        self.assertEqual(names["_2026"], "'One'!$C$4")     # a name cannot start with a digit
        self.assertEqual(names["A1_"], "'One'!$C$5")       # nor look like a cell reference
        for name in names:
            self.assertTrue(build_xlsx.defined_name(name) == name, name)

    # ------------------------------------------ W10: three part number formats

    def test_w10_a_zero_shows_as_a_dash(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [table(
            ["Line", "Value"], [["Empty", 0], ["Negative", -1200]],
            colFormats=["text", "currency"])]}])
        cells = dict((c["a"], c) for c in snap["S"]["cells"])
        self.assertEqual(wb["S"]["C4"].number_format, '$#,##0;($#,##0);"-"')
        self.assertEqual(cells["C4"]["d"], "-")
        self.assertEqual(cells["C5"]["d"], "($1,200)")

    def test_w10_an_unformatted_number_still_gets_a_three_part_format(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Line", "Value"], [["Count", 12], ["Decimal", 1.5]])]}])
        self.assertEqual(wb["S"]["C4"].number_format, build_xlsx.FORMATS["count"])
        self.assertEqual(wb["S"]["C5"].number_format, build_xlsx.FORMATS["number"])

    def test_w10_a_format_the_builder_does_not_know_warns(self):
        wb, result, _ = self.build([{"name": "S", "blocks": [table(
            ["Line", "Value"], [["Score", 4]], colFormats=["text", "score"])]}])
        self.assertTrue(any("format 'score' is not one of" in w for w in result["warnings"]))
        self.assertEqual(wb["S"]["C4"].number_format, build_xlsx.FORMATS["number"])

    # ------------------------------------- W11: the cover carries the answer

    def cover_book(self, **readme):
        sheets = [
            {"name": "Assumptions", "purpose": "Every input the model reads",
             "blocks": [pairs([{"label": "Growth", "value": 0.1}], id="b_in")]},
            {"name": "Revenue Ramp", "blocks": [pairs(
                [{"label": "Year 5 revenue ($M)", "value": {"f": "=1+1"}, "format": "currency"}],
                id="b_out", type="output")]},
        ]
        base = {"purpose": "Sizes the market."}
        base.update(readme)
        return self.build(sheets, name="cover_%d" % len(readme), readme=base)

    def test_w11_the_cover_carries_the_answer_line(self):
        wb, _, snap = self.cover_book(answer="The market is $1.9B and the ramp reaches $124M.")
        rm = wb["Read Me"]
        column_a = [rm.cell(r, 1).value for r in range(1, rm.max_row + 1)]
        self.assertIn("Answer", column_a)
        self.assertTrue(any(v and "1.9B" in str(v) for v in column_a))
        self.assertEqual(set(c["k"] for c in snap["Read Me"]["cells"]), set(["readme"]))

    def test_w11_the_takeaway_from_the_project_is_the_fallback(self):
        wb, _, _ = self.build(
            [{"name": "S", "blocks": [pairs([{"label": "Growth", "value": 0.1}])]}],
            project={"name": "Q", "brief": "b", "takeaway": "Ship it in the second half."},
            name="takeaway")
        rm = wb["Read Me"]
        self.assertTrue(any(str(rm.cell(r, 1).value or "").startswith("Ship it")
                            for r in range(1, rm.max_row + 1)))

    def test_w11_results_are_live_references_to_the_output_blocks(self):
        wb, _, _ = self.cover_book(results=True)
        rm = wb["Read Me"]
        rows = [r for r in range(1, rm.max_row + 1) if rm.cell(r, 1).value == "Year 5 revenue ($M)"]
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(rm.cell(row, 2).value, "='Revenue Ramp'!C3")
        self.assertEqual(rm.cell(row, 2).number_format, build_xlsx.FORMATS["currency"])
        self.assertEqual(rm.cell(row, 3).value, "Revenue Ramp")

    def test_w11_no_result_rows_unless_the_payload_asks(self):
        wb, _, _ = self.cover_book()
        rm = wb["Read Me"]
        self.assertFalse([r for r in range(1, rm.max_row + 1)
                          if str(rm.cell(r, 2).value or "").startswith("=")])

    def test_w11_the_cover_has_one_contents_row_per_tab(self):
        wb, _, _ = self.cover_book(contents=[{"tab": "Assumptions", "holds": "The inputs"},
                                             {"tab": "Gone", "holds": "Not a tab"}])
        rm = wb["Read Me"]
        holds = dict((rm.cell(r, 1).value, rm.cell(r, 2).value) for r in range(1, rm.max_row + 1))
        self.assertEqual(holds["Assumptions"], "The inputs")
        self.assertIn("Revenue Ramp", holds)          # no contents entry, still a row
        self.assertNotIn("Gone", holds)

    def test_w11_a_contents_entry_for_no_tab_warns(self):
        _, result, _ = self.cover_book(contents=[{"tab": "Gone", "holds": "Not a tab"}])
        self.assertTrue(any("not a tab" in w for w in result["warnings"]))

    def test_w11_a_tab_purpose_fills_the_contents_row(self):
        wb, _, _ = self.cover_book()
        rm = wb["Read Me"]
        holds = dict((rm.cell(r, 1).value, rm.cell(r, 2).value) for r in range(1, rm.max_row + 1))
        self.assertEqual(holds["Assumptions"], "Every input the model reads")

    def test_w11_the_cover_prints_portrait_on_one_page(self):
        wb, _, snap = self.cover_book()
        self.assertEqual(snap["Read Me"]["widths"], {"A": 30, "B": 45, "C": 14})
        self.assertEqual(sum(snap["Read Me"]["widths"].values()), 89)

    # -------------------------------------------------- W12: chart geometry

    def test_w12_a_chart_covers_exactly_its_block_rect(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [
            table(["Segment", "Year 1"], [["SMB", 1], ["Mid", 2]],
                  id="b_t", colFormats=["text", "currency"]),
            {"id": "b_c", "type": "chart", "title": "By segment", "row": 8, "col": 2,
             "cols": 5, "rows": 12,
             "chart": {"kind": "bar", "sourceBlock": "b_t", "categoriesCol": 0,
                       "seriesCols": [1]}},
        ]}])
        chart = wb["S"]._charts[0]
        anchor = chart.anchor
        self.assertEqual((anchor._from.col, anchor._from.row), (1, 7))     # B8, zero based
        self.assertEqual((anchor.to.col, anchor.to.row), (1 + 5, 7 + 12))  # five columns, twelve rows
        self.assertEqual(snap["S"]["charts"][0]["anchor"], "B8")

    def test_w12_a_single_series_chart_has_no_legend_and_no_gridlines(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [
            table(["Segment", "Year 1"], [["SMB", 1], ["Mid", 2]], id="b_t"),
            {"id": "b_c", "type": "chart", "title": "By segment", "row": 8, "col": 2,
             "cols": 4, "rows": 10,
             "chart": {"kind": "bar", "sourceBlock": "b_t", "categoriesCol": 0,
                       "seriesCols": [1]}},
        ]}])
        chart = wb["S"]._charts[0]
        self.assertIsNone(chart.legend)
        self.assertIsNone(chart.y_axis.majorGridlines)
        self.assertIsNotNone(chart.layout.manualLayout)

    # ------------------------------------------- W13: one block width per sheet

    def test_w13_a_narrower_table_snaps_to_the_widest_and_is_recorded(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [
            dict(table(["Line", "2026", "2027", "2028"],
                       [["Wide", 1, 2, 3]], id="b_wide"), row=2, col=2),
            dict(table(["Line", "2026", "2027"], [["Narrow", 1, 2]], id="b_narrow"),
                 row=6, col=2),
        ]}])
        blocks = dict((b["id"], b) for b in snap["S"]["blocks"])
        self.assertEqual(blocks["b_narrow"]["cols"], 4)
        self.assertEqual(blocks["b_wide"]["cols"], 4)
        self.assertIn({"blockId": "b_narrow", "from": 3, "to": 4, "reason": "snap"},
                      snap["S"]["adjusted"])
        merges = snap["S"]["merges"]
        self.assertIn("B2:E2", merges)
        self.assertIn("B6:E6", merges)      # both title rules end on the same edge

    def test_w13_a_move_of_more_than_two_columns_is_not_taken(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [
            dict(table(["Line"] + ["Y%d" % i for i in range(7)],
                       [["Wide"] + [1] * 7], id="b_wide"), row=2, col=2),
            dict(table(["Line", "2026"], [["Narrow", 1]], id="b_narrow"), row=6, col=2),
        ]}])
        blocks = dict((b["id"], b) for b in snap["S"]["blocks"])
        self.assertEqual(blocks["b_narrow"]["cols"], 2)
        self.assertEqual(blocks["b_wide"]["cols"], 8)

    def test_w13_a_pair_block_keeps_its_two_columns(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [
            pairs([{"label": "Growth", "value": 0.1}], id="b_p", cols=4),
            dict(table(["Line", "2026", "2027"], [["Row", 1, 2]], id="b_t"), row=6, col=2),
        ]}])
        blocks = dict((b["id"], b) for b in snap["S"]["blocks"])
        self.assertEqual(blocks["b_p"]["cols"], 2)
        self.assertIn({"blockId": "b_p", "from": 4, "to": 2, "reason": "fit"},
                      snap["S"]["adjusted"])
        self.assertIn("B2:C2", snap["S"]["merges"])

    def test_w13_a_table_narrower_than_its_plan_loses_the_empty_column(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [dict(table(
            ["Line", "2026"], [["Row", 1]], id="b_t", rowNotes=["a note"]), cols=4)]}])
        self.assertEqual(snap["S"]["blocks"][0]["cols"], 3)     # two columns plus the note
        self.assertEqual(wb["S"]["D4"].value, "a note")         # adjacent, not at E

    # --------------------------------------------------- W14: the link cell kind

    def test_w14_a_link_cell_is_a_hyperlink_in_the_link_colour(self):
        url = "https://www.census.gov/programs-surveys/cbp.html"
        wb, _, snap = self.build([{"name": "S", "blocks": [table(
            ["Ref", "Link"], [["R1", {"v": "Census CBP", "link": url}]],
            colFormats=["text", "text"])]}])
        cell = wb["S"]["C4"]
        self.assertEqual(cell.value, "Census CBP")
        self.assertEqual(cell.hyperlink.target, url)
        self.assertEqual(cell.font.color.rgb[-6:], build_xlsx.BLACK)   # W14: the house ink
        self.assertEqual(cell.font.underline, "single")
        self.assertEqual(dict((c["a"], c) for c in snap["S"]["cells"])["C4"]["k"], "link")

    def test_w14_a_bare_url_becomes_a_hyperlink_too(self):
        url = "https://advocacy.sba.gov/"
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Ref", "Link"], [["R1", url]], colFormats=["text", "text"])]}])
        self.assertEqual(wb["S"]["C4"].hyperlink.target, url)

    def test_w14_an_internal_link_points_at_a_cell(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [table(
            ["Ref", "Link"], [["R1", {"v": "See the ramp", "link": "'Ramp'!B4"}]],
            colFormats=["text", "text"])]}])
        cell = wb["S"]["C4"]
        self.assertIsNone(cell.hyperlink.target)
        self.assertEqual(cell.hyperlink.location, "'Ramp'!B4")

    def link_book(self, mode="house", name=None, **extra):
        """A sheet with a hyperlink cell and a bare cross-sheet pass-through on it."""
        url = "https://advocacy.sba.gov/"
        sheets = [
            {"name": "Src", "blocks": [pairs([{"label": "Growth", "value": 0.1}], id="b_in")]},
            {"name": "S", "blocks": [table(
                ["Ref", "Link", "Pulled"],
                [["R1", {"v": "SBA", "link": url}, {"f": "={{b_in!Growth}}"}]],
                colFormats=["text", "text", "percent"])]},
        ]
        return self.build(sheets, name=name or self.id().rsplit(".", 1)[-1], mode=mode, **extra)

    def test_w14_house_draws_both_kinds_of_link_in_the_ink_colour(self):
        """The house skill bans green: the link takes the ink and the underline marks it."""
        wb, _, snap = self.link_book()
        cells = dict((c["a"], c) for c in snap["S"]["cells"])
        self.assertEqual(build_xlsx.HOUSE_DEFAULTS["link_color"], "")      # "" means the ink
        self.assertEqual(build_xlsx.HouseStyler().link_color, build_xlsx.BLACK)
        for a in ("C4", "D4"):                     # the hyperlink and the pass-through
            self.assertEqual(cells[a]["k"], "link", a)
            self.assertEqual(wb["S"][a].font.color.rgb[-6:], build_xlsx.BLACK, a)
            self.assertEqual(wb["S"][a].font.underline, "single", a)

    def test_w14_plain_keeps_office_hyperlink_blue(self):
        wb, _, _ = self.link_book(mode="plain")
        for a in ("C4", "D4"):
            self.assertEqual(wb["S"][a].font.color.rgb[-6:], build_xlsx.LINK_BLUE, a)
            self.assertEqual(wb["S"][a].font.underline, "single", a)

    def test_w14_a_skill_link_colour_wins_over_the_ink(self):
        path = os.path.join(self.tmp, "linked.md")
        write_text(path, "---\nname: linked\nformat:\n  link_color: 0563C1\n---\n# linked\n")
        wb, result, _ = self.link_book(mode="skill", skill={"path": path})
        self.assertEqual(result["warnings"], [])
        self.assertEqual(wb["S"]["C4"].font.color.rgb[-6:], build_xlsx.LINK_BLUE)

    def test_w14_link_underline_off_still_underlines_a_real_hyperlink(self):
        """A pass-through is only underlined when the style asks; a hyperlink always is."""
        path = os.path.join(self.tmp, "nounderline.md")
        write_text(path, "---\nname: flat\nformat:\n  link_underline: false\n---\n# flat\n")
        wb, result, _ = self.link_book(mode="skill", skill={"path": path})
        self.assertEqual(result["warnings"], [])
        self.assertEqual(wb["S"]["C4"].font.underline, "single")     # the hyperlink
        self.assertIsNone(wb["S"]["D4"].font.underline)              # the pass-through

    def test_w14_a_skill_can_ask_for_the_ink_by_name(self):
        path = os.path.join(self.tmp, "inked.md")
        write_text(path, "---\nname: inked\nformat:\n  formula_color: 333333\n"
                         "  link_color: none\n---\n# inked\n")
        wb, result, _ = self.link_book(mode="skill", skill={"path": path})
        self.assertEqual(result["warnings"], [])
        self.assertEqual(wb["S"]["C4"].font.color.rgb[-6:], "333333")

    def test_w14_a_template_draws_a_link_in_its_own_ink(self):
        path = os.path.join(self.tmp, "link_template.xlsx")
        wb = Workbook()
        body = wb.active
        body.title = "Sample"
        body["A1"] = "Body copy"
        body["A1"].font = build_xlsx.Font(name="Georgia", size=12)
        style = wb.create_sheet("Style")
        style["A1"], style["B1"] = "Formula", "1,234"
        style["B1"].font = build_xlsx.Font(name="Georgia", size=12, color="333F50")
        wb.save(path)
        book, result, _ = self.link_book(mode="template", templatePath=path)
        self.assertEqual(book["S"]["C4"].font.color.rgb[-6:], "333F50")
        self.assertEqual(book["S"]["C4"].font.underline, "single")

    def test_w14_a_link_that_is_neither_warns_and_stays_text(self):
        wb, result, snap = self.build([{"name": "S", "blocks": [table(
            ["Ref", "Link"], [["R1", {"v": "Somewhere", "link": "over the rainbow"}]],
            colFormats=["text", "text"])]}])
        self.assertTrue(any("is not a URL or a Sheet!A1 reference" in w for w in result["warnings"]))
        self.assertIsNone(wb["S"]["C4"].hyperlink)
        self.assertEqual(dict((c["a"], c) for c in snap["S"]["cells"])["C4"]["k"], "text")

    # ----------------------------------------------------- W16: tab colours

    def test_w16_tab_colour_by_sheet_role(self):
        wb, _, _ = self.build([
            {"name": "Assumptions", "blocks": [pairs([{"label": "Growth", "value": 0.1}], id="b_i")]},
            {"name": "Build", "blocks": [table(["Line", "N"], [["Row", 1]], id="b_t")]},
            {"name": "Outcome", "blocks": [pairs([{"label": "Answer", "value": {"f": "=1+1"}}],
                                                 id="b_o", type="output")]},
        ], name="tabs")
        colour = lambda name: (wb[name].sheet_properties.tabColor.rgb[-6:]
                               if wb[name].sheet_properties.tabColor else None)
        self.assertEqual(colour("Read Me"), build_xlsx.HOUSE_DEFAULTS["tab_cover"])
        self.assertEqual(colour("Assumptions"), build_xlsx.HOUSE_DEFAULTS["tab_inputs"])
        self.assertEqual(colour("Outcome"), build_xlsx.HOUSE_DEFAULTS["tab_output"])
        self.assertIsNone(colour("Build"))          # a calc tab keeps Excel's default

    def test_w16_plain_mode_leaves_every_tab_alone(self):
        wb, _, _ = self.build([{"name": "S", "blocks": [pairs(
            [{"label": "Growth", "value": 0.1}])]}], mode="plain", name="tabs_plain")
        self.assertIsNone(wb["S"].sheet_properties.tabColor)

    def test_w16_a_skill_can_change_or_drop_the_tab_colours(self):
        path = os.path.join(self.tmp, "tabs_skill.md")
        write_text(path, "---\nname: tinted\nformat:\n  tab_inputs: 7030A0\n"
                         "  tab_cover: none\n  link_color: 0563C1\n---\n# tinted\n")
        out = os.path.join(self.tmp, "tabs_skill.xlsx")
        spec = quality_spec(out, [
            {"name": "Assumptions", "blocks": [pairs([{"label": "Growth", "value": 0.1}], id="b_i")]},
        ], mode="skill")
        spec["skill"] = {"path": path}
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        self.assertEqual(result["warnings"], [])
        wb = load_workbook(out)
        self.assertEqual(wb["Assumptions"].sheet_properties.tabColor.rgb[-6:], "7030A0")
        self.assertIsNone(wb["Read Me"].sheet_properties.tabColor)


# ------------------------------------------------------------- W11 fidelity


LONG_PARA = ("Every figure on this tab is an assumption until a source row carries it, and "
             "the ones that are not sourced are the ones to argue about first, so read the "
             "note beside each of them before you change anything at all in this model.")


def text_block(body=LONG_PARA, **kw):
    block = {"id": "b_x", "type": "text", "title": "Scope", "row": 2, "col": 2, "cols": 2,
             "text": body}
    block.update(kw)
    return block


def chart_block(source, **kw):
    block = {"id": "b_ch", "type": "chart", "title": "By year", "row": 10, "col": 2,
             "cols": 3, "rows": 10,
             "chart": {"kind": "bar", "sourceBlock": source, "categoriesCol": 0,
                       "seriesCols": [1]}}
    block.update(kw)
    return block


class FidelityTests(unittest.TestCase):
    """W11: what `balanced` may move, what `strict` may not, and how it is recorded."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="workbook_fidelity_test_")

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def build(self, sheets, name=None, ok=True, **extra):
        path = os.path.join(self.tmp, (name or self.id().rsplit(".", 1)[-1]) + ".xlsx")
        result = build_xlsx.build(quality_spec(path, sheets, "house", **extra))
        self.assertEqual(result["ok"], ok, result["errors"])
        if not ok:
            return None, result, {}
        with open(result["snapshot"]) as handle:
            snap = json.load(handle)
        return (load_workbook(path), result,
                dict((s["name"], s) for s in snap["sheets"]))

    @staticmethod
    def moves(result, reason=None):
        out = result.get("adjusted", [])
        return [m for m in out if reason is None or m["reason"] == reason]

    @staticmethod
    def blocks(snap, sheet="S"):
        return dict((b["id"], b) for b in snap[sheet]["blocks"])

    def stack(self, rows, **kw):
        """One table per row number, three columns wide, stacked down the sheet."""
        return [{"name": "S", "blocks": [
            dict(table(["Line", "2026", "2027"], [["A", 1, 2], ["B", 3, 4]],
                       id="b_%d" % i), row=r)
            for i, r in enumerate(rows)]}]

    # ------------------------------------------------------- the setting

    def test_w11_fidelity_defaults_to_balanced(self):
        _, result, _ = self.build(self.stack([2]))
        self.assertEqual(result["fidelity"], "balanced")
        self.assertEqual(result["warnings"], [])

    def test_w11_strict_is_reported_as_asked(self):
        _, result, _ = self.build(self.stack([2]), fidelity="strict")
        self.assertEqual(result["fidelity"], "strict")

    def test_w11_an_unknown_fidelity_warns_and_falls_back(self):
        _, result, _ = self.build(self.stack([2]), fidelity="loose")
        self.assertEqual(result["fidelity"], "balanced")
        self.assertTrue(any("fidelity 'loose' is not one of" in w for w in result["warnings"]),
                        result["warnings"])

    # --------------------------------------------------------- (d) spacing

    def test_w11_balanced_opens_a_blank_row_between_stacked_blocks(self):
        wb, result, snap = self.build(self.stack([2, 6]))
        self.assertEqual(self.blocks(snap)["b_1"]["row"], 7)      # 6 would touch the block above
        self.assertIsNone(wb["S"]["B6"].value)
        move = self.moves(result, "shift")[0]
        self.assertEqual(move["blockId"], "b_1")
        self.assertEqual(move["from"], {"row": 6, "col": 2, "rows": 4, "cols": 3})
        self.assertEqual(move["to"], {"row": 7, "col": 2, "rows": 4, "cols": 3})
        self.assertEqual(move["sheet"], "S")

    def test_w11_a_shift_cascades_down_the_sheet(self):
        _, result, snap = self.build(self.stack([2, 6, 10]))
        rows = [self.blocks(snap)["b_%d" % i]["row"] for i in range(3)]
        self.assertEqual(rows, [2, 7, 12])                        # one blank row between each
        self.assertEqual(len(self.moves(result, "shift")), 2)

    def test_w11_no_block_shifts_more_than_two_rows(self):
        """The fourth block of a stack would need three rows to clear, so it stops at two."""
        _, result, snap = self.build(self.stack([2, 6, 10, 14]))
        rows = [self.blocks(snap)["b_%d" % i]["row"] for i in range(4)]
        self.assertEqual(rows, [2, 7, 12, 16])
        for move in self.moves(result, "shift"):
            self.assertLessEqual(move["to"]["row"] - move["from"]["row"], 2, move)

    def test_w11_a_plan_that_already_leaves_a_blank_row_is_not_touched(self):
        _, result, snap = self.build(self.stack([2, 7]))
        self.assertEqual(self.blocks(snap)["b_1"]["row"], 7)
        self.assertEqual(self.moves(result, "shift"), [])

    def test_w11_a_shift_onto_a_block_that_cannot_move_is_not_taken(self):
        """The bottom block is pinned by a tall neighbour, so the stack stays as planned."""
        blocks = [dict(table(["Line", "2026"], [["A", 1]], id="b_a"), row=2),
                  dict(table(["Line", "2026"], [["A", 1]], id="b_b"), row=5),
                  {"id": "b_tall", "type": "chart", "title": "Tall", "row": 2, "col": 6,
                   "cols": 3, "rows": 20,
                   "chart": {"kind": "bar", "sourceBlock": "b_a", "categoriesCol": 0,
                             "seriesCols": [1]}}]
        _, result, snap = self.build([{"name": "S", "blocks": blocks}])
        self.assertEqual(self.blocks(snap)["b_b"]["row"], 6)   # one row of room, and it took it
        self.assertEqual(self.blocks(snap)["b_tall"]["row"], 2)

    def test_w11_strict_leaves_stacked_blocks_touching(self):
        wb, result, snap = self.build(self.stack([2, 6]), fidelity="strict")
        self.assertEqual(self.blocks(snap)["b_1"]["row"], 6)
        self.assertEqual(self.moves(result, "shift"), [])

    def test_w11_a_plan_that_overlaps_is_still_an_error(self):
        """Spacing opens a blank row; it never papers over two blocks written on each other."""
        _, result, _ = self.build(self.stack([2, 4]), ok=False)
        self.assertTrue(any("overlap" in e for e in result["errors"]), result["errors"])

    # ----------------------------------------------------------- (a) widen

    def test_w11_a_long_paragraph_takes_the_free_columns_beside_it(self):
        wb, result, snap = self.build([{"name": "S", "blocks": [text_block()]}])
        self.assertEqual(self.blocks(snap)["b_x"]["cols"], 4)    # two columns, plus two taken
        move = self.moves(result, "widen")[0]
        self.assertEqual((move["from"]["cols"], move["to"]["cols"]), (2, 4))
        self.assertIn("B3:E3", snap["S"]["merges"])

    def test_w11_a_paragraph_that_fits_its_columns_is_left_alone(self):
        _, result, snap = self.build([{"name": "S", "blocks": [text_block("Short enough.")]}])
        self.assertEqual(self.blocks(snap)["b_x"]["cols"], 2)
        self.assertEqual(self.moves(result, "widen"), [])

    def test_w11_a_widen_never_takes_a_column_a_neighbour_holds(self):
        blocks = [text_block(), dict(table(["Line", "2026"], [["A", 1]], id="b_t"),
                                     row=2, col=4)]
        _, result, snap = self.build([{"name": "S", "blocks": blocks}])
        self.assertEqual(self.blocks(snap)["b_x"]["cols"], 2)
        self.assertEqual(self.moves(result, "widen"), [])

    def test_w11_a_chart_grows_to_the_table_it_plots(self):
        blocks = [dict(table(["Line", "2026", "2027", "2028", "2029", "2030"],
                             [["A", 1, 2, 3, 4, 5]], id="b_t"), row=2),
                  chart_block("b_t")]
        _, result, snap = self.build([{"name": "S", "blocks": blocks}])
        self.assertEqual(self.blocks(snap)["b_ch"]["cols"], 5)   # 3 planned, capped at 2 more
        self.assertEqual(self.moves(result, "widen")[0]["blockId"], "b_ch")

    def test_w11_strict_leaves_a_long_paragraph_in_its_columns(self):
        _, result, snap = self.build([{"name": "S", "blocks": [text_block()]}],
                                     fidelity="strict")
        self.assertEqual(self.blocks(snap)["b_x"]["cols"], 2)
        self.assertEqual(self.moves(result, "widen"), [])

    # ------------------------------------------------------ (b) total row

    def total_table(self, **kw):
        block = table(["Segment", "Accounts (K)", "Seats per account", "Price per seat",
                       "TAM ($M)"],
                      [["Micro", 4100, 6, 39, 959], ["Small", 1450, 18, 39, 1018]],
                      id="b_t", colFormats=["text", "count", "count", "currency", "currency"])
        block.update(kw)
        return block

    def test_w11_total_true_adds_a_sum_row_and_records_it(self):
        wb, result, snap = self.build([{"name": "S", "blocks": [self.total_table(total=True)]}])
        ws = wb["S"]
        self.assertEqual(ws["B6"].value, "Total")
        self.assertEqual(ws["C6"].value, "=SUM(C4:C5)")
        self.assertTrue(ws["B6"].font.b and ws["C6"].font.b)
        self.assertEqual(self.blocks(snap)["b_t"]["rows"], 5)
        move = self.moves(result, "total")[0]
        self.assertEqual((move["from"]["rows"], move["to"]["rows"]), (4, 5))

    def test_w11_a_total_row_leaves_the_columns_that_do_not_add_up_blank(self):
        """Accounts and TAM add up; seats per account and price per seat are rates."""
        wb, _, _ = self.build([{"name": "S", "blocks": [self.total_table(total=True)]}])
        ws = wb["S"]
        self.assertEqual([ws.cell(6, c).value for c in range(2, 7)],
                         ["Total", "=SUM(C4:C5)", None, None, "=SUM(F4:F5)"])

    def test_w11_a_percent_or_year_column_is_never_summed(self):
        block = table(["Year", "Penetration", "Revenue ($M)"],
                      [[2026, 0.01, 12], [2027, 0.02, 24]], id="b_t",
                      colFormats=["year", "percent", "currency"], total=True)
        wb, _, _ = self.build([{"name": "S", "blocks": [block]}])
        self.assertEqual([wb["S"].cell(6, c).value for c in range(2, 5)],
                         ["Total", None, "=SUM(D4:D5)"])

    def test_w11_the_total_row_carries_the_column_number_formats(self):
        wb, _, snap = self.build([{"name": "S", "blocks": [self.total_table(total=True)]}])
        cells = dict((c["a"], c) for c in snap["S"]["cells"])
        self.assertEqual(cells["F6"]["k"], "total")
        self.assertEqual(wb["S"]["F6"].number_format, build_xlsx.FORMATS["currency"])

    def test_w11_an_approved_total_row_is_never_regenerated(self):
        block = self.total_table(total=True, totalRow=["All", {"f": "=SUM(C4:C5)"},
                                                       None, None, None])
        wb, result, snap = self.build([{"name": "S", "blocks": [block]}])
        self.assertEqual(wb["S"]["B6"].value, "All")
        self.assertEqual(self.blocks(snap)["b_t"]["rows"], 5)
        self.assertEqual(self.moves(result, "total"), [])

    def test_w11_strict_adds_no_total_row(self):
        wb, result, snap = self.build([{"name": "S", "blocks": [self.total_table(total=True)]}],
                                      fidelity="strict")
        self.assertIsNone(wb["S"]["B6"].value)
        self.assertEqual(self.blocks(snap)["b_t"]["rows"], 4)
        self.assertEqual(self.moves(result, "total"), [])

    def test_w11_a_total_row_pushes_the_block_below_down_two_rows(self):
        blocks = [self.total_table(total=True, row=2),
                  dict(table(["Line", "2026"], [["A", 1]], id="b_u"), row=6)]
        _, result, snap = self.build([{"name": "S", "blocks": blocks}])
        self.assertEqual(self.blocks(snap)["b_u"]["row"], 8)     # the total took row 6
        shift = self.moves(result, "shift")[0]
        self.assertEqual(shift["to"]["row"] - shift["from"]["row"], 2)

    def test_w11_a_total_row_with_no_room_is_given_back_with_a_warning(self):
        blocks = [self.total_table(total=True, row=2),
                  dict(table(["Line", "2026"], [["A", 1]], id="b_u"), row=6),
                  dict(table(["Line", "2026"], [["A", 1]], id="b_v"), row=9)]
        wb, result, snap = self.build([{"name": "S", "blocks": blocks}])
        self.assertEqual(self.blocks(snap)["b_t"]["rows"], 4)
        written = [c.value for row in wb["S"].iter_rows() for c in row]
        self.assertNotIn("Total", written)
        self.assertTrue(any("no free row under the block" in w for w in result["warnings"]),
                        result["warnings"])

    def test_w11_a_table_with_nothing_to_add_up_warns(self):
        block = table(["Ref", "Source"], [["R1", "Census"]], id="b_t",
                      colFormats=["text", "text"], total=True)
        _, result, _ = self.build([{"name": "S", "blocks": [block]}])
        self.assertTrue(any("no column of the table adds up" in w for w in result["warnings"]),
                        result["warnings"])

    # ------------------------------------------------------ (c) note column

    def band(self, **extra):
        blocks = [
            dict(table(["L", "a"], [["y", 1]], id="b_n", note="beside me"), row=2, col=2),
            dict(table(["L", "a", "b", "c", "d"], [["x", 1, 2, 3, 4]], id="b_w"),
                 row=2, col=6),
        ]
        return self.build([{"name": "S", "blocks": blocks}], **extra)

    def test_w11_balanced_moves_a_note_to_the_band_edge_and_records_it(self):
        wb, result, snap = self.band()
        self.assertEqual(self.blocks(snap)["b_n"]["noteCol"], 11)
        self.assertEqual(wb["S"]["K2"].value, "beside me")
        self.assertEqual(self.moves(result, "note")[0]["blockId"], "b_n")

    def test_w11_strict_keeps_a_note_beside_its_own_block(self):
        wb, result, snap = self.band(fidelity="strict", name="band_strict")
        self.assertEqual(wb["S"]["D2"].value, "beside me")
        self.assertNotIn("noteCol", self.blocks(snap)["b_n"])
        self.assertEqual(self.moves(result, "note"), [])

    # ------------------------------------------------------- how it is recorded

    def test_w11_a_width_move_keeps_its_own_shape(self):
        """The W13 records stay {blockId, from: cols, to: cols}; only fidelity moves are rects."""
        _, result, snap = self.build([{"name": "S", "blocks": [
            dict(table(["L", "a", "b", "c"], [["x", 1, 2, 3]], id="b_w"), row=2),
            dict(table(["L", "a"], [["y", 1]], id="b_n"), row=6)]}])
        fit = [m for m in result["adjusted"] if m["reason"] == "snap"][0]
        self.assertEqual(fit, {"blockId": "b_n", "from": 2, "to": 4, "reason": "snap",
                               "sheet": "S"})
        self.assertEqual(self.blocks(snap)["b_n"]["adjusted"], "snap")

    def test_w11_a_block_carries_its_own_move_in_the_snapshot(self):
        _, result, snap = self.build([{"name": "S", "blocks": [text_block()]}])
        adjusted = self.blocks(snap)["b_x"]["adjusted"]
        self.assertEqual(adjusted["reason"], "widen")
        self.assertEqual(adjusted["from"]["cols"], 2)
        self.assertEqual(adjusted["to"]["cols"], 4)

    def test_w11_two_moves_on_one_block_read_as_one_reason(self):
        blocks = [dict(table(["L", "a"], [["y", 1]], id="b_a"), row=2),
                  self.total_table(total=True, row=5, id="b_t")]
        _, _, snap = self.build([{"name": "S", "blocks": blocks}])
        adjusted = self.blocks(snap)["b_t"]["adjusted"]
        self.assertEqual(adjusted["reason"], "total+shift")
        self.assertEqual((adjusted["from"]["row"], adjusted["to"]["row"]), (5, 6))
        self.assertEqual((adjusted["from"]["rows"], adjusted["to"]["rows"]), (4, 5))

    def test_w11_every_move_names_its_sheet(self):
        _, result, _ = self.build([
            {"name": "S", "blocks": [text_block()]},
            {"name": "T", "blocks": [self.total_table(total=True)]},
        ])
        self.assertEqual(sorted(set(m["sheet"] for m in result["adjusted"])), ["S", "T"])

    def test_w11_a_sheet_that_moved_nothing_carries_no_adjusted_list(self):
        _, result, snap = self.build(self.stack([2]))
        self.assertNotIn("adjusted", result)
        self.assertNotIn("adjusted", snap["S"])

    def test_w11_the_shipped_sample_exercises_the_balanced_moves(self):
        """The sample's plan leaves its blocks a blank row apart, so the moves it shows are
        the two the plan cannot make itself: the ramp's total row and the long paragraphs."""
        spec_path = os.path.join(HERE, "sample_spec.json")
        with open(spec_path) as handle:
            spec = json.load(handle)
        spec["outputPath"] = os.path.join(self.tmp, "sample_fidelity.xlsx")
        result = build_xlsx.build(spec)
        self.assertTrue(result["ok"], result["errors"])
        self.assertEqual(result["warnings"], [])
        reasons = set(m["reason"] for m in result["adjusted"])
        self.assertTrue({"widen", "total"} <= reasons, sorted(reasons))
        self.assertNotIn("shift", reasons)
        strict = dict(spec, fidelity="strict",
                      outputPath=os.path.join(self.tmp, "sample_fidelity_strict.xlsx"))
        result = build_xlsx.build(strict)
        self.assertTrue(result["ok"], result["errors"])
        self.assertEqual(set(m["reason"] for m in result.get("adjusted", [])), {"fit"})


# ----------------------------------------------------------- N9 comments export


THREAD = {"blockId": "b_t", "entries": [
    {"by": "Jackson", "text": "Split this by segment.", "status": "resolved"},
    {"by": "Claude", "text": "Split into SMB and Enterprise.", "status": "resolved"},
]}


class CommentTests(unittest.TestCase):
    """N9: the block's resolved threads, written as one Excel comment on its title cell."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="workbook_comment_test_")

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def build(self, threads=None, comments=None, blocks=None, name=None):
        path = os.path.join(self.tmp, (name or self.id().rsplit(".", 1)[-1]) + ".xlsx")
        extra = {}
        if threads is not None:
            extra["threads"] = threads
        if comments is not None:
            extra["comments"] = comments
        sheets = [{"name": "S", "blocks": blocks or [table(["Line", "2026"], [["A", 1]])]}]
        result = build_xlsx.build(quality_spec(path, sheets, "house", **extra))
        self.assertTrue(result["ok"], result["errors"])
        with open(result["snapshot"]) as handle:
            snap = json.load(handle)
        return (load_workbook(path), result,
                dict((s["name"], s) for s in snap["sheets"]))

    def test_n9_comments_are_off_by_default(self):
        wb, result, snap = self.build(threads=[THREAD])
        self.assertIsNone(wb["S"]["B2"].comment)
        self.assertEqual(result["warnings"], [])
        self.assertNotIn("hasComment", dict((c["a"], c) for c in snap["S"]["cells"])["B2"])

    def test_n9_a_thread_becomes_one_comment_on_the_block_title(self):
        wb, result, snap = self.build(threads=[THREAD], comments=True)
        comment = wb["S"]["B2"].comment
        self.assertIsNotNone(comment)
        self.assertEqual(comment.author, "Blueprint")
        self.assertEqual(comment.text,
                         "Jackson: Split this by segment.\n"
                         "Claude: Split into SMB and Enterprise.")
        self.assertTrue(dict((c["a"], c) for c in snap["S"]["cells"])["B2"]["hasComment"])
        self.assertEqual(result["warnings"], [])

    def test_n9_the_newest_entry_comes_last(self):
        wb, _, _ = self.build(threads=[THREAD], comments=True)
        lines = wb["S"]["B2"].comment.text.split("\n")
        self.assertEqual(lines[-1], "Claude: Split into SMB and Enterprise.")

    def test_n9_several_threads_on_one_block_share_one_comment(self):
        second = {"blockId": "b_t", "entries": [{"by": "Jackson", "text": "Add a total row."}]}
        wb, _, _ = self.build(threads=[THREAD, second], comments=True)
        text = wb["S"]["B2"].comment.text
        self.assertIn("Split this by segment.", text)
        self.assertEqual(text.split("\n\n")[-1], "Jackson: Add a total row.")

    def test_n9_an_entry_that_was_declined_says_so(self):
        thread = {"blockId": "b_t", "entries": [
            {"by": "Jackson", "text": "Make it monthly.", "status": "declined"}]}
        wb, _, _ = self.build(threads=[thread], comments=True)
        self.assertEqual(wb["S"]["B2"].comment.text, "Jackson: Make it monthly. [declined]")

    def test_n9_an_entry_with_no_author_is_written_plain(self):
        thread = {"blockId": "b_t", "entries": [{"text": "Carried over from the last model."}]}
        wb, _, _ = self.build(threads=[thread], comments=True)
        self.assertEqual(wb["S"]["B2"].comment.text, "Carried over from the last model.")

    def test_n9_only_the_title_cell_carries_a_comment(self):
        wb, _, _ = self.build(threads=[THREAD], comments=True)
        ws = wb["S"]
        with_comments = [c.coordinate for row in ws.iter_rows() for c in row if c.comment]
        self.assertEqual(with_comments, ["B2"])

    def test_n9_a_thread_for_a_block_that_is_not_here_warns(self):
        thread = {"blockId": "b_missing", "entries": [{"by": "Jackson", "text": "Hello."}]}
        _, result, _ = self.build(threads=[thread], comments=True)
        self.assertTrue(any("no block 'b_missing'" in w for w in result["warnings"]),
                        result["warnings"])

    def test_n9_a_thread_with_no_block_id_warns(self):
        _, result, _ = self.build(threads=[{"entries": [{"text": "Hello."}]}], comments=True)
        self.assertTrue(any("names no blockId" in w for w in result["warnings"]),
                        result["warnings"])

    def test_n9_threads_that_are_not_a_list_warn(self):
        _, result, _ = self.build(threads={"blockId": "b_t"}, comments=True)
        self.assertTrue(any("threads must be a list" in w for w in result["warnings"]),
                        result["warnings"])

    def test_n9_an_empty_thread_writes_nothing(self):
        wb, result, _ = self.build(threads=[{"blockId": "b_t", "entries": []}], comments=True)
        self.assertIsNone(wb["S"]["B2"].comment)
        self.assertEqual(result["warnings"], [])

    def test_n9_a_comment_survives_the_file_round_trip(self):
        path = os.path.join(self.tmp, "roundtrip.xlsx")
        sheets = [{"name": "S", "blocks": [table(["Line", "2026"], [["A", 1]])]}]
        result = build_xlsx.build(quality_spec(path, sheets, "house", comments=True,
                                               threads=[THREAD]))
        self.assertTrue(result["ok"], result["errors"])
        comment = load_workbook(path)["S"]["B2"].comment
        self.assertIn("Split into SMB and Enterprise.", comment.text)
        self.assertEqual(comment.author, "Blueprint")


if __name__ == "__main__":
    unittest.main(verbosity=2)

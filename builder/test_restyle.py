#!/usr/bin/env python3
"""Tests for restyle.py. Run: python3 builder/test_restyle.py

Restyles sample_plain.xlsx to house and checks fonts, borders, alignment, colours,
gridlines and the title merge on known cells against the house rules, with every
value and formula unchanged. Restyles every corpus workbook (make_corpus.py) both
ways with values unchanged. Round trips sample_house.xlsx through plain and back.
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

from openpyxl import load_workbook  # noqa: E402

import build_xlsx  # noqa: E402
import import_scan  # noqa: E402
import make_corpus  # noqa: E402
import restyle  # noqa: E402

RESTYLE = os.path.join(HERE, "restyle.py")
SAMPLE_SPEC = os.path.join(HERE, "sample_spec.json")
SAMPLE_PLAIN = os.path.join(HERE, "sample_plain.xlsx")
SAMPLE_HOUSE = os.path.join(HERE, "sample_house.xlsx")
NAVY, GREY, BLUE, BLACK = build_xlsx.NAVY, build_xlsx.GREY, build_xlsx.BLUE, build_xlsx.BLACK


def values(path):
    """{(sheet, coordinate): value} for every non-empty cell, formulas as strings."""
    wb = load_workbook(path)
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    out[(ws.title, c.coordinate)] = c.value
    return out


def rgb(color):
    return color.rgb[-6:].upper() if color is not None and isinstance(color.rgb, str) else None


def side(border, name):
    s = getattr(border, name)
    return s.style if s is not None else None


class RestyleTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="restyle-")
        for sample, mode in ((SAMPLE_PLAIN, "plain"), (SAMPLE_HOUSE, "house")):
            if not os.path.exists(sample):
                with open(SAMPLE_SPEC) as fh:
                    spec = json.load(fh)
                spec["outputPath"], spec["formatting"] = sample, mode
                build_xlsx.build(spec)
        cls.plain_scan = import_scan.scan(SAMPLE_PLAIN)
        cls.house_path = os.path.join(cls.tmp, "plain_to_house.xlsx")
        spec = {"inputPath": SAMPLE_PLAIN, "outputPath": cls.house_path, "formatting": "house", "scan": cls.plain_scan}
        spec_path = os.path.join(cls.tmp, "spec.json")
        with open(spec_path, "w") as fh:
            json.dump(spec, fh)
        proc = subprocess.run([sys.executable, RESTYLE, spec_path], capture_output=True, text=True)
        lines = [ln for ln in proc.stdout.splitlines() if ln.strip()]
        cls.result = json.loads(lines[-1]) if lines else {"ok": False, "errors": [proc.stderr]}
        cls.rc = proc.returncode

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    # ------------------------------------------------------ plain -> house

    def test_cli_result(self):
        self.assertEqual(self.rc, 0, self.result)
        self.assertTrue(self.result["ok"], self.result)
        self.assertEqual(self.result["path"], self.house_path)
        names = [s["name"] for s in self.result["sheets"]]
        self.assertEqual(names, ["Assumptions", "TAM Build", "Revenue Ramp", "Sources & Appendix"])
        for s in self.result["sheets"]:
            self.assertGreater(s["blocks"], 0, s)
            self.assertGreater(s["cells"], 0, s)
        self.assertTrue(os.path.exists(self.house_path))

    def test_values_and_formulas_unchanged(self):
        before, after = values(SAMPLE_PLAIN), values(self.house_path)
        self.assertEqual(before, after)
        formulas = [v for v in before.values() if isinstance(v, str) and v.startswith("=")]
        self.assertGreater(len(formulas), 10)
        self.assertEqual(after[("TAM Build", "C11")], "=SUM(C7:C10)")

    def test_gridlines_off_on_every_restyled_sheet(self):
        wb = load_workbook(self.house_path)
        for name in ("Assumptions", "TAM Build", "Revenue Ramp", "Sources & Appendix"):
            self.assertFalse(wb[name].sheet_view.showGridLines, name)
        self.assertTrue(load_workbook(SAMPLE_PLAIN)["Assumptions"].sheet_view.showGridLines)

    def test_house_title_rule_and_merge(self):
        ws = load_workbook(self.house_path)["Assumptions"]
        title = ws["B5"]
        self.assertEqual(title.value, "Market frame")
        self.assertEqual((title.font.name, title.font.sz, title.font.b), ("Arial", 10, True))
        self.assertEqual(title.alignment.horizontal, "center")
        self.assertEqual(side(title.border, "top"), "thin")
        self.assertEqual(side(title.border, "bottom"), "medium")
        self.assertEqual(rgb(title.border.bottom.color), NAVY)
        merges = set(str(m) for m in ws.merged_cells.ranges)
        self.assertIn("B5:C5", merges)   # title spans the block's own label and value columns
        self.assertIn("B2:G2", merges)   # text block title, a column wider than the plan (W11)
        self.assertIn("D6:F6", merges)   # existing note merge kept
        self.assertNotIn("B5:E5", merges)   # and never the note column

    def test_house_pairs_labels_values_notes(self):
        ws = load_workbook(self.house_path)["Assumptions"]
        label, value, note, block_note = ws["B6"], ws["C6"], ws["D6"], ws["D5"]
        self.assertEqual(label.font.name, "Arial")
        self.assertEqual(side(label.border, "right"), "dashed")
        self.assertEqual(side(label.border, "left"), "thin")
        self.assertEqual(label.alignment.horizontal, "left")
        self.assertEqual(rgb(value.font.color), BLACK)   # "2026", stored as text, is a year
        self.assertEqual(rgb(ws["C7"].font.color), BLUE)  # 33.2, a literal a reader can change
        self.assertEqual(value.alignment.horizontal, "center")
        self.assertTrue(all(side(value.border, s) == "thin" for s in ("left", "right", "top", "bottom")))
        for cell in (note, block_note):
            self.assertTrue(cell.font.i)
            self.assertEqual(rgb(cell.font.color), GREY)
            self.assertFalse(any(side(cell.border, s) for s in ("left", "right", "top", "bottom")))
        self.assertIsNone(value.fill.fill_type)

    def test_house_table_header_body_total(self):
        ws = load_workbook(self.house_path)["TAM Build"]
        header, first_header, body, label, total, total_label = ws["C6"], ws["B6"], ws["C7"], ws["B7"], ws["C11"], ws["B11"]
        self.assertTrue(header.font.b and first_header.font.b)
        self.assertEqual(header.alignment.horizontal, "center")
        self.assertEqual(side(first_header.border, "right"), "dashed")
        self.assertEqual(side(header.border, "right"), "thin")
        self.assertEqual(rgb(body.font.color), BLUE)    # a literal in any block is an input (W2)
        self.assertFalse(body.font.b)
        self.assertEqual(side(label.border, "right"), "dashed")
        self.assertEqual(label.alignment.horizontal, "left")
        self.assertTrue(total.font.b and total_label.font.b)
        self.assertEqual(rgb(total.font.color), BLACK)  # formula
        self.assertEqual(total.number_format, ws["C7"].number_format)  # existing formats untouched

    def test_house_matches_builder_output(self):
        """Every cell a house build styles looks the same after restyling the same plain build.

        Both sides are built from sample_spec.json here, not read from the committed sample
        workbooks: those are other suites' regression anchors and predate the current styling
        rules, and what this test is for is restyle agreeing with the builder as it is now.
        """
        fresh_plain = os.path.join(self.tmp, "fresh_plain.xlsx")
        fresh_house = os.path.join(self.tmp, "fresh_house.xlsx")
        restyled = os.path.join(self.tmp, "fresh_plain_to_house.xlsx")
        with open(SAMPLE_SPEC) as fh:
            spec = json.load(fh)
        for path, mode in ((fresh_plain, "plain"), (fresh_house, "house")):
            local = dict(spec, outputPath=path, formatting=mode)
            self.assertTrue(build_xlsx.build(local)["ok"], mode)
        res = restyle.restyle({"inputPath": fresh_plain, "outputPath": restyled,
                               "formatting": "house", "scan": import_scan.scan(fresh_plain)})
        self.assertTrue(res["ok"], res)
        ref, out = load_workbook(fresh_house), load_workbook(restyled)
        checked = 0
        for name in ("Assumptions", "TAM Build", "Revenue Ramp", "Sources & Appendix"):
            for row in ref[name].iter_rows():
                for rc in row:
                    if rc.value is None:
                        continue
                    oc = out[name][rc.coordinate]
                    where = "%s!%s" % (name, rc.coordinate)
                    self.assertEqual((oc.font.name, oc.font.sz, bool(oc.font.b), bool(oc.font.i), rgb(oc.font.color)),
                                     (rc.font.name, rc.font.sz, bool(rc.font.b), bool(rc.font.i), rgb(rc.font.color)), where)
                    self.assertEqual(tuple(side(oc.border, s) for s in ("left", "right", "top", "bottom")),
                                     tuple(side(rc.border, s) for s in ("left", "right", "top", "bottom")), where)
                    self.assertEqual((oc.alignment.horizontal, oc.alignment.vertical, bool(oc.alignment.wrap_text)),
                                     (rc.alignment.horizontal, rc.alignment.vertical, bool(rc.alignment.wrap_text)), where)
                    self.assertEqual(oc.number_format, rc.number_format, where)
                    checked += 1
        self.assertGreater(checked, 150)

    def test_charts_survive(self):
        out = load_workbook(self.house_path)
        self.assertEqual([len(ws._charts) for ws in out.worksheets], [0, 0, 1, 1, 0])

    def test_cells_outside_blocks_untouched(self):
        src = os.path.join(self.tmp, "stray.xlsx")
        shutil.copy(SAMPLE_PLAIN, src)
        wb = load_workbook(src)
        ws = wb["Assumptions"]
        ws["J40"] = "stray"
        ws["J40"].font = build_xlsx.Font(name="Courier", size=14, bold=True)
        wb.save(src)
        dst = os.path.join(self.tmp, "stray_house.xlsx")
        res = restyle.restyle({"inputPath": src, "outputPath": dst, "formatting": "house", "scan": import_scan.scan(src)})
        self.assertTrue(res["ok"], res)
        cell = load_workbook(dst)["Assumptions"]["J40"]
        self.assertEqual((cell.value, cell.font.name, cell.font.sz, cell.font.b), ("stray", "Courier", 14, True))

    def test_sheets_filter(self):
        dst = os.path.join(self.tmp, "one_sheet.xlsx")
        res = restyle.restyle({"inputPath": SAMPLE_PLAIN, "outputPath": dst, "formatting": "house",
                               "scan": self.plain_scan, "sheets": ["TAM Build", "Nope"]})
        self.assertTrue(res["ok"], res)
        self.assertEqual([s["name"] for s in res["sheets"]], ["TAM Build"])
        self.assertTrue(any("Nope" in w for w in res["warnings"]))
        wb = load_workbook(dst)
        self.assertFalse(wb["TAM Build"].sheet_view.showGridLines)
        self.assertTrue(wb["Assumptions"].sheet_view.showGridLines)
        self.assertEqual(wb["Assumptions"]["B5"].font.name, "Calibri")

    def test_never_overwrites_input(self):
        res = restyle.restyle({"inputPath": SAMPLE_PLAIN, "outputPath": SAMPLE_PLAIN, "formatting": "house",
                               "scan": self.plain_scan})
        self.assertFalse(res["ok"])
        self.assertTrue(any("never overwritten" in e for e in res["errors"]))

    def test_scan_without_geometry_is_rescanned(self):
        scan = json.loads(json.dumps(self.plain_scan))
        for tab in scan["tabs"]:
            for b in tab["blocks"]:
                b.pop("geometry", None)
        dst = os.path.join(self.tmp, "rescanned.xlsx")
        res = restyle.restyle({"inputPath": SAMPLE_PLAIN, "outputPath": dst, "formatting": "house", "scan": scan})
        self.assertTrue(res["ok"], res)
        self.assertTrue(any("re-scanned" in w for w in res["warnings"]))
        self.assertEqual(load_workbook(dst)["Assumptions"]["B5"].font.name, "Arial")

    # --------------------------------------------------------- corpus

    def test_corpus_values_unchanged(self):
        corpus = os.path.join(self.tmp, "corpus")
        make_corpus.make(corpus)
        seen = 0
        for name in sorted(os.listdir(corpus)):
            if not name.endswith(".xlsx"):
                continue
            src = os.path.join(corpus, name)
            scan = import_scan.scan(src)
            for mode in ("house", "plain"):
                dst = os.path.join(self.tmp, "%s_%s" % (mode, name))
                res = restyle.restyle({"inputPath": src, "outputPath": dst, "formatting": mode, "scan": scan})
                self.assertTrue(res["ok"], (name, mode, res))
                self.assertEqual(values(src), values(dst), (name, mode))
                self.assertGreater(sum(s["cells"] for s in res["sheets"]), 0, (name, mode))
                seen += 1
        self.assertEqual(seen, 2 * (len(make_corpus.FILES) + len(make_corpus.EXTRA_FILES)))

    # ------------------------------------------------------ round trip

    def test_house_to_plain_and_back_keeps_values(self):
        plain = os.path.join(self.tmp, "house_to_plain.xlsx")
        res = restyle.restyle({"inputPath": SAMPLE_HOUSE, "outputPath": plain, "formatting": "plain",
                               "scan": import_scan.scan(SAMPLE_HOUSE)})
        self.assertTrue(res["ok"], res)
        self.assertEqual([s["name"] for s in res["sheets"]], ["Assumptions", "TAM Build", "Revenue Ramp", "Sources & Appendix"])
        self.assertTrue(any("Read Me" in w and "left as is" in w for w in res["warnings"]))
        wb = load_workbook(plain)
        ws = wb["Assumptions"]
        self.assertTrue(ws.sheet_view.showGridLines)
        self.assertEqual((ws["B5"].font.name, ws["B5"].font.sz, ws["B5"].font.b), ("Calibri", 11, True))
        self.assertIsNone(side(ws["B5"].border, "bottom"))
        self.assertNotIn("B5:E5", set(str(m) for m in ws.merged_cells.ranges))  # plain has no title merge
        self.assertEqual(rgb(ws["C6"].font.color), BLACK)  # plain never colours inputs blue
        self.assertEqual(wb["Read Me"]["A1"].font.sz, 16)  # cover page untouched
        self.assertEqual(values(SAMPLE_HOUSE), values(plain))

        back = os.path.join(self.tmp, "plain_to_house_again.xlsx")
        res = restyle.restyle({"inputPath": plain, "outputPath": back, "formatting": "house",
                               "scan": import_scan.scan(plain)})
        self.assertTrue(res["ok"], res)
        self.assertEqual(values(SAMPLE_HOUSE), values(back))
        ws = load_workbook(back)["Assumptions"]
        self.assertIn("B5:C5", set(str(m) for m in ws.merged_cells.ranges))
        self.assertEqual(rgb(ws["C6"].font.color), BLACK)   # "2026" is a year, not an input
        self.assertEqual(rgb(ws["C7"].font.color), BLUE)

    def test_table_row_notes_restyle_like_the_builder(self):
        """A plain table with rowNotes restyled to house: grey italic unbordered notes, title merge stops before them."""
        src = os.path.join(self.tmp, "rownotes_plain.xlsx")
        spec = {"outputPath": src, "formatting": "plain", "project": {"name": "Notes"},
                "sheets": [{"name": "Tbl", "blocks": [{
                    "id": "b_n", "type": "table", "title": "Ramp", "row": 2, "col": 2, "cols": 3,
                    "columns": ["Segment", "Year 1", "Year 2"], "colFormats": ["text", "currency", "currency"],
                    "data": [["SMB", 1, 2], ["Enterprise", 3, 4], ["Mid", 5, 6]],
                    "rowNotes": ["From CRM", None, "Estimate"],
                    "totalRow": ["Total", {"f": "=SUM({{b_n!col:1}})"}, {"f": "=SUM({{b_n!col:2}})"}],
                    "totalNote": "Sum check"}]}]}
        self.assertTrue(build_xlsx.build(spec)["ok"])
        scan = import_scan.scan(src)
        block = scan["tabs"][0]["blocks"][0]
        self.assertEqual(block["type"], "table")
        self.assertIn("rowNotes", block)
        self.assertEqual(block["geometry"]["noteCells"], [[4, 5], [6, 5], [7, 5]])
        dst = os.path.join(self.tmp, "rownotes_house.xlsx")
        res = restyle.restyle({"inputPath": src, "outputPath": dst, "formatting": "house", "scan": scan})
        self.assertTrue(res["ok"], res)
        self.assertEqual(values(src), values(dst))
        ws = load_workbook(dst)["Tbl"]
        for a in ("E4", "E6", "E7"):
            cell = ws[a]
            self.assertEqual(cell.font.name, "Arial", a)
            self.assertTrue(cell.font.i, a)
            self.assertEqual(rgb(cell.font.color), GREY, a)
            self.assertFalse(any(side(cell.border, s) for s in ("left", "right", "top", "bottom")), a)
        self.assertEqual(side(ws["D4"].border, "right"), "thin")
        self.assertTrue(ws["B7"].font.b and ws["C7"].font.b)
        merges = set(str(m) for m in ws.merged_cells.ranges)
        self.assertIn("B2:D2", merges)     # as the builder writes it: notes sit outside the title rule
        self.assertNotIn("B2:E2", merges)
        self.assertGreaterEqual(ws.column_dimensions["E"].width, len("Sum check"))
        # the same table built straight to house looks identical, cell for cell
        ref_path = os.path.join(self.tmp, "rownotes_ref.xlsx")
        spec["outputPath"], spec["formatting"] = ref_path, "house"
        self.assertTrue(build_xlsx.build(spec)["ok"])
        ref = load_workbook(ref_path)["Tbl"]
        for row in ref.iter_rows():
            for rc in row:
                if rc.value is None:
                    continue
                oc = ws[rc.coordinate]
                self.assertEqual((oc.font.name, bool(oc.font.b), bool(oc.font.i), rgb(oc.font.color)),
                                 (rc.font.name, bool(rc.font.b), bool(rc.font.i), rgb(rc.font.color)), rc.coordinate)
                self.assertEqual(tuple(side(oc.border, s) for s in ("left", "right", "top", "bottom")),
                                 tuple(side(rc.border, s) for s in ("left", "right", "top", "bottom")), rc.coordinate)

    # ------------------------------------------------- W14: links follow the style

    def link_source(self, name="links.xlsx"):
        """A one-sheet workbook holding a hyperlink cell and a bare cross-sheet pass-through."""
        src = os.path.join(self.tmp, name)
        spec = {"outputPath": src, "formatting": "plain", "project": {"name": "Links"},
                "sheets": [
                    {"name": "Src", "blocks": [{
                        "id": "b_in", "type": "inputs", "title": "Inputs", "row": 2, "col": 2,
                        "cols": 2, "rows": [{"label": "Growth", "value": 0.1}]}]},
                    {"name": "Tbl", "blocks": [{
                        "id": "b_l", "type": "table", "title": "Sources", "row": 2, "col": 2,
                        "cols": 3, "columns": ["Ref", "Link", "Pulled"],
                        "colFormats": ["text", "text", "percent"],
                        "data": [["R1", {"v": "SBA", "link": "https://advocacy.sba.gov/"},
                                  {"f": "={{b_in!Growth}}"}]]}]},
                ]}
        self.assertTrue(build_xlsx.build(spec)["ok"])
        return src

    def restyled_links(self, mode, name):
        src = self.link_source("links_%s.xlsx" % mode)
        dst = os.path.join(self.tmp, name)
        res = restyle.restyle({"inputPath": src, "outputPath": dst, "formatting": mode,
                               "scan": import_scan.scan(src)})
        self.assertTrue(res["ok"], res)
        self.assertEqual(values(src), values(dst))
        return load_workbook(dst)["Tbl"]

    def test_w14_restyle_to_house_draws_links_in_the_ink_with_an_underline(self):
        """The same rule as a build: no green, the underline is what marks a link."""
        ws = self.restyled_links("house", "links_to_house.xlsx")
        for a in ("C4", "D4"):                    # the hyperlink and the pass-through
            self.assertEqual(rgb(ws[a].font.color), BLACK, a)
            self.assertEqual(ws[a].font.u, "single", a)

    def test_w14_restyle_to_plain_uses_office_hyperlink_blue(self):
        ws = self.restyled_links("plain", "links_to_plain.xlsx")
        for a in ("C4", "D4"):
            self.assertEqual(rgb(ws[a].font.color), build_xlsx.LINK_BLUE, a)
            self.assertEqual(ws[a].font.u, "single", a)

    def test_w14_a_style_that_wants_no_underline_still_underlines_a_hyperlink(self):
        path = os.path.join(self.tmp, "flat.md")
        with open(path, "w") as fh:
            fh.write("---\nname: flat\nformat:\n  link_underline: false\n---\n# flat\n")
        src = self.link_source("links_flat.xlsx")
        dst = os.path.join(self.tmp, "links_flat_out.xlsx")
        res = restyle.restyle({"inputPath": src, "outputPath": dst, "formatting": "skill",
                               "skill": {"path": path}, "scan": import_scan.scan(src)})
        self.assertTrue(res["ok"], res)
        ws = load_workbook(dst)["Tbl"]
        self.assertEqual(ws["C4"].font.u, "single")
        self.assertIsNone(ws["D4"].font.u)

    def test_template_mode_uses_template_fonts(self):
        dst = os.path.join(self.tmp, "templated.xlsx")
        res = restyle.restyle({"inputPath": SAMPLE_PLAIN, "outputPath": dst, "formatting": "template",
                               "templatePath": SAMPLE_HOUSE, "scan": self.plain_scan})
        self.assertTrue(res["ok"], res)
        ws = load_workbook(dst)["TAM Build"]
        self.assertEqual(ws["C7"].font.name, "Arial")  # base font read from the template
        self.assertTrue(ws["B5"].font.b)
        self.assertTrue(ws.sheet_view.showGridLines)  # template builds like plain


if __name__ == "__main__":
    unittest.main(verbosity=1)

#!/usr/bin/env python3
"""Tests for import_pptx.py. Run: python3 builder/test_import_pptx.py

Generates the two fixture decks (fixtures-deck/make_fixtures.py) into a temp
dir, scans each through the CLI, and asserts the roles, the counts, the
sections, the 30 pt grid snapping and the confidence floor on ambiguous shapes.
The fixtures place every shape on exact grid coordinates, so snapping is
asserted without tolerance.
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.dont_write_bytecode = True  # keep a __pycache__ out of the fixtures dir
sys.path.insert(0, os.path.join(HERE, "fixtures-deck"))

import make_fixtures  # noqa: E402

SCANNER = os.path.join(HERE, "import_pptx.py")
TMP = {}


def setUpModule():
    TMP["dir"] = tempfile.mkdtemp(prefix="import_pptx_")
    make_fixtures.build_all(TMP["dir"])


def tearDownModule():
    shutil.rmtree(TMP.get("dir", ""), ignore_errors=True)


def run_scanner(src, out):
    proc = subprocess.run([sys.executable, SCANNER, src, out], capture_output=True, text=True)
    lines = [ln for ln in proc.stdout.splitlines() if ln.strip()]
    return proc.returncode, json.loads(lines[-1]) if lines else None, proc.stderr


def scan_fixture(name):
    src = os.path.join(TMP["dir"], name)
    out = os.path.join(TMP["dir"], name.replace(".pptx", ".json"))
    code, result, err = run_scanner(src, out)
    assert code == 0, "%s failed: %s %s" % (name, result, err)
    with open(out) as fh:
        return result, json.load(fh)


def blocks_of(plan):
    return [b for s in plan["slides"] for b in s["blocks"]]


def typed(slide, kind):
    return [b for b in slide["blocks"] if b["type"] == kind]


class ProtoDeckTest(unittest.TestCase):
    """Harbor shape: kicker, action title, three cards, furniture, a divider."""

    @classmethod
    def setUpClass(cls):
        cls.result, cls.plan = scan_fixture("harbor_deck.pptx")
        cls.slides = cls.plan["slides"]

    def test_cli_reports_counts(self):
        self.assertTrue(self.result["ok"])
        self.assertEqual(self.result["slides"], 6)
        self.assertEqual(self.result["dividers"], 3)
        self.assertEqual(self.result["blocks"], len(blocks_of(self.plan)))

    def test_slide_size_is_the_32x18_grid(self):
        size = self.plan["slideSize"]
        self.assertEqual((size["w"], size["h"]), (960.0, 540.0))
        self.assertEqual((size["cols"], size["rows"], size["cell"]), (32, 18, 30.0))

    def test_title_slide_headline_comes_from_the_title_placeholder(self):
        s = self.slides[0]
        self.assertEqual(s["headline"], "Referral playbook for the fall ask")
        self.assertEqual(s["headlineSignal"], "title placeholder")

    def test_headline_falls_back_to_the_largest_text_on_the_top_band(self):
        s = self.slides[1]
        self.assertEqual(s["headline"], "Referrals stall at the ask, not at the intro")
        self.assertEqual(s["headlineSignal"], "largest text on the top band")

    def test_slide_name_is_two_to_five_words_of_the_headline(self):
        for s in self.slides:
            self.assertTrue(2 <= len(s["name"].split()) <= 5, s["name"])
        self.assertEqual(self.slides[1]["name"], "Referrals stall at the ask")
        self.assertEqual(self.slides[0]["name"], "Referral playbook for the fall")

    def test_kicker_is_the_short_line_above_the_headline(self):
        self.assertEqual(self.slides[1]["kicker"], "Growth | Referrals")
        self.assertEqual(self.slides[3]["kicker"], "Growth | Fixes")
        self.assertIsNone(self.slides[0].get("kicker"))

    def test_headline_and_kicker_are_furniture_never_blocks(self):
        kinds = {b["type"] for b in blocks_of(self.plan)}
        self.assertNotIn("headline", kinds)
        self.assertNotIn("kicker", kinds)
        texts = {b["text"] for b in blocks_of(self.plan)}
        self.assertNotIn("Growth | Referrals", texts)
        self.assertNotIn("Referrals stall at the ask, not at the intro", texts)

    def test_three_round_rects_become_one_cards_block(self):
        cards = typed(self.slides[1], "cards")
        self.assertEqual(len(cards), 1)
        c = cards[0]
        self.assertEqual(c["count"], 3)
        self.assertEqual(c["content"], "3 cards: Open networks / Warm intros / Referral loop")
        self.assertGreaterEqual(c["confidence"], 0.8)
        self.assertTrue(any("rounded rectangles" in sig for sig in c["signals"]), c["signals"])
        self.assertTrue(any("row" in sig for sig in c["signals"]), c["signals"])

    def test_cards_span_the_card_box_on_the_grid(self):
        c = typed(self.slides[1], "cards")[0]
        # three 240 pt cards from x=60 to x=900, y=180 to y=480
        self.assertEqual((c["col"], c["row"], c["cols"], c["rows"]), (3, 7, 28, 10))
        self.assertEqual(c["shape"], {"x": 60.0, "y": 180.0, "w": 840.0, "h": 300.0})

    def test_page_number_and_logo_are_dropped(self):
        for s in self.slides:
            self.assertEqual([b for b in s["blocks"] if b["text"].strip().isdigit()], [])
            self.assertEqual(typed(s, "image"), [])

    def test_source_line_is_a_notes_block_on_the_bottom_row(self):
        notes = typed(self.slides[1], "notes")
        self.assertEqual(len(notes), 1)
        self.assertEqual(notes[0]["row"], 18)
        self.assertIn("Pilot data, 2026", notes[0]["content"])
        self.assertGreaterEqual(notes[0]["confidence"], 0.9)

    def test_divider_names_the_section_and_holds_no_blocks(self):
        d = self.slides[2]
        self.assertTrue(d["isDivider"])
        self.assertEqual(d["section"], "Where referrals stall")
        self.assertEqual(d["blocks"], [])
        self.assertIsNone(d.get("headline"))
        self.assertEqual(d["name"], "Where referrals stall")

    def test_a_full_bleed_rectangle_reads_as_a_divider_background(self):
        d = self.slides[4]
        self.assertTrue(d["isDivider"])
        self.assertEqual(d["section"], "How we fix it")
        self.assertEqual(d["blocks"], [])

    def test_a_layout_background_reads_as_a_divider_background(self):
        d = self.slides[5]
        self.assertTrue(d["isDivider"])
        self.assertEqual(d["section"], "What it costs")
        self.assertEqual(d["blocks"], [])

    def test_section_carries_to_the_following_slides(self):
        self.assertIsNone(self.slides[1].get("section"))
        self.assertEqual(self.slides[3]["section"], "Where referrals stall")
        self.assertFalse(self.slides[3]["isDivider"])

    def test_takeaway_bar_sits_on_the_bottom_band(self):
        t = typed(self.slides[3], "takeaway")
        self.assertEqual(len(t), 1)
        self.assertEqual((t[0]["col"], t[0]["row"], t[0]["cols"], t[0]["rows"]), (3, 17, 28, 1))
        self.assertIn("Fix the ask", t[0]["content"])
        self.assertTrue(any("bottom band" in sig for sig in t[0]["signals"]))

    def test_story_is_left_for_the_user_to_write(self):
        for s in self.slides:
            self.assertEqual(s["story"], "")

    def test_ids_follow_the_slide_index(self):
        for s in self.slides:
            for n, b in enumerate(s["blocks"], 1):
                self.assertEqual(b["id"], "sld_%d_%d" % (s["index"], n))

    def test_point_geometry_is_preserved(self):
        notes = typed(self.slides[1], "notes")[0]
        self.assertEqual(notes["shape"], {"x": 60.0, "y": 510.0, "w": 500.0, "h": 18.0})


class PlainDeckTest(unittest.TestCase):
    """A plain deck: bullets, chart, table, picture, stat, quote, placeholder, noise."""

    @classmethod
    def setUpClass(cls):
        cls.result, cls.plan = scan_fixture("plain_deck.pptx")
        cls.slides = cls.plan["slides"]

    def test_bullets_beat_prose(self):
        b = typed(self.slides[0], "bullets")
        self.assertEqual(len(b), 1)
        self.assertTrue(b[0]["content"].startswith("4 bullets:"))
        self.assertIn("Partners forget to ask", b[0]["content"])
        self.assertTrue(any("bullet character" in s for s in b[0]["signals"]), b[0]["signals"])
        self.assertTrue(any("levels" in s for s in b[0]["signals"]), b[0]["signals"])
        self.assertGreaterEqual(b[0]["confidence"], 0.8)

    def test_native_chart_carries_its_kind(self):
        c = typed(self.slides[1], "chart")
        self.assertEqual(len(c), 1)
        self.assertEqual(c[0]["chartKind"], "bar")
        self.assertEqual(c[0]["content"], "bar chart: 1 series over 4 categories")
        self.assertGreaterEqual(c[0]["confidence"], 0.9)

    def test_table_reports_its_shape_and_headers(self):
        t = typed(self.slides[2], "table")
        self.assertEqual(len(t), 1)
        self.assertEqual(t[0]["content"], "4x3 table: Channel / Donors / Cost")
        self.assertEqual((t[0]["col"], t[0]["row"], t[0]["cols"], t[0]["rows"]), (3, 7, 28, 9))

    def test_picture_that_is_not_a_logo_is_an_image_block(self):
        i = typed(self.slides[3], "image")
        self.assertEqual(len(i), 1)
        self.assertEqual((i[0]["col"], i[0]["row"]), (19, 7))
        self.assertTrue(any("picture" in s for s in i[0]["signals"]))

    def test_stat_is_a_big_number_with_a_small_label(self):
        s = typed(self.slides[3], "stat")
        self.assertEqual(len(s), 1)
        self.assertEqual(s[0]["content"], "stat 62% — of donors give again")
        self.assertEqual(s[0]["title"], "of donors give again")
        self.assertTrue(any("44pt" in sig for sig in s[0]["signals"]), s[0]["signals"])
        self.assertGreaterEqual(s[0]["confidence"], 0.8)

    def test_quote_is_centred_and_quoted(self):
        q = typed(self.slides[4], "quote")
        self.assertEqual(len(q), 1)
        self.assertIn("They ask, we deliver", q[0]["content"])
        self.assertTrue(any("centred" in s for s in q[0]["signals"]), q[0]["signals"])

    def test_dashed_outline_and_marker_make_a_placeholder(self):
        p = typed(self.slides[5], "placeholder")
        self.assertEqual(len(p), 1)
        self.assertEqual(p[0]["content"], "placeholder: Timeline of the pilot, quarter by quarter")
        self.assertTrue(any("dashed" in s for s in p[0]["signals"]), p[0]["signals"])
        self.assertTrue(any("PLACEHOLDER" in s for s in p[0]["signals"]), p[0]["signals"])

    def test_ambiguous_little_box_is_low_confidence(self):
        noise = [b for b in self.slides[5]["blocks"] if b["title"] == "Q3 push"]
        self.assertEqual(len(noise), 1)
        self.assertEqual(noise[0]["type"], "text")
        self.assertLess(noise[0]["confidence"], 0.6)
        self.assertTrue(any("small floating box" in s for s in noise[0]["signals"]), noise[0]["signals"])
        self.assertGreaterEqual(self.result["lowConfidence"], 1)

    def test_every_block_sits_inside_the_grid(self):
        for b in blocks_of(self.plan):
            self.assertGreaterEqual(b["col"], 1)
            self.assertGreaterEqual(b["row"], 1)
            self.assertGreaterEqual(b["cols"], 1)
            self.assertGreaterEqual(b["rows"], 1)
            self.assertLessEqual(b["col"] + b["cols"] - 1, 32)
            self.assertLessEqual(b["row"] + b["rows"] - 1, 18)

    def test_every_block_carries_confidence_and_signals(self):
        for b in blocks_of(self.plan):
            self.assertTrue(0.0 <= b["confidence"] <= 1.0, b)
            self.assertTrue(b["signals"], b)
            self.assertTrue(b["title"], b)
            self.assertTrue(b["content"], b)
            self.assertLessEqual(len(b["text"]), 200)
            self.assertEqual(b["text"], b["text"].strip())

    def test_headlines_are_found_on_every_content_slide(self):
        for s in self.slides:
            self.assertTrue(s.get("headline"), s["index"])
            self.assertFalse(s["isDivider"])


class CliTest(unittest.TestCase):
    def test_missing_file_fails_cleanly(self):
        tmp = tempfile.mkdtemp(prefix="import_pptx_cli_")
        try:
            code, result, _ = run_scanner(os.path.join(tmp, "nope.pptx"), os.path.join(tmp, "out.json"))
            self.assertEqual(code, 1)
            self.assertFalse(result["ok"])
            self.assertTrue(result["errors"])
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_usage(self):
        proc = subprocess.run([sys.executable, SCANNER], capture_output=True, text=True)
        self.assertEqual(proc.returncode, 1)
        self.assertFalse(json.loads(proc.stdout)["ok"])

    def test_not_a_deck_fails_cleanly(self):
        tmp = tempfile.mkdtemp(prefix="import_pptx_cli_")
        try:
            bad = os.path.join(tmp, "notes.pptx")
            with open(bad, "w") as fh:
                fh.write("this is not a presentation")
            code, result, _ = run_scanner(bad, os.path.join(tmp, "out.json"))
            self.assertEqual(code, 1)
            self.assertFalse(result["ok"])
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()

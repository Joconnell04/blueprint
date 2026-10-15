#!/usr/bin/env python3
"""Tests for diagram_layout.py. Run: python3 builder/test_diagram_layout.py"""

import math
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import build_pptx  # noqa: E402
import diagram_layout as dl  # noqa: E402

RECT = (60.0, 180.0, 840.0, 300.0)


def nodes(*labels):
    return [{"id": "n%d" % (i + 1), "label": label} for i, label in enumerate(labels)]


def chain(*ids):
    return [{"from": a, "to": b} for a, b in zip(ids, ids[1:])]


def by_id(plan):
    return dict((n["id"], n) for n in plan["nodes"])


def inside(node, rect, slack=0.51):
    x, y, w, h = rect
    return (node["x"] >= x - slack and node["y"] >= y - slack
            and node["x"] + node["w"] <= x + w + slack
            and node["y"] + node["h"] <= y + h + slack)


class MeasureTests(unittest.TestCase):
    """The layout must measure text exactly as the builder does."""

    def test_cpl_and_wrap_match_the_builder(self):
        for width in (60.0, 137.5, 280.0):
            for size in (8.0, 12.0, 17.0):
                self.assertEqual(dl.cpl(width, size, 0.5),
                                 build_pptx.cpl(width, size, 0.5))
        text = "Script the ask at the moment it is easy"
        self.assertEqual(dl.wrap_lines(text, 18), build_pptx.wrap_lines(text, 18))

    def test_options_merge_over_the_defaults(self):
        opts = dl.options({"gutterPt": 24, "bodyPt": 14, "nope": 3})
        self.assertEqual(opts["gutterPt"], 24.0)
        self.assertEqual(opts["bodyPt"], 14.0)
        self.assertEqual(opts["floorPt"], dl.DEFAULTS["floorPt"])
        self.assertNotIn("nope", opts)


class FlowTests(unittest.TestCase):

    def test_flow_lr_geometry_sits_inside_the_rect(self):
        plan = dl.plan({"kind": "flow", "nodes": nodes("Prepare", "Ask", "Connect"),
                        "edges": chain("n1", "n2", "n3")}, RECT)
        self.assertEqual(plan["kind"], "flow")
        self.assertEqual(plan["direction"], "LR")
        self.assertEqual(len(plan["nodes"]), 3)
        for node in plan["nodes"]:
            self.assertTrue(inside(node, RECT), node)
        self.assertFalse(dl.any_overlap(plan["nodes"]))
        xs = [n["x"] for n in plan["nodes"]]
        self.assertEqual(xs, sorted(xs))
        self.assertEqual(len(set(round(n["w"], 2) for n in plan["nodes"])), 1)

    def test_flow_tb_stacks_the_ranks_downwards(self):
        plan = dl.plan({"kind": "flow", "direction": "TB",
                        "nodes": nodes("Spot", "Script", "Hand off"),
                        "edges": chain("n1", "n2", "n3")}, RECT)
        ys = [n["y"] for n in plan["nodes"]]
        self.assertEqual(ys, sorted(ys))
        self.assertGreater(ys[1] - ys[0], plan["nodes"][0]["h"])
        for node in plan["nodes"]:
            self.assertTrue(inside(node, RECT), node)

    def test_rank_is_the_longest_path_from_the_sources(self):
        spec = {"kind": "flow", "nodes": nodes("A", "B", "C"),
                "edges": [{"from": "n1", "to": "n2"}, {"from": "n1", "to": "n3"},
                          {"from": "n3", "to": "n2"}]}
        plan = dl.plan(spec, RECT)
        ranks = dict((n["id"], n["rank"]) for n in plan["nodes"])
        self.assertEqual(ranks, {"n1": 0, "n3": 1, "n2": 2})

    def test_order_within_a_rank_follows_the_barycenter(self):
        spec = {"kind": "flow", "nodes": nodes("A", "B", "X", "Y"),
                "edges": [{"from": "n1", "to": "n4"}, {"from": "n2", "to": "n3"}]}
        plan = dl.plan(spec, RECT)
        found = by_id(plan)
        self.assertEqual(found["n4"]["order"], 0)
        self.assertEqual(found["n3"]["order"], 1)
        self.assertLess(found["n4"]["y"], found["n3"]["y"])

    def test_a_cycle_in_a_flow_does_not_hang(self):
        plan = dl.plan({"kind": "flow", "nodes": nodes("A", "B", "C"),
                        "edges": chain("n1", "n2", "n3") + [{"from": "n3", "to": "n1"}]}, RECT)
        self.assertEqual(len(plan["nodes"]), 3)
        self.assertEqual(len(plan["edges"]), 3)

    def test_connectors_leave_and_meet_the_boxes(self):
        plan = dl.plan({"kind": "flow", "nodes": nodes("Prepare", "Ask"),
                        "edges": [{"from": "n1", "to": "n2", "label": "prompts"}]}, RECT)
        edge = plan["edges"][0]
        found = by_id(plan)
        self.assertEqual(edge["kind"], "straight")
        self.assertAlmostEqual(edge["points"][0][0],
                               found["n1"]["x"] + found["n1"]["w"] + dl.DEFAULTS["edgeGapPt"], 1)
        self.assertAlmostEqual(edge["points"][-1][0],
                               found["n2"]["x"] - dl.DEFAULTS["edgeGapPt"], 1)
        self.assertEqual(edge["label"], "prompts")
        mid = edge["labelXY"]
        self.assertTrue(edge["points"][0][0] < mid[0] < edge["points"][-1][0])

    def test_an_offset_rank_gets_an_elbow(self):
        spec = {"kind": "flow", "nodes": nodes("A", "B", "C"),
                "edges": [{"from": "n1", "to": "n2"}, {"from": "n1", "to": "n3"}]}
        plan = dl.plan(spec, RECT)
        kinds = set(e["kind"] for e in plan["edges"])
        self.assertEqual(kinds, set(["elbow"]))
        for edge in plan["edges"]:
            self.assertEqual(len(edge["points"]), 4)


class CycleTests(unittest.TestCase):

    def test_ring_spacing_is_even_and_clockwise_from_the_top(self):
        plan = dl.plan({"kind": "cycle", "nodes": nodes("Plan", "Build", "Learn", "Adjust")},
                       RECT)
        centres = [(n["x"] + n["w"] / 2.0, n["y"] + n["h"] / 2.0) for n in plan["nodes"]]
        cx = RECT[0] + RECT[2] / 2.0
        cy = RECT[1] + RECT[3] / 2.0
        self.assertAlmostEqual(centres[0][0], cx, 1)
        self.assertLess(centres[0][1], cy)          # first node at the top
        self.assertGreater(centres[1][0], cx)       # second to the right: clockwise
        self.assertGreater(centres[2][1], cy)       # third at the bottom
        self.assertLess(centres[3][0], cx)          # fourth to the left

    def test_ring_nodes_are_equally_spaced_and_never_overlap(self):
        for count in range(3, 9):
            plan = dl.plan({"kind": "cycle",
                            "nodes": nodes(*["Step %d" % (i + 1) for i in range(count)])}, RECT)
            self.assertEqual(len(plan["nodes"]), count, plan["warnings"])
            self.assertFalse(dl.any_overlap(plan["nodes"]), count)
            for node in plan["nodes"]:
                self.assertTrue(inside(node, RECT), (count, node))
            ring = plan["ring"]
            self.assertLessEqual(ring["rx"], ring["ry"] * dl.DEFAULTS["ringRatio"] + 0.01)
            steps = []
            for i, node in enumerate(plan["nodes"]):
                dx = (node["x"] + node["w"] / 2.0 - ring["x"]) / ring["rx"]
                dy = (node["y"] + node["h"] / 2.0 - ring["y"]) / ring["ry"]
                self.assertAlmostEqual(dx * dx + dy * dy, 1.0, 2)   # on the ring
                steps.append(round(math.degrees(math.atan2(dy, dx)) % 360.0, 3))
            gaps = [round((steps[(i + 1) % count] - steps[i]) % 360.0, 2)
                    for i in range(count)]
            self.assertEqual(len(set(gaps)), 1, "uneven ring: %s" % gaps)
            self.assertAlmostEqual(gaps[0], 360.0 / count, 1)

    def test_a_ring_with_no_edges_gets_its_arrows(self):
        plan = dl.plan({"kind": "cycle", "nodes": nodes("Plan", "Build", "Learn")}, RECT)
        self.assertEqual(len(plan["edges"]), 3)
        pairs = set((e["from"], e["to"]) for e in plan["edges"])
        self.assertEqual(pairs, set([("n1", "n2"), ("n2", "n3"), ("n3", "n1")]))


class TimelineTests(unittest.TestCase):

    def test_boxes_alternate_around_the_spine(self):
        plan = dl.plan({"kind": "timeline",
                        "nodes": nodes("Q1 pilot", "Q2 rollout", "Q3 review", "Q4 renew")},
                       RECT)
        spine = plan["axis"][0]
        self.assertEqual(spine["kind"], "rule")
        self.assertAlmostEqual(spine["y"], RECT[1] + RECT[3] / 2.0, 1)
        sides = [n["side"] for n in plan["nodes"]]
        self.assertEqual(sides, ["above", "below", "above", "below"])
        for node in plan["nodes"]:
            self.assertTrue(inside(node, RECT), node)
            if node["side"] == "above":
                self.assertLessEqual(node["y"] + node["h"], spine["y"])
            else:
                self.assertGreaterEqual(node["y"], spine["y"])
        xs = [n["x"] for n in plan["nodes"]]
        self.assertEqual(xs, sorted(xs))
        self.assertFalse(dl.any_overlap(plan["nodes"]))


class TreeTests(unittest.TestCase):

    def test_tree_layers_top_down_from_the_root(self):
        spec = {"kind": "tree", "nodes": nodes("Brand", "Voice", "Visual", "Tokens"),
                "edges": [{"from": "n1", "to": "n2"}, {"from": "n1", "to": "n3"},
                          {"from": "n2", "to": "n4"}]}
        plan = dl.plan(spec, RECT)
        self.assertEqual(plan["direction"], "TB")
        self.assertEqual(plan["root"], "n1")
        found = by_id(plan)
        self.assertEqual(found["n1"]["rank"], 0)
        self.assertEqual(found["n2"]["rank"], 1)
        self.assertEqual(found["n3"]["rank"], 1)
        self.assertEqual(found["n4"]["rank"], 2)
        self.assertLess(found["n1"]["y"], found["n2"]["y"])
        self.assertLess(found["n2"]["y"], found["n4"]["y"])
        self.assertAlmostEqual(found["n2"]["y"], found["n3"]["y"], 1)
        self.assertFalse(dl.any_overlap(plan["nodes"]))


class MatrixTests(unittest.TestCase):

    def test_axis_labels_then_four_square_cells(self):
        spec = {"kind": "matrix",
                "nodes": nodes("Effort", "Value", "Quick wins", "Big bets",
                               "Fill ins", "Money pits")}
        plan = dl.plan(spec, RECT)
        self.assertEqual(len(plan["nodes"]), 6)
        axes = [n for n in plan["nodes"] if n["shape"] == "text"]
        cells = [n for n in plan["nodes"] if n["shape"] == "roundRect"]
        self.assertEqual([n["axis"] for n in axes], ["x", "y"])
        self.assertEqual([n["cell"] for n in cells], ["tl", "tr", "bl", "br"])
        self.assertAlmostEqual(cells[0]["w"], cells[0]["h"], 1)
        self.assertEqual(len(set(round(c["w"], 1) for c in cells)), 1)
        self.assertAlmostEqual(cells[0]["y"], cells[1]["y"], 1)
        self.assertAlmostEqual(cells[0]["x"], cells[2]["x"], 1)
        self.assertGreater(cells[2]["y"], cells[0]["y"])
        self.assertGreater(cells[1]["x"], cells[0]["x"])
        for node in plan["nodes"]:
            self.assertTrue(inside(node, RECT), node)
        self.assertFalse(dl.any_overlap(plan["nodes"]))

    def test_matrix_drops_edges_and_extra_nodes_with_a_warning(self):
        spec = {"kind": "matrix",
                "nodes": nodes("X", "Y", "A", "B", "C", "D", "E"),
                "edges": [{"from": "n3", "to": "n4"}]}
        plan = dl.plan(spec, RECT)
        self.assertEqual(plan["edges"], [])
        self.assertEqual(len(plan["nodes"]), 6)
        text = " ".join(plan["warnings"])
        self.assertIn("draws no edges", text)
        self.assertIn("four cells", text)


class FitTests(unittest.TestCase):

    def test_font_steps_down_for_long_labels_and_stops_at_the_floor(self):
        short = dl.plan({"kind": "flow", "nodes": nodes("Ask", "Refer", "Close"),
                         "edges": chain("n1", "n2", "n3")}, RECT)
        self.assertEqual(short["fontPt"], dl.DEFAULTS["bodyPt"])
        long_labels = ["Partner describes the offer in their own words" for _ in range(8)]
        tight = dl.plan({"kind": "flow", "nodes": nodes(*long_labels)},
                        (60.0, 180.0, 300.0, 120.0))
        self.assertLess(tight["fontPt"], dl.DEFAULTS["bodyPt"])
        self.assertGreaterEqual(tight["fontPt"], dl.DEFAULTS["floorPt"])

    def test_node_width_comes_from_the_longest_label(self):
        opts = dl.options({})
        short = dl.fit_nodes(["Ask", "Refer"], 400.0, 200.0, opts)
        long_one = dl.fit_nodes(["Ask the partner at the easy moment", "Refer"],
                                400.0, 200.0, opts)
        self.assertGreater(long_one[1], short[1])
        self.assertLessEqual(long_one[1], dl.DEFAULTS["nodeMaxW"] + 0.01)

    def test_a_short_label_still_grows_to_the_target_width(self):
        opts = dl.options({})
        loose = dl.fit_nodes(["Ask"], 400.0, 200.0, opts, target_w=180.0)
        self.assertAlmostEqual(loose[1], 180.0, 1)
        capped = dl.fit_nodes(["Ask"], 90.0, 200.0, opts, target_w=180.0)
        self.assertAlmostEqual(capped[1], 90.0, 1)

    def test_a_node_box_is_wider_than_it_is_tall(self):
        for kind in ("flow", "timeline", "tree"):
            plan = dl.plan({"kind": kind,
                            "nodes": nodes("Spot", "Script", "Hand off"),
                            "edges": chain("n1", "n2", "n3")}, RECT)
            for node in plan["nodes"]:
                if node["shape"] != "roundRect":
                    continue
                self.assertGreater(node["w"], node["h"], (kind, node))
                self.assertGreaterEqual(node["h"], dl.DEFAULTS["nodeMinH"] - 0.01)

    def test_an_unfittable_label_is_truncated_with_a_warning(self):
        label = "an extremely long label that no small node box could ever hold at all"
        plan = dl.plan({"kind": "flow", "nodes": nodes(label, "Next", "Last")},
                       (60.0, 180.0, 240.0, 90.0))
        self.assertTrue(any("truncated" in w for w in plan["warnings"]), plan["warnings"])
        self.assertTrue(plan["nodes"][0]["label"].endswith(dl.ELLIPSIS))
        self.assertGreaterEqual(plan["fontPt"], dl.DEFAULTS["floorPt"])
        self.assertTrue(any("floor" in w for w in plan["warnings"]), plan["warnings"])

    def test_over_the_cap_truncates_the_node_list(self):
        plan = dl.plan({"kind": "flow",
                        "nodes": nodes(*["Step %d" % (i + 1) for i in range(15)])}, RECT)
        self.assertEqual(len(plan["nodes"]), dl.NODE_CAP)
        self.assertTrue(any("cap" in w for w in plan["warnings"]), plan["warnings"])


class InputTests(unittest.TestCase):

    def test_an_edge_to_a_missing_node_is_dropped_and_named(self):
        plan = dl.plan({"kind": "flow", "nodes": nodes("A", "B"),
                        "edges": [{"from": "n1", "to": "n9"}]}, RECT)
        self.assertEqual(plan["edges"], [])
        self.assertTrue(any("does not exist" in w and "n9" in w for w in plan["warnings"]),
                        plan["warnings"])

    def test_unknown_kind_and_direction_fall_back(self):
        plan = dl.plan({"kind": "sankey", "direction": "ZZ", "nodes": nodes("A", "B")}, RECT)
        self.assertEqual(plan["kind"], "flow")
        self.assertEqual(plan["direction"], "LR")
        text = " ".join(plan["warnings"])
        self.assertIn("unknown diagram kind", text)
        self.assertIn("unknown direction", text)

    def test_no_nodes_is_an_empty_plan_with_a_warning(self):
        plan = dl.plan({"kind": "flow", "nodes": []}, RECT)
        self.assertEqual(plan["nodes"], [])
        self.assertEqual(plan["edges"], [])
        self.assertTrue(any("no nodes" in w for w in plan["warnings"]))

    def test_plain_string_nodes_get_ids(self):
        plan = dl.plan({"kind": "flow", "nodes": ["Spot", "Ask"]}, RECT)
        self.assertEqual([n["id"] for n in plan["nodes"]], ["n1", "n2"])
        self.assertEqual([n["label"] for n in plan["nodes"]], ["Spot", "Ask"])

    def test_a_tiny_block_warns_instead_of_drawing(self):
        plan = dl.plan({"kind": "flow", "nodes": nodes("A", "B")}, (60.0, 180.0, 20.0, 20.0))
        self.assertEqual(plan["nodes"], [])
        self.assertTrue(any("too small" in w for w in plan["warnings"]))

    def test_every_kind_stays_inside_the_rect_and_never_overlaps(self):
        labels = ["Spot the moment", "Script the ask", "Hand off", "Follow up",
                  "Review", "Renew"]
        for kind in dl.KINDS:
            for count in (3, 4, 6):
                spec = {"kind": kind, "nodes": nodes(*labels[:count]),
                        "edges": chain(*["n%d" % (i + 1) for i in range(count)])}
                plan = dl.plan(spec, RECT)
                self.assertTrue(plan["nodes"], (kind, count, plan["warnings"]))
                for node in plan["nodes"]:
                    self.assertTrue(inside(node, RECT), (kind, count, node))
                self.assertFalse(dl.any_overlap(plan["nodes"]), (kind, count))
                self.assertGreaterEqual(plan["fontPt"], dl.DEFAULTS["floorPt"])

    def test_summary_names_the_kind_and_the_labels(self):
        text = dl.summary({"kind": "flow", "nodes": nodes("Prepare", "Ask", "Hand off")})
        self.assertIn("flow", text)
        self.assertIn("Prepare", text)
        self.assertIn("Hand off", text)
        self.assertIn("→", text)
        self.assertIn("Quick wins", dl.summary({"kind": "matrix",
                                                "nodes": nodes("Effort", "Value", "Quick wins")}))


if __name__ == "__main__":
    unittest.main(verbosity=2)

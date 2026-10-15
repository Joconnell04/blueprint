#!/usr/bin/env python3
"""diagram_layout.py: geometry for the deck `diagram` block.

Pure functions. No python-pptx, no I/O, no state: `plan(content, rect, opts)`
takes the block's content and its rectangle in points and returns every node
box, every connector polyline and every warning. build_pptx.py draws the result
and the canvas sketches from the same numbers, so both agree on the fit.

Five kinds:

    flow      layered left to right (or top to bottom): rank by the longest path
              from the sources, order within a rank by barycenter
    cycle     nodes on a ring, arrows clockwise from the top
    timeline  nodes on a horizontal spine, boxes alternating above and below
    tree      layered top to bottom from the root
    matrix    a square 2x2 grid; nodes[0] and nodes[1] are the axis labels and
              the next four fill the cells in reading order

Fit: the node width comes from the longest label at the body size and steps down
one point at a time to the caption floor. Over `NODE_CAP` nodes, or a diagram
that still does not fit at the floor, is a warning plus truncation - never an
overlap and never type below the floor.

Contract: builder/SPEC-deck.md. Python 3.9, standard library only.
"""

import math

KINDS = ("flow", "cycle", "timeline", "tree", "matrix")
DIRECTIONS = ("LR", "TB")
NODE_CAP = 12
MATRIX_CELLS = 4
CYCLE_MIN = 3
ELLIPSIS = "…"

# Geometry defaults. A style's `diagram` token block overrides any key; the
# builder passes bodyPt from the diagramNode role and floorPt from the caption
# floor, so a style that scales its type scales its diagrams too.
DEFAULTS = {
    "bodyPt": 12.0,          # the node label size the fit starts from
    "floorPt": 8.0,          # the caption floor; nothing is ever set below it
    "k": 0.5,                # average glyph width in ems, as build_pptx.cpl uses
    "leading": 1.2,
    "maxNodeLines": 2,
    "insetPt": 6.0,          # text inset inside a node box
    "gutterPt": 18.0,        # between node boxes and between ranks
    "nodeMinW": 54.0,
    "nodeMaxW": 260.0,
    "nodeMinH": 26.0,
    "nodeMaxH": 130.0,
    "edgeGapPt": 4.0,        # between a node box and the end of its connector
    "edgeLabelGapPt": 7.0,   # an edge label clears its connector by this much
    "fillShare": 0.78,       # LR: the share of a rank band one node box takes
    "stackW": 110.0,         # TB and ring: the node width a short label grows to
    "aspect": 0.38,          # node height as a share of its width, before the text
    "slotShare": 0.45,       # a node box grows to this share of the height it is given
    "maxRatio": 0.7,         # and never taller than this share of its own width
    "timelineShare": 0.86,   # timeline: the share of its slot a node box takes
    "ringRatio": 1.6,        # cycle: the widest the ring may stretch horizontally
    "spineGapPt": 4.0,       # timeline: between the spine and a node box
    "axisGapPt": 10.0,       # matrix: between an axis label and the grid
    "labelMaxW": 120.0,      # matrix: the widest y-axis label column
    "labelMaxH": 26.0,       # matrix: the tallest x-axis label strip
    "axisPt": 8.0,           # matrix: the axis label size (the caption size)
}


# ------------------------------------------------------------------- measuring


def cpl(width, size, k):
    """Characters per line at this width and size. Mirrors build_pptx.cpl."""
    return max(int(math.floor(float(width) / max(float(size) * float(k), 0.1))), 1)


def wrap_lines(text, chars):
    """Greedy word wrap, one entry per rendered line. Mirrors build_pptx.wrap_lines."""
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


def options(over=None):
    """DEFAULTS with `over` merged on top, non-numeric entries ignored."""
    out = dict(DEFAULTS)
    for key, value in (over or {}).items():
        if key not in out or value is None:
            continue
        try:
            out[key] = float(value) if not isinstance(value, bool) else out[key]
        except (TypeError, ValueError):
            continue
    out["maxNodeLines"] = max(int(out["maxNodeLines"]), 1)
    return out


def wrap_label(label, width, size, o):
    """The lines `label` takes inside a box of this outer width."""
    inner = max(float(width) - 2.0 * o["insetPt"], 12.0)
    return wrap_lines(label, cpl(inner, size, o["k"]))


def label_height(lines, size, o):
    return len(lines) * float(size) * o["leading"] + 2.0 * o["insetPt"]


def label_fits(label, width, height, size, o):
    lines = wrap_label(label, width, size, o)
    if len(lines) > o["maxNodeLines"]:
        return False
    return label_height(lines, size, o) <= float(height) + 0.51


def ideal_width(labels, size, o):
    """The node width the longest label wants at `size`, inside the min and max."""
    need = o["nodeMinW"]
    for label in labels:
        words = str(label or "").split()
        if not words:
            continue
        longest = max(len(w) for w in words)
        whole = len(" ".join(words))
        chars = max(longest, int(math.ceil(whole / float(o["maxNodeLines"]))))
        need = max(need, chars * float(size) * o["k"] + 2.0 * o["insetPt"])
    return min(need, o["nodeMaxW"])


def truncate_label(label, width, height, size, o):
    """Cut a label to the lines that fit and mark the cut with an ellipsis."""
    inner = max(float(width) - 2.0 * o["insetPt"], 12.0)
    chars = cpl(inner, size, o["k"])
    keep = max(int(o["maxNodeLines"]), 1)
    room = max(int(math.floor((float(height) - 2.0 * o["insetPt"])
                              / max(float(size) * o["leading"], 1.0))), 1)
    keep = max(min(keep, room), 1)
    lines = wrap_lines(label, chars)
    if len(lines) <= keep:
        return str(label), lines
    cut = lines[:keep]
    tail = cut[-1]
    if len(tail) + 1 > chars and len(tail) > 1:
        tail = tail[:max(chars - 1, 1)].rstrip()
    cut[-1] = tail + ELLIPSIS
    return " ".join(cut), cut


def fit_nodes(labels, slot_w, slot_h, o, target_w=0.0, aspect=None):
    """Fit every node label into one shared size and box.

    Returns (size, node_w, node_h, lines_per_label, labels_out, warnings). The
    width comes from the longest label at the body size, grown to `target_w` so
    a short label still reads as a box, and stepped down one point at a time; at
    the floor an unfittable label is truncated with a warning. The height is the
    taller of the text and `aspect` of the width, inside the slot.
    """
    warnings = []
    labels = [str(x if x is not None else "") for x in labels]
    slot_w = max(float(slot_w), 12.0)
    slot_h = max(float(slot_h), 12.0)
    size = float(o["bodyPt"])
    floor = float(o["floorPt"])
    share = float(o["aspect"]) if aspect is None else float(aspect)
    node_w = min(max(o["nodeMinW"], 0.0), slot_w)
    node_h = min(max(o["nodeMinH"], 0.0), slot_h)
    while True:
        want = max(ideal_width(labels, size, o), float(target_w or 0.0))
        node_w = min(max(want, min(o["nodeMinW"], slot_w)), slot_w, o["nodeMaxW"])
        tallest = 1
        for label in labels:
            tallest = max(tallest, len(wrap_label(label, node_w, size, o)))
        tallest = min(tallest, int(o["maxNodeLines"]))
        text_h = label_height([""] * tallest, size, o)
        grown = max(text_h, o["nodeMinH"], node_w * share, slot_h * o["slotShare"])
        node_h = min(grown, o["nodeMaxH"], slot_h, max(node_w * o["maxRatio"], text_h))
        if all(label_fits(x, node_w, node_h, size, o) for x in labels):
            break
        if size - 1.0 < floor - 0.001:
            break
        size -= 1.0
    out_labels = []
    lines = []
    for label in labels:
        if label_fits(label, node_w, node_h, size, o):
            out_labels.append(label)
            lines.append(wrap_label(label, node_w, size, o))
            continue
        cut, cut_lines = truncate_label(label, node_w, node_h, size, o)
        out_labels.append(cut)
        lines.append(cut_lines)
        warnings.append("label '%s' does not fit at %g pt; truncated" % (label, size))
    return size, node_w, node_h, lines, out_labels, warnings


# -------------------------------------------------------------------- reading


def normalise(content):
    """(kind, direction, nodes, edges, warnings) from a block's content."""
    warnings = []
    data = content if isinstance(content, dict) else {}
    kind = str(data.get("kind") or "flow").strip().lower()
    if kind not in KINDS:
        if kind:
            warnings.append("unknown diagram kind '%s'; drew a flow" % kind)
        kind = "flow"
    direction = str(data.get("direction") or "").strip().upper()
    if direction not in DIRECTIONS:
        if direction:
            warnings.append("unknown direction '%s'; used %s"
                            % (direction, "TB" if kind == "tree" else "LR"))
        direction = "TB" if kind == "tree" else "LR"

    nodes = []
    seen = set()
    for i, raw in enumerate(data.get("nodes") or []):
        if isinstance(raw, dict):
            node_id = str(raw.get("id") or "").strip()
            label = raw.get("label")
            if label is None:
                label = raw.get("text")
        else:
            node_id, label = "", raw
        if not node_id:
            node_id = "n%d" % (i + 1)
        if node_id in seen:
            warnings.append("duplicate node id '%s' dropped" % node_id)
            continue
        seen.add(node_id)
        nodes.append({"id": node_id, "label": str(label if label is not None else "")})
    if len(nodes) > NODE_CAP:
        warnings.append("%d nodes is over the cap of %d; kept the first %d"
                        % (len(nodes), NODE_CAP, NODE_CAP))
        nodes = nodes[:NODE_CAP]
        seen = set(n["id"] for n in nodes)

    edges = []
    missing = []
    for raw in data.get("edges") or []:
        if not isinstance(raw, dict):
            continue
        src = str(raw.get("from") or raw.get("source") or "").strip()
        dst = str(raw.get("to") or raw.get("target") or "").strip()
        label = raw.get("label")
        if src not in seen or dst not in seen:
            missing.append("%s->%s" % (src or "?", dst or "?"))
            continue
        if src == dst:
            warnings.append("edge %s->%s points at itself; dropped" % (src, dst))
            continue
        edges.append({"from": src, "to": dst,
                      "label": str(label) if label not in (None, "") else ""})
    if missing:
        warnings.append("edge to a node that does not exist: %s" % ", ".join(missing))
    if not nodes:
        warnings.append("diagram has no nodes")
    return kind, direction, nodes, edges, warnings


# -------------------------------------------------------------------- helpers


def _centre(box):
    return (box["x"] + box["w"] / 2.0, box["y"] + box["h"] / 2.0)


def _overlap(a, b):
    dx = min(a["x"] + a["w"], b["x"] + b["w"]) - max(a["x"], b["x"])
    dy = min(a["y"] + a["h"], b["y"] + b["h"]) - max(a["y"], b["y"])
    return dx > 0.01 and dy > 0.01


def any_overlap(nodes):
    """True when two node boxes cover any common area. The layouts never do."""
    boxes = [n for n in nodes if n.get("shape") != "text"]
    for i, a in enumerate(boxes):
        for b in boxes[i + 1:]:
            if _overlap(a, b):
                return True
    return False


def _exit_point(box, toward, gap):
    """Where a straight line to `toward` leaves `box`, pulled back by `gap`."""
    cx, cy = _centre(box)
    dx, dy = toward[0] - cx, toward[1] - cy
    if abs(dx) < 0.001 and abs(dy) < 0.001:
        return (cx, cy)
    half_w, half_h = box["w"] / 2.0 + gap, box["h"] / 2.0 + gap
    scale = float("inf")
    if abs(dx) > 0.001:
        scale = min(scale, half_w / abs(dx))
    if abs(dy) > 0.001:
        scale = min(scale, half_h / abs(dy))
    return (cx + dx * scale, cy + dy * scale)


def _label_spot(points, gap):
    """The label anchor: half way along the polyline, clear of the line itself."""
    point = _mid_point(points)
    dx, dy = 0.0, 0.0
    for i in range(len(points) - 1):
        dx = points[i + 1][0] - points[i][0]
        dy = points[i + 1][1] - points[i][1]
        if abs(dx) > 0.01 or abs(dy) > 0.01:
            length = math.hypot(dx, dy)
            if length > 0:
                mid = _mid_point(points)
                if abs(mid[0] - points[i][0]) + abs(mid[1] - points[i][1]) <= length + 0.5:
                    break
    if abs(dx) >= abs(dy):
        return (round(point[0], 2), round(point[1] - gap, 2))
    return (round(point[0] + gap, 2), round(point[1], 2))


def _mid_point(points):
    """The point half way along a polyline, and the segment it sits on."""
    if len(points) < 2:
        return points[0] if points else (0.0, 0.0)
    spans = []
    total = 0.0
    for i in range(len(points) - 1):
        (x1, y1), (x2, y2) = points[i], points[i + 1]
        length = math.hypot(x2 - x1, y2 - y1)
        spans.append(length)
        total += length
    if total <= 0:
        return points[0]
    want = total / 2.0
    run = 0.0
    for i, length in enumerate(spans):
        if run + length >= want:
            share = (want - run) / max(length, 0.001)
            (x1, y1), (x2, y2) = points[i], points[i + 1]
            return (x1 + (x2 - x1) * share, y1 + (y2 - y1) * share)
        run += length
    return points[-1]


def rank_nodes(nodes, edges):
    """Longest path from the sources. A cycle saturates instead of looping."""
    index = dict((n["id"], i) for i, n in enumerate(nodes))
    ranks = [0] * len(nodes)
    for _ in range(len(nodes)):
        moved = False
        for edge in edges:
            a, b = index[edge["from"]], index[edge["to"]]
            if ranks[b] < ranks[a] + 1 and ranks[a] + 1 <= len(nodes) - 1:
                ranks[b] = ranks[a] + 1
                moved = True
        if not moved:
            break
    return ranks


def order_ranks(nodes, edges, ranks, sweeps=2):
    """Order within each rank by the barycenter of the neighbours, input order first."""
    index = dict((n["id"], i) for i, n in enumerate(nodes))
    groups = {}
    for i, rank in enumerate(ranks):
        groups.setdefault(rank, []).append(i)
    preds = dict((i, []) for i in range(len(nodes)))
    succs = dict((i, []) for i in range(len(nodes)))
    for edge in edges:
        a, b = index[edge["from"]], index[edge["to"]]
        preds[b].append(a)
        succs[a].append(b)
    place = {}
    for rank in sorted(groups):
        for pos, node in enumerate(groups[rank]):
            place[node] = float(pos)

    def sweep(order_keys, neighbours):
        for rank in order_keys:
            group = groups[rank]
            scores = []
            for pos, node in enumerate(group):
                near = [place[x] for x in neighbours[node] if x in place]
                scores.append((sum(near) / len(near) if near else place[node], pos, node))
            scores.sort(key=lambda item: (item[0], item[1]))
            groups[rank] = [node for _, _, node in scores]
            for pos, node in enumerate(groups[rank]):
                place[node] = float(pos)

    keys = sorted(groups)
    for _ in range(max(int(sweeps), 1)):
        sweep(keys[1:], preds)
        sweep(list(reversed(keys[:-1])), succs)
    return groups


# -------------------------------------------------------------------- layouts


def _layered(nodes, edges, rect, direction, o, ranks, groups):
    """Shared geometry for flow and tree: ranks across, order within."""
    x, y, w, h = rect
    gut = o["gutterPt"]
    keys = sorted(groups)
    depth = len(keys)
    widest = max(len(groups[k]) for k in keys)
    labels = [n["label"] for n in nodes]

    if direction == "LR":
        band = (w - gut * (depth - 1)) / float(depth)
        slot = (h - gut * (widest - 1)) / float(widest)
        size, node_w, node_h, lines, texts, warns = fit_nodes(
            labels, band, slot, o, target_w=band * o["fillShare"])
    else:
        band = (h - gut * (depth - 1)) / float(depth)
        slot = (w - gut * (widest - 1)) / float(widest)
        size, node_w, node_h, lines, texts, warns = fit_nodes(
            labels, slot, band, o, target_w=min(o["stackW"], slot),
            aspect=o["aspect"] * 0.9)
        # leave a connector run between the ranks, however tight the band is
        node_h = min(node_h, max(band - 2 * o["edgeGapPt"], o["nodeMinH"]))

    out = [None] * len(nodes)
    for step, key in enumerate(keys):
        group = groups[key]
        count = len(group)
        if direction == "LR":
            centre = x + step * (band + gut) + band / 2.0
            total = count * node_h + (count - 1) * gut
            start = y + (h - total) / 2.0
            for pos, node in enumerate(group):
                out[node] = {"x": centre - node_w / 2.0, "y": start + pos * (node_h + gut),
                             "w": node_w, "h": node_h}
        else:
            centre = y + step * (band + gut) + band / 2.0
            total = count * node_w + (count - 1) * gut
            start = x + (w - total) / 2.0
            for pos, node in enumerate(group):
                out[node] = {"x": start + pos * (node_w + gut), "y": centre - node_h / 2.0,
                             "w": node_w, "h": node_h}
        for pos, node in enumerate(group):
            out[node]["rank"] = step
            out[node]["order"] = pos

    placed = []
    for i, node in enumerate(nodes):
        box = out[i]
        placed.append({"id": node["id"], "label": texts[i], "lines": lines[i],
                       "x": round(box["x"], 2), "y": round(box["y"], 2),
                       "w": round(box["w"], 2), "h": round(box["h"], 2),
                       "fontPt": size, "rank": box["rank"], "order": box["order"],
                       "shape": "roundRect", "align": "center"})
    return placed, size, warns, ranks


def _layered_edges(placed, edges, direction, o):
    """Elbow connectors between ranks, straight when the centres line up."""
    by_id = dict((n["id"], n) for n in placed)
    gap = o["edgeGapPt"]
    out = []
    for edge in edges:
        a, b = by_id[edge["from"]], by_id[edge["to"]]
        ax, ay = _centre(a)
        bx, by = _centre(b)
        if direction == "LR" and b["rank"] > a["rank"]:
            start = (a["x"] + a["w"] + gap, ay)
            end = (b["x"] - gap, by)
            mid = (start[0] + end[0]) / 2.0
            points = ([start, end] if abs(ay - by) < 1.0
                      else [start, (mid, ay), (mid, by), end])
        elif direction == "TB" and b["rank"] > a["rank"]:
            start = (ax, a["y"] + a["h"] + gap)
            end = (bx, b["y"] - gap)
            mid = (start[1] + end[1]) / 2.0
            points = ([start, end] if abs(ax - bx) < 1.0
                      else [start, (ax, mid), (bx, mid), end])
        else:
            points = [_exit_point(a, (bx, by), gap), _exit_point(b, (ax, ay), gap)]
        points = [(round(px, 2), round(py, 2)) for px, py in points]
        out.append({"from": edge["from"], "to": edge["to"], "label": edge["label"],
                    "points": points, "labelXY": _label_spot(points, o["edgeLabelGapPt"]),
                    "kind": "elbow" if len(points) > 2 else "straight"})
    return out


def layout_flow(nodes, edges, rect, direction, o):
    ranks = rank_nodes(nodes, edges)
    groups = order_ranks(nodes, edges, ranks)
    placed, size, warns, ranks = _layered(nodes, edges, rect, direction, o, ranks, groups)
    return {"nodes": placed, "edges": _layered_edges(placed, edges, direction, o),
            "axis": [], "fontPt": size, "warnings": warns}


def layout_tree(nodes, edges, rect, direction, o):
    targets = set(e["to"] for e in edges)
    root = None
    for node in nodes:
        if node["id"] not in targets:
            root = node["id"]
            break
    if root is None:
        root = nodes[0]["id"]
    ranks = rank_nodes(nodes, edges)
    groups = order_ranks(nodes, edges, ranks)
    placed, size, warns, ranks = _layered(nodes, edges, rect, direction, o, ranks, groups)
    out = {"nodes": placed, "edges": _layered_edges(placed, edges, direction, o),
           "axis": [], "fontPt": size, "warnings": warns}
    out["root"] = root
    return out


def layout_cycle(nodes, edges, rect, o):
    """Nodes on a ring from the top, clockwise. Ring edges when none are given."""
    warns = []
    nodes = list(nodes)
    if not edges:
        edges = [{"from": nodes[i]["id"], "to": nodes[(i + 1) % len(nodes)]["id"], "label": ""}
                 for i in range(len(nodes))] if len(nodes) > 1 else []
    x, y, w, h = rect
    gut = o["gutterPt"]
    placed = []
    while True:
        count = len(nodes)
        labels = [n["label"] for n in nodes]
        slot_w = max((w - 2 * gut) / 3.0, o["nodeMinW"])
        slot_h = max((h - 2 * gut) / 3.0, o["nodeMinH"])
        size, node_w, node_h, lines, texts, fit_warns = fit_nodes(
            labels, slot_w, slot_h, o, target_w=min(o["stackW"], slot_w),
            aspect=o["aspect"] * 0.9)
        shrink = 0
        while True:
            use_w = max(node_w - shrink, min(o["nodeMinW"], slot_w))
            use_h = max(node_h - shrink * (node_h / max(node_w, 1.0)), o["nodeMinH"])
            radius_y = max((h - use_h) / 2.0, 1.0)
            radius_x = max(min((w - use_w) / 2.0, radius_y * o["ringRatio"]), 1.0)
            cx, cy = x + w / 2.0, y + h / 2.0
            ring = {"x": round(cx, 2), "y": round(cy, 2),
                    "rx": round(radius_x, 2), "ry": round(radius_y, 2)}
            trial = []
            for i in range(count):
                angle = math.radians(-90.0 + 360.0 * i / float(max(count, 1)))
                px = cx + radius_x * math.cos(angle)
                py = cy + radius_y * math.sin(angle)
                trial.append({"x": px - use_w / 2.0, "y": py - use_h / 2.0,
                              "w": use_w, "h": use_h})
            if count < 2 or not any_overlap([dict(t, shape="roundRect") for t in trial]):
                placed = []
                for i, node in enumerate(nodes):
                    box = trial[i]
                    placed.append({"id": node["id"], "label": texts[i], "lines": lines[i],
                                   "x": round(box["x"], 2), "y": round(box["y"], 2),
                                   "w": round(box["w"], 2), "h": round(box["h"], 2),
                                   "fontPt": size, "rank": 0, "order": i,
                                   "shape": "roundRect", "align": "center"})
                break
            if use_w <= min(o["nodeMinW"], slot_w) + 0.01:
                placed = []
                break
            shrink += 4.0
        if placed or len(nodes) <= CYCLE_MIN:
            warns.extend(fit_warns)
            break
        warns.append("a ring of %d nodes does not fit; dropped '%s'"
                     % (len(nodes), nodes[-1]["label"]))
        dropped = nodes.pop()["id"]
        edges = [e for e in edges if e["from"] != dropped and e["to"] != dropped]
    if not placed:
        warns.append("the ring does not fit this block")
        return {"nodes": [], "edges": [], "axis": [], "fontPt": o["floorPt"], "warnings": warns}

    by_id = dict((n["id"], n) for n in placed)
    out_edges = []
    for edge in edges:
        if edge["from"] not in by_id or edge["to"] not in by_id:
            continue
        a, b = by_id[edge["from"]], by_id[edge["to"]]
        start = _exit_point(a, _centre(b), o["edgeGapPt"])
        end = _exit_point(b, _centre(a), o["edgeGapPt"])
        points = [(round(start[0], 2), round(start[1], 2)),
                  (round(end[0], 2), round(end[1], 2))]
        out_edges.append({"from": edge["from"], "to": edge["to"], "label": edge["label"],
                          "points": points, "labelXY": _label_spot(points, o["edgeLabelGapPt"]),
                          "kind": "straight"})
    return {"nodes": placed, "edges": out_edges, "axis": [], "fontPt": placed[0]["fontPt"],
            "ring": ring, "warnings": warns}


def layout_timeline(nodes, edges, rect, o):
    """A horizontal spine with the boxes alternating above and below it."""
    x, y, w, h = rect
    gut = o["gutterPt"]
    count = len(nodes)
    labels = [n["label"] for n in nodes]
    slot = (w - gut * (count - 1)) / float(count)
    half = max(h / 2.0 - o["spineGapPt"], o["nodeMinH"])
    size, node_w, node_h, lines, texts, warns = fit_nodes(
        labels, slot, half, o, target_w=slot * o["timelineShare"],
        aspect=o["aspect"] * 0.9)
    node_h = min(node_h, half)
    spine_y = y + h / 2.0
    placed = []
    for i, node in enumerate(nodes):
        left = x + i * (slot + gut) + (slot - node_w) / 2.0
        above = (i % 2 == 0)
        top = spine_y - o["spineGapPt"] - node_h if above else spine_y + o["spineGapPt"]
        placed.append({"id": node["id"], "label": texts[i], "lines": lines[i],
                       "x": round(left, 2), "y": round(top, 2),
                       "w": round(node_w, 2), "h": round(node_h, 2),
                       "fontPt": size, "rank": i, "order": 0 if above else 1,
                       "shape": "roundRect", "align": "center",
                       "side": "above" if above else "below"})
    first = placed[0]
    last = placed[-1]
    spine_x = first["x"] + first["w"] / 2.0
    spine_w = max(last["x"] + last["w"] / 2.0 - spine_x, 1.0)
    if count == 1:
        spine_x, spine_w = x, w
    axis = [{"kind": "rule", "x": round(spine_x, 2), "y": round(spine_y, 2),
             "w": round(spine_w, 2), "h": 1.0}]
    by_id = dict((n["id"], n) for n in placed)
    out_edges = []
    for edge in edges:
        a, b = by_id[edge["from"]], by_id[edge["to"]]
        start = _exit_point(a, _centre(b), o["edgeGapPt"])
        end = _exit_point(b, _centre(a), o["edgeGapPt"])
        points = [(round(start[0], 2), round(start[1], 2)),
                  (round(end[0], 2), round(end[1], 2))]
        out_edges.append({"from": edge["from"], "to": edge["to"], "label": edge["label"],
                          "points": points, "labelXY": _label_spot(points, o["edgeLabelGapPt"]),
                          "kind": "straight"})
    return {"nodes": placed, "edges": out_edges, "axis": axis, "fontPt": size,
            "warnings": warns}


def layout_matrix(nodes, edges, rect, o):
    """A square 2x2: nodes[0] and nodes[1] label the axes, the next four fill the cells."""
    warns = []
    if edges:
        warns.append("a matrix draws no edges; %d dropped" % len(edges))
    if len(nodes) < 3:
        warns.append("a matrix needs two axis labels and at least one cell; got %d"
                     % len(nodes))
    cells = nodes[2:2 + MATRIX_CELLS]
    if len(nodes) > 2 + MATRIX_CELLS:
        warns.append("a matrix holds four cells; dropped %d node(s)"
                     % (len(nodes) - 2 - MATRIX_CELLS))
    if len(nodes) >= 3 and len(cells) < MATRIX_CELLS:
        warns.append("a matrix has four cells; %d given" % len(cells))

    x, y, w, h = rect
    gut = o["gutterPt"]
    label_w = min(o["labelMaxW"], w * 0.18)
    label_h = min(o["labelMaxH"], h * 0.12)
    grid_w = max(w - label_w - o["axisGapPt"], 24.0)
    grid_h = max(h - label_h - o["axisGapPt"], 24.0)
    side = min(grid_w, grid_h)
    # centre the whole assembly, so the y-axis label sits against the grid
    spare = max(w - (label_w + o["axisGapPt"] + side), 0.0)
    label_x = x + spare / 2.0
    grid_x = label_x + label_w + o["axisGapPt"]
    grid_y = y + (grid_h - side) / 2.0
    cell_w = (side - gut) / 2.0
    cell_h = (side - gut) / 2.0

    labels = [n["label"] for n in cells]
    size, _, _, lines, texts, fit_warns = fit_nodes(labels or [""], cell_w, cell_h, o)
    warns.extend(fit_warns)
    spots = [("tl", grid_x, grid_y), ("tr", grid_x + cell_w + gut, grid_y),
             ("bl", grid_x, grid_y + cell_h + gut),
             ("br", grid_x + cell_w + gut, grid_y + cell_h + gut)]
    placed = []
    if len(nodes) >= 1:
        placed.append({"id": nodes[0]["id"], "label": nodes[0]["label"],
                       "lines": [nodes[0]["label"]],
                       "x": round(grid_x, 2), "y": round(grid_y + side + o["axisGapPt"], 2),
                       "w": round(side, 2), "h": round(label_h, 2),
                       "fontPt": o["axisPt"], "rank": 0, "order": 0,
                       "shape": "text", "align": "center", "axis": "x"})
    if len(nodes) >= 2:
        placed.append({"id": nodes[1]["id"], "label": nodes[1]["label"],
                       "lines": [nodes[1]["label"]],
                       "x": round(label_x, 2), "y": round(grid_y, 2),
                       "w": round(label_w, 2), "h": round(side, 2),
                       "fontPt": o["axisPt"], "rank": 0, "order": 1,
                       "shape": "text", "align": "right", "axis": "y"})
    for i, node in enumerate(cells):
        cell, cx, cy = spots[i]
        placed.append({"id": node["id"], "label": texts[i], "lines": lines[i],
                       "x": round(cx, 2), "y": round(cy, 2),
                       "w": round(cell_w, 2), "h": round(cell_h, 2),
                       "fontPt": size, "rank": 0 if cell in ("tl", "tr") else 1,
                       "order": 0 if cell in ("tl", "bl") else 1,
                       "shape": "roundRect", "align": "center", "cell": cell})
    return {"nodes": placed, "edges": [], "axis": [], "fontPt": size, "warnings": warns}


# ----------------------------------------------------------------- the planner


def plan(content, rect, opts=None):
    """Lay out one diagram block inside `rect` = (x, y, w, h) in points.

    Returns {kind, direction, fontPt, nodes, edges, axis, warnings}. Every node
    carries x/y/w/h/label/lines/fontPt and `shape` ("roundRect", or "text" for a
    matrix axis label); every edge carries a `points` polyline whose last
    segment takes the arrowhead, plus `labelXY` for its caption.
    """
    kind, direction, nodes, edges, warnings = normalise(content)
    o = options(opts)
    x, y, w, h = [float(v) for v in rect]
    body = {"kind": kind, "direction": direction, "nodes": [], "edges": [], "axis": [],
            "fontPt": o["bodyPt"], "warnings": list(warnings)}
    if not nodes or w < 24 or h < 24:
        if w < 24 or h < 24:
            body["warnings"].append("the block is too small for a diagram")
        return body
    inner = (x, y, w, h)
    if kind == "cycle":
        out = layout_cycle(nodes, edges, inner, o)
    elif kind == "timeline":
        out = layout_timeline(nodes, edges, inner, o)
    elif kind == "tree":
        out = layout_tree(nodes, edges, inner, direction, o)
    elif kind == "matrix":
        out = layout_matrix(nodes, edges, inner, o)
    else:
        out = layout_flow(nodes, edges, inner, direction, o)
    body["nodes"] = out["nodes"]
    body["edges"] = out["edges"]
    body["axis"] = out.get("axis") or []
    body["fontPt"] = out.get("fontPt") or o["bodyPt"]
    body["warnings"].extend(out.get("warnings") or [])
    if body["fontPt"] <= o["floorPt"] + 0.001 and body["nodes"]:
        body["warnings"].append("diagram labels are at the %g pt floor" % o["floorPt"])
    for key in ("root", "ring"):
        if key in out:
            body[key] = out[key]
    return body


def summary(content):
    """The one-line description a placeholder carries: the kind and the node labels."""
    kind, _, nodes, _, _ = normalise(content)
    labels = [n["label"] for n in nodes if str(n["label"]).strip()]
    joiner = ", " if kind == "matrix" else " → "
    body = joiner.join(labels)
    head = "Designed %s diagram" % kind
    return ("%s: %s" % (head, body)) if body else head

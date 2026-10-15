#!/usr/bin/env python3
"""fit.py: the fidelity pass and the deck-wide type chooser for build_pptx.py.

Pure functions over integer grid cells. No python-pptx, no file IO and no
measurement of its own, so the canvas can run the same numbers from the same
source. build_pptx.py owns the type scale and therefore owns measurement; it
hands this module the facts (what a block needs, what slack it has) and this
module decides the geometry.

Three fidelities, from the spec's `fidelity` field:

* `strict`   the approved geometry is drawn exactly as planned.
* `balanced` (the default) every left edge and width snaps to the style's
  column module; a block may grow up to 2 rows or columns into free space when
  its content does not fit at the role's floor, taking the cells from a
  neighbour that has slack; an empty band deeper than 3 rows at the bottom of
  the body is rebalanced; edges within one cell of each other snap.
* `loose`    balanced, plus: content that clearly does not suit its slot (a
  seven-column table in twelve columns, five cards in a three-card slot, a
  stat that is really a sentence) is re-picked into the archetype slot that
  fits it, from docs/research/deck-planning.md section 5.

Every deviation comes back as `{"blockId", "from", "to", "reason"}` so the
build records it in the snapshot and the canvas can show it.

Python 3.9, standard library only.
"""

import math

# --------------------------------------------------------------- constants

FIDELITIES = ("strict", "balanced", "loose")
DEFAULT_FIDELITY = "balanced"

BODY_TOP = 5            # the body area, 1-based rows, inclusive
BODY_BOTTOM = 17
MAX_GROW = 2            # cells a block may gain in one dimension
ALIGN_SNAP = 1          # edges this close snap together
EMPTY_BAND = 3          # a band deeper than this is rebalanced
MIN_CELLS = 1           # no block is left thinner than this

# F-3: a slide with a source line needs the last body row for it. The solver
# keeps row 17 clear, shifting the band that reaches it up by one row when the
# space is there, and reporting the block when it is not.
SOURCE_ROW = BODY_BOTTOM
SOURCE_REASON = "moved up to keep row %d for the source line" % SOURCE_ROW

# Q21 (loose): a stat band beside a text or bullets rail is a figure band that
# was boxed in. The rail goes below and the stats take the full-width slot.
STAT_RAIL_TYPES = ("text", "bullets")
STAT_BAND_REASON = "stat-band"
STAT_BAND_SHARE = 0.75  # a stat already this wide is not boxed in

# The column module (Q9). The style's `furniture.body` is the content measure;
# the module cuts it into MODULE_COLS columns with a one-cell gutter, so every
# block on every slide lands on one of the same few left edges. An edge reaches
# for a module edge up to MODULE_SNAP cells away and no further.
CELL = 30.0
MODULE_COLS = 12
MODULE_SNAP = 2
MODULE_REASON = "module"

# Archetype slots, docs/research/deck-planning.md section 5. Only the slots a
# re-pick can land in: the catalogue's job here is to answer "where does this
# content belong instead", not to re-plan the slide.
ARCHETYPE_SLOTS = {
    "table": {
        "table": {"col": 2, "row": 7, "cols": 30, "rows": 9},
    },
    "three-cards": {
        "cards": {"col": 3, "row": 7, "cols": 28, "rows": 10},
    },
    "two-by-two": {
        "cards": {"col": 3, "row": 7, "cols": 28, "rows": 9},
    },
    "wide-cards": {
        "cards": {"col": 2, "row": 7, "cols": 30, "rows": 10},
    },
    "stat-band": {
        "stat": {"col": 2, "row": 7, "cols": 30, "rows": 5},
    },
    "hero-stat": {
        "stat": {"col": 3, "row": 7, "cols": 11, "rows": 6},
    },
    "statement": {
        "statement": {"col": 3, "row": 8, "cols": 22, "rows": 6},
    },
    "points": {
        "list": {"col": 3, "row": 7, "cols": 20, "rows": 10},
    },
    "image-full": {
        "visual": {"col": 1, "row": 1, "cols": 32, "rows": 18},
    },
}


def fidelity_of(spec, warn=None):
    """The spec's fidelity, defaulted and validated."""
    raw = spec.get("fidelity") if isinstance(spec, dict) else None
    name = str(raw or "").strip().lower()
    if not name:
        return DEFAULT_FIDELITY
    if name not in FIDELITIES:
        if warn:
            warn("unknown fidelity '%s'; used %s" % (raw, DEFAULT_FIDELITY))
        return DEFAULT_FIDELITY
    return name


# ------------------------------------------------------------------- cells


def cells(col, row, cols, rows):
    return {"col": int(col), "row": int(row),
            "cols": max(int(cols), MIN_CELLS), "rows": max(int(rows), MIN_CELLS)}


def copy_cells(cell):
    return cells(cell["col"], cell["row"], cell["cols"], cell["rows"])


def same_cells(a, b):
    return all(int(a[k]) == int(b[k]) for k in ("col", "row", "cols", "rows"))


def left(cell):
    return cell["col"]


def right(cell):
    """One past the last column, so abutting blocks share the value."""
    return cell["col"] + cell["cols"]


def top(cell):
    return cell["row"]


def bottom(cell):
    return cell["row"] + cell["rows"]


def cells_overlap(a, b):
    return (left(a) < right(b) and right(a) > left(b)
            and top(a) < bottom(b) and bottom(a) > top(b))


def cells_to_rect(cell, cell_w, cell_h):
    return ((cell["col"] - 1) * cell_w, (cell["row"] - 1) * cell_h,
            cell["cols"] * cell_w, cell["rows"] * cell_h)


def rect_to_cells(rect, cell_w, cell_h):
    """A drawn rectangle back to cells, never rounding outwards.

    The guard can leave a trimmed block a fraction off the row grid; rounding
    its size up would hand the fidelity pass cells the guard just took away.
    """
    col = int(round(rect[0] / max(cell_w, 0.01))) + 1
    row = int(round(rect[1] / max(cell_h, 0.01))) + 1
    cols = int(math.floor(rect[2] / max(cell_w, 0.01) + 0.02))
    rows = int(math.floor(rect[3] / max(cell_h, 0.01) + 0.02))
    return cells(max(col, 1), max(row, 1), max(cols, MIN_CELLS), max(rows, MIN_CELLS))


def module_of(body, cell=CELL, columns=MODULE_COLS):
    """The style's column module in points, or None when it has no body slot.

    `module = (body.w - (columns - 1) * gutter) / columns` with the gutter one
    cell wide. `lefts` is the left edge of each column; `rights[k]` is the right
    edge of a block that starts on column k and is one module wide, so a block
    spanning n modules from column k ends on `rights[k + n - 1]`.
    """
    if not isinstance(body, dict):
        return None
    gutter = float(cell)
    columns = max(int(columns), 1)
    x = float(body.get("x") or 0.0)
    w = float(body.get("w") or 0.0)
    module = (w - gutter * (columns - 1)) / float(columns)
    if module <= 0.5:
        return None
    pitch = module + gutter
    lefts = [x + i * pitch for i in range(columns)]
    return {"x": x, "w": w, "cols": columns, "gutter": gutter,
            "module": module, "pitch": pitch,
            "lefts": lefts, "rights": [left + module for left in lefts]}


def nearest_edge(edges, value, tolerance, low=None, high=None):
    """The module edge this one should sit on, or None to leave it alone.

    None also means "already on the module": both answers are "do nothing".
    """
    best = None
    for edge in edges:
        if abs(edge - value) > tolerance + 0.001:
            continue
        if low is not None and edge < low - 0.001:
            continue
        if high is not None and edge > high + 0.001:
            continue
        if best is None or abs(edge - value) < abs(best - value) - 0.001:
            best = edge
    if best is None or abs(best - value) < 0.01:
        return None
    return best


def rect_to_band(rect, cell_w, cell_h):
    """A furniture rectangle to the cells it touches, rounded outwards.

    A band is an obstacle, so every cell it reaches into counts as taken.
    """
    c0 = int(math.floor(rect[0] / max(cell_w, 0.01) + 0.02)) + 1
    r0 = int(math.floor(rect[1] / max(cell_h, 0.01) + 0.02)) + 1
    c1 = int(math.ceil((rect[0] + rect[2]) / max(cell_w, 0.01) - 0.02))
    r1 = int(math.ceil((rect[1] + rect[3]) / max(cell_h, 0.01) - 0.02))
    return cells(max(c0, 1), max(r0, 1), max(c1 - c0 + 1, 1), max(r1 - r0 + 1, 1))


# -------------------------------------------------------------------- block


class Item(object):
    """One block in the fidelity pass: its cells and what it needs.

    `need_rows` / `need_cols` are the cells the content asks for beyond the
    slot at the role's floor size; `slack_rows` / `slack_cols` are the cells
    the block could give up and still hold its content. `repick` is the
    request build_pptx raises when content plainly does not suit the slot:
    `{"archetype": "table", "slot": "table", "reason": "..."}`.
    """

    def __init__(self, block, cell, need_rows=0, need_cols=0, slack_rows=0,
                 slack_cols=0, repick=None, movable=True):
        self.block = block or {}
        self.id = self.block.get("id")
        self.cells = copy_cells(cell)
        self.start = copy_cells(cell)
        self.need_rows = max(int(need_rows), 0)
        self.need_cols = max(int(need_cols), 0)
        self.slack_rows = max(int(slack_rows), 0)
        self.slack_cols = max(int(slack_cols), 0)
        self.repick = repick
        self.movable = bool(movable)
        # Set by the module pass when the module misses the cell grid: the
        # points to draw instead of the cells, and the cells they belong to.
        self.pt = None
        self.pt_cols = None

    def pt_rect(self):
        """(x, w) in points when the module snap left this block off the grid.

        A later pass that changes the block's columns takes the override with
        it: the cells are the record, the points are only the finer answer for
        the columns they were measured against.
        """
        if not self.pt:
            return None
        if self.pt_cols != (self.cells["col"], self.cells["cols"]):
            return None
        return self.pt

    def __repr__(self):  # pragma: no cover - debugging only
        return "Item(%s %s need=%d/%d slack=%d/%d)" % (
            self.id, self.cells, self.need_rows, self.need_cols,
            self.slack_rows, self.slack_cols)


# --------------------------------------------------------------------- pass


class Plan(object):
    """The fidelity pass over one slide."""

    def __init__(self, items, bands, grid_cols, grid_rows, fidelity, module=None,
                 cell_w=CELL, source=False, warn=None):
        self.items = [i for i in items]
        self.bands = list(bands or [])
        self.grid_cols = max(int(grid_cols), 1)
        self.grid_rows = max(int(grid_rows), 1)
        self.fidelity = fidelity
        self.module = module if isinstance(module, dict) else None
        self.cell_w = float(cell_w) or CELL
        self.source = bool(source)
        self.warn = warn if callable(warn) else (lambda message: None)
        self.reasons = {}        # id(item) -> why it moved, in order

    def body_bottom(self):
        """The last row a block may use: row 17, or 16 with a source line (F-3)."""
        bottom_row = SOURCE_ROW - 1 if self.source else BODY_BOTTOM
        return min(bottom_row, self.grid_rows)

    # -- space

    def in_grid(self, cell):
        return (cell["col"] >= 1 and cell["row"] >= 1
                and right(cell) <= self.grid_cols + 1
                and bottom(cell) <= self.grid_rows + 1)

    def blockers(self, cell, skip=()):
        """The blocks in the way. A block that draws nothing never blocks."""
        out = []
        for other in self.items:
            if other in skip or not other.movable:
                continue
            if cells_overlap(cell, other.cells):
                out.append(other)
        return out

    def hits_band(self, cell):
        return any(cells_overlap(cell, band) for band in self.bands)

    def free(self, cell, skip=()):
        if not self.in_grid(cell) or self.hits_band(cell):
            return False
        return not self.blockers(cell, skip=skip)

    def note(self, item, reason):
        """Record why a block moved. A block can move twice; both reasons stand."""
        seen = self.reasons.setdefault(id(item), [])
        if reason not in seen:
            seen.append(reason)

    # -- the pass

    def run(self):
        if self.fidelity == "strict":
            return []
        self.clear_source_row()
        if self.fidelity == "loose":
            self.free_stat_band()
            self.repick_slots()
        self.snap_module()
        self.grow_to_fit()
        self.rebalance()
        self.align_edges()
        return self.adjustments()

    def adjustments(self):
        out = []
        for item in self.items:
            snapped = item.pt_rect()
            if same_cells(item.start, item.cells) and not snapped:
                continue
            move = {
                "blockId": item.id,
                "from": copy_cells(item.start),
                "to": copy_cells(item.cells),
                "reason": "; ".join(self.reasons.get(id(item)) or []) or "adjusted to fit",
            }
            if snapped:
                move["snappedPt"] = {"x": round(snapped[0], 2), "w": round(snapped[1], 2)}
            out.append(move)
        return out

    # -- (0) snap every left edge and width to the style's column module

    def pt_span(self, item):
        """(x, w) in points of a block's current cells."""
        return ((item.cells["col"] - 1) * self.cell_w,
                item.cells["cols"] * self.cell_w)

    def pt_bounds(self, item):
        """(low, high) in points the block's edges may not cross.

        The midpoint of the gap to every neighbour and every band that shares
        its rows: each block stays on its own side of every midpoint, so no
        snap can put two blocks on top of each other.
        """
        low, high = 0.0, self.grid_cols * self.cell_w
        x0, w0 = self.pt_span(item)
        spans = []
        for other in self.items:
            if other is item or not other.movable:
                continue
            if not (top(item.cells) < bottom(other.cells)
                    and bottom(item.cells) > top(other.cells)):
                continue
            spans.append(self.pt_span(other))
        for band in self.bands:
            if not (top(item.cells) < bottom(band) and bottom(item.cells) > top(band)):
                continue
            spans.append(((band["col"] - 1) * self.cell_w, band["cols"] * self.cell_w))
        for ox, ow in spans:
            if ox + ow <= x0 + 0.001:
                low = max(low, (ox + ow + x0) / 2.0)
            elif ox >= x0 + w0 - 0.001:
                high = min(high, (x0 + w0 + ox) / 2.0)
        return (low, high)

    def snap_module(self):
        for item in sorted(self.items, key=lambda i: (top(i.cells), left(i.cells))):
            if item.movable:
                self.snap_to_module(item)

    def snap_to_module(self, item):
        """Left edge to the nearest module left, right edge to the nearest module right."""
        mod = self.module
        if not mod:
            return False
        x0, w0 = self.pt_span(item)
        low, high = self.pt_bounds(item)
        tolerance = MODULE_SNAP * self.cell_w
        x1 = nearest_edge(mod["lefts"], x0, tolerance, low, high)
        r1 = nearest_edge(mod["rights"], x0 + w0, tolerance, low, high)
        if x1 is None and r1 is None:
            return False
        x1 = x0 if x1 is None else x1
        r1 = x0 + w0 if r1 is None else r1
        if r1 - x1 < min(mod["module"], w0) - 0.5:
            return False
        cell = cells(int(round(x1 / self.cell_w)) + 1, item.cells["row"],
                     int(round((r1 - x1) / self.cell_w)), item.cells["rows"])
        if not self.in_grid(cell) or self.hits_band(cell):
            return False
        on_grid = (abs(x1 - (cell["col"] - 1) * self.cell_w) < 0.01
                   and abs((r1 - x1) - cell["cols"] * self.cell_w) < 0.01)
        item.cells = cell
        if on_grid:
            item.pt = None
            item.pt_cols = None
        else:
            item.pt = (x1, r1 - x1)
            item.pt_cols = (cell["col"], cell["cols"])
        self.note(item, MODULE_REASON)
        return True

    # -- (F-3) row 17 belongs to the source line

    def clear_source_row(self):
        """Lift the band that reaches row 17 when the slide carries a source.

        A source line and a block cannot share the last body row. The whole row
        group moves up together, so the columns that were level stay level; when
        there is nothing above to move into, the block is named in a warning
        rather than silently drawn under the source.
        """
        if not self.source:
            return
        over = [i for i in self.items
                if i.movable and bottom(i.cells) - 1 >= SOURCE_ROW]
        if not over:
            return
        group = [i for i in self.items if i.movable
                 and any(top(i.cells) < bottom(o.cells)
                         and bottom(i.cells) > top(o.cells) for o in over)]
        if self.shift_clear(group, -1):
            for item in group:
                self.note(item, SOURCE_REASON)
            return
        for item in over:
            self.warn("block '%s' reaches row %d, which the source line holds"
                      % (item.id or "?", SOURCE_ROW))

    def shift_clear(self, group, offset):
        """Move a whole group by `offset` rows, or refuse and change nothing."""
        trial = []
        for item in group:
            moved = copy_cells(item.cells)
            moved["row"] += offset
            if moved["row"] < 1 or not self.in_grid(moved) or self.hits_band(moved):
                return False
            if self.blockers(moved, skip=tuple(group)):
                return False
            trial.append((item, moved))
        if not trial:
            return False
        for item, moved in trial:
            item.cells = moved
        return True

    # -- (loose) free a stat band that a neighbouring rail boxed in

    def type_of(self, item):
        return str((item.block or {}).get("type") or "").strip().lower()

    def free_stat_band(self):
        """A half-width stat band beside a text rail takes the full band (Q21).

        Three figures squeezed into half a slide beside a bullet list are three
        figures the reader cannot see. The rail has no reason to be beside them:
        it drops below, and the stats take the full-width stat-band slot.
        """
        slot = ARCHETYPE_SLOTS["stat-band"]["stat"]
        for item in list(self.items):
            if not item.movable or self.type_of(item) != "stat":
                continue
            if item.cells["cols"] > slot["cols"] * STAT_BAND_SHARE:
                continue
            rails = [o for o in self.items
                     if o is not item and o.movable
                     and top(item.cells) < bottom(o.cells)
                     and bottom(item.cells) > top(o.cells)]
            if len(rails) != 1 or self.type_of(rails[0]) not in STAT_RAIL_TYPES:
                continue
            rail = rails[0]
            if self.lay_out_stat_band(item, rail, slot):
                self.note(item, STAT_BAND_REASON)
                self.note(rail, STAT_BAND_REASON)

    def lay_out_stat_band(self, item, rail, slot):
        """Stats across the band, the rail under them, or leave both alone."""
        low, high = self.free_band([item, rail])
        if low is None:
            return False
        top_row = max(low, min(top(item.cells), top(rail.cells)))
        rows = min(item.cells["rows"], slot["rows"], high - top_row + 1)
        if rows < MIN_CELLS:
            return False
        band = cells(slot["col"], top_row, slot["cols"], rows)
        below_rows = high - bottom(band) + 1
        if below_rows < MIN_CELLS:
            return False                 # no room under the band for the rail
        below = cells(slot["col"], bottom(band), slot["cols"], below_rows)
        if not self.in_grid(band) or not self.in_grid(below):
            return False
        if self.hits_band(band) or self.hits_band(below):
            return False
        skip = (item, rail)
        if self.blockers(band, skip=skip) or self.blockers(below, skip=skip):
            return False
        item.cells = band
        rail.cells = below
        item.need_cols = 0
        return True

    # -- (loose) re-pick the archetype slot

    def repick_slots(self):
        for item in self.items:
            request = item.repick
            if not request or not item.movable:
                continue
            slot = (ARCHETYPE_SLOTS.get(str(request.get("archetype")))
                    or {}).get(str(request.get("slot")))
            if not slot:
                continue
            want = cells(slot["col"], slot["row"], slot["cols"], slot["rows"])
            want = self.fit_into_band(want, item)
            if want is None or same_cells(want, item.cells):
                continue
            if not self.free(want, skip=(item,)):
                continue
            item.cells = want
            item.need_rows = 0
            item.need_cols = 0
            self.note(item, str(request.get("reason") or "re-picked the archetype slot"))

    def fit_into_band(self, want, item):
        """Pull a catalogue slot inside the free band, or give up on it."""
        low, high = self.free_band([item], want)
        if low is None:
            return None
        want = copy_cells(want)
        if want["rows"] > high - low + 1:
            want["rows"] = high - low + 1
        if want["row"] < low:
            want["row"] = low
        if bottom(want) - 1 > high:
            want["row"] = max(low, high - want["rows"] + 1)
        if want["rows"] < MIN_CELLS or not self.in_grid(want):
            return None
        return want

    # -- (a) grow a block whose content does not fit

    def grow_to_fit(self):
        for item in self.items:
            if not item.movable:
                continue
            if item.need_rows > 0:
                want = min(item.need_rows, MAX_GROW)
                got = self.grow_rows(item, want)
                if got:
                    self.note(item, "grown %d row%s to fit its content"
                              % (got, "" if got == 1 else "s"))
            if item.need_cols > 0:
                want = min(item.need_cols, MAX_GROW)
                got = self.grow_cols(item, want)
                if got:
                    self.note(item, "grown %d column%s to fit its content"
                              % (got, "" if got == 1 else "s"))

    def grow_rows(self, item, want):
        for step in range(want, 0, -1):
            for direction in ("down", "up"):
                if self.try_grow(item, step, direction):
                    return step
        return 0

    def grow_cols(self, item, want):
        for step in range(want, 0, -1):
            for direction in ("right", "left"):
                if self.try_grow(item, step, direction):
                    return step
        return 0

    def try_grow(self, item, step, direction):
        grown = copy_cells(item.cells)
        if direction == "down":
            grown["rows"] += step
        elif direction == "up":
            grown["row"] -= step
            grown["rows"] += step
        elif direction == "right":
            grown["cols"] += step
        else:
            grown["col"] -= step
            grown["cols"] += step
        if not self.in_grid(grown) or self.hits_band(grown):
            return False
        # Growing down never takes the row the source line needs (F-3).
        if (self.source and bottom(grown) - 1 >= SOURCE_ROW
                and bottom(grown) > bottom(item.cells)):
            return False
        hit = self.blockers(grown, skip=(item,))
        if not hit:
            item.cells = grown
            return True
        if len(hit) != 1:
            return False
        neighbour = hit[0]
        if not self.shrink(neighbour, step, direction):
            return False
        item.cells = grown
        return True

    def shrink(self, neighbour, step, direction):
        """Take `step` cells off a neighbour that has the slack, or refuse."""
        if not neighbour.movable:
            return False
        shrunk = copy_cells(neighbour.cells)
        if direction in ("down", "up"):
            if neighbour.slack_rows < step or shrunk["rows"] - step < MIN_CELLS:
                return False
            if direction == "down":
                shrunk["row"] += step         # the neighbour below loses its top
                shrunk["rows"] -= step
            else:
                shrunk["rows"] -= step        # the neighbour above loses its foot
        else:
            if neighbour.slack_cols < step or shrunk["cols"] - step < MIN_CELLS:
                return False
            if direction == "right":
                shrunk["col"] += step         # the neighbour right loses its left
                shrunk["cols"] -= step
            else:
                shrunk["cols"] -= step
        if not self.in_grid(shrunk) or self.hits_band(shrunk):
            return False
        if self.blockers(shrunk, skip=(neighbour,)):
            return False
        neighbour.cells = shrunk
        neighbour.slack_rows = max(neighbour.slack_rows - step, 0) \
            if direction in ("down", "up") else neighbour.slack_rows
        neighbour.slack_cols = max(neighbour.slack_cols - step, 0) \
            if direction in ("right", "left") else neighbour.slack_cols
        self.note(neighbour, "shrunk %d cell%s to make room"
                  % (step, "" if step == 1 else "s"))
        return True

    # -- (b) rebalance the whitespace under the last block

    def free_band(self, group, probe=None):
        """(first, last) free row for this group inside the body area."""
        low, high = BODY_TOP, self.body_bottom()
        spans = [i.cells for i in group] + ([probe] if probe else [])
        for band in self.bands:
            if not any(left(c) < right(band) and right(c) > left(band) for c in spans):
                continue
            if bottom(band) - 1 < low or top(band) > high:
                continue
            if top(band) <= low:
                low = max(low, bottom(band))
            elif bottom(band) - 1 >= high:
                high = min(high, top(band) - 1)
            else:
                # a band across the middle: keep the deeper side
                above = top(band) - low
                below = high - (bottom(band) - 1)
                if below >= above:
                    low = bottom(band)
                else:
                    high = top(band) - 1
        if high < low:
            return (None, None)
        return (low, high)

    def rebalance(self):
        live = [i for i in self.items if i.movable]
        if not live:
            return
        low, high = self.free_band(live)
        if low is None:
            return
        first = min(top(i.cells) for i in live)
        last = max(bottom(i.cells) for i in live) - 1
        if first < low or last > high:
            return                       # the blocks are not inside the body band
        slack = high - last
        if slack <= EMPTY_BAND:
            return
        groups = self.row_groups(live)
        if len(groups) == 1:
            offset = low + (high - low + 1 - (last - first + 1)) // 2 - first
            if offset == 0:
                return
            if not self.shift(groups[0], offset):
                return
            for item in groups[0]:
                self.note(item, "centred in the body band")
            return
        extra = (high - first + 1) - (last - first + 1)
        heights = [self.group_height(group) for group in groups]
        total = float(sum(heights)) or 1.0
        offset = 0
        given = 0
        for index, group in enumerate(groups):
            share = int(math.floor(extra * heights[index] / total))
            if index == len(groups) - 1:
                share = extra - given
            given += share
            for item in group:
                if offset:
                    item.cells["row"] += offset
                if share:
                    item.cells["rows"] += share
                if offset or share:
                    self.note(item, "spread to fill the body band")
            offset += share

    @staticmethod
    def group_height(group):
        return (max(bottom(i.cells) for i in group)
                - min(top(i.cells) for i in group))

    @staticmethod
    def row_groups(items):
        """Blocks clustered into the rows they share, top to bottom."""
        out = []
        for item in sorted(items, key=lambda i: (top(i.cells), left(i.cells))):
            placed = False
            for group in out:
                if any(top(item.cells) < bottom(o.cells)
                       and bottom(item.cells) > top(o.cells) for o in group):
                    group.append(item)
                    placed = True
                    break
            if not placed:
                out.append([item])
        return out

    def shift(self, group, offset):
        trial = []
        for item in group:
            moved = copy_cells(item.cells)
            moved["row"] += offset
            if not self.in_grid(moved) or self.hits_band(moved):
                return False
            trial.append((item, moved))
        for item, moved in trial:
            item.cells = moved
        return True

    # -- (c) snap edges that nearly line up

    EDGES = ("col", "right", "row", "bottom")

    def align_edges(self):
        for edge in self.EDGES:
            self.align_one(edge)

    def edge_of(self, item, edge):
        cell = item.cells
        if edge == "col":
            return left(cell)
        if edge == "right":
            return right(cell)
        if edge == "row":
            return top(cell)
        return bottom(cell)

    def align_one(self, edge):
        live = [i for i in self.items if i.movable]
        if len(live) < 2:
            return
        values = {}
        for item in live:
            values.setdefault(self.edge_of(item, edge), []).append(item)
        for cluster in self.clusters(sorted(values)):
            if len(cluster) < 2:
                continue
            members = [i for value in cluster for i in values[value]]
            if len({self.edge_of(i, edge) for i in members}) < 2:
                continue
            target = self.target_edge(edge, cluster, values)
            for item in members:
                if self.edge_of(item, edge) == target:
                    continue
                if self.snap(item, edge, target):
                    self.note(item, "edge aligned to %s %d" % (self.EDGE_NAMES[edge],
                                                               target))

    EDGE_NAMES = {"col": "column", "right": "column", "row": "row", "bottom": "row"}

    @staticmethod
    def clusters(sorted_values):
        out = []
        for value in sorted_values:
            if out and value - out[-1][-1] <= ALIGN_SNAP:
                out[-1].append(value)
            else:
                out.append([value])
        return out

    @staticmethod
    def target_edge(edge, cluster, values):
        counts = [(len(values[v]), v) for v in cluster]
        best = max(c for c, _ in counts)
        tied = sorted(v for c, v in counts if c == best)
        if edge in ("col", "row"):
            return tied[0]
        return tied[-1]

    def snap(self, item, edge, target):
        moved = copy_cells(item.cells)
        if edge == "col":
            moved["cols"] += moved["col"] - target
            moved["col"] = target
        elif edge == "right":
            moved["cols"] = target - moved["col"]
        elif edge == "row":
            moved["rows"] += moved["row"] - target
            moved["row"] = target
        else:
            moved["rows"] = target - moved["row"]
        if moved["cols"] < MIN_CELLS or moved["rows"] < MIN_CELLS:
            return False
        if not self.free(moved, skip=(item,)):
            return False
        item.cells = moved
        return True


def plan_slide(items, bands, grid_cols, grid_rows, fidelity, module=None, cell_w=CELL,
               source=False, warn=None):
    """Run the fidelity pass. Mutates each Item's cells; returns the deviations.

    `source` says the slide carries a source line, so row 17 is spoken for
    (F-3); `warn` takes the messages about blocks the pass could not settle.
    """
    return Plan(items, bands, grid_cols, grid_rows, fidelity, module=module,
                cell_w=cell_w, source=source, warn=warn).run()


# ------------------------------------------------------- deck-wide type


def size_ladder(spec, top=None):
    """Every size a role may take, largest first: the range walked by the fit rule.

    The ladder starts at the size the style asked for, not at the ceiling of the
    range: a deck-wide size only ever steps a role *down* from the look its
    tokens describe, so an extracted style keeps its own type. The step is the
    one in docs/research/deck-planning.md section 7: 2 pt down to 14 pt, 1 pt
    below it, and never past the role's floor.
    """
    top_size = float(top if top is not None else spec.get("size", spec.get("max", 12)))
    floor = float(spec.get("min", 9))
    if top_size < floor:
        top_size = floor
    out = [top_size]
    size = top_size
    guard = 0
    while size > floor and guard < 80:
        guard += 1
        step = 2.0 if size > 14 else 1.0
        size = max(floor, size - step)
        out.append(size)
    return out

import assert from "node:assert/strict";
import { cornerResize, resolveOverlaps, stackBlocks, flowBlocks, blockAbove, mergeBlocks, splitBlock, alignLeft, alignTop, sameWidth, sameHeight, distributeVertically, distributeHorizontally, colKind, fitCols, formatSample, headersFit, labelPct, valueCell } from "../src/app/layout-view.js";
import { MODULE_COLS, balanceBlocks, columnModule, contentLines, diagramId, diagramText, fitBlocks, fitRows, fitToSlide, parseDiagram, resolveDeckOverlaps, rowBands, snapToModule } from "../src/app/deck-layout.js";
import { ARCHETYPES, CYCLE, detectArchetype, nextArchetype, remix, type SlotRole } from "../src/app/remix.js";
import { resolveBandOverlaps } from "../src/shared/archetypes.js";
import { DECK_COLS, DECK_ROWS, type BlockType } from "../src/shared/types.js";
import { fakeValue, type Block } from "../src/app/model.js";

const mk = (id: string, row: number, col: number, rows: number, cols: number): Block => ({ id, type: "table", title: id, row, col, rows, cols });
const pick = (b: Block) => [b.row, b.col, b.rows, b.cols];

// moved block dropped onto another: the other moves below it
{
  const a = mk("a", 2, 2, 6, 5), b = mk("b", 15, 2, 4, 3);
  b.row = 3; // dropped onto a
  const changed = resolveOverlaps([a, b], b);
  assert.ok(changed); assert.equal(b.row, 3, "moved block keeps its spot"); assert.equal(a.row, 3 + 4 + 1, "other block pushed below");
}
// chain: pushing a onto c pushes c too
{
  const a = mk("a", 2, 2, 6, 5), b = mk("b", 20, 2, 4, 5), c = mk("c", 9, 2, 5, 5);
  b.row = 2; resolveOverlaps([a, b, c], b);
  assert.equal(a.row, 7); assert.equal(c.row, 14);
}
// no overlap: nothing moves
{
  const a = mk("a", 2, 2, 6, 5), b = mk("b", 2, 9, 4, 3);
  assert.equal(resolveOverlaps([a, b], b), false); assert.equal(a.row, 2);
}
// stack and flow
{
  const bs = [mk("a", 9, 4, 3, 4), mk("b", 2, 2, 5, 6), mk("c", 30, 1, 4, 9)];
  stackBlocks(bs); assert.deepEqual(bs.map((b) => [b.row, b.col]), [[8, 2], [2, 2], [12, 2]]);
  flowBlocks(bs); assert.deepEqual(bs.map((b) => [b.row, b.col]), [[2, 9], [2, 2], [8, 2]], "c wraps to a new band because 9 cols do not fit after a");
}
// merge: a bare title block above a table donates its title
{
  const title: Block = { id: "t", type: "text", title: "Revenue by band", row: 2, col: 2, rows: 1, cols: 4 };
  const table = mk("tbl", 4, 2, 6, 4); table.columns = ["Band", "Firms"];
  const bs = [title, table];
  assert.equal(blockAbove(bs, table), title);
  const m = mergeBlocks(bs, title, table);
  assert.equal(bs.length, 1); assert.equal(m.title, "Revenue by band"); assert.equal(m.type, "table"); assert.equal(m.row, 2); assert.equal(m.rows, 8); assert.deepEqual(m.columns, ["Band", "Firms"]);
}
// split: a 10-row table (8 data rows) at data row 3 gives 3 + 5 data rows
{
  const t = mk("t", 2, 2, 10, 3); t.rowLabels = ["a", "b", "c", "d", "e", "f", "g", "h"];
  const bs = [t];
  const second = splitBlock(bs, t, 3)!;
  assert.equal(t.rows, 5); assert.deepEqual(t.rowLabels, ["a", "b", "c"]);
  assert.equal(second.rows, 7); assert.deepEqual(second.rowLabels, ["d", "e", "f", "g", "h"]); assert.equal(second.row, t.row + t.rows + 1);
  assert.equal(splitBlock(bs, t, 9), undefined, "cannot split past the data area");
}
// group drop: the moved group stays put, the block under it goes below the lower member, members never push each other
{
  const a = mk("a", 2, 2, 4, 3), b = mk("b", 7, 2, 5, 3), c = mk("c", 20, 2, 6, 3);
  a.row = 18; b.row = 23; // dropped together onto c
  assert.ok(resolveOverlaps([a, b, c], [a, b]));
  assert.equal(a.row, 18); assert.equal(b.row, 23); assert.equal(c.row, 23 + 5 + 1);
}
// align left and top: the group's leftmost column and topmost row win
{
  const bs = [mk("a", 2, 5, 3, 3), mk("b", 6, 2, 3, 3), mk("c", 10, 8, 3, 3)];
  alignLeft(bs); assert.deepEqual(bs.map((b) => b.col), [2, 2, 2]);
  alignTop(bs); assert.deepEqual(bs.map((b) => b.row), [2, 2, 2]);
  alignLeft([]); alignTop([]); // no-ops
}
// same width and height take the reference block's size (the primary), default the first
{
  const bs = [mk("a", 2, 2, 3, 3), mk("b", 6, 2, 5, 7), mk("c", 12, 2, 2, 4)];
  sameWidth(bs, bs[1]); assert.deepEqual(bs.map((b) => b.cols), [7, 7, 7]);
  sameHeight(bs); assert.deepEqual(bs.map((b) => b.rows), [3, 3, 3]);
}
// distribute vertically: order kept, first row kept, equal blank rows between (11 free rows over two gaps = 6 each, rounded)
{
  const a = mk("a", 2, 2, 4, 3), b = mk("b", 8, 2, 3, 3), c = mk("c", 20, 2, 5, 3);
  distributeVertically([c, a, b]);
  assert.deepEqual([a.row, b.row, c.row], [2, 12, 21]);
  assert.equal(b.row - (a.row + a.rows), c.row - (b.row + b.rows), "gaps are equal");
  // packed or overlapping blocks still get one blank row between
  const d = mk("d", 2, 2, 4, 3), e = mk("e", 3, 2, 4, 3), f = mk("f", 4, 2, 4, 3);
  distributeVertically([d, e, f]); assert.deepEqual([d.row, e.row, f.row], [2, 7, 12]);
  // two blocks: the outer edges already define the gap, nothing moves
  const g = mk("g", 2, 2, 4, 3), hh = mk("h", 10, 2, 4, 3);
  distributeVertically([g, hh]); assert.deepEqual([g.row, hh.row], [2, 10]);
}
// distribute horizontally: same on columns (11 free over two gaps = 6, rounded)
{
  const a = mk("a", 2, 2, 3, 3), b = mk("b", 2, 6, 3, 4), c = mk("c", 2, 20, 3, 3);
  distributeHorizontally([b, c, a]);
  assert.deepEqual([a.col, b.col, c.col], [2, 11, 21]);
}
// ---- U2: a placeholder value never reads as data, a real sample takes its column's format ----

// the generator's own output is a placeholder, whoever put it in the plan
{
  assert.deepEqual(valueCell("Firms", 0, undefined), { text: fakeValue("Firms", 0), ph: true }, "no sample: the guess, marked");
  assert.deepEqual(valueCell("Firms", 1, ""), { text: fakeValue("Firms", 1), ph: true }, "an empty sample is no sample");
  assert.deepEqual(valueCell("Revenue", 2, fakeValue("Revenue", 2)), { text: fakeValue("Revenue", 2), ph: true }, "a sample equal to the guess is still a guess");
  assert.deepEqual(valueCell("Businesses", 0, 32000000), { text: "32,000,000", ph: false }, "a real sample, in the column's format");
  assert.deepEqual(valueCell("Segment", 0, "SMB"), { text: "SMB", ph: false });
  // the generator's guess for row 3 is not the guess for row 0: the rule is per row
  assert.equal(valueCell("Firms", 3, fakeValue("Firms", 0)).ph, false);
}
// number formats by column header, the way the builder writes them
{
  assert.equal(colKind("Revenue ($M)"), "money");
  assert.equal(colKind("Share"), "percent");
  assert.equal(colKind("Multiple"), "multiple");
  assert.equal(colKind("Year"), "year");
  assert.equal(colKind("Businesses"), "count");
  assert.equal(formatSample("Revenue ($M)", 2034), "$2,034");
  assert.equal(formatSample("Price", 12.5), "$12.50");
  assert.equal(formatSample("Share", 0.2), "20%", "a fraction reads as a percent, as the builder's percent format does");
  assert.equal(formatSample("Share", 20), "20%");
  assert.equal(formatSample("Growth", 0.125), "12.5%");
  assert.equal(formatSample("Multiple", 2.4), "2.4x");
  assert.equal(formatSample("Year", 2026), "2026", "a year keeps no thousands separator");
  assert.equal(formatSample("Year 1", 375000), "375,000", "a Year N column is unsigned: the sample may be a count");
  assert.equal(formatSample("Stage", "All businesses"), "All businesses", "a string is the user's own text");
}

// ---- U3 and U4: a block as wide as its columns, headers one step down before they truncate ----

{
  assert.equal(fitCols(["Item", "Year 1", "Year 2", "Year 3"]), 4, "one cell per column");
  assert.equal(fitCols(["Segment", "Firms", "Spend"]), 3);
  assert.equal(fitCols(["Item", "Value"]), 3, "a label column plus two at the least");
  assert.equal(fitCols(undefined), 3);
  assert.equal(fitCols(Array.from({ length: 40 }, (_, i) => `c${i}`)), 25, "capped at the sheet's right edge");
  assert.equal(labelPct("inputs"), 60);
  assert.equal(labelPct("table", ["Line", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5"]), 29, "the label column narrows as columns are added");
  assert.equal(labelPct("table", ["Segment", "Firms"]), 46, "clamped");
}
{
  const years = ["Line", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5"];
  assert.ok(headersFit(years, fitCols(years)), "six columns at their own width fit at the cell size");
  assert.ok(!headersFit(years, 5), "the same six squeezed into five cells do not");
  assert.ok(headersFit(years, 5, 9), "and do fit one step down the type scale");
  assert.ok(!headersFit(years, 4) && !headersFit(years, 4, 9), "four cells truncate even after the step");
  assert.ok(!headersFit(["Stage", "Businesses", "Penetration"], 3), "long headers in three cells need the step");
  assert.ok(headersFit([], 2), "no columns, nothing to fit");
}

// ---- Remix: the archetype catalogue (docs/research/deck-planning.md section 5) ----

const dk = (id: string, type: BlockType, geo: Partial<Block> = {}): Block => ({ id, type, title: id, row: 7, col: 3, rows: 4, cols: 8, ...geo });
/** A block type whose first-choice slot role is this one, so a round-trip set fills every slot exactly. */
const OF_ROLE: Record<SlotRole, BlockType> = { text: "text", visual: "chart", card: "cards", stat: "stat", list: "bullets", table: "table", quote: "quote", statement: "takeaway", caption: "notes" };
const onSlide = (bs: Block[], where: string) => {
  for (const b of bs) {
    assert.ok(b.col >= 1 && b.col + b.cols - 1 <= DECK_COLS, `${where}: ${b.id} columns ${b.col}+${b.cols} off the slide`);
    assert.ok(b.row >= 1 && b.row + b.rows - 1 <= DECK_ROWS, `${where}: ${b.id} rows ${b.row}+${b.rows} off the slide`);
  }
};

// 1. every archetype round-trips its own slots: blocks laid out on the slots land back on them, no placeholders
for (const a of ARCHETYPES) {
  const blocks = a.slots.map((s, i) => dk(`b${i}`, OF_ROLE[s.role], { row: s.row, col: s.col, rows: s.rows, cols: s.cols }));
  const out = remix(blocks, a.name);
  assert.equal(out.archetype, a.name);
  assert.equal(out.blocks.length, a.slots.length, `${a.name}: one block per slot, nothing added`);
  a.slots.forEach((s, i) => {
    const b = out.blocks.find((x) => x.id === `b${i}`)!;
    assert.deepEqual([b.col, b.row, b.cols, b.rows], [s.col, s.row, s.cols, s.rows], `${a.name} slot ${s.name}`);
  });
  onSlide(out.blocks, a.name);
}

// 2. more blocks than slots: extras keep their ids and stack below the body, never off the slide
{
  const blocks = [dk("t1", "text"), dk("t2", "text"), dk("t3", "text"), dk("c1", "chart"), dk("t4", "text")];
  const out = remix(blocks, "text-visual");
  const at = (id: string) => out.blocks.find((b) => b.id === id)!;
  assert.deepEqual([at("t1").col, at("t1").row, at("t1").cols, at("t1").rows], [3, 7, 12, 10], "first text takes the text slot");
  assert.deepEqual([at("c1").col, at("c1").row, at("c1").cols, at("c1").rows], [17, 7, 14, 10], "the chart takes the visual slot");
  for (const id of ["t2", "t3", "t4"]) assert.equal(at(id).row, 17, `${id} stacks in the band below the body`);
  assert.deepEqual([at("t2").col, at("t3").col, at("t4").col], [3, 12, 21], "extras share the band instead of piling up");
  onSlide(out.blocks, "extras");
  assert.equal(out.blocks.length, 5, "nothing dropped, nothing invented");
}

// 3. fewer blocks than slots: the empty slot becomes a placeholder titled by its role
{
  const out = remix([dk("t1", "text"), dk("t2", "text")], "three-cards");
  const ph = out.blocks.filter((b) => b.type === "placeholder");
  assert.equal(ph.length, 1);
  assert.equal(ph[0].title, "Card 3");
  assert.deepEqual([ph[0].col, ph[0].row, ph[0].cols, ph[0].rows], [23, 7, 8, 10]);
  assert.deepEqual(out.blocks.filter((b) => b.type === "text").map((b) => b.col), [3, 13], "text fills a card slot");
}

// 4. role mapping: chart to the visual slot, text to the text slot, bullets has nowhere in chart-insight
{
  const out = remix([dk("ch", "chart"), dk("tx", "text"), dk("bu", "bullets")], "chart-insight");
  const at = (id: string) => out.blocks.find((b) => b.id === id)!;
  assert.deepEqual([at("ch").col, at("ch").cols], [3, 18], "chart in the chart slot");
  assert.deepEqual([at("tx").col, at("tx").cols], [23, 8], "text in the insight rail");
  assert.equal(at("bu").row, 19 - 2, "bullets stacks below");
  assert.equal(out.blocks.find((b) => b.type === "placeholder")!.title, "Caption", "the source slot is a placeholder");
}

// 5 and 6. ids preserved and nothing dropped, whichever archetype
{
  const blocks = [dk("a", "stat"), dk("b", "cards", { count: 4 }), dk("c", "quote"), dk("d", "table"), dk("e", "notes")];
  for (const a of ARCHETYPES) {
    const out = remix(blocks, a.name);
    for (const b of blocks) assert.ok(out.blocks.some((x) => x.id === b.id), `${a.name} dropped ${b.id}`);
    assert.equal(new Set(out.blocks.map((b) => b.id)).size, out.blocks.length, `${a.name}: duplicate ids`);
    onSlide(out.blocks, a.name);
  }
}

// 7. cycling walks the content archetypes and never lands on one that owns its furniture
{
  assert.ok(CYCLE.every((a) => !a.special));
  assert.equal(nextArchetype(undefined), "three-cards");
  assert.equal(nextArchetype("three-cards"), "two-by-two");
  assert.equal(nextArchetype("matrix"), "three-cards", "the cycle wraps");
  assert.equal(nextArchetype("cover"), "three-cards", "a special archetype leaves the cycle at the start");
  const once = remix([dk("t", "text"), dk("c", "chart")]);
  assert.equal(once.archetype, "three-cards", "a hand-arranged slide starts the cycle");
  assert.equal(remix(once.blocks).archetype, "two-by-two", "the next call steps on");
}

// 8. a named archetype wins over the cycle
assert.equal(remix([dk("q", "quote")], "quote").archetype, "quote");

// 9. one cards block spans the card band and keeps its count; without a count it takes the slot count
{
  const three = remix([dk("cards", "cards", { count: 3 })], "three-cards");
  const b = three.blocks[0];
  assert.deepEqual([b.col, b.row, b.cols, b.rows], [3, 7, 28, 10], "one cards block fills the band");
  assert.equal(b.count, 3, "count preserved");
  assert.equal(three.blocks.length, 1, "spanned card slots do not leave placeholders");
  const four = remix([dk("cards", "cards")], "two-by-two");
  assert.deepEqual([four.blocks[0].col, four.blocks[0].row, four.blocks[0].cols, four.blocks[0].rows], [3, 7, 28, 9]);
  assert.equal(four.blocks[0].count, 4, "no count yet: take the archetype's");
}

// 10. furniture never moves (rule 8)
{
  const head = dk("h", "headline", { row: 3, col: 3, rows: 3, cols: 28 });
  const kick = dk("k", "kicker", { row: 2, col: 3, rows: 1, cols: 16 });
  const out = remix([head, kick, dk("t", "text")], "quote");
  const at = (id: string) => out.blocks.find((b) => b.id === id)!;
  assert.deepEqual([at("h").row, at("h").col, at("h").rows, at("h").cols], [3, 3, 3, 28]);
  assert.deepEqual([at("k").row, at("k").col, at("k").rows, at("k").cols], [2, 3, 1, 16]);
}

// 11. emphasis wins the first slot of its role (rule 7)
{
  const out = remix([dk("plain", "text"), dk("pop", "text", { emphasis: true })], "two-column");
  assert.equal(out.blocks.find((b) => b.col === 3)!.id, "pop");
  assert.equal(out.blocks.find((b) => b.col === 18)!.id, "plain");
}

// 12. pure and deterministic: the input is untouched and two runs agree
{
  const blocks = [dk("a", "text"), dk("b", "chart"), dk("c", "stat")];
  const before = JSON.stringify(blocks);
  const one = remix(blocks, "stat-band");
  const two = remix(blocks, "stat-band");
  assert.equal(JSON.stringify(blocks), before, "inputs mutated");
  assert.deepEqual(one.blocks, two.blocks, "not deterministic");
}

// 13. a repeated row re-solves for the count (rules 2 and 3): four cards grid, five wrap 3 over 2
{
  const four = remix([dk("1", "cards"), dk("2", "cards"), dk("3", "cards"), dk("4", "cards")], "three-cards");
  assert.deepEqual(four.blocks.map((b) => [b.col, b.row, b.cols, b.rows]), [[3, 7, 13, 4], [18, 7, 13, 4], [3, 12, 13, 4], [18, 12, 13, 4]]);
  const five = remix([dk("1", "cards"), dk("2", "cards"), dk("3", "cards"), dk("4", "cards"), dk("5", "cards")], "three-cards");
  assert.deepEqual(five.blocks.map((b) => [b.col, b.row]), [[3, 7], [13, 7], [23, 7], [8, 12], [18, 12]], "second row centred");
  assert.equal(five.blocks.filter((b) => b.type === "placeholder").length, 0);
  const stats = remix([dk("1", "stat"), dk("2", "stat"), dk("3", "stat"), dk("4", "stat")], "stat-band");
  assert.deepEqual(stats.blocks.filter((b) => b.type === "stat").map((b) => [b.col, b.cols]), [[2, 6], [10, 6], [18, 6], [26, 6]]);
}

// 14. three-cards with a takeaway bar: the cards shorten and the bar takes row 16 (doc, three-cards)
{
  const out = remix([dk("c1", "cards"), dk("c2", "cards"), dk("c3", "cards"), dk("bar", "takeaway")], "three-cards");
  assert.deepEqual(out.blocks.filter((b) => b.type === "cards").map((b) => [b.col, b.row, b.cols, b.rows]), [[3, 7, 8, 8], [13, 7, 8, 8], [23, 7, 8, 8]]);
  const bar = out.blocks.find((b) => b.id === "bar")!;
  assert.deepEqual([bar.col, bar.row, bar.cols, bar.rows], [3, 16, 28, 2]);
  assert.equal(detectArchetype(out.blocks), "three-cards", "the takeaway variant still reads as three-cards");
}

// 15. detection: a remixed slide reports its archetype, a hand-arranged one reports none
{
  for (const a of ARCHETYPES) {
    const blocks = a.slots.map((s, i) => dk(`b${i}`, OF_ROLE[s.role], { row: s.row, col: s.col, rows: s.rows, cols: s.cols }));
    assert.equal(detectArchetype(remix(blocks, a.name).blocks), a.name, `${a.name} not detected`);
  }
  assert.equal(detectArchetype([dk("t", "text", { row: 9, col: 5, rows: 3, cols: 7 })]), undefined);
  assert.equal(detectArchetype([]), undefined);
  assert.equal(detectArchetype(remix([dk("cards", "cards", { count: 3 })], "three-cards").blocks), "three-cards", "the spanned card band still reads as three-cards");
}

// ---- F19: corner resize maths ----

// the south and east edges change the size only
{
  const start = { row: 6, col: 4, rows: 5, cols: 8 };
  assert.deepEqual(cornerResize(start, "se", 2, 3), { row: 6, col: 4, rows: 7, cols: 11 });
  assert.deepEqual(cornerResize(start, "se", -9, -9), { row: 6, col: 4, rows: 1, cols: 1 }, "never smaller than one cell");
  // the deck's walls: the body band bottom and the slide's right edge
  assert.deepEqual(cornerResize(start, "se", 40, 40, { top: 5, bottom: 17, left: 1, right: 32 }), { row: 6, col: 4, rows: 12, cols: 29 });
}
// the north and west edges move the block's cell as they resize it
{
  const start = { row: 6, col: 4, rows: 5, cols: 8 };
  assert.deepEqual(cornerResize(start, "nw", 2, 1), { row: 8, col: 5, rows: 3, cols: 7 }, "dragging in shrinks from the top left");
  assert.deepEqual(cornerResize(start, "nw", -3, -2), { row: 3, col: 2, rows: 8, cols: 10 }, "dragging out grows up and left");
  assert.deepEqual(cornerResize(start, "nw", -9, -9, { top: 5, left: 1 }), { row: 5, col: 1, rows: 6, cols: 11 }, "stops at the band and the edge");
  assert.deepEqual(cornerResize(start, "nw", 9, 9), { row: 10, col: 11, rows: 1, cols: 1 }, "never flips through itself");
  assert.deepEqual(cornerResize(start, "ne", 1, 2), { row: 7, col: 4, rows: 4, cols: 10 }, "north edge moves, east edge grows");
  assert.deepEqual(cornerResize(start, "sw", 1, 2), { row: 6, col: 6, rows: 6, cols: 6 }, "south edge grows, west edge moves");
}

// ---- F11: the furniture band and the footer row are walls ----
{
  assert.deepEqual(pick(fitToSlide(mk("a", 1, 3, 4, 10))), [5, 3, 4, 10], "a block in the band snaps to row 5");
  assert.deepEqual(pick(fitToSlide(mk("b", 16, 3, 4, 10))), [14, 3, 4, 10], "a block over the footer row is pulled up");
  assert.deepEqual(pick(fitToSlide(mk("c", 5, 3, 30, 40))), [5, 1, 13, 32], "no taller than the band and no wider than the slide");
  assert.deepEqual(pick(fitToSlide(mk("d", 7, 3, 6, 8))), [7, 3, 6, 8], "a block already in the band is left alone");
  // a block with nowhere to go either way shrinks onto the free rows rather than sliding into the footer
  const a = mk("a", 5, 1, 7, 32), b = mk("b", 6, 1, 10, 32);
  assert.equal(resolveDeckOverlaps([a, b]), true);
  assert.deepEqual(pick(b), [12, 1, 6, 32], "no room below or right: it takes what the band has left");
  assert.equal(a.row, 5, "the earlier block never moves");
}

/**
 * F-2: a drop on the deck canvas resolves the way the server will, so the preview and the saved plan
 * agree. The rule is reading order, not who was dragged: the later block gives way, even when the later
 * block is the one the hand moved.
 */
{
  // the dragged block lands below the one already there, so the dragged block is the one that moves on
  const a = mk("a", 5, 3, 5, 14), b = mk("b", 13, 3, 4, 14);
  b.row = 7;                                        // dropped onto a
  assert.equal(resolveDeckOverlaps([a, b]), true);
  assert.deepEqual([a.row, b.row], [5, 10], "the later block in reading order gives way, not the earlier one");
  // the same drop through the server's own pass, cell for cell
  const server = [mk("a", 5, 3, 5, 14), mk("b", 7, 3, 4, 14)];
  resolveBandOverlaps(server);
  assert.deepEqual(server.map(pick), [a, b].map(pick), "canvas and server land on the same slide");

  // dragged the other way: the block that lands higher now reads first, so the one it hit moves down
  const x = mk("x", 9, 3, 4, 14), y = mk("y", 6, 3, 5, 14);
  x.row = 5;                                        // dropped onto y
  assert.equal(resolveDeckOverlaps([x, y]), true);
  assert.deepEqual([x.row, y.row], [5, 9], "the dragged block keeps the cell it was dropped on");
  const also = [mk("x", 5, 3, 4, 14), mk("y", 6, 3, 5, 14)];
  resolveBandOverlaps(also);
  assert.deepEqual(also.map(pick), [x, y].map(pick));

  // nothing overlapping: nothing moves, and nothing is reported
  const p = mk("p", 5, 3, 4, 14), q = mk("q", 10, 3, 4, 14);
  assert.equal(resolveDeckOverlaps([p, q]), false);
  assert.deepEqual([p.row, q.row], [5, 10]);
}

// ---- ranked change 1: the column module the body area implies ----
{
  const house = columnModule(undefined);
  assert.equal(house.starts.length, MODULE_COLS);
  assert.equal(house.gutter, 1);
  // the grid's own body is two cells of margin either side: 28 cells, so a module is 42.5 pt
  assert.equal(Math.round(house.unit * 30 * 10) / 10, 42.5);
  assert.deepEqual(house.starts, [3, 5, 8, 10, 13, 15, 18, 20, 22, 25, 27, 30]);
  assert.deepEqual(house.ends, [3, 6, 8, 11, 13, 16, 18, 20, 23, 25, 28, 30]);
  assert.equal(house.starts[0], 3, "the first module starts where the furniture does");
  assert.equal(house.ends[MODULE_COLS - 1], 30, "the last one ends at the body's right edge");
  // a style with its own body area moves every edge: harbor's margin is 32 pt, not 60
  const harbor = columnModule({ furniture: { body: { x: 32, y: 197, w: 868, h: 253 } } });
  assert.deepEqual(harbor.starts.slice(0, 3), [2, 5, 7]);
  assert.equal(harbor.ends[MODULE_COLS - 1], 30);
  // a body the tokens do not describe falls back to the grid
  assert.deepEqual(columnModule({ furniture: { kicker: { x: 60, y: 42, w: 600 } } }).starts, house.starts);
  assert.deepEqual(columnModule({ furniture: { body: { x: Number.NaN, y: 0, w: 0 } } }).starts, house.starts);
}
// snap: both edges take the nearest module edge, and only inside half a module
{
  const m = columnModule(undefined);
  assert.deepEqual(snapToModule({ col: 3, cols: 28 }, m), { col: 3, cols: 28 }, "the full body band is already modular");
  assert.deepEqual(snapToModule({ col: 2, cols: 30 }, m), { col: 3, cols: 28 }, "one cell out either side snaps in");
  assert.deepEqual(snapToModule({ col: 4, cols: 10 }, m), { col: 3, cols: 11 });
  assert.deepEqual(snapToModule({ col: 23, cols: 8 }, m), { col: 22, cols: 9 });
  assert.deepEqual(snapToModule({ col: 18, cols: 12 }, m), { col: 18, cols: 11 }, "a left edge already on a module stays");
  // two cells out is where the user meant it: half a module is 1.2 cells, so nothing moves
  assert.deepEqual(snapToModule({ col: 1, cols: 32 }, m), { col: 1, cols: 32 });
  // a drop slides the block onto the module and keeps the width it had
  assert.deepEqual(snapToModule({ col: 2, cols: 18 }, m, "move"), { col: 3, cols: 18 });
  assert.deepEqual(snapToModule({ col: 1, cols: 18 }, m, "move"), { col: 1, cols: 18 }, "two cells out stays");
  assert.deepEqual(snapToModule({ col: 29, cols: 6 }, m, "move"), { col: 29, cols: 6 }, "never slides off the right edge");
  // a corner resize moves the edge the grip moved and pins the other one
  assert.deepEqual(snapToModule({ col: 4, cols: 10 }, m, "left"), { col: 3, cols: 11 });
  assert.deepEqual(snapToModule({ col: 4, cols: 10 }, m, "right"), { col: 4, cols: 10 }, "the right edge is already on one");
  assert.deepEqual(snapToModule({ col: 23, cols: 9 }, m, "right"), { col: 23, cols: 8 });
  assert.deepEqual(snapToModule({ col: 23, cols: 10 }, m, "right"), { col: 23, cols: 10 }, "two cells past the last module stays");
  // never off the slide, never narrower than a cell
  const wide = snapToModule({ col: 30, cols: 3 }, m);
  assert.ok(wide.col + wide.cols - 1 <= DECK_COLS, "stays on the slide");
  assert.ok(wide.cols >= 1);
}

// ---- Fit: the rows a block's content asks for (ranked change 9) ----
{
  const fb = (type: BlockType, o: Partial<Block> = {}): Block => ({ id: "f", type, title: "Title", row: 7, col: 3, rows: 9, cols: 8, ...o });
  // a chart, a diagram and an image are drawn to their box: Fit leaves them where they are
  for (const type of ["chart", "diagram", "image"] as BlockType[]) assert.equal(fitRows(fb(type)), null, type);
  assert.equal(fitRows(fb("stat")), 4, "a stat is its number and its caption");
  // the description, wrapped at the block's own width: 23 characters in 7 cells is one line
  assert.equal(contentLines("census sources per step", 210), 1);
  assert.equal(contentLines("x".repeat(200), 420), 4);
  assert.equal(fitRows(fb("text", { cols: 7, content: "census sources per step" })), 2);
  assert.equal(fitRows(fb("bullets", { cols: 14, content: "x".repeat(200) })), 3, "four lines of body take two rows and the label one");
  assert.ok(fitRows(fb("text", { cols: 4, content: "x".repeat(200) }))! > fitRows(fb("text", { cols: 14, content: "x".repeat(200) }))!, "narrower wraps more");
  // a table is its header plus its rows; four data rows when the plan names none
  assert.equal(fitRows(fb("table", { cols: 30, columns: ["Player", "Days", "Trend"] })), 4);
  assert.equal(fitRows(fb("table", { cols: 30, rowLabels: ["a", "b", "c", "d", "e", "f"] })), 6);
  // cards: the count narrows each card, so the same description wraps further and the band grows
  assert.equal(fitRows(fb("cards", { cols: 28, count: 3, content: "owner and first milestone per move" })), 3);
  assert.ok(fitRows(fb("cards", { cols: 28, count: 6, content: "x".repeat(120) }))! > fitRows(fb("cards", { cols: 28, count: 2, content: "x".repeat(120) }))!);
  // fitBlocks reports whether anything moved, and never leaves the band
  const bs: Block[] = [fb("stat", { id: "s", rows: 9 }), fb("chart", { id: "c", rows: 9 })];
  assert.equal(fitBlocks(bs), true);
  assert.deepEqual(bs.map((b) => b.rows), [4, 9], "the chart keeps its height");
  assert.equal(fitBlocks(bs), false, "nothing to do the second time");
  const tall: Block[] = [fb("table", { rowLabels: Array.from({ length: 40 }, (_, i) => `r${i}`) })];
  fitBlocks(tall);
  assert.equal(tall[0].rows, 13, "no taller than the body band");
}

// ---- Balance: equal gaps down the body band (ranked change 8) ----
{
  const bb = (id: string, row: number, rows: number, col = 3, cols = 28): Block => ({ id, type: "text", title: id, row, col, rows, cols });
  // blocks that share rows are one band and move together; the band is centred when it does not fill
  const side = [bb("chart", 7, 10, 3, 18), bb("stat", 7, 5, 23, 8), bb("aside", 13, 4, 24, 7)];
  assert.deepEqual(rowBands(side).map((band) => band.map((b) => b.id)), [["chart", "stat", "aside"]]);
  assert.equal(balanceBlocks(side), true);
  assert.deepEqual(side.map((b) => [b.row, b.col]), [[6, 3], [6, 23], [12, 24]], "the whole band shifts up one, columns kept");
  // two bands keep the gap the slide already has, and the pair is centred in the band
  const two = [bb("cards", 6, 8), bb("takeaway", 16, 2)];
  assert.equal(balanceBlocks(two), true);
  assert.deepEqual(two.map((b) => b.row), [5, 15]);
  assert.equal(two[1].row - (two[0].row + two[0].rows), 2, "the gap it had");
  // three even bands: one blank row between, one above and one below
  const three = [bb("a", 5, 3), bb("b", 9, 3), bb("c", 13, 3)];
  assert.equal(balanceBlocks(three), true);
  assert.deepEqual(three.map((b) => b.row), [6, 10, 14]);
  assert.equal(balanceBlocks(three), false, "already balanced");
  // a group that fills the band lets its bands abut rather than run past row 17
  const packed = [bb("a", 5, 5), bb("b", 12, 4), bb("c", 16, 3)];
  balanceBlocks(packed);
  assert.deepEqual(packed.map((b) => b.row), [5, 10, 14]);
  assert.equal(packed[2].row + packed[2].rows - 1, 16, "inside rows 5 to 17");
  assert.equal(balanceBlocks([]), false);
}

// ---- F18: the Edges field ----
{
  const one = parseDiagram("Signup -> First invoice: two days");
  assert.deepEqual(one.nodes, [{ id: "signup", label: "Signup" }, { id: "first-invoice", label: "First invoice" }]);
  assert.deepEqual(one.edges, [{ from: "signup", to: "first-invoice", label: "two days" }]);
  // a chain is one edge per hop and the trailing label belongs to the last one
  const chain = parseDiagram("A -> B -> C: then");
  assert.deepEqual(chain.nodes.map((n) => n.id), ["a", "b", "c"]);
  assert.deepEqual(chain.edges, [{ from: "a", to: "b" }, { from: "b", to: "c", label: "then" }]);
  // a lone line is a node, colon and all; blank lines are skipped; the same label is the same node
  const mixed = parseDiagram("Onboarding: week one\n\nOnboarding -> Billing\n  \nRetention");
  assert.deepEqual(mixed.nodes.map((n) => n.label), ["Onboarding: week one", "Onboarding", "Billing", "Retention"]);
  assert.deepEqual(mixed.edges, [{ from: "onboarding", to: "billing" }]);
  assert.equal(diagramId("Year 1 revenue ($M)"), "year-1-revenue-m");
  assert.equal(diagramId("!!!"), "node", "a label with nothing to slug still gets an id");
  assert.deepEqual(parseDiagram("A → B"), { nodes: [{ id: "a", label: "A" }, { id: "b", label: "B" }], edges: [{ from: "a", to: "b" }] }, "an arrow glyph counts too");
  // the field shows back what the block holds, unlinked nodes last
  const d = { kind: "flow" as const, direction: "LR" as const, ...parseDiagram("Signup -> Invoice: two days\nChurn") };
  assert.equal(diagramText(d), "Signup -> Invoice: two days\nChurn");
  assert.equal(diagramText(undefined), "");
  assert.deepEqual(parseDiagram(diagramText(d)), parseDiagram("Signup -> Invoice: two days\nChurn"), "round trip");
}

console.log("LAYOUT OK");

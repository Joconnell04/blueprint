import { h, clear, svg, icon, twoStep, popover } from "./dom.js";
import { fadeIn, pulse, slideIn } from "./motion.js";
import type { Controller } from "./controller.js";
import { CELL_H, CELL_W, TYPE_LABEL, typeLabel, addr, colLetter, fakeValue, overlaps, uid, isReadMe, type Block, type BlockType, type Note, type Warning, type WorkbookBlockType, type Tab } from "./model.js";
import { notesPanel } from "./notes.js";

const MIN_COLS = 26, MIN_ROWS = 44, FLOW_MAX_COLS = 16;
export const DEFAULT_SIZE: Record<WorkbookBlockType, [number, number]> = { inputs: [5, 3], output: [4, 3], table: [6, 5], calc: [6, 5], chart: [12, 6], text: [3, 6] };
const DEFAULT_TITLE: Record<WorkbookBlockType, string> = { inputs: "Key Assumptions", table: "Data Table", calc: "Calculation", output: "Key Outputs", chart: "Chart", text: "Notes" };
const TYPE_DESC: Record<WorkbookBlockType, string> = { inputs: "labels and values", table: "raw data", calc: "formulas", output: "results", chart: "plots a table", text: "paragraph" };

/**
 * A table or calc block's honest width: one cell per column, a label column plus two at the least, never
 * past the sheet's right edge. The builder writes one column per header, so a wider default draws a
 * footprint the sheet will not have (ux-audit item 12: a three-column table five cells wide).
 */
export function fitCols(columns?: string[]): number {
  return Math.min(MIN_COLS - 1, Math.max(3, columns?.length ?? 0));
}

/** Px per character of the canvas's mono face, as a fraction of the font size. */
const CHAR_W = 0.62;
/** The canvas cell size, mirroring --fs-micro. A narrow header row steps one down to --fs-tag (the `.tight` class). */
const FS_CELL = 10;

/** The label column's share of a block's width: 60% for a label/value pair, narrower the more columns there are. */
export function labelPct(type: WorkbookBlockType, columns?: string[]): number {
  if (type === "inputs" || type === "output") return 60;
  return Math.max(22, Math.min(46, Math.round(100 / Math.max(1, columns?.length ?? 4)) + 12));
}

/**
 * Does every header fit its column at `fs`? A block too narrow for its headers steps them one size down
 * (--fs-micro to --fs-tag) before it truncates, so six columns stop reading "Yea…" (ux-audit item 1).
 * Mirrors the CSS: the first cell takes labelPct of the width, the rest share what is left, 4px padding
 * each side, 1px border each side.
 */
export function headersFit(columns: string[], cols: number, fs: number = FS_CELL): boolean {
  const n = columns.length;
  if (!n) return true;
  const w = cols * CELL_W - 2;
  const label = (w * labelPct("table", columns)) / 100;
  const each = n > 1 ? (w - label) / (n - 1) : w;
  const fits = (box: number, text: string) => Math.floor(Math.max(0, box - 8) / (fs * CHAR_W)) >= text.length;
  return fits(label, columns[0]) && columns.slice(1).every((c) => fits(each, c));
}

/** What a column holds, read from its header. The families are fakeValue's, so a sample and a guess take the same shape. */
export type ColKind = "money" | "percent" | "multiple" | "year" | "count";
export function colKind(header: string): ColKind {
  const h = header.toLowerCase();
  if (/\$|revenue|cost|price|arr|spend|sales|ebitda|cash|capex|opex|value|fee|budget|income/.test(h)) return "money";
  if (/%|rate|share|margin|growth|mix|penetration|conversion|churn|yield|pct/.test(h)) return "percent";
  if (/multiple|x$/.test(h)) return "multiple";
  if (/^(year|fy|date|period|month|quarter)$/.test(h.trim())) return "year";
  return "count";
}

/**
 * A real sample in its column's number format, the way the builder will write it: 32000000 under
 * "Businesses" reads 32,000,000 rather than bare digits (ux-feedback item 3). A string is the user's own
 * text and passes through. A `Year 1` column is left unsigned: the generator guesses money there, but a
 * sample under it is as often a count.
 */
export function formatSample(header: string, v: string | number): string {
  if (typeof v !== "number" || !Number.isFinite(v)) return String(v).trim();
  const num = (x: number, dp = 0) => x.toLocaleString("en-US", { minimumFractionDigits: dp, maximumFractionDigits: dp });
  switch (colKind(header)) {
    // whole dollars above 100, cents below it: the builder's currency and currency2 formats
    case "money": return `$${num(v, Math.abs(v) >= 100 ? 0 : 2)}`;
    // a fraction reads as a percent, as Excel's percent format does; one decimal, and none when it is whole
    case "percent": { const pct = Math.round((Math.abs(v) <= 1 ? v * 100 : v) * 10) / 10; return `${num(pct, Number.isInteger(pct) ? 0 : 1)}%`; }
    case "multiple": return `${num(v, 1)}x`;
    case "year": return String(Math.round(v));
    default: return num(v, Number.isInteger(v) ? 0 : 1);
  }
}

/**
 * One value cell: the sample the plan carries, or the generator's guess. A guess is marked `.ph` (muted,
 * dotted) so nobody reads it as data, and a sample that is the generator's own output for that column and
 * row is a guess too, whoever pushed it into the plan (there is no flag on the wire to say so).
 */
export function valueCell(header: string, i: number, raw: string | number | undefined): { text: string; ph: boolean } {
  const gen = fakeValue(header, i);
  if (raw == null || String(raw).trim() === "" || String(raw) === gen) return { text: gen, ph: true };
  return { text: formatSample(header, raw), ph: false };
}

/** Vertical stack: one column, reading order top to bottom, a blank row between blocks. */
export function stackBlocks(blocks: Block[]) {
  let r = 2;
  for (const b of [...blocks].sort((a, b) => a.row - b.row || a.col - b.col)) { b.row = r; b.col = 2; r += b.rows + 1; }
}
/** Horizontal flow: fill left to right, a blank column between blocks, wrap to a new band when the row is full. */
export function flowBlocks(blocks: Block[]) {
  let r = 2, col = 2, bandH = 0;
  for (const b of [...blocks].sort((a, b) => a.row - b.row || a.col - b.col)) {
    if (col > 2 && col + b.cols - 1 > FLOW_MAX_COLS) { r += bandH + 1; col = 2; bandH = 0; }
    b.row = r; b.col = col; col += b.cols + 1; bandH = Math.max(bandH, b.rows);
  }
}

/**
 * Push every block that overlaps `moved` out of the way: down in reading order, keeping columns. Repeats until stable.
 * `moved` may be a group (a multi-selection dropped together): the group stays put and only other blocks move.
 */
export function resolveOverlaps(blocks: Block[], moved: Block | Block[], maxPasses = 20): boolean {
  const fixed = new Set(Array.isArray(moved) ? moved : [moved]);
  let changed = false;
  for (let pass = 0; pass < maxPasses; pass++) {
    let any = false;
    for (const o of [...blocks].sort((a, b) => a.row - b.row)) {
      if (fixed.has(o)) continue;
      const hit = blocks.some((x) => x !== o && overlaps(x, o) && (fixed.has(x) || x.row <= o.row));
      if (!hit) continue;
      const blocker = blocks.filter((x) => x !== o && overlaps(x, o)).sort((a, b) => b.row + b.rows - (a.row + a.rows))[0];
      o.row = blocker.row + blocker.rows + 1; any = true; changed = true;
    }
    if (!any) break;
  }
  return changed;
}

// ---- align and distribute (pure, on the selected blocks) ----
/** Every block takes the leftmost column of the group. */
export function alignLeft(sel: Block[]) { if (!sel.length) return; const c = Math.min(...sel.map((b) => b.col)); for (const b of sel) b.col = c; }
/** Every block takes the topmost row of the group. */
export function alignTop(sel: Block[]) { if (!sel.length) return; const r = Math.min(...sel.map((b) => b.row)); for (const b of sel) b.row = r; }
/** Every block takes the width of `ref` (the primary selection by default). */
export function sameWidth(sel: Block[], ref: Block = sel[0]) { if (!ref) return; for (const b of sel) b.cols = ref.cols; }
/** Every block takes the height of `ref` (the primary selection by default). */
export function sameHeight(sel: Block[], ref: Block = sel[0]) { if (!ref) return; for (const b of sel) b.rows = ref.rows; }
/** Keep the current top-to-bottom order and the first block's row; put the same number of blank rows (at least one) between neighbours, spreading the span the group already covers. */
export function distributeVertically(sel: Block[]) {
  if (sel.length < 2) return;
  const order = [...sel].sort((a, b) => a.row - b.row || a.col - b.col);
  const span = Math.max(...order.map((b) => b.row + b.rows)) - order[0].row;
  const total = order.reduce((s, b) => s + b.rows, 0);
  const gap = Math.max(1, Math.round((span - total) / (order.length - 1)));
  let r = order[0].row;
  for (const b of order) { b.row = r; r += b.rows + gap; }
}
/** Same as distributeVertically, left to right on columns. */
export function distributeHorizontally(sel: Block[]) {
  if (sel.length < 2) return;
  const order = [...sel].sort((a, b) => a.col - b.col || a.row - b.row);
  const span = Math.max(...order.map((b) => b.col + b.cols)) - order[0].col;
  const total = order.reduce((s, b) => s + b.cols, 0);
  const gap = Math.max(1, Math.round((span - total) / (order.length - 1)));
  let c = order[0].col;
  for (const b of order) { b.col = c; c += b.cols + gap; }
}

/** The nearest block directly above `b` (column overlap, at most two blank rows between). */
export function blockAbove(blocks: Block[], b: Block): Block | undefined {
  return blocks.filter((o) => o !== b && o.row + o.rows <= b.row && b.row - (o.row + o.rows) <= 2 && o.col < b.col + b.cols && b.col < o.col + o.cols).sort((x, y) => (y.row + y.rows) - (x.row + x.rows))[0];
}
/** Merge `upper` into `lower`: union rectangle; a bare title (text block or a block with no columns or line items) donates its title, otherwise the lower block's identity wins. */
export function mergeBlocks(blocks: Block[], upper: Block, lower: Block): Block {
  const bare = upper.type === "text" || (!(upper.columns?.length) && !(upper.rowLabels?.length));
  const row = Math.min(upper.row, lower.row), col = Math.min(upper.col, lower.col);
  const bottom = Math.max(upper.row + upper.rows, lower.row + lower.rows), right = Math.max(upper.col + upper.cols, lower.col + lower.cols);
  const merged: Block = { ...lower, row, col, rows: bottom - row, cols: right - col, title: bare ? upper.title : lower.title, detected: undefined };
  const i = blocks.indexOf(lower); blocks.splice(i, 1, merged);
  blocks.splice(blocks.indexOf(upper), 1);
  return merged;
}
/** Split a table-like block at a data row (1-based within the data area): the second part gets the same columns and a numbered title. */
export function splitBlock(blocks: Block[], b: Block, atDataRow: number): Block | undefined {
  const headerRows = b.type === "table" || b.type === "calc" ? 2 : 1;
  const dataRows = b.rows - headerRows;
  if (atDataRow < 1 || atDataRow >= dataRows) return undefined;
  const second: Block = { ...b, id: uid("blk"), title: `${b.title} (2)`, row: b.row + headerRows + atDataRow, rows: dataRows - atDataRow + headerRows, rowLabels: b.rowLabels?.slice(atDataRow), sample: b.sample?.slice(atDataRow), detected: undefined };
  b.rows = headerRows + atDataRow; b.rowLabels = b.rowLabels?.slice(0, atDataRow); b.sample = b.sample?.slice(0, atDataRow);
  // make room: the second part starts one blank row below the first
  second.row = b.row + b.rows + 1;
  blocks.splice(blocks.indexOf(b) + 1, 0, second);
  return second;
}

// ---- corner resize (F19): both canvases hang four grips off every block ----

export type Corner = "nw" | "ne" | "sw" | "se";
export const CORNERS: Corner[] = ["nw", "ne", "sw", "se"];
export interface Rect { row: number; col: number; rows: number; cols: number }
/** Inclusive cell limits. Omitted sides are open, which is what the sheet canvas wants below and right. */
export interface Bounds { top?: number; bottom?: number; left?: number; right?: number }

/**
 * Corner drag maths, in cells. The south and east edges change the size only; the north and west edges
 * move the block's row or column as they change it, and stop at the opposite edge so the block never
 * flips through itself. Pure, so test/layout.test.ts can check it without a pointer.
 */
export function cornerResize(start: Rect, corner: Corner, dRow: number, dCol: number, bounds: Bounds = {}): Rect {
  const top = bounds.top ?? 1, left = bounds.left ?? 1;
  const bottom = bounds.bottom ?? Number.MAX_SAFE_INTEGER, right = bounds.right ?? Number.MAX_SAFE_INTEGER;
  const out: Rect = { ...start };
  if (corner === "nw" || corner === "ne") {
    const row = Math.min(Math.max(top, start.row + dRow), start.row + start.rows - 1);
    out.row = row; out.rows = start.row + start.rows - row;
  } else out.rows = Math.max(1, Math.min(bottom - start.row + 1, start.rows + dRow));
  if (corner === "nw" || corner === "sw") {
    const col = Math.min(Math.max(left, start.col + dCol), start.col + start.cols - 1);
    out.col = col; out.cols = start.col + start.cols - col;
  } else out.cols = Math.max(1, Math.min(right - start.col + 1, start.cols + dCol));
  return out;
}

/** The four grips a block carries. The stylesheet places each one and turns its arrow. */
export function resizeHandles(): HTMLElement[] {
  return CORNERS.map((c) => h("div", { class: `rs ${c}`, title: "Resize" }));
}

/**
 * Wire all four grips on one block. One gesture is one undo step: `done` is called on pointer up, and
 * only when the geometry actually changed. It is handed the corner that was dragged, so a canvas that
 * snaps can move the edge the user moved and leave the opposite one alone.
 */
export function wireCornerResize(opts: {
  el: HTMLElement; block: Block;
  cell: { w: number; h: number };
  bounds?: Bounds;
  zoom: () => number;
  place: () => void;
  moved?: () => void;
  done: (corner: Corner) => void;
  interacting?: (on: boolean) => void;
}): void {
  const b = opts.block;
  for (const corner of CORNERS) {
    const grip = opts.el.querySelector<HTMLElement>(`.rs.${corner}`);
    if (!grip) continue;
    grip.onpointerdown = (e) => {
      e.preventDefault(); e.stopPropagation();
      opts.interacting?.(true);
      try { grip.setPointerCapture(e.pointerId); } catch { /* synthetic pointer */ }
      const start: Rect = { row: b.row, col: b.col, rows: b.rows, cols: b.cols };
      // Deltas between two moves, never from the press: inside the host's nested iframes a captured
      // pointer can report another frame's coordinates, and the press and the moves then disagree.
      let from: { x: number; y: number } | null = null;
      const onMove = (ev: PointerEvent) => {
        from ??= { x: ev.clientX, y: ev.clientY };
        const z = opts.zoom();
        const dRow = Math.round((ev.clientY - from.y) / z / opts.cell.h), dCol = Math.round((ev.clientX - from.x) / z / opts.cell.w);
        Object.assign(b, cornerResize(start, corner, dRow, dCol, opts.bounds));
        opts.place(); opts.moved?.();
      };
      const onUp = () => {
        grip.removeEventListener("pointermove", onMove); grip.removeEventListener("pointerup", onUp);
        opts.interacting?.(false);
        if (b.row !== start.row || b.col !== start.col || b.rows !== start.rows || b.cols !== start.cols) opts.done(corner);
      };
      grip.addEventListener("pointermove", onMove); grip.addEventListener("pointerup", onUp);
    };
  }
}

/**
 * The Readiness list, both canvases: one line per entry, a chevron to read the rest of it. A warning
 * that names a block selects it when the line is clicked.
 */
export function readinessList(warns: Warning[], pick: (blockId: string) => void): HTMLElement {
  return h("ul", { class: "warnlist" }, ...warns.map((w) => {
    const line = h("span", { class: "m trunc" }, w.message);
    const li = h("li", { class: w.blockId ? "link" : "" }, h("span", { class: `dot ${w.kind === "overlap" ? "bad" : "warn"}` }), line);
    const chev = h("button", { class: "chev", title: "Expand", onclick: (e: Event) => {
      e.stopPropagation();
      const open = li.classList.toggle("open");
      line.classList.toggle("trunc", !open);
      chev.title = open ? "Collapse" : "Expand";
    } }, icon("chev", 9));
    li.append(chev);
    if (w.blockId) li.onclick = () => pick(w.blockId!);
    return li;
  }));
}

export const HISTORY_MAX = 60;
const histories = new Map<string, { past: string[]; future: string[] }>();
/** One undo step: the blocks as they stand. Shared by both canvases so a tab keeps one history. */
export function snapshotBlocks(blocks: Block[]): string { return JSON.stringify(blocks); }
export function historyFor(tabId: string): { past: string[]; future: string[] } { let h = histories.get(tabId); if (!h) { h = { past: [], future: [] }; histories.set(tabId, h); } return h; }

export interface Selection { setSel(ids: string[], primary: string | null): void; toggleSel(id: string): void; pressSel(id: string): void }
/**
 * Multi-select behaviour, shared by both canvases: `selectedBlocks` is the group, `selectedBlock` the
 * primary the inspector edits and the one a group action measures from.
 */
export function selection(c: Controller, refresh: () => void): Selection {
  const setSel = (ids: string[], primary: string | null) => { c.selectedBlocks = new Set(ids); c.selectedBlock = primary ?? ids[ids.length - 1] ?? null; refresh(); };
  const toggleSel = (id: string) => {
    const s = new Set(c.selectedBlocks);
    if (s.has(id)) { s.delete(id); if (c.selectedBlock === id) c.selectedBlock = [...s].pop() ?? null; }
    else { s.add(id); c.selectedBlock = id; }
    c.selectedBlocks = s; refresh();
  };
  /** A plain press on a block: a member becomes the primary and keeps the group; anything else becomes the only selection. */
  const pressSel = (id: string) => { if (!c.selectedBlocks.has(id)) setSel([id], id); else if (c.selectedBlock !== id) { c.selectedBlock = id; refresh(); } };
  return { setSel, toggleSel, pressSel };
}

/** The align and distribute bar, identical in both canvases. Hidden below two selected blocks. */
export function fillSelectionBar(bar: HTMLElement, opts: { count: number; primary?: Block; run: (fn: (sel: Block[]) => void) => void; deselect: () => void }): void {
  clear(bar);
  bar.hidden = opts.count < 2;
  if (opts.count < 2) return;
  const name = opts.primary?.title ?? "the first";
  bar.append(h("span", { class: "count" }, `${opts.count} selected`), h("span", { class: "sep" }),
    h("button", { class: "ghost", title: "Leftmost column", onclick: () => opts.run(alignLeft) }, selIcon("alignleft"), "Align left"),
    h("button", { class: "ghost", title: "Topmost row", onclick: () => opts.run(alignTop) }, selIcon("aligntop"), "Align top"),
    h("button", { class: "ghost", title: `Width of "${name}"`, onclick: () => opts.run((sel) => sameWidth(sel, opts.primary ?? sel[0])) }, selIcon("samew"), "Same width"),
    h("button", { class: "ghost", title: `Height of "${name}"`, onclick: () => opts.run((sel) => sameHeight(sel, opts.primary ?? sel[0])) }, selIcon("sameh"), "Same height"),
    h("span", { class: "sep" }),
    h("button", { class: "ghost", title: "Equal gaps down", onclick: () => opts.run(distributeVertically) }, selIcon("distv"), "Distribute vertically"),
    h("button", { class: "ghost", title: "Equal gaps across", onclick: () => opts.run(distributeHorizontally) }, selIcon("disth"), "Distribute horizontally"),
    h("span", { class: "spacer" }),
    h("button", { class: "ghost", title: "Esc", onclick: opts.deselect }, "Deselect"));
}

export interface Minimap { el: HTMLElement; draw(): void; viewport(): void }
/**
 * The canvas at a glance with the viewport as an outline; click or drag it to scroll. Hidden while the
 * whole canvas fits on screen. `rect` maps a block to canvas coordinates, which differ between the two
 * canvases (the sheet has row and column headers, the slide does not).
 */
export function createMinimap(opts: {
  W: number; H: number; wrap: HTMLElement;
  zoom: () => number;
  blocks: () => Block[];
  rect: (b: Block) => [number, number, number, number];
  selected: (b: Block) => boolean;
  interacting?: (on: boolean) => void;
}): Minimap {
  const { W, H, wrap } = opts;
  const scale = Math.min(MINIMAP_W / W, MINIMAP_MAX_H / H);
  const el = h("div", { class: "minimap", hidden: true });
  const ns = "http://www.w3.org/2000/svg";
  const svgEl = document.createElementNS(ns, "svg");
  svgEl.setAttribute("viewBox", `0 0 ${W} ${H}`); svgEl.setAttribute("width", String(Math.round(W * scale))); svgEl.setAttribute("height", String(Math.round(H * scale))); svgEl.setAttribute("preserveAspectRatio", "xMinYMin meet");
  const g = document.createElementNS(ns, "g");
  const view = document.createElementNS(ns, "rect"); view.setAttribute("class", "mv");
  svgEl.append(g, view); el.append(svgEl);
  const viewport = () => {
    const z = opts.zoom();
    el.hidden = W * z <= wrap.clientWidth + 1 && H * z <= wrap.clientHeight + 1;
    view.setAttribute("x", String(wrap.scrollLeft / z)); view.setAttribute("y", String(wrap.scrollTop / z));
    view.setAttribute("width", String(Math.min(W, wrap.clientWidth / z))); view.setAttribute("height", String(Math.min(H, wrap.clientHeight / z)));
  };
  const draw = () => {
    clear(g);
    for (const b of opts.blocks()) {
      const r = document.createElementNS(ns, "rect");
      const [x, y, w, hh] = opts.rect(b);
      r.setAttribute("class", `mb ${b.type}${opts.selected(b) ? " sel" : ""}`);
      r.setAttribute("x", String(x)); r.setAttribute("y", String(y)); r.setAttribute("width", String(w)); r.setAttribute("height", String(hh));
      g.append(r);
    }
    viewport();
  };
  const scrollTo = (ev: PointerEvent) => {
    const mr = svgEl.getBoundingClientRect();
    const z = opts.zoom();
    const x = (ev.clientX - mr.left) / scale, y = (ev.clientY - mr.top) / scale; // the point under the pointer, centred in the viewport
    wrap.scrollLeft = Math.max(0, x * z - wrap.clientWidth / 2); wrap.scrollTop = Math.max(0, y * z - wrap.clientHeight / 2);
  };
  el.onpointerdown = (e) => {
    e.preventDefault(); e.stopPropagation(); opts.interacting?.(true);
    try { el.setPointerCapture(e.pointerId); } catch { /* not capturable */ }
    scrollTo(e);
    // The press centres the viewport on the point pressed; after that the drag is deltas, for the same
    // reason the canvas drags are (a captured pointer's coordinates may be another frame's).
    let from: { x: number; y: number } | null = null;
    let at: [number, number] = [wrap.scrollLeft, wrap.scrollTop];
    const onMove = (ev: PointerEvent) => {
      if (!from) { from = { x: ev.clientX, y: ev.clientY }; at = [wrap.scrollLeft, wrap.scrollTop]; return; }
      const z = opts.zoom();
      wrap.scrollLeft = Math.max(0, at[0] + ((ev.clientX - from.x) / scale) * z);
      wrap.scrollTop = Math.max(0, at[1] + ((ev.clientY - from.y) / scale) * z);
    };
    const onUp = () => { el.removeEventListener("pointermove", onMove); el.removeEventListener("pointerup", onUp); opts.interacting?.(false); };
    el.addEventListener("pointermove", onMove); el.addEventListener("pointerup", onUp);
  };
  return { el, draw, viewport };
}

/** The question, instruction and closed-thread badges on a block header. */
export function blockBadges(notes: Note[]): Node[] {
  const q = notes.filter((n) => n.status === "open" && n.kind === "question").length;
  const instr = notes.filter((n) => n.status === "open" && n.kind === "instruction").length;
  const closed = notes.filter((n) => n.status !== "open").length;
  const out: Node[] = [];
  if (q) out.push(h("span", { class: "badge q", title: `${q} question${q === 1 ? "" : "s"} from Claude waiting on you` }, `${q}?`));
  if (instr) out.push(h("span", { class: "badge n", title: `${instr} open instruction${instr === 1 ? "" : "s"}` }, String(instr)));
  if (!q && !instr && closed) out.push(h("span", { class: "badge ok", title: `${closed} closed` }, icon("check", 9)));
  return out;
}

/** Canvas zoom, shared by every tab and kept across re-renders. */
export const ZOOM_STEPS = [0.5, 0.75, 1, 1.25];
let zoom = 1;
/** Scroll position per tab so a re-render (every save) does not jump the view. */
const scrolls = new Map<string, [number, number]>();
const MINIMAP_W = 160, MINIMAP_MAX_H = 160;

export function renderLayoutView(c: Controller, root: HTMLElement) {
  const p = c.project;
  const tab = c.tab;
  const strip = h("div", { class: "tabstrip", "data-scroll": "strip" }, ...p.tabs.map((t) => {
    const open = p.notes.filter((n) => n.tabId === t.id && n.status === "open").length;
    const cover = isReadMe(t.name) && t.blocks.length === 0;
    return h("button", { class: t.id === c.selectedTab ? "on" : "", onclick: () => { c.selectedTab = t.id; c.selectedBlock = null; c.setState(c.state!); } },
      h("span", { class: `dot ${t.layoutApproved || cover ? "ok" : t.blocks.length ? "warn" : "none"}` }), t.name, open ? h("span", { class: "n" }, String(open)) : null);
  }));
  const main = h("div", { class: "main" }, strip);
  const side = h("div", { class: "side" });
  root.append(h("div", { class: "body" }, main, side));
  if (!tab) { main.append(h("div", { class: "center" }, "No tabs yet")); return; }
  const blocks = tab.blocks;
  const hist = historyFor(tab.id);
  let before = snapshotBlocks(blocks);
  // Selection invariants: members exist on this tab; the primary is a member; no primary means nothing selected.
  for (const id of [...c.selectedBlocks]) if (!blocks.some((b) => b.id === id)) c.selectedBlocks.delete(id);
  if (!c.selectedBlock) c.selectedBlocks.clear();
  else if (!c.selectedBlocks.has(c.selectedBlock)) c.selectedBlocks = new Set([c.selectedBlock]);
  const selected = () => blocks.filter((b) => c.selectedBlocks.has(b.id));

  const persist = (opts: { record?: boolean; resolve?: Block | Block[] } = {}) => {
    if (opts.resolve && resolveOverlaps(blocks, opts.resolve)) c.toast?.("Moved overlapping blocks down");
    const after = snapshotBlocks(blocks);
    if (opts.record !== false && after !== before) { hist.past.push(before); if (hist.past.length > HISTORY_MAX) hist.past.shift(); hist.future = []; }
    before = after;
    c.saveLayout(tab, blocks);
  };
  const restore = (json: string) => { const next = JSON.parse(json) as Block[]; blocks.splice(0, blocks.length, ...next); before = json; c.saveLayout(tab, blocks); };
  const undo = () => { const prev = hist.past.pop(); if (prev == null) { c.toast?.("Nothing to undo"); return; } hist.future.push(snapshotBlocks(blocks)); restore(prev); };
  const redo = () => { const next = hist.future.pop(); if (next == null) return; hist.past.push(snapshotBlocks(blocks)); restore(next); };
  document.onkeydown = (e) => {
    const inField = (e.target as HTMLElement).closest("input, textarea, select");
    const mod = e.metaKey || e.ctrlKey;
    if (mod && e.key.toLowerCase() === "z") { if (inField) return; e.preventDefault(); if (e.shiftKey) redo(); else undo(); return; }
    if (mod && (e.key === "=" || e.key === "+")) { e.preventDefault(); setZoom(ZOOM_STEPS[Math.min(ZOOM_STEPS.length - 1, ZOOM_STEPS.indexOf(zoom) + 1)]); return; }
    if (mod && (e.key === "-" || e.key === "_")) { e.preventDefault(); setZoom(ZOOM_STEPS[Math.max(0, ZOOM_STEPS.indexOf(zoom) - 1)]); return; }
    if (mod && e.key === "0") { e.preventDefault(); setZoom(1); return; }
    if (inField) return;
    const b = c.block;
    if (mod && e.key.toLowerCase() === "d" && b) { e.preventDefault(); const copy: Block = { ...b, id: uid("blk"), title: `${b.title} copy`, row: b.row + b.rows + 1 }; blocks.push(copy); c.selectedBlock = copy.id; c.selectedBlocks = new Set([copy.id]); persist({ resolve: copy }); return; }
    if (e.key === "Escape") { document.querySelectorAll(".pop").forEach((x) => x.remove()); c.selectedBlock = null; c.selectedBlocks.clear(); c.setState(c.state!); return; }
    if (e.key === "?" ) { e.preventDefault(); openCheatSheet(root); return; }
    if (/^[1-9]$/.test(e.key) && !mod) { const t = p.tabs[Number(e.key) - 1]; if (t) { c.selectedTab = t.id; c.selectedBlock = null; c.setState(c.state!); } return; }
    if (!b) return;
    const sel = selected();
    const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key as string];
    if (d) {
      e.preventDefault();
      if (e.shiftKey) { for (const s of sel) { s.rows = Math.max(1, s.rows + d[0]); s.cols = Math.max(1, s.cols + d[1]); } }
      else {
        // nudge the whole selection by one cell, clamped so the group keeps its shape at the sheet edge
        const dRow = Math.max(1 - Math.min(...sel.map((s) => s.row)), d[0]), dCol = Math.max(1 - Math.min(...sel.map((s) => s.col)), d[1]);
        for (const s of sel) { s.row += dRow; s.col += dCol; }
      }
      persist(); return;
    }
    if (e.key === "Backspace" || e.key === "Delete") { e.preventDefault(); for (const s of sel) { const i = blocks.indexOf(s); if (i >= 0) blocks.splice(i, 1); } c.selectedBlock = null; c.selectedBlocks.clear(); persist(); }
  };

  // ---- toolbar ----
  const addBlock = (type: WorkbookBlockType) => {
    const bottom = blocks.reduce((m, b) => Math.max(m, b.row + b.rows), 1);
    const [rows, cols] = DEFAULT_SIZE[type];
    const b: Block = { id: uid("blk"), type, title: DEFAULT_TITLE[type], row: bottom + 1, col: 2, rows, cols };
    // A table or calc is as wide as its columns, not the type default: the canvas draws the footprint the sheet will have.
    if (type === "table" || type === "calc") { b.columns = ["Item", "Year 1", "Year 2", "Year 3"]; b.cols = fitCols(b.columns); }
    if (type === "inputs" || type === "output") b.rowLabels = ["Line item 1", "Line item 2"];
    blocks.push(b); c.selectedBlock = b.id; c.selectedBlocks = new Set([b.id]); persist({ resolve: b });
    requestAnimationFrame(() => { scrollTo(b); (side.querySelector("input.title") as HTMLInputElement | null)?.select(); });
  };
  const addBtn = h("button", { onclick: (e: Event) => popover(e.currentTarget as HTMLElement, (close) => h("div", {}, ...(Object.keys(TYPE_LABEL) as WorkbookBlockType[]).map((t) => h("button", { onclick: () => { close(); addBlock(t); } }, h("span", { class: `ty ${t}` }, TYPE_LABEL[t]), h("span", { class: "desc" }, TYPE_DESC[t]))))) }, icon("plus", 11), "Block");
  const warnForTab = c.state!.warnings.filter((w) => w.tabId === tab.id);
  const cover = isReadMe(tab.name) && blocks.length === 0;
  const zoomBtns = ZOOM_STEPS.map((z) => h("button", { class: z === zoom ? "on" : "", title: `${Math.round(z * 100)}%`, onclick: () => setZoom(z) }, `${Math.round(z * 100)}%`));
  main.append(h("div", { class: "toolbar" }, addBtn, h("span", { class: "sep" }),
    h("button", { class: "ghost", title: "One column", onclick: () => { stackBlocks(blocks); persist(); } }, icon("stack", 11), "Stack"),
    h("button", { class: "ghost", title: "Side by side", onclick: () => { flowBlocks(blocks); persist(); } }, icon("flow", 11), "Flow"),
    h("span", { class: "sep" }),
    h("button", { class: "ghost icon", title: "Undo (⌘Z)", disabled: hist.past.length === 0, onclick: undo }, icon("undo", 12)),
    h("button", { class: "ghost icon", title: "Redo (⇧⌘Z)", disabled: hist.future.length === 0, onclick: redo }, icon("redo", 12)),
    h("span", { class: "purpose", title: tab.purpose }, tab.purpose),
    h("div", { class: "seg zoom", title: "Zoom" }, ...zoomBtns),
    h("button", { class: "ghost icon", title: "Shortcuts (?)", onclick: () => openCheatSheet(root) }, "?"),
    warnForTab.length ? h("span", { class: "chip warn", title: warnForTab.map((w) => w.message).join("\n") }, h("span", { class: "dot warn" }), `${warnForTab.length} flagged`) : h("span", { class: "chip ok" }, h("span", { class: "dot ok" }), tab.layoutApproved || cover ? "approved" : "ready")));
  // Selection bar: align and distribute, shown only with two or more blocks selected. Each action is one undo step.
  const selbar = h("div", { class: "toolbar selbar", hidden: true });
  main.append(selbar);
  const run = (fn: (sel: Block[]) => void) => { const sel = selected(); if (sel.length < 2) return; fn(sel); persist({ resolve: sel }); };
  const refreshSelbar = () => fillSelectionBar(selbar, { count: c.selectedBlocks.size, primary: c.block, run, deselect: () => setSel([], null) });

  // ---- canvas ----
  const maxCol = Math.max(MIN_COLS, ...blocks.map((b) => b.col + b.cols + 4));
  const maxRow = Math.max(MIN_ROWS, ...blocks.map((b) => b.row + b.rows + 8));
  const W = 30 + maxCol * CELL_W, H = 20 + maxRow * CELL_H;
  const wrap = h("div", { class: "canvaswrap", tabindex: "0", "data-scroll": `canvas:${tab.id}` });
  // The scaler is the scrollable footprint (W×H times the zoom); the canvas inside keeps its unscaled coordinates and is scaled with a transform.
  const scaler = h("div", { class: "canvasscale" });
  const canvas = h("div", { class: "canvas", style: { width: `${W}px`, height: `${H}px` } });
  const colhdr = h("div", { class: "colhdr", style: { left: "30px", top: 0, width: `${maxCol * CELL_W}px` } }, ...Array.from({ length: maxCol }, (_, i) => h("div", { style: { width: `${CELL_W}px` } }, colLetter(i + 1))));
  const rowhdr = h("div", { class: "rowhdr", style: { top: "20px", left: 0, height: `${maxRow * CELL_H}px` } }, ...Array.from({ length: maxRow }, (_, i) => h("div", { style: { height: `${CELL_H}px`, lineHeight: `${CELL_H}px` } }, String(i + 1))));
  const corner = h("div", { class: "corner", style: { top: 0, left: 0 } });
  canvas.append(
    h("div", { class: "grid", style: { left: "30px", top: "20px", width: `${maxCol * CELL_W}px`, height: `${maxRow * CELL_H}px`, backgroundSize: `${CELL_W}px ${CELL_H}px` } }),
    colhdr, rowhdr, corner);
  if (cover) canvas.append(h("div", { class: "cover" }, h("b", {}, "Cover page"), "Generated at build"));
  scaler.append(canvas); wrap.append(scaler); main.append(wrap);
  // Sticky headers do not survive a transform on their ancestor, so they are absolute and follow the scroll by hand.
  const syncHeaders = () => { const sx = wrap.scrollLeft / zoom, sy = wrap.scrollTop / zoom; colhdr.style.transform = `translateY(${sy}px)`; rowhdr.style.transform = `translateX(${sx}px)`; corner.style.transform = `translate(${sx}px, ${sy}px)`; };
  const applyZoom = () => { scaler.style.width = `${W * zoom}px`; scaler.style.height = `${H * zoom}px`; canvas.style.transform = `scale(${zoom})`; zoomBtns.forEach((btn, i) => btn.classList.toggle("on", ZOOM_STEPS[i] === zoom)); syncHeaders(); drawMini(); };
  const setZoom = (z: number) => {
    if (z === zoom) return;
    const cx = (wrap.scrollLeft + wrap.clientWidth / 2) / zoom, cy = (wrap.scrollTop + wrap.clientHeight / 2) / zoom; // keep the same sheet point centred
    zoom = z; applyZoom();
    wrap.scrollLeft = Math.max(0, cx * zoom - wrap.clientWidth / 2); wrap.scrollTop = Math.max(0, cy * zoom - wrap.clientHeight / 2);
    syncHeaders(); drawMini();
  };
  wrap.onscroll = () => { syncHeaders(); drawViewport(); scrolls.set(tab.id, [wrap.scrollLeft, wrap.scrollTop]); };

  const blockEls = new Map<string, HTMLElement>();
  const place = (el: HTMLElement, b: Block) => { el.style.left = `${30 + (b.col - 1) * CELL_W}px`; el.style.top = `${20 + (b.row - 1) * CELL_H}px`; el.style.width = `${b.cols * CELL_W - 1}px`; el.style.height = `${b.rows * CELL_H - 1}px`; };
  const scrollTo = (b: Block) => { wrap.scrollTop = Math.max(0, (b.row - 2) * CELL_H * zoom); const x = (b.col - 1) * CELL_W * zoom; if (x < wrap.scrollLeft || x > wrap.scrollLeft + wrap.clientWidth - 200) wrap.scrollLeft = Math.max(0, x - 40); };
  const isVisible = (b: Block) => { const l = (30 + (b.col - 1) * CELL_W) * zoom, t = (20 + (b.row - 1) * CELL_H) * zoom, r = l + b.cols * CELL_W * zoom, btm = t + b.rows * CELL_H * zoom; return l >= wrap.scrollLeft && r <= wrap.scrollLeft + wrap.clientWidth && t >= wrap.scrollTop && btm <= wrap.scrollTop + wrap.clientHeight; };

  const { setSel, toggleSel, pressSel } = selection(c, () => refreshSelection());

  for (const b of blocks) {
    const el = renderBlock(b, p.notes.filter((n) => n.blockId === b.id));
    place(el, b);
    el.classList.toggle("sel", c.selectedBlocks.has(b.id));
    el.classList.toggle("primary", c.selectedBlock === b.id);
    el.classList.toggle("overlap", blocks.some((o) => o !== b && overlaps(o, b)));
    if (c.changed.has(b.id)) pulse(el);
    blockEls.set(b.id, el);
    el.addEventListener("pointerdown", (e) => { e.stopPropagation(); wrap.focus({ preventScroll: true }); if (e.shiftKey) toggleSel(b.id); else pressSel(b.id); });
    const header = el.querySelector(".bh") as HTMLElement;
    header.onpointerdown = (e) => {
      e.preventDefault(); e.stopPropagation();
      wrap.focus({ preventScroll: true });
      if (e.shiftKey) { toggleSel(b.id); return; }
      pressSel(b.id);
      c.interacting = true;
      try { header.setPointerCapture(e.pointerId); } catch { /* synthetic pointer */ }
      // Drag the whole selection: every member keeps its offset from the pressed block; the group is clamped as one at the sheet edge.
      const group = selected();
      const starts = new Map(group.map((g) => [g.id, { row: g.row, col: g.col }]));
      const minRow = Math.min(...group.map((g) => g.row)), minCol = Math.min(...group.map((g) => g.col));
      let from: { x: number; y: number } | null = null;   // the first move is the origin, not the press
      el.classList.add("dragging");
      const onMove = (ev: PointerEvent) => {
        from ??= { x: ev.clientX, y: ev.clientY };
        const dCol = Math.max(1 - minCol, Math.round((ev.clientX - from.x) / zoom / CELL_W)), dRow = Math.max(1 - minRow, Math.round((ev.clientY - from.y) / zoom / CELL_H));
        for (const g of group) { const s = starts.get(g.id)!; g.col = s.col + dCol; g.row = s.row + dRow; place(blockEls.get(g.id)!, g); }
        markOverlaps(); drawMini();
      };
      const onUp = () => {
        header.removeEventListener("pointermove", onMove); header.removeEventListener("pointerup", onUp); el.classList.remove("dragging"); c.interacting = false;
        if (group.some((g) => { const s = starts.get(g.id)!; return g.row !== s.row || g.col !== s.col; })) persist({ resolve: group });
        refreshSide();
      };
      header.addEventListener("pointermove", onMove); header.addEventListener("pointerup", onUp);
    };
    header.ondblclick = () => { c.selectedBlock = b.id; refreshSelection(); const input = side.querySelector<HTMLInputElement>("input.title"); input?.focus(); input?.select(); };
    // Four corners (F19): the north and west grips move the block's cell as they resize it.
    wireCornerResize({
      el, block: b, cell: { w: CELL_W, h: CELL_H }, zoom: () => zoom,
      place: () => place(el, b),
      moved: () => { markOverlaps(); drawMini(); },
      done: () => { persist({ resolve: b }); refreshSide(); },
      interacting: (on) => { c.interacting = on; },
    });
    canvas.append(el);
  }
  const markOverlaps = () => { for (const b of blocks) blockEls.get(b.id)?.classList.toggle("overlap", blocks.some((o) => o !== b && overlaps(o, b))); };
  const refreshSelection = (full = true) => {
    for (const [id, el] of blockEls) { el.classList.toggle("sel", c.selectedBlocks.has(id)); el.classList.toggle("primary", id === c.selectedBlock); }
    drawMini();
    if (full) { refreshSide(); refreshSelbar(); }
  };

  // Marquee: press on empty grid and drag; every block the rectangle touches is selected (shift keeps the current selection). A plain press clears it.
  wrap.onpointerdown = (e) => {
    if (e.button !== 0 || (e.target as HTMLElement).closest(".blk")) return;
    const wr = wrap.getBoundingClientRect();
    if (e.clientX - wr.left >= wrap.clientWidth || e.clientY - wr.top >= wrap.clientHeight) return; // scrollbar
    const keep = e.shiftKey ? [...c.selectedBlocks] : [];
    if (!e.shiftKey) setSel([], null);
    wrap.focus({ preventScroll: true });
    const toCanvas = (ev: PointerEvent): [number, number] => { const cr = canvas.getBoundingClientRect(); return [(ev.clientX - cr.left) / zoom, (ev.clientY - cr.top) / zoom]; };
    const [x0, y0] = toCanvas(e);   // the press is in this frame's own coordinates; the moves may not be
    let from: { x: number; y: number } | null = null;
    let box: HTMLElement | null = null;
    try { wrap.setPointerCapture(e.pointerId); } catch { /* not capturable */ }
    c.interacting = true;
    const onMove = (ev: PointerEvent) => {
      from ??= { x: ev.clientX, y: ev.clientY };
      const dx = ev.clientX - from.x, dy = ev.clientY - from.y;
      if (!box) { if (Math.abs(dx) < 3 && Math.abs(dy) < 3) return; box = h("div", { class: "marquee" }); canvas.append(box); }
      const x1 = x0 + dx / zoom, y1 = y0 + dy / zoom;
      const l = Math.min(x0, x1), t = Math.min(y0, y1), r = Math.max(x0, x1), btm = Math.max(y0, y1);
      Object.assign(box.style, { left: `${l}px`, top: `${t}px`, width: `${r - l}px`, height: `${btm - t}px` });
      const hits = blocks.filter((b) => { const bl = 30 + (b.col - 1) * CELL_W, bt = 20 + (b.row - 1) * CELL_H; return bl < r && l < bl + b.cols * CELL_W && bt < btm && t < bt + b.rows * CELL_H; }).map((b) => b.id);
      const ids = [...new Set([...keep, ...hits])];
      c.selectedBlocks = new Set(ids); c.selectedBlock = hits[hits.length - 1] ?? ids[ids.length - 1] ?? null;
      refreshSelection(false);
    };
    const onUp = () => { wrap.removeEventListener("pointermove", onMove); wrap.removeEventListener("pointerup", onUp); wrap.removeEventListener("pointercancel", onUp); box?.remove(); c.interacting = false; refreshSelection(true); };
    wrap.addEventListener("pointermove", onMove); wrap.addEventListener("pointerup", onUp); wrap.addEventListener("pointercancel", onUp);
  };

  // ---- minimap: the sheet at a glance, the viewport as an outline; click or drag to scroll. Hidden while the sheet fits. ----
  const mini = createMinimap({
    W, H, wrap, zoom: () => zoom, blocks: () => blocks,
    rect: (b) => [30 + (b.col - 1) * CELL_W, 20 + (b.row - 1) * CELL_H, b.cols * CELL_W, b.rows * CELL_H],
    selected: (b) => c.selectedBlocks.has(b.id),
    interacting: (on) => { c.interacting = on; },
  });
  main.append(mini.el);
  const drawMini = () => mini.draw();
  const drawViewport = () => mini.viewport();
  new ResizeObserver(() => drawViewport()).observe(wrap);

  applyZoom();
  const saved = scrolls.get(tab.id);
  if (saved) { wrap.scrollLeft = saved[0]; wrap.scrollTop = saved[1]; }
  syncHeaders(); drawMini();
  if (c.selectedBlock) { const b = blocks.find((b) => b.id === c.selectedBlock); if (b && !isVisible(b)) requestAnimationFrame(() => scrollTo(b)); }

  const refreshSide = () => { clear(side); renderSide(c, tab, blocks, side, persist); };
  refreshSide(); refreshSelbar(); slideIn(side);
  c.changed.clear();

  // ---- footer ----
  const isDone = (t: Tab) => t.layoutApproved || (isReadMe(t.name) && t.blocks.length === 0);
  const allApproved = p.tabs.every(isDone);
  const pending = p.tabs.filter((t) => !isDone(t)).length;
  const openNotes = p.notes.filter((n) => n.status === "open").length;
  // Layout warnings only: an open instruction has its own count, and the missing takeaway belongs to the brief card.
  const flagged = c.state!.warnings.filter((w) => w.kind !== "open-note" && w.kind !== "no-takeaway" && !!w.tabId);
  const totalWarn = flagged.length;
  const approveAll = h("button", { class: "primary", onclick: (e: Event) => {
    if (!totalWarn) { void c.approve("all_layouts"); return; }
    popover(e.currentTarget as HTMLElement, (close) => h("div", { class: "confirm" },
      h("div", { class: "t" }, `${totalWarn} still flagged`),
      h("ul", { class: "warnlist" }, ...flagged.slice(0, 5).map((w) => h("li", {}, h("span", { class: `dot ${w.kind === "overlap" ? "bad" : "warn"}` }), h("span", {}, `${w.tabName}: ${w.message}`)))),
      h("div", { class: "acts" }, h("button", { onclick: close }, "Cancel"), h("button", { class: "primary", onclick: () => { close(); void c.approve("all_layouts"); } }, "Approve anyway"))), { align: "right" });
  } }, "Approve all layouts");
  root.append(h("div", { class: "ftr" },
    h("span", { class: "msg" }, allApproved ? "Layouts approved" : `${p.tabs.length - pending} of ${p.tabs.length} approved${openNotes ? ` · ${openNotes} open note${openNotes === 1 ? "" : "s"}` : ""}${totalWarn ? ` · ${totalWarn} flagged` : ""}`),
    h("span", { class: "spacer" }),
    cover ? null : tab.layoutApproved
      ? h("button", { onclick: () => void c.approve("layout", tab.id, false) }, `Unapprove "${tab.name}"`)
      : h("button", { onclick: () => void c.approve("layout", tab.id) }, `Approve "${tab.name}"`),
    allApproved ? null : approveAll));
}

const SEL_ICONS: Record<string, string> = {
  alignleft: '<path d="M3 2v12"/><rect x="5" y="4" width="8" height="3" rx="0.8"/><rect x="5" y="9" width="5" height="3" rx="0.8"/>',
  aligntop: '<path d="M2 3h12"/><rect x="4" y="5" width="3" height="8" rx="0.8"/><rect x="9" y="5" width="3" height="5" rx="0.8"/>',
  samew: '<rect x="3" y="3" width="10" height="3.5" rx="0.8"/><rect x="3" y="9.5" width="10" height="3.5" rx="0.8"/>',
  sameh: '<rect x="3" y="3" width="3.5" height="10" rx="0.8"/><rect x="9.5" y="3" width="3.5" height="10" rx="0.8"/>',
  distv: '<path d="M3 2h10M3 14h10"/><rect x="5" y="4.5" width="6" height="2.5" rx="0.6"/><rect x="5" y="9" width="6" height="2.5" rx="0.6"/>',
  disth: '<path d="M2 3v10M14 3v10"/><rect x="4.5" y="5" width="2.5" height="6" rx="0.6"/><rect x="9" y="5" width="2.5" height="6" rx="0.6"/>',
};
/** The align and distribute glyphs; the selection bar in both canvases draws from this set. */
export function selIcon(name: string, size = 11): SVGElement {
  const el = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  el.setAttribute("width", String(size)); el.setAttribute("height", String(size)); el.setAttribute("viewBox", "0 0 16 16");
  el.setAttribute("fill", "none"); el.setAttribute("stroke", "currentColor"); el.setAttribute("stroke-width", "1.6"); el.setAttribute("stroke-linecap", "round"); el.setAttribute("stroke-linejoin", "round");
  el.innerHTML = SEL_ICONS[name] ?? ""; return el;
}

function renderBlock(b: Block, notes: Note[]): HTMLElement {
  const type = b.type as WorkbookBlockType;
  // Headers one step down rather than truncated: a table or calc narrower than its own headers need.
  const tight = (b.type === "table" || b.type === "calc") && !!b.columns?.length && !headersFit(b.columns, b.cols);
  const body = h("div", { class: `bc${tight ? " tight" : ""}` }, ...blockBody(b));
  body.style.setProperty("--lblw", `${labelPct(type, b.columns)}%`);
  return h("div", { class: `blk ${b.type}`, "data-id": b.id },
    h("div", { class: "bh" }, h("span", { class: `ty ${b.type}` }, typeLabel(b.type)), h("span", { class: "t", title: `${b.title} at ${addr(b.row, b.col)}` }, b.title),
      ...blockBadges(notes),
      b.detected ? h("span", { class: `badge det${b.detected.confidence < 0.6 ? " low" : ""}`, title: `Detected at ${Math.round(b.detected.confidence * 100)}%: ${b.detected.signals.join(", ")}` }, `${Math.round(b.detected.confidence * 100)}%`) : null),
    body,
    ...resizeHandles());
}

function blockBody(b: Block): Node[] {
  const row = (cells: { text: string; cls?: string }[], header = false) => h("div", { class: `r${header ? " h" : ""}` }, ...cells.map((c) => h("div", { class: c.cls ?? "" }, c.text)));
  /** A guess is muted and dotted (.ph); a real sample takes the column's number format. */
  const cell = (header: string, i: number, raw: string | number | undefined, cls: string) => {
    const v = valueCell(header, i, raw);
    return { text: v.text, cls: v.ph ? `${cls} ph`.trim() : cls };
  };
  if (b.type === "inputs" || b.type === "output") {
    const labels = b.rowLabels?.length ? b.rowLabels : [];
    const n = Math.max(1, b.rows - 1);
    return Array.from({ length: n }, (_, i) => { const l = labels[i]; if (!l) return row([{ text: "", cls: "e" }, { text: "", cls: "e" }]); return row([{ text: l }, cell(l, i, b.sample?.[i]?.[0], b.type === "inputs" ? "v" : "f")]); });
  }
  if (b.type === "table" || b.type === "calc") {
    const cols = b.columns?.length ? b.columns : ["No columns"];
    const n = Math.max(1, b.rows - 2);
    const hasSample = !!b.sample?.length;
    const out: Node[] = [row(cols.map((c) => ({ text: c })), true)];
    for (let r = 0; r < n; r++) {
      const sampleRow = b.sample?.[r]; const lbl = b.rowLabels?.[r];
      const fill = hasSample ? !!sampleRow : (!!lbl || r < 3);
      out.push(row(cols.map((col, ci) => {
        if (!fill) return { text: "", cls: "e" };
        if (ci === 0) return lbl ? { text: lbl } : cell(col, r, sampleRow?.[0], "");
        return cell(col, r, sampleRow?.[ci], b.type === "calc" ? "f" : "v");
      })));
    }
    return out;
  }
  if (b.type === "chart") return [chartSketch(b.chartKind ?? "bar")];
  return [h("div", {}, b.description || "Written at build")];
}

export function chartSketch(kind: "bar" | "line" | "pie"): SVGElement {
  const vals = [35, 52, 44, 70, 62, 85];
  if (kind === "pie") {
    const total = vals.reduce((a, b) => a + b, 0); let acc = 0;
    const slices = vals.map((v, i) => { const a0 = (acc / total) * Math.PI * 2; acc += v; const a1 = (acc / total) * Math.PI * 2; const x0 = 50 + 26 * Math.cos(a0), y0 = 30 + 26 * Math.sin(a0), x1 = 50 + 26 * Math.cos(a1), y1 = 30 + 26 * Math.sin(a1); return `<path d="M50 30 L${x0.toFixed(1)} ${y0.toFixed(1)} A26 26 0 ${a1 - a0 > Math.PI ? 1 : 0} 1 ${x1.toFixed(1)} ${y1.toFixed(1)} Z" fill="var(--chart)" opacity="${0.3 + i * 0.12}"/>`; }).join("");
    const el = svg(slices); el.setAttribute("preserveAspectRatio", "xMidYMid meet"); return el;
  }
  if (kind === "line") {
    const pts = vals.map((v, i) => `${10 + i * 16},${58 - v * 0.6}`).join(" ");
    return svg(`<line x1="4" y1="58" x2="98" y2="58" stroke="currentColor" opacity="0.3"/><polyline points="${pts}" fill="none" stroke="var(--chart)" stroke-width="1.8"/><polyline points="${vals.map((v, i) => `${10 + i * 16},${62 - v * 0.45}`).join(" ")}" fill="none" stroke="var(--chart)" stroke-width="1.2" opacity="0.45"/>${vals.map((v, i) => `<circle cx="${10 + i * 16}" cy="${58 - v * 0.6}" r="1.6" fill="var(--chart)"/>`).join("")}`);
  }
  const bars = vals.map((v, i) => `<rect x="${8 + i * 15}" y="${58 - v * 0.6}" width="10" height="${v * 0.6}" fill="var(--chart)" opacity="${0.3 + i * 0.1}"/>`).join("");
  return svg(`<line x1="4" y1="58" x2="98" y2="58" stroke="currentColor" opacity="0.3"/>${bars}`);
}

function chipsEditor(items: string[], placeholder: string, onChange: (next: string[]) => void): HTMLElement {
  const wrap = h("div", { class: "chips" });
  const draw = () => {
    clear(wrap);
    items.forEach((it, i) => wrap.append(h("div", { class: "c" }, h("span", { title: it }, it), h("button", { title: "Remove", onclick: () => { items.splice(i, 1); onChange([...items]); draw(); } }, icon("x", 10)))));
    const add = h("button", { class: "add", onclick: () => {
      const input = h("input", { placeholder });
      let done = false; // guard: Enter commits and re-renders, which blurs the detached input
      const commit = (again: boolean) => {
        if (done) return; done = true;
        const v = input.value.trim();
        if (v) { items.push(v); onChange([...items]); }
        draw();
        if (again && v) (wrap.querySelector("button.add") as HTMLButtonElement)?.click();
      };
      input.onkeydown = (e) => { if (e.key === "Enter") { e.preventDefault(); commit(true); } if (e.key === "Escape") { done = true; draw(); } };
      input.onblur = () => commit(false);
      add.replaceWith(input); input.focus();
    } }, `+ ${placeholder}`);
    wrap.append(add);
  };
  draw();
  return wrap;
}

function renderSide(c: Controller, tab: Tab, blocks: Block[], side: HTMLElement, persist: (opts?: { record?: boolean; resolve?: Block }) => void) {
  const b = c.block;
  const p = c.project;
  if (!b) {
    const warns = c.state!.warnings.filter((w) => w.tabId === tab.id);
    const body = h("div", { class: "scroll compact", "data-scroll": "side" },
      h("div", { class: "field" }, h("span", { class: "label" }, "Readiness"), warns.length ? readinessList(warns, (id) => { c.selectedBlock = id; c.setState(c.state!); }) : h("div", { class: "ok-text" }, "Nothing flagged")),
      h("div", { class: "field" }, h("span", { class: "label" }, "Blocks"), h("div", { class: "blocklist" }, ...blocks.map((x) => h("div", { class: "b", onclick: () => { c.selectedBlock = x.id; c.setState(c.state!); } }, h("span", { class: "addr" }, addr(x.row, x.col)), h("span", { class: "t" }, x.title), (() => { const oq = p.notes.filter((n) => n.blockId === x.id && n.status === "open" && n.kind === "question").length; return oq ? h("span", { class: "badge q" }, `${oq}?`) : null; })())), blocks.length ? null : h("span", { class: "muted small-text" }, isReadMe(tab.name) ? "Generated at build" : "No blocks yet"))),
      h("div", { class: "field" }, h("span", { class: "label" }, "Thread"), notesPanel(c, tab.id, undefined, c.notesFor(tab.id))));
    side.append(h("div", { class: "sh" }, h("span", { class: "label" }, "Tab"), h("span", { class: "t" }, tab.name)), body);
    fadeIn(body);
    return;
  }
  const notes = c.notesFor(tab.id, b.id);
  const scroll = h("div", { class: "scroll compact", "data-scroll": "side" },
    h("div", { class: "field" }, h("span", { class: "label" }, "Title"), h("input", { class: "title", value: b.title, spellcheck: "false", onchange: (e: Event) => { b.title = (e.target as HTMLInputElement).value.trim() || b.title; persist(); }, onkeydown: (e: KeyboardEvent) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); } })),
    h("div", { class: "field" }, h("span", { class: "label" }, "Type"), h("select", { onchange: (e: Event) => { b.type = (e.target as HTMLSelectElement).value as WorkbookBlockType; persist(); } }, ...(Object.keys(TYPE_LABEL) as WorkbookBlockType[]).map((t) => h("option", { value: t, selected: t === b.type }, `${TYPE_LABEL[t]}: ${TYPE_DESC[t]}`)))),
    b.type === "table" || b.type === "calc" ? h("div", { class: "field" }, h("span", { class: "label" }, "Columns"), chipsEditor([...(b.columns ?? [])], "column", (next) => { b.columns = next.length ? next : undefined; b.cols = Math.max(b.cols, fitCols(next)); persist(); })) : null,
    b.type === "table" || b.type === "calc" ? h("div", { class: "field" }, h("span", { class: "label" }, "Row labels"), chipsEditor([...(b.rowLabels ?? [])], "row", (next) => { b.rowLabels = next.length ? next : undefined; b.rows = Math.max(b.rows, next.length + 2); persist(); })) : null,
    b.type === "inputs" || b.type === "output" ? h("div", { class: "field" }, h("span", { class: "label" }, "Line items"), chipsEditor([...(b.rowLabels ?? [])], "line item", (next) => { b.rowLabels = next.length ? next : undefined; b.rows = Math.max(b.rows, next.length + 1); persist(); })) : null,
    b.type === "chart" ? h("div", { class: "field" }, h("span", { class: "label" }, "Kind"), h("select", { onchange: (e: Event) => { b.chartKind = (e.target as HTMLSelectElement).value as Block["chartKind"]; persist(); } }, ...(["bar", "line", "pie"] as const).map((k) => h("option", { value: k, selected: (b.chartKind ?? "bar") === k }, k)))) : null,
    b.type === "chart" ? h("div", { class: "field" }, h("span", { class: "label" }, "Plots"), h("select", { onchange: (e: Event) => { b.source = (e.target as HTMLSelectElement).value || undefined; persist(); } }, h("option", { value: "" }, "Choose a table"), ...blocks.filter((x) => x.type === "table" || x.type === "calc").map((x) => h("option", { value: x.id, selected: x.id === b.source }, x.title)))) : null,
    b.detected ? h("div", { class: "field" }, h("span", { class: "label" }, "Detected"), h("div", { class: "card" }, h("div", { class: "inline" }, h("span", { class: `chip ${b.detected.confidence < 0.6 ? "warn" : "ok"}` }, h("span", { class: `dot ${b.detected.confidence < 0.6 ? "warn" : "ok"}` }), `${Math.round(b.detected.confidence * 100)}% ${typeLabel(b.type).toLowerCase()}`), h("span", { class: "mono muted xs" }, b.detected.range), h("span", { class: "spacer" }), h("button", { class: "small", title: "Clear the badge", onclick: () => { b.detected = undefined; persist(); } }, "Looks right")), h("div", { class: "ct muted" }, b.detected.signals.join(" · ")))) : null,
    h("div", { class: "field" }, h("span", { class: "label" }, "Description"), h("textarea", { rows: 2, value: b.description ?? "", placeholder: "What it does", onchange: (e: Event) => { b.description = (e.target as HTMLTextAreaElement).value.trim() || undefined; persist(); } })),
    h("div", { class: "field" }, h("span", { class: "label" }, "Thread"), notesPanel(c, tab.id, b.id, notes)),
    (() => {
      const above = blockAbove(blocks, b);
      const dataRows = b.rows - (b.type === "table" || b.type === "calc" ? 2 : 1);
      const splitAt = h("input", { class: "w-num", type: "number", min: "1", max: String(Math.max(1, dataRows - 1)), value: String(Math.max(1, Math.floor(dataRows / 2))), title: "Split before row" });
      return h("div", { class: "field" }, h("span", { class: "label" }, "Reshape"), h("div", { class: "inline wrap tight" },
        h("button", { class: "small", disabled: !above, title: above ? `Merge with "${above.title}"` : "Nothing above", onclick: () => { if (!above) return; const m = mergeBlocks(blocks, above, b); c.selectedBlock = m.id; persist({ resolve: m }); } }, "Merge up"),
        h("button", { class: "small", disabled: dataRows < 2, onclick: () => { const second = splitBlock(blocks, b, Number(splitAt.value) || 1); if (second) { c.selectedBlock = second.id; persist({ resolve: second }); } } }, "Split"), splitAt));
    })(),
    h("div", { class: "inline tight" },
      h("button", { onclick: () => { const copy: Block = { ...b, id: uid("blk"), title: `${b.title} copy`, row: b.row + b.rows + 1 }; blocks.push(copy); c.selectedBlock = copy.id; persist({ resolve: copy }); } }, icon("copy", 12), "Duplicate"),
      twoStep(h("span", {}, "Remove"), "Remove block?", () => { const i = blocks.indexOf(b); if (i >= 0) { blocks.splice(i, 1); c.selectedBlock = null; p.notes = p.notes.filter((n) => n.blockId !== b.id); persist(); } }, "danger")));
  side.append(h("div", { class: "sh" }, h("span", { class: `tybox ty ${b.type}` }, typeLabel(b.type)), h("span", { class: "t" }, b.title), h("span", { class: "mono muted xs" }, `${addr(b.row, b.col)} · ${b.rows}×${b.cols}`)), scroll);
  fadeIn(scroll);
}

const SHORTCUTS: [string, string][] = [["⌘Z / ⇧⌘Z", "Undo, redo"], ["↑ ↓ ← →", "Nudge one cell"], ["⇧ + arrows", "Resize"], ["⌘D", "Duplicate"], ["⌫", "Remove"], ["Esc", "Deselect"], ["1 to 9", "Switch tab"], ["⌘+ ⌘- ⌘0", "Zoom"], ["Click", "Select"], ["⇧ click", "Add to selection"], ["Drag grid", "Select by area"], ["Drag header", "Move"], ["Drag corner", "Resize"], ["Double-click", "Rename a block"], ["⌘↩ in a note", "Add note"], ["⌘↑ ⌘↓", "Select row, Tabs"], ["⌥↑ ⌥↓", "Move row, Tabs"], ["A", "Arrows, built view"], ["?", "Shortcuts"]];
/** The `?` sheet. The deck canvas passes its own rows, which name Remix and the slide keys. */
export function openCheatSheet(root: HTMLElement, rows: [string, string][] = SHORTCUTS) {
  document.querySelectorAll(".pop").forEach((x) => x.remove());
  const pop = h("div", { class: "pop cheats" }, h("div", { class: "label pophdr" }, "Keyboard"), ...rows.map(([k, d]) => h("div", { class: "cheat" }, h("span", { class: "key" }, k), h("span", {}, d))));
  pop.style.left = "50%"; pop.style.top = "60px"; pop.style.transform = "translateX(-50%)";
  root.append(pop);
  const close = () => { pop.remove(); document.removeEventListener("pointerdown", onDoc, true); document.removeEventListener("keydown", onKey, true); };
  const onDoc = (e: PointerEvent) => { if (!pop.contains(e.target as Node)) close(); };
  const onKey = (e: KeyboardEvent) => { if (e.key === "Escape" || e.key === "?") { e.stopPropagation(); close(); } };
  setTimeout(() => { document.addEventListener("pointerdown", onDoc, true); document.addEventListener("keydown", onKey, true); }, 0);
}

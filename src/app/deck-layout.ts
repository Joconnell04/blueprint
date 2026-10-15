/**
 * Deck Layer 2: one 16:9 slide on the 30 pt grid (32 columns by 18 rows), drawn at a zoom that fits the
 * panel. Furniture (kicker, headline band, hairline, source line, page number) comes from the selected
 * style's own token JSON, so the sketch sits where the built slide will; without tokens it falls back to
 * the house table in docs/research/deck-planning.md section 6.
 *
 * Rows 1 to 4 are the furniture band and row 18 the footer: a block cannot be dragged, dropped or
 * resized into either (feedback F11). A divider row draws the format's section slide instead of a grid.
 *
 * Everything the workbook canvas already solved is imported from layout-view.ts rather than copied:
 * history, selection, corner resize, the align bar, the minimap, the note badges, the chart sketch,
 * the readiness list and the `?` sheet.
 */
import { h, clear, svg, icon, twoStep, popover } from "./dom.js";
import { fadeIn, pulse, slideIn } from "./motion.js";
import type { Controller } from "./controller.js";
import { fakeValue, overlaps, uid, type Block, type Note, type State, type Tab } from "./model.js";
import {
  DECK_COLS, DECK_PLACEABLE_TYPES, DECK_ROWS,
  type DeckBlockType, type Diagram, type DiagramEdge, type DiagramKind, type DiagramNode,
} from "../shared/types.js";
import {
  HISTORY_MAX, ZOOM_STEPS, blockBadges, chartSketch, createMinimap, fillSelectionBar, historyFor,
  openCheatSheet, readinessList, resizeHandles, selection, snapshotBlocks, wireCornerResize,
} from "./layout-view.js";
import { notesPanel } from "./notes.js";
import { ARCHETYPES, DECK_BODY_BOTTOM, DECK_BODY_TOP, FURNITURE_TYPES, clampToBody, detectArchetype, remix } from "./remix.js";
// Straight from the shared solver, not through remix.ts: the canvas and the server must run the one pass (F-2).
import { resolveBandOverlaps } from "../shared/archetypes.js";

/** A 30 pt cell drawn as this many pixels at zoom 1, so the slide is 832 x 468 and 16:9 exactly. */
export const DECK_CELL = 26;
const PT = DECK_CELL / 30;                 // px per point, for the furniture table and the type scale
const SLIDE_W = DECK_COLS * DECK_CELL, SLIDE_H = DECK_ROWS * DECK_CELL;
const PAD = 20;                            // air around the slide inside the scroll footprint

// ---- the selected style's tokens (backlog D11) ----

/** One furniture box in points on the 960 x 540 slide. */
export interface TokenBox { x: number; y: number; w: number; h?: number; align?: string }
/** One line of the style's type scale: size in points, colour as a palette key or a bare hex. */
export interface TokenType { size?: number; color?: string; bold?: boolean; caps?: boolean }
/** As much of builder/styles/*.json as the canvas draws. Everything is optional: extracted styles are partial. */
export interface DeckTokens {
  palette?: Record<string, string | string[]>;
  fonts?: Record<string, string>;
  typeScale?: Record<string, TokenType>;
  furniture?: Record<string, TokenBox>;
  backgrounds?: Record<string, Record<string, string>>;
  kicker?: { pattern?: string };
  divider?: { numbered?: boolean };
  rules?: Record<string, boolean>;
}
export interface DeckStyle { id?: string; kind?: string; tokens?: DeckTokens }

/** The style on the state, or the copy the project carries; either way the tokens are the builder's own. */
export function deckStyle(c: Controller): DeckStyle | undefined {
  const fromState = (c.state as State & { style?: DeckStyle }).style;
  return fromState ?? (c.project as { style?: DeckStyle }).style;
}

/** House furniture, in points (research doc section 6). The fallback whenever the style names no box. */
const HOUSE: Record<string, TokenBox> = {
  kicker: { x: 60, y: 42, w: 600, h: 14 },
  headline: { x: 60, y: 66, w: 840, h: 84 },
  hairline: { x: 60, y: 156, w: 840, h: 1 },
  source: { x: 60, y: 514, w: 600, h: 14 },
  pageNumber: { x: 870, y: 514, w: 60, h: 14, align: "right" },
  dividerTitle: { x: 60, y: 180, w: 840, h: 150 },
};
/** The box for one piece of furniture: the style's, else the house one. */
export function furnitureBox(t: DeckTokens | undefined, name: keyof typeof HOUSE | string): TokenBox {
  const box = t?.furniture?.[name];
  return box && Number.isFinite(box.x) && Number.isFinite(box.y) ? box : HOUSE[name] ?? HOUSE.headline;
}
/** Point geometry to slide pixels. Furniture lives inside the slide, so the canvas padding is not in here. */
const box = (f: TokenBox) => ({
  left: `${f.x * PT}px`, top: `${f.y * PT}px`, width: `${f.w * PT}px`, ...(f.h ? { height: `${f.h * PT}px` } : {}),
});

const HEX = /^#?[0-9a-fA-F]{6}$/;
const hex = (v: unknown): string | undefined => (typeof v === "string" && HEX.test(v.trim()) ? `#${v.trim().replace("#", "")}` : undefined);
/**
 * One colour by name: the slide kind's own background override first (a divider inverts the ink), then
 * the palette, then the name read as a bare hex, which is what an extracted style writes.
 */
export function tokenColour(t: DeckTokens | undefined, bg: Record<string, string> | undefined, name?: string): string | undefined {
  if (!name) return undefined;
  const pal = t?.palette?.[name];
  return hex(bg?.[name]) ?? hex(typeof pal === "string" ? pal : undefined) ?? hex(name);
}

/**
 * The CSS custom properties one slide draws from: its background, its ink and one entry per furniture
 * role. Only what the tokens give is returned, so deck.css keeps the canvas's own look without them.
 */
export function slideVars(t: DeckTokens | undefined, kind: "content" | "divider"): Record<string, string> {
  const out: Record<string, string> = {};
  if (!t) return out;
  const bg = t.backgrounds?.[kind];
  const set = (k: string, v?: string) => { if (v) out[k] = v; };
  set("--sl-bg", hex(bg?.fill) ?? tokenColour(t, undefined, "bg"));
  set("--sl-ink", tokenColour(t, bg, "ink"));
  set("--sl-muted", tokenColour(t, bg, "muted"));
  set("--sl-accent", tokenColour(t, bg, "accent"));
  set("--sl-line", tokenColour(t, bg, "line"));
  const font = (name: string) => { const f = t.fonts?.[name]; return f ? `"${f}", var(--font)` : undefined; };
  set("--sl-display", font("display"));
  set("--sl-body", font("body") ?? font("display"));
  // one colour and one size per role; a size in points becomes pixels at the canvas's own scale
  const role = (prop: string, ...names: string[]) => {
    const ts = names.map((n) => t.typeScale?.[n]).find(Boolean);
    if (!ts) return;
    set(`--c-${prop}`, tokenColour(t, bg, ts.color));
    if (ts.size) set(`--fs-${prop}`, `${Math.round(ts.size * PT * 10) / 10}px`);
  };
  role("kicker", "kicker");
  role("headline", "headline");
  role("source", "source", "caption");
  role("page", "pageNumber", "caption");
  role("dtitle", "dividerTitle", "coverTitle");   // no headline fallback: a section title is not a headline
  return out;
}

/**
 * The kicker line: the style's own pattern with the section and the slide's name in it, an empty slot's
 * separator dropped. The section is the nearest divider above, which the server derives on every save.
 */
export function kickerText(t: DeckTokens | undefined, section: string | undefined, topic: string): string {
  const pattern = t?.kicker?.pattern ?? "{section} | {topic}";
  return pattern.replace("{section}", section?.trim() ?? "").replace("{topic}", topic.trim())
    .replace(/^[\s|·–—]+/, "").replace(/[\s|·–—]+$/, "").replace(/\s{2,}/g, " ");
}

/** Default sizes, mirroring DECK_SIZE in src/server/store.ts so a block added here matches one added in chat. */
export const DECK_DEFAULT_SIZE: Record<DeckBlockType, [number, number]> = {
  headline: [2, 32], kicker: [1, 16], text: [8, 14], bullets: [8, 14], cards: [7, 32], stat: [4, 8],
  chart: [10, 20], diagram: [9, 20], table: [9, 32], image: [10, 14], quote: [5, 24], takeaway: [2, 32],
  placeholder: [8, 14], notes: [3, 32],
};
/** One word per type: the palette, the block tag and the inspector all read from here. */
export const DECK_LABEL: Record<DeckBlockType, string> = {
  headline: "Headline", kicker: "Kicker", text: "Text", bullets: "Bullets", cards: "Cards", stat: "Stat",
  chart: "Chart", diagram: "Diagram", table: "Table", image: "Image", quote: "Quote", takeaway: "Takeaway",
  placeholder: "Placeholder", notes: "Notes",
};
export const deckLabel = (t: Block["type"]): string => DECK_LABEL[t as DeckBlockType] ?? String(t);
/** The palette: every type a user places by hand. The headline is the slide's own field, never a block. */
const PALETTE = DECK_PLACEABLE_TYPES;
const COUNTED = new Set<Block["type"]>(["cards", "stat"]);

const SHORTCUTS: [string, string][] = [["⌥R", "Remix"], ["⌘Z / ⇧⌘Z", "Undo, redo"], ["⌘← ⌘→", "Previous, next slide"], ["↑ ↓ ← →", "Nudge one cell"], ["⇧ + arrows", "Resize"], ["⌘D", "Duplicate"], ["⌫", "Remove"], ["Esc", "Deselect"], ["1 to 9", "Switch slide"], ["⌘+ ⌘- ⌘0", "Zoom, fit"], ["Click", "Select"], ["⇧ click", "Add to selection"], ["Drag slide", "Select by area"], ["Drag header", "Move"], ["Drag corner", "Resize"], ["Click headline", "Edit headline"], ["⌘↩ in a note", "Add note"], ["⌘↑ ⌘↓", "Select row, Slides"], ["⌥↑ ⌥↓", "Move row, Slides"], ["?", "Shortcuts"]];

/**
 * Save one slide's headline. It patches the one field through ui_save_tab_fields, which leaves every
 * approval alone: editing a headline in Layer 2 must not send the user back to Layer 1.
 */
function saveHeadline(c: Controller, tab: Tab, text: string): void {
  tab.message = text || undefined;
  c.saveMessage(tab.id, text);
}

/** Keep a block inside the body band: rows 5 to 17, 1-based, never wider than the slide. */
export function fitToSlide(b: Block): Block {
  clampToBody([b]);
  return b;
}

/**
 * Push overlapping blocks apart, by the server's own rule (F-2). The sheet canvas keeps the block the
 * user dragged and moves everything under it; the server moves whichever block is later in reading order
 * and re-runs the pass on every save. Those two disagreed, so a drop previewed one slide and saved
 * another. This is `resolveBandOverlaps` itself: later block first down, then right, and onto the last
 * free rows when the body band has room for neither.
 */
export function resolveDeckOverlaps(blocks: Block[]): boolean {
  return resolveBandOverlaps(blocks);
}

// ---- the column module (docs/research/output-quality.md, ranked change 1) ----

/** A slide cell in points: the plan's own grid, 960 x 540 over 32 by 18. */
const CELL_PT = 30;
/** How many modules the body is divided into. Twelve divides by 2, 3, 4 and 6, which is every deck band. */
export const MODULE_COLS = 12;
/** One cell of air between modules. */
const MODULE_GUTTER = 1;

export interface ColumnModule {
  /** module width and gutter in cells; a module is rarely a whole number of them */
  unit: number;
  gutter: number;
  /** the module left edges, as 1-based columns */
  starts: number[];
  /** the module right edges, as 1-based inclusive columns */
  ends: number[];
}

/**
 * The 12-column module the style's body area implies: the body width less eleven one-cell gutters, split
 * twelve ways, every edge expressed back in 30 pt cells and rounded. Without tokens the body is the grid's
 * own, two cells of margin either side, which is where the house furniture and every archetype already sit.
 */
export function columnModule(tokens?: DeckTokens, count = MODULE_COLS): ColumnModule {
  const named = tokens?.furniture?.body;
  const body = named && Number.isFinite(named.x) && Number.isFinite(named.w) && named.w > 0
    ? { x: named.x, w: named.w }
    : { x: 2 * CELL_PT, w: (DECK_COLS - 4) * CELL_PT };
  const gutter = MODULE_GUTTER * CELL_PT;
  const unit = Math.max(CELL_PT / 2, (body.w - gutter * (count - 1)) / count);
  const starts: number[] = [], ends: number[] = [];
  for (let i = 0; i < count; i++) {
    const left = body.x + i * (unit + gutter);
    const col = Math.min(DECK_COLS, Math.max(1, Math.round(left / CELL_PT) + 1));
    starts.push(col);
    ends.push(Math.min(DECK_COLS, Math.max(col, Math.round((left + unit) / CELL_PT))));
  }
  return { unit: unit / CELL_PT, gutter: MODULE_GUTTER, starts, ends };
}

/** The nearest of `xs` to `v`, lower one on a tie, and how far off it was. */
function nearest(xs: number[], v: number): { at: number; off: number } {
  let at = v, off = Infinity;
  for (const x of xs) { const d = Math.abs(x - v); if (d < off) { at = x; off = d; } }
  return { at, off };
}

/** Which edge the gesture moved, so the snap moves that one and leaves the other where the user put it. */
export type SnapEdge = "move" | "left" | "right" | "both";

/**
 * Snap one block onto the module, in cells and nothing else. An edge only moves when it is within half a
 * module (the module and its gutter) of one: further out is where the user meant it. `how` says what the
 * gesture was: "move" slides the block so its left edge lands on a module edge and keeps the width, "left"
 * and "right" pin the opposite edge (one corner of a resize), "both" takes the nearest edge on each side.
 */
export function snapToModule(b: { col: number; cols: number }, m: ColumnModule, how: SnapEdge = "both"): { col: number; cols: number } {
  const tol = (m.unit + m.gutter) / 2;
  const right = b.col + b.cols - 1;
  const l = nearest(m.starts, b.col), r = nearest(m.ends, right);
  const takeLeft = how !== "right" && l.off <= tol;
  const takeRight = (how === "right" || how === "both") && r.off <= tol;
  if (how === "move") {
    // a slide that would push the block off the right edge is not a snap; the block keeps the cell it landed on
    const room = takeLeft && l.at >= 1 && l.at + b.cols - 1 <= DECK_COLS;
    return { col: room ? l.at : b.col, cols: b.cols };
  }
  const col = takeLeft ? l.at : b.col;
  const end = takeRight ? r.at : right;
  return { col, cols: Math.max(1, Math.min(DECK_COLS - col + 1, end - col + 1)) };
}

// ---- Fit: the rows a block's content asks for (ranked change 9) ----

const BODY_PT = 14;        // the deck builder's own default body size
const LINE = 1.16;         // leading, the way the builder measures a line
const CHAR_PT = 0.5;       // characters per point of width (builder/SPEC-deck.md, "Text fit")
const TITLE_PT = 16;       // the block's own label, above its content
const CARD_HEAD_PT = 34;   // a card's label and its sub line
const CARD_LINES = 3;      // a card always has room for a short paragraph
const TABLE_ROW_PT = 24;   // one table row, the header included
const TABLE_ROWS = 4;      // data rows assumed when the plan names none
const STAT_ROWS = 4;       // the number and its caption
const BAND_ROWS = DECK_BODY_BOTTOM - DECK_BODY_TOP + 1;

/** Lines a description needs at a width, by the deck builder's own measure. */
export function contentLines(text: string | undefined, widthPt: number, sizePt = BODY_PT): number {
  const per = Math.max(3, Math.floor(widthPt / (sizePt * CHAR_PT)));
  return Math.max(1, Math.ceil((text ?? "").trim().length / per));
}

const rowsFor = (pt: number) => Math.max(1, Math.ceil(pt / CELL_PT));

/**
 * The rows this block's content asks for, or null when its height is the user's call. A chart, a diagram
 * and an image are drawn to their box, so Fit leaves them alone; everything else is text, and the estimate
 * is the description the plan carries, wrapped at the block's own width.
 */
export function fitRows(b: Block): number | null {
  const widthPt = Math.max(CELL_PT, b.cols * CELL_PT);
  switch (b.type) {
    case "chart": case "diagram": case "image": return null;
    case "stat": return STAT_ROWS;
    case "table": {
      const data = b.rowLabels?.length || b.sample?.length || TABLE_ROWS;
      return rowsFor((data + 1) * TABLE_ROW_PT);
    }
    case "cards": {
      const n = Math.max(1, Math.min(6, b.count ?? 3));
      const each = Math.max(CELL_PT, (widthPt - (n - 1) * CELL_PT) / n);
      const lines = Math.max(CARD_LINES, contentLines(b.content ?? b.title, each));
      return rowsFor(CARD_HEAD_PT + lines * BODY_PT * LINE);
    }
    default: {
      const lines = contentLines(b.content ?? b.title, widthPt);
      return rowsFor(TITLE_PT + lines * BODY_PT * LINE);
    }
  }
}

/** Fit every block to what it holds. Columns and rows stay; only the height changes. True when any did. */
export function fitBlocks(blocks: Block[]): boolean {
  let changed = false;
  for (const b of blocks) {
    const want = fitRows(b);
    if (want == null) continue;
    const rows = Math.max(1, Math.min(BAND_ROWS, want));
    if (rows !== b.rows) { b.rows = rows; changed = true; }
  }
  return changed;
}

// ---- Balance: equal gaps down the body band (ranked change 8) ----

/** Blocks that share rows are one band: they sit side by side, so they move together. */
export function rowBands(blocks: Block[]): Block[][] {
  const bands: Block[][] = [];
  let bottom = -1;
  for (const b of [...blocks].sort((x, y) => x.row - y.row || x.col - y.col)) {
    if (bands.length && b.row <= bottom) bands[bands.length - 1].push(b);
    else { bands.push([b]); bottom = -1; }
    bottom = Math.max(bottom, b.row + b.rows - 1);
  }
  return bands;
}

const bandTop = (band: Block[]) => Math.min(...band.map((b) => b.row));
const bandRows = (band: Block[]) => Math.max(...band.map((b) => b.row + b.rows - 1)) - bandTop(band) + 1;

/**
 * Balance: the same gap between every band down rows 5 to 17, and the group centred when it does not fill
 * the band. The gap is the one the slide already has, averaged and at least one row, tightened only when
 * the group would not fit. Order and columns are kept, and a band moves as one, so a staggered pair stays
 * staggered. True when anything moved.
 */
export function balanceBlocks(blocks: Block[]): boolean {
  const bands = rowBands(blocks);
  if (!bands.length) return false;
  const heights = bands.map(bandRows);
  const total = heights.reduce((a, b) => a + b, 0);
  const gaps = bands.length - 1;
  let gap = 0;
  if (gaps) {
    let sum = 0;
    for (let i = 1; i < bands.length; i++) sum += bandTop(bands[i]) - (bandTop(bands[i - 1]) + heights[i - 1]);
    gap = Math.max(1, Math.round(sum / gaps));
    while (gap > 0 && total + gap * gaps > BAND_ROWS) gap -= 1;   // bands abut before any of them leaves the band
  }
  let top = DECK_BODY_TOP + Math.max(0, Math.floor((BAND_ROWS - (total + gap * gaps)) / 2));
  let moved = false;
  bands.forEach((band, i) => {
    const d = top - bandTop(band);
    if (d) { for (const b of band) b.row += d; moved = true; }
    top += heights[i] + gap;
  });
  return moved;
}

/** The fallback when the host's sandbox refuses storage: the toggle still works for this canvas's life. */
const snapMemory = new Map<string, boolean>();
/** Snap is on until the user turns it off, and the answer is remembered per project. */
export function snapRemembered(projectId: string): boolean {
  try { return localStorage.getItem(`wp.snap.${projectId}`) !== "0"; } catch { return snapMemory.get(projectId) !== false; }
}
export function rememberSnap(projectId: string, on: boolean): void {
  snapMemory.set(projectId, on);
  try { localStorage.setItem(`wp.snap.${projectId}`, on ? "1" : "0"); } catch { /* sandboxed host */ }
}

// ---- sketches: what a block will hold, never the content itself ----

const WIDTHS = [96, 88, 92, 74, 62, 84];
const line = (w: number, cls = "") => h("div", { class: `ln${cls ? ` ${cls}` : ""}`, style: { width: `${w}%` } });
const lines = (n: number, cls = "") => h("div", { class: "lines" }, ...Array.from({ length: n }, (_, i) => line(WIDTHS[i % WIDTHS.length], cls)));

function tableSketch(b: Block): HTMLElement {
  const cols = b.columns?.length ? b.columns.slice(0, 8) : ["", "", "", ""];
  const rows = Math.max(2, Math.min(8, b.rows - 2));
  return h("div", { class: "tbl" },
    h("div", { class: "tr th" }, ...cols.map((c) => h("div", { class: "tc" }, c || line(62, "thin")))),
    ...Array.from({ length: rows }, () => h("div", { class: "tr" }, ...cols.map(() => h("div", { class: "tc" }, line(70, "thin"))))));
}

// ---- the diagram block (F18) ----

export const DIAGRAM_KINDS: DiagramKind[] = ["flow", "cycle", "timeline", "tree", "matrix"];
/** Drawn when the block has no nodes yet: the shape of the thing, with nothing claimed about it. */
const GENERIC: Diagram = {
  kind: "flow", direction: "LR",
  nodes: [1, 2, 3, 4].map((i) => ({ id: `n${i}`, label: "" })),
  edges: [{ from: "n1", to: "n2" }, { from: "n2", to: "n3" }, { from: "n3", to: "n4" }],
};

/** A node id from its label: lower case, one dash per gap. The same label is the same node. */
export function diagramId(label: string): string {
  return label.trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "").slice(0, 24) || "node";
}

/**
 * The Edges field, one line each: `A -> B: label`. A line with no arrow is a node on its own, a chain
 * `A -> B -> C` makes one edge per hop, and a label after the last node belongs to the last hop.
 */
export function parseDiagram(text: string): { nodes: DiagramNode[]; edges: DiagramEdge[] } {
  const nodes: DiagramNode[] = [];
  const edges: DiagramEdge[] = [];
  const add = (label: string): string => {
    const clean = label.trim();
    if (!clean) return "";
    const id = diagramId(clean);
    if (!nodes.some((n) => n.id === id)) nodes.push({ id, label: clean });
    return id;
  };
  for (const raw of text.split(/\r?\n/)) {
    const row = raw.trim();
    if (!row) continue;
    const hops = row.split(/\s*(?:-+>|→)\s*/).filter((s) => s.trim());
    if (hops.length < 2) { add(row); continue; }
    const last = hops[hops.length - 1];
    const at = last.indexOf(":");
    const label = at >= 0 ? last.slice(at + 1).trim() : "";
    hops[hops.length - 1] = at >= 0 ? last.slice(0, at) : last;
    const ids = hops.map(add).filter(Boolean);
    for (let i = 1; i < ids.length; i++) {
      const tail = i === ids.length - 1 && label;
      edges.push(tail ? { from: ids[i - 1], to: ids[i], label } : { from: ids[i - 1], to: ids[i] });
    }
  }
  return { nodes, edges };
}

/** The same lines back, so the field shows what the block holds. Unlinked nodes go last, one per line. */
export function diagramText(d?: Diagram): string {
  if (!d) return "";
  const label = (id: string) => d.nodes.find((n) => n.id === id)?.label ?? id;
  const out = d.edges.map((e) => `${label(e.from)} -> ${label(e.to)}${e.label ? `: ${e.label}` : ""}`);
  const linked = new Set(d.edges.flatMap((e) => [e.from, e.to]));
  for (const n of d.nodes) if (!linked.has(n.id)) out.push(n.label);
  return out.join("\n");
}

const NS = "http://www.w3.org/2000/svg";
function sv<K extends keyof SVGElementTagNameMap>(tag: K, attrs: Record<string, string | number> = {}): SVGElementTagNameMap[K] {
  const e = document.createElementNS(NS, tag);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, typeof v === "number" ? String(Math.round(v * 10) / 10) : v);
  return e;
}
interface DgBox { x: number; y: number; w: number; h: number }
/** Cut a label to what its box holds; the sketch never wraps. */
const cut = (s: string, px: number) => { const max = Math.max(1, Math.floor(px / 4.7)); return s.length > max ? `${s.slice(0, Math.max(1, max - 1))}…` : s; };

function dgNode(n: DiagramNode, r: DgBox): SVGElement {
  const g = sv("g", { class: "nd" });
  g.append(sv("rect", { x: r.x, y: r.y, width: r.w, height: r.h, rx: 3 }));
  if (n.label) {
    const t = sv("text", { x: r.x + r.w / 2, y: r.y + r.h / 2, class: "lb" });
    t.textContent = cut(n.label, r.w - 6);
    g.append(t);
  } else g.append(sv("rect", { x: r.x + 5, y: r.y + r.h / 2 - 1, width: Math.max(4, r.w - 10), height: 2, rx: 1, class: "ph" }));
  return g;
}

const centre = (r: DgBox) => [r.x + r.w / 2, r.y + r.h / 2] as const;
/** Where the line between two centres leaves a box. */
function edgePoint(r: DgBox, dx: number, dy: number): readonly [number, number] {
  const [cx, cy] = centre(r);
  const tx = dx ? (r.w / 2) / Math.abs(dx) : Infinity, ty = dy ? (r.h / 2) / Math.abs(dy) : Infinity;
  const t = Math.min(tx, ty);
  return [cx + dx * t, cy + dy * t];
}
function dgArrow(a: DgBox, b: DgBox, label?: string): SVGElement[] {
  const [ax, ay] = centre(a), [bx, by] = centre(b);
  const dx = bx - ax, dy = by - ay;
  const len = Math.hypot(dx, dy) || 1;
  const ux = dx / len, uy = dy / len;
  const [x1, y1] = edgePoint(a, dx, dy);
  const [x2, y2] = edgePoint(b, -dx, -dy);
  const out: SVGElement[] = [sv("line", { x1, y1, x2: x2 - ux * 3.6, y2: y2 - uy * 3.6, class: "ar" })];
  const hx = -ux * 4.4, hy = -uy * 4.4, px = -uy * 2.3, py = ux * 2.3;
  out.push(sv("path", { class: "hd", d: `M${x2.toFixed(1)} ${y2.toFixed(1)}L${(x2 + hx + px).toFixed(1)} ${(y2 + hy + py).toFixed(1)}L${(x2 + hx - px).toFixed(1)} ${(y2 + hy - py).toFixed(1)}Z` }));
  if (label) {
    // beside a vertical arrow, above a horizontal one: a label on top of the line hides it
    const down = Math.abs(uy) > Math.abs(ux);
    const t = sv("text", { class: `el${down ? " side" : ""}`, x: (x1 + x2) / 2 + (down ? 5 : 0), y: (y1 + y2) / 2 - (down ? 0 : 2.5) });
    t.textContent = cut(label, down ? 60 : Math.max(18, len - 12));
    out.push(t);
  }
  return out;
}

/** Node boxes per kind: a row or column for flow, a ring, a line, a tree by depth, a 2 x 2 grid. */
function dgLayout(d: Diagram, W: number, H: number): { boxes: Map<string, DgBox>; extra: SVGElement[] } {
  const boxes = new Map<string, DgBox>();
  const extra: SVGElement[] = [];
  const n = d.nodes.length;
  const gap = 12;
  if (d.kind === "cycle") {
    const bw = Math.min(70, Math.max(30, W / 3.1)), bh = Math.min(22, Math.max(14, H / 3.4));
    const rx = Math.max(6, W / 2 - bw / 2 - 2), ry = Math.max(6, H / 2 - bh / 2 - 2);
    d.nodes.forEach((node, i) => {
      const a = -Math.PI / 2 + (i * 2 * Math.PI) / n;
      boxes.set(node.id, { x: W / 2 + Math.cos(a) * rx - bw / 2, y: H / 2 + Math.sin(a) * ry - bh / 2, w: bw, h: bh });
    });
    return { boxes, extra };
  }
  if (d.kind === "timeline") {
    const bh = Math.min(20, Math.max(12, H * 0.42));
    const bw = Math.max(18, (W - gap * (n - 1)) / n);
    const axis = H - bh / 2 - 4;
    extra.push(sv("line", { x1: 0, y1: axis, x2: W, y2: axis, class: "ax" }));
    d.nodes.forEach((node, i) => {
      const x = i * (bw + gap);
      boxes.set(node.id, { x, y: Math.max(0, axis - bh - 10), w: bw, h: bh });
      extra.push(sv("circle", { cx: x + bw / 2, cy: axis, r: 2.4, class: "tick" }));
      extra.push(sv("line", { x1: x + bw / 2, y1: axis - 10, x2: x + bw / 2, y2: axis, class: "ar" }));
    });
    return { boxes, extra };
  }
  if (d.kind === "matrix") {
    extra.push(sv("line", { x1: W / 2, y1: 0, x2: W / 2, y2: H, class: "ax" }));
    extra.push(sv("line", { x1: 0, y1: H / 2, x2: W, y2: H / 2, class: "ax" }));
    const rows = Math.max(2, Math.ceil(n / 2));
    const bw = (W - gap * 3) / 2, bh = Math.min(26, (H - gap * (rows + 1)) / rows);
    d.nodes.forEach((node, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      boxes.set(node.id, { x: gap + col * (bw + gap), y: gap + row * (bh + gap), w: bw, h: bh });
    });
    return { boxes, extra };
  }
  if (d.kind === "tree") {
    const depth = new Map<string, number>();
    const incoming = new Set(d.edges.map((e) => e.to));
    const roots = d.nodes.filter((x) => !incoming.has(x.id));
    const queue = (roots.length ? roots : d.nodes.slice(0, 1)).map((x) => x.id);
    for (const id of queue) depth.set(id, 0);
    for (let i = 0; i < queue.length; i++) {
      const at = queue[i], next = (depth.get(at) ?? 0) + 1;
      for (const e of d.edges.filter((x) => x.from === at)) {
        if (depth.has(e.to)) continue;
        depth.set(e.to, next); queue.push(e.to);
      }
    }
    for (const node of d.nodes) if (!depth.has(node.id)) depth.set(node.id, 0);
    const levels = Math.max(...[...depth.values()]) + 1;
    const bh = Math.min(22, Math.max(12, (H - gap * (levels - 1)) / levels));
    for (let lv = 0; lv < levels; lv++) {
      const row = d.nodes.filter((node) => depth.get(node.id) === lv);
      const bw = Math.max(20, (W - gap * (row.length - 1)) / row.length);
      row.forEach((node, i) => boxes.set(node.id, { x: i * (bw + gap), y: lv * (bh + gap), w: bw, h: bh }));
    }
    return { boxes, extra };
  }
  if (d.direction === "TB") {
    const down = 18;                                   // room for the arrow and its label between rows
    const bh = Math.max(12, Math.min(34, (H - down * (n - 1)) / n));
    const bw = Math.min(W, 140);
    d.nodes.forEach((node, i) => boxes.set(node.id, { x: (W - bw) / 2, y: i * (bh + down), w: bw, h: bh }));
    return { boxes, extra };
  }
  const bw = Math.max(18, (W - gap * (n - 1)) / n), bh = Math.min(30, Math.max(14, H * 0.5));
  d.nodes.forEach((node, i) => boxes.set(node.id, { x: i * (bw + gap), y: (H - bh) / 2, w: bw, h: bh }));
  return { boxes, extra };
}

/** The sketch: nodes as small rounded boxes, arrows by kind, drawn from `block.diagram` when it has any. */
export function diagramSketch(b: Block): SVGElement {
  const d = b.diagram?.nodes?.length ? b.diagram : GENERIC;
  const W = Math.max(60, b.cols * DECK_CELL - 10), H = Math.max(26, (b.rows - 1) * DECK_CELL - 8);
  const root = sv("svg", { class: "dg", viewBox: `0 0 ${Math.round(W)} ${Math.round(H)}`, preserveAspectRatio: "none" });
  const { boxes, extra } = dgLayout(d, W, H);
  root.append(...extra);
  const ring: DiagramEdge[] = d.kind === "cycle" ? d.nodes.map((n, i) => ({ from: n.id, to: d.nodes[(i + 1) % d.nodes.length].id })) : [];
  const drawn = d.kind === "cycle" && !d.edges.length ? ring : d.edges;
  if (d.kind !== "matrix" && d.kind !== "timeline") {
    for (const e of drawn) {
      const a = boxes.get(e.from), z = boxes.get(e.to);
      if (a && z && a !== z) root.append(...dgArrow(a, z, e.label));
    }
  }
  for (const node of d.nodes) { const r = boxes.get(node.id); if (r) root.append(dgNode(node, r)); }
  return root;
}

function deckBody(b: Block): Node[] {
  switch (b.type as DeckBlockType) {
    case "kicker": return [h("div", { class: "caps" }, b.content || "Category | Topic")];
    // A headline block from an older plan is the slide's own headline now (F14): it draws nothing here.
    case "headline": return [];
    case "text": return [lines(Math.max(3, Math.min(5, b.rows - 2)))];
    case "bullets": {
      const n = Math.max(3, Math.min(6, b.rows - 2));
      return [h("div", { class: "lines" }, ...Array.from({ length: n }, (_, i) => h("div", { class: "bl" }, h("span", { class: "bd" }), line(WIDTHS[i % WIDTHS.length]))))];
    }
    case "cards": {
      const n = Math.max(2, Math.min(6, b.count ?? 3));
      // Header, sub, then three or four lines: the sketch keeps the card's proportion instead of filling it to the bottom.
      const body = Math.max(2, Math.min(4, b.rows - 4));
      return [h("div", { class: "cards" }, ...Array.from({ length: n }, () => h("div", { class: "ck" },
        line(68, "hd"), line(44, "sub"),
        h("div", { class: "bdy" }, ...Array.from({ length: body }, (_, i) => line(WIDTHS[i % WIDTHS.length], "thin"))))))];
    }
    case "stat": {
      const n = Math.max(1, Math.min(4, b.count ?? 1));
      // The number is the generator's guess, so it is drawn as a placeholder: muted and dotted, never as data.
      return [h("div", { class: `stats${n > 1 ? " many" : ""}` }, ...Array.from({ length: n }, (_, i) => h("div", { class: "sv" }, h("div", { class: "num ph" }, fakeValue(b.title || "value", i)), line(64))))];
    }
    case "chart": return [chartSketch(b.chartKind ?? "bar")];
    case "diagram": return [diagramSketch(b)];
    case "table": return [tableSketch(b)];
    case "image": return [h("div", { class: "img" }, svg('<line x1="0" y1="0" x2="100" y2="60" stroke="currentColor" opacity="0.3"/>'), h("span", { class: "lbl" }, b.content || b.title || "Image"))];
    case "quote": return [h("div", { class: "qt" }, h("span", { class: "qm" }, "“"), lines(2, "big"))];
    case "takeaway": return [h("div", { class: "bar" }, h("span", { class: "trunc" }, b.content || b.title))];
    case "placeholder": return [h("div", { class: "ph" }, h("span", { class: "trunc" }, b.content || b.title))];
    case "notes": return [h("div", { class: "src" }, line(54, "thin"), line(38, "thin"))];
    default: return [lines(2)];
  }
}

function renderDeckBlock(b: Block, notes: Note[]): HTMLElement {
  const flat = b.rows <= 1;
  return h("div", { class: `blk deck ${b.type}${b.emphasis ? " emph" : ""}${flat ? " flat" : ""}`, "data-id": b.id },
    h("div", { class: "bh" },
      h("span", { class: `ty ${b.type}` }, deckLabel(b.type)),
      h("span", { class: "t", title: b.content ? `${b.title}: ${b.content}` : b.title }, b.title),
      ...blockBadges(notes),
      b.emphasis ? h("span", { class: "badge emph", title: "This is what should pop" }, "pop") : null,
      b.detected ? h("span", { class: `badge det${b.detected.confidence < 0.6 ? " low" : ""}`, title: `Detected at ${Math.round(b.detected.confidence * 100)}%: ${b.detected.signals.join(", ")}` }, `${Math.round(b.detected.confidence * 100)}%`) : null),
    flat ? null : h("div", { class: "bc" }, ...deckBody(b)),
    ...resizeHandles());
}

// ---- the view ----

/** Scroll position per slide so a save does not jump the view; zoom is null while it follows the panel width. */
const scrolls = new Map<string, [number, number]>();
/** The slide strip keeps its own scroll across re-renders (F21), one position per plan. */
const stripScroll = new Map<string, number>();
let deckZoom: number | null = null;
/** The archetype a Remix just applied, per slide, so the label names it until the slide is edited again. */
const lastRemix = new Map<string, { archetype: string; snap: string }>();
/** Geometry only, order independent: survives the round trip through the server, which rebuilds the blocks. */
const geomSig = (blocks: Block[]) => blocks.map((b) => `${b.id}:${b.row},${b.col},${b.rows},${b.cols}`).sort().join("|");

/** Two digits, the way a deck numbers its sections: "03 Competitive Analysis". */
const sectionNo = (n: number) => String(n).padStart(2, "0");

export function renderDeckLayoutView(c: Controller, root: HTMLElement) {
  const p = c.project;
  const tab = c.tab;
  const tokens = deckStyle(c)?.tokens;
  const done = (t: Tab) => t.layoutApproved || !!t.divider;   // a divider has nothing to lay out
  const slides = p.tabs.filter((t) => !t.divider);

  // ---- the strip: every row, dividers included, its scroll kept across draws ----
  const strip = h("div", { class: "tabstrip", "data-scroll": "strip" }, ...p.tabs.map((t) => {
    const open = p.notes.filter((n) => n.tabId === t.id && n.status === "open").length;
    return h("button", { class: `${t.divider ? "dv " : ""}${t.id === c.selectedTab ? "on" : ""}`.trim(), onclick: () => { c.selectedTab = t.id; c.selectedBlock = null; c.selectedBlocks.clear(); c.setState(c.state!); } },
      h("span", { class: `dot ${done(t) ? "ok" : t.blocks.length ? "warn" : "none"}` }), t.name, open ? h("span", { class: "n" }, String(open)) : null);
  }));
  strip.onscroll = () => stripScroll.set(p.id, strip.scrollLeft);
  const main = h("div", { class: "main" }, strip);
  const side = h("div", { class: "side" });
  root.append(h("div", { class: "body" }, main, side));
  strip.scrollLeft = stripScroll.get(p.id) ?? 0;
  // the chip only comes into view when it is out of it, so clicking one never drags the strip back left
  requestAnimationFrame(() => {
    const chip = strip.querySelector<HTMLElement>("button.on");
    if (!chip) return;
    const left = chip.offsetLeft, right = left + chip.offsetWidth;
    if (left < strip.scrollLeft || right > strip.scrollLeft + strip.clientWidth) chip.scrollIntoView({ inline: "nearest", block: "nearest" });
  });
  if (!tab) { main.append(h("div", { class: "center" }, "No slides yet")); return; }

  const pageNo = p.tabs.findIndex((t) => t.id === tab.id) + 1;
  const cycleSlide = (d: number) => {
    const i = p.tabs.findIndex((t) => t.id === tab.id);
    const next = p.tabs[(i + d + p.tabs.length) % p.tabs.length];
    if (!next || next.id === tab.id) return;
    c.selectedTab = next.id; c.selectedBlock = null; c.selectedBlocks.clear(); c.setState(c.state!);
  };

  // ---- a divider row draws the format's section slide: no palette, no blocks, no Remix (F16) ----
  if (tab.divider) {
    renderDividerSlide(c, root, main, side, { tokens, tab, pageNo, cycleSlide });
    renderDeckFooter(c, root, { slides, done });
    return;
  }

  const blocks = tab.blocks;
  // A `headline` block from an older plan is the slide's headline: kept in the plan, drawn nowhere (F14).
  const body = blocks.filter((b) => b.type !== "headline");
  const hist = historyFor(tab.id);
  let before = snapshotBlocks(blocks);
  // Selection invariants: members exist on this slide; the primary is a member; no primary means nothing selected.
  for (const id of [...c.selectedBlocks]) if (!body.some((b) => b.id === id)) c.selectedBlocks.delete(id);
  if (!c.selectedBlock) c.selectedBlocks.clear();
  else if (!c.selectedBlocks.has(c.selectedBlock)) c.selectedBlocks = new Set([c.selectedBlock]);
  const selected = () => body.filter((b) => c.selectedBlocks.has(b.id));

  const persist = (opts: { record?: boolean; resolve?: Block | Block[] } = {}) => {
    clampToBody(body);
    // `resolve` says a gesture just moved something; which block gives way is the server's rule, not the drag's.
    if (opts.resolve && resolveDeckOverlaps(body)) c.toast?.("Moved to fit");
    const after = snapshotBlocks(blocks);
    if (opts.record !== false && after !== before) { hist.past.push(before); if (hist.past.length > HISTORY_MAX) hist.past.shift(); hist.future = []; }
    before = after;
    c.saveLayout(tab, blocks);
  };
  const restore = (json: string) => { const next = JSON.parse(json) as Block[]; blocks.splice(0, blocks.length, ...next); before = json; c.saveLayout(tab, blocks); };
  const undo = () => { const prev = hist.past.pop(); if (prev == null) { c.toast?.("Nothing to undo"); return; } hist.future.push(snapshotBlocks(blocks)); restore(prev); };
  const redo = () => { const next = hist.future.pop(); if (next == null) return; hist.past.push(snapshotBlocks(blocks)); restore(next); };
  const removeBlocks = (gone: Block[]) => {
    if (!gone.length) return;
    for (const b of gone) { const i = blocks.indexOf(b); if (i >= 0) blocks.splice(i, 1); }
    const ids = new Set(gone.map((b) => b.id));
    p.notes = p.notes.filter((n) => !n.blockId || !ids.has(n.blockId));
    c.selectedBlock = null; c.selectedBlocks.clear();
    persist();
  };

  // ---- remix ----
  const stuck = lastRemix.get(tab.id);
  const archetype = detectArchetype(blocks) ?? (stuck && stuck.snap === geomSig(blocks) ? stuck.archetype : undefined);
  const doRemix = (name?: string) => {
    const next = remix(blocks, name);
    blocks.splice(0, blocks.length, ...next.blocks);
    clampToBody(blocks.filter((b) => b.type !== "headline"));
    lastRemix.set(tab.id, { archetype: next.archetype, snap: geomSig(blocks) }); // before persist: saving re-renders
    persist();
    c.toast?.(`Remixed: ${next.archetype}`);
  };

  document.onkeydown = (e) => {
    const inField = !!(e.target as HTMLElement).closest("input, textarea, select");
    const mod = e.metaKey || e.ctrlKey;
    if (mod && e.key.toLowerCase() === "z") { if (inField) return; e.preventDefault(); if (e.shiftKey) redo(); else undo(); return; }
    if (mod && (e.key === "=" || e.key === "+")) { e.preventDefault(); setZoom(ZOOM_STEPS[Math.min(ZOOM_STEPS.length - 1, ZOOM_STEPS.indexOf(zoom()) + 1)]); return; }
    if (mod && (e.key === "-" || e.key === "_")) { e.preventDefault(); setZoom(ZOOM_STEPS[Math.max(0, ZOOM_STEPS.indexOf(zoom()) - 1)]); return; }
    if (mod && e.key === "0") { e.preventDefault(); setZoom(null); return; }
    if (inField) return;
    // ⌘← and ⌘→ walk the strip (F15); the plain arrows stay with the selected block
    if (mod && (e.key === "ArrowLeft" || e.key === "ArrowRight")) { e.preventDefault(); cycleSlide(e.key === "ArrowLeft" ? -1 : 1); return; }
    const b = c.block;
    if (mod && e.key.toLowerCase() === "d" && b) { e.preventDefault(); duplicate(b); return; }
    if (e.key === "Escape") { document.querySelectorAll(".pop").forEach((x) => x.remove()); c.selectedBlock = null; c.selectedBlocks.clear(); c.setState(c.state!); return; }
    if (e.key === "?") { e.preventDefault(); openCheatSheet(root, SHORTCUTS); return; }
    // Remix is ⌥R and nothing else (F13). Option+R reports "®" on a Mac keyboard, so the code decides.
    if (e.altKey && !mod && (e.code === "KeyR" || e.key.toLowerCase() === "r")) { e.preventDefault(); doRemix(); return; }
    if (e.key === "Backspace" || e.key === "Delete") { const sel = selected(); if (!sel.length) return; e.preventDefault(); removeBlocks(sel); return; }
    if (/^[1-9]$/.test(e.key) && !mod) { const t = p.tabs[Number(e.key) - 1]; if (t) { c.selectedTab = t.id; c.selectedBlock = null; c.setState(c.state!); } return; }
    if (!b) return;
    const sel = selected();
    const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key as string];
    if (d) {
      e.preventDefault();
      if (e.shiftKey) for (const s of sel) { s.rows = Math.max(1, s.rows + d[0]); s.cols = Math.max(1, s.cols + d[1]); }
      else {
        // move the group by one cell, clamped as one group at the edges of the body band
        const top = Math.min(...sel.map((s) => s.row)), left = Math.min(...sel.map((s) => s.col));
        const btm = Math.max(...sel.map((s) => s.row + s.rows - 1)), right = Math.max(...sel.map((s) => s.col + s.cols - 1));
        const dRow = Math.max(DECK_BODY_TOP - top, Math.min(DECK_BODY_BOTTOM - btm, d[0])), dCol = Math.max(1 - left, Math.min(DECK_COLS - right, d[1]));
        for (const s of sel) { s.row += dRow; s.col += dCol; }
      }
      persist(); return;
    }
  };

  // ---- toolbar ----
  const duplicate = (b: Block) => {
    const copy: Block = fitToSlide({ ...b, id: uid("blk"), title: `${b.title} copy`, row: b.row + b.rows + 1 });
    blocks.push(copy); body.push(copy);
    c.selectedBlock = copy.id; c.selectedBlocks = new Set([copy.id]); persist({ resolve: copy });
  };
  const addBlock = (type: DeckBlockType) => {
    const [rows, cols] = DECK_DEFAULT_SIZE[type];
    const bottom = body.reduce((m, x) => Math.max(m, x.row + x.rows), DECK_BODY_TOP - 1);
    const b: Block = fitToSlide({ id: uid("blk"), type, title: DECK_LABEL[type], row: type === "takeaway" ? DECK_BODY_BOTTOM - rows + 1 : bottom + 1, col: Math.min(3, DECK_COLS - cols + 1), rows, cols });
    if (type === "cards") b.count = 3;
    if (type === "table") b.columns = ["Item", "Now", "Target"];
    if (type === "diagram") b.diagram = { kind: "flow", direction: "LR", nodes: [], edges: [] };
    blocks.push(b); body.push(b);
    c.selectedBlock = b.id; c.selectedBlocks = new Set([b.id]); persist({ resolve: b });
    requestAnimationFrame(() => (side.querySelector("input.title") as HTMLInputElement | null)?.select());
  };
  const addBtn = h("button", { onclick: (e: Event) => popover(e.currentTarget as HTMLElement, (close) => h("div", { class: "deckpal" }, ...PALETTE.map((t) => h("button", { onclick: () => { close(); addBlock(t); } }, h("span", { class: `ty ${t}` }, DECK_LABEL[t]))))) }, icon("plus", 11), "Block");
  const remixBtn = h("button", { class: "ghost", title: "Re-lay as the next archetype", onclick: () => doRemix() }, icon("flow", 11), "Remix", h("span", { class: "key" }, "⌥R"));

  // ---- Fit, Balance and the column module: the balanced passes, on this slide only ----
  const mod = columnModule(tokens);
  const movable = () => body.filter((b) => !FURNITURE_TYPES.has(b.type));
  /** The selection, or the whole slide when nothing is selected. Furniture is never in it (rule 8). */
  const targets = () => { const sel = selected().filter((b) => !FURNITURE_TYPES.has(b.type)); return sel.length ? sel : movable(); };
  /** "3 blocks" for a toast: what actually moved, not what was asked. */
  const moved = (sel: Block[], was: number[], key: "rows" | "row") => sel.filter((b, i) => b[key] !== was[i]).length;
  const doFit = () => {
    const sel = targets();
    if (!sel.length) { c.toast?.("Nothing to fit"); return; }
    const was = sel.map((b) => b.rows);
    if (!fitBlocks(sel)) { c.toast?.("Already fits"); return; }
    persist({ resolve: sel });
    const n = moved(sel, was, "rows");
    c.toast?.(`Fitted ${n} block${n === 1 ? "" : "s"}`);
  };
  const doBalance = () => {
    const sel = targets();
    if (!sel.length) { c.toast?.("Nothing to balance"); return; }
    const was = sel.map((b) => b.row);
    if (!balanceBlocks(sel)) { c.toast?.("Already balanced"); return; }
    persist({ resolve: sel });
    const n = moved(sel, was, "row");
    c.toast?.(`Balanced ${n} block${n === 1 ? "" : "s"}`);
  };
  const fitBtn = h("button", { class: "ghost", title: "Size to the content", onclick: doFit }, icon("shrink", 11), "Fit");
  const balanceBtn = h("button", { class: "ghost", title: "Equal gaps down the band", onclick: doBalance }, icon("stack", 11), "Balance");
  const snapBtn: HTMLButtonElement = h("button", { class: `ghost snapbtn${snapRemembered(p.id) ? " on" : ""}`, title: "Snap to the column module", onclick: () => {
    const on = !snapRemembered(p.id);
    rememberSnap(p.id, on);
    snapBtn.classList.toggle("on", on);
    snapBtn.setAttribute("aria-pressed", String(on));
    c.toast?.(on ? "Snap on" : "Snap off");
  }, "aria-pressed": String(snapRemembered(p.id)) }, "Snap");
  /** A dropped block slides onto the module and keeps its width; a group shifts as one, so a band holds. */
  const snapDrop = (group: Block[]) => {
    if (!snapRemembered(p.id) || !group.length) return;
    if (group.length === 1) { Object.assign(group[0], snapToModule(group[0], mod, "move")); return; }
    const left = Math.min(...group.map((g) => g.col)), right = Math.max(...group.map((g) => g.col + g.cols - 1));
    const d = snapToModule({ col: left, cols: right - left + 1 }, mod, "move").col - left;
    const shift = Math.max(1 - left, Math.min(DECK_COLS - right, d));
    if (shift) for (const g of group) g.col += shift;
  };
  /** A resized block takes the module on the edge the grip moved; the opposite edge stays put. */
  const snapResize = (b: Block, corner?: string) => {
    if (!snapRemembered(p.id)) return;
    Object.assign(b, snapToModule(b, mod, corner === "nw" || corner === "sw" ? "left" : "right"));
  };
  const archBtn = h("button", { class: "chip line archchip", title: "Pick an archetype", onclick: (e: Event) => popover(e.currentTarget as HTMLElement, (close) => h("div", { class: "archlist" }, ...ARCHETYPES.map((a) => h("button", { class: a.name === archetype ? "on" : "", onclick: () => { close(); doRemix(a.name); } }, h("span", { class: "t" }, a.name), h("span", { class: "n" }, String(a.slots.length)))))) }, archetype ?? "custom", icon("chev", 9));
  const warnForTab = c.state!.warnings.filter((w) => w.tabId === tab.id);
  const zoomBtns = ZOOM_STEPS.map((z) => h("button", { title: `${Math.round(z * 100)}%`, onclick: () => setZoom(z) }, `${Math.round(z * 100)}%`));
  main.append(h("div", { class: "toolbar" }, addBtn, h("span", { class: "sep" }), remixBtn, archBtn, fitBtn, balanceBtn, snapBtn, h("span", { class: "sep" }),
    h("button", { class: "ghost icon", title: "Undo (⌘Z)", disabled: hist.past.length === 0, onclick: undo }, icon("undo", 12)),
    h("button", { class: "ghost icon", title: "Redo (⇧⌘Z)", disabled: hist.future.length === 0, onclick: redo }, icon("redo", 12)),
    h("span", { class: "purpose", title: tab.story ?? tab.purpose }, tab.section ? `${tab.section} · ${tab.purpose}` : tab.purpose),
    h("div", { class: "seg zoom", title: "Zoom · Fit ⌘0" }, ...zoomBtns),
    h("button", { class: "ghost icon", title: "Shortcuts (?)", onclick: () => openCheatSheet(root, SHORTCUTS) }, "?"),
    warnForTab.length ? h("span", { class: "chip warn", title: warnForTab.map((w) => w.message).join("\n") }, h("span", { class: "dot warn" }), `${warnForTab.length} flagged`) : h("span", { class: "chip ok" }, h("span", { class: "dot ok" }), tab.layoutApproved ? "approved" : "ready")));
  const selbar = h("div", { class: "toolbar selbar", hidden: true });
  main.append(selbar);
  const run = (fn: (sel: Block[]) => void) => { const sel = selected(); if (sel.length < 2) return; fn(sel); persist({ resolve: sel }); };
  const refreshSelbar = () => fillSelectionBar(selbar, { count: c.selectedBlocks.size, primary: c.block, run, deselect: () => setSel([], null) });

  // ---- slide ----
  const W = SLIDE_W + PAD * 2, H = SLIDE_H + PAD * 2;
  const wrap = h("div", { class: "canvaswrap deckwrap", tabindex: "0", "data-scroll": `canvas:${tab.id}` });
  const scaler = h("div", { class: "canvasscale" });
  const canvas = h("div", { class: "canvas deckcanvas", style: { width: `${W}px`, height: `${H}px` } });
  const slide = h("div", { class: "slide", style: { left: `${PAD}px`, top: `${PAD}px`, width: `${SLIDE_W}px`, height: `${SLIDE_H}px` } });
  paintSlide(slide, tokens, "content");
  const dots = h("div", { class: "dots", style: { backgroundSize: `${DECK_CELL}px ${DECK_CELL}px` } });
  slide.append(dots);
  // the module's own edges, faint and only while a block is moving, so a drop reads as landing on them
  const guides = h("div", { class: "guides" },
    ...mod.starts.map((col) => h("div", { class: "gd", style: { left: `${(col - 1) * DECK_CELL}px` } })),
    h("div", { class: "gd", style: { left: `${mod.ends[mod.ends.length - 1] * DECK_CELL}px` } }));
  slide.append(guides);
  const showGuides = (on: boolean) => guides.classList.toggle("on", on && snapRemembered(p.id));
  canvas.append(slide);
  scaler.append(canvas); wrap.append(scaler); main.append(wrap);

  // furniture, under the blocks: the headline band is the only part that takes a click
  const hasData = body.some((b) => b.type === "chart" || b.type === "table" || b.type === "stat");
  const kicker = tab.kicker?.trim() || kickerText(tokens, tab.section, tab.name);
  // rows 1 to 4 are the furniture's own band: subtly marked, and nothing drops into it
  slide.append(h("div", { class: "band", style: { height: `${(DECK_BODY_TOP - 1) * DECK_CELL}px` } }));
  if (tokens?.rules?.kicker !== false) slide.append(h("div", { class: `dk kicker${body.some((b) => b.type === "kicker") ? " dim" : ""}`, style: box(furnitureBox(tokens, "kicker")) }, kicker));
  // the rule under the headline is optional furniture: drawn where the style asks for it, and nowhere else
  const hairline = tokens ? tokens.furniture?.hairline : HOUSE.hairline;
  if (hairline) slide.append(h("div", { class: "dk hairline", style: box({ ...hairline, h: undefined }) }));
  if (hasData && tokens?.rules?.source !== false) slide.append(h("div", { class: "dk source", style: box(furnitureBox(tokens, "source")) }, "Source"));
  if (tokens?.rules?.pageNumber !== false) slide.append(h("div", { class: "dk page num", style: box(furnitureBox(tokens, "pageNumber")) }, String(pageNo)));

  const headBox = furnitureBox(tokens, "headline");
  const headline = h("div", { class: "dk headline", style: box(headBox), title: "Click to edit" });
  const drawHeadline = () => { clear(headline); const t = tab.message?.trim(); headline.classList.toggle("empty", !t); headline.append(t || "Add the headline"); };
  const commitHeadline = (text: string) => {
    const next = text.replace(/\s+/g, " ").trim();
    if (next === (tab.message ?? "")) { drawHeadline(); return; }
    saveHeadline(c, tab, next);
    drawHeadline();
  };
  headline.onpointerdown = (e) => e.stopPropagation();
  headline.onclick = () => {
    const ta = h("textarea", { class: "dk hledit", style: box(headBox), value: tab.message ?? "", spellcheck: "false", placeholder: "The slide's message, as a sentence" });
    headline.replaceWith(ta);
    ta.focus(); ta.select();
    let done = false;
    const close = (save: boolean) => { if (done) return; done = true; const v = ta.value; ta.replaceWith(headline); if (save) commitHeadline(v); else drawHeadline(); };
    ta.onkeydown = (e) => { e.stopPropagation(); if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); close(true); } if (e.key === "Escape") { e.preventDefault(); close(false); } };
    ta.onblur = () => close(true);
  };
  drawHeadline();
  slide.append(headline);

  // ---- zoom: the largest step whose slide fits the panel, until the user picks one ----
  const fitZoom = () => {
    const room = wrap.clientWidth - 6;
    const fits = ZOOM_STEPS.filter((z) => W * z <= room);
    return fits.length ? fits[fits.length - 1] : ZOOM_STEPS[0];
  };
  const zoom = () => deckZoom ?? fitZoom();
  const applyZoom = () => {
    const z = zoom();
    scaler.style.width = `${W * z}px`; scaler.style.height = `${H * z}px`;
    canvas.style.transform = `scale(${z})`;
    zoomBtns.forEach((btn, i) => btn.classList.toggle("on", ZOOM_STEPS[i] === z));
    mini.draw();
  };
  const setZoom = (z: number | null) => {
    const from = zoom();
    const cx = (wrap.scrollLeft + wrap.clientWidth / 2) / from, cy = (wrap.scrollTop + wrap.clientHeight / 2) / from;
    deckZoom = z;
    const to = zoom();
    applyZoom();
    if (to !== from) { wrap.scrollLeft = Math.max(0, cx * to - wrap.clientWidth / 2); wrap.scrollTop = Math.max(0, cy * to - wrap.clientHeight / 2); }
    mini.viewport();
  };
  wrap.onscroll = () => { mini.viewport(); scrolls.set(tab.id, [wrap.scrollLeft, wrap.scrollTop]); };

  // ---- blocks ----
  const blockEls = new Map<string, HTMLElement>();
  const place = (el: HTMLElement, b: Block) => {
    el.style.left = `${(b.col - 1) * DECK_CELL}px`; el.style.top = `${(b.row - 1) * DECK_CELL}px`;
    el.style.width = `${b.cols * DECK_CELL}px`; el.style.height = `${b.rows * DECK_CELL}px`;
  };
  const markOverlaps = () => { for (const b of body) blockEls.get(b.id)?.classList.toggle("overlap", body.some((o) => o !== b && overlaps(o, b))); };
  const refreshSelection = (full = true) => {
    for (const [id, el] of blockEls) { el.classList.toggle("sel", c.selectedBlocks.has(id)); el.classList.toggle("primary", id === c.selectedBlock); }
    mini.draw();
    if (full) { refreshSide(); refreshSelbar(); }
  };
  const { setSel, toggleSel, pressSel } = selection(c, () => refreshSelection());

  for (const b of body) {
    const el = renderDeckBlock(b, p.notes.filter((n) => n.blockId === b.id));
    place(el, b);
    el.classList.toggle("sel", c.selectedBlocks.has(b.id));
    el.classList.toggle("primary", c.selectedBlock === b.id);
    el.classList.toggle("overlap", body.some((o) => o !== b && overlaps(o, b)));
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
      // the whole selection moves together, clamped as one group inside the body band
      const group = selected();
      const starts = new Map(group.map((g) => [g.id, { row: g.row, col: g.col }]));
      const top = Math.min(...group.map((g) => g.row)), left = Math.min(...group.map((g) => g.col));
      const btm = Math.max(...group.map((g) => g.row + g.rows - 1)), right = Math.max(...group.map((g) => g.col + g.cols - 1));
      // Deltas between two moves, never from the press: a captured pointer inside the host's nested
      // iframes can report another frame's coordinates, and then the press and the moves disagree.
      let from: { x: number; y: number } | null = null;
      el.classList.add("dragging");
      showGuides(true);
      const onMove = (ev: PointerEvent) => {
        from ??= { x: ev.clientX, y: ev.clientY };
        const z = zoom();
        const dCol = Math.max(1 - left, Math.min(DECK_COLS - right, Math.round((ev.clientX - from.x) / z / DECK_CELL)));
        const dRow = Math.max(DECK_BODY_TOP - top, Math.min(DECK_BODY_BOTTOM - btm, Math.round((ev.clientY - from.y) / z / DECK_CELL)));
        for (const g of group) { const s = starts.get(g.id)!; g.col = s.col + dCol; g.row = s.row + dRow; place(blockEls.get(g.id)!, g); }
        markOverlaps(); mini.draw();
      };
      const onUp = () => {
        header.removeEventListener("pointermove", onMove); header.removeEventListener("pointerup", onUp); el.classList.remove("dragging"); c.interacting = false;
        showGuides(false);
        if (group.some((g) => { const s = starts.get(g.id)!; return g.row !== s.row || g.col !== s.col; })) { snapDrop(group); persist({ resolve: group }); }
        refreshSide();
      };
      header.addEventListener("pointermove", onMove); header.addEventListener("pointerup", onUp);
    };
    header.ondblclick = () => { c.selectedBlock = b.id; refreshSelection(); const input = side.querySelector<HTMLInputElement>("input.title"); input?.focus(); input?.select(); };
    // all four corners (F19); the band and the footer row are walls, as they are for a drag
    wireCornerResize({
      el, block: b, cell: { w: DECK_CELL, h: DECK_CELL },
      bounds: { top: DECK_BODY_TOP, bottom: DECK_BODY_BOTTOM, left: 1, right: DECK_COLS },
      zoom, place: () => place(el, b),
      moved: () => { markOverlaps(); mini.draw(); },
      done: (corner) => { snapResize(b, corner); place(el, b); persist({ resolve: b }); refreshSide(); },
      interacting: (on) => { c.interacting = on; showGuides(on); },
    });
    slide.append(el);
  }

  // marquee: press on the slide and drag; every block the rectangle touches is selected
  wrap.onpointerdown = (e) => {
    if (e.button !== 0 || (e.target as HTMLElement).closest(".blk, .dk")) return;
    const wr = wrap.getBoundingClientRect();
    if (e.clientX - wr.left >= wrap.clientWidth || e.clientY - wr.top >= wrap.clientHeight) return; // scrollbar
    const keep = e.shiftKey ? [...c.selectedBlocks] : [];
    if (!e.shiftKey) setSel([], null);
    wrap.focus({ preventScroll: true });
    const toSlide = (ev: PointerEvent): [number, number] => { const sr = slide.getBoundingClientRect(); const z = zoom(); return [(ev.clientX - sr.left) / z, (ev.clientY - sr.top) / z]; };
    const [x0, y0] = toSlide(e);   // the press is in this frame's own coordinates; the moves may not be
    let from: { x: number; y: number } | null = null;
    let marq: HTMLElement | null = null;
    try { wrap.setPointerCapture(e.pointerId); } catch { /* not capturable */ }
    c.interacting = true;
    const onMove = (ev: PointerEvent) => {
      from ??= { x: ev.clientX, y: ev.clientY };
      const z = zoom();
      const dx = (ev.clientX - from.x) / z, dy = (ev.clientY - from.y) / z;
      if (!marq) { if (Math.abs(dx) * z < 3 && Math.abs(dy) * z < 3) return; marq = h("div", { class: "marquee" }); slide.append(marq); }
      const x1 = x0 + dx, y1 = y0 + dy;
      const l = Math.min(x0, x1), t = Math.min(y0, y1), r = Math.max(x0, x1), btm = Math.max(y0, y1);
      Object.assign(marq.style, { left: `${l}px`, top: `${t}px`, width: `${r - l}px`, height: `${btm - t}px` });
      const hits = body.filter((b) => { const bl = (b.col - 1) * DECK_CELL, bt = (b.row - 1) * DECK_CELL; return bl < r && l < bl + b.cols * DECK_CELL && bt < btm && t < bt + b.rows * DECK_CELL; }).map((b) => b.id);
      const ids = [...new Set([...keep, ...hits])];
      c.selectedBlocks = new Set(ids); c.selectedBlock = hits[hits.length - 1] ?? ids[ids.length - 1] ?? null;
      refreshSelection(false);
    };
    const onUp = () => { wrap.removeEventListener("pointermove", onMove); wrap.removeEventListener("pointerup", onUp); wrap.removeEventListener("pointercancel", onUp); marq?.remove(); c.interacting = false; refreshSelection(true); };
    wrap.addEventListener("pointermove", onMove); wrap.addEventListener("pointerup", onUp); wrap.addEventListener("pointercancel", onUp);
  };

  const mini = createMinimap({
    W, H, wrap, zoom, blocks: () => body,
    rect: (b) => [PAD + (b.col - 1) * DECK_CELL, PAD + (b.row - 1) * DECK_CELL, b.cols * DECK_CELL, b.rows * DECK_CELL],
    selected: (b) => c.selectedBlocks.has(b.id),
    interacting: (on) => { c.interacting = on; },
  });
  main.append(mini.el);
  new ResizeObserver(() => { if (deckZoom == null) applyZoom(); mini.viewport(); }).observe(wrap);

  applyZoom();
  const saved = scrolls.get(tab.id);
  if (saved) { wrap.scrollLeft = saved[0]; wrap.scrollTop = saved[1]; }
  mini.draw();

  const refreshSide = () => { clear(side); renderDeckSide(c, tab, body, side, persist, duplicate, removeBlocks); };
  refreshSide(); refreshSelbar(); slideIn(side);
  c.changed.clear();

  renderDeckFooter(c, root, { slides, done });
}

/** Paint the slide from the style's tokens: its background and the colours and sizes its furniture uses. */
function paintSlide(slide: HTMLElement, tokens: DeckTokens | undefined, kind: "content" | "divider"): void {
  const vars = slideVars(tokens, kind);
  for (const [k, v] of Object.entries(vars)) slide.style.setProperty(k, v);
  if (vars["--sl-bg"]) slide.dataset.bg = kind;
}

/**
 * A divider slide: the format's own section slide. Background, title box and type come from the tokens,
 * the number counts the dividers in order, and the title is edited in place (F16).
 */
function renderDividerSlide(c: Controller, root: HTMLElement, main: HTMLElement, side: HTMLElement, o: {
  tokens: DeckTokens | undefined; tab: Tab; pageNo: number; cycleSlide: (d: number) => void;
}): void {
  const { tokens, tab } = o;
  const p = c.project;
  const no = p.tabs.filter((t) => t.divider).findIndex((t) => t.id === tab.id) + 1;
  const numbered = tokens?.divider?.numbered !== false;

  document.onkeydown = (e) => {
    const inField = !!(e.target as HTMLElement).closest("input, textarea, select");
    const mod = e.metaKey || e.ctrlKey;
    if (inField) return;
    if (mod && (e.key === "ArrowLeft" || e.key === "ArrowRight")) { e.preventDefault(); o.cycleSlide(e.key === "ArrowLeft" ? -1 : 1); return; }
    if (e.key === "Escape") { document.querySelectorAll(".pop").forEach((x) => x.remove()); return; }
    if (e.key === "?") { e.preventDefault(); openCheatSheet(root, SHORTCUTS); return; }
    if (/^[1-9]$/.test(e.key) && !mod) { const t = p.tabs[Number(e.key) - 1]; if (t) { c.selectedTab = t.id; c.selectedBlock = null; c.setState(c.state!); } }
  };

  main.append(h("div", { class: "toolbar" },
    h("span", { class: "purpose" }, "Section divider"),
    h("button", { class: "ghost icon", title: "Shortcuts (?)", onclick: () => openCheatSheet(root, SHORTCUTS) }, "?"),
    h("span", { class: "chip ok" }, h("span", { class: "dot ok" }), "divider")));

  const W = SLIDE_W + PAD * 2, H = SLIDE_H + PAD * 2;
  const wrap = h("div", { class: "canvaswrap deckwrap", tabindex: "0", "data-scroll": `canvas:${tab.id}` });
  const scaler = h("div", { class: "canvasscale" });
  const canvas = h("div", { class: "canvas deckcanvas", style: { width: `${W}px`, height: `${H}px` } });
  const slide = h("div", { class: "slide divider", style: { left: `${PAD}px`, top: `${PAD}px`, width: `${SLIDE_W}px`, height: `${SLIDE_H}px` } });
  paintSlide(slide, tokens, "divider");
  canvas.append(slide); scaler.append(canvas); wrap.append(scaler); main.append(wrap);

  const titleBox = furnitureBox(tokens, "dividerTitle");
  const title = h("div", { class: "dk dtitle", style: box(titleBox), title: "Click to edit" });
  const draw = () => {
    clear(title);
    if (numbered) title.append(h("span", { class: "no" }, sectionNo(no)));
    title.append(h("span", { class: "t" }, tab.name));
  };
  const commit = (text: string) => {
    const next = text.replace(/\s+/g, " ").trim();
    if (next && next !== tab.name) c.saveTabFields(tab.id, { name: next });
    draw();
  };
  title.onclick = () => {
    const input = h("input", { class: "dk dtedit", style: box(titleBox), value: tab.name, spellcheck: "false" });
    title.replaceWith(input);
    input.focus(); input.select();
    let closed = false;
    const close = (save: boolean) => { if (closed) return; closed = true; const v = input.value; input.replaceWith(title); if (save) commit(v); else draw(); };
    input.onkeydown = (e) => { e.stopPropagation(); if (e.key === "Enter") { e.preventDefault(); close(true); } if (e.key === "Escape") { e.preventDefault(); close(false); } };
    input.onblur = () => close(true);
  };
  draw();
  slide.append(title);
  if (tokens?.rules?.pageNumber !== false) slide.append(h("div", { class: "dk page num", style: box(furnitureBox(tokens, "pageNumber")) }, String(o.pageNo)));

  // zoom follows the panel; a divider has nothing to place, so there is no zoom control
  const fit = () => { const room = wrap.clientWidth - 6; const fits = ZOOM_STEPS.filter((z) => W * z <= room); return fits.length ? fits[fits.length - 1] : ZOOM_STEPS[0]; };
  const applyZoom = () => { const z = fit(); scaler.style.width = `${W * z}px`; scaler.style.height = `${H * z}px`; canvas.style.transform = `scale(${z})`; };
  new ResizeObserver(applyZoom).observe(wrap);
  applyZoom();

  // the inspector: the title and the thread, nothing else
  const body = h("div", { class: "scroll", "data-scroll": "side" },
    h("div", { class: "field" }, h("span", { class: "label" }, "Title"),
      h("input", { class: "title", value: tab.name, spellcheck: "false", onchange: (e: Event) => commitField(c, tab, (e.target as HTMLInputElement).value), onkeydown: (e: KeyboardEvent) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); } })),
    h("div", { class: "field" }, h("span", { class: "label" }, "Thread"), notesPanel(c, tab.id, undefined, c.notesFor(tab.id))));
  side.append(h("div", { class: "sh" }, h("span", { class: "label" }, "Divider"), h("span", { class: "t" }, tab.name)), body);
  fadeIn(body); slideIn(side);
}

function commitField(c: Controller, tab: Tab, value: string): void {
  const next = value.replace(/\s+/g, " ").trim();
  if (next && next !== tab.name) c.saveTabFields(tab.id, { name: next });
}

/** The footer: how many slide layouts are approved, and the two approve buttons. Dividers never count. */
function renderDeckFooter(c: Controller, root: HTMLElement, o: { slides: Tab[]; done: (t: Tab) => boolean }): void {
  const p = c.project;
  const tab = c.tab!;
  const allApproved = p.tabs.every(o.done);
  const approved = o.slides.filter((t) => t.layoutApproved).length;
  const openNotes = p.notes.filter((n) => n.status === "open").length;
  // Layout warnings only: an open instruction has its own count, and the missing takeaway belongs to the brief card.
  const flagged = c.state!.warnings.filter((w) => w.kind !== "open-note" && w.kind !== "no-takeaway" && !!w.tabId);
  const approveAll = h("button", { class: "primary", onclick: (e: Event) => {
    if (!flagged.length) { void c.approve("all_layouts"); return; }
    popover(e.currentTarget as HTMLElement, (close) => h("div", { class: "confirm" },
      h("div", { class: "t" }, `${flagged.length} still flagged`),
      h("ul", { class: "warnlist" }, ...flagged.slice(0, 5).map((w) => h("li", {}, h("span", { class: `dot ${w.kind === "overlap" ? "bad" : "warn"}` }), h("span", {}, `${w.tabName}: ${w.message}`)))),
      h("div", { class: "acts" }, h("button", { onclick: close }, "Cancel"), h("button", { class: "primary", onclick: () => { close(); void c.approve("all_layouts"); } }, "Approve anyway"))), { align: "right" });
  } }, "Approve all layouts");
  root.append(h("div", { class: "ftr" },
    h("span", { class: "msg" }, allApproved ? "Layouts approved" : `${approved} of ${o.slides.length} approved${openNotes ? ` · ${openNotes} open note${openNotes === 1 ? "" : "s"}` : ""}${flagged.length ? ` · ${flagged.length} flagged` : ""}`),
    h("span", { class: "spacer" }),
    tab.divider ? null : tab.layoutApproved
      ? h("button", { onclick: () => void c.approve("layout", tab.id, false) }, `Unapprove "${tab.name}"`)
      : h("button", { onclick: () => void c.approve("layout", tab.id) }, `Approve "${tab.name}"`),
    allApproved ? null : approveAll));
}

/** The slide's story, read only: four lines, the rest behind "more" (F20). */
function storyField(tab: Tab): HTMLElement | null {
  const story = tab.story?.trim();
  if (!story) return null;
  const text = h("div", { class: "story clamp" }, story);
  const more = h("button", { class: "more", onclick: () => { const open = !text.classList.toggle("clamp"); more.textContent = open ? "less" : "more"; } }, "more");
  requestAnimationFrame(() => { if (text.scrollHeight <= text.clientHeight + 1) more.hidden = true; });
  return h("div", { class: "field" }, h("span", { class: "label" }, "Story"), h("div", { class: "storywrap" }, text, more));
}

/** Slide-level notes, marked as the slide's, above a block's own thread. Clicking one opens that thread. */
function slideNotes(c: Controller, tab: Tab): HTMLElement | null {
  const notes = c.notesFor(tab.id);
  if (!notes.length) return null;
  const open = () => { c.selectedBlock = null; c.selectedBlocks.clear(); c.setState(c.state!); };
  return h("div", { class: "notes onslide" }, ...notes.map((n) => h("div", { class: `note ${n.status}${n.kind === "question" ? " question" : ""}`, onclick: open },
    h("div", { class: "nh" },
      h("span", { class: `dot ${n.status === "open" ? "warn" : n.status === "resolved" ? "ok" : "bad"}` }),
      h("span", { class: "st" }, "Slide"),
      h("span", { class: "nt trunc", title: n.text }, n.text)))));
}

/** The inspector: the same fields and classes as the workbook one, with the deck's own per-block values. */
function renderDeckSide(
  c: Controller, tab: Tab, blocks: Block[], side: HTMLElement,
  persist: (opts?: { record?: boolean; resolve?: Block }) => void,
  duplicate: (b: Block) => void,
  removeBlocks: (gone: Block[]) => void,
) {
  const b = c.block;
  const p = c.project;
  const at = (x: Block) => `${x.row}:${x.col}`;
  if (!b) {
    const warns = c.state!.warnings.filter((w) => w.tabId === tab.id);
    const body = h("div", { class: "scroll compact", "data-scroll": "side" },
      h("div", { class: "field" }, h("span", { class: "label" }, "Headline"), h("textarea", { rows: 2, value: tab.message ?? "", placeholder: "The slide's message", onchange: (e: Event) => {
        saveHeadline(c, tab, (e.target as HTMLTextAreaElement).value.replace(/\s+/g, " ").trim());
      } })),
      storyField(tab),
      h("div", { class: "field" }, h("span", { class: "label" }, "Readiness"), warns.length ? readinessList(warns, (id) => { c.selectedBlock = id; c.setState(c.state!); }) : h("div", { class: "ok-text" }, "Nothing flagged")),
      h("div", { class: "field" }, h("span", { class: "label" }, "Blocks"), h("div", { class: "blocklist" }, ...blocks.map((x) => h("div", { class: `b${x.id === c.selectedBlock ? " sel" : ""}`, onclick: () => { c.selectedBlock = x.id; c.setState(c.state!); } },
        h("span", { class: "addr" }, at(x)), h("span", { class: "t" }, x.title),
        x.emphasis ? h("span", { class: "badge emph" }, "pop") : null,
        (() => { const oq = p.notes.filter((n) => n.blockId === x.id && n.status === "open" && n.kind === "question").length; return oq ? h("span", { class: "badge q" }, `${oq}?`) : null; })(),
        h("span", { class: "acts" }, twoStep(icon("trash", 11), "Delete?", () => removeBlocks([x]), "")))),
        blocks.length ? null : h("span", { class: "muted small-text" }, "No blocks yet"))),
      h("div", { class: "field" }, h("span", { class: "label" }, "Thread"), notesPanel(c, tab.id, undefined, c.notesFor(tab.id))));
    side.append(h("div", { class: "sh" }, h("span", { class: "label" }, "Slide"), h("span", { class: "t" }, tab.name)), body);
    fadeIn(body);
    return;
  }
  const notes = c.notesFor(tab.id, b.id);
  const scroll = h("div", { class: "scroll compact", "data-scroll": "side" },
    h("div", { class: "field" }, h("span", { class: "label" }, "Title"), h("input", { class: "title", value: b.title, spellcheck: "false", onchange: (e: Event) => { b.title = (e.target as HTMLInputElement).value.trim() || b.title; persist(); }, onkeydown: (e: KeyboardEvent) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); } })),
    h("div", { class: "field" }, h("span", { class: "label" }, "Type"), h("select", { onchange: (e: Event) => { b.type = (e.target as HTMLSelectElement).value as DeckBlockType; persist(); } }, ...PALETTE.map((t) => h("option", { value: t, selected: t === b.type }, DECK_LABEL[t])))),
    h("div", { class: "field" }, h("span", { class: "label" }, "Content"), h("textarea", { rows: 2, value: b.content ?? "", placeholder: "What goes here", onchange: (e: Event) => { b.content = (e.target as HTMLTextAreaElement).value.trim() || undefined; persist(); } })),
    h("div", { class: "field row2" }, h("span", { class: "label" }, "Emphasis"), h("span", { class: "spacer" }), h("div", { class: "seg" },
      h("button", { class: b.emphasis ? "" : "on", onclick: () => { b.emphasis = undefined; persist(); } }, "Off"),
      h("button", { class: b.emphasis ? "on" : "", onclick: () => { b.emphasis = true; persist(); } }, "Pop"))),
    COUNTED.has(b.type) ? h("div", { class: "field row2" }, h("span", { class: "label" }, "Count"), h("span", { class: "spacer" }), h("input", { class: "w-num", type: "number", min: "1", max: "6", value: String(b.count ?? (b.type === "cards" ? 3 : 1)), onchange: (e: Event) => { const n = Number((e.target as HTMLInputElement).value); b.count = Math.max(1, Math.min(6, n || 1)); persist(); } })) : null,
    b.type === "chart" ? h("div", { class: "field" }, h("span", { class: "label" }, "Kind"), h("select", { onchange: (e: Event) => { b.chartKind = (e.target as HTMLSelectElement).value as Block["chartKind"]; persist(); } }, ...(["bar", "line", "pie"] as const).map((k) => h("option", { value: k, selected: (b.chartKind ?? "bar") === k }, k)))) : null,
    ...(b.type === "diagram" ? diagramFields(b, persist) : []),
    b.detected ? h("div", { class: "field" }, h("span", { class: "label" }, "Detected"), h("div", { class: "card" }, h("div", { class: "inline" }, h("span", { class: `chip ${b.detected.confidence < 0.6 ? "warn" : "ok"}` }, h("span", { class: `dot ${b.detected.confidence < 0.6 ? "warn" : "ok"}` }), `${Math.round(b.detected.confidence * 100)}% ${deckLabel(b.type).toLowerCase()}`), h("span", { class: "spacer" }), h("button", { class: "small", title: "Clear the badge", onclick: () => { b.detected = undefined; persist(); } }, "Looks right")), h("div", { class: "ct muted" }, b.detected.signals.join(" · ")))) : null,
    h("div", { class: "field" }, h("span", { class: "label" }, "Thread"), slideNotes(c, tab), notesPanel(c, tab.id, b.id, notes)),
    h("div", { class: "inline tight" },
      h("button", { onclick: () => duplicate(b) }, icon("copy", 12), "Duplicate"),
      twoStep(h("span", {}, "Remove"), "Remove block?", () => removeBlocks([b]), "danger")));
  side.append(h("div", { class: "sh" }, h("span", { class: `tybox ty ${b.type}` }, deckLabel(b.type)), h("span", { class: "t" }, b.title), h("span", { class: "mono muted xs" }, `${at(b)} · ${b.rows}×${b.cols}`)), scroll);
  fadeIn(scroll);
}

/** Kind, direction and the one Edges field a diagram block is described with (F18). */
function diagramFields(b: Block, persist: (opts?: { record?: boolean; resolve?: Block }) => void): (HTMLElement | null)[] {
  const d: Diagram = b.diagram ?? { kind: "flow", direction: "LR", nodes: [], edges: [] };
  const save = (next: Partial<Diagram>) => { b.diagram = { ...d, ...next }; persist(); };
  const directional = d.kind === "flow" || d.kind === "tree";
  return [
    h("div", { class: "field" }, h("span", { class: "label" }, "Kind"), h("select", { onchange: (e: Event) => save({ kind: (e.target as HTMLSelectElement).value as DiagramKind }) },
      ...DIAGRAM_KINDS.map((k) => h("option", { value: k, selected: d.kind === k }, k)))),
    directional ? h("div", { class: "field row2" }, h("span", { class: "label" }, "Direction"), h("span", { class: "spacer" }), h("select", { class: "w-dir", onchange: (e: Event) => save({ direction: (e.target as HTMLSelectElement).value as "LR" | "TB" }) },
      ...(["LR", "TB"] as const).map((v) => h("option", { value: v, selected: (d.direction ?? "LR") === v }, v)))) : null,
    h("div", { class: "field" }, h("span", { class: "label" }, "Edges"), h("textarea", { class: "edges", rows: 4, spellcheck: "false", value: diagramText(d), placeholder: "A -> B: label", onchange: (e: Event) => save(parseDiagram((e.target as HTMLTextAreaElement).value)) })),
  ];
}

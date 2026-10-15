/**
 * The deck archetype catalogue and the pure layout solver, shared by the server and the canvas.
 *
 * Copied from src/app/remix.ts (which now imports from here) so the server can lay out a slide
 * Claude proposed without geometry, exactly as Remix would on the canvas. Dependency-free on
 * purpose: types only, no DOM, no node. Catalogue geometry is docs/research/deck-planning.md
 * section 5, cell for cell.
 *
 * The rules it implements, in that doc's numbering:
 *  1. Role first. Blocks map to slots by role, order within a role preserved.
 *  2. A repeated row re-solves for the count rather than leaving a hole (three-cards, stat-band).
 *  3. Five or more wrap to two rows instead of shrinking.
 *  5. More blocks than slots: extras stack full width below the last slot.
 *  6. Nothing is dropped, ever.
 *  7. Emphasis wins the best slot of its role.
 *  8. Headline and kicker are furniture: never slots, never moved.
 */
import { DECK_COLS, DECK_ROWS, type Block, type BlockType } from "./types.js";

export type SlotRole = "text" | "visual" | "card" | "stat" | "list" | "table" | "quote" | "statement" | "caption";
export interface Slot { name: string; role: SlotRole; col: number; row: number; cols: number; rows: number }
/** `special` archetypes carry their own furniture (cover, divider, agenda) or need a real image; they are picked by name, never by cycling. */
export interface Archetype { name: string; slots: Slot[]; special?: boolean }

const S = (name: string, role: SlotRole, col: number, row: number, cols: number, rows: number): Slot => ({ name, role, col, row, cols, rows });

/** The catalogue, cell for cell as written in docs/research/deck-planning.md section 5. */
export const ARCHETYPES: Archetype[] = [
  { name: "three-cards", slots: [S("card-1", "card", 3, 7, 8, 10), S("card-2", "card", 13, 7, 8, 10), S("card-3", "card", 23, 7, 8, 10)] },
  { name: "two-by-two", slots: [S("card-1", "card", 3, 7, 13, 4), S("card-2", "card", 18, 7, 13, 4), S("card-3", "card", 3, 12, 13, 4), S("card-4", "card", 18, 12, 13, 4)] },
  { name: "two-column", slots: [S("left", "text", 3, 7, 13, 10), S("right", "text", 18, 7, 13, 10)] },
  { name: "text-visual", slots: [S("text", "text", 3, 7, 12, 10), S("visual", "visual", 17, 7, 14, 10)] },
  { name: "visual-text", slots: [S("visual", "visual", 3, 7, 14, 10), S("text", "text", 19, 7, 12, 10)] },
  { name: "chart-insight", slots: [S("chart", "visual", 3, 7, 18, 10), S("insight", "text", 23, 7, 8, 10), S("source", "caption", 3, 18, 14, 1)] },
  { name: "stat-band", slots: [S("stat-1", "stat", 2, 7, 8, 5), S("stat-2", "stat", 13, 7, 8, 5), S("stat-3", "stat", 24, 7, 8, 5), S("support", "text", 3, 13, 28, 4)] },
  { name: "hero-stat", slots: [S("stat", "stat", 3, 7, 11, 6), S("proof", "text", 16, 7, 15, 10)] },
  { name: "statement", slots: [S("statement", "statement", 3, 8, 22, 6), S("sub", "text", 3, 15, 22, 2)] },
  { name: "quote", slots: [S("quote", "quote", 5, 8, 24, 5), S("attribution", "caption", 5, 14, 24, 2)] },
  { name: "table", slots: [S("table", "table", 2, 7, 30, 9), S("source", "caption", 2, 17, 16, 1)] },
  { name: "points", slots: [S("list", "list", 3, 7, 20, 10), S("aside", "text", 24, 7, 7, 10)] },
  { name: "timeline-placeholder", slots: [S("track", "visual", 2, 8, 30, 5), S("caption", "text", 3, 14, 28, 2)] },
  { name: "matrix", slots: [S("plot", "visual", 12, 7, 10, 10), S("y-axis", "caption", 8, 7, 3, 10), S("x-axis", "caption", 12, 17, 10, 1), S("callouts", "text", 23, 7, 8, 10)] },
  { name: "image-full", special: true, slots: [S("image", "visual", 1, 1, 32, 18), S("overlay", "statement", 3, 12, 20, 4)] },
  { name: "cover", special: true, slots: [S("wordmark", "caption", 2, 2, 10, 2), S("title", "statement", 2, 10, 22, 3), S("subtitle", "text", 2, 13, 22, 2), S("date", "caption", 2, 15, 12, 1)] },
  { name: "divider", special: true, slots: [S("section", "statement", 2, 4, 28, 3), S("subtitle", "text", 2, 8, 22, 2)] },
  { name: "agenda", special: true, slots: [S("eyebrow", "caption", 2, 2, 10, 1), S("statement", "statement", 2, 4, 26, 3), S("entries-label", "caption", 2, 8, 10, 1), S("entry-1", "text", 2, 9, 28, 2), S("entry-2", "text", 2, 11, 28, 2), S("entry-3", "text", 2, 13, 28, 2), S("entry-4", "text", 2, 15, 28, 2)] },
];
export const ARCHETYPE_NAMES: string[] = ARCHETYPES.map((a) => a.name);
/** What the R key walks: every archetype that keeps the standing furniture. */
export const CYCLE: Archetype[] = ARCHETYPES.filter((a) => !a.special);

/** Furniture, rule 8: never a slot, never moved by a remix. */
export const FURNITURE_TYPES: ReadonlySet<BlockType> = new Set<BlockType>(["headline", "kicker"]);
const FURNITURE = FURNITURE_TYPES;

/**
 * The band a content block may occupy. Rows 1 to 4 are the kicker, the headline and the hairline;
 * row 18 is the source line and the page number. Nothing else may sit there (feedback F11).
 */
export const DECK_BODY_TOP = 5;
export const DECK_BODY_BOTTOM = 17;

/**
 * Slot roles a block type can take, best first. A chart lands in a visual slot, and in a table slot when
 * that is all there is. Most things can fill a card slot, which is what the doc's "fills per card slot"
 * lines say: text, stat, image, quote and placeholder all sit in a card.
 */
const PREF: Record<string, SlotRole[]> = {
  text: ["text", "list", "statement", "card", "caption"],
  bullets: ["list", "text", "card", "statement"],
  cards: ["card", "text", "visual"],
  stat: ["stat", "card", "text", "statement"],
  chart: ["visual", "table", "card"],
  diagram: ["visual", "card", "table", "text"],
  image: ["visual", "card", "table"],
  placeholder: ["visual", "card", "table", "text"],
  table: ["table", "visual", "card"],
  quote: ["quote", "card", "statement", "text"],
  takeaway: ["statement", "text", "caption"],
  notes: ["caption", "text"],
};
const ROLE_TITLE: Record<SlotRole, string> = { text: "Text", visual: "Visual", card: "Card", stat: "Stat", list: "List", table: "Table", quote: "Quote", statement: "Statement", caption: "Caption" };

const readingOrder = (a: Slot, b: Slot) => a.row - b.row || a.col - b.col;
const rank = (b: Block, role: SlotRole) => (PREF[b.type] ?? ["text"]).indexOf(role);
const bbox = (slots: Slot[]): Pick<Slot, "col" | "row" | "cols" | "rows"> => {
  const col = Math.min(...slots.map((s) => s.col)), row = Math.min(...slots.map((s) => s.row));
  return { col, row, cols: Math.max(...slots.map((s) => s.col + s.cols)) - col, rows: Math.max(...slots.map((s) => s.row + s.rows)) - row };
};

/**
 * Rules 2 and 3 for the two row archetypes. The geometry is the doc's own: a pair takes the two-column
 * split, four cards take the two-by-two grid, five or six wrap to the 3x2 grid (the second row centred),
 * four stats take the wide-box four-up. Any other count keeps the archetype's own slots.
 */
export function solveSlots(arch: Archetype, blocks: Block[]): Slot[] {
  const count = (role: SlotRole) => blocks.filter((b) => rank(b, role) === 0).length;
  if (arch.name === "three-cards") {
    const k = count("card");
    const card = (i: number, col: number, row: number, cols: number, rows: number) => S(`card-${i}`, "card", col, row, cols, rows);
    // "With a takeaway bar: cards become rows 8 (rows 7 to 14) and takeaway {col 3, row 16, cols 28, rows 2}"
    const bar = blocks.some((b) => b.type === "takeaway");
    const rows = bar ? 8 : 10;
    const band = bar ? [S("takeaway", "statement", 3, 16, 28, 2)] : [];
    if (k === 2) return [card(1, 3, 7, 13, rows), card(2, 18, 7, 13, rows), ...band];
    if (k === 4) return [card(1, 3, 7, 13, 4), card(2, 18, 7, 13, 4), card(3, 3, 12, 13, 4), card(4, 18, 12, 13, 4)];
    if (k === 5) return [card(1, 3, 7, 8, 4), card(2, 13, 7, 8, 4), card(3, 23, 7, 8, 4), card(4, 8, 12, 8, 4), card(5, 18, 12, 8, 4)];
    if (k >= 6) return [card(1, 3, 7, 8, 4), card(2, 13, 7, 8, 4), card(3, 23, 7, 8, 4), card(4, 3, 12, 8, 4), card(5, 13, 12, 8, 4), card(6, 23, 12, 8, 4)];
    if (bar) return [card(1, 3, 7, 8, rows), card(2, 13, 7, 8, rows), card(3, 23, 7, 8, rows), ...band];
  }
  if (arch.name === "stat-band" && count("stat") === 4) {
    return [S("stat-1", "stat", 2, 7, 6, 5), S("stat-2", "stat", 10, 7, 6, 5), S("stat-3", "stat", 18, 7, 6, 5), S("stat-4", "stat", 26, 7, 6, 5), S("support", "text", 3, 13, 28, 4)];
  }
  return arch.slots;
}

/** Keep every block on the slide: 32 columns by 18 rows, 1-based. */
function clamp(b: Block): Block {
  const cols = Math.max(1, Math.min(b.cols, DECK_COLS)), rows = Math.max(1, Math.min(b.rows, DECK_ROWS));
  return { ...b, cols, rows, col: Math.min(Math.max(1, b.col), DECK_COLS - cols + 1), row: Math.min(Math.max(1, b.row), DECK_ROWS - rows + 1) };
}

/** The archetype these blocks are already laid out as, or undefined when the slide is hand-arranged. */
export function detectArchetype(blocks: Block[]): string | undefined {
  const movable = blocks.filter((b) => !FURNITURE.has(b.type));
  if (!movable.length) return undefined;
  // Own slots first, across the whole catalogue, then the re-solved rows: a 2x2 of cards is two-by-two,
  // not three-cards re-solved for four, even though the two geometries are the same.
  for (const own of [true, false]) {
    for (const arch of ARCHETYPES) {
      const slots = own ? arch.slots : solveSlots(arch, movable);
      const left = [...slots];
      const cards = left.filter((s) => s.role === "card");
      const union = cards.length > 1 ? bbox(cards) : null;
      const ok = movable.every((b) => {
        const i = left.findIndex((s) => s.col === b.col && s.row === b.row && s.cols === b.cols && s.rows === b.rows);
        if (i >= 0) { left.splice(i, 1); return true; }
        // one cards block spanning the whole card band (three-cards with count 3)
        if (union && b.col === union.col && b.row === union.row && b.cols === union.cols && b.rows === union.rows) {
          for (const c of cards) { const j = left.indexOf(c); if (j >= 0) left.splice(j, 1); }
          return true;
        }
        return false;
      });
      if (ok) return arch.name;
    }
  }
  return undefined;
}

/** The archetype after `current` in the cycle; the first one when the slide is hand-arranged or on a special archetype. */
export function nextArchetype(current?: string): string {
  const i = CYCLE.findIndex((a) => a.name === current);
  return CYCLE[(i + 1) % CYCLE.length].name;
}

export interface RemixOptions {
  /** Fill an empty slot with a `placeholder` block. Remix does; the server's auto-layout never invents blocks. */
  placeholders?: boolean;
  /**
   * Where a slot can take several block types equally, follow the order the blocks arrived in rather than
   * the role ranking. The auto-layout does: Claude sends blocks in reading order and expects them back in it.
   */
  orderFirst?: boolean;
}

/**
 * Re-lay `blocks` as `archetype`, or as the next one in the cycle. Returns new block objects; the input
 * is left alone. Furniture keeps its geometry, every block keeps its id, title and content, missing slots
 * arrive as placeholders and anything left over stacks below the body.
 */
export function remix(blocks: Block[], archetype?: string, opts: RemixOptions = { placeholders: true }): { blocks: Block[]; archetype: string } {
  const furniture = blocks.filter((b) => FURNITURE.has(b.type)).map((b) => ({ ...b }));
  const movable = blocks.filter((b) => !FURNITURE.has(b.type));
  const name = archetype ?? nextArchetype(detectArchetype(blocks));
  const arch = ARCHETYPES.find((a) => a.name === name) ?? CYCLE[0];
  const slots = [...solveSlots(arch, movable)].sort(readingOrder);

  // Two passes so a caption slot cannot eat the text block a later text slot needs: exact roles first,
  // then the best of what is left. Within a role, emphasis first (rule 7), otherwise reading order.
  const taken = new Map<Slot, Block>();
  const used = new Set<Block>();
  const pick = (slot: Slot, exact: boolean) => {
    const pool = movable.filter((b) => !used.has(b) && (exact ? rank(b, slot.role) === 0 : rank(b, slot.role) >= 0));
    if (!pool.length) return;
    const best = opts.orderFirst && !exact ? pool[0] : pool.reduce((a, b) => {
      const ra = rank(a, slot.role), rb = rank(b, slot.role);
      if (ra !== rb) return ra < rb ? a : b;
      if (!!a.emphasis !== !!b.emphasis) return a.emphasis ? a : b;
      return movable.indexOf(a) < movable.indexOf(b) ? a : b;
    });
    taken.set(slot, best); used.add(best);
  };
  for (const s of slots) pick(s, true);
  for (const s of slots) if (!taken.has(s)) pick(s, false);

  // One cards block fills the whole card band rather than sitting in the first card slot alone.
  const spare = new Set(slots.filter((s) => !taken.has(s)));
  const only = movable.filter((b) => b.type === "cards");
  const cardSlots = slots.filter((s) => s.role === "card");
  let span: Pick<Slot, "col" | "row" | "cols" | "rows"> | null = null;
  if (only.length === 1 && cardSlots.length > 1 && cardSlots.some((s) => taken.get(s) === only[0])) {
    const rest = cardSlots.filter((s) => spare.has(s));
    if (rest.length) { span = bbox(cardSlots); for (const s of rest) spare.delete(s); }
  }

  const out: Block[] = [...furniture];
  const seen = new Set(blocks.map((b) => b.id));
  for (const s of slots) {
    const b = taken.get(s);
    if (b) {
      const geo = span && s.role === "card" ? span : s;
      const next: Block = { ...b, col: geo.col, row: geo.row, cols: geo.cols, rows: geo.rows };
      if (b.type === "cards" && next.count == null) next.count = Math.max(2, cardSlots.length);
      out.push(clamp(next));
      continue;
    }
    if (!opts.placeholders || !spare.has(s)) continue;
    const dupes = slots.filter((x) => x.role === s.role);
    const label = dupes.length > 1 ? `${ROLE_TITLE[s.role]} ${dupes.indexOf(s) + 1}` : ROLE_TITLE[s.role];
    let id = `ph_${arch.name}_${s.name}`;
    while (seen.has(id)) id += "x";
    seen.add(id);
    out.push(clamp({ id, type: "placeholder", title: label, col: s.col, row: s.row, cols: s.cols, rows: s.rows }));
  }

  // Rule 5: extras stack under the body in two-row bands, full width. When the bands run out before the
  // extras do they share the last ones side by side rather than piling up off the slide.
  const extras = movable.filter((b) => !used.has(b));
  if (extras.length) {
    const bottom = slots.length ? Math.max(...slots.map((s) => s.row + s.rows - 1)) : 6;
    const start = Math.min(DECK_ROWS - 1, bottom + 1);
    const bands = Math.max(1, Math.floor((DECK_ROWS - start + 1) / 2));
    const per = Math.ceil(extras.length / bands);
    const w = Math.max(1, Math.floor((28 - (per - 1)) / per));
    extras.forEach((b, i) => {
      out.push(clamp({ ...b, col: 3 + (i % per) * (w + 1), cols: w, row: start + Math.floor(i / per) * 2, rows: 2 }));
    });
  }
  return { blocks: out, archetype: arch.name };
}

// ---------- the solver the server lays a proposal out with ----------

/**
 * The archetype the shape of this evidence asks for, or undefined when nothing in the catalogue fits
 * and the blocks should just stack in the body band. Roles and counts only: three parallel things, two
 * contrasted things, one chart plus its so-what, a band of stats.
 */
export function pickArchetype(blocks: Block[]): string | undefined {
  const body = blocks.filter((b) => !FURNITURE.has(b.type));
  if (!body.length) return undefined;
  const of = (type: string) => body.filter((b) => b.type === type);
  // A takeaway bar and a notes line are furniture-ish: they ride along, they do not pick the archetype.
  const core = body.filter((b) => b.type !== "takeaway" && b.type !== "notes");
  const k = core.length;
  if (!k) return undefined;
  const visuals = core.filter((b) => rank(b, "visual") === 0);
  const texts = core.filter((b) => b.type === "text" || b.type === "bullets");
  const stats = of("stat").filter((b) => core.includes(b));
  const cards = of("cards").filter((b) => core.includes(b));

  // One cards block carries its own count; 4 cards read better as a 2x2 than as a four-up row.
  if (cards.length === 1 && k === 1) return (cards[0].count ?? 3) === 4 ? "two-by-two" : "three-cards";
  if (stats.length >= 3 && stats.length === k) return "stat-band";
  if (k === 1) {
    const b = core[0];
    if (b.type === "table") return "table";
    if (b.type === "quote") return "quote";
    if (b.type === "stat") return "hero-stat";
    return undefined; // one block owns the whole body band
  }
  if (k === 2) {
    if (visuals.length === 1 && texts.length === 1) {
      const chart = visuals[0].type === "chart" || visuals[0].type === "diagram";
      if (chart) return "chart-insight";
      return core.indexOf(visuals[0]) === 0 ? "visual-text" : "text-visual";
    }
    if (stats.length === 1 && texts.length === 1) return "hero-stat";
    if (of("quote").length === 1) return "quote";
    return "two-column";
  }
  if (k === 3) return "three-cards";
  if (k === 4) return "two-by-two";
  return undefined; // five or more: stack in bands
}

/** Pull every block into rows 5 to 17: the furniture band above and the footer row below are not ours. */
export function clampToBody(blocks: Block[]): void {
  for (const b of blocks) {
    b.cols = Math.max(1, Math.min(Math.round(b.cols) || 1, DECK_COLS));
    b.rows = Math.max(1, Math.min(Math.round(b.rows) || 1, DECK_BODY_BOTTOM - DECK_BODY_TOP + 1));
    b.col = Math.min(Math.max(1, Math.round(b.col) || 1), DECK_COLS - b.cols + 1);
    b.row = Math.min(Math.max(DECK_BODY_TOP, Math.round(b.row) || DECK_BODY_TOP), DECK_BODY_BOTTOM - b.rows + 1);
  }
}

const hits = (a: Block, b: Block) => a.row < b.row + b.rows && b.row < a.row + a.rows && a.col < b.col + b.cols && b.col < a.col + a.cols;

/**
 * Push overlapping blocks apart inside the body band: the later block in reading order moves down, and
 * right when there is no room below. Same idea as the canvas's resolveOverlaps, with the band as the wall.
 */
export function resolveBandOverlaps(blocks: Block[], maxPasses = 20): boolean {
  const order = [...blocks].sort((a, b) => a.row - b.row || a.col - b.col);
  let changed = false;
  for (let pass = 0; pass < maxPasses; pass++) {
    let any = false;
    for (let i = 1; i < order.length; i++) {
      const b = order[i];
      const over = order.slice(0, i).filter((a) => hits(a, b));
      if (!over.length) continue;
      const below = Math.max(...over.map((a) => a.row + a.rows));
      if (below + b.rows - 1 <= DECK_BODY_BOTTOM) { b.row = below; any = true; changed = true; continue; }
      const right = Math.max(...over.map((a) => a.col + a.cols));
      if (right + b.cols - 1 <= DECK_COLS) { b.col = right; any = true; changed = true; continue; }
      // No room either way: shrink onto the last free rows rather than sit on top of another block.
      const top = Math.min(DECK_BODY_BOTTOM, below);
      if (b.row !== top || b.rows !== Math.max(1, DECK_BODY_BOTTOM - top + 1)) {
        b.row = top; b.rows = Math.max(1, DECK_BODY_BOTTOM - top + 1); any = true; changed = true;
      }
    }
    if (!any) break;
  }
  return changed;
}

/**
 * Lay these blocks out: as `archetype` when one is given, else as the archetype their roles ask for, else
 * stacked in bands down the body. Returns new block objects in the input's order, nothing added, nothing
 * dropped, everything inside rows 5 to 17.
 */
export function layoutBlocks(blocks: Block[], archetype?: string): Block[] {
  if (!blocks.length) return [];
  const name = archetype ?? pickArchetype(blocks);
  const out = new Map<string, Block>();
  if (name) {
    for (const b of remix(blocks, name, { placeholders: false, orderFirst: true }).blocks) out.set(b.id, b);
  } else {
    const furniture = blocks.filter((b) => FURNITURE.has(b.type));
    for (const b of furniture) out.set(b.id, { ...b });
    for (const b of stackInBody(blocks.filter((b) => !FURNITURE.has(b.type)))) out.set(b.id, b);
  }
  const laid = blocks.map((b) => out.get(b.id) ?? { ...b });
  const body = laid.filter((b) => !FURNITURE.has(b.type));
  clampToBody(body);
  resolveBandOverlaps(body);
  return laid;
}

/** The fallback: full-width bands down rows 5 to 17, two rows each once there are too many to breathe. */
function stackInBody(blocks: Block[]): Block[] {
  const span = DECK_BODY_BOTTOM - DECK_BODY_TOP + 1;
  const n = blocks.length;
  if (!n) return [];
  let rows = Math.floor((span - (n - 1)) / n);
  let per = 1;
  if (rows < 2) { rows = 2; per = Math.ceil(n / Math.max(1, Math.floor((span + 1) / (rows + 1)))); }
  const cols = per > 1 ? Math.max(1, Math.floor((28 - (per - 1)) / per)) : 28;
  return blocks.map((b, i) => {
    const row = DECK_BODY_TOP + Math.floor(i / per) * (rows + 1);
    return clamp({ ...b, col: 3 + (i % per) * (cols + 1), cols, row: Math.min(row, DECK_BODY_BOTTOM - rows + 1), rows });
  });
}

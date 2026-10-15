import fs from "node:fs";
import path from "node:path";
import { PLANNER_HOME } from "./playwright.config";

export const DEMO_ID = "demo";
export const TAB_IDS = { readme: "tab_readme", assumptions: "tab_assump", funnel: "tab_funnel", ramp: "tab_ramp" };
export const TAB_NAMES = ["Read Me", "Assumptions", "Market Funnel", "Revenue Ramp"];

export const DECK_ID = "deck";
export const SLIDE_IDS = { where: "sl_where", size: "sl_size", gap: "sl_gap", move: "sl_move", proof: "sl_proof", ask: "sl_ask" };
export const SLIDE_NAMES = ["Where we play", "The market", "The gap", "Three moves", "Why it works", "The ask"];
export const SECTIONS = ["Today", "The plan"];

const NOW = "2026-01-01T00:00:00.000Z";

function tab(id: string, name: string, purpose: string, approved: boolean, message?: string) {
  return { id, name, purpose, message, approved, layoutProposed: false, layoutApproved: false, blocks: [] as unknown[] };
}

/** The demo project as it looks after propose_outline: four tabs, a part-filled brief, one resolved note. */
export function demoProject(opts: { tabsApproved?: boolean } = {}) {
  const approved = opts.tabsApproved ?? false;
  return {
    id: DEMO_ID,
    name: "Demo Market Sizing",
    kind: "workbook",
    brief: "Market sizing model for a B2B payments product aimed at US SMBs",
    briefCard: { audience: "The CFO", goal: "Size the US SMB payments market" }, // no takeaway: the readiness dot shows
    formatting: "house",
    skipLayouts: false,
    tabsApproved: approved,
    tabs: [
      tab(TAB_IDS.readme, TAB_NAMES[0], "Cover page generated at build", approved),
      tab(TAB_IDS.assumptions, TAB_NAMES[1], "Inputs everything else reads", approved, "Every number traces to one input"),
      tab(TAB_IDS.funnel, TAB_NAMES[2], "Top-down TAM to SOM", approved),
      tab(TAB_IDS.ramp, TAB_NAMES[3], "Five-year revenue build from share and price", approved, "Year 5 revenue is the number to argue with"),
    ],
    notes: [
      { id: "note_seed1", tabId: TAB_IDS.assumptions, text: "Keep every input on this tab, none inline", status: "resolved", response: "Done: all inputs live on Assumptions.", createdBy: "user", createdAt: NOW, resolvedAt: NOW },
    ],
    revision: 1,
    createdAt: NOW,
    updatedAt: NOW,
    renderSeq: 0,
  };
}

function slide(id: string, name: string, section: string, story: string, message: string) {
  return { id, name, purpose: name, section, story, message, approved: false, layoutProposed: false, layoutApproved: false, blocks: [] as unknown[] };
}

/** A deck outline: six slides, stories, two sections, a full brief. The long story exercises the expander. */
export function deckProject() {
  return {
    id: DECK_ID,
    name: "Payments Partner Review",
    kind: "deck",
    brief: "Partner review deck arguing for three moves in SMB payments",
    briefCard: { audience: "The partner group", goal: "Win approval for three moves", takeaway: "We win SMB payments by owning onboarding", emphasis: ["Onboarding", "Year 5 revenue"] },
    formatting: "house",
    skipLayouts: false,
    tabsApproved: false,
    tabs: [
      slide(SLIDE_IDS.where, SLIDE_NAMES[0], SECTIONS[0], "Sets the frame: we sell payments to US SMBs and nothing else. Names the three segments the rest of the deck returns to. Ends on the one we are losing.", "We compete in one market, not four"),
      slide(SLIDE_IDS.size, SLIDE_NAMES[1], SECTIONS[0], "Sizes the opportunity top down, from all US businesses to the ones we can actually serve. Shows the funnel as three steps so the partner can argue with each one. The number that matters is the serviceable count, not the headline TAM, and the appendix carries the census sources behind every step of the build.", "Six million businesses are reachable today"),
      slide(SLIDE_IDS.gap, SLIDE_NAMES[2], SECTIONS[0], "Contrasts our onboarding time with the two competitors who are taking share. Makes the loss concrete in weeks and in lost accounts.", "We lose on week one, not on price"),
      slide(SLIDE_IDS.move, SLIDE_NAMES[3], SECTIONS[1], "Lays out the three moves as cards so they read as one plan. Each card carries the owner and the first milestone.", "Three moves, one owner each"),
      slide(SLIDE_IDS.proof, SLIDE_NAMES[4], SECTIONS[1], "Backs the plan with the pilot result and the ramp it implies. Keeps the chart to one series so the claim is unmissable.", "The pilot already moved onboarding to two days"),
      slide(SLIDE_IDS.ask, SLIDE_NAMES[5], SECTIONS[1], "States the ask, the money and the decision date. Nothing else on the slide.", "Approve the three moves and the budget"),
    ],
    notes: [
      { id: "note_deck1", tabId: SLIDE_IDS.gap, kind: "question", text: "Which two competitors should this name?", status: "open", createdBy: "claude", createdAt: NOW },
    ],
    revision: 1,
    createdAt: NOW,
    updatedAt: NOW,
    renderSeq: 0,
  };
}

export const DECK_DIV_ID = "deckdiv";
export const DECK_LONG_ID = "decklong";
export const DIVIDER_IDS = { today: "sl_div1", plan: "sl_div2", appendix: "sl_div3" };

/** A divider row: its name is the section title, and it carries nothing else (feedback F4). */
function divider(id: string, name: string) {
  return { id, name, purpose: "", divider: true, section: name, approved: false, layoutProposed: false, layoutApproved: false, blocks: [] as unknown[] };
}

/**
 * The same deck once sections are rows: two divider rows, six slides, and every slide's `section` set to
 * what the server re-derives from the divider above it.
 */
export function deckDividerProject(opts: { id?: string } = {}) {
  const base = deckProject();
  // The story that exercises the expander runs past two lines at desktop width, where the row is widest.
  const long = { ...base.tabs[1], story: `${base.tabs[1].story} Every step names the source it came from, so the partner can take the argument apart one line at a time and still land on the same count.` };
  return {
    ...base,
    id: opts.id ?? DECK_DIV_ID,
    tabs: [divider(DIVIDER_IDS.today, SECTIONS[0]), base.tabs[0], long, base.tabs[2], divider(DIVIDER_IDS.plan, SECTIONS[1]), ...base.tabs.slice(3)],
  };
}

/** Long enough that the outline list scrolls: the divider deck with its slides repeated under a third section. */
export function longDeckProject() {
  const base = deckDividerProject({ id: DECK_LONG_ID });
  const again = base.tabs.filter((t) => !t.divider).map((t, i) => ({ ...t, id: `${t.id}_b`, name: `${t.name} again`, section: "Appendix" }));
  return { ...base, tabs: [...base.tabs, divider(DIVIDER_IDS.appendix, "Appendix"), ...again] };
}

// ---------------------------------------------------------------- a built deck
//
// build_deck writes the .pptx, the sibling snapshot and the PDF, then renders the built view with the
// snapshot. Until that tool lands, this writes the same three things by hand: the project JSON with
// buildPath / snapshotPath / pdfPath / buildSummary, and the snapshot file per builder/SPEC-deck.md.

export const DECK_BUILT_ID = "deckbuilt";
export const CELL = 30, SLIDE = { w: 960, h: 540 };
/** Slide kinds for the six seeded slides: a cover, four content slides, a closing divider. */
const KINDS = ["title", "content", "content", "content", "content", "divider"];
const HEADLINES = [
  "Payments Partner Review",
  "Six million businesses are reachable today",
  "We lose on week one, not on price",
  "Three moves, one owner each",
  "The pilot already moved onboarding to two days",
  "The ask",
];

interface SeedBlock { id: string; type: string; title: string; row: number; col: number; rows: number; cols: number; content?: string; count?: number; emphasis?: boolean }
/** Plan blocks per slide, in slide-grid cells (32 columns by 18 rows). */
const DECK_BLOCKS: SeedBlock[][] = [
  [],
  [{ id: "b_funnel", type: "chart", title: "Market funnel", row: 6, col: 3, rows: 10, cols: 16, content: "three steps from all businesses to reachable" },
   { id: "b_reach", type: "stat", title: "Reachable today", row: 6, col: 21, rows: 4, cols: 10, count: 2, emphasis: true }],
  [{ id: "b_breaks", type: "cards", title: "Where it breaks", row: 6, col: 3, rows: 7, cols: 28, count: 3, content: "the three moments referrals stall" },
   { id: "b_week", type: "takeaway", title: "Week one is the loss", row: 16, col: 3, rows: 2, cols: 28 }],
  [{ id: "b_moves", type: "cards", title: "The three moves", row: 6, col: 3, rows: 8, cols: 28, count: 3, content: "owner and first milestone per move" },
   { id: "b_notes", type: "notes", title: "Speaker notes", row: 15, col: 3, rows: 3, cols: 28 }],
  [{ id: "b_ramp", type: "chart", title: "Pilot ramp", row: 6, col: 3, rows: 10, cols: 16 },
   { id: "b_voice", type: "quote", title: "Partner quote", row: 6, col: 21, rows: 5, cols: 10 }],
  [],
];
/** The one block the builder did not put where the plan asked, so the plan diff has something to say. */
const MOVED = { id: "b_moves", rowShift: 1 };

const CARD_TEXT = ["No moment Timing Partners never hear a reason to refer", "No reward Incentive Nothing lands when they do", "No proof Evidence They cannot see it worked"];
const MOVE_TEXT = ["Own onboarding Ana Two-day pilot by cohort two", "Pay for proof Devi Referral credit live", "Show the ramp Sam Partner dashboard"];
const STAT_TEXT = ["6.4M", "42%"];

function shape(o: { id: string; blockId?: string | null; role: string; kind: string; x: number; y: number; w: number; h: number; text?: string; fontPt?: number }) {
  return { blockId: null, ...o };
}

/** Furniture plus one entry per drawn shape, the way build_pptx.py reports it. */
function slideShapes(i: number): ReturnType<typeof shape>[] {
  const n = i + 1;
  const kind = KINDS[i];
  const out: ReturnType<typeof shape>[] = [];
  let seq = 0;
  const id = () => `s${n}-${++seq}`;
  if (kind === "title") {
    out.push(shape({ id: id(), role: "wordmark", kind: "text", x: 60, y: 60, w: 300, h: 18, text: "Blueprint", fontPt: 12 }));
    out.push(shape({ id: id(), role: "coverTitle", kind: "text", x: 60, y: 200, w: 780, h: 120, text: HEADLINES[i], fontPt: 40 }));
    out.push(shape({ id: id(), role: "coverSubtitle", kind: "text", x: 60, y: 332, w: 600, h: 40, text: "Three moves in SMB payments", fontPt: 18 }));
    out.push(shape({ id: id(), role: "coverDate", kind: "text", x: 60, y: 470, w: 400, h: 16, text: "Draft for discussion", fontPt: 11 }));
    return out;
  }
  if (kind === "divider") {
    out.push(shape({ id: id(), role: "dividerTitle", kind: "text", x: 60, y: 210, w: 780, h: 90, text: HEADLINES[i], fontPt: 34 }));
    out.push(shape({ id: id(), role: "dividerSubtitle", kind: "text", x: 60, y: 312, w: 600, h: 40, text: "Approve the three moves and the budget", fontPt: 16 }));
    out.push(shape({ id: id(), role: "pageNumber", kind: "text", x: 870, y: 514, w: 60, h: 14, text: String(n), fontPt: 9 }));
    return out;
  }
  out.push(shape({ id: id(), role: "kicker", kind: "text", x: 60, y: 42, w: 600, h: 14, text: `${SECTIONS[i < 3 ? 0 : 1]} | ${SLIDE_NAMES[i]}`, fontPt: 11 }));
  out.push(shape({ id: id(), role: "headline", kind: "text", x: 60, y: 66, w: 840, h: 84, text: HEADLINES[i], fontPt: 26 }));
  out.push(shape({ id: id(), role: "hairline", kind: "rule", x: 60, y: 156, w: 840, h: 1 }));
  for (const b of DECK_BLOCKS[i]) {
    const shift = MOVED.id === b.id ? MOVED.rowShift : 0;
    const x = (b.col - 1) * CELL, y = (b.row - 1 + shift) * CELL, w = b.cols * CELL, h = b.rows * CELL;
    if (b.type === "cards" || b.type === "stat") {
      const items = b.count ?? 3;
      const gut = 27, each = (w - gut * (items - 1)) / items;
      const text = b.type === "stat" ? STAT_TEXT : b.id === "b_moves" ? MOVE_TEXT : CARD_TEXT;
      for (let k = 0; k < items; k++) {
        out.push(shape({ id: id(), blockId: b.id, role: b.type === "stat" ? "statValue" : "card", kind: b.type === "stat" ? "stat" : "card", x: x + k * (each + gut), y, w: each, h, text: text[k % text.length], fontPt: b.type === "stat" ? 32 : 17 }));
      }
      continue;
    }
    const role = b.type === "chart" ? "chart" : b.type === "table" ? "table" : b.type === "takeaway" ? "takeaway" : b.type === "quote" ? "quote" : b.type === "notes" ? "notes" : "body";
    const text = b.type === "chart" ? undefined : b.type === "takeaway" ? "Fix week one and the rest follows" : b.type === "quote" ? "Two days felt like a different company" : b.type === "notes" ? "Walk the cards left to right" : b.content;
    out.push(shape({ id: id(), blockId: b.id, role, kind: b.type, x, y, w, h, text, fontPt: b.type === "takeaway" ? 18 : b.type === "quote" ? 20 : 14 }));
  }
  out.push(shape({ id: id(), role: "source", kind: "text", x: 60, y: 514, w: 600, h: 14, text: "Source: Partner CRM", fontPt: 9 }));
  out.push(shape({ id: id(), role: "pageNumber", kind: "text", x: 870, y: 514, w: 60, h: 14, text: String(n), fontPt: 9 }));
  return out;
}

export function deckSnapshot() {
  const ids = Object.values(SLIDE_IDS);
  return {
    name: "Payments Partner Review",
    slideSize: SLIDE,
    grid: { cols: 32, rows: 18, cell: CELL },
    style: { mode: "house", name: "House" },
    slides: SLIDE_NAMES.map((name, i) => ({
      index: i + 1, id: ids[i], name, kind: KINDS[i], section: SECTIONS[i < 3 ? 0 : 1],
      headline: HEADLINES[i], shapes: slideShapes(i),
    })),
  };
}

/** Findings from check_deck.py: one error, two warnings, one of them furniture (no blockId). */
export const DECK_FINDINGS = [
  { slide: 2, role: "headline", kind: "widow", severity: "warning", message: "the headline's last line is one word" },
  { slide: 4, blockId: "b_moves", role: "cardBody", kind: "overflow", severity: "warning", message: "block 'The three moves' needs 7 lines and the box holds 5" },
  { slide: 5, blockId: "b_voice", role: "quote", kind: "off-slide", severity: "error", message: "the quote sits 12 pt past the right edge" },
];

/** The deck project as it looks after build_deck: layouts approved, threads closed, a file on disk. */
export function deckBuiltProject(opts: { pdf?: boolean } = {}) {
  const base = deckProject();
  const snapshot = deckSnapshot();
  const out = path.join(PLANNER_HOME, "out");
  const buildPath = path.join(out, "Payments Partner Review.pptx");
  const pdfPath = path.join(out, "Payments Partner Review.pdf");
  return {
    ...base,
    id: DECK_BUILT_ID,
    tabsApproved: true,
    tabs: base.tabs.map((t, i) => ({ ...t, approved: true, layoutProposed: true, layoutApproved: true, blocks: DECK_BLOCKS[i] })),
    notes: [
      { id: "note_db1", tabId: SLIDE_IDS.gap, kind: "question", text: "Which two competitors should this name?", status: "resolved", response: "Stripe and Adyen.", createdBy: "claude", createdAt: NOW, resolvedAt: NOW },
      { id: "note_db2", tabId: SLIDE_IDS.move, kind: "instruction", text: "One owner per card, no committees", status: "resolved", response: "Done: each card names one owner.", createdBy: "user", createdAt: NOW, resolvedAt: NOW },
    ],
    buildPath,
    snapshotPath: path.join(out, "Payments Partner Review.snapshot.json"),
    pdfPath: opts.pdf === false ? undefined : pdfPath,
    thumbnails: SLIDE_NAMES.map((_, i) => path.join(out, "Payments Partner Review.slides", `slide-${i + 1}.png`)),
    buildSummary: {
      kind: "deck",
      slides: snapshot.slides.map((s) => ({ index: s.index, tabId: s.id, name: s.name, headline: s.headline, shapes: s.shapes.length, thumbnail: path.join(out, "Payments Partner Review.slides", `slide-${s.index}.png`) })),
      findings: DECK_FINDINGS,
      warnings: ["Lato is not installed; using Arial", "Slide 4 body stepped down to 14 pt"],
      pdfPath: opts.pdf === false ? undefined : pdfPath,
      // The canvas reads the snapshot from the build's own result; this copy is what survives a re-render.
      snapshot,
    },
  };
}

/** Write the built deck: the project JSON, plus the snapshot file where snapshotPath says it is. */
export function seedDeckBuilt(opts: { pdf?: boolean } = {}) {
  const project = deckBuiltProject(opts);
  fs.mkdirSync(path.join(PLANNER_HOME, "out"), { recursive: true });
  fs.writeFileSync(String(project.snapshotPath), JSON.stringify(deckSnapshot(), null, 2));
  seed(project);
  return project;
}

export function projectPath(id = DEMO_ID): string {
  return path.join(PLANNER_HOME, "projects", `${id}.json`);
}

/** Write a project JSON to the planner home. The server reads from disk on every call, so this is a full reset. */
export function seed(project: Record<string, unknown> = demoProject()): void {
  fs.mkdirSync(path.join(PLANNER_HOME, "projects"), { recursive: true });
  fs.writeFileSync(projectPath(String(project.id)), JSON.stringify(project, null, 2));
}

export function readProject(id = DEMO_ID): any {
  return JSON.parse(fs.readFileSync(projectPath(id), "utf8"));
}

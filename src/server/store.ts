import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import crypto from "node:crypto";
import type { Block, Brief, BuildOptions, DeckBlockType, Diagram, DiagramEdge, DiagramKind, DiagramNode, Kind, Note, Project, Stage, Tab, WorkbookBlockType } from "../shared/types.js";
import { BUILD_DEFAULTS, DECK_COLS, DECK_ROWS } from "../shared/types.js";
import { DECK_BODY_BOTTOM, DECK_BODY_TOP, clampToBody, layoutBlocks, resolveBandOverlaps } from "../shared/archetypes.js";

const DEFAULT_HOME = path.join(os.homedir(), ".blueprint");
export const HOME = process.env.BLUEPRINT_HOME || process.env.WORKBOOK_PLANNER_HOME || DEFAULT_HOME;
export const PROJECTS_DIR = path.join(HOME, "projects");
export const OUTPUT_DIR = process.env.BLUEPRINT_OUTPUT || process.env.WORKBOOK_PLANNER_OUTPUT || path.join(os.homedir(), "Documents", "Blueprint");

/** First run under the new name: carry a Workbook Planner home over, once. */
if (HOME === DEFAULT_HOME && !fs.existsSync(PROJECTS_DIR)) {
  const old = path.join(os.homedir(), ".workbook-planner");
  if (fs.existsSync(old)) {
    try {
      fs.mkdirSync(HOME, { recursive: true });
      for (const entry of ["projects", "templates", "styles", "defaults.json"]) {
        const from = path.join(old, entry);
        if (fs.existsSync(from)) fs.cpSync(from, path.join(HOME, entry), { recursive: true });
      }
      console.error(`Blueprint: copied projects, templates and defaults from ${old} to ${HOME}.`);
    } catch (e: any) {
      console.error(`Blueprint: could not copy ${old} to ${HOME}: ${e?.message ?? e}`);
    }
  }
}

fs.mkdirSync(PROJECTS_DIR, { recursive: true });

export function newId(prefix: string): string {
  return `${prefix}_${crypto.randomBytes(3).toString("hex")}`;
}

export function slugify(name: string): string {
  return name.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "").slice(0, 40) || "project";
}

function projectPath(id: string): string {
  return path.join(PROJECTS_DIR, `${id}.json`);
}

export function listProjects(): (Pick<Project, "id" | "name" | "kind" | "updatedAt" | "revision"> & { stage: Stage })[] {
  return fs
    .readdirSync(PROJECTS_DIR)
    .filter((f) => f.endsWith(".json"))
    .map((f) => JSON.parse(fs.readFileSync(path.join(PROJECTS_DIR, f), "utf8")) as Project)
    .map((p) => ({ id: p.id, name: p.name, kind: p.kind ?? "workbook", updatedAt: p.updatedAt, revision: p.revision, stage: stageOf(p) }))
    .sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : -1));
}

export function loadProject(id: string): Project {
  const p = projectPath(id);
  if (!fs.existsSync(p)) throw new Error(`No project with id "${id}". Call list_projects to see existing projects or start_project to create one.`);
  const project = JSON.parse(fs.readFileSync(p, "utf8")) as Project;
  migrate(project);
  return project;
}

export function saveProject(project: Project): Project {
  normalizeOutline(project);
  project.revision += 1;
  project.updatedAt = new Date().toISOString();
  const tmp = projectPath(project.id) + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(project, null, 2));
  fs.renameSync(tmp, projectPath(project.id));
  return project;
}

const DEFAULTS_FILE = path.join(HOME, "defaults.json");
export interface Defaults { styleId?: string; formatting?: Project["formatting"]; outputDir?: string; templatePath?: string }
export function loadDefaults(): Defaults { try { return JSON.parse(fs.readFileSync(DEFAULTS_FILE, "utf8")); } catch { return {}; } }
export function saveDefaults(d: Defaults): void { fs.writeFileSync(DEFAULTS_FILE, JSON.stringify({ ...loadDefaults(), ...d }, null, 2)); }

/** The brief is three fields. Emphasis chips are deprecated: accepted from an old payload, never stored (F7). */
export function cleanBrief(b: Brief | undefined): Brief | undefined {
  if (!b) return undefined;
  const out: Brief = { audience: b.audience?.trim() || undefined, goal: b.goal?.trim() || undefined, takeaway: b.takeaway?.trim() || undefined };
  return out.audience || out.goal || out.takeaway ? out : undefined;
}

// ---------- Build options ----------

/** The options with every default filled in. `p.build` absent means all four defaults. */
export function buildOptions(p: Project): Required<BuildOptions> {
  return { ...BUILD_DEFAULTS, ...(p.build ?? {}) };
}

/** Patch the options, keeping the fields the caller left out. Nothing patched leaves `p.build` alone. */
export function setBuildOptions(p: Project, patch: BuildOptions): Required<BuildOptions> {
  const next: BuildOptions = { ...(p.build ?? {}) };
  for (const [key, value] of Object.entries(patch)) if (value !== undefined) (next as Record<string, unknown>)[key] = value;
  p.build = Object.keys(next).length ? next : undefined;
  return buildOptions(p);
}

/** One line for the model, only when the build is not left entirely on its defaults. */
export function buildLine(p: Project): string {
  const b = buildOptions(p);
  const defaulted = (Object.keys(BUILD_DEFAULTS) as (keyof Required<BuildOptions>)[]).every((k) => b[k] === BUILD_DEFAULTS[k]);
  if (defaulted) return "";
  const flag = (on: boolean, label: string) => (on ? label : `no ${label}`);
  return `build: ${b.fidelity}, ${flag(b.notesFromStory, "notes from story")}, ${flag(b.cover, "cover")}, ${flag(b.agenda, "agenda")}, ${flag(b.comments, "comments")}`;
}

export function createProject(input: { name: string; brief: string; kind?: Kind; briefCard?: Brief; formatting?: Project["formatting"]; templatePath?: string; outputDir?: string; skipLayouts?: boolean }): Project {
  let id = slugify(input.name);
  if (fs.existsSync(projectPath(id))) id = `${id}-${crypto.randomBytes(2).toString("hex")}`;
  const now = new Date().toISOString();
  const d = loadDefaults();
  const project: Project = {
    id,
    name: input.name,
    kind: input.kind ?? "workbook",
    brief: input.brief,
    briefCard: cleanBrief(input.briefCard),
    formatting: input.formatting ?? d.formatting ?? "plain",
    styleId: input.formatting ?? d.styleId ?? d.formatting ?? "plain",
    templatePath: input.templatePath ?? d.templatePath,
    outputDir: input.outputDir ?? d.outputDir,
    skipLayouts: input.skipLayouts ?? false,
    tabsApproved: false,
    tabs: [],
    notes: [],
    revision: 0,
    createdAt: now,
    updatedAt: now,
  };
  return saveProject(project);
}

export function stageOf(p: Project): Stage {
  if (p.buildPath) return "built";
  if (p.tabs.length === 0) return "brief";
  if (!p.tabsApproved) return "tabs";
  if (p.skipLayouts) return "ready";
  const allApproved = p.tabs.every((t) => t.layoutApproved || t.divider || (t.name.trim().toLowerCase().replace(/\s+/g, "") === "readme" && t.blocks.length === 0));
  return allApproved && p.tabs.some((t) => t.blocks.length > 0) ? "ready" : "layouts";
}

// ---------- Tabs ----------

export interface OutlineRow { id?: string; name: string; purpose?: string; message?: string; story?: string; section?: string; divider?: boolean; agendaLine?: string; kicker?: string }

export function makeTab(input: OutlineRow): Tab {
  const t: Tab = { id: newId("tab"), name: input.name, purpose: input.purpose ?? "", message: input.message, story: input.story, section: input.section, approved: false, layoutProposed: false, layoutApproved: false, blocks: [] };
  if (input.divider) t.divider = true;
  if (input.agendaLine?.trim()) t.agendaLine = input.agendaLine.trim();
  if (input.kicker) t.kicker = input.kicker;
  return t;
}

/** Sentence case on a headline: the first letter, nothing else (feedback F11). */
export function capFirst(s: string): string {
  const t = s.trimStart();
  return t ? s.replace(t[0], t[0].toLocaleUpperCase()) : s;
}

const sameName = (a: string, b: string) => a.trim().toLowerCase() === b.trim().toLowerCase();

/** A divider immediately followed by another divider goes: the run keeps the last name (feedback F10). */
function collapseDividers<T extends { name: string; divider?: boolean }>(rows: T[]): T[] {
  return rows.filter((r, i) => !(r.divider && rows[i + 1]?.divider));
}

/**
 * An imported divider slide, or one from a plan written before dividers were rows: it carries a section,
 * nothing else, and its name is that section. A row with a story or a headline is never promoted.
 */
function looksLikeDivider(t: Tab): boolean {
  const section = (t.section ?? "").trim();
  if (!section || t.blocks.length || t.story?.trim() || t.message?.trim()) return false;
  const name = t.name.trim().toLowerCase();
  return !!name && (sameName(name, section) || section.toLowerCase().startsWith(name));
}

/**
 * A `section` named on an incoming row becomes a divider row above it (feedback F4). Only the row as it
 * arrives counts: the `section` stored on a row is derived from position, so re-reading it would spawn a
 * divider every time a slide moved.
 */
export function spliceDividers(rows: { tab: Tab; section?: string }[]): Tab[] {
  const out: Tab[] = [];
  let running = "";
  for (const { tab, section } of rows) {
    if (tab.divider) { running = tab.name.trim(); out.push(tab); continue; }
    const want = (section ?? "").trim();
    if (want && !sameName(want, running)) { out.push(makeTab({ name: want, divider: true })); running = want; }
    out.push(tab);
  }
  return collapseDividers(out);
}

/**
 * Run after every save. Dividers are rows, sections are not fields: adjacent dividers merge, a divider
 * carries nothing but its title, and every row's `section` is re-derived from the nearest divider at or
 * above it. Headlines are sentence case. Nothing else writes `section`.
 */
export function normalizeOutline(p: Project): void {
  for (const t of p.tabs) {
    if (t.message) t.message = capFirst(t.message.trim()) || undefined;
    if (t.divider === undefined && looksLikeDivider(t)) t.divider = true;
  }
  if (p.kind !== "deck") return;

  let section = "";
  for (const t of p.tabs) {
    if (t.divider) {
      section = t.name.trim();
      t.story = undefined;
      t.message = undefined;
      t.kicker = undefined;
      if (t.blocks.length) { const ids = new Set(t.blocks.map((b) => b.id)); p.notes = p.notes.filter((n) => !n.blockId || !ids.has(n.blockId)); t.blocks = []; }
      t.layoutProposed = false;
      t.layoutApproved = false;
    } else t.agendaLine = undefined;   // the agenda line is a divider's own field, like the section title
    t.section = section;
  }
}

/** The headline block Claude is not supposed to send any more: its text becomes the slide headline (F14). */
export function foldHeadlineBlocks(p: Project, tab: Tab): string | undefined {
  const heads = tab.blocks.filter((b) => b.type === "headline");
  if (!heads.length) return undefined;
  tab.blocks = tab.blocks.filter((b) => b.type !== "headline");
  const ids = new Set(heads.map((b) => b.id));
  p.notes = p.notes.filter((n) => !n.blockId || !ids.has(n.blockId));
  const text = heads.map((b) => (b.title || b.content || "").trim()).find(Boolean);
  if (text && !tab.message?.trim()) tab.message = capFirst(text);
  return text || undefined;
}

/** Replace the outline, keeping ids, blocks and notes for rows whose id or name matches. */
export function replaceTabs(p: Project, tabs: OutlineRow[]): void {
  const byId = new Map(p.tabs.map((t) => [t.id, t]));
  const byName = new Map(p.tabs.map((t) => [t.name.trim().toLowerCase(), t]));
  const next: Tab[] = tabs.map((t) => {
    const existing = (t.id && byId.get(t.id)) || byName.get(t.name.trim().toLowerCase());
    if (existing) return { ...existing, name: t.name, purpose: t.purpose || existing.purpose, message: t.message ?? existing.message, story: t.story ?? existing.story, divider: t.divider ?? existing.divider, agendaLine: t.agendaLine?.trim() ?? existing.agendaLine, approved: false };
    return makeTab(t);
  });
  const keep = new Set(next.map((t) => t.id));
  p.notes = p.notes.filter((n) => keep.has(n.tabId));
  p.tabs = spliceDividers(next.map((t, i) => ({ tab: t, section: tabs[i].section })));
  p.tabsApproved = false;
}

export function findTab(p: Project, ref: string): Tab {
  const t = p.tabs.find((t) => t.id === ref) || p.tabs.find((t) => t.name.trim().toLowerCase() === ref.trim().toLowerCase());
  if (!t) throw new Error(`No tab "${ref}". Tabs: ${p.tabs.map((t) => `${t.name} (${t.id})`).join(", ") || "none"}`);
  return t;
}

// ---------- Blocks ----------

export const DEFAULT_SIZE: Record<WorkbookBlockType, { rows: number; cols: number }> = {
  inputs: { rows: 5, cols: 3 },
  output: { rows: 4, cols: 3 },
  table: { rows: 7, cols: 5 },
  calc: { rows: 7, cols: 5 },
  chart: { rows: 12, cols: 6 },
  text: { rows: 3, cols: 6 },
};

/** The slide grid, from shared/types.ts: 32 columns by 18 rows (spec section 2). */
export { DECK_COLS, DECK_ROWS };

export const DECK_SIZE: Record<DeckBlockType, { rows: number; cols: number }> = {
  headline: { rows: 2, cols: 32 },
  kicker: { rows: 1, cols: 16 },
  text: { rows: 8, cols: 14 },
  bullets: { rows: 8, cols: 14 },
  cards: { rows: 7, cols: 32 },
  stat: { rows: 4, cols: 8 },
  chart: { rows: 10, cols: 20 },
  diagram: { rows: 9, cols: 20 },
  table: { rows: 9, cols: 32 },
  image: { rows: 10, cols: 14 },
  quote: { rows: 5, cols: 24 },
  takeaway: { rows: 2, cols: 32 },
  placeholder: { rows: 8, cols: 14 },
  notes: { rows: 3, cols: 32 },
};

export function defaultSize(type: Block["type"], kind: Kind = "workbook"): { rows: number; cols: number } {
  if (kind === "deck") return DECK_SIZE[type as DeckBlockType] ?? { rows: 8, cols: 14 };
  return DEFAULT_SIZE[type as WorkbookBlockType] ?? { rows: 3, cols: 6 };
}

export function makeBlock(input: Partial<Block> & { type: Block["type"]; title: string; questions?: string[] }, kind: Kind = "workbook"): Block {
  const size = defaultSize(input.type, kind);
  const deck = kind === "deck";
  // A table or calc is as wide as its columns (three at the least, capped at 25), mirroring fitCols in the canvas.
  const tabular = (input.type === "table" || input.type === "calc") && (input.columns?.length ?? 0) > 0;
  const cols = input.cols ?? (deck ? size.cols : tabular ? Math.min(25, Math.max(3, input.columns!.length)) : Math.max(input.columns?.length ?? 0, input.type === "inputs" || input.type === "output" ? 3 : size.cols));
  const rows = input.rows ?? (deck ? size.rows : input.type === "inputs" || input.type === "output" ? Math.max(2, (input.rowLabels?.length ?? 0) + 1) : input.type === "table" || input.type === "calc" ? Math.max(3, (input.sample?.length ?? 0) + 2, (input.rowLabels?.length ?? 0) + 2) : size.rows);
  return {
    id: input.id ?? newId("blk"),
    type: input.type,
    title: input.title,
    row: input.row ?? (deck ? 1 : 2),
    col: input.col ?? (deck ? 1 : 2),
    rows: Math.max(1, rows),
    cols: Math.max(1, cols),
    columns: input.columns,
    rowLabels: input.rowLabels,
    sample: input.sample,
    description: input.description,
    content: input.content,
    emphasis: input.emphasis,
    count: input.count,
    source: input.source,
    detected: input.detected,
    chartKind: input.chartKind,
    diagram: input.diagram,
  };
}

// ---------- Placeholders that are really diagrams (research item 4) ----------

/**
 * Relationship words, each mapped to the diagram kind that draws it. flow is last so "process cycle"
 * reads as a cycle and "phased rollout" as a timeline.
 */
const RELATIONSHIPS: [DiagramKind, RegExp][] = [
  ["cycle", /\b(cycle|loop)\b/i],
  ["timeline", /\b(timeline|roadmap|phases?|phased)\b/i],
  ["tree", /\b(hierarchy|tree|org\s*chart)\b/i],
  ["matrix", /\b(2\s*x\s*2|matrix|versus|vs|overlaps?|overlapping|venn)\b/i],
  ["flow", /\b(steps?|flows?|process|funnel|journey|hand-?off)\b|→|->/i],
];

/** A real image, or a visual someone will design and hand over, is a placeholder for good reason. */
const SUPPLIED = /\b(photos?|photograph|picture|logos?|screenshots?|headshot|icons?|illustration|artwork|mock-?up|render|designs?|designed|designer|art direction)\b|\.(png|jpe?g|gif|svg|webp|tiff?|pdf)\b|(^|\s)[~.]*\/\S+/i;

/** Separators a written-out sequence uses: "A → B → C", "A, B and C", "A + B = C", "cost vs speed". */
const SEQUENCE = /\s*(?:→|➜|⇒|->|=>|\+|=|,|;|\/|\bthen\b|\band\b|\bversus\b|\bvs\b\.?)\s*/i;

/** A word that names the relationship or glues the title together, so never a node on its own. */
const NOT_A_NODE = /^(?:steps?|flows?|process|cycle|loop|timeline|roadmap|phases?|phased|hierarchy|tree|org|chart|2x2|matrix|versus|vs|overlaps?|overlapping|venn|funnel|journey|diagram|of|for|the|a|an|and|to|in|on|with|by|from|our|their|this|that)$/i;

/** One node label: drop the brackets a placeholder sentence comes in and the relationship word in front. */
function nodeLabel(raw: string): string {
  let t = raw.replace(/[\[\]{}()"“”'’]/g, " ").replace(/[.,;:!?]+$/, " ").replace(/\s+/g, " ").trim();
  for (let i = 0; i < 3; i++) {
    const cut = t
      .replace(/^(?:the|a|an)\s+/i, "")
      .replace(/^(?:2\s*x\s*2|matrix|venn|overlaps?|overlapping|timeline|roadmap|phases?|cycle|loop|flows?|process|funnel|journey|hand-?off|hierarchy|tree|org\s*chart|steps?|diagram|placeholder|tb[dc])\b\s*(?:of|between|for|from|across|showing|in)?\s*/i, "")
      .trim();
    if (!cut || cut === t) break;
    t = cut;
  }
  return capFirst(t);
}

/** Nodes read off the first field that spells a sequence out, else undefined. The label before a colon goes. */
function sequenceNodes(fields: (string | undefined)[]): DiagramNode[] | undefined {
  for (const field of fields) {
    const text = (field ?? "").trim();
    if (!text) continue;
    const colon = text.indexOf(":");
    for (const src of [colon >= 0 ? text.slice(colon + 1) : "", text]) {
      if (!src.trim()) continue;
      const parts = src.split(SEQUENCE).map(nodeLabel).filter((s) => s && !NOT_A_NODE.test(s) && s.length <= 40 && s.split(/\s+/).length <= 6);
      if (parts.length >= 2) return parts.slice(0, 7).map((label, i) => ({ id: `n${i + 1}`, label }));
    }
  }
  return undefined;
}

/** No sequence in the text: name the nodes after the title so the slots say what they are for. */
function genericNodes(title: string, kind: DiagramKind): DiagramNode[] {
  const words = title.replace(/[^\w\s-]/g, " ").split(/\s+/).filter((w) => w && !NOT_A_NODE.test(w));
  const stem = capFirst(words.slice(0, 2).join(" ")) || (kind === "timeline" ? "Stage" : "Step");
  const count = kind === "matrix" ? 2 : 3;
  return Array.from({ length: count }, (_, i) => ({ id: `n${i + 1}`, label: `${stem} ${i + 1}` }));
}

/** A matrix takes six nodes: the two axes, then four cells. Name the cells after the axes so the slots say what they hold. */
function matrixCells(axes: DiagramNode[]): DiagramNode[] {
  const a = (axes[0]?.label ?? "Axis 1").toLowerCase();
  const b = (axes[1]?.label ?? "Axis 2").toLowerCase();
  return [
    { id: "c1", label: `High ${a}, high ${b}` },
    { id: "c2", label: `High ${a}, low ${b}` },
    { id: "c3", label: `Low ${a}, high ${b}` },
    { id: "c4", label: `Low ${a}, low ${b}` },
  ];
}

function diagramEdges(kind: DiagramKind, nodes: DiagramNode[]): DiagramEdge[] {
  if (kind === "matrix") return [];                                                    // a matrix draws axes and cells, no arrows
  if (kind === "tree") return nodes.slice(1).map((n) => ({ from: nodes[0].id, to: n.id }));
  const chain = nodes.slice(1).map((n, i) => ({ from: nodes[i].id, to: n.id }));
  if (kind === "cycle" && nodes.length > 2) chain.push({ from: nodes[nodes.length - 1].id, to: nodes[0].id });
  return chain;
}

/**
 * Research item 4, "never ship a placeholder": relationships shipped as dashed boxes holding a sentence
 * like "[a + b = c]" although the diagram block draws a Venn, a flow or a timeline. A deck placeholder
 * or image whose words name a relationship becomes a diagram block instead, with nodes read off the
 * text where it spells a sequence out. A real image or a designed visual someone will supply is left
 * alone. Returns the patch to apply, or undefined when the block stays as it is.
 */
export function promotePlaceholder(b: { type: Block["type"]; title?: string; content?: string; description?: string; diagram?: Diagram }): { type: DeckBlockType; diagram: Diagram } | undefined {
  if (b.type !== "placeholder" && b.type !== "image") return undefined;
  const given = b.diagram?.nodes?.length ? b.diagram : undefined;
  const text = [b.title, b.content, b.description].filter(Boolean).join(". ");
  if (!given && (!text.trim() || SUPPLIED.test(text))) return undefined;
  const kind = RELATIONSHIPS.find(([, re]) => re.test(text))?.[0] ?? given?.kind;
  if (!kind) return undefined;
  let nodes = given?.nodes ?? sequenceNodes([b.title, b.content, b.description]) ?? genericNodes(b.title ?? "", kind);
  if (kind === "matrix" && nodes.length === 2) nodes = [...nodes, ...matrixCells(nodes)]; // build_deck wants the two axes and four cells
  const direction = given?.direction ?? (kind === "flow" ? "LR" : kind === "tree" ? "TB" : undefined);
  return {
    type: "diagram",
    diagram: { kind, ...(direction ? { direction } : {}), nodes, edges: given?.edges?.length ? given.edges : diagramEdges(kind, nodes) },
  };
}

/** Place blocks that have no explicit position: stack them top to bottom with a blank row between. */
export function autoPlace(blocks: Block[], kind: Kind = "workbook"): void {
  if (kind === "deck") return autoPlaceDeck(blocks);
  let nextRow = 2;
  for (const b of blocks) {
    if (b.row <= 0 || b.col <= 0 || Number.isNaN(b.row)) {
      b.row = nextRow;
      b.col = 2;
    }
    nextRow = Math.max(nextRow, b.row + b.rows + 1);
  }
}

/**
 * Rows 1 to 4 and row 18 are furniture: kicker, headline, hairline, source, page number (feedback F11).
 * A block never sits there. Blocks proposed without geometry go through the archetype solver, the rest
 * keep what they were given, and overlaps are pushed apart inside the band.
 */
function autoPlaceDeck(blocks: Block[]): void {
  const furniture = blocks.filter((b) => b.type === "headline" || b.type === "kicker");
  const body = blocks.filter((b) => !furniture.includes(b));
  const isPlaced = (b: Block) => b.row > 0 && b.col > 0 && !Number.isNaN(b.row);
  const loose = body.filter((b) => !isPlaced(b));
  if (body.length && loose.length === body.length) {
    const laid = new Map(layoutBlocks(body).map((b) => [b.id, b]));
    for (const b of body) {
      const l = laid.get(b.id);
      if (!l) continue;
      b.row = l.row; b.col = l.col; b.rows = l.rows; b.cols = l.cols;
      if (l.count != null && b.count == null) b.count = l.count;
    }
  } else if (loose.length) {
    // Some blocks are placed by hand: the rest stack under them, full width, inside the band.
    let next = body.filter(isPlaced).reduce((m, b) => Math.max(m, b.row + b.rows), DECK_BODY_TOP - 1) + 1;
    for (const b of loose) {
      b.col = 3; b.cols = Math.min(b.cols, 28); b.row = next; next = b.row + b.rows + 1;
    }
  }
  for (const b of furniture) { b.col = 1; b.row = b.type === "kicker" ? 1 : 2; b.cols = Math.min(b.cols, DECK_COLS); }
  clampToBody(body);
  resolveBandOverlaps(body);
}

/** The band clamp on its own, for a hand edit or a chat edit that moved one block. */
export function fitDeckBlocks(blocks: Block[]): void {
  const body = blocks.filter((b) => b.type !== "headline" && b.type !== "kicker");
  clampToBody(body);
  resolveBandOverlaps(body);
}

export function findBlock(p: Project, blockId: string): { tab: Tab; block: Block } {
  for (const tab of p.tabs) {
    const block = tab.blocks.find((b) => b.id === blockId);
    if (block) return { tab, block };
  }
  throw new Error(`No block with id "${blockId}".`);
}

export function colLetter(col: number): string {
  let s = "";
  let c = col;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

export function addr(row: number, col: number): string {
  return `${colLetter(col)}${row}`;
}

/** Cell addresses a builder will use for this block; mirrors builder/SPEC.md geometry. */
export function blockCells(b: Block): Record<string, unknown> {
  const r = b.row, c = b.col;
  if (b.type === "inputs" || b.type === "output") {
    const labels = b.rowLabels ?? [];
    const valueCells: Record<string, string> = {};
    labels.forEach((l, i) => (valueCells[l] = addr(r + 1 + i, c + 1)));
    return { title: addr(r, c), labels: `${addr(r + 1, c)}:${addr(r + Math.max(1, labels.length), c)}`, values: `${addr(r + 1, c + 1)}:${addr(r + Math.max(1, labels.length), c + 1)}`, valueCells, ref: `{{${b.id}!<label>}}` };
  }
  if (b.type === "table" || b.type === "calc") {
    const n = Math.max(1, b.rows - 2);
    const w = Math.max(1, b.columns?.length ?? b.cols);
    return { title: addr(r, c), header: `${addr(r + 1, c)}:${addr(r + 1, c + w - 1)}`, data: `${addr(r + 2, c)}:${addr(r + 1 + n, c + w - 1)}`, firstDataRow: r + 2, firstCol: colLetter(c), ref: `{{${b.id}!row,col}} or {{${b.id}!col:c}}` };
  }
  return { anchor: addr(r, c) };
}

// ---------- Notes ----------

export function addNote(p: Project, input: { tabId: string; blockId?: string; text: string; createdBy: "user" | "claude"; kind?: Note["kind"] }): Note {
  const kind = input.kind ?? (input.createdBy === "claude" ? "question" : "instruction");
  const note: Note = { id: newId("note"), tabId: input.tabId, blockId: input.blockId, kind, text: input.text, status: "open", createdBy: input.createdBy, createdAt: new Date().toISOString() };
  p.notes.push(note);
  return note;
}

/** Claude's open questions on a block become question threads; duplicates (same text, still open) are skipped. */
export function addQuestions(p: Project, tabId: string, blockId: string, questions: string[] | undefined): void {
  for (const q of questions ?? []) {
    const text = q.trim();
    if (!text) continue;
    if (p.notes.some((n) => n.blockId === blockId && n.kind === "question" && n.status === "open" && n.text === text)) continue;
    addNote(p, { tabId, blockId, text, createdBy: "claude", kind: "question" });
  }
}

/**
 * Old plans predate deck mode, stored questions on blocks, had notes without a kind and put the headline
 * in a block of its own; normalise on load. The headline is the slide's own field now (feedback F14).
 */
export function migrate(p: Project): void {
  if (!p.kind) p.kind = "workbook";
  if (p.briefCard?.emphasis) p.briefCard = cleanBrief(p.briefCard);
  for (const n of p.notes) if (!n.kind) n.kind = n.createdBy === "claude" ? "question" : "instruction";
  for (const t of p.tabs) for (const b of t.blocks as (Block & { questions?: string[] })[]) {
    if (b.questions) { addQuestions(p, t.id, b.id, b.questions); delete b.questions; }
  }
  if (p.kind === "deck") {
    for (const t of p.tabs) foldHeadlineBlocks(p, t);
    // Plans written before dividers were rows: turn each section string into the row it should have been.
    if (!p.tabs.some((t) => t.divider) && p.tabs.some((t) => (t.section ?? "").trim())) {
      p.tabs = spliceDividers(p.tabs.map((t) => ({ tab: t, section: t.section })));
    }
  }
}

// ---------- Readiness ----------

export interface Warning { tabId: string; tabName: string; blockId?: string; blockTitle?: string; kind: string; message: string }

/** The checks block, by title: "Checks", "Model checks", "Sense check". Nothing else marks one. */
export function isChecksBlock(b: Block): boolean {
  return /\bchecks?\b/i.test(b.title);
}

export function readiness(p: Project): Warning[] {
  const out: Warning[] = [];
  const stage = stageOf(p);
  const deck = p.kind === "deck";
  // Once, not per slide: the brief's one takeaway is what every headline has to serve.
  if (p.tabs.length > 0 && !p.briefCard?.takeaway && (deck || p.briefCard)) out.push({ tabId: "", tabName: "", kind: "no-takeaway", message: "The brief has no takeaway yet" });
  for (const t of p.tabs) {
    const cover = t.name.trim().toLowerCase().replace(/\s+/g, "") === "readme";
    if ((stage === "layouts" || stage === "ready") && t.blocks.length === 0 && !p.skipLayouts && !cover && !t.divider) out.push({ tabId: t.id, tabName: t.name, kind: "empty-tab", message: `"${t.name}" has no blocks yet` });
    // The headline is the slide, so a laid-out content slide without one is not ready to build (F12).
    if (deck && !t.divider && t.layoutProposed && !t.message?.trim()) out.push({ tabId: t.id, tabName: t.name, kind: "no-headline", message: "Headline missing" });
    for (const b of t.blocks) {
      if (!deck && (b.type === "table" || b.type === "calc") && !(b.columns?.length)) out.push({ tabId: t.id, tabName: t.name, blockId: b.id, blockTitle: b.title, kind: "no-columns", message: `"${b.title}" has no columns defined` });
      if (!deck && (b.type === "inputs" || b.type === "output") && !(b.rowLabels?.length)) out.push({ tabId: t.id, tabName: t.name, blockId: b.id, blockTitle: b.title, kind: "no-rows", message: `"${b.title}" has no line items defined` });
      // Research item 4: a placeholder ships as a dashed box with a bracketed sentence in it. Never ship one.
      if (deck && b.type === "placeholder") out.push({ tabId: t.id, tabName: t.name, blockId: b.id, blockTitle: b.title, kind: "placeholder", message: "placeholder: make it a diagram, image or drop it" });
      if (b.detected && b.detected.confidence < 0.6) out.push({ tabId: t.id, tabName: t.name, blockId: b.id, blockTitle: b.title, kind: "low-confidence", message: `"${b.title}" detected as ${b.type} at ${Math.round(b.detected.confidence * 100)}% confidence; check the type` });
      if (!deck && b.type === "chart" && !(b.source && t.blocks.some((x) => x.id === b.source))) out.push({ tabId: t.id, tabName: t.name, blockId: b.id, blockTitle: b.title, kind: "chart-source", message: `Chart "${b.title}" has no source table on this tab` });
      for (const o of t.blocks) {
        if (o === b || o.id < b.id) continue;
        if (b.row < o.row + o.rows && o.row < b.row + b.rows && b.col < o.col + o.cols && o.col < b.col + b.cols)
          out.push({ tabId: t.id, tabName: t.name, blockId: b.id, blockTitle: b.title, kind: "overlap", message: `"${b.title}" overlaps "${o.title}"` });
      }
    }
  }
  // W8: a model that calculates needs one Checks block whose rows read OK or name what broke. Recognised by
  // its title, since a block carries no checks flag; propose it at plan time, on the outputs tab.
  if (!deck && p.tabs.some((t) => t.blocks.some((b) => b.type === "calc")) && !p.tabs.some((t) => t.blocks.some(isChecksBlock)))
    out.push({ tabId: "", tabName: "", kind: "no-checks", message: "No Checks block: put one on the outputs tab (totals reconcile, no #REF, every input read)" });
  for (const n of p.notes.filter((n) => n.status === "open")) {
    const t = p.tabs.find((t) => t.id === n.tabId);
    const b = n.blockId ? t?.blocks.find((b) => b.id === n.blockId) : undefined;
    const where = b ? `"${b.title}"` : `tab "${t?.name}"`;
    if (n.kind === "question") out.push({ tabId: n.tabId, tabName: t?.name ?? "?", blockId: n.blockId, blockTitle: b?.title, kind: "question", message: `Claude asks on ${where}: ${n.text}` });
    else out.push({ tabId: n.tabId, tabName: t?.name ?? "?", blockId: n.blockId, blockTitle: b?.title, kind: "open-note", message: `Open instruction on ${where}: ${n.text}` });
  }
  return out;
}

/** One line of the brief card, or "" when nothing is filled in yet. */
export function briefLine(p: Project): string {
  const b = p.briefCard;
  if (!b) return "";
  const parts = [b.audience && `audience: ${b.audience}`, b.goal && `goal: ${b.goal}`, b.takeaway && `takeaway: ${b.takeaway}`].filter(Boolean);
  return parts.length ? `brief: ${parts.join(" | ")}` : "";
}

/** Compact, model-facing summary of the plan (used for tool results and updateModelContext). */
export function summarize(p: Project): string {
  const stage = stageOf(p);
  const deck = p.kind === "deck";
  const lines: string[] = [];
  lines.push(`project: ${p.id} | name: ${p.name} | kind: ${p.kind} | stage: ${stage} | formatting: ${p.formatting} | revision: ${p.revision}${p.skipLayouts ? " | skipLayouts" : ""}${deck ? ` | slide grid: ${DECK_ROWS} rows x ${DECK_COLS} cols` : ""}`);
  const brief = briefLine(p);
  if (brief) lines.push(brief);
  const build = buildLine(p);
  if (build) lines.push(build);
  p.tabs.forEach((t, i) => {
    if (t.divider) { lines.push(`${i + 1}. ${t.name} (${t.id}) [divider: a section slide, no blocks]${t.agendaLine ? ` [agenda line: ${t.agendaLine}]` : ""}`); return; }
    lines.push(`${i + 1}. ${t.name} (${t.id})${t.section ? ` [${t.section}]` : ""}${p.tabsApproved ? " [outline approved]" : ""}${t.layoutApproved ? " [layout approved]" : t.layoutProposed ? " [layout proposed]" : ""}${t.purpose ? `: ${t.purpose}` : ""}`);
    if (t.story) lines.push(`   story: ${t.story}`);
    if (t.message) lines.push(`   ${deck ? "headline" : "message"}: ${t.message}`);
    if (t.kicker) lines.push(`   kicker: ${t.kicker}`);
    for (const b of t.blocks) {
      if (deck) {
        const detail = [b.content && `content: ${b.content}`, b.count && `count: ${b.count}`, b.emphasis && "emphasis", b.type === "chart" && `${b.chartKind ?? "bar"} chart`, b.diagram && `${b.diagram.kind} diagram, ${b.diagram.nodes.length} nodes`].filter(Boolean).join("; ");
        lines.push(`   - [${b.type}] ${b.title} (${b.id}) at row ${b.row}, col ${b.col}, ${b.rows} rows x ${b.cols} cols${detail ? `; ${detail}` : ""}`);
        continue;
      }
      const cells = blockCells(b);
      const detail = b.type === "inputs" || b.type === "output" ? `lines: ${(b.rowLabels ?? []).join(", ") || "(none)"}` : b.type === "table" || b.type === "calc" ? `columns: ${(b.columns ?? []).join(", ") || "(none)"}${b.rowLabels?.length ? `; rows: ${b.rowLabels.join(", ")}` : ""}` : b.type === "chart" ? `${b.chartKind ?? "bar"} chart, source: ${b.source ?? "(none)"}` : "";
      lines.push(`   - [${b.type}] ${b.title} (${b.id}) at ${addr(b.row, b.col)}, ${b.rows}x${b.cols} cells; ${detail}${(cells as any).data ? `; data ${(cells as any).data}` : (cells as any).values ? `; values ${(cells as any).values}` : ""}`);
    }
  });
  const where = (n: Note) => { const t = p.tabs.find((t) => t.id === n.tabId); const b = n.blockId ? t?.blocks.find((b) => b.id === n.blockId) : undefined; return `${t?.name ?? "?"}${b ? ` / ${b.title}` : ""}`; };
  const openInstr = p.notes.filter((n) => n.status === "open" && n.kind === "instruction");
  const openQ = p.notes.filter((n) => n.status === "open" && n.kind === "question");
  const answered = p.notes.filter((n) => n.kind === "question" && n.status !== "open" && n.response);
  if (openInstr.length) {
    lines.push(`open instructions from the user (${openInstr.length}), resolve or decline each with resolve_notes before building:`);
    for (const n of openInstr) lines.push(`   - ${n.id} on ${where(n)}: ${n.text}`);
  }
  if (openQ.length) {
    lines.push(`your open questions (${openQ.length}), waiting on the user; do not build until answered or withdrawn (resolve_notes with status declined withdraws one):`);
    for (const n of openQ) lines.push(`   - ${n.id} on ${where(n)}: ${n.text}`);
  }
  if (answered.length) {
    lines.push(`answers from the user (${answered.length}), honour these in the build:`);
    for (const n of answered) lines.push(`   - ${n.id} on ${where(n)}: Q "${n.text}" A "${n.response}"`);
  }
  const closed = p.notes.filter((n) => n.status !== "open" && n.kind === "instruction");
  if (closed.length) lines.push(`closed instructions: ${closed.length}`);
  const warn = readiness(p).filter((w) => w.kind !== "open-note");
  if (warn.length) {
    lines.push(`readiness warnings (${warn.length}):`);
    for (const w of warn) lines.push(`   - ${w.message}`);
  }
  return lines.join("\n");
}

/**
 * build_deck: the real .pptx from the approved deck plan.
 *
 * Mirrors tools-build.ts for workbooks. The model supplies real content per plan block id;
 * geometry, type and emphasis come from the approved plan, never from the payload. The build
 * refuses while a thread is open or a layout is unapproved. builder/SPEC-deck.md is the
 * builder contract: spec JSON in, .pptx + snapshot (+ PDF and slide PNGs) out.
 */
import fs from "node:fs";
import path from "node:path";
import type { McpServer } from "@modelcontextprotocol/server";
import { registerAppTool } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import type { Block, DeckBlockType, Project, Tab } from "../shared/types.js";
import { DECK_COLS, DECK_ROWS, FIDELITY_HELP } from "../shared/types.js";
import { DECK_SIZE, HOME, OUTPUT_DIR, buildOptions, capFirst, findTab, loadProject, saveProject } from "./store.js";
import { DeckBlockTypeSchema, DiagramSchema, ROOT, fail, logTiming, render, runPython, ui } from "./common.js";
import type { DeckAdjustment, DeckBuildSummary, DeckBuiltSlide, DeckTypeScale } from "./tools-review.js";
import { IMAGES_LINE, MAX_IMAGES, rangeText, slideImages, thumbFor } from "./tools-review.js";
import { builderStyleFields } from "./styles.js";

const DECK_BUILDER = path.join(ROOT, "builder", "build_pptx.py");
const DECK_CHECKER = path.join(ROOT, "builder", "check_deck.py");
/** 960 x 540 pt over the 32 x 18 grid the canvas and the plan both use. */
const CELL = 30;
const SLIDE_W = DECK_COLS * CELL;
const SLIDE_H = DECK_ROWS * CELL;
/** Deck tokens shipped with the builder, when the style's own resolver has none. */
const BUNDLED_TOKENS: { match: RegExp; file: string }[] = [
  { match: /harbor-deck/i, file: path.join(ROOT, "builder", "styles", "harbor_deck.json") },
];

/** The archetype slot a block came from; the builder records it in the snapshot. */
const ROLE_OF: Record<DeckBlockType, string> = {
  headline: "headline", kicker: "kicker", text: "text", bullets: "list", cards: "card",
  stat: "stat", chart: "visual", diagram: "visual", table: "table", image: "visual", quote: "quote",
  takeaway: "statement", placeholder: "visual", notes: "notes",
};

/** What to do about each deck-check finding, in the plan. The fix is never a hand-built file. */
const FIX: Record<string, string> = {
  overlap: "move one of them with update_plan move_block, or make it fewer rows",
  "off-slide": "move it into rows 5 to 17 with update_plan move_block",
  "missing-headline": "write the headline with update_plan set_headline",
  overflow: "shorten this block's content, or give it more rows with update_plan move_block",
  "font-floor": "shorten the copy: the type is already at its smallest for that role",
  empty: "send real content for it, or drop it with update_plan remove_block",
  widow: "re-word the last line so it does not end on a single word",
  "diagram-edge": "fix the edge so both ends name an existing node id",
  "diagram-overlap": "fewer nodes, shorter labels, or give the diagram block more rows or columns with update_plan move_block",
  "diagram-font": "shorten the node labels: they are at the smallest readable size",
};

// ---------- content schema (builder/SPEC-deck.md, "Block") ----------

const RunSchema = z.object({
  text: z.string(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  color: z.string().optional().describe("6 hex digits or a palette key"),
});
const ParagraphSchema = z.union([
  z.string(),
  z.object({ text: z.string().optional(), runs: z.array(RunSchema).optional() }),
]);
/** Bullets nest one level here; the builder itself takes any depth. */
const BulletSchema = z.union([
  z.string(),
  z.object({ text: z.string(), children: z.array(z.union([z.string(), z.object({ text: z.string() })])).optional().describe("Sub-bullets, indented 18 pt") }),
]);
const CardSchema = z.object({ header: z.string().describe("2 to 4 words"), sub: z.string().optional().describe("Optional label above the body"), body: z.string().describe("1 to 3 sentences") });
const StatSchema = z.object({ value: z.union([z.string(), z.number()]).describe("The number as it should read, e.g. 4.1x or 38%"), label: z.string().describe("What it measures, one short line") });

const DeckContentSchema = z.object({
  text: z.string().optional().describe("headline, kicker, quote, takeaway, notes: the line itself. A text block may use this for one paragraph"),
  paragraphs: z.array(ParagraphSchema).optional().describe("text: one entry per paragraph, a string or {text} or {runs: [{text, bold?, italic?, color?}]}"),
  runs: z.array(RunSchema).optional().describe("text: a single paragraph of styled runs"),
  items: z.array(z.union([z.string(), CardSchema, StatSchema, BulletSchema])).optional().describe("bullets: strings or {text, children}. cards: {header, sub?, body}. stat: {value, label}"),
  kind: z.enum(["bar", "column", "line", "pie", "stacked"]).optional().describe("chart: the plot. Default column"),
  categories: z.array(z.string()).optional().describe("chart: the x axis labels"),
  series: z.array(z.object({ name: z.string(), values: z.array(z.number().nullable()) })).optional().describe("chart: one entry per series, values aligned to categories"),
  headers: z.array(z.string()).optional().describe("table: column headers"),
  rows: z.array(z.array(z.union([z.string(), z.number()]))).optional().describe("table: one array per body row; 9 rows or fewer belong on a slide"),
  path: z.string().optional().describe("image: absolute path to a png or jpg"),
  placeholder: z.string().optional().describe("image: a label to draw instead when you have no file"),
  label: z.string().optional().describe("placeholder: what the designed element will hold, e.g. 'Next-steps strip: Decide, Fund, Assign'"),
  attribution: z.string().optional().describe("quote: who said it"),
  diagram: DiagramSchema.optional().describe("diagram: the nodes and edges as the reader sees them"),
});

export const DeckBuildBlockSchema = z.object({
  blockId: z.string().describe("Plan block id from get_plan"),
  type: DeckBlockTypeSchema.optional().describe("Only when skipLayouts is on and the block is not in the plan; otherwise the plan's type wins"),
  title: z.string().optional().describe("Optional label drawn above the content; omit for none (the plan's block title is planning shorthand, not slide copy)"),
  content: DeckContentSchema.describe("The real content for this block's type"),
  col: z.number().int().min(1).max(DECK_COLS).optional().describe("skipLayouts only: 1-based column on the 32-wide grid"),
  row: z.number().int().min(1).max(DECK_ROWS).optional().describe("skipLayouts only: 1-based row on the 18-high grid"),
  cols: z.number().int().min(1).max(DECK_COLS).optional().describe("skipLayouts only: width in cells"),
  rows: z.number().int().min(1).max(DECK_ROWS).optional().describe("skipLayouts only: height in cells"),
});

const DeckSlideSchema = z.object({
  tabId: z.string().describe("Slide id (or exact name) from the plan"),
  kind: z.enum(["title", "cover", "agenda", "divider", "section", "close", "content"]).optional().describe("Default: content, or divider for a plan row with a section and no blocks. title draws the wordmark, headline, subtitle and date; agenda draws its eyebrow and statement"),
  headline: z.string().optional().describe("The one line this slide argues, serving the slide's message and the brief's takeaway. Required on every slide except a divider"),
  kicker: z.string().optional().describe("Eyebrow above the headline, e.g. 'Evidence | Where growth comes from'"),
  subtitle: z.string().optional().describe("Title and divider slides"),
  date: z.string().optional().describe("Title slide, e.g. 'Draft for discussion'"),
  source: z.string().optional().describe("Content slides: drawn bottom left as 'Source: ...'"),
  pageNumber: z.union([z.number().int(), z.literal(false)]).optional().describe("Default: the slide's position in the deck. false drops it"),
  notes: z.string().optional().describe("Speaker notes for this slide"),
  blocks: z.array(DeckBuildBlockSchema.passthrough().partial({ content: true }).extend({ blockId: z.string() })).default([]).describe("One entry per plan block on this slide, by blockId. Title and divider slides have none"),
});

type DeckSlideInput = z.infer<typeof DeckSlideSchema>;
type SlideKind = "title" | "agenda" | "divider" | "content";

/** A divider row is a divider slide (feedback F4); the payload's kind only picks title, agenda or close. */
function slideKindOf(tab: Tab, given: string | undefined): SlideKind {
  if (tab.divider) return "divider";
  const k = (given ?? "").toLowerCase();
  if (k === "title" || k === "cover") return "title";
  if (k === "agenda") return "agenda";
  if (k === "divider" || k === "section" || k === "close") return "divider";
  return "content";
}

/** What a type needs in `content`, and the message when it is not there. */
function contentProblem(type: DeckBlockType, c: z.infer<typeof DeckContentSchema>): string | undefined {
  const items = c.items ?? [];
  const has = (k: keyof typeof c) => c[k] !== undefined && c[k] !== null && String(c[k]).length > 0;
  switch (type) {
    case "headline": case "kicker": case "quote": case "takeaway": case "notes":
      return has("text") ? undefined : `a ${type} block needs content.text`;
    case "text":
      return c.paragraphs?.length || c.runs?.length || has("text") ? undefined : "a text block needs content.paragraphs (or content.text for one paragraph)";
    case "bullets":
      if (!items.length) return "a bullets block needs content.items: strings, or {text, children}";
      return items.every((i) => typeof i === "string" || "text" in i) ? undefined : "every bullets item is a string or {text, children}";
    case "cards":
      if (!items.length) return "a cards block needs content.items: {header, sub?, body}";
      return items.every((i) => typeof i === "object" && "header" in i) ? undefined : "every card is {header, sub?, body}";
    case "stat":
      if (!items.length) return "a stat block needs content.items: {value, label}";
      return items.every((i) => typeof i === "object" && "value" in i) ? undefined : "every stat is {value, label}";
    case "chart":
      if (!c.categories?.length || !c.series?.length) return "a chart block needs content.categories and content.series: [{name, values}]";
      return c.series.every((s) => s.values.length === c.categories!.length) ? undefined : `every series needs ${c.categories.length} values, one per category`;
    case "table":
      if (!c.headers?.length || !c.rows?.length) return "a table block needs content.headers and content.rows";
      return c.rows.every((r) => r.length === c.headers!.length) ? undefined : `every row needs ${c.headers.length} cells, one per header`;
    case "image":
      return has("path") || has("placeholder") ? undefined : "an image block needs content.path, or content.placeholder as an honest label";
    case "placeholder":
      return has("label") ? undefined : "a placeholder block needs content.label, naming what will be designed there";
    case "diagram": {
      const d = c.diagram;
      if (!d?.nodes?.length) return "a diagram block needs content.diagram: {kind, nodes: [{id, label}], edges: [{from, to, label?}]}";
      if (d.nodes.length < 2) return "a diagram needs at least two nodes, and reads best at 3 to 7";
      if (d.nodes.length > 12) return `a diagram takes 12 nodes at most, and reads best at 3 to 7; this one has ${d.nodes.length}`;
      const ids = new Set<string>();
      for (const n of d.nodes) { if (ids.has(n.id)) return `two diagram nodes share the id "${n.id}"`; ids.add(n.id); }
      const bad = (d.edges ?? []).find((e) => !ids.has(e.from) || !ids.has(e.to));
      if (bad) return `diagram edge ${bad.from} to ${bad.to} names a node that is not in the list`;
      if (d.kind === "matrix" && d.nodes.length !== 6) return "a matrix diagram takes exactly six nodes: the two axis labels first, then the four cells";
      return undefined;
    }
    default:
      return undefined;
  }
}

/** skipLayouts: the payload places the block itself, clamped to the slide. */
function payloadGeometry(type: DeckBlockType, raw: { col?: number; row?: number; cols?: number; rows?: number }) {
  const size = DECK_SIZE[type] ?? { rows: 8, cols: 14 };
  const cols = Math.max(1, Math.min(raw.cols ?? size.cols, DECK_COLS));
  const rows = Math.max(1, Math.min(raw.rows ?? size.rows, DECK_ROWS));
  return {
    col: Math.min(Math.max(1, raw.col ?? 1), DECK_COLS - cols + 1),
    row: Math.min(Math.max(1, raw.row ?? 7), DECK_ROWS - rows + 1),
    cols, rows,
  };
}

/** The result text stays inside this many lines, so the fix list is what grows and shrinks. */
const MAX_TEXT_LINES = 12;
/** The deck-wide sizes worth one line, in the order a reader notices them. */
const TYPE_ROLES = ["headline", "body", "bullets", "cardBody", "statValue", "tableCell"];

/** "Type: headline 28, body 12" from the builder's `typeScale`, or nothing when it reported none. */
function typeLine(scale: DeckTypeScale | undefined): string | undefined {
  if (!scale) return undefined;
  const size = (role: string) => (typeof scale[role] === "number" ? (scale[role] as number) : undefined);
  const roles = [...TYPE_ROLES, ...Object.keys(scale)].filter((r, i, all) => size(r) !== undefined && all.indexOf(r) === i);
  if (!roles.length) return undefined;
  const floored = scale.atFloor ?? [];
  return `Type: ${roles.slice(0, 3).map((r) => `${r} ${size(r)}`).join(", ")}${floored.length ? `; ${floored.join(" and ")} at the floor, shorten that copy` : ""}`;
}

/** Deck tokens for a skill the style resolver could not place: my sidecar, then a bundled file. */
function fallbackTokens(name: string): string | undefined {
  const slug = name.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "");
  const mine = path.join(HOME, "styles", `${slug}.json`);
  if (fs.existsSync(mine)) return mine;
  for (const b of BUNDLED_TOKENS) if (b.match.test(name) && fs.existsSync(b.file)) return b.file;
  return undefined;
}

/** The builder's `style` block, plus any warning the fallback had to make. */
function deckStyle(p: Project, warnings: string[]): Record<string, unknown> {
  const fields = builderStyleFields(p.styleId, p.formatting, "deck", p.templatePath);
  const style: Record<string, unknown> = { mode: fields.styleMode };
  if (fields.tokensPath) style.tokensPath = fields.tokensPath;
  if (fields.templatePath) style.templatePath = fields.templatePath;
  if (fields.styleMode === "skill" && !fields.tokensPath) {
    const name = (p.styleId ?? p.formatting).replace(/^skill:/, "");
    const tokens = fallbackTokens(name);
    if (tokens) style.tokensPath = tokens;
    else { style.mode = "house"; warnings.push(`the "${name}" style has no deck tokens; built in the house style (run extract_format on a deck in that format to make them)`); }
  }
  if (fields.styleMode === "template" && !fields.templatePath) {
    style.mode = "house";
    warnings.push("no template path on the plan; built in the house style");
  }
  return style;
}

export function registerBuildDeckTool(server: McpServer): void {
  registerAppTool(server, "build_deck", {
    title: "Build the .pptx",
    description: "Build the real deck from the approved plan: a headline per slide and, by plan block id, the real content for that block's type (text paragraphs, bullet items, cards {header, sub, body}, stats {value, label}, chart {kind, categories, series}, diagram {kind, nodes, edges}, table {headers, rows}, quote, takeaway, image or placeholder label). Diagram rules: one idea per diagram; 3 to 7 nodes; every edge carries a verb; a flow or timeline reads left to right and a tree top to bottom; no decorative shapes; a timeline puts the date in each node label; a matrix takes exactly six nodes, the two axis labels then the four cells; where the style bans shape diagrams the builder draws a labelled placeholder instead. Positions and types come from the plan, not from this call, `fidelity` here overrides the plan's own setting for this one build, and it refuses while a thread is open or a layout is unapproved: when the deck check reports something, fix the plan with update_plan and call build_deck again, never hand-build the deck or replace the file, and look at the slide images this result carries so the next fix answers what a reader would notice.",
    inputSchema: z.object({
      project: z.string(),
      slides: z.array(DeckSlideSchema).min(1).describe("One entry per plan slide, in any order; the deck is built in plan order"),
      fileName: z.string().optional().describe("File name without extension. Default: project name"),
      exportPdf: z.boolean().optional().describe("Also write a PDF beside the deck when LibreOffice is installed"),
      fidelity: z.enum(["strict", "balanced", "loose"]).optional().describe(`This build only, leaving the plan's setting alone. ${FIDELITY_HELP}`),
    }),
    _meta: { ui },
  }, async ({ project, slides, fileName, exportPdf, fidelity }) => {
    const t0 = Date.now();
    const p = loadProject(project);
    if (p.kind !== "deck") return fail(`"${p.name}" (${p.id}) is a ${p.kind} plan, not a deck. Build it with build_workbook, or start a deck with start_project kind deck.`);
    if (!p.tabsApproved) return fail("The outline is not approved. Ask the user to approve the slide list first.");
    if (!p.skipLayouts) {
      const pending = p.tabs.filter((t) => t.layoutProposed && !t.layoutApproved);
      if (pending.length) return fail(`Layouts not yet approved for: ${pending.map((t) => t.name).join(", ")}. Wait for the user to approve on the canvas.`);
      const noLayout = p.tabs.filter((t) => !t.layoutProposed && !t.divider);
      if (noLayout.length) return fail(`No layout proposed for: ${noLayout.map((t) => t.name).join(", ")}. Call propose_layouts for every slide, or set skipLayouts via update_plan.`);
    }
    const open = p.notes.filter((n) => n.status === "open");
    if (open.length) return fail(`${open.length} thread(s) are still open. Instructions: resolve or decline them with resolve_notes. Your questions: wait for the user's answer on the canvas, or withdraw with resolve_notes status declined.\n${open.map((n) => `- ${n.id} [${n.kind}]: ${n.text}`).join("\n")}`);

    const errors: string[] = [];
    const extraWarnings: string[] = [];

    // One payload slide per plan slide, matched by id or name, before anything is built.
    const byTab = new Map<string, DeckSlideInput>();
    for (const s of slides) {
      let tab: Tab | undefined;
      try { tab = findTab(p, s.tabId); } catch { errors.push(`Slide "${s.tabId}" is not in the plan. Slides: ${p.tabs.map((t) => `${t.name} (${t.id})`).join(", ")}`); continue; }
      if (byTab.has(tab.id)) { errors.push(`Slide "${tab.name}" (${tab.id}) is given twice; one entry per slide`); continue; }
      byTab.set(tab.id, s);
    }
    // A divider slide needs nothing from the payload: its title is the row name and the builder draws it.
    for (const t of p.tabs) if (!byTab.has(t.id)) {
      if (t.divider) byTab.set(t.id, { tabId: t.id, kind: "divider", blocks: [] } as DeckSlideInput);
      else errors.push(`Slide "${t.name}" (${t.id}) has no entry in the build`);
    }
    if (errors.length) return fail(`Build refused:\n${errors.map((e) => `- ${e}`).join("\n")}`);

    const specSlides: unknown[] = [];
    /** What each plan slide reads at the top, by tab id: the built view's fallback headline. */
    const titleOf = new Map<string, string>();
    /** The kinds the plan brings itself, so the builder is never asked for a cover the deck already has. */
    const planKinds = new Set<SlideKind>();
    for (const t of p.tabs) {
      const s = byTab.get(t.id)!;
      const kind = slideKindOf(t, s.kind);
      planKinds.add(kind);
      const seen = new Set<string>();
      const specBlocks: Record<string, unknown>[] = [];
      for (const raw of (s.blocks ?? []) as any[]) {
        const plan: Block | undefined = t.blocks.find((b) => b.id === raw.blockId);
        const who = plan ? `block "${plan.title}" (${plan.id})` : raw.blockId ? `block id ${raw.blockId}` : "a block with no blockId";
        if (plan) { if (seen.has(plan.id)) { errors.push(`Slide "${t.name}", ${who}: given twice; one entry per plan block`); continue; } seen.add(plan.id); }
        else if (!p.skipLayouts) { errors.push(`Slide "${t.name}": ${who} is not in the approved plan. Add it with update_plan first, or drop it from the payload`); continue; }
        const type = (plan?.type ?? raw.type) as DeckBlockType | undefined;
        if (!type) { errors.push(`Slide "${t.name}", ${who}: type is required when the block is not in the plan`); continue; }
        if (plan && raw.type && raw.type !== plan.type) { errors.push(`Slide "${t.name}", ${who}: the plan has it as a ${plan.type} block, the payload says ${raw.type}. Send ${plan.type} content, or change the type with update_plan first`); continue; }
        // Per-block parse, so a bad field is named by slide and block rather than by array index.
        const parsed = DeckBuildBlockSchema.safeParse({ ...raw, type });
        if (!parsed.success) {
          for (const issue of parsed.error.issues.slice(0, 5)) errors.push(`Slide "${t.name}", ${who}: ${issue.path.join(".") || "block"}: ${issue.message}`);
          continue;
        }
        const problem = contentProblem(type, parsed.data.content);
        if (problem) { errors.push(`Slide "${t.name}", ${who}: ${problem}`); continue; }
        const geometry = plan ? { col: plan.col, row: plan.row, cols: plan.cols, rows: plan.rows } : payloadGeometry(type, parsed.data);
        specBlocks.push({ id: parsed.data.blockId, type, role: ROLE_OF[type], title: parsed.data.title, emphasis: plan?.emphasis || undefined, ...geometry, content: parsed.data.content });
      }
      for (const b of t.blocks) if (!seen.has(b.id)) errors.push(`Slide "${t.name}": block "${b.title}" (${b.id}) is in the approved plan but missing from the build`);

      // The headline lives on the plan row: the payload may restate it, and the plan's own is the fallback.
      const headline = kind === "divider" ? "" : capFirst(((s.headline ?? "").trim() || (t.message ?? "").trim()));
      if (!headline && kind !== "divider") errors.push(`Slide "${t.name}" (${t.id}): a ${kind} slide needs a headline; write the one line it argues for the reader, or set it on the plan with update_plan set_headline. Only divider slides go without`);
      titleOf.set(t.id, headline || t.name);   // a divider reads as its own row name, which is the section title
      specSlides.push({
        id: t.id, kind, name: t.name, story: t.story, section: t.section, message: t.message,
        headline: headline || undefined, kicker: s.kicker ?? t.kicker, subtitle: s.subtitle, date: s.date, source: s.source,
        pageNumber: s.pageNumber, notes: s.notes, blocks: specBlocks,
        agendaLine: kind === "divider" ? t.agendaLine : undefined,   // the generated agenda prefers this over the first story sentence under the divider
      });
    }
    if (errors.length) return fail(`Build refused:\n${errors.map((e) => `- ${e}`).join("\n")}`);

    const outDir = p.outputDir || OUTPUT_DIR;
    fs.mkdirSync(outDir, { recursive: true });
    const base = (fileName || p.fileName || p.name).replace(/[\\/:*?"<>|]/g, "-").trim();
    let outPath = path.join(outDir, `${base}.pptx`);
    for (let v = 2; fs.existsSync(outPath); v++) outPath = path.join(outDir, `${base} v${v}.pptx`);
    // What the builder may change between the plan and the file. `fidelity` on this call is for this
    // build only; `cover` and `agenda` mean "generate one when the outline has none", so a plan that
    // brings its own title or agenda slide asks for none and keeps its own.
    const opts = buildOptions(p);
    const askedFidelity = fidelity ?? opts.fidelity;
    const spec = {
      name: p.name,
      outputPath: outPath,
      exportPdf: exportPdf ?? p.exportPdf ?? false,
      thumbnails: true,
      style: deckStyle(p, extraWarnings),
      slideSize: { w: SLIDE_W, h: SLIDE_H },
      grid: { cols: DECK_COLS, rows: DECK_ROWS, cell: CELL },
      fidelity: askedFidelity,
      notesFromStory: opts.notesFromStory,
      cover: opts.cover && !planKinds.has("title"),
      agenda: opts.agenda && !planKinds.has("agenda"),
      // The agenda statement is the one thing the audience should remember; the goal is the fallback the builder already uses.
      agendaStatement: p.briefCard?.takeaway || undefined,
      brief: p.briefCard,
      slides: specSlides,
    };
    const buildsDir = path.join(HOME, "builds");
    fs.mkdirSync(buildsDir, { recursive: true });
    const specPath = path.join(buildsDir, `${p.id}-deck-${Date.now()}.json`);
    fs.writeFileSync(specPath, JSON.stringify(spec, null, 2));

    const result = await runPython(DECK_BUILDER, [specPath]);
    if (!result.ok) return fail(`Builder failed:\n${(result.errors ?? [result.stderr ?? "unknown error"]).map((e: string) => `- ${e}`).join("\n")}\nFix the payload or the plan and call build_deck again; do not hand-build the deck instead. Spec saved at ${specPath}.`);

    // Deck check: geometry, headlines, overlap and text fit, read back off the built file.
    const checkPath = specPath.replace(/\.json$/, ".check.json");
    let check: any = await runPython(DECK_CHECKER, [outPath, checkPath, "--spec", specPath]);
    try { if (fs.existsSync(checkPath)) check = JSON.parse(fs.readFileSync(checkPath, "utf8")); } catch { /* keep whatever stdout gave */ }
    const findings: any[] = check?.ok ? check.findings ?? [] : [];
    const checkErrors = findings.filter((f) => f.severity === "error");

    const warnings: string[] = [...(result.warnings ?? []), ...extraWarnings];
    const thumbnails: string[] = Array.isArray(result.thumbnails) ? result.thumbnails : [];
    const snapshotPath = typeof result.snapshotPath === "string" ? result.snapshotPath : undefined;
    let snapshot: any;
    if (snapshotPath) {
      try { snapshot = JSON.parse(fs.readFileSync(snapshotPath, "utf8")); }
      catch (e) { process.stderr.write(`snapshot unreadable at ${snapshotPath}: ${(e as Error).message}\n`); }
    }
    // The deck as the builder built it: its snapshot is the order, because a generated cover or agenda
    // shifts every plan slide down. Plan rows come back by tab id, never by position.
    const snapSlides: any[] = Array.isArray(snapshot?.slides) ? snapshot.slides : [];
    const reportRows: any[] = Array.isArray(result.slides) ? result.slides : [];
    const order = snapSlides.length
      ? snapSlides.map((snap: any, i: number) => ({ index: Number(snap?.index ?? i + 1), snap }))
      : reportRows.map((row: any, i: number) => ({ index: Number(row?.index ?? i + 1), snap: undefined as any }));
    const summarySlides: DeckBuiltSlide[] = order.map(({ index, snap }, i) => {
      const row = reportRows.find((r: any) => Number(r?.index ?? 0) === index) ?? reportRows[i];
      // With no snapshot to read ids from, the plan only lines up when the builder added nothing.
      const tab = snap ? p.tabs.find((t) => t.id === snap.id) : order.length === p.tabs.length ? p.tabs[i] : undefined;
      return {
        index,
        tabId: tab?.id,
        name: snap?.name ?? tab?.name ?? row?.title,
        headline: snap?.headline || (tab ? titleOf.get(tab.id) : undefined) || row?.title,
        shapes: Array.isArray(snap?.shapes) ? snap.shapes.length : row?.shapes ?? 0,
        thumbnail: thumbFor(thumbnails, index),
        generated: snap && !tab ? true : undefined,
      };
    });
    const adjusted: DeckAdjustment[] = Array.isArray(result.adjusted) ? result.adjusted : [];
    const typeScale = result.typeScale && typeof result.typeScale === "object" ? (result.typeScale as DeckTypeScale) : undefined;
    // The canvas's deck built view reads this shape and the snapshot JSON; keep the names.
    const buildSummary: DeckBuildSummary = {
      kind: "deck",
      slides: summarySlides,
      findings: findings.slice(0, 200),
      warnings,
      adjusted: adjusted.slice(0, 60),
      typeScale,
      pdfPath: typeof result.pdfPath === "string" ? result.pdfPath : undefined,
    };
    p.buildPath = outPath;
    p.snapshotPath = snapshotPath;
    p.specPath = specPath;   // the spec of this build, so the next one can be diffed against it
    p.buildSummary = buildSummary;
    p.pdfPath = buildSummary.pdfPath;
    p.thumbnails = thumbnails;
    saveProject(p);

    const byKind = Object.entries(findings.reduce((acc: Record<string, number>, f: any) => ({ ...acc, [f.kind]: (acc[f.kind] ?? 0) + 1 }), {})).map(([k, n]) => `${k} ${n}`).join(", ");
    const closed = p.notes.filter((n) => n.status !== "open");
    const checkLine = check?.ok ? `Deck check: ${check.summary?.errors ?? 0} errors, ${check.summary?.warnings ?? 0} warnings${byKind ? ` (${byKind})` : ""}` : "Deck check did not run.";
    const lines = [
      `Built ${outPath}`,
      `${summarySlides.length} slides, ${summarySlides.reduce((n: number, s) => n + (s.shapes ?? 0), 0)} shapes${thumbnails.length ? `, ${thumbnails.length} slide PNGs` : ""}`,
      // The two counts read as one line, so the fix list keeps the room inside MAX_TEXT_LINES.
      `${checkLine}; builder warnings: ${warnings.length}${warnings.length && !findings.length ? ` - ${warnings.slice(0, 3).join("; ")}` : ""}`,
    ];
    if (adjusted.length) lines.push(`Adjusted ${adjusted.length} block${adjusted.length === 1 ? "" : "s"} for fit (${result.fidelity ?? askedFidelity}); see the Built view`);
    const scaleLine = typeLine(typeScale);
    if (scaleLine) lines.push(scaleLine);
    if (buildSummary.pdfPath) lines.push(`PDF: ${buildSummary.pdfPath}`);
    // Every finding is fixed inside Blueprint, on the plan the user approved. Never by hand.
    const worst = [...checkErrors, ...findings.filter((f) => f.severity !== "error")]
      // A block at the font floor is the same fix as its overflow; say it once.
      .filter((f, _i, all) => !(f.kind === "font-floor" && all.some((o) => o.kind === "overflow" && o.blockId === f.blockId)));
    // The last two lines are always these, and the fix list takes the room that is left.
    const closing = [
      `Notes honoured: ${closed.filter((n) => n.status === "resolved").length} resolved, ${closed.filter((n) => n.status === "declined").length} declined; tell the user the path and how each note was handled.`,
      worst.length ? `Next: update_plan for each fix above, then build_deck again with the same payload.` : `Next: give the user the path; re-open the canvas with show_plan if they want it back.`,
    ];
    if (worst.length) {
      lines.push(`Fix these on the plan and call build_deck again; do not hand-build the deck or replace the file, the approved layout is the user's.`);
      const room = MAX_TEXT_LINES - lines.length - closing.length;
      const show = Math.max(1, worst.length > room ? room - 1 : worst.length);
      // A deck-wide finding (the one-size-per-role type) is about the deck, not about slide 0.
      for (const f of worst.slice(0, show)) lines.push(`- ${f.scope === "deck" || !f.slide ? "the deck" : `slide ${f.slide}`}: ${f.kind}${f.blockId ? ` on ${f.blockId}` : ""}: ${FIX[f.kind] ?? "fix it on the plan"} (${f.message})`);
      if (worst.length > show) lines.push(`- and ${worst.length - show} more, all of those kinds; the whole list is in the build summary`);
    }
    lines.push(...closing);

    // The slides as pictures, so the next fix answers what the deck looks like, not only what the check counted.
    const asked = summarySlides.map((s: any) => s.index as number);
    const images = await slideImages(asked.slice(0, MAX_IMAGES).map((index: number) => ({ index, file: thumbFor(thumbnails, index) })));
    const rest = asked.slice(MAX_IMAGES);
    const tail: string[] = [];
    if (images.blocks.length) tail.push(IMAGES_LINE);
    if (rest.length) tail.push(`${images.blocks.length ? `Images stop at slide ${images.shown[images.shown.length - 1]}` : "No images here"}; use review_deck for slides ${rangeText(rest)}.`);
    tail.push(...images.notes);

    logTiming("build_deck", Date.now() - t0, `${summarySlides.length} slides, ${images.blocks.length} images`);
    const out = render(p, "built", { build: buildSummary, snapshot }, lines.join("\n"));
    if (!tail.length) return out;
    return { ...out, content: [...out.content, { type: "text" as const, text: tail.join("\n") }, ...images.blocks] };
  });
}

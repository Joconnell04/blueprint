import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";
import { z } from "zod";
import type { Project } from "../shared/types.js";
import { HOME, readiness, saveProject, stageOf, summarize } from "./store.js";
import { styleState } from "./styles.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
// dist/server/*.js -> project root is two levels up; src/server/*.ts -> also two levels up.
export const ROOT = path.resolve(__dirname, "..", "..");
export const APP_HTML = path.join(ROOT, "dist", "mcp-app.html");
export const BUILDER = path.join(ROOT, "builder", "build_xlsx.py");
export const SCANNER = path.join(ROOT, "builder", "import_scan.py");
export const CHECKER = path.join(ROOT, "builder", "check_formulas.py");
export const WB_CHECKER = path.join(ROOT, "builder", "check_workbook.py");
// Versioned so a host that caches UI resources by URI picks up a new build (Claude Desktop kept serving the first canvas it ever read).
export const VERSION = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8")).version as string;
export const RESOURCE_URI = `ui://blueprint/canvas-${VERSION}.html`;

export type View = "tabs" | "layouts" | "built";
export const ui = { resourceUri: RESOURCE_URI };
export const appOnly = { resourceUri: RESOURCE_URI, visibility: ["app" as const] };
export const modelOnly = { visibility: ["model" as const] };


export function render(p: Project, view: View, extra: Record<string, unknown> = {}, text?: string) {
  p.renderSeq = (p.renderSeq ?? 0) + 1;
  saveProject(p);
  const warnings = readiness(p);
  // The canvas draws furniture, dividers and cards from the same tokens the builder uses (feedback F5, F16).
  if (extra.style === undefined) extra = { ...extra, style: styleState(p) };
  // A built project re-rendered later (show_plan, apply_template) still needs its preview: the canvas cannot read files.
  if (view === "built" && extra.snapshot === undefined && p.snapshotPath) {
    try { extra = { ...extra, snapshot: JSON.parse(fs.readFileSync(p.snapshotPath, "utf8")) }; }
    catch (e) { process.stderr.write(`snapshot unreadable at ${p.snapshotPath}: ${(e as Error).message}\n`); }
  }
  if (view === "built" && extra.build === undefined && p.buildSummary) extra = { ...extra, build: p.buildSummary };
  return {
    content: [{ type: "text" as const, text: text ?? summarize(p) }],
    structuredContent: { view, stage: stageOf(p), project: p, warnings, renderId: p.renderSeq, ...extra },
  };
}

export function fail(message: string) {
  return { isError: true as const, content: [{ type: "text" as const, text: message }] };
}

export const WorkbookBlockTypeSchema = z.enum(["inputs", "table", "calc", "output", "chart", "text"]);
/** Deck blocks Claude may place. No headline: that is the slide's own field (feedback F14). */
export const DeckBlockTypeSchema = z.enum(["kicker", "text", "bullets", "cards", "stat", "chart", "diagram", "table", "image", "quote", "takeaway", "placeholder", "notes"]);
/** Both vocabularies in one enum; the project's kind says which half applies. */
export const BlockTypeSchema = z.enum(["inputs", "table", "calc", "output", "chart", "text", "kicker", "bullets", "cards", "stat", "diagram", "image", "quote", "takeaway", "placeholder", "notes", "headline"]);

export const BriefSchema = z.object({
  audience: z.string().optional().describe("Who reads this, in their own words: 'partnership leadership team'. Not a process note"),
  goal: z.string().optional().describe("What it has to do for them: the decision they make or the thing they understand"),
  takeaway: z.string().optional().describe("Optional, one sentence: the thing the audience should remember. Never a slide count, a format name or a note about how you are working"),
  emphasis: z.array(z.string()).optional().describe("Deprecated, ignored. Do not send it"),
});

export const DiagramSchema = z.object({
  kind: z.enum(["flow", "cycle", "timeline", "tree", "matrix"]).describe("flow = steps with arrows, cycle = a loop, timeline = dated stops, tree = a hierarchy, matrix = four cells"),
  direction: z.enum(["LR", "TB"]).optional().describe("Reading direction: LR left to right (default), TB top to bottom"),
  nodes: z.array(z.object({ id: z.string().describe("Short unique id such as n1"), label: z.string().describe("2 to 6 words as the reader sees them") })).describe("3 to 7 nodes reads best, 12 at most, ids unique. A matrix takes the two axis labels first, then its four cells"),
  edges: z.array(z.object({ from: z.string(), to: z.string(), label: z.string().optional().describe("The verb on the arrow, 1 to 3 words") })).describe("Each end is a node id"),
});

export const BlockInputSchema = z.object({
  id: z.string().optional().describe("Keep an existing block id to update it; omit for new blocks"),
  type: BlockTypeSchema.describe("Workbook: inputs (label/value the user changes), output (label/value by formula), table, calc, chart, text. Deck: text, bullets, cards, stat, chart, diagram, table, image, quote, takeaway, placeholder, notes. headline is deprecated: the headline is the slide's own field, and a headline block is folded into it. A relationship (steps, flow, cycle, timeline, hierarchy, 2x2, overlap, Venn) is a diagram with nodes and edges; a placeholder is only a real image or a designed visual a human will supply, and one whose words name a relationship is turned into a diagram for you"),
  title: z.string().describe("Block title as the reader sees it, e.g. Key Assumptions, Membership Ramp"),
  row: z.number().int().min(1).optional().describe("Top-left row, 1-based: an Excel row, or 5 to 17 on a slide (rows 1 to 4 and row 18 are furniture). Omit and the whole slide is laid out for you"),
  col: z.number().int().min(1).optional().describe("Top-left column, 1-based: 1 = A in Excel, or 1 to 32 on a slide. Omit to auto-place"),
  rows: z.number().int().min(1).optional().describe("Planned height in cells (title + header + data). Defaults from content"),
  cols: z.number().int().min(1).optional().describe("Planned width in cells. Defaults from columns"),
  columns: z.array(z.string()).optional().describe("table/calc: header labels, first one is the row-label column if rows are labelled"),
  rowLabels: z.array(z.string()).optional().describe("inputs/output: the line items (labels). table/calc: optional row labels"),
  sample: z.array(z.array(z.union([z.string(), z.number()]))).optional().describe("2 to 4 rows of fake placeholder data for the canvas only, matching columns"),
  description: z.string().optional().describe("One line on what this block does and what it feeds"),
  content: z.string().optional().describe("Deck: one line on what goes here, not the content itself (e.g. 'three proof points from the pilot')"),
  emphasis: z.boolean().optional().describe("Deck: this is the block that should pop on the slide"),
  count: z.number().int().min(1).max(8).optional().describe("Deck: how many cards or stats"),
  questions: z.array(z.string()).optional().describe("Questions for the user about this block. Each becomes an open thread on the block that the user answers on the canvas; the build waits until every question is answered or withdrawn"),
  source: z.string().optional().describe("chart: id of the table/calc block on the same tab it plots"),
  chartKind: z.enum(["bar", "line", "pie"]).optional().describe("chart: bar (default), line or pie"),
  diagram: DiagramSchema.optional().describe("diagram: the nodes and edges to draw. One idea per diagram, 3 to 7 nodes, verbs on the arrows"),
});

/** What a cell is, when the colour rule would read it wrong: an input the user edits, a formula, a link, a note. */
export const CellRoleSchema = z.enum(["input", "formula", "link", "note"]);
export const CellSchema = z.union([z.number(), z.string(), z.boolean(), z.null(), z.object({ f: z.string() }).describe("formula, must start with ="), z.object({ v: z.union([z.number(), z.string()]), format: z.string().optional(), role: CellRoleSchema.optional() })]);
export const FormatSchema = z.enum(["currency", "currency2", "percent", "multiple", "count", "number", "text", "year"]);

const ValidationSchema = z.union([
  z.object({ list: z.union([z.array(z.string()), z.string().describe("=range, {{block!col:c}} allowed")]) }),
  z.object({ type: z.enum(["whole", "decimal"]).optional(), min: z.number().optional(), max: z.number().optional() }),
]).describe("Excel data validation: a dropdown list, or numeric bounds");

export const BuildBlockSchema = z.object({
  id: z.string().describe("Plan block id from get_plan (required when the layout was approved)"),
  type: BlockTypeSchema.optional().describe("Only when skipLayouts is on and the block is not in the plan"),
  title: z.string().optional(),
  row: z.number().int().optional(),
  col: z.number().int().optional(),
  cols: z.number().int().optional(),
  note: z.string().optional().describe("Short grey note in the column right of the block (takes one extra column; leave a gap). One line, under 100 characters"),
  rows: z.array(z.object({ label: z.string(), value: CellSchema, format: FormatSchema.optional(), note: z.string().optional(), validation: ValidationSchema.optional(), role: CellRoleSchema.optional().describe("What this value is where the colour rule would read it wrong: a base year the user must be able to change is role \"input\"") })).optional().describe("inputs/output blocks: one entry per line item. A scenario selector is an inputs row with validation {list: [...]} and formulas elsewhere that CHOOSE/INDEX on it"),
  columns: z.array(z.string()).optional().describe("table/calc: header labels"),
  colFormats: z.array(FormatSchema).optional().describe("table/calc: one format per column, for a table whose columns are the metrics"),
  // W3: a label-plus-years table puts four quantities down one column, so a column format cannot say what each is.
  rowFormats: z.array(FormatSchema.nullable()).optional().describe("table/calc: one format per data row, for a table whose rows are the metrics (label plus years). Wins over colFormats on the cells it covers; null leaves that row to colFormats"),
  data: z.array(z.array(CellSchema)).optional().describe("table/calc: data rows. Use {f: '=...'} for formulas. Reference other blocks with {{blockId!Label}}, {{blockId!r,c}}, {{blockId!col:c}}"),
  totalRow: z.array(CellSchema).optional().describe("table/calc: optional bold total row"),
  rowNotes: z.array(z.string().nullable()).optional().describe("table/calc: one short note per data row, written right of the block (adds one column to its extent)"),
  totalNote: z.string().optional(),
  colValidation: z.array(ValidationSchema.nullable()).optional().describe("table/calc: validation per column, applied to every data cell"),
  cellRoles: z.record(z.string(), CellRoleSchema).optional().describe("table/calc: what a cell is where the colour rule reads it wrong, keyed \"row,col\" zero-based in the data area, e.g. {\"0,1\": \"input\"} for a base year the user edits"),
  text: z.string().optional().describe("text blocks"),
  chart: z.object({ kind: z.enum(["bar", "line", "pie"]), sourceBlock: z.string(), categoriesCol: z.number().int(), seriesCols: z.array(z.number().int()), yTitle: z.string().optional() }).optional(),
});

const WORKFLOW_WORKBOOK = `Blueprint, workbook mode. You plan; the canvas renders in this chat and the user edits it by hand.
1. Ask 2 to 5 clarifying questions first: audience, horizon, granularity, scenarios, the one output number, what data they have. Wait.
2. propose_outline: every tab, name + one-line purpose + one-line message. House format opens with a "Read Me" tab and no blocks on it.
3. Wait for approval on the canvas, then get_plan and propose_layouts for every tab in one call.
4. Give the outputs tab a "Checks" block: one row per check, each reading OK or naming what broke (totals reconcile, no #REF, every input read). A table's own totals go in its total row ({{blockId!total:c}}), never in a block below it.
5. Keep the inputs editable: a toggle row (Yes/No, a scenario switch) carries validation {list: [...]} so it stays an input cell, and a base year the user must be able to change carries role: "input".
6. Wait for layout approval. Chat edits go through update_plan; the open canvas refreshes itself.
7. get_plan, close every thread with resolve_notes, push the real labels into the plan, then build_workbook.
8. Number formats: a table whose rows are the metrics (a label plus years) carries rowFormats, one per data row; a table whose columns are the metrics carries colFormats. rowFormats wins where both cover a cell.
9. Look at every sheet image the build returns: say what a reader would notice (clipped labels, wrong formats, widths, empty bands), fix it through the plan, then build_workbook again. Two rebuild cycles at most unless the user asks for more. Never hand-patch the .xlsx: the plan cannot learn it and the next build loses it.
Each tool result says what to do next. skipLayouts: true builds straight from the approved outline.`;

const WORKFLOW_DECK = `Blueprint, deck mode. You plan; the canvas renders in this chat and the user edits it by hand.
1. Ask first, then wait: audience, goal, the one takeaway, how long the deck runs, which sections it needs.
2. Brief: audience = who reads it, goal = what it must do for them, takeaway = one sentence they should remember. Never a slide count, a format name or a note about how you are working.
3. propose_outline: one row per slide, name = 2 to 5 words of purpose (not a headline), story = 2 to 4 sentences on what the slide does for the reader. A section is a row of its own with divider: true, named after the section.
4. Wait for approval on the canvas, then get_plan.
5. Write every slide's headline from its story, in sentence case, one line, a claim with a verb, then read them top to bottom as the storyline and fix any that break it before you lay anything out.
6. propose_layouts: per slide the headline, an optional kicker, and blocks on the 32x18 grid, rows 5 to 17. Never a headline block. Omit geometry and the slide is laid out for you.
7. A relationship (steps, flow, process, cycle, loop, timeline, roadmap, hierarchy, tree, 2x2, versus, overlap, Venn) is a diagram block with nodes and edges, never a placeholder and never a text block; a placeholder is only for a real image or a designed visual a human will supply.
8. Wait for layout approval. Chat edits go through update_plan; the open canvas refreshes itself.
9. get_plan, close every thread with resolve_notes, then build_deck.
10. Look at every slide image the build returns: for each problem say what a reader would notice, fix it with update_plan or new content, and build_deck again. Two rebuild cycles at most unless the user asks for more; never hand-build the deck or replace the file.
Every headline serves the brief's takeaway. Each tool result says what to do next.`;

export function workflow(kind: "workbook" | "deck"): string {
  return kind === "deck" ? WORKFLOW_DECK : WORKFLOW_WORKBOOK;
}

// ---------- server ----------


export function isReadMe(name: string): boolean {
  return name.trim().toLowerCase().replace(/\s+/g, "") === "readme";
}

export function sheetName(name: string): string {
  return name.replace(/[\[\]:*?\/\\]/g, " ").trim().slice(0, 31) || "Sheet";
}


export function runBuilder(specPath: string): Promise<any> { return runPython(BUILDER, [specPath]); }

/** Append one line per slow step to ~/.blueprint/log/planner.log (created on first use). */
export function logTiming(label: string, ms: number, extra = ""): void {
  try {
    const dir = path.join(HOME, "log"); fs.mkdirSync(dir, { recursive: true });
    fs.appendFileSync(path.join(dir, "planner.log"), `${new Date().toISOString()} ${label} ${ms}ms${extra ? ` ${extra}` : ""}\n`);
  } catch { /* logging never fails a call */ }
}

export function runPython(script: string, args: string[]): Promise<any> {
  const t0 = Date.now();
  return new Promise((resolve0) => {
    const resolve = (r: any) => { logTiming(`python ${path.basename(script)}`, Date.now() - t0, r?.ok === false ? "failed" : ""); resolve0(r); };
    const py = process.env.BLUEPRINT_PYTHON || process.env.WORKBOOK_PLANNER_PYTHON || "python3";
    const child = spawn(py, [script, ...args], { stdio: ["ignore", "pipe", "pipe"] });
    let out = "", err = "";
    child.stdout.on("data", (d) => (out += d));
    child.stderr.on("data", (d) => (err += d));
    child.on("error", (e) => resolve({ ok: false, errors: [`Could not run ${py}: ${e.message}. Install Python 3 with openpyxl and python-pptx (pip3 install openpyxl python-pptx).`] }));
    child.on("close", () => {
      try { let r: any; try { r = JSON.parse(out.trim() || "{}"); } catch { r = JSON.parse(out.trim().split("\n").pop() || "{}"); } if (r.ok === undefined) r.ok = false; if (!r.ok && !r.errors?.length) r.errors = [err.trim() || "builder produced no output"]; resolve(r); }
      catch { resolve({ ok: false, errors: [`Builder output was not JSON. stdout: ${out.slice(-800)} stderr: ${err.slice(-800)}`] }); }
    });
  });
}


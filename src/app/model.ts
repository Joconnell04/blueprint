// Shared client-side types and helpers (mirror of src/server/types.ts, kept dependency-free for the bundle)
import type { BlockType, WorkbookBlockType, Block, Kind, Note, Project } from "../shared/types.js";
export type { BlockType, WorkbookBlockType, Block, Brief, Kind, Tab, Note, Project } from "../shared/types.js";
export interface Warning { tabId: string; tabName: string; blockId?: string; blockTitle?: string; kind: string; message: string }
export type View = "tabs" | "layouts" | "built";
export interface State { view: View; stage: string; project: Project; warnings: Warning[]; focusTab?: string; build?: any; snapshot?: any }

export const CELL_W = 64;
export const CELL_H = 22;
export const TYPE_LABEL: Record<WorkbookBlockType, string> = { inputs: "Inputs", table: "Table", calc: "Calc", output: "Output", chart: "Chart", text: "Text" };
/** Label for any block type, including the deck types this view does not draw. */
export function typeLabel(t: BlockType): string { return (TYPE_LABEL as Record<string, string>)[t] ?? t; }

export function colLetter(col: number): string {
  let s = ""; let c = col;
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}
export function addr(row: number, col: number): string { return `${colLetter(col)}${row}`; }

/** The chrome's words for the mode: a workbook has tabs, a deck has slides. Nothing else in the UI branches on kind. */
export interface Vocab {
  one: string; many: string; add: string; addDivider: string; approve: string; approved: string;
  count: (n: number) => string;
  /** The footer line: how many rows, and how many sections the dividers make. */
  footer: (rows: number, sections: number) => string;
}
export function vocab(kind: Kind): Vocab {
  const deck = kind === "deck";
  const one = deck ? "Slide" : "Tab";
  const many = deck ? "Slides" : "Tabs";
  const count = (n: number) => `${n} ${one.toLowerCase()}${n === 1 ? "" : "s"}`;
  return {
    one, many,
    add: deck ? "Add slide" : "Add tab",
    addDivider: "Add divider",
    approve: deck ? "Approve slides" : "Approve tabs",
    approved: `${many} approved`,
    count,
    footer: (rows, sections) => `${count(rows)}${deck && sections ? ` · ${sections} section${sections === 1 ? "" : "s"}` : ""}`,
  };
}

export function isReadMe(name: string): boolean { return name.trim().toLowerCase().replace(/\s+/g, "") === "readme"; }
export function uid(prefix: string): string { return `${prefix}_${Math.random().toString(16).slice(2, 8)}`; }

export function overlaps(a: Block, b: Block): boolean {
  return a.row < b.row + b.rows && b.row < a.row + a.rows && a.col < b.col + b.cols && b.col < a.col + a.cols;
}

/** Plausible placeholder value from a column name. */
export function fakeValue(header: string, i: number): string {
  const h = header.toLowerCase();
  const n = (base: number) => Math.round(base * (1 + 0.37 * ((i * 7) % 5))).toLocaleString();
  if (/\$|revenue|cost|price|arr|spend|sales|ebitda|cash|capex|opex|value|fee|budget|income/.test(h)) return `$${n(1240)}`;
  if (/%|rate|share|margin|growth|mix|penetration|conversion|churn|yield|pct/.test(h)) return `${(12.5 + i * 3.1).toFixed(1)}%`;
  if (/multiple|x$/.test(h)) return `${(2.4 + i * 0.6).toFixed(1)}x`;
  if (/^(year|fy|date|period|month|quarter)$/.test(h.trim())) return `${2026 + i}`;
  if (/^(year|fy|q|quarter|month)\s*\d/.test(h)) return `$${n(820)}`;
  if (/count|users|customers|members|units|#|headcount|volume|number|households|people|clients|accounts/.test(h)) return n(1200);
  if (/segment|name|label|region|channel|product|category|item|line|scenario|source|type|tier/.test(h)) return ["Segment A", "Segment B", "Segment C", "Segment D", "Segment E"][i % 5];
  return n(48);
}

export function summarizePlan(p: Project, stage: string): string {
  const kind = p.kind ?? "workbook";
  const lines: string[] = [`${kind === "deck" ? "Deck" : "Workbook"} plan "${p.name}" (id ${p.id}), stage ${stage}, revision ${p.revision}. The user edited this on the canvas.`];
  const bc = p.briefCard;
  if (bc) {
    const parts = [bc.audience && `audience: ${bc.audience}`, bc.goal && `goal: ${bc.goal}`, bc.takeaway && `takeaway: ${bc.takeaway}`].filter(Boolean);
    if (parts.length) lines.push(`Brief: ${parts.join(" | ")}`);
  }
  let n = 0; // dividers carry no number: the count is slides, not rows
  p.tabs.forEach((t) => {
    if (t.divider) { lines.push(`-- section divider: ${t.name} (${t.id})`); return; }
    n += 1;
    lines.push(`${n}. ${t.name} (${t.id})${t.section ? ` [${t.section}]` : ""}${t.layoutApproved ? " [layout approved]" : t.layoutProposed ? " [layout proposed]" : ""}: ${t.purpose}`);
    if (t.story) lines.push(`   story: ${t.story}`);
    if (t.message) lines.push(`   message: ${t.message}`);
    for (const b of t.blocks) {
      const detail = b.type === "inputs" || b.type === "output" ? `lines: ${(b.rowLabels ?? []).join(", ") || "(none)"}` : b.type === "table" || b.type === "calc" ? `columns: ${(b.columns ?? []).join(", ") || "(none)"}` : b.type === "chart" ? `source: ${b.source ?? "(none)"}` : "";
      lines.push(`   - [${b.type}] ${b.title} (${b.id}) at ${addr(b.row, b.col)}, ${b.rows}x${b.cols}; ${detail}`);
    }
  });
  const where = (n: Note) => { const t = p.tabs.find((t) => t.id === n.tabId); const b = n.blockId ? t?.blocks.find((b) => b.id === n.blockId) : undefined; return `${t?.name}${b ? ` / ${b.title}` : ""}`; };
  const open = p.notes.filter((n) => n.status === "open" && n.kind === "instruction");
  if (open.length) { lines.push(`Open instructions (${open.length}):`); for (const n of open) lines.push(`   - ${n.id} on ${where(n)}: ${n.text}`); }
  const openQ = p.notes.filter((n) => n.status === "open" && n.kind === "question");
  if (openQ.length) lines.push(`Your questions still unanswered: ${openQ.length}.`);
  const answered = p.notes.filter((n) => n.kind === "question" && n.response);
  if (answered.length) { lines.push(`Answered questions (honour these):`); for (const n of answered) lines.push(`   - on ${where(n)}: Q "${n.text}" A "${n.response}"`); }
  lines.push(`${vocab(kind).many} approved: ${p.tabsApproved ? "yes" : "no"}. Call get_plan before proposing or building.`);
  return lines.join("\n");
}

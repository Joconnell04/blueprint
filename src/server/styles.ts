import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";
import type { Kind, Project } from "../shared/types.js";
import { HOME, slugify } from "./store.js";

// Its own copy of the project root: common.ts imports this module for the canvas's style block, so this
// one must not import common.ts back.
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");

/** Which builder a style can dress. A skill with both blocks, or both descriptions, is "both". */
export type StyleKind = "workbook" | "deck" | "both";

/** A formatting source the user can pick for a project. */
export interface StyleOption {
  id: string;            // "plain" | "house" | "skill:<name>" | "template"
  source: "plain" | "house" | "skill" | "template";
  name: string;
  description: string;
  kind: StyleKind;       // workbook styles, deck styles, or both; the popover filters on the project's kind
  skillPath?: string;    // SKILL.md for skills
  tokensPath?: string;   // deck tokens the deck builder can load: a user sidecar, else the bundled one
  builderMode: "plain" | "house" | "template" | "skill"; // what the Python builder applies mechanically
}

const BUILTIN: StyleOption[] = [
  { id: "plain", source: "plain", name: "Plain", description: "Bold titles and headers, thin borders, number formats, no colour", kind: "both", builderMode: "plain" },
  { id: "house", source: "house", name: "House format", description: "Arial 10, blue inputs, dashed label divider, navy title rules, Read Me cover", kind: "both", builderMode: "house" },
];

/** Deck tokens shipped with the builder, keyed by the skill they were derived from. */
const BUNDLED_TOKENS: { match: RegExp; file: string }[] = [
  { match: /harbor-deck|harbor\b/i, file: path.join(ROOT, "builder", "styles", "harbor_deck.json") },
];
/** The neutral default every deck style falls back to, field by field. */
export const HOUSE_DECK_TOKENS = path.join(ROOT, "builder", "styles", "house_deck.json");

/** Derived deck tokens for skills we may not write to (plugin skills are read-only). */
export const SIDECAR_DIR = path.join(HOME, "styles");

export function sidecarPath(name: string): string {
  return path.join(SIDECAR_DIR, `${slugify(name)}.json`);
}

/** Every directory the skill scan reads. BLUEPRINT_SKILLS (or the old name) comes first. */
function skillDirs(): string[] {
  const dirs = new Set<string>();
  for (const d of envSkillDirs()) dirs.add(d);
  dirs.add(path.join(os.homedir(), ".claude", "skills"));
  const sessions = path.join(os.homedir(), "Library", "Application Support", "Claude", "local-agent-mode-sessions", "skills-plugin");
  try {
    for (const org of fs.readdirSync(sessions)) for (const inst of fs.readdirSync(path.join(sessions, org))) {
      const d = path.join(sessions, org, inst, "skills"); if (fs.existsSync(d)) dirs.add(d);
    }
  } catch { /* not installed */ }
  return [...dirs];
}

function envSkillDirs(): string[] {
  const env = process.env.BLUEPRINT_SKILLS || process.env.WORKBOOK_PLANNER_SKILLS;
  return env ? env.split(path.delimiter).filter(Boolean) : [];
}

/** Where a new skill is written: the first configured skills dir, else ~/.claude/skills. */
export function skillsWriteDir(): string {
  return envSkillDirs()[0] || path.join(os.homedir(), ".claude", "skills");
}

function frontMatterText(md: string): string {
  const m = /^---\n([\s\S]*?)\n---/.exec(md);
  return m ? m[1] : "";
}

function frontMatter(md: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const line of frontMatterText(md).split("\n")) { const i = line.indexOf(":"); if (i > 0) out[line.slice(0, i).trim()] = line.slice(i + 1).trim().replace(/^"|"$/g, ""); }
  return out;
}

/** "house format for any spreadsheet", "workbook style guide", "spreadsheet formatting rules": the skill governs how a sheet looks, not what it computes. */
const FORMAT_SKILL = /(spreadsheet|workbook|financial model|xlsx|excel)\s+(format|formatting|style|styling|look)|(format|style|styling)\s+(rules|guide|skill)?\s*for\s+(any\s+)?(spreadsheet|workbook|xlsx|excel|financial model)/i;

/** "deck format", "slide design", "a PowerPoint in the Acme client format": the skill governs how a slide looks. */
const DECK_SKILL = /(deck|decks|slide|slides|presentation|presentations|powerpoint|pptx|potx)[\s-]+(format|formatting|style|styling|look|design|template|theme|palette)|(format|formatting|style|styling|template|theme)\s+(rules|guide|skill|system)?\s*for\s+(any\s+)?(deck|decks|slides?|presentation|powerpoint|pptx)|(in|matching)\s+the\s+[\w\s.()-]{0,40}?(deck|slide|client|house|brand)\s+format/i;

/** Generic file-mechanics skills: they build the file, they do not say what it should look like. */
const MECHANICS = new Set(["xlsx", "pptx", "docx", "pdf"]);

export function describesWorkbookStyle(description: string): boolean { return FORMAT_SKILL.test(description); }
export function describesDeckStyle(description: string): boolean { return DECK_SKILL.test(description); }

/** A SKILL.md on disk, found by directory name or front-matter name. */
export interface FoundSkill { name: string; dir: string; file: string; md: string; writable: boolean }

export function findSkill(name: string): FoundSkill | undefined {
  const want = name.replace(/^skill:/, "").trim().toLowerCase();
  if (!want) return undefined;
  const writeRoot = path.resolve(skillsWriteDir());
  for (const dir of skillDirs()) {
    let names: string[] = [];
    try { names = fs.readdirSync(dir); } catch { continue; }
    for (const n of names) {
      const file = path.join(dir, n, "SKILL.md");
      if (!fs.existsSync(file)) continue;
      let md = ""; try { md = fs.readFileSync(file, "utf8"); } catch { continue; }
      const fm = frontMatter(md);
      const skillName = fm.name || n;
      if (skillName.toLowerCase() !== want && n.toLowerCase() !== want) continue;
      return { name: skillName, dir: path.join(dir, n), file, md, writable: path.resolve(dir) === writeRoot };
    }
  }
  return undefined;
}

/** The deck tokens a skill can be built from: the user's sidecar first, then a bundled one. */
function tokensFor(name: string): string | undefined {
  const mine = sidecarPath(name);
  if (fs.existsSync(mine)) return mine;
  for (const b of BUNDLED_TOKENS) if (b.match.test(name) && fs.existsSync(b.file)) return b.file;
  return undefined;
}

/** Skills whose description or front-matter blocks say they govern how a workbook or a deck looks. */
export function listSkillStyles(): StyleOption[] {
  const seen = new Map<string, StyleOption>();
  for (const dir of skillDirs()) {
    let names: string[] = [];
    try { names = fs.readdirSync(dir); } catch { continue; }
    for (const n of names) {
      const file = path.join(dir, n, "SKILL.md");
      if (!fs.existsSync(file)) continue;
      let md = ""; try { md = fs.readFileSync(file, "utf8"); } catch { continue; }
      const fm = frontMatter(md);
      const fmText = frontMatterText(md);
      const name = fm.name || n;
      const desc = fm.description || "";
      if (MECHANICS.has(name) || seen.has(name)) continue;
      const tokensPath = tokensFor(name);
      const workbook = fm.format === "spreadsheet" || /^format:[ \t]*$/m.test(fmText) || FORMAT_SKILL.test(desc);
      const deck = fm.format === "deck" || /^deck:[ \t]*$/m.test(fmText) || DECK_SKILL.test(desc) || !!tokensPath;
      if (!workbook && !deck) continue;
      seen.set(name, {
        id: `skill:${name}`, source: "skill", name,
        description: desc.length > 140 ? desc.slice(0, 137) + "..." : desc,
        kind: workbook && deck ? "both" : deck ? "deck" : "workbook",
        skillPath: file, tokensPath: deck ? tokensPath : undefined, builderMode: "skill",
      });
    }
  }
  return [...seen.values()];
}

export function listStyles(): StyleOption[] {
  return [...BUILTIN, ...listSkillStyles(), { id: "template", source: "template", name: "Template file", description: "Copy fonts and styles from an .xlsx or .pptx you supply", kind: "both", builderMode: "template" }];
}

export function resolveStyle(id: string | undefined): StyleOption {
  if (!id) return BUILTIN[0];
  return listStyles().find((s) => s.id === id) ?? (id === "house" ? BUILTIN[1] : BUILTIN[0]);
}

/** The skill text Claude should follow at build time, trimmed to a sane size. */
export function skillText(opt: StyleOption, max = 9000): string | undefined {
  if (opt.source !== "skill" || !opt.skillPath) return undefined;
  try { const md = fs.readFileSync(opt.skillPath, "utf8").replace(/^---\n[\s\S]*?\n---\n?/, ""); return md.length > max ? md.slice(0, max) + "\n[truncated]" : md; } catch { return undefined; }
}

/** build_xlsx.py spec fields: the mode plus, for skills, the SKILL.md its format block is read from. */
export interface WorkbookStyleFields { formatting: string; skill?: { path: string } }
/** build_pptx.py reads these as its spec's `style` block: {mode: styleMode, tokensPath, templatePath}. */
export interface DeckStyleFields { styleMode: "plain" | "house" | "skill" | "template"; tokensPath?: string; templatePath?: string }

export function builderStyleFields(styleId: string | undefined, formatting: string, kind?: "workbook"): WorkbookStyleFields;
export function builderStyleFields(styleId: string | undefined, formatting: string, kind: "deck", templatePath?: string): DeckStyleFields;
export function builderStyleFields(styleId: string | undefined, formatting: string, kind: "workbook" | "deck" = "workbook", templatePath?: string): WorkbookStyleFields | DeckStyleFields {
  const opt = resolveStyle(styleId ?? formatting);
  if (kind === "deck") {
    const fields: DeckStyleFields = { styleMode: opt.builderMode };
    if (opt.builderMode === "skill" && opt.tokensPath) fields.tokensPath = opt.tokensPath;
    if (opt.builderMode === "template" && templatePath) fields.templatePath = templatePath;
    return fields;
  }
  if (opt.source === "skill" && opt.skillPath) return { formatting: "skill", skill: { path: opt.skillPath } };
  return { formatting: opt.builderMode };
}

// ---------- resolved tokens for the canvas (feedback F5, F11, F16) ----------

/** The token file a deck project draws its furniture, palette and type scale from. */
export function deckTokensPath(styleId: string | undefined, formatting?: string): string {
  const opt = resolveStyle(styleId ?? formatting);
  if (opt.tokensPath && fs.existsSync(opt.tokensPath)) return opt.tokensPath;
  if (opt.source === "skill") {
    const mine = sidecarPath(opt.name);
    if (fs.existsSync(mine)) return mine;
    for (const b of BUNDLED_TOKENS) if (b.match.test(opt.name) && fs.existsSync(b.file)) return b.file;
  }
  return HOUSE_DECK_TOKENS; // plain, house and template all render off the house tokens
}

/** The style block the canvas draws from: the id it is set to, and for a deck its resolved tokens. */
export interface StyleState { id: string; kind: Kind; tokens?: Record<string, unknown>; tokensPath?: string }

const tokenCache = new Map<string, { mtime: number; tokens: Record<string, unknown> }>();

/** Read a token file at most once per change: cached by path and mtime. */
export function readTokens(file: string): Record<string, unknown> | undefined {
  try {
    const mtime = fs.statSync(file).mtimeMs;
    const hit = tokenCache.get(file);
    if (hit && hit.mtime === mtime) return hit.tokens;
    const tokens = JSON.parse(fs.readFileSync(file, "utf8")) as Record<string, unknown>;
    tokenCache.set(file, { mtime, tokens });
    return tokens;
  } catch (e) {
    process.stderr.write(`style tokens unreadable at ${file}: ${(e as Error).message}\n`);
    return undefined;
  }
}

export function styleState(p: Project): StyleState {
  const id = p.styleId ?? p.formatting;
  if (p.kind !== "deck") return { id, kind: "workbook" };
  const file = deckTokensPath(p.styleId, p.formatting);
  const tokens = readTokens(file) ?? (file === HOUSE_DECK_TOKENS ? undefined : readTokens(HOUSE_DECK_TOKENS));
  return { id, kind: "deck", tokens, tokensPath: file };
}

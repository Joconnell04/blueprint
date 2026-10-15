import fs from "node:fs";
import path from "node:path";
import { spawn } from "node:child_process";
import type { McpServer } from "@modelcontextprotocol/server";
import { registerAppTool } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import { HOME, slugify } from "./store.js";
import { ROOT, fail } from "./common.js";
import { SIDECAR_DIR, describesDeckStyle, describesWorkbookStyle, findSkill, listStyles, sidecarPath, skillsWriteDir } from "./styles.js";
import type { StyleOption } from "./styles.js";

/** builder/extract_format.py: reads example files and writes a formatting skill's tokens. */
export const EXTRACTOR = path.join(ROOT, "builder", "extract_format.py");

const XLSX = /\.(xlsx|xlsm|xltx)$/i;
const PPTX = /\.(pptx|potx|ppsx)$/i;
const IMAGE = /\.(png|jpe?g|gif|webp)$/i;
const MIME: Record<string, string> = { ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif", ".webp": "image/webp" };
/** Claude Desktop publishes no per-image cap; 4 MB is the line at which we send a note instead. */
const MAX_IMAGE_BYTES = 4 * 1024 * 1024;

interface ImageBlock { type: "image"; data: string; mimeType: string; annotations: { audience: ("user" | "assistant")[] } }

interface ExtractResult {
  format: { workbook?: Record<string, unknown>; deck?: Record<string, unknown> };
  confidence: Record<string, number>;
  evidence: Record<string, string>;
  unknowns: string[];
  rules: string[];
  samples: { palette: string[]; fonts: string[] };
}

/** One recovered field, ranked. `key` is the section ("workbook") or a field under it ("workbook.font"). */
interface Field { key: string; value?: unknown; confidence?: number; evidence?: string }
/** One thing the files cannot answer: a row's worth of words, and the whole question behind it. */
interface Ask { question: string; detail: string }

/** extract_format.py prints a SKILL.md on stdout for --skill-md, so this one keeps the bytes raw. */
function runPythonText(script: string, args: string[]): Promise<{ ok: boolean; out: string; err: string }> {
  return new Promise((resolve) => {
    const py = process.env.BLUEPRINT_PYTHON || process.env.WORKBOOK_PLANNER_PYTHON || "python3";
    const child = spawn(py, [script, ...args], { stdio: ["ignore", "pipe", "pipe"] });
    let out = "", err = "";
    child.stdout.on("data", (d) => (out += d));
    child.stderr.on("data", (d) => (err += d));
    child.on("error", (e) => resolve({ ok: false, out: "", err: `Could not run ${py}: ${e.message}. Install Python 3 with openpyxl and python-pptx (pip3 install openpyxl python-pptx).` }));
    child.on("close", (code) => resolve({ ok: code === 0, out, err }));
  });
}

/** The extractor prints one JSON line when it gives up; show its errors, not the whole line. */
function pyError(r: { out: string; err: string }): string {
  try { const j = JSON.parse(r.out.trim().split("\n").pop() || "{}"); if (Array.isArray(j.errors) && j.errors.length) return j.errors.join("; "); } catch { /* not JSON */ }
  return (r.err || r.out || "no output").trim().slice(-600);
}

const ObservationsSchema = z.object({
  format: z.object({
    workbook: z.record(z.string(), z.any()).optional().describe("Workbook keys: font, size, input_color, title_rule_color, label_divider, gridlines, currency, ..."),
    deck: z.record(z.string(), z.any()).optional().describe("Nested deck tokens: {palette: {primary, secondary, ink, muted, bg}, fonts: {major, minor}, typeScale: {headline: {size, bold, color, case}}, card: {shape, radius, border, borderColor, fill}, furniture: {logo: {x, y, w, h}}, backgrounds: {title, divider, content}, slideSize: {width, height}}"),
  }).optional(),
  confidence: z.record(z.string(), z.number()).optional().describe("0 to 1 per dotted path, e.g. {'palette.primary': 0.9}"),
  evidence: z.record(z.string(), z.string()).optional().describe("One line per field on what you saw"),
  rules: z.array(z.string()).optional().describe("What tokens cannot hold, one sentence each; they become the skill's prose"),
  unknowns: z.array(z.string()).optional().describe("What the screenshots still cannot answer, in the words to ask the user"),
}).describe("Your reading of the screenshots in the extractor's schema (docs/skill-format-block.md). A field a file read at 0.8 or better wins; the rest is yours.");

type Observations = z.infer<typeof ObservationsSchema>;
type Kind = "workbook" | "deck" | "both";

function pct(v: number | undefined): string { return v === undefined ? "?" : `${Math.round(v * 100)}%`; }

/** Every scalar leaf of a token block as a dotted path, the shape the confidence table is keyed by. */
function leaves(node: unknown, prefix = ""): { path: string; value: unknown }[] {
  if (node && typeof node === "object" && !Array.isArray(node)) {
    return Object.entries(node as Record<string, unknown>).flatMap(([k, v]) => leaves(v, prefix ? `${prefix}.${k}` : k));
  }
  return prefix ? [{ path: prefix, value: node }] : [];
}

function short(v: unknown, max = 22): string {
  const s = Array.isArray(v) ? v.join(",") : String(v);
  return s.length > max ? `${s.slice(0, max - 1)}...` : s;
}

/** The fields worth a line of a 12-line result; everything else only shows when these are absent. */
const HEADLINE_FIELDS = [
  "font", "size", "input_color", "formula_color", "title_rule_color", "label_divider", "gridlines", "currency",
  "palette.primary", "palette.secondary", "palette.ink", "fonts.major", "typeScale.headline.size", "card.border", "backgrounds.title", "slideSize.width",
];

/** A section's leaves, headline fields first, then by confidence: the one order the tool and the canvas share. */
function ranked(section: string, block: Record<string, unknown>, conf: Record<string, number>): { path: string; value: unknown }[] {
  const rank = (p: string) => { const i = HEADLINE_FIELDS.indexOf(p); return i < 0 ? HEADLINE_FIELDS.length : i; };
  return leaves(block)
    .sort((a, b) => rank(a.path) - rank(b.path) || (conf[`${section}.${b.path}`] ?? 0) - (conf[`${section}.${a.path}`] ?? 0) || a.path.localeCompare(b.path));
}

/** One line per section: the fields that carry the style, with the confidence Claude should weigh them by. */
function fieldLine(section: string, block: Record<string, unknown>, conf: Record<string, number>): string {
  const all = ranked(section, block, conf);
  const head = all.slice(0, 5).map((l) => `${l.path} ${short(l.value)} ${pct(conf[`${section}.${l.path}`])}`);
  return `${section} ${pct(conf[section])} over ${all.length} field(s): ${head.join(", ")}`;
}

/**
 * Every recovered field in that same order, each section headed by its own summary row, so the canvas
 * renders the preview from this list alone instead of re-ranking the token blocks itself.
 */
function fieldsOf(r: ExtractResult): Field[] {
  const conf = r.confidence ?? {};
  const why = r.evidence ?? {};
  const out: Field[] = [];
  for (const section of sectionsOf(r)) {
    out.push({ key: section, confidence: conf[section], evidence: why[section] });
    for (const l of ranked(section, r.format[section as "workbook" | "deck"]!, conf)) {
      const key = `${section}.${l.path}`;
      out.push({ key, value: l.value, confidence: conf[key], evidence: why[key] });
    }
  }
  return out;
}

/** The extractor's unknowns are written for a CLI; Claude reaches the same door through `observations`. */
function ask(unknown: string): string {
  return `Ask the user: ${unknown.replace(/pass the reading back with --merge/g, "pass your reading back as observations")}`;
}

const MAX_QUESTION_WORDS = 12;
/** Words an unknown's own question starts with, when the sentence carries more than one. */
const OPENERS = /^(who|what|which|where|when|why|how|should|does|do|is|are|can|could|would|will|has|have)\b/i;
function words(s: string): string[] { return s.split(/\s+/).filter(Boolean); }

/** An unknown is a paragraph: the question inside it, short enough for one row. */
function question(text: string): string {
  const one = text.replace(/\s+/g, " ").trim();
  const sentences = one.split(/(?<=[.?!])\s+/);
  const sentence = sentences.find((s) => s.includes("?")) ?? sentences[0] ?? one;
  // The sentence whole, then its clauses: "How is a title set off: a rule, a fill?" asks in its first seven words.
  for (const piece of [sentence, ...sentence.split(/\s*[:;]\s*|,\s+(?=(?:and|or|but)\b)/i)]) {
    const p = (piece ?? "").trim().replace(/[,;:]+$/, "");
    if (!p || words(p).length > MAX_QUESTION_WORDS) continue;
    if (!p.endsWith("?") && !OPENERS.test(p)) continue;
    return p.charAt(0).toUpperCase() + (p.endsWith("?") ? p.slice(1) : `${p.slice(1)}?`);
  }
  return `${words(sentence).slice(0, MAX_QUESTION_WORDS - 1).join(" ").replace(/[,;:]+$/, "")}…`;
}

function asks(unknowns: string[] | undefined): Ask[] {
  return (unknowns ?? []).map((u) => ({ question: question(u), detail: u.replace(/\s+/g, " ").trim() }));
}

function sectionsOf(r: ExtractResult): string[] {
  return ["workbook", "deck"].filter((s) => r.format[s as "workbook" | "deck"]);
}

function defaultDescription(name: string, sections: string[]): string {
  const both = sections.includes("workbook") && sections.includes("deck");
  const what = both ? "workbook format and deck format" : sections.includes("deck") ? "deck format" : "workbook format";
  const files = both ? "an .xlsx or .pptx" : sections.includes("deck") ? "a .pptx" : "an .xlsx";
  return `${name}: ${what} derived from example files. Use whenever building ${files} in this style.`;
}

/** The style scan reads the description, so a skill whose description hides its kind would never list. */
function ensureDiscoverable(desc: string, sections: string[]): string {
  const add: string[] = [];
  if (sections.includes("workbook") && !describesWorkbookStyle(desc)) add.push("workbook format");
  if (sections.includes("deck") && !describesDeckStyle(desc)) add.push("deck format");
  if (!add.length) return desc;
  return `${desc.replace(/[\s.]*$/, "")}. It is a ${add.join(" and ")} for Blueprint.`;
}

/** Screenshots go to the model, not the view: audience ["assistant"] says "Claude should look at this". */
function imageBlocks(images: string[]): { blocks: ImageBlock[]; notes: string[] } {
  const blocks: ImageBlock[] = [];
  const notes: string[] = [];
  for (const file of images) {
    try {
      const size = fs.statSync(file).size;
      if (size > MAX_IMAGE_BYTES) { notes.push(`${path.basename(file)} is ${(size / 1048576).toFixed(1)} MB, over the 4 MB inline cap: crop or downscale it, or describe it in observations`); continue; }
      blocks.push({ type: "image", data: fs.readFileSync(file).toString("base64"), mimeType: MIME[path.extname(file).toLowerCase()] ?? "image/png", annotations: { audience: ["assistant"] } });
    } catch (e) { notes.push(`${path.basename(file)} could not be read (${(e as Error).message})`); }
  }
  return { blocks, notes };
}

/** Only the half the caller asked for: a workbook-only run never merges deck observations. */
function obsForKind(obs: Observations, kind: Kind): Observations {
  if (!obs.format) return obs;
  const format: Observations["format"] = {};
  if (kind !== "deck" && obs.format.workbook) format.workbook = obs.format.workbook;
  if (kind !== "workbook" && obs.format.deck) format.deck = obs.format.deck;
  return { ...obs, format };
}

function workDir(): string { const d = path.join(HOME, "extracts"); fs.mkdirSync(d, { recursive: true }); return d; }

function entryFor(id: string): StyleOption | undefined { return listStyles().find((s) => s.id === id); }

export function registerExtractTool(server: McpServer): void {
  registerAppTool(server, "extract_format", {
    title: "Extract a format from example files",
    description: "Derive a formatting skill from examples: .xlsx and .pptx files are parsed into style tokens, screenshots come back as images for you to read. Call it once with no save to see what was recovered and what has to be asked, pass your reading of any screenshots back as observations, then call again with save: true to write ~/.claude/skills/<name>/SKILL.md and, for deck tokens, the builder sidecar. fromSkill derives deck tokens for a prose-only skill instead of from files.",
    inputSchema: z.object({
      paths: z.array(z.string()).default([]).describe("Absolute paths to examples: .xlsx, .pptx, .png, .jpg. Pass paths as the user gave them; never base64"),
      name: z.string().optional().describe("Style name. The skill directory is its slug. Required to save"),
      description: z.string().optional().describe("One line for the skill's front matter: when to use this style. Written for you if omitted"),
      kind: z.enum(["workbook", "deck", "both"]).default("both").describe("Which half to derive; files of the other kind are ignored"),
      observations: ObservationsSchema.optional(),
      fromSkill: z.string().optional().describe("Name of an existing skill (e.g. harbor-deck) to derive tokens for. Returns its prose; send observations back to write its sidecar"),
      save: z.boolean().default(false).describe("false: preview only. true: write the SKILL.md and the deck tokens sidecar"),
      overwrite: z.boolean().default(false).describe("Replace an existing skill directory of the same name"),
    }),
    _meta: { ui: { visibility: ["model", "app"] } },
  }, async ({ paths, name, description, kind, observations, fromSkill, save, overwrite }) => {
    const stamp = Date.now();
    const resultPath = path.join(workDir(), `${stamp}.json`);

    // ---- fromSkill: a prose skill has no tokens; Claude derives them from the prose, we write the sidecar.
    if (fromSkill) {
      const found = findSkill(fromSkill);
      if (!found) {
        const known = listStyles().filter((s) => s.source === "skill").map((s) => s.name);
        return fail(`No skill named ${fromSkill}. Style skills found: ${known.join(", ") || "none"}.`);
      }
      const prose = found.md.replace(/^---\n[\s\S]*?\n---\n?/, "").trim();
      const out = sidecarPath(found.name);
      if (!observations) {
        const next = [
          `${found.name}: prose only, no tokens. ${found.writable ? "Writable" : "Read-only plugin skill"}, so tokens go to ${out}.`,
          `Read the prose above and call extract_format again with fromSkill: "${found.name}", save: true and observations holding format.deck:`,
          `palette (primary, secondary, accent, ink, muted, bg), fonts (major, minor), typeScale per role (size, bold, color, case),`,
          `card (shape, radius, border, borderColor, fill), furniture boxes (logo, kicker, headline, pageNumber, source), backgrounds (title, divider, content), slideSize.`,
          `Put what the prose cannot pin down in unknowns; put authoring rules in rules.`,
        ];
        return {
          content: [
            { type: "text" as const, text: `${found.name} SKILL.md (${found.file}):\n\n${prose.length > 9000 ? `${prose.slice(0, 9000)}\n[truncated]` : prose}` },
            { type: "text" as const, text: next.join("\n") },
          ],
          structuredContent: { skill: found.name, skillPath: found.file, writable: found.writable, sidecarPath: out, needsObservations: true, saved: false },
        };
      }
      const obsPath = path.join(workDir(), `${stamp}-obs.json`);
      fs.writeFileSync(obsPath, JSON.stringify(obsForKind(observations, "deck")));
      const tokensTmp = path.join(workDir(), `${stamp}-tokens.json`);
      const r = await runPythonText(EXTRACTOR, ["--out", resultPath, "--merge", obsPath, "--tokens", tokensTmp]);
      if (!r.ok || !fs.existsSync(resultPath)) return fail(`Extractor failed:\n${pyError(r)}`);
      const result = JSON.parse(fs.readFileSync(resultPath, "utf8")) as ExtractResult;
      if (!result.format.deck) return fail(`Those observations held no deck block. Send observations.format.deck with palette, fonts, typeScale, card, furniture and backgrounds.`);
      if (!fs.existsSync(tokensTmp)) return fail(`The extractor wrote no tokens for ${found.name}. Check observations.format.deck.`);
      const tokens = JSON.parse(fs.readFileSync(tokensTmp, "utf8")) as Record<string, unknown>;
      const groups = Object.keys(tokens).filter((k) => k !== "name");
      const lines: string[] = [];
      if (save) {
        fs.mkdirSync(SIDECAR_DIR, { recursive: true });
        fs.renameSync(tokensTmp, out);
        lines.push(`Wrote deck tokens for ${found.name}: ${out} (${groups.join(", ")}).`);
        lines.push(`${found.file} was not touched; styles.ts merges the sidecar, so skill:${found.name} is buildable now.`);
      } else {
        lines.push(`${found.name} would get ${groups.length} token group(s): ${groups.join(", ")}.`);
        lines.push(`Nothing written. Call again with save: true to write ${out}.`);
      }
      lines.push(fieldLine("deck", result.format.deck, result.confidence));
      for (const u of (result.unknowns ?? []).slice(0, 3)) lines.push(ask(u));
      return {
        content: [{ type: "text" as const, text: lines.slice(0, 12).join("\n") }],
        structuredContent: { skill: found.name, skillPath: found.file, sidecarPath: out, saved: save, tokens, fields: fieldsOf(result), confidence: result.confidence, unknowns: asks(result.unknowns), style: entryFor(`skill:${found.name}`) },
      };
    }

    // ---- files: parse what can be parsed, hand screenshots to Claude, write on the second call.
    const missing = paths.filter((p) => !fs.existsSync(p));
    const present = paths.filter((p) => fs.existsSync(p));
    if (paths.length && !present.length) return fail(`No file at ${paths.join(", ")}`);
    const ignored = present.filter((p) => (kind === "deck" && XLSX.test(p)) || (kind === "workbook" && PPTX.test(p)));
    const use = present.filter((p) => !ignored.includes(p));
    const images = use.filter((p) => IMAGE.test(p));
    if (!use.length && !observations) return fail(`Nothing to read. Give at least one .xlsx, .pptx or screenshot path for kind ${kind}, or pass observations, or use fromSkill.`);

    let obsPath: string | undefined;
    if (observations) {
      obsPath = path.join(workDir(), `${stamp}-obs.json`);
      fs.writeFileSync(obsPath, JSON.stringify(obsForKind(observations, kind)));
    }

    const slug = name ? slugify(name) : "";
    if (save && !slug) return fail("Saving needs a name: it becomes the skill directory and the style's name.");
    const skillDir = path.join(skillsWriteDir(), slug);
    const skillFile = path.join(skillDir, "SKILL.md");
    // A preview reports this too, so the popover can offer Replace instead of learning it from a refused save.
    const exists = !!slug && fs.existsSync(skillDir);
    if (save && exists && !overwrite) return fail(`A skill already exists at ${skillFile}. Call again with overwrite: true to replace it, or pick another name.`);

    const args = ["--out", resultPath, ...(obsPath ? ["--merge", obsPath] : [])];
    const tokensTmp = path.join(workDir(), `${stamp}-tokens.json`);
    if (save) args.push("--tokens", tokensTmp);
    // The description depends on which sections came back, so a save runs the extractor twice: once to see, once to write.
    let result: ExtractResult;
    {
      const r = await runPythonText(EXTRACTOR, [...args, ...use]);
      if (!r.ok || !fs.existsSync(resultPath)) return fail(`Extraction failed:\n${pyError(r)}`);
      result = JSON.parse(fs.readFileSync(resultPath, "utf8")) as ExtractResult;
    }
    const sections = sectionsOf(result);
    const { blocks, notes } = observations ? { blocks: [], notes: [] } : imageBlocks(images);

    const lines: string[] = [];
    const read = use.filter((p) => !IMAGE.test(p)).map((p) => path.basename(p));
    lines.push(`${read.length ? `Read ${read.join(", ")}${sections.length ? `: ${sections.join(" + ")} tokens.` : ": no tokens recovered."}` : "No file to parse."}${images.length ? ` ${images.length} screenshot(s) ${observations ? "already read" : "returned above for you to read"}.` : ""}${ignored.length ? ` Ignored for kind ${kind}: ${ignored.map((p) => path.basename(p)).join(", ")}.` : ""}${missing.length ? ` Missing: ${missing.join(", ")}.` : ""}`);
    for (const s of sections) lines.push(fieldLine(s, result.format[s as "workbook" | "deck"]!, result.confidence));

    if (save) {
      const desc = ensureDiscoverable((description || defaultDescription(name!, sections)).replace(/\s+/g, " ").trim(), sections);
      const md = await runPythonText(EXTRACTOR, [...args, "--skill-md", slug, desc, ...use]);
      if (!md.ok || !md.out.startsWith("---")) return fail(`Could not write the skill:\n${pyError(md)}`);
      const existed = fs.existsSync(skillFile);
      fs.mkdirSync(skillDir, { recursive: true });
      // skill_md() restates the tokens; the rules that are not tokens (a frozen header row, an
      // inputs-only tab, what Claude read off a screenshot) are appended so they reach the prose.
      const extra = /^##+ Rules$/m.test(md.out) ? [] : (result.rules ?? []).map((x) => x.trim()).filter((x) => x && !md.out.includes(x)).slice(0, 12);
      fs.writeFileSync(skillFile, extra.length ? `${md.out.replace(/\s*$/, "\n")}${extra.map((x) => `- ${x}`).join("\n")}\n` : md.out);
      let tokensOut: string | undefined;
      if (kind !== "workbook" && result.format.deck && fs.existsSync(tokensTmp)) {
        fs.mkdirSync(SIDECAR_DIR, { recursive: true });
        tokensOut = sidecarPath(slug);
        fs.renameSync(tokensTmp, tokensOut);
      } else if (fs.existsSync(tokensTmp)) fs.rmSync(tokensTmp, { force: true });
      const entry = entryFor(`skill:${slug}`);
      lines.push(`Wrote ${skillFile}${existed ? " (replaced)" : ""}.`);
      if (tokensOut) lines.push(`Deck tokens for build_deck: ${tokensOut}.`);
      lines.push(entry ? `Lists as skill:${slug} (kind ${entry.kind}); set it with ui_set_style or update_plan.` : `Written, but the style scan did not pick it up: check the description in ${skillFile}.`);
      for (const u of (result.unknowns ?? []).slice(0, 3)) lines.push(ask(u));
      for (const n of notes) lines.push(n);
      return {
        content: [...blocks, { type: "text" as const, text: lines.slice(0, 12).join("\n") }],
        structuredContent: { saved: true, name: slug, kind, sections, exists: true, skillPath: skillFile, tokensPath: tokensOut, style: entry, fields: fieldsOf(result), format: result.format, confidence: result.confidence, unknowns: asks(result.unknowns), rules: result.rules, samples: result.samples },
      };
    }

    if (fs.existsSync(tokensTmp)) fs.rmSync(tokensTmp, { force: true });
    for (const u of (result.unknowns ?? []).slice(0, 4)) lines.push(ask(u));
    for (const n of notes) lines.push(n);
    if (blocks.length) lines.push(`Read the screenshot(s) above, then call again with observations (palette, fonts, type sizes, card and furniture geometry)${slug ? "" : ", a name"} and save: true.`);
    else if (exists) lines.push(`${skillFile} already exists. Call again with save: true and overwrite: true to replace it.`);
    else lines.push(`Nothing written yet. Call again with save: true${slug ? "" : " and a name"} to write the SKILL.md${kind === "workbook" || !result.format.deck ? "" : " and the deck tokens"}.`);
    return {
      content: [...blocks, { type: "text" as const, text: lines.slice(0, 12).join("\n") }],
      structuredContent: { saved: false, name: slug || undefined, kind, sections, needsObservations: blocks.length > 0, exists, skillPath: slug ? skillFile : undefined, fields: fieldsOf(result), format: result.format, confidence: result.confidence, evidence: result.evidence, unknowns: asks(result.unknowns), rules: result.rules, samples: result.samples },
    };
  });
}

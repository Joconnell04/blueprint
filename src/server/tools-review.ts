/**
 * review_deck and review_workbook: eyes on the built file.
 *
 * The deck builder renders one PNG per slide (soffice to PDF, then pdftoppm; builder/build_pptx.py).
 * The workbook renders here, one page per sheet, because build_xlsx.py writes no images. Both turn the
 * files into image content blocks the model looks at, and both fix what they see on the plan, never in
 * the built file.
 */
import { execFile } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import type { McpServer } from "@modelcontextprotocol/server";
import { registerAppTool } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import { HOME, loadProject } from "./store.js";
import { fail, logTiming, modelOnly, sheetName } from "./common.js";

const run = promisify(execFile);

/** One result carries 12 slide images at most; the rest come from another review_deck call. */
export const MAX_IMAGES = 12;
/** A PNG over this is resampled narrower before it rides in a result. */
const MAX_PNG_BYTES = 400 * 1024;
/** And the whole result stays under this much base64. */
const MAX_TOTAL_BASE64 = 3 * 1024 * 1024;
/** Resample widths for an oversized PNG; the builder renders 1280 wide at 96 dpi. */
const WIDTHS = [960, 640];

/** The line that precedes the images, in every result that carries them. */
export const IMAGES_LINE = "Slide images follow: look at each, list what a reader would notice, fix it through the plan, rebuild.";
/** The same line for a workbook, which is read for different faults (research W7). */
export const SHEET_IMAGES_LINE = "Sheet images follow: look at each, list what a reader would notice (clipped labels, wrong formats, widths, empty bands), fix it through the plan, rebuild.";

/** An image content block for the model (core MCP; `audience` decides who it is for). */
export type ImageBlock = {
  type: "image";
  data: string;
  mimeType: string;
  annotations: { audience: ("user" | "assistant")[]; priority: number };
};

/** The build summary build_deck saves on the project, as much of it as a review needs. */
export type DeckFinding = { slide?: number; kind?: string; blockId?: string; role?: string; message?: string; severity?: string; scope?: string };
/** A slide of the built deck. `generated` is one the builder added (a cover, an agenda): no plan row, so no tabId. */
export type DeckBuiltSlide = { index: number; tabId?: string; name?: string; headline?: string; shapes?: number; thumbnail?: string; generated?: boolean };
/** Cells as the fidelity pass reports them (builder/SPEC-deck.md, "Fidelity"). */
export type DeckCells = { col: number; row: number; cols: number; rows: number };
export type DeckAdjustment = { slide?: number; blockId?: string; from?: DeckCells; to?: DeckCells; reason?: string };
/** Deck-wide type: the pt size per role, plus `atFloor`, the roles the builder took to their floor. */
export type DeckTypeScale = { atFloor?: string[] } & { [role: string]: number | string[] | undefined };
export type DeckBuildSummary = {
  kind: "deck";
  slides: DeckBuiltSlide[];
  findings: DeckFinding[];
  warnings: string[];
  /** What the fidelity pass moved, one entry per block per slide. */
  adjusted?: DeckAdjustment[];
  /** The sizes the builder pinned for the whole deck, when it reported them. */
  typeScale?: DeckTypeScale;
  pdfPath?: string;
};

/** The PNG for page `index`: pdftoppm writes slide-1.png, or slide-01.png past nine pages. */
export function thumbFor(thumbnails: string[], index: number, prefix = "slide"): string | undefined {
  return thumbnails.find((f) => new RegExp(`${prefix}-0*${index}\\.png$`).test(f)) ?? thumbnails[index - 1];
}

/** "4", "13 to 22" or "2, 5, 9": slide numbers as a reader says them. */
export function rangeText(nums: number[]): string {
  if (!nums.length) return "";
  const contiguous = nums.every((n, i) => i === 0 || n === nums[i - 1] + 1);
  if (nums.length > 2 && contiguous) return `${nums[0]} to ${nums[nums.length - 1]}`;
  return nums.slice(0, 12).join(", ") + (nums.length > 12 ? ", and more" : "");
}

/** One line, collapsed and cut to length, for a result line that has to stay one line. */
function oneLine(s: string | undefined, max: number): string {
  const flat = (s ?? "").replace(/\s+/g, " ").trim();
  return flat.length > max ? `${flat.slice(0, max - 3)}...` : flat;
}

const kb = (bytes: number) => `${Math.round(bytes / 1024)} KB`;

/**
 * A PNG over the per-image cap, resampled narrower with sips into ~/.blueprint/thumbs.
 * Returns the smaller file, or undefined when nothing fit (then the caller skips the slide).
 */
async function resampled(file: string, mtimeMs: number): Promise<string | undefined> {
  const dir = path.join(HOME, "thumbs");
  const key = crypto.createHash("sha1").update(file).digest("hex").slice(0, 8);
  for (const width of WIDTHS) {
    const out = path.join(dir, `${key}-${Math.round(mtimeMs)}-w${width}.png`);
    try {
      if (!fs.existsSync(out)) {
        fs.mkdirSync(dir, { recursive: true });
        await run("sips", ["--resampleWidth", String(width), file, "--out", out], { timeout: 30_000 });
      }
      if (fs.statSync(out).size <= MAX_PNG_BYTES) return out;
    } catch { return undefined; }   // no sips (or it failed): the slide is skipped with a note
  }
  return undefined;
}

export type SlideImages = { blocks: ImageBlock[]; shown: number[]; notes: string[] };
/** What the notes call a picture, and which tool fetches the rest: a slide of a deck, a sheet of a workbook. */
type Noun = { one: string; tool: string };
const SLIDE: Noun = { one: "Slide", tool: "review_deck" };
const SHEET: Noun = { one: "Sheet", tool: "review_workbook" };

/**
 * The requested pages as image blocks: 12 at most, each PNG under 400 KB, the lot under 3 MB of
 * base64. Anything left out comes back as a note the caller prints beside the images.
 */
export async function slideImages(slides: { index: number; file?: string }[], noun: Noun = SLIDE): Promise<SlideImages> {
  const blocks: ImageBlock[] = [];
  const shown: number[] = [];
  const notes: string[] = [];
  const missing: number[] = [];
  const dropped: number[] = [];
  let total = 0;
  for (const { index, file } of slides.slice(0, MAX_IMAGES)) {
    if (!file || !fs.existsSync(file)) { missing.push(index); continue; }
    let use = file;
    const stat = fs.statSync(file);
    const size = stat.size;
    if (size > MAX_PNG_BYTES) {
      const small = await resampled(file, stat.mtimeMs);
      if (!small) { notes.push(`${noun.one} ${index} image is ${kb(size)} and would not resample; skipped.`); continue; }
      use = small;
    }
    const data = fs.readFileSync(use).toString("base64");
    if (total + data.length > MAX_TOTAL_BASE64) { dropped.push(index); continue; }
    total += data.length;
    blocks.push({ type: "image", data, mimeType: "image/png", annotations: { audience: ["assistant"], priority: 0.6 } });
    shown.push(index);
  }
  if (missing.length) notes.push(`No PNG for ${noun.one.toLowerCase()}${missing.length > 1 ? "s" : ""} ${rangeText(missing)}: LibreOffice or pdftoppm is missing, or the build wrote none.`);
  if (dropped.length) notes.push(`${noun.one}${dropped.length > 1 ? "s" : ""} ${rangeText(dropped)} left out to keep the result small; ask for them with ${noun.tool}.`);
  return { blocks, shown, notes };
}

// ---------- workbook: the sheets as pictures ----------

/** Where LibreOffice and pdftoppm live when they are not on PATH; the env names build_pptx.py already honours. */
const SOFFICE_PATHS = ["/opt/homebrew/bin/soffice", "/usr/local/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"];
const PDFTOPPM_PATHS = ["/opt/homebrew/bin/pdftoppm", "/usr/local/bin/pdftoppm"];
const PDFINFO_PATHS = ["/opt/homebrew/bin/pdfinfo", "/usr/local/bin/pdfinfo"];
/** 60 dpi is a page a reader can judge at about 45 KB. */
const SHEET_DPI = 60;
const SOFFICE_TIMEOUT = 90_000;
const PDFTOPPM_TIMEOUT = 45_000;
/** One page per sheet, so image N is sheet N (LibreOffice 7.2 and later). */
const SINGLE_PAGE = 'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":"true"}}';

function onPath(exe: string): string | undefined {
  for (const dir of (process.env.PATH ?? "").split(path.delimiter)) {
    if (!dir) continue;
    try { const full = path.join(dir, exe); if (fs.existsSync(full)) return full; } catch { /* unreadable dir */ }
  }
  return undefined;
}

/** An env override wins and, when it points nowhere, means "not installed" (build_pptx.py does the same). */
function findTool(envName: string, exe: string, extra: string[]): string | undefined {
  const override = process.env[envName];
  if (override) return fs.existsSync(override) ? override : undefined;
  return onPath(exe) ?? extra.find((f) => fs.existsSync(f));
}

/** One rendered page. `sheet` is set only when the PDF came out one page per sheet. */
export type SheetImage = { index: number; sheet?: string };
export type SheetRender = { files: string[]; images: SheetImage[]; notes: string[] };

/** PNGs for this build live under ~/.blueprint/sheets/<project>/p<first page>. */
export function sheetsDir(projectId: string, from: number): string {
  return path.join(HOME, "sheets", projectId, `p${from}`);
}

async function pdfPages(pdf: string): Promise<number | undefined> {
  const tool = findTool("BLUEPRINT_PDFINFO", "pdfinfo", PDFINFO_PATHS);
  if (!tool) return undefined;
  try {
    const { stdout } = await run(tool, [pdf], { timeout: 20_000 });
    const m = /^Pages:\s+(\d+)/m.exec(stdout);
    return m ? Number(m[1]) : undefined;
  } catch { return undefined; }
}

/**
 * The built workbook as one PNG per sheet: LibreOffice to PDF, then pdftoppm, 12 pages from `from`.
 * A missing tool is a note, never an error; the build itself already succeeded.
 */
export async function renderSheets(xlsxPath: string, projectId: string, sheets: string[], from = 1): Promise<SheetRender> {
  const notes: string[] = [];
  const none = (note?: string): SheetRender => ({ files: [], images: [], notes: note ? [...notes, note] : notes });
  if (process.env.BLUEPRINT_NO_SHEET_IMAGES) return none("Sheet images are switched off (BLUEPRINT_NO_SHEET_IMAGES).");
  const soffice = findTool("BLUEPRINT_SOFFICE", "soffice", SOFFICE_PATHS);
  const toppm = findTool("BLUEPRINT_PDFTOPPM", "pdftoppm", PDFTOPPM_PATHS);
  if (!soffice || !toppm) {
    const missing = [!soffice && "LibreOffice (soffice)", !toppm && "pdftoppm (poppler)"].filter(Boolean).join(" and ");
    return none(`No sheet images: ${missing} not installed, so nothing has looked at the sheets. Open the file yourself, or say so to the user.`);
  }
  const dir = sheetsDir(projectId, from);
  try { fs.rmSync(dir, { recursive: true, force: true }); fs.mkdirSync(dir, { recursive: true }); }
  catch (e) { return none(`No sheet images: ${dir} is not writable (${(e as Error).message}).`); }
  const profile = path.join(os.tmpdir(), "blueprint-lo-profile");
  const pdf = path.join(dir, `${path.basename(xlsxPath).replace(/\.xlsx$/i, "")}.pdf`);
  const convert = (filter: string) =>
    run(soffice, [`-env:UserInstallation=file://${profile}`, "--headless", "--norestore", "--convert-to", filter, "--outdir", dir, xlsxPath], { timeout: SOFFICE_TIMEOUT });
  let perSheet = true;
  try { await convert(SINGLE_PAGE); } catch { /* the retry below covers it */ }
  if (!fs.existsSync(pdf)) {   // an older LibreOffice does not take the filter, so print it as it is set up
    perSheet = false;
    try { await convert("pdf"); } catch { /* the check below reports it */ }
  }
  if (!fs.existsSync(pdf)) return none("No sheet images: LibreOffice wrote no PDF.");
  try {
    await run(toppm, ["-r", String(SHEET_DPI), "-png", "-f", String(from), "-l", String(from + MAX_IMAGES - 1), pdf, path.join(dir, "sheet")], { timeout: PDFTOPPM_TIMEOUT });
  } catch (e) { return none(`No sheet images: pdftoppm failed (${(e as Error).message.split("\n")[0].slice(0, 120)}).`); }
  const files = fs.readdirSync(dir)
    .filter((f) => /^sheet-\d+\.png$/.test(f))
    .map((f) => ({ n: Number(/(\d+)\.png$/.exec(f)![1]), file: path.join(dir, f) }))
    .sort((a, b) => a.n - b.n);
  if (!files.length) return none("No sheet images: pdftoppm wrote no PNG.");
  const pages = await pdfPages(pdf);
  // Pages only stand for sheets when the export fitted each sheet on one page.
  const oneEach = perSheet && (pages !== undefined ? pages === sheets.length : files.length === sheets.length);
  const images = files.map(({ n }) => ({ index: n, sheet: oneEach ? sheets[n - 1] : undefined }));
  if (!oneEach) notes.push(`${pages ?? files.length} printed pages for ${sheets.length} sheets, so an image is a page, not a sheet. Set printArea and fitToWidth so a sheet prints as one page.`);
  const last = from + files.length - 1;
  if (pages !== undefined && pages > last) notes.push(`Images cover ${oneEach ? "sheets" : "pages"} ${from} to ${last} of ${pages}; ask for the rest with review_workbook.`);
  return { files: files.map((f) => f.file), images, notes };
}

/** slideImages, worded for a workbook's sheets. */
export function sheetImages(pages: { index: number; file?: string }[]): Promise<SlideImages> {
  return slideImages(pages, SHEET);
}

/** What build_workbook saves on the project, as much of it as a review needs. */
export type WorkbookSheet = { name: string; cells?: number; formulas?: number; charts?: number };
export type FormulaFinding = { sheet?: string; cell?: string; blockId?: string; code?: string; message?: string; severity?: string };
export type WorkbookBuildSummary = {
  sheets?: WorkbookSheet[];
  warnings?: string[];
  check?: { engine?: string; evaluated?: number; summary?: { errors?: number; warnings?: number }; findings?: FormulaFinding[]; error?: string };
  /** One entry per rendered page; `sheet` is set when a page is a sheet. */
  sheetImages?: SheetImage[];
};

export function registerReviewTool(server: McpServer): void {
  registerAppTool(server, "review_deck", {
    title: "Look at the built slides",
    description: "Read the last build back as pictures: the slide images plus each slide's headline, blocks and open findings. Ask for `slides` to look at a few again, or `what: \"issues\"` for only the slides the deck check flagged. Fix what you see with update_plan and build_deck; never edit the .pptx by hand.",
    inputSchema: z.object({
      project: z.string(),
      slides: z.array(z.number().int().min(1)).max(40).optional().describe("Slide numbers, 1-based as the build numbers them. Default: every slide, 12 images per call"),
      what: z.enum(["all", "issues"]).optional().describe("issues: only the slides the deck check flagged. Default all"),
    }),
    _meta: { ui: modelOnly },
  }, async ({ project, slides, what }) => {
    const t0 = Date.now();
    const p = loadProject(project);
    if (p.kind !== "deck") return fail(`"${p.name}" (${p.id}) is a ${p.kind} plan, not a deck. review_deck reads a built deck's slide images.`);
    const summary = p.buildSummary as DeckBuildSummary | undefined;
    if (!p.buildPath || summary?.kind !== "deck" || !summary.slides?.length) {
      return fail(`No build yet for "${p.name}" (${p.id}). Call build_deck first; review_deck reads the last build's slide images.`);
    }

    const rows = summary.slides;
    const findings = summary.findings ?? [];
    // A deck-wide finding (the sizes the builder pins once) belongs to no slide, so it is said once.
    const deckWide = findings.filter((f) => f.scope === "deck" || !f.slide);
    const deckLine = deckWide.length
      ? `Deck-wide: ${deckWide.slice(0, 4).map((f) => `${f.kind}${f.role ? ` on ${f.role}` : ""} (${oneLine(f.message, 60)})`).join("; ")}`
      : undefined;
    const flagged = (index: number) => findings.filter((f) => f.slide === index);
    const known = new Set(rows.map((r) => r.index));
    const asked = slides?.length ? slides.filter((n) => known.has(n)) : rows.map((r) => r.index);
    const unknown = (slides ?? []).filter((n) => !known.has(n));
    const wanted = what === "issues" ? asked.filter((n) => flagged(n).length) : asked;
    if (!wanted.length) {
      return {
        content: [{ type: "text" as const, text: what === "issues"
          ? deckLine ?? `No findings on the last build of "${p.name}". Call review_deck without what to look at every slide.`
          : `No slide ${rangeText(unknown) || "asked for"} in the last build; it has ${rows.length} slides.` }],
      };
    }

    const take = wanted.slice(0, MAX_IMAGES);
    const rest = wanted.slice(MAX_IMAGES);
    const thumbnails = p.thumbnails ?? [];
    const lines = take.map((index) => {
      const row = rows.find((r) => r.index === index)!;
      const tab = p.tabs.find((t) => t.id === row.tabId);
      const blocks = (tab?.blocks ?? []).map((b) => `${b.type} ${oneLine(b.title, 24)}`.trim());
      const found = flagged(index);
      // A divider reads as its own section name, so its headline is not repeated after the colon.
      const head = row.headline && row.headline !== row.name ? `: ${oneLine(row.headline, 90)}` : "";
      // A cover or agenda the builder added has no plan row, so it has no blocks to list either.
      const parts = [`slide ${index} ${oneLine(row.name, 30)}${row.generated ? " (generated)" : ""}${head}`];
      if (blocks.length) parts.push(`blocks: ${blocks.slice(0, 6).join(", ")}${blocks.length > 6 ? `, and ${blocks.length - 6} more` : ""}`);
      if (found.length) parts.push(`findings: ${found.slice(0, 4).map((f) => `${f.kind}${f.blockId ? ` on ${f.blockId}` : ""} (${oneLine(f.message, 50)})`).join("; ")}`);
      return parts.join(" | ");
    });

    const images = await slideImages(take.map((index) => ({ index, file: rows.find((r) => r.index === index)?.thumbnail ?? thumbFor(thumbnails, index) })));
    const tail: string[] = [];
    if (images.blocks.length) tail.push(IMAGES_LINE);
    if (rest.length) tail.push(`Showing ${take.length} of ${wanted.length} slides; call review_deck again for slides ${rangeText(rest)}.`);
    if (unknown.length) tail.push(`No slide ${rangeText(unknown)} in this deck; it has ${rows.length} slides.`);
    tail.push(...images.notes);
    if (deckLine) tail.push(deckLine);
    logTiming("review_deck", Date.now() - t0, `${images.blocks.length} images`);
    return {
      content: [
        { type: "text" as const, text: lines.join("\n") },
        ...(tail.length ? [{ type: "text" as const, text: tail.join("\n") }] : []),
        ...images.blocks,
      ],
    };
  });

  registerAppTool(server, "review_workbook", {
    title: "Look at the built sheets",
    description: "Read the last build back as pictures: one image per sheet, with that sheet's blocks and formula findings. Name `sheets` to look at a few again, or leave it out for all of them, 12 images per call. Fix what you see with update_plan and build_workbook; never patch the .xlsx by hand.",
    inputSchema: z.object({
      project: z.string(),
      sheets: z.array(z.string()).max(40).optional().describe("Sheet names as the build names them. Default: every sheet"),
    }),
    _meta: { ui: modelOnly },
  }, async ({ project, sheets }) => {
    const t0 = Date.now();
    const p = loadProject(project);
    if (p.kind !== "workbook") return fail(`"${p.name}" (${p.id}) is a ${p.kind} plan, not a workbook. review_deck reads a built deck.`);
    const summary = (p.buildSummary ?? {}) as WorkbookBuildSummary;
    const built = summary.sheets ?? [];
    if (!p.buildPath || !built.length) return fail(`No build yet for "${p.name}" (${p.id}). Call build_workbook first; review_workbook reads the last build's sheet images.`);

    const names = built.map((s) => s.name);
    // A sheet is asked for by its built name, or by the plan tab it came from.
    const wanted: string[] = [];
    const unknown: string[] = [];
    for (const raw of sheets ?? []) {
      const q = raw.trim().toLowerCase();
      const tab = p.tabs.find((t) => t.id === raw.trim() || t.name.trim().toLowerCase() === q);
      const hit = names.find((n) => n.toLowerCase() === q)
        ?? names.find((n) => n.toLowerCase() === sheetName(raw.trim()).toLowerCase())
        ?? (tab ? names.find((n) => n === sheetName(tab.name)) : undefined);
      if (!hit) unknown.push(raw.trim());
      else if (!wanted.includes(hit)) wanted.push(hit);
    }
    if (!sheets?.length) wanted.push(...names);
    if (!wanted.length) return { content: [{ type: "text" as const, text: `No sheet named ${unknown.join(", ")}. This build has: ${names.join(", ")}.` }] };

    const map = summary.sheetImages ?? [];
    const perSheet = map.some((m) => m.sheet);
    const pageOf = (name: string): number | undefined =>
      perSheet ? map.find((m) => m.sheet === name)?.index : map.length ? undefined : names.indexOf(name) + 1;
    const take = wanted.slice(0, MAX_IMAGES);
    const rest = wanted.slice(MAX_IMAGES);

    // The build renders the first 12 pages; a later sheet, or a build that rendered none, renders now.
    let files = p.thumbnails ?? [];
    const asked = take.map(pageOf).filter((n): n is number => !!n);
    const notes: string[] = [];
    if (asked.length && !asked.every((n) => files.some((f) => new RegExp(`sheet-0*${n}\\.png$`).test(f))) && fs.existsSync(p.buildPath)) {
      const again = await renderSheets(p.buildPath, p.id, names, Math.min(...asked));
      if (again.files.length) files = again.files;
      notes.push(...again.notes);
    }

    const findings = summary.check?.findings ?? [];
    const lines = take.map((name) => {
      const row = built.find((s) => s.name === name)!;
      const tab = p.tabs.find((t) => sheetName(t.name) === name);
      const blocks = (tab?.blocks ?? []).map((b) => `${b.type} ${oneLine(b.title, 24)}`.trim());
      const found = findings.filter((f) => f.sheet === name);
      const parts = [`${name}: ${row.cells ?? 0} cells, ${row.formulas ?? 0} formulas${row.charts ? `, ${row.charts} charts` : ""}`];
      if (blocks.length) parts.push(`blocks: ${blocks.slice(0, 6).join(", ")}${blocks.length > 6 ? `, and ${blocks.length - 6} more` : ""}`);
      if (found.length) parts.push(`findings: ${found.slice(0, 4).map((f) => `${f.code} at ${f.cell}${f.blockId ? ` (${f.blockId})` : ""}`).join("; ")}`);
      return parts.join(" | ");
    });

    const shots = perSheet || !map.length
      ? take.map((name) => { const page = pageOf(name); return { index: page ?? 0, file: page ? thumbFor(files, page, "sheet") : undefined }; })
      : files.slice(0, MAX_IMAGES).map((file, i) => ({ index: i + 1, file }));
    const images = await sheetImages(shots);
    const tail: string[] = [];
    if (images.blocks.length) tail.push(SHEET_IMAGES_LINE);
    if (images.blocks.length && map.length && !perSheet) tail.push("An image is a printed page, not a sheet: set printArea and fitToWidth so a sheet prints as one page.");
    if (rest.length) tail.push(`Showing ${take.length} of ${wanted.length} sheets; call review_workbook again for ${rest.slice(0, 6).join(", ")}${rest.length > 6 ? ", and more" : ""}.`);
    if (unknown.length) tail.push(`No sheet named ${unknown.join(", ")}; this build has ${names.join(", ")}.`);
    tail.push(...notes, ...images.notes);
    logTiming("review_workbook", Date.now() - t0, `${images.blocks.length} images`);
    return {
      content: [
        { type: "text" as const, text: lines.join("\n") },
        ...(tail.length ? [{ type: "text" as const, text: tail.join("\n") }] : []),
        ...images.blocks,
      ],
    };
  });
}

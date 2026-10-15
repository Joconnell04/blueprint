import fs from "node:fs";
import path from "node:path";
import type { McpServer } from "@modelcontextprotocol/server";
import { registerAppTool } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import type { Note, Project } from "../shared/types.js";
import { HOME, OUTPUT_DIR, buildOptions, findTab, loadProject, saveProject, stageOf } from "./store.js";
import { BuildBlockSchema, CHECKER, WB_CHECKER, fail, isReadMe, logTiming, render, runBuilder, runPython, sheetName, ui } from "./common.js";
import { SHEET_IMAGES_LINE, renderSheets, sheetImages, thumbFor } from "./tools-review.js";
import { builderStyleFields } from "./styles.js";

/** The result text stays inside one screen; the error list takes the room that is left. */
const MAX_TEXT_LINES = 12;

interface ThreadEntry { by: Note["createdBy"]; text: string; status: Note["status"] }
interface Thread { tab: string; blockId: string; entries: ThreadEntry[] }

/**
 * N9: the closed threads for the builder, grouped by the block whose title cell carries the comment, in the
 * order they were written. A note and its response are two entries of one thread, so the file holds the
 * answer as well as the ask. The builder hangs a comment on a block, so a note left on the tab rather than
 * on a block goes on the first block of that tab; a tab with no blocks has nowhere to put one.
 */
function threadsOf(p: Project): Thread[] {
  const groups = new Map<string, Thread>();
  for (const n of p.notes) {
    if (n.status === "open") continue;
    const tab = p.tabs.find((t) => t.id === n.tabId);
    if (!tab) continue;
    const blockId = n.blockId ?? tab.blocks[0]?.id;
    if (!blockId) continue;
    const key = `${tab.id}:${n.blockId ?? ""}`;
    let g = groups.get(key);
    if (!g) { g = { tab: sheetName(tab.name), blockId, entries: [] }; groups.set(key, g); }
    g.entries.push({ by: n.createdBy, text: n.text, status: n.status });
    const answer = n.response?.trim();
    if (answer) g.entries.push({ by: n.createdBy === "user" ? "claude" : "user", text: answer, status: n.status });
  }
  return [...groups.values()];
}

export function registerBuildTools(server: McpServer): void {
  registerAppTool(server, "build_workbook", {
    title: "Build the .xlsx",
    description: "Build the real workbook from the approved plan. Supply content for every block on every tab, by plan block id: real labels, real numbers, working formulas (use {{blockId!Label}} / {{blockId!r,c}} / {{blockId!col:c}} references from get_plan so addresses resolve, including across sheets), rowFormats on a table whose rows are the metrics and colFormats on a table whose columns are the metrics. It refuses while notes are open or layouts are unapproved, takes every position from the plan rather than from this call, and returns one image per sheet to look at before the file goes to the user.",
    inputSchema: z.object({
      project: z.string(),
      sheets: z.array(z.object({
        tab: z.string().describe("Tab id or name"),
        blocks: z.array(BuildBlockSchema.passthrough().partial({ id: true }).extend({ id: z.string() })).describe("One entry per plan block, by id"),
        freeze: z.string().optional().describe("Top-left cell of the scrolling area, e.g. B3 (freeze panes)"),
        printArea: z.string().optional().describe("'auto' or a range like B2:F40"),
        landscape: z.boolean().optional(),
        fitToWidth: z.boolean().optional(),
        zoom: z.number().int().min(10).max(400).optional(),
      })).min(1),
      readme: z.object({ purpose: z.string(), contents: z.array(z.object({ tab: z.string(), holds: z.string() })).optional(), rules: z.array(z.string()).optional() }).optional().describe("Cover page text, used only when the approved plan has a tab named Read Me with no blocks; the builder writes the cover there"),
      fileName: z.string().optional().describe("File name without extension. Default: project name"),
      fidelity: z.enum(["strict", "balanced"]).optional().describe("This build only, leaving the plan's setting alone. strict keeps the approved widths; balanced lets the builder fit a block to its content and align it to the sheet's block width, within two columns"),
    }),
    _meta: { ui },
  }, async ({ project, sheets, readme, fileName, fidelity }) => {
    const t0 = Date.now();
    const p = loadProject(project);
    const stage = stageOf(p);
    if (!p.tabsApproved) return fail("Tabs are not approved. Ask the user to approve the tab list first.");
    if (!p.skipLayouts) {
      const pending = p.tabs.filter((t) => t.layoutProposed && !t.layoutApproved);
      if (pending.length) return fail(`Layouts not yet approved for: ${pending.map((t) => t.name).join(", ")}. Wait for the user to approve on the canvas.`);
      const noLayout = p.tabs.filter((t) => !t.layoutProposed && !(isReadMe(t.name) && t.blocks.length === 0));
      if (noLayout.length && stage !== "ready") return fail(`No layout proposed for: ${noLayout.map((t) => t.name).join(", ")}. Call propose_layouts, or set skipLayouts via update_plan.`);
    }
    const open = p.notes.filter((n) => n.status === "open");
    if (open.length) return fail(`${open.length} thread(s) are still open. Instructions: resolve or decline them with resolve_notes. Your questions: wait for the user's answer on the canvas, or withdraw with resolve_notes status declined.\n${open.map((n) => `- ${n.id} [${n.kind}]: ${n.text}`).join("\n")}`);
    const errors: string[] = [];
    const specSheets: unknown[] = [];
    const seenTabs = new Set<string>();
    // Validate each block against the schema here, so a bad field is reported by tab and block rather than by array index.
    for (const s of sheets) {
      const tabName = (() => { try { return findTab(p, s.tab).name; } catch { return s.tab; } })();
      s.blocks.forEach((raw: any, i: number) => {
        const parsed = BuildBlockSchema.safeParse(raw);
        if (!parsed.success) {
          const planBlock = p.tabs.flatMap((t) => t.blocks).find((b) => b.id === raw?.id);
          const who = planBlock ? `"${planBlock.title}" (${planBlock.id})` : raw?.id ? `id ${raw.id}` : `block #${i + 1}`;
          for (const issue of parsed.error.issues.slice(0, 5)) errors.push(`Tab "${tabName}", ${who}: ${issue.path.join(".") || "block"}: ${issue.message}`);
        } else s.blocks[i] = parsed.data;
      });
    }
    if (errors.length) return fail(`Build refused (payload errors):\n${errors.map((e) => `- ${e}`).join("\n")}`);
    for (const s of sheets) {
      const tab = findTab(p, s.tab);
      seenTabs.add(tab.id);
      const planIds = new Set(tab.blocks.map((b) => b.id));
      const givenIds = new Set(s.blocks.map((b) => b.id));
      for (const b of tab.blocks) if (!givenIds.has(b.id)) errors.push(`Tab "${tab.name}": block "${b.title}" (${b.id}) is in the approved plan but missing from the build`);
      const blocks = s.blocks.map((bb) => {
        const plan = tab.blocks.find((b) => b.id === bb.id);
        if (!plan && !p.skipLayouts) { errors.push(`Tab "${tab.name}": block id ${bb.id} is not in the approved plan`); return null; }
        const base = plan ?? { type: bb.type!, title: bb.title ?? bb.id, row: bb.row ?? 2, col: bb.col ?? 2, cols: bb.cols ?? 3, rows: 1 };
        if (!plan && !bb.type) errors.push(`Block ${bb.id}: type is required when not in the plan`);
        const { id, type: _t, title: _ti, row: _r, col: _c, cols: _co, ...content } = bb;
        if (plan?.type === "chart" && !bb.chart) errors.push(`Chart block "${plan.title}" (${plan.id}) needs a chart spec`);
        if (plan?.type === "chart" && bb.chart && plan.chartKind && bb.chart.kind !== plan.chartKind) errors.push(`Chart block "${plan.title}" (${plan.id}) was approved as a ${plan.chartKind} chart; the payload says ${bb.chart.kind}`);
        if ((base.type === "inputs" || base.type === "output") && !bb.rows?.length) errors.push(`Block "${base.title}" (${id}) needs rows (label/value pairs)`);
        if ((base.type === "table" || base.type === "calc") && !(bb.columns?.length && bb.data?.length)) errors.push(`Block "${base.title}" (${id}) needs columns and data`);
        // rowFormats is a table shape; an inputs or output row carries its own format instead.
        if (bb.rowFormats?.length && base.type !== "table" && base.type !== "calc") errors.push(`Block "${base.title}" (${id}) is ${base.type}, so it takes format per row, not rowFormats`);
        return { id, type: base.type, title: bb.title ?? base.title, row: base.row, col: base.col, cols: base.cols, ...content };
      }).filter(Boolean);
      specSheets.push({ name: sheetName(tab.name), blocks, freeze: s.freeze, printArea: s.printArea, landscape: s.landscape, fitToWidth: s.fitToWidth, zoom: s.zoom });
      void planIds;
    }
    for (const t of p.tabs) if (!seenTabs.has(t.id)) {
      if (isReadMe(t.name) && t.blocks.length === 0) { specSheets.splice(p.tabs.indexOf(t), 0, { name: sheetName(t.name), blocks: [] }); seenTabs.add(t.id); }
      else errors.push(`Tab "${t.name}" (${t.id}) has no sheet in the build`);
    }
    if (errors.length) return fail(`Build refused:\n${errors.map((e) => `- ${e}`).join("\n")}`);

    const outDir = p.outputDir || OUTPUT_DIR;
    fs.mkdirSync(outDir, { recursive: true });
    const base = (fileName || p.fileName || p.name).replace(/[\\/:*?"<>|]/g, "-").trim();
    let outPath = path.join(outDir, `${base}.xlsx`);
    for (let v = 2; fs.existsSync(outPath); v++) outPath = path.join(outDir, `${base} v${v}.xlsx`);
    const hasReadMe = p.tabs.some((t) => isReadMe(t.name));
    // What the builder may change between the approved plan and the file. `fidelity` on this call is for
    // this build only; `comments` says whether the closed threads are written into the file as cell comments.
    const opts = buildOptions(p);
    const askedFidelity = fidelity ?? opts.fidelity;
    const threads = threadsOf(p);
    const spec = { outputPath: outPath, ...builderStyleFields(p.styleId, p.formatting), templatePath: p.templatePath, project: { name: p.name, brief: p.brief }, fidelity: askedFidelity, comments: opts.comments, threads: threads.length ? threads : undefined, readme: hasReadMe ? (readme ?? { purpose: p.brief, contents: p.tabs.filter((t) => !isReadMe(t.name)).map((t) => ({ tab: sheetName(t.name), holds: t.purpose })) }) : undefined, sheets: specSheets };
    const buildsDir = path.join(HOME, "builds");
    fs.mkdirSync(buildsDir, { recursive: true });
    const specPath = path.join(buildsDir, `${p.id}-${Date.now()}.json`);
    fs.writeFileSync(specPath, JSON.stringify(spec, null, 2));

    const result = await runBuilder(specPath);
    if (!result.ok) return fail(`Builder failed:\n${(result.errors ?? [result.stderr ?? "unknown error"]).map((e: string) => `- ${e}`).join("\n")}\nFix the payload and call build_workbook again. Spec saved at ${specPath}.`);
    // Formula check: static references plus evaluation (LibreOffice when installed, else the bundled engine).
    const checkPath = specPath.replace(/\.json$/, ".check.json");
    let check: any = await runPython(CHECKER, [outPath, checkPath]);
    try { if (fs.existsSync(checkPath)) check = JSON.parse(fs.readFileSync(checkPath, "utf8")); } catch { /* keep whatever stdout gave */ }
    const findings: any[] = check?.ok ? check.findings ?? [] : [];
    const errs = findings.filter((f) => f.severity === "error");
    const builtSheets: any[] = Array.isArray(result.sheets) ? result.sheets : [];
    // W7: the sheets as pictures. A clean formula check said nothing about the 58 clipped cells that shipped.
    const rendered = await renderSheets(outPath, p.id, builtSheets.map((s: any) => String(s?.name ?? "")));
    p.buildPath = outPath;
    p.snapshotPath = typeof result.snapshot === "string" ? result.snapshot : undefined;
    let snapshot: any;
    if (p.snapshotPath) {
      try { snapshot = JSON.parse(fs.readFileSync(p.snapshotPath, "utf8")); }
      catch (e) { process.stderr.write(`snapshot unreadable at ${p.snapshotPath}: ${(e as Error).message}\n`); }
    }
    // W11: every width move the builder took, from its result when it reports them, else off the snapshot's per-sheet lists.
    const adjusted: any[] = Array.isArray(result.adjusted)
      ? result.adjusted
      : (Array.isArray(snapshot?.sheets) ? snapshot.sheets : []).flatMap((sh: any) => (Array.isArray(sh?.adjusted) ? sh.adjusted.map((a: any) => ({ sheet: sh.name, ...a })) : []));
    p.buildSummary = { ...result, adjusted: adjusted.slice(0, 60), fidelity: result.fidelity ?? askedFidelity, check: check?.ok ? { engine: check.engine, evaluated: check.evaluated, summary: check.summary, findings: findings.slice(0, 200) } : { error: "formula check did not run" }, sheetImages: rendered.images };
    p.specPath = specPath;   // the spec of this build, so the next one can be diffed against it
    p.thumbnails = rendered.files;
    // Workbook check (docs/research/workbook-quality.md): clipping, formats, hardcodes, unused inputs, empty bands, cover, totals.
    let quality: any = undefined;
    if (p.snapshotPath) {
      const qPath = specPath.replace(/\.json$/, ".quality.json");
      try {
        const q: any = await runPython(WB_CHECKER, [outPath, p.snapshotPath, specPath, "--out", qPath]);
        const parsed = fs.existsSync(qPath) ? JSON.parse(fs.readFileSync(qPath, "utf8")) : q;
        if (parsed?.ok) quality = { summary: parsed.summary, byKind: parsed.byKind, findings: (parsed.findings ?? []).slice(0, 200) };
      } catch (e) { process.stderr.write(`workbook check did not run: ${(e as Error).message}\n`); }
    }
    if (quality) (p.buildSummary as any).quality = quality;
    saveProject(p);

    const images = await sheetImages(rendered.images.map((im) => ({ index: im.index, file: thumbFor(rendered.files, im.index, "sheet") })));
    const closed = p.notes.filter((n) => n.status !== "open");
    const sum = (k: string) => builtSheets.reduce((n: number, s: any) => n + (Number(s?.[k]) || 0), 0);
    const charts = sum("charts");
    const lines = [
      `Built ${outPath}`,
      `${builtSheets.length} sheets, ${sum("cells")} cells, ${sum("formulas")} formulas${charts ? `, ${charts} charts` : ""}${images.blocks.length ? `, ${images.blocks.length} sheet images` : ""}`,
      check?.ok ? `Formula check (${check.engine}): ${check.summary?.errors ?? 0} errors, ${check.summary?.warnings ?? 0} warnings` : "Formula check did not run.",
    ];
    if (quality) {
      const kinds = Object.entries(quality.byKind ?? {}).sort((a: any, b: any) => (b[1] as number) - (a[1] as number)).slice(0, 4).map(([k, n]) => `${n} ${k}`).join(", ");
      lines.push(`Workbook check: ${quality.summary?.errors ?? 0} errors, ${quality.summary?.warnings ?? 0} warnings${kinds ? ` (${kinds})` : ""}; the list is in the Built view`);
    }
    if (adjusted.length) lines.push(`Adjusted ${adjusted.length} block${adjusted.length === 1 ? "" : "s"} (${result.fidelity ?? askedFidelity})`);
    if (result.warnings?.length) lines.push(`Builder warnings: ${result.warnings.slice(0, 3).join("; ")}${result.warnings.length > 3 ? `, and ${result.warnings.length - 3} more` : ""}`);
    const closing = [
      `Notes honoured: ${closed.filter((n) => n.status === "resolved").length} resolved, ${closed.filter((n) => n.status === "declined").length} declined; tell the user the path and how each note was handled.`,
      errs.length ? "Next: fix each error above, then build_workbook again." : "Next: read the sheet images, fix what a reader would notice through the plan, then give the user the path.",
    ];
    if (errs.length) {
      lines.push("Fix these on the plan or in the payload and rebuild; never patch the .xlsx by hand.");
      const room = MAX_TEXT_LINES - lines.length - closing.length;
      const show = Math.max(1, errs.length > room ? room - 1 : errs.length);
      for (const f of errs.slice(0, show)) lines.push(`- ${f.sheet}!${f.cell}${f.blockId ? ` (${f.blockId})` : ""}: ${f.code}: ${f.message}`);
      if (errs.length > show) lines.push(`- and ${errs.length - show} more; the whole list is in the build summary`);
    }
    lines.push(...closing);

    const tail: string[] = [];
    if (images.blocks.length) tail.push(SHEET_IMAGES_LINE);
    const named = rendered.images.filter((im) => im.sheet).map((im) => `${im.index} ${im.sheet}`);
    if (named.length && images.blocks.length) tail.push(`In order: ${named.join(", ")}.`);
    tail.push(...rendered.notes, ...images.notes);
    logTiming("build_workbook", Date.now() - t0, `${builtSheets.length} sheets, ${images.blocks.length} images`);
    const out = render(p, "built", { build: p.buildSummary, snapshot }, lines.join("\n"));
    if (!tail.length && !images.blocks.length) return out;
    return { ...out, content: [...out.content, ...(tail.length ? [{ type: "text" as const, text: tail.join("\n") }] : []), ...images.blocks] };
  });

  // ---------- app-only tools (called by the canvas) ----------

}

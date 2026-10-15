import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/server";
import { registerAppTool } from "@modelcontextprotocol/ext-apps/server";
import { builderStyleFields } from "./styles.js";
import type { Project } from "../shared/types.js";
import { loadProject, saveProject } from "./store.js";

// dist/server/restyle.js and src/server/restyle.ts both sit two levels below the project root.
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
const SCANNER = path.join(ROOT, "builder", "import_scan.py");
const RESTYLER = path.join(ROOT, "builder", "restyle.py");

type Render = (p: Project, view: "tabs" | "layouts" | "built", extra?: Record<string, unknown>, text?: string) => any;
type RunPython = (script: string, args: string[]) => Promise<any>;

function fail(message: string) {
  return { isError: true as const, content: [{ type: "text" as const, text: message }] };
}

/**
 * restyle_workbook: roadmap C11, the fast path from an imported plan. Re-uses the scan saved by
 * import_workbook (project.scanPath) when it is newer than the source, otherwise scans again, then
 * runs builder/restyle.py to rewrite formats on the original file into a new .xlsx.
 */
export function registerRestyleTool(
  server: McpServer,
  ui: { resourceUri: string },
  _appOnly: { resourceUri: string; visibility: "app"[] },
  render: Render,
  runPython: RunPython,
  HOME: string,
  OUTPUT_DIR: string,
) {
  registerAppTool(server, "restyle_workbook", {
    title: "Restyle the imported workbook in place",
    description: "Fast path for an imported workbook: rewrite formats only (fonts, colours, borders, alignment, number formats where a cell has none, column widths, gridlines, title merges) on every block the scanner detected in the original file, in the chosen style. Values, formulas, charts, data validation and conditional formats are untouched; cells outside detected blocks keep their formatting; a Read Me cover sheet is left as is. Writes a new file in the project's output folder, never over the original. Use when the user wants the file to follow the house (or plain, or template) format without rebuilding it.",
    inputSchema: z.object({
      project: z.string().describe("An imported project id (from import_workbook)"),
      formatting: z.enum(["plain", "house", "template"]).optional().describe("Default: the project's formatting"),
      templatePath: z.string().optional().describe("template mode: the .xlsx whose fonts to reuse"),
      fileName: z.string().optional().describe("File name without extension. Default: <original name> restyled"),
    }),
    _meta: { ui },
  }, async ({ project, formatting, templatePath, fileName }) => {
    const p = loadProject(project);
    if (!p.importedFrom) return fail("This project was not imported. restyle_workbook only works on a project created by import_workbook; use build_workbook instead.");
    if (!fs.existsSync(p.importedFrom)) return fail(`The imported file is no longer at ${p.importedFrom}`);
    const mode = formatting ?? p.formatting;
    const tpl = templatePath ?? p.templatePath;
    if (mode === "template" && !tpl) return fail("template formatting needs a templatePath (or set one on the project).");

    // Reuse the scan saved at import unless the source changed since; otherwise scan again.
    let scan: unknown;
    const fresh = p.scanPath && fs.existsSync(p.scanPath) && fs.statSync(p.scanPath).mtimeMs >= fs.statSync(p.importedFrom).mtimeMs;
    if (fresh) { try { scan = JSON.parse(fs.readFileSync(p.scanPath!, "utf8")); } catch { scan = undefined; } }
    if (!scan) {
      const scanPath = path.join(HOME, "imports", `${Date.now()}.json`);
      fs.mkdirSync(path.dirname(scanPath), { recursive: true });
      const r = await runPython(SCANNER, [p.importedFrom, scanPath]);
      if (!r.ok) return fail(`Scan failed:\n${(r.errors ?? ["unknown"]).map((e: string) => `- ${e}`).join("\n")}`);
      scan = JSON.parse(fs.readFileSync(scanPath, "utf8"));
      p.scanPath = scanPath;
    }

    const outDir = p.outputDir || OUTPUT_DIR;
    fs.mkdirSync(outDir, { recursive: true });
    const original = path.basename(p.importedFrom, path.extname(p.importedFrom));
    const base = (fileName || `${original} restyled`).replace(/[\\/:*?"<>|]/g, "-").trim();
    let outPath = path.join(outDir, `${base}.xlsx`);
    for (let v = 2; fs.existsSync(outPath) || path.resolve(outPath) === path.resolve(p.importedFrom); v++) outPath = path.join(outDir, `${base} v${v}.xlsx`);

    const style = formatting ? { formatting: mode } : builderStyleFields(p.styleId, p.formatting);
    const spec = { inputPath: p.importedFrom, outputPath: outPath, ...style, templatePath: tpl, scan };
    const buildsDir = path.join(HOME, "builds");
    fs.mkdirSync(buildsDir, { recursive: true });
    const specPath = path.join(buildsDir, `${p.id}-${Date.now()}-restyle.json`);
    fs.writeFileSync(specPath, JSON.stringify(spec));
    const result = await runPython(RESTYLER, [specPath]);
    if (!result.ok) return fail(`Restyle failed:\n${(result.errors ?? ["unknown error"]).map((e: string) => `- ${e}`).join("\n")}\nSpec saved at ${specPath}.`);

    if (formatting) { p.formatting = formatting; p.styleId = formatting; }
    if (templatePath) p.templatePath = templatePath;
    p.buildPath = outPath;
    p.buildSummary = result;
    p.snapshotPath = undefined; // a restyle has no cell snapshot; the Built view shows the path
    saveProject(p);
    const sheets = (result.sheets ?? []) as { name: string; blocks: number; cells: number }[];
    return render(p, "built", { build: result }, `Restyled ${p.importedFrom} in ${mode} format -> ${outPath}\n${sheets.map((s) => `- ${s.name}: ${s.blocks} blocks, ${s.cells} cells restyled`).join("\n")}\n${result.warnings?.length ? `warnings: ${result.warnings.join("; ")}\n` : ""}Values, formulas and charts are unchanged. Tell the user the path and any warning that needs a decision (a title left unmerged, a sheet skipped).`);
  });
}

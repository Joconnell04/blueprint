import fs from "node:fs";
import path from "node:path";
import type { McpServer } from "@modelcontextprotocol/server";
import { registerAppTool } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import { HOME, addNote, createProject, makeBlock, makeTab, saveProject, summarize } from "./store.js";
import { blockNotes } from "./import-formulas.js";
import { SCANNER, fail, render, runPython, ui } from "./common.js";

export function registerImportTools(server: McpServer): void {
  registerAppTool(server, "import_workbook", {
    title: "Import an existing workbook",
    description: "Scan an existing .xlsx into a plan: one tab per sheet, blocks detected from borders, bold headers, formula density and chart anchors, each with a confidence. Opens the layout canvas with Detected badges so the user can fix mislabelled blocks before anything is rebuilt. Tabs are marked approved (the sheets already exist); layouts are not.",
    inputSchema: z.object({
      path: z.string().describe("Absolute path to the .xlsx"),
      name: z.string().optional().describe("Project name; default: the file name"),
      brief: z.string().optional().describe("What the user wants to do with it (restyle, extend, rebuild)"),
      formatting: z.enum(["plain", "house", "template"]).default("plain"),
    }),
    _meta: { ui },
  }, async ({ path: file, name, brief, formatting }) => {
    if (!fs.existsSync(file)) return fail(`No file at ${file}`);
    const outPath = path.join(HOME, "imports", `${Date.now()}.json`);
    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    const r = await runPython(SCANNER, [file, outPath]);
    if (!r.ok) return fail(`Scan failed:\n${(r.errors ?? ["unknown"]).map((e: string) => `- ${e}`).join("\n")}`);
    const scan = JSON.parse(fs.readFileSync(outPath, "utf8"));
    const p = createProject({ name: name || path.basename(file, path.extname(file)), brief: brief || `Imported from ${file}`, formatting });
    p.styleId = formatting; p.importedFrom = file; p.scanPath = outPath;
    const low: string[] = [];
    for (const t of scan.tabs as any[]) {
      const tab = makeTab({ name: t.name, purpose: t.purpose || "" });
      tab.approved = true; tab.layoutProposed = t.blocks.length > 0;
      tab.blocks = (t.blocks as any[]).map((b) => {
        // U6: the notes the sheet already carried. The header note is what the block is for, so it becomes
        // the description; the row and total notes ride on the scan and reach Claude through get_plan.
        const notes = blockNotes(b);
        const blk = makeBlock({ id: b.id, type: b.type, title: b.title || `${b.type} at ${b.range}`, row: b.row, col: b.col, rows: b.rows, cols: b.cols, columns: b.columns, rowLabels: b.rowLabels, sample: (b.sample ?? []).slice(0, 4), source: b.source, description: notes.description ?? b.detected?.reason });
        blk.detected = { confidence: b.confidence, signals: b.detected?.signals ?? [], range: b.range };
        if (b.confidence < 0.6) low.push(`${t.name}: ${blk.title} (${b.range}) at ${Math.round(b.confidence * 100)}%`);
        return blk;
      });
      for (const u of t.unclassified ?? []) addNote(p, { tabId: tab.id, text: `Unclassified region ${u.range}: ${u.reason}. Add a block for it or ignore.`, createdBy: "claude", kind: "question" });
      p.tabs.push(tab);
    }
    p.tabsApproved = true;
    saveProject(p);
    return render(p, "layouts", { focusTab: p.tabs[0]?.id }, `Imported ${scan.tabs.length} sheet(s), ${p.tabs.reduce((n, t) => n + t.blocks.length, 0)} blocks.${low.length ? `\nLow-confidence detections to check with the user:\n${low.map((l) => `- ${l}`).join("\n")}` : ""}${scan.warnings?.length ? `\nScanner warnings: ${scan.warnings.join("; ")}` : ""}\nThe user reviews on the canvas; then either propose changes with update_plan or wait for layout approval and build.\n\n${summarize(p)}`);
  });

}

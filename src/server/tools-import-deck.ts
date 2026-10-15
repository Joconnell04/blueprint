import fs from "node:fs";
import path from "node:path";
import type { McpServer } from "@modelcontextprotocol/server";
import { registerAppTool } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import { DECK_BLOCK_TYPES } from "../shared/types.js";
import type { Block, DeckBlockType, Tab } from "../shared/types.js";
import { HOME, addr, autoPlace, createProject, makeBlock, makeTab } from "./store.js";
import { ROOT, fail, render, runPython, ui } from "./common.js";

/** builder/import_pptx.py: the deck half of the import pair (import_scan.py does workbooks). */
export const DECK_SCANNER = path.join(ROOT, "builder", "import_pptx.py");

interface ScanBlock {
  id: string; type: string; title: string; content: string; text: string;
  row: number; col: number; rows: number; cols: number;
  confidence: number; signals: string[]; count?: number; chartKind?: Block["chartKind"];
  shape?: { x: number; y: number; w: number; h: number };
}
interface ScanSlide {
  index: number; name: string; story: string; section?: string; isDivider: boolean;
  headline?: string; kicker?: string; blocks: ScanBlock[];
}

const DECK_TYPES = new Set<string>(DECK_BLOCK_TYPES);

/** A slide's one line for Claude: the headline is the message, the kicker has nowhere better to live. */
function outlineRow(s: ScanSlide): Tab {
  const tab = makeTab({
    name: s.name || `Slide ${s.index}`,
    purpose: s.isDivider ? "Section divider" : s.kicker ? `Kicker: ${s.kicker}` : "",
    message: s.headline ?? "",
    story: s.story ?? "",
    section: s.section,
  });
  tab.layoutProposed = s.blocks.length > 0;
  return tab;
}

export function registerImportDeckTool(server: McpServer): void {
  registerAppTool(server, "import_deck", {
    title: "Import an existing deck",
    description: "Scan an existing .pptx into a plan: one outline row per slide, blocks detected by role (cards, stat, chart, table, image, quote, takeaway, placeholder, bullets, text), each geometry snapped to the 32x18 slide grid with a confidence. The headline and kicker come off each slide as furniture, dividers become sections, page numbers and the logo are dropped. Opens the layout canvas with Detected badges; nothing is approved, so the user fixes mislabelled blocks before anything is rebuilt.",
    inputSchema: z.object({
      path: z.string().describe("Absolute path to the .pptx"),
      name: z.string().optional().describe("Project name; default: the file name"),
    }),
    _meta: { ui },
  }, async ({ path: file, name }) => {
    if (!fs.existsSync(file)) return fail(`No file at ${file}`);
    const outPath = path.join(HOME, "imports", `deck-${Date.now()}.json`);
    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    const r = await runPython(DECK_SCANNER, [file, outPath]);
    if (!r.ok) return fail(`Deck scan failed:\n${(r.errors ?? ["unknown"]).map((e: string) => `- ${e}`).join("\n")}`);
    const scan = JSON.parse(fs.readFileSync(outPath, "utf8")) as { slides: ScanSlide[]; warnings?: string[] };

    const p = createProject({ name: name || path.basename(file, path.extname(file)), brief: `Imported from ${file}`, kind: "deck" });
    p.importedFrom = file;
    p.scanPath = outPath;

    const low: string[] = [];
    const unknown: string[] = [];
    for (const s of scan.slides) {
      const tab = outlineRow(s);
      tab.blocks = s.blocks.map((b) => {
        if (!DECK_TYPES.has(b.type)) unknown.push(`${tab.name}: ${b.type}`);
        const type = (DECK_TYPES.has(b.type) ? b.type : "text") as DeckBlockType;
        const blk = makeBlock({
          id: b.id, type, title: b.title || type, row: b.row, col: b.col, rows: b.rows, cols: b.cols,
          content: b.content, count: b.count, chartKind: b.chartKind,
        }, "deck");
        blk.detected = {
          confidence: b.confidence,
          signals: b.signals ?? [],
          range: `${addr(blk.row, blk.col)}:${addr(blk.row + blk.rows - 1, blk.col + blk.cols - 1)}`,
        };
        if (b.confidence < 0.6) low.push(`${blk.title} on ${tab.name} (${Math.round(b.confidence * 100)}%)`);
        return blk;
      });
      autoPlace(tab.blocks, "deck");
      p.tabs.push(tab);
    }

    const blocks = p.tabs.reduce((n, t) => n + t.blocks.length, 0);
    const sections = [...new Set(p.tabs.map((t) => t.section).filter(Boolean))] as string[];
    const warnings = [...(scan.warnings ?? []), ...(unknown.length ? [`unknown block types imported as text: ${unknown.join(", ")}`] : [])];
    const lines = [
      `Imported ${scan.slides.length} slide(s), ${blocks} block(s) from ${path.basename(file)} into plan ${p.id}.`,
      sections.length ? `Sections: ${sections.join(" / ")}.` : `No section dividers found.`,
      low.length ? `Low confidence (check the type with the user): ${low.slice(0, 3).join("; ")}${low.length > 3 ? ` and ${low.length - 3} more` : ""}.` : `Every block detected above 60% confidence.`,
      warnings.length ? `Scanner: ${warnings.slice(0, 2).join("; ")}.` : "",
      `Headlines are on each slide's message, kickers on its purpose; stories are empty for you to write.`,
      `Nothing is approved: the user fixes types and geometry on the canvas.`,
      `Next: get_plan for the detail, update_plan for fixes, then build_deck once the layouts are approved.`,
    ].filter(Boolean);
    return render(p, "layouts", { focusTab: p.tabs[0]?.id }, lines.join("\n"));
  });
}

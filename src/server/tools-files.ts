import fs from "node:fs";
import path from "node:path";
import type { McpServer } from "@modelcontextprotocol/server";
import { registerAppTool } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import type { Project } from "../shared/types.js";
import { PROJECTS_DIR, loadProject } from "./store.js";
import { appOnly, fail } from "./common.js";

/**
 * Bytes of a built file, for the canvas.
 *
 * The canvas is a sandboxed iframe: it cannot read a file and cannot download one, so "Save copy" pulls the
 * file through this app-only tool and hands the base64 to the host's `ui/download-file`. Hosts cap a tool
 * result at an unpublished size and the reference apps chunk at 500 KB, so this does the same
 * (docs/research/platform-docs.md §1 and §3). App-only keeps the base64 out of model context.
 */

/** 500 KB, the reference chunk in ext-apps' patterns.md. The view asks for this and never more. */
export const MAX_CHUNK = 512_000;

const MIME: Record<string, string> = {
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".pdf": "application/pdf",
  ".json": "application/json",
};

/** A file name the host can write: no separators, no Windows-illegal characters. */
function safeName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, "-").replace(/^\.+/, "").trim();
}

/**
 * The three files a project can hand over, and nothing else. The view names a kind, never a path, so there
 * is no path to escape from: `build` and `pdf` come from what the build recorded on the project, `plan` is
 * the project's own JSON in the planner home.
 */
function resolveFile(p: Project, which: "build" | "pdf" | "plan"): { file: string; name: string } | null {
  if (which === "build") return p.buildPath ? { file: p.buildPath, name: safeName(path.basename(p.buildPath)) } : null;
  if (which === "pdf") {
    const pdf = (p as Project & { pdfPath?: string }).pdfPath ?? (p.buildSummary as { pdfPath?: string } | undefined)?.pdfPath;
    return pdf ? { file: pdf, name: safeName(path.basename(pdf)) } : null;
  }
  const base = safeName(p.fileName || p.name) || p.id;
  return { file: path.join(PROJECTS_DIR, `${p.id}.json`), name: `${base}.plan.json` };
}

export function registerFileTools(server: McpServer): void {
  registerAppTool(server, "ui_file_chunk", {
    description: "Canvas: one chunk of a built file as base64, so the view can hand it to the host's download",
    inputSchema: z.object({
      project: z.string(),
      which: z.enum(["build", "pdf", "plan"]).default("build").describe("build: the .xlsx or .pptx. pdf: the exported deck. plan: the project JSON"),
      offset: z.number().int().min(0).default(0).describe("Byte offset to read from"),
      size: z.number().int().min(1).max(MAX_CHUNK).default(MAX_CHUNK).describe(`Bytes to read, at most ${MAX_CHUNK}`),
    }),
    _meta: { ui: appOnly },
  }, async ({ project, which, offset, size }) => {
    const p = loadProject(project);
    const target = resolveFile(p, which);
    if (!target) return fail(which === "pdf" ? "This project has no PDF export." : "This project has no built file yet.");

    let stat: fs.Stats;
    try { stat = fs.statSync(target.file); }
    catch { return fail(`No file at ${target.file} any more. Rebuild it.`); }
    if (!stat.isFile()) return fail(`${target.file} is not a file.`);

    const total = stat.size;
    const at = Math.min(Math.max(0, offset), total);
    const want = Math.min(size, MAX_CHUNK, total - at);
    let bytes = Buffer.alloc(0);
    if (want > 0) {
      const buf = Buffer.alloc(want);
      const fd = fs.openSync(target.file, "r");
      try { bytes = buf.subarray(0, fs.readSync(fd, buf, 0, want, at)); }
      finally { fs.closeSync(fd); }
    }
    const mimeType = MIME[path.extname(target.name).toLowerCase()] ?? "application/octet-stream";
    return {
      // The base64 rides in structuredContent only; this line is what a text-only host would see.
      content: [{ type: "text" as const, text: `${target.name}: ${bytes.length} bytes at ${at} of ${total}` }],
      structuredContent: { name: target.name, mimeType, total, offset: at, data: bytes.toString("base64") },
    };
  });
}

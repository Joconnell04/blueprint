import fs from "node:fs";
import path from "node:path";
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/server";
import { registerAppTool } from "@modelcontextprotocol/ext-apps/server";
import type { Project, Tab } from "../shared/types.js";
import { HOME, loadProject, saveProject, slugify, newId, readiness, stageOf } from "./store.js";

export const TEMPLATES_DIR = path.join(HOME, "templates");
fs.mkdirSync(TEMPLATES_DIR, { recursive: true });

export interface Template { id: string; name: string; description: string; formatting: Project["formatting"]; styleId?: string; tabs: Tab[]; createdAt: string; fromProject?: string }

export function listTemplates(): Template[] {
  fs.mkdirSync(TEMPLATES_DIR, { recursive: true });
  return fs.readdirSync(TEMPLATES_DIR).filter((f) => f.endsWith(".json")).map((f) => JSON.parse(fs.readFileSync(path.join(TEMPLATES_DIR, f), "utf8")) as Template).sort((a, b) => (a.createdAt < b.createdAt ? 1 : -1));
}

/** Snapshot an approved plan as a template: tabs and blocks, without notes, approvals, samples or detection metadata. */
export function saveTemplate(p: Project, name: string, description: string): Template {
  let id = slugify(name);
  if (fs.existsSync(path.join(TEMPLATES_DIR, `${id}.json`))) id = `${id}-${newId("").slice(1)}`;
  const tabs: Tab[] = p.tabs.map((t) => ({ id: newId("tab"), name: t.name, purpose: t.purpose, approved: false, layoutProposed: t.blocks.length > 0, layoutApproved: false, blocks: t.blocks.map((b) => ({ id: newId("blk"), type: b.type, title: b.title, row: b.row, col: b.col, rows: b.rows, cols: b.cols, columns: b.columns, rowLabels: b.rowLabels, description: b.description, source: undefined })) }));
  // re-point chart sources to the copied ids
  p.tabs.forEach((t, ti) => t.blocks.forEach((b, bi) => { if (b.source) { const si = t.blocks.findIndex((x) => x.id === b.source); if (si >= 0) tabs[ti].blocks[bi].source = tabs[ti].blocks[si].id; } }));
  const tpl: Template = { id, name, description, formatting: p.formatting, styleId: p.styleId, tabs, createdAt: new Date().toISOString(), fromProject: p.id };
  fs.writeFileSync(path.join(TEMPLATES_DIR, `${id}.json`), JSON.stringify(tpl, null, 2));
  return tpl;
}

export function applyTemplate(p: Project, tpl: Template): void {
  p.tabs = tpl.tabs.map((t) => ({ ...t, id: newId("tab"), blocks: t.blocks.map((b) => ({ ...b, id: newId("blk") })) }));
  p.tabs.forEach((t, ti) => t.blocks.forEach((b, bi) => { const src = tpl.tabs[ti].blocks[bi].source; if (src) { const si = tpl.tabs[ti].blocks.findIndex((x) => x.id === src); if (si >= 0) b.source = t.blocks[si].id; } }));
  p.formatting = tpl.formatting; p.styleId = tpl.styleId ?? tpl.formatting;
  p.tabsApproved = false;
  p.notes = [];
}

export function registerTemplateTools(server: McpServer, ui: { resourceUri: string }, appOnly: { resourceUri: string; visibility: "app"[] }, render: (p: Project, view: "tabs" | "layouts" | "built", extra?: Record<string, unknown>, text?: string) => any) {
  registerAppTool(server, "list_templates", {
    title: "List plan templates",
    description: "Templates saved from approved plans (tabs and block layouts, no content). Use one with start_project's template argument or apply_template to a new project so a repeat job is a review, not a proposal.",
    inputSchema: z.object({}),
    _meta: { ui: { visibility: ["model"] } },
  }, async () => {
    const all = listTemplates();
    return { content: [{ type: "text", text: all.length ? all.map((t) => `${t.id}: ${t.name} (${t.tabs.length} tabs, ${t.tabs.reduce((n, x) => n + x.blocks.length, 0)} blocks, ${t.formatting}) ${t.description}`).join("\n") : "No templates yet. The user can save one from an approved plan on the canvas, or call save_template." }] };
  });

  registerAppTool(server, "save_template", {
    title: "Save the plan as a template",
    description: "Save the current tabs and layouts as a reusable template (notes and content are not kept). Only on the user's say-so.",
    inputSchema: z.object({ project: z.string(), name: z.string(), description: z.string().default("") }),
    _meta: { ui: { visibility: ["model"] } },
  }, async ({ project, name, description }) => {
    const tpl = saveTemplate(loadProject(project), name, description);
    return { content: [{ type: "text", text: `Saved template ${tpl.id} (${tpl.tabs.length} tabs).` }] };
  });

  registerAppTool(server, "apply_template", {
    title: "Start a project from a template",
    description: "Replace a project's tabs and layouts with a template's. Tabs come back unapproved so the user reviews them on the canvas; then approve, propose or adjust layouts, and build as usual.",
    inputSchema: z.object({ project: z.string(), template: z.string().describe("Template id from list_templates") }),
    _meta: { ui },
  }, async ({ project, template }) => {
    const p = loadProject(project);
    const tpl = listTemplates().find((t) => t.id === template || t.name.toLowerCase() === template.toLowerCase());
    if (!tpl) return { isError: true as const, content: [{ type: "text" as const, text: `No template "${template}". Call list_templates.` }] };
    applyTemplate(p, tpl); saveProject(p);
    return render(p, "tabs", {}, `Applied template ${tpl.name}. The user reviews the tabs on the canvas.`);
  });

  registerAppTool(server, "ui_templates", {
    description: "Canvas: list templates / save current plan as one / apply one",
    inputSchema: z.object({ project: z.string(), action: z.enum(["list", "save", "apply", "delete"]), name: z.string().optional(), description: z.string().optional(), template: z.string().optional() }),
    _meta: { ui: appOnly },
  }, async ({ project, action, name, description, template }) => {
    const p = loadProject(project);
    if (action === "save" && name) saveTemplate(p, name, description ?? "");
    if (action === "apply" && template) { const tpl = listTemplates().find((t) => t.id === template); if (tpl) { applyTemplate(p, tpl); saveProject(p); } }
    if (action === "delete" && template) { const f = path.join(TEMPLATES_DIR, `${template}.json`); if (fs.existsSync(f)) fs.unlinkSync(f); }
    return { content: [{ type: "text", text: "ok" }], structuredContent: { templates: listTemplates().map((t) => ({ id: t.id, name: t.name, description: t.description, tabs: t.tabs.length, blocks: t.tabs.reduce((n, x) => n + x.blocks.length, 0), formatting: t.formatting })), project: p, warnings: readiness(p), stage: stageOf(p) } };
  });
}

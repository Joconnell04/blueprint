import fs from "node:fs";
import type { McpServer } from "@modelcontextprotocol/server";
import { registerAppTool } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import type { Block } from "../shared/types.js";
import { FIDELITY_HELP } from "../shared/types.js";
import { OUTPUT_DIR, addNote, autoPlace, buildOptions, findTab, foldHeadlineBlocks, loadProject, makeBlock, makeTab, readiness, saveProject, setBuildOptions, spliceDividers, stageOf, saveDefaults } from "./store.js";
import { listStyles, resolveStyle, styleState } from "./styles.js";
import { appOnly } from "./common.js";

export function registerUiTools(server: McpServer): void {
  registerAppTool(server, "ui_get_state", {
    description: "Canvas polling: current project state",
    inputSchema: z.object({ project: z.string(), withSnapshot: z.boolean().optional() }),
    _meta: { ui: appOnly },
  }, async ({ project, withSnapshot }) => {
    const p = loadProject(project);
    // The canvas cannot read files, so a built view that lost its preview asks the poll for the snapshot once.
    const extra: Record<string, unknown> = {};
    if (withSnapshot) {
      if (p.buildSummary !== undefined) extra.build = p.buildSummary;
      if (p.snapshotPath) {
        try { extra.snapshot = JSON.parse(fs.readFileSync(p.snapshotPath, "utf8")); }
        catch (e) { process.stderr.write(`snapshot unreadable at ${p.snapshotPath}: ${(e as Error).message}\n`); }
      }
    }
    return { content: [{ type: "text", text: String(p.revision) }], structuredContent: { project: p, warnings: readiness(p), stage: stageOf(p), style: styleState(p), ...extra } };
  });

  registerAppTool(server, "ui_save_tabs", {
    description: "Canvas: persist the hand-edited outline",
    inputSchema: z.object({ project: z.string(), tabs: z.array(z.object({ id: z.string().optional(), name: z.string(), purpose: z.string().optional(), story: z.string().optional(), divider: z.boolean().optional(), agendaLine: z.string().optional(), message: z.string().optional(), section: z.string().optional() })) }),
    _meta: { ui: appOnly },
  }, async ({ project, tabs }) => {
    const p = loadProject(project);
    const byId = new Map(p.tabs.map((t) => [t.id, t]));
    const text = (v: string | undefined, fallback?: string) => (v === undefined ? fallback : v.trim() || undefined);
    const next = tabs.map((t) => {
      const e = t.id ? byId.get(t.id) : undefined;
      return e ? { ...e, name: t.name, purpose: t.purpose ?? e.purpose, message: text(t.message, e.message), story: text(t.story, e.story), divider: t.divider ?? e.divider, agendaLine: text(t.agendaLine, e.agendaLine) } : makeTab(t);
    });
    const keep = new Set(next.map((t) => t.id));
    p.notes = p.notes.filter((n) => keep.has(n.tabId));
    p.tabs = spliceDividers(next.map((t, i) => ({ tab: t, section: tabs[i].section })));
    p.tabsApproved = false;
    saveProject(p);
    return { content: [{ type: "text", text: "saved" }], structuredContent: { project: p, warnings: readiness(p), stage: stageOf(p) } };
  });

  registerAppTool(server, "ui_save_tab_fields", {
    description: "Canvas: patch one tab or slide's fields",
    inputSchema: z.object({ project: z.string(), tabId: z.string(), message: z.string().optional(), story: z.string().optional(), divider: z.boolean().optional(), agendaLine: z.string().optional(), kicker: z.string().optional(), purpose: z.string().optional(), name: z.string().optional(), section: z.string().optional() }),
    _meta: { ui: appOnly },
  }, async ({ project, tabId, message, story, divider, agendaLine, kicker, purpose, name }) => {
    const p = loadProject(project);
    const t = findTab(p, tabId);
    const text = (v: string | undefined, current?: string) => (v === undefined ? current : v.trim() || undefined);
    // An omitted field keeps what is stored, an empty one clears it. The approval flags, the other rows
    // and the row order are untouched, so a story or a headline typed after approval keeps the approval.
    if (name !== undefined && name.trim()) t.name = name.trim();
    if (purpose !== undefined) t.purpose = purpose.trim();
    if (divider !== undefined) t.divider = divider || undefined;
    t.message = text(message, t.message);
    t.story = text(story, t.story);
    t.agendaLine = text(agendaLine, t.agendaLine);
    t.kicker = text(kicker, t.kicker);
    saveProject(p);   // dividers merge, sections re-derive and the headline takes sentence case
    return { content: [{ type: "text", text: "saved" }], structuredContent: { project: p, warnings: readiness(p), stage: stageOf(p) } };
  });

  registerAppTool(server, "ui_save_brief", {
    description: "Canvas: persist the brief card",
    inputSchema: z.object({ project: z.string(), audience: z.string().optional(), goal: z.string().optional(), takeaway: z.string().optional(), emphasis: z.array(z.string()).optional().describe("Deprecated, ignored") }),
    _meta: { ui: appOnly },
  }, async ({ project, audience, goal, takeaway }) => {
    const p = loadProject(project);
    const text = (v: string | undefined, current?: string) => (v === undefined ? current : v.trim() || undefined);
    // Three fields, nothing else: the emphasis chips are gone (feedback F7).
    p.briefCard = {
      audience: text(audience, p.briefCard?.audience),
      goal: text(goal, p.briefCard?.goal),
      takeaway: text(takeaway, p.briefCard?.takeaway),
    };
    saveProject(p);
    return { content: [{ type: "text", text: "saved" }], structuredContent: { project: p, warnings: readiness(p), stage: stageOf(p) } };
  });

  registerAppTool(server, "ui_save_layout", {
    description: "Canvas: persist hand-edited blocks for one tab or slide",
    inputSchema: z.object({ project: z.string(), tab: z.string(), blocks: z.array(z.any()) }),
    _meta: { ui: appOnly },
  }, async ({ project, tab, blocks }) => {
    const p = loadProject(project);
    const t = findTab(p, tab);
    t.blocks = (blocks as Block[]).map((b) => makeBlock({ ...b, id: b.id }, p.kind));
    const keep = new Set(t.blocks.map((b) => b.id));
    p.notes = p.notes.filter((n) => n.tabId !== t.id || !n.blockId || keep.has(n.blockId));
    if (p.kind === "deck") foldHeadlineBlocks(p, t);
    autoPlace(t.blocks, p.kind);   // deck: rows 5 to 17, overlaps pushed apart
    t.layoutProposed = true;
    t.layoutApproved = false;
    saveProject(p);
    return { content: [{ type: "text", text: "saved" }], structuredContent: { project: p, warnings: readiness(p), stage: stageOf(p) } };
  });

  registerAppTool(server, "ui_note", {
    description: "Canvas: add, edit, delete or reopen a note",
    inputSchema: z.object({ project: z.string(), action: z.enum(["add", "edit", "delete", "reopen", "answer"]), noteId: z.string().optional(), tabId: z.string().optional(), blockId: z.string().optional(), text: z.string().optional() }),
    _meta: { ui: appOnly },
  }, async ({ project, action, noteId, tabId, blockId, text }) => {
    const p = loadProject(project);
    if (action === "add") addNote(p, { tabId: tabId!, blockId, text: text ?? "", createdBy: "user", kind: "instruction" });
    else {
      const n = p.notes.find((n) => n.id === noteId);
      if (n) {
        if (action === "edit") n.text = text ?? n.text;
        if (action === "delete") p.notes = p.notes.filter((x) => x !== n);
        if (action === "reopen") { n.status = "open"; n.response = undefined; n.resolvedAt = undefined; }
        if (action === "answer") { n.response = text ?? ""; n.status = "resolved"; n.resolvedAt = new Date().toISOString(); }
      }
    }
    saveProject(p);
    return { content: [{ type: "text", text: "saved" }], structuredContent: { project: p, warnings: readiness(p), stage: stageOf(p) } };
  });

  registerAppTool(server, "ui_styles", {
    description: "Canvas: list format sources (built-in, skills found on this Mac, template)",
    inputSchema: z.object({ project: z.string() }),
    _meta: { ui: appOnly },
  }, async ({ project }) => { const p = loadProject(project); return { content: [{ type: "text", text: "ok" }], structuredContent: { styles: listStyles(), current: p.styleId ?? p.formatting, templatePath: p.templatePath ?? "", outputDir: p.outputDir ?? OUTPUT_DIR, fileName: p.fileName ?? p.name } }; });

  registerAppTool(server, "ui_set_style", {
    description: "Canvas: set the formatting source, template path, output folder and file name",
    inputSchema: z.object({ project: z.string(), styleId: z.string().optional(), templatePath: z.string().optional(), outputDir: z.string().optional(), fileName: z.string().optional() }),
    _meta: { ui: appOnly },
  }, async ({ project, styleId, templatePath, outputDir, fileName }) => {
    const p = loadProject(project);
    if (styleId) { const st = resolveStyle(styleId); p.styleId = st.id; p.formatting = st.builderMode; }
    if (templatePath !== undefined) p.templatePath = templatePath.trim() || undefined;
    if (outputDir !== undefined) p.outputDir = outputDir.trim() || undefined;
    if (fileName !== undefined) p.fileName = fileName.trim() || undefined;
    saveProject(p);
    saveDefaults({ styleId: p.styleId, formatting: p.formatting, outputDir: p.outputDir, templatePath: p.templatePath }); // the next project starts from these
    return { content: [{ type: "text", text: "saved" }], structuredContent: { project: p, warnings: readiness(p), stage: stageOf(p) } };
  });

  registerAppTool(server, "ui_set_build_options", {
    description: "Canvas: set what the builder may change between the approved plan and the built file. Defaults are balanced fidelity, notes from the story, a cover and an agenda.",
    inputSchema: z.object({
      project: z.string(),
      fidelity: z.enum(["strict", "balanced", "loose"]).optional().describe(FIDELITY_HELP),
      notesFromStory: z.boolean().optional().describe("Speaker notes on every content slide default to its story"),
      cover: z.boolean().optional().describe("Generate a cover slide when the outline has none"),
      agenda: z.boolean().optional().describe("Generate an agenda from the divider list when the outline has none"),
      comments: z.boolean().optional().describe("Workbook: write the closed threads into the file as cell comments"),
    }),
    _meta: { ui: appOnly },
  }, async ({ project, fidelity, notesFromStory, cover, agenda, comments }) => {
    const p = loadProject(project);
    setBuildOptions(p, { fidelity, notesFromStory, cover, agenda, comments });
    saveProject(p);
    return { content: [{ type: "text", text: "saved" }], structuredContent: { project: p, warnings: readiness(p), stage: stageOf(p), buildOptions: buildOptions(p) } };
  });

  registerAppTool(server, "ui_approve", {
    description: "Canvas: approve the outline or a layout",
    inputSchema: z.object({ project: z.string(), what: z.enum(["tabs", "layout", "all_layouts"]), tab: z.string().optional(), approved: z.boolean().default(true) }),
    _meta: { ui: appOnly },
  }, async ({ project, what, tab, approved }) => {
    const p = loadProject(project);
    if (what === "tabs") { p.tabsApproved = approved; p.tabs.forEach((t) => (t.approved = approved)); }
    else if (what === "layout") { const t = findTab(p, tab!); t.layoutApproved = approved; if (approved) t.layoutProposed = true; }
    else { p.tabs.forEach((t) => { t.layoutApproved = approved; if (approved) t.layoutProposed = true; }); if (!approved) { p.buildPath = undefined; p.snapshotPath = undefined; } }
    saveProject(p);
    return { content: [{ type: "text", text: "saved" }], structuredContent: { project: p, warnings: readiness(p), stage: stageOf(p) } };
  });

}

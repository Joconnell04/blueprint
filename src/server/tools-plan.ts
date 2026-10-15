import fs from "node:fs";
import type { McpServer } from "@modelcontextprotocol/server";
import { registerAppTool } from "@modelcontextprotocol/ext-apps/server";
import { z } from "zod";
import type { Block, Brief, Project } from "../shared/types.js";
import { FIDELITY_HELP } from "../shared/types.js";
import { addNote, addQuestions, autoPlace, blockCells, buildOptions, capFirst, cleanBrief, createProject, findBlock, findTab, fitDeckBlocks, foldHeadlineBlocks, listProjects, loadProject, makeBlock, makeTab, promotePlaceholder, readiness, replaceTabs, saveProject, setBuildOptions, stageOf, summarize } from "./store.js";
import { resolveStyle, skillText } from "./styles.js";
import { originalContent } from "./import-formulas.js";
import { listTemplates, applyTemplate } from "./templates.js";
import { BlockInputSchema, BriefSchema, DeckBlockTypeSchema, WorkbookBlockTypeSchema, fail, render, ui, workflow, type View } from "./common.js";

/**
 * P8: which style's skill prose this plan has already been handed. It lives on the project JSON rather than
 * on the Project type because only get_plan reads it; the canvas never sees it.
 */
type PlanWithSkillSent = Project & { skillTextSentFor?: string };

/** One line for the result when a placeholder that named a relationship was turned into a diagram. */
function promotionLine(titles: string[]): string {
  return titles.length
    ? `promoted ${titles.length} placeholder${titles.length === 1 ? "" : "s"} to diagrams: ${titles.join(", ")}. Check the kind, the nodes and the edges, and fix them with update_plan.`
    : "";
}

export function registerPlanTools(server: McpServer): void {
  registerAppTool(server, "start_project", {
    title: "Start a plan",
    description: "Create a plan and get the workflow for it. kind workbook ends in an .xlsx, kind deck in a .pptx. Call this first, then ask the user for whatever the brief is still missing.",
    inputSchema: z.object({
      name: z.string().describe("Short name, becomes the file name, e.g. 'Acme Market Sizing'"),
      kind: z.enum(["workbook", "deck"]).default("workbook").describe("workbook = Excel, deck = PowerPoint"),
      brief: z.string().describe("The user's ask plus any context they gave, verbatim where possible"),
      briefCard: BriefSchema.optional().describe("Audience, goal and the one takeaway, in the user's own words. Fill in what they already said and ask for the rest before proposing; never put a slide count, a format name or a note about your process in here"),
      formatting: z.enum(["plain", "house", "template"]).optional().describe("plain = clean default; house = the built-in house format; template = copy the look of templatePath. The user can change this on the canvas (style popover), including picking a format skill"),
      templatePath: z.string().optional().describe("Absolute path to an .xlsx or .pptx template, only for formatting = template"),
      outputDir: z.string().optional().describe("Absolute directory for the built file. Default: ~/Documents/Blueprint"),
      skipLayouts: z.boolean().default(false).describe("true to skip the layout stage for small jobs"),
      template: z.string().optional().describe("Template id from list_templates to start from; the outline and layouts are pre-filled for the user to review"),
    }),
    _meta: { ui: { visibility: ["model"] } },
  }, async (args) => {
    const p = createProject(args);
    if (args.formatting) p.styleId = args.formatting;
    let tplNote = "";
    if (args.template) { const tpl = listTemplates().find((t) => t.id === args.template || t.name.toLowerCase() === args.template!.toLowerCase()); if (tpl) { applyTemplate(p, tpl); tplNote = `\nFrom template "${tpl.name}": ${tpl.tabs.length} rows and their layouts are filled in. Skip the proposal; call show_plan so the user reviews and approves.`; } else tplNote = `\nNo template named "${args.template}"; starting blank.`; }
    saveProject(p);
    return { content: [{ type: "text", text: `Created "${p.name}" (id: ${p.id}), a ${p.kind} plan.${tplNote}\n\n${workflow(p.kind)}` }], structuredContent: { projectId: p.id, kind: p.kind } };
  });

  registerAppTool(server, "list_projects", {
    title: "List plans",
    description: "List saved plans with their kind, id and stage, so a conversation can resume one.",
    inputSchema: z.object({}),
    _meta: { ui: { visibility: ["model"] } },
  }, async () => {
    const all = listProjects();
    return { content: [{ type: "text", text: all.length ? all.map((p) => `${p.id}: ${p.name} [${p.kind}, ${p.stage}] updated ${p.updatedAt}`).join("\n") : "No plans yet. Call start_project." }] };
  });

  registerAppTool(server, "propose_outline", {
    title: "Propose the outline",
    description: "Layer 1. Set the full ordered outline (tabs for a workbook, slides for a deck) and render the canvas for the user to edit and approve. A deck section is a row of its own with divider true; re-calling replaces the list and rows matched by id or name keep their notes and layouts.",
    inputSchema: z.object({
      project: z.string().describe("Project id"),
      rows: z.array(z.object({
        id: z.string().optional().describe("Keep an existing row id to update it; omit for new rows"),
        name: z.string().describe("Workbook: the tab name, standing alone out of context ('Membership Ramp', not 'Ramp'). Deck: the slide's purpose in 2 to 5 words, not its headline. On a divider row: the section title"),
        purpose: z.string().optional().describe("Workbook: one line on what the tab holds and what it feeds"),
        story: z.string().optional().describe("Deck: 2 to 4 sentences on what this slide does for the reader and what they should believe after it"),
        divider: z.boolean().optional().describe("Deck: this row is a section divider, named after the section. It carries nothing else and becomes a divider slide"),
        agendaLine: z.string().optional().describe("Deck, divider rows only: the one line the generated agenda prints under this section title. Written for the reader, not a label; a content row's is dropped"),
        section: z.string().optional().describe("Deprecated. Sections are divider rows; a section here is turned into one"),
        message: z.string().optional().describe("Workbook only: one line, the point of this tab. A slide's headline is written at layout time"),
      })).min(1).describe("One row per tab (workbook) or slide (deck), in order, dividers included"),
    }),
    _meta: { ui },
  }, async ({ project, rows }) => {
    const p = loadProject(project);
    replaceTabs(p, rows);
    saveProject(p);   // saveProject turns any legacy section into a divider row and merges adjacent ones
    const deck = p.kind === "deck";
    const dividers = p.tabs.filter((t) => t.divider).length;
    const next = deck
      ? `Outline rendered: ${p.tabs.length - dividers} slides and ${dividers} divider${dividers === 1 ? "" : "s"}. A divider is a row in the list, not a field on a slide; drag it to move the section boundary.\nWait for the user to edit and approve it on the canvas. Then: get_plan, write each slide's headline from its story, read the headlines in order as the storyline, and call propose_layouts once for every slide.`
      : `Outline rendered: ${p.tabs.length} tabs. Wait for the user to edit and approve it on the canvas.\nThen: get_plan, then propose_layouts for every tab in one call. Place blocks top to bottom: the tab's own assumptions, the widest raw table, the derivation, the single answer last, each calc below what it reads.`;
    return render(p, "tabs", {}, `${next}\n\n${summarize(p)}`);
  });

  registerAppTool(server, "propose_layouts", {
    title: "Propose layouts",
    description: "Layer 2. For one or more tabs or slides, set the blocks with their positions, content and any questions you have, and render the layout canvas: every deck slide also needs its headline here, one line in sentence case written from that slide's story, and the headlines read top to bottom must still tell the storyline. Call get_plan first so hand edits survive; re-calling for a row replaces that row's blocks, and blocks matched by id keep their notes.",
    inputSchema: z.object({
      project: z.string(),
      layouts: z.array(z.object({
        tab: z.string().describe("Tab or slide id, or exact name"),
        headline: z.string().optional().describe("Deck: the one line this slide argues, in sentence case. Required on every slide that is not a divider"),
        kicker: z.string().optional().describe("Deck: the eyebrow above the headline. Defaults to the section name"),
        blocks: z.array(BlockInputSchema),
      })).min(1),
    }),
    _meta: { ui },
  }, async ({ project, layouts }) => {
    const p = loadProject(project);
    const deck = p.kind === "deck";
    if (!p.tabsApproved) return fail(`The outline is not approved yet. Ask the user to approve it on the canvas, or call update_plan with the approve_tabs op once they say so in chat.`);
    const allowed = deck ? DeckBlockTypeSchema.options : WorkbookBlockTypeSchema.options;
    const wrong = layouts.flatMap((l) => l.blocks.map((b) => b.type)).filter((t) => !allowed.includes(t as never) && t !== "headline");
    if (wrong.length) return fail(`Block types not available in ${p.kind} mode: ${[...new Set(wrong)].join(", ")}. Use: ${allowed.join(", ")}.`);
    // The headline is the slide, not a block: refuse before anything is written rather than build a nameless slide.
    if (deck) {
      const missing = layouts.filter((l) => {
        const t = p.tabs.find((x) => x.id === l.tab) ?? p.tabs.find((x) => x.name.trim().toLowerCase() === l.tab.trim().toLowerCase());
        if (!t || t.divider) return false;
        const own = l.blocks.find((b) => b.type === "headline");
        return !(l.headline ?? "").trim() && !(own?.title ?? own?.content ?? "").trim() && !t.message?.trim();
      });
      if (missing.length) return fail(`No headline for: ${missing.map((l) => `"${l.tab}"`).join(", ")}. Write each one from that slide's story: one line, sentence case, a claim with a verb. Send it as the slide's headline, never as a block.`);
    }
    let focus: string | undefined;
    const asides: string[] = [];
    const promoted: string[] = [];
    for (const l of layouts) {
      const tab = findTab(p, l.tab);
      if (deck && tab.divider) { asides.push(`"${tab.name}" is a divider: no blocks, the builder draws the section slide`); continue; }
      const oldById = new Map(tab.blocks.map((b) => [b.id, b]));
      const blocks: Block[] = l.blocks.map((b) => {
        // A placeholder that names a relationship is a diagram: promote it before it is stored (research item 4).
        const input = { ...b, id: b.id && oldById.has(b.id) ? b.id : undefined, row: b.row ?? 0, col: b.col ?? 0 };
        const promotion = deck ? promotePlaceholder(input) : undefined;
        if (promotion) { Object.assign(input, promotion); promoted.push(input.title); }
        return makeBlock(input, p.kind);
      });
      const keep = new Set(blocks.map((b) => b.id));
      p.notes = p.notes.filter((n) => n.tabId !== tab.id || !n.blockId || keep.has(n.blockId));
      tab.blocks = blocks;
      l.blocks.forEach((b, i) => addQuestions(p, tab.id, blocks[i].id, b.questions));
      if (deck) {
        if ((l.headline ?? "").trim()) tab.message = capFirst(l.headline!.trim());
        if ((l.kicker ?? "").trim()) tab.kicker = l.kicker!.trim();
        const head = foldHeadlineBlocks(p, tab);
        if (head) asides.push(`"${tab.name}": the headline block is now the slide headline`);
      }
      autoPlace(tab.blocks, p.kind);
      tab.layoutProposed = true;
      tab.layoutApproved = false;
      focus ??= tab.id;
    }
    saveProject(p);
    const next = deck
      ? `Layout canvas rendered. Blocks live in rows 5 to 17; rows 1 to 4 and row 18 are furniture, and overlaps were pushed apart.${asides.length ? `\n${asides.join("; ")}.` : ""}${promoted.length ? `\n${promotionLine(promoted)}` : ""}\nWait for the user to arrange and approve it. Then: get_plan, close every thread with resolve_notes, and call build_deck with the real content per block.`
      : `Layout canvas rendered. Wait for the user to arrange and approve it.\nThen: push the real columns, line items and 2 to 4 real sample rows into the plan with update_plan so the canvas stops showing placeholders, call get_plan, close every thread with resolve_notes, and build_workbook using the {{blockId!...}} references.`;
    return render(p, "layouts", { focusTab: focus }, `${next}\n\n${summarize(p)}`);
  });

  registerAppTool(server, "get_plan", {
    title: "Get the current plan",
    description: "Read the plan as it is now: the brief, every row with its headline, story and blocks (a divider row is a section slide), the build options, open threads, readiness warnings, and for a workbook the cell addresses each block will occupy. Call it before propose_layouts, resolve_notes and any build.",
    inputSchema: z.object({ project: z.string(), skillText: z.boolean().optional().describe("Send the format skill's rules again even though this plan has had them; they are sent once per style and referenced after") }),
    _meta: { ui: { visibility: ["model"] } },
  }, async ({ project, skillText: again }) => {
    const p = loadProject(project);
    const cells: Record<string, unknown> = {};
    if (p.kind !== "deck") for (const t of p.tabs) for (const b of t.blocks) cells[b.id] = { tab: t.name, ...blockCells(b) };
    const style = resolveStyle(p.styleId ?? p.formatting);
    const skill = skillText(style);
    // P8: the skill prose runs to thousands of tokens and never changes mid-plan. Send it the first time this
    // plan uses the style, then reference it; a style change or `skillText: true` sends it again.
    const sendSkill = !!skill && (again === true || (p as PlanWithSkillSent).skillTextSentFor !== style.id);
    if (sendSkill) { (p as PlanWithSkillSent).skillTextSentFor = style.id; saveProject(p); }
    const skillNote = !skill ? ""
      : sendSkill ? `. Follow this formatting skill for everything the builder cannot do mechanically (naming, notes, reading order, what to leave out):\n---\n${skill}\n---`
      : `\nStyle: ${style.name} (rules sent earlier; call get_plan with \`skillText: true\` to see them again)`;
    const styleNote = `formatting: ${style.name} (${style.source}); the builder applies ${style.builderMode} mechanics${skillNote}`;
    let original = "";
    if (p.scanPath && fs.existsSync(p.scanPath)) {
      try {
        const oc = originalContent(p.scanPath);
        const lines = Object.entries(oc).filter(([id]) => p.tabs.some((t) => t.blocks.some((b) => b.id === id))).map(([id, o]) => `${id} (${o.tab}): columns ${JSON.stringify(o.columns ?? [])}; rows ${JSON.stringify(o.rows.slice(0, 12))}${o.rows.length > 12 ? ` ... ${o.rows.length} rows` : ""}${o.rowNotes ? `; rowNotes ${JSON.stringify(o.rowNotes)}` : ""}${o.totalNote ? `; totalNote ${JSON.stringify(o.totalNote)}` : ""}${o.unresolved.length ? `; unresolved refs: ${o.unresolved.slice(0, 6).join(", ")}` : ""}`);
        if (lines.length) original = `\n\noriginal content from the imported file, formulas already translated to block references (reuse these in build_workbook, rowNotes and totalNote included, so the rebuild keeps the notes the sheet carried; unresolved refs point outside any detected block and need your judgement):\n${lines.join("\n")}`;
      } catch (e: any) { original = `\n\n(could not read the import scan: ${e?.message ?? e})`; }
    }
    const stage = stageOf(p);
    const build = p.kind === "deck" ? "build_deck" : "build_workbook";
    const next = stage === "brief" ? "next: propose_outline." : stage === "tabs" ? "next: wait for the outline to be approved, then propose_layouts." : stage === "layouts" ? "next: propose or fix layouts, then wait for approval." : stage === "ready" ? `next: close every open thread with resolve_notes, then ${build}.` : "next: the file is built; re-open it with show_plan.";
    const geometry = p.kind === "deck"
      ? `slide geometry: 32 columns by 18 rows, 960x540 pt. Block positions above are row, col, rows, cols on that grid. Blocks live in rows 5 to 17: rows 1 to 4 hold the kicker, the headline and the hairline, row 18 the source and the page number. A slide's headline is its own line above, not a block.`
      : `cell addresses by block id (use {{blockId!...}} in formulas and the builder resolves them):\n${JSON.stringify(cells, null, 1)}`;
    return { content: [{ type: "text", text: `${summarize(p)}\n\n${styleNote}\n\n${geometry}${original}\n\n${next}` }], structuredContent: { project: p, kind: p.kind, brief: p.briefCard, buildOptions: buildOptions(p), warnings: readiness(p), cells } };
  });

  registerAppTool(server, "update_plan", {
    title: "Edit the plan from chat",
    description: "Apply targeted edits the user asked for in chat: the brief, a slide's headline or story, a divider's agenda line, adding and moving rows, dividers and blocks, notes, skipLayouts, formatting, the build options, and approvals on the user's explicit say-so. A headline change is set_headline on the slide, never a new block. The canvas already on screen updates itself; this does not open a new one.",
    inputSchema: z.object({
      project: z.string(),
      ops: z.array(z.discriminatedUnion("op", [
        z.object({ op: z.literal("add_tab"), name: z.string(), purpose: z.string().optional(), message: z.string().optional(), story: z.string().optional(), divider: z.boolean().optional().describe("Deck: add a section divider row named `name`"), section: z.string().optional().describe("Deprecated: a section here becomes a divider row above this one"), position: z.number().int().optional().describe("0-based index; default end") }),
        z.object({ op: z.literal("rename_tab"), tab: z.string(), name: z.string().optional(), purpose: z.string().optional() }),
        z.object({ op: z.literal("set_brief"), brief: BriefSchema.describe("Only the fields you are changing") }),
        z.object({ op: z.literal("set_message"), tab: z.string(), message: z.string().describe("One line: the point of this tab. On a slide this is the headline; set_headline says so plainly") }),
        z.object({ op: z.literal("set_headline"), tabId: z.string().describe("Slide id, or exact name"), headline: z.string().describe("The one line this slide argues, sentence case. Seeds the slide's headline; never add a headline block") }),
        z.object({ op: z.literal("set_kicker"), tabId: z.string(), kicker: z.string().describe("Deck: the eyebrow above the headline; empty clears it back to the section name") }),
        z.object({ op: z.literal("set_agenda_line"), tabId: z.string().describe("Divider row id, or exact name"), agendaLine: z.string().describe("Deck: the one line the generated agenda prints under this section title; empty clears it and the agenda falls back to the first story under the divider") }),
        z.object({ op: z.literal("set_story"), tab: z.string(), story: z.string().describe("Deck: 2 to 4 sentences on what this slide does for the reader") }),
        z.object({ op: z.literal("set_section"), tab: z.string(), section: z.string().describe("Deprecated. A section is a divider row: this inserts one above the slide, and an empty string removes the divider directly above it") }),
        z.object({ op: z.literal("remove_tab"), tab: z.string() }),
        z.object({ op: z.literal("reorder_tabs"), order: z.array(z.string()).describe("All tab ids or names in the new order") }),
        z.object({ op: z.literal("add_block"), tab: z.string(), block: BlockInputSchema }),
        z.object({ op: z.literal("update_block"), blockId: z.string(), patch: BlockInputSchema.partial() }),
        z.object({ op: z.literal("move_block"), blockId: z.string(), row: z.number().int().optional(), col: z.number().int().optional(), rows: z.number().int().optional(), cols: z.number().int().optional(), toTab: z.string().optional() }),
        z.object({ op: z.literal("remove_block"), blockId: z.string() }),
        z.object({ op: z.literal("add_note"), tab: z.string(), blockId: z.string().optional(), text: z.string(), kind: z.enum(["question", "instruction"]).default("question").describe("question = you are asking the user; instruction = you are recording something the user said in chat") }),
        z.object({ op: z.literal("set"), skipLayouts: z.boolean().optional(), formatting: z.enum(["plain", "house", "template"]).optional(), styleId: z.string().optional().describe("plain | house | skill:<name> | template"), templatePath: z.string().optional(), outputDir: z.string().optional(), fileName: z.string().optional(), name: z.string().optional(), exportPdf: z.boolean().optional().describe("Deck: also write a PDF when LibreOffice is installed") }),
        z.object({ op: z.literal("set_build_options"), fidelity: z.enum(["strict", "balanced", "loose"]).optional().describe(FIDELITY_HELP), notesFromStory: z.boolean().optional().describe("Speaker notes on every content slide default to its story"), cover: z.boolean().optional().describe("Generate a cover slide when the outline has none"), agenda: z.boolean().optional().describe("Generate an agenda from the divider list when the outline has none"), comments: z.boolean().optional().describe("Workbook: write the closed threads into the file as cell comments"), }),
        z.object({ op: z.literal("approve_tabs") }),
        z.object({ op: z.literal("approve_layout"), tab: z.string().optional().describe("Omit to approve all proposed layouts") }),
        z.object({ op: z.literal("unapprove"), what: z.enum(["tabs", "layouts"]) }),
      ])).min(1),
    }),
    _meta: { ui: { visibility: ["model"] } },
  }, async ({ project, ops }) => {
    const p = loadProject(project);
    let view: View = stageOf(p) === "tabs" || stageOf(p) === "brief" ? "tabs" : "layouts";
    void view;
    let focus: string | undefined;
    const headlineFolds: string[] = [];
    const promoted: string[] = [];
    for (const op of ops) {
      switch (op.op) {
        case "add_tab": {
          const t = makeTab(op);
          const at = op.position ?? p.tabs.length;
          // A section named here is a divider row above the new one, not a field on it.
          const above = (op.section ?? "").trim();
          const rows: typeof p.tabs = above && !p.tabs[at - 1]?.divider ? [makeTab({ name: above, divider: true }), t] : [t];
          p.tabs.splice(at, 0, ...rows);
          p.tabsApproved = false; view = "tabs"; break;
        }
        case "set_headline": { const t = findTab(p, op.tabId); t.message = op.headline.trim() ? capFirst(op.headline.trim()) : undefined; focus = t.id; break; }
        case "set_kicker": { const t = findTab(p, op.tabId); t.kicker = op.kicker.trim() || undefined; focus = t.id; break; }
        case "set_agenda_line": { const t = findTab(p, op.tabId); t.agendaLine = op.agendaLine.trim() || undefined; focus = t.id; break; }
        case "rename_tab": { const t = findTab(p, op.tab); if (op.name) t.name = op.name; if (op.purpose) t.purpose = op.purpose; break; }
        case "set_brief": { const patch = cleanBrief({ ...p.briefCard, ...Object.fromEntries(Object.entries(op.brief).filter(([, v]) => v !== undefined)) } as Brief); p.briefCard = patch ?? {}; break; }
        case "set_message": { findTab(p, op.tab).message = op.message.trim() || undefined; break; }
        case "set_story": { findTab(p, op.tab).story = op.story.trim() || undefined; break; }
        case "set_section": {
          // Sections are rows: name one and a divider goes in above the slide, clear it and the divider above goes.
          const t = findTab(p, op.tab);
          const at = p.tabs.indexOf(t);
          const want = op.section.trim();
          const above = p.tabs[at - 1];
          if (!want) { if (above?.divider) { p.tabs.splice(at - 1, 1); p.notes = p.notes.filter((n) => n.tabId !== above.id); } }
          else if (above?.divider) above.name = want;
          else p.tabs.splice(at, 0, makeTab({ name: want, divider: true }));
          view = "tabs"; break;
        }
        case "remove_tab": { const t = findTab(p, op.tab); p.tabs = p.tabs.filter((x) => x !== t); p.notes = p.notes.filter((n) => n.tabId !== t.id); p.tabsApproved = false; view = "tabs"; break; }
        case "reorder_tabs": { const ordered = op.order.map((r) => findTab(p, r)); const rest = p.tabs.filter((t) => !ordered.includes(t)); p.tabs = [...ordered, ...rest]; view = "tabs"; break; }
        case "add_block": {
          const t = findTab(p, op.tab);
          const input = { ...op.block, id: undefined, row: op.block.row ?? 0, col: op.block.col ?? 0 };
          const promotion = p.kind === "deck" ? promotePlaceholder(input) : undefined;
          if (promotion) { Object.assign(input, promotion); promoted.push(input.title); }
          const b = makeBlock(input, p.kind);
          if (!b.row && p.kind !== "deck") { const maxRow = t.blocks.reduce((m, x) => Math.max(m, x.row + x.rows), 1); b.row = maxRow + 1; b.col = 2; }
          t.blocks.push(b);
          autoPlace(t.blocks, p.kind);
          if (p.kind === "deck") { const head = foldHeadlineBlocks(p, t); if (head) headlineFolds.push(`"${t.name}": a headline block became the slide headline`); }
          addQuestions(p, t.id, b.id, op.block.questions); t.layoutProposed = true; t.layoutApproved = false; view = "layouts"; focus = t.id; break;
        }
        case "update_block": {
          const { tab, block } = findBlock(p, op.blockId);
          const { questions, ...patch } = op.patch;
          Object.assign(block, Object.fromEntries(Object.entries(patch).filter(([, v]) => v !== undefined)));
          if (p.kind === "deck") {
            const promotion = promotePlaceholder(block);
            if (promotion) { Object.assign(block, promotion); promoted.push(block.title); }
            fitDeckBlocks(tab.blocks);
          }
          addQuestions(p, tab.id, block.id, questions); tab.layoutApproved = false; view = "layouts"; focus = tab.id; break;
        }
        case "move_block": {
          const { tab, block } = findBlock(p, op.blockId);
          if (op.toTab) { const dest = findTab(p, op.toTab); tab.blocks = tab.blocks.filter((b) => b !== block); dest.blocks.push(block); p.notes.filter((n) => n.blockId === block.id).forEach((n) => (n.tabId = dest.id)); dest.layoutApproved = false; focus = dest.id; } else focus = tab.id;
          if (op.row) block.row = op.row; if (op.col) block.col = op.col; if (op.rows) block.rows = op.rows; if (op.cols) block.cols = op.cols;
          if (p.kind === "deck") { const dest = op.toTab ? findTab(p, op.toTab) : tab; fitDeckBlocks(dest.blocks); }
          tab.layoutApproved = false; view = "layouts"; break;
        }
        case "remove_block": { const { tab, block } = findBlock(p, op.blockId); tab.blocks = tab.blocks.filter((b) => b !== block); p.notes = p.notes.filter((n) => n.blockId !== block.id); tab.layoutApproved = false; view = "layouts"; focus = tab.id; break; }
        case "add_note": { const t = findTab(p, op.tab); if (op.blockId) findBlock(p, op.blockId); addNote(p, { tabId: t.id, blockId: op.blockId, text: op.text, createdBy: op.kind === "instruction" ? "user" : "claude", kind: op.kind }); focus = t.id; break; }
        case "set": { if (op.skipLayouts != null) p.skipLayouts = op.skipLayouts; if (op.formatting) { p.formatting = op.formatting; p.styleId = op.formatting; } if (op.styleId) { const st = resolveStyle(op.styleId); p.styleId = st.id; p.formatting = st.builderMode; } if (op.templatePath) p.templatePath = op.templatePath; if (op.outputDir) p.outputDir = op.outputDir; if (op.fileName) p.fileName = op.fileName; if (op.name) p.name = op.name; if (op.exportPdf != null) p.exportPdf = op.exportPdf; break; }
        case "set_build_options": { setBuildOptions(p, { fidelity: op.fidelity, notesFromStory: op.notesFromStory, cover: op.cover, agenda: op.agenda, comments: op.comments }); break; }
        case "approve_tabs": { p.tabsApproved = true; p.tabs.forEach((t) => (t.approved = true)); view = "layouts"; break; }
        case "approve_layout": { const tabs = op.tab ? [findTab(p, op.tab)] : p.tabs.filter((t) => t.layoutProposed); tabs.forEach((t) => (t.layoutApproved = true)); view = "layouts"; break; }
        case "unapprove": { if (op.what === "tabs") p.tabsApproved = false; else { p.tabs.forEach((t) => (t.layoutApproved = false)); p.buildPath = undefined; p.snapshotPath = undefined; } break; }
      }
    }
    saveProject(p);
    void focus;
    return { content: [{ type: "text", text: `Applied ${ops.length} edit(s).${headlineFolds.length ? ` ${headlineFolds.join("; ")}.` : ""}${promoted.length ? ` ${promotionLine(promoted)}` : ""} The open canvas refreshes itself; call show_plan only if the user cannot see one.\n\n${summarize(p)}` }], structuredContent: { project: p, warnings: readiness(p), stage: stageOf(p) } };
  });

  registerAppTool(server, "show_plan", {
    title: "Show the plan canvas",
    description: "Re-render the canvas at the current stage without changing anything (e.g. the user asks to see it again).",
    inputSchema: z.object({ project: z.string(), view: z.enum(["tabs", "layouts", "built"]).optional(), tab: z.string().optional() }),
    _meta: { ui },
  }, async ({ project, view, tab }) => {
    const p = loadProject(project);
    const stage = stageOf(p);
    const v: View = view ?? (stage === "built" ? "built" : stage === "tabs" || stage === "brief" ? "tabs" : "layouts");
    return render(p, v, { focusTab: tab ? findTab(p, tab).id : undefined });
  });

  registerAppTool(server, "resolve_notes", {
    title: "Resolve notes",
    description: "Close threads, which the build refuses to run around. Resolve one of the user's instructions only when the build will honour it (say what you did), or decline it with the reason. Your own questions wait for the user's answer on the canvas; withdraw one (status declined) only when it stops mattering, never to unblock a build.",
    inputSchema: z.object({
      project: z.string(),
      resolutions: z.array(z.object({ noteId: z.string(), status: z.enum(["resolved", "declined"]), response: z.string().describe("One or two sentences the user will read") })).min(1),
    }),
    _meta: { ui: { visibility: ["model"] } },
  }, async ({ project, resolutions }) => {
    const p = loadProject(project);
    const missing: string[] = [];
    for (const r of resolutions) {
      const n = p.notes.find((n) => n.id === r.noteId);
      if (!n) { missing.push(r.noteId); continue; }
      n.status = r.status; n.response = r.response; n.resolvedAt = new Date().toISOString();
    }
    saveProject(p);
    const open = p.notes.filter((n) => n.status === "open").length;
    return { content: [{ type: "text", text: `${missing.length ? `Unknown note ids: ${missing.join(", ")}. ` : ""}${open} note(s) still open.\n\n${summarize(p)}` }], structuredContent: { project: p, warnings: readiness(p), stage: stageOf(p) } };
  });

}

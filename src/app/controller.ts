import { App, applyHostFonts, applyHostStyleVariables } from "@modelcontextprotocol/ext-apps";
import type { AppToolCallback, McpUiHostStyles, RegisteredAppTool, StandardSchemaV1 } from "@modelcontextprotocol/ext-apps";
import { z } from "zod/v4";
import type { Block, Brief, Note, Project, State, Tab, View } from "./model.js";
import { summarizePlan, colLetter } from "./model.js";

type Listener = () => void;

/** Build stamp: vite replaces the identifier when the config defines it, otherwise this falls back. */
declare const __APP_VERSION__: string | undefined;
const VERSION = typeof __APP_VERSION__ === "string" ? __APP_VERSION__ : "0.3.0";

/**
 * One outline row on the wire to ui_save_tabs. An omitted field keeps what the server already has.
 * `divider: true` is a section row; the server derives every slide's `section` from where the dividers sit,
 * so the canvas never sends one.
 */
export interface OutlineRow { id?: string; name: string; purpose?: string; story?: string; divider?: boolean; agendaLine?: string }

/** The fields ui_save_tab_fields can patch on one row. Omitted keeps what is stored, "" clears it. */
export interface TabFields { name?: string; purpose?: string; message?: string; story?: string; divider?: boolean; agendaLine?: string }

/** The server drops a blank optional field; normalise the local copy the same way so a poll does not bounce it back. */
function normalizeFields(f: TabFields): Partial<Tab> {
  const out: Partial<Tab> = {};
  if (f.name !== undefined && f.name.trim()) out.name = f.name.trim();
  if (f.purpose !== undefined) out.purpose = f.purpose.trim();
  if (f.divider !== undefined) out.divider = f.divider || undefined;
  for (const k of ["message", "story", "agendaLine"] as const) if (f[k] !== undefined) out[k] = f[k]!.trim() || undefined;
  return out;
}

/**
 * How tall the canvas asks to be, per view. The outline sizes to its content so a four-row plan is not a
 * 760 px box with an empty band under it; the grid views keep the fixed box they scroll inside.
 */
const HEIGHT: Record<View, { min: number; max: number; fit: boolean }> = {
  tabs: { min: 320, max: 760, fit: true },
  layouts: { min: 440, max: 760, fit: false },
  built: { min: 440, max: 760, fit: false },
};

/** Clipboard in a sandboxed iframe: the async API first, the old selection trick when the host blocks it. */
async function copyText(text: string): Promise<boolean> {
  try { await navigator.clipboard.writeText(text); return true; } catch { /* blocked in the sandbox */ }
  try {
    const ta = document.createElement("textarea");
    ta.value = text; ta.setAttribute("readonly", ""); ta.style.position = "fixed"; ta.style.opacity = "0";
    document.body.append(ta); ta.select();
    const ok = document.execCommand("copy"); ta.remove();
    return ok;
  } catch { return false; }
}

/** Owns the App connection, current state, persistence calls, model-context sync and the poll loop. */
export class Controller {
  app: App;
  state: State | null = null;
  selectedTab: string | null = null;
  selectedBlock: string | null = null;
  /** Every selected block id; `selectedBlock` is the primary (last clicked) and is always a member when set. */
  selectedBlocks = new Set<string>();
  interacting = false;
  renderId = 0;        // the render sequence this canvas was created with
  superseded = false;  // a newer canvas exists further down the chat
  pinned = false;      // user chose to keep using this canvas anyway
  private listeners: Listener[] = [];
  private pollTimer: number | null = null;
  private ctxTimer: number | null = null;
  private saveTimers = new Map<string, number>();
  /** Per-row field patches still inside their debounce window, merged so the save carries all of them. */
  private tabPatches = new Map<string, TabFields>();
  /** The call each pending debounce will make, kept so a teardown can make it without waiting the window out. */
  private pendingSaves = new Map<string, () => Promise<unknown>>();
  toast: ((msg: string) => void) | null = null;
  /** Block ids changed by the last remote (chat) edit; the layout view flashes them once. */
  changed = new Set<string>();
  /** Partial tool arguments while Claude is still streaming the call that created this canvas. */
  partial: { tool: string; args: any } | null = null;
  /** Error text from the tool call that created this canvas, when it produced no plan. */
  error: string | null = null;
  /** The host cancelled the call that would have filled this canvas. */
  cancelled = false;
  /** Brief card: null follows the fields (a filled one collapses it), true and false are the user's choice. */
  briefOpen: boolean | null = null;
  /** A `data-focus` key the next draw should put the caret in, for a field the draw itself creates. */
  pendingFocus: string | null = null;

  constructor() {
    // `tools.listChanged` is what lets the host call the canvas back (registerTools below, backlog P5).
    this.app = new App({ name: "Blueprint", version: VERSION }, { tools: { listChanged: true } }, { autoResize: true });
  }

  onChange(fn: Listener) { this.listeners.push(fn); }
  emit() { for (const l of this.listeners) l(); }

  setState(s: State) {
    this.state = s;
    if (s.focusTab) { this.selectedTab = s.focusTab; s.focusTab = undefined; } // consume once; later re-renders keep the user's choice
    if (!this.selectedTab || !s.project.tabs.some((t) => t.id === this.selectedTab)) this.selectedTab = s.project.tabs[0]?.id ?? null;
    if (this.selectedBlock && !s.project.tabs.some((t) => t.blocks.some((b) => b.id === this.selectedBlock))) this.selectedBlock = null;
    this.emit();
    if (this.wantsSnapshot()) void this.fetchSnapshot(); // a built state that arrived without its preview
  }

  // ---- built preview ----

  /**
   * The built view draws from the snapshot JSON, which only the server can read: a state that reaches this
   * canvas from a poll or from another tool carries no snapshot, so ask for one. Once per revision, so a
   * project whose snapshot file is gone is not asked on every poll.
   */
  private snapAskedAt = -1;
  private wantsSnapshot(): boolean {
    return !!this.state && this.state.stage === "built" && !this.state.snapshot && this.snapAskedAt !== this.project.revision;
  }
  private takeSnapshot(sc: any): boolean {
    if (!sc?.snapshot || !this.state) return false;
    this.state.snapshot = sc.snapshot;
    if (sc.build !== undefined) this.state.build = sc.build;
    return true;
  }
  /** Fetch it now, rather than leaving the view without a preview until the next poll. */
  async fetchSnapshot(): Promise<void> {
    if (!this.wantsSnapshot()) return;
    this.snapAskedAt = this.project.revision;
    try {
      const r = await this.app.callServerTool({ name: "ui_get_state", arguments: { project: this.project.id, withSnapshot: true } });
      if (this.takeSnapshot(r.structuredContent)) this.emit();
    } catch { this.snapAskedAt = -1; } // transport hiccup: let the next poll ask again
  }

  get project(): Project { return this.state!.project; }
  get tab(): Tab | undefined { return this.project.tabs.find((t) => t.id === this.selectedTab); }
  get block(): Block | undefined { return this.tab?.blocks.find((b) => b.id === this.selectedBlock); }
  notesFor(tabId: string, blockId?: string): Note[] { return this.project.notes.filter((n) => n.tabId === tabId && (blockId ? n.blockId === blockId : !n.blockId)); }

  async connect() {
    this.app.ontoolresult = (result) => {
      const sc = result.structuredContent as (State & { renderId?: number }) | undefined;
      if (sc && sc.project) { this.partial = null; this.cancelled = false; this.renderId = sc.renderId ?? sc.project.renderSeq ?? 0; this.setState(sc); }
      else if (result.isError) { this.partial = null; this.cancelled = false; this.error = (result.content?.find((c: any) => c.type === "text") as any)?.text ?? "The tool returned an error."; this.emit(); }
    };
    const sketch = (params: any) => {
      if (this.state) return;
      const tool = String(this.app.getHostContext()?.toolInfo?.tool?.name ?? "");
      const a = params.arguments ?? {};
      if (!(Array.isArray(a.tabs) || Array.isArray(a.layouts))) return; // only proposals have something to sketch
      this.partial = { tool, args: a }; this.emit();
    };
    this.app.ontoolinputpartial = sketch;
    this.app.ontoolinput = sketch;
    // The user stopped the call: drop the half-drawn sketch rather than leave it frozen mid-stream.
    this.app.ontoolcancelled = () => { this.partial = null; this.error = null; this.cancelled = true; this.emit(); this.toast?.("Cancelled"); };
    this.app.onhostcontextchanged = (ctx) => this.applyHost(ctx);
    this.wireTeardown();   // registered before connect, so a teardown request cannot arrive unhandled
    this.registerTools();  // and before connect, so the tool list rides out with the handshake
    await this.app.connect();
    this.applyHost(this.app.getHostContext());
    // A host that read the list during the handshake has it already; this is for one that reads it later.
    void this.app.sendToolListChanged().catch(() => { /* the host takes no app tools */ });
    this.startPolling();
  }

  // ---- app tools: the host drives the open canvas (docs/research/platform-docs.md §1, backlog P5) ----

  /** Every tool this canvas registered, by name. The host calls them; a test reads them without one. */
  readonly appTools = new Map<string, RegisteredAppTool>();

  /**
   * Four tools that move the selection or the view and re-render in place, so Claude can point at a
   * slide or a block without spending a turn on a whole new canvas. Each one answers in a single line.
   * Nothing here calls the server: if the host never calls them, the canvas is exactly what it was.
   */
  private registerTools(): void {
    const add = <S extends StandardSchemaV1>(name: string, description: string, inputSchema: S, run: (args: StandardSchemaV1.InferOutput<S>) => string) => {
      // The SDK's callback type is conditional on the schema, which a generic S cannot resolve; the shape is the same.
      const cb = ((args: StandardSchemaV1.InferOutput<S>) => ({ content: [{ type: "text" as const, text: run(args) }] })) as AppToolCallback<S, undefined>;
      this.appTools.set(name, this.app.registerTool<undefined, S>(name, { description, inputSchema }, cb));
    };
    add("focus_slide", "Show one slide or tab of the open plan on the canvas.",
      z.object({ tabId: z.string().optional(), name: z.string().optional() }),
      ({ tabId, name }) => {
        if (!this.state) return "No plan open";
        const t = this.findTab(tabId, name);
        if (!t) return `No ${this.rowWord} "${name ?? tabId ?? ""}"`;
        this.focusTab(t.id);
        return `Focused ${t.name}`;
      });
    add("select_block", "Select one block on the open canvas.",
      z.object({ blockId: z.string() }),
      ({ blockId }) => {
        if (!this.state) return "No plan open";
        const tab = this.project.tabs.find((t) => t.blocks.some((b) => b.id === blockId));
        const block = tab?.blocks.find((b) => b.id === blockId);
        if (!tab || !block) return `No block ${blockId}`;
        this.selectedTab = tab.id;
        this.selectedBlock = blockId;
        this.selectedBlocks = new Set([blockId]);
        this.emit();
        return `Selected ${block.title}`;
      });
    add("scroll_to", "Bring one slide, tab or block into view on the canvas.",
      z.object({ tabId: z.string(), blockId: z.string().optional() }),
      ({ tabId, blockId }) => {
        if (!this.state) return "No plan open";
        const t = this.findTab(tabId);
        if (!t) return `No ${this.rowWord} "${tabId}"`;
        const block = blockId ? t.blocks.find((b) => b.id === blockId) : undefined;
        if (blockId && !block) return `No block ${blockId}`;
        this.focusTab(t.id);
        if (block) { this.selectedBlock = block.id; this.selectedBlocks = new Set([block.id]); this.emit(); }
        // emit draws synchronously, so the element the host asked for is on the page by now
        this.scrollIntoView(block ? block.id : t.id);
        return `Scrolled to ${block?.title ?? t.name}`;
      });
    add("show_view", "Switch the canvas between the outline, the layouts and the built file.",
      z.object({ view: z.enum(["outline", "layouts", "built"]) }),
      ({ view }) => {
        if (!this.state) return "No plan open";
        if (view === "layouts" && !this.project.tabsApproved) return "Outline not approved";
        if (view === "built" && !this.project.buildPath) return "Nothing built yet";
        const next: View = view === "outline" ? "tabs" : view;
        if (this.state.view !== next) { this.state.view = next; this.emit(); }
        return `Showing ${view}`;
      });
  }

  /** Call one app tool the way the host would, validation and all. The door a test without a host uses. */
  callAppTool(name: string, args: Record<string, unknown> = {}): Promise<{ content?: unknown[]; isError?: boolean }> {
    const tool = this.appTools.get(name);
    if (!tool) return Promise.reject(new Error(`No app tool "${name}"`));
    return tool.handler(args, undefined as never);
  }

  private get rowWord(): string { return (this.project.kind ?? "workbook") === "deck" ? "slide" : "tab"; }
  /** By id first, then by name, so the host may name a slide the way the user reads it. */
  private findTab(tabId?: string, name?: string): Tab | undefined {
    const tabs = this.project.tabs;
    const byId = tabId ? tabs.find((t) => t.id === tabId) : undefined;
    if (byId) return byId;
    const want = (name ?? tabId ?? "").trim().toLowerCase();
    return want ? tabs.find((t) => t.name.trim().toLowerCase() === want) : undefined;
  }
  private focusTab(id: string): void {
    this.selectedTab = id;
    this.selectedBlock = null;
    this.selectedBlocks.clear();
    this.emit();
  }
  /**
   * Outline rows and blocks carry `data-id`, which is also the key FLIP animates on. The grid views name
   * their rows only in the strip, so a tab with no element of its own scrolls the strip's chip into view.
   */
  private scrollIntoView(id: string): void {
    const el = document.querySelector<HTMLElement>(`[data-id="${id.replace(/["\\]/g, "\\$&")}"]`)
      ?? document.querySelector<HTMLElement>(".tabstrip button.on, .slidestrip button.on");
    el?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  private hostKey = "";
  private hostStyles: unknown = null;
  /**
   * The host's own CSS variables and font faces, when it ships them (backlog P9). They are named for
   * Claude Desktop's palette, not ours, so tokens.css keeps working as the fallback and the theme
   * attribute below still decides light from dark. The context is merged, so the object only changes
   * when the host sends a new one: apply it once rather than on every resize frame.
   */
  private applyHostStyles(styles: McpUiHostStyles | undefined): void {
    if (!styles || styles === this.hostStyles) return;
    this.hostStyles = styles;
    if (styles.variables) { try { applyHostStyleVariables(styles.variables); } catch { /* host sent something odd */ } }
    if (styles.css?.fonts) { try { applyHostFonts(styles.css.fonts); } catch { /* injected once, or not at all */ } }
  }

  /** Notifications carry only the changed fields, so always read the merged context. Re-render only when something visible changed, and never mid-drag. */
  private applyHost(_partial: any) {
    const ctx: any = this.app.getHostContext() ?? _partial;
    if (!ctx) return;
    this.applyHostStyles(ctx.styles);
    const theme = ctx.theme === "light" ? "light" : "dark";
    const full = ctx.displayMode === "fullscreen";
    document.documentElement.setAttribute("data-theme", theme);
    document.documentElement.classList.toggle("fullscreen", full);
    this.applyHeight();
    // Height is applied through the CSS variables and needs no re-render; a re-render mid-click or mid-drag loses the interaction.
    const key = `${theme}|${full}|${(ctx.availableDisplayModes ?? []).join()}`;
    if (key === this.hostKey || this.interacting) { this.hostKey = key; return; }
    this.hostKey = key;
    this.emit();
  }

  /**
   * Set the box the host sizes the iframe to, from the policy for the view being drawn. `fit` leaves the
   * height to the content between min and max (autoResize reports what the document measures); the grid
   * views take a fixed height and scroll inside it. Called again on every draw, since the view changes.
   */
  applyHeight(): void {
    const ctx: any = this.app.getHostContext();
    if (ctx?.displayMode === "fullscreen") return; // .fullscreen takes the viewport
    const pol = this.state ? HEIGHT[this.state.view] : HEIGHT.layouts; // before a plan lands, the streaming sketch wants the old fixed box
    const maxH = ctx?.containerDimensions?.maxHeight ?? ctx?.containerDimensions?.height;
    const room = Math.max(pol.min, Math.min(maxH ? maxH - 4 : 640, pol.max));
    const s = document.documentElement.style;
    s.setProperty("--app-height", pol.fit ? "auto" : `${room}px`);
    s.setProperty("--app-min-height", `${pol.min}px`);
    s.setProperty("--app-max-height", pol.fit ? `${room}px` : "none");
  }

  canFullscreen(): boolean { return !!this.app.getHostContext()?.availableDisplayModes?.includes("fullscreen"); }
  isFullscreen(): boolean { return this.app.getHostContext()?.displayMode === "fullscreen"; }
  async toggleFullscreen() {
    const mode = this.isFullscreen() ? "inline" : "fullscreen";
    try { await this.app.requestDisplayMode({ mode }); } catch { this.toast?.("No fullscreen here"); }
  }

  /** Show a built file in the Finder. Hosts that take open-link get the file URL; the rest get the path on the clipboard. */
  async reveal(path: string) {
    if ((this.app.getHostCapabilities() as { openLinks?: unknown } | undefined)?.openLinks) {
      const dir = path.slice(0, path.lastIndexOf("/")) || path;
      const url = `file://${dir.split("/").map(encodeURIComponent).join("/")}`;
      try { const r = await this.app.openLink({ url }); if (!r?.isError) { this.toast?.("Opening folder"); return; } } catch { /* fall through to copy */ }
    }
    this.toast?.((await copyText(path)) ? "Copied" : "Copy failed");
  }

  // ---- persistence ----

  private absorb(sc: any) {
    if (!sc?.project) return;
    const view = this.state!.view;
    this.setState({ ...this.state!, ...sc, view: sc.stage === "built" && view === "built" ? "built" : view, focusTab: undefined });
    this.scheduleContext();
  }

  private async call(name: string, args: Record<string, unknown>): Promise<boolean> {
    try {
      const r = await this.app.callServerTool({ name, arguments: { project: this.project.id, ...args } });
      if (r.isError) { this.toast?.((r.content?.find((c: any) => c.type === "text") as any)?.text ?? "Save failed"); return false; }
      this.absorb(r.structuredContent);
      return true;
    } catch { this.toast?.("Save failed"); return false; }
  }

  saveTabs(tabs: OutlineRow[]) { return this.call("ui_save_tabs", { tabs }); }
  /**
   * Patch one row instead of replacing the outline. ui_save_tabs clears `tabsApproved`, so a story or a
   * headline typed after approval would send the user back to Layer 1; this leaves every approval alone.
   * Debounced per row, and patches stack inside the window so two fields edited together both land.
   * Renames, reorders and add/delete still go through saveTabs: they change the outline itself.
   */
  saveTabFields(tabId: string, fields: TabFields, debounceMs = 250) {
    const t = this.project.tabs.find((x) => x.id === tabId);
    if (t) Object.assign(t, normalizeFields(fields)); // the local copy first, so the field does not flicker
    const patch = { ...this.tabPatches.get(tabId), ...fields };
    this.tabPatches.set(tabId, patch);
    this.schedule(`tab:${tabId}`, debounceMs, () => { this.tabPatches.delete(tabId); return this.call("ui_save_tab_fields", { tabId, ...patch }); });
  }
  /** The row's one-line message: the point of the tab, and the headline Layer 2 draws on the slide. */
  saveMessage(tabId: string, text: string) { this.saveTabFields(tabId, { message: text }); }
  /**
   * Patch the brief card. The patch lands on the local copy first so the fields do not flicker, then the
   * whole card goes out once the typing stops, so two fields edited inside one window both survive.
   */
  saveBrief(patch: Partial<Brief>, debounceMs = 250) {
    const p = this.project;
    p.briefCard = { ...p.briefCard, ...patch };
    this.schedule("brief", debounceMs, () => {
      const b = this.project.briefCard ?? {};
      return this.call("ui_save_brief", { audience: b.audience ?? "", goal: b.goal ?? "", takeaway: b.takeaway ?? "" });
    });
  }
  saveLayout(tab: Tab, blocks: Block[], debounceMs = 250) {
    tab.blocks = blocks; tab.layoutApproved = false; this.emit();
    this.schedule(`layout:${tab.id}`, debounceMs, () => this.call("ui_save_layout", { tab: tab.id, blocks }));
  }

  /** Debounce one save per key. The call is held as well as the timer, so `flushSaves` can make it now. */
  private schedule(key: string, ms: number, run: () => Promise<unknown>) {
    const prev = this.saveTimers.get(key); if (prev) clearTimeout(prev);
    this.pendingSaves.set(key, run);
    this.saveTimers.set(key, window.setTimeout(() => { this.saveTimers.delete(key); this.pendingSaves.delete(key); void run(); }, ms));
  }

  /** Make every pending debounced save now, and resolve when the server has taken them. */
  flushSaves(): Promise<unknown> {
    const runs = [...this.pendingSaves.values()];
    for (const t of this.saveTimers.values()) clearTimeout(t);
    this.saveTimers.clear(); this.pendingSaves.clear();
    return Promise.all(runs.map((run) => { try { return run(); } catch { return undefined; } }));
  }

  /**
   * The host unmounts this canvas as soon as `onteardown` resolves, so an edit still inside its 250 ms
   * window would be dropped: flush first, and cap the wait so a stalled transport cannot hold the
   * unmount open. visibilitychange and pagehide catch the unloads a host tears us down without asking.
   */
  private wireTeardown() {
    this.app.onteardown = async () => {
      await Promise.race([this.flushSaves(), new Promise((done) => window.setTimeout(done, 1200))]);
      return {};
    };
    document.addEventListener("visibilitychange", () => { if (document.hidden) void this.flushSaves(); });
    window.addEventListener("pagehide", () => void this.flushSaves());
  }
  note(action: "add" | "edit" | "delete" | "reopen" | "answer", args: { noteId?: string; tabId?: string; blockId?: string; text?: string }) { return this.call("ui_note", { action, ...args }); }
  async approve(what: "tabs" | "layout" | "all_layouts", tab?: string, approved = true) {
    const ok = await this.call("ui_approve", { what, tab, approved });
    if (!ok || !approved) return;
    const deck = this.project.kind === "deck";
    const row = deck ? "slide" : "tab";
    const build = deck ? "build_deck" : "build_workbook";
    if (what === "tabs") await this.sendMessage(`I approved the ${row} outline on the canvas. Call get_plan, then propose_layouts for every ${row} in one call.${this.project.skipLayouts ? ` skipLayouts is on, so go straight to resolving notes and ${build}.` : ""}`);
    else if (this.project.tabs.every((t) => t.layoutApproved || (t.blocks.length === 0 && /^read\s*me$/i.test(t.name.trim())))) await this.sendMessage(`I approved every ${row} layout on the canvas. Call get_plan, resolve every open note with resolve_notes, then ${build}.`);
  }

  async styles(): Promise<any | null> {
    try { const r = await this.app.callServerTool({ name: "ui_styles", arguments: { project: this.project.id } }); return r.structuredContent ?? null; } catch { return null; }
  }
  setStyle(args: { styleId?: string; templatePath?: string; outputDir?: string; fileName?: string }) { return this.call("ui_set_style", args); }
  async templates(action: "list" | "save" | "apply" | "delete", args: { name?: string; description?: string; template?: string } = {}): Promise<any[] | null> {
    try {
      const r = await this.app.callServerTool({ name: "ui_templates", arguments: { project: this.project.id, action, ...args } });
      const sc = r.structuredContent as any;
      if (action === "apply" && sc?.project) { this.absorb({ ...sc, stage: sc.stage }); this.state!.view = "tabs"; this.selectedTab = this.project.tabs[0]?.id ?? null; this.selectedBlock = null; this.emit(); }
      return sc?.templates ?? null;
    } catch { return null; }
  }

  // ---- model sync ----

  scheduleContext() {
    if (this.ctxTimer) clearTimeout(this.ctxTimer);
    this.ctxTimer = window.setTimeout(() => void this.pushContext(), 800);
  }
  async pushContext() {
    if (!this.state) return;
    // A host that declared its capabilities without this one never takes a context update; do not spend the round trip.
    const caps = this.app.getHostCapabilities();
    if (caps && !caps.updateModelContext) return;
    try { await this.app.updateModelContext({ content: [{ type: "text", text: summarizePlan(this.project, this.state.stage) }] }); } catch { /* transport */ }
  }
  /**
   * The approve → continue handoff. Hosts under-declare `message` (the reference basic-host takes
   * `ui/message` without listing the capability), so an undeclared host is still asked; one refusal is
   * remembered and the user is told to carry on in chat instead.
   */
  private noMessages = false;
  async sendMessage(text: string) {
    if (this.noMessages) { this.toast?.("Tell Claude to continue"); return; }
    try { await this.app.sendMessage({ role: "user", content: [{ type: "text", text }] }); this.toast?.("Sent to Claude"); }
    catch { this.noMessages = true; this.toast?.("Tell Claude to continue"); }
  }

  /** Describe what changed between two plan revisions, for the toast and the flash. */
  private diff(prev: Project, next: Project): { changed: Set<string>; text: string; focusTab?: string } {
    const changed = new Set<string>(); const parts: string[] = []; let focusTab: string | undefined;
    const prevTabs = new Map(prev.tabs.map((t) => [t.id, t])); const nextTabs = new Map(next.tabs.map((t) => [t.id, t]));
    for (const t of next.tabs) if (!prevTabs.has(t.id)) parts.push(`added tab ${t.name}`);
    for (const t of prev.tabs) if (!nextTabs.has(t.id)) parts.push(`removed tab ${t.name}`);
    if (prev.tabs.map((t) => t.id).join() !== next.tabs.map((t) => t.id).join() && next.tabs.length === prev.tabs.length && next.tabs.every((t) => prevTabs.has(t.id))) parts.push("reordered tabs");
    for (const t of next.tabs) {
      const pt = prevTabs.get(t.id); if (!pt) continue;
      if (pt.name !== t.name) parts.push(`renamed ${pt.name} to ${t.name}`);
      const prevBlocks = new Map(pt.blocks.map((b) => [b.id, b]));
      for (const b of t.blocks) {
        const pb = prevBlocks.get(b.id);
        if (!pb) { changed.add(b.id); parts.push(`added ${b.title}`); focusTab ??= t.id; continue; }
        const moved = pb.row !== b.row || pb.col !== b.col; const resized = pb.rows !== b.rows || pb.cols !== b.cols;
        const edited = pb.title !== b.title || pb.type !== b.type || JSON.stringify(pb.columns) !== JSON.stringify(b.columns) || JSON.stringify(pb.rowLabels) !== JSON.stringify(b.rowLabels);
        if (moved || resized || edited) { changed.add(b.id); focusTab ??= t.id; parts.push(moved ? `moved ${b.title} to ${colLetter(b.col)}${b.row}` : resized ? `resized ${b.title}` : `updated ${b.title}`); }
      }
      for (const pb of pt.blocks) if (!t.blocks.some((b) => b.id === pb.id)) parts.push(`removed ${pb.title}`);
    }
    const newNotes = next.notes.filter((n) => !prev.notes.some((m) => m.id === n.id) && n.createdBy === "claude");
    if (newNotes.length) parts.push(`${newNotes.length} new question${newNotes.length === 1 ? "" : "s"} from Claude`);
    const answered = next.notes.filter((n) => n.status !== "open" && prev.notes.some((m) => m.id === n.id && m.status === "open") && n.kind === "instruction");
    if (answered.length) parts.push(`${answered.length} note${answered.length === 1 ? "" : "s"} answered`);
    const text = parts.length ? `Claude: ${parts.slice(0, 2).join(", ")}${parts.length > 2 ? ` +${parts.length - 2}` : ""}` : "";
    return { changed, text, focusTab };
  }

  // ---- polling for chat edits and newer canvases ----

  private startPolling() {
    if (this.pollTimer) return;
    this.pollTimer = window.setInterval(async () => {
      if (!this.state || this.interacting || document.hidden) return;
      try {
        const want = this.wantsSnapshot(); // ordinary polls stay lean; only a preview-less built view asks
        if (want) this.snapAskedAt = this.project.revision;
        const r = await this.app.callServerTool({ name: "ui_get_state", arguments: { project: this.project.id, ...(want ? { withSnapshot: true } : {}) } });
        const sc = r.structuredContent as any;
        if (!sc?.project) return;
        const newer = (sc.project.renderSeq ?? 0) > this.renderId;
        if (newer !== this.superseded && !this.pinned) { this.superseded = newer; this.emit(); }
        if (sc.project.revision > this.project.revision && !this.saveTimers.size) {
          const view: View = this.state.view === "built" && sc.stage !== "built" ? "layouts" : sc.stage === "built" ? "built" : this.state.view;
          const summary = this.diff(this.project, sc.project);
          if (summary.changed.size) { this.changed = summary.changed; if (summary.focusTab) this.selectedTab = summary.focusTab; }
          this.setState({ ...this.state, ...sc, view, focusTab: undefined });
          if (summary.text) this.toast?.(summary.text);
        } else if (want && this.takeSnapshot(sc)) this.emit(); // nothing changed, but the preview came back
      } catch { /* ignore */ }
    }, 2500);
  }
}

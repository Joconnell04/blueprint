import { h, clear, popover, icon, twoStep } from "./dom.js";
import { fadeIn } from "./motion.js";
import type { Controller } from "./controller.js";
import type { Kind, Project } from "./model.js";
import { BUILD_DEFAULTS, type BuildOptions, type Fidelity } from "../shared/types.js";

type StyleKind = "workbook" | "deck" | "both";

/** The three fidelities in order, with the one line the sheet says about each. No help paragraphs. */
const FIDELITIES: [Fidelity, string, string][] = [
  ["strict", "Strict", "Draws the plan exactly."],
  ["balanced", "Balanced", "Resizes and balances within two cells."],
  ["loose", "Loose", "May re-pick the layout."],
];
/** The three generated extras, label first: what the builder adds when the plan has none. */
const BUILD_TOGGLES: [keyof Omit<BuildOptions, "fidelity">, string][] = [
  ["notesFromStory", "Notes from story"], ["cover", "Cover"], ["agenda", "Agenda"],
];

/** What the builder may change, with every default filled in. */
function buildOptions(p: Project): Required<BuildOptions> {
  return { ...BUILD_DEFAULTS, ...(p.build ?? {}) };
}

/**
 * Save one build option, the moment it is picked. The local copy is patched first so the control never
 * waits for the round trip, and what the server stored is taken back onto it, revision included: without
 * that the next poll would see a newer plan, redraw the canvas, and take this open sheet with it.
 */
async function saveBuild(c: Controller, patch: BuildOptions): Promise<void> {
  const p = c.project;
  p.build = { ...(p.build ?? {}), ...patch };
  try {
    const r = await c.app.callServerTool({ name: "ui_set_build_options", arguments: { project: p.id, ...patch } });
    if (!r.isError) {
      const stored = (r.structuredContent as { project?: Project } | undefined)?.project;
      if (stored) { p.build = stored.build; p.revision = stored.revision; p.updatedAt = stored.updatedAt; }
      return;
    }
    const blocks = (r.content ?? []) as unknown as { type?: string; text?: string }[];
    c.toast?.(briefly(blocks.filter((b) => b.type === "text").map((b) => b.text ?? "").join("\n")));
  } catch { c.toast?.("Save failed"); }
}

/** One entry of ui_styles.styles. */
interface StyleEntry {
  id: string;
  source: "plain" | "house" | "skill" | "template";
  name: string;
  description: string;
  kind: StyleKind;
  skillPath?: string;
  tokensPath?: string;
  builderMode: string;
}
interface StylesData { styles: StyleEntry[]; current: string; templatePath?: string; outputDir?: string; fileName?: string }

/** One field extract_format recovered: `key` is a section ("deck") or a field under it ("deck.palette.primary"). */
interface Field { key: string; value?: unknown; confidence?: number; evidence?: string }
/** One thing the files cannot answer: the question for the row, the whole of it for the tooltip. */
interface Ask { question: string; detail: string }

/** The half of extract_format's structuredContent this popover reads. It arrives pre-ranked. */
interface Extracted {
  saved?: boolean;
  name?: string;
  needsObservations?: boolean;
  exists?: boolean;      // a skill of this name is already on disk, so saving replaces it
  skillPath?: string;
  style?: StyleEntry;
  fields?: Field[];
  unknowns?: Ask[];
}

const MAX_FIELDS = 5;    // recovered fields listed per section; the rest are counted
const MAX_UNKNOWNS = 3;  // questions listed; Claude gets the whole list when it calls the tool itself
const DROP_HINT = "Drop files here";
const NO_PATH_HINT = "Attach in chat, then paste the path";

/** extract_format is registered for the model and the app, and takes no project id, so it goes out raw. */
async function extract(c: Controller, args: Record<string, unknown>): Promise<{ ok: boolean; text: string; sc: Extracted }> {
  try {
    const r = await c.app.callServerTool({ name: "extract_format", arguments: args });
    const blocks = (r.content ?? []) as unknown as { type?: string; text?: string }[];
    const text = blocks.filter((b) => b.type === "text").map((b) => b.text ?? "").join("\n");
    return { ok: !r.isError, text, sc: (r.structuredContent ?? {}) as Extracted };
  } catch (e) {
    return { ok: false, text: (e as Error)?.message || "The tool did not answer", sc: {} };
  }
}

function lines(v: string): string[] { return v.split("\n").map((s) => s.trim()).filter(Boolean); }
function firstLine(t: string): string { return (t.split("\n").find((l) => l.trim()) ?? "That did not work").trim(); }
/** A toast is six words at most: the first sentence of a tool error, not the whole refusal. */
function briefly(t: string): string {
  const s = firstLine(t).split(/(?<=[.!?])\s/)[0];
  return s.length > 48 ? `${s.slice(0, 47)}…` : s;
}
function pct(v: number | undefined): string { return v === undefined ? "?" : `${Math.round(v * 100)}%`; }
function dotFor(v: number | undefined): string { return v === undefined ? "none" : v >= 0.8 ? "ok" : v >= 0.5 ? "warn" : "bad"; }
function short(v: unknown): string {
  const s = Array.isArray(v) ? v.join(", ") : String(v);
  return s.length > 28 ? `${s.slice(0, 27)}…` : s;
}
function title(s: string): string { return s.charAt(0).toUpperCase() + s.slice(1); }

/**
 * Files dropped on the paths field. A sandboxed iframe gets no `File.path`, so the only usable drop is
 * one that carries the path as text (Finder sends `text/uri-list`); anything else leaves the hint to say
 * where the path comes from instead.
 */
function droppedPaths(e: DragEvent): string[] {
  const dt = e.dataTransfer;
  if (!dt) return [];
  const out: string[] = [];
  for (const f of Array.from(dt.files)) {
    const p = (f as File & { path?: string }).path;
    if (p) out.push(p);
  }
  if (out.length) return out;
  for (const raw of (dt.getData("text/uri-list") || dt.getData("text/plain") || "").split(/[\r\n]+/)) {
    const t = raw.trim();
    if (!t || t.startsWith("#")) continue;
    if (/^file:\/\//i.test(t)) { try { out.push(decodeURIComponent(new URL(t).pathname)); } catch { /* not a URL after all */ } }
    else if (t.startsWith("/") || t.startsWith("~")) out.push(t);
  }
  return out;
}

/** A button that shows what it is doing and comes back to its own label. */
async function busy(btn: HTMLButtonElement, label: string, fn: () => Promise<void>): Promise<void> {
  const was = btn.textContent ?? "";
  btn.disabled = true;
  btn.textContent = label;
  try { await fn(); } finally {
    if (btn.isConnected) { btn.disabled = false; btn.textContent = was; }
  }
}

/**
 * Header popover: which format the build follows (built-in, a skill on this Mac, or a template file),
 * where the file goes, and a new style extracted from example files. The picker lists only the styles
 * that can dress this project's kind.
 */
export function openStylePopover(c: Controller, anchor: HTMLElement) {
  popover(anchor, (close) => {
    const box = h("div", { class: "stylepop stylefmt" }, h("div", { class: "muted small-text pophdr" }, "Loading…"));
    void c.styles().then((d) => {
      const data = d as StylesData | null;
      clear(box);
      if (!data?.styles) { box.append(h("div", { class: "muted small-text pophdr" }, "No styles")); return; }
      fill(c, box, data, close);
      fadeIn(box);
    });
    return box;
  }, { align: "auto" });
}

function fill(c: Controller, box: HTMLElement, data: StylesData, close: () => void): void {
  const kind: Kind = c.project.kind ?? "workbook";
  const deckMode = kind === "deck";
  let current = data.current;

  // ---- picker: the styles that fit this kind, one line each
  const list = h("div", { class: "stylelist" });
  const tplField = h("div", { class: "field tplf", hidden: !current.startsWith("template") });
  const tpl = h("input", { value: data.templatePath ?? "", placeholder: deckMode ? "/path/to/deck.pptx" : "/path/to/workbook.xlsx", spellcheck: "false" });
  tplField.append(h("span", { class: "label" }, "Template"), tpl);
  const fits = (s: StyleEntry) => s.kind === kind || s.kind === "both";

  const drawList = () => {
    clear(list);
    for (const s of data.styles.filter(fits)) {
      const on = s.id === current;
      const deckSkill = deckMode && s.source === "skill";
      const buildable = deckSkill && !!s.tokensPath;   // a deck skill the builder can dress a slide from
      const proseOnly = deckSkill && !s.tokensPath;    // prose the builder cannot read: Claude has to derive tokens
      const pick = h("button", {
        class: `styleopt${on ? " on" : ""}`,
        onclick: () => { current = s.id; tplField.hidden = s.source !== "template"; drawList(); },
      },
        h("span", { class: "radio" }, on ? icon("check", 10) : null),
        h("span", { class: "sname" }, h("span", { class: "strunc" }, s.name), buildable ? h("span", { class: "dot ok buildable", title: "Tokens ready" }) : null),
        h("span", { class: "sdesc" },
          s.source === "skill" ? h("span", { class: "chip line" }, "skill") : null,
          proseOnly ? h("span", { class: "chip line tokmiss" }, "tokens missing") : null,
          h("span", { class: "strunc" }, s.description)));
      list.append(h("div", { class: "srow", "data-style": s.id }, pick, proseOnly ? deriveButton(c, s) : null));
    }
  };
  drawList();

  // ---- where the file goes: one row saying the file name, and the two fields behind it
  const out = h("input", { value: data.outputDir ?? "", placeholder: "~/Documents", spellcheck: "false" });
  const fname = h("input", { value: data.fileName ?? "", placeholder: c.project.name, spellcheck: "false" });
  const ext = deckMode ? ".pptx" : ".xlsx";   // the builder names the file from this field and its kind
  const outNow = h("span", { class: "onow trunc" });
  const sayName = () => { outNow.textContent = `${fname.value.trim() || c.project.name}${ext}`; };
  sayName();
  fname.oninput = sayName;
  const outFields = h("div", { class: "split outf" },
    h("div", { class: "field" }, h("span", { class: "label" }, "Folder"), out),
    h("div", { class: "field" }, h("span", { class: "label" }, "File name"), fname));

  // ---- what the builder may change between the approved plan and the file (deck only)
  const opts = buildOptions(c.project);
  const fidSeg = h("div", { class: "seg fidseg" });
  const fidLine = h("div", { class: "muted xs fidline" });
  const drawFid = () => {
    clear(fidSeg);
    for (const [value, label] of FIDELITIES) {
      fidSeg.append(h("button", { class: value === opts.fidelity ? "on" : "", "data-fid": value, onclick: () => {
        if (opts.fidelity === value) return;
        opts.fidelity = value;
        drawFid();
        void saveBuild(c, { fidelity: value });
      } }, label));
    }
    fidLine.textContent = FIDELITIES.find(([v]) => v === opts.fidelity)?.[2] ?? "";
  };
  drawFid();
  const toggles = h("div", { class: "inline tight wrap bopts" });
  const drawToggles = () => {
    clear(toggles);
    for (const [key, label] of BUILD_TOGGLES) {
      const on = opts[key];
      toggles.append(h("button", { class: `small ghost bopt${on ? " on" : ""}`, "data-opt": key, "aria-pressed": String(on), onclick: () => {
        opts[key] = !on;
        drawToggles();
        void saveBuild(c, { [key]: !on });
      } }, on ? icon("check", 10) : null, label));
    }
  };
  drawToggles();
  const buildFields = h("div", { class: "buildf" },
    h("div", { class: "field row2 fidf" }, h("span", { class: "label" }, "Fidelity"), h("span", { class: "spacer" }), fidSeg),
    fidLine,
    h("div", { class: "field boptf" }, h("span", { class: "label" }, "Include"), toggles));

  let outOpen = false;
  const outRow: HTMLButtonElement = h("button", { class: "ghost small outrow", onclick: () => setOut(!outOpen) },
    h("span", { class: "label" }, "Output"), outNow, icon("chev", 10));
  /** Folder and File name are almost never the question, so the sheet opens with them away. */
  function setOut(open: boolean): void {
    outOpen = open;
    outFields.hidden = !open;
    buildFields.hidden = !open || !deckMode;   // a workbook has no fidelity and nothing generated
    outRow.classList.toggle("open", open);
    outRow.setAttribute("aria-expanded", String(open));
    if (open) { fadeIn(outFields); if (deckMode) fadeIn(buildFields); out.focus(); }
  }
  setOut(false);   // collapsed on every open; it stays as the user left it while the sheet lives

  // ---- new style from example files
  const nname = h("input", { class: "nname", placeholder: "Name", spellcheck: "false" });
  const npaths = h("textarea", { class: "npaths", placeholder: "/path/to/deck.pptx", spellcheck: "false", rows: "2" });
  let nkind: StyleKind = kind;
  const seg = h("div", { class: "seg kindseg" });
  const drawSeg = () => {
    clear(seg);
    for (const k of ["workbook", "deck", "both"] as StyleKind[]) {
      seg.append(h("button", { class: k === nkind ? "on" : "", onclick: () => { nkind = k; drawSeg(); } }, k === "workbook" ? "Workbook" : k === "deck" ? "Deck" : "Both"));
    }
  };
  drawSeg();

  const hint = h("div", { class: "drophint muted xs" }, DROP_HINT);
  const zone = h("div", { class: "dropz" }, npaths, hint);
  zone.ondragover = (e: DragEvent) => { e.preventDefault(); zone.classList.add("over"); };
  zone.ondragleave = () => zone.classList.remove("over");
  zone.ondrop = (e: DragEvent) => {
    e.preventDefault();
    zone.classList.remove("over");
    const dropped = droppedPaths(e);
    if (!dropped.length) { hint.textContent = NO_PATH_HINT; return; }  // the bytes are no use to us; the path is
    npaths.value = [...lines(npaths.value), ...dropped].join("\n");
    hint.textContent = DROP_HINT;
  };

  const flash = h("div", { class: "card xerr", hidden: true });
  const prev = h("div", { class: "xprev", hidden: true });
  let preview: Extracted | null = null;   // the reading on screen, if the user has extracted one

  const problem = (text: string) => {
    clear(flash);
    flash.hidden = false;
    flash.append(h("div", { class: "ct" }, firstLine(text)));
    fadeIn(flash);
  };

  /**
   * What the files gave up, straight off `fields`: the tool ranks them and heads each section with its
   * own confidence, so the rows here are the list in order. Then the questions only Claude can carry.
   */
  const showPreview = (sc: Extracted, name: string) => {
    clear(prev);
    prev.hidden = false;
    const rows: HTMLElement[] = [];
    let shown = 0, more = 0;
    for (const f of sc.fields ?? []) {
      const dot = f.key.indexOf(".");
      if (dot < 0) {  // a section: its label and the mean confidence over the fields under it
        shown = 0;
        rows.push(h("div", { class: "xsec" }, h("span", { class: "label" }, title(f.key)), h("span", { class: "chip line num" }, pct(f.confidence))));
        continue;
      }
      if (shown++ >= MAX_FIELDS) { more++; continue; }
      const path = f.key.slice(dot + 1);
      rows.push(h("div", { class: "xrow" },
        h("span", { class: `dot ${dotFor(f.confidence)}`, title: pct(f.confidence) }),
        h("span", { class: "xk trunc", title: f.evidence || path }, path),
        h("span", { class: "xv mono trunc" }, short(f.value))));
    }
    prev.append(h("div", { class: "label pophdr" }, "Found"));
    if (rows.length) prev.append(...rows); else prev.append(h("div", { class: "muted small-text pophdr" }, "Nothing recovered"));
    if (more) prev.append(h("div", { class: "muted xs xmore" }, `+${more} more`));

    const unknowns = (sc.unknowns ?? []).slice(0, MAX_UNKNOWNS);
    if (unknowns.length) {
      prev.append(h("div", { class: "label pophdr sectop" }, "Ask Claude"),
        h("ul", { class: "warnlist xask" }, ...unknowns.map((u) => h("li", { title: u.detail }, h("span", { class: "dot warn" }), h("span", {}, u.question)))));
    }

    // Saving is the footer's job; a screenshot Claude has to read is the one action that belongs here.
    if (sc.needsObservations) {
      const read: HTMLButtonElement = h("button", { class: "small", onclick: () => void busy(read, "Sending", () => c.sendMessage(`Read the attached screenshots and call extract_format again with observations for ${name}, save: true`)) }, "Read screenshots");
      prev.append(h("div", { class: "inline end tight xacts" }, read));
    }
    fadeIn(prev);
  };

  const discard = () => { preview = null; clear(prev); prev.hidden = true; flash.hidden = true; drawFoot(); };

  /** The saved skill lists like any other style, so refresh the picker and select it. */
  const selectSaved = async (sc: Extracted) => {
    const id = sc.style?.id ?? (sc.name ? `skill:${sc.name}` : "");
    const fresh = (await c.styles()) as StylesData | null;
    if (fresh?.styles) data.styles = fresh.styles;
    const entry = data.styles.find((s) => s.id === id);
    discard();
    if (!entry || !fits(entry)) { drawList(); c.toast?.("Style added"); return; } // it dresses the other builder
    current = id;
    tplField.hidden = true;
    drawList();
    // Persists it with whatever else the sheet holds, and re-renders the canvas over this popover.
    await c.setStyle({ styleId: id, templatePath: tpl.value, outputDir: out.value, fileName: fname.value });
    c.toast?.("Style added");
  };

  const run = async (write: boolean, overwrite = false): Promise<void> => {
    const name = nname.value.trim();
    if (!name) { nname.focus(); c.toast?.("Name the style"); return; }
    const paths = lines(npaths.value);
    if (!paths.length) { npaths.focus(); c.toast?.("Add a file path"); return; }
    flash.hidden = true;
    const r = await extract(c, { paths, name, kind: nkind, save: write, ...(overwrite ? { overwrite: true } : {}) });
    if (!r.ok) { problem(r.text); drawFoot(); return; }
    if (write) { await selectSaved(r.sc); return; }
    preview = r.sc;
    showPreview(r.sc, name);
    drawFoot();
  };

  const go: HTMLButtonElement = h("button", { class: "small", onclick: () => void busy(go, "Reading", () => run(false)) }, "Extract");

  /**
   * One footer for the sheet, one primary in it: the picked style until a reading is on screen, and
   * then that reading. The preview already says whether the skill is on disk, so Replace is offered
   * up front rather than after a refused save.
   */
  const foot = h("div", { class: "inline end tight popfoot" });
  function drawFoot(): void {
    clear(foot);
    if (!preview) {
      const apply: HTMLButtonElement = h("button", {
        class: "primary small",
        onclick: () => void busy(apply, "Saving", async () => {
          await c.setStyle({ styleId: current, templatePath: tpl.value, outputDir: out.value, fileName: fname.value });
          close();
          c.toast?.("Style saved");
        }),
      }, "Save");
      foot.append(h("button", { class: "small ghost", onclick: close }, "Cancel"), apply);
      return;
    }
    const keep: HTMLButtonElement = preview.exists
      ? twoStep("Replace style", "Replace?", () => void busy(keep, "Saving", () => run(true, true)), "primary small")
      : h("button", { class: "primary small", onclick: () => void busy(keep, "Saving", () => run(true)) }, "Save style");
    foot.append(h("button", { class: "small ghost", onclick: discard }, "Cancel"), keep);
  }
  drawFoot();

  // The body scrolls, the footer does not: a long preview never pushes the primary action out of the sheet.
  box.append(
    h("div", { class: "stylebody" },
      h("div", { class: "label pophdr" }, "Style"),
      list,
      tplField,
      outRow,
      outFields,
      buildFields,
      h("div", { class: "label pophdr sectop" }, "New style"),
      h("div", { class: "stylenew" },
        h("div", { class: "split" },
          h("div", { class: "field" }, h("span", { class: "label" }, "Name"), nname),
          h("div", { class: "field kindf" }, h("span", { class: "label" }, "Kind"), seg)),
        h("div", { class: "field" }, h("span", { class: "label" }, "Files"), zone),
        h("div", { class: "inline end tight" }, go),
        flash,
        prev)),
    foot);
}

/** Tokens the builder can read have to come from the prose, and only Claude can read prose. */
function deriveButton(c: Controller, s: StyleEntry): HTMLButtonElement {
  const btn: HTMLButtonElement = h("button", {
    class: "small ghost derive",
    title: "Derive tokens from this skill",
    onclick: () => void busy(btn, "Asking", async () => {
      const r = await extract(c, { fromSkill: s.name, kind: "deck" });
      if (!r.ok) { c.toast?.(briefly(r.text)); return; }
      await c.sendMessage(`Derive deck tokens for ${s.name} from its skill text and save them`);
    }),
  }, "Derive");
  return btn;
}

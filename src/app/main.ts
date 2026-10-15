import { Controller } from "./controller.js";
import { h, clear, icon } from "./dom.js";
import { flip, toast as showToast } from "./motion.js";
import { vocab } from "./model.js";
import { renderOutlineView } from "./outline-view.js";
import { renderLayoutView } from "./layout-view.js";
import { renderDeckLayoutView } from "./deck-layout.js";
import { renderBuiltView } from "./built-view.js";
import { renderStreaming } from "./streaming-view.js";
import { openStylePopover } from "./style-popover.js";
import { openTemplatesPopover } from "./templates-popover.js";

declare global { interface Window { blueprint?: Controller } }

const root = document.getElementById("root")!;
const c = new Controller();
// The host drives this canvas through the app tools in controller.ts; this is the same registry for a
// test or a debug console inside Claude Desktop, where nothing else can reach into the iframe.
window.blueprint = c;
let toastTimer = 0;
c.toast = (msg) => {
  const t = showToast(root, msg);
  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(t.close, msg.startsWith("Claude:") ? 5000 : 2600);
};

function styleLabel(p: { styleId?: string; formatting: string }): string {
  const id = p.styleId ?? p.formatting;
  if (id.startsWith("skill:")) return id.slice(6);
  return id === "house" ? "house format" : id === "template" ? "template" : "plain";
}

/** Every re-render runs through FLIP, so a row or block that moved slides to its new place. */
function render() { flip(root, draw); }

/** Where every scrolling box was, keyed by its `data-scroll` id. */
type Scrolls = Map<string, [number, number]>;
/** The focused field, keyed by its `data-focus` id, with the caret or selection it held. */
interface Focused { key: string; start: number | null; end: number | null }

function takeScrolls(): Scrolls {
  const at: Scrolls = new Map();
  for (const el of root.querySelectorAll<HTMLElement>("[data-scroll]")) at.set(el.dataset.scroll!, [el.scrollLeft, el.scrollTop]);
  return at;
}
function takeFocus(): Focused | null {
  const host = (document.activeElement as HTMLElement | null)?.closest<HTMLElement>("[data-focus]");
  const key = host?.dataset.focus;
  if (!key) return null;
  const f = host as HTMLInputElement | HTMLTextAreaElement;
  try { return { key, start: f.selectionStart ?? null, end: f.selectionEnd ?? null }; } catch { return { key, start: null, end: null }; }
}
/**
 * A draw replaces the whole canvas, so the caret and every scroll offset would be lost with it: ⌘↩ on a
 * note box sent the outline back to the top (F8). Put the caret back first, then the offsets, because
 * focusing a field scrolls its box.
 */
function restore(at: Scrolls, was: Focused | null, pending: string | null): void {
  const key = pending ?? was?.key;
  const el = key ? root.querySelector<HTMLElement>(`[data-focus="${key.replace(/["\\]/g, "\\$&")}"]`) : null;
  if (el) {
    el.focus({ preventScroll: true });
    const f = el as HTMLInputElement | HTMLTextAreaElement;
    try { if (pending) f.select?.(); else if (was?.start != null) f.setSelectionRange(was.start, was.end); } catch { /* not a text field */ }
  }
  for (const box of root.querySelectorAll<HTMLElement>("[data-scroll]")) {
    const to = at.get(box.dataset.scroll!);
    if (to) { box.scrollLeft = to[0]; box.scrollTop = to[1]; }
  }
}
/**
 * The scrolling boxes the other views own, tagged from here until their own files carry the attribute.
 * One id per box, stable across draws because the markup order is.
 */
function tagScrolls(): void {
  const tag = (sel: string, id: string) => root.querySelectorAll<HTMLElement>(sel).forEach((el, i) => { if (!el.dataset.scroll) el.dataset.scroll = i ? `${id}${i}` : id; });
  tag(".tabstrip", "strip");
  tag(".slidestrip", "slidestrip");
  tag(".canvaswrap", "canvas");
  tag(".side .scroll", "side");
}

function draw() {
  const at = takeScrolls();
  const was = takeFocus();
  const pending = c.pendingFocus;
  c.pendingFocus = null;
  paint();
  tagScrolls();
  restore(at, was, pending);
}

function paint() {
  clear(root);
  document.onkeydown = null;
  c.applyHeight(); // the outline sizes to content, the grid views take a fixed box
  const s = c.state;
  if (!s) {
    root.classList.remove("collapsed");
    if (c.error) { root.classList.add("collapsed"); root.append(h("div", { class: "superseded" }, h("span", { class: "dot bad" }), h("span", { class: "trunc" }, c.error.split("\n")[0]), h("span", { class: "spacer" }), h("span", { class: "muted" }, "See the tool result"))); return; }
    if (c.partial) renderStreaming(c.partial, root); else root.append(h("div", { class: "empty" }, c.cancelled ? "Cancelled" : "Waiting for the plan"));
    return;
  }
  const p = s.project;
  if (c.superseded && !c.pinned) {
    root.classList.add("collapsed");
    root.append(h("div", { class: "superseded" },
      h("span", { class: "logo" }, icon("logo", 12)),
      h("span", { class: "trunc" }, p.name),
      h("span", { class: "muted" }, "Newer canvas below"),
      h("span", { class: "spacer" }),
      h("button", { class: "small ghost", onclick: () => { c.pinned = true; c.superseded = false; render(); } }, "Keep this one")));
    return;
  }
  root.classList.remove("collapsed");
  const v = vocab(p.kind ?? "workbook");
  const stageIdx = s.view === "tabs" ? 0 : s.view === "layouts" ? 1 : 2;
  const steps = [v.many, "Layouts", "Build"].map((label, i) => h("button", { class: i === stageIdx ? "on" : i < stageIdx ? "done" : "", disabled: (i === 1 && !p.tabsApproved) || (i === 2 && !p.buildPath), onclick: () => { s.view = i === 0 ? "tabs" : i === 1 ? "layouts" : "built"; c.setState(s); } }, label));
  root.append(h("div", { class: "hdr" },
    h("div", { class: "logo" }, icon("logo", 12)),
    h("span", { class: "name", title: p.brief }, p.name),
    h("span", { class: "chip line kind", title: `${p.kind ?? "workbook"} plan` }, p.kind ?? "workbook"),
    p.importedFrom ? h("span", { class: "chip line", title: p.importedFrom }, "imported") : null,
    h("button", { class: "chip line stylebtn", title: "Style", onclick: (e: Event) => openStylePopover(c, e.currentTarget as HTMLElement) }, styleLabel(p), icon("chev", 9)),
    h("span", { class: "spacer" }),
    h("button", { class: "ghost small", title: "Templates", onclick: (e: Event) => openTemplatesPopover(c, e.currentTarget as HTMLElement) }, icon("template", 12), "Templates"),
    h("div", { class: "seg" }, ...steps),
    c.canFullscreen() ? h("button", { class: "ghost icon", title: c.isFullscreen() ? "Exit fullscreen" : "Fullscreen", onclick: () => void c.toggleFullscreen() }, icon(c.isFullscreen() ? "shrink" : "expand", 13)) : null));
  if (s.view === "tabs") renderOutlineView(c, root);
  else if (s.view === "layouts") (p.kind === "deck" ? renderDeckLayoutView : renderLayoutView)(c, root);
  else renderBuiltView(c, root);
}

c.onChange(render);
render();
c.connect().catch((e) => { root.append(h("div", { class: "empty" }, `No host connection: ${e?.message ?? e}`)); });

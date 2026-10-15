import { fadeIn } from "./motion.js";

export function h<K extends keyof HTMLElementTagNameMap>(tag: K, attrs: Record<string, any> = {}, ...children: (Node | string | null | undefined | false)[]): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v == null || v === false) continue;
    if (k === "class") el.className = v;
    else if (k === "style" && typeof v === "object") Object.assign(el.style, v);
    else if (k.startsWith("on") && typeof v === "function") (el as any)[k.toLowerCase()] = v;
    else if (k === "value") (el as any).value = v;
    else if (k === "checked" || k === "disabled" || k === "selected") (el as any)[k] = !!v;
    else el.setAttribute(k, String(v));
  }
  for (const c of children) if (c != null && c !== false) el.append(c instanceof Node ? c : document.createTextNode(String(c)));
  return el;
}
export function clear(el: Element): void { while (el.firstChild) el.removeChild(el.firstChild); }

const ICONS: Record<string, string> = {
  grip: '<circle cx="5" cy="3" r="1.4"/><circle cx="11" cy="3" r="1.4"/><circle cx="5" cy="8" r="1.4"/><circle cx="11" cy="8" r="1.4"/><circle cx="5" cy="13" r="1.4"/><circle cx="11" cy="13" r="1.4"/>',
  up: '<path d="M8 13V3M4 7l4-4 4 4"/>', down: '<path d="M8 3v10M4 9l4 4 4-4"/>',
  trash: '<path d="M3 4h10M6 4V2h4v2M5 4l1 10h4l1-10"/>', plus: '<path d="M8 3v10M3 8h10"/>', x: '<path d="M4 4l8 8M12 4l-8 8"/>',
  stack: '<rect x="3" y="2" width="10" height="4" rx="1"/><rect x="3" y="10" width="10" height="4" rx="1"/>',
  flow: '<rect x="2" y="3" width="4" height="10" rx="1"/><rect x="10" y="3" width="4" height="10" rx="1"/>',
  expand: '<path d="M9 2h5v5M14 2 9 7M7 14H2V9M2 14l5-5"/>', shrink: '<path d="M14 7H9V2M9 7l5-5M2 9h5v5M7 9l-5 5"/>',
  undo: '<path d="M6 4L2.5 7.5 6 11M3 7.5h6.5a3.5 3.5 0 0 1 0 7H8"/>', redo: '<path d="M10 4l3.5 3.5L10 11M13 7.5H6.5a3.5 3.5 0 0 0 0 7H8"/>',
  template: '<rect x="2" y="2" width="12" height="12" rx="1.5"/><path d="M2 6h12M6 6v8M6 10h8"/>',
  chev: '<path d="M4 6l4 4 4-4"/>',
  logo: '<rect x="2" y="2" width="12" height="12" rx="1.5"/><path d="M2 6h12M6 6v8"/>', check: '<path d="M3 8.5l3 3 7-7"/>', copy: '<rect x="5" y="5" width="9" height="9" rx="1.5"/><path d="M11 5V3.5A1.5 1.5 0 0 0 9.5 2h-6A1.5 1.5 0 0 0 2 3.5v6A1.5 1.5 0 0 0 3.5 11H5"/>',
  reveal: '<path d="M2 4.5A1.5 1.5 0 0 1 3.5 3h3l1.5 2h4.5A1.5 1.5 0 0 1 14 6.5v5A1.5 1.5 0 0 1 12.5 13h-9A1.5 1.5 0 0 1 2 11.5z"/>',
  arrows: '<path d="M3 12V6a2 2 0 0 1 2-2h6"/><path d="M8.5 1.5 11.5 4 8.5 6.5"/>',
};
export function icon(name: string, size = 14, fill = false): SVGElement {
  const el = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  el.setAttribute("width", String(size)); el.setAttribute("height", String(size)); el.setAttribute("viewBox", "0 0 16 16");
  if (fill) el.setAttribute("fill", "currentColor"); else { el.setAttribute("fill", "none"); el.setAttribute("stroke", "currentColor"); el.setAttribute("stroke-width", "1.6"); el.setAttribute("stroke-linecap", "round"); el.setAttribute("stroke-linejoin", "round"); }
  el.innerHTML = ICONS[name] ?? ""; return el;
}
export function svg(inner: string, viewBox = "0 0 100 60"): SVGElement {
  const el = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  el.setAttribute("viewBox", viewBox); el.setAttribute("preserveAspectRatio", "none"); el.classList.add("chart");
  el.innerHTML = inner; return el;
}
/** A button that asks for a second click before doing something destructive. */
export function twoStep(label: Node | string, confirmLabel: string, onConfirm: () => void, cls = "ghost icon"): HTMLButtonElement {
  let armed = false; let timer = 0;
  const btn = h("button", { class: cls, title: typeof label === "string" ? label : undefined }, label);
  const reset = () => { armed = false; btn.classList.remove("confirm"); clear(btn); btn.append(typeof label === "string" ? label : label.cloneNode(true)); };
  btn.onclick = (e) => {
    e.stopPropagation();
    if (!armed) { armed = true; btn.classList.add("confirm"); clear(btn); btn.append(confirmLabel); timer = window.setTimeout(reset, 3000); return; }
    clearTimeout(timer); onConfirm();
  };
  return btn;
}
/** Popover anchored under an element; closes on outside click or Escape. */
export function popover(anchor: HTMLElement, build: (close: () => void) => HTMLElement, opts: { align?: "left" | "right" | "auto" } = {}): void {
  document.querySelectorAll(".pop").forEach((p) => p.remove());
  const root = document.getElementById("root")!;
  const rr = root.getBoundingClientRect(); const ar = anchor.getBoundingClientRect();
  const close = () => { pop.remove(); document.removeEventListener("pointerdown", onDoc, true); document.removeEventListener("keydown", onKey, true); };
  const pop = build(close);
  pop.classList.add("pop");
  // Align to whichever side of the anchor has room; never let the popover leave the canvas.
  const align = opts.align && opts.align !== "auto" ? opts.align : (ar.left + ar.width / 2 - rr.left) > rr.width / 2 ? "right" : "left";
  if (align === "right") pop.style.right = `${Math.max(4, rr.right - ar.right)}px`; else pop.style.left = `${Math.max(4, ar.left - rr.left)}px`;
  root.append(pop);
  const pw = pop.getBoundingClientRect().width;
  if (align === "left" && ar.left - rr.left + pw > rr.width - 4) { pop.style.left = `${Math.max(4, rr.width - pw - 4)}px`; }
  if (align === "right" && rr.right - ar.right + pw > rr.width - 4) { pop.style.right = "4px"; }
  // Open below the anchor, or above it when there is no room below.
  const ph = pop.getBoundingClientRect().height;
  const below = ar.bottom - rr.top + 4;
  pop.style.top = below + ph > rr.height - 4 && ar.top - rr.top - ph - 4 > 0 ? `${ar.top - rr.top - ph - 4}px` : `${below}px`;
  const onDoc = (e: PointerEvent) => { if (!pop.contains(e.target as Node) && e.target !== anchor) close(); };
  const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") close(); };
  setTimeout(() => { document.addEventListener("pointerdown", onDoc, true); document.addEventListener("keydown", onKey, true); }, 0);
  fadeIn(pop);
  const first = pop.querySelector<HTMLElement>("input, button"); first?.focus();
}

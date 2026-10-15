/**
 * The canvas motion language. Four moves, all quiet, all driven by the duration and easing tokens in
 * tokens.css, all instant when the user asks for reduced motion. Views never write their own transitions.
 *
 * - flip(container, mutate): list reorders and block geometry changes (Stack, Flow, align, undo, drag drop)
 * - fadeIn(el): something appears in place (popover, panel content, a new row)
 * - slideIn(el): a panel or bar arrives from an edge
 * - pulse(el): a block or cell Claude changed
 * - toast(root, text): the one status line, with its own enter and exit
 */

export function prefersReducedMotion(): boolean {
  try { return window.matchMedia("(prefers-reduced-motion: reduce)").matches; } catch { return false; }
}

const tokens = new Map<string, string>();
function token(name: string, fallback: string): string {
  let v = tokens.get(name);
  if (v == null) {
    v = getComputedStyle(document.documentElement).getPropertyValue(name).trim() || fallback;
    tokens.set(name, v);
  }
  return v;
}
/** A duration token in milliseconds. */
function ms(name: "--dur-1" | "--dur-2" | "--dur-3"): number {
  const raw = token(name, name === "--dur-1" ? "120ms" : name === "--dur-2" ? "180ms" : "220ms");
  const n = parseFloat(raw);
  return raw.endsWith("s") && !raw.endsWith("ms") ? n * 1000 : n;
}
const easeOut = () => token("--ease-out", "cubic-bezier(0.2, 0, 0, 1)");
const easeInOut = () => token("--ease-in-out", "cubic-bezier(0.4, 0, 0.2, 1)");
const accent = () => token("--accent", "#4d8dff");

function canAnimate(el: Element): boolean {
  return typeof (el as HTMLElement).animate === "function" && !prefersReducedMotion();
}

/**
 * FLIP: measure every keyed child, run `mutate`, then animate each child from where it was to where it
 * now is. Keys come from `data-id` (tab rows, blocks), so a mutation may also add, remove or resize items:
 * new items fade in, moved items slide, resized items grow. Reduced motion runs `mutate` and stops.
 */
export function flip(container: HTMLElement, mutate: () => void, opts: { selector?: string; duration?: number } = {}): void {
  const selector = opts.selector ?? "[data-id]";
  if (prefersReducedMotion() || !container.isConnected) { mutate(); return; }
  const items = () => [...container.querySelectorAll<HTMLElement>(selector)];
  const before = new Map<string, DOMRect>();
  for (const el of items()) { const k = el.dataset.id; if (k) before.set(k, el.getBoundingClientRect()); }
  mutate();
  const duration = opts.duration ?? ms("--dur-3");
  for (const el of items()) {
    const k = el.dataset.id; if (!k || !canAnimate(el)) continue;
    const b = before.get(k);
    const a = el.getBoundingClientRect();
    if (!b) { fadeIn(el); continue; }
    const dx = b.left - a.left, dy = b.top - a.top;
    const dw = b.width - a.width, dh = b.height - a.height;
    if (Math.abs(dx) < 0.5 && Math.abs(dy) < 0.5 && Math.abs(dw) < 0.5 && Math.abs(dh) < 0.5) continue;
    const from: Keyframe = { transform: `translate(${dx}px, ${dy}px)`, transformOrigin: "0 0" };
    const to: Keyframe = { transform: "translate(0px, 0px)", transformOrigin: "0 0" };
    if (Math.abs(dw) >= 0.5 || Math.abs(dh) >= 0.5) {
      from.width = `${b.width}px`; from.height = `${b.height}px`;
      to.width = `${a.width}px`; to.height = `${a.height}px`;
    }
    el.animate([from, to], { duration, easing: easeOut() });
  }
}

/** Something appears where it is: a popover, a panel's contents, a row added to a list. */
export function fadeIn(el: HTMLElement, opts: { duration?: number; rise?: number } = {}): void {
  if (!canAnimate(el)) return;
  const rise = opts.rise ?? 3;
  el.animate([{ opacity: 0, transform: `translateY(${rise}px)` }, { opacity: 1, transform: "translateY(0)" }],
    { duration: opts.duration ?? ms("--dur-2"), easing: easeOut() });
}

/** A panel or bar arrives from an edge. */
export function slideIn(el: HTMLElement, opts: { from?: "right" | "left" | "top" | "bottom"; distance?: number } = {}): void {
  if (!canAnimate(el)) return;
  const d = opts.distance ?? 10;
  const from = opts.from ?? "right";
  const off = from === "right" ? `${d}px, 0` : from === "left" ? `${-d}px, 0` : from === "bottom" ? `0, ${d}px` : `0, ${-d}px`;
  el.animate([{ opacity: 0, transform: `translate(${off})` }, { opacity: 1, transform: "translate(0, 0)" }],
    { duration: ms("--dur-3"), easing: easeOut() });
}

/** One ring, once: this is the block or cell Claude just changed. */
export function pulse(el: HTMLElement): void {
  if (!canAnimate(el)) return;
  el.animate(
    [{ boxShadow: `0 0 0 0 ${accent()}` }, { boxShadow: `0 0 0 3px ${accent()}`, offset: 0.25 }, { boxShadow: "0 0 0 6px transparent" }],
    { duration: ms("--dur-3") * 3, easing: easeInOut() },
  );
}

export interface Toast { el: HTMLElement; close: () => void }
/** The one status line. Replaces any toast already up; `close` plays the exit and removes it. */
export function toast(root: HTMLElement, text: string, cls = "toast"): Toast {
  root.querySelectorAll(".toast").forEach((t) => t.remove());
  const el = document.createElement("div");
  el.className = cls;
  el.textContent = text;
  root.append(el);
  if (canAnimate(el)) el.animate([{ opacity: 0, transform: "translate(-50%, 6px)" }, { opacity: 1, transform: "translate(-50%, 0)" }], { duration: ms("--dur-2"), easing: easeOut() });
  let closed = false;
  const close = () => {
    if (closed) return;
    closed = true;
    if (!canAnimate(el)) { el.remove(); return; }
    const a = el.animate([{ opacity: 1, transform: "translate(-50%, 0)" }, { opacity: 0, transform: "translate(-50%, 4px)" }], { duration: ms("--dur-1"), easing: easeInOut() });
    a.onfinish = () => el.remove();
  };
  return { el, close };
}

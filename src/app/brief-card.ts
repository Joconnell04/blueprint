import { h, icon } from "./dom.js";
import { fadeIn } from "./motion.js";
import type { Controller } from "./controller.js";
import type { Brief } from "./model.js";

/**
 * The brief card: who it is for, what it must do, the one thing to remember. Three fields, nothing else:
 * the emphasis chips went with F7. One compact card above the outline in both modes, folded to a single
 * line once the audience or the goal is set, and it carries the readiness dot for a missing takeaway,
 * which is shown nowhere else.
 */
export function briefCard(c: Controller): HTMLElement {
  const b = c.project.briefCard ?? {};
  const warn = c.state?.warnings.find((w) => w.kind === "no-takeaway");
  const open = c.briefOpen ?? !(b.audience || b.goal || b.takeaway);
  const dot = () => (warn ? h("span", { class: "dot warn", title: warn.message }) : null);
  const edit = () => { c.briefOpen = true; c.emit(); };

  if (!open) {
    const rest = b.takeaway || b.goal || "";
    const line = h("span", { class: "bl trunc" },
      b.audience ? h("span", { class: "who" }, `For ${b.audience}${rest ? ": " : ""}`) : null, rest);
    const one = h("div", { class: "brief row link" },
      h("span", { class: "label" }, "Brief"), dot(), line, h("span", { class: "spacer" }),
      h("button", { class: "link", onclick: edit }, "Edit"));
    one.onclick = (e) => { if (!(e.target as HTMLElement).closest("button")) edit(); };
    return one;
  }

  const card = h("div", { class: "brief open" });

  const had = !!b.takeaway;
  const field = (label: string, key: "audience" | "goal" | "takeaway", placeholder: string) =>
    h("div", { class: `field ${key}` }, h("span", { class: "label" }, label),
      h("input", {
        class: `b${key}`, value: b[key] ?? "", placeholder, spellcheck: "false", "data-focus": `brief:${key}`,
        onchange: (e: Event) => {
          const val = (e.target as HTMLInputElement).value.trim();
          const patch: Partial<Brief> = {}; patch[key] = val;
          c.saveBrief(patch);
          if (key === "takeaway" && val && !had) { c.briefOpen = false; c.emit(); } // filled: fold it to one line
        },
        onkeydown: (e: KeyboardEvent) => { if (e.key === "Enter" || e.key === "Escape") (e.target as HTMLInputElement).blur(); },
      }));

  card.append(
    h("div", { class: "bhead" }, h("span", { class: "label" }, "Brief"), dot(), h("span", { class: "spacer" }),
      h("button", { class: "ghost icon", title: "Collapse", onclick: () => { c.briefOpen = false; c.emit(); } }, icon("chev", 12))),
    h("div", { class: "bgrid" },
      field("Audience", "audience", "Who reads it"),
      field("Goal", "goal", "What it must do"),
      field("Takeaway", "takeaway", "One sentence they should remember")));
  queueMicrotask(() => fadeIn(card)); // appended by the caller on the next tick
  return card;
}

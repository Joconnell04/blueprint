import { h, clear, popover, icon, twoStep } from "./dom.js";
import type { Controller } from "./controller.js";

/** Header popover: save the current plan as a template, or start this project from one. */
export function openTemplatesPopover(c: Controller, anchor: HTMLElement) {
  popover(anchor, (close) => {
    const box = h("div", { class: "stylepop" }, h("div", { class: "muted small-text pophdr" }, "Loading…"));
    const draw = (list: any[] | null) => {
      clear(box);
      if (!list) { box.append(h("div", { class: "muted small-text pophdr" }, "No templates")); return; }
      box.append(h("div", { class: "label pophdr" }, "Templates"));
      if (!list.length) box.append(h("div", { class: "muted small-text pophdr" }, "None yet"));
      for (const t of list) box.append(h("div", { class: "tplrow" },
        h("div", { class: "tplinfo" }, h("div", { class: "tplname trunc" }, t.name), h("div", { class: "tplmeta trunc" }, `${t.tabs} tabs · ${t.blocks} blocks · ${t.formatting}${t.description ? ` · ${t.description}` : ""}`)),
        twoStep(h("span", {}, "Use"), "Replace this plan?", async () => { await c.templates("apply", { template: t.id }); close(); c.toast?.(`Using ${t.name}`); }, "small"),
        twoStep(icon("trash", 12), "Delete?", async () => draw(await c.templates("delete", { template: t.id })), "ghost icon")));
      const hasPlan = c.project.tabs.length > 0;
      const name = h("input", { placeholder: "Name", spellcheck: "false", disabled: !hasPlan });
      const desc = h("input", { placeholder: "When to use it", spellcheck: "false", disabled: !hasPlan });
      const save = h("button", { class: "small primary", disabled: !hasPlan, onclick: async () => { const n = name.value.trim(); if (!n) { name.focus(); return; } save.disabled = true; draw(await c.templates("save", { name: n, description: desc.value.trim() })); c.toast?.("Template saved"); } }, "Save template");
      name.onkeydown = (e) => { if (e.key === "Enter") save.click(); };
      box.append(h("div", { class: "label pophdr sectop" }, "Save this plan"), h("div", { class: "inline" }, name, desc), h("div", { class: "inline end tight" }, h("button", { class: "small ghost", onclick: close }, "Close"), save));
    };
    void c.templates("list").then(draw);
    return box;
  }, { align: "auto" });
}

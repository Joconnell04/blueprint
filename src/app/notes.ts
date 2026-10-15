import { h, clear, twoStep, icon } from "./dom.js";
import { fadeIn } from "./motion.js";
import type { Controller } from "./controller.js";
import type { Note } from "./model.js";

/** Marker word per entry: who wrote it and where it stands. One word, never a sentence. */
function marker(n: Note): string {
  if (n.kind === "question") return n.status === "open" ? "Asks" : n.status === "resolved" ? "Answered" : "Withdrawn";
  return n.status === "open" ? "Open" : n.status === "resolved" ? "Done" : "Declined";
}

/** The thread on a tab or a block: one compact line per entry, the answer box inline under an open question. */
export function notesPanel(c: Controller, tabId: string, blockId: string | undefined, notes: Note[]): HTMLElement {
  const list = h("div", { class: "notes" });
  if (!notes.length) list.append(h("div", { class: "muted small-text" }, "No notes"));
  const sorted = [...notes].sort((a, b) => (a.status === "open" ? 0 : 1) - (b.status === "open" ? 0 : 1) || (a.kind === "question" ? 0 : 1) - (b.kind === "question" ? 0 : 1));
  for (const n of sorted) list.append(noteCard(c, n));
  const ta = h("textarea", { rows: 2, placeholder: "Note to Claude", "data-focus": `note:${blockId ?? tabId}` });
  const add = h("button", { class: "small", onclick: async () => { const t = ta.value.trim(); if (!t) return; add.disabled = true; await c.note("add", { tabId, blockId, text: t }); ta.value = ""; add.disabled = false; } }, "Add", h("span", { class: "key" }, "⌘↩"));
  ta.onkeydown = (e) => { if ((e.metaKey || e.ctrlKey) && e.key === "Enter") { e.preventDefault(); add.click(); } };
  return h("div", { class: "stack" }, list, h("div", { class: "addnote" }, ta, h("div", { class: "inline end" }, add)));
}

function noteCard(c: Controller, n: Note): HTMLElement {
  const card = h("div", { class: `note ${n.status}${n.kind === "question" ? " question" : ""}` });
  const draw = () => {
    clear(card);
    const isQ = n.kind === "question";
    const head = h("div", { class: "nh" },
      h("span", { class: `dot ${n.status === "open" ? "warn" : n.status === "resolved" ? "ok" : "bad"}` }),
      h("span", { class: "st" }, marker(n)),
      h("span", { class: "nt trunc", title: n.text }, n.text),
      h("span", { class: "acts" }));
    const acts = head.querySelector(".acts") as HTMLElement;
    if (n.status === "open" && !isQ) acts.append(h("button", { title: "Edit", onclick: () => edit() }, "Edit"));
    if (n.status !== "open") acts.append(h("button", { title: "Reopen", onclick: () => void c.note("reopen", { noteId: n.id }) }, "Reopen"));
    acts.append(twoStep(icon("trash", 11), "Delete?", () => void c.note("delete", { noteId: n.id }), ""));
    card.append(head);
    if (isQ && n.status === "open") {
      const input = h("input", { placeholder: "Answer" });
      const send = h("button", { class: "small primary", onclick: async () => { const t = input.value.trim(); if (!t) { input.focus(); return; } send.disabled = true; await c.note("answer", { noteId: n.id, text: t }); } }, "Send");
      input.onkeydown = (e) => { if (e.key === "Enter") send.click(); };
      card.append(h("div", { class: "inline tight" }, input, send));
    } else if (n.response) card.append(h("div", { class: "nr" }, h("div", { class: "who" }, isQ ? "you" : "Claude"), h("div", { class: "txt trunc", title: n.response }, n.response)));
  };
  const edit = () => {
    clear(card);
    const ta = h("textarea", { value: n.text });
    const save = h("button", { class: "small primary", onclick: async () => { const t = ta.value.trim(); if (t && t !== n.text) await c.note("edit", { noteId: n.id, text: t }); else draw(); } }, "Save");
    const cancel = h("button", { class: "small ghost", onclick: draw }, "Cancel");
    ta.onkeydown = (e) => { if (e.key === "Escape") draw(); if ((e.metaKey || e.ctrlKey) && e.key === "Enter") save.click(); };
    card.append(ta, h("div", { class: "inline tight end" }, save, cancel));
    fadeIn(card);
    ta.focus();
  };
  draw();
  return card;
}

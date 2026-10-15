import { h, clear, icon, twoStep } from "./dom.js";
import { fadeIn, slideIn } from "./motion.js";
import type { Controller, OutlineRow, TabFields } from "./controller.js";
import { vocab, type Tab } from "./model.js";
import { briefCard } from "./brief-card.js";
import { notesPanel } from "./notes.js";

/**
 * Layer 1, both modes. One row is a handle, a number, the name, and under it the story (deck) or the
 * purpose (workbook) clamped to two lines with a quiet expander. Nothing else: no message line, no section
 * pill, and the row actions float over the right gutter only while the row is hovered or selected.
 *
 * A divider is an item in the same list (`divider: true`): its name is the section title, drawn in caps
 * with a rule, and under it the agenda line the generated agenda prints for the section. It takes no
 * number. Sections come from where the dividers sit, so the server derives every slide's `section` and
 * nothing here reads it.
 *
 * Story, purpose, a divider's name and its agenda line patch their own row through ui_save_tab_fields,
 * which keeps the approval; renames, reorders and add/delete replace the outline through ui_save_tabs,
 * which withdraws it.
 */
export function renderOutlineView(c: Controller, root: HTMLElement) {
  const p = c.project;
  const deck = (p.kind ?? "workbook") === "deck";
  const v = vocab(p.kind ?? "workbook");
  const list = h("div", { class: `tablist${deck ? " deck" : ""}`, "data-scroll": "outline" });
  list.append(h("div", { class: "thead" }, h("div", { class: "grip-space" }), h("div", { class: "idx label" }, "#"), h("div", { class: "grow label" }, v.one)));

  const same = (a: string, b: string) => a.trim().toLowerCase() === b.trim().toLowerCase();
  /** Two dividers in a row saying the same thing are one divider. Merged here, before the save, not after it. */
  const merge = (rows: OutlineRow[]) => rows.filter((r, i) => { const prev = rows[i - 1]; return !(r.divider && prev?.divider && same(prev.name, r.name)); });

  /** Read every row back out of the DOM, in order. An omitted field keeps what the server already holds. */
  const rowsFromDom = (): OutlineRow[] => [...list.querySelectorAll<HTMLElement>(".trow")].map((r) => {
    const val = (sel: string) => (r.querySelector(sel) as HTMLInputElement | HTMLTextAreaElement | null)?.value.trim();
    const id = r.dataset.id || undefined;
    if (r.dataset.divider) return { id, name: val("input.dname") || "Section", divider: true, agendaLine: val("textarea.dagenda") };
    const row: OutlineRow = { id, name: val("input.tname") || `Untitled ${v.one.toLowerCase()}` };
    if (deck) row.story = val("textarea.tstory"); else row.purpose = val("textarea.tpurpose");
    return row;
  });
  const persistFromDom = () => void c.saveTabs(merge(rowsFromDom()));
  const fromModel = (): OutlineRow[] => p.tabs.map((t) => (t.divider
    ? { id: t.id || undefined, name: t.name, divider: true, agendaLine: t.agendaLine }
    : { id: t.id || undefined, name: t.name, purpose: t.purpose, story: t.story }));
  const save = () => void c.saveTabs(merge(fromModel()));
  const commitKey = (e: KeyboardEvent) => { if (e.key === "Enter" || e.key === "Escape") { e.preventDefault(); (e.target as HTMLElement).blur(); } };
  /** One field on one row. A row added this render has no id yet, so that one still saves the outline. */
  const saveField = (t: Tab, fields: TabFields) => { if (t.id) c.saveTabFields(t.id, fields); else persistFromDom(); };

  const move = (from: number, to: number) => {
    if (from === to || to < 0 || to >= p.tabs.length) return;
    const t = p.tabs.splice(from, 1)[0]; p.tabs.splice(to, 0, t);
    save();
  };
  const select = (t: Tab) => () => {
    c.selectedTab = t.id;
    list.querySelectorAll<HTMLElement>(".trow").forEach((x) => x.classList.toggle("sel", x.dataset.id === t.id));
    rerenderSide();
  };

  // ---- one flat drag: the insertion line moves, no row does, and the drop saves once ----
  const line = h("div", { class: "insline", hidden: true });
  const dragFrom = (row: HTMLElement, handle: HTMLElement) => {
    handle.onpointerdown = (e) => {
      e.preventDefault(); e.stopPropagation();
      const rows = [...list.querySelectorAll<HTMLElement>(".trow")];
      const from = rows.indexOf(row);
      if (from < 0) return;
      c.interacting = true;
      try { handle.setPointerCapture(e.pointerId); } catch { /* the host's nested iframes sometimes refuse it */ }
      // Measured once: nothing resizes during a drag, so a move only has to translate the line.
      const top = list.getBoundingClientRect().top - list.scrollTop;
      const edges = rows.map((r) => { const b = r.getBoundingClientRect(); return { mid: b.top + b.height / 2, y: b.top - top }; });
      const tail = rows[rows.length - 1].getBoundingClientRect().bottom - top;
      const anchor = edges[from].mid; // where the pointer is, in the list's own coordinates
      row.classList.add("dragging");
      let at = -1;
      const place = (i: number) => { at = i; line.hidden = false; line.style.transform = `translateY(${i < edges.length ? edges[i].y : tail}px)`; };
      place(from);
      // A captured pointer inside the host's nested iframes can report another frame's coordinates, so the
      // pointer is followed by how far it has moved, never by where it says it is.
      let base: number | null = null;
      const onMove = (ev: PointerEvent) => {
        if (base === null) base = ev.clientY;
        const y = anchor + ev.clientY - base;
        let i = 0; while (i < edges.length && y > edges[i].mid) i += 1;
        if (i !== at) place(i);
      };
      let done = false;
      const onUp = () => {
        if (done) return; done = true; // the capture fires it on the handle, the backstop on the window
        handle.removeEventListener("pointermove", onMove); handle.removeEventListener("pointerup", onUp); handle.removeEventListener("pointercancel", onUp);
        window.removeEventListener("pointerup", onUp, true);
        row.classList.remove("dragging"); line.hidden = true; c.interacting = false;
        const to = at > from ? at - 1 : at;
        if (to !== from) move(from, to);
      };
      handle.addEventListener("pointermove", onMove); handle.addEventListener("pointerup", onUp); handle.addEventListener("pointercancel", onUp);
      window.addEventListener("pointerup", onUp, true); // a release the capture never reports cannot leave the drag on
    };
  };

  /** Up, down and a two-step delete. Absolute over the row's right gutter, so they reserve no width. */
  const actions = (i: number) => h("div", { class: "acts" },
    h("button", { class: "ghost icon", title: "Move up", disabled: i === 0, onclick: (e: Event) => { e.stopPropagation(); move(i, i - 1); } }, icon("up", 12)),
    h("button", { class: "ghost icon", title: "Move down", disabled: i === p.tabs.length - 1, onclick: (e: Event) => { e.stopPropagation(); move(i, i + 1); } }, icon("down", 12)),
    twoStep(icon("trash", 12), "Delete?", () => { p.tabs.splice(i, 1); save(); }));

  const grip = () => h("div", { class: "grip", title: "Reorder" }, icon("grip", 14, true));

  let n = 0; // numbering counts slides only; a divider shows none
  p.tabs.forEach((t, i) => {
    const sel = c.selectedTab === t.id ? " sel" : "";
    let row: HTMLElement;
    if (t.divider) {
      const name = h("input", { class: "dname", value: t.name, placeholder: "Section", spellcheck: "false", "data-focus": `dname:${i}` });
      name.oninput = () => sizeName(name);
      name.onkeydown = commitKey;
      name.onchange = () => {
        const next = name.value.trim() || "Section";
        const merged = [p.tabs[i - 1], p.tabs[i + 1]].some((x) => x?.divider && same(x.name, next));
        if (merged) { name.value = next; persistFromDom(); } else saveField(t, { name: next });
      };
      // the section title on one line, the agenda line the deck prints for the section under it
      const head = h("div", { class: "dhead", style: { display: "flex", alignItems: "center", gap: "var(--sp-2)" } },
        name, h("span", { class: "dmeasure" }), h("span", { class: "rule" }));
      const agenda = textBox(t.agendaLine ?? "", "dagenda", "Agenda line", `agenda:${i}`, (text) => saveField(t, { agendaLine: text }), commitKey);
      agenda.style.marginLeft = "var(--sp-1)";   // the caps name keeps no negative gutter, so the words line up under it
      row = h("div", { class: `row trow div${sel}`, "data-id": t.id, "data-divider": "1", onclick: select(t) },
        grip(), h("div", { class: "idx" }), h("span", { class: "ndot" }), h("div", { class: "fields" }, head, agenda), actions(i));
    } else {
      n += 1;
      const notes = p.notes.filter((x) => x.tabId === t.id);
      const open = notes.filter((x) => x.status === "open").length;
      row = h("div", { class: `row trow${deck ? " deck" : ""}${sel}`, "data-id": t.id, onclick: select(t) },
        grip(), h("div", { class: "idx" }, String(n)),
        // the thread's state in front of the row, in a slot every row keeps so no name is indented differently
        h("span", { class: `ndot${notes.length ? (open ? " dot warn" : " dot ok") : ""}`, title: open ? `${open} open` : notes.length ? "Answered" : undefined }),
        h("div", { class: "fields" },
          h("input", { class: "tname", value: t.name, spellcheck: "false", "data-focus": `name:${t.id}`, onchange: persistFromDom, onkeydown: commitKey }),
          deck
            ? textBox(t.story ?? "", "tstory", "Story", `story:${t.id}`, (text) => saveField(t, { story: text }), commitKey)
            : textBox(t.purpose ?? "", "tpurpose", "Purpose", `purpose:${t.id}`, (text) => saveField(t, { purpose: text }), commitKey)),
        actions(i));
    }
    dragFrom(row, row.querySelector(".grip") as HTMLElement);
    list.append(row);
  });
  list.append(line);

  // ---- add: one button for a row, a quieter one for a divider ----
  const addRow = h("div", { class: "addrow" });
  const addDivider = () => {
    const at = p.tabs.findIndex((t) => t.id === c.selectedTab);
    let i = at >= 0 ? at + 1 : p.tabs.length;
    // A run of dividers merges, so a new one never lands straight above or below another: take the next gap.
    while (i < p.tabs.length && (p.tabs[i].divider || p.tabs[i - 1]?.divider)) i += 1;
    p.tabs.splice(i, 0, { id: "", name: "Section", purpose: "", divider: true, approved: false, layoutProposed: false, layoutApproved: false, blocks: [] });
    c.pendingFocus = `dname:${i}`; // the draw that follows the save puts the caret in the new name
    save();
  };
  const showTrigger = () => {
    clear(addRow);
    addRow.append(h("button", { class: "trigger", onclick: showForm }, icon("plus", 12), v.add));
    if (deck) addRow.append(h("button", { class: "link adddiv", onclick: addDivider }, v.addDivider));
  };
  const showForm = () => {
    clear(addRow);
    const name = h("input", { class: "tname", placeholder: "Name", spellcheck: "false" });
    const second = h("input", { class: deck ? "tstory" : "tpurpose", placeholder: deck ? "Story" : "Purpose", spellcheck: "false" });
    const submit = () => {
      const nm = name.value.trim(); if (!nm) { name.focus(); return; }
      const extra = second.value.trim();
      p.tabs.push({ id: "", name: nm, purpose: deck ? "" : extra, story: deck ? extra || undefined : undefined, approved: false, layoutProposed: false, layoutApproved: false, blocks: [] });
      save();
    };
    const key = (e: KeyboardEvent) => { if (e.key === "Enter") submit(); if (e.key === "Escape") showTrigger(); };
    name.onkeydown = key; second.onkeydown = key;
    addRow.append(h("div", { class: "grip-space" }), h("div", { class: "idx-space" }), h("div", { class: "inputs" }, name, second), h("div", { class: "hint" }, h("span", { class: "key" }, "↩"), "add", h("span", { class: "key" }, "esc")));
    fadeIn(addRow);
    name.focus();
  };
  showTrigger();
  list.append(addRow);

  const side = h("div", { class: "side" });
  const rerenderSide = () => {
    clear(side);
    const t = c.tab;
    if (!t) { side.append(h("div", { class: "sh" }, h("span", { class: "label" }, "Thread")), h("div", { class: "scroll muted" }, `Select a ${v.one.toLowerCase()}`)); return; }
    const panel = notesPanel(c, t.id, undefined, c.notesFor(t.id));
    // notes.ts tags nothing of its own yet; the shell needs a key to put the caret back after ⌘↩ (F8).
    panel.querySelector(".addnote textarea")?.setAttribute("data-focus", `note:${t.id}`);
    const body = h("div", { class: "scroll", "data-scroll": "side" }, panel);
    side.append(h("div", { class: "sh" }, h("span", { class: "label" }, "Thread"), h("span", { class: "t" }, t.name)), body);
    fadeIn(body);
  };
  rerenderSide();
  slideIn(side);

  const rows = p.tabs.filter((t) => !t.divider).length;
  const sections = p.tabs.filter((t) => t.divider).length;
  const openNotes = p.notes.filter((x) => x.status === "open").length;
  const footer = h("div", { class: "ftr" },
    h("span", { class: "msg" }, p.tabsApproved ? v.approved : `${v.footer(rows, sections)}${openNotes ? ` · ${openNotes} open note${openNotes === 1 ? "" : "s"}` : ""}`),
    h("span", { class: "spacer" }),
    p.tabsApproved
      ? h("button", { onclick: () => void c.approve("tabs", undefined, false) }, "Unapprove")
      : h("button", { class: "primary", disabled: p.tabs.length === 0, onclick: () => void c.approve("tabs") }, v.approve, h("span", { class: "key" }, "⌘↩")),
  );
  root.append(h("div", { class: "body" }, h("div", { class: "main" }, briefCard(c), list), side), footer);
  // Stories take the height they need up to two lines; only the ones with more to show get the expander.
  list.querySelectorAll<HTMLTextAreaElement>(".storywrap textarea").forEach((ta) => sizeText(ta, false));
  // A divider's name is as wide as its text, so the rule starts where the words end.
  list.querySelectorAll<HTMLInputElement>("input.dname").forEach((el) => sizeName(el));

  document.onkeydown = (e) => {
    const target = e.target as HTMLElement;
    if ((e.metaKey || e.ctrlKey) && e.key === "Enter") {
      if (!target.closest("textarea, .addrow") && !p.tabsApproved && p.tabs.length) void c.approve("tabs");
      return;
    }
    const down = e.key === "ArrowDown", up = e.key === "ArrowUp";
    if (!down && !up) return;
    const i = p.tabs.findIndex((t) => t.id === c.selectedTab);
    if ((e.metaKey || e.ctrlKey) && !e.altKey) { // ⌘↑ / ⌘↓: the selection walks the rows
      e.preventDefault();
      const t = p.tabs[i < 0 ? 0 : Math.min(p.tabs.length - 1, Math.max(0, i + (down ? 1 : -1)))];
      if (!t) return;
      select(t)();
      list.querySelector<HTMLElement>(`.trow[data-id="${t.id}"]`)?.scrollIntoView({ block: "nearest" });
      return;
    }
    if (e.altKey && !e.metaKey && !e.ctrlKey && i >= 0) { // ⌥↑ / ⌥↓: the selected row moves
      e.preventDefault();
      move(i, i + (down ? 1 : -1));
    }
  };
}

/** Two lines of the row's story or purpose, with a quiet expander when there is more. Editable where it sits. */
function textBox(value: string, cls: string, placeholder: string, key: string, commit: (text: string) => void, commitKey: (e: KeyboardEvent) => void): HTMLElement {
  const box = h("textarea", {
    class: cls, rows: "2", value, placeholder, spellcheck: "false", "data-focus": key,
    onchange: () => commit(box.value.trim()), onkeydown: commitKey,
    // while it is being written, show all of it; back to two lines when the cursor leaves
    onfocus: () => sizeText(box, true), oninput: () => sizeText(box, true), onblur: () => sizeText(box, box.classList.contains("open")),
  });
  const more = h("button", { class: "link more", hidden: true }, "more");
  more.onclick = (e) => {
    e.stopPropagation();
    sizeText(box, box.classList.toggle("open"));
    more.textContent = box.classList.contains("open") ? "less" : "more";
  };
  return h("div", { class: "storywrap" }, box, more);
}

/**
 * A textarea cannot size itself, so measure it: the content height up to two lines when closed, all of it
 * when open. The expander shows only when there is a third line to reveal.
 */
function sizeText(ta: HTMLTextAreaElement, open: boolean): void {
  const wrap = ta.parentElement;
  const btn = wrap?.querySelector<HTMLButtonElement>("button.more");
  const line = parseFloat(getComputedStyle(ta).lineHeight) || 0;
  const border = ta.offsetHeight - ta.clientHeight; // border-box: the height must carry the borders too
  const cap = Math.round(line * 2);
  wrap?.classList.remove("hasmore");
  ta.style.height = "0px"; // scrollHeight never reports less than the box, so collapse it before measuring
  let full = ta.scrollHeight;
  const over = full > cap + 1;
  if (over) { wrap?.classList.add("hasmore"); ta.style.height = "0px"; full = ta.scrollHeight; } // the expander keeps a gutter, which rewraps
  ta.style.height = `${(open ? full : Math.min(full, cap)) + border}px`;
  if (btn) btn.hidden = !over;
}

/** A divider's name is an input that looks like the caps label, so it is as wide as the words in it. */
function sizeName(el: HTMLInputElement): void {
  const span = el.parentElement?.querySelector<HTMLElement>(".dmeasure");
  if (!span) return;
  span.textContent = el.value.trim() || el.placeholder;
  el.style.width = `${Math.min(280, span.offsetWidth + 14)}px`;
}

/** Old name, same view. */
export const renderTabsView = renderOutlineView;

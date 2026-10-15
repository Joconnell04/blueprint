import { h, clear, icon, popover } from "./dom.js";
import { fadeIn, pulse, slideIn } from "./motion.js";
import type { Controller } from "./controller.js";
import { colLetter } from "./model.js";
import type { Block, Project, Tab } from "./model.js";

interface SnapCell { a: string; v: unknown; k: string; fmt?: string; d?: string }
/**
 * One block as the workbook builder wrote it, cells included. `adjusted` is the reason it moved: the
 * builder leaves a bare string here and the whole record on the sheet (backlog W11).
 */
interface SnapBlock { id: string; row: number; col: number; rows: number; cols: number; type: string; noteCol?: number; adjusted?: string | { reason?: string } }
/** One deviation the workbook build took: the block, the cells either end, and why. */
interface SheetAdjustment { sheet?: string | number; blockId?: string; id?: string; from?: unknown; to?: unknown; reason?: string }
interface SnapSheet { name: string; widths?: Record<string, number>; maxRow: number; maxCol: number; cells: SnapCell[]; merges?: string[]; blocks?: SnapBlock[]; charts?: { blockId: string; anchor: string; kind: string; series: number }[]; adjusted?: SheetAdjustment[]; truncated?: boolean }
interface Snapshot { sheets: SnapSheet[] }

function parseAddr(a: string): { row: number; col: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(a); if (!m) return { row: 0, col: 0 };
  let col = 0; for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]), col };
}

/** Dependency arrows are off until asked for, and the answer is remembered per project. */
function arrowsRemembered(projectId: string): boolean {
  try { return localStorage.getItem(`wp.arrows.${projectId}`) === "1"; } catch { return false; }
}
function rememberArrows(projectId: string, on: boolean): void {
  try { localStorage.setItem(`wp.arrows.${projectId}`, on ? "1" : "0"); } catch { /* sandboxed host */ }
}

/** Layer 3. A workbook shows the built sheets, a deck shows the built slides; the chrome is the same. */
export function renderBuiltView(c: Controller, root: HTMLElement) {
  if ((c.project.kind ?? "workbook") === "deck") renderDeckBuilt(c, root);
  else renderWorkbookBuilt(c, root);
}

// ---- shared panel pieces: both views name the file the same way and carry the same thread ledger ----

/** The panel header: file name, any state chips, the save progress chip, Save copy, Reveal. The full path is the name's tooltip. */
function fileHeader(c: Controller, path: string, chips: (Node | null)[] = [], pdfPath?: string): HTMLElement {
  return h("div", { class: "sh" },
    h("span", { class: "t", title: path || undefined }, path.split("/").pop() || "(no file)"),
    ...chips,
    ...saveControls(c, pdfPath),
    h("button", { class: "ghost icon", title: "Reveal", onclick: () => { if (path) void c.reveal(path); } }, icon("reveal", 13)));
}

// ---- Save copy -----------------------------------------------------------------------------------
//
// The canvas is a sandboxed iframe: it cannot read the built file and cannot download it. So the bytes
// come through the app-only ui_file_chunk tool, 500 KB at a time, and the host writes them
// (docs/research/platform-docs.md §1 and §3). Absent entirely when the host takes no downloads.

type Which = "build" | "pdf" | "plan";
interface Chunk { name: string; mimeType: string; total: number; offset: number; data: string }

/** The server's cap, so one request is one chunk. */
const CHUNK = 512_000;

function decode(b64: string): Uint8Array {
  const s = atob(b64);
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}

/** Base64 for the assembled file. Each slice divides by 3, so the encoded slices join without padding in the middle. */
function encode(bytes: Uint8Array): string {
  const STEP = 32_766;
  let out = "";
  for (let i = 0; i < bytes.length; i += STEP) {
    const end = Math.min(i + STEP, bytes.length);
    let s = "";
    for (let j = i; j < end; j++) s += String.fromCharCode(bytes[j]);
    out += btoa(s);
  }
  return out;
}

function sizeLabel(bytes: number): string {
  return bytes >= 1_048_576 ? `${(bytes / 1_048_576).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`;
}

/** Pull one file through ui_file_chunk until `total` bytes have arrived. Null when the server refused or the file changed under us. */
async function fetchFile(c: Controller, which: Which, onBytes: (n: number) => void): Promise<{ name: string; mimeType: string; base64: string } | null> {
  const parts: Uint8Array[] = [];
  let at = 0, got = 0, total = 0, done = false, name = "file", mimeType = "application/octet-stream";
  for (let guard = 0; guard < 2000 && !done; guard++) {
    const r = await c.app.callServerTool({ name: "ui_file_chunk", arguments: { project: c.project.id, which, offset: at, size: CHUNK } });
    const ch = r.structuredContent as Chunk | undefined;
    if (r.isError || !ch || typeof ch.data !== "string") return null;
    name = ch.name; mimeType = ch.mimeType; total = ch.total;
    const bytes = decode(ch.data);
    parts.push(bytes); got += bytes.length; at = ch.offset + bytes.length;
    onBytes(got);
    if (got >= total) done = true;
    else if (!bytes.length) return null;
  }
  if (!done) return null;
  const all = new Uint8Array(got);
  let o = 0;
  for (const part of parts) { all.set(part, o); o += part.length; }
  return { name, mimeType, base64: encode(all) };
}

/** The bytes-so-far chip and the Save copy button, or nothing when the host has no downloadFile. */
function saveControls(c: Controller, pdfPath?: string): (Node | null)[] {
  if (!c.app.getHostCapabilities()?.downloadFile) return [];
  const deck = (c.project.kind ?? "workbook") === "deck";
  const chip = h("span", { class: "chip line num savechip", hidden: true }, "0 KB");
  const btn = h("button", { class: "ghost icon", title: "Save copy" }, icon("down", 13));
  let busy = false;
  const save = async (which: Which) => {
    if (busy) return;
    busy = true; btn.disabled = true; chip.hidden = false; chip.textContent = "0 KB";
    try {
      const file = await fetchFile(c, which, (n) => { chip.textContent = sizeLabel(n); });
      if (!file) { c.toast?.("Nothing to save"); return; }
      const r = await c.app.downloadFile({ contents: [{ type: "resource", resource: { uri: `file:///${encodeURI(file.name)}`, mimeType: file.mimeType, blob: file.base64 } }] });
      c.toast?.(r?.isError ? "Host refused" : `Saved ${file.name}`);
    } catch { c.toast?.("Host refused"); }
    finally { busy = false; btn.disabled = false; chip.hidden = true; }
  };
  const entries: { label: string; ext: string; which: Which }[] = [
    { label: deck ? "Deck" : "Workbook", ext: deck ? ".pptx" : ".xlsx", which: "build" },
    ...(pdfPath ? [{ label: "PDF", ext: ".pdf", which: "pdf" as Which }] : []),
    { label: "Plan JSON", ext: ".json", which: "plan" },
  ];
  btn.onclick = () => popover(btn, (close) => h("div", { class: "savepop" },
    h("div", { class: "label pophdr" }, "Save copy"),
    ...entries.map((e) => h("button", { "data-which": e.which, onclick: () => { close(); void save(e.which); } }, e.label, h("span", { class: "desc" }, e.ext)))));
  return [chip, btn];
}

/** Builder warnings, collapsed behind a count chip. */
function warningsField(warnings: string[]): HTMLElement | null {
  if (!warnings.length) return null;
  const list = h("ul", { class: "warnlist", hidden: true }, ...warnings.map((w) => h("li", {}, h("span", { class: "dot warn" }), h("span", {}, w))));
  const chip = h("button", { class: "chip warn", onclick: () => { list.hidden = !list.hidden; if (!list.hidden) fadeIn(list); } },
    h("span", { class: "dot warn" }), `${warnings.length} warning${warnings.length === 1 ? "" : "s"}`);
  return h("div", { class: "field" }, h("span", { class: "label" }, "Builder"), h("div", { class: "inline" }, chip), list);
}

/** Every note on the plan, one line each, with the response under it. */
function threadField(p: Project): HTMLElement {
  const notes = p.notes;
  return h("div", { class: "field" }, h("span", { class: "label" }, `Thread (${notes.length})`),
    notes.length ? h("div", { class: "notes" }, ...notes.map((n) => {
      const t = p.tabs.find((t) => t.id === n.tabId); const blk = n.blockId ? t?.blocks.find((x) => x.id === n.blockId) : undefined;
      const mark = n.kind === "question" ? (n.status === "resolved" ? "Answered" : n.status === "declined" ? "Withdrawn" : "Unanswered") : n.status === "resolved" ? "Done" : n.status === "declined" ? "Declined" : "Open";
      return h("div", { class: `note ${n.status}` },
        h("div", { class: "nh" }, h("span", { class: `dot ${n.status === "resolved" ? "ok" : n.status === "declined" ? "bad" : "warn"}` }), h("span", { class: "st" }, mark), h("span", { class: "nt trunc", title: n.text }, n.text), h("span", { class: "muted xs" }, `${t?.name ?? "?"}${blk ? ` / ${blk.title}` : ""}`)),
        n.response ? h("div", { class: "nr" }, h("div", { class: "who" }, n.kind === "question" ? "you" : "Claude"), h("div", { class: "txt trunc", title: n.response }, n.response)) : null);
    })) : h("div", { class: "muted small-text" }, "No notes"));
}

/** The one commit in this view: drop back to the layouts so the plan can change and be rebuilt. */
function reopenFooter(c: Controller, msg: string): HTMLElement {
  return h("div", { class: "ftr" },
    h("span", { class: "msg num" }, msg),
    h("span", { class: "spacer" }),
    h("button", { onclick: async () => { await c.approve("all_layouts", undefined, false); c.state!.view = "layouts"; c.setState(c.state!); } }, "Reopen layouts"));
}

// ---- the adjusted chip, shared by both kinds (docs/research/output-quality.md, backlog Q6 and W11) ----

/** One block's cells on the grid, the way the plan and the build both write them. */
interface Cells { col: number; row: number; cols: number; rows: number }
/** One line of the adjusted list: where it happened, what moved, from where to where and why. */
interface AdjRow { ix: string; title: string; from: Cells; to: Cells; reason?: string; pick: () => void }

const sameCells = (a: Cells, b: Cells) => a.col === b.col && a.row === b.row && a.cols === b.cols && a.rows === b.rows;
const cellsText = (a: Cells) => `r${a.row} c${a.col} ${a.rows}×${a.cols}`;
const isCells = (a: unknown): a is Cells => !!a && ["col", "row", "cols", "rows"].every((k) => Number.isFinite((a as Record<string, unknown>)[k]));

/** The chip's list: one line per move, and the one button that writes them all into the plan. */
function openAdjusted(anchor: HTMLElement, rows: AdjRow[], keep: () => void): void {
  popover(anchor, (close) => h("div", { class: "adjlist" },
    h("div", { class: "label pophdr" }, "Adjusted"),
    ...rows.slice(0, 40).map((a) => h("button", { class: "arow", onclick: () => { close(); a.pick(); } },
      h("span", { class: "mono ix" }, a.ix),
      h("span", { class: "t trunc" }, a.title),
      h("span", { class: "mono geo" }, `${cellsText(a.from)} → ${cellsText(a.to)}`),
      a.reason ? h("span", { class: "why" }, a.reason) : null)),
    h("div", { class: "inline end tight afoot" },
      h("button", { class: "primary small", onclick: () => { close(); keep(); } }, "Keep positions"))), { align: "auto" });
}

/** What the builder moved, as one toolbar chip. Absent when the build took the plan as it was. */
function adjustedChip(rows: AdjRow[], keep: () => void): HTMLElement | null {
  if (!rows.length) return null;
  return h("button", { class: "chip line num adjchip", title: "What the builder moved", onclick: (e: Event) => openAdjusted(e.currentTarget as HTMLElement, rows, keep) },
    `adjusted ${rows.length}`, icon("chev", 9));
}

/**
 * Write the built geometry back into the plan, one ui_save_layout per row, so the next build draws what
 * this one drew. Every record then matches the plan, which is what takes the chip away.
 */
function keepGeometry(c: Controller, groups: { tab?: Tab; moves: { blockId?: string; to: Cells }[] }[], word: string): void {
  let kept = 0;
  for (const g of groups) {
    if (!g.tab) continue;
    let touched = false;
    const next = g.tab.blocks.map((bl) => {
      const move = g.moves.find((m) => m.blockId === bl.id);
      if (!move || sameCells(bl, move.to)) return bl;
      touched = true;
      return { ...bl, col: move.to.col, row: move.to.row, cols: move.to.cols, rows: move.to.rows };
    });
    if (!touched) continue;
    c.saveLayout(g.tab, next);
    kept += 1;
  }
  c.toast?.(kept ? `Kept ${kept} ${word}${kept === 1 ? "" : "s"}` : "Nothing to keep");
}

// ---------------------------------------------------------------- workbook

/** One move on one sheet, with both ends resolved to cells. */
interface SheetMove { sheet: string; index: number; blockId: string; title: string; from: Cells; to: Cells; reason?: string }

/**
 * What the builder moved, for a workbook (backlog W11). Three places carry it and they say the same
 * thing: the build summary's own list, the snapshot's per-sheet list, and the bare reason string the
 * snapshot leaves on a block it adjusted. Geometry a record does not carry comes from the two ends
 * themselves: the plan block is the "from", the snapshot's block extent is the "to". A record whose plan
 * block already sits where the build put it is dropped, so writing the geometry back clears the chip.
 */
function sheetAdjustments(b: { adjusted?: SheetAdjustment[] }, snap: Snapshot | null, p: Project): SheetMove[] {
  const sheets = snap?.sheets ?? [];
  const out: SheetMove[] = [];
  const seen = new Set<string>();
  const take = (index: number, rec: SheetAdjustment) => {
    const sheet = sheets[index];
    const blockId = rec.blockId ?? rec.id;
    if (!sheet || !blockId || seen.has(`${index}:${blockId}`)) return;
    const built = sheet.blocks?.find((x) => x.id === blockId);
    const plan = p.tabs.find((t) => t.name === sheet.name)?.blocks.find((x) => x.id === blockId);
    const from = isCells(rec.from) ? rec.from : plan;
    const to = isCells(rec.to) ? rec.to : built;
    if (!from || !to || sameCells(from, to)) return;
    if (plan && sameCells(plan, to)) return;   // the plan already says what the build did
    seen.add(`${index}:${blockId}`);
    out.push({ sheet: sheet.name, index, blockId, title: plan?.title ?? blockId, from, to, reason: rec.reason });
  };
  for (const rec of b.adjusted ?? []) take(typeof rec.sheet === "number" ? rec.sheet - 1 : sheets.findIndex((s) => s.name === rec.sheet), rec);
  sheets.forEach((s, i) => {
    for (const rec of s.adjusted ?? []) take(i, rec);
    for (const bl of s.blocks ?? []) if (bl.adjusted) take(i, { blockId: bl.id, reason: typeof bl.adjusted === "string" ? bl.adjusted : bl.adjusted.reason });
  });
  return out;
}

function renderWorkbookBuilt(c: Controller, root: HTMLElement) {
  const p = c.project;
  const b = c.state!.build ?? p.buildSummary ?? {};
  const snap = (c.state!.snapshot ?? null) as Snapshot | null;
  let sheetIdx = Math.max(0, snap?.sheets.findIndex((s) => s.name === (c.tab?.name ?? "")) ?? 0);
  let arrows = arrowsRemembered(p.id);
  const adjusted = sheetAdjustments(b as { adjusted?: SheetAdjustment[] }, snap, p);
  const keepPositions = () => {
    const bySheet = new Map<number, SheetMove[]>();
    for (const m of adjusted) bySheet.set(m.index, [...(bySheet.get(m.index) ?? []), m]);
    keepGeometry(c, [...bySheet].map(([i, moves]) => ({ tab: p.tabs.find((t) => t.name === snap?.sheets[i]?.name), moves })), "sheet");
  };
  /** The chip's lines, built at draw time so each one can walk to the sheet it names. */
  const adjRows = (): AdjRow[] => adjusted.map((m) => ({
    ix: String(m.index + 1), title: m.title, from: m.from, to: m.to, reason: m.reason,
    pick: () => { sheetIdx = m.index; drawSheet(); flashBlock(m.blockId); },
  }));

  const main = h("div", { class: "main" });
  const side = h("div", { class: "side" });
  root.append(h("div", { class: "body" }, main, side));

  const drawSheet = () => {
    clear(main);
    if (!snap) { main.append(h("div", { class: "center" }, h("div", {}, "Built"), h("div", { class: "mono xs muted" }, p.buildPath ?? ""), h("div", { class: "muted small-text" }, "Rebuild for a preview"))); return; }
    const strip = h("div", { class: "tabstrip" }, ...snap.sheets.map((s, i) => h("button", { class: i === sheetIdx ? "on" : "", onclick: () => { sheetIdx = i; drawSheet(); } }, s.name)));
    const sheet = snap.sheets[sheetIdx];
    const plan = p.tabs.find((t) => t.name === sheet.name);
    const diffs = diffBlocks(plan?.blocks ?? [], sheet.blocks ?? []);
    const deps = dependencies(sheet);
    for (const bl of sheet.blocks ?? []) if ((bl.type === "calc" || bl.type === "output") && !deps.some((d) => d.from === bl.id)) diffs.push({ id: bl.id, text: `${plan?.blocks.find((x) => x.id === bl.id)?.title ?? bl.id}: reads no other block` });
    const arrowBtn = h("button", { class: `ghost${arrows ? " on" : ""}`, title: "Arrows (A)", onclick: () => setArrows(!arrows) }, icon("arrows", 11), "Arrows");
    const toolbar = h("div", { class: "toolbar" },
      h("span", { class: "chip line num" }, `${sheet.cells.length}${sheet.truncated ? "+" : ""} cells`),
      h("span", { class: "chip line num" }, `${sheet.cells.filter((x) => x.k === "formula").length} formulas`),
      (sheet.charts?.length ?? 0) ? h("span", { class: "chip line num" }, `${sheet.charts!.length} chart${sheet.charts!.length === 1 ? "" : "s"}`) : null,
      adjustedChip(adjRows(), keepPositions),
      h("span", { class: "purpose" }),
      arrowBtn,
      diffs.length ? h("span", { class: "chip warn", title: diffs.map((d) => d.text).join("\n") }, h("span", { class: "dot warn" }), `${diffs.length} off plan`) : h("span", { class: "chip ok" }, h("span", { class: "dot ok" }), "matches plan"));
    main.append(strip, toolbar, grid(sheet, diffs));
  };
  const setArrows = (on: boolean) => { arrows = on; rememberArrows(p.id, on); rememberScroll(); drawSheet(); restoreScroll(); c.toast?.(on ? "Arrows on" : "Arrows off"); };
  let scrollAt: [number, number] = [0, 0];
  const rememberScroll = () => { const w = main.querySelector<HTMLElement>(".canvaswrap"); if (w) scrollAt = [w.scrollLeft, w.scrollTop]; };
  const restoreScroll = () => { const w = main.querySelector<HTMLElement>(".canvaswrap"); if (w) { w.scrollLeft = scrollAt[0]; w.scrollTop = scrollAt[1]; } };

  const grid = (sheet: SnapSheet, diffs: Diff[]) => {
    const plan = p.tabs.find((t) => t.name === sheet.name);
    const CW = 8, RH = 22; // px per width unit, px per row
    const cols = Math.max(sheet.maxCol + 1, 8), rows = Math.max(sheet.maxRow + 2, 12);
    const widthOf = (col: number) => Math.round(((sheet.widths?.[colLetter(col)] ?? 8.43) + 0.7) * CW);
    const xs: number[] = [30]; for (let c = 1; c <= cols; c++) xs.push(xs[c - 1] + widthOf(c));
    const wrap = h("div", { class: "canvaswrap" });
    const canvas = h("div", { class: "canvas", style: { width: `${xs[cols]}px`, height: `${20 + rows * RH}px` } });
    canvas.append(
      h("div", { class: "colhdr", style: { marginLeft: "30px", width: `${xs[cols] - 30}px` } }, ...Array.from({ length: cols }, (_, i) => h("div", { style: { width: `${widthOf(i + 1)}px` } }, colLetter(i + 1)))),
      h("div", { class: "rowhdr", style: { position: "absolute", top: "20px", left: 0, height: `${rows * RH}px` } }, ...Array.from({ length: rows }, (_, i) => h("div", { style: { height: `${RH}px`, lineHeight: `${RH}px` } }, String(i + 1)))),
      h("div", { class: "corner", style: { position: "absolute", top: 0, left: 0 } }));
    const merged = new Map<string, number>();
    for (const m of sheet.merges ?? []) { const [a, z] = m.split(":"); const s = parseAddr(a), e = parseAddr(z); merged.set(a, e.col - s.col + 1); }
    for (const cell of sheet.cells) {
      const { row, col } = parseAddr(cell.a); if (!row) continue;
      const span = merged.get(cell.a) ?? 1;
      const text = cell.d ?? (typeof cell.v === "number" ? cell.v.toLocaleString() : String(cell.v ?? ""));
      const el = h("div", { class: `sc k-${cell.k}`, "data-a": cell.a, title: cell.k === "formula" ? String(cell.v) : undefined, style: { left: `${xs[col - 1]}px`, top: `${20 + (row - 1) * RH}px`, width: `${xs[col - 1 + span] - xs[col - 1] - 1}px`, height: `${RH - 1}px` } }, text);
      canvas.append(el);
    }
    for (const bl of sheet.blocks ?? []) {
      const d = diffs.find((x) => x.id === bl.id);
      const planTitle = plan?.blocks.find((x) => x.id === bl.id)?.title ?? bl.id;
      const outline = h("div", { class: `sblk${d ? " diff" : ""}`, "data-b": bl.id, title: planTitle, style: { left: `${xs[bl.col - 1] - 2}px`, top: `${20 + (bl.row - 1) * RH - 2}px`, width: `${xs[bl.col - 1 + bl.cols] - xs[bl.col - 1] + 3}px`, height: `${bl.rows * RH + 3}px` } });
      // one stray click should not cost a model turn, so ask first
      outline.onclick = () => askExplain(c, outline, planTitle, `Explain the block "${planTitle}" on sheet "${sheet.name}" of the built workbook: what each formula does, which cells it reads, and any assumption baked in. Keep it short.`);
      canvas.append(outline);
    }
    if (arrows) {
      // from the reading block's edge to the block it reads
      const deps = dependencies(sheet);
      const svgNS = "http://www.w3.org/2000/svg";
      const overlay = document.createElementNS(svgNS, "svg");
      overlay.setAttribute("class", "deps"); overlay.setAttribute("width", String(xs[cols])); overlay.setAttribute("height", String(20 + rows * RH));
      overlay.innerHTML = `<defs><marker id="arr" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="currentColor"/></marker></defs>`;
      const center = (bl: { row: number; col: number; rows: number; cols: number }) => ({ x: (xs[bl.col - 1] + xs[bl.col - 1 + bl.cols]) / 2, top: 20 + (bl.row - 1) * RH, bottom: 20 + (bl.row - 1 + bl.rows) * RH });
      for (const d of deps) {
        const a = sheet.blocks?.find((b) => b.id === d.from); if (!a) continue;
        const ca = center(a);
        if (d.cross) { const t = document.createElementNS(svgNS, "text"); t.setAttribute("x", String(xs[a.col - 1] + 4)); t.setAttribute("y", String(ca.top - 4)); t.setAttribute("class", "deplbl"); t.textContent = "reads other sheets"; overlay.append(t); continue; }
        const b = sheet.blocks?.find((x) => x.id === d.to); if (!b) continue;
        const cb = center(b);
        const up = cb.bottom <= ca.top;
        const y1 = up ? ca.top : ca.bottom, y2 = up ? cb.bottom : cb.top;
        const path = document.createElementNS(svgNS, "path");
        const dx = Math.abs(cb.x - ca.x) < 8 ? 24 : 0;
        path.setAttribute("d", `M${ca.x + dx} ${y1} C ${ca.x + dx} ${(y1 + y2) / 2}, ${cb.x + dx} ${(y1 + y2) / 2}, ${cb.x + dx} ${y2}`);
        path.setAttribute("marker-end", "url(#arr)");
        overlay.append(path);
      }
      canvas.append(overlay);
    }
    for (const ch of sheet.charts ?? []) {
      const { row, col } = parseAddr(ch.anchor);
      canvas.append(h("div", { class: "schart", style: { left: `${xs[col - 1]}px`, top: `${20 + (row - 1) * RH}px`, width: `${xs[Math.min(cols, col + 5)] - xs[col - 1]}px`, height: `${10 * RH}px` } }, `${ch.kind} chart, ${ch.series} series`));
    }
    wrap.append(canvas);
    return wrap;
  };

  drawSheet();
  const flashCell = (a: string) => { const el = main.querySelector<HTMLElement>(`.sc[data-a="${a}"]`); if (!el) return; el.scrollIntoView({ block: "center", inline: "center" }); pulse(el); };
  function flashBlock(id: string): void { const el = main.querySelector<HTMLElement>(`.sblk[data-b="${id}"]`); if (!el) return; el.scrollIntoView({ block: "center", inline: "center" }); pulse(el); }

  // ---- side: file, sheets, one status line, warnings behind a count, the thread ledger ----
  const check = b.check as { engine?: string; error?: string; evaluated?: number; findings?: { sheet: string; cell: string; code: string; message: string; severity: string; formula?: string }[] } | undefined;
  const findings = check?.findings ?? [];

  const checkFindings = h("ul", { class: "warnlist", hidden: true }, ...findings.slice(0, 40).map((f) => h("li", { class: "link", title: f.formula, onclick: () => { const i = snap?.sheets.findIndex((x) => x.name === f.sheet) ?? -1; if (i >= 0) { sheetIdx = i; drawSheet(); flashCell(f.cell); } } }, h("span", { class: `dot ${f.severity === "error" ? "bad" : "warn"}` }), h("span", {}, h("span", { class: "mono" }, `${f.sheet}!${f.cell}`), ` ${f.code}: ${f.message}`))));
  const checkLine = check
    ? check.error
      ? h("div", { class: "inline tight small-text muted", title: check.error }, h("span", { class: "dot warn" }), "Check unavailable")
      : findings.length
        ? statusButton(`${findings.length} problem${findings.length === 1 ? "" : "s"}`, checkFindings)
        : h("div", { class: "inline tight small-text" }, h("span", { class: "dot ok" }), h("span", { class: "muted" }, `${check.evaluated ?? 0} formulas, no problems`))
    : null;

  const body = h("div", { class: "scroll" },
    h("div", { class: "field" }, h("span", { class: "label" }, "Sheets"),
      h("div", { class: "rows" }, ...((b.sheets ?? []) as { name: string; cells?: number; formulas?: number }[]).map((s) => {
        const i = snap?.sheets.findIndex((x) => x.name === s.name) ?? -1;
        return h("div", { class: `sheetrow${i >= 0 ? " link" : ""}`, onclick: () => { if (i >= 0) { sheetIdx = i; drawSheet(); } } },
          h("span", { class: "t" }, s.name), h("span", { class: "n" }, `${s.cells ?? 0} cells`), h("span", { class: "n" }, `${s.formulas ?? 0} f`));
      }))),
    checkLine ? h("div", { class: "field" }, h("span", { class: "label" }, "Formulas"), checkLine, checkFindings) : null,
    warningsField((b.warnings ?? []) as string[]),
    threadField(p));

  side.append(fileHeader(c, p.buildPath ?? ""), body);
  slideIn(side);

  const sheetCount = ((b.sheets ?? []) as unknown[]).length;
  const formulaCount = ((b.sheets ?? []) as { formulas?: number }[]).reduce((n, s) => n + (s.formulas ?? 0), 0);
  root.append(reopenFooter(c, `${sheetCount} sheet${sheetCount === 1 ? "" : "s"} · ${formulaCount} formulas`));

  document.onkeydown = (e) => {
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    if ((e.target as HTMLElement).closest("input, textarea, select")) return;
    if (e.key.toLowerCase() === "a") { e.preventDefault(); setArrows(!arrows); }
  };
}

/** The status line that expands a findings list. */
function statusButton(label: string, list: HTMLElement): HTMLElement {
  return h("button", { class: "status", onclick: (e: Event) => { const btn = e.currentTarget as HTMLElement; list.hidden = !list.hidden; btn.classList.toggle("open", !list.hidden); if (!list.hidden) fadeIn(list); } },
    h("span", { class: "dot warn" }), label, icon("chev", 10));
}

/** One stray click should not cost a model turn, so ask first. */
function askExplain(c: Controller, anchor: Element, title: string, message: string): void {
  popover(anchor as HTMLElement, (close) => h("div", { class: "confirm" },
    h("div", { class: "t" }, `Explain "${title}"?`),
    h("div", { class: "acts" },
      h("button", { onclick: close }, "Cancel"),
      h("button", { class: "primary", onclick: () => { close(); void c.sendMessage(message); } }, "Ask"))), { align: "auto" });
}

// ---------------------------------------------------------------- deck

/**
 * One deviation the builder took between the approved plan and the file (docs/research/output-quality.md,
 * "The fidelity setting"). The build summary lists them with their slide; the snapshot carries the same
 * records per slide, and repeats each one on the shapes it drew.
 */
interface Adjustment { slide: number; blockId?: string; from: Cells; to: Cells; reason?: string }

/** One drawn shape in the deck snapshot (builder/SPEC-deck.md, "Build snapshot"): points on the slide. */
interface SnapShape { id: string; blockId?: string | null; role?: string; x: number; y: number; w: number; h: number; kind?: string; text?: string; fontPt?: number; adjusted?: Omit<Adjustment, "slide"> }
interface SnapSlide { index?: number; id?: string; name?: string; kind?: string; section?: string; headline?: string; shapes?: SnapShape[]; adjusted?: Omit<Adjustment, "slide">[] }
interface DeckSnapshot { name?: string; slideSize?: { w: number; h: number }; grid?: { cols: number; rows: number; cell: number }; style?: { mode?: string; name?: string }; slides?: SnapSlide[] }
interface DeckFinding { slide: number; blockId?: string; role?: string; kind: string; message: string; severity?: string }
interface DeckSummary { kind?: string; slides?: { index: number; tabId?: string; name?: string; headline?: string; shapes?: number; thumbnail?: string }[]; findings?: DeckFinding[]; warnings?: string[]; adjusted?: Adjustment[]; pdfPath?: string; snapshot?: DeckSnapshot }
interface SlideRow { index: number; name: string; headline?: string; shapeCount: number; findings: DeckFinding[]; adjusted: Adjustment[]; slide?: SnapSlide; tab?: Tab }

/**
 * What the builder moved: the build summary's own list, else the per-slide records the snapshot carries.
 * A record whose plan block already sits where the build put it is dropped, so writing the geometry back
 * into the plan is what clears the chip.
 */
function adjustments(b: DeckSummary, snap: DeckSnapshot | null, planOf: (slide: number) => Tab | undefined): Adjustment[] {
  const out: Adjustment[] = [];
  const seen = new Set<string>();
  const add = (a: Adjustment) => {
    if (!isCells(a?.from) || !isCells(a?.to)) return;
    const key = `${a.slide}:${a.blockId ?? ""}`;
    if (seen.has(key)) return;
    seen.add(key);
    out.push(a);
  };
  for (const a of b.adjusted ?? []) add(a);
  for (const s of snap?.slides ?? []) for (const a of s.adjusted ?? []) add({ ...a, slide: s.index ?? 0 });
  return out.filter((a) => {
    const block = planOf(a.slide)?.blocks.find((x) => x.id === a.blockId);
    return !block || !sameCells(block, a.to);
  }).sort((x, y) => x.slide - y.slide);
}

const SVGNS = "http://www.w3.org/2000/svg";
function svgEl<K extends keyof SVGElementTagNameMap>(tag: K, attrs: Record<string, string | number | undefined> = {}, text?: string): SVGElementTagNameMap[K] {
  const e = document.createElementNS(SVGNS, tag);
  for (const [k, v] of Object.entries(attrs)) if (v != null) e.setAttribute(k, String(v));
  if (text != null) e.textContent = text;
  return e;
}

/** Slide-point geometry for the preview: the role tag's height, and the floor a thumbnail's display type is lifted to. */
const TAG_PT = 10, THUMB_MIN_PT = 52;
/** Roles the builder draws as the slide's ink; a thumbnail shows only these. */
const DISPLAY_ROLES = new Set(["headline", "statement", "coverTitle", "dividerTitle", "agendaStatement", "takeaway", "quote", "statValue"]);
const QUIET_ROLES = new Set(["kicker", "wordmark", "agendaEyebrow", "caption", "source", "attribution", "coverDate", "coverSubtitle", "dividerSubtitle", "statLabel", "cardSub"]);
/** Block types the builder draws at the style's slot or not at all, so their cells never match the plan. */
const FURNITURE_TYPES = new Set(["headline", "kicker", "notes"]);

function inkClass(role: string): string {
  if (role === "statValue") return "val";
  if (role === "pageNumber") return "page";
  if (DISPLAY_ROLES.has(role)) return "hl";
  if (QUIET_ROLES.has(role)) return "kick";
  return "";
}

/** Greedy wrap at the characters per line the deck builder measures with (SPEC-deck, "Text fit", k = 0.5). */
function fitLines(text: string, widthPt: number, sizePt: number, maxLines: number): string[] {
  const per = Math.max(3, Math.floor(widthPt / (sizePt * 0.5)));
  const words = text.split(/\s+/).filter(Boolean);
  const lines: string[] = []; let line = "";
  for (const w of words) {
    const next = line ? `${line} ${w}` : w;
    if (next.length > per && line) { if (lines.length + 1 >= maxLines) { lines.push(`${line}…`); return lines; } lines.push(line); line = w; }
    else line = next;
  }
  if (line) lines.push(line);
  return lines;
}

/** The shapes of one plan block, as the rectangle that covers them all, in slide points. */
function blockBoxes(shapes: SnapShape[]): Map<string, { x: number; y: number; w: number; h: number }> {
  const out = new Map<string, { x: number; y: number; w: number; h: number }>();
  for (const s of shapes) {
    if (!s.blockId) continue;
    const cur = out.get(s.blockId);
    if (!cur) { out.set(s.blockId, { x: s.x, y: s.y, w: s.w, h: s.h }); continue; }
    const x = Math.min(cur.x, s.x), y = Math.min(cur.y, s.y);
    out.set(s.blockId, { x, y, w: Math.max(cur.x + cur.w, s.x + s.w) - x, h: Math.max(cur.y + cur.h, s.y + s.h) - y });
  }
  return out;
}

/**
 * One slide as vector: a box per drawn shape where it sits on the slide, the role tag and the first words
 * at slide scale, and a hit rectangle per plan block. A thumbnail keeps the geometry and the display type
 * only, with that type lifted to a floor so a 96 px preview still reads.
 */
function slideSvg(slide: SnapSlide, size: { w: number; h: number }, opts: { detail: "thumb" | "large"; diffs?: Set<string>; marks?: Set<string>; adjusted?: Set<string>; titles?: Map<string, string>; onBlock?: (blockId: string, anchor: SVGElement) => void } = { detail: "thumb" }): SVGSVGElement {
  const thumb = opts.detail === "thumb";
  const kind = slide.kind ?? "content";
  const svg = svgEl("svg", { class: `sprev ${opts.detail} k-${kind}`, viewBox: `0 0 ${size.w} ${size.h}`, preserveAspectRatio: "xMidYMid meet" });
  svg.append(svgEl("rect", { class: "bg", x: 0, y: 0, width: size.w, height: size.h }));
  const shapes = (slide.shapes ?? []).filter((s) => s.kind !== "notes");
  for (const s of shapes) {
    const k = s.kind ?? "text";
    if (k === "rule") { svg.append(svgEl("rect", { class: "shp rule", x: s.x, y: s.y, width: s.w, height: Math.max(1.5, s.h) })); continue; }
    if (s.blockId) svg.append(svgEl("rect", { class: `shp k-${k}`, x: s.x, y: s.y, width: s.w, height: s.h, rx: 2 }));
    const text = (s.text ?? "").trim();
    if (!text) continue;
    const cls = inkClass(s.role ?? "");
    if (thumb && !(cls === "hl" || cls === "val")) continue;
    const pt = thumb ? Math.max(s.fontPt ?? 14, THUMB_MIN_PT) : s.fontPt ?? 14;
    const pad = Math.min(8, s.w * 0.05);
    const tag = !thumb && !!s.blockId && !!s.role;
    if (tag) svg.append(svgEl("text", { class: "tag", x: s.x + pad, y: s.y + TAG_PT, "font-size": TAG_PT }, (s.role ?? "").replace(/([a-z])([A-Z])/g, "$1 $2").toUpperCase()));
    const top = s.y + (tag ? TAG_PT + 4 : 0);
    const room = Math.max(1, Math.floor((s.h - (top - s.y)) / (pt * 1.16)));
    const lines = fitLines(text, s.w - pad * 2, pt, Math.min(room, thumb ? 2 : 4));
    lines.forEach((ln, i) => svg.append(svgEl("text", { class: `tx ${cls}`, x: s.x + pad, y: top + pt * (0.84 + i * 1.16), "font-size": pt }, ln)));
  }
  if (thumb) return svg;
  for (const [id, r] of blockBoxes(shapes)) {
    const rect = svgEl("rect", { class: `svblk${opts.diffs?.has(id) ? " diff" : ""}${opts.adjusted?.has(id) ? " adj" : ""}${opts.marks?.has(id) ? " mark" : ""}`, x: r.x - 2, y: r.y - 2, width: r.w + 4, height: r.h + 4, rx: 3 });
    rect.append(svgEl("title", {}, opts.titles?.get(id) ?? id));
    if (opts.onBlock) rect.addEventListener("click", () => opts.onBlock!(id, rect));
    svg.append(rect);
  }
  return svg;
}

/**
 * Built geometry against the plan, on the slide grid rather than in cells. `detected.range` on an imported
 * deck is A1-style over the 32 columns, so it parses like a sheet address.
 */
function diffDeck(plan: Block[], shapes: SnapShape[], cell: { w: number; h: number }): Diff[] {
  const boxes = blockBoxes(shapes);
  const cellsOf = (r: { x: number; y: number; w: number; h: number }) => ({ row: Math.round(r.y / cell.h) + 1, col: Math.round(r.x / cell.w) + 1, rows: Math.max(1, Math.round(r.h / cell.h)), cols: Math.max(1, Math.round(r.w / cell.w)) });
  const out: Diff[] = [];
  for (const pb of plan) {
    if (FURNITURE_TYPES.has(pb.type)) continue;
    const box = boxes.get(pb.id);
    if (!box) { out.push({ id: pb.id, text: `${pb.title}: planned, not built` }); continue; }
    const g = cellsOf(box);
    if (g.row !== pb.row || g.col !== pb.col) out.push({ id: pb.id, text: `${pb.title}: built at r${g.row} c${g.col}, planned r${pb.row} c${pb.col}` });
    else if (g.rows !== pb.rows || g.cols !== pb.cols) out.push({ id: pb.id, text: `${pb.title}: built ${g.rows}×${g.cols}, planned ${pb.rows}×${pb.cols}` });
    if (pb.detected?.range) {
      const [a] = pb.detected.range.split(":"); const o = parseAddr(a);
      if (o.row && (o.row !== g.row || o.col !== g.col)) out.push({ id: pb.id, text: `${pb.title}: was ${pb.detected.range}, now r${g.row} c${g.col}` });
    }
  }
  for (const id of boxes.keys()) if (!plan.some((b) => b.id === id)) out.push({ id, text: `${id}: built, not planned` });
  return out;
}

function renderDeckBuilt(c: Controller, root: HTMLElement) {
  const p = c.project;
  const b = (c.state!.build ?? p.buildSummary ?? {}) as DeckSummary;
  // The build passes the snapshot with its result; a re-render only has the copy the summary carries.
  const snap = (c.state!.snapshot ?? b.snapshot ?? null) as DeckSnapshot | null;
  const size = snap?.slideSize ?? { w: 960, h: 540 };
  const grid = snap?.grid ?? { cols: 32, rows: 18, cell: 30 };
  const cell = { w: size.w / grid.cols, h: size.h / grid.rows };
  const findings = b.findings ?? [];
  const pdfPath = b.pdfPath ?? (p as { pdfPath?: string }).pdfPath;

  const snapSlides = snap?.slides ?? [];
  const sumSlides = b.slides ?? [];
  const rows: SlideRow[] = Array.from({ length: Math.max(snapSlides.length, sumSlides.length) }, (_, i) => {
    const s = snapSlides[i];
    const sum = sumSlides.find((x) => x.index === (s?.index ?? i + 1)) ?? sumSlides[i];
    const index = s?.index ?? sum?.index ?? i + 1;
    const tab = p.tabs.find((t) => t.id === (s?.id ?? sum?.tabId)) ?? p.tabs[i];
    const drawn = (s?.shapes ?? []).filter((x) => x.kind !== "notes");
    return { index, name: s?.name ?? sum?.name ?? tab?.name ?? `Slide ${index}`, headline: s?.headline ?? sum?.headline, shapeCount: s ? drawn.length : sum?.shapes ?? 0, findings: findings.filter((f) => f.slide === index), adjusted: [], slide: s, tab };
  });
  const adjusted = adjustments(b, snap, (slide) => rows.find((r) => r.index === slide)?.tab);
  for (const r of rows) r.adjusted = adjusted.filter((a) => a.slide === r.index);
  const titleOf = (a: Adjustment) => rows.find((r) => r.index === a.slide)?.tab?.blocks.find((x) => x.id === a.blockId)?.title ?? a.blockId ?? "block";
  let idx = Math.max(0, rows.findIndex((r) => r.tab?.id === c.selectedTab));
  let mark: string | null = null; // the shape a finding pointed at, cleared by the next hand-picked slide

  const keepPositions = () => {
    const bySlide = new Map<number, Adjustment[]>();
    for (const a of adjusted) bySlide.set(a.slide, [...(bySlide.get(a.slide) ?? []), a]);
    keepGeometry(c, [...bySlide].map(([slide, moves]) => ({ tab: rows.find((r) => r.index === slide)?.tab, moves })), "slide");
  };
  /** The chip's lines, built at draw time so each one can walk to the slide it names. */
  const adjRows = (): AdjRow[] => adjusted.map((a) => ({
    ix: String(a.slide), title: titleOf(a), from: a.from, to: a.to, reason: a.reason,
    pick: () => select(rows.findIndex((r) => r.index === a.slide), a.blockId ?? null),
  }));

  const main = h("div", { class: "main" });
  const side = h("div", { class: "side" });
  root.append(h("div", { class: "body" }, main, side));

  const pane = h("div", { class: "pane" });
  const strip = h("div", { class: "slidestrip" }, ...rows.map((r, i) => h("button", { class: `sl${i === idx ? " on" : ""}`, title: `${r.index}. ${r.name}`, onclick: () => select(i) },
    r.slide ? slideSvg(r.slide, size, { detail: "thumb" }) : h("span", { class: "ix num" }, String(r.index)),
    r.slide ? h("span", { class: "ix num" }, String(r.index)) : null,
    r.findings.length ? h("span", { class: `dot ${r.findings.some((f) => f.severity === "error") ? "bad" : "warn"}` }) : null)));
  if (rows.length) main.append(strip);
  main.append(pane);

  const drawSlide = () => {
    clear(pane);
    const r = rows[idx];
    if (!r?.slide) {
      pane.append(h("div", { class: "center" }, h("div", {}, "Built"), h("div", { class: "mono xs muted" }, p.buildPath ?? ""), h("div", { class: "muted small-text" }, "Rebuild for a preview")));
      return;
    }
    const shapes = (r.slide.shapes ?? []).filter((x) => x.kind !== "notes");
    const diffs = diffDeck(r.tab?.blocks ?? [], shapes, cell);
    const titles = new Map((r.tab?.blocks ?? []).map((bl) => [bl.id, bl.title]));
    const toolbar = h("div", { class: "toolbar" },
      h("span", { class: "chip line num" }, `${r.shapeCount} shape${r.shapeCount === 1 ? "" : "s"}`),
      r.findings.length ? h("span", { class: "chip warn num", title: r.findings.map((f) => `${f.kind}: ${f.message}`).join("\n") }, h("span", { class: `dot ${r.findings.some((f) => f.severity === "error") ? "bad" : "warn"}` }), `${r.findings.length} problem${r.findings.length === 1 ? "" : "s"}`) : null,
      adjustedChip(adjRows(), keepPositions),
      h("span", { class: "purpose", title: r.headline }, r.headline ?? r.name),
      diffs.length ? h("span", { class: "chip warn", title: diffs.map((d) => d.text).join("\n") }, h("span", { class: "dot warn" }), `${diffs.length} off plan`) : h("span", { class: "chip ok" }, h("span", { class: "dot ok" }), "matches plan"));
    const svg = slideSvg(r.slide, size, {
      detail: "large",
      diffs: new Set(diffs.map((d) => d.id)),
      marks: mark ? new Set([mark]) : undefined,
      adjusted: new Set(r.adjusted.map((a) => a.blockId).filter((x): x is string => !!x)),
      titles,
      onBlock: (id, anchor) => askExplain(c, anchor, titles.get(id) ?? id, `Explain "${titles.get(id) ?? id}" on slide ${r.index} of the built deck: what it holds, how it serves the slide's message, and anything that does not. Keep it short.`),
    });
    const slide = h("div", { class: "sslide", style: { width: `min(100cqw, calc(100cqh * ${size.w / size.h}))`, aspectRatio: `${size.w} / ${size.h}` } }, svg);
    const stage = h("div", { class: "canvaswrap sstage" }, slide);
    pane.append(toolbar, stage);
    fadeIn(slide);
  };

  const select = (i: number, markId: string | null = null) => {
    idx = Math.max(0, Math.min(rows.length - 1, i));
    mark = markId;
    strip.querySelectorAll("button.sl").forEach((el, j) => el.classList.toggle("on", j === idx));
    side.querySelectorAll(".sheetrow").forEach((el, j) => el.classList.toggle("sel", j === idx));
    strip.querySelectorAll("button.sl")[idx]?.scrollIntoView({ block: "nearest", inline: "nearest" });
    if (rows[idx]?.tab) c.selectedTab = rows[idx].tab!.id;
    drawSlide();
  };

  // ---- side: file, slides, one status line, warnings behind a count, the thread ledger ----
  const checks = rows.reduce((n, r) => n + r.shapeCount, 0);
  const findingList = h("ul", { class: "warnlist", hidden: true }, ...findings.slice(0, 40).map((f) => h("li", { class: "link", title: f.role, onclick: () => { const i = rows.findIndex((r) => r.index === f.slide); select(i >= 0 ? i : idx, f.blockId ?? null); } },
    h("span", { class: `dot ${f.severity === "error" ? "bad" : "warn"}` }),
    h("span", {}, h("span", { class: "mono" }, `Slide ${f.slide}`), ` ${f.kind}: ${f.message}`))));
  const statusLine = findings.length
    ? statusButton(`${findings.length} problem${findings.length === 1 ? "" : "s"}`, findingList)
    : h("div", { class: "inline tight small-text" }, h("span", { class: "dot ok" }), h("span", { class: "muted" }, `${checks} checks, no problems`));

  const body = h("div", { class: "scroll" },
    h("div", { class: "field" }, h("span", { class: "label" }, "Slides"),
      h("div", { class: "rows" }, ...rows.map((r, i) => h("div", { class: `sheetrow deck link${i === idx ? " sel" : ""}`, onclick: () => select(i) },
        h("span", { class: "ix num" }, String(r.index)),
        h("span", { class: "t", title: r.headline }, r.name),
        h("span", { class: "n" }, `${r.shapeCount} shapes`),
        h("span", { class: `n${r.findings.length ? " warn" : ""}` }, r.findings.length ? String(r.findings.length) : ""))))),
    h("div", { class: "field" }, h("span", { class: "label" }, "Checks"), statusLine, findingList),
    warningsField(b.warnings ?? []),
    threadField(p));

  side.append(fileHeader(c, p.buildPath ?? "", [pdfPath ? h("span", { class: "chip line", title: pdfPath }, "PDF") : null], pdfPath), body);
  slideIn(side);
  drawSlide();

  root.append(reopenFooter(c, `${rows.length} slide${rows.length === 1 ? "" : "s"} · ${checks} shapes`));

  document.onkeydown = (e) => {
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    if ((e.target as HTMLElement).closest("input, textarea, select")) return;
    if (e.key === "ArrowLeft") { e.preventDefault(); select(idx - 1); }
    else if (e.key === "ArrowRight") { e.preventDefault(); select(idx + 1); }
  };
}

// ---------------------------------------------------------------- shared plan diff

/** Block-to-block dependencies on one sheet, from formula references in the snapshot. Cross-sheet refs count as "reads other sheet". */
function dependencies(sheet: SnapSheet): { from: string; to: string; cross: boolean }[] {
  const blocks = sheet.blocks ?? [];
  const owner = (row: number, col: number) => blocks.find((b) => row >= b.row && row < b.row + b.rows && col >= b.col && col < b.col + b.cols)?.id;
  const edges = new Map<string, { from: string; to: string; cross: boolean }>();
  const REF = /(?:'([^']+)'|([A-Za-z0-9_]+))?!?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/g;
  for (const cell of sheet.cells) {
    if (cell.k !== "formula" || typeof cell.v !== "string") continue;
    const { row, col } = parseAddr(cell.a); const from = owner(row, col); if (!from) continue;
    const f = cell.v;
    let m: RegExpExecArray | null; REF.lastIndex = 0;
    while ((m = REF.exec(f))) {
      const sheetName = m[1] ?? m[2];
      const qualified = !!sheetName && m[0].includes("!");
      if (qualified && sheetName !== sheet.name) { edges.set(`${from}>x`, { from, to: "", cross: true }); continue; }
      const a = parseAddr(`${m[3]}${m[4]}`); const b = m[5] ? parseAddr(`${m[5]}${m[6]}`) : a;
      for (let r = a.row; r <= b.row; r++) for (let c = a.col; c <= b.col; c++) { const to = owner(r, c); if (to && to !== from) edges.set(`${from}>${to}`, { from, to, cross: false }); }
    }
  }
  return [...edges.values()];
}

interface Diff { id: string; text: string }
function diffBlocks(plan: { id: string; title: string; row: number; col: number; rows: number; cols: number; detected?: { range: string } }[], built: { id: string; row: number; col: number; rows: number; cols: number }[]): Diff[] {
  const out: Diff[] = [];
  for (const pb of plan) {
    const bb = built.find((x) => x.id === pb.id);
    if (!bb) { out.push({ id: pb.id, text: `${pb.title}: planned, not built` }); continue; }
    if (bb.row !== pb.row || bb.col !== pb.col) out.push({ id: pb.id, text: `${pb.title}: built at ${colLetter(bb.col)}${bb.row}, planned ${colLetter(pb.col)}${pb.row}` });
    else if (bb.rows !== pb.rows || bb.cols !== pb.cols) out.push({ id: pb.id, text: `${pb.title}: built ${bb.rows}×${bb.cols}, planned ${pb.rows}×${pb.cols}` });
    // imported projects: say where the block sat in the original file when it moved
    if (pb.detected?.range) {
      const [a] = pb.detected.range.split(":"); const o = parseAddr(a);
      if (o.row && (o.row !== bb.row || o.col !== bb.col)) out.push({ id: pb.id, text: `${pb.title}: was ${pb.detected.range}, now ${colLetter(bb.col)}${bb.row}` });
    }
  }
  for (const bb of built) if (!plan.some((x) => x.id === bb.id)) out.push({ id: bb.id, text: `${bb.id}: built, not planned` });
  return out;
}

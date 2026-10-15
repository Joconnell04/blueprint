import { h, icon } from "./dom.js";
import { CELL_H, CELL_W, TYPE_LABEL, typeLabel, colLetter, type Block, type WorkbookBlockType } from "./model.js";
import { DEFAULT_SIZE } from "./layout-view.js";

/** A live sketch of the proposal while Claude is still streaming the tool call. Partial JSON is healed by the host, so treat every field as optional. */
export function renderStreaming(partial: { tool: string; args: any }, root: HTMLElement) {
  const a = partial.args ?? {};
  root.append(h("div", { class: "hdr" }, h("div", { class: "logo" }, icon("logo", 12)), h("span", { class: "name" }, "Claude is drafting"), h("span", { class: "chip warn" }, h("span", { class: "dot warn" }), "streaming"), h("span", { class: "spacer" })));
  if (Array.isArray(a.tabs)) {
    const list = h("div", { class: "tablist" }, h("div", { class: "thead" }, h("div", { class: "grip-space" }), h("div", { class: "idx label" }, "#"), h("div", { class: "grow label" }, "Tab")));
    a.tabs.forEach((t: any, i: number) => list.append(h("div", { class: "row trow" }, h("div", { class: "grip" }), h("div", { class: "idx" }, String(i + 1)), h("div", { class: "fields" }, h("div", { class: "nm trunc" }, String(t?.name ?? "…")), h("div", { class: "muted small-text trunc" }, String(t?.purpose ?? ""))))));
    root.append(h("div", { class: "body" }, h("div", { class: "main" }, list)));
    return;
  }
  const layouts: any[] = Array.isArray(a.layouts) ? a.layouts : [];
  const current = layouts[layouts.length - 1];
  const strip = h("div", { class: "tabstrip" }, ...layouts.map((l, i) => h("button", { class: i === layouts.length - 1 ? "on" : "" }, h("span", { class: "dot none" }), String(l?.tab ?? "…"))));
  const blocks: Block[] = [];
  let nextRow = 2;
  for (const raw of (current?.blocks ?? []) as any[]) {
    const type = (Object.keys(TYPE_LABEL).includes(raw?.type) ? raw.type : "table") as WorkbookBlockType;
    const [rows, cols] = DEFAULT_SIZE[type];
    const b: Block = { id: `s${blocks.length}`, type, title: String(raw?.title ?? "…"), row: raw?.row ?? nextRow, col: raw?.col ?? 2, rows: raw?.rows ?? Math.max(rows, (raw?.rowLabels?.length ?? 0) + 2), cols: raw?.cols ?? ((type === "table" || type === "calc") && raw?.columns?.length ? Math.min(25, Math.max(3, raw.columns.length)) : Math.max(cols, raw?.columns?.length ?? 0)), columns: raw?.columns, rowLabels: raw?.rowLabels };
    nextRow = Math.max(nextRow, b.row + b.rows + 1);
    blocks.push(b);
  }
  const maxCol = 20, maxRow = Math.max(30, nextRow + 4);
  const canvas = h("div", { class: "canvas", style: { width: `${30 + maxCol * CELL_W}px`, height: `${20 + maxRow * CELL_H}px` } },
    h("div", { class: "grid", style: { left: "30px", top: "20px", width: `${maxCol * CELL_W}px`, height: `${maxRow * CELL_H}px`, backgroundSize: `${CELL_W}px ${CELL_H}px` } }),
    h("div", { class: "colhdr", style: { marginLeft: "30px", width: `${maxCol * CELL_W}px` } }, ...Array.from({ length: maxCol }, (_, i) => h("div", { style: { width: `${CELL_W}px` } }, colLetter(i + 1)))),
    h("div", { class: "rowhdr", style: { position: "absolute", top: "20px", left: 0, height: `${maxRow * CELL_H}px` } }, ...Array.from({ length: maxRow }, (_, i) => h("div", { style: { height: `${CELL_H}px`, lineHeight: `${CELL_H}px` } }, String(i + 1)))),
    h("div", { class: "corner", style: { position: "absolute", top: 0, left: 0 } }));
  for (const b of blocks) canvas.append(h("div", { class: `blk ${b.type} ghost`, style: { left: `${30 + (b.col - 1) * CELL_W}px`, top: `${20 + (b.row - 1) * CELL_H}px`, width: `${b.cols * CELL_W - 1}px`, height: `${b.rows * CELL_H - 1}px` } },
    h("div", { class: "bh" }, h("span", { class: `ty ${b.type}` }, typeLabel(b.type)), h("span", { class: "t" }, b.title)),
    h("div", { class: "bc ghostbody" }, (b.columns ?? b.rowLabels ?? []).join(" · "))));
  root.append(h("div", { class: "body" }, h("div", { class: "main" }, strip, h("div", { class: "canvaswrap" }, canvas))));
}

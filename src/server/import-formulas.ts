import fs from "node:fs";

/**
 * Translate formulas from an imported workbook into the builder's block references, using the scanner's
 * geometry: a value cell of an inputs/output block becomes {{id!Label}}, a data cell of a table/calc block
 * becomes {{id!r,c}} (zero-based within the data area), a column of data becomes {{id!col:c}}. References
 * to cells outside every detected block are left as they were, so nothing is silently dropped.
 */

interface ScanBlock { id: string; type: string; sample?: (string | number | null)[][]; rowLabels?: string[]; columns?: string[]; note?: string; rowNotes?: (string | null)[]; geometry?: { titleRow: number; headerRows: number[]; labelCol: number; contentCols: number[]; bodyRows: number[]; totalRow: number | null; subtotalRows?: number[]; box: [number, number, number, number] } }
interface ScanTab { name: string; blocks: ScanBlock[] }
export interface Scan { tabs: ScanTab[] }

export interface BlockNotes { description?: string; rowNotes?: (string | null)[]; totalNote?: string }

/**
 * U6: the scanner writes one note per occupied row of a block, the header and total rows included, so a
 * naive read puts every note one row out. Split them by what the row is instead: the block's own note and
 * the header row's become the description, the total row's becomes totalNote, the rest line up with the
 * data rows, which is the order `sample` comes back in. A rebuild then writes each note where it was.
 */
export function blockNotes(b: ScanBlock): BlockNotes {
  const header = [b.note?.trim()].filter(Boolean) as string[];
  const raw = (b.rowNotes ?? []).map((n) => (typeof n === "string" ? n.trim() : ""));
  const g = b.geometry;
  if (!raw.some(Boolean) || !g) {
    const rowNotes = raw.some(Boolean) ? raw.map((n) => n || null) : undefined;
    return { description: header.join("; ") || undefined, rowNotes };
  }
  const total = g.totalRow ?? undefined;
  const rows = [...new Set([...(g.headerRows ?? []), ...(g.bodyRows ?? []), ...(g.subtotalRows ?? []), ...(total != null ? [total] : [])])]
    .filter((r) => r > g.box[0])
    .sort((x, y) => x - y);
  const byRow = new Map<number, string>();
  rows.forEach((r, i) => { if (raw[i]) byRow.set(r, raw[i]); });
  for (const r of g.headerRows ?? []) { const v = byRow.get(r); if (v) header.push(v); }
  const rowNotes = (g.bodyRows ?? []).map((r) => byRow.get(r) ?? null);
  return {
    description: header.join("; ") || undefined,
    rowNotes: rowNotes.some(Boolean) ? rowNotes : undefined,
    totalNote: total != null ? byRow.get(total) : undefined,
  };
}

function colNum(letters: string): number { let n = 0; for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
function colLetters(n: number): string { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

type CellRef = { ref: string; blockId?: string };

/** Map (sheet, row, col) to a block reference, or undefined when no block owns the cell. */
export function cellToRef(scan: Scan, sheet: string, row: number, col: number): CellRef | undefined {
  const tab = scan.tabs.find((t) => t.name === sheet); if (!tab) return undefined;
  for (const b of tab.blocks) {
    const g = b.geometry; if (!g) continue;
    const [r0, c0, r1, c1] = g.box; if (row < r0 || row > r1 || col < c0 || col > c1) continue;
    const bodyIdx = g.bodyRows.indexOf(row);
    if (b.type === "inputs" || b.type === "output") {
      if (bodyIdx >= 0 && col === g.labelCol + 1) { const label = b.rowLabels?.[bodyIdx] ?? b.sample?.[bodyIdx]?.[0]; if (label != null && String(label).trim()) return { ref: `{{${b.id}!${String(label).trim()}}}`, blockId: b.id }; }
      return { ref: `${colLetters(col)}${row}`, blockId: b.id };
    }
    if (b.type === "table" || b.type === "calc") {
      const ci = g.contentCols.indexOf(col);
      if (bodyIdx >= 0 && ci >= 0) return { ref: `{{${b.id}!${bodyIdx},${ci}}}`, blockId: b.id };
      if (g.totalRow === row && ci >= 0) return { ref: `{{${b.id}!total:${ci}}}`, blockId: b.id };
      return { ref: `${colLetters(col)}${row}`, blockId: b.id };
    }
    return { ref: `${colLetters(col)}${row}`, blockId: b.id };
  }
  return undefined;
}

const REF_RE = /(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))?(!?)\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/g;

/** Rewrite every cell and range reference in `formula` (written on `sheet`) into block references where possible. */
export function translateFormula(formula: string, sheet: string, scan: Scan): { formula: string; unresolved: string[] } {
  const unresolved: string[] = [];
  const out = formula.replace(REF_RE, (m, q, bare, bang, c1, r1, c2, r2) => {
    const hasSheet = bang === "!";
    if (!hasSheet && (q || bare)) return m; // a function name or similar glued to a cell-looking token, e.g. SUM(A1 is split by the paren, but LOG10 must not be touched
    const target = hasSheet ? (q ? q.replace(/''/g, "'") : bare) : sheet;
    const a = cellToRef(scan, target, Number(r1), colNum(c1));
    if (!c2) { if (!a) { unresolved.push(m); return m; } return a.ref; }
    // a range: whole data column of one block, else leave as is
    const b = cellToRef(scan, target, Number(r2), colNum(c2));
    if (a && b && a.blockId === b.blockId && colNum(c1) === colNum(c2)) {
      const mA = /^\{\{([^!]+)!(\d+),(\d+)\}\}$/.exec(a.ref), mB = /^\{\{([^!]+)!(\d+),(\d+)\}\}$/.exec(b.ref);
      if (mA && mB && mA[3] === mB[3]) { const tab = scan.tabs.find((t) => t.name === target); const blk = tab?.blocks.find((x) => x.id === mA[1]); const n = blk?.geometry?.bodyRows.length ?? 0; if (Number(mA[2]) === 0 && Number(mB[2]) === n - 1) return `{{${mA[1]}!col:${mA[3]}}}`; }
    }
    unresolved.push(m); return m;
  });
  return { formula: out, unresolved };
}

export interface OriginalBlock extends BlockNotes { tab: string; columns?: string[]; rows: (string | number | null)[][]; unresolved: string[] }

/** Per-block original content for a rebuild: columns/labels, the notes that were on it, and rows with formulas translated. */
export function originalContent(scanPath: string): Record<string, OriginalBlock> {
  const scan = JSON.parse(fs.readFileSync(scanPath, "utf8")) as Scan;
  const out: Record<string, OriginalBlock> = {};
  for (const t of scan.tabs) for (const b of t.blocks) {
    if (!b.sample?.length) continue;
    const unresolved = new Set<string>();
    const rows = b.sample.slice(0, 60).map((r) => r.map((v) => {
      if (typeof v === "string" && v.startsWith("=")) { const tr = translateFormula(v, t.name, scan); tr.unresolved.forEach((u) => unresolved.add(u)); return tr.formula; }
      return v;
    }));
    out[b.id] = { tab: t.name, columns: b.columns, rows, unresolved: [...unresolved], ...blockNotes(b) };
  }
  return out;
}

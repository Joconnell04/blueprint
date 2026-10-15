/** Blueprint plans two kinds of file: a workbook (.xlsx) or a deck (.pptx). Same three layers in both. */
export type Kind = "workbook" | "deck";

export type WorkbookBlockType = "inputs" | "table" | "calc" | "output" | "chart" | "text";
/** Deck blocks (spec section 3). text, table and chart are shared with the workbook list; `kind` says how they render and build. */
export type DeckBlockType =
  | "headline" | "kicker" | "text" | "bullets" | "cards" | "stat"
  | "chart" | "diagram" | "table" | "image" | "quote" | "takeaway" | "placeholder" | "notes";
export type BlockType = WorkbookBlockType | DeckBlockType;

/** Build a map over one vocabulary with these rather than over BlockType, which holds both. */
export const WORKBOOK_BLOCK_TYPES: WorkbookBlockType[] = ["inputs", "table", "calc", "output", "chart", "text"];
export const DECK_BLOCK_TYPES: DeckBlockType[] = ["headline", "kicker", "text", "bullets", "cards", "stat", "chart", "diagram", "table", "image", "quote", "takeaway", "placeholder", "notes"];
/** What Claude and the canvas may place. `headline` is the slide's own field, never a block (feedback F14). */
export const DECK_PLACEABLE_TYPES: DeckBlockType[] = DECK_BLOCK_TYPES.filter((t) => t !== "headline");

/** A slide is 960x540 pt on a 30 pt grid: 32 columns by 18 rows. Shared by the server and the canvas. */
export const DECK_COLS = 32;
export const DECK_ROWS = 18;

/**
 * What the builder may do between the approved plan and the built file (docs/research/output-quality.md,
 * "The fidelity setting"). Absent means every default: balanced, notes from the story, a cover, an agenda.
 */
export type Fidelity = "strict" | "balanced" | "loose";
export interface BuildOptions {
  fidelity?: Fidelity;
  notesFromStory?: boolean;  // every content slide's speaker notes default to its story
  cover?: boolean;           // generate a cover slide when the outline has none
  agenda?: boolean;          // generate an agenda from the divider list when the outline has none
  comments?: boolean;        // workbook: write the closed threads into the file as cell comments (N9)
}
export const BUILD_DEFAULTS: Required<BuildOptions> = { fidelity: "balanced", notesFromStory: true, cover: true, agenda: true, comments: false };
/** One sentence per value, shared by the tool schemas and the canvas popover. */
export const FIDELITY_HELP = "strict keeps the approved geometry exactly. balanced lets the builder resize within two cells, balance whitespace and align edges. loose lets it re-pick the layout from the content.";

export type NoteStatus = "open" | "resolved" | "declined";
export type Formatting = "plain" | "house" | "template" | "skill";

/** The brief card above the outline: who it is for, what it must do, the one thing to remember. */
export interface Brief {
  audience?: string;   // who reads it
  goal?: string;       // what it has to do for them
  takeaway?: string;   // the one sentence they should remember, optional
  /** @deprecated Emphasis chips were misread as process notes (feedback F7). Old plans keep theirs; nothing writes or reads them. */
  emphasis?: string[];
}

/** A diagram block's own shape (feedback F18): the builder draws nodes and connectors from the style tokens. */
export type DiagramKind = "flow" | "cycle" | "timeline" | "tree" | "matrix";
export interface DiagramNode { id: string; label: string }
export interface DiagramEdge { from: string; to: string; label?: string }
export interface Diagram {
  kind: DiagramKind;
  direction?: "LR" | "TB";
  nodes: DiagramNode[];
  edges: DiagramEdge[];
}

export interface Block {
  id: string;
  type: BlockType;
  title: string;
  row: number; // workbook: 1-based Excel row. deck: 1-based row on the 18-row slide grid
  col: number; // workbook: 1-based Excel column (1 = A). deck: 1-based column on the 32-column slide grid
  rows: number; // planned height in cells
  cols: number; // planned width in cells
  columns?: string[]; // table/calc header labels; inputs/output: label names
  rowLabels?: string[]; // inputs/output labels or table row labels (planning only)
  sample?: (string | number)[][]; // fake placeholder rows for the canvas
  description?: string;
  content?: string; // deck: one line on what goes here, not the content itself
  emphasis?: boolean; // deck: this block is what should pop on the slide
  count?: number; // deck: how many cards or stats
  source?: string; // chart: id of the block it plots
  chartKind?: "bar" | "line" | "pie"; // chart: how it plots (default bar)
  diagram?: Diagram; // diagram: the nodes and edges to draw
  detected?: { confidence: number; signals: string[]; range: string }; // set by import_workbook / import_deck
}

/** One outline row: a tab in a workbook, a slide in a deck. */
export interface Tab {
  id: string;
  name: string;
  purpose: string; // workbook: what the tab holds and what it feeds
  message?: string; // the headline: one line, the point of this tab or slide. Sentence case
  story?: string; // deck: 2 to 4 sentences on what the slide does for the reader
  divider?: boolean; // deck: this row is a section divider. Its name is the section title; no story, no message, no blocks
  agendaLine?: string; // deck, divider rows only: the one line the generated agenda prints under this section title
  kicker?: string; // deck: eyebrow above the headline; the section name when empty
  section?: string; // derived on every save from the nearest divider above; nothing else writes it
  approved: boolean;
  layoutProposed: boolean;
  layoutApproved: boolean;
  blocks: Block[];
}

/**
 * One thread entry. Two kinds share the lifecycle:
 * - instruction: written by the user, Claude resolves or declines it with a response.
 * - question: written by Claude, the user answers it (response) which closes it; the answer is an instruction to honour.
 */
export interface Note {
  id: string;
  tabId: string;
  blockId?: string;
  kind: "instruction" | "question";
  text: string;
  status: NoteStatus;
  response?: string;
  createdBy: "user" | "claude";
  createdAt: string;
  resolvedAt?: string;
}

export interface Project {
  id: string;
  name: string;
  kind: Kind;                  // workbook (.xlsx) or deck (.pptx); migrate() sets workbook on old plans
  brief: string;               // the user's ask, verbatim where possible
  briefCard?: Brief;           // the structured brief shown above the outline
  formatting: Formatting;      // what the Python builder applies mechanically
  styleId?: string;            // "plain" | "house" | "skill:<name>" | "template" (see styles.ts)
  templatePath?: string;
  fileName?: string;
  outputDir?: string;
  exportPdf?: boolean;         // deck: also write a .pdf when LibreOffice is present
  build?: BuildOptions;        // what the builder may change; absent means every default (BUILD_DEFAULTS)
  skipLayouts: boolean;
  tabsApproved: boolean;
  tabs: Tab[];
  notes: Note[];
  revision: number;
  createdAt: string;
  updatedAt: string;
  buildPath?: string;
  snapshotPath?: string;
  buildSummary?: unknown;
  renderSeq?: number;
  importedFrom?: string;
  scanPath?: string;           // import_workbook: the scanner's out.json, reused by restyle_workbook
  specPath?: string;           // the builder spec of the last build, so a rebuild can be diffed
  pdfPath?: string;            // deck: the PDF beside the .pptx, when LibreOffice wrote one
  thumbnails?: string[];       // deck: slide PNGs the built view shows
}

export type Stage = "brief" | "tabs" | "layouts" | "ready" | "built";

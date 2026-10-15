# Blueprint: design

An MCP App for Claude Desktop. Plan a workbook (.xlsx) or a deck (.pptx) with Claude in three layers
(outline, then each layout, then the build), annotate it on a canvas, approve it, then Claude builds the real file.

- Runs on the Claude subscription. The server never calls a model. All reasoning happens in the chat; the server stores state, renders the canvas, validates, and writes the file.
- Local only: a stdio MCP server spawned by Claude Desktop, state in `~/.blueprint/`.

## Stack

| Part | Choice | Why |
|---|---|---|
| MCP server | TypeScript, `@modelcontextprotocol/server` 2.0 + `@modelcontextprotocol/ext-apps` 2.0, stdio | The only SDK with first-class MCP Apps helpers; stdio is what Claude Desktop launches |
| Canvas UI | Vanilla TypeScript, Vite + `vite-plugin-singlefile`, one HTML resource | Smallest bundle, no framework lifecycle to fight inside the sandbox iframe, pointer-event drag works without libraries |
| Plan state | JSON files, one per project, `~/.blueprint/projects/<id>.json` | Survives restarts, inspectable, no DB |
| xlsx builder | Python 3 + openpyxl | ExcelJS cannot write charts; openpyxl can, and the user's xlsx skill already uses it |
| pptx builder | Python 3 + python-pptx | Native charts and tables, real masters, PDF and PNG through LibreOffice |

## Two modes

A plan carries a `kind`. Everything else is shared: the same tools, the same canvas, the same threads and approvals.

| | Workbook | Deck |
|---|---|---|
| Outline row | Tab: name, purpose, message | Slide: name (2 to 5 words of purpose), story, section, message |
| Layout grid | Real Excel cells | 960x540 pt slide on a 30 pt grid: 32 columns by 18 rows |
| Blocks | inputs, table, calc, output, chart, text | headline, kicker, text, bullets, cards, stat, chart, table, image, quote, takeaway, placeholder, notes |
| Layout tools | Stack, Flow, align, distribute | Remix (ten archetypes), align, distribute |
| Build | `build_workbook` then the formula check | `build_deck` then the deck check, plus an optional PDF |
| Import | `import_workbook` | `import_deck` |

## Data model

```
Project  { id, name, kind, brief, briefCard, formatting, styleId?, templatePath?, fileName?, outputDir?,
           exportPdf?, skipLayouts, tabsApproved, tabs: Tab[], notes: Note[], revision, buildPath?,
           snapshotPath?, buildSummary?, renderSeq?, importedFrom?, scanPath? }
Brief    { audience?, goal?, takeaway?, emphasis?: string[] }
Tab      { id, name, purpose, message?, story?, section?, approved, layoutProposed, layoutApproved, blocks: Block[] }
Block    { id, type, title, row, col, rows, cols, columns?, rowLabels?, sample?, description?,
           content? (deck: what goes here), emphasis? (deck), count? (cards/stats), source? + chartKind? (charts),
           detected?: { confidence, signals, range } (set by the importers) }
Note     { id, tabId, blockId?, kind: instruction|question, text, status: open|resolved|declined,
           response?, createdBy: user|claude, createdAt, resolvedAt? }
           instruction: written by the user, Claude resolves or declines it with a response
           question: written by Claude, the user answers inline; the answer reads back as an instruction to honour
```

`row`, `col`, `rows`, `cols` are 1-based cells in both modes: Excel cells for a workbook, slide-grid cells for a deck.
A block occupies exactly that rectangle, so what the user drags is what gets written.

Stage: `brief -> tabs -> layouts -> built`, derived from the approvals rather than stored. The outline approved moves
to layouts; all layouts approved (or `skipLayouts`) allows the build.

## Modules

| Path | Holds |
|---|---|
| `src/shared/types.ts` | Plan types and the deck grid constants, used by server and canvas |
| `src/server/store.ts` | Home dir, project JSON, outline and block helpers, default sizes per mode, readiness, the summary Claude reads |
| `src/server/common.ts` | Paths, render and fail helpers, zod schemas, the per-mode workflow text, Python runners, timing log |
| `src/server/tools-plan.ts` | start_project, list_projects, propose_outline, propose_layouts, get_plan, update_plan, show_plan, resolve_notes |
| `src/server/tools-build.ts`, `tools-build-deck.ts` | build_workbook, build_deck |
| `src/server/tools-import.ts`, `tools-import-deck.ts`, `restyle.ts` | import_workbook, import_deck, restyle_workbook |
| `src/server/tools-extract.ts`, `styles.ts` | extract_format, the format sources the picker lists, deck token sidecars |
| `src/server/tools-ui.ts`, `templates.ts` | Canvas-only tools, plan templates |
| `src/server/import-formulas.ts` | Translates imported formulas to block references |
| `src/app/*` | Canvas: controller (host bridge, polling, diffs), outline view, layout view, deck layout, remix, built view, popovers |
| `builder/build_xlsx.py`, `build_pptx.py` | The two builders and their snapshots; plain, house, skill and template stylers |
| `builder/import_scan.py`, `import_pptx.py`, `restyle.py` | Scanners and the in-place restyle |
| `builder/check_formulas.py`, `check_deck.py` | Formula check, deck check |
| `builder/extract_format.py` | Reads example files into style tokens and writes a SKILL.md |

## Tools (model-visible)

| Tool | Purpose |
|---|---|
| `start_project` | Create a plan of a `kind` from name, brief and brief card. Returns the workflow for that mode so Claude asks before proposing. |
| `propose_outline` | Replace the outline rows, tabs or slides. Rows matched by id or name keep their notes and layouts. Renders Layer 1. |
| `propose_layouts` | Set blocks for one or more rows, with positions, content lines and Claude's questions. Renders Layer 2. |
| `update_plan` | Targeted edits from chat: brief, message, story, section, rows, blocks, notes, skipLayouts, formatting, approvals. Re-renders nothing. |
| `get_plan` | Full current state: brief, rows, blocks, open threads, readiness warnings, and for a workbook the cell addresses each block will occupy. |
| `resolve_notes` | Close a thread with what was done, or decline it with a reason. |
| `build_workbook` | Claude supplies real content per block. The server validates against the approved plan, runs openpyxl, writes the .xlsx, renders the Built view. |
| `build_deck` | Claude supplies a headline per slide and real content per block. The server validates, runs python-pptx, writes the .pptx, the snapshot, the PNGs and any PDF. |
| `import_workbook` | Scan an existing .xlsx into a plan, block by block, with a confidence each. |
| `import_deck` | Scan an existing .pptx into a plan: slides to rows, shapes to blocks by role. |
| `extract_format` | Derive a formatting skill from example files and Claude's reading of screenshots (section below). |
| `restyle_workbook` | Reformat an imported workbook in place: formats only, values and formulas untouched. |
| `list_templates`, `save_template`, `apply_template` | Save an approved plan's structure and start the next one from it. |
| `show_plan` | Re-render the canvas at the current stage. |
| `list_projects` | Find an existing plan so a new chat can resume it. |

App-only (`visibility: ["app"]`, hidden from the model, called by the canvas): `ui_get_state`, `ui_save_tabs`,
`ui_save_brief`, `ui_save_layout`, `ui_note` (add, edit, delete, reopen, answer), `ui_approve`, `ui_styles`,
`ui_set_style`, `ui_templates`.

## Chat and canvas sync

- Canvas to Claude: every hand edit persists through an app-only tool, then the view pushes a compact plan summary with `updateModelContext` (gated on `getHostCapabilities()`). Approve buttons also `sendMessage` a short prompt, so the flow advances without the user typing. Tool descriptions tell Claude to call `get_plan` before proposing or building, so hand edits survive even if the host drops context updates.
- Claude to canvas: `propose_outline`, `propose_layouts`, `show_plan`, both builds, both imports and `apply_template` render a canvas; `update_plan` and `resolve_notes` do not. Every canvas polls `ui_get_state` every 2.5 s: a newer revision redraws with the changed blocks pulsing and a toast, and a newer render sequence collapses the older canvas to a one-line bar.

## Surfacing gaps before approval

The readiness panel lists, per row: blocks with open questions from Claude, blocks with no columns or no content line,
chart blocks with no source, rows with no blocks, slides with no headline, and open notes. Blocks with questions carry a
badge, approval is allowed with warnings but the count sits in the button label, and a missing takeaway is flagged once
for the plan rather than once per row.

## Skipping a layer

`skipLayouts: true` (from `start_project` or `update_plan`) lets a build run straight from the approved outline; Claude then supplies block positions in the build payload. A row with a proposed but unapproved layout still blocks the build.

## Build output

The workbook builder writes the .xlsx and a `.snapshot.json` beside it (cells with kind and display string, widths,
merges, block extents, charts). `check_formulas.py` then runs, recalculating with LibreOffice when it is installed and
with a bundled evaluator otherwise; findings come back by cell and block.

The deck builder writes the .pptx, a `.snapshot.json` (per slide: shapes with role, geometry, text preview and font
size), slide PNGs under `<name>.slides/`, and a PDF when `exportPdf` is set and LibreOffice is present. `check_deck.py`
reads the built file back for overflow, overlap, off-slide geometry, missing headlines and empty placeholders.

The Built view renders the snapshot: a grid with a plan-vs-build diff and dependency arrows for a workbook, a row and a vector thumbnail per slide for a deck. Both show one status line with the findings behind it.

## Formatting modes

- `plain`: bold titles, bold header row, thin borders, number formats, no colour. Decks get a neutral sans, charcoal ink, one accent.
- `house`: my spreadsheet format encoded in the builder (Arial 10, blue inputs, black formulas, grey italic notes, dashed label divider, navy title rules, Read Me cover), and a clean consulting default for decks (kicker, action title, cards).
- `skill`: a formatting skill found on this Mac. Its `format:` block drives the workbook styler and its `deck:` block the deck styler; the prose reaches Claude through `get_plan`. Schema in `docs/skill-format-block.md`.
- `template`: a file you supply. An .xlsx with a `Style` sheet of labelled example cells is cloned per role; a .pptx or .potx lends its master, its first layout and its placeholder positions.

The style picker scans `~/.claude/skills`, any local-agent-mode plugin skills and `$BLUEPRINT_SKILLS`, and tags each
entry workbook, deck or both from its front matter and description. Plugin skills are read-only, so their deck tokens
go to a sidecar at `~/.blueprint/styles/<slug>.json` that `styles.ts` merges in. `extract_format` writes both halves: a
`SKILL.md` under `~/.claude/skills/<slug>/` and, for decks, the sidecar the builder loads. A skill with prose but no
tokens builds in the house style, with a warning naming it.

## Notes in the finished file

Notes are agent instructions, not reader annotations, so they are never written into cells or onto slides. They live in
the plan JSON with their resolutions and show in the Built view. The house format's own note column is a separate thing,
filled from block content Claude supplies.

## Testing

1. Python unit suites per builder, scanner, checker and the extractor.
2. Script clients over stdio: a full workbook flow and a full deck flow, asserting the built file with openpyxl and python-pptx.
3. Layout units for the remix mapping, Playwright against the vendored basic-host, a visual pass over both modes and themes, then `docs/VERIFY.md` in Claude Desktop.

# Changelog

## 0.3.3

Backlog finished: every item except the Claude Desktop pass, which is a hand check.

Decks
- Style weights and bold per role (light headlines when the format says so), display-role floors at 0.85 × size, loose mode frees a boxed-in stat band, stat grid layout, table label-column token, kicker case token, source row kept clear.
- Chrome tokens (dots, lines, wordmarks, logos), cover/divider/agenda furniture derived from the style's own measure, sources shortened to one line, duplicate content dropped, parallel columns harmonised, chart discipline, overflow paragraphs moved to notes, template layouts mapped to blocks; stat rule position token; an empty slide is a check error.
- Extractor records the content measure (`furniture.body`) and per-role weights, so the column module and type match the real format; a test style was re-derived from a hand-built deck.

Workbooks
- Fidelity (strict / balanced): widen, generated totals, shift, note band, recorded as adjustments and shown in the Built view with Keep positions; house links are ink underlined, never green; comments export from resolved threads (off by default); toggle rows and base years keep their input role through the workflow text; scanner row notes survive a rebuild; the checker follows the underline rule.
- Skill prose is sent once per style and referenced after.

Canvas and platform
- Light theme rebuilt on real surface separation with AA semantics; link, node and edge cells styled; clipboard permission on the resource.
- App-side tools (focus_slide, select_block, scroll_to, show_view) so Claude can drive the open canvas; host style variables applied; drop overlaps resolve exactly as the server does.

## 0.3.2

Output quality round: the built file is judged against what a reader wants, and the builder may leave the approved geometry to get there.

Decks
- Fidelity setting on the project (strict, balanced, loose; default balanced): balanced grows blocks to fit, rebalances whitespace, aligns edges and snaps to the style's 12-column module; loose re-picks the layout from the content. Every deviation is recorded and shown as an "adjusted" chip in the Built view, with Keep positions to write it back into the plan.
- One type scale across the deck (largest headline size that fits every headline, one body size per role), the style's own sizes and fonts win over house defaults, readability floors only.
- Emphasis rendered per block type; tables on Tufte defaults (no header fill, accent rule, hairlines, right-aligned numbers, rating circles); palette-derived backgrounds; seven consistency findings (kicker drift, headline case, parallel cards, empty band, missing sources, bullet count, duplicate headline).
- Speaker notes from the story; generated cover and agenda (agenda statement from the takeaway, agenda lines on divider rows).
- build_deck returns the slide images to Claude and review_deck fetches any slide; the workflow says look at every slide, fix through the plan, rebuild, never hand-build.
- Placeholders that name a relationship become diagram blocks; placeholders are flagged before build.
- Canvas: Fidelity and Include controls in the style sheet, Fit and Balance actions, Snap toggle with module guides.

Workbooks
- Cells coloured by role, not by block; label columns fit their labels; widths from values with wrapping headers; one note column per band; computed freeze panes, print area, landscape and fit to width; named ranges for inputs; three-part number formats; row formats beside column formats; link cells; tab colours; the cover carries the answer.
- check_workbook.py: fourteen finding kinds (clipping, mixed formats, hardcodes, row consistency on period tables, unused inputs, empty bands, freeze, cover, totals) run after every build.
- build_workbook returns sheet images to Claude and review_workbook fetches any sheet; a plan with calcs and no Checks block is flagged; the workflow says never hand-patch the .xlsx.

## 0.3.1

Feedback round 1, after the first real deck run.

Outline
- Rows carry name and story only; the message line and section pills are gone. Divider rows are items in the list (Add slide / Add divider), renamed inline, dragged like any row; sections follow position and adjacent duplicates merge.
- One flat drag with a single insertion line; ⌘↑/⌘↓ selects rows, ⌥↑/⌥↓ moves them. Scroll position and caret survive every re-render (Cmd+Enter no longer jumps to the top).
- Brief is audience, goal, takeaway; emphasis removed. Tool text tells Claude what never goes in a field.

Layouts
- Headline required for every content slide in propose_layouts; headline blocks are folded into the headline slot; set_headline op; headlines start with a capital.
- Rows 1–4 and 18 are the furniture band: the server and both canvases keep blocks out; proposals without geometry are laid out by the shared archetype solver; overlaps resolve on save.
- Furniture, palette and type drawn from the selected style's tokens; divider slides render like the format's, numbered.
- ⌥R remixes (plain R does nothing); ⌘←/→ cycles slides; Delete removes blocks; blocks delete from the Blocks list; four-corner resize on both canvases; story and slide notes shown in the panel; the strip keeps its scroll.
- New diagram block (flow, cycle, timeline, tree, matrix) with an Edges editor.

Builder
- Native diagrams from the style tokens, placeholder mode for formats that ban shape diagrams; deck-diagrams skill in skills/ (copy to ~/.claude/skills).
- Fonts resolve weight suffixes ("Montserrat Light" → Montserrat) and fall back to Arial, never Helvetica. Missing table, takeaway and placeholder tokens derive from the palette (no more white-on-white headers on dark styles). Blocks that intrude into the furniture band are shifted before drawing. Kicker follows the format's pattern; dividers numbered per format.
- build_deck findings are actionable and the result says to fix them through Blueprint, never by hand-building the file.

Style sheet
- Folder and file name sit behind an Output disclosure.

## 0.3.0

### Renamed to Blueprint

- Renamed the package, bundle (`blueprint.mcpb`), display name and MCP server, so the model now calls `mcp__Blueprint__*`.
- Moved the home to `~/.blueprint` and the default output to `~/Documents/Blueprint`; added `BLUEPRINT_HOME|OUTPUT|PYTHON|SKILLS` and kept honouring the old `WORKBOOK_PLANNER_*` names.
- Copied projects, templates, styles and defaults from `~/.workbook-planner` on the first run.
- Kept the version in the canvas resource URI (`ui://blueprint/canvas-<version>.html`) and stamped the same version into the canvas build.
- Made the install one uninstall and one drag: the new extension id does not inherit the old output folder and Python path, so both are set again.

### Two modes

- Gave every plan a `kind`: workbook (.xlsx) or deck (.pptx), set at `start_project` and migrated onto old plans.
- Replaced `propose_tabs` with `propose_outline`, which takes rows of tabs or slides.
- Added the brief card (audience, goal, the one takeaway, emphasis chips), settable at `start_project`, from chat with `update_plan`, or on the canvas with `ui_save_brief`. A missing takeaway is flagged once per plan, not once per row.
- Added a one-line message to every tab and slide, plus a story and a section on slides; all of it reaches Claude through `get_plan`.
- Added the deck block types (headline, kicker, text, bullets, cards, stat, chart, table, image, quote, takeaway, placeholder, notes) on a 32x18 slide grid with their own default sizes; `propose_layouts` refuses block types from the other mode.
- Split the workflow text `start_project` returns per mode, cut it under 12 lines, and moved the per-step guidance into each tool's result.

### Design system

- Added `src/app/tokens.css` and `src/app/motion.ts`, rebuilt `styles.css` on the tokens, and moved every view onto the shared flip, fade, pulse and toast helpers.
- Trimmed 87 strings to the copy rules: labels of three words or fewer, verbs on buttons, no help paragraphs outside the `?` sheet.
- Rebuilt the Built view header: file name with Reveal, compact sheet list, one status line, warnings behind a count, and an Arrows toggle (`A`) remembered per project.
- Added a Playwright visual pass over both modes and both themes at desktop width.

### Outline view

- Put the brief card above the rows, editable in place and collapsible to one line once filled.
- Gave slide rows a name, the story with a quiet expander, a section group header and the message line; tabs keep name, purpose and message.
- Sized the canvas to its content instead of a fixed height.

### Deck mode

- Added `builder/build_pptx.py`: python-pptx builder with plain, house, skill and template styles, title, agenda, divider and content slide kinds, furniture from the style tokens, native charts and tables, text auto-fit, a snapshot JSON, slide PNGs and an optional PDF through LibreOffice.
- Added `builder/check_deck.py`: overflow, overlap, off-slide geometry, missing headline and empty placeholder findings, read back off the built file.
- Added `build_deck`: per-type content schema, refusal while a thread is open or a layout is unapproved, validation against the approved geometry and types, deck check findings and a built summary in the result.
- Added the deck layout canvas: 16:9 slide, 30 pt grid, style furniture drawn behind the blocks, a deck block palette, and an inspector with the content line and the emphasis toggle.
- Added Remix: ten archetypes, role-based slot mapping, extras kept below, undoable, on `R` or from the archetype chip.
- Added the deck Built view: a row and a vector thumbnail per slide from the snapshot, click to zoom, deck check findings and the plan diff.
- Added `import_deck`: a python-pptx scan that turns slides into outline rows and shapes into blocks by role with a confidence each, dividers into sections, and drops page numbers and logos.
- Added a deck flow test over stdio that asserts the built .pptx structure with python-pptx.
- Listed deck skills in the style picker with their kind, and shipped house deck and Harbor token sets.

### Format extraction

- Added `builder/extract_format.py`: reads .xlsx and .pptx into style tokens and writes a `format:` block, a `deck:` block, a SKILL.md and the deck tokens the builder loads.
- Added `extract_format`: takes file paths, hands screenshots back as images for Claude to read, takes those observations on the next call, then writes `~/.claude/skills/<slug>/SKILL.md`. Read-only plugin skills get a token sidecar under `~/.blueprint/styles` instead, and `fromSkill` derives tokens for a prose-only skill.
- Wrote the rules the extractor could not turn into tokens into the SKILL.md prose, and phrased every unknown as a question for the user.
- Documented the `deck:` block in `docs/skill-format-block.md`.

### Platform fixes

- Cleared the streaming sketch on `ontoolcancelled`, not only on the error path.
- Gated `updateModelContext` on `getHostCapabilities()`, and made `sendMessage` ask once and then remember a refusal.
- Flushed the debounced layout save on `requestTeardown` and `onteardown`, so a closing canvas keeps the last drag.
- Added a 512x512 `icon.png` to the bundle and listed every model-visible tool in the manifest.

## 0.2.3

- Canvas resource URI carries the version so Claude Desktop stops serving a cached older canvas after an upgrade.

## 0.2.2

- Row notes and total notes on tables, data validation on inputs and table columns, freeze panes, print area, landscape, fit to width and zoom per sheet.
- Rebuild from import: original formulas translated to block references.
- Audit fixes: build payload schema visible to Claude, reopen layouts, error bar instead of a stuck streaming sketch, template field, templates on a fresh home, double-click rename.

## 0.2.1

- Threads replace notes plus questions; answers read back to Claude.
- Multi-select, align and distribute, zoom and minimap, merge up and split at row, chart kinds, shortcut cheat sheet.
- Formatting skills drive the builder through a `format:` front-matter block; template workbooks can carry a Style sheet.
- Restyle an imported workbook in place; import diff; explain a built block.
- Payload errors name the tab and block; timing log; remembered style defaults.
- Full test suite under `npm test`; docs/VERIFY.md checklist for Claude Desktop.

## 0.2.0

- Dark-first Raycast-style canvas; light theme follows the host.
- Inline editors and popovers everywhere (the sandboxed iframe blocks browser dialogs).
- Live drag reorder for tabs; block drag, resize, Stack and Flow; undo/redo and keyboard shortcuts; overlap resolution on drop.
- One thread model: your instructions and Claude's questions share a lifecycle per tab or block; answers read back to Claude.
- Only one canvas visible per chat: older canvases collapse; chat edits refresh the live one, flash the changed blocks and toast a summary.
- Streaming sketch while Claude is still writing the proposal.
- Built view: read-only preview grid from a build snapshot, plan-vs-build diff, dependency arrows, formula-check findings you can click through to the cell.
- Formula check after every build (LibreOffice when installed, else a bundled evaluator).
- Style popover: plain, house, formatting skills found on this Mac, template workbook, output folder, file name.
- Templates: save an approved plan and start the next project from it.
- Import an existing .xlsx into a plan with per-block confidence; review on the canvas before rebuilding.
- Read Me cover only when the plan has a Read Me tab; content-fit column widths.
- Playwright e2e suite against a vendored reference host, plus a visual-pass spec.

## 0.1.0

- First version: tabs and layouts on a canvas, notes, house-format builder, .mcpb bundle.

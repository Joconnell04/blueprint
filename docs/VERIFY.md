# Verify in Claude Desktop

Three paths: a workbook, a deck, and extracting a style. Run them in order after a fresh install.
Each step says what to do, what you should see, and where to look when it does not.

When a step fails, in this order:

- the plan JSON, `~/.blueprint/projects/<id>.json` (what the server actually stored)
- `~/.blueprint/log/planner.log` (one line per slow step, with timings)
- `~/Library/Logs/Claude/mcp-server-Blueprint.log` (the server's stderr, Python errors included)

Paste the file and the step number into a Code session.

## A. Workbook

**A1. Install.** Settings > Extensions: uninstall **Workbook Planner**, then drag `blueprint.mcpb` onto the drop zone.
See: "Blueprint" under Installed. Set the output folder and Python path again; the rename does not inherit them.
Stuck: the Claude Desktop log names a failed spawn or a missing module.

**A2. Start.** New chat: "Use Blueprint. Market sizing model for a B2B payments product aimed at US SMBs, house format. Ask me what you need first."
See: 2 to 5 questions before any canvas, then the brief card filled from your answers (audience, goal, the one takeaway).
Stuck: if Claude proposes without asking, the workflow text did not arrive; check `start_project` in the server log.

**A3. Outline canvas.** Drag a tab by its grip, click a name to rename, add a tab on the dashed row, leave a note, then click **Approve tabs**.
See: one dark canvas, a Read Me tab first (house format), rows that shift as you drag, a toast saying the approval went to Claude, and Claude continuing without you typing.
Stuck: reload the chat; if the edit is missing from the plan JSON, `ui_save_tabs` failed in the server log.

**A4. Layouts canvas.** Drag a block onto another, press ⌘Z, press `?`, answer a "Claude asks" thread inline, then **Approve all layouts**.
See: real column letters and row numbers, the block underneath moving down, the undo landing, the cheat sheet, and the answer reading back to Claude.
Stuck: the plan JSON holds the block coordinates; a stale canvas means the poll is not running (server log, `ui_get_state`).

**A5. One canvas only.** Scroll up to the outline canvas.
See: it has collapsed to a one-line "moved to a newer message" bar.
Stuck: check the render sequence in the plan JSON; two live canvases mean the collapse did not fire.

**A6. Chat edits.** Type "move the ramp table to the top of its tab".
See: the live canvas flashes the moved block and toasts what Claude did. No second canvas appears.
Stuck: `update_plan` in the server log says whether the edit was applied.

**A7. Build.** Let Claude resolve the threads and build.
See: the Built view with the preview grid, "matches plan" or the differences, the Arrows toggle (`A`), and "Formula check: 0 errors". The .xlsx opens from `~/Documents/Blueprint`.
Stuck: a refusal names the open thread or the unapproved layout; a builder failure prints the Python traceback in the server log and keeps the spec under `~/.blueprint/builds/`.

**A8. Style popover.** Click the format chip in the header.
See: plain, house, your spreadsheet-format skill, and Template file. Change the output folder; the next project starts with it.
Stuck: a missing skill means the description did not match the format-skill test; check `~/.blueprint/defaults.json` for the remembered folder.

**A9. Templates.** Header > Templates > Save current plan. New chat: "start a Blueprint plan from the market sizing template".
See: the template listed, then the tabs arriving pre-filled and unapproved.
Stuck: `~/.blueprint/templates/` holds the saved file.

**A10. Import and restyle.** "Import ~/Documents/some-model.xlsx with Blueprint", then "restyle it in the house format".
See: one tab per sheet, Detected badges with confidences, nothing approved; the restyle writes a new file and leaves the original alone.
Stuck: the scan JSON is under `~/.blueprint/imports/`.

## B. Deck

**B1. Start.** New chat: "Plan a deck with Blueprint: Q3 board update for the investors. Ask me what you need first."
See: the questions first, then a deck plan (`kind: "deck"` in the plan JSON) and the brief card.
Stuck: if the plan comes back as a workbook, `start_project` got no kind; the tool result in the server log shows the arguments.

**B2. Outline rows.** Read the slide rows, expand a story, edit a section, fix a message.
See: a two to five word name (not a headline), two lines of story with an expander, a section group header, the message line.
Stuck: `ui_save_tabs` carries story, section and message; compare with the plan JSON.

**B3. Approve.** Click **Approve slides**.
See: the header says Slides approved, and Claude proposes layouts on its own.
Stuck: the button label still says Approve slides after a reload means `ui_approve` did not persist.

**B4. Layouts and Remix.** Press `R` on a slide, then pick an archetype from the chip next to it. Drag a block, resize it, set emphasis in the inspector.
See: the slide re-laid on the 32x18 grid, the headline band and page number drawn behind the blocks, a toast naming the archetype, and nothing dropped (extras stack below).
Stuck: the archetype is derived from the geometry, so check the block coordinates in the plan JSON; `⌘Z` undoes a remix.

**B5. Notes.** Leave a note on a block ("this chart should be the quarterly trend, not the mix") and let Claude answer it.
See: the thread on the block, then Claude's resolution with what it changed. Approve all layouts.
Stuck: a build that refuses lists every open thread by id.

**B6. Build.** "Build it, and export a PDF too."
See: the result says the path, the slide and shape counts, how many PNGs, and "Deck check: 0 errors". The Built view lists a row and a thumbnail per slide; click one to zoom.
Stuck: a `check_deck` error names the slide and the block; the spec and the check JSON stay under `~/.blueprint/builds/`.

**B7. Open the file.** Open the .pptx from `~/Documents/Blueprint`.
See: each headline reads as the argument of its slide, the kicker and page number sit where the style puts them, charts and tables are native objects, nothing runs off the slide.
Stuck: compare the file with `<name>.snapshot.json` beside it; the snapshot is what the canvas drew.

**B8. PDF and thumbnails.** Look beside the .pptx.
See: `<name>.pdf` and `<name>.slides/slide-1.png` and up.
Stuck: without LibreOffice both are skipped with a warning, never an error. `BLUEPRINT_SOFFICE` points at `soffice`; PNGs also want `pdftoppm`.

## C. Extract a style

**C1. Give it examples.** Attach a formatted .xlsx or .pptx in chat: "Blueprint: make a style called Acme deck from this file."
See: Claude calls `extract_format` with the file's path and no save, then reports which tokens came back and what it has to ask you.
Stuck: if Claude cannot pass a path, paste the absolute path instead; the tool takes paths, never base64.

**C2. Answer the unknowns.** Answer the questions it asks. If you attached screenshots, Claude reads them back as observations.
See: a second call with `observations`, now covering palette, fonts, the type scale, card and furniture geometry.
Stuck: `~/.blueprint/extracts/` holds every result JSON, newest last.

**C3. Save.** "Save it."
See: `~/.claude/skills/acme-deck/SKILL.md` written, deck tokens at `~/.blueprint/styles/acme-deck.json`, and a line saying it lists as `skill:acme-deck`.
Stuck: if the style does not list, the front-matter description does not read as a format skill; edit it in the SKILL.md.

**C4. Use it.** Open the format chip on a deck plan and pick Acme deck, then build.
See: the style in the chip, and the built deck in those colours and fonts.
Stuck: `build_deck` warns when a skill has no tokens and falls back to the house style; the warning names the skill.

## Known limits

- The streaming sketch (blocks appearing while Claude writes) needs the host to send partial tool input. The reference host does not, so this checklist is the only place it gets tested.
- The deck Built view draws slides from the snapshot, not from the PNGs. The PNGs are for opening beside the file.

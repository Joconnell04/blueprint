# build_xlsx.py contract

Invocation: `python3 builder/build_xlsx.py <spec.json>`. Writes the workbook, prints one JSON object to stdout, exit 0 on success, exit 1 with `{"ok": false, "errors": [...]}` on failure. Nothing else on stdout. Python 3.9 compatible, openpyxl only (no pandas).

## Input spec

```jsonc
{
  "outputPath": "/abs/path/file.xlsx",
  "formatting": "plain" | "house" | "skill" | "template",
  "skill": { "path": "/abs/path/SKILL.md" },   // only for skill mode
  "templatePath": "/abs/path/template.xlsx",   // only for template mode
  "fidelity": "strict" | "balanced",           // default balanced, see Fidelity
  "comments": true,                            // default false, see Comments
  "threads": [ { "blockId": "b_in", "entries": [ { "by": "Jackson", "text": "...", "status": "resolved" } ] } ],
  "project": { "name": "...", "brief": "..." },
  "readme": {                                   // optional. house mode always writes a "Read Me" first tab from this (fallbacks if absent)
    "purpose": "one paragraph",
    "answer": "the one line the model exists to produce",   // optional, the cover's Answer section
    "results": true,                            // optional, see "The cover" below
    "contents": [ { "tab": "Assumptions", "holds": "what the tab holds" } ],
    "rules": [ "Blue cells are inputs", "..." ]
  },
  "sheets": [
    {
      "name": "Assumptions",                    // Excel sheet name (<=31 chars, no []:*?/\)
      "purpose": "what the tab holds",          // optional, fills this tab's Contents row on the cover
      "blocks": [ Block, ... ],
      "freeze": "B3",                           // optional sheet setup, see "Sheet setup" below
      "printArea": "auto", "landscape": true, "fitToWidth": true, "zoom": 90, "titleRows": "2:3"
    }
  ]
}
```

### Sheet setup

Every field is optional. The builder works out its own setup for every sheet from what it wrote, so a sheet that says nothing still prints; a field the spec sets wins over the computed default. The cover tab takes the same fields.

| field | value | effect |
|---|---|---|
| `freeze` | a cell such as `"B3"` (case does not matter) | `ws.freeze_panes`: rows above and columns left of that cell stay put when scrolling |
| `printArea` | `"auto"` or a range such as `"B2:F40"` | `ws.print_area`; `auto` is the used range, read after the sheet is written |
| `landscape` | `true` or `false` | page orientation; `false` sets portrait explicitly |
| `fitToWidth` | `true` | page setup fitToWidth 1, fitToHeight 0 (one page wide, as many tall as needed); `false` leaves the page setup alone |
| `titleRows` | a row range such as `"2:3"` | `ws.print_title_rows`: those rows repeat at the top of every printed page |
| `zoom` | 10 to 400 | `ws.sheet_view.zoomScale` |

**Computed defaults.** Unless the spec says otherwise, a sheet with blocks gets:

- `freeze` at the first data cell of its first `table` or `calc` block (label column and header row both held), or, with no table on the sheet, at the first value cell of its first `inputs` or `output` block. A blockless cover gets no freeze.
- `printArea` on the used range, widened to cover every block's rectangle (a chart holds no values, so without this it would print blank) and every merged span.
- `landscape` once that range is wider than 90 width units, which is what fits a portrait page, else portrait.
- `fitToWidth`, so a sheet is one page wide however many columns it has.
- `titleRows` from the first table's title row through its header row, once the used range is more than 45 rows tall.

A bad value adds a warning naming the sheet and the field and falls back to the computed default; the rest still apply.

### Block

Common fields: `id`, `type`, `title`, `row`, `col` (1-based Excel coordinates of the block's top-left cell; `col` 1 = A), `cols` (planned width in cells), `note` (optional, a short grey note placed right of the title row), `cellRoles` (optional, see Cell roles below).

**How wide a block ends up.** `cols` is what the plan reserved; what the block occupies is what it writes:

- `inputs` and `output` are two columns, a label and its value, whatever `cols` says. The columns the plan reserved beyond that are note room.
- `table` and `calc` are as wide as their data, so a planned column the data does not fill cannot leave an empty column between the block and its note.
- `text` and `chart` keep `cols`: a paragraph needs its merge width and a chart its rectangle.
- Then blocks whose left edges sit within two columns of each other grow to the widest table in that band, so the title rules end on one edge; a block never grows more than two columns, and a pair block never grows at all. Every move is recorded in the snapshot's `adjusted` list as `{blockId, from, to, reason}` with reason `fit` (measured off its planned `cols`) or `snap` (grown to the sheet's block width), and on the block itself as `"adjusted": "fit" | "snap"`. The band snap is a `balanced` move: under `strict` only `fit` runs.

Beyond that, what may move is the `fidelity` setting's to say; see Fidelity.

**table / calc** also take `"total": true`, which asks for a total row the approved rows do not carry. Under `balanced` the builder writes one; under `strict` it does not. See Fidelity.

Fixed geometry, which the server's `get_plan` tool also uses to hand Claude cell addresses, so do not deviate:

| type | row 0 | row 1 | rows 2.. |
|---|---|---|---|
| `inputs`, `output` | title | first label/value pair | more pairs |
| `table`, `calc` | title | header row | data rows |
| `text` | title | text (merged across `cols`) | |
| `chart` | chart anchored at (row, col) | | |

**inputs / output**
```jsonc
{ "type": "inputs", "rows": [
    { "label": "Market size ($M)", "value": 1200, "format": "currency", "note": "Analyst estimate" },
    { "label": "Growth", "value": 0.12, "format": "percent" },
    { "label": "Blended price", "value": { "f": "={{pricing!Avg price}}*1.1" }, "format": "currency" }
] }
```
Label in the block's first column, value in the second, note in the row band's note column (see Notes below). Rows beyond the planned block height are still written (the planned height is a hint).

A row may carry `"validation"` (see Data validation below); it goes on the value cell, and `"role"` (see Cell roles), which fixes what that value is.

**table / calc**
```jsonc
{ "type": "table", "columns": ["Segment", "Year 1", "Year 2"],
  "colFormats": ["text", "currency", "currency"],
  "data": [
    ["SMB", 120, { "f": "=B5*(1+{{assumptions!Growth}})" }],
    ["Enterprise", 300, { "f": "=C5*1.2" }]
  ],
  "rowFormats": ["percent", "count", "currency"],                         // optional, one per data row
  "totalRow": ["Total", { "f": "=SUM(C5:C6)" }, { "f": "=SUM(D5:D6)" }],  // optional, bold
  "rowNotes": ["From CRM export", null],                                  // optional, one per data row
  "totalNote": "Ties to the P&L",                                         // optional, beside the total row
  "colValidation": [null, { "min": 0 }, { "list": "={{b_scen!col:0}}" }], // optional, one per column
  "cellRoles": { "0,1": "input" }                                         // optional, see Cell roles
}
```
Cell values: number, string, boolean, null, or `{ "f": "=..." }` formula, `{ "v": 5, "format": "percent" }` for a per-cell override, or `{ "v": "Census CBP", "link": "https://..." }` for a hyperlink (see Links).

**Which format a cell takes.** `colFormats` is per column and `rowFormats` is per data row, one entry each, null to skip. A row format wins on every cell of its row except where the column format is `text`, which keeps the label column text; a null row entry, or a row past the end of the list, falls back to the column format; a per-cell `{"v":.., "format":..}` wins over both. The total row keeps the column formats. This is what a label-plus-years table needs, where one row is a share, the next a count and the next a currency, and one column holds all three. `rowFormats` that is not a list adds a warning and is ignored.

Row notes: `rowNotes` holds one string or null per data row (extra entries are ignored; a short list leaves the remaining rows without a note) and `totalNote` is the note beside the total row. They go in the row band's note column, styled like inputs row notes: grey italic, unbordered, wrap off, subject to the house length check. When any row note or total note is present the block's extent grows by one column. The title rule and merge never extend over the note column. The snapshot records them with kind `note`.

`colValidation` holds one validation rule or null per column, applied to every data cell of that column (not the header, not the total row). A column with a `list` rule is a picklist, so its literal strings count as inputs.

### Data validation

A rule is one of:

```jsonc
{ "list": ["Low", "Base", "High"] }             // in-cell dropdown; items may not hold commas or quotes, joined list under 255 chars
{ "list": "={{b_scen!col:0}}" }                 // dropdown from a range; templates resolve, a plain "=C4:C9" works too
{ "min": 0, "max": 1 }                          // decimal between; either bound may be left out
{ "type": "whole", "min": 1 }                   // "whole" or "decimal" (default decimal), with min and/or max
```

Written as an openpyxl `DataValidation` on the target cells: blanks allowed, `showErrorMessage` on with a short message ("Pick a value from the list.", "Enter a number between 0 and 1.", "Enter a whole number of at least 1."). A malformed rule (not an object, unknown `type`, no list and no bound, bad list) adds a warning and writes nothing. The snapshot marks each validated cell that holds a value with `"validation": true`.

**text**: `{ "type": "text", "text": "paragraph" }` merged across `cols`, wrap on.

**chart**
```jsonc
{ "type": "chart", "title": "Revenue by year", "chart": {
    "kind": "bar" | "line" | "pie",
    "sourceBlock": "b_ramp",              // a table/calc block id in the same sheet
    "categoriesCol": 0,                    // zero-based column index within the source block's data
    "seriesCols": [1, 2],                  // zero-based column indexes to plot; header cells become series names
    "yTitle": "USD M"                      // optional
} }
```
Chart size and place: anchored from the block's top-left cell to its bottom-right, so the drawing covers exactly the cells the plan reserved and never sits over a written one (a two-cell anchor, `editAs="oneCell"`). `chart.width` and `chart.height` are also set, in centimetres, from the real column widths and row heights the block spans. The plot area is inset so the drawing does not run into its own frame, gridlines are off, and a chart of one series has no legend. The print area covers the chart, so it prints with its sheet.

### Formats

Three part, so a zero reads as a dash and a negative reads in parentheses, the way the house model does it:

`currency` -> `$#,##0;($#,##0);"-"`, `currency2` -> `$#,##0.00;($#,##0.00);"-"`, `percent` -> `0.0%;(0.0%);"-"`, `multiple` -> `0.0"x";(0.0"x");"-"`, `count` -> `#,##0;(#,##0);"-"`, `number` -> `#,##0.00;(#,##0.00);"-"`, `text`, `year` (write as text). A number with no format takes `count` when it is a whole number and `number` when it is not. A `format` name the builder does not know adds a warning naming the cell and falls back to `number`, rather than silently writing `4` as `4.00`.

### Cell roles

What a cell is decides its colour, and it is decided per cell, never by the block it sits in. The roles are `input`, `formula`, `link`, `note`, `label` and `text`, and they are what the snapshot records as `k`.

Derived, in this order: a hyperlink or `{"link": ..}` cell is a `link`; a formula that is nothing but a reference to another sheet (`='Ramp'!C4`) is a `link` too, any other formula is a `formula`; the literal text `TBD` is an `input`; the first column of a table or calc block is a `label`; a literal number, or a string that reads as a number such as a year, is an `input` wherever it sits; a literal string in a column with a `list` validation is an `input`, because it is a picklist value; anything else is `text`.

Two overrides, for the cells that rule gets wrong:

- `cellRoles` on a block: `{"r,c": "input" | "formula" | "link" | "note"}`, keys zero-based in the block's data area. For a table or calc block `r` is the data row (0 is the first row under the header) and `c` the column; for inputs and output `r` is the pair row and `c` is 0 for the label, 1 for the value. A key that is not `row,col` or a role the builder does not know adds a warning and is ignored.
- `role` on an inputs or output row, or inside a `{"v": .., "role": ..}` cell, which wins over everything.

### Links

A cell value of `{ "v": "Census CBP", "link": "https://..." }` is written as a hyperlink: the style's `link_color`, underlined, role `link`. A formula that is nothing but a reference to another sheet is a link too, and it takes the same colour; its underline is the style's `link_underline`, which is on in the house format because there the link colour is the ink colour. `link` takes a URL (`http`, `https` or `mailto`) or an internal reference (`'Sheet Name'!A1`, `Sheet!A1`, with or without a leading `#`); anything else adds a warning naming the cell and the text is written plain. With no `v` the link text is the target itself. A bare string that is a URL becomes a hyperlink on its own, which is what the Sources tab's Link column wants.

### Named ranges

Every input cell gets a workbook-scoped defined name, so `='Key Assumptions'!C3` can be written `=Take_rate` and survives a row being inserted. The name comes from the cell's label: a pair row's label, or a table cell's row label and column header. Characters other than letters, digits, `_` and `.` become `_`, a name cannot start with a digit or look like a cell reference, and it is capped at 60 characters. The first cell to want a name gets the plain one; a later cell with the same label takes its sheet name as a prefix (`Two_Growth`), then a counter (`Growth_2`). Formula templates still resolve to addresses: the plan and `get_plan` speak in cells.

### Formula reference templates

Inside any formula string, resolve `{{ ... }}` before writing:

- `{{blockId!Label text}}` -> the value cell of the row whose label equals the text (case-insensitive, trimmed) in an inputs/output block.
- `{{blockId!r,c}}` -> the data cell at zero-based (row, col) within a table/calc block's data area (row 0 = first data row, not the header).
- `{{blockId!col:c}}` -> the full data column range of a table/calc block, e.g. `C5:C9`.
- `{{blockId!total:c}}` -> the total row cell for column c (error if no totalRow).

If the referenced block is on another sheet, prefix with the quoted sheet name (`'Sheet Name'!C4`). Unresolvable templates are errors (fail the build, list them all). Block ids are unique across the whole workbook.

Also resolve `{{sheet:blockId}}` to the quoted sheet name of that block (useful for hand-written ranges).

## Formatting modes

**plain**: Calibri 11 default. Block title bold, table header bold with bottom border, thin borders around table cells, no fills, number formats applied, widths from the content. Gridlines left on. Roles still apply to links, which take Office's hyperlink blue `0563C1` underlined, but a literal input is not coloured: plain mode carries no meaning in its font colour.

**house** (the built-in house format; follow exactly):
- Arial 10 everywhere. Read Me title is Arial 16 bold navy `1F3864`; its subtitle is grey `595959` italic.
- Block title: bold, centered across the block width (merge), thin top border, medium bottom border navy `1F3864` across the block width.
- Column headers (table/calc): bold, centered. Total row bold.
- Thin black border around every body cell. A dashed vertical border on the right edge of the label column (first column) of every block, so the label column is separated from values by a dashed line.
- Font color by cell role, not by block: an `input` -> blue `0000FF`, a `formula` -> black, a `link` -> black and underlined, a `note` -> grey `595959` italic, not bordered, wrap off. A literal number in a table or calc block is an input, so it is blue there too. The house palette is blue, black and grey, with no green for a cross-sheet link, so a link takes the ink colour and the underline is what marks it (W14); a style that wants its own link colour sets `link_color`.
- No fills anywhere. Gridlines off on every sheet (`ws.sheet_view.showGridLines = False`).
- Values centered in their column; labels left-aligned.
- Number formats as above. Years written as text.
- `Read Me` tab first: title = project name, subtitle = "v1, built <YYYY-MM-DD>", navy rule, then sections "Purpose" (merged wrapped paragraph, explicit row height), "Answer" (see The cover), "Contents" (two columns: tab name, what it holds; one row per sheet), "Reading a Cell" (blue = input, black = formula, grey italic = note), "Rules" (from spec or defaults). No hard-wrapped prose across rows. Columns 30, 45 and 14, which is 89 units: the paragraph merges to the portrait print width, not past it.
- Tab colour by role, from the style: the cover grey `808080`, a tab whose blocks are all inputs blue `0070C0`, a tab with an `output` block green `00B050`, anything else Excel's default.
- Column widths: see Column widths below.
- No em dashes or en dashes in any cell: replace with a comma or colon.

**skill**: the house format parameterised by a formatting skill. `skill.path` points at a SKILL.md; the builder reads an optional `format:` block in its YAML front-matter (nested keys, two-space indent, scalars only; lists and deeper nesting are ignored) and builds exactly like `house` with those values swapped in. Everything the block does not set keeps the house default. A SKILL.md with no `format:` block builds as house with a warning naming the skill; an unreadable file does the same. Unknown keys and bad values each add a warning and keep the default. The prose of the skill is not read here: the server hands it to Claude, who follows it when writing the plan. Schema (defaults are the house format):

| key | default | meaning |
|---|---|---|
| `font` | `Arial` | font name everywhere |
| `size` | `10` | font size everywhere except the cover title |
| `title_size` | `16` | cover page (Read Me) title size |
| `input_color` | `0000FF` | any literal a reader can change, wherever it sits, and the text `TBD` |
| `formula_color` | `000000` | formulas and every other body value |
| `link_color` | the ink (`formula_color`) | a hyperlink and a bare cross-sheet reference; `none` is the ink colour, which is the house palette (W14) |
| `link_underline` | `true` | underline a bare cross-sheet reference too; a real hyperlink is always underlined |
| `note_color` | `595959` | notes and the cover subtitle |
| `note_italic` | `true` | notes in italic |
| `title_rule_color` | `1F3864` | block title rule, cover title colour, cover section rules |
| `title_rule` | `medium` | `medium`, `thin` or `none`: the bottom border of every block title row |
| `label_divider` | `dashed` | `dashed`, `thin` or `none`: the right border of the label column |
| `body_borders` | `thin` | `thin` or `none`: the box around every body cell |
| `gridlines` | `false` | sheet gridlines |
| `fills` | `false` | `true` fills column header cells with `header_fill` |
| `header_fill` | `F2F2F2` | the header fill colour when `fills` is on |
| `cover_tab` | `true` | write the cover page into a blockless `Read Me` sheet |
| `currency`, `percent`, `multiple`, `count` | house formats | Excel number format strings for those spec formats |
| `label_width`, `value_width`, `note_width` | `30`, `12`, `40` | minimum label and value widths; cap on the note column |
| `label_cap`, `value_cap` | `60`, `20` | cap on a label or text column, cap on a column of numbers |
| `tab_cover`, `tab_inputs`, `tab_calc`, `tab_output` | `808080`, `0070C0`, none, `00B050` | tab colour per sheet role; `none` leaves that role alone |
| `dash_replace` | `true` | replace em and en dashes with a comma (warning per cell) |
| `max_note_chars` | `100` | warn on cell text longer than this outside Read Me and text blocks; `0` disables |

Colours are six hex digits, with or without `#`. `docs/skill-format-block.md` shows the block on the spreadsheet-format skill's front-matter. The server's skill discovery only reads top-level `name:` and `description:`, so the block never changes which skills are listed.

**template**: open `templatePath` with openpyxl, read the default font (name, size) from the first non-empty cell of the first sheet that is not named `Style` and the first bold cell's font for titles; build like `plain` but with those fonts. A template names no link colour, so a link takes the template's ink, which is the `Formula` role's colour when its `Style` sheet has one and black otherwise, with the underline marking it (W14). Do not copy sheets from the template. If the template cannot be read, fall back to plain and add a warning.

If the template has a sheet named `Style`, it is a style sample: column A names a role and column B on the same row is the example cell. Roles, matched by the exact label (case does not matter): `Title`, `Header`, `Label`, `Input`, `Formula`, `Total`, `Note`, `Text`. The builder clones font, fill, border, alignment and number format from each example onto every cell of that role (the number format only where the spec set none). Column widths for label, value and note columns come from the `Style` sheet's B, C and D widths when they are set (label and value are minimums, the note width is used as is). A role without a row falls back to the plain behaviour in the template's base font, so a `Style` sheet with only `Title` and `Header` is fine. The block title is merged across the block when the `Title` example is centred. With an `Input` row present, literal inputs get that style, so the template can carry its own input colour. `builder/corpus/style_template.xlsx` (from `python3 builder/make_corpus.py`) is an example.

### Column widths

Widths come from the content, never from the plan, because nobody approved a column width.

- A **label column** (a pair block's labels, a table's first column) is at least `label_width` and as wide as its longest label, capped at `label_cap`. At the cap the label wraps and its row grows, so a label never clips.
- A column of **numbers** is as wide as its widest value once its own number format is applied, floored at `value_width` and capped at `value_cap`. A header does not widen it.
- A column holding **text** is capped at `label_cap`; a text cell longer than the column at its cap wraps and the row height follows the wrapped lines.
- A **header** wraps rather than widening its column, over two lines at most: a header too long for two lines at that width widens the column, still inside the cap, and the header row gets the height its wrapped lines need. `wrap_text` set by the width pass is never overwritten by the cell styling.
- A **note column** is as wide as its longest note, capped at `note_width`, and the note's merge (see Notes) gives it the rest of the room. Empty columns inside that merge are widened to fit the longest note, up to 25 each.

### Notes

One note column per row band. Blocks whose rows overlap share one column, immediately right of the widest of them, so a sheet never grows two note columns for one band and never leaves an empty column between a block and its note. The block title rule stops before it.

Every note is merged from that column to the sheet's last used column plus two, which is what gives a one-line note room to read; the merge stops before any cell that holds a value or is already merged, so it never covers anything. Notes stay grey italic with wrap off: a note that needs two lines is too long, and the house length check says so.

### The cover

`readme.answer` (or `project.takeaway`) is written as an Answer section under Purpose: the one line the model exists to produce. `readme.results` adds a row per result under it, each one the label, a live cross-sheet reference to the value and the tab it lives on: `true` takes every row of every `output` block, and a list of `{label, tab, ref, format}` says exactly which (a `ref` must start with `=`). Contents carries one row per tab in sheet order whatever the payload holds, taking `holds` from the matching `readme.contents` entry, else from that sheet's own `purpose`; a contents entry naming no tab is dropped with a warning.

## Fidelity

`fidelity` says how far the builder may move the approved geometry to make the sheet readable. Every deviation is recorded, never silent. The deck builder's setting works the same way (`builder/SPEC-deck.md`), with different moves: what a workbook wants to deviate on is width, and width is not in the plan at all.

**strict**: the planned cells are the cells. Widths still come from the content, because a column width is not plan geometry and nobody approved 24 characters, and a `table` or `calc` is still as wide as its data, because you cannot write five columns into four. Nothing else moves: no band snap, no widen, no generated total row, and a note stays beside its own block unless another block is written in that column.

**balanced** (the default), in this order:

1. **Widen.** A block whose columns cannot hold its content takes up to 2 more columns, and only when they are free. Only `text` and `chart` blocks can be short of their content, because they are the two that keep the `cols` the plan gave them: a paragraph that would run past three lines takes the columns that bring it back to three, and a chart narrower than the table it plots grows towards that table's width. A widen that would touch another block is not taken. The reason is `widen`.
2. **Total.** A `table` or `calc` carrying `"total": true` whose approved rows hold no `totalRow` gets one: the label `Total`, then `=SUM()` down every column that adds up. A column adds up when its format is `currency`, `currency2`, `count` or `number` and its header does not read as a rate, a price or an average (`per`, `each`, `average`, `avg`, `median`, `mean`, `rate`, `share`, `margin`, `price`, `ratio`, `multiple`, `growth`, `cagr`, `yield`, `uplift`, `discount`); a `percent`, `multiple`, `year` or `text` column is left blank, as is a column with nothing numeric in it. A table where nothing adds up warns and gets no row. The reason is `total`.
3. **Space.** A block the plan left touching the block above it shifts down by up to 2 rows to open one blank row. Top down, so a shift cascades; a shift onto a block that has not moved is skipped, and the pass runs again while anything moved. Two blocks the plan already overlaps are left exactly where they are, so the overlap is still an error naming both. The reason is `shift`.
4. **Note.** The band's note column, as in Notes above, which is a move for any block narrower than the widest in its band. The reason is `note`.

A generated total row that still lands on another block once the sheet has been spaced is given back, with a warning naming the block; that record's reason is `no-total`. Nothing else ever gives a move back.

**How a move is recorded.** Per block in the snapshot as `"adjusted": {"from": {row, col, rows, cols}, "to": {...}, "reason": "..."}`, and once per move in the sheet's `adjusted` list and in the result JSON as `{blockId, from, to, reason, sheet}`. The rectangle is the one the snapshot's block entry carries: the block plus its note column when that column is adjacent. A block with more than one move carries the first `from`, the last `to`, and the reasons joined with `+` (`"total+shift"`). The width records (`fit`, `snap`) keep their own shape, `from` and `to` being column counts, and a block whose only move is a width move still carries `"adjusted": "fit" | "snap"` as before.

Because a generated total row does not exist under `strict`, never point a `{{blockId!total:c}}` template at one: write the total into the plan when another block reads it.

## Comments

`comments: true` writes each block's threads into the workbook as one Excel cell comment on the block's title cell, author `Blueprint`. Off by default: a workbook carries no comments unless the spec asks.

`threads` is a list of `{blockId, entries: [{by, text, status}]}`, passed by the server after `resolve_notes` has closed them. Every thread naming the same block goes into that block's one comment, in the order given with a blank line between threads, so the newest entry is last. One entry is one line, `by: text`, and an entry whose `status` is anything other than `resolved` carries it (`Jackson: Make it monthly. [declined]`). An entry with no text is skipped, and a thread with nothing left writes no comment. A thread naming no `blockId`, a block this workbook does not hold, or a `threads` value that is not a list each add a warning and write nothing; a block with no title has no cell to hang a comment on and warns too. The snapshot marks the title cell `"hasComment": true`.

## Output

```json
{ "ok": true, "path": "...", "snapshot": "...", "fidelity": "balanced",
  "sheets": [ { "name": "Assumptions", "cells": 42, "formulas": 10, "charts": 1 } ],
  "adjusted": [ { "blockId": "b_in", "from": {"row": 2, "col": 2, "rows": 4, "cols": 3},
                  "to": {"row": 2, "col": 2, "rows": 5, "cols": 3},
                  "reason": "total", "sheet": "Assumptions" } ],
  "warnings": ["..."], "errors": [] }
```

`fidelity` is the setting the build actually ran at, so an unknown value comes back as `balanced`. `adjusted` is every move on every sheet, each naming its `sheet`; it is left out when nothing moved. See Fidelity.

## QA the script performs before returning
- Every formula string starts with `=`.
- Every `{{...}}` template resolved.
- Sheet names valid and unique.
- No block overlaps on a sheet (compare the blocks' own rectangles, after measuring, snapping and the fidelity pass); overlap is an error listing both block ids, and two blocks the plan already overlaps are never spaced apart to hide it. A note column belongs to a whole row band and always sits right of every block in it, so it cannot collide with one.
- In house mode: no cell text over 100 chars outside Read Me, no em/en dashes (auto-replaced, counted as warning).

## Tests
`builder/test_build.py` (plain unittest, no pytest needed): builds a 3-sheet sample spec in plain, house and template modes into a temp dir, re-opens with openpyxl, asserts titles at expected cells, formulas present with templates resolved to real addresses, chart count, house mode fonts/gridlines, skill mode with and without a `format:` block, a template `Style` sheet, table row notes (cells, styling, extent growth), data validation on inputs rows and table columns, sheet setup (freeze, print area, orientation, fit to width, zoom), and that an unresolved template fails with a listed error. A second class covers the quality rules: labels fitting their column and a clip detector over the shipped sample, roles and their overrides, row formats, the note column and its merge, widths from the values with wrapped headers, the computed sheet setup, named ranges, three part formats, the cover's Answer and Contents, chart anchoring, the width snap, links in every mode and tab colours. `FidelityTests` covers the two fidelities: the default and an unknown value, the widen, the generated total row and the columns it leaves blank, spacing and its two-row cap, a plan that already overlaps, the note column under each setting, and how every move reads in the snapshot and the result. `CommentTests` covers N9: off by default, one comment per block on its title cell, several threads in one comment, a declined entry, the warnings, and the file round trip. Run: `python3 builder/test_build.py`.

## Build snapshot (for the canvas preview)

Alongside the workbook, the builder writes `<outputPath minus .xlsx>.snapshot.json` and reports it as `"snapshot": "<path>"` in the result JSON. The canvas renders it as a read-only grid and diffs block extents against the plan.

```jsonc
{
  "sheets": [
    {
      "name": "Assumptions",
      "widths": { "B": 30, "C": 15 },          // column letter -> width (Excel units), only set columns
      "maxRow": 24, "maxCol": 6,
      "cells": [
        // one entry per written cell, in row-major order
        { "a": "B2", "v": "Key Assumptions", "k": "title", "hasComment": true },
        { "a": "B3", "v": "Market size ($M)", "k": "label" },
        { "a": "C3", "v": 1200, "k": "input", "fmt": "$#,##0;($#,##0);\"-\"", "d": "$1,200" },
        { "a": "C5", "v": "=C3*(1+C4)", "k": "formula", "fmt": "0.0%;(0.0%);\"-\"" },
        { "a": "C6", "v": "='Ramp'!C4", "k": "link", "fmt": "#,##0;(#,##0);\"-\"" },
        { "a": "D3", "v": "Analyst estimate", "k": "note" }
      ],
      "merges": ["B2:C2", "D3:F3"],
      "blocks": [
        // the cells each block ended up occupying, its note column included
        { "id": "b_in", "row": 2, "col": 2, "rows": 5, "cols": 3, "type": "inputs", "adjusted": "fit" },
        { "id": "b_tam", "row": 8, "col": 2, "rows": 6, "cols": 5, "type": "table",
          // a fidelity move, from rectangle to rectangle (Fidelity)
          "adjusted": { "from": { "row": 8, "col": 2, "rows": 5, "cols": 5 },
                        "to": { "row": 8, "col": 2, "rows": 6, "cols": 5 }, "reason": "total" } }
      ],
      "adjusted": [
        // every move the builder took on this sheet: a width move by column count, a
        // fidelity move by rectangle
        { "blockId": "b_in", "from": 4, "to": 2, "reason": "fit" },
        { "blockId": "b_tam", "from": { "row": 8, "col": 2, "rows": 5, "cols": 5 },
          "to": { "row": 8, "col": 2, "rows": 6, "cols": 5 }, "reason": "total" }
      ],
      "charts": [ { "blockId": "b_chart", "anchor": "B14", "kind": "bar", "series": 2 } ]
    }
  ]
}
```

`k` (kind) is one of `title`, `header`, `label`, `input`, `formula`, `link`, `total`, `note`, `text`, `readme`: the cell's role, with the structural kinds (`title`, `header`, `total`, `readme`) winning over it. A cell under a data validation also carries `"validation": true`. A block whose note column is not adjacent to it, because its row band has a wider block in it, also carries `"noteCol": <column number>`. A cell carrying an Excel comment is marked `"hasComment": true` (see Comments). `d` is the display string the builder would expect Excel to show for literals (apply the number format to the value in Python; omit for formulas and strings). Keep the file under 2 MB: if a sheet has more than 5,000 cells, write only the first 5,000 and set `"truncated": true` on the sheet.

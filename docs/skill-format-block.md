# The `format:` and `deck:` blocks in a formatting skill

A formatting skill's front-matter carries up to two blocks of style tokens:

- `format:` holds the workbook tokens. `build_xlsx.py` in `"formatting": "skill"` mode
  reads it and applies it mechanically: fonts, colours, borders, gridlines,
  number formats, widths.
- `deck:` holds the deck tokens, read by the deck builder: slide size, palette,
  fonts, type scale, card style, furniture positions, backgrounds, margins.

Everything else in the skill is prose for Claude, who follows it when writing the
plan. A skill with no block builds as the house format, with a warning naming the
skill. `builder/extract_format.py` writes both blocks from example files; see
[Deriving a block from files](#deriving-a-block-from-files) at the end.

## The workbook block

Below is the front-matter of the `spreadsheet-format` skill with the block added.
Every value shown is the default, so this block changes nothing; it documents the
schema. The user's skill file itself is not modified.

```yaml
---
name: spreadsheet-format
description: "A house format for any spreadsheet or financial model. Use whenever building or editing an .xlsx, alongside the xlsx skill. Governs layout, reading order, typography, notes, references, the cover page, naming, and the QA sweep before shipping."
format:
  font: Arial
  size: 10
  title_size: 16
  input_color: 0000FF
  formula_color: 000000
  link_color: none
  link_underline: true
  note_color: 595959
  note_italic: true
  title_rule_color: 1F3864
  title_rule: medium
  label_divider: dashed
  body_borders: thin
  gridlines: false
  fills: false
  header_fill: F2F2F2
  cover_tab: true
  currency: '$#,##0;($#,##0);"-"'
  percent: '0.0%;(0.0%);"-"'
  multiple: '0.0"x";(0.0"x");"-"'
  count: '#,##0;(#,##0);"-"'
  label_width: 30
  label_cap: 60
  value_width: 12
  value_cap: 20
  note_width: 40
  tab_cover: "808080"
  tab_inputs: 0070C0
  tab_calc: none
  tab_output: 00B050
  dash_replace: true
  max_note_chars: 100
---
```

## Parsing rules

- The block starts at a top-level line reading exactly `format:` with nothing after
  the colon. Its keys are the following lines indented by two spaces, `key: value`,
  one scalar per line. The block ends at the next unindented line.
- Only scalars are read. Lists, deeper nesting and multi-line values are ignored.
  This is why the deck tokens are a sibling block with dotted keys rather than a
  nested `format.deck` map: the parser is one level deep, and an unknown key inside
  `format:` costs a warning on every workbook build.
- Quotes around a value are stripped, so `currency: "$#,##0"` and
  `currency: $#,##0` are the same. Quote a value when it holds `#` or `:`.
- Every key is optional. A missing key keeps the house default above.
- Unknown keys, and values that fail the checks below, each add one warning to the
  build result and keep the default. The build never fails because of the block.
- A top-level `format: spreadsheet` tag (a scalar, not a block) is not a format
  block; the skill builds as house with a warning.
- The server's skill discovery reads only top-level `name:` and `description:`,
  so adding either block never changes which skills are listed.

## Keys

| key | type | default | effect |
|---|---|---|---|
| `font` | text | `Arial` | font name everywhere |
| `size` | number | `10` | font size everywhere except the cover title |
| `title_size` | number | `16` | Read Me cover title size |
| `input_color` | hex | `0000FF` | any literal a reader can change, wherever it sits, and the text `TBD` |
| `formula_color` | hex | `000000` | formulas and every other body value |
| `link_color` | hex or `none` | `none` | a hyperlink cell and a bare cross-sheet reference; `none` takes `formula_color`, which keeps the house palette blue, black and grey, and a hex such as `008000` colours links green |
| `link_underline` | bool | `true` | underline a bare cross-sheet reference too, not only a hyperlink; with no link colour the underline is the only mark a link has |
| `note_color` | hex | `595959` | notes and the cover subtitle |
| `note_italic` | bool | `true` | notes in italic |
| `title_rule_color` | hex | `1F3864` | block title rule, cover title, cover section rules |
| `title_rule` | `medium`, `thin`, `none` | `medium` | bottom border of every block title row (with `none`, no top border either) |
| `label_divider` | `dashed`, `thin`, `none` | `dashed` | right border of the label column; `none` leaves whatever `body_borders` gives |
| `body_borders` | `thin`, `none` | `thin` | box around every body cell |
| `gridlines` | bool | `false` | sheet gridlines on every sheet |
| `fills` | bool | `false` | `true` fills column header cells with `header_fill` |
| `header_fill` | hex | `F2F2F2` | the header fill when `fills` is on |
| `cover_tab` | bool | `true` | write the cover into a blockless `Read Me` sheet; `false` leaves it empty with a warning |
| `currency` | text | `$#,##0;($#,##0);"-"` | Excel number format for the `currency` spec format |
| `percent` | text | `0.0%;(0.0%);"-"` | for `percent` |
| `multiple` | text | `0.0"x";(0.0"x");"-"` | for `multiple` |
| `count` | text | `#,##0;(#,##0);"-"` | for `count` |
| `label_width` | number | `30` | minimum width of a block's label column; the column still grows to fit its longest label |
| `label_cap` | number | `60` | cap on a label or text column; at the cap the label wraps and the row grows instead of clipping |
| `value_width` | number | `12` | minimum width of a value column |
| `value_cap` | number | `20` | cap on a column of numbers, whose width comes from the formatted values |
| `note_width` | number | `40` | cap on the note column's content fit |
| `tab_cover` | hex or `none` | `808080` | tab colour of the `Read Me` sheet |
| `tab_inputs` | hex or `none` | `0070C0` | tab colour of a sheet whose blocks are inputs |
| `tab_calc` | hex or `none` | `none` | tab colour of a calc sheet; `none` leaves the tab alone, which is the default |
| `tab_output` | hex or `none` | `00B050` | tab colour of an output sheet |
| `dash_replace` | bool | `true` | em and en dashes become a comma, one warning per cell |
| `max_note_chars` | number | `100` | warn on cell text longer than this outside Read Me and text blocks; `0` disables |

Hex colours are six digits, with or without a leading `#`. Booleans accept
`true`/`false`, `yes`/`no`, `on`/`off`, `1`/`0`. Keys spelt `_colour` are accepted
as `_color`. `link_color` and the four `tab_*` colours also accept `none`.

**Number formats are three-part.** Each one is `positive;negative;zero`, so a
model shows `1,200`, `(1,200)` and `-` rather than `0` in an empty cell, and a
column of figures lines up on the dash. A two-part value like `$#,##0;($#,##0)`
is honoured as written, so an existing block keeps working; write the third part
when you want the dash. `currency2` is not configurable and stays
`$#,##0.00;($#,##0.00);"-"`.

## The deck block

`deck:` is a second top-level block, a sibling of `format:`, not nested inside it.
Its keys are flat scalars like the workbook block's, and groups are written with
dots: `palette.primary`, `typeScale.headline.size`, `furniture.logo.x`. A reader
expands the dots into the nested shape the deck builder wants; `extract_format.py`
emits that nested shape as JSON and flattens it when it writes a SKILL.md.

```yaml
deck:
  slideSize.width: 960
  slideSize.height: 540
  slideSize.grid: 30
  palette.primary: 1E3A5F
  palette.secondary: 2A9D8F
  palette.accent: E9C46A
  palette.ink: 1F2937
  palette.muted: 666666
  palette.bg: FFFFFF
  fonts.major: Lato
  fonts.minor: Lato
  typeScale.headline.size: 28
  typeScale.headline.weight: light
  typeScale.headline.bold: false
  typeScale.headline.color: 1F2937
  typeScale.headline.case: sentence
  typeScale.kicker.size: 16
  typeScale.kicker.case: title
  typeScale.cardHeader.size: 17
  typeScale.cardHeader.weight: semibold
  typeScale.cardHeader.bold: true
  typeScale.cardHeader.color: 1E3A5F
  table.headerFill: E3EEF2
  table.ruleColor: 9FB8C8
  table.rowHeightPt: 22
  divider.numbered: true
  diagrams: shapes
  card.shape: roundRect
  card.radius: 8000
  card.border: 2.25
  card.borderColor: 1E3A5F
  card.fill: FFFFFF
  furniture.logo.x: 29
  furniture.logo.y: 21
  furniture.logo.w: 55
  furniture.logo.h: 32
  furniture.logo.cell: "1,1"
  furniture.headline.x: 32
  furniture.headline.y: 108
  furniture.body.x: 32
  furniture.body.y: 178
  furniture.body.w: 868
  furniture.body.h: 322
  backgrounds.title: 1E3A5F
  backgrounds.divider: 2A9D8F
  backgrounds.content: FFFFFF
  margins.left: 32
```

### Deck keys

| key | type | effect |
|---|---|---|
| `slideSize.width`, `slideSize.height` | number, pt | slide size; the canvas assumes 960 x 540 |
| `slideSize.grid` | number, pt | the layout cell; 30, so 32 x 18 cells |
| `palette.primary` | hex | card borders, card headers, the title slide background, emphasis |
| `palette.secondary` | hex | section dividers, card subheads |
| `palette.accent` | hex | the sparing third colour: callouts, highlights |
| `palette.ink` | hex | headline and kicker text on a light slide |
| `palette.muted` | hex | body copy, page numbers, source lines |
| `palette.bg` | hex | the content slide background |
| `fonts.major` | text | display and headline font |
| `fonts.minor` | text | body font |
| `fonts.mono` | text | optional: numerals, page numbers, labels |
| `typeScale.<role>.size` | number, pt | type size for that role |
| `typeScale.<role>.weight` | `light`, `regular`, `medium`, `semibold`, `bold` | the face weight that role is set in, e.g. `typeScale.headline.weight: light` for a deck headlined in Montserrat Light |
| `typeScale.<role>.bold` | bool | the bold flag written into the file; it follows `weight`, so only `semibold` and `bold` are true |
| `typeScale.<role>.color` | hex | colour for that role |
| `typeScale.<role>.case` | `upper`, `title`, `sentence` | how the text is cased; `typeScale.kicker.case: title` is the kicker casing a format like Harbor keeps, and `upper` becomes `caps` in the builder tokens |
| `card.shape` | `roundRect`, `rect` | the preset geometry of a card |
| `card.radius` | number | the `prstGeom` adj value, 0 to 50000; 8000 is a soft fillet |
| `card.border` | number, pt | card border weight |
| `card.borderColor` | hex | card border colour |
| `card.fill` | hex | card fill |
| `table.headerFill` | hex | fill behind a table's header row; omit it for the no-fill Tufte default |
| `table.ruleColor` | hex | the hairline between rows, e.g. `table.ruleColor: 9FB8C8` |
| `table.rowHeightPt` | number, pt | the row height a table is measured at, e.g. `table.rowHeightPt: 22` |
| `divider.numbered` | bool | `divider.numbered: true` reads a section divider as `03 Competitive analysis`; `false` shows the bare title |
| `diagrams` | `shapes`, `placeholder` | `diagrams: placeholder` draws the labelled dashed box a format that bans shape diagrams mandates; `shapes` draws the diagram |
| `furniture.<role>.x`, `.y`, `.w`, `.h` | number, pt | where a standing element sits, from the top left |
| `furniture.<role>.cell` | `"col,row"` | the same position as 1-based grid cells |
| `furniture.body` | box, pt | the content measure every block is laid out inside, e.g. `furniture.body.x: 32` with `furniture.body.w: 868`; the builder cuts its twelve-column module from it, so a deck without it falls back to the house measure and lands on the wrong column edges |
| `backgrounds.<kind>` | hex | slide background by kind: `title`, `divider`, `content`, `agenda` |
| `margins.left`, `.right`, `.top`, `.bottom` | number, pt | where content starts and stops |

Type roles: `headline`, `kicker`, `body`, `caption`, `cardHeader`, `cardSub`,
`cardBody`, `pageNumber`, and the optional `display` for a divider's large title.
Furniture roles: `logo`, `pageNumber`, `kicker`, `headline`, `body`, `source`. A
reader that only knows the first eight type roles ignores `display` safely; every
key is optional and a missing one keeps the builder's own default.

`weight` and `bold` say the same thing twice on purpose: the builder writes a
bold flag into the .pptx, and the weight names the face the family should be
asked for. A deck set in Montserrat Light was being rewritten in bold because the
only reading was a flag; with the weight recorded, `light` stays light.

## Deriving a block from files

`builder/extract_format.py` reads example files and writes both blocks.

```
python3 builder/extract_format.py --out result.json model.xlsx deck.pptx
python3 builder/extract_format.py --out result.json --merge observations.json deck.pptx
python3 builder/extract_format.py --skill-md my-format "When to use it" model.xlsx
python3 builder/extract_format.py --tokens style.json deck.pptx
```

`result.json`:

```json
{"format": {"workbook": {...}, "deck": {...}},
 "confidence": {"workbook.font": 0.97, "workbook": 0.94},
 "evidence": {"workbook.font": "Arial on 100% of 214 styled cells"},
 "unknowns": ["no coloured input cells found: ask which colour marks an editable input"],
 "rules": ["Inputs live on their own tab (Assumptions); every other tab reads from it."],
 "samples": {"palette": ["1E3A5F", "2A9D8F"], "fonts": ["Lato"]}}
```

- `format.workbook` uses the workbook keys above and nothing else, so the block
  can be pasted into a SKILL.md and built with no change to `build_xlsx.py`.
- `confidence` is keyed by dotted path, 0 to 1. One observation never reaches 0.8;
  three cells or shapes that agree do. `confidence.workbook` and `confidence.deck`
  are the mean over the key fields, so a file with no style reads low overall.
- `evidence` says what was counted, in one line, for every field.
- `unknowns` names what the files cannot answer, in the words to ask the user.
- `rules` holds the observations that are not tokens (a frozen header row, an
  inputs-only tab, a page number that flips colour); they become the prose in a
  generated SKILL.md.
- `samples` lists every colour and font by how often it was used, for a picker.

Extra confidence keys carry facts that have no schema key: `workbook.freeze_panes`
and `workbook.inputs_separated`. Their evidence lines say what was seen.

Two deck readings are measured rather than counted, because the modal answer is
the wrong one for both. `furniture.body` is the union of the body shapes on the
content slides, trimmed to the 10th percentile left edge and the 90th percentile
right edge, and banded from the foot of the headline to the top of the source
line; a legend parked in the margin or one block run full width therefore cannot
widen the measure for the whole deck. `furniture.headline.x` is the median left
edge of the headlines, so one slide whose biggest run is a stat rather than a
headline cannot move the column. The kicker, page number, logo and source are
furniture and are excluded from both.

`--merge observations.json` takes Claude's reading of a screenshot in the same
schema and merges it over the file-derived block: **the file wins any field it
read with confidence 0.8 or better; the observations win everything else**,
including fields no file could show. Images passed as paths are not parsed; they
land in `unknowns` for Claude to read in chat and pass back this way.

`--tokens style.json` writes the deck block in the shape `build_pptx.py` loads
through `style.tokensPath`, so an extracted deck style can be built without a
hand edit. The two vocabularies differ in six names and one wrapper:

| deck block | deck builder tokens |
|---|---|
| `palette.primary` | `palette.accent` |
| `palette.secondary` | `palette.accent2` |
| `palette.accent` | third entry of `palette.series` |
| `palette.ink`, `.muted`, `.bg` | the same |
| `fonts.major`, `fonts.minor` | `fonts.display`, `fonts.body` |
| `card.radius`, `card.border` | `card.radiusAdj`, `card.borderPt` |
| `typeScale.display` | `typeScale.dividerTitle` |
| `typeScale.<role>.case: upper` | `typeScale.<role>.caps: true` |
| `backgrounds.<kind>: hex` | `backgrounds.<kind>: {"fill": hex}` |

`typeScale.<role>.weight`, `furniture.body`, `table`, `kicker`, `divider`,
`diagram` and `diagrams` are spelt the same on both sides and pass through as
they stand, so a block that carries them builds with them.

Two things the translation decides rather than reads: a role's colour becomes a
palette name (`ink`, `accent`, `muted`) whenever the hex matches one, because a
slide kind can override a name and cannot override a hex; and a dark background
gets `ink` and `muted` set to the colour its own display text was, so white text
survives on a dark slide. Everything else passes through unchanged.

`--skill-md name description` prints a complete SKILL.md: front-matter with the
name, the description, the `format:` block and the `deck:` block, then a short
prose section listing the rules the tokens imply.

What files cannot show, and Claude should ask about: which colour means "editable"
when a workbook has no coloured inputs; whether the format wants a cover tab when
the example has none; `dash_replace` and `max_note_chars`, which are authoring
rules rather than marks on a page; the intent behind a one-off value that only
appears once; and anything on a slide that is an image rather than a shape.

# build_pptx.py contract

Invocation: `python3 builder/build_pptx.py <spec.json>`. Writes the deck, prints one JSON
object to stdout, exit 0 on success, exit 1 with `{"ok": false, "errors": [...]}` on failure.
Nothing else on stdout. Python 3.9 compatible, python-pptx only
(`python3 -m pip install python-pptx`). Geometry, type ranges and furniture follow
`docs/research/deck-planning.md`.

## Input spec

```jsonc
{
  "name": "Referral Growth Review",
  "outputPath": "/abs/path/deck.pptx",
  "exportPdf": false,                   // true writes <name>.pdf when LibreOffice is present
  "thumbnails": true,                   // default true: slide PNGs under <name>.slides/
  "style": { "mode": "plain" | "house" | "skill" | "template",
             "tokens": { ... },         // skill mode: inline tokens (schema below)
             "tokensPath": "/abs/x.json",  // skill mode: the same tokens as a file
             "templatePath": "/abs/t.pptx" },  // template mode
  "slideSize": { "w": 960, "h": 540 },
  "grid": { "cols": 32, "rows": 18, "cell": 30 },
  "fidelity": "strict" | "balanced" | "loose",   // default balanced, see Fidelity
  "cover": true,                        // default true when no slide is a title
  "agenda": true,                       // default true when no slide is an agenda
  "agendaStatement": "...",             // optional, the generated agenda's statement
  "notesFromStory": true,               // default true: a slide's story becomes its notes
  "date": "Draft for discussion",       // optional, the generated cover's date line
  "brief": { "audience": "...", "goal": "...", "takeaway": "...", "emphasis": ["..."] },
  "slides": [ Slide, ... ]
}
```

`brief` is metadata: `takeaway` lands in the file's comments, `goal` feeds the generated
cover and agenda, nothing else is drawn.

### Slide

```jsonc
{
  "id": "s_stall",
  "kind": "title" | "agenda" | "divider" | "content",   // optional, see below
  "name": "Where referrals stall",   // Layer 1 shorthand purpose
  "story": "...",                    // Layer 1 note, not drawn
  "section": "Why referrals",        // the nearest divider above; feeds the kicker
  "message": "...",                  // one line, not drawn on content slides
  "kicker": "Strategic Rationale | Why Referrals",   // optional; wins over the pattern
  "headline": "Referrals stall at three moments, and each one has a different fix",
  "subtitle": "...",                 // title and divider slides
  "date": "Draft for discussion",    // title slide
  "source": "Partner CRM",           // renders as "Source: Partner CRM" bottom left
  "pageNumber": 7 | false,           // default: the slide's index, none on a title slide
  "notes": "speaker notes",
  "blocks": [ Block, ... ]
}
```

**Slide kind** decides the background, the furniture and the type roles. When `kind` is
absent: a slide with a `section` and no blocks is a `divider`, everything else is `content`.
`cover` is an alias for `title`; `close`, `thanks` and `section` are aliases for `divider`.

| kind | draws | from |
|---|---|---|
| `title` | wordmark, title, subtitle, date. No page number | `style.copy.wordmark`, `headline`, `subtitle`/`kicker`, `date` |
| `agenda` | eyebrow, statement, page number, then the blocks | `style.copy.agendaEyebrow`, `headline` |
| `divider` | "NN Section", subtitle, page number | `headline` or `section`, then `subtitle`/`message` |
| `content` | kicker, headline, hairline, source, page number, then the blocks | the slide fields |

### Block

Common fields: `id` (unique across the deck), `type`, `role` (the archetype slot it came
from, recorded in the snapshot), `title` (optional, drawn above the content), `col`, `row`,
`cols`, `rows` (1-based grid cells), `emphasis`, `content` (per type).

Geometry: `x = (col - 1) * cell`, `y = (row - 1) * cell`, `w = cols * cell`, `h = rows * cell`.
A block that runs past the slide is clamped to it and adds a warning. When a template's slide
size differs from `slideSize`, the template wins and the cell size becomes
`slideWidth / grid.cols` by `slideHeight / grid.rows`, so the grid still spans the slide.

| type | `content` | notes |
|---|---|---|
| `headline` | `{text}` | Furniture: folded into the slide's own headline when it is empty, dropped with a warning when it is not. Never drawn at the block's cells |
| `kicker` | `{text}` | Furniture, same rule. Also dropped with a warning in a style that has no kicker |
| `text` | `{paragraphs: [string \| {text, runs}]}`, or `{runs: [...]}`, or `{text}` | A run is `{text, bold?, italic?, color?}` |
| `bullets` | `{items: [string \| {text, children: [...]}]}` | Nested to any depth; each level indents 18 pt |
| `cards` | `{items: [{header, sub?, body}]}` | Laid across the block rect with the style's gutter; 5 or more wrap to two rows |
| `stat` | `{items: [{value, label}]}`, `layout?: "band" \| "grid"` | Same tiling, no box. See [Figure grids](#figure-grids) |
| `chart` | `{kind: bar\|column\|line\|pie\|stacked, categories: [...], series: [{name, values}], emphasisSeries?}` | A native PowerPoint chart |
| `table` | `{headers: [...], rows: [[...]]}` | A native table on Tufte defaults. See [Tables](#tables) |
| `image` | `{path}` or `{placeholder: "label"}` | Fitted inside the rect, aspect kept. A missing file becomes a labelled placeholder plus a warning |
| `quote` | `{text, attribution?}` | Curly quotes added |
| `takeaway` | `{text}` | The filled, borderless bar |
| `placeholder` | `{label}` | A dashed box with the label centred. An honest placeholder, never shape art |
| `diagram` | `{kind, direction?, nodes: [{id, label}], edges: [{from, to, label?}]}` | A native flow, cycle, timeline, tree or matrix. See [Diagrams](#diagrams) |
| `notes` | `{text}` | Appended to the speaker notes. Draws nothing, but keeps a snapshot entry |

Empty content draws `[Title]` in muted italic and adds a warning; it never draws nothing.

### Emphasis

`emphasis: true` says this block carries the slide, and the builder gives it the style's
accent treatment for its type. The accent gives way to `ink` when it has less than a 3:1
contrast ratio against that slide kind's background, so the treatment never costs the
reader the words.

| type | treatment |
|---|---|
| `cards` | the border in the accent at `card.borderPt` + 0.75, and the card header in the accent |
| `stat` | the value is already the accent; a hairline is drawn under it, the width of the tile |
| `text`, `bullets` | a 4 pt accent bar down the left edge, the copy inset past it, and the first sentence bold |
| `table` | the rule under the column heads at 1.5 pt instead of 1 pt |
| `chart` | one series in the accent and the others muted: `content.emphasisSeries` by name or index, else the first |
| `takeaway` | inverse: the bar in the accent, the text in the slide's background colour |

Every shape of an emphasised block carries `emphasis: true` in the snapshot. The measure
and the deck-wide type pass see the inset box a text or bullets block gets, so the accent
bar never costs it a size.

### Tables

No fill, rules where they carry meaning, figures on the right edge (Tufte, Few):

* **No header fill** unless the style itself names `table.headerFill`. A fill inherited
  from house is not the style asking for one, so a partial token file gets none.
* **Column heads** in caps in the accent (or ink when the accent fails contrast), at the
  deck's `cardSub` size and never larger than the cell size: a column name is a label.
* **One 1 pt accent rule** under the heads, 0.5 pt hairlines in `line` *between* the rows,
  and nothing under the last row.
* **Numeric columns right-aligned**, text left. A cell is numeric when it reads as a
  number, a percent, a currency or an "x" multiple: `12`, `7.7%`, `$60K`, `£1,200`,
  `4.1x`, `(3)`. `2 weeks` is text. Alignment is a column decision: a column goes right
  only when every cell in it is a figure, because one figure pulled to the right edge of a
  column of phrases reads as a mistake.
* **The first column set by `table.labelColumn`** when it reads as a label column:
  every cell short text, with figures or ratings to its right. The token is
  `{case: "caps" | "none", color: "accent" | "ink" | "muted", size, bold}`; house and
  harbor ship `none` / `ink` / bold, which is the family's medium weight, and a format whose
  row labels are copper caps in the gutter says `{case: "caps", color: "accent", size: 8.5,
  bold: false}`. `size` defaults to the deck's cell size and sets the row height with it;
  `color: "accent"` gives way to ink below the 3:1 contrast ratio, as everywhere else.
* **Rating cells as circles** in the accent, sized to the row: filled, half and empty.
  A column has to agree before any of it is drawn as a rating, because a lone digit is
  only a rating in a column of ratings; one `●○◐` glyph or a `3/5` settles a column on its
  own, bare digits 1 to 5 need two rows saying the same thing. Each circle is a snapshot
  shape at `role: "tableRating"`.
* **Row height from the deck-wide `tableCell` size**, plus the cell padding, per row: a
  table is as tall as its rows need and is never stretched to fill its block. A cell that
  wraps gets a taller row. A table taller than its block warns. An empty cell is written as
  one sized non-breaking space: a bare empty run leaves the renderer to pick its own font
  size for the line box, the row grows past the height the table was measured at, and the
  rating marks drawn over it would miss their rows.

### Figure grids

A `stat` block is a band of headline numbers by default: one tile per figure, value over
label, laid across the rect. `layout: "grid"` sets it as a figure list instead, which is
what a deck does with supporting numbers: **number and label pairs in two columns, in
reading order, with a hairline under each pair**. The number sits in a fixed left gutter,
sized to the widest value so every label lines up, in the accent at `statValue` x 0.6; the
label is set in the `body` role beside it.

The builder picks `grid` on its own when a stat block holds **4 or more items in a block
more than 6 rows deep**: four figures that deep are supporting figures, not the slide's
headline numbers. `layout` on the block settles it either way; an unknown value warns and
draws the band. Every shape of a stat block carries its `layout` in the snapshot.

**Which side the rule sits on (F-5).** `stat.rule` is `"below"` (the default, and house),
`"above"` or `"none"`, and one token settles every rule a figure carries: the accent hairline
under an emphasised figure in the band and the hairlines of a figure grid. `"above"` heads
each figure with its rule and moves the figure clear of it; the block's footprint is the
same either way. `"none"` draws neither. An unknown value warns and rules below. Every rule
shape a stat block draws carries `rulePosition` in the snapshot.

### Text fit

Per `docs/research/deck-planning.md` section 7, and the canvas must use the same numbers:

1. Characters per line `= slotWidth / (size * k)`, `k` = 0.5 body, 0.48 bold display,
   0.55 all caps, 0.6 numerals; the value per role lives in the type scale.
2. Greedy word wrap gives the lines; height `= lines * size * leading`.
3. While the text is taller than the box or over the role's `maxLines`, step down 2 pt
   (1 pt below 14 pt). Paragraphs sharing a role step together, so a list keeps one size.
4. At the role's floor, stop and warn. Nothing is ever set below the floor.
5. A last line of one word steps down once when that pulls it up, and is flagged otherwise.
   A role the deck sets once (below) takes that step for the whole deck or not at all.

## Fidelity

`fidelity` says how far the builder may move the approved geometry to make the slide
readable. Every deviation is recorded, never silent.

**strict**: the planned cells are drawn. The furniture guard still runs, because a block
under the headline is a collision, not a preference.

**balanced** (the default), in this order:

0. **Module.** Every left edge and width snaps to the style's column module, so one deck
   uses a handful of left edges instead of one per slide. The module comes from
   `furniture.body`: twelve columns with a one-cell gutter, `module = (body.w - 11g) / 12`
   with `g` = one cell (30 pt), so a block spanning n columns is `n * module + (n-1) * g`
   wide. A left edge takes the nearest module left and a right edge the nearest module
   right, up to 2 cells away, and never across the midpoint of the gap to a neighbour or a
   furniture band, so a snap cannot put two blocks on top of each other. The module rarely
   divides the 30 pt grid evenly: the move is recorded in cells as usual and the points it
   really landed on come back as `snappedPt: {x, w}`, which is what the builder draws. A
   style with no `furniture.body` has no module and the pass does nothing. The reason is
   `module`.

0b. **The source row (F-3).** A slide that drew a source line needs row 17 for it, so the
   body band ends at row 16. The row group that reaches row 17 moves up one row as a group,
   keeping the columns level; when there is nothing above it to move into, nothing moves and
   each block that still reaches row 17 is named in a warning. Nothing grown, rebalanced or
   snapped afterwards takes the row back. The reason is
   `moved up to keep row 17 for the source line`.

1. **Grow.** A block whose content does not fit at the role's floor takes up to 2 rows (or
   2 columns) of free space, down before up and right before left. When the space is taken,
   the one neighbour in the way gives up the same cells if it has the slack to spare, which
   is measured at its own size, not at its floor. Nothing moves onto the furniture.
2. **Rebalance.** When the body area (rows 5 to 17, clipped to the furniture this slide
   actually drew) has an empty band deeper than 3 rows below the last block, the slack is
   shared out: one row of blocks is centred in the band, several rows grow proportionally to
   their heights and keep their gutters.
3. **Align.** Like edges within one cell of each other snap together: lefts with lefts,
   rights with rights, tops with tops, bottoms with bottoms. The value most blocks already
   share wins, ties going to the outer edge. A snap that would collide is skipped.

4. **Copy.** Three passes that change the words rather than the cells, all reported
   (`from`/`to` carry text, not cells):

   * **The source line (Q15).** A source that does not fit its slot on one line at the caption
     floor is shortened until it does: the detail in brackets goes first, then the clauses
     after each `;` from the back, then the trailing sentences, and last of all the line is cut
     on a word with an ellipsis. The full wording goes to the speaker notes, so nothing is
     lost, and the move is reported with the reason `source` and no block id.
   * **Duplicate content (Q15).** A block whose **first paragraph repeats the headline** (the
     words that carry meaning, with at least 80% of the larger set shared) loses that
     paragraph; a block that says nothing else is dropped entirely. Reason `duplicate`. Two
     blocks on one slide carrying the same title is a warning, never a change.
   * **One treatment per row (Q15).** Blocks planned on the same row read as one row, so they
     look like one: when they are **all text-like** and any of them is `bullets`, every `text`
     block among them is drawn as `bullets`, keeping its title as a title rather than as the
     first item. Bullets win because a reader can see a marker and cannot see its absence. A
     row that also holds a table, a stat or a chart is not a row of parallel copy and is left
     alone. Reason `parallel`; the runs inside a converted paragraph become plain text.

5. **The shrink floor (Q15).** After the deck-wide type pass, a `text` or `bullets` block that
   still overflows at its role's floor gives up its **last paragraph to the speaker notes**,
   one at a time until it fits or only one is left. Cutting a sentence in half is the one thing
   the fit may never do. Reason `moved to notes`.

**loose**: balanced, plus a re-pick when the content plainly does not suit the slot. The
slot comes from the archetype catalogue in `docs/research/deck-planning.md` section 5, for
that role and count, and is taken only when it is free:

| trigger | goes to |
|---|---|
| a table of 5 or more columns with under 64 pt a column | `table`, the full-width wide box |
| 4 or more cards with under 150 pt a card | the wide card box |
| a `stat` whose value runs past 3 words or 20 characters | `statement`, or the stat band |
| a `stat` under three quarters of the band wide, beside exactly one `text` or `bullets` block on its rows | the full-width stat band, the rail moved under it |

The last one is the boxed-in stat band (Q21): figures squeezed into half a slide beside a
bullet list are figures the reader cannot see, and the rail has no reason to be beside them.
The stats take the `stat-band` slot at the top of the free band and the rail takes the rest
of it underneath, at the same left edge and width. Both moves are recorded with the reason
`stat-band`, and the pass gives up rather than move either block when the band is not free.

**strict** runs none of the copy passes and keeps the shrink floor's overflow: strict draws
the plan as approved and names the problem in the findings instead. The title warning and the
one-line source test still run, because neither is a change to the plan.

Deviations come back per block in the snapshot and in the result as
`{blockId, from: {col,row,cols,rows}, to: {...}, reason, snappedPt?}`. An entry whose reason is
`source`, `duplicate`, `parallel` or `moved to notes` carries `from: {text}` / `to: {text}` (or
`{type}` for `parallel`) instead of cells, because the change is to the copy: those entries
appear on the slide and in the result, and never badge a shape. `from` is the
rectangle the slide had after the furniture guard, so the reason and the move always
describe the same step. A block whose columns did not change keeps the height the guard
gave it, so a module snap never hands back a row the guard trimmed.

## Deck-wide type

A deck reads as one deck when a role is one size throughout, so the builder picks the size
once and pins it: the largest size, at or below what the style's tokens ask for, at which
every case in the deck fits. Never below the role's floor.

| set once | from |
|---|---|
| `headline` | every content slide's headline in its slot, in no more lines than the role allows. The tightest slot sets the size |
| `body`, `bullets` | every `text` and `bullets` block in its rectangle |
| `cardHeader`, `cardSub`, `cardBody` | every `cards` block, the three roles stepping together because they share a card |
| `tableHeader`, `tableCell` | every `table` block, cell by cell against its own row height. A table cell may wrap; `maxLines` is a rule for prose |
| `statValue`, `statLabel` | every `stat` block, the pair stepping together |

Two things keep the rule honest. A block that will not fit even at the floor cannot be
satisfied at any size, so it is reported as `overflow` instead of dragging the deck down
with it. And when the chosen size leaves a one-word last line anywhere, the role steps down
once more if that clears every widow in the deck and does not land on the floor: one size
for the deck is worth more than one slide's widow, and the floor is never spent on a widow.

**Floors.** A role's floor is the style's own `min`. A style that names a size and no
floor gets a readability floor instead of the house one, never above the size it asked
for: 8 pt for a caption, a page number, a table cell or header, a stat label, an
attribution or a diagram label, and 9 pt for anything carrying a sentence. A **display
role** floors at **0.85 of the size it asked for** instead (Q19): `headline`,
`dividerTitle`, `coverTitle`, `statValue` and `agendaStatement` are the shape of the slide,
so a headline that will not fit is worth an overflow finding and is not worth setting at a
body size. A style that names its own `min` still wins, and the ratio never lifts a size. So a token file
that says 9.5 pt kicker renders at 9.5, not at the house floor of 11, and the deck-wide
pass then walks that role's own range. The shipped house and plain styles are their own
token files, so they keep the built-in ranges. A ceiling never cuts a size the style asked
for either.

**Weight (Q20).** `typeScale.<role>.bold` is honoured as given, and `weight` names a face:
`"light"`, `"regular"`, `"medium"`, `"semibold"` or `"bold"`. `"regular"` and `"bold"` are
the family's own RIBBI faces, so they are written as the bare family with the bold attribute
off and on: there is no font family called `Montserrat Regular` for a renderer to resolve, and
asking for one is how a deck ends up in a substituted serif. The other weights do ship as
families of their own, so when the face is installed the builder writes it by name, the way a
designer sets it (`Montserrat Light`, `Montserrat SemiBold`), and leaves the attribute off; when it
is not installed, the family is written with bold on for semibold and off for light and
medium. A run that asks for
bold on its own (an emphasised lead sentence) drops to the base family and takes the
attribute, because bold on top of a Light face is not a bold word. An unknown weight warns
and keeps the family.

The headline's weight is the style's call: house and plain say `bold` in their own token
files, and a style that names the role without naming a weight takes the builder's default
for it rather than the house file's. The resolved typeface per role comes back as
`typeScale.faces` in the result and `typeFaces` in the snapshot, so a weight that is not
installed is visible instead of silently synthesised.

Kicker, page number, source, cover, divider and agenda type stay as the tokens set them.
The sizes come back as `typeScale` in the result and in the snapshot; a role that was
stepped all the way to its floor is listed in `typeScale.atFloor`, and `check_deck` reports
it once as a `font-floor` finding carrying `slide: 0` and `scope: "deck"`.

## Cover and agenda

`cover` and `agenda` default to true when the deck has no slide of that kind, and a
generated slide is built from the spec, never invented:

* the **cover** is a `title` slide: the deck `name` as the title, the leading clause of
  `brief.goal` (or `brief.takeaway`) as the subtitle, `date` when the spec gives one, and no
  page number.
* the **agenda** is an `agenda` slide in the style's own agenda dressing: the statement is
  `agendaStatement`, else the leading clause of `brief.goal`, else the deck name; then one
  entry per section divider, `NN Section` where the style numbers dividers, followed by the
  first sentence of the first story under that divider. Entries sit at rows 9, 11, 13 and 15
  per the agenda archetype; more than four warns, because the geometry reads four.

A generated cover is not counted: page numbers start at 1 on the slide after it. The deck's
own title slide is left as it is, numbering included, and a spec that already has a title or
an agenda slide keeps it and gets a warning if it asked for another.

**Speaker notes.** With `notesFromStory` (the default) a slide with a `story` and no `notes`
writes the story into the notes, so the Layer 1 note reaches the person presenting.

## Furniture

**Kicker (F5).** A content slide draws `kicker` when the spec gives one. Otherwise
the builder fills the style's `kicker.pattern` (default `"{section} | {topic}"`) with the
slide's `section` and its `name`; a style may set `"{section}"` or `"{topic}"` instead. An
empty value takes its separator with it, so a slide with no section reads `The prize`, not
`| The prize`. A style with `rules.kicker: false` drops the line with a warning.

`kicker.case` recases a **derived** kicker to the format's own convention (F-1):
`"title"` (harbor), `"sentence"`, `"caps"` or `"none"` (the default, and house). The
separators split the line, so each field is cased on its own; Title Case keeps the small
joining words low inside a field, and both title and sentence case leave a word that
carries its own capitals alone, so `GTM` never comes back as `Gtm`. A kicker the payload
wrote is the writer's wording and is set exactly as given.

**Headline casing (F11).** The first character of every headline, cover title, agenda
statement and divider title is capitalised; nothing else is touched. A first word that
carries its own capital (`iPhone`, `eBay`) is left alone.

**Numbered dividers (F16).** With `divider.numbered` true (the default, and house) a
divider reads `03 Competitive analysis`: the number counts section dividers in deck order,
zero padded to two digits, and lands in the snapshot as the slide's `ordinal`. A `close`,
`thanks` or `thankyou` slide draws like a divider but is not a section, so it is never
numbered. `harbor_deck.json` sets `numbered: false`, because that format shows bare titles.
The divider subtitle is unchanged.

**Headline and kicker blocks.** A `headline` or `kicker` block is furniture, not content:
it fills the slide's own field when that field is empty (and keeps its block id in the
snapshot), and is dropped with a warning when the slide already has one. Two headlines are
never drawn.

**Every slot is only as tall as its text.** A furniture box is drawn at the height of the
copy it holds, clamped to the slot. A style with a deep headline slot therefore reserves
the band it uses, not the band it declares.

**The furniture guard.** After the furniture is drawn, every block is measured against it.
A block that intrudes into a band above it (kicker, headline, hairline, cover or agenda
furniture) drops to the next whole row below that band, keeping its bottom edge. A block
that runs into a band below it (the source line, only when the slide has a source, and the
page number) moves up by whole rows when it can keep its height and nothing else is in the
way, and is trimmed when it cannot. Each move warns and names the block and the furniture.
Rows stay whole, so columns on one slide keep their alignment.

## Diagrams

`diagrams` picks the mode: `"shapes"` (the default) draws the diagram, `"placeholder"`
draws the dashed labelled box a format that bans shape diagrams mandates, naming the kind
and every node label. `builder/styles/harbor_deck.json` is `"placeholder"`.

Geometry comes from `builder/diagram_layout.py`, pure functions that take the block
rectangle in points and return every node box and connector polyline, so the canvas sketch
and the deck agree. `plan(content, rect, opts)` returns
`{kind, direction, fontPt, nodes, edges, axis, warnings}` plus `ring` for a cycle and
`root` for a tree.

| kind | layout |
|---|---|
| `flow` | Layered: rank by the longest path from the sources, order within a rank by barycenter, `LR` (default) or `TB`. Elbow connectors, straight when the centres line up |
| `cycle` | Nodes on a ring from the top, clockwise. Ring arrows are drawn for you when the spec gives no edges |
| `timeline` | Nodes on a horizontal spine, boxes alternating above and below, dates in the labels. The spine spans the first node to the last |
| `tree` | Layered top to bottom from the root (the first node with no incoming edge) |
| `matrix` | A square 2x2. `nodes[0]` labels the x axis, `nodes[1]` the y axis, and the next four fill the cells top-left, top-right, bottom-left, bottom-right. Edges are dropped with a warning |

Nodes draw as the style's card (`card.shape`, `fill`, `borderColor`, `borderPt`,
`radiusAdj`) with the label in the body font at the `diagramNode` role; matrix axis labels
draw as text at the `diagramLabel` role; connectors take the `diagram.edgeColor` line with
an `a:tailEnd` triangle; an edge label is caption-sized and sits clear of the line at its
midpoint.

**Fit.** The node width comes from the longest label at the `diagramNode` size, grown to
fill its share of the block and capped by `diagram.nodeMaxW`; the height is the taller of
the text and its share of the slot, never taller than the box is wide. When a label will
not fit, the type steps down one point at a time to the `diagramNode` floor (8 pt, the
caption floor). At the floor the label is truncated with an ellipsis and a warning. Over
12 nodes the list is truncated with a warning. A block too small for any diagram, or one
with no nodes, draws the labelled placeholder instead.

## Charts

A chart earns its place by making a comparison the reader could not make from the numbers
(Zelazny, Few, Knaflic), and everything that is not that comparison comes off (Q15):

* **No legend for one series.** A legend that names the only thing on the chart is a label
  the reader has to walk back to. Two or more series keep the legend, at the bottom.
* **Gridlines run horizontally or not at all.** The value axis carries them on a `column`,
  `line` or `stacked` chart; a `bar` chart's value axis is the horizontal one, so its
  gridlines would be vertical rules and it gets none. The category axis never has them.
  `chart.gridlines: false` turns off what is left.
* **Data labels on when the reader can count the marks**: 8 points or fewer across all
  series. A pie's labels carry the category name, so a pie has no legend either.
* **One accent against muted others.** `content.emphasisSeries` (a name or an index) picks
  the series that carries the point, with or without `emphasis: true` on the block.
* **Bar charts are ranked**, largest first as the reader sees it: a single-series `bar` or
  `column` chart is sorted descending unless its categories are a time series (a year, a
  quarter, a month, a week, a phase), which keeps its own order. PowerPoint plots the first
  category at the bottom of a bar chart and at the left of a column chart, so the data order
  is reversed for one and not the other. The snapshot entry carries `sorted: "descending"`.
* **A chart title never repeats the title above it.** `content.title` is drawn as the chart
  title unless it repeats the block title (or, when the block has none, the headline), and
  is dropped with a warning when it does.
* **Number formats come from the values.** A value may arrive written the way the deck says
  it: `"7.7%"`, `"$60K"`, `"4.1x"`, `"1,200"`. The number drives the chart and the unit
  drives the axis and label format (`#,##0.0"%"`, `"$"#,##0`, `0"x"`), which comes back on
  the snapshot entry as `numberFormat`. Axis lines and tick labels are muted, with no tick
  marks.

## Style modes

**plain**: Arial, charcoal ink `1F2937`, one accent `3B5BDB`, square cards, no furniture
beyond the headline and the page number. A kicker is dropped with a warning.

**house** (the default): `builder/styles/house_deck.json`. Neutral consulting: kicker, action
headline, hairline, rounded 1 pt cards, muted grey body, source line, page number.

**skill**: `style.tokens` or `style.tokensPath`, merged over house, so a partial token file
still renders. `builder/styles/harbor_deck.json` is the Harbor format derived
from the `harbor-deck` skill: Lato with an Arial fallback, 2.25 pt `1E3A5F` roundRect
cards, sky-blue dividers, dark-blue title slide, slate agenda, page number top right.

**template** (Q8): opens the .pptx or .potx, keeps its master, layouts and theme, drops its
slides, and then **maps its layouts to the slides** rather than drawing over them. An
unreadable template falls back to house with a warning.

* Every layout is profiled by the placeholders it has: `titleSlide` (a title and a subtitle),
  `section` (a title, and "section", "divider", "chapter" or "part" in its name), `twoContent`
  (a title and two or more body or picture placeholders), `titleBody`, `titleOnly`, `blank`.
  Date, footer, header and slide-number placeholders are ignored.
* Each slide asks for what it needs: a title slide wants `titleSlide`, a divider `section`, a
  content slide with no blocks `titleOnly`, with one text-like block `titleBody` and with two
  or more `twoContent`. A slide holding an image wants a picture placeholder more than it
  wants any of that, and the best-scoring layout wins, ties going to the simpler one.
* The **headline fills the title placeholder** (the divider title on a divider, the cover
  title on a cover), the slide's subtitle the subtitle placeholder, the **first image block**
  the picture placeholder and the **first text or bullets blocks** the body placeholders, one
  each. A filled block is not drawn a second time; every block that is left is drawn as usual
  at its cells, guarded against the placeholders the template just filled, whose rectangles
  are bands like any other furniture (`role: "templateBody"`, `"templatePicture"`).
* Only a content slide hands its blocks over: a cover, a divider and an agenda fill the title
  and the subtitle and keep their own blocks.
* **The master's theme wins.** Its major and minor fonts become the display and body
  families, and its colour scheme becomes the palette: `dk1` the ink, `lt1` the background,
  `accent1`/`accent2` the accents, with a muted derived from the pair. House's per-kind
  backgrounds are dropped with it, because the master dresses every slide kind. The body
  measure (and so the column module) comes from the layout most content slides will take.
* Wherever the template drew the title, the builder draws no furniture of its own beyond the
  chrome, the kicker and the page number, and leaves the master's background alone. Unused
  placeholders are removed.

**Fonts.** Every role resolves its family from the style's own `fonts` first, so a table
cell, a stat label, a diagram label and a page number are set in the style's faces and
never fall through to a house family. A weight or slant suffix resolves to its base family
before the installed check, so `"Montserrat Light"` on a machine with Montserrat Regular writes
`Montserrat` and measures Montserrat; a face installed under its own name (`"Montserrat SemiBold"`)
is written as asked. A family that is not installed tries the style's other faces (body,
then display) before `fonts.fallback` and then Arial, never a serif. Only the `mono` slot
may be a mono face, and only when the style itself names one: otherwise a mono role is set
in the body font. Every substitution warns once.

**Derived tokens.** A partial token file inherits house, which is a light style, so a
style that names a dark background derives its own values instead: its rules become
`cardBorder`, the takeaway bar becomes `cardFill`, `line` (the hairline, chart gridlines
and table rules) becomes `cardBorder`, and the header fill it never named is dropped
rather than inherited. A palette whose `muted` equals its `ink` gets a muted of 65% ink
over the background, so body copy and captions read as a second voice. A style that lays
out its own `furniture` and never mentions `hairline` does not get the house one drawn
through its headline.

**Background scopes.** House dresses its own dividers, title and agenda, and none of that
belongs to a style that brought its own palette. So for every slide kind the style itself
scopes, a key it did not name is dropped and resolves from the style's palette instead:
`accent` from `palette.accent`, `line` from `palette.line`, and so on. `muted` is the one
key that is derived rather than dropped, as 65% of that kind's ink over that kind's fill,
because the palette's muted was mixed for the palette's background and not for this one.
A style that names `palette.bg` and scopes no content fill has said what colour its slides
are, so the content background follows the palette; the other kinds keep their own scoping.

## Style token schema

```jsonc
{
  "name": "Harbor",
  "palette": { "bg", "ink", "muted", "accent", "accent2", "line", "cardFill", "cardBorder",
               "takeawayFill", "inverse", "series": ["...", "..."] },   // 6 hex digits
  "fonts": { "display", "body", "mono", "fallback" },
  "card": { "shape": "roundRect" | "rect", "radiusAdj": 8000, "borderPt": 2.25,
            "borderColor": "accent", "fill": "cardFill", "insetPt": 14, "gutterPt": 27 },
  "table": { "headerFill", "ruleColor", "rowHeightPt",
             "labelColumn": { "case": "caps" | "none", "color": "accent" | "ink" | "muted",
                              "size": 8.5, "bold": false } },
  "chart": { "legend": "bottom", "gridlines": true },
  "diagrams": "shapes" | "placeholder",          // default shapes
  "diagram": { "gutterPt": 18, "insetPt": 6, "nodeMinW": 54, "nodeMaxW": 260,
               "nodeMinH": 26, "nodeMaxH": 130, "maxNodeLines": 2, "edgePt": 1.25,
               "edgeColor": "muted" },
  "kicker": { "pattern": "{section} | {topic}",
              "case": "title" | "sentence" | "caps" | "none" },   // derived kickers only
  "divider": { "numbered": true },
  "stat": { "rule": "above" | "below" | "none" },   // default below

  "chrome": [ { "kind": "dots" | "line" | "wordmark" | "logo",
                "x", "y", "w", "h", "color", "fill", "linePt",
                "count": 3, "text": "HARBOR", "size": 9, "path": "/abs/logo.png",
                "kinds": ["content", "divider"] } ],
  "typeScale": { "<role>": { "size", "min", "max", "maxLines", "bold", "italic", "caps",
                             "weight": "light|regular|medium|semibold|bold",
                             "font": "display|body|mono", "color", "leading", "k", "align" } },
  "furniture": { "<slot>": { "x", "y", "w", "h", "align" } },
  "backgrounds": { "content|title|divider|agenda": { "fill", "ink", "muted", "accent",
                                                     "accent2", "line" } },
  "copy": { "wordmark", "agendaEyebrow" },
  "rules": { "kicker", "pageNumber", "source", "pageNumberOnTitle" }
}
```

Any colour is a palette key or a hex value. A `backgrounds` entry overrides palette keys for
every shape on a slide of that kind, which is how white text on the blue divider works
without a second palette.

Roles: `coverTitle`, `coverSubtitle`, `coverDate`, `wordmark`, `agendaEyebrow`,
`agendaStatement`, `dividerTitle`, `dividerSubtitle`, `kicker`, `headline`, `statement`,
`blockTitle`, `body`, `bullets`, `cardHeader`, `cardSub`, `cardBody`, `statValue`,
`statLabel`, `quote`, `attribution`, `tableHeader`, `tableCell`, `caption`, `pageNumber`,
`takeaway`, `placeholder`, `diagramNode`, `diagramLabel`. Defaults are the ranges in
deck-planning section 7.

Furniture slots: `kicker`, `headline`, `hairline` (null to drop it), `body`, `takeaway`,
`source`, `pageNumber`, `dividerTitle`, `dividerSubtitle`, `coverWordmark`, `coverTitle`,
`coverSubtitle`, `coverDate`, `agendaEyebrow`, `agendaStatement`, `agendaList`. A slot may
also be nested under a slide kind (`"furniture": {"divider": {"pageNumber": {...}}}`) to move
it on that kind only.

**Derived slots (Q15).** A style that lays out its own `furniture` has said where its slides
start and stop, so house's 60/840 divider title is a fact about house and not about it. Every
divider, cover and agenda slot such a style did not name is derived from the measure it did:
`furniture.body` when it has one, otherwise the headline column, with the band running from
the body's top (or the headline's bottom) to the body's bottom (or the source line). The
divider title sets **top-left in that band** at the measure's left edge, its subtitle under it
(F-5): a section title heads the band it opens rather than hanging off its floor. The cover
title, subtitle and date still bottom out together; `agendaEyebrow` takes the kicker slot,
`agendaStatement` the headline slot and `agendaList` the body. Each box is as tall as two
lines of its role plus a quarter of an em of descender slack. A style that brought no
furniture of its own *is* house's, and nothing is derived. Type is untouched: the sizes are
whatever the merged `typeScale` says.

**Chrome (Q15).** `chrome` is the list of marks a format repeats on every slide and that are
neither content nor a type role: the three circles a dark deck carries top left, the
hairline under its kicker row, a wordmark, a logo. Each entry is a rectangle in points plus a
`kind`:

| kind | draws |
|---|---|
| `dots` | `count` circles (default 3) of diameter `h`, spread across `w` with equal gaps. `fill` fills them, otherwise they are outlined in `color` at `linePt` (1 pt) |
| `line` | a filled rule, at least 0.75 pt tall |
| `wordmark` | `text` (else `copy.wordmark`) at the `wordmark` role, at `size` when the entry names one |
| `logo` | the image at `path`, fitted inside the box. A missing file warns and draws nothing |

An entry draws on content slides; `kinds` names the slide kinds it draws on instead. The
marks go down before the furniture, so the guard keeps blocks off them, and every one is a
snapshot shape at `role: "chrome"` with no block id. `check_deck` never counts chrome in the
overlap check. House and harbor define none, and the extractor fills them in.

## Output

```json
{ "ok": true, "path": "...", "snapshotPath": "...", "pdfPath": "...",
  "thumbnails": ["/.../deck.slides/slide-1.png"],
  "slides": [ { "index": 1, "title": "...", "shapes": 8 } ],
  "fidelity": "balanced",
  "typeScale": { "headline": 27, "body": 11, "bullets": 14, "cardHeader": 15,
                 "cardSub": 14, "cardBody": 12, "tableHeader": 11, "tableCell": 10,
                 "statValue": 44, "statLabel": 12, "atFloor": ["cardHeader"],
                 "faces": { "headline": "Montserrat Light", "body": "Montserrat" } },
  "adjusted": [ { "slide": 4, "blockId": "b_take",
                  "from": { "col": 11, "row": 7, "cols": 20, "rows": 5 },
                  "to": { "col": 11, "row": 9, "cols": 20, "rows": 5 },
                  "reason": "centred in the body band" } ],
  "warnings": ["..."], "errors": [] }
```

`adjusted` only appears when there is something to report, and `atFloor` only when a role
was stepped all the way down. `typeScale.faces` is always there: it is the typeface every
role really landed on, which is not always the one the tokens named (Q20). A generated cover or agenda is a slide like any
other in `slides`, counted in `index`.

`pdfPath` only with `exportPdf` and LibreOffice; `thumbnails` only when the PNGs were
rendered. Both are skipped with a warning, never an error, when the tools are missing.

**PDF and PNG**: `soffice` is looked up in `$BLUEPRINT_SOFFICE`, then on `PATH`, then at
`/opt/homebrew/bin`, `/usr/local/bin` and inside LibreOffice.app. The deck converts to PDF
with `--headless --convert-to pdf` under its own user profile and a timeout; the PDF becomes
PNGs with `pdftoppm -png -r 96` (`$BLUEPRINT_PDFTOPPM` overrides the lookup), falling back to
one `soffice --convert-to png` page when pdftoppm is missing.

## Build snapshot

Alongside the deck, `<outputPath minus .pptx>.snapshot.json`, reported as `snapshotPath`. The
canvas draws a vector preview from it without LibreOffice.

```jsonc
{
  "name": "Referral Growth Review",
  "slideSize": { "w": 960, "h": 540 },
  "grid": { "cols": 32, "rows": 18, "cell": 30 },
  "style": { "mode": "house", "name": "House" },
  "fidelity": "balanced",
  "typeScale": { "headline": 26, "body": 12, "bullets": 14 },
  "typeScaleFloored": ["cardBody"],
  "typeFaces": { "headline": "Montserrat Light", "body": "Montserrat", "pageNumber": "Menlo" },
  "slides": [
    {
      "index": 3, "id": "s_stall", "name": "Where referrals stall", "kind": "content",
      "section": "Why referrals",
      "headline": "Referrals stall at three moments, and each one has a different fix",
      "shapes": [
        { "id": "s3-2", "blockId": null, "role": "headline", "x": 60, "y": 66,
          "w": 840, "h": 84, "kind": "text", "text": "Referrals stall at three...",
          "fontPt": 26 },
        { "id": "s3-5", "blockId": "b_stall_cards", "role": "card", "x": 60, "y": 180,
          "w": 260, "h": 240, "kind": "card", "text": "No moment Timing Partners...",
          "fontPt": 17 }
      ]
    }
  ]
}
```

One entry per drawn shape. Furniture has `blockId: null`; a `cards` or `stat` block has one
entry per item; a `notes` block has one entry of kind `notes` that nothing draws, so a preview
skips it. `text` is the first 80 characters.

A slide whose blocks moved carries `adjusted: [{blockId, from, to, reason, snappedPt?}]`, and
every shape of a moved block repeats its own `{from, to, reason, snappedPt?}` under
`adjusted`, so a preview can badge the shape without looking anything up. `snappedPt` is
`{x, w}` in points, present when the column module did not divide the grid evenly, and it is
what was drawn. Every shape of a block the plan marked `emphasis: true` carries
`emphasis: true`. Every shape of a `stat` block carries its `layout`, `"band"` or `"grid"`
(Q22). `typeScale` is the deck-wide sizes, `typeScaleFloored` the roles that were stepped to
their floor, and `typeFaces` the typeface every role landed on. `kind` is one of `text`,
`card`, `stat`, `quote`, `takeaway`, `chart`, `table`, `image`, `placeholder`, `rule`,
`notes`, `node`, `edge`. A chrome mark is a `rule` (or an `image` for a logo) at
`role: "chrome"`; a chart entry carries `sorted` and `numberFormat` when it has them. A hairline, an accent bar and a table rating mark are `rule` entries
at `role: "hairline"` or `role: "tableRating"`; the hairline a figure-grid pair carries is a
`rule` at `role: "statRule"`. Every `rule` a stat block draws carries `rulePosition`,
`"above"` or `"below"` (F-5).

A `diagram` block records exactly one `node` entry per node and one `edge` entry per edge,
every one carrying the block id: a node box is `kind: "node"`, `role: "diagramNode"` (a
matrix axis label is `role: "diagramLabel"`), a connector is `kind: "edge"`,
`role: "diagramEdge"` and carries its polyline as `points: [[x, y], ...]` with the label in
`text`, an edge label is a `text` entry at `role: "diagramEdge"`, and a timeline spine is a
`rule` entry at `role: "diagramAxis"`. A divider slide carries its `ordinal`.

## check_deck.py

`python3 builder/check_deck.py <deck.pptx> [<out.json>] [--spec <spec.json>]`. Reads the deck,
the sibling snapshot (for block ids and roles) and the spec when given (for the style, so the
font floors are the ones the deck was built with). Exit 0 whenever the check ran, 1 only if it
crashed.

```json
{ "ok": true, "slides": 8, "shapes": 46,
  "findings": [ { "slide": 4, "blockId": "b_over", "role": "body", "kind": "overflow",
                  "severity": "warning", "message": "block 'b_over' needs 7 lines and the box holds 4" } ],
  "summary": { "errors": 0, "warnings": 2 } }
```

| kind | severity | fires when |
|---|---|---|
| `off-slide` | error | a shape sits outside the slide rectangle |
| `overlap` | error | two content shapes cover more than 10% of the smaller one |
| `missing-headline` | error | a content slide has no headline shape |
| `empty-slide` | error | a content slide holds nothing but furniture: every block went, or it never had one. Names the slide, and the fix is to add a block or drop the slide |
| `overflow` | warning | the measured text is taller than its box |
| `font-floor` | warning | a role whose default is above its floor was shrunk to the floor. A deck-wide role fires once, with `slide: 0` and `scope: "deck"` |
| `empty` | warning | an empty box, or one that only holds `[placeholder]` copy |
| `widow` | warning | a paragraph whose last line is one word |
| `diagram-edge` | error | a diagram edge names a node that does not exist (needs `--spec`) |
| `diagram-overlap` | error | two node boxes in one diagram cover each other |
| `diagram-font` | warning | a diagram's labels sit at the readable floor |
| `kicker-drift` | warning | a slide's kicker does not follow the style's own `kicker.pattern`: it carries a separator but not the pattern's shape. A single field is not drift |
| `headline-case` | warning | a headline is title case, all caps, or ends in a full stop. A word with its own capitals (`AI`, `MetaDesign`) is not title case |
| `not-parallel` | warning | card bodies in one row do not close alike, or their line counts differ by more than one |
| `empty-band` | warning | more than 4 empty rows under the last block on a content slide. Strict only: balanced fills the band |
| `source-missing` | warning | fewer than half the slides that show figures carry a source line. One finding, `slide: 0`, `scope: "deck"` |
| `bullet-count` | warning | more than 7 bullets in one block |
| `duplicate-headline` | warning | two consecutive slides carry the same headline |

Every consistency finding names the slide, and the block where there is one, and ends with
the one line that fixes it. The spec slide behind a built slide is found by the snapshot's
slide `id`, not by position, because a generated cover and agenda shift every index; a
generated slide has no spec slide and is skipped by the spec-based checks.

Backgrounds (over 90% of the slide), hairlines, chrome marks, accent bars and table rating
marks are excluded from the overlap check; charts, tables and pictures are excluded from the text
checks. Diagram connectors,
edge labels and axis rules cross the boxes they belong to, so they are excluded from the
overlap check, and every diagram shape is checked as a diagram rather than as loose text.

## Tests

`builder/test_diagram_layout.py` (29 tests): every kind's geometry inside its rectangle,
longest-path ranking, barycenter ordering, even ring spacing, the square matrix cells, the
fit step-down and the node cap, and that its text measurement matches build_pptx's.

`builder/test_build_pptx.py` (382 tests, plain unittest, no pytest): builds a five-slide spec covering
every block type in all four style modes, then asserts geometry within 1 pt, one snapshot
entry per block, divider and title behaviour, page numbers, speaker notes, the card and token
styling per mode, template layouts, theme fonts and slide size, the fit ladder, overflow,
clamping, empty copy, and the PDF and PNG steps both with LibreOffice and with it patched out.
`CheckDeckTests` plants an overlap, an off-slide shape, a missing headline, overflow, an empty
box and a widow. `DiagramBlockTests` builds all five kinds, counts one shape per node and per
edge, reads the arrowhead XML back, checks placeholder mode and the diagram findings.
`FurnitureTests` covers the kicker pattern, headline casing, numbered dividers and the
furniture guard; `DarkStyleTests` covers the derived tokens and the font fallback.
`SampleDeckTests` builds `builder/sample_deck_spec.json` (nine slides, one of them a
diagram) in house and in the Harbor style. `FidelityTests` covers the three fidelities: strict
leaves the plan alone, balanced grows a block and shrinks the neighbour with slack, rebalances
an empty band, centres one row of blocks and snaps near edges, loose re-picks a seven-column
table, five cards and a wordy stat, and refuses a re-pick that would collide; every deviation
is checked in the result and in the snapshot. `FitModuleTests` covers `builder/fit.py` on its
own. `DeckTypeTests` covers the deck-wide sizes: one headline size on every slide that fits
the longest headline, one bullet and one stat size across the deck, the card roles stepping
together, the floor, the widow step and the single `font-floor` finding. `ChromeTests` covers
the story becoming the notes, the generated cover and agenda, a deck that brings its own, and
page numbers skipping the cover. `TypeScaleTests` covers the readability floors and a token
size the house floor used to lift; `FontChainTests` covers the style's own fonts winning,
mono falling back to the body face, and a built deck with no house font in it;
`EmphasisTests` covers the six treatments, the snapshot marks and a plain block getting
none of them; `TableStyleTests` covers the Tufte defaults, what counts as a number, the
label column, the rating circles and the natural row height; `ModuleTests` covers the
module maths, three columns landing on it, the recorded `snappedPt`, strict leaving it
alone and two snapped blocks not colliding; `BackgroundMergeTests` covers a scoped
background taking its missing keys from the palette; `ConsistencyTests` plants each of the
seven consistency findings and a clean deck that fires none; `SpecSlideMappingTests` builds
a deck with a generated cover and agenda and checks the diagram finding lands on the right
slide.

`DisplayFloorTests` covers the 0.85 display floor and a text role keeping the readability
one; `WeightTests` covers `weight` resolving to an installed face, falling back to the
family, the RIBBI weights staying on the bare family, a run overruling the role,
and the faces landing in the result, with the installed font set patched so the answers do
not depend on the machine; `StatBandTests` covers loose freeing a stat band a text rail
boxed in, the rail moving under it, balanced leaving the pair alone and a table neighbour
not counting as a rail; `StatGridTests` covers the auto-pick, the block asking for either
layout, the 0.6 number size, the body label, one hairline a pair and the reading order;
`StatRuleTests` covers `stat.rule`: the default and the house token saying below, a style
asking for above, an unknown value warning, the band hairline and the grid hairlines each
following it, `rulePosition` in the snapshot, the block's footprint not moving, and a plain
band staying unruled;
`LabelColumnTests` covers the house default, a caps-and-accent format, the accent giving way
on white and a data first column being left alone; `KickerCaseTests` covers the four cases,
an acronym surviving, an explicit kicker going in as written and an unknown case warning;
`SourceRowTests` covers row 17 being kept for the source line, a whole row group moving
together, a block that cannot move being named, and growth never taking the row back.

`ChromeTokenTests` covers the four chrome kinds, the count and spacing of the dots, the
snapshot role, the default content-only scope and an entry that names its own kinds, a
missing logo, an unknown kind and a block guarded off the chrome band; `DerivedFurnitureTests`
covers a style with its own furniture deriving its divider, cover and agenda slots from its
body (and from its headline column when it has no body), the divider title heading the band
and its subtitle sitting under it, the cover group still bottoming out, a slot the style
named being left alone (title or subtitle), house keeping its own coordinates, and
`agendaList` placing the generated agenda entries; `SourceLineTests` covers a source left alone, brackets
going first, trailing clauses going next, the full text reaching the notes, the reported
adjustment and the one-line test at the caption floor; `DuplicateContentTests` covers a first
paragraph dropped, a block that says nothing else dropped whole, a paragraph that merely
shares words kept, twin titles warning and strict keeping all of it; `EmptySlideTests` covers
the `empty-slide` error on a slide the duplicate pass emptied and on one that never had a
block, the builder's warning still firing, a folded headline block not counting as content,
and a divider, a slide that kept a block and strict all firing nothing; `ParallelRowTests` covers
a text block taking the row's marker, its title staying a title, the report, a row of only
text and a text rail beside a table both left alone; `ChartDisciplineTests` covers the legend
rule, horizontal gridlines, the data-label threshold both ways, the ranking of a column and of
a bar chart, time categories keeping their order, the chart title dropped and kept, the value
parser and axis formats, and a named series taking the accent; `NotesFloorTests` covers the
last paragraph moving to the notes, the first staying, nothing cut mid-sentence, strict moving
nothing and the snapshot entry; `TemplateLayoutTests` builds a template with two layouts of
its own and covers the layout picked per slide (title and body, two content, title only,
section, title slide, picture), the placeholders filled, the blocks left drawn as usual, the
body measure and the theme colours winning.

The fixtures that assert the approved geometry pass `"fidelity": "strict"` with `cover` and
`agenda` off, so the plan they assert is the plan that is drawn.
Run: `python3 builder/test_build_pptx.py`.

# Design system

The canvas is dark first, Raycast-dense: compact rows, hairline borders on dark surfaces, one restrained
accent, no gradients, no glass, no glow, no purple. Three files hold it:

| File | Holds |
|---|---|
| `src/app/tokens.css` | every colour, size, radius, shadow, z-index, duration |
| `src/app/styles.css` | one class family per component, built only from tokens |
| `src/app/motion.ts` | `flip`, `fadeIn`, `slideIn`, `pulse`, `toast` |

Rule: no literal colour, duration or font size anywhere but `tokens.css`. Views set geometry inline (cell
maths) and nothing else; anything visual is a class.

## Tokens

**Colour.** Surfaces `--bg0` (page) to `--bg4` (raised), plus `--hover` and `--sel` washes. Hairlines
`--line1` faint, `--line2` default, `--line3` strong. Ink `--fg1` (primary) to `--fg5` (faintest).
Semantic: `--accent` / `--accent-soft` / `--accent-ink`, `--ok`, `--warn`, `--bad`, each with a `-soft`
fill, `--chart` for chart sketches, `--value` for a number the user supplies, `--link` for a cell that only
passes another sheet's value through, `--btn-bg` / `--btn-fg` for the filled button. `:root` is dark;
`:root[data-theme="light"]` overrides the same names, so no rule is written twice.

**Light.** Light keeps dark's ordering and reverses only its direction: `--bg0` is the page and every step
after it is whiter, so "further forward" still means "further from the page". `--bg4` is the single
deliberate flip: white has nothing above it, so the active/pressed tone (button hover, segmented on, a
block's header) goes the other way. Two consequences worth knowing before touching the light block:

- A surface that reads in dark must read in light at the *same* ratio. Every neighbouring pair is at least
  1.05, `--bg0`/`--bg3` is 1.20, and `--bg2`/`--bg4` is 1.20, which matches or beats dark on every pair.
  The pre-DS5 light block had `--bg1` = `--bg3` and `--bg0` ≈ `--bg2`, so a button vanished into the side
  panel, a card vanished into a slide, and the built sheet had no body on the canvas.
- Washes and hairlines are alpha, so light needs more of it than dark to land on the same ratio: `--hover`
  .06 / `--sel` .10, hairlines .10 / .16 / .26. The faint inks (`--fg4`, `--fg5`) and the semantics are
  darker than a naive inversion, so an 11 px label, a keycap and a chip on its own `-soft` fill all clear
  4.5:1, the ratio an 11 px label needs and the thing a bright background loses first.

**Type.** System stack (`--font`), SF Mono (`--mono`) for numerals, addresses and keycaps, always with
`font-variant-numeric: tabular-nums` (use `.mono` or `.num`). Scale `--fs-1` 11px (labels, keycaps, meta),
`--fs-2` 12px (secondary), `--fs-3` 13px (body, the default), `--fs-4` 15px, `--fs-5` 18px. Line heights
`--lh-1` 1.3 and `--lh-2` 1.45. `.label` is the uppercase 11px field label. The dense canvas needs two
sizes below the scale, and uses them nowhere else: `--fs-micro` 10px (cell text in a block or a built
sheet) and `--fs-tag` 9px (block type tags, arrow labels).

**Spacing.** 4-pt scale `--sp-1` … `--sp-8` (plus `--sp-0` 2px for icon gaps). Control heights
`--h-ctl` 28, `--h-ctl-sm` 24, `--h-row` 28, `--h-hdr` 46, `--h-ftr` 48, `--h-bar` 36, `--h-strip` 38.

**Radii.** `--r-sm` 5 (controls, rows, blocks), `--r` 7 (popovers, cards, segmented), `--r-lg` 10 (the app).

**Elevation.** `--e1` a block lifted off the grid; `--e2` anything floating over the canvas (popover,
minimap, toast, a dragged row). Nothing else casts a shadow.

**Z ladder.** `--z-canvas` 1 → `--z-block` 5 → `--z-drag` 6 → `--z-marquee` 7 → `--z-rowhdr` 11 /
`--z-colhdr` 12 / `--z-corner` 13 → `--z-float` 20 (minimap) → `--z-pop` 30 → `--z-toast` 40.

**Motion.** `--dur-1` 120ms (hover, colour), `--dur-2` 180ms (appear, toast in), `--dur-3` 220ms (move,
resize, panel). `--ease-out: cubic-bezier(0.2,0,0,1)` for anything arriving, `--ease-in-out` for anything
reversing or leaving. No bounce, no spring.

## Components

One class family each. Usage line per family:

| Family | Use it for |
|---|---|
| `button` + `.primary .ghost .icon .danger .small .confirm .on` | every action; `.primary` is the one commit per view, `.ghost` for toolbars, `.confirm` is the armed half of a two-step |
| `.key` | a keycap inside a button or hint (`⌘↩`, `esc`) |
| `.chip` + `.ok .warn .bad .line .num` | a count or state read at a glance; never a control unless it is a `button.chip` that expands a list |
| `.dot` + `.ok .warn .bad .none` | the 6px state marker in front of a row, chip or note |
| `.seg` | the stage control and the zoom steps: two to four exclusive options |
| `.pop` + `.confirm .cheats` | anything anchored: menus, the style and template sheets, a confirm, the `?` sheet |
| `.toast` | one transient line at the bottom of the canvas, six words or fewer |
| `.side` + `.sh` `.scroll` | the right panel: header bar with a title, scrolling body of fields |
| `.row` (`.trow`, `.sheetrow`, `.blocklist .b`) | one compact line in a list: 28px, hover wash, `.sel` for the selected one |
| `.field` + `.label` | a labelled control in the panel or a popover |
| `.card` | a small boxed aside inside a field (detection summary, an explanation) |
| `.notes` `.note` (thread) | one line per entry: dot, marker word, text, hover actions; the answer box sits inline under an open question |

Also: `.blk` (a planned block on the grid), `.sc` (a built cell), `.canvaswrap`/`.canvas` (scroll
footprint and sheet), `.minimap`, `.chips` (columns and row labels editor), `.warnlist` (one line per
flagged item).

**Snapshot kinds.** The two built views draw straight from the build snapshot, so every kind the builder
records needs a rule or it falls back to a plain box. A workbook cell is `.sc.k-<kind>`: `title`, `header`,
`label`, `input`, `formula`, `total`, `note`, `text`, `readme`, `link` (`--link`, underlined; a bare
cross-sheet reference, never a formula). A slide shape is `.sprev .shp.k-<kind>`: `card`, `table`, `chart`,
`image`, `placeholder`, `stat`, `text`, `quote`, `takeaway`, `node` (a diagram node, drawn as a card),
`edge` (only the connector's bounding box is recorded, so it is a dashed hint and never a filled shape),
plus `rule`; `notes` is dropped before drawing. A slide itself is `.sprev.k-<kind>`, where `title` and
`divider` differ and `content` and `agenda` share the default.

## Host

The canvas is a sandboxed iframe, so anything the browser gates on a permission has to be declared on the
resource. `src/server/server.ts` sets `_meta.ui` to `prefersBorder: true` and
`permissions: { clipboardWrite: {} }`; without the second one `navigator.clipboard.writeText` is blocked in
Claude Desktop and `controller.copyText()` falls through to its `document.execCommand("copy")` path. The
spec still says not to assume the permission is granted, which is why the fallback stays.

## Motion rules

Every view calls these; no view writes its own transition.

| Helper | When |
|---|---|
| `flip(container, mutate)` | anything keyed by `data-id` that moves or resizes: tab reorder, Stack, Flow, align, distribute, undo, redo, a drop. `main.ts` wraps the whole re-render in it, so a row or block that ends up somewhere new slides there. |
| `fadeIn(el)` | something appearing in place: a popover, the panel's contents after a selection, an expanded list |
| `slideIn(el)` | the side panel when a view opens |
| `pulse(el)` | the one block or cell Claude just changed (replaces the old 2.4s flash) |
| `toast(root, text)` | the status line, with its own enter and exit |

`prefers-reduced-motion: reduce` turns every helper into an instant state change, and the stylesheet
clamps CSS transitions to 1ms. Durations and easings are read from the tokens at runtime, so changing
`--dur-3` changes the motion everywhere.

## Copy rules

Labels three words or fewer. Buttons are verbs. No sentences in the chrome and no help paragraphs: the
`?` sheet holds the shortcuts and the only explanations. Toasts six words or fewer. Readiness warnings and
empty states are one line. Section labels drop "on" and "the" ("Thread", not "Notes on this tab").

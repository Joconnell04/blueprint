---
name: deck-diagrams
description: Rules for a diagram on a deck slide. Use any time a slide needs a diagram, flow, process, cycle, loop, timeline, roadmap, tree, hierarchy or 2x2 matrix, before writing the `diagram` block's content for build_deck. Covers which kind fits the relationship, how many nodes, how to word node labels and edges, and what the builder does in a format that bans shape diagrams.
---

# Deck diagrams

A diagram earns its place when the relationship between things is the point. If
the point is the list of things, use cards or bullets instead.

## One idea per diagram

A diagram shows one relationship. Two relationships means two diagrams, or one
diagram and one takeaway bar. Never draw a second idea into the corner.

The diagram never repeats the headline. The headline says the so-what; the
diagram shows the mechanism. If the diagram would read as the headline in
boxes, cut the diagram.

## Choose the kind by the relationship

| The relationship | kind | direction |
|---|---|---|
| One thing leads to the next: a sequence, a process, a hand-off | `flow` | `LR` |
| The end feeds the start: a loop that keeps running | `cycle` | n/a |
| Events sit on a date line | `timeline` | n/a |
| One thing breaks into parts: a hierarchy, an owner tree | `tree` | `TB` |
| Two independent axes sort four options | `matrix` | n/a |

Reading order: `LR` for a process, `TB` for a hierarchy. A process that runs
`TB` reads as a hierarchy and confuses the room, so only switch when the slide
is tall and narrow.

## Size

- 3 to 7 nodes. Under three is a sentence; over seven is a table.
- 12 nodes is the hard cap. The builder drops the rest and warns.
- A `matrix` takes exactly six nodes: the x-axis label, the y-axis label, then
  the four cells in reading order (top-left, top-right, bottom-left,
  bottom-right).

## Words

- Node labels: four words at most, and parallel grammar across every node.
  "Spot the moment / Script the ask / Hand off warm" reads; "Spotting /
  The script / Warm hand-off from the partner lead" does not.
- Start a node with a verb in a `flow`, `cycle` or `tree`; a `timeline` node
  carries its date in the label ("Q2 rollout"); a `matrix` cell is a noun.
- Edges carry verbs, one or two words: `prompts`, `warms`, `owns`, `feeds
  back`. An edge with no verb is decoration; leave its label out.
- No sentences, no punctuation at the end of a label, no "etc.".

## What never goes in

- Decorative shapes: no chevrons, no funnels, no arcs, no icons, no numbered
  medallions. The builder draws node boxes and connectors from the style
  tokens, and that is the whole vocabulary.
- Repeated furniture: no title inside the diagram, no legend for five nodes.
- A node that is only there to balance the picture.

## Formats that ban shape diagrams

Some formats (Harbor, for one) forbid hand-drawn shape
diagrams: an honest labelled placeholder always beats a fake visual. Write the
`diagram` block anyway. The builder reads the style's `diagrams` token and, in
`placeholder` mode, draws the dashed labelled box the format mandates, naming
the kind and every node label, for a designer to replace. The content you write
is what the placeholder says, so the wording rules above still apply.

## The spec the builder takes

```jsonc
{
  "id": "b_flow", "type": "diagram", "role": "visual",
  "col": 3, "row": 7, "cols": 28, "rows": 7,
  "content": {
    "kind": "flow",              // flow | cycle | timeline | tree | matrix
    "direction": "LR",           // LR or TB, flow and tree only
    "nodes": [
      {"id": "spot", "label": "Spot the moment"},
      {"id": "ask",  "label": "Script the ask"},
      {"id": "hand", "label": "Hand off warm"}
    ],
    "edges": [
      {"from": "spot", "to": "ask",  "label": "prompts"},
      {"from": "ask",  "to": "hand", "label": "owns"}
    ]
  }
}
```

Node ids are yours and must be unique inside the block; every edge names two of
them, or the builder drops the edge and `check_deck` reports it. A `cycle` with
no edges gets its ring arrows drawn for you.

Geometry: give the block the full body width and four to seven rows. The
builder sizes the boxes from the longest label, steps the type down to the
caption floor, and reports a warning instead of shrinking past it: if a label
will not fit, cut the label.

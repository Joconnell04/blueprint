# deck-diagrams

The condensed rules the deck tools carry in their own text, and the skill that
holds the long form.

Skill: `skills/deck-diagrams/SKILL.md` (front-matter name `deck-diagrams`).
Load it before writing any `diagram` block content.

## Tool text (89 words)

> One idea per diagram, 3 to 7 nodes (12 max), never a restatement of the
> headline. Pick the kind by relationship: sequence is `flow`, a loop is
> `cycle`, dates are `timeline`, a hierarchy is `tree`, two axes are `matrix`
> (first two nodes are the axes, the next four the cells). Read `LR` for a
> process, `TB` for a hierarchy. Node labels take four words, parallel grammar;
> edges carry a verb. No decorative shapes. A format that bans shape diagrams
> still takes the spec: the builder draws a labelled placeholder.

## Builder side

- Block: `{type: "diagram", content: {kind, direction?, nodes:[{id,label}], edges:[{from,to,label?}]}}`.
- Layout: `builder/diagram_layout.py`, pure functions, points in and points out.
- Style tokens: `diagrams` (`"shapes"` or `"placeholder"`), the `diagram`
  geometry block, the `diagramNode` and `diagramLabel` type-scale roles, and
  the `card` block for the node fill, border and radius.
- `check_deck` findings: `diagram-edge` (an edge names a node that does not
  exist), `diagram-overlap` (two node boxes cover each other), `diagram-font`
  (labels sit at the readable floor).

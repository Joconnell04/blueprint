/**
 * Remix: re-lay a slide's blocks as another archetype (spec section 2, catalogue in
 * docs/research/deck-planning.md section 5).
 *
 * The catalogue and the solver live in src/shared/archetypes.ts so the server lays a proposal out exactly
 * as Remix does on the canvas. This file is the canvas's door onto it: the names the deck canvas and the
 * layout tests already import, nothing of its own.
 */
export type { Archetype, Slot, SlotRole } from "../shared/archetypes.js";
export {
  ARCHETYPES, ARCHETYPE_NAMES, CYCLE, DECK_BODY_BOTTOM, DECK_BODY_TOP, FURNITURE_TYPES,
  clampToBody, detectArchetype, layoutBlocks, nextArchetype, remix,
} from "../shared/archetypes.js";

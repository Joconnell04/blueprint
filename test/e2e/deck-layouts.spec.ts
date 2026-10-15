import { test, expect, type Browser, type FrameLocator, type Locator } from "@playwright/test";
import fs from "node:fs";
import path from "node:path";
import { Host, dragBy, screenshotOnFailure } from "./host";
import { ARTIFACTS } from "./playwright.config";
import { DECK_DIV_ID, DECK_LONG_ID, DIVIDER_IDS, SECTIONS, SLIDE_IDS, SLIDE_NAMES, deckDividerProject, deckProject, longDeckProject, readProject, seed } from "./seed";

/** The slide grid as the canvas draws it: 32 by 18 cells of 26 px, the slide inset by PAD inside the canvas. */
const CELL = 26, COLS = 32, ROWS = 18;
const at = (n: number) => `${(n - 1) * CELL}px`;
const size = (n: number) => `${n * CELL}px`;
const SHOTS = path.join(ARTIFACTS, "visual");
/** The module edges the house body implies: twelve left edges in cells, from columnModule(undefined). */
const MODULE_STARTS = [3, 5, 8, 10, 13, 15, 18, 20, 22, 25, 27, 30];

/** Press and move a block's own header inside the frame, so the drag is live while the test looks at it. */
async function pressAndMove(header: Locator, dx: number): Promise<void> {
  await header.evaluate((el, d) => {
    const r = el.getBoundingClientRect();
    const x = r.left + r.width / 2, y = r.top + r.height / 2;
    const fire = (type: string, cx: number, cy: number) => el.dispatchEvent(new PointerEvent(type, { pointerId: 1, pointerType: "mouse", isPrimary: true, button: type === "pointermove" ? -1 : 0, buttons: 1, clientX: cx, clientY: cy, bubbles: true, cancelable: true, composed: true }));
    fire("pointerdown", x, y);
    fire("pointermove", x + 2, y);
    fire("pointermove", x + d, y);
  }, dx);
}
async function releaseAt(header: Locator): Promise<void> {
  await header.evaluate((el) => el.dispatchEvent(new PointerEvent("pointerup", { pointerId: 1, pointerType: "mouse", isPrimary: true, button: 0, buttons: 0, bubbles: true, cancelable: true, composed: true })));
}

/**
 * Three of the six seeded slides get layouts, so the other three stay flagged as empty: "Three moves" is
 * the three-cards archetype (one cards block spanning the band), "The market" carries the blocks the drag,
 * resize, nudge and align tests move, "The gap" a table under its open question.
 */
const LAYOUTS = [
  { tab: SLIDE_IDS.move, blocks: [
    { type: "cards", title: "The three moves", row: 7, col: 3, rows: 10, cols: 28, count: 3, content: "owner and first milestone per move", emphasis: true },
  ] },
  { tab: SLIDE_IDS.size, blocks: [
    { type: "chart", title: "Market funnel", row: 7, col: 3, rows: 10, cols: 18, chartKind: "bar" },
    { type: "stat", title: "Reachable today", row: 7, col: 23, rows: 5, cols: 8, count: 1 },
    { type: "text", title: "Why it holds", row: 13, col: 24, rows: 4, cols: 7, content: "census sources per step" },
  ] },
  { tab: SLIDE_IDS.gap, blocks: [
    { type: "table", title: "Onboarding days", row: 7, col: 2, rows: 9, cols: 30, columns: ["Player", "Days", "Trend"] },
  ] },
];

/** The outline is two divider rows and six slides, so a slide's page number is its row, dividers counted. */
const PAGE = { where: 2, size: 3, gap: 4, move: 6, proof: 7, ask: 8 };

test.describe.serial("deck layout view", () => {
  let host: Host;
  let app: FrameLocator;
  screenshotOnFailure(() => host?.page);

  const block = (title: string): Locator => app.locator(".blk").filter({ has: app.locator(".bh .t", { hasText: title }) });
  const slide = (name: string) => app.locator(".tabstrip button", { hasText: name });
  const onDisk = (tabId: string) => readProject(DECK_DIV_ID).tabs.find((t: any) => t.id === tabId);
  const blockOnDisk = (tabId: string, title: string) => onDisk(tabId).blocks.find((b: any) => b.title === title);
  // null while a debounced save is still in flight, so expect.poll can wait for it
  const geom = (tabId: string, title: string) => { const b = blockOnDisk(tabId, title); return b ? [b.row, b.col, b.rows, b.cols] : null; };
  /** The canvas transform: a drag in page pixels is divided by it before it snaps to cells. */
  const zoom = () => app.locator(".canvas.deckcanvas").evaluate((el) => new DOMMatrixReadOnly(getComputedStyle(el).transform).a);

  test.beforeAll(async ({ browser }: { browser: Browser }) => {
    seed(deckDividerProject());
    host = await Host.open(browser);
    await host.callPlain("update_plan", { ops: [{ op: "approve_tabs" }] }, { project: DECK_DIV_ID });
    app = await host.call("propose_layouts", { layouts: LAYOUTS }, { project: DECK_DIV_ID });
  });
  test.afterAll(async () => { await host?.page.close(); });

  test("the slide strip and the slide render at a zoom that fits", async () => {
    // eight rows: six slides and the two section dividers, which the strip marks as its own kind of row
    await expect(app.locator(".tabstrip button")).toHaveCount(8);
    await expect(app.locator(".tabstrip button.dv")).toHaveCount(2);
    await expect(app.locator(".tabstrip button.on")).toHaveText(new RegExp(SLIDE_NAMES[3])); // focusTab: the first layout proposed
    await expect(app.locator(".slide")).toHaveCSS("width", size(COLS));
    await expect(app.locator(".slide")).toHaveCSS("height", size(ROWS));
    // one zoom step is on, the whole slide fits the panel, so the minimap stays away
    await expect(app.locator(".toolbar .zoom button.on")).toHaveCount(1);
    const fits = await app.locator(".canvaswrap").evaluate((w) => { const s = w.querySelector(".canvasscale") as HTMLElement; return s.offsetWidth <= w.clientWidth && s.offsetHeight <= w.clientHeight; });
    expect(fits).toBe(true);
    await expect(app.locator(".minimap")).toBeHidden();
    // furniture: the kicker is the section and the slide, the band is walled off, the page number counts rows
    await expect(app.locator(".dk.kicker")).toHaveText(`${SECTIONS[1]} | ${SLIDE_NAMES[3]}`);
    await expect(app.locator(".dk.headline")).toHaveText("Three moves, one owner each");
    await expect(app.locator(".dk.page")).toHaveText(String(PAGE.move));
    await expect(app.locator(".dk.hairline")).toHaveCount(1);
    await expect(app.locator(".slide .band")).toHaveCSS("height", size(4));
    await expect(block("The three moves")).toHaveCSS("left", at(3));
    await expect(block("The three moves")).toHaveCSS("top", at(7));
    await expect(block("The three moves")).toHaveCSS("width", size(28));
    await expect(block("The three moves")).toHaveCSS("height", size(10));
    await expect(block("The three moves").locator(".badge.emph")).toHaveText("pop");
    await expect(app.locator(".toolbar .archchip")).toContainText("three-cards");
    // Fit, Balance and Snap sit next to Remix; Snap is on until the user says otherwise
    await expect(app.locator(".toolbar button", { hasText: "Fit" })).toBeVisible();
    await expect(app.locator(".toolbar button", { hasText: "Balance" })).toBeVisible();
    await expect(app.locator(".toolbar button.snapbtn")).toHaveAttribute("aria-pressed", "true");
    // the module the house body implies: twelve edges and the body's right edge, away until a block moves
    await expect(app.locator(".slide .guides .gd")).toHaveCount(13);
    await expect(app.locator(".slide .guides")).not.toHaveClass(/(^|\s)on(\s|$)/);
    expect(await app.locator(".slide .guides .gd").evaluateAll((els) => els.map((e) => Math.round(parseFloat((e as HTMLElement).style.left)) / 26 + 1)))
      .toEqual([...MODULE_STARTS, 31]);
    // the story sits under the headline, read only, and the block has all four corner grips
    await expect(app.locator(".side .field", { hasText: "Story" }).locator(".story")).toContainText("Lays out the three moves");
    await expect(block("The three moves").locator(".rs")).toHaveCount(4);
    // three slides have no blocks yet and Claude has one open question: the missing takeaway is not in this count
    await expect(app.locator(".ftr .msg")).toHaveText("0 of 6 approved · 1 open note · 4 flagged");
  });

  test("the palette lists every deck block but the headline", async () => {
    await slide(SLIDE_NAMES[5]).click(); // The ask: an empty slide
    await expect(app.locator(".blk")).toHaveCount(0);
    await app.locator(".toolbar button", { hasText: "Block" }).first().click();
    const pop = app.locator(".pop.deckpal");
    await expect(pop.locator("button")).toHaveCount(13);
    for (const t of ["Kicker", "Text", "Bullets", "Cards", "Stat", "Chart", "Diagram", "Table", "Image", "Quote", "Takeaway", "Placeholder", "Notes"]) {
      await expect(pop.locator(".ty", { hasText: new RegExp(`^${t}$`) })).toHaveCount(1);
    }
    await expect(pop.locator(".ty", { hasText: /^Headline$/ })).toHaveCount(0);
  });

  test("adding a block selects it and the inspector shows Content and Emphasis", async () => {
    await app.locator(".pop.deckpal button").filter({ has: app.locator(".ty", { hasText: /^Stat$/ }) }).click();
    await expect(app.locator(".blk")).toHaveCount(1);
    const added = block("Stat");
    await expect(added).toHaveClass(/sel/);
    await expect(added).toHaveCSS("left", at(3));   // the body band's top left, under the furniture
    await expect(added).toHaveCSS("top", at(5));
    await expect(added).toHaveCSS("width", size(8));
    await expect.poll(() => geom(SLIDE_IDS.ask, "Stat")).toEqual([5, 3, 4, 8]);
    const side = app.locator(".side");
    await expect(side.locator(".sh")).toContainText("Stat");
    const content = side.locator(".field", { hasText: "Content" }).locator("textarea");
    await expect(content).toHaveAttribute("placeholder", "What goes here");
    await content.fill("the ask, in one number");
    await content.blur();
    await expect.poll(() => blockOnDisk(SLIDE_IDS.ask, "Stat")?.content).toBe("the ask, in one number");
    // emphasis is a two-state segment, off until the user says this is the block that should pop
    const emphasis = side.locator(".field", { hasText: "Emphasis" }).locator(".seg");
    await expect(emphasis.locator("button.on")).toHaveText("Off");
    await emphasis.locator("button", { hasText: "Pop" }).click();
    await expect(app.locator(".blk .badge.emph")).toHaveText("pop");
    await expect.poll(() => blockOnDisk(SLIDE_IDS.ask, "Stat")?.emphasis).toBe(true);
    // the slide had no blocks a moment ago, so the flagged count drops by one
    await expect(app.locator(".ftr .msg")).toHaveText("0 of 6 approved · 1 open note · 3 flagged");
  });

  test("a diagram block sketches its nodes and the Edges field describes them", async () => {
    await app.locator(".toolbar button", { hasText: "Block" }).first().click();
    await app.locator(".pop.deckpal button").filter({ has: app.locator(".ty", { hasText: /^Diagram$/ }) }).click();
    const dg = block("Diagram");
    await expect(dg).toHaveCSS("width", size(20));
    await expect(dg).toHaveCSS("height", size(9));
    await expect.poll(() => geom(SLIDE_IDS.ask, "Diagram")).toEqual([9, 3, 9, 20]);
    // nothing described yet: the generic four-node flow, no labels claimed
    await expect(dg.locator(".dg .nd")).toHaveCount(4);
    await expect(dg.locator(".dg text.lb")).toHaveCount(0);
    const side = app.locator(".side");
    await expect(side.locator(".field", { hasText: "Kind" }).locator("select")).toHaveValue("flow");
    await expect(side.locator(".field", { hasText: "Direction" }).locator("select")).toHaveValue("LR");
    const edges = side.locator("textarea.edges");
    await expect(edges).toHaveAttribute("placeholder", "A -> B: label");
    await edges.fill("Signup -> Invoice: two days\nChurn");
    await edges.blur();
    await expect(dg.locator(".dg .nd")).toHaveCount(3);
    await expect(dg.locator(".dg text.lb")).toHaveText(["Signup", "Invoice", "Churn"]);
    await expect(dg.locator(".dg text.el")).toHaveText(["two days"]);
    await expect.poll(() => blockOnDisk(SLIDE_IDS.ask, "Diagram")?.diagram).toEqual({
      kind: "flow", direction: "LR",
      nodes: [{ id: "signup", label: "Signup" }, { id: "invoice", label: "Invoice" }, { id: "churn", label: "Churn" }],
      edges: [{ from: "signup", to: "invoice", label: "two days" }],
    });
    // a ring has no direction to pick, and the field shows the same lines back
    await side.locator(".field", { hasText: "Kind" }).locator("select").selectOption("cycle");
    await expect(side.locator(".field", { hasText: "Direction" })).toHaveCount(0);
    await expect(side.locator("textarea.edges")).toHaveValue("Signup -> Invoice: two days\nChurn");
    await expect(block("Diagram").locator(".dg .nd")).toHaveCount(3);
    await expect.poll(() => blockOnDisk(SLIDE_IDS.ask, "Diagram")?.diagram.kind).toBe("cycle");
  });

  test("the Blocks list deletes a block in two clicks, and the Delete key removes the selection", async () => {
    await app.locator(".canvaswrap").press("Escape");
    const rows = app.locator(".side .blocklist .b");
    await expect(rows).toHaveCount(2);
    // a click on the row selects the block, so the list and the canvas agree
    await rows.filter({ hasText: "Stat" }).click();
    await expect(block("Stat")).toHaveClass(/primary/);
    await app.locator(".canvaswrap").press("Escape");
    const statRow = rows.filter({ hasText: "Stat" });
    await statRow.hover();
    const trash = statRow.locator(".acts button");
    await trash.click();
    await expect(trash).toHaveText("Delete?");
    await trash.click();
    await expect(app.locator(".blk")).toHaveCount(1);
    await expect.poll(() => onDisk(SLIDE_IDS.ask).blocks.map((b: any) => b.title)).toEqual(["Diagram"]);
    // and the key does the same for whatever is selected on the canvas
    await block("Diagram").click();
    await app.locator(".canvaswrap").press("Backspace");
    await expect(app.locator(".blk")).toHaveCount(0);
    await expect.poll(() => onDisk(SLIDE_IDS.ask).blocks).toEqual([]);
    await expect(app.locator(".ftr .msg")).toHaveText("0 of 6 approved · 1 open note · 4 flagged");
  });

  test("dragging a block header snaps it to the 30 pt grid", async () => {
    await slide(SLIDE_NAMES[1]).click();
    await expect(app.locator(".blk")).toHaveCount(3);
    // the column module has its own test below; with Snap off a drop lands on the cell it was dropped on
    const snap = app.locator(".toolbar button.snapbtn");
    await snap.click();
    await expect(snap).toHaveAttribute("aria-pressed", "false");
    await expect(app.locator(".toast")).toHaveText("Snap off");
    const chart = block("Market funnel");
    const z = await zoom();
    // one cell left, and 1.4 cells down: the drop rounds to the nearest cell rather than landing between them
    await dragBy(host, chart.locator(".bh"), -1 * CELL * z, 1.4 * CELL * z);
    await expect(chart).toHaveCSS("left", at(2));
    await expect(chart).toHaveCSS("top", at(8));
    await expect.poll(() => geom(SLIDE_IDS.size, "Market funnel")).toEqual([8, 2, 10, 18]);
    await expect(app.locator(".side .sh")).toContainText("8:2");
  });

  test("a long story clamps to four lines and expands", async () => {
    await app.locator(".canvaswrap").press("Escape"); // the story sits under the headline, with no block selected
    const story = app.locator(".side .field", { hasText: "Story" });
    await expect(story.locator(".story")).toContainText("Sizes the opportunity top down");
    const more = story.locator("button.more");
    await expect(more).toHaveText("more");
    const clamped = await story.locator(".story").evaluate((el) => el.scrollHeight > el.clientHeight + 1);
    expect(clamped).toBe(true);
    await more.click();
    await expect(more).toHaveText("less");
    await expect.poll(() => story.locator(".story").evaluate((el) => el.scrollHeight <= el.clientHeight + 1)).toBe(true);
  });

  test("the furniture band refuses a block dragged into it", async () => {
    const chart = block("Market funnel");
    const z = await zoom();
    // six cells up from row 8 is row 2; rows 1 to 4 are the furniture's, so the block stops at row 5
    await dragBy(host, chart.locator(".bh"), 0, -6 * CELL * z);
    await expect(chart).toHaveCSS("top", at(5));
    await expect.poll(() => geom(SLIDE_IDS.size, "Market funnel")).toEqual([5, 2, 10, 18]);
  });

  test("resizing from the corner grows the block in cells", async () => {
    const stat = block("Reachable today");
    await stat.click();
    const z = await zoom();
    await dragBy(host, stat.locator(".rs.se"), 2 * CELL * z, 1 * CELL * z);
    await expect(stat).toHaveCSS("width", size(10));
    await expect(stat).toHaveCSS("height", size(6));
    await expect.poll(() => geom(SLIDE_IDS.size, "Reachable today")).toEqual([7, 23, 6, 10]);
  });

  test("resizing from the top left corner moves the block's cell, and one gesture is one undo", async () => {
    const stat = block("Reachable today");
    const z = await zoom();
    // dragging the top left grip in by (1, 2) keeps the bottom right corner where it was
    await dragBy(host, stat.locator(".rs.nw"), 1 * CELL * z, 2 * CELL * z);
    await expect(stat).toHaveCSS("left", at(24));
    await expect(stat).toHaveCSS("top", at(9));
    await expect(stat).toHaveCSS("width", size(9));
    await expect(stat).toHaveCSS("height", size(4));
    await expect.poll(() => geom(SLIDE_IDS.size, "Reachable today")).toEqual([9, 24, 4, 9]);
    await app.locator(".toolbar button[title^='Undo']").click();
    await expect(block("Reachable today")).toHaveCSS("top", at(7));
    await expect.poll(() => geom(SLIDE_IDS.size, "Reachable today")).toEqual([7, 23, 6, 10]);
  });

  test("arrows nudge the selected block one cell", async () => {
    await block("Why it holds").click();
    await expect(app.locator(".blk.primary .bh .t")).toHaveText("Why it holds");
    await app.locator(".canvaswrap").press("ArrowRight");
    await app.locator(".canvaswrap").press("ArrowDown");
    await expect(block("Why it holds")).toHaveCSS("left", at(25));
    await expect(block("Why it holds")).toHaveCSS("top", at(14));
    await expect.poll(() => geom(SLIDE_IDS.size, "Why it holds")).toEqual([14, 25, 4, 7]);
  });

  test("shift-click selects two blocks and Align left shares their leftmost column", async () => {
    await host.settle();
    await block("Reachable today").click();
    await block("Why it holds").click({ modifiers: ["Shift"] });
    await expect(app.locator(".blk.sel")).toHaveCount(2);
    await expect(app.locator(".selbar .count")).toHaveText("2 selected");
    await app.locator(".selbar button", { hasText: "Align left" }).click();
    await expect(block("Reachable today")).toHaveCSS("left", at(23));
    await expect(block("Why it holds")).toHaveCSS("left", at(23));
    await expect.poll(() => [blockOnDisk(SLIDE_IDS.size, "Reachable today")?.col, blockOnDisk(SLIDE_IDS.size, "Why it holds")?.col]).toEqual([23, 23]);
    await app.locator(".selbar button", { hasText: "Deselect" }).click();
    await expect(app.locator(".selbar")).toBeHidden();
  });

  /**
   * Ranked change 1: a drop lands on the module the style's body implies, and the module's edges show
   * while the block is moving. The chart is at column 2 by now, one cell off the edge at column 3.
   */
  test("Snap lands a drop on the column module", async () => {
    const snap = app.locator(".toolbar button.snapbtn");
    await snap.click();
    await expect(app.locator(".toast")).toHaveText("Snap on");
    await expect(snap).toHaveAttribute("aria-pressed", "true");
    const chart = block("Market funnel");
    const z = await zoom();
    expect(geom(SLIDE_IDS.size, "Market funnel")).toEqual([5, 2, 10, 18]);
    // the guides come up while the block moves, and go away when it lands
    const header = chart.locator(".bh");
    await pressAndMove(header, 2 * CELL * z);
    await expect(app.locator(".slide .guides")).toHaveClass(/(^|\s)on(\s|$)/);
    await releaseAt(header);
    await expect(app.locator(".slide .guides")).not.toHaveClass(/(^|\s)on(\s|$)/);
    // dropped on column 4, one cell past the module edge: it lands on column 3 and keeps its width
    await expect(chart).toHaveCSS("left", at(3));
    await expect(chart).toHaveCSS("width", size(18));
    await expect.poll(() => geom(SLIDE_IDS.size, "Market funnel")).toEqual([5, 3, 10, 18]);
    // three cells on is column 6, one off the next edge, so that is where it lands
    await dragBy(host, chart.locator(".bh"), 3 * CELL * z, 0);
    await expect(chart).toHaveCSS("left", at(5));
    await expect.poll(() => geom(SLIDE_IDS.size, "Market funnel")).toEqual([5, 5, 10, 18]);
    // and a resize takes the module on the edge the grip moved, leaving the other one alone
    const stat = block("Reachable today");
    await stat.click();
    expect(geom(SLIDE_IDS.size, "Reachable today")).toEqual([7, 23, 6, 10]);
    await dragBy(host, stat.locator(".rs.se"), -1 * CELL * z, 0);
    await expect.poll(() => geom(SLIDE_IDS.size, "Reachable today")).toEqual([7, 23, 6, 8]);
    await expect(stat).toHaveCSS("left", at(23));
    await expect(stat).toHaveCSS("width", size(8));
  });

  test("Remix is ⌥R: plain R does nothing, and undo puts the slide back", async () => {
    await slide(SLIDE_NAMES[3]).click();
    expect(geom(SLIDE_IDS.move, "The three moves")).toEqual([7, 3, 10, 28]);
    await expect(app.locator(".toolbar button", { hasText: "Remix" }).locator(".key")).toHaveText("⌥R");
    // plain R is not a shortcut any more (F13): the slide is left exactly as it was
    await app.locator(".canvaswrap").press("r");
    await expect(app.locator(".toolbar .archchip")).toContainText("three-cards");
    await expect(block("The three moves")).toHaveCSS("height", size(10));
    expect(geom(SLIDE_IDS.move, "The three moves")).toEqual([7, 3, 10, 28]);
    await app.locator(".canvaswrap").press("Alt+r");
    await expect(app.locator(".toast")).toHaveText("Remixed: two-by-two");
    await expect(app.locator(".toolbar .archchip")).toContainText("two-by-two");
    // one cards block fills the whole card band of the new archetype, and nothing else is invented
    await expect(block("The three moves")).toHaveCSS("height", size(9));
    await expect(app.locator(".blk")).toHaveCount(1);
    await expect.poll(() => geom(SLIDE_IDS.move, "The three moves")).toEqual([7, 3, 9, 28]);
    await app.locator(".toolbar button[title^='Undo']").click();
    await expect(block("The three moves")).toHaveCSS("height", size(10));
    await expect.poll(() => geom(SLIDE_IDS.move, "The three moves")).toEqual([7, 3, 10, 28]);
    await expect(app.locator(".toolbar .archchip")).toContainText("three-cards");
  });

  test("the archetype popover lays the slide out by name and fills the empty slot", async () => {
    await app.locator(".toolbar .archchip").click();
    const pop = app.locator(".pop.archlist");
    await expect(pop.locator("button")).toHaveCount(18);
    await pop.locator("button", { hasText: "text-visual" }).click();
    await expect(app.locator(".toolbar .archchip")).toContainText("text-visual");
    await expect(app.locator(".blk")).toHaveCount(2);
    await expect(block("The three moves")).toHaveCSS("width", size(12));
    await expect(block("Visual")).toHaveCSS("left", at(17));
    await expect.poll(() => onDisk(SLIDE_IDS.move).blocks.map((b: any) => [b.type, b.col, b.cols])).toEqual([["cards", 3, 12], ["placeholder", 17, 14]]);
  });

  test("⌘left and ⌘right walk the strip, dividers included", async () => {
    await app.locator(".canvaswrap").press("Meta+ArrowLeft");
    await expect(app.locator(".tabstrip button.on")).toHaveText(new RegExp(SECTIONS[1]));
    await expect(app.locator(".slide.divider")).toBeVisible();
    await app.locator(".canvaswrap").press("Meta+ArrowRight");
    await expect(app.locator(".tabstrip button.on")).toHaveText(new RegExp(SLIDE_NAMES[3]));
    await expect(app.locator(".slide.divider")).toHaveCount(0);
  });

  test("the headline edits inline, persists, and keeps the outline approved", async () => {
    const text = "Three moves, one owner, one milestone each";
    await app.locator(".dk.headline").click();
    const editor = app.locator("textarea.hledit");
    await expect(editor).toBeVisible();
    await editor.fill(text);
    await editor.press("Enter");
    await expect(app.locator(".dk.headline")).toHaveText(text);
    await expect.poll(() => onDisk(SLIDE_IDS.move).message).toBe(text);
    // the one-field save leaves the approvals and the other rows alone
    expect(readProject(DECK_DIV_ID).tabsApproved).toBe(true);
    expect(onDisk(SLIDE_IDS.gap).message).toBe("We lose on week one, not on price");
    expect(onDisk(SLIDE_IDS.move).story).toContain("Lays out the three moves");
    // and the inspector shows the same line once nothing is selected
    await app.locator(".canvaswrap").press("Escape");
    await expect(app.locator(".side .field", { hasText: "Headline" }).locator("textarea")).toHaveValue(text);
  });

  test("a block's thread lists the slide's own notes, marked", async () => {
    await slide(SLIDE_NAMES[2]).click(); // The gap: Claude's question sits on the slide, not on a block
    await block("Onboarding days").click();
    const thread = app.locator(".side .field", { hasText: "Thread" });
    const onSlide = thread.locator(".notes.onslide .note");
    await expect(onSlide).toHaveCount(1);
    await expect(onSlide.locator(".st")).toHaveText("Slide");
    await expect(onSlide.locator(".nt")).toHaveText("Which two competitors should this name?");
    // clicking it drops the block selection, which is where that thread can be answered
    await onSlide.click();
    await expect(app.locator(".side .sh")).toContainText("Slide");
    await expect(app.locator(".blk.sel")).toHaveCount(0);
  });

  test("a divider row draws the numbered section slide and renames in place", async () => {
    await app.locator(".tabstrip button.dv").nth(1).click();
    await expect(app.locator(".slide.divider")).toBeVisible();
    await expect(app.locator(".slide.divider .dtitle .no")).toHaveText("02");
    await expect(app.locator(".slide.divider .dtitle .t")).toHaveText(SECTIONS[1]);
    // no palette, no Remix, no blocks, and nothing to approve on a section slide
    await expect(app.locator(".toolbar button", { hasText: "Block" })).toHaveCount(0);
    await expect(app.locator(".toolbar button", { hasText: "Remix" })).toHaveCount(0);
    await expect(app.locator(".toolbar .archchip")).toHaveCount(0);
    await expect(app.locator(".blk")).toHaveCount(0);
    await expect(app.locator(".ftr button", { hasText: "Approve \"" })).toHaveCount(0);
    // the inspector is the title and the thread, nothing else
    await expect(app.locator(".side .sh")).toContainText("Divider");
    await expect(app.locator(".side .field")).toHaveCount(2);
    await expect(app.locator(".side .field .label").nth(0)).toHaveText("Title");
    await expect(app.locator(".side .field .label").nth(1)).toHaveText("Thread");
    // the title is the section name, edited on the slide itself
    await app.locator(".slide.divider .dtitle").click();
    const editor = app.locator("input.dtedit");
    await expect(editor).toHaveValue(SECTIONS[1]);
    await editor.fill("The plan we back");
    await editor.press("Enter");
    await expect(app.locator(".slide.divider .dtitle .t")).toHaveText("The plan we back");
    await expect.poll(() => onDisk(DIVIDER_IDS.plan).name).toBe("The plan we back");
    // the first divider keeps its own number
    await app.locator(".tabstrip button.dv").nth(0).click();
    await expect(app.locator(".slide.divider .dtitle .no")).toHaveText("01");
  });

  test("approves one slide, then every layout after the warning", async () => {
    await slide(SLIDE_NAMES[3]).click();
    await app.locator(".ftr button", { hasText: `Approve "${SLIDE_NAMES[3]}"` }).click();
    await expect(app.locator(".ftr button", { hasText: "Unapprove" })).toBeVisible();
    // the archetype filled an empty slot with a placeholder, which the server flags: never ship one
    await expect(app.locator(".toolbar .chip.warn")).toContainText("1 flagged");
    await expect.poll(() => onDisk(SLIDE_IDS.move).layoutApproved).toBe(true);
    await expect(app.locator(".ftr .msg")).toContainText("1 of 6 approved");
    // three slides have no blocks, one placeholder is flagged and Claude's question is open: Approve all asks
    await app.locator(".ftr button.primary", { hasText: "Approve all layouts" }).click();
    const confirm = app.locator(".pop.confirm");
    await expect(confirm.locator(".t")).toHaveText("5 still flagged");
    await expect(confirm.locator(".warnlist li")).toHaveCount(5);
    await confirm.locator("button", { hasText: "Cancel" }).click();
    expect(readProject(DECK_DIV_ID).tabs.filter((t: any) => t.layoutApproved).length).toBe(1);
    await app.locator(".ftr button.primary", { hasText: "Approve all layouts" }).click();
    await app.locator(".pop.confirm button", { hasText: "Approve anyway" }).click();
    await expect(app.locator(".ftr .msg")).toHaveText("Layouts approved");
    await expect.poll(() => readProject(DECK_DIV_ID).tabs.every((t: any) => t.layoutApproved || t.divider)).toBe(true);
  });
});

/**
 * F21: the strip is a scroller, and every save re-renders it. It keeps its own scrollLeft (the shell
 * restores anything tagged `data-scroll` after a draw) and only pulls a chip into view when it is hidden.
 */
test("the slide strip keeps its scroll when a chip is clicked", async ({ browser }: { browser: Browser }) => {
  seed(longDeckProject());
  const host = await Host.open(browser);
  await host.callPlain("update_plan", { ops: [{ op: "approve_tabs" }] }, { project: DECK_LONG_ID });
  const app = await host.call("propose_layouts", { layouts: [{ tab: SLIDE_IDS.move, blocks: [{ type: "cards", title: "The three moves", row: 7, col: 3, rows: 10, cols: 28, count: 3 }] }] }, { project: DECK_LONG_ID });
  const strip = app.locator(".tabstrip");
  await expect(strip).toHaveAttribute("data-scroll", "strip");
  await host.settle();
  // scroll the strip along, then click a chip that is already in view
  const room = await strip.evaluate((el) => el.scrollWidth - el.clientWidth);
  expect(room).toBeGreaterThan(100);
  await strip.evaluate((el) => { el.scrollLeft = 120; });
  const inView = await strip.evaluate((el) => {
    const chips = [...el.querySelectorAll("button")] as HTMLElement[];
    const hit = chips.find((c) => c.offsetLeft >= el.scrollLeft && c.offsetLeft + c.offsetWidth <= el.scrollLeft + el.clientWidth && !c.classList.contains("on"));
    return hit ? chips.indexOf(hit) : -1;
  });
  expect(inView).toBeGreaterThanOrEqual(0);
  await app.locator(".tabstrip button").nth(inView).click();
  await expect(app.locator(".tabstrip button.on")).toHaveCount(1);
  await expect.poll(() => strip.evaluate((el) => el.scrollLeft)).toBe(120);
  await host.page.close();
});

/**
 * Fit and Balance: the balanced passes of the fidelity setting, run on the current slide so the user can
 * see what the setting would do before building (docs/research/output-quality.md, "The fidelity setting").
 */
test("Fit sizes blocks to their content, Balance spreads them down the band", async ({ browser }: { browser: Browser }) => {
  const id = "deckfit";
  seed({ ...deckProject(), id });
  const host = await Host.open(browser);
  await host.callPlain("update_plan", { ops: [{ op: "approve_tabs" }] }, { project: id });
  const app = await host.call("propose_layouts", { layouts: [{ tab: SLIDE_IDS.size, blocks: [
    { type: "bullets", title: "Why it holds", row: 5, col: 3, rows: 6, cols: 12, content: "x".repeat(200) },
    { type: "table", title: "Onboarding days", row: 5, col: 17, rows: 9, cols: 14, columns: ["Player", "Days", "Trend"] },
    { type: "stat", title: "Reachable today", row: 12, col: 3, rows: 5, cols: 12 },
  ] }] }, { project: id });
  const blk = (title: string) => app.locator(".blk").filter({ has: app.locator(".bh .t", { hasText: title }) });
  const geoms = () => readProject(id).tabs.find((t: any) => t.id === SLIDE_IDS.size).blocks
    .map((b: any) => [b.title, b.row, b.col, b.rows, b.cols]);
  expect(geoms()).toEqual([["Why it holds", 5, 3, 6, 12], ["Onboarding days", 5, 17, 9, 14], ["Reachable today", 12, 3, 5, 12]]);

  // the selection first: four wrapped lines of body and a label are three rows, and only that block moves
  await blk("Why it holds").click();
  await app.locator(".toolbar button", { hasText: "Fit" }).click();
  await expect(app.locator(".toast")).toHaveText("Fitted 1 block");
  await expect.poll(geoms).toEqual([["Why it holds", 5, 3, 3, 12], ["Onboarding days", 5, 17, 9, 14], ["Reachable today", 12, 3, 5, 12]]);

  // nothing selected: the whole slide. A table is its header plus its rows, a stat is four rows.
  await app.locator(".canvaswrap").press("Escape");
  await app.locator(".toolbar button", { hasText: "Fit" }).click();
  await expect(app.locator(".toast")).toHaveText("Fitted 2 blocks");
  await expect.poll(geoms).toEqual([["Why it holds", 5, 3, 3, 12], ["Onboarding days", 5, 17, 4, 14], ["Reachable today", 12, 3, 4, 12]]);
  await app.locator(".toolbar button", { hasText: "Fit" }).click();
  await expect(app.locator(".toast")).toHaveText("Already fits");

  // Balance: the pair that share rows 5 to 9 move as one band, and the two bands sit on equal gaps
  await app.locator(".toolbar button", { hasText: "Balance" }).click();
  await expect(app.locator(".toast")).toHaveText("Balanced 3 blocks");
  await expect.poll(geoms).toEqual([["Why it holds", 6, 3, 3, 12], ["Onboarding days", 6, 17, 4, 14], ["Reachable today", 13, 3, 4, 12]]);
  await expect(blk("Why it holds")).toHaveCSS("top", at(6));
  await expect(blk("Reachable today")).toHaveCSS("top", at(13));
  await app.locator(".toolbar button", { hasText: "Balance" }).click();
  await expect(app.locator(".toast")).toHaveText("Already balanced");

  // one gesture, one undo step: the fitted slide comes back
  await app.locator(".toolbar button[title^='Undo']").click();
  await expect.poll(geoms).toEqual([["Why it holds", 5, 3, 3, 12], ["Onboarding days", 5, 17, 4, 14], ["Reachable today", 12, 3, 4, 12]]);
  await host.page.close();
});

/** The shot the design review reads: a modular slide with the module's own guides up, dark and light. */
test("shot: the deck layout with the column guides", async ({ browser }: { browser: Browser }) => {
  const id = "deckshot";
  seed({ ...deckProject(), id });
  fs.mkdirSync(SHOTS, { recursive: true });
  const host = await Host.open(browser);
  await host.callPlain("update_plan", { ops: [{ op: "approve_tabs" }] }, { project: id });
  const app = await host.call("propose_layouts", { layouts: [{ tab: SLIDE_IDS.size, blocks: [
    { type: "chart", title: "Market funnel", row: 6, col: 3, rows: 10, cols: 16, chartKind: "bar", content: "three steps from all businesses to reachable" },
    { type: "stat", title: "Reachable today", row: 6, col: 20, rows: 5, cols: 11, count: 2, emphasis: true },
    { type: "bullets", title: "Why it holds", row: 12, col: 20, rows: 4, cols: 11, content: "census sources per step" },
  ] }] }, { project: id });
  const idx = (await host.iframes.count()) - 1;
  await app.locator(".blk").filter({ has: app.locator(".bh .t", { hasText: "Market funnel" }) }).click();
  let dark = false;
  for (const theme of ["dark", "light"] as const) {
    if (dark !== (theme === "dark")) { await host.themeToggle.click(); dark = theme === "dark"; }
    await host.settle();
    await expect.poll(() => app.locator("html").getAttribute("data-theme")).toBe(theme);
    // the guides show while a block is moving; hold them up for the shot rather than freezing a drag
    await app.locator(".slide .guides").evaluate((el) => el.classList.add("on"));
    await host.page.waitForTimeout(300);
    await host.iframes.nth(idx).screenshot({ path: path.join(SHOTS, `c1-deck-guides-${theme}.png`) });
  }
  await host.page.close();
});

/**
 * The brief's missing takeaway is a readiness warning with no slide behind it, so it belongs to the brief
 * card in Layer 1 and not to the layout counters (spec section 5, layout-view.ts `flagged`).
 */
test("the flagged count leaves the missing takeaway to the brief card", async ({ browser }: { browser: Browser }) => {
  const base = deckProject();
  const id = "decknotake";
  seed({ ...base, id, briefCard: { audience: base.briefCard.audience, goal: base.briefCard.goal }, tabs: [base.tabs[3]], notes: [] });
  const host = await Host.open(browser);
  await host.callPlain("update_plan", { ops: [{ op: "approve_tabs" }] }, { project: id });
  const app = await host.call("propose_layouts", { layouts: [{ tab: SLIDE_IDS.move, blocks: [{ type: "cards", title: "The three moves", row: 7, col: 3, rows: 10, cols: 28, count: 3 }] }] }, { project: id });
  // nothing on the slide is flagged, and the footer counts nothing either
  await expect(app.locator(".toolbar .chip.ok")).toHaveText("ready");
  await expect(app.locator(".side .field", { hasText: "Readiness" }).locator(".ok-text")).toHaveText("Nothing flagged");
  await expect(app.locator(".ftr .msg")).toHaveText("0 of 1 approved");
  // the warning itself is alive: the brief card in Layer 1 still shows its dot
  await app.locator(".hdr .seg button").first().click();
  await expect(app.locator(".brief .dot.warn")).toHaveCount(1);
  await host.page.close();
});

/**
 * F-2: a drop that overlaps is resolved by reading order, which is the rule the server re-runs on every
 * save (`resolveBandOverlaps`). The canvas used to keep whatever the hand had dragged and push the other
 * block down, so a drop previewed one slide and stored another: here the dragged table lands lower than
 * the chart, so the table is the one that moves on, and the plan on disk says the same.
 */
test("a drop that overlaps moves the later block, canvas and plan alike", async ({ browser }: { browser: Browser }) => {
  const id = "deckdrop";
  seed({ ...deckProject(), id });
  const host = await Host.open(browser);
  await host.callPlain("update_plan", { ops: [{ op: "approve_tabs" }] }, { project: id });
  const app = await host.call("propose_layouts", { layouts: [{ tab: SLIDE_IDS.size, blocks: [
    { type: "chart", title: "Market funnel", row: 5, col: 3, rows: 5, cols: 14, chartKind: "bar" },
    { type: "table", title: "Onboarding days", row: 12, col: 3, rows: 4, cols: 14, columns: ["Player", "Days", "Trend"] },
  ] }] }, { project: id });
  const blk = (title: string) => app.locator(".blk").filter({ has: app.locator(".bh .t", { hasText: title }) });
  const geoms = () => readProject(id).tabs.find((t: any) => t.id === SLIDE_IDS.size).blocks
    .map((b: any) => [b.title, b.row, b.col, b.rows, b.cols]);
  const zoom = () => app.locator(".canvas.deckcanvas").evaluate((el) => new DOMMatrixReadOnly(getComputedStyle(el).transform).a);
  expect(geoms()).toEqual([["Market funnel", 5, 3, 5, 14], ["Onboarding days", 12, 3, 4, 14]]);

  // Snap off, so the drop lands on the cell it was dropped on and only the overlap rule moves anything
  await app.locator(".toolbar button.snapbtn").click();
  await expect(app.locator(".toolbar button.snapbtn")).toHaveAttribute("aria-pressed", "false");
  const z = await zoom();
  // five cells up puts the table at row 7, inside the chart's rows 5 to 9
  // (no toast assertion: dragBy can spend longer replaying the gesture than a toast stays up)
  await dragBy(host, blk("Onboarding days").locator(".bh"), 0, -5 * CELL * z);
  await expect(blk("Onboarding days")).toHaveCSS("top", at(10));
  await expect(blk("Market funnel")).toHaveCSS("top", at(5));
  await expect.poll(geoms).toEqual([["Market funnel", 5, 3, 5, 14], ["Onboarding days", 10, 3, 4, 14]]);
  await host.page.close();
});

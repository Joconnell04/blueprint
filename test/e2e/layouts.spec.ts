import { test, expect, type Browser, type FrameLocator, type Locator } from "@playwright/test";
import { Host, dragBy, screenshotOnFailure } from "./host";
import { TAB_IDS, demoProject, readProject, seed } from "./seed";

const CELL_W = 64, CELL_H = 22, LEFT = 30, TOP = 20;
const left = (col: number) => `${LEFT + (col - 1) * CELL_W}px`;
const top = (row: number) => `${TOP + (row - 1) * CELL_H}px`;

/**
 * Layouts for the Assumptions tab: an inputs block, a table, a calc with an open question, a chart with
 * no source (a warning). Positions leave room so the drags and resizes below never overlap another block
 * (the canvas may push overlapping blocks down on drop).
 */
const LAYOUTS = [
  { tab: TAB_IDS.assumptions, blocks: [
    { type: "inputs", title: "Key Assumptions", row: 2, col: 2, rows: 4, cols: 3, rowLabels: ["Price per unit", "Growth rate", "Churn"] },
    { type: "table", title: "Segments", row: 7, col: 2, rows: 5, cols: 4, columns: ["Segment", "Firms", "Spend"], sample: [["SMB", 1200, 40], ["Mid", 300, 90]] },
    { type: "calc", title: "Funnel", row: 7, col: 10, rows: 5, cols: 4, columns: ["Step", "Year 1", "Year 2"], questions: ["Annual or monthly?"] },
    { type: "chart", title: "Ramp Chart", row: 20, col: 2, rows: 6, cols: 5 },
  ] },
  { tab: TAB_IDS.funnel, blocks: [{ type: "output", title: "SOM", row: 2, col: 2, rows: 3, cols: 3, rowLabels: ["SOM year 3", "Share"] }] },
];

test.describe.serial("layout view", () => {
  let host: Host;
  let app: FrameLocator;
  screenshotOnFailure(() => host?.page);

  const block = (title: string): Locator => app.locator(".blk").filter({ has: app.locator(".bh .t", { hasText: title }) });
  const assumptions = () => readProject().tabs.find((t: any) => t.id === TAB_IDS.assumptions);
  const blockOnDisk = (title: string) => assumptions().blocks.find((b: any) => b.title === title);

  test.beforeAll(async ({ browser }: { browser: Browser }) => {
    seed(demoProject({ tabsApproved: true }));
    host = await Host.open(browser);
    app = await host.call("propose_layouts", { layouts: LAYOUTS });
  });
  test.afterAll(async () => { await host?.page.close(); });

  test("blocks render at their cells", async () => {
    await expect(app.locator(".tabstrip button.on")).toHaveText(/Assumptions/);
    await expect(app.locator(".blk")).toHaveCount(4);
    // left = 30 + (col-1)*64, top = 20 + (row-1)*22, size = cells * (64, 22) minus the 1px border
    await expect(block("Key Assumptions")).toHaveCSS("left", left(2));
    await expect(block("Key Assumptions")).toHaveCSS("top", top(2));
    await expect(block("Key Assumptions")).toHaveCSS("width", `${3 * CELL_W - 1}px`);
    await expect(block("Key Assumptions")).toHaveCSS("height", `${4 * CELL_H - 1}px`);
    await expect(block("Segments")).toHaveCSS("left", left(2));
    await expect(block("Segments")).toHaveCSS("top", top(7));
    await expect(block("Funnel")).toHaveCSS("left", left(10));
    await expect(block("Funnel")).toHaveCSS("top", top(7));
    await expect(block("Ramp Chart")).toHaveCSS("top", top(20));
    // header rows and grid geometry
    await expect(app.locator(".colhdr > div").nth(1)).toHaveText("B");
    await expect(app.locator(".colhdr > div").nth(1)).toHaveCSS("width", `${CELL_W}px`);
    await expect(app.locator(".rowhdr > div").nth(6)).toHaveText("7");
    // the question shows as a badge and the chart with no source is flagged
    await expect(block("Funnel").locator(".badge.q")).toHaveText("1?");
    await expect(app.locator(".toolbar .chip.warn")).toHaveText("2 flagged"); // the question and the chart with no source
  });

  test("a guessed value reads as a placeholder, a real sample takes its column's format", async () => {
    // Segments carries two real sample rows: formatted the way the builder will write them, and not marked
    const seg = block("Segments");
    const cells = (title: string, r: number) => block(title).locator(".bc .r").nth(r).locator("> div");
    await expect(cells("Segments", 1).nth(0)).toHaveText("SMB");
    await expect(cells("Segments", 1).nth(1)).toHaveText("1,200");
    await expect(cells("Segments", 1).nth(2)).toHaveText("$40.00");
    await expect(seg.locator(".bc .ph")).toHaveCount(0);
    // Funnel has no samples, so every value is the generator's guess: muted, dotted, never mistaken for data
    const guess = cells("Funnel", 1).nth(1);
    await expect(guess).toHaveClass(/ph/);
    await expect(guess).toHaveCSS("text-decoration-style", "dotted");
    const colour = (l: Locator) => l.evaluate((el) => getComputedStyle(el).color);
    expect(await colour(guess)).not.toBe(await colour(cells("Segments", 1).nth(1)));
    // and the same for a line item with no value behind it
    await expect(block("Key Assumptions").locator(".bc .v.ph")).toHaveCount(3);
  });

  test("a block too narrow for its headers steps them one size down", async () => {
    const body = block("Funnel").locator(".bc");
    const header = block("Funnel").locator(".bc .r.h > div").first();
    await expect(body).not.toHaveClass(/tight/);
    await expect(header).toHaveCSS("font-size", "10px");
    // shift + arrows resize: two columns narrower and "Year 1" no longer fits at the cell size
    await block("Funnel").click();
    await app.locator(".canvaswrap").press("Shift+ArrowLeft");
    await app.locator(".canvaswrap").press("Shift+ArrowLeft");
    await expect(block("Funnel")).toHaveCSS("width", `${2 * CELL_W - 1}px`);
    await expect(body).toHaveClass(/tight/);
    await expect(header).toHaveCSS("font-size", "9px");
    await expect.poll(() => blockOnDisk("Funnel").cols).toBe(2);
    // back to its own width, back to the cell size
    await app.locator(".canvaswrap").press("Shift+ArrowRight");
    await app.locator(".canvaswrap").press("Shift+ArrowRight");
    await expect(body).not.toHaveClass(/tight/);
    await expect(header).toHaveCSS("font-size", "10px");
    await expect.poll(() => blockOnDisk("Funnel").cols).toBe(4);
  });

  test("+ Block popover adds a table", async () => {
    await app.locator(".toolbar button", { hasText: "Block" }).first().click();
    const pop = app.locator(".pop");
    await expect(pop).toBeVisible();
    await expect(pop.locator("button")).toHaveCount(6);
    await pop.locator("button").filter({ has: app.locator(".ty", { hasText: /^Table$/ }) }).click();
    await expect(app.locator(".blk")).toHaveCount(5);
    const added = block("Data Table");
    await expect(added).toHaveClass(/sel/);
    // placed one blank row below the lowest block (chart ends at row 25) at column B
    await expect(added).toHaveCSS("top", top(27));
    await expect(added).toHaveCSS("left", left(2));
    // four columns, four cells: the canvas draws the footprint the sheet will have, not the type default of five
    await expect(added).toHaveCSS("width", `${4 * CELL_W - 1}px`);
    await expect(added.locator(".bc .r.h > div")).toHaveCount(4);
    await expect(app.locator(".side input.title")).toHaveValue("Data Table");
    await expect.poll(() => blockOnDisk("Data Table")?.columns).toEqual(["Item", "Year 1", "Year 2", "Year 3"]);
    expect(blockOnDisk("Data Table").row).toBe(27);
    expect(blockOnDisk("Data Table").cols).toBe(4);
  });

  test("dragging a block header moves it and persists", async () => {
    const seg = block("Segments");
    await dragBy(host, seg.locator(".bh"), 1 * CELL_W, 2 * CELL_H);
    await expect(seg).toHaveCSS("left", left(3));
    await expect(seg).toHaveCSS("top", top(9));
    await expect.poll(() => [blockOnDisk("Segments").row, blockOnDisk("Segments").col]).toEqual([9, 3]);
    // the side panel shows the new address
    await expect(app.locator(".side .sh")).toContainText("C9");
  });

  test("resizing from the corner grows the block and persists", async () => {
    const seg = block("Segments");
    await expect(seg.locator(".rs")).toHaveCount(4);   // all four corners resize (F19)
    await dragBy(host, seg.locator(".rs.se"), 2 * CELL_W, 3 * CELL_H);
    await expect(seg).toHaveCSS("width", `${6 * CELL_W - 1}px`);
    await expect(seg).toHaveCSS("height", `${8 * CELL_H - 1}px`);
    await expect.poll(() => [blockOnDisk("Segments").rows, blockOnDisk("Segments").cols]).toEqual([8, 6]);
  });

  test("resizing from the top left moves the block's cell, and one gesture is one undo", async () => {
    // Segments is at C9, 8 rows by 6 cols; pulling the top left grip in keeps the bottom right corner put
    const seg = block("Segments");
    await dragBy(host, seg.locator(".rs.nw"), 1 * CELL_W, 2 * CELL_H);
    await expect(seg).toHaveCSS("left", left(4));
    await expect(seg).toHaveCSS("top", top(11));
    await expect(seg).toHaveCSS("width", `${5 * CELL_W - 1}px`);
    await expect(seg).toHaveCSS("height", `${6 * CELL_H - 1}px`);
    await expect.poll(() => [blockOnDisk("Segments").row, blockOnDisk("Segments").col, blockOnDisk("Segments").rows, blockOnDisk("Segments").cols]).toEqual([11, 4, 6, 5]);
    await expect(app.locator(".side .sh")).toContainText("D11");
    // the whole drag is one step, so one undo puts the block back where it was
    await app.locator(".toolbar button[title^='Undo']").click();
    await expect(seg).toHaveCSS("left", left(3));
    await expect(seg).toHaveCSS("top", top(9));
    await expect.poll(() => [blockOnDisk("Segments").row, blockOnDisk("Segments").col, blockOnDisk("Segments").rows, blockOnDisk("Segments").cols]).toEqual([9, 3, 8, 6]);
  });

  test("Stack puts every block in one column in reading order", async () => {
    await app.locator(".toolbar button", { hasText: "Stack" }).click();
    // reading order by (row, col): Key Assumptions (2,2), Funnel (7,10), Segments (9,3), Ramp Chart (20,2), Data Table (27,2)
    await expect(block("Key Assumptions")).toHaveCSS("top", top(2));
    await expect(block("Funnel")).toHaveCSS("top", top(7));
    await expect(block("Funnel")).toHaveCSS("left", left(2));
    await expect(block("Segments")).toHaveCSS("top", top(13));
    await expect(block("Segments")).toHaveCSS("left", left(2));
    await expect(block("Ramp Chart")).toHaveCSS("top", top(22));
    await expect(block("Data Table")).toHaveCSS("top", top(29));
    await expect.poll(() => assumptions().blocks.map((b: any) => [b.title, b.row, b.col])).toEqual([
      ["Key Assumptions", 2, 2], ["Segments", 13, 2], ["Funnel", 7, 2], ["Ramp Chart", 22, 2], ["Data Table", 29, 2],
    ]);
  });

  test("Flow lays blocks side by side and wraps at column P", async () => {
    await app.locator(".toolbar button", { hasText: "Flow" }).click();
    // widths: 3, 4, 6, 5, 5 with a blank column between: B, F, K then M+5-1=17 > 16 so wrap
    await expect(block("Key Assumptions")).toHaveCSS("left", left(2));
    await expect(block("Key Assumptions")).toHaveCSS("top", top(2));
    await expect(block("Funnel")).toHaveCSS("left", left(6));
    await expect(block("Funnel")).toHaveCSS("top", top(2));
    await expect(block("Segments")).toHaveCSS("left", left(11));
    await expect(block("Segments")).toHaveCSS("top", top(2));
    // band height = tallest so far (Segments, 8 rows) + 1 blank row
    await expect(block("Ramp Chart")).toHaveCSS("left", left(2));
    await expect(block("Ramp Chart")).toHaveCSS("top", top(11));
    await expect(block("Data Table")).toHaveCSS("left", left(8));
    await expect(block("Data Table")).toHaveCSS("top", top(11));
    await expect.poll(() => blockOnDisk("Data Table").col).toBe(8);
  });

  // State here (after Flow): Key Assumptions (2,2) 4x3, Funnel (2,6) 5x4, Segments (2,11) 8x6, Ramp Chart (11,2) 6x5, Data Table (11,8) 6x4.
  test("shift-click adds to and removes from the selection; the last click is the primary", async () => {
    await host.settle();
    await block("Key Assumptions").click();
    await expect(block("Key Assumptions")).toHaveClass(/primary/);
    await expect(app.locator(".selbar")).toBeHidden();
    await block("Funnel").click({ modifiers: ["Shift"] });
    await expect(block("Funnel")).toHaveClass(/primary/);
    await expect(block("Key Assumptions")).toHaveClass(/sel/);
    await expect(block("Key Assumptions")).not.toHaveClass(/primary/);
    await expect(app.locator(".blk.sel")).toHaveCount(2);
    await expect(app.locator(".selbar .count")).toHaveText("2 selected");
    await expect(app.locator(".side .sh")).toContainText("Funnel"); // the inspector follows the primary
    // shift-click again removes it: the other block becomes primary and the bar goes away
    await block("Funnel").click({ modifiers: ["Shift"] });
    await expect(block("Funnel")).not.toHaveClass(/sel/);
    await expect(block("Key Assumptions")).toHaveClass(/primary/);
    await expect(app.locator(".selbar")).toBeHidden();
    await block("Funnel").click({ modifiers: ["Shift"] });
    await expect(app.locator(".blk.sel")).toHaveCount(2);
    await expect(app.locator(".minimap rect.mb.sel")).toHaveCount(2);
  });

  test("dragging a selected header moves the whole selection by the same delta", async () => {
    await dragBy(host, block("Funnel").locator(".bh"), 1 * CELL_W, 2 * CELL_H);
    await expect(block("Funnel")).toHaveCSS("left", left(7));
    await expect(block("Funnel")).toHaveCSS("top", top(4));
    await expect(block("Key Assumptions")).toHaveCSS("left", left(3));
    await expect(block("Key Assumptions")).toHaveCSS("top", top(4));
    await expect(block("Segments")).toHaveCSS("left", left(11)); // not selected, not moved
    await expect.poll(() => [blockOnDisk("Funnel").row, blockOnDisk("Funnel").col, blockOnDisk("Key Assumptions").row, blockOnDisk("Key Assumptions").col]).toEqual([4, 7, 4, 3]);
    await expect(app.locator(".blk.sel")).toHaveCount(2); // the selection survives the save and re-render
  });

  test("Align left moves the group to its leftmost column, one undo step", async () => {
    await block("Data Table").click(); // a plain click outside the selection replaces it
    await expect(app.locator(".blk.sel")).toHaveCount(1);
    await block("Funnel").click({ modifiers: ["Shift"] });
    await app.locator(".selbar button", { hasText: "Align left" }).click();
    await expect(block("Data Table")).toHaveCSS("left", left(7));
    await expect(block("Funnel")).toHaveCSS("left", left(7));
    await expect.poll(() => blockOnDisk("Data Table").col).toBe(7);
    await app.locator(".toolbar button[title^='Undo']").click();
    await expect(block("Data Table")).toHaveCSS("left", left(8));
    await app.locator(".toolbar button[title^='Redo']").click();
    await expect(block("Data Table")).toHaveCSS("left", left(7));
  });

  test("Distribute vertically spaces three blocks with equal blank rows, keeping their order", async () => {
    await app.locator(".selbar button", { hasText: "Deselect" }).click();
    await expect(app.locator(".blk.sel")).toHaveCount(0);
    // make room first: Data Table from row 11 to 17, so the three are Funnel 4-8, Ramp Chart 11-16, Data Table 17-21
    await dragBy(host, block("Data Table").locator(".bh"), 0, 6 * CELL_H);
    await expect(block("Data Table")).toHaveCSS("top", top(17));
    await block("Ramp Chart").click({ modifiers: ["Shift"] });
    await block("Funnel").click({ modifiers: ["Shift"] });
    await expect(app.locator(".selbar .count")).toHaveText("3 selected");
    await app.locator(".selbar button", { hasText: "Distribute vertically" }).click();
    // span 4..22 is 18 rows, the blocks take 5 + 6 + 5 = 16, so one blank row between each: 4, 10, 17
    await expect(block("Funnel")).toHaveCSS("top", top(4));
    await expect(block("Ramp Chart")).toHaveCSS("top", top(10));
    await expect(block("Data Table")).toHaveCSS("top", top(17));
    await expect.poll(() => [blockOnDisk("Funnel").row, blockOnDisk("Ramp Chart").row, blockOnDisk("Data Table").row]).toEqual([4, 10, 17]);
  });

  test("zoom scales the canvas with a transform; the minimap shows only while the sheet overflows", async () => {
    const canvas = app.locator(".canvas");
    await expect(canvas).toHaveCSS("transform", "matrix(1, 0, 0, 1, 0, 0)");
    await expect(app.locator(".minimap")).toBeVisible(); // 26 columns are wider than the viewport
    await expect(app.locator(".minimap rect.mb")).toHaveCount(5);
    await app.locator(".toolbar .zoom button", { hasText: "50%" }).click();
    await expect(canvas).toHaveCSS("transform", "matrix(0.5, 0, 0, 0.5, 0, 0)");
    await expect(app.locator(".toolbar .zoom button.on")).toHaveText("50%");
    await expect(block("Funnel")).toHaveCSS("left", left(7)); // blocks keep sheet coordinates; only the transform changes
    const fits = await app.locator(".canvaswrap").evaluate((w) => { const s = w.querySelector(".canvasscale") as HTMLElement; return s.offsetWidth <= w.clientWidth && s.offsetHeight <= w.clientHeight; });
    expect(await app.locator(".minimap").isHidden()).toBe(fits);
    // keys: ⌘0 back to 100%, ⌘- one step down
    await app.locator(".canvaswrap").press("Meta+0");
    await expect(canvas).toHaveCSS("transform", "matrix(1, 0, 0, 1, 0, 0)");
    await app.locator(".canvaswrap").press("Meta+Minus");
    await expect(canvas).toHaveCSS("transform", "matrix(0.75, 0, 0, 0.75, 0, 0)");
    await app.locator(".toolbar .zoom button", { hasText: "100%" }).click();
    await expect(canvas).toHaveCSS("transform", "matrix(1, 0, 0, 1, 0, 0)");
  });

  test("chips editor adds a column (commit on blur)", async () => {
    await block("Segments").click();
    await expect(app.locator(".side .sh")).toContainText("Segments");
    const chips = app.locator(".side .field", { hasText: "Columns" }).locator(".chips");
    await expect(chips.locator(".c")).toHaveCount(3);
    await chips.locator("button.add").click();
    const input = chips.locator("input");
    await input.fill("Growth");
    await input.blur();
    await expect(chips.locator(".c")).toHaveCount(4);
    await expect(chips.locator(".c").nth(3)).toContainText("Growth");
    await expect(block("Segments").locator(".bc .r.h > div")).toHaveCount(4);
    await expect.poll(() => blockOnDisk("Segments").columns).toEqual(["Segment", "Firms", "Spend", "Growth"]);
  });

  test("chips editor adds a column once on Enter", async () => {
    // Enter commits and re-renders, which blurs the detached input and would commit the same value twice;
    // chipsEditor's `done` guard is what keeps it to one (src/app/layout-view.ts).
    const chips = app.locator(".side .field", { hasText: "Columns" }).locator(".chips");
    await chips.locator("button.add").click();
    const input = chips.locator("input");
    await input.fill("Margin");
    await input.press("Enter");
    await expect(chips.locator(".c")).toHaveCount(5, { timeout: 2000 });
    await expect.poll(() => blockOnDisk("Segments").columns, { timeout: 2000 }).toEqual(["Segment", "Firms", "Spend", "Growth", "Margin"]);
  });

  test("Remove needs two clicks and then drops the block", async () => {
    await block("Data Table").click();
    const remove = app.locator(".side button.danger");
    await expect(remove).toHaveText("Remove");
    await remove.click();
    await expect(remove).toHaveText("Remove block?");
    await remove.click();
    await expect(app.locator(".blk")).toHaveCount(4);
    await expect(block("Data Table")).toHaveCount(0);
    await expect.poll(() => assumptions().blocks.map((b: any) => b.title)).not.toContain("Data Table");
    // nothing selected: the side panel is back to the tab summary
    await expect(app.locator(".side .sh")).toContainText("Tab");
  });

  test("Approve all layouts asks first while anything is flagged", async () => {
    await expect(app.locator(".ftr .msg")).toContainText("flagged");
    await app.locator(".ftr button.primary", { hasText: "Approve all layouts" }).click();
    const confirm = app.locator(".pop.confirm");
    await expect(confirm).toBeVisible();
    await expect(confirm.locator(".t")).toContainText(/still flagged/);
    await confirm.locator("button", { hasText: "Cancel" }).click();
    await expect(app.locator(".pop")).toHaveCount(0);
    expect(readProject().tabs.every((t: any) => !t.layoutApproved)).toBe(true);
    // approving anyway goes through
    await app.locator(".ftr button.primary", { hasText: "Approve all layouts" }).click();
    await app.locator(".pop.confirm button", { hasText: "Approve anyway" }).click();
    await expect(app.locator(".ftr .msg")).toHaveText("Layouts approved");
    await expect.poll(() => readProject().tabs.every((t: any) => t.layoutApproved)).toBe(true);
  });
});

test("switching tabs in the strip sticks after a fresh propose_layouts", async ({ browser }) => {
  const host = await Host.open(browser);
  await host.reset();
  await host.callPlain("update_plan", { ops: [{ op: "approve_tabs" }] });
  const app = await host.call("propose_layouts", { layouts: [
    { tab: TAB_IDS.assumptions, blocks: [{ type: "inputs", title: "Key Assumptions", rowLabels: ["Growth"] }] },
    { tab: TAB_IDS.ramp, blocks: [{ type: "calc", title: "Five Year Ramp", columns: ["Line", "Year 1"] }] },
  ] });
  await host.settle();
  await app.locator(".tabstrip button", { hasText: "Revenue Ramp" }).click();
  await expect(app.locator(".blk .bh", { hasText: "Five Year Ramp" })).toBeVisible();
  await app.locator(".blk .bh", { hasText: "Five Year Ramp" }).click();
  await expect(app.locator(".side .sh", { hasText: "Five Year Ramp" })).toBeVisible();
});

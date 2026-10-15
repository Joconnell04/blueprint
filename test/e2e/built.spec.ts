/**
 * Layer 3, both kinds. The workbook half builds a real .xlsx (python + the formula check) and asserts the
 * chrome around it; the deck half renders the seeded build snapshot, since build_deck is not wired up yet.
 */
import { test, expect, type Browser, type FrameLocator } from "@playwright/test";
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import fs from "node:fs";
import path from "node:path";
import { Host, screenshotOnFailure } from "./host";
import { ARTIFACTS, PLANNER_HOME, ROOT } from "./playwright.config";
import { CELL, DECK_BUILT_ID, SLIDE_IDS, TAB_IDS, TAB_NAMES, deckBuiltProject, deckSnapshot, demoProject, readProject, seed, seedDeckBuilt } from "./seed";

const SHOTS = path.join(ARTIFACTS, "visual");

/** One deviation the fidelity pass recorded, as the builder writes it: cells before, cells after, why. */
interface Move { slide: number; blockId: string; from: { col: number; row: number; cols: number; rows: number }; to: { col: number; row: number; cols: number; rows: number }; reason: string }
const MOVES: Move[] = [
  { slide: 2, blockId: "b_reach", from: { col: 21, row: 6, cols: 10, rows: 4 }, to: { col: 21, row: 6, cols: 10, rows: 5 }, reason: "fit" },
  { slide: 4, blockId: "b_moves", from: { col: 3, row: 6, cols: 28, rows: 8 }, to: { col: 3, row: 7, cols: 28, rows: 8 }, reason: "balance" },
];

/**
 * The same built deck with two deviations recorded: one block grew a row to fit, one band moved down to
 * balance. The snapshot's own shapes are put where the build put them and carry the record, so the
 * preview, the plan diff and the adjusted list all tell the same story.
 */
function seedAdjusted(): void {
  const project = deckBuiltProject();
  const snap = deckSnapshot() as { slides: any[] };
  for (const move of MOVES) {
    const slide = snap.slides.find((s) => s.index === move.slide);
    const record = { blockId: move.blockId, from: move.from, to: move.to, reason: move.reason };
    slide.adjusted = [record];
    for (const shape of slide.shapes as any[]) {
      if (shape.blockId !== move.blockId) continue;
      shape.y = (move.to.row - 1) * CELL;
      shape.h = move.to.rows * CELL;
      shape.adjusted = { from: move.from, to: move.to, reason: move.reason };
    }
  }
  fs.mkdirSync(path.join(PLANNER_HOME, "out"), { recursive: true });
  fs.writeFileSync(String(project.snapshotPath), JSON.stringify(snap, null, 2));
  seed({ ...project, buildSummary: { ...(project.buildSummary as object), adjusted: MOVES, snapshot: snap } });
}

const planBlock = (tabId: string, blockId: string) =>
  readProject(DECK_BUILT_ID).tabs.find((t: any) => t.id === tabId)?.blocks.find((b: any) => b.id === blockId);

/**
 * Does the vendored basic-host tell the canvas it takes downloads? It builds its AppBridge with one
 * literal capability object; today that object is `{openLinks:{}, serverTools, serverResources,
 * updateModelContext}` and carries no `downloadFile`, so the Save copy button must stay unbuilt.
 */
function hostAdvertisesDownload(): boolean {
  const bundle = fs.readFileSync(path.join(ROOT, "test/host/index.html"), "utf8");
  const at = bundle.indexOf("openLinks:{}");
  return at >= 0 && bundle.slice(at, at + 400).includes("downloadFile");
}

/** The Save copy button and its progress chip behave the same in both kinds, so both halves run this. */
async function expectSaveCopyMatchesHost(app: FrameLocator): Promise<void> {
  const save = app.locator('.side .sh button[title="Save copy"]');
  if (hostAdvertisesDownload()) {
    await expect(save).toBeVisible();
    await save.click();
    await expect(app.locator(".pop.savepop button")).not.toHaveCount(0);
    await app.locator(".pop.savepop button").first().press("Escape");
  } else {
    console.log("the vendored host advertises no downloadFile: asserting Save copy is hidden");
    await expect(save).toHaveCount(0);
    await expect(app.locator(".side .sh .savechip")).toHaveCount(0);
  }
  await expect(app.locator('.side .sh button[title="Reveal"]')).toBeVisible();
}

test.describe.configure({ mode: "serial", timeout: 120_000 });

/** skipLayouts, so one build_workbook call is the whole setup. A bad freeze plants a builder warning. */
const SHEETS = [
  { tab: TAB_IDS.readme, blocks: [] },
  { tab: TAB_IDS.assumptions, blocks: [
    { id: "a_in", type: "inputs", title: "Key Assumptions", row: 2, col: 2, cols: 3, rows: [
      { label: "Addressable businesses", value: 6400000, format: "count" },
      { label: "Penetration, Year 5", value: 0.04, format: "percent" },
      { label: "Average price ($)", value: 2400, format: "currency" }] }] },
  { tab: TAB_IDS.funnel, blocks: [
    { id: "f_tab", type: "table", title: "Funnel", row: 2, col: 2, cols: 3, columns: ["Stage", "Businesses", "Share"], colFormats: ["text", "count", "percent"],
      data: [["All US businesses", 32000000, 1], ["Addressable", 6400000, 0.2]] }] },
  { tab: TAB_IDS.ramp, freeze: "nope", blocks: [
    { id: "r_calc", type: "calc", title: "Five Year Ramp", row: 2, col: 2, cols: 3, columns: ["Line", "Year 1", "Year 2"], colFormats: ["text", "count", "count"], data: [
      ["Customers", { f: "={{a_in!Addressable businesses}}/5" }, { f: "=C4*(1+{{a_in!Penetration, Year 5}})" }],
      ["Revenue ($M)", { f: "=C4*{{a_in!Average price ($)}}/1000000" }, { f: "=D4*{{a_in!Average price ($)}}/1000000" }]] },
    { id: "r_out", type: "output", title: "Year 5 Outcome", row: 9, col: 2, cols: 3, rows: [
      { label: "Year 5 revenue ($M)", value: { f: "={{r_calc!1,2}}" }, format: "currency" },
      { label: "Spare capacity", value: { f: "=B60" }, format: "number" }] }] },
];

test.describe("built view: workbook", () => {
  let host: Host;
  let app: FrameLocator;
  screenshotOnFailure(() => host?.page);

  const field = (label: string) => app.locator(".side .field").filter({ hasText: label });

  test.beforeAll(async ({ browser }: { browser: Browser }) => {
    seed({ ...demoProject({ tabsApproved: true }), skipLayouts: true });
    host = await Host.open(browser);
    app = await host.call("build_workbook", { sheets: SHEETS });
    await expect(app.locator(".tabstrip button", { hasText: "Revenue Ramp" })).toBeVisible({ timeout: 60_000 });
  });
  test.afterAll(async () => { await host?.page.close(); });

  test("names the file, keeps the path in the tooltip and offers Reveal", async () => {
    const name = app.locator(".side .sh .t");
    await expect(name).toHaveText(/^Demo Market Sizing( v\d+)?\.xlsx$/);
    const path = (await name.getAttribute("title")) ?? "";
    expect(path.split("/").pop()).toBe(await name.textContent());
    expect(path).toContain("/out/");
    await expect(app.locator('.side .sh button[title="Reveal"]')).toBeVisible();
    await expect(app.locator(".ftr .msg")).toContainText("4 sheets");
    await expect(app.locator(".ftr button", { hasText: "Reopen layouts" })).toBeVisible();
  });

  test("one status line, expandable to the findings", async () => {
    const formulas = field("Formulas");
    await expect(formulas).toContainText(/problems?/);
    const status = formulas.locator("button.status");
    await expect(status).toHaveText(/\d+ problems?/);
    await expect(formulas.locator(".warnlist")).toBeHidden();
    await status.click();
    await expect(formulas.locator(".warnlist li").first()).toBeVisible();
    await expect(formulas.locator(".warnlist li .mono").first()).toContainText("!");
    await status.click();
    await expect(formulas.locator(".warnlist")).toBeHidden();
  });

  test("builder warnings sit behind a count chip", async () => {
    const builder = field("Builder");
    const chip = builder.locator("button.chip.warn");
    await expect(chip).toHaveText(/1 warning/);
    await expect(builder.locator(".warnlist")).toBeHidden();
    await chip.click();
    await expect(builder.locator(".warnlist li")).toHaveCount(1);
    await expect(builder.locator(".warnlist li")).toContainText("freeze");
  });

  test("the sheet list keeps its formulas column inside the panel, right aligned", async () => {
    const rows = app.locator(".side .sheetrow");
    await expect(rows).toHaveCount(4);
    const panel = (await app.locator(".side").boundingBox())!;
    const rights: number[] = [];
    for (let i = 0; i < 4; i++) {
      const box = (await rows.nth(i).locator(".n").last().boundingBox())!;
      expect(box.width).toBeGreaterThan(0);
      expect(box.x + box.width).toBeLessThanOrEqual(panel.x + panel.width - 10);
      rights.push(Math.round(box.x + box.width));
    }
    expect(new Set(rights).size).toBe(1); // one column, not four
  });

  test("Save copy follows the host's downloadFile capability", async () => {
    await expectSaveCopyMatchesHost(app);
  });

  test("the arrows toggle draws the arrows and is remembered across a re-render", async () => {
    const arrows = app.locator(".toolbar button", { hasText: "Arrows" });
    await expect(arrows).toBeVisible();
    await expect(arrows).not.toHaveClass(/(^|\s)on(\s|$)/);
    await arrows.click();
    await expect(app.locator(".toolbar button.on", { hasText: "Arrows" })).toBeVisible();
    await expect(app.locator("svg.deps")).toBeAttached();
    // the theme flip redraws the whole canvas, so the toggle can only be on if it was remembered
    await host.themeToggle.click();
    await host.settle();
    await expect(app.locator(".toolbar button.on", { hasText: "Arrows" })).toBeVisible();
    await host.themeToggle.click();
    await host.settle();
    await app.locator(".toolbar button.on", { hasText: "Arrows" }).click();
    await expect(app.locator("svg.deps")).toHaveCount(0);
  });
});

// ---------------------------------------------------------------- W11: the workbook's adjusted chip

const WB_ADJ_ID = "wbadjusted";
/**
 * A built workbook with two deviations, one in each shape the builder writes them: the sheet's own list
 * (a width snap, recorded as columns) and the bare reason string it leaves on the block it adjusted (a
 * fit). Neither carries cells, so the canvas reads the plan for the "from" and the snapshot's block
 * extent for the "to", which is also what Keep positions writes back.
 */
const WB_SHEETS: { name: string; plan: any[]; built: any[]; adjusted?: any[] }[] = [
  { name: TAB_NAMES[0], plan: [], built: [] },
  { name: TAB_NAMES[1],
    plan: [{ id: "a_in", type: "inputs", title: "Key Assumptions", row: 2, col: 2, rows: 4, cols: 2 }],
    built: [{ id: "a_in", type: "inputs", row: 2, col: 2, rows: 4, cols: 4, adjusted: "snap" }],
    adjusted: [{ blockId: "a_in", from: 2, to: 4, reason: "snap" }] },
  { name: TAB_NAMES[2], plan: [], built: [] },
  { name: TAB_NAMES[3],
    plan: [{ id: "r_calc", type: "calc", title: "Five Year Ramp", row: 2, col: 2, rows: 5, cols: 3, columns: ["Line", "Year 1", "Year 2"] }],
    built: [{ id: "r_calc", type: "calc", row: 3, col: 2, rows: 5, cols: 3, adjusted: "fit" }] },
];

function seedWorkbookAdjusted() {
  const base = demoProject({ tabsApproved: true });
  const out = path.join(PLANNER_HOME, "out");
  const buildPath = path.join(out, "Demo Market Sizing.xlsx");
  const snapshotPath = path.join(out, "Demo Market Sizing.snapshot.json");
  const snapshot = { sheets: WB_SHEETS.map((s) => ({
    name: s.name, widths: { A: 3, B: 30, C: 14, D: 14 }, maxRow: 10, maxCol: 6, merges: [], charts: [],
    cells: [{ a: "B2", v: s.plan[0]?.title ?? s.name, k: "title" }],
    blocks: s.built, ...(s.adjusted ? { adjusted: s.adjusted } : {}),
  })) };
  const project = {
    ...base,
    id: WB_ADJ_ID,
    tabs: base.tabs.map((t, i) => ({ ...t, approved: true, layoutProposed: true, layoutApproved: true, blocks: WB_SHEETS[i].plan })),
    buildPath,
    snapshotPath,
    buildSummary: { kind: "workbook", sheets: WB_SHEETS.map((s) => ({ name: s.name, cells: 1, formulas: 0 })), warnings: [] },
  };
  fs.mkdirSync(out, { recursive: true });
  fs.writeFileSync(snapshotPath, JSON.stringify(snapshot, null, 2));
  fs.writeFileSync(buildPath, "not really a workbook");
  seed(project);
  return project;
}

const wbBlock = (tabId: string, blockId: string) =>
  readProject(WB_ADJ_ID).tabs.find((t: any) => t.id === tabId)?.blocks.find((b: any) => b.id === blockId);

test.describe("built view: workbook adjustments", () => {
  let host: Host;
  let app: FrameLocator;
  screenshotOnFailure(() => host?.page);

  const open = async () => {
    app = await host.call("show_plan", { view: "built" }, { project: WB_ADJ_ID });
    await expect(app.locator(".tabstrip button").first()).toBeVisible({ timeout: 20_000 });
  };

  test.beforeAll(async ({ browser }: { browser: Browser }) => {
    seedWorkbookAdjusted();
    host = await Host.open(browser);
    await open();
  });
  test.afterAll(async () => { await host?.page.close(); });

  test("the adjusted chip lists what the builder moved, both shapes of record", async () => {
    const chip = app.locator(".toolbar button.adjchip");
    await expect(chip).toHaveText(/^adjusted 2/);
    await chip.click();
    const rows = app.locator(".pop.adjlist button.arow");
    await expect(rows).toHaveCount(2);
    // the sheet's own record: the block grew two columns to match its band
    await expect(rows.nth(0).locator(".ix")).toHaveText("2");
    await expect(rows.nth(0).locator(".t")).toHaveText("Key Assumptions");
    await expect(rows.nth(0).locator(".geo")).toHaveText("r2 c2 4×2 → r2 c2 4×4");
    await expect(rows.nth(0).locator(".why")).toHaveText("snap");
    // and the block's own reason string, with the geometry read off the two ends
    await expect(rows.nth(1).locator(".ix")).toHaveText("4");
    await expect(rows.nth(1).locator(".t")).toHaveText("Five Year Ramp");
    await expect(rows.nth(1).locator(".geo")).toHaveText("r2 c2 5×3 → r3 c2 5×3");
    await expect(rows.nth(1).locator(".why")).toHaveText("fit");
    // a line walks to the sheet it names
    await rows.nth(1).click();
    await expect(app.locator(".pop.adjlist")).toHaveCount(0);
    await expect(app.locator(".tabstrip button.on")).toHaveText(TAB_NAMES[3]);
  });

  test("Keep positions writes the built geometry into the plan and clears the chip", async () => {
    expect(wbBlock(TAB_IDS.assumptions, "a_in").cols).toBe(2);
    expect(wbBlock(TAB_IDS.ramp, "r_calc").row).toBe(2);
    await app.locator(".toolbar button.adjchip").click();
    await app.locator(".pop.adjlist button", { hasText: "Keep positions" }).click();
    await expect(app.locator(".toast")).toHaveText("Kept 2 sheets");
    await expect.poll(() => wbBlock(TAB_IDS.assumptions, "a_in")?.cols, { timeout: 15_000 }).toBe(4);
    await expect.poll(() => wbBlock(TAB_IDS.ramp, "r_calc")?.row).toBe(3);
    // the plan now says what the build did, so there is nothing left to report
    await expect(app.locator(".toolbar button.adjchip")).toHaveCount(0);
    await expect(app.locator(".toolbar .chip.ok", { hasText: "matches plan" })).toBeVisible();
  });
});

test.describe("built view: deck", () => {
  let host: Host;
  let app: FrameLocator;
  screenshotOnFailure(() => host?.page);

  const field = (label: string) => app.locator(".side .field").filter({ hasText: label });
  const open = async () => {
    app = await host.call("show_plan", { view: "built" }, { project: DECK_BUILT_ID });
    await expect(app.locator(".slidestrip .sl").first()).toBeVisible({ timeout: 20_000 });
  };

  test.beforeAll(async ({ browser }: { browser: Browser }) => {
    seedDeckBuilt();
    host = await Host.open(browser);
    await open();
  });
  test.afterAll(async () => { await host?.page.close(); });

  test("a thumbnail per slide, the first one large", async () => {
    await expect(app.locator(".hdr .chip.kind")).toHaveText("deck");
    await expect(app.locator(".slidestrip .sl")).toHaveCount(6);
    await expect(app.locator(".slidestrip .sl.on")).toHaveCount(1);
    await expect(app.locator(".slidestrip .sl").nth(0)).toHaveClass(/(^|\s)on(\s|$)/);
    // findings mark their own thumbnail: one warning on slide 4, one error on slide 5
    await expect(app.locator(".slidestrip .sl .dot.warn")).toHaveCount(2);
    await expect(app.locator(".slidestrip .sl .dot.bad")).toHaveCount(1);
    await expect(app.locator(".sslide .sprev.large")).toBeVisible();
    await expect(app.locator(".sslide .sprev.large text").first()).toContainText("Blueprint");
    await expect(app.locator(".ftr .msg")).toContainText("6 slides");
    await expect(app.locator(".toolbar .chip", { hasText: "shapes" })).toBeVisible();
    // the deck has no dependency arrows
    await expect(app.locator(".toolbar button", { hasText: "Arrows" })).toHaveCount(0);
  });

  test("lists the slides compactly and selects one", async () => {
    const rows = app.locator(".side .sheetrow.deck");
    await expect(rows).toHaveCount(6);
    await expect(rows.nth(0).locator(".ix")).toHaveText("1");
    await expect(rows.nth(0)).toHaveClass(/(^|\s)sel(\s|$)/);
    await expect(rows.nth(3).locator(".n").first()).toHaveText(/\d+ shapes/);
    await expect(rows.nth(3).locator(".n").last()).toHaveText("1"); // the overflow finding on slide 4
    await rows.nth(3).click();
    await expect(rows.nth(3)).toHaveClass(/(^|\s)sel(\s|$)/);
    await expect(app.locator(".slidestrip .sl").nth(3)).toHaveClass(/(^|\s)on(\s|$)/);
    await expect(app.locator(".toolbar .purpose")).toHaveText("Three moves, one owner each");
    // the cards block was built one row lower than the plan asked
    await expect(app.locator(".toolbar .chip.warn", { hasText: "off plan" })).toHaveText("1 off plan");
    await expect(app.locator(".sprev.large .svblk.diff")).toHaveCount(1);
  });

  test("the arrow keys walk the strip", async () => {
    await app.locator(".slidestrip .sl").nth(3).press("ArrowRight");
    await expect(app.locator(".slidestrip .sl").nth(4)).toHaveClass(/(^|\s)on(\s|$)/);
    await expect(app.locator(".toolbar .purpose")).toContainText("pilot");
    await app.locator(".slidestrip .sl").nth(4).press("ArrowLeft");
    await expect(app.locator(".slidestrip .sl").nth(3)).toHaveClass(/(^|\s)on(\s|$)/);
  });

  test("the status line expands, and a finding selects its slide and outlines the shape", async () => {
    const checks = field("Checks");
    const status = checks.locator("button.status");
    await expect(status).toHaveText("3 problems");
    await expect(checks.locator(".warnlist")).toBeHidden();
    await status.click();
    await expect(checks.locator(".warnlist li")).toHaveCount(3);
    await expect(checks.locator(".warnlist li .mono").nth(2)).toHaveText("Slide 5");
    await checks.locator(".warnlist li").nth(2).click();
    await expect(app.locator(".slidestrip .sl").nth(4)).toHaveClass(/(^|\s)on(\s|$)/);
    await expect(app.locator(".sprev.large .svblk.mark")).toHaveCount(1);
  });

  test("a block asks before it spends a turn", async () => {
    await app.locator(".side .sheetrow.deck").nth(3).click();
    await app.locator(".sprev.large .svblk").first().click();
    const pop = app.locator(".pop.confirm");
    await expect(pop.locator(".t")).toHaveText('Explain "The three moves"?');
    await expect(pop.locator("button", { hasText: "Ask" })).toBeVisible();
    await pop.locator("button", { hasText: "Cancel" }).click();
    await expect(pop).toHaveCount(0);
  });

  test("Save copy follows the host's downloadFile capability, deck too", async () => {
    await expectSaveCopyMatchesHost(app);
  });

  test("a PDF chip, only when the deck was exported", async () => {
    await expect(app.locator(".side .sh .chip")).toHaveText("PDF");
    expect(await app.locator(".side .sh .chip").getAttribute("title")).toMatch(/\.pdf$/);
    seedDeckBuilt({ pdf: false });
    await open();
    await expect(app.locator(".side .sh .chip")).toHaveCount(0);
    await expect(app.locator('.side .sh button[title="Reveal"]')).toBeVisible();
  });

  /**
   * The fidelity setting's own report: what the builder moved between the approved plan and the file,
   * and the one button that makes the plan agree with it (docs/research/output-quality.md).
   */
  test("the adjusted chip lists what the builder moved", async () => {
    seedAdjusted();
    await open();
    const chip = app.locator(".toolbar button.adjchip");
    await expect(chip).toHaveText(/^adjusted 2/);
    await chip.click();
    const rows = app.locator(".pop.adjlist button.arow");
    await expect(rows).toHaveCount(2);
    await expect(rows.nth(0).locator(".ix")).toHaveText("2");
    await expect(rows.nth(0).locator(".t")).toHaveText("Reachable today");
    await expect(rows.nth(0).locator(".geo")).toHaveText("r6 c21 4×10 → r6 c21 5×10");
    await expect(rows.nth(0).locator(".why")).toHaveText("fit");
    await expect(rows.nth(1).locator(".t")).toHaveText("The three moves");
    await expect(rows.nth(1).locator(".why")).toHaveText("balance");
    // a line walks to its slide, where the shape it names carries a dashed outline
    await rows.nth(1).click();
    await expect(app.locator(".pop.adjlist")).toHaveCount(0);
    await expect(app.locator(".slidestrip .sl").nth(3)).toHaveClass(/(^|\s)on(\s|$)/);
    await expect(app.locator(".sprev.large .svblk.adj")).toHaveCount(1);
  });

  test("Keep positions writes the built geometry into the plan and clears the chip", async () => {
    expect(planBlock(SLIDE_IDS.size, "b_reach").rows).toBe(4);
    expect(planBlock(SLIDE_IDS.move, "b_moves").row).toBe(6);
    await app.locator(".toolbar button.adjchip").click();
    await app.locator(".pop.adjlist button", { hasText: "Keep positions" }).click();
    await expect(app.locator(".toast")).toHaveText("Kept 2 slides");
    await expect.poll(() => planBlock(SLIDE_IDS.size, "b_reach")?.rows, { timeout: 15_000 }).toBe(5);
    await expect.poll(() => planBlock(SLIDE_IDS.move, "b_moves")?.row).toBe(7);
    // the plan now says what the build did, so there is nothing left to report on either slide
    await expect(app.locator(".toolbar button.adjchip")).toHaveCount(0);
    await expect(app.locator(".sprev.large .svblk.adj")).toHaveCount(0);
    await expect(app.locator(".toolbar .chip.ok", { hasText: "matches plan" })).toBeVisible();
  });

  test("shot: the built deck with the adjusted chip, dark and light", async () => {
    fs.mkdirSync(SHOTS, { recursive: true });
    seedAdjusted();
    await open();
    const idx = (await host.iframes.count()) - 1;
    await app.locator(".side .sheetrow.deck").nth(3).click();
    let dark = false;
    for (const theme of ["dark", "light"] as const) {
      if (dark !== (theme === "dark")) { await host.themeToggle.click(); dark = theme === "dark"; }
      await host.settle();
      await expect.poll(() => app.locator("html").getAttribute("data-theme")).toBe(theme);
      await app.locator(".toolbar button.adjchip").click();
      await expect(app.locator(".pop.adjlist")).toBeVisible();
      await host.page.waitForTimeout(300);
      await host.iframes.nth(idx).screenshot({ path: path.join(SHOTS, `c1-built-adjusted-${theme}.png`) });
      await app.locator(".sstage").press("Escape");
    }
    if (dark) await host.themeToggle.click();
    seedDeckBuilt();
  });
});

/**
 * The server half of Save copy. The canvas pulls the built file through ui_file_chunk and assembles the
 * base64; this drives the same loop over a stdio client against the same planner home, so the byte count
 * is checked against the file on disk rather than against the view's own arithmetic.
 */
test.describe("save copy: ui_file_chunk", () => {
  const CHUNK = 512_000;
  const out = path.join(PLANNER_HOME, "out");
  /** 1.2 MB, so the loop takes three chunks, and deterministic, so the assembled bytes can be compared. */
  const deckBytes = Buffer.from(Array.from({ length: 1_200_000 }, (_, i) => (i * 31 + 7) & 0xff));
  const pdfBytes = Buffer.from("%PDF-1.7 not really a pdf\n");
  let client: Client;
  let buildPath = "", pdfPath = "";

  const chunk = async (args: Record<string, unknown>) => {
    const r: any = await client.callTool({ name: "ui_file_chunk", arguments: { project: DECK_BUILT_ID, ...args } }).catch((e) => ({ isError: true, content: [{ type: "text", text: String(e) }] }));
    return { isError: !!r.isError, sc: r.structuredContent as { name: string; mimeType: string; total: number; offset: number; data: string } | undefined, text: r.content?.map((c: any) => c.text).join("\n") ?? "" };
  };

  test.beforeAll(async () => {
    const project = seedDeckBuilt();
    buildPath = String(project.buildPath); pdfPath = String(project.pdfPath);
    fs.mkdirSync(out, { recursive: true });
    fs.writeFileSync(buildPath, deckBytes);
    fs.writeFileSync(pdfPath, pdfBytes);
    client = new Client({ name: "built-spec", version: "0" });
    await client.connect(new StdioClientTransport({
      command: process.execPath,
      args: [path.join(ROOT, "dist", "server", "main.js")],
      cwd: ROOT,
      env: { ...process.env, BLUEPRINT_HOME: PLANNER_HOME, BLUEPRINT_OUTPUT: out } as Record<string, string>,
      stderr: "ignore",
    }));
  });
  test.afterAll(async () => { await client?.close(); seedDeckBuilt(); });

  test("names the file, reports the total and encodes exactly what it read", async () => {
    const { isError, sc } = await chunk({ which: "build", offset: 0, size: 300 });
    expect(isError).toBe(false);
    expect(sc!.name).toBe(path.basename(buildPath));
    expect(sc!.mimeType).toBe("application/vnd.openxmlformats-officedocument.presentationml.presentation");
    expect(sc!.total).toBe(deckBytes.length);
    expect(sc!.offset).toBe(0);
    expect(sc!.data.length).toBe(400); // base64 of 300 bytes, no padding
    expect(Buffer.from(sc!.data, "base64").equals(deckBytes.subarray(0, 300))).toBe(true);
  });

  test("the chunk loop assembles the right byte length", async () => {
    const parts: Buffer[] = [];
    let at = 0, total = 0, calls = 0;
    do {
      const { isError, sc } = await chunk({ which: "build", offset: at, size: CHUNK });
      expect(isError).toBe(false);
      calls++;
      total = sc!.total;
      expect(sc!.offset).toBe(at);
      const bytes = Buffer.from(sc!.data, "base64");
      expect(bytes.length).toBeGreaterThan(0);
      expect(bytes.length).toBeLessThanOrEqual(CHUNK);
      parts.push(bytes);
      at += bytes.length;
    } while (at < total);
    const whole = Buffer.concat(parts);
    expect(calls).toBe(3);
    expect(whole.length).toBe(deckBytes.length);
    expect(whole.equals(deckBytes)).toBe(true);
    // the tail chunk is short, and reading past the end is empty rather than an error
    expect(parts[2].length).toBe(deckBytes.length - 2 * CHUNK);
    const past = await chunk({ which: "build", offset: deckBytes.length + 10, size: 1000 });
    expect(past.isError).toBe(false);
    expect(past.sc!.data).toBe("");
    expect(past.sc!.offset).toBe(deckBytes.length);
  });

  test("the PDF and the plan JSON come through the same tool", async () => {
    const pdf = await chunk({ which: "pdf" });
    expect(pdf.isError).toBe(false);
    expect(pdf.sc!.name).toBe(path.basename(pdfPath));
    expect(pdf.sc!.mimeType).toBe("application/pdf");
    expect(Buffer.from(pdf.sc!.data, "base64").equals(pdfBytes)).toBe(true);

    const plan = await chunk({ which: "plan" });
    expect(plan.isError).toBe(false);
    expect(plan.sc!.name).toBe("Payments Partner Review.plan.json");
    expect(plan.sc!.mimeType).toBe("application/json");
    const parsed = JSON.parse(Buffer.from(plan.sc!.data, "base64").toString("utf8"));
    expect(parsed.id).toBe(DECK_BUILT_ID);
    expect(parsed.tabs).toHaveLength(6);
  });

  test("refuses a missing export, an unknown kind and an oversized chunk", async () => {
    expect((await chunk({ which: "sneaky" })).isError).toBe(true);
    expect((await chunk({ which: "build", size: CHUNK + 1 })).isError).toBe(true);
    expect((await chunk({ which: "build", offset: -1 })).isError).toBe(true);
    seedDeckBuilt({ pdf: false });
    const pdf = await chunk({ which: "pdf" });
    expect(pdf.isError).toBe(true);
    expect(pdf.text).toContain("no PDF");
    // the build itself is still readable, and there is no way to name any other path
    expect((await chunk({ which: "build", size: 16 })).isError).toBe(false);
  });
});

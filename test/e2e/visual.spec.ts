/**
 * Visual pass: every view, dark and light, at desktop width, into test/e2e/artifacts/visual/*.png.
 * The outline is also shot at the inline chat width (a 760 px page, so a canvas of about 700), which is
 * where the rows have to hold without overlapping. Each shot is the canvas iframe alone, not the host page.
 * Not a gate; run with `npx playwright test -c test/e2e/playwright.config.ts visual`.
 */
import { test, expect, type Browser, type FrameLocator } from "@playwright/test";
import fs from "node:fs";
import path from "node:path";
import { Host } from "./host";
import { TAB_IDS, DECK_BUILT_ID, DECK_DIV_ID, DEMO_ID, deckDividerProject, demoProject, seed, seedDeckBuilt } from "./seed";
import { PLANNER_HOME, ARTIFACTS } from "./playwright.config";
const readPlan = () => JSON.parse(fs.readFileSync(path.join(PLANNER_HOME, "projects", `${DEMO_ID}.json`), "utf8"));

const OUT = path.join(ARTIFACTS, "visual");
const LAYOUTS = [
  { tab: TAB_IDS.assumptions, blocks: [
    { type: "inputs", title: "Key Assumptions", rowLabels: ["Addressable businesses", "Penetration, Year 5", "Average price ($)", "Growth"], questions: ["Is price annual or monthly?"] },
    { type: "inputs", title: "Scenario Variables", rowLabels: ["Scenario", "Churn"] } ] },
  { tab: TAB_IDS.ramp, blocks: [
    { type: "calc", title: "Five Year Ramp", columns: ["Line", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5"], rowLabels: ["Customers", "Revenue ($M)"], sample: [["Customers", 375000, 450000, 525000, 600000, 675000], ["Revenue ($M)", 2034, 2440, 2847, 3254, 3661]] },
    { type: "chart", title: "Revenue by Year", row: 2, col: 9 },
    { type: "output", title: "Year 5 Outcome", rowLabels: ["Year 5 revenue ($M)"] } ] },
  { tab: TAB_IDS.funnel, blocks: [{ type: "table", title: "Funnel", columns: ["Stage", "Businesses", "Share"], sample: [["All businesses", 32000000, "100%"], ["Addressable", 6400000, "20%"]] }, { type: "output", title: "Serviceable Market", rowLabels: ["Serviceable businesses"] }] },
];

test.describe.configure({ mode: "serial", timeout: 120_000 });

test.describe("visual pass", () => {
  let host: Host;
  let idx = 0;      // the app iframe this pass is shooting
  let dark = false; // the vendored host starts in Playwright's default light scheme

  test.beforeAll(async ({ browser }: { browser: Browser }) => {
    fs.mkdirSync(OUT, { recursive: true });
    seed(demoProject());
    host = await Host.open(browser);
    await host.page.setViewportSize({ width: 1280, height: 800 });
  });
  test.afterAll(async () => { await host?.page.close(); });

  const setTheme = async (want: "dark" | "light") => {
    if (dark === (want === "dark")) return;
    await host.themeToggle.click();
    dark = want === "dark";
    await host.settle();
  };
  /**
   * One name, two files. The theme flip re-renders the canvas (and drops any popover), so `prep` runs
   * again for each theme to put the canvas back into the state being photographed.
   */
  const shot = async (name: string, prep?: () => Promise<void>) => {
    for (const theme of ["dark", "light"] as const) {
      await setTheme(theme);
      if (prep) await prep();
      await host.settle();
      await host.iframes.nth(idx).screenshot({ path: path.join(OUT, `${name}-${theme}.png`) });
    }
  };
  const track = async (app: FrameLocator) => { idx = (await host.iframes.count()) - 1; return app; };
  /** The inline chat width, where the rows are tightest; back to the desktop box afterwards. */
  const narrow = async (name: string, prep?: () => Promise<void>) => {
    await host.page.setViewportSize({ width: 760, height: 900 });
    await host.settle();
    await shot(name, prep);
    await host.page.setViewportSize({ width: 1280, height: 800 });
    await host.settle();
  };

  test("tabs view", async () => {
    const app = await track(await host.call("show_plan", { view: "tabs" }));
    await host.settle();
    await shot("01-tabs", async () => { await app.locator(".trow").nth(1).locator(".idx").click(); });
    await narrow("01n-tabs-narrow", async () => { await app.locator(".trow").nth(1).locator(".idx").click(); });
    await shot("02-style-popover", async () => { await app.locator(".stylebtn").click(); await expect(app.locator(".stylepop .stylelist")).toBeVisible(); });
    await shot("03-templates-popover", async () => { await app.getByRole("button", { name: "Templates" }).click(); await expect(app.locator(".stylepop")).toBeVisible(); });
  });

  test("deck outline", async () => {
    seed(deckDividerProject());
    const app = await track(await host.call("show_plan", { view: "tabs" }, { project: DECK_DIV_ID }));
    await host.settle();
    // a slide selected, so the row actions are in the shot where they live
    const select = async () => { await app.locator(".trow:not(.div)").nth(3).locator(".idx").click(); };
    await shot("10-deck-outline", select);
    await narrow("10n-deck-outline-narrow", select);
    await shot("11-deck-outline-story", async () => {
      await select();
      const more = app.locator(".trow:not(.div)").nth(1).locator("button.more");
      if (await more.isVisible() && (await more.textContent()) === "more") await more.click();
      await expect(app.locator(".trow:not(.div)").nth(1).locator("button.more")).toHaveText("less");
    });
  });

  test("layouts view", async () => {
    await host.callPlain("update_plan", { ops: [{ op: "approve_tabs" }] });
    const app = await track(await host.call("propose_layouts", { layouts: LAYOUTS }));
    await host.settle();
    const pick = (tab: string, block: string) => async () => {
      await app.locator(".tabstrip button", { hasText: tab }).click();
      await app.locator(".blk .bh", { hasText: block }).click();
    };
    await shot("04-layouts", pick("Revenue Ramp", "Five Year Ramp"));
    await shot("05-layouts-question", pick("Assumptions", "Key Assumptions"));
    await shot("06-layouts-selection", async () => {
      await app.locator(".tabstrip button", { hasText: "Assumptions" }).click();
      await app.locator(".blk .bh", { hasText: "Key Assumptions" }).click();
      await app.locator(".blk .bh", { hasText: "Scenario Variables" }).click({ modifiers: ["Shift"] });
      await expect(app.locator(".selbar")).toBeVisible();
    });
    await shot("07-cheatsheet", async () => {
      await app.locator(".tabstrip button", { hasText: "Revenue Ramp" }).click();
      await app.locator(".toolbar button", { hasText: "?" }).click();
      await expect(app.locator(".pop.cheats")).toBeVisible();
    });
  });

  test("built view", async () => {
    await host.callPlain("update_plan", { ops: [{ op: "approve_layout" }] });
    // answer the question, then build with real content
    const P = readPlan(); const notes = P.notes.filter((n: any) => n.status === "open");
    for (const n of notes) await host.callPlain("resolve_notes", { resolutions: [{ noteId: n.id, status: "declined", response: "Visual pass" }] });
    const bid = (tab: string, type: string) => P.tabs.find((t: any) => t.id === tab).blocks.find((b: any) => b.type === type).id;
    const A = bid(TAB_IDS.assumptions, "inputs"), S = P.tabs.find((t: any) => t.id === TAB_IDS.assumptions).blocks[1].id, F = bid(TAB_IDS.funnel, "table"), SM = bid(TAB_IDS.funnel, "output"), R = bid(TAB_IDS.ramp, "calc"), C = bid(TAB_IDS.ramp, "chart"), Y = bid(TAB_IDS.ramp, "output");
    const app = await track(await host.call("build_workbook", { sheets: [
      { tab: TAB_IDS.readme, blocks: [] },
      { tab: TAB_IDS.assumptions, blocks: [
        { id: A, rows: [{ label: "Addressable businesses", value: 6400000, format: "count", note: "Census count; see R1" }, { label: "Penetration, Year 5", value: 0.04, format: "percent" }, { label: "Average price ($)", value: 2400, format: "currency" }, { label: "Growth", value: 0.35, format: "percent" }] },
        { id: S, rows: [{ label: "Scenario", value: "Base" }, { label: "Churn", value: 0.1, format: "percent" }] } ] },
      { tab: TAB_IDS.funnel, blocks: [
        { id: F, columns: ["Stage", "Businesses", "Share"], colFormats: ["text", "count", "percent"], data: [["All US businesses", 32000000, 1], ["Addressable", { f: `={{${A}!Addressable businesses}}` }, { f: "=B5/B4" }]] },
        { id: SM, rows: [{ label: "Serviceable businesses", value: { f: `={{${F}!1,1}}*{{${A}!Penetration, Year 5}}` }, format: "count" }] } ] },
      { tab: TAB_IDS.ramp, blocks: [
        { id: R, columns: ["Line", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5"], colFormats: ["text", "count", "count", "count", "count", "count"], data: [
          ["Customers", { f: `={{${SM}!Serviceable businesses}}/5` }, { f: `=C4*(1+{{${A}!Growth}})` }, { f: `=D4*(1+{{${A}!Growth}})` }, { f: `=E4*(1+{{${A}!Growth}})` }, { f: `=F4*(1+{{${A}!Growth}})` }],
          ["Revenue ($M)", { f: `=C4*{{${A}!Average price ($)}}/1000000` }, { f: `=D4*{{${A}!Average price ($)}}/1000000` }, { f: `=E4*{{${A}!Average price ($)}}/1000000` }, { f: `=F4*{{${A}!Average price ($)}}/1000000` }, { f: `=G4*{{${A}!Average price ($)}}/1000000` }] ] },
        { id: C, chart: { kind: "bar", sourceBlock: R, categoriesCol: 0, seriesCols: [1, 2, 3, 4, 5] } },
        { id: Y, rows: [{ label: "Year 5 revenue ($M)", value: { f: `={{${R}!1,5}}` }, format: "currency" }] } ] },
    ] }));
    await expect(app.locator(".tabstrip button", { hasText: "Revenue Ramp" })).toBeVisible({ timeout: 20000 });
    const sheet = async () => { await app.locator(".tabstrip button", { hasText: "Revenue Ramp" }).click(); };
    await shot("08-built", sheet);
    await shot("09-built-arrows", async () => {
      await sheet();
      if (!(await app.locator(".toolbar button.on", { hasText: "Arrows" }).count())) await app.locator(".toolbar button", { hasText: "Arrows" }).click();
      await expect(app.locator(".toolbar button.on", { hasText: "Arrows" })).toBeVisible();
    });
    // Market Funnel is the only sheet with a link cell (a bare cross-sheet reference), so `.sc.k-link` is
    // the one cell kind the other built shots never photograph.
    await shot("08b-built-link", async () => {
      if (await app.locator(".toolbar button.on", { hasText: "Arrows" }).count()) await app.locator(".toolbar button", { hasText: "Arrows" }).click();
      await app.locator(".tabstrip button", { hasText: "Market Funnel" }).click();
      await expect(app.locator(".sc.k-link")).toBeVisible();
    });
  });
  test("deck built", async () => {
    seedDeckBuilt();
    const app = await track(await host.call("show_plan", { view: "built" }, { project: DECK_BUILT_ID }));
    await expect(app.locator(".slidestrip .sl").first()).toBeVisible({ timeout: 20_000 });
    await host.settle();
    const checks = app.locator(".side .field").filter({ hasText: "Checks" });
    await shot("12-deck-built", async () => {
      await app.locator(".side .sheetrow.deck").nth(3).click();
      if (await checks.locator(".warnlist").isHidden()) await checks.locator("button.status").click();
      await expect(checks.locator(".warnlist li").first()).toBeVisible();
    });
  });
});

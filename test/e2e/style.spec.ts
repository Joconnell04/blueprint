/**
 * The style popover: the picker filtered to the project's kind, a deck skill's buildable dot and the
 * Derive button a prose-only one gets, the "Output" row that keeps Folder and File name folded away,
 * and "New style" — extract_format over a fixture workbook, the preview it comes back with, and the
 * save that writes the SKILL.md and selects it.
 *
 * The suite's planner (playwright.config.ts) runs with BLUEPRINT_SKILLS inside the test home, so the
 * skills this spec plants and the one it saves land under test/e2e/artifacts, never in ~/.claude/skills.
 * global-setup wipes that home after the server starts, so beforeAll creates the dir itself.
 */
import { test, expect, type Browser, type FrameLocator } from "@playwright/test";
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { Host, screenshotOnFailure } from "./host";
import { DECK_ID, DEMO_ID, deckProject, demoProject, readProject, seed } from "./seed";
import { ARTIFACTS, PLANNER_HOME, ROOT } from "./playwright.config";

const SKILLS = path.join(PLANNER_HOME, "skills");     // BLUEPRINT_SKILLS: where a saved style is written
const SIDECARS = path.join(PLANNER_HOME, "styles");   // HOME/styles: deck tokens styles.ts merges in
const FIXTURES = path.join(PLANNER_HOME, "fixtures");
const SHOTS = path.join(ARTIFACTS, "visual");

const DECK_SKILL = "harbor-deck";  // a deck skill with tokens: buildable
const PROSE_SKILL = "acme-deck";        // a deck skill without them: tokens missing, Derive offered
const BOOK_SKILL = "acme-book";         // a workbook skill: never listed for a deck
const NEW_NAME = "E2E House";
const NEW_SLUG = "e2e-house";
const OUT_DIR = path.join(PLANNER_HOME, "out", "styled");  // what the Output fields are set to, inside the run's home
const OUT_NAME = "Q3 Model";

const tokens = (name: string) => JSON.stringify({ name, palette: { primary: "1F3864", secondary: "2E74B5", ink: "111111" }, fonts: { major: "Lato", minor: "Lato" } });

function plantSkill(name: string, description: string, body: string): void {
  const dir = path.join(SKILLS, name);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, "SKILL.md"), `---\nname: ${name}\ndescription: "${description}"\n---\n# ${name}\n${body}\n`);
}

function plantSidecar(slug: string): void {
  fs.mkdirSync(SIDECARS, { recursive: true });
  fs.writeFileSync(path.join(SIDECARS, `${slug}.json`), tokens(slug));
}

test.describe.configure({ mode: "serial", timeout: 180_000 });

test.describe("style popover", () => {
  let host: Host;
  let book: FrameLocator;   // the workbook project's canvas
  let deck: FrameLocator;   // the deck project's canvas
  let deckIdx = 0;          // which app iframe the deck canvas is, for the screenshots
  let dark = false;         // the vendored host starts in Playwright's default light scheme
  screenshotOnFailure(() => host?.page);

  const pop = (app: FrameLocator) => app.locator(".stylepop.stylefmt");
  /** The sheet's one footer: Cancel and the single primary, whatever that currently is. */
  const foot = (app: FrameLocator) => pop(app).locator(".popfoot");
  const rowIds = (app: FrameLocator) => pop(app).locator(".stylelist .srow").evaluateAll((els) => els.map((e) => e.getAttribute("data-style") ?? "").sort());
  /** Reopen the sheet: clicking the anchor again replaces the popover, so every open reads ui_styles afresh. */
  const openSheet = async (app: FrameLocator) => {
    await app.locator(".hdr .stylebtn").click();
    await expect(pop(app).locator(".stylelist .srow").first()).toBeVisible();
    return pop(app);
  };

  test.beforeAll(async ({ browser }: { browser: Browser }) => {
    fs.mkdirSync(SKILLS, { recursive: true });
    fs.mkdirSync(SHOTS, { recursive: true });
    plantSkill(DECK_SKILL, "Harbor deck format: navy palette, Lato type, pipe kickers.", "Kickers carry a pipe. Cards take a 2.25pt border.");
    plantSkill(PROSE_SKILL, "Acme slide format: prose rules for every Acme deck.", "Titles sentence case. One idea per slide.");
    plantSkill(BOOK_SKILL, "Acme workbook format: Arial 10, blue inputs, no gridlines.", "Inputs are blue. Totals carry a top rule.");
    plantSidecar(DECK_SKILL);
    execFileSync("python3", ["builder/fixtures/make_fixtures.py", FIXTURES], { cwd: ROOT, stdio: "pipe" });
    for (const f of ["house.xlsx", "deck.pptx", "logo.png"]) expect(fs.existsSync(path.join(FIXTURES, f)), `fixture ${f}`).toBe(true);

    seed(demoProject());
    seed(deckProject());
    host = await Host.open(browser);
    await host.page.setViewportSize({ width: 1280, height: 800 });
    book = await host.call("show_plan", { view: "tabs" });
    deck = await host.call("show_plan", { view: "tabs" }, { project: DECK_ID });
    deckIdx = (await host.iframes.count()) - 1;
    await host.settle();
  });

  test.afterAll(async () => { await host?.page.close(); });

  test("lists the styles that fit a workbook", async () => {
    const sheet = await openSheet(book);
    // containment, not equality: whatever format skills this Mac has are listed too
    const ids = await rowIds(book);
    expect(ids).toEqual(expect.arrayContaining(["plain", "house", "template", `skill:${BOOK_SKILL}`]));
    expect(ids).not.toContain(`skill:${DECK_SKILL}`);
    expect(ids).not.toContain(`skill:${PROSE_SKILL}`);
    await expect(sheet.locator(`.srow[data-style="skill:${BOOK_SKILL}"] .chip.line`).first()).toHaveText("skill");
    await expect(sheet.locator(".srow .dot.buildable")).toHaveCount(0); // a workbook skill builds from its own prose
    // the template row keeps its path field, and the field appears when the row is picked
    await expect(sheet.locator(".tplf")).toBeHidden();
    await sheet.locator('.srow[data-style="template"] .styleopt').click();
    await expect(sheet.locator(".tplf input")).toHaveAttribute("placeholder", "/path/to/workbook.xlsx");
    await sheet.locator('.srow[data-style="house"] .styleopt').click();
    await expect(sheet.locator(".tplf")).toBeHidden();
  });

  test("lists deck styles for a deck, with a buildable dot and Derive", async () => {
    const sheet = await openSheet(deck);
    const ids = await rowIds(deck);
    expect(ids).toEqual(expect.arrayContaining(["plain", "house", "template", `skill:${DECK_SKILL}`, `skill:${PROSE_SKILL}`]));
    expect(ids).not.toContain(`skill:${BOOK_SKILL}`); // a workbook-only skill cannot dress a slide
    await expect(sheet.locator(`.srow[data-style="skill:${BOOK_SKILL}"]`)).toHaveCount(0);
    // tokens on disk: the deck builder can dress a slide from this one
    await expect(sheet.locator(`.srow[data-style="skill:${DECK_SKILL}"] .dot.buildable`)).toBeVisible();
    // prose only: no tokens, so the row says so and offers to have Claude derive them
    const prose = sheet.locator(`.srow[data-style="skill:${PROSE_SKILL}"]`);
    await expect(prose.locator(".chip.tokmiss")).toHaveText("tokens missing");
    await expect(prose.locator(".dot.buildable")).toHaveCount(0);
    await expect(sheet.locator(".stylenew .kindseg button.on")).toHaveText("Deck"); // new styles default to the project's kind
    await expect(sheet.locator(".stylenew textarea.npaths")).toHaveAttribute("placeholder", "/path/to/deck.pptx");

    await prose.locator("button.derive").click();
    await expect(deck.locator(".toast")).toBeVisible();
    expect(["Sent to Claude", "Tell Claude to continue"]).toContain((await deck.locator(".toast").innerText()).trim());
  });

  test("a tokens sidecar turns the same skill buildable", async () => {
    plantSidecar(PROSE_SKILL); // what extract_format writes when Claude derives tokens from the prose
    const sheet = await openSheet(deck);
    await expect(sheet.locator(`.srow[data-style="skill:${PROSE_SKILL}"] .dot.buildable`)).toBeVisible();
    await expect(sheet.locator(`.srow[data-style="skill:${PROSE_SKILL}"] .chip.tokmiss`)).toHaveCount(0);
    await expect(sheet.locator(`.srow[data-style="skill:${PROSE_SKILL}"] button.derive`)).toHaveCount(0);
  });

  test("Output keeps the folder and file name folded away", async () => {
    const sheet = await openSheet(book);
    // collapsed on open: one row, and the file name on it
    await expect(sheet.locator(".outf")).toBeHidden();
    await expect(sheet.locator(".outrow .label")).toHaveText("Output");
    await expect(sheet.locator(".outrow .onow")).toHaveText("Demo Market Sizing.xlsx");
    await expect(sheet.locator(".outrow")).toHaveAttribute("aria-expanded", "false");

    await sheet.locator(".outrow").click();
    await expect(sheet.locator(".outf")).toBeVisible();
    await expect(sheet.locator(".outrow")).toHaveAttribute("aria-expanded", "true");
    await expect(sheet.locator(".outf .field .label")).toHaveText(["Folder", "File name"]);
    // two fields is all it adds: one footer under the sheet, one primary in it
    await expect(pop(book).locator(".popfoot")).toHaveCount(1);
    await expect(pop(book).locator("button.primary")).toHaveCount(1);
    await expect(foot(book).locator("button.primary")).toHaveText("Save");

    // it stays open while the sheet is open, and shuts on a second click
    await sheet.locator('.srow[data-style="plain"] .styleopt').click();
    await expect(sheet.locator(".outf")).toBeVisible();
    await sheet.locator(".outrow").click();
    await expect(sheet.locator(".outf")).toBeHidden();
    // the next sheet starts collapsed again
    await expect((await openSheet(book)).locator(".outf")).toBeHidden();
  });

  test("Output values save with the style", async () => {
    const sheet = await openSheet(book);
    await sheet.locator(".outrow").click();
    await sheet.locator(".outf input").first().fill(OUT_DIR);
    await sheet.locator(".outf input").nth(1).fill(OUT_NAME);
    await sheet.locator(".outrow").click();  // shut again: the row reads back what was typed
    await expect(sheet.locator(".outrow .onow")).toHaveText(`${OUT_NAME}.xlsx`);

    // the footer's one Save carries both fields, folded or not
    await foot(book).locator("button.primary").click();
    await expect.poll(() => readProject(DEMO_ID).outputDir, { timeout: 15_000 }).toBe(OUT_DIR);
    expect(readProject(DEMO_ID).fileName).toBe(OUT_NAME);
    await host.settle();

    const again = await openSheet(book);
    await expect(again.locator(".outrow .onow")).toHaveText(`${OUT_NAME}.xlsx`);
    await again.locator(".outrow").click();
    await expect(again.locator(".outf input").first()).toHaveValue(OUT_DIR);
    await expect(again.locator(".outf input").nth(1)).toHaveValue(OUT_NAME);
  });

  /**
   * What the builder may change between the approved plan and the file (docs/research/output-quality.md,
   * "The fidelity setting"): one segmented control and three extras, under Folder and File name.
   */
  test("Fidelity and the build extras sit under Output, deck only", async () => {
    const sheet = await openSheet(deck);
    await expect(sheet.locator(".buildf")).toBeHidden();
    await sheet.locator(".outrow").click();
    await expect(sheet.locator(".buildf")).toBeVisible();
    await expect(sheet.locator(".buildf .fidf .label")).toHaveText("Fidelity");
    await expect(sheet.locator(".buildf .fidseg button")).toHaveText(["Strict", "Balanced", "Loose"]);
    await expect(sheet.locator(".buildf .fidseg button.on")).toHaveText("Balanced");
    await expect(sheet.locator(".buildf .fidline")).toHaveText("Resizes and balances within two cells.");
    await expect(sheet.locator(".buildf .boptf .label")).toHaveText("Include");
    await expect(sheet.locator(".buildf .bopt")).toHaveText(["Notes from story", "Cover", "Agenda"]);
    for (const opt of ["notesFromStory", "cover", "agenda"]) {
      await expect(sheet.locator(`.buildf .bopt[data-opt="${opt}"]`)).toHaveAttribute("aria-pressed", "true");
    }
    // one footer, one primary: the sheet still has exactly one commit in it
    await expect(pop(deck).locator("button.primary")).toHaveCount(1);
    await expect(foot(deck).locator("button.primary")).toHaveText("Save");
    // a workbook has no fidelity and nothing generated, so its Output row is the two fields alone
    const book2 = await openSheet(book);
    await book2.locator(".outrow").click();
    await expect(book2.locator(".outf")).toBeVisible();
    await expect(book2.locator(".buildf")).toBeHidden();
  });

  test("Fidelity and the extras save as they are picked", async () => {
    const sheet = await openSheet(deck);
    await sheet.locator(".outrow").click();
    await sheet.locator(".buildf .fidseg button", { hasText: "Strict" }).click();
    await expect(sheet.locator(".buildf .fidseg button.on")).toHaveText("Strict");
    await expect(sheet.locator(".buildf .fidline")).toHaveText("Draws the plan exactly.");
    await expect.poll(() => readProject(DECK_ID).build?.fidelity, { timeout: 15_000 }).toBe("strict");
    const cover = sheet.locator('.buildf .bopt[data-opt="cover"]');
    await cover.click();
    await expect(cover).toHaveAttribute("aria-pressed", "false");
    await expect.poll(() => readProject(DECK_ID).build?.cover, { timeout: 15_000 }).toBe(false);
    // the other two are left alone, and the next sheet reads back what was stored
    expect(readProject(DECK_ID).build?.agenda).toBe(undefined);
    // saving on change must not pull the canvas out from under the sheet: it is still the same popover
    await host.page.waitForTimeout(3000);   // one poll interval and a bit
    await expect(sheet.locator(".buildf .fidseg button.on")).toHaveText("Strict");
    await host.settle();
    const again = await openSheet(deck);
    await again.locator(".outrow").click();
    await expect(again.locator(".buildf .fidseg button.on")).toHaveText("Strict");
    await expect(again.locator('.buildf .bopt[data-opt="cover"]')).toHaveAttribute("aria-pressed", "false");
    await expect(again.locator('.buildf .bopt[data-opt="agenda"]')).toHaveAttribute("aria-pressed", "true");
    // back to the defaults, so the later shots show the sheet as it ships
    await again.locator(".buildf .fidseg button", { hasText: "Balanced" }).click();
    await again.locator('.buildf .bopt[data-opt="cover"]').click();
    await expect.poll(() => readProject(DECK_ID).build?.fidelity, { timeout: 15_000 }).toBe("balanced");
  });

  test("shot: the Output row open on a deck, dark and light", async () => {
    for (const theme of ["dark", "light"] as const) {
      if (dark !== (theme === "dark")) { await host.themeToggle.click(); dark = theme === "dark"; await host.settle(); }
      await expect.poll(() => deck.locator("html").getAttribute("data-theme")).toBe(theme);
      const sheet = await openSheet(deck);
      await sheet.locator(".outrow").click();
      await expect(sheet.locator(".buildf")).toBeVisible();
      await host.settle();
      await host.page.waitForTimeout(300);
      await host.iframes.nth(deckIdx).screenshot({ path: path.join(SHOTS, `c1-style-output-${theme}.png`) });
    }
  });

  test("a drop fills the paths field only when it carries a path", async () => {
    const sheet = await openSheet(book);
    const zone = sheet.locator(".dropz");
    // Finder sends the path as text/uri-list: that is the one drop the sandbox can use
    await zone.evaluate((el, uri) => {
      const dt = new DataTransfer();
      dt.setData("text/uri-list", uri);
      el.dispatchEvent(new DragEvent("drop", { dataTransfer: dt, bubbles: true, cancelable: true }));
    }, "file:///tmp/from-finder.pptx");
    await expect(sheet.locator("textarea.npaths")).toHaveValue("/tmp/from-finder.pptx");
    await expect(sheet.locator(".drophint")).toHaveText("Drop files here");
    // bytes only: File.path does not exist in the sandbox, so the hint says where the path comes from
    await zone.evaluate((el) => {
      const dt = new DataTransfer();
      dt.items.add(new File(["x"], "sales.pptx", { type: "application/octet-stream" }));
      el.dispatchEvent(new DragEvent("drop", { dataTransfer: dt, bubbles: true, cancelable: true }));
    });
    await expect(sheet.locator(".drophint")).toHaveText("Attach in chat, then paste the path");
    await expect(sheet.locator("textarea.npaths")).toHaveValue("/tmp/from-finder.pptx"); // and nothing else
  });

  test("Extract previews what a fixture workbook gave up", async () => {
    const sheet = await openSheet(book);
    await expect(sheet.locator(".stylenew .kindseg button.on")).toHaveText("Workbook");
    await sheet.locator(".stylenew input.nname").fill(NEW_NAME);
    await sheet.locator(".stylenew textarea.npaths").fill(path.join(FIXTURES, "house.xlsx"));
    await sheet.getByRole("button", { name: "Extract" }).click();
    await expect(sheet.locator(".xprev .xrow").first()).toBeVisible({ timeout: 60_000 });
    await expect(sheet.locator(".xprev .xsec .label")).toHaveText("Workbook");
    await expect(sheet.locator(".xprev .xsec .chip")).toHaveText(/%$/);
    await expect(sheet.locator(".xprev .xrow .dot.ok").first()).toBeVisible();
    const keys = await sheet.locator(".xprev .xrow .xk").allInnerTexts();
    expect(keys.length).toBeGreaterThan(0);
    expect(keys.some((k) => ["font", "size", "input_color", "title_rule_color", "label_divider", "currency"].includes(k))).toBe(true);
    await expect(sheet.locator(".xprev .xrow").filter({ hasText: "font" }).first().locator(".xv")).toHaveText("Arial");
    for (const v of await sheet.locator(".xprev .xrow .xv").allInnerTexts()) expect(v.trim()).not.toBe("");
    // the questions the files cannot answer go to Claude, not into a token: twelve words, the rest on hover
    const ask = sheet.locator(".xprev .xask li").first();
    await expect(ask).toBeVisible();
    const question = (await ask.innerText()).trim();
    expect(question.split(/\s+/).length, question).toBeLessThanOrEqual(12);
    expect(((await ask.getAttribute("title")) ?? "").length).toBeGreaterThan(question.length);
    // one footer, one primary: the second Save that used to sit under the preview is gone
    await expect(sheet.locator("button.primary")).toHaveCount(1);
    await expect(foot(book).locator("button.primary")).toHaveText("Save style");
    expect(fs.existsSync(path.join(SKILLS, NEW_SLUG))).toBe(false); // a preview writes nothing
  });

  test("Save writes the skill and selects it", async () => {
    await foot(book).locator("button.primary").click();
    await expect.poll(() => fs.existsSync(path.join(SKILLS, NEW_SLUG, "SKILL.md")), { timeout: 60_000 }).toBe(true);
    const md = fs.readFileSync(path.join(SKILLS, NEW_SLUG, "SKILL.md"), "utf8");
    expect(md).toMatch(/^---\nname: e2e-house\n/);
    expect(md).toMatch(/\nformat:\n/);
    await expect.poll(() => readProject(DEMO_ID).styleId).toBe(`skill:${NEW_SLUG}`);
    await expect(book.locator(".hdr .stylebtn")).toHaveText(NEW_SLUG);
    const sheet = await openSheet(book);
    await expect(sheet.locator(`.srow[data-style="skill:${NEW_SLUG}"] .styleopt.on`)).toBeVisible();
    await expect(sheet.locator(`.srow[data-style="skill:${NEW_SLUG}"] .chip.line`).first()).toHaveText("skill");
  });

  test("saving the same name again asks before replacing", async () => {
    const sheet = await openSheet(book);
    await sheet.locator(".stylenew input.nname").fill(NEW_NAME);
    await sheet.locator(".stylenew textarea.npaths").fill(path.join(FIXTURES, "house.xlsx"));
    await sheet.getByRole("button", { name: "Extract" }).click();
    await expect(sheet.locator(".xprev .xrow").first()).toBeVisible({ timeout: 60_000 });
    const before = fs.statSync(path.join(SKILLS, NEW_SLUG, "SKILL.md")).mtimeMs;
    // the preview says the skill is already there, so the footer offers Replace without a refused save
    const replace = foot(book).locator("button.primary");
    await expect(replace).toHaveText("Replace style");
    await expect(sheet.locator(".xerr")).toBeHidden();
    await replace.click();
    await expect(replace).toHaveText("Replace?");
    await replace.click();
    await expect.poll(() => fs.statSync(path.join(SKILLS, NEW_SLUG, "SKILL.md")).mtimeMs, { timeout: 60_000 }).toBeGreaterThan(before);
  });

  test("a screenshot among the paths asks Claude to read it", async () => {
    const sheet = await openSheet(book);
    await sheet.locator(".stylenew input.nname").fill("Shot Style");
    await sheet.locator(".stylenew textarea.npaths").fill([path.join(FIXTURES, "house.xlsx"), path.join(FIXTURES, "logo.png")].join("\n"));
    await sheet.getByRole("button", { name: "Extract" }).click();
    // the tool hands the image to Claude, not to the canvas, so the preview asks Claude to read it
    const read = sheet.locator(".xacts button", { hasText: "Read screenshots" });
    await expect(read).toBeVisible({ timeout: 60_000 });
    await read.click();
    await expect(book.locator(".toast")).toBeVisible();
    expect(fs.existsSync(path.join(SKILLS, "shot-style"))).toBe(false); // still a preview
  });

  test("shot: the new style sheet, dark and light", async () => {
    for (const theme of ["dark", "light"] as const) {
      if (dark !== (theme === "dark")) { await host.themeToggle.click(); dark = theme === "dark"; await host.settle(); }
      const sheet = await openSheet(deck);
      // a name no skill has yet, so the footer shows the shot the sheet is really for: one Save style
      await sheet.locator(".stylenew input.nname").fill("Acme Slides");
      await sheet.locator(".stylenew textarea.npaths").fill(path.join(FIXTURES, "deck.pptx"));
      await sheet.getByRole("button", { name: "Extract" }).click();
      await expect(sheet.locator(".xprev .xrow").first()).toBeVisible({ timeout: 60_000 });
      await expect(foot(deck).locator("button.primary")).toHaveText("Save style");
      // the sheet's body scrolls under a fixed footer: put the "New style" heading at its top so the shot is the whole section
      await sheet.evaluate((el) => {
        const body = el.querySelector<HTMLElement>(".stylebody");
        const head = el.querySelector<HTMLElement>(".sectop");
        if (body && head) body.scrollTop = head.offsetTop - body.offsetTop - 4;
      });
      await host.settle();
      await host.page.waitForTimeout(400); // one frame past --dur-3, so no row is mid-FLIP behind the sheet
      await host.iframes.nth(deckIdx).screenshot({ path: path.join(SHOTS, `13-style-new-${theme}.png`) });
    }
  });
});

import { test, expect, type Browser, type FrameLocator } from "@playwright/test";
import { Host, dragBy, screenshotOnFailure } from "./host";
import { DECK_DIV_ID, DECK_LONG_ID, DIVIDER_IDS, SECTIONS, SLIDE_NAMES, TAB_NAMES, deckDividerProject, demoProject, longDeckProject, readProject, seed } from "./seed";

test.describe.serial("tabs view", () => {
  let host: Host;
  let app: FrameLocator;
  screenshotOnFailure(() => host?.page);

  const names = () => app.locator(".tablist .trow input.tname");
  const values = async () => names().evaluateAll((els) => els.map((e) => (e as HTMLInputElement).value));
  // Names are input values (properties, not attributes), so find the row by index.
  const row = async (name: string) => { const i = (await values()).indexOf(name); expect(i, `tab "${name}" present`).toBeGreaterThanOrEqual(0); return app.locator(".tablist .trow").nth(i); };

  test.beforeAll(async ({ browser }: { browser: Browser }) => {
    seed(demoProject());
    host = await Host.open(browser);
    app = await host.call("show_plan", { view: "tabs" });
  });
  test.afterAll(async () => { await host?.page.close(); });

  test("renders the seeded tabs", async () => {
    await expect(app.locator(".hdr .name")).toHaveText("Demo Market Sizing");
    await expect(app.locator(".hdr .chip.kind")).toHaveText("workbook");
    await expect(app.locator(".hdr .seg button").first()).toHaveText("Tabs");
    await expect(names()).toHaveCount(4);
    expect(await values()).toEqual(TAB_NAMES);
    // the row carries its name and its purpose and nothing else: no message line, no section pill
    await expect(app.locator(".tablist input.tmessage")).toHaveCount(0);
    await expect(app.locator(".tablist button.sect")).toHaveCount(0);
    // the resolved note on Assumptions leaves the row's one state dot
    await expect((await row("Assumptions")).locator(".dot.ok")).toHaveCount(1);
    await expect(app.locator(".ftr .msg")).toHaveText("4 tabs");
  });

  test("folds the brief to one line and drops its dot once the takeaway is filled", async () => {
    // seeded with an audience and a goal but no takeaway: folded already, and readiness flags it here only
    await expect(app.locator(".brief .bl")).toContainText("For The CFO: Size the US SMB payments market");
    await expect(app.locator(".brief .dot.warn")).toHaveCount(1);
    await app.locator(".brief button.link", { hasText: "Edit" }).click();
    await expect(app.locator(".brief input.baudience")).toHaveValue("The CFO");
    await expect(app.locator(".brief .chips")).toHaveCount(0); // emphasis is gone from the card
    const takeaway = app.locator(".brief input.btakeaway");
    await expect(takeaway).toHaveAttribute("placeholder", "One sentence they should remember");
    await takeaway.fill("Year 5 revenue clears $3.6B");
    await takeaway.press("Enter");
    await expect.poll(() => readProject().briefCard.takeaway).toBe("Year 5 revenue clears $3.6B");
    // filled: the card folds to one line and the dot is gone
    await expect(app.locator(".brief .bl")).toContainText("For The CFO: Year 5 revenue clears $3.6B");
    await expect(app.locator(".brief .dot.warn")).toHaveCount(0);
  });

  test("adds a tab from the inline row with Enter", async () => {
    await app.locator(".addrow button.trigger").click();
    const name = app.locator(".addrow input.tname");
    await name.fill("Sources");
    await app.locator(".addrow input.tpurpose").fill("Every reference cited");
    await name.press("Enter");
    await expect(names()).toHaveCount(5);
    expect(await values()).toEqual([...TAB_NAMES, "Sources"]);
    await expect.poll(() => readProject().tabs.map((t: any) => t.name)).toEqual([...TAB_NAMES, "Sources"]);
    // the add row is back to its trigger after the re-render, and a workbook is offered no divider
    await expect(app.locator(".addrow button.trigger")).toHaveText("Add tab");
    await expect(app.locator(".addrow button.adddiv")).toHaveCount(0);
  });

  test("renames a tab by editing the name input", async () => {
    const input = (await row("Market Funnel")).locator("input.tname");
    await input.fill("TAM Build");
    await input.press("Enter");
    await expect.poll(values).toEqual(["Read Me", "Assumptions", "TAM Build", "Revenue Ramp", "Sources"]);
    await expect.poll(() => readProject().tabs[2].name).toBe("TAM Build");
    expect(readProject().tabs[2].id).toBe("tab_funnel"); // id kept on rename
  });

  test("edits the purpose in place and keeps the other rows", async () => {
    const purpose = (await row("Assumptions")).locator("textarea.tpurpose");
    await expect(purpose).toHaveValue("Inputs everything else reads");
    await purpose.fill("One input, one place");
    await purpose.press("Enter");
    await expect.poll(() => readProject().tabs.find((t: any) => t.id === "tab_assump").purpose).toBe("One input, one place");
    expect(readProject().tabs.find((t: any) => t.id === "tab_ramp").purpose).toBe("Five-year revenue build from share and price");
  });

  test("reorders by dragging the grip and the order survives a fresh show_plan", async () => {
    // Drag "Assumptions" (index 1) below "TAM Build" (index 2).
    await host.settle();
    const grip = (await row("Assumptions")).locator(".grip");
    const target = await row("TAM Build");
    const gb = (await grip.boundingBox())!;
    const tb = (await target.boundingBox())!;
    const dy = tb.y + tb.height * 0.9 - (gb.y + gb.height / 2);
    await dragBy(host, grip, 0, dy, 16);
    const want = ["Read Me", "TAM Build", "Assumptions", "Revenue Ramp", "Sources"];
    await expect.poll(values).toEqual(want);
    await expect.poll(() => readProject().tabs.map((t: any) => t.name)).toEqual(want);
    // fresh canvas from the server
    const again = await host.call("show_plan", { view: "tabs" });
    expect(await again.locator(".tablist .trow input.tname").evaluateAll((els) => els.map((e) => (e as HTMLInputElement).value))).toEqual(want);
    app = again;
  });

  test("deletes a tab with the two-step trash button", async () => {
    const r = await row("Sources");
    await r.hover(); // the action buttons only show on hover or selection
    const trash = r.locator(".acts button").nth(2);
    await trash.click();
    await expect(trash).toHaveText("Delete?");
    await trash.click();
    await expect(names()).toHaveCount(4);
    await expect.poll(() => readProject().tabs.map((t: any) => t.name)).toEqual(["Read Me", "TAM Build", "Assumptions", "Revenue Ramp"]);
  });

  test("adds a note with the textarea and Cmd/Ctrl+Enter", async () => {
    await (await row("Assumptions")).click();
    await expect(app.locator(".side .sh")).toContainText("Assumptions");
    await expect(app.locator(".side .note.resolved")).toHaveCount(1);
    const ta = app.locator(".side .addnote textarea");
    await ta.fill("Blue inputs only, no hard-coded numbers elsewhere");
    await ta.press("Control+Enter");
    await expect(app.locator(".side .note.open .nt")).toHaveText("Blue inputs only, no hard-coded numbers elsewhere");
    await expect((await row("Assumptions")).locator(".dot.warn")).toHaveCount(1);
    await expect.poll(() => readProject().notes.filter((n: any) => n.status === "open").length).toBe(1);
    await expect(app.locator(".ftr .msg")).toHaveText("4 tabs · 1 open note");
  });

  test("approves the tab list", async () => {
    await app.locator(".ftr button.primary", { hasText: "Approve tabs" }).click();
    await expect(app.locator(".ftr .msg")).toHaveText("Tabs approved");
    await expect(app.locator(".ftr button", { hasText: "Unapprove" })).toBeVisible();
    await expect.poll(() => readProject().tabsApproved).toBe(true);
    expect(readProject().tabs.every((t: any) => t.approved)).toBe(true);
    // the canvas asked the host to send a message so Claude continues
    await expect(host.page.getByText("Messages")).toBeVisible();
    // step 2 in the header becomes reachable
    await expect(app.locator(".hdr .seg button").nth(1)).toBeEnabled();
  });
});

test.describe.serial("deck outline", () => {
  let host: Host;
  let app: FrameLocator;
  screenshotOnFailure(() => host?.page);

  const rows = () => app.locator(".tablist .trow");
  const slides = () => app.locator(".tablist .trow:not(.div)");
  const dividers = () => app.locator(".tablist .trow.div input.dname");
  const dividerNames = () => dividers().evaluateAll((els) => els.map((e) => (e as HTMLInputElement).value));
  const plan = () => readProject(DECK_DIV_ID);
  const names = () => plan().tabs.map((t: any) => t.name);
  const STORY = "Contrasts our onboarding time with the two rivals taking share, in weeks and in lost accounts.";
  const AGENDA = "Where we are losing today, in one line.";

  test.beforeAll(async ({ browser }: { browser: Browser }) => {
    seed(deckDividerProject());
    host = await Host.open(browser);
    app = await host.call("show_plan", { view: "tabs" }, { project: DECK_DIV_ID });
  });
  test.afterAll(async () => { await host?.page.close(); });

  test("speaks deck in the chrome", async () => {
    await expect(app.locator(".hdr .chip.kind")).toHaveText("deck");
    await expect(app.locator(".hdr .seg button").first()).toHaveText("Slides");
    await expect(app.locator(".thead .grow")).toHaveText("Slide");
    await expect(rows()).toHaveCount(8);   // six slides and two dividers, one flat list
    await expect(slides()).toHaveCount(6);
    await expect(app.locator(".addrow button.trigger")).toHaveText("Add slide");
    await expect(app.locator(".addrow button.adddiv")).toHaveText("Add divider");
    await expect(app.locator(".ftr button.primary")).toContainText("Approve slides");
    await expect(app.locator(".ftr .msg")).toHaveText("6 slides · 2 sections · 1 open note");
  });

  test("numbers the slides only and says nothing else on a row", async () => {
    expect(await slides().locator(".idx").allInnerTexts()).toEqual(["1", "2", "3", "4", "5", "6"]);
    expect(await app.locator(".tablist .trow.div .idx").allInnerTexts()).toEqual(["", ""]);
    expect(await dividerNames()).toEqual(SECTIONS);
    // the section pill, the section label row and the message line are all gone
    await expect(app.locator(".tablist button.sect")).toHaveCount(0);
    await expect(app.locator(".tablist .srow")).toHaveCount(0);
    await expect(app.locator(".tablist input.tmessage")).toHaveCount(0);
  });

  test("shows the brief folded, with no emphasis behind Edit", async () => {
    await expect(app.locator(".brief .bl")).toContainText("For The partner group: We win SMB payments by owning onboarding");
    await expect(app.locator(".brief .dot.warn")).toHaveCount(0);
    await app.locator(".brief button.link", { hasText: "Edit" }).click();
    await expect(app.locator(".brief .field")).toHaveCount(3);
    await expect(app.locator(".brief .chips")).toHaveCount(0);
    await app.locator(".brief .bhead button.icon").click(); // fold it back for the rest of the pass
    await expect(app.locator(".brief .bl")).toBeVisible();
  });

  test("expands a long story with the more button", async () => {
    // only the story that overflows two lines offers the expander
    await expect(slides().nth(1).locator("button.more")).toBeVisible();
    await expect(slides().nth(2).locator("button.more")).toBeHidden();
    const tall = async () => (await slides().nth(1).locator("textarea.tstory").boundingBox())!.height;
    const before = await tall();
    await slides().nth(1).locator("button.more").click();
    await expect(slides().nth(1).locator("button.more")).toHaveText("less");
    expect(await tall()).toBeGreaterThan(before);
  });

  test("edits a story and it round-trips through the server", async () => {
    const story = slides().nth(2).locator("textarea.tstory");
    await story.fill(STORY);
    await story.press("Enter");
    await expect.poll(() => plan().tabs[3].story).toBe(STORY);
    const again = await host.call("show_plan", { view: "tabs" }, { project: DECK_DIV_ID });
    await expect(again.locator(".tablist .trow:not(.div)").nth(2).locator("textarea.tstory")).toHaveValue(STORY);
    app = again;
  });

  test("renames a divider in place and the slides under it follow", async () => {
    const name = dividers().nth(1);
    await name.fill("The moves");
    await name.press("Enter");
    await expect.poll(() => plan().tabs[4].name).toBe("The moves");
    expect(plan().tabs[4].divider).toBe(true);
    await expect.poll(() => plan().tabs[5].section).toBe("The moves");
    expect(plan().tabs[1].section).toBe(SECTIONS[0]); // the other section is untouched
    await expect(app.locator(".ftr .msg")).toContainText("2 sections");
  });

  test("edits the agenda line on a divider row", async () => {
    const lines = () => app.locator(".tablist .trow.div textarea.dagenda");
    await expect(lines()).toHaveCount(2);
    await expect(app.locator(".tablist .trow:not(.div) textarea.dagenda")).toHaveCount(0); // a slide carries none
    await expect(lines().first()).toHaveAttribute("placeholder", "Agenda line");
    await expect(lines().first()).toHaveValue("");
    await lines().first().fill(AGENDA);
    await lines().first().press("Enter");
    await expect.poll(() => plan().tabs[0].agendaLine).toBe(AGENDA);
    expect(plan().tabs[0].divider).toBe(true);
    expect(plan().tabs[4].agendaLine).toBe(undefined);   // the other divider keeps none
    expect(plan().tabs[1].story).toBeTruthy();           // and the slides under it are untouched
    // a fresh canvas from the server draws it back
    const again = await host.call("show_plan", { view: "tabs" }, { project: DECK_DIV_ID });
    await expect(again.locator(".tablist .trow.div textarea.dagenda").first()).toHaveValue(AGENDA);
    app = again;
  });

  test("reorders with one flat drag, no placeholder row", async () => {
    await host.settle();
    const grip = slides().nth(0).locator(".grip");
    const target = slides().nth(1);
    const gb = (await grip.boundingBox())!;
    const tb = (await target.boundingBox())!;
    await dragBy(host, grip, 0, tb.y + tb.height * 0.9 - (gb.y + gb.height / 2), 16);
    const want = [SECTIONS[0], SLIDE_NAMES[1], SLIDE_NAMES[0], SLIDE_NAMES[2], "The moves", ...SLIDE_NAMES.slice(3)];
    await expect.poll(names).toEqual(want);
    await expect(app.locator(".tablist .trow.ph")).toHaveCount(0);
    await expect(app.locator(".tablist .insline")).toHaveCount(1);
  });

  test("Cmd+Down moves the selection to the next row", async () => {
    await rows().nth(1).click();
    await expect(app.locator(".trow.sel")).toHaveAttribute("data-id", plan().tabs[1].id);
    await app.locator(".tablist").press("Meta+ArrowDown");
    await expect(app.locator(".trow.sel")).toHaveAttribute("data-id", plan().tabs[2].id);
    await expect(app.locator(".side .sh .t")).toHaveText(plan().tabs[2].name);
    await app.locator(".tablist").press("Meta+ArrowUp");
    await expect(app.locator(".trow.sel")).toHaveAttribute("data-id", plan().tabs[1].id);
  });

  test("Alt+Down moves the selected row, Alt+Up puts it back", async () => {
    const order = () => names().slice(1, 3);
    const before = order();
    await rows().nth(1).click();
    await app.locator(".tablist").press("Alt+ArrowDown");
    await expect.poll(order).toEqual([before[1], before[0]]);
    await app.locator(".tablist").press("Alt+ArrowUp");
    await expect.poll(order).toEqual(before);
  });

  test("adds a divider after the selected row and puts the caret in its name", async () => {
    await rows().nth(2).click();
    await app.locator(".addrow button.adddiv").click();
    await expect(rows()).toHaveCount(9);
    await expect.poll(() => plan().tabs[3].divider).toBe(true);
    expect(plan().tabs[3].name).toBe("Section");
    await expect(dividers().nth(1)).toBeFocused();
    await expect(app.locator(".ftr .msg")).toContainText("3 sections");
  });

  test("merges two dividers with the same name, keeping the first", async () => {
    const added = dividers().nth(1);
    await added.fill(SECTIONS[0]);        // the same name as the divider at the top, but not next to it yet
    await added.press("Enter");
    await expect.poll(() => plan().tabs[3].name).toBe(SECTIONS[0]);
    await expect(rows()).toHaveCount(9);
    await rows().nth(3).click();
    await app.locator(".tablist").press("Alt+ArrowUp");   // index 3 to 2: still between slides
    await expect.poll(() => plan().tabs[2].divider).toBe(true);
    await app.locator(".tablist").press("Alt+ArrowUp");   // index 2 to 1: now under the first divider, so it merges
    await expect(rows()).toHaveCount(8);
    expect(await dividerNames()).toEqual([SECTIONS[0], "The moves"]);
    expect(plan().tabs[0].id).toBe(DIVIDER_IDS.today);    // the client merged before the save, so the first survived
    expect(plan().tabs.filter((t: any) => t.divider).length).toBe(2);
  });

  test("deletes a divider from its row actions", async () => {
    const first = rows().nth(0);
    await first.hover();
    const trash = first.locator(".acts button").nth(2);
    await trash.click();
    await trash.click();
    await expect(rows()).toHaveCount(7);
    await expect.poll(() => plan().tabs.filter((t: any) => t.divider).length).toBe(1);
  });

  test("a new divider never lands next to another one", async () => {
    const was = plan().tabs.findIndex((t: any) => t.divider);
    await rows().nth(was).click();               // the selected row is itself a divider
    await app.locator(".addrow button.adddiv").click();
    await expect(rows()).toHaveCount(8);
    const at = plan().tabs.map((t: any) => !!t.divider);
    expect(at.filter(Boolean).length).toBe(2);
    expect(at.findIndex((d: boolean, i: number) => d && at[i + 1])).toBe(-1); // a row sits between them
  });
});

test.describe.serial("outline scroll and focus", () => {
  let host: Host;
  let app: FrameLocator;
  screenshotOnFailure(() => host?.page);

  test.beforeAll(async ({ browser }: { browser: Browser }) => {
    seed(longDeckProject());
    host = await Host.open(browser);
    app = await host.call("show_plan", { view: "tabs" }, { project: DECK_LONG_ID });
    await host.settle();
  });
  test.afterAll(async () => { await host?.page.close(); });

  test("keeps the list scroll and the caret when Cmd+Enter files a note", async () => {
    const list = app.locator(".tablist");
    await app.locator(".tablist .trow:not(.div)").nth(4).click();
    const at = await list.evaluate((el) => { el.scrollTop = 160; return el.scrollTop; });
    expect(at, "the seeded outline is long enough to scroll").toBeGreaterThan(0);
    const ta = app.locator(".side .addnote textarea");
    await ta.fill("Name the two rivals on this slide");
    await ta.press("Control+Enter");
    await expect(app.locator(".side .note.open .nt")).toHaveText("Name the two rivals on this slide");
    expect(await list.evaluate((el) => el.scrollTop)).toBe(at);
    await expect(ta).toBeFocused();
  });
});

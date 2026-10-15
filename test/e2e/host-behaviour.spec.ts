import { test, expect, type Browser } from "@playwright/test";
import { Host, screenshotOnFailure } from "./host";
import { demoProject, seed } from "./seed";

test.describe.serial("host behaviour", () => {
  let host: Host;
  screenshotOnFailure(() => host?.page);

  test.beforeAll(async ({ browser }: { browser: Browser }) => {
    seed(demoProject());
    host = await Host.open(browser);
  });
  test.afterAll(async () => { await host?.page.close(); });

  test("an older canvas collapses when a newer one renders, and Keep using this one restores it", async () => {
    const first = await host.call("show_plan", { view: "tabs" });
    await expect(first.locator(".tablist .trow")).toHaveCount(4);
    const second = await host.call("show_plan", { view: "tabs" });
    await expect(second.locator(".tablist .trow")).toHaveCount(4);
    // The first canvas polls ui_get_state every 2.5 s and sees a higher renderSeq.
    const bar = first.locator(".superseded");
    await expect(bar).toBeVisible({ timeout: 6000 });
    await expect(bar).toContainText("Demo Market Sizing");
    await expect(bar).toContainText("Newer canvas below");
    await expect(first.locator("#root")).toHaveClass(/collapsed/);
    await expect(first.locator(".tablist")).toHaveCount(0);
    // the newer canvas is untouched
    await expect(second.locator(".superseded")).toHaveCount(0);
    await expect(second.locator(".tablist .trow")).toHaveCount(4);
    // pin the old one (after the host has finished animating the iframe down to the bar's height)
    await host.settle();
    await bar.locator("button", { hasText: "Keep this one" }).click();
    await expect(first.locator(".superseded")).toHaveCount(0);
    await expect(first.locator("#root")).not.toHaveClass(/collapsed/);
    await expect(first.locator(".tablist .trow")).toHaveCount(4);
    // and it stays pinned through the next poll
    await host.page.waitForTimeout(3000);
    await expect(first.locator(".superseded")).toHaveCount(0);
  });

  test("the host theme toggle switches every canvas to dark", async () => {
    // Playwright's default colour scheme is light, so the host starts light.
    await expect(host.themeToggle).toHaveAttribute("title", "Switch to dark mode");
    await host.themeToggle.click();
    await expect(host.themeToggle).toHaveAttribute("title", "Switch to light mode");
    await expect(host.app(1).locator("html")).toHaveAttribute("data-theme", "dark");
    await expect(host.app(0).locator("html")).toHaveAttribute("data-theme", "dark");
  });

  test("the light theme survives a container resize", async () => {
    // Known bug (src/app/controller.ts, applyHost): the host-context-changed handler is given only the
    // changed fields, and applyHost reads ctx.theme from that partial object, so a containerDimensions-only
    // notification (the host sends one whenever the iframe resizes) flips a light canvas back to dark.
    // Expected to fail until applyHost reads app.getHostContext() instead.
    await host.themeToggle.click();
    await expect(host.themeToggle).toHaveAttribute("title", "Switch to dark mode");
    await expect(host.app(1).locator("html")).toHaveAttribute("data-theme", "light");
    await host.page.setViewportSize({ width: 1300, height: 1000 });
    await host.settle();
    await expect(host.app(1).locator("html")).toHaveAttribute("data-theme", "light", { timeout: 2000 });
    await expect(host.app(0).locator("html")).toHaveAttribute("data-theme", "light", { timeout: 2000 });
  });

  /**
   * P5. The vendored basic-host has an AppBridge that can call `tools/list` and `tools/call` on the app,
   * but no UI that does, so this drives the App's own handler and registry inside the frame: exactly what
   * a host reaches. `onlisttools` also proves the schemas serialise, which is the part that can break.
   */
  test("the canvas registers four app tools, schemas and all", async () => {
    const app = await host.call("show_plan", { view: "tabs" });
    const listed = await app.locator("html").evaluate(async () => {
      const r = await (window as any).blueprint.app.onlisttools({}, {});
      return (r.tools as any[]).map((t) => [t.name, Object.keys(t.inputSchema?.properties ?? {}).sort().join(" "), (t.inputSchema?.required ?? []).join(" ")]);
    });
    expect(listed).toEqual([
      ["focus_slide", "name tabId", ""],
      ["select_block", "blockId", "blockId"],
      ["scroll_to", "blockId tabId", "tabId"],
      ["show_view", "view", "view"],
    ]);
    const views = await app.locator("html").evaluate(async () => {
      const r = await (window as any).blueprint.app.onlisttools({}, {});
      return (r.tools as any[]).find((t) => t.name === "show_view").inputSchema.properties.view.enum;
    });
    expect(views).toEqual(["outline", "layouts", "built"]);
  });

  test("an app tool moves the canvas in place, without a render from the server", async () => {
    const app = host.app((await host.iframes.count()) - 1);
    const before = await host.iframes.count();
    const call = (name: string, args: Record<string, unknown>) =>
      app.locator("html").evaluate((_el, [n, a]: [string, Record<string, unknown>]) => (window as any).blueprint.callAppTool(n, a).then((r: any) => r.content[0].text as string), [name, args] as [string, Record<string, unknown>]);

    // the plan is not approved, so the layouts are not open: one line back, canvas untouched
    expect(await call("show_view", { view: "layouts" })).toBe("Outline not approved");
    await expect(app.locator(".tablist .trow")).toHaveCount(4);

    // focus_slide takes the name a user would read, and the row it names becomes the selected one
    expect(await call("focus_slide", { name: "Revenue Ramp" })).toBe("Focused Revenue Ramp");
    await expect(app.locator(".tablist .trow.sel")).toHaveCount(1);
    await expect(app.locator(".tablist .trow.sel .tname")).toHaveValue("Revenue Ramp");
    expect(await call("scroll_to", { tabId: "tab_assump" })).toBe("Scrolled to Assumptions");
    await expect(app.locator(".tablist .trow.sel .tname")).toHaveValue("Assumptions");

    // a row or a block nobody has is a refusal, not a throw, and the selection stays where it was
    expect(await call("focus_slide", { name: "Nope" })).toBe('No tab "Nope"');
    expect(await call("select_block", { blockId: "blk_nope" })).toBe("No block blk_nope");
    await expect(app.locator(".tablist .trow.sel .tname")).toHaveValue("Assumptions");

    // nothing here went through the server, so no new canvas was rendered
    expect(await host.iframes.count()).toBe(before);
  });
});

import { expect, test, type Browser, type FrameLocator, type Locator, type Page, type TestInfo } from "@playwright/test";
import { HOST_URL } from "./playwright.config";
import { DEMO_ID } from "./seed";

/**
 * Drives the vendored ext-apps basic-host (test/host). The page has a Server select, a Tool select,
 * a JSON textarea and a "Call Tool" button; every app tool call appends a panel whose <iframe> is the
 * sandbox (port 8081), which wraps a second <iframe> holding the planner canvas.
 */
export class Host {
  constructor(readonly page: Page) {}

  static async open(browser: Browser): Promise<Host> {
    const page = await browser.newPage();
    const host = new Host(page);
    await page.goto(HOST_URL);
    // Server list is fetched then each server connected; the Tool select fills once that resolves.
    await expect(host.toolSelect.locator("option").first()).toBeAttached({ timeout: 15_000 });
    return host;
  }

  get toolSelect(): Locator { return this.page.locator('select[class*="toolSelect"]'); }
  get toolInput(): Locator { return this.page.locator('textarea[class*="toolInput"]'); }
  get themeToggle(): Locator { return this.page.locator('button[class*="themeToggle"]'); }
  /** All sandbox iframes on the page, one per app tool call, in call order. */
  get iframes(): Locator { return this.page.locator("iframe"); }

  /** The planner canvas inside the i-th app panel (0-based). */
  app(i: number): FrameLocator { return this.page.frameLocator("iframe").nth(i).frameLocator("iframe"); }

  /** Call an app tool through the form and return the canvas it rendered, once the canvas has drawn its header. */
  async call(tool: string, args: Record<string, unknown>, opts: { project?: string } = {}): Promise<FrameLocator> {
    const before = await this.iframes.count();
    await this.toolSelect.selectOption(tool);
    await this.toolInput.fill(JSON.stringify({ project: opts.project ?? DEMO_ID, ...args }));
    await this.page.getByRole("button", { name: "Call Tool" }).click();
    await expect(this.iframes).toHaveCount(before + 1, { timeout: 15_000 });
    const app = this.app(before);
    await expect(app.locator(".hdr .name")).toBeVisible({ timeout: 15_000 });
    return app;
  }

  /** Call a tool that has no UI (update_plan, get_plan, resolve_notes). */
  async callPlain(tool: string, args: Record<string, unknown>, opts: { project?: string } = {}): Promise<void> {
    const panels = this.page.locator('div[class*="toolCallInfoPanel"]');
    const before = await panels.count();
    await this.toolSelect.selectOption(tool);
    await this.toolInput.fill(JSON.stringify({ project: opts.project ?? DEMO_ID, ...args }));
    await this.page.getByRole("button", { name: "Call Tool" }).click();
    await expect(panels).toHaveCount(before + 1);
    await expect(panels.nth(before).getByText("Tool Result")).toBeVisible();
  }

  /**
   * Wait until no app iframe is changing size. The host animates the iframe height for 300 ms after
   * every size-changed notification and sends a host-context-changed on every frame of it; the canvas
   * re-renders on each one, which drops any pointer drag in progress (see the bug list in the report).
   */
  async settle(quietMs = 400, timeoutMs = 6000): Promise<void> {
    const sizes = async () => JSON.stringify(await this.iframes.evaluateAll((els) => els.map((e) => { const r = e.getBoundingClientRect(); return [Math.round(r.width), Math.round(r.height)]; })));
    const start = Date.now();
    let last = await sizes();
    let quietSince = Date.now();
    while (Date.now() - start < timeoutMs) {
      await this.page.waitForTimeout(100);
      const now = await sizes();
      if (now !== last) { last = now; quietSince = Date.now(); }
      else if (Date.now() - quietSince >= quietMs) return;
    }
  }

  /** Reload the host page so later calls start from a clean panel list. */
  async reset(): Promise<void> {
    await this.page.goto(HOST_URL);
    await expect(this.toolSelect.locator("option").first()).toBeAttached({ timeout: 15_000 });
  }
}

/**
 * Drag from the centre of `from` to the point (dx, dy) away, in steps. First with real mouse events
 * (the canvas uses pointer events plus setPointerCapture on the handle). Inside the host's nested
 * iframes, headless Chromium ignores setPointerCapture in roughly half of all browser sessions: the call
 * returns, hasPointerCapture is false straight away, and it stays that way for every press on that page
 * (a plain page never does this, so it is the cross-frame mouse routing, not the app). When that happens
 * the handle's own up-handler is fired to undo the press, and the drag is replayed as pointer events
 * dispatched at the handle inside the frame, which is where the canvas listens anyway.
 */
export async function dragBy(host: Host, from: Locator, dx: number, dy: number, steps = 12): Promise<void> {
  const page = host.page;
  await host.settle();
  await from.scrollIntoViewIfNeeded();
  const box = (await from.boundingBox())!;
  const x = box.x + box.width / 2, y = box.y + box.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x + 2, y + 2); // the first move after the press is when Chromium applies the capture
  const captured = await from.evaluate((el) => el.hasPointerCapture(1)); // Playwright's mouse is pointerId 1
  if (captured) {
    for (let i = 1; i <= steps; i++) await page.mouse.move(x + (dx * i) / steps, y + (dy * i) / steps);
    await page.mouse.up();
    return;
  }
  await from.evaluate((el) => el.dispatchEvent(new PointerEvent("pointerup", { pointerId: 1, bubbles: true })));
  await page.mouse.up();
  await host.settle(200);
  console.log("dragBy: Chromium ignored setPointerCapture on this page; replaying the drag as in-frame pointer events");
  await from.evaluate((el, o) => {
    const r = el.getBoundingClientRect();
    const x0 = r.left + r.width / 2, y0 = r.top + r.height / 2;
    const fire = (type: string, cx: number, cy: number) => el.dispatchEvent(new PointerEvent(type, { pointerId: 1, pointerType: "mouse", isPrimary: true, button: type === "pointermove" ? -1 : 0, buttons: type === "pointerup" ? 0 : 1, clientX: cx, clientY: cy, bubbles: true, cancelable: true, composed: true }));
    fire("pointerdown", x0, y0);
    fire("pointermove", x0 + 2, y0 + 2);
    for (let i = 1; i <= o.steps; i++) fire("pointermove", x0 + (o.dx * i) / o.steps, y0 + (o.dy * i) / o.steps);
    fire("pointerup", x0 + o.dx, y0 + o.dy);
  }, { dx, dy, steps });
}

/** Screenshot the whole host page when a test fails; the shared page from beforeAll is not the `page` fixture, so do it by hand. */
export function screenshotOnFailure(getPage: () => Page | undefined) {
  test.afterEach(async ({}, testInfo: TestInfo) => {
    const page = getPage();
    if (!page || page.isClosed() || testInfo.status === testInfo.expectedStatus) return;
    const file = testInfo.outputPath("failure.png");
    await page.screenshot({ path: file, fullPage: true }).catch(() => {});
    testInfo.attachments.push({ name: "failure", path: file, contentType: "image/png" });
  });
}

import { defineConfig } from "@playwright/test";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
export const ROOT = path.resolve(here, "..", "..");
export const ARTIFACTS = path.join(here, "artifacts");
/** Planner home for the run: wiped and re-seeded by global-setup, rewritten by each spec. */
export const PLANNER_HOME = path.join(ARTIFACTS, "home");
// Ports are deliberately off the dev defaults (3001/8080/8081) so a dev host or a second suite can run alongside.
export const PLANNER_PORT = 13001;
export const HOST_PORT = 18080;
export const SANDBOX_PORT = 18081; // hard-coded in the vendored test/host/index.html (patched from 8081)
export const HOST_URL = `http://localhost:${HOST_PORT}/index.html`;

export default defineConfig({
  testDir: here,
  testMatch: /.*\.spec\.ts/,
  // Playwright wipes outputDir on every run; keep it in a subfolder so artifacts/visual survives concurrent runs.
  outputDir: path.join(ARTIFACTS, "runs"),
  globalSetup: path.join(here, "global-setup.ts"),
  fullyParallel: false,
  workers: 1,
  retries: 0,
  timeout: 30_000,
  expect: { timeout: 8_000 },
  reporter: [["list"]],
  use: {
    baseURL: HOST_URL,
    headless: true,
    channel: "chromium",
    launchOptions: { args: process.env.PW_ARGS ? process.env.PW_ARGS.split(" ") : [] },
    viewport: { width: 1400, height: 1000 },
    screenshot: "only-on-failure",
    trace: "off",
  },
  webServer: [
    {
      command: `node dist/server/main.js --http ${PLANNER_PORT}`,
      cwd: ROOT,
      port: PLANNER_PORT,
      reuseExistingServer: false,
      timeout: 20_000,
      // The skills dir is inside the run's own home, so a spec that saves a style never touches ~/.claude/skills.
      // global-setup wipes the home after this server starts, so the spec that needs the dir creates it.
      env: { WORKBOOK_PLANNER_HOME: PLANNER_HOME, WORKBOOK_PLANNER_OUTPUT: path.join(PLANNER_HOME, "out"), BLUEPRINT_SKILLS: path.join(PLANNER_HOME, "skills"), BLUEPRINT_NO_SHEET_IMAGES: "1" /* the canvas never sees sheet images; skip the render per build */ },
    },
    {
      command: `npx tsx test/host/serve.ts`,
      cwd: ROOT,
      url: HOST_URL,
      reuseExistingServer: false,
      timeout: 30_000,
      env: { SERVERS: JSON.stringify([`http://localhost:${PLANNER_PORT}/mcp`]), HOST_PORT: String(HOST_PORT), SANDBOX_PORT: String(SANDBOX_PORT) },
    },
  ],
});

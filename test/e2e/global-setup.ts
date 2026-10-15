import fs from "node:fs";
import path from "node:path";
import { PLANNER_HOME, ROOT } from "./playwright.config";
import { seed } from "./seed";

/** Wipe the planner home and seed the demo project. Runs once before the suite (the web servers may already be up; they read from disk on each call). */
export default async function globalSetup() {
  for (const f of ["dist/mcp-app.html", "dist/server/main.js"]) {
    if (!fs.existsSync(path.join(ROOT, f))) throw new Error(`${f} is missing. Run \`npm run build\` before \`npm run test:e2e\`.`);
  }
  fs.rmSync(PLANNER_HOME, { recursive: true, force: true });
  seed();
}

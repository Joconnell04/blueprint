import fs from "node:fs";
import { McpServer } from "@modelcontextprotocol/server";
import { registerAppResource, RESOURCE_MIME_TYPE } from "@modelcontextprotocol/ext-apps/server";
import { HOME, OUTPUT_DIR } from "./store.js";
import { APP_HTML, RESOURCE_URI, VERSION, appOnly, render, runPython, ui } from "./common.js";
import { registerPlanTools } from "./tools-plan.js";
import { registerImportTools } from "./tools-import.js";
import { registerImportDeckTool } from "./tools-import-deck.js";
import { registerExtractTool } from "./tools-extract.js";
import { registerBuildTools } from "./tools-build.js";
import { registerBuildDeckTool } from "./tools-build-deck.js";
import { registerReviewTool } from "./tools-review.js";
import { registerUiTools } from "./tools-ui.js";
import { registerFileTools } from "./tools-files.js";
import { registerTemplateTools } from "./templates.js";
import { registerRestyleTool } from "./restyle.js";

export { isReadMe } from "./common.js";

/** The MCP server: one tool group per module, one HTML resource for the canvas. */
export function createServer(): McpServer {
  const server = new McpServer({ name: "Blueprint", version: VERSION });
  registerPlanTools(server);
  registerImportTools(server);
  registerImportDeckTool(server);
  registerExtractTool(server);
  registerBuildTools(server);
  registerBuildDeckTool(server);
  registerReviewTool(server);
  registerUiTools(server);
  registerFileTools(server);
  registerTemplateTools(server, ui, appOnly, render);
  registerRestyleTool(server, ui, appOnly, render, runPython, HOME, OUTPUT_DIR);

  registerAppResource(server, "Blueprint canvas", RESOURCE_URI, { mimeType: RESOURCE_MIME_TYPE }, async () => {
    const html = fs.readFileSync(APP_HTML, "utf8");
    return { contents: [{ uri: RESOURCE_URI, mimeType: RESOURCE_MIME_TYPE, text: html, _meta: { ui: { prefersBorder: true, permissions: { clipboardWrite: {} } } } }] };
  });
  return server;
}

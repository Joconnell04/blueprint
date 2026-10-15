/** import_workbook over stdio: scan the builder's sample house workbook into a plan. Run: npx tsx test/import.test.ts */
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const home = fs.mkdtempSync(path.join(os.tmpdir(), "wp-import-"));
const sample = path.resolve("builder/sample_house.xlsx");
if (!fs.existsSync(sample)) { const { execSync } = await import("node:child_process"); execSync("python3 builder/build_xlsx.py builder/sample_spec.json"); }
const client = new Client({ name: "import-test", version: "0" });
await client.connect(new StdioClientTransport({ command: "node", args: ["dist/server/main.js"], env: { ...process.env, WORKBOOK_PLANNER_HOME: home } }));
const r: any = await client.callTool({ name: "import_workbook", arguments: { path: sample, formatting: "house" } });
assert.ok(!r.isError, r.content?.[0]?.text);
const sc = r.structuredContent as any;
assert.equal(sc.view, "layouts");
const p = sc.project;
assert.ok(p.tabsApproved && p.importedFrom === sample);
assert.ok(p.tabs.length >= 4, `tabs: ${p.tabs.map((t: any) => t.name)}`);
const blocks = p.tabs.flatMap((t: any) => t.blocks);
assert.ok(blocks.length >= 10, `blocks: ${blocks.length}`);
assert.ok(blocks.every((b: any) => b.detected && typeof b.detected.confidence === "number" && b.row >= 1 && b.col >= 1));
assert.ok(blocks.some((b: any) => b.type === "chart" && b.source), "charts resolve to a source block");
assert.ok(blocks.some((b: any) => b.type === "inputs"), "an inputs block was detected");
const text = r.content[0].text as string;
assert.ok(/Imported \d+ sheet/.test(text));
await client.close();
console.log("IMPORT OK", p.tabs.map((t: any) => `${t.name}:${t.blocks.length}`).join(" "));

// ---- restyle in place: the imported workbook comes back reformatted, values untouched
{
  const c2 = new Client({ name: "import-test-2", version: "0" });
  await c2.connect(new StdioClientTransport({ command: "node", args: ["dist/server/main.js"], env: { ...process.env, WORKBOOK_PLANNER_HOME: home, WORKBOOK_PLANNER_OUTPUT: path.join(home, "out") } }));
  const rr: any = await c2.callTool({ name: "restyle_workbook", arguments: { project: p.id, formatting: "plain" } });
  assert.ok(!rr.isError, rr.content?.[0]?.text);
  const p2 = (rr.structuredContent as any).project;
  assert.ok(p2.buildPath && fs.existsSync(p2.buildPath) && p2.buildPath !== sample, p2.buildPath);
  await c2.close();
  console.log("RESTYLE OK", path.basename(p2.buildPath));
}

// ---- rebuild path: get_plan on an imported project carries the original content with translated formulas
{
  const c3 = new Client({ name: "import-test-3", version: "0" });
  await c3.connect(new StdioClientTransport({ command: "node", args: ["dist/server/main.js"], env: { ...process.env, WORKBOOK_PLANNER_HOME: home } }));
  const g: any = await c3.callTool({ name: "get_plan", arguments: { project: p.id } });
  const text = g.content[0].text as string;
  assert.ok(/original content from the imported file/.test(text) && /\{\{imp_\d+_\d+!/.test(text), text.slice(-600));
  await c3.close();
  console.log("REBUILD CONTENT OK");
}

// ---- U6: the notes the sheet already carried come back on the plan, split by the row each sat on.
// The scanner counts the header and total rows, so a naive read puts every note one row out.
{
  const { execFileSync } = await import("node:child_process");
  const spec = JSON.parse(fs.readFileSync("builder/sample_spec.json", "utf8"));
  const tam = spec.sheets.flatMap((sh: any) => sh.blocks).find((b: any) => b.id === "b_tam");
  tam.rowNotes = ["Licensed firmographic file", null, "Sales-qualified only", null];
  tam.totalNote = "Ties to the funnel total";
  spec.outputPath = path.join(home, "notes.xlsx");
  fs.writeFileSync(path.join(home, "notes-spec.json"), JSON.stringify(spec));
  const built = JSON.parse(execFileSync("python3", ["builder/build_xlsx.py", path.join(home, "notes-spec.json")]).toString().trim().split("\n").pop()!);
  assert.ok(built.ok, JSON.stringify(built.errors));

  const c4 = new Client({ name: "import-test-4", version: "0" });
  await c4.connect(new StdioClientTransport({ command: "node", args: ["dist/server/main.js"], env: { ...process.env, WORKBOOK_PLANNER_HOME: home } }));
  const imp: any = await c4.callTool({ name: "import_workbook", arguments: { path: spec.outputPath, formatting: "house" } });
  assert.ok(!imp.isError, imp.content?.[0]?.text);
  const np = (imp.structuredContent as any).project;

  // the header note says what the block is for, so it lands on the block as its description
  const table = np.tabs.find((t: any) => t.name === "TAM Build").blocks.find((b: any) => b.type === "table");
  assert.equal(table.description, "Accounts from the licensed firmographic file", JSON.stringify(table));
  const inputs = np.tabs.find((t: any) => t.name === "Assumptions").blocks.find((b: any) => b.type === "inputs");
  assert.equal(inputs.description, "Refresh when the SBA profile updates");

  const g: any = await c4.callTool({ name: "get_plan", arguments: { project: np.id } });
  const plan = g.content[0].text as string;
  const line = plan.split("\n").find((l: string) => l.startsWith(`${table.id} (TAM Build)`));
  assert.ok(line, plan.slice(-900));
  assert.ok(line!.includes('rowNotes ["Licensed firmographic file",null,"Sales-qualified only",null]'), line!);
  assert.ok(line!.includes('totalNote "Ties to the funnel total"'), line!);
  // an inputs block has no header or total row, so its notes map straight onto its rows
  const inLine = plan.split("\n").find((l: string) => l.startsWith(`${inputs.id} (Assumptions)`));
  assert.ok(inLine!.includes('rowNotes ["The model starts here","SBA 2024 small business profile","Firms with 5 to 200 employees",null]'), inLine!);
  await c4.close();
  console.log("ROW NOTES OK");
}

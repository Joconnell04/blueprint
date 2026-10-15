/** import_deck over stdio: scan the generated fixture decks into a plan. Run: npx tsx test/import-deck.test.ts */
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import assert from "node:assert/strict";
import { execSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const home = fs.mkdtempSync(path.join(os.tmpdir(), "bp-import-deck-"));
const fixtures = fs.mkdtempSync(path.join(os.tmpdir(), "bp-deck-fixtures-"));
execSync(`python3 builder/fixtures-deck/make_fixtures.py "${fixtures}"`, { stdio: "pipe" });
const sample = path.join(fixtures, "harbor_deck.pptx");
assert.ok(fs.existsSync(sample), `fixture missing: ${sample}`);

const connect = async (tag: string) => {
  const c = new Client({ name: tag, version: "0" });
  await c.connect(new StdioClientTransport({ command: "node", args: ["dist/server/main.js"], env: { ...process.env, BLUEPRINT_HOME: home } }));
  return c;
};

const client = await connect("import-deck-test");
const r: any = await client.callTool({ name: "import_deck", arguments: { path: sample } });
assert.ok(!r.isError, r.content?.[0]?.text);
const sc = r.structuredContent as any;
assert.equal(sc.view, "layouts");
const p = sc.project;

// ---- the project is a deck, imported, and nothing is approved
assert.equal(p.kind, "deck");
assert.equal(p.importedFrom, sample);
assert.ok(p.scanPath && fs.existsSync(p.scanPath), p.scanPath);
assert.equal(p.tabs.length, 6, `slides: ${p.tabs.map((t: any) => t.name)}`);
assert.ok(!p.tabsApproved && p.tabs.every((t: any) => !t.approved && !t.layoutApproved), "nothing is approved on import");

// ---- outline rows: headline on message, kicker on purpose, story left for the user
const slide2 = p.tabs[1];
assert.equal(slide2.name, "Referrals stall at the ask");
assert.equal(slide2.message, "Referrals stall at the ask, not at the intro");
assert.equal(slide2.purpose, "Kicker: Growth | Referrals");
assert.equal(slide2.story, "");
assert.equal(p.tabs[2].purpose, "Section divider");

// ---- the divider names the section and the next slide inherits it
assert.equal(p.tabs[2].section, "Where referrals stall");
assert.equal(p.tabs[3].section, "Where referrals stall");
assert.equal(p.tabs[2].blocks.length, 0, "a divider carries no blocks");
assert.deepEqual(p.tabs.filter((t: any) => t.purpose === "Section divider").map((t: any) => t.section), ["Where referrals stall", "How we fix it", "What it costs"]);

// ---- blocks: detected, on the 32x18 grid, furniture excluded
const blocks = p.tabs.flatMap((t: any) => t.blocks);
assert.ok(blocks.length >= 4, `blocks: ${blocks.length}`);
assert.ok(blocks.every((b: any) => b.detected && typeof b.detected.confidence === "number" && b.detected.signals.length > 0 && b.detected.range),
  "every block carries detected confidence, signals and a range");
assert.ok(blocks.every((b: any) => b.row >= 1 && b.col >= 1 && b.row + b.rows - 1 <= 18 && b.col + b.cols - 1 <= 32), "geometry stays on the slide grid");
assert.ok(!blocks.some((b: any) => b.type === "headline" || b.type === "kicker"), "headline and kicker stay furniture");
const cards = blocks.find((b: any) => b.type === "cards");
assert.ok(cards && cards.count === 3 && /3 cards: Open networks/.test(cards.content), JSON.stringify(cards));
assert.equal(`${cards.row},${cards.col},${cards.rows},${cards.cols}`, "7,3,10,28");
assert.ok(blocks.some((b: any) => b.type === "takeaway"), "the bottom bar is a takeaway");
assert.ok(blocks.some((b: any) => b.type === "notes"), "the source line is a notes block");

const text = r.content[0].text as string;
assert.ok(/Imported 6 slide\(s\), \d+ block\(s\)/.test(text), text);
assert.ok(/Where referrals stall/.test(text), text);
assert.ok(text.split("\n").length <= 8, `result text must stay under 8 lines:\n${text}`);
await client.close();
console.log("IMPORT DECK OK", p.tabs.map((t: any) => `${t.name}:${t.blocks.length}`).join(" | "));

// ---- list_projects shows it as a deck, get_plan carries the detections and sections
{
  const c2 = await connect("import-deck-test-2");
  const l: any = await c2.callTool({ name: "list_projects", arguments: {} });
  assert.ok(new RegExp(`${p.id}: .*\\[deck, `).test(l.content[0].text as string), l.content[0].text);

  const g: any = await c2.callTool({ name: "get_plan", arguments: { project: p.id } });
  const gt = g.content[0].text as string;
  assert.ok(/kind: deck/.test(gt), gt.slice(0, 200));
  assert.ok(/\[Where referrals stall\]/.test(gt), "sections show on the outline");
  assert.ok(/\[cards\] Cards .* at row 7, col 3, 10 rows x 28 cols/.test(gt), gt);
  assert.ok(/detected as text at \d+% confidence/.test(gt), "low-confidence detections reach the readiness warnings");
  const gp = (g.structuredContent as any).project;
  assert.ok(gp.tabs.flatMap((t: any) => t.blocks).every((b: any) => b.detected), "detections survive a reload");
  await c2.close();
  console.log("DECK PLAN OK");
}

// ---- the plain deck: chart, table, image, stat, quote and placeholder all survive the round trip
{
  const c3 = await connect("import-deck-test-3");
  const plain: any = await c3.callTool({ name: "import_deck", arguments: { path: path.join(fixtures, "plain_deck.pptx"), name: "Plain deck" } });
  assert.ok(!plain.isError, plain.content?.[0]?.text);
  const pp = (plain.structuredContent as any).project;
  assert.equal(pp.kind, "deck");
  assert.equal(pp.tabs.length, 6);
  const kinds = new Set(pp.tabs.flatMap((t: any) => t.blocks).map((b: any) => b.type));
  for (const k of ["bullets", "chart", "table", "image", "stat", "quote", "placeholder"]) assert.ok(kinds.has(k), `missing ${k}; got ${[...kinds]}`);
  const chart = pp.tabs.flatMap((t: any) => t.blocks).find((b: any) => b.type === "chart");
  assert.equal(chart.chartKind, "bar");
  const lowest = Math.min(...pp.tabs.flatMap((t: any) => t.blocks).map((b: any) => b.detected.confidence));
  assert.ok(lowest < 0.6, "ambiguous shapes come in under 60%");
  assert.ok(/Low confidence/.test(plain.content[0].text as string), plain.content[0].text);
  await c3.close();
  console.log("PLAIN DECK OK", [...kinds].join(" "));
}

fs.rmSync(fixtures, { recursive: true, force: true });

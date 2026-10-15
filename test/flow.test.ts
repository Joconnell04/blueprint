/**
 * End-to-end flow over stdio: start_project -> propose_outline -> hand edits (ui_* tools) -> approve ->
 * propose_layouts -> notes -> resolve -> build_workbook, then templates, a format skill and a deck outline.
 * Run: npx tsx test/flow.test.ts
 */
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const home = fs.mkdtempSync(path.join(os.tmpdir(), "blueprint-test-"));
const out = path.join(home, "out");
const client = new Client({ name: "flow-test", version: "0" });
await client.connect(new StdioClientTransport({ command: "node", args: ["dist/server/main.js"], env: { ...process.env, BLUEPRINT_HOME: home, BLUEPRINT_OUTPUT: out } }));

const call = async (name: string, args: Record<string, unknown>) => {
  const r: any = await client.callTool({ name, arguments: args });
  const text = r.content?.map((c: any) => c.text).join("\n") ?? "";
  return { r, text, sc: r.structuredContent as any, isError: !!r.isError };
};

// tool list: app-only tools must still be listed (host filters by _meta), model tools have descriptions
const tools = await client.listTools();
const names = tools.tools.map((t) => t.name);
for (const n of ["start_project", "propose_outline", "propose_layouts", "get_plan", "update_plan", "resolve_notes", "build_workbook", "review_workbook", "ui_get_state", "ui_save_layout", "ui_save_brief", "ui_save_tab_fields", "ui_set_build_options"]) assert.ok(names.includes(n), `missing tool ${n}`);
assert.ok(!names.includes("propose_tabs"), "propose_tabs was renamed to propose_outline");
const pt = tools.tools.find((t) => t.name === "propose_outline")!;
assert.match(String((pt._meta as any)?.ui?.resourceUri), /^ui:\/\/blueprint\/canvas-\d+\.\d+\.\d+\.html$/);
const res = await client.listResources();
const canvasUri = res.resources.find((r) => r.uri.startsWith("ui://blueprint/canvas-"))!.uri; assert.ok(canvasUri);
const html = await client.readResource({ uri: canvasUri });
assert.ok((html.contents[0] as any).text.includes("<script"), "canvas html served");

// 1. start
let s = await call("start_project", { name: "Acme Market Sizing", brief: "Market sizing model for Acme's B2B payments product", formatting: "house" });
assert.ok(!s.isError, s.text);
const pid = s.sc.projectId as string;
assert.equal(s.sc.kind, "workbook", "kind defaults to workbook");
assert.ok(s.text.includes("clarifying questions"));
assert.ok(s.text.split("\n").length <= 16, "start_project result stays short");

// 2. tabs
s = await call("propose_outline", { project: pid, rows: [
  { name: "Assumptions", purpose: "Inputs everything else reads", message: "Change numbers here, nowhere else" },
  { name: "TAM Build", purpose: "Top-down market funnel to serviceable market" },
  { name: "Revenue Ramp", purpose: "Five-year revenue build from share and price" },
  { name: "Sources & Appendix", purpose: "Every reference cited in the model" },
] });
assert.equal(s.sc.view, "tabs"); assert.equal(s.sc.project.tabs.length, 4);
assert.equal(s.sc.project.tabs[0].message, "Change numbers here, nowhere else");
assert.ok(/message: Change numbers here, nowhere else/.test(s.text), "the message reaches Claude in the summary");
const tabs = s.sc.project.tabs;

// hand edit: reorder + rename via the canvas tool, add a tab note
s = await call("ui_save_tabs", { project: pid, tabs: [tabs[0], tabs[2], tabs[1], tabs[3]].map((t: any) => ({ id: t.id, name: t.name === "TAM Build" ? "Market Funnel" : t.name, purpose: t.purpose })) });
assert.equal(s.sc.project.tabs[1].name, "Revenue Ramp"); assert.equal(s.sc.project.tabs[2].name, "Market Funnel"); assert.equal(s.sc.project.tabs[2].id, tabs[1].id, "id kept on rename");
assert.equal(s.sc.project.tabs[0].message, "Change numbers here, nowhere else", "hand edits keep the message");
s = await call("ui_save_brief", { project: pid, audience: "Acme's board", takeaway: "The serviceable market supports the raise" });
assert.equal(s.sc.project.briefCard.takeaway, "The serviceable market supports the raise");
s = await call("ui_note", { project: pid, action: "add", tabId: tabs[0].id, text: "Keep every input on this tab, none inline" });
// layouts refused before approval
s = await call("propose_layouts", { project: pid, layouts: [{ tab: "Assumptions", blocks: [] }] });
assert.ok(s.isError && /not approved/.test(s.text));
s = await call("ui_approve", { project: pid, what: "tabs" });
assert.equal(s.sc.project.tabsApproved, true);

// 3. layouts
s = await call("get_plan", { project: pid });
assert.ok(/Market Funnel/.test(s.text) && /open instructions from the user \(1\)/.test(s.text));
s = await call("propose_layouts", { project: pid, layouts: [
  { tab: "Assumptions", blocks: [
    { type: "inputs", title: "Key Assumptions", rowLabels: ["Addressable businesses", "Penetration, Year 5", "Average price ($)", "Growth"], questions: ["Is price annual or monthly?"] },
  ] },
  { tab: "Market Funnel", blocks: [
    { type: "table", title: "Funnel", columns: ["Stage", "Businesses", "Share"], sample: [["All businesses", 32000000, "100%"], ["Addressable", 6400000, "20%"]] },
    { type: "output", title: "Serviceable Market", rowLabels: ["Serviceable businesses"] },
  ] },
  { tab: "Revenue Ramp", blocks: [
    { type: "calc", title: "Five Year Ramp", columns: ["Line", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5"], rowLabels: ["Customers", "Revenue ($M)"], row: 14, col: 2 },
    { type: "chart", title: "Revenue by Year", row: 2, col: 2 },
    { type: "output", title: "Year 5 Outcome", rowLabels: ["Year 5 revenue ($M)"], row: 22, col: 2 },
  ] },
  { tab: "Sources & Appendix", blocks: [ { type: "table", title: "Sources", columns: ["Ref", "Source", "Key figures", "Link"] } ] },
] });
assert.equal(s.sc.view, "layouts");
const ramp = s.sc.project.tabs.find((t: any) => t.name === "Revenue Ramp");
const rampCalc = ramp.blocks.find((b: any) => b.type === "calc"); const chart = ramp.blocks.find((b: any) => b.type === "chart");
assert.equal(rampCalc.row, 14); assert.ok(s.sc.warnings.some((w: any) => w.kind === "chart-source"), "chart without source flagged");
assert.ok(s.sc.warnings.some((w: any) => w.kind === "question"));
const qNote = s.sc.project.notes.find((n: any) => n.kind === "question");
assert.ok(qNote && qNote.createdBy === "claude" && qNote.status === "open", "question became an open thread");

// hand edit: drag the ramp table to the top, chart below it, point chart at the table (ui_save_layout)
const moved = ramp.blocks.map((b: any) => b.id === rampCalc.id ? { ...b, row: 2 } : b.id === chart.id ? { ...b, row: 8, source: rampCalc.id } : b);
s = await call("ui_save_layout", { project: pid, tab: ramp.id, blocks: moved });
assert.equal(s.sc.project.tabs.find((t: any) => t.id === ramp.id).blocks.find((b: any) => b.id === rampCalc.id).row, 2);
assert.ok(!s.sc.warnings.some((w: any) => w.kind === "chart-source"));
// block note from the canvas
s = await call("ui_note", { project: pid, action: "add", tabId: ramp.id, blockId: rampCalc.id, text: "Monthly granularity for years 1-2 is NOT needed; keep annual" });
// chat edit: add a block via update_plan
s = await call("update_plan", { project: pid, ops: [{ op: "add_block", tab: "Assumptions", block: { type: "text", title: "How to read", cols: 4 } }] });
assert.equal(s.sc.project.tabs[0].blocks.length, 2);
// the user answers Claude's question on the canvas; that closes it and the answer reads back
s = await call("ui_note", { project: pid, action: "answer", noteId: qNote.id, text: "Annual." });
const answered = s.sc.project.notes.find((n: any) => n.id === qNote.id);
assert.equal(answered.status, "resolved"); assert.equal(answered.response, "Annual.");
// W8: a plan that calculates and carries no Checks block is a readiness warning, and a Checks block clears it
{
  const before = await call("ui_get_state", { project: pid });
  assert.ok(before.sc.warnings.some((w: any) => w.kind === "no-checks"), JSON.stringify(before.sc.warnings.map((w: any) => w.kind)));
  const added = await call("update_plan", { project: pid, ops: [{ op: "add_block", tab: "Revenue Ramp", block: { type: "output", title: "Checks", rowLabels: ["Ramp ties to the Year 5 outcome", "No error values", "Every input read"] } }] });
  assert.ok(!added.isError, added.text);
  assert.ok(!added.sc.warnings.some((w: any) => w.kind === "no-checks"), "a Checks block clears it");
  assert.ok(!/No Checks block/.test(added.text), added.text);
  const checksId = added.sc.project.tabs.find((t: any) => t.name === "Revenue Ramp").blocks.find((b: any) => b.title === "Checks").id;
  const gone = await call("update_plan", { project: pid, ops: [{ op: "remove_block", blockId: checksId }] });
  assert.ok(gone.sc.warnings.some((w: any) => w.kind === "no-checks"), "and it comes back when the block goes");
  assert.ok(/No Checks block: put one on the outputs tab/.test(gone.text), gone.text);
}
// build refused: layouts unapproved
s = await call("build_workbook", { project: pid, sheets: [{ tab: "Assumptions", blocks: [] }] });
assert.ok(s.isError && /not yet approved/.test(s.text), s.text);
s = await call("ui_approve", { project: pid, what: "all_layouts" });
// build refused: open notes
s = await call("build_workbook", { project: pid, sheets: [{ tab: "Assumptions", blocks: [] }] });
assert.ok(s.isError && /thread\(s\) are still open/.test(s.text), s.text);
const plan = await call("get_plan", { project: pid });
assert.ok(/answers from the user \(1\)/.test(plan.text), "answered question surfaces to Claude");
const notes = plan.sc.project.notes.filter((n: any) => n.status === "open");
assert.equal(notes.length, 2);
s = await call("resolve_notes", { project: pid, resolutions: notes.map((n: any, i: number) => ({ noteId: n.id, status: i === 0 ? "resolved" : "declined", response: i === 0 ? "All inputs live on Assumptions." : "Annual kept; monthly would triple the row count." })) });
assert.ok(/0 note\(s\) still open/.test(s.text));
// build refused: missing block
s = await call("build_workbook", { project: pid, sheets: [{ tab: "Assumptions", blocks: [] }] });
assert.ok(s.isError && /missing from the build/.test(s.text), s.text);

// payload errors name the block, not the array index
{
  const P0 = plan.sc.project; const first = P0.tabs[0].blocks[0];
  const bad = await call("build_workbook", { project: pid, sheets: [{ tab: "Assumptions", blocks: [{ id: first.id, rows: [{ label: "x", value: 1, format: "dollars" }] }] }] });
  assert.ok(bad.isError && /rows\.0\.format/.test(bad.text), bad.text); // schema errors carry the path; semantic errors name the block
}

// real build
const P = plan.sc.project; const cells = plan.sc.cells;
const bid = (tab: string, type: string) => P.tabs.find((t: any) => t.name === tab).blocks.find((b: any) => b.type === type).id;
const A = bid("Assumptions", "inputs"), T = bid("Assumptions", "text"), F = bid("Market Funnel", "table"), SM = bid("Market Funnel", "output"), R = bid("Revenue Ramp", "calc"), C = bid("Revenue Ramp", "chart"), Y5 = bid("Revenue Ramp", "output"), SRC = bid("Sources & Appendix", "table");
assert.ok(cells[A].valueCells["Growth"], "get_plan hands out value cell addresses");
// rowFormats on an inputs block is refused: its rows carry their own format
{
  const bad = await call("build_workbook", { project: pid, sheets: [{ tab: "Assumptions", blocks: [{ id: A, rowFormats: ["currency"], rows: [{ label: "Growth", value: 0.35 }] }] }] });
  assert.ok(bad.isError && /takes format per row, not rowFormats/.test(bad.text), bad.text);
}
const READ_ME = { purpose: "Sizes the market for Acme's B2B payments product and ramps revenue over five years." };
const BUILD_SHEETS: any[] = [
  { tab: "Assumptions", blocks: [
    { id: A, rows: [ { label: "Addressable businesses", value: 6400000, format: "count", note: "Census count of firms with 10 to 500 staff; see R1" }, { label: "Penetration, Year 5", value: 0.04, format: "percent", note: "Management assumption; no published figure" }, { label: "Average price ($)", value: 2400, format: "currency" }, { label: "Growth", value: 0.35, format: "percent", validation: { type: "decimal", min: 0, max: 1 } }, { label: "Scenario", value: "Base", validation: { list: ["Low", "Base", "High"] } } ] },
    { id: T, text: "Blue cells are inputs. Everything else is a formula." },
  ] },
  { tab: "Market Funnel", blocks: [
    { id: F, columns: ["Stage", "Businesses", "Share"], colFormats: ["text", "count", "percent"], data: [["All US businesses", 32000000, 1], ["Addressable", { f: "={{" + A + "!Addressable businesses}}" }, { f: "=C5/C4" }]] },
    { id: SM, rows: [ { label: "Serviceable businesses", value: { f: "={{" + F + "!1,1}}*{{" + A + "!Penetration, Year 5}}" }, format: "count" } ] },
  ] },
  { tab: "Revenue Ramp", freeze: "C4", landscape: true, printArea: "auto", blocks: [
    // rows are the metrics here, so the row formats win over the column ones (W3)
    { id: R, columns: ["Line", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5"], colFormats: ["text", "count", "count", "count", "count", "count"], rowFormats: ["count", "currency"], rowNotes: ["Serviceable market over five years", "Customers x price"], data: [
      ["Customers", { f: "={{" + SM + "!Serviceable businesses}}/5" }, { f: "=C4*(1+{{" + A + "!Growth}})" }, { f: "=D4*(1+{{" + A + "!Growth}})" }, { f: "=E4*(1+{{" + A + "!Growth}})" }, { f: "=F4*(1+{{" + A + "!Growth}})" }],
      ["Revenue ($M)", { f: "=C4*{{" + A + "!Average price ($)}}/1000000" }, { f: "=D4*{{" + A + "!Average price ($)}}/1000000" }, { f: "=E4*{{" + A + "!Average price ($)}}/1000000" }, { f: "=F4*{{" + A + "!Average price ($)}}/1000000" }, { f: "=G4*{{" + A + "!Average price ($)}}/1000000" }],
    ] },
    { id: C, chart: { kind: "bar", sourceBlock: R, categoriesCol: 0, seriesCols: [1, 2, 3, 4, 5] } },
    { id: Y5, rows: [ { label: "Year 5 revenue ($M)", value: { f: "={{" + R + "!1,5}}" }, format: "currency" } ] },
  ] },
  { tab: "Sources & Appendix", blocks: [ { id: SRC, columns: ["Ref", "Source", "Key figures", "Link"], data: [["R1", "US Census, County Business Patterns 2024", "6.4M firms with 10 to 500 employees", "https://www.census.gov/programs-surveys/cbp.html"]] } ] },
];
s = await call("build_workbook", { project: pid, readme: READ_ME, sheets: BUILD_SHEETS });
assert.ok(!s.isError, s.text);
assert.equal(s.sc.view, "built");
assert.ok(/Formula check \((libreoffice|python)\): 0 errors/.test(s.text), s.text.split("\n").find((l: string) => l.startsWith("Formula check")));
{
  // the result stays inside one screen; the images ride behind it, not in it
  const head = (s.r.content as any[])[0].text as string;
  assert.ok(head.split("\n").length <= 12, `build_workbook result is ${head.split("\n").length} lines:\n${head}`);
}
// W3: rowFormats reaches the builder spec, beside colFormats
{
  const specPath = s.sc.project.specPath as string;
  assert.ok(specPath && fs.existsSync(specPath), `build_workbook records the spec it sent: ${specPath}`);
  const spec = JSON.parse(fs.readFileSync(specPath, "utf8"));
  const rampBlock = spec.sheets.flatMap((sh: any) => sh.blocks ?? []).find((b: any) => b.id === R);
  assert.deepEqual(rampBlock.rowFormats, ["count", "currency"], "rowFormats passed to the builder");
  assert.deepEqual(rampBlock.colFormats, ["text", "count", "count", "count", "count", "count"], "and colFormats with it");
  // W11: the plan's fidelity, with nothing set, is the default the builder is told about
  assert.equal(spec.fidelity, "balanced", "the plan's fidelity reaches the builder");
  // N9: comments are off until the user asks, and the closed threads ride along either way, grouped by block
  assert.equal(spec.comments, false, "comments stay off until the user turns them on");
  const onRamp = (spec.threads as any[]).find((t: any) => t.blockId === R);
  assert.ok(onRamp, `threads by block: ${JSON.stringify((spec.threads as any[]).map((t: any) => [t.tab, t.blockId]))}`);
  assert.equal(onRamp.tab, "Revenue Ramp");
  assert.deepEqual(onRamp.entries, [
    { by: "user", text: "Monthly granularity for years 1-2 is NOT needed; keep annual", status: "declined" },
    { by: "claude", text: "Annual kept; monthly would triple the row count.", status: "declined" },
  ], "the ask and the answer are two entries of one thread");
  // a note left on the tab rather than on a block hangs on that tab's first block, which is where the comment goes
  const onAssumptions = (spec.threads as any[]).filter((t: any) => t.blockId === A);
  assert.equal(onAssumptions.length, 2, `the tab note and the question: ${JSON.stringify(onAssumptions)}`);
  assert.deepEqual(onAssumptions[0].entries, [
    { by: "user", text: "Keep every input on this tab, none inline", status: "resolved" },
    { by: "claude", text: "All inputs live on Assumptions.", status: "resolved" },
  ], "the tab note, and what Claude did about it");
  assert.deepEqual(onAssumptions[1].entries, [
    { by: "claude", text: "Is price annual or monthly?", status: "resolved" },
    { by: "user", text: "Annual.", status: "resolved" },
  ], "a question Claude asked carries the user's answer back");
  assert.ok((spec.threads as any[]).every((t: any) => t.blockId), "every thread names the block whose cell carries it");
  assert.ok((spec.threads as any[]).every((t: any) => t.entries.every((e: any) => e.status !== "open")), "an open thread never reaches the builder; the build refuses first");
}
// W11: the width moves the builder took are counted in the summary and get one line in the result
{
  const moves = (s.sc.snapshot.sheets as any[]).flatMap((sh: any) => sh.adjusted ?? []);
  assert.deepEqual((s.sc.build.adjusted as any[]).map((a: any) => a.blockId).sort(), moves.map((a: any) => a.blockId).sort(), "every move the snapshot records reaches the build summary");
  assert.equal(s.sc.build.fidelity, "balanced");
  if (moves.length) assert.ok(new RegExp(`^Adjusted ${moves.length} blocks? \\(balanced\\)$`, "m").test(s.text), s.text);
  else assert.ok(!/^Adjusted /m.test(s.text), `nothing was adjusted, so nothing is said: ${s.text}`);
  console.log(`ADJUSTED ${moves.length}`);
}
// W7: the sheets come back as pictures, one per sheet, and the paths are kept for review_workbook
const sheetImgs = (s.r.content as any[]).filter((c) => c.type === "image");
if (sheetImgs.length) {
  assert.equal(sheetImgs.length, 4, "one image per sheet");
  assert.equal(sheetImgs[0].mimeType, "image/png");
  assert.deepEqual(sheetImgs[0].annotations, { audience: ["assistant"], priority: 0.6 });
  assert.ok(/Sheet images follow: look at each/.test(s.text), s.text);
  assert.ok(/In order: 1 Assumptions, 2 Market Funnel, 3 Revenue Ramp, 4 Sources & Appendix\./.test(s.text), s.text);
  assert.equal((s.sc.project.thumbnails as string[]).length, 4, "PNG paths kept on the project");
  assert.deepEqual((s.sc.build.sheetImages as any[]).map((x) => x.sheet), ["Assumptions", "Market Funnel", "Revenue Ramp", "Sources & Appendix"]);
} else {
  assert.ok(/No sheet images: /.test(s.text), `with no renderer the result has to say so: ${s.text}`);
  console.log("SHEET IMAGES SKIPPED: no soffice or pdftoppm");
}
const built = s.sc.project.buildPath as string;
assert.ok(fs.existsSync(built), "xlsx written");
{
  const { execFileSync } = await import("node:child_process");
  const probe = execFileSync("python3", ["-c", `import openpyxl,sys; wb=openpyxl.load_workbook(sys.argv[1]); a=wb["Assumptions"]; r=wb["Revenue Ramp"]; print(len(a.data_validations.dataValidation), r.freeze_panes, r.page_setup.orientation, r["H4"].value, bool(r.print_area))`, built]).toString().trim();
  assert.equal(probe, "2 C4 landscape Serviceable market over five years True", probe);
  // The builder's half of W3: the revenue row is currency once build_xlsx.py reads rowFormats. Said, not asserted,
  // because the spec carries them either way and that is this side's contract.
  const fmt = execFileSync("python3", ["-c", `import openpyxl,sys; print(openpyxl.load_workbook(sys.argv[1])["Revenue Ramp"].cell(5, 3).number_format)`, built]).toString().trim();
  console.log(fmt.includes("$") ? "ROWFORMATS applied by the builder" : `ROWFORMATS in the spec, not yet applied by the builder (${fmt})`);
}
const snapPath = s.sc.project.snapshotPath as string;
assert.equal(snapPath, built.replace(/\.xlsx$/, ".snapshot.json"));
assert.ok(fs.existsSync(snapPath), "snapshot written");
assert.equal(s.sc.build.snapshot, snapPath);
const snapSheets = s.sc.snapshot.sheets as any[];
assert.deepEqual(snapSheets.map((x) => x.name), s.sc.build.sheets.map((x: any) => x.name));
assert.equal(snapSheets.length, 4);
const rampSnap = snapSheets.find((x) => x.name === "Revenue Ramp");
assert.equal(rampSnap.charts.length, 1); assert.equal(rampSnap.charts[0].series, 5);
assert.ok(rampSnap.cells.some((c: any) => c.k === "formula"), "formula kinds recorded");
assert.ok(rampSnap.blocks.some((b: any) => b.id === R), "block extents recorded");
// the canvas cannot read files, so a built view with no preview asks the poll for the snapshot
{
  const lean = await call("ui_get_state", { project: pid });
  assert.equal(lean.sc.snapshot, undefined, "an ordinary poll stays lean");
  assert.equal(lean.sc.style.kind, "workbook", "a workbook carries its style id and no deck tokens");
  assert.equal(lean.sc.style.tokens, undefined);
  assert.ok(lean.sc.style.id, "the style id is there");
  const full = await call("ui_get_state", { project: pid, withSnapshot: true });
  assert.equal((full.sc.snapshot.sheets as any[]).length, 4, "withSnapshot carries the preview");
  assert.ok(full.sc.build, "and the build summary with it");
}
// W7: review_workbook hands the sheet back on demand, with its blocks and its formula findings
{
  const one = await call("review_workbook", { project: pid, sheets: ["Revenue Ramp"] });
  assert.ok(!one.isError, one.text);
  const head = (one.r.content as any[])[0].text as string;
  assert.match(head, /^Revenue Ramp: \d+ cells, \d+ formulas, 1 charts \| blocks: calc Five Year Ramp, chart Revenue by Year, output Year 5 Outcome$/);
  if (sheetImgs.length) {
    assert.equal((one.r.content as any[]).filter((c) => c.type === "image").length, 1, "one sheet, one image");
    assert.ok(/Sheet images follow: look at each/.test(one.text), one.text);
  }
  const all = await call("review_workbook", { project: pid });
  assert.equal(((all.r.content as any[])[0].text as string).split("\n").length, 4, "every sheet by default, one line each");
  const miss = await call("review_workbook", { project: pid, sheets: ["Nope"] });
  assert.ok(/No sheet named Nope\. This build has: Assumptions, Market Funnel/.test(miss.text), miss.text);
}
// N9: the user turns comments on for this plan; W11: `fidelity` on the call is for that build only
{
  const o = await call("ui_set_build_options", { project: pid, comments: true, fidelity: "strict" });
  assert.ok(!o.isError, o.text);
  assert.deepEqual(o.sc.buildOptions, { fidelity: "strict", notesFromStory: true, cover: true, agenda: true, comments: true }, JSON.stringify(o.sc.buildOptions));
  const g = await call("get_plan", { project: pid });
  assert.ok(/^build: strict, notes from story, cover, agenda, comments$/m.test(g.text), g.text.split("\n").find((l: string) => l.startsWith("build:")) ?? g.text);
}
// W7: no LibreOffice, no images, and the result says so instead of going quiet
{
  const c0 = new Client({ name: "flow-test-no-soffice", version: "0" });
  await c0.connect(new StdioClientTransport({ command: "node", args: ["dist/server/main.js"], env: { ...process.env, BLUEPRINT_HOME: home, BLUEPRINT_OUTPUT: out, BLUEPRINT_SOFFICE: path.join(home, "no-soffice") } }));
  const r: any = await c0.callTool({ name: "build_workbook", arguments: { project: pid, readme: READ_ME, sheets: BUILD_SHEETS, fidelity: "balanced" } });
  const text = r.content?.map((c: any) => c.text).join("\n") ?? "";
  assert.ok(!r.isError, text);
  assert.equal((r.content as any[]).filter((c: any) => c.type === "image").length, 0, "no renderer, no images");
  assert.ok(/No sheet images: LibreOffice \(soffice\) not installed/.test(text), text);
  assert.ok(!/Sheet images follow/.test(text), text);
  {
    const builds = path.join(home, "builds");
    const latest = fs.readdirSync(builds).filter((f) => f.endsWith(".json") && !f.endsWith(".check.json") && !f.endsWith(".quality.json")).sort().pop()!;
    const spec2 = JSON.parse(fs.readFileSync(path.join(builds, latest), "utf8"));
    assert.equal(spec2.comments, true, "the comment setting reaches the builder");
    assert.equal(spec2.fidelity, "balanced", "fidelity on the call wins over the plan's strict, for this build only");
    assert.ok(spec2.threads?.length, "and the closed threads are still there");
  }
  await c0.close();
}
console.log("BUILT", built);
console.log(s.text);
fs.writeFileSync(path.join(home, "BUILT_PATH"), built);
await client.close();
console.log("FLOW OK", home);

// ---- templates: save the approved plan, start a new project from it
{
  const c2 = new Client({ name: "flow-test-2", version: "0" });
  await c2.connect(new StdioClientTransport({ command: "node", args: ["dist/server/main.js"], env: { ...process.env, BLUEPRINT_HOME: home, BLUEPRINT_OUTPUT: out } }));
  const call2 = async (name: string, args: Record<string, unknown>) => { const r: any = await c2.callTool({ name, arguments: args }); return { r, text: r.content?.map((c: any) => c.text).join("\n") ?? "", sc: r.structuredContent as any, isError: !!r.isError }; };
  let t = await call2("save_template", { project: pid, name: "Market sizing", description: "Funnel, ramp, sources" });
  assert.ok(!t.isError && /Saved template market-sizing/.test(t.text), t.text);
  t = await call2("list_templates", {});
  assert.ok(/market-sizing: Market sizing \(4 tabs/.test(t.text), t.text);
  t = await call2("start_project", { name: "Beta Sizing", brief: "Second model", template: "market-sizing" });
  assert.ok(/From template/.test(t.text), t.text);
  const pid2 = t.sc.projectId as string;
  t = await call2("get_plan", { project: pid2 });
  assert.equal(t.sc.project.tabs.length, 4); assert.equal(t.sc.project.tabsApproved, false);
  assert.equal(t.sc.project.notes.length, 0, "notes are not carried into templates");
  const ramp2 = t.sc.project.tabs.find((x: any) => x.name === "Revenue Ramp");
  const chart2 = ramp2.blocks.find((b: any) => b.type === "chart"); const calc2 = ramp2.blocks.find((b: any) => b.type === "calc");
  assert.equal(chart2.source, calc2.id, "chart source re-pointed to the copied block id");
  assert.notEqual(calc2.id, R, "block ids are fresh");
  await c2.close();
  console.log("TEMPLATES OK");
}

// ---- skill style: a formatting skill with a format: block drives the builder
{
  const skillDir = path.join(home, "skills", "test-format"); fs.mkdirSync(skillDir, { recursive: true });
  fs.writeFileSync(path.join(skillDir, "SKILL.md"), `---\nname: test-format\ndescription: "Test workbook format for the flow test."\nformat:\n  font: Calibri\n  size: 9\n  input_color: 1F5F99\n  label_divider: none\n  cover_tab: false\n---\n# Test format\nLabels in plain English.\n`);
  const c3 = new Client({ name: "flow-test-3", version: "0" }); // old env names still honoured
  await c3.connect(new StdioClientTransport({ command: "node", args: ["dist/server/main.js"], env: { ...process.env, WORKBOOK_PLANNER_HOME: home, WORKBOOK_PLANNER_OUTPUT: out, WORKBOOK_PLANNER_SKILLS: path.join(home, "skills") } }));
  const call3 = async (name: string, args: Record<string, unknown>) => { const r: any = await c3.callTool({ name, arguments: args }); return { r, text: r.content?.map((c: any) => c.text).join("\n") ?? "", sc: r.structuredContent as any, isError: !!r.isError }; };
  let t = await call3("ui_styles", { project: pid });
  assert.ok(t.sc.styles.some((s: any) => s.id === "skill:test-format" && s.builderMode === "skill"), JSON.stringify(t.sc.styles.map((s: any) => s.id)));
  t = await call3("ui_set_style", { project: pid, styleId: "skill:test-format" });
  assert.equal(t.sc.project.formatting, "skill");
  t = await call3("get_plan", { project: pid });
  assert.ok(/Labels in plain English/.test(t.text), "skill text reaches Claude");
  // P8: once per style. The next call references the rules instead of repeating thousands of tokens.
  t = await call3("get_plan", { project: pid });
  assert.ok(!/Labels in plain English/.test(t.text), `the prose should not repeat:\n${t.text.slice(0, 400)}`);
  assert.ok(/^Style: test-format \(rules sent earlier; call get_plan with `skillText: true` to see them again\)$/m.test(t.text), t.text.split("\n").find((l: string) => l.startsWith("Style:")) ?? t.text);
  t = await call3("get_plan", { project: pid, skillText: true });
  assert.ok(/Labels in plain English/.test(t.text), "and comes back when it is asked for");
  t = await call3("get_plan", { project: pid });
  assert.ok(!/Labels in plain English/.test(t.text), "then goes quiet again");
  // a built-in format has no prose, so there is nothing to reference either
  await call3("ui_set_style", { project: pid, styleId: "house" });
  t = await call3("get_plan", { project: pid });
  assert.ok(!/rules sent earlier/.test(t.text) && !/Labels in plain English/.test(t.text), t.text);
  // a different skill is a different set of rules: those are sent in full
  const other = path.join(home, "skills", "test-format-b"); fs.mkdirSync(other, { recursive: true });
  fs.writeFileSync(path.join(other, "SKILL.md"), `---\nname: test-format-b\ndescription: "Second test workbook format for the flow test."\nformat:\n  font: Arial\n---\n# Second format\nRow labels in the past tense.\n`);
  await call3("ui_set_style", { project: pid, styleId: "skill:test-format-b" });
  t = await call3("get_plan", { project: pid });
  assert.ok(/Row labels in the past tense/.test(t.text), "a style the plan has not had sends its rules");
  t = await call3("get_plan", { project: pid });
  assert.ok(/^Style: test-format-b \(rules sent earlier/m.test(t.text), t.text.split("\n").find((l: string) => l.startsWith("Style:")) ?? t.text);
  await call3("ui_set_style", { project: pid, styleId: "skill:test-format" });
  // rebuild with the same payload by reusing the saved spec through build_workbook is verbose; instead run the builder on the last spec with the skill fields
  const specs = fs.readdirSync(path.join(home, "builds")).filter((f) => f.endsWith(".json") && !f.endsWith(".check.json") && !f.endsWith(".quality.json")).sort();
  const spec = JSON.parse(fs.readFileSync(path.join(home, "builds", specs[specs.length - 1]), "utf8"));
  spec.formatting = "skill"; spec.skill = { path: path.join(skillDir, "SKILL.md") }; spec.outputPath = path.join(out, "skill-build.xlsx");
  fs.writeFileSync(path.join(home, "skill-spec.json"), JSON.stringify(spec));
  const { execFileSync } = await import("node:child_process");
  const res = JSON.parse(execFileSync("python3", ["builder/build_xlsx.py", path.join(home, "skill-spec.json")]).toString().trim().split("\n").pop()!);
  assert.ok(res.ok, JSON.stringify(res.errors));
  const py = execFileSync("python3", ["-c", `import openpyxl,sys; wb=openpyxl.load_workbook(sys.argv[1]); ws=wb["Assumptions"]; c=ws["C3"]; print(c.font.name, c.font.sz, c.font.color.rgb, ws["B3"].border.right.style, "Read Me" in wb.sheetnames)`, spec.outputPath]).toString().trim();
  assert.equal(py, "Calibri 9.0 001F5F99 thin False", py); // divider none: the ordinary thin body border remains
  await c3.close();
  console.log("SKILL STYLE OK");
}

// ---- deck mode: dividers, headlines, the furniture band and the archetype solver. The build has its own suite.
{
  const c4 = new Client({ name: "flow-test-deck", version: "0" });
  await c4.connect(new StdioClientTransport({ command: "node", args: ["dist/server/main.js"], env: { ...process.env, BLUEPRINT_HOME: home, BLUEPRINT_OUTPUT: out } }));
  const call4 = async (name: string, args: Record<string, unknown>) => { const r: any = await c4.callTool({ name, arguments: args }); return { r, text: r.content?.map((c: any) => c.text).join("\n") ?? "", sc: r.structuredContent as any, isError: !!r.isError }; };

  // the brief is three fields: an emphasis chip from an old payload is accepted and dropped (F7)
  let d = await call4("start_project", { name: "Q3 Board Update", kind: "deck", brief: "Board deck for the Q3 review", briefCard: { audience: "The board", goal: "Approve the next raise", takeaway: "Retention is compounding; fund two more pods", emphasis: ["retention", "cash"] } });
  assert.ok(!d.isError, d.text);
  const dpid = d.sc.projectId as string;
  assert.equal(d.sc.kind, "deck");
  assert.ok(/deck plan/.test(d.text) && /propose_outline/.test(d.text), d.text);
  assert.ok(d.text.split("\n").length <= 16, "start_project result stays short");
  assert.ok(/divider: true/.test(d.text), "the workflow says a section is a divider row");
  assert.ok(/never a headline block/i.test(d.text), "and that the headline is not a block");
  {
    // the deck workflow is the model's whole briefing: one screen, 12 lines at most (research item 4)
    const wf = d.text.split("\n\n").slice(1).join("\n\n").split("\n");
    assert.ok(wf.length <= 12, `the deck workflow is ${wf.length} lines, 12 at most:\n${wf.join("\n")}`);
    assert.ok(/a diagram block with nodes and edges/.test(d.text), "and says a relationship is a diagram");
    assert.ok(/never a placeholder/.test(d.text) && /a real image or a designed visual/.test(d.text), d.text);
  }

  d = await call4("get_plan", { project: dpid });
  assert.equal(d.sc.project.tabs.length, 0, "a deck starts with no slides");
  assert.equal(d.sc.project.kind, "deck");
  assert.equal(d.sc.project.briefCard.emphasis, undefined, "emphasis is never stored");
  assert.ok(/kind: deck/.test(d.text) && /takeaway: Retention is compounding/.test(d.text), d.text);
  assert.ok(!/emphasis/.test(d.text), "and never reaches Claude");

  // ---- dividers are rows: sent as rows, and a legacy section on a row becomes one (F4, F10)
  d = await call4("propose_outline", { project: dpid, rows: [
    { name: "Context", divider: true },
    { name: "Where we landed", story: "Opens on the quarter's result so the board can park the number and listen. Names the one driver that moved." },
    { name: "What changed", story: "Shows the two cohort curves side by side. The reader should see the gap without being told." },
    { name: "What we need", story: "Turns the result into the ask: two more pods, funded now, and what they buy.", section: "Ask" },
  ] });
  assert.ok(!d.isError, d.text);
  assert.equal(d.sc.view, "tabs");
  assert.deepEqual(d.sc.project.tabs.map((t: any) => t.name), ["Context", "Where we landed", "What changed", "Ask", "What we need"], "the legacy section became a divider row above its slide");
  assert.deepEqual(d.sc.project.tabs.map((t: any) => !!t.divider), [true, false, false, true, false]);
  assert.deepEqual(d.sc.project.tabs.map((t: any) => t.section), ["Context", "Context", "Context", "Ask", "Ask"], "sections are derived from the nearest divider above");
  assert.ok(/3 slides and 2 dividers/.test(d.text), d.text);
  assert.ok(/divider is a row in the list/.test(d.text), d.text);
  assert.ok(/\[divider: a section slide, no blocks\]/.test(d.text), "the summary marks divider rows");

  // adjacent dividers merge, and a run of them keeps the last name
  {
    const m = await call4("propose_outline", { project: dpid, rows: [
      { name: "Context", divider: true },
      { name: "context", divider: true },
      { name: "Opening", story: "One slide." },
      { name: "First", divider: true },
      { name: "Second", divider: true },
      { name: "Closing", story: "Another slide.", section: "Second" },
    ] });
    assert.deepEqual(m.sc.project.tabs.map((t: any) => t.name), ["context", "Opening", "Second", "Closing"], "same-name dividers merge; a run keeps the last name");
    assert.deepEqual(m.sc.project.tabs.map((t: any) => t.section), ["context", "context", "Second", "Second"]);
  }

  // back to the real outline
  d = await call4("propose_outline", { project: dpid, rows: [
    { name: "Context", divider: true },
    { name: "Where we landed", story: "Opens on the quarter's result so the board can park the number and listen. Names the one driver that moved." },
    { name: "What changed", story: "Shows the two cohort curves side by side. The reader should see the gap without being told." },
    { name: "Ask", divider: true },
    { name: "What we need", story: "Turns the result into the ask: two more pods, funded now, and what they buy." },
  ] });
  assert.equal(d.sc.project.tabs.length, 5);

  d = await call4("get_plan", { project: dpid });
  assert.ok(/story: Opens on the quarter's result/.test(d.text), "stories reach Claude");
  assert.ok(/\[Context\]/.test(d.text), "derived sections reach Claude");
  assert.ok(/slide geometry: 32 columns by 18 rows/.test(d.text) && /rows 5 to 17/.test(d.text), d.text);

  // ---- a divider's agenda line: the one line the generated agenda prints under the section title (Q17)
  {
    const AGENDA = "Where the quarter landed, in one number.";
    // propose_outline takes it on a divider row; a content row's is dropped
    let a = await call4("propose_outline", { project: dpid, rows: [
      { name: "Context", divider: true, agendaLine: AGENDA },
      { name: "Where we landed", story: "Opens on the quarter's result so the board can park the number and listen. Names the one driver that moved.", agendaLine: "Never printed." },
      { name: "What changed", story: "Shows the two cohort curves side by side. The reader should see the gap without being told." },
      { name: "Ask", divider: true },
      { name: "What we need", story: "Turns the result into the ask: two more pods, funded now, and what they buy." },
    ] });
    assert.ok(!a.isError, a.text);
    assert.equal(a.sc.project.tabs[0].agendaLine, AGENDA);
    assert.equal(a.sc.project.tabs[1].agendaLine, undefined, "only a divider row carries one");
    assert.ok(/\[divider: a section slide, no blocks\] \[agenda line: Where the quarter landed, in one number\.\]/.test(a.text), a.text);

    // from chat, on the other divider
    a = await call4("update_plan", { project: dpid, ops: [{ op: "set_agenda_line", tabId: "Ask", agendaLine: "  What we need approved today.  " }] });
    assert.equal(a.sc.project.tabs[3].agendaLine, "What we need approved today.", "trimmed");
    assert.ok(/\[agenda line: What we need approved today\.\]/.test(a.text), a.text);

    // the canvas patches the one row, and an outline save that omits it keeps it
    const div = a.sc.project.tabs[0];
    a = await call4("ui_save_tab_fields", { project: dpid, tabId: div.id, agendaLine: "  Where the quarter landed  " });
    assert.ok(!a.isError, a.text);
    assert.equal(a.sc.project.tabs[0].agendaLine, "Where the quarter landed");
    assert.equal(a.sc.project.tabs[0].name, div.name, "and leaves the section title alone");
    assert.ok(/^Opens on the quarter's result/.test(a.sc.project.tabs[1].story), "and the rows under it keep their stories");
    a = await call4("ui_save_tabs", { project: dpid, tabs: a.sc.project.tabs.map((t: any) => ({ id: t.id, name: t.name, divider: t.divider })) });
    assert.equal(a.sc.project.tabs[0].agendaLine, "Where the quarter landed", "an outline save that omits it keeps it");
    assert.equal(a.sc.project.tabs[3].agendaLine, "What we need approved today.");

    // an empty one clears it, and the summary then says nothing
    a = await call4("ui_save_tab_fields", { project: dpid, tabId: div.id, agendaLine: "" });
    assert.equal(a.sc.project.tabs[0].agendaLine, undefined);
    a = await call4("update_plan", { project: dpid, ops: [{ op: "set_agenda_line", tabId: "Ask", agendaLine: "" }] });
    assert.ok(!/agenda line/.test(a.text), a.text);
    assert.ok(/\[divider: a section slide, no blocks\]/.test(a.text), "the divider marker stays");
  }

  // ---- the headline is the slide's own field, in sentence case (F11, F12, F14)
  d = await call4("update_plan", { project: dpid, ops: [{ op: "set_headline", tabId: "What changed", headline: "churn moved, not price." }] });
  assert.equal(d.sc.project.tabs[2].message, "Churn moved, not price.", "the server capitalises the first letter and leaves the rest");
  assert.ok(/headline: Churn moved, not price\./.test(d.text), d.text);
  d = await call4("update_plan", { project: dpid, ops: [{ op: "set_story", tab: "What changed", story: "One chart, two cohorts, no commentary." }] });
  assert.equal(d.sc.project.tabs[2].story, "One chart, two cohorts, no commentary.");

  // a legacy set_section inserts a divider row rather than writing a field
  d = await call4("update_plan", { project: dpid, ops: [{ op: "set_section", tab: "What changed", section: "Evidence" }] });
  assert.deepEqual(d.sc.project.tabs.map((t: any) => t.name), ["Context", "Where we landed", "Evidence", "What changed", "Ask", "What we need"]);
  assert.equal(d.sc.project.tabs.find((t: any) => t.name === "What changed").section, "Evidence");
  d = await call4("update_plan", { project: dpid, ops: [{ op: "set_section", tab: "What changed", section: "" }] });
  assert.equal(d.sc.project.tabs.length, 5, "clearing it takes the divider above out again");

  // a missing takeaway is flagged once for the whole plan, not once per slide
  d = await call4("update_plan", { project: dpid, ops: [{ op: "set_brief", brief: { takeaway: "" } }] });
  assert.equal(d.sc.warnings.filter((w: any) => w.kind === "no-takeaway").length, 1, JSON.stringify(d.sc.warnings));
  d = await call4("update_plan", { project: dpid, ops: [{ op: "set_brief", brief: { takeaway: "Retention is compounding; fund two more pods" } }] });
  assert.equal(d.sc.warnings.filter((w: any) => w.kind === "no-takeaway").length, 0);
  assert.equal(d.sc.project.briefCard.audience, "The board", "set_brief patches, it does not replace");

  // an old canvas still sending emphasis chips is accepted and they are dropped
  {
    const b = await call4("ui_save_brief", { project: dpid, goal: "Approve the next raise", emphasis: ["retention"] });
    assert.ok(!b.isError, b.text);
    assert.equal(b.sc.project.briefCard.emphasis, undefined);
    assert.equal(b.sc.project.briefCard.goal, "Approve the next raise");
  }

  // layouts: approval gate, mode-specific block types
  d = await call4("propose_layouts", { project: dpid, layouts: [{ tab: "Where we landed", headline: "x", blocks: [] }] });
  assert.ok(d.isError && /not approved/.test(d.text), d.text);
  await call4("update_plan", { project: dpid, ops: [{ op: "approve_tabs" }] });

  // one field on one row: the approval survives, so a headline typed in Layer 2 does not send the user back to Layer 1
  {
    const before = await call4("ui_get_state", { project: dpid });
    assert.equal(before.sc.project.tabsApproved, true);
    const slide = before.sc.project.tabs[2];
    const patched = await call4("ui_save_tab_fields", { project: dpid, tabId: slide.id, message: "Two cohorts, one gap." });
    assert.ok(!patched.isError, patched.text);
    assert.equal(patched.sc.project.tabsApproved, true, "a field patch leaves the outline approved");
    assert.ok(patched.sc.project.tabs.every((t: any) => t.approved), "and leaves every row approved");
    assert.equal(patched.sc.project.tabs[2].message, "Two cohorts, one gap.");
    assert.equal(patched.sc.project.tabs[2].story, slide.story, "the row's other fields are untouched");
    assert.equal(patched.sc.project.tabs[1].message, before.sc.project.tabs[1].message, "the other rows are untouched");
    assert.equal(patched.sc.project.revision, before.sc.project.revision + 1, "one save, one revision");
    const plan = await call4("get_plan", { project: dpid });
    assert.ok(/headline: Two cohorts, one gap\./.test(plan.text), plan.text);
    assert.equal(plan.sc.project.tabsApproved, true);
  }

  // the style block carries the resolved builder tokens, so the canvas draws the real furniture (F5, F16)
  {
    const st = await call4("ui_get_state", { project: dpid });
    assert.equal(st.sc.style.kind, "deck");
    const tk = st.sc.style.tokens;
    assert.ok(tk && tk.palette && tk.typeScale.headline && tk.furniture.headline, `deck tokens missing: ${JSON.stringify(Object.keys(tk ?? {}))}`);
    assert.equal(typeof tk.furniture.kicker.y, "number");
    const shown = await call4("show_plan", { project: dpid });
    assert.ok(shown.sc.style.tokens.furniture.pageNumber, "show_plan carries them too");
  }

  d = await call4("propose_layouts", { project: dpid, layouts: [{ tab: "Where we landed", headline: "x", blocks: [{ type: "inputs", title: "Assumptions" }] }] });
  assert.ok(d.isError && /not available in deck mode/.test(d.text), d.text);
  // a content slide with no headline anywhere is refused by name
  d = await call4("propose_layouts", { project: dpid, layouts: [{ tab: "What we need", blocks: [{ type: "text", title: "The ask" }] }] });
  assert.ok(d.isError && /No headline for: "What we need"/.test(d.text), d.text);

  // ---- a headline block is folded into the slide headline and dropped (F14)
  d = await call4("propose_layouts", { project: dpid, layouts: [{ tab: "Where we landed", blocks: [
    { type: "headline", title: "retention carried the quarter" },
    { type: "cards", title: "Three drivers", content: "one card per driver, each with the number", count: 3, emphasis: true },
    { type: "takeaway", title: "Fund two more pods" },
  ] }] });
  assert.ok(!d.isError, d.text);
  let slide = d.sc.project.tabs.find((t: any) => t.name === "Where we landed");
  assert.equal(slide.message, "Retention carried the quarter", "the headline block became the slide headline, sentence case");
  assert.ok(!slide.blocks.some((b: any) => b.type === "headline"), "and the block is gone");
  assert.ok(/headline block is now the slide headline/.test(d.text), d.text);
  const cards = slide.blocks.find((b: any) => b.type === "cards");
  assert.deepEqual([cards.row, cards.col, cards.rows, cards.cols], [7, 3, 8, 28], "three-cards with a takeaway bar: the card band");
  assert.equal(cards.count, 3); assert.equal(cards.emphasis, true);
  const takeaway = slide.blocks.find((b: any) => b.type === "takeaway");
  assert.deepEqual([takeaway.row, takeaway.col, takeaway.rows, takeaway.cols], [16, 3, 2, 28], "the takeaway bar sits above the footer row");
  assert.ok(/content: one card per driver/.test(d.text), "content descriptions reach Claude");
  assert.ok(slide.blocks.every((b: any) => b.row >= 5 && b.row + b.rows - 1 <= 17 && b.col + b.cols - 1 <= 32), "every block sits in the body band");

  // ---- three cards with no geometry: the archetype solver lays them out (F11)
  d = await call4("propose_layouts", { project: dpid, layouts: [{ tab: "What changed", headline: "Two cohorts, one gap", kicker: "Evidence | Cohorts", blocks: [
    { type: "cards", title: "2023 cohort", content: "retention curve for the 2023 cohort" },
    { type: "cards", title: "2024 cohort", content: "retention curve for the 2024 cohort" },
    { type: "cards", title: "The gap", content: "the spread between them at month 12" },
  ] }] });
  assert.ok(!d.isError, d.text);
  slide = d.sc.project.tabs.find((t: any) => t.name === "What changed");
  assert.deepEqual(slide.blocks.map((b: any) => [b.col, b.row, b.cols, b.rows]), [[3, 7, 8, 10], [13, 7, 8, 10], [23, 7, 8, 10]], "three cards take the three-cards slots");
  assert.equal(slide.kicker, "Evidence | Cohorts");
  assert.equal(slide.message, "Two cohorts, one gap");

  // ---- the furniture band: a block asked for row 1 lands at row 5, and overlaps are pushed apart (F11)
  d = await call4("propose_layouts", { project: dpid, layouts: [{ tab: "What we need", headline: "Fund two more pods this quarter", blocks: [
    { type: "text", title: "The ask", content: "two pods, what they cost", row: 1, col: 3, rows: 4, cols: 14 },
    { type: "stat", title: "The cost", content: "the number", row: 2, col: 3, rows: 4, cols: 14 },
    { type: "table", title: "What they buy", content: "pod by pod", row: 7, col: 3, rows: 6, cols: 28 },
  ] }] });
  assert.ok(!d.isError, d.text);
  slide = d.sc.project.tabs.find((t: any) => t.name === "What we need");
  const at = (title: string) => { const b = slide.blocks.find((x: any) => x.title === title); return [b.col, b.row, b.cols, b.rows]; };
  assert.deepEqual(at("The ask"), [3, 5, 14, 4], "row 1 is furniture: the block starts at row 5");
  assert.equal(slide.blocks.filter((b: any) => b.row < 5 || b.row + b.rows - 1 > 17).length, 0, "nothing sits in the furniture band or on the footer row");
  const pairs = slide.blocks.flatMap((a: any, i: number) => slide.blocks.slice(i + 1).map((b: any) => [a, b]));
  assert.ok(!pairs.some(([a, b]: any[]) => a.row < b.row + b.rows && b.row < a.row + a.rows && a.col < b.col + b.cols && b.col < a.col + a.cols), `blocks still overlap: ${JSON.stringify(slide.blocks.map((b: any) => [b.title, b.col, b.row, b.cols, b.rows]))}`);
  assert.equal(d.sc.warnings.filter((w: any) => w.kind === "overlap").length, 0, "so no overlap warning survives the save");

  // a hand edit from the canvas is clamped the same way
  {
    const dragged = slide.blocks.map((b: any) => (b.title === "The cost" ? { ...b, row: 2, col: 3 } : b));
    const saved = await call4("ui_save_layout", { project: dpid, tab: slide.id, blocks: dragged });
    const after = saved.sc.project.tabs.find((t: any) => t.name === "What we need");
    assert.ok(after.blocks.every((b: any) => b.row >= 5 && b.row + b.rows - 1 <= 17), "a dragged block cannot sit in the furniture band");
  }

  // a divider row needs no layout, and the plan is ready once the content slides are approved
  d = await call4("ui_approve", { project: dpid, what: "all_layouts" });
  assert.equal(d.sc.stage, "ready", `dividers must not hold the plan in layouts: ${JSON.stringify(d.sc.project.tabs.map((t: any) => [t.name, t.divider, t.layoutApproved]))}`);
  assert.equal(d.sc.warnings.filter((w: any) => w.kind === "no-headline").length, 0, JSON.stringify(d.sc.warnings));
  assert.equal(d.sc.warnings.filter((w: any) => w.kind === "empty-tab").length, 0, "a divider is not an empty slide");

  // a content slide whose headline is cleared is not ready to build
  d = await call4("ui_save_tab_fields", { project: dpid, tabId: d.sc.project.tabs.find((t: any) => t.name === "What changed").id, message: "" });
  const missing = d.sc.warnings.find((w: any) => w.kind === "no-headline");
  assert.ok(missing && missing.message === "Headline missing" && missing.tabId, JSON.stringify(d.sc.warnings));

  // ---- build options: the fidelity setting and the generated slides (Q3, Q5, Q6)
  {
    const before = await call4("get_plan", { project: dpid });
    assert.equal(before.sc.project.build, undefined, "a plan carries no build options until one is set");
    assert.deepEqual(before.sc.buildOptions, { fidelity: "balanced", notesFromStory: true, cover: true, agenda: true, comments: false }, "get_plan reports the defaults under buildOptions");
    assert.ok(!/^build: /m.test(before.text), "and the summary says nothing while they are all default");

    let o = await call4("ui_set_build_options", { project: dpid, fidelity: "strict", notesFromStory: false });
    assert.ok(!o.isError, o.text);
    assert.deepEqual(o.sc.project.build, { fidelity: "strict", notesFromStory: false }, "the canvas patches only the fields it sent");
    assert.ok(o.sc.project.id === dpid && Array.isArray(o.sc.warnings) && typeof o.sc.stage === "string", "and gets the usual project, warnings and stage back");
    // the canvas keeps the build summary in `build`, so the options ride under their own key
    assert.deepEqual(o.sc.buildOptions, { fidelity: "strict", notesFromStory: false, cover: true, agenda: true, comments: false }, "the filled-in options come back as buildOptions");
    assert.equal(o.sc.build, undefined, "and nothing lands on `build`, which would blank the built view");

    o = await call4("update_plan", { project: dpid, ops: [{ op: "set_build_options", cover: false }] });
    assert.deepEqual(o.sc.project.build, { fidelity: "strict", notesFromStory: false, cover: false }, "a chat edit patches too, it does not replace");
    assert.ok(/^build: strict, no notes from story, no cover, agenda, no comments$/m.test(o.text), o.text.split("\n").find((l: string) => l.startsWith("build:")) ?? o.text);

    const after = await call4("get_plan", { project: dpid });
    assert.deepEqual(after.sc.buildOptions, { fidelity: "strict", notesFromStory: false, cover: false, agenda: true, comments: false }, "get_plan fills the untouched field in from the defaults");

    // back on every default: the line goes away again
    o = await call4("update_plan", { project: dpid, ops: [{ op: "set_build_options", fidelity: "balanced", notesFromStory: true, cover: true, agenda: true }] });
    assert.ok(!/^build: /m.test(o.text), o.text);
    o = await call4("update_plan", { project: dpid, ops: [{ op: "set_build_options", fidelity: "loose" }] });
    assert.ok(/^build: loose, notes from story, cover, agenda, no comments$/m.test(o.text), o.text);
    await call4("update_plan", { project: dpid, ops: [{ op: "set_build_options", fidelity: "balanced" }] });
  }

  // ---- a placeholder is never shipped: the outline and the layouts flag it (research item 4).
  // "Logo, top-left" is a real image someone supplies, so it stays a placeholder and warns; a
  // placeholder that names a relationship is promoted to a diagram instead (the next block).
  {
    const state = await call4("ui_get_state", { project: dpid });
    const slide = state.sc.project.tabs.find((t: any) => t.name === "What we need");
    let ph = await call4("update_plan", { project: dpid, ops: [{ op: "add_block", tab: slide.id, block: { type: "placeholder", title: "Logo, top-left" } }] });
    const block = ph.sc.project.tabs.find((t: any) => t.id === slide.id).blocks.find((b: any) => b.type === "placeholder");
    const w = ph.sc.warnings.find((x: any) => x.kind === "placeholder");
    assert.ok(w, JSON.stringify(ph.sc.warnings.map((x: any) => x.kind)));
    assert.equal(w.message, "placeholder: make it a diagram, image or drop it");
    assert.equal(w.tabId, slide.id); assert.equal(w.blockId, block.id, "the warning points at the block");
    assert.ok(/placeholder: make it a diagram, image or drop it/.test(ph.text), "and reaches Claude in the summary");
    ph = await call4("update_plan", { project: dpid, ops: [{ op: "remove_block", blockId: block.id }] });
    assert.equal(ph.sc.warnings.filter((x: any) => x.kind === "placeholder").length, 0, "and goes when the block does");
  }

  // ---- a relationship is a diagram, not a dashed box with a sentence in it (research item 4)
  {
    const state = await call4("ui_get_state", { project: dpid });
    const slide = state.sc.project.tabs.find((t: any) => t.name === "What we need");

    // propose_layouts: the flow is promoted, the logo is left alone
    let pl = await call4("propose_layouts", { project: dpid, layouts: [{ tab: slide.id, headline: "Fund two more pods this quarter", blocks: [
      { type: "placeholder", title: "Handoff flow: Agency → client → team" },
      { type: "placeholder", title: "Logo, top-left" },
    ] }] });
    assert.ok(!pl.isError, pl.text);
    let laid = pl.sc.project.tabs.find((t: any) => t.id === slide.id);
    const flow = laid.blocks.find((b: any) => b.type === "diagram");
    assert.ok(flow, `the placeholder should be a diagram: ${JSON.stringify(laid.blocks.map((b: any) => [b.type, b.title]))}`);
    assert.equal(flow.title, "Handoff flow: Agency → client → team", "the title the user wrote survives");
    assert.equal(flow.diagram.kind, "flow");
    assert.equal(flow.diagram.direction, "LR", "a flow reads left to right");
    assert.deepEqual(flow.diagram.nodes.map((n: any) => n.label), ["Agency", "Client", "Team"], "three nodes read off the arrows");
    assert.deepEqual(flow.diagram.edges.map((e: any) => [e.from, e.to]), [[flow.diagram.nodes[0].id, flow.diagram.nodes[1].id], [flow.diagram.nodes[1].id, flow.diagram.nodes[2].id]], "two edges, in order");
    assert.ok(/promoted 1 placeholder to diagrams: Handoff flow: Agency → client → team/.test(pl.text), pl.text);
    const logo = laid.blocks.find((b: any) => b.title === "Logo, top-left");
    assert.equal(logo.type, "placeholder", "a real image someone supplies stays a placeholder");
    assert.equal(logo.diagram, undefined);
    assert.equal(pl.sc.warnings.filter((x: any) => x.kind === "placeholder").length, 1, "and is the only placeholder left to warn about");

    // add_block from chat: a 2x2 is a matrix, and the axes come off the "vs"
    const add = await call4("update_plan", { project: dpid, ops: [{ op: "add_block", tab: slide.id, block: { type: "placeholder", title: "2x2 of cost vs speed" } }] });
    assert.ok(!add.isError, add.text);
    laid = add.sc.project.tabs.find((t: any) => t.id === slide.id);
    const matrix = laid.blocks.find((b: any) => b.title === "2x2 of cost vs speed");
    assert.equal(matrix.type, "diagram");
    assert.equal(matrix.diagram.kind, "matrix");
    assert.deepEqual(matrix.diagram.nodes.map((n: any) => n.label), ["Cost", "Speed", "High cost, high speed", "High cost, low speed", "Low cost, high speed", "Low cost, low speed"], "the two axis labels first, then the four cells build_deck wants");
    assert.deepEqual(matrix.diagram.edges, [], "a matrix draws axes and cells, no arrows");
    assert.ok(/promoted 1 placeholder to diagrams: 2x2 of cost vs speed/.test(add.text), add.text);

    // update_block from chat: the relationship arrives in the patch, so the type follows
    const patched = await call4("update_plan", { project: dpid, ops: [{ op: "update_block", blockId: logo.id, patch: { title: "Rollout timeline", content: "discover, pilot, scale" } }] });
    assert.ok(!patched.isError, patched.text);
    const timeline = patched.sc.project.tabs.find((t: any) => t.id === slide.id).blocks.find((b: any) => b.id === logo.id);
    assert.equal(timeline.type, "diagram");
    assert.equal(timeline.diagram.kind, "timeline");
    assert.deepEqual(timeline.diagram.nodes.map((n: any) => n.label), ["Discover", "Pilot", "Scale"]);
    assert.ok(/promoted 1 placeholder to diagrams: Rollout timeline/.test(patched.text), patched.text);
    assert.equal(patched.sc.warnings.filter((x: any) => x.kind === "placeholder").length, 0, "no placeholder is left on the slide");
  }

  await c4.close();
  console.log("DECK OUTLINE OK");
}

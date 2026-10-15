/**
 * End-to-end deck flow over stdio: start_project kind deck -> propose_outline -> propose_layouts ->
 * approvals and notes on the canvas -> build_deck, then the .pptx, the snapshot and the deck check
 * are read back, and every refusal is exercised.
 * Run: npx tsx test/deck-flow.test.ts  (from the project root, after npm run build)
 */
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const home = fs.mkdtempSync(path.join(os.tmpdir(), "blueprint-deck-"));
const out = path.join(home, "out");
const client = new Client({ name: "deck-flow-test", version: "0" });
await client.connect(new StdioClientTransport({ command: "node", args: ["dist/server/main.js"], env: { ...process.env, BLUEPRINT_HOME: home, BLUEPRINT_OUTPUT: out } }));

const call = async (name: string, args: Record<string, unknown>) => {
  const r: any = await client.callTool({ name, arguments: args });
  // A result may carry image blocks beside the text; keep the two apart so a line count means something.
  const texts: string[] = (r.content ?? []).filter((c: any) => c.type === "text").map((c: any) => c.text as string);
  const images: any[] = (r.content ?? []).filter((c: any) => c.type === "image");
  return { r, text: texts.join("\n"), texts, images, sc: r.structuredContent as any, isError: !!r.isError };
};

// ---- the tool is listed, model-visible, renders the canvas, and shows the content schema
const tools = await client.listTools();
const bd = tools.tools.find((t) => t.name === "build_deck");
assert.ok(bd, "build_deck is listed");
assert.match(String((bd!._meta as any)?.ui?.resourceUri), /^ui:\/\/blueprint\/canvas-\d+\.\d+\.\d+\.html$/);
assert.ok(bd!.description!.split(/(?<=[.!?])\s+/).length <= 3, `tool description is ${bd!.description!.split(/(?<=[.!?])\s+/).length} sentences`);
const schema = JSON.stringify(bd!.inputSchema);
for (const key of ["header", "attribution", "categories", "headers", "placeholder", "takeaway", "diagram", "nodes", "edges"]) assert.ok(schema.includes(key), `the ${key} content field is exposed to the model`);
assert.ok(/3 to 7 nodes/.test(bd!.description!) && /one idea per diagram/i.test(bd!.description!), "the diagram rules ride on the description");
assert.ok(/never hand-build|do not hand-build/i.test(bd!.description!) && /update_plan/.test(bd!.description!), "and so does what to do about a finding");

// ---- 1. the plan
let s = await call("start_project", {
  name: "Referral Growth Review", kind: "deck", formatting: "house",
  brief: "Working review of referral growth for the partnership leadership team",
  briefCard: { audience: "Partnership leadership team", goal: "Agree the two referral moves we fund next", takeaway: "Referrals stall at the ask, not at the relationship", emphasis: ["conversion gap", "the ask moment"] },
});
assert.ok(!s.isError, s.text);
const pid = s.sc.projectId as string;
assert.equal(s.sc.kind, "deck");

s = await call("propose_outline", { project: pid, rows: [
  { name: "Cover", story: "Sets the room: a working review of referral growth, not a status update." },
  { name: "Agenda", story: "Frames the session in one sentence and names the four stops." },
  { name: "Why referrals", divider: true },
  { name: "Where referrals stall", story: "Names the three failure modes so the room stops arguing about whether partners care. Each one has its own fix." },
  { name: "Where growth comes from", story: "Puts the channel mix on screen so funding is argued over two channels, not six. The table gives the team the moments to script." },
  { name: "What we do next", divider: true },
  { name: "The ask", story: "Closes on the decision: two moves, an owner each, nothing else." },
] });
assert.ok(!s.isError, s.text);
assert.equal(s.sc.project.tabs.length, 7);
assert.deepEqual(s.sc.project.tabs.filter((t: any) => t.divider).map((t: any) => t.name), ["Why referrals", "What we do next"]);
assert.deepEqual(s.sc.project.tabs.map((t: any) => t.section), ["", "", "Why referrals", "Why referrals", "Why referrals", "What we do next", "What we do next"], "sections come from the dividers");
assert.ok(/5 slides and 2 dividers/.test(s.text), s.text);

// layouts and the build both wait for the outline
s = await call("build_deck", { project: pid, slides: [{ tabId: "Cover", headline: "x" }] });
assert.ok(s.isError && /outline is not approved/.test(s.text), s.text);
await call("ui_approve", { project: pid, what: "tabs" });

s = await call("propose_layouts", { project: pid, layouts: [
  { tab: "Cover", headline: "Growing partnerships through referrals", blocks: [] },
  { tab: "Agenda", headline: "We agree where referral growth comes from and what we change next", blocks: [
    { type: "bullets", title: "Four stops", content: "one line per stop, in order", row: 9, col: 2, rows: 4, cols: 28 },
    { type: "stat", title: "The prize", content: "three numbers the team already trusts", count: 3, row: 14, col: 2, rows: 3, cols: 28, questions: ["Is the 4.1x conversion figure blended or partner-sourced only?"] },
  ] },
  { tab: "Where referrals stall", headline: "referrals stall at three moments, and each one has a different fix", kicker: "Strategic Rationale | Why Referrals", blocks: [
    { type: "cards", title: "Three stall moments", content: "one card per failure mode, named and explained", count: 3, emphasis: true, row: 7, col: 3, rows: 8, cols: 28 },
    { type: "takeaway", title: "Fix the ask", content: "the one line the room should leave with", row: 16, col: 3, rows: 2, cols: 28 },
  ] },
  { tab: "Where growth comes from", headline: "Two channels drive nine of every ten introductions", blocks: [
    { type: "chart", title: "Channel mix", content: "introductions by channel, last year against this year", chartKind: "bar", row: 7, col: 3, rows: 10, cols: 14 },
    { type: "table", title: "Scripted moments", content: "the moment, the owner and the prompt", row: 7, col: 18, rows: 5, cols: 14 },
    { type: "text", title: "What it means", content: "two sentences on what to fund", row: 13, col: 18, rows: 4, cols: 14 },
  ] },
  { tab: "The ask", headline: "Pick the two moves we fund this quarter and name an owner for each", blocks: [
    { type: "quote", title: "Partner voice", content: "the line from the partner interviews", row: 7, col: 4, rows: 4, cols: 26 },
    { type: "placeholder", title: "Next steps strip", content: "a designed four-step strip", row: 12, col: 4, rows: 5, cols: 26 },
  ] },
] });
assert.ok(!s.isError, s.text);
const P = s.sc.project;
const tabOf = (name: string) => P.tabs.find((t: any) => t.name === name);
const bid = (tab: string, type: string) => tabOf(tab).blocks.find((b: any) => b.type === type).id;
const ids = {
  bullets: bid("Agenda", "bullets"), stat: bid("Agenda", "stat"),
  cards: bid("Where referrals stall", "cards"), takeaway: bid("Where referrals stall", "takeaway"),
  chart: bid("Where growth comes from", "chart"), table: bid("Where growth comes from", "table"), text: bid("Where growth comes from", "text"),
  quote: bid("The ask", "quote"), placeholder: bid("The ask", "placeholder"),
};
const cardsBlock = tabOf("Where referrals stall").blocks.find((b: any) => b.type === "cards");
assert.deepEqual([cardsBlock.col, cardsBlock.row, cardsBlock.cols, cardsBlock.rows], [3, 7, 28, 8], "the plan keeps the geometry it was given");
assert.equal(tabOf("Where referrals stall").message, "Referrals stall at three moments, and each one has a different fix", "the headline is the slide's own field, sentence case");
assert.equal(tabOf("Where referrals stall").kicker, "Strategic Rationale | Why Referrals");
assert.equal(tabOf("Why referrals").blocks.length, 0, "a divider takes no blocks");
assert.ok(P.tabs.every((t: any) => t.blocks.every((b: any) => b.row >= 5 && b.row + b.rows - 1 <= 17)), "every block sits in rows 5 to 17");
const qNote = s.sc.project.notes.find((n: any) => n.kind === "question");
assert.ok(qNote && qNote.status === "open", "the question on the stat block became an open thread");

// ---- 2. the payload: real content per plan block
const baseSlides = () => JSON.parse(JSON.stringify([
  { tabId: tabOf("Cover").id, kind: "title", subtitle: "Partnership growth review", date: "Draft for discussion", blocks: [] },
  { tabId: tabOf("Agenda").id, kind: "agenda", blocks: [
    { blockId: ids.bullets, content: { items: [
      "Why referrals: the conversion gap we are leaving on the table",
      "Where they stall: three moments, each with a different fix",
      { text: "What the data says: two channels carry the volume", children: ["Events are down on purpose"] },
      "What we decide: two moves, one owner each",
    ] } },
    { blockId: ids.stat, content: { items: [
      { value: "4.1x", label: "Conversion versus a cold list" },
      { value: "38%", label: "Of new partners came from a referral" },
      { value: "11 days", label: "Shorter from contact to signed" },
    ] } },
  ] },
  { tabId: tabOf("Where referrals stall").id,
    notes: "Ask the room which of the three they recognise before showing the takeaway bar.", blocks: [
      { blockId: ids.cards, content: { items: [
        { header: "No moment", sub: "Timing", body: "Partners meet the right person weeks before they think of us. Nothing in the meeting flow prompts an introduction, so the moment goes by." },
        { header: "No words", sub: "Language", body: "When the moment comes, partners describe us in their own words. The introduction lands flat and the other side cannot tell why it matters." },
        { header: "No follow", sub: "Hand-off", body: "Introductions arrive in a partner inbox and sit there. Nobody owns the first reply, so a warm lead cools to the temperature of a cold one." },
      ] } },
      { blockId: ids.takeaway, content: { text: "Fix the ask, not the relationship: a scripted moment beats another reminder email" } },
    ] },
  { tabId: tabOf("Where growth comes from").id, kicker: "Evidence | Where Growth Comes From",
    headline: "Two channels drive nine of every ten introductions", source: "Partner CRM", blocks: [
      { blockId: ids.chart, content: { kind: "column", categories: ["Events", "Existing partners", "Advisors", "Inbound"], series: [
        { name: "Last year", values: [22, 61, 12, 9] },
        { name: "This year", values: [18, 84, 14, 7] },
      ] } },
      { blockId: ids.table, content: { headers: ["Moment", "Owner", "Prompt"], rows: [
        ["Kickoff call", "Partner lead", "Who else should hear this"],
        ["First win", "Account lead", "Name one peer"],
        ["Quarterly review", "Partner lead", "Two introductions"],
        ["Renewal", "Account lead", "Warm hand-off"],
      ] } },
      { blockId: ids.text, content: { paragraphs: [
        "Existing partners grew 38% this year while events fell.",
        { runs: [{ text: "Fund the two that compound", bold: true }, { text: "; stop staffing the rest." }] },
      ] } },
    ] },
  { tabId: tabOf("What we do next").id, subtitle: "Three moves, one owner each.", blocks: [] },
  { tabId: tabOf("The ask").id, kicker: "Close | The Ask", blocks: [
      { blockId: ids.quote, content: { text: "We never stopped wanting to refer you. Nobody ever asked us at a moment when it was easy.", attribution: "Partner interview, regional network" } },
      { blockId: ids.placeholder, content: { label: "Designed next-steps strip: Decide, Fund, Assign, Review" } },
    ] },
]));

// the build waits for the layouts, then for the threads
s = await call("build_deck", { project: pid, slides: baseSlides() });
assert.ok(s.isError && /Layouts not yet approved/.test(s.text), s.text);
await call("ui_approve", { project: pid, what: "all_layouts" });
await call("ui_note", { project: pid, action: "add", tabId: tabOf("Where referrals stall").id, blockId: ids.cards, text: "Keep the three stall moments in this order." });
await call("update_plan", { project: pid, ops: [{ op: "add_note", tab: "The ask", text: "Add a fourth move on pricing", kind: "instruction" }] });
s = await call("build_deck", { project: pid, slides: baseSlides() });
assert.ok(s.isError && /thread\(s\) are still open/.test(s.text), s.text);

// the user answers Claude's question on the canvas; Claude closes the two instructions
s = await call("ui_note", { project: pid, action: "answer", noteId: qNote.id, text: "Partner-sourced only." });
assert.equal(s.sc.project.notes.find((n: any) => n.id === qNote.id).status, "resolved");
const openNotes = s.sc.project.notes.filter((n: any) => n.status === "open");
assert.equal(openNotes.length, 2);
s = await call("resolve_notes", { project: pid, resolutions: [
  { noteId: openNotes.find((n: any) => n.blockId === ids.cards).id, status: "resolved", response: "Cards run No moment, No words, No follow, in that order." },
  { noteId: openNotes.find((n: any) => !n.blockId).id, status: "declined", response: "Pricing is a separate decision; this deck stops at two moves." },
] });
assert.ok(/0 note\(s\) still open/.test(s.text), s.text);

// ---- 3. the build, with the plan's build options behind it
s = await call("update_plan", { project: pid, ops: [{ op: "set_build_options", fidelity: "loose" }] });
assert.ok(!s.isError, s.text);
assert.equal(s.sc.project.build.fidelity, "loose", "fidelity is a plan setting, not a payload field");

s = await call("build_deck", { project: pid, slides: baseSlides() });
assert.ok(!s.isError, s.text);
const buildText = s.text as string;
assert.equal(s.sc.view, "built");
assert.equal(s.sc.stage, "built");
assert.ok(s.texts[0].split("\n").length <= 12, `result text is ${s.texts[0].split("\n").length} lines:\n${s.texts[0]}`);
assert.match(s.text, /Notes honoured: 2 resolved, 1 declined/);
assert.match(s.texts[0].split("\n").pop()!, /^Next: /, "the findings text ends with the call to make next");
if (/Deck check: [1-9]/.test(s.text) || s.sc.build.findings.length) {
  assert.ok(/Fix these on the plan and call build_deck again/.test(s.text) && /do not hand-build/.test(s.text), s.text);
  assert.ok(s.text.split("\n").filter((l: string) => l.startsWith("- slide ")).length <= 12, "the finding list is capped at 12");
}
assert.ok(s.sc.style.tokens.furniture.headline, "the built render carries the style tokens too");
assert.ok(s.sc.project.specPath && fs.existsSync(s.sc.project.specPath), "the builder spec is kept for the next diff");

// the four build options ride in the spec the builder was handed
const spec = JSON.parse(fs.readFileSync(s.sc.project.specPath as string, "utf8"));
for (const key of ["fidelity", "notesFromStory", "cover", "agenda"]) assert.ok(key in spec, `the build spec carries ${key}: ${Object.keys(spec).join(", ")}`);
assert.equal(spec.fidelity, "loose", "the plan's fidelity reaches the builder");
assert.equal(spec.notesFromStory, true);
assert.equal(spec.cover, false, "this outline brings its own cover slide, so the builder is not asked for one");
assert.equal(spec.agenda, false, "and its own agenda slide");
const built = s.sc.project.buildPath as string;
assert.equal(built, path.join(out, "Referral Growth Review.pptx"));
assert.ok(fs.existsSync(built), "pptx written");
const snapPath = s.sc.project.snapshotPath as string;
assert.equal(snapPath, built.replace(/\.pptx$/, ".snapshot.json"));
assert.ok(fs.existsSync(snapPath), "snapshot written");

// buildSummary: the shape the canvas built view reads
const summary = s.sc.build;
assert.equal(summary.kind, "deck");
assert.equal(summary.slides.length, 7);
assert.deepEqual(summary.slides.map((x: any) => x.index), [1, 2, 3, 4, 5, 6, 7]);
assert.deepEqual(summary.slides.map((x: any) => x.tabId), P.tabs.map((t: any) => t.id));
assert.deepEqual(summary.slides.map((x: any) => x.name), ["Cover", "Agenda", "Why referrals", "Where referrals stall", "Where growth comes from", "What we do next", "The ask"]);
assert.equal(summary.slides[3].headline, "Referrals stall at three moments, and each one has a different fix", "the plan's headline is built when the payload leaves it out");
assert.equal(summary.slides[2].headline, "Why referrals", "a divider slide reads as its row name, and needed no payload entry");
assert.equal(summary.slides[5].headline, "What we do next");
assert.ok(summary.slides.every((x: any) => x.shapes > 0), JSON.stringify(summary.slides));
assert.ok(Array.isArray(summary.findings) && Array.isArray(summary.warnings), "findings and warnings are lists");
const realWarnings = summary.warnings.filter((w: string) => !/LibreOffice|pdftoppm|PNG/i.test(w));
assert.deepEqual(realWarnings, [], `builder warnings: ${JSON.stringify(summary.warnings)}`);
// A deck-scope finding is about the one-size-per-role type, not about a block on a slide.
const bad = summary.findings.filter((f: any) => f.kind !== "overflow" && f.scope !== "deck");
assert.equal(bad.length, 0, `deck check findings beyond overflow: ${JSON.stringify(bad)}`);

// what the fidelity pass moved, and the sizes the builder pinned for the whole deck
assert.ok(Array.isArray(summary.adjusted), `the build summary carries the adjustments: ${JSON.stringify(summary.adjusted)}`);
assert.ok(summary.adjusted.length <= 60, "the adjustment list is capped at 60");
for (const a of summary.adjusted) {
  assert.equal(typeof a.slide, "number", `an adjustment names its slide: ${JSON.stringify(a)}`);
  assert.ok(a.from && a.to && a.reason, `an adjustment carries from, to and why: ${JSON.stringify(a)}`);
  assert.ok(summary.slides.some((r: any) => r.index === a.slide), `adjustment on slide ${a.slide}, which the built deck does not have`);
}
if (summary.adjusted.length) assert.match(buildText, /^Adjusted \d+ blocks? for fit \((strict|balanced|loose)\); see the Built view$/m);
else assert.ok(!/^Adjusted /m.test(buildText), "nothing moved, so nothing is claimed");
if (!summary.typeScale) console.log("SKIP typeScale: this builder reports no deck-wide type sizes");
else {
  assert.equal(typeof summary.typeScale.headline, "number", `a deck-wide headline size: ${JSON.stringify(summary.typeScale)}`);
  const line = buildText.split("\n").find((l: string) => l.startsWith("Type: "));
  assert.ok(line, `the result says what type the deck reads at:\n${buildText}`);
  assert.ok(line!.startsWith(`Type: headline ${summary.typeScale.headline}`), line!);
}
if (summary.pdfPath) assert.ok(fs.existsSync(summary.pdfPath), "the PDF is where the result says");
for (const t of summary.slides) if (t.thumbnail) assert.ok(fs.existsSync(t.thumbnail), `thumbnail missing: ${t.thumbnail}`);

// ---- 3b. the slides ride back as images the model can look at
const thumbs = (s.sc.project.thumbnails ?? []) as string[];
const pngHead = (data: string) => Buffer.from(data, "base64").subarray(0, 4).toString("hex");
if (!thumbs.length) {
  console.log(`SKIP slide images: the build wrote no slide PNGs (${summary.warnings.join("; ") || "no LibreOffice or pdftoppm on this machine"})`);
  assert.equal(s.images.length, 0, "no PNGs, no image blocks");
} else {
  assert.ok(s.images.length >= 1, "the build result carries slide images");
  assert.equal(s.images.length, Math.min(summary.slides.length, 12), "one image per slide, 12 at most");
  for (const img of s.images) {
    assert.equal(img.mimeType, "image/png");
    assert.deepEqual(img.annotations?.audience, ["assistant"], "the images are addressed to the model");
    assert.equal(img.annotations?.priority, 0.6);
    assert.equal(pngHead(img.data), "89504e47", "each block carries real PNG bytes");
    assert.ok(Buffer.from(img.data, "base64").length <= 400 * 1024, `a slide image is ${Buffer.from(img.data, "base64").length} bytes`);
  }
  assert.equal(s.texts.length, 2, `the text is the findings plus one line about the images: ${JSON.stringify(s.texts)}`);
  assert.match(s.texts[1], /^Slide images follow: look at each, list what a reader would notice, fix it through the plan, rebuild\.$/);
  const totalBase64 = s.images.reduce((n: number, i: any) => n + i.data.length, 0);
  assert.ok(totalBase64 <= 3 * 1024 * 1024, `the images are ${Math.round(totalBase64 / 1024)} KB of base64`);
  console.log(`slide images: ${s.images.length}, ${s.images.map((i: any) => Math.round(Buffer.from(i.data, "base64").length / 1024) + "KB").join(" ")}`);
}

// the snapshot lists one shape per block (a cards block, one per card)
const snap = JSON.parse(fs.readFileSync(snapPath, "utf8"));
assert.equal(snap.slides.length, 7);
assert.deepEqual(snap.grid, { cols: 32, rows: 18, cell: 30 });
const shapes = snap.slides.flatMap((x: any) => x.shapes);
for (const [name, id] of Object.entries(ids)) assert.ok(shapes.some((sh: any) => sh.blockId === id), `no snapshot shape for the ${name} block`);
assert.equal(shapes.filter((sh: any) => sh.blockId === ids.cards).length, 3, "one snapshot entry per card");
assert.ok(shapes.some((sh: any) => sh.blockId === ids.chart && sh.kind === "chart"), "the chart is recorded as a chart");
assert.ok(shapes.some((sh: any) => sh.blockId === ids.table && sh.kind === "table"), "the table is recorded as a table");
assert.equal(s.sc.snapshot.name, "Referral Growth Review", "the snapshot reaches the canvas");

// ---- 3c. fidelity on the call overrides the plan for that build, and only for that build
{
  const once = await call("build_deck", { project: pid, slides: baseSlides(), fidelity: "strict" });
  assert.ok(!once.isError, once.text);
  assert.equal(JSON.parse(fs.readFileSync(once.sc.project.specPath as string, "utf8")).fidelity, "strict", "the payload's fidelity reaches the builder");
  assert.equal(once.sc.project.build.fidelity, "loose", "and the plan keeps its own setting");
  assert.match(JSON.stringify((await client.listTools()).tools.find((t) => t.name === "build_deck")!.inputSchema), /fidelity/, "fidelity is exposed on the call");
}

// ---- 4. the file itself, through python-pptx
const probe = JSON.parse(execFileSync("python3", ["-c", `
import json, sys
from pptx import Presentation
EMU = 12700.0
prs = Presentation(sys.argv[1])
want_x, want_y, head = float(sys.argv[2]), float(sys.argv[3]), sys.argv[4]
slides = list(prs.slides)
def texts(slide):
    return [sh.text_frame.text for sh in slide.shapes if sh.has_text_frame]
cards = [max(abs(sh.left / EMU - want_x), abs(sh.top / EMU - want_y))
         for sh in slides[3].shapes if sh.has_text_frame]
print(json.dumps({
    "slides": len(slides),
    "size": [prs.slide_width / EMU, prs.slide_height / EMU],
    "headlineOnSlide4": any(head in t for t in texts(slides[3])),
    "charts": sum(1 for s in slides for sh in s.shapes if sh.has_chart),
    "tables": sum(1 for s in slides for sh in s.shapes if sh.has_table),
    "cardOffset": min(cards) if cards else 999.0,
    "notes4": slides[3].notes_slide.notes_text_frame.text if slides[3].has_notes_slide else "",
    "cover": texts(slides[0]),
}))
`, built, String((cardsBlock.col - 1) * 30), String((cardsBlock.row - 1) * 30), "Referrals stall at three moments"]).toString());
assert.equal(probe.slides, 7, "seven slides in the deck");
assert.deepEqual(probe.size, [960, 540]);
assert.ok(probe.headlineOnSlide4, `slide 4 headline missing: ${JSON.stringify(probe)}`);
assert.ok(probe.charts >= 1, "a native chart exists");
assert.ok(probe.tables >= 1, "a native table exists");
assert.ok(probe.cardOffset <= 1, `the first card sits ${probe.cardOffset} pt from its plan cell`);
assert.match(probe.notes4, /which of the three they recognise/);
assert.ok(probe.cover.some((t: string) => t.includes("Draft for discussion")), `title slide furniture: ${JSON.stringify(probe.cover)}`);

// ---- 5. refusals
{
  const wrong = baseSlides();
  wrong[2].blocks[0] = { blockId: ids.cards, type: "chart", content: { kind: "bar", categories: ["a"], series: [{ name: "s", values: [1] }] } };
  const r = await call("build_deck", { project: pid, slides: wrong });
  assert.ok(r.isError && /Where referrals stall/.test(r.text) && /the plan has it as a cards block, the payload says chart/.test(r.text), r.text);
}
{
  // the payload may leave the headline out because the plan holds it; with neither, the build refuses
  await call("update_plan", { project: pid, ops: [{ op: "set_headline", tabId: "Where referrals stall", headline: "" }] });
  const r = await call("build_deck", { project: pid, slides: baseSlides() });
  assert.ok(r.isError && /Slide "Where referrals stall".*needs a headline/.test(r.text) && /set_headline/.test(r.text), r.text);
  const back = await call("update_plan", { project: pid, ops: [{ op: "set_headline", tabId: "Where referrals stall", headline: "referrals stall at three moments, and each one has a different fix" }] });
  assert.equal(back.sc.project.tabs.find((t: any) => t.name === "Where referrals stall").message, "Referrals stall at three moments, and each one has a different fix");
}
{
  const short = baseSlides();
  short[2].blocks.pop();
  const r = await call("build_deck", { project: pid, slides: short });
  assert.ok(r.isError && /is in the approved plan but missing from the build/.test(r.text), r.text);
}
{
  await call("update_plan", { project: pid, ops: [{ op: "add_note", tab: "Cover", text: "Use the new wordmark", kind: "instruction" }] });
  const r = await call("build_deck", { project: pid, slides: baseSlides() });
  assert.ok(r.isError && /thread\(s\) are still open/.test(r.text), r.text);
}
{
  const w = await call("start_project", { name: "Acme Sizing", kind: "workbook", brief: "A workbook, not a deck" });
  const r = await call("build_deck", { project: w.sc.projectId, slides: [{ tabId: "Assumptions", headline: "x" }] });
  assert.ok(r.isError && /is a workbook plan, not a deck/.test(r.text), r.text);
}

// ---- 6. diagram blocks: the plan carries the sketch, the build carries the real nodes (F18)
{
  const open = (await call("get_plan", { project: pid })).sc.project.notes.filter((n: any) => n.status === "open");
  if (open.length) await call("resolve_notes", { project: pid, resolutions: open.map((n: any) => ({ noteId: n.id, status: "resolved", response: "Handled in this pass." })) });

  const node = (i: number) => ({ id: `n${i}`, label: `Step ${i}` });
  const flow = {
    kind: "flow", direction: "LR",
    nodes: [
      { id: "n1", label: "Partner meeting" }, { id: "n2", label: "Scripted ask" }, { id: "n3", label: "Warm introduction" },
      { id: "n4", label: "First reply" }, { id: "n5", label: "Signed partner" },
    ],
    edges: [
      { from: "n1", to: "n2", label: "prompts" }, { from: "n2", to: "n3", label: "sends" },
      { from: "n3", to: "n4", label: "owned by us" }, { from: "n4", to: "n5", label: "closes" },
    ],
  };

  let d = await call("propose_layouts", { project: pid, layouts: [{ tab: "The ask", headline: "Pick the two moves we fund this quarter and name an owner for each", blocks: [
    { type: "diagram", title: "How a referral moves", content: "the five steps from the moment to a signed partner", diagram: flow, row: 7, col: 3 },
  ] }] });
  assert.ok(!d.isError, d.text);
  const dg = d.sc.project.tabs.find((t: any) => t.name === "The ask").blocks[0];
  assert.equal(dg.type, "diagram");
  assert.deepEqual([dg.col, dg.row, dg.cols, dg.rows], [3, 7, 20, 9], "a diagram block defaults to 20 by 9 cells");
  assert.equal(dg.diagram.nodes.length, 5, "the plan keeps the sketch");
  assert.ok(/diagram diagram, 5 nodes|flow diagram, 5 nodes/.test(d.text), d.text);
  await call("ui_approve", { project: pid, what: "all_layouts" });

  const withDiagram = (content: unknown) => baseSlides().map((x: any) => (x.tabId === tabOf("The ask").id ? { ...x, blocks: [{ blockId: dg.id, content }] } : x));
  const cases: [unknown, RegExp][] = [
    [{}, /needs content\.diagram/],
    [{ diagram: { ...flow, nodes: Array.from({ length: 13 }, (_, i) => node(i + 1)), edges: [] } }, /12 nodes at most/],
    [{ diagram: { ...flow, nodes: [node(1), node(1)], edges: [] } }, /share the id "n1"/],
    [{ diagram: { ...flow, edges: [{ from: "n1", to: "n9" }] } }, /names a node that is not in the list/],
    [{ diagram: { ...flow, kind: "matrix" } }, /matrix diagram takes exactly six nodes/],
  ];
  for (const [bad, re] of cases) {
    const r = await call("build_deck", { project: pid, slides: withDiagram(bad) });
    assert.ok(r.isError && re.test(r.text), `expected ${re}, got: ${r.text}`);
    assert.ok(/The ask/.test(r.text) && /How a referral moves/.test(r.text), `the refusal must name the slide and the block: ${r.text}`);
  }

  const good = await call("build_deck", { project: pid, slides: withDiagram({ diagram: flow }) });
  assert.ok(!good.isError, good.text);
  const unknown = (good.sc.build.warnings as string[]).find((w) => /unknown type '?diagram/.test(w));
  if (unknown) console.log(`SKIP diagram drawing: build_pptx.py does not take diagram blocks yet (${unknown})`);
  else {
    const snap2 = JSON.parse(fs.readFileSync(good.sc.project.snapshotPath as string, "utf8"));
    const shapes2 = snap2.slides.flatMap((x: any) => x.shapes);
    assert.ok(shapes2.some((sh: any) => sh.blockId === dg.id), `the diagram has no snapshot shape: ${JSON.stringify(shapes2.map((sh: any) => [sh.blockId, sh.kind]))}`);
  }
}

// ---- 7. review_deck: the built slides as pictures, plus what each one holds
{
  const rv = (await client.listTools()).tools.find((t) => t.name === "review_deck");
  assert.ok(rv, "review_deck is listed");
  const sentences = rv!.description!.split(/(?<=[.!?])\s+/).length;
  assert.ok(sentences <= 3, `review_deck description is ${sentences} sentences`);
  assert.deepEqual((rv!._meta as any)?.ui?.visibility, ["model"], "review_deck answers the model, it does not re-render the canvas");
  assert.ok(!/—/.test(rv!.description!), "no em dashes in the description");
  assert.match(JSON.stringify(rv!.inputSchema), /issues/, "what: issues is exposed to the model");

  // before any build there is nothing to look at, and it says so
  const fresh = await call("start_project", { name: "Nothing Built Yet", kind: "deck", brief: "A deck plan with no build behind it" });
  const before = await call("review_deck", { project: fresh.sc.projectId });
  assert.ok(before.isError && /No build yet/.test(before.text) && /build_deck first/.test(before.text), before.text);

  // one slide: its image, its headline and its blocks
  const one = await call("review_deck", { project: pid, slides: [4] });
  assert.ok(!one.isError, one.text);
  assert.equal(one.texts[0].split("\n").length, 1, `one line for one slide: ${one.texts[0]}`);
  assert.match(one.texts[0], /^slide 4 Where referrals stall: Referrals stall at three moments, and each one has a different fix/);
  assert.match(one.texts[0], /blocks: cards Three stall moments, takeaway Fix the ask/);
  if (thumbs.length) {
    assert.equal(one.images.length, 1, "one slide asked for, one image back");
    assert.equal(one.images[0].mimeType, "image/png");
    assert.deepEqual(one.images[0].annotations.audience, ["assistant"]);
    assert.equal(pngHead(one.images[0].data), "89504e47");
    assert.match(one.texts[1], /^Slide images follow: /);
  } else assert.equal(one.images.length, 0);

  // every slide by default, one line each, 12 images at most
  const all = await call("review_deck", { project: pid });
  assert.equal(all.texts[0].split("\n").length, 7, `one line per slide:\n${all.texts[0]}`);
  assert.ok(all.texts[0].split("\n").every((l: string) => /^slide \d+ /.test(l)), all.texts[0]);
  if (thumbs.length) assert.equal(all.images.length, 7);

  // issues mode: only the slides the deck check flagged
  const lastBuild = (await call("show_plan", { project: pid, view: "built" })).sc.build;
  // A deck-wide finding names no slide, so it is not one of the flagged slides.
  const deckWide = (lastBuild.findings ?? []).filter((f: any) => f.scope === "deck" || !f.slide);
  const flagged = [...new Set((lastBuild.findings ?? []).filter((f: any) => f.slide).map((f: any) => f.slide))];
  const issues = await call("review_deck", { project: pid, what: "issues" });
  if (flagged.length) {
    assert.equal(issues.texts[0].split("\n").length, flagged.length, `one line per flagged slide:\n${issues.texts[0]}`);
    assert.ok(issues.texts[0].split("\n").every((l: string) => /findings: /.test(l)), issues.texts[0]);
  } else if (deckWide.length) {
    assert.match(issues.text, /^Deck-wide: /, `the deck's own finding is still said once: ${issues.text}`);
    assert.equal(issues.images.length, 0, "no slide was flagged, so there is nothing to look at");
  } else {
    assert.ok(/No findings on the last build/.test(issues.text), issues.text);
    assert.equal(issues.images.length, 0, "nothing flagged, nothing to look at");
  }
  if (deckWide.length) assert.match(all.texts[1] ?? "", /Deck-wide: /, `every-slide review carries the deck's own finding: ${JSON.stringify(all.texts)}`);

  // a slide number the deck does not have
  const missing = await call("review_deck", { project: pid, slides: [99] });
  assert.match(missing.text, /No slide 99 .*7 slides/);
  assert.equal(missing.images.length, 0);

  // a PNG over 400 KB is resampled smaller, not dropped (the busiest slide's file, blown up on disk)
  const heaviest = (lastBuild.slides ?? [])
    .filter((r: any) => r.thumbnail && fs.existsSync(r.thumbnail))
    .sort((a: any, b: any) => fs.statSync(b.thumbnail).size - fs.statSync(a.thumbnail).size)[0];
  if (heaviest) {
    const file = heaviest.thumbnail as string;
    execFileSync("sips", ["--resampleWidth", "3000", file, "--out", `${file}.big.png`], { stdio: "ignore" });
    fs.renameSync(`${file}.big.png`, file);
    const size = fs.statSync(file).size;
    if (size <= 400 * 1024) console.log(`SKIP oversized image: slide ${heaviest.index} is only ${size} bytes at 3000 px wide`);
    else {
      const big = await call("review_deck", { project: pid, slides: [heaviest.index] });
      assert.equal(big.images.length, 1, `a ${Math.round(size / 1024)} KB PNG is resampled, not dropped: ${big.text}`);
      const got = Buffer.from(big.images[0].data, "base64").length;
      assert.ok(got <= 400 * 1024, `the resampled image is still ${Math.round(got / 1024)} KB`);
      assert.equal(pngHead(big.images[0].data), "89504e47");
      console.log(`oversized image: ${Math.round(size / 1024)}KB on disk, ${Math.round(got / 1024)}KB in the result`);
    }
  }
}

// ---- 8. a deck that brings no cover: the builder adds one, and every number follows the built deck
{
  const g = await call("start_project", {
    name: "Onboarding Handover", kind: "deck", formatting: "house",
    brief: "Short internal review of partner onboarding",
    briefCard: { audience: "Partner operations", goal: "Agree the two onboarding fixes we fund next", takeaway: "Onboarding stalls at the handover" },
  });
  const gid = g.sc.projectId as string;
  let o = await call("propose_outline", { project: gid, rows: [
    { name: "Why onboarding", divider: true },
    { name: "Where it stalls", story: "Names the handover as the stall, so the room stops blaming the partner." },
  ] });
  assert.ok(!o.isError, o.text);
  await call("ui_approve", { project: gid, what: "tabs" });
  // Two rows for six lines: the fidelity pass has to grow the block, so the build reports the move.
  o = await call("propose_layouts", { project: gid, layouts: [
    { tab: "Where it stalls", headline: "Onboarding stalls at the handover, not at the sale", blocks: [
      { type: "bullets", title: "Stall points", content: "one line per stall", row: 7, col: 2, rows: 2, cols: 28 },
    ] },
  ] });
  assert.ok(!o.isError, o.text);
  await call("ui_approve", { project: gid, what: "all_layouts" });
  const gtabs = o.sc.project.tabs as any[];
  const gcontent = gtabs.find((t: any) => t.name === "Where it stalls");
  const gb = await call("build_deck", { project: gid, slides: [
    { tabId: gcontent.id, blocks: [{ blockId: gcontent.blocks[0].id, content: { items: [
      "The handover email arrives with nobody named on it, so it waits",
      "The first task waits on an account we have not opened yet",
      "The first review lands after the partner has stopped asking for one",
      "Nobody owns the first reply, so a warm start cools to a cold one",
      "The partner learns the product from us twice, then from nobody",
      "The first invoice arrives before the first result does",
    ] } }] },
  ] });
  assert.ok(!gb.isError, gb.text);
  const gspec = JSON.parse(fs.readFileSync(gb.sc.project.specPath as string, "utf8"));
  assert.equal(gspec.cover, true, "no title slide in this outline, so the builder is asked for a cover");
  assert.equal(gspec.agenda, true, "and for an agenda");
  const gsnap = JSON.parse(fs.readFileSync(gb.sc.project.snapshotPath as string, "utf8"));
  const gindices = gsnap.slides.map((x: any, i: number) => x.index ?? i + 1);
  const grows = gb.sc.build.slides as any[];
  assert.deepEqual(grows.map((r: any) => r.index), gindices, "the built view numbers slides in the builder's order");
  if (gsnap.slides.length === gtabs.length) console.log("SKIP generated chrome: this builder added no cover or agenda");
  else {
    assert.ok(grows[0].generated && !grows[0].tabId, `the first built slide is the builder's own: ${JSON.stringify(grows[0])}`);
    const row = grows.find((r: any) => r.tabId === gcontent.id);
    assert.ok(row, `the plan slide is somewhere in the built list: ${JSON.stringify(grows)}`);
    assert.equal(row.index, gsnap.slides.findIndex((x: any) => x.id === gcontent.id) + 1, "a plan slide is numbered where the builder put it, not where the plan has it");
    assert.ok(row.index > 1, "the generated cover shifted it down");
    assert.equal(row.name, "Where it stalls", "and it is still its own plan row, matched by id");
    // review_deck counts the same slides the same way
    const one = await call("review_deck", { project: gid, slides: [row.index] });
    assert.ok(!one.isError, one.text);
    assert.ok(one.texts[0].startsWith(`slide ${row.index} Where it stalls`), one.texts[0]);
    assert.match(one.texts[0], /blocks: bullets Stall points/);
    // an adjustment is numbered by the built deck too, not by the plan
    const moves = gb.sc.build.adjusted as any[];
    if (!moves.length) console.log("SKIP adjusted: the fidelity pass moved nothing on this deck");
    else {
      assert.ok(moves.every((a: any) => gindices.includes(a.slide)), `adjustments name built slides: ${JSON.stringify(moves)}`);
      const mine = moves.find((a: any) => a.blockId === gcontent.blocks[0].id);
      assert.ok(mine, `the block that had to grow is in the list: ${JSON.stringify(moves)}`);
      assert.equal(mine.slide, row.index, "and it is numbered where the builder put its slide");
      assert.match(gb.text, /^Adjusted \d+ blocks? for fit \(balanced\); see the Built view$/m, gb.text);
    }
    const every = await call("review_deck", { project: gid });
    const lines = every.texts[0].split("\n");
    assert.deepEqual(lines.map((l: string) => Number(l.match(/^slide (\d+) /)![1])), gindices, `review_deck numbers by the built deck:\n${every.texts[0]}`);
    assert.match(lines[0], /^slide 1 .*\(generated\)/, lines[0]);
  }
}

console.log(`BUILT ${built}`);
console.log(buildText);
await client.close();
console.log("DECK FLOW OK", home);

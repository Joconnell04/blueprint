/**
 * extract_format over stdio: preview, save, the deck tokens sidecar, screenshots as image blocks,
 * the overwrite refusal, and fromSkill on a prose-only deck skill.
 * Every write is redirected into a temp BLUEPRINT_HOME and BLUEPRINT_SKILLS, so the real
 * ~/.claude/skills and ~/.blueprint are never touched.
 * Run: npx tsx test/extract.test.ts   (after tsc -p tsconfig.server.json)
 */
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const home = fs.mkdtempSync(path.join(os.tmpdir(), "blueprint-extract-"));
const skills = path.join(home, "skills");
const fixtures = path.join(home, "fixtures");
fs.mkdirSync(skills, { recursive: true });
execFileSync("python3", ["builder/fixtures/make_fixtures.py", fixtures], { stdio: "pipe" });

const house = path.join(fixtures, "house.xlsx");
const deck = path.join(fixtures, "deck.pptx");
const shot = path.join(fixtures, "logo.png");
for (const f of [house, deck, shot]) assert.ok(fs.existsSync(f), `fixture missing: ${f}`);

const client = new Client({ name: "extract-test", version: "0" });
await client.connect(new StdioClientTransport({
  command: "node", args: ["dist/server/main.js"],
  env: { ...process.env, BLUEPRINT_HOME: home, BLUEPRINT_OUTPUT: path.join(home, "out"), BLUEPRINT_SKILLS: skills },
}));

const call = async (name: string, args: Record<string, unknown>) => {
  const r: any = await client.callTool({ name, arguments: args });
  const text = r.content?.filter((c: any) => c.type === "text").map((c: any) => c.text).join("\n") ?? "";
  return { r, text, sc: (r.structuredContent ?? {}) as any, isError: !!r.isError };
};
const lines = (t: string) => t.split("\n").length;
const sidecar = (slug: string) => path.join(home, "styles", `${slug}.json`);
const skillFile = (slug: string) => path.join(skills, slug, "SKILL.md");
/** Every scalar leaf of a token block, to check `fields` drops none of them. */
const leafCount = (n: any): number => (n && typeof n === "object" && !Array.isArray(n) ? Object.values(n).reduce((t: number, v) => t + leafCount(v), 0) : 1);
const words = (s: string) => s.split(/\s+/).filter(Boolean).length;
/** The canvas renders a preview from `fields` alone: a summary row per section, then its fields, pre-ranked. */
const checkFields = (sc: any, sections: string[]) => {
  const keys = sc.fields.map((f: any) => f.key);
  assert.deepEqual(keys.filter((k: string) => !k.includes(".")), sections, `sections head the list: ${keys.slice(0, 3)}`);
  let expected = sections.length;
  for (const s of sections) {
    expected += leafCount(sc.format[s]);
    const at = keys.indexOf(s);
    assert.ok(sc.fields[at].confidence > 0, `${s} carries its own confidence`);
    assert.equal(sc.fields[at].value, undefined, "a section row is a heading, not a field");
    // the five rows the canvas shows are the ranked head of the section, each one weighable
    for (const f of sc.fields.slice(at + 1, at + 6)) {
      assert.ok(f.key.startsWith(`${s}.`), `${f.key} is not a field of ${s}`);
      assert.ok(f.confidence > 0, `${f.key} has no confidence`);
      assert.ok((f.evidence ?? "").length > 0, `${f.key} has no evidence`);
    }
  }
  assert.equal(sc.fields.length, expected, "every leaf is exported once");
  for (const f of sc.fields.filter((x: any) => x.key.includes("."))) assert.ok(f.value !== undefined && f.value !== null, `${f.key} has no value`);
};
/** Each unknown is one row's question plus the whole thing for the tooltip. */
const checkAsks = (sc: any) => {
  assert.ok(sc.unknowns.length > 0, "the files cannot answer everything");
  for (const u of sc.unknowns) {
    assert.ok(words(u.question) <= 12, `question ran to ${words(u.question)} words: ${u.question}`);
    assert.match(u.question, /^[A-Z].*[?…]$/);
    assert.ok(u.detail.length >= u.question.length, `detail is shorter than the question: ${u.detail}`);
  }
};

// ---- the tool is model-visible and describes itself in three sentences
const tools = await client.listTools();
const tool = tools.tools.find((t) => t.name === "extract_format")!;
assert.ok(tool, "extract_format is registered");
assert.ok(((tool._meta as any)?.ui?.visibility ?? []).includes("model"), "extract_format is visible to the model");
assert.equal((tool._meta as any)?.ui?.resourceUri, undefined, "no canvas resource: the result is text plus images");
assert.ok((tool.description ?? "").split(/(?<=\.)\s/).length <= 4, tool.description);

const p = await call("start_project", { name: "Extract probe", brief: "Somewhere for ui_styles to read from" });
const pid = p.sc.projectId as string;

// ---- 1. preview a workbook: key fields recovered with confidence, nothing written
let r = await call("extract_format", { paths: [house], name: "Acme House", kind: "workbook" });
assert.ok(!r.isError, r.text);
assert.equal(r.sc.saved, false);
assert.deepEqual(r.sc.sections, ["workbook"]);
for (const k of ["workbook.font", "workbook.size", "workbook.input_color", "workbook.title_rule_color", "workbook.label_divider", "workbook.currency"]) {
  assert.ok(r.sc.confidence[k] >= 0.8, `${k} came back at ${r.sc.confidence[k]}, under 0.8`);
}
assert.ok(r.sc.confidence.workbook >= 0.8, `overall ${r.sc.confidence.workbook}`);
assert.equal(r.sc.format.workbook.font, "Arial");
assert.equal(r.sc.format.workbook.input_color, "0000FF");
assert.equal(r.sc.format.workbook.label_divider, "dashed");
assert.ok(/Ask the user:/.test(r.text), r.text);
assert.ok(lines(r.text) <= 12, `preview text ran to ${lines(r.text)} lines`);
assert.ok(!fs.existsSync(path.join(skills, "acme-house")), "a preview writes nothing");

// the preview names the file a save would write, and says nothing is there yet
assert.equal(r.sc.exists, false, "nothing to replace");
assert.equal(r.sc.skillPath, skillFile("acme-house"));

// `fields` is the list the canvas draws: headline fields first, each with its confidence and evidence
checkFields(r.sc, ["workbook"]);
assert.deepEqual(r.sc.fields.slice(1, 4).map((f: any) => f.key), ["workbook.font", "workbook.size", "workbook.input_color"]);
assert.equal(r.sc.fields[1].value, "Arial");
assert.match(r.sc.fields[1].evidence, /Arial/);

// the questions are asked in twelve words, with the whole question kept for the tooltip
checkAsks(r.sc);
const dash = r.sc.unknowns.find((u: any) => /dash/i.test(u.question));
assert.equal(dash.question, "Should em and en dashes be rewritten (dash_replace)?");
assert.match(dash.detail, /max_note_chars/);

// ---- 2. save: a SKILL.md in the temp skills dir, listed by ui_styles with kind workbook
r = await call("extract_format", { paths: [house], name: "Acme House", kind: "workbook", save: true });
assert.ok(!r.isError, r.text);
assert.equal(r.sc.saved, true);
assert.equal(r.sc.skillPath, skillFile("acme-house"));
assert.ok(fs.existsSync(skillFile("acme-house")), "SKILL.md written");
const md = fs.readFileSync(skillFile("acme-house"), "utf8");
assert.match(md, /^---\nname: acme-house\n/);
assert.match(md, /\nformat:\n/);
assert.match(md, /\n {2}font: Arial\n/);
assert.match(md, /\n {2}input_color: "0000FF"\n/);
assert.ok(!/\ndeck:\n/.test(md), "kind workbook writes no deck block");
assert.equal(r.sc.tokensPath, undefined, "a workbook style has no deck tokens");
assert.ok(!fs.existsSync(sidecar("acme-house")), "no sidecar for a workbook style");
assert.ok(lines(r.text) <= 12, `save text ran to ${lines(r.text)} lines`);

let s = await call("ui_styles", { project: pid });
const acme = s.sc.styles.find((x: any) => x.id === "skill:acme-house");
assert.ok(acme, JSON.stringify(s.sc.styles.map((x: any) => x.id)));
assert.equal(acme.kind, "workbook");
assert.equal(acme.source, "skill");
assert.equal(acme.builderMode, "skill");
assert.equal(acme.skillPath, skillFile("acme-house"));
assert.equal(acme.tokensPath, undefined);
for (const b of ["plain", "house", "template"]) {
  assert.equal(s.sc.styles.find((x: any) => x.id === b).kind, "both", `${b} dresses both builders`);
}

// the style is selectable and drives the workbook builder as a skill
s = await call("ui_set_style", { project: pid, styleId: "skill:acme-house" });
assert.equal(s.sc.project.formatting, "skill");
assert.equal(s.sc.project.styleId, "skill:acme-house");

// ---- 3. an existing skill directory is never overwritten by accident
// the preview says it is there, so the popover offers Replace without a refused save first
r = await call("extract_format", { paths: [house], name: "Acme House", kind: "workbook" });
assert.ok(!r.isError, r.text);
assert.equal(r.sc.exists, true, "the skill is on disk now");
assert.equal(r.sc.skillPath, skillFile("acme-house"));
assert.ok(r.text.includes(skillFile("acme-house")) && /overwrite: true/.test(r.text), r.text);
assert.ok(lines(r.text) <= 12, `preview text ran to ${lines(r.text)} lines`);

r = await call("extract_format", { paths: [house], name: "Acme House", kind: "workbook", save: true });
assert.ok(r.isError, "a second save refuses");
assert.ok(r.text.includes(skillFile("acme-house")) && /overwrite: true/.test(r.text), r.text);
r = await call("extract_format", { paths: [house], name: "Acme House", kind: "workbook", save: true, overwrite: true });
assert.ok(!r.isError, r.text);
assert.ok(/replaced/.test(r.text), r.text);

// ---- 4. a deck: deck block in the SKILL.md plus builder tokens in the sidecar
r = await call("extract_format", { paths: [deck, house], name: "Harbor Deck", kind: "deck", save: true });
assert.ok(!r.isError, r.text);
assert.deepEqual(r.sc.sections, ["deck"]);
assert.ok(/Ignored for kind deck: house.xlsx/.test(r.text), r.text);
const dmd = fs.readFileSync(skillFile("harbor-deck"), "utf8");
assert.match(dmd, /\ndeck:\n/);
assert.match(dmd, /\n {2}palette\.primary: "1E3A5F"\n/);
assert.match(dmd, /\n {2}fonts\.major: Lato\n/);
assert.match(dmd, /\n {2}card\.border: 2\.25\n/);
assert.match(dmd, /\n {2}slideSize\.width: 960\n/);
assert.ok(!/\nformat:\n/.test(dmd), "kind deck writes no workbook block");
checkFields(r.sc, ["deck"]);
assert.deepEqual(r.sc.fields.slice(1, 5).map((f: any) => f.key), ["deck.palette.primary", "deck.palette.secondary", "deck.palette.ink", "deck.fonts.major"]);
assert.equal(r.sc.fields.find((f: any) => f.key === "deck.slideSize.width").value, 960);
assert.equal(r.sc.tokensPath, sidecar("harbor-deck"));
assert.ok(fs.existsSync(sidecar("harbor-deck")), "deck tokens sidecar written");
const tokens = JSON.parse(fs.readFileSync(sidecar("harbor-deck"), "utf8"));
assert.equal(tokens.palette.accent, "1E3A5F", "deck palette.primary becomes the builder's accent");
assert.equal(tokens.palette.accent2, "2A9D8F");
assert.equal(tokens.fonts.display, "Lato");
assert.equal(tokens.card.borderPt, 2.25);
assert.equal(tokens.slideSize.width, 960);
assert.equal(tokens.backgrounds.divider.fill, "2A9D8F");
assert.equal(tokens.backgrounds.divider.ink, "FFFFFF", "a dark slide flips its text colours");

s = await call("ui_styles", { project: pid });
const st = s.sc.styles.find((x: any) => x.id === "skill:harbor-deck");
assert.ok(st, JSON.stringify(s.sc.styles.map((x: any) => x.id)));
assert.equal(st.kind, "deck");
assert.equal(st.tokensPath, sidecar("harbor-deck"), "the popover can see the style is buildable");

// ---- 5. a screenshot path comes back as an image block for Claude, not as parsed tokens
r = await call("extract_format", { paths: [shot], name: "Shot only" });
assert.ok(!r.isError, r.text);
const img = r.r.content.find((c: any) => c.type === "image");
assert.ok(img, `no image block in ${JSON.stringify(r.r.content.map((c: any) => c.type))}`);
assert.equal(img.mimeType, "image/png");
assert.deepEqual(img.annotations.audience, ["assistant"]);
assert.equal(img.data, fs.readFileSync(shot).toString("base64"));
assert.equal(r.sc.needsObservations, true);
assert.deepEqual(r.sc.sections, []);
assert.deepEqual(r.sc.fields, [], "a screenshot gives up no fields until Claude reads it");
assert.ok(/observations/.test(r.text), r.text);
assert.ok(lines(r.text) <= 12, `image preview ran to ${lines(r.text)} lines`);
// the question is pulled out of the middle of the sentence and still reads as one
checkAsks(r.sc);
assert.equal(r.sc.unknowns[0].question, "What palette, fonts, type sizes and card style does it show?");
assert.match(r.sc.unknowns[0].detail, /^logo\.png is a screenshot/);

// ---- 6. second call: the reading of the screenshot is merged and written, and no image comes back
r = await call("extract_format", {
  paths: [shot], name: "Shot Deck", kind: "deck", save: true,
  observations: {
    format: { deck: { palette: { primary: "112233", secondary: "445566", ink: "1F2937" }, fonts: { major: "Inter", minor: "Inter" }, typeScale: { headline: { size: 30, bold: true } } } },
    confidence: { "palette.primary": 0.9 },
    rules: ["Headlines sit on two lines at most."],
  },
});
assert.ok(!r.isError, r.text);
assert.ok(!r.r.content.some((c: any) => c.type === "image"), "screenshots already read are not sent again");
const shotMd = fs.readFileSync(skillFile("shot-deck"), "utf8");
assert.match(shotMd, /\n {2}palette\.primary: "112233"\n/);
assert.match(shotMd, /\n {2}fonts\.major: Inter\n/);
assert.match(shotMd, /Headlines sit on two lines at most\./);
assert.equal(JSON.parse(fs.readFileSync(sidecar("shot-deck"), "utf8")).palette.accent, "112233");

// ---- 7. fromSkill: a prose-only deck skill gets its prose read back, then a sidecar of its own
const proseDir = path.join(skills, "test-deck-prose");
fs.mkdirSync(proseDir, { recursive: true });
const proseFile = path.join(proseDir, "SKILL.md");
fs.writeFileSync(proseFile, `---\nname: test-deck-prose\ndescription: "Test deck format for the extract test: navy palette, Lato type, bordered cards."\n---\n# Test deck format\nCards carry a 2.25 pt navy border and a soft fillet.\nHeadlines are sentence case, 28 pt, ink.\n`);
const proseBefore = fs.readFileSync(proseFile, "utf8");

s = await call("ui_styles", { project: pid });
const prose0 = s.sc.styles.find((x: any) => x.id === "skill:test-deck-prose");
assert.ok(prose0, "a prose-only deck skill lists from its description alone");
assert.equal(prose0.kind, "deck");
assert.equal(prose0.tokensPath, undefined, "prose only: nothing to build from yet");

r = await call("extract_format", { fromSkill: "test-deck-prose" });
assert.ok(!r.isError, r.text);
assert.ok(r.text.includes("Cards carry a 2.25 pt navy border"), "the prose comes back for Claude to read");
assert.equal(r.sc.needsObservations, true);
assert.equal(r.sc.saved, false);
assert.equal(r.sc.sidecarPath, sidecar("test-deck-prose"));
assert.equal(r.sc.skillPath, proseFile);

r = await call("extract_format", {
  fromSkill: "test-deck-prose", save: true,
  observations: {
    format: { deck: { palette: { primary: "1E3A5F", secondary: "2A9D8F", ink: "1F2937", bg: "FFFFFF" }, fonts: { major: "Lato", minor: "Lato" }, card: { shape: "roundRect", radius: 8000, border: 2.25, borderColor: "1E3A5F" }, typeScale: { headline: { size: 28, case: "sentence" } } } },
    rules: ["Cards carry a 2.25 pt navy border."],
  },
});
assert.ok(!r.isError, r.text);
assert.equal(r.sc.saved, true);
assert.ok(fs.existsSync(sidecar("test-deck-prose")), "the sidecar is written");
const pt = JSON.parse(fs.readFileSync(sidecar("test-deck-prose"), "utf8"));
assert.equal(pt.palette.accent, "1E3A5F");
assert.equal(pt.card.borderPt, 2.25);
assert.equal(pt.card.borderColor, "accent", "a colour that has a palette name is written as the name");
assert.equal(pt.typeScale.headline.size, 28);
assert.equal(fs.readFileSync(proseFile, "utf8"), proseBefore, "fromSkill never edits the skill itself");
assert.deepEqual(fs.readdirSync(proseDir), ["SKILL.md"], "nothing else lands in the skill directory");

s = await call("ui_styles", { project: pid });
const prose1 = s.sc.styles.find((x: any) => x.id === "skill:test-deck-prose");
assert.equal(prose1.tokensPath, sidecar("test-deck-prose"), "the sidecar makes the prose skill buildable");
assert.equal(prose1.kind, "deck");

r = await call("extract_format", { fromSkill: "no-such-skill" });
assert.ok(r.isError && /No skill named no-such-skill/.test(r.text), r.text);

// ---- 8. the bundled tokens stand in for a harbor-deck skill that has no sidecar
const bundledDir = path.join(skills, "harbor-deck-bundled");
fs.mkdirSync(bundledDir, { recursive: true });
fs.writeFileSync(path.join(bundledDir, "SKILL.md"), `---\nname: harbor-deck-bundled\ndescription: "Harbor deck format: navy palette, Lato type, bordered cards."\n---\n# Harbor\nPipe kicker, action title, three cards.\n`);
s = await call("ui_styles", { project: pid });
const bundled = s.sc.styles.find((x: any) => x.id === "skill:harbor-deck-bundled");
assert.ok(bundled, "a harbor-deck skill lists");
assert.equal(bundled.kind, "deck");
assert.equal(bundled.tokensPath, path.resolve("builder/styles/harbor_deck.json"), "no user sidecar: the bundled tokens are referenced, not copied");
assert.ok(!fs.existsSync(sidecar("harbor-deck-bundled")), "nothing is written into the user's styles dir");

// ---- 9. bad input fails politely
r = await call("extract_format", { paths: ["/nope/missing.xlsx"], name: "Nope" });
assert.ok(r.isError && /No file at/.test(r.text), r.text);
r = await call("extract_format", { paths: [house], save: true });
assert.ok(r.isError && /needs a name/.test(r.text), r.text);
r = await call("extract_format", { paths: [house], kind: "deck" });
assert.ok(r.isError && /Nothing to read/.test(r.text), r.text);

await client.close();
fs.rmSync(home, { recursive: true, force: true });
console.log("EXTRACT OK");

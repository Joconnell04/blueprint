import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { translateFormula, originalContent } from "../src/server/import-formulas.js";

const scanPath = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "wp-if-")), "scan.json");
if (!fs.existsSync("builder/sample_house.xlsx")) execFileSync("python3", ["builder/build_xlsx.py", "builder/sample_spec.json"]);
execFileSync("python3", ["builder/import_scan.py", "builder/sample_house.xlsx", scanPath]);
const scan = JSON.parse(fs.readFileSync(scanPath, "utf8"));

const ramp = scan.tabs.find((t: any) => t.name === "Revenue Ramp").blocks.find((b: any) => b.type === "calc");
const first = ramp.sample[0];
// "='Assumptions'!C18" -> an inputs value cell on Assumptions -> {{id!Label}}
const t1 = translateFormula(first[1], "Revenue Ramp", scan);
assert.match(t1.formula, /^=\{\{imp_\d+_\d+![^}]+\}\}$/, t1.formula);
assert.equal(t1.unresolved.length, 0);
// "='TAM Build'!C11*1000*C5": cross-sheet total cell and a same-sheet data cell of this block
const t2 = translateFormula(first[2], "Revenue Ramp", scan);
assert.match(t2.formula, /\{\{imp_\d+_\d+!(total:\d+|\d+,\d+)\}\}\*1000\*\{\{imp_\d+_\d+!0,\d+\}\}/, t2.formula);
// function names are not mangled
const t3 = translateFormula("=LOG10(C5)+SUM(C5:C9)", "Revenue Ramp", scan);
assert.ok(t3.formula.startsWith("=LOG10("), t3.formula);
assert.match(t3.formula, /SUM\(\{\{imp_\d+_\d+!col:\d+\}\}\)/, t3.formula);
const oc = originalContent(scanPath);
assert.ok(oc[ramp.id] && oc[ramp.id].rows.length === 5 && oc[ramp.id].unresolved.length === 0, JSON.stringify(oc[ramp.id]));
console.log("IMPORT FORMULAS OK", t1.formula, "|", t2.formula);

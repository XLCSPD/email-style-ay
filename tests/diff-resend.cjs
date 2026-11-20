const fs = require("node:fs");
const vm = require("node:vm");
const context = { Office: { actions: { associate() {} }, context: {} }, Promise, String };
vm.createContext(context);
vm.runInContext(fs.readFileSync(process.argv[2], "utf8"), context);

const cell = (t) => `<td style="border:1px solid rgb(171,171,171);padding:4px"><div class="elementToProof" style="color:rgb(0,111,201);font-size:12pt">${t}</div></td>`;
const inserted = `<table style="border-collapse:collapse"><tbody><tr>${cell("Region")}${cell("Sales")}</tr><tr>${cell("EMEA")}${cell("1.2M")}</tr></tbody></table>`;
const pasted = `<table class="MsoTable"><tr><td style="background:#C6EFCE;color:#006100">+12%</td></tr></table>`;
const email = `<html><body><div style="color:rgb(0,111,201)">Numbers:</div>${inserted}${pasted}</body></html>`;

const passes = [email];
for (let i = 0; i < 4; i++) passes.push(context.styleEmail(passes[passes.length - 1]).html);
for (let i = 2; i < passes.length; i++) {
  const a = passes[i - 1], b = passes[i];
  if (a === b) { console.log(`pass ${i}: stable`); continue; }
  let at = 0;
  while (a[at] === b[at]) at++;
  console.log(`pass ${i}: differs at ${at}`);
  console.log("  before:", a.slice(Math.max(0, at - 120), at + 120));
  console.log("  after: ", b.slice(Math.max(0, at - 120), at + 120));
}

// Runs every check against site/launchevent.js. Exit code 1 if anything fails.
// The PASS/FAIL suites report each check; the older suites print values, so their
// failure lines are recognised by pattern.
const { spawnSync } = require("node:child_process");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const engine = path.join(root, "site", "launchevent.js");
const site = path.join(root, "site");
const manifest = path.join(root, "site", "manifest.xml");

const suites = [
  ["check-smart", [engine]],
  ["check-kpi", [engine]],
  ["check-word-safe", [engine]],
  ["check-replies-attach", [engine]],
  ["check-mention-agenda", [engine]],
  ["check-pane", [site]],
  ["check-setup", [engine, manifest]],
  ["check-pasted", [engine]],
  ["check-brands", [engine, site]],
  ["check-hardening", [engine]],
  ["check-personal-reply", [engine]],
];
const legacy = ["check-nested", "check-numbered", "check-design-b", "check-edge-cases", "check-new-features", "check-bare-title", "check-reply", "check-subject"];
const LEGACY_PROBLEM = /error|text lost: [^n]|balanced: false|untouched: false|intact: false|styled: false|branded: false|stable: false|restyled: false|undefined|: false$/i;
const LEGACY_EXPECTED_FALSE = /body rewritten: false|allowed: false|rewritten: false|touched: false|li's div: false/;

const run = (name, args) => spawnSync(process.execPath, [path.join(__dirname, name + ".cjs"), ...args], { encoding: "utf8" });

let failed = 0;
let passed = 0;
for (const [name, args] of suites) {
  const result = run(name, args);
  const lines = (result.stdout + result.stderr).split(/\r?\n/);
  const pass = lines.filter((line) => line.startsWith("PASS")).length;
  const fail = lines.filter((line) => line.startsWith("FAIL") || /Error/.test(line));
  passed += pass;
  failed += fail.length + (result.status ? 1 : 0);
  console.log(`${name.padEnd(22)} ${String(pass).padStart(3)} pass${fail.length ? ", " + fail.length + " fail" : ""}`);
  fail.forEach((line) => console.log("   " + line));
}

for (const name of legacy) {
  const result = run(name, [engine]);
  const problems = (result.stdout + result.stderr)
    .split(/\r?\n/)
    .filter((line) => LEGACY_PROBLEM.test(line) && !LEGACY_EXPECTED_FALSE.test(line));
  failed += problems.length + (result.status ? 1 : 0);
  console.log(`${name.padEnd(22)} ${problems.length ? problems.length + " problem(s)" : "ok"}`);
  problems.forEach((line) => console.log("   " + line));
}

const resend = run("diff-resend", [engine]).stdout.trim().split(/\r?\n/).pop();
const stable = /stable/.test(resend);
if (!stable) failed++;
console.log(`${"diff-resend".padEnd(22)} ${stable ? "stable" : resend}`);

console.log(`\n${passed} checks passed${failed ? `, ${failed} problem(s)` : ""}.`);
process.exit(failed ? 1 : 0);

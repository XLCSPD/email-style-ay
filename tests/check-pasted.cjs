const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync(process.argv[2], "utf8");
function load() {
  const context = { Office: { actions: { associate() {} }, context: {} }, Promise, String };
  vm.createContext(context);
  vm.runInContext(source, context);
  return context;
}
const report = (name, ok, detail = "") => console.log(`${ok ? "PASS" : "FAIL"} | ${name}${detail ? " | " + detail : ""}`);
const count = (s, re) => (s.match(re) || []).length;
const st = "font-family:Aptos;font-size:12pt;color:rgb(0,111,201)";
const li = (t) => `<li style="${st}"><div role="presentation">${t}</div></li>`;
const div = (t) => `<div class="elementToProof" style="${st}">${t}</div>`;
const email = (own) => `<html><head></head><body>${own}</body></html>`;
const own = (html) => html.slice(0, html.search(/id="?appendonsend|<\/body>/i));
const lost = (out, words) => words.filter((w) => !out.includes(w));
const c = load();

// What Outlook makes of a status update pasted from a document: every numbered decision is its own list,
// and its sub-points are a bullet list nested in a bullet list, with no item to hang from.
const pasted =
  div("<b>Decisions needed</b>") +
  `<div><ol style="margin-top:0"><li style="${st}">Should we adopt the new reporting tool?</li></ol></div>` +
  `<div><ul><ul style="list-style-type:circle">${li("sign the agreement")}${li("review the terms")}${li("verify a test run")}</ul></ul></div>` +
  `<div><ol start="2"><li style="${st}">Should we run a pilot?</li></ol></div>` +
  `<div><ul><ul style="list-style-type:circle">${li("A pilot costs about $14–16 per person")}${li("The goal is 90 of 100 reports")}</ul></ul></div>` +
  div("Spent so far: about $40");
let out = own(c.applyVersatexStyle(email(pasted)));
report("pasted lists: no Outlook list left", !/<(ul|ol|li)\b/i.test(out));
report("pasted lists: badges 1 and 2 kept", /<span[^>]*>1<\/span><\/td>/.test(out) && /<span[^>]*>2<\/span><\/td>/.test(out));
report("pasted lists: 5 grey hexagons", count(out, /color:#58595b;font-size:18px[^>]*>&#11042;/g) === 5, String(count(out, /&#11042;/g)));
const firstItem = out.slice(out.indexOf("Should we adopt"), out.indexOf("Should we run"));
report("pasted lists: sub-points sit under their numbered item", ["sign the agreement", "review the terms", "verify a test run"].every((w) => firstItem.includes(w)));
report("pasted lists: no words lost", lost(out, ["sign the agreement", "review the terms", "verify a test run", "A pilot costs", "90 of 100", "Spent so far"]).length === 0);

// Same thing inside one numbered list (sub-list between items, indented twice).
out = own(c.applyVersatexStyle(email(`<ol>${li("One")}<ul><ul>${li("one-a")}${li("one-b")}</ul></ul>${li("Two")}</ol>`)));
report("double-indented sub-list between items -> hexagons", !/<(ul|ol|li)\b/i.test(out) && count(out, /&#11042;/g) === 2 && lost(out, ["One", "one-a", "one-b", "Two"]).length === 0);

out = own(c.applyVersatexStyle(email(`${div("Notes")}<ul><ul>${li("alpha")}${li("beta")}</ul></ul>`)));
report("parentless list with nothing above -> plain hexagon list", !/<(ul|ol|li)\b/i.test(out) && count(out, /&#11042;/g) === 2);

out = own(c.applyVersatexStyle(email(`<ol>${li("Top")}</ol><ul><ul><ul>${li("deep")}</ul></ul></ul>`)));
report("triple-indented paste -> hexagon under the item", !/<(ul|ol|li)\b/i.test(out) && out.includes("deep") && count(out, /&#11042;/g) === 1);

out = own(c.applyVersatexStyle(email(`<ol>${li("Top")}</ol>${div("A normal paragraph")}<ul><ul>${li("not attached")}</ul></ul>`)));
report("a paragraph in between stops the attaching", out.indexOf("A normal paragraph") < out.indexOf("not attached") && !/<(ul|ol|li)\b/i.test(out));

// Loose text inside lists is kept.
const shapes = {
  "lines as divs inside a sub-list": `<ol>${li("A")}<ul>${li("a1")}${div("STAYING")}${div("IFTRIAL")}</ul>${li("C")}</ol>`,
  "an item written as a div": `<ol>${li("A")}<ul>${li("a1")}</ul>${div("C-VERSION2")}</ol>`,
  "bare text between items": `<ol>${li("A")}<ul>${li("a1")}STAYING IFTRIAL</ul>C-VERSION2</ol>`,
  "unclosed list items": `<ul><li>first<li>second<li>third</ul>`,
};
for (const [name, html] of Object.entries(shapes)) {
  out = own(c.applyVersatexStyle(email(html)));
  const words = (html.replace(/<[^>]+>/g, " ").match(/[A-Za-z0-9-]+/g) || []).filter((w) => w.length > 1);
  report(`loose text kept: ${name}`, lost(out, words).length === 0, lost(out, words).join(","));
}

// Ranges are one figure.
out = own(c.applyVersatexStyle(email(div("About $14–16 or $41-43, maybe $14–$16; 10–20% of runs; 7–15 hours; up 6%."))));
const figures = [...out.matchAll(/class="vstx-figure"[^>]*>([^<]+)</g)].map((m) => m[1]);
report("ranges marked as one figure", figures.join("|") === "$14–16|$41-43|$14–$16|10–20%|6%", figures.join("|"));

// Safety net: a styling bug that drops words must not reach the recipient.
const safe = load();
const lossy = (html) => html.replace(/IMPORTANT FACT/g, "");
safe.restyleOwnPart = lossy;
out = safe.applyVersatexStyle(email(div("Keep this IMPORTANT FACT please") + `<ul>${li("item")}</ul>`));
report("safety net: words kept when styling would drop them", out.includes("IMPORTANT FACT") && out.includes("vstx-bar") && out.includes("vstx-signature"));
report("safety net: falls back to basic styling (fonts, no smart parts)", /font-family:Roboto/.test(own(out)) && !/vstx-list/.test(out));
safe.brandText = lossy;
out = safe.applyVersatexStyle(email(div("Keep this IMPORTANT FACT please")));
report("safety net: if even basic styling loses words, the text goes as written", out.includes("IMPORTANT FACT"));

const fresh = load();
const wordsOf = fresh.ownWords;
report("word check ignores what styling legitimately changes",
  fresh.keepsWords(email(div("KPI: Spend $4.1M -8% [On track]") + div('<a href="https://app.powerbi.com/groups/me/r/1">https://app.powerbi.com/groups/me/r/1</a>') + div("Action: ship it")),
    fresh.applyVersatexStyle(email(div("KPI: Spend $4.1M -8% [On track]") + div('<a href="https://app.powerbi.com/groups/me/r/1">https://app.powerbi.com/groups/me/r/1</a>') + div("Action: ship it")))));
report("word check notices a missing word", !fresh.keepsWords(email(div("alpha beta gamma")), email(div("alpha gamma"))));
report("word check counts repeats", !fresh.keepsWords(email(div("go go go")), email(div("go go"))));

const statusUpdate = email(div("Good morning, team,") + pasted + div("<b>Risks to know</b>") + `<ul>${li("Numbers may change after the next data refresh.")}${li("The pilot needs one more reviewer.")}</ul>`);
const styled = c.applyVersatexStyle(statusUpdate);
report("whole status update: styled with every word kept", c.keepsWords(statusUpdate, styled) && /vstx-list/.test(styled));
report("whole status update: re-send stable", c.applyVersatexStyle(styled) === c.applyVersatexStyle(c.applyVersatexStyle(styled)));

// A bullet list copied from an earlier styled email: Outlook on the web drops the vstx- class names and mso-
// properties, so it arrives as a plain table. It must go out as the list it was, not branded as a data table.
const earlier = c.styleEmail(email(`<ul>${li("The report now loads faster.")}${li("Filters remember your last choice.")}</ul>`)).html;
const copiedList = /<div class="vstx-space"[^>]*><table class="vstx-list"[\s\S]*?<\/table><\/div>/.exec(earlier)[0]
  .replace(/\sclass="[^"]*"/g, "")
  .replace(/mso-[a-z-]+:[^;"]*;?/g, "");
const withCopiedList = own(c.styleEmail(email(div("Two small things:") + copiedList + div("Thanks"))).html);
const DATA_TABLE_HEADER = /background:#231f20[^"]*border-bottom:3px solid #fdc00f/;
report("copied hexagon list is not branded as a data table", !DATA_TABLE_HEADER.test(withCopiedList) && !/vstx-table/.test(withCopiedList));
report("copied hexagon list keeps its bullets and words", count(withCopiedList, /&#11042;|\u2B22/g) === 2 && withCopiedList.includes("The report now loads faster.") && withCopiedList.includes("Filters remember your last choice."));
const insertedTable = own(c.styleEmail(email(`<table><tr><td>Name</td><td>Role</td></tr><tr><td>Sam</td><td>Mac</td></tr></table>`)).html);
report("a table you insert yourself is still branded", /vstx-table/.test(insertedTable) && DATA_TABLE_HEADER.test(insertedTable));

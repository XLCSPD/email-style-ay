const fs = require("node:fs");
const vm = require("node:vm");

const context = { Office: { actions: { associate() {} } }, Promise, String };
vm.createContext(context);
vm.runInContext(fs.readFileSync(process.argv[2], "utf8"), context);

const defaultDiv = (inner) => `<div class="elementToProof" style="font-family:Aptos,Calibri,sans-serif;font-size:12pt;color:rgb(15,108,189)">${inner}</div>`;
const body = (inner) => `<html><body>${inner}</body></html>`;
const tagBefore = (out, marker) => {
  const at = out.indexOf(marker);
  const open = out.lastIndexOf("<", at);
  return out.slice(open, out.indexOf(">", open) + 1);
};
const effective = (tag, prop) => {
  const style = (/style="([^"]*)"/.exec(tag) || [])[1] || "";
  const values = [...style.matchAll(new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([^;]+)`, "gi"))].map((m) => m[1].trim());
  return values[values.length - 1] || "(inherits)";
};

const cases = [
  ["Default typed text", body(defaultDiv("Hello team")), (o) => `colour ${effective(tagBefore(o, "Hello team"), "color")}, size ${effective(tagBefore(o, "Hello team"), "font-size")}, font ${effective(tagBefore(o, "Hello team"), "font-family")}`],
  ["Red warning text", body(defaultDiv(`Status: <span style="color:rgb(209,52,56)">DEADLINE MOVED</span>`)), (o) => `colour ${effective(tagBefore(o, "DEADLINE MOVED"), "color")}`],
  ["Large emphasis text", body(defaultDiv(`<span style="font-size:20pt">Key number: 42%</span>`)), (o) => `size ${effective(tagBefore(o, "Key number"), "font-size")}`],
  ["Code / DAX snippet", body(`<div style="font-family:Consolas,monospace;font-size:10pt">CALCULATE([Sales], ALL(Date))</div>`), (o) => `font ${effective(tagBefore(o, "CALCULATE"), "font-family")}, size ${effective(tagBefore(o, "CALCULATE"), "font-size")}`],
  ["Blue code line", body(`<div style="font-family:Consolas,monospace;font-size:10pt;color:rgb(0,111,201)">Here is an example</div>` + defaultDiv("normal").replace("rgb(15,108,189)", "rgb(0,111,201)")), (o) => `colour ${effective(tagBefore(o, "Here is an example"), "color")}, font ${effective(tagBefore(o, "Here is an example"), "font-family")}, size ${effective(tagBefore(o, "Here is an example"), "font-size")}`],
  ["Code block with spans", body(`<div style="font-family:Consolas,monospace;font-size:10pt;color:rgb(15,108,189)"><span style="color:rgb(0,0,255)">EVALUATE</span> <span style="font-size:12pt">TOPN</span></div>` + defaultDiv("normal")), (o) => `keyword font ${effective(tagBefore(o, "EVALUATE"), "font-family")}, keyword colour ${effective(tagBefore(o, "EVALUATE"), "color")}, TOPN size ${effective(tagBefore(o, "TOPN"), "font-size")}`],
  ["Pasted Excel table", body(defaultDiv("Results:") + `<table><tr><td style="background:#C6EFCE;color:#006100;font-size:9pt;font-family:Calibri">+12%</td><td style="background:#FFC7CE;color:#9C0006;font-size:9pt">-8%</td></tr></table>`), (o) => `+12% colour ${effective(tagBefore(o, "+12%"), "color")}, size ${effective(tagBefore(o, "+12%"), "font-size")}, font ${effective(tagBefore(o, "+12%"), "font-family")}`],
  ["Quoted family entity", body(`<div style="font-family:&quot;Segoe UI&quot;,sans-serif;font-size:12pt;color:rgb(15,108,189)">Entity font</div>`), (o) => `font ${effective(tagBefore(o, "Entity font"), "font-family")}, colour ${effective(tagBefore(o, "Entity font"), "color")}`],
  ["Link in text", body(defaultDiv(`See <a href="https://example.com" style="color:#0f6cbd">the report</a>`)), (o) => `link colour ${effective(tagBefore(o, "the report"), "color")}`],
  ["Logo image", body(defaultDiv("x")), (o) => `logo ${/defoxx-logo\.png/.test(o)}, motif ${/defoxx-x\.png/.test(o)}, alt styled: ${/alt="D\.E\. Foxx &amp; Associates"[^>]*font-weight:700/.test(o)}`],
];

for (const [name, html, check] of cases) console.log(name.padEnd(22), "|", check(context.applyVersatexStyle(html)));

const shades = context.applyVersatexStyle(body(`<ol><li>Top</li><ul><li>Second</li><ol><li>Third</li></ol></ul></ol><ul><li>Bullet</li><ul><li>Sub bullet</li></ul></ul>`));
const colourOf = (word) => (new RegExp(`color:(#[0-9a-f]{6})[^>]*>&#11042;</span></td><td[^>]*><div[^>]*>${word}`).exec(shades) || [])[1];
console.log("Hexagon shades".padEnd(22), "| second", colourOf("Second"), "| third", colourOf("Third"), "| top bullet", colourOf("Bullet"), "| sub bullet", colourOf("Sub bullet"));

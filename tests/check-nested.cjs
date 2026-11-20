const fs = require("node:fs");
const vm = require("node:vm");

const context = { Office: { actions: { associate() {} } }, Promise, String };
vm.createContext(context);
vm.runInContext(fs.readFileSync(process.argv[2], "utf8"), context);

const li = (t, extra = "") => `<li style="font-family:Aptos;font-size:12pt"><div role="presentation" class="elementToProof">${t}</div>${extra}</li>`;
const subUl = (...items) => `<ul style="list-style-type:circle">${items.map((t) => `<li style="list-style-type:circle"><div role="presentation">${t}</div></li>`).join("")}</ul>`;
const wrap = (inner) => `<html><body><div>Test the new add-in</div>${inner}<div>Thanks</div></body></html>`;

const cases = {
  userCase: wrap(`<ol>${li("Plan for the launch", subUl("Draft the summary", "Share it for review", "Publish on Friday"))}${li("Update the slides")}${li("Bulleted lists")}${li("Numbered lists")}${li("Unchanged")}${li("Checks")}</ol>`),
  bulletsInBullets: wrap(`<ul>${li("Outer one", subUl("Inner a", "Inner b"))}${li("Outer two")}</ul>`),
  numbersInBullets: wrap(`<ul>${li("Outer", `<ol>${li("Step")}${li("Step")}</ol>`)}${li("Outer two")}</ul>`),
  numbersInNumbers: wrap(`<ol>${li("Phase", `<ol>${li("Sub")}${li("Sub")}</ol>`)}${li("Phase two")}</ol>`),
  outlookSiblingUl: wrap(`<ol>${li("Plan for the launch")}${subUl("Draft the summary", "Share it for review", "Publish on Friday")}${li("Update the slides")}${li("Checks")}</ol>`),
  outlookSiblingOl: wrap(`<ol>${li("Phase")}<ol>${li("Sub a")}${li("Sub b")}</ol>${li("Phase two")}</ol>`),
  outlookUlInUl: wrap(`<ul>${li("Outer")}${subUl("Inner a", "Inner b")}${li("Outer two")}</ul>`),
  numberedSubSteps: wrap(`<ol>${li("Plan for the launch")}<ol style="list-style-type:lower-alpha">${li("Draft the summary")}${li("Share it for review")}${li("Publish on Friday")}</ol>${li("Update the slides")}${li("Bulleted lists")}${li("Numbered lists")}${li("Unchanged")}${li("Checks")}</ol>`),
};

for (const [name, html] of Object.entries(cases)) {
  const out = context.applyVersatexStyle(html);
  const badges = [...out.matchAll(/text-align:center"><span[^>]*>(\d+)<\/span><\/td>/g)].map((m) => m[1]);
  const hexagons = (out.match(/&#11042;/g) || []).length;
  const leftoverOl = (out.match(/<ol\b/gi) || []).length;
  const unstyledLi = [...out.matchAll(/<li\b[^>]*>/gi)].filter((m) => !/list-style:none/.test(m[0])).length;
  const hexOnOwnLine = /&#11042;<\/span>&nbsp;&nbsp;\s*<div/i.test(out);
  const balanced = (out.match(/<table\b/gi) || []).length === (out.match(/<\/table>/gi) || []).length;
  const words = ["Draft the summary", "Share it for review", "Publish on Friday", "Sub a", "Sub b", "Inner a", "Inner b"];
  const lost = words.filter((w) => html.includes(w) && !out.includes(w));
  console.log(
    "| text lost:", lost.length ? lost.join("/") : "none",
    name.padEnd(17),
    "| badges:", badges.join(",") || "-",
    "| hexagons:", hexagons,
    "| <ol> left:", leftoverOl,
    "| bullet <li> without list-style:none:", unstyledLi,
    "| hexagon before a block:", hexOnOwnLine,
    "| tables balanced:", balanced
  );
}

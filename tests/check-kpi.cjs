const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync(process.argv[2], "utf8");
function load(features) {
  const context = { Office: { actions: { associate() {} }, context: { roamingSettings: { get: (k) => (k === "vstxFeatures" ? features : undefined) } } }, Promise, String };
  vm.createContext(context);
  vm.runInContext(source, context);
  return context;
}
const report = (name, ok, detail = "") => console.log(`${ok ? "PASS" : "FAIL"} | ${name}${detail ? " | " + detail : ""}`);
const count = (s, re) => (s.match(re) || []).length;
const blue = "font-family:Aptos;font-size:12pt;color:rgb(0,111,201)";
const div = (t) => `<div style="${blue}">${t}</div>`;
const email = (...blocks) => `<html><head></head><body>${blocks.join("")}</body></html>`;

const c = load({});
const p = (s) => JSON.parse(JSON.stringify(c.parseKpi(s)));
const same = (got, want) => Object.keys(want).every((k) => got[k] === want[k]);

let k = p("Spend $4.1M -8% vs plan [On track]");
report("full segment parsed", same(k, { label: "Spend", value: "$4.1M", change: "-8%", note: "vs plan", status: "On track" }), JSON.stringify(k));
k = p("Margin 18%");
report("value only", same(k, { label: "Margin", value: "18%", change: null, note: "" }), JSON.stringify(k));
k = p("EMEA -4%");
report("a lone signed number is the value, not a change", same(k, { label: "EMEA", value: "-4%", change: null }), JSON.stringify(k));
k = p("Q3 revenue $1.2M (+6%)");
report("'Q3' isn't read as a number; brackets dropped", same(k, { label: "Q3 revenue", value: "$1.2M", change: "+6%", note: "" }), JSON.stringify(k));
k = p("FY2026 savings $3M");
report("'FY2026' isn't read as a number", same(k, { label: "FY2026 savings", value: "$3M" }), JSON.stringify(k));
k = p("$620K saved +12%");
report("number-first segment uses the words as the label", same(k, { label: "saved", value: "$620K", change: "+12%" }), JSON.stringify(k));
k = p("PO cycle 3.2 days -0.4");
report("time units stay with the number", same(k, { label: "PO cycle", value: "3.2 days", change: "-0.4" }), JSON.stringify(k));
k = p("Margin 18% +2 pts");
report("points as a change", same(k, { value: "18%", change: "+2 pts" }), JSON.stringify(k));

let out = c.applyVersatexStyle(email(div("KPI: Spend $4.1M -8% [On track] | Savings $620K +12% | Off-contract 14% +3% [At risk]")));
const kpi = /<table class="vstx-kpi"[\s\S]*?<\/table>/.exec(out)[0];
report("three tiles", count(kpi, /class="vstx-kpi-tile"/g) === 3);
report("labels written in capitals (Outlook on the web ignores text-transform)", />SPEND<\/div>/.test(kpi) && />OFF-CONTRACT<\/div>/.test(kpi) && !/text-transform/.test(kpi));
report("arrows show direction", count(kpi, /&#9660;/g) === 1 && count(kpi, /&#9650;/g) === 2);
report("arrows are never coloured", ![...kpi.matchAll(/color:(#[0-9a-f]{6})[^>]*>&#96(?:50|60);/g)].some((m) => m[1] !== "#231f20"));
report("colour only where a status tag was added", count(kpi, /class="vstx-status"/g) === 2 && kpi.includes("#2e7d32") && kpi.includes("#e07b00") && !kpi.includes("#c62828"));
report("change shown without its sign", />8%<\/span>/.test(kpi) && />12%<\/span>/.test(kpi));
report("no yellow figure markers inside tiles", !kpi.includes("vstx-figure"));
report("tiles have the yellow top rule", count(kpi, /border-top:3px solid #fdc00f/g) === 3);

out = c.applyVersatexStyle(email(div("KPIs: A 1 | B 2 | C 3 | D 4 | E 5")));
report("'KPIs:' works; at most 3 tiles per row", /vstx-kpi/.test(out) && count(out, /<tr>(?:(?!<\/tr>)[\s\S])*?vstx-kpi-tile/g) === 2);
report("tile widths leave room for gaps", /vstx-kpi-tile" width="32%"/.test(out) && /vstx-kpi-gap" width="2%"/.test(out));

out = c.applyVersatexStyle(email(div("KPI: Revenue $1.2M")));
report("a single KPI is half width, not full width", /vstx-kpi-tile" width="49%"/.test(out));

out = c.applyVersatexStyle(email(div("KPI: something without numbers")));
report("no numbers -> left as text", !out.includes("vstx-kpi") && out.includes("KPI: something without numbers"));

out = c.applyVersatexStyle(email(div("KPI: R&amp;D spend $1M &lt;b&gt;")));
report("text is escaped", out.includes("R&amp;D SPEND") && !out.includes("<b>&gt;") && !/R&amp;amp;D/.test(out));

out = c.applyVersatexStyle(email(`<ul><li style="${blue}">KPI: Revenue $1.2M</li></ul>`));
report("not converted inside a list", !out.includes("vstx-kpi"));

out = load({ kpis: false }).applyVersatexStyle(email(div("KPI: Revenue $1.2M +6%")));
report("switch off -> plain line", !out.includes("vstx-kpi"));

const first = c.applyVersatexStyle(email(div("Hi"), div("KPI: Spend $4.1M -8% [On track] | Margin 18%"), div("Thanks")));
const second = c.applyVersatexStyle(first);
report("re-send stable", second === c.applyVersatexStyle(second) && count(second, /class="vstx-kpi-tile"/g) === 2 && count(second, /class="vstx-status"/g) === 1 && !second.includes("vstx-heading"));

out = c.applyVersatexStyle(email(div("Thanks"), '<div id="appendonsend"></div><hr>', div("KPI: Old $1M")));
report("quoted part untouched", !out.slice(out.indexOf("appendonsend")).includes("vstx-kpi"));

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
const div = (t) => `<div class="elementToProof" style="${blue}">${t}</div>`;
const email = (...blocks) => `<html><head></head><body>${blocks.join("")}</body></html>`;
const c = load({ mentions: true });
const run = (...blocks) => c.applyVersatexStyle(email(...blocks));
report("mention tags are off by default", !load({}).applyVersatexStyle(email(div('Hi <a id="OWAAM1" href="mailto:a@x.com">@A</a>'))).includes("vstx-mention"));

// ---- @mentions ----
const owaMention = '<a id="OWAAM4F1C2D" href="mailto:mark@example.com" style="color:rgb(0,120,212)">@Mark Smith</a>';
let out = run(div(`Action: ${owaMention} to send the deck`));
const link = /<a\b[^>]*>[\s\S]*?<\/a>/.exec(out)[0];
report("mention keeps Outlook's id and address", /id="OWAAM4F1C2D"/.test(link) && /href="mailto:mark@example.com"/.test(link));
report("mention becomes a name tag", /<span class="vstx-mention"[^>]*>@Mark Smith<\/span><\/a>/.test(link));
report("tag style: bold ink, grey, yellow underline", /vstx-mention" style="[^"]*font-weight:700;color:#231f20;background:#f4f4f4;border-bottom:2px solid #fdc00f/.test(link));
report("link loses Outlook blue and underline", /<a [^>]*style="[^"]*color:#231f20[^"]*text-decoration:none/.test(link) && !/rgb\(0,120,212\)/.test(link));
report("mention inside a callout still works", /vstx-callout-text[\s\S]*vstx-mention/.test(out));

out = run(div('Ping <a href="mailto:ann@x.com">@Ann</a> about it'));
report("@name mailto link without an Outlook id is tagged", out.includes('class="vstx-mention"'));

out = run(div('Write to <a href="mailto:ann@x.com">ann@x.com</a> or <a href="https://x.com">@x</a>'));
report("ordinary email and web links untouched", !out.includes("vstx-mention"));

out = run(div(`${owaMention.replace('id="OWAAM4F1C2D"', 'id="x_OWAAM4F1C2D"')} said yes`));
report("quoted-style x_OWAAM id recognised", out.includes("vstx-mention"));

const once = run(div(`Thanks ${owaMention}`));
report("re-send keeps one tag", c.applyVersatexStyle(once) === c.applyVersatexStyle(c.applyVersatexStyle(once)) && count(c.applyVersatexStyle(once), /vstx-mention/g) === 1);

out = c.applyVersatexStyle(email(div("Thanks"), '<div id="appendonsend"></div><hr>', div(`Earlier: ${owaMention}`)));
report("quoted mentions untouched", !out.includes("vstx-mention"));

out = load({ mentions: false }).applyVersatexStyle(email(div(`Hi ${owaMention}`)));
report("switch off -> mention left as Outlook made it", !out.includes("vstx-mention") && out.includes(owaMention));

// ---- Agenda ----
out = run(div("Agenda:"), div("10:00 Welcome"), div("10:05-10:45 Q3 review [On track]"), div("10:45 – 11:30 Vendor roadmap"), div("11:30am Wrap-up"), div("Thanks"));
const agenda = /<table class="vstx-agenda"[\s\S]*?<\/table>/.exec(out);
report("4 timed lines become one agenda", agenda && count(agenda[0], /class="vstx-agenda-time"/g) === 4);
report("time ranges with an en dash", agenda && agenda[0].includes(">10:05&#8211;10:45<") && agenda[0].includes(">10:45&#8211;11:30<"));
report("durations computed", agenda && agenda[0].includes(">40 min<") && agenda[0].includes(">45 min<") && count(agenda[0], /vstx-agenda-length/g) === 2);
report("item text kept, time prefix removed", agenda && />Welcome<\/div>/.test(agenda[0]) && /vstx-agenda-text[^>]*>Vendor roadmap/.test(agenda[0]) && !/vstx-agenda-text[^>]*>\s*10/.test(agenda[0]));
report("status tags still work inside the agenda", agenda && /Q3 review <span class="vstx-status"/.test(agenda[0]));
report("lines before and after stay text", out.includes(">Agenda:</div>") && out.includes(">Thanks</div>"));

out = run(div("9:00 Kickoff"), div("12:30pm Lunch"), div("1:30pm-3pm Workshop"));
report("am/pm and 'pm' only times", out.includes("vstx-agenda") && out.includes(">1 h 30 min<"));

out = run(div("10:00 call with Mark"), div("Then we review the numbers."));
report("a single timed line stays a sentence", !out.includes("vstx-agenda"));

out = run(div("10:00 Welcome"), div("<br>"), div("11:00 Review"));
report("a blank line breaks the run", !out.includes("vstx-agenda"));

out = run(div("3 people joined"), div("5 decisions made"));
report("plain numbers aren't times", !out.includes("vstx-agenda"));

out = run(div("25:00 nope"), div("10:75 nope"));
report("impossible times rejected", !out.includes("vstx-agenda"));

out = run(`<ul><li style="${blue}">10:00 Welcome</li><li style="${blue}">10:30 Review</li></ul>`);
report("lines inside a list aren't turned into an agenda", !out.includes("vstx-agenda"));

const agendaEmail = run(div("10:00 Welcome"), div("10:30 Review"));
report("re-send stable", c.applyVersatexStyle(agendaEmail) === c.applyVersatexStyle(c.applyVersatexStyle(agendaEmail)) && count(c.applyVersatexStyle(agendaEmail), /class="vstx-agenda"/g) === 1);

out = load({ agenda: false }).applyVersatexStyle(email(div("10:00 Welcome"), div("10:30 Review")));
report("switch off -> plain lines", !out.includes("vstx-agenda"));

out = run(div("<b>10:00</b> Welcome"), div("10:30 <i>Review</i>"));
report("formatting around the time handled", out.includes("vstx-agenda") && out.includes("<i>Review</i>"));

out = run(div("1.50 per unit"), div("2.25 per unit"));
report("decimal prices aren't times", !out.includes("vstx-agenda"));
out = run(div("10:00am: Welcome"), div("10:30am | Review"));
report("separators after the time accepted", out.includes("vstx-agenda") && />Welcome<\/div>/.test(out) && />Review<\/div>/.test(out));

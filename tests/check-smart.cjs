const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync(process.argv[2], "utf8");
function load(features) {
  const roaming = features ? { get: (key) => (key === "vstxFeatures" ? features : undefined) } : null;
  const context = { Office: { actions: { associate() {} }, context: roaming ? { roamingSettings: roaming } : {} }, Promise, String };
  vm.createContext(context);
  vm.runInContext(source, context);
  return context;
}
const blue = "font-family:Aptos,Calibri,sans-serif;font-size:12pt;color:rgb(0,111,201)";
const div = (inner) => `<div class="elementToProof" style="${blue}">${inner}</div>`;
const email = (...blocks) => `<html><head><style>P {margin-top:0;margin-bottom:0;}</style></head><body>${blocks.join("")}</body></html>`;
const count = (s, re) => (s.match(re) || []).length;
const report = (name, ok, detail = "") => console.log(`${ok ? "PASS" : "FAIL"} | ${name}${detail ? " | " + detail : ""}`);

const ctx = load(null);
const run = (html) => ctx.applyVersatexStyle(html);

// Headings
let out = run(email(div("Hi team,"), div("<b>Q3 Highlights</b>"), div("Revenue grew."), div("<strong>Next steps:</strong>"), div("<b>Thanks</b>"), div("<b>Sales</b> were strong this quarter.")));
report("bold line becomes heading", /class="vstx-heading-text"[^>]*>Q3 Highlights</.test(out));
report("trailing colon dropped in heading", /vstx-heading-text"[^>]*>Next steps</.test(out));
report("greeting/sign-off not a heading", !/vstx-heading-text"[^>]*>Thanks/.test(out) && !/vstx-heading-text"[^>]*>Hi team/.test(out));
report("partly-bold sentence not a heading", !/vstx-heading-text"[^>]*>Sales/.test(out));

// Callouts
out = run(email(div("Action: send the deck to Mark by Friday"), div("<b>Decision:</b> we go with option B"), div("Note: numbers are preliminary"), div("Next&nbsp;steps: review on Monday"), div("Notes from the call are attached")));
report("Action callout: yellow edge, label on an ink tag", /vstx-callout-label[^>]*border-left:3px solid #fdc00f[^>]*><span[^>]*background:#231f20;color:#ffffff[^>]*>ACTION</.test(out) && /vstx-callout-text[^>]*>send the deck/.test(out));
report("bold label stripped, body kept", /vstx-callout-text[^>]*><b><\/b>\s*we go with option B/.test(out) || /vstx-callout-text[^>]*>(<b><\/b>)?we go with option B/.test(out), out.match(/vstx-callout-text[^>]*>[^<]{0,5}(<b><\/b>)?[^<]{0,40}/g)?.[1]);
report("soft label (Note): grey text and edge, no yellow", /vstx-callout-label"[^>]*border-left:3px solid #a7a9ac[^>]*><span style="[^"]*color:#58595b">NOTE</.test(out));
const deadline = run(email(div("Deadline: numbers due Friday")));
report("strong label without a tag (Deadline): ink text, yellow edge", /vstx-callout-label"[^>]*border-left:3px solid #fdc00f[^>]*><span style="[^"]*color:#231f20">DEADLINE</.test(deadline) && !/background:#231f20;color:#ffffff[^>]*>DEADLINE/.test(deadline));
report("callout panel is light, never a black block", !/vstx-callout-label"[^>]*bgcolor="#(?:231f20|58595b)"/.test(out + deadline));
report("Next&nbsp;steps recognised", />NEXT STEPS</.test(out));
report("'Notes from…' not a callout", !/>Notes</.test(out) && out.includes("Notes from the call"));

// Key figures
out = run(email(div("Revenue was $1.2M, up 6% (EMEA -4%). Headcount 120. Call 513-555-0142 on Nov 5 about Q3."), div('See <a href="https://x.com">the 5% report</a>')));
const figures = [...out.matchAll(/class="vstx-figure"[^>]*>([^<]+)</g)].map((m) => m[1]);
report("figures marked", figures.join("|") === "$1.2M|6%|-4%", figures.join("|"));
report("yellow highlight when few figures", /vstx-figure" style="font-weight:700;background:#fff0c2/.test(out) && !/vstx-figure" style="[^"]*border-bottom/.test(out));
report("no figure inside links", !/<a [^>]*>[^<]*<span class="vstx-figure"/.test(out));
const many = run(email(div("1% 2% 3% 4% 5% 6% 7% 8% 9% 10%")));
report("many figures -> bold only", count(many, /vstx-figure" style="font-weight:700"/g) === 10 && !many.includes("background:#fff0c2"));
const codeFig = run(email(`<div style="font-family:Consolas">SELECT 50% FROM t</div>`));
report("no figures in code", !codeFig.includes("vstx-figure"));

// Link buttons
out = run(email(div('<a href="https://app.powerbi.com/groups/me/reports/123">https://app.powerbi.com/groups/me/reports/123</a>'), div('<a href="https://contoso.sharepoint.com/:x:/r/sites/fin/Book.xlsx">https://contoso.sharepoint.com/:x:/r/sites/fin/Book.xlsx</a>'), div('<a href="https://example.com/q3">Q3 dashboard</a>'), div('Read <a href="https://example.com/a">this</a> first')));
report("Power BI link -> button", /vstx-button[\s\S]*?>Open Power BI report&nbsp;&#8594;<\/a>/.test(out) && />app\.powerbi\.com</.test(out));
report("SharePoint Excel label", out.includes("Open the Excel workbook&nbsp;&#8594;"));
report("named link keeps its text", out.includes("Q3 dashboard&nbsp;&#8594;") && out.includes(">example.com<"));
report("link inside a sentence untouched", out.includes('Read <a href="https://example.com/a">this</a> first'));
const misleading = [
  "https://sharepoint.com:x@evil.example/doc",
  "https://app.powerbi.com@evil.example/r",
  "https://sharepoint.com:x&#64;evil.example/doc",
  "https://sh\u0430repoint.com/doc",
];
misleading.forEach((href) => {
  const styled = run(email(div("Hi team,"), div(`<a href="${href}">${href}</a>`), div("Thanks")));
  report(`misleading link stays a plain link: ${href}`, styled.includes("vstx-signature") && !/vstx-button/.test(styled) && styled.includes(`<a href="${href}">${href}</a>`));
});
out = run(email(div('<a href="https://evil.example\\@sharepoint.com/doc">https://evil.example\\@sharepoint.com/doc</a>')));
report("backslash link: caption names the site a browser opens", />evil\.example</.test(out) && !out.includes("Open in SharePoint"));
out = run(email(div('<a href="https://contoso.sharepoint.com:443/:w:/r/sites/fin/Plan.docx">https://contoso.sharepoint.com:443/:w:/r/sites/fin/Plan.docx</a>')));
report("link with a port: button and caption without the port", out.includes("Open the Word document&nbsp;&#8594;") && />contoso\.sharepoint\.com</.test(out));

// Status tags
out = run(email(div("Migration [On track], data feed [At risk], vendor [blocked], API [Done]."), `<div style="font-family:Consolas">[Done]</div>`));
report("status pills", count(out, /class="vstx-status"/g) === 4 && out.includes("#2e7d32") && out.includes("#e07b00") && out.includes("#c62828"));
report("status labels sentence case", />&#11042;<\/span>&nbsp;On track</.test(out) && /&nbsp;Blocked</.test(out));
report("no pill in code", out.includes(">[Done]</div>"));

// Lists + smart formatting together, and stability on re-send
const listEmail = email(div("<b>Plan</b>"), `<ol><li style="${blue}">Ship [On track] by 15%</li><li style="${blue}">Action: not a callout inside a list</li></ol>`, div("Deadline: Friday"), div('<a href="https://app.powerbi.com/r/1">https://app.powerbi.com/r/1</a>'));
const first = ctx.styleEmail(listEmail).html;
const second = ctx.styleEmail(first).html;
const third = ctx.styleEmail(second).html;
report("list item keeps pill + figure", /vstx-status[\s\S]*?On track[\s\S]*?vstx-figure[^>]*>15%/.test(first));
report("no callout inside lists", !/vstx-callout-text[^>]*>not a callout/.test(first));
report("re-send stable", second === third && count(second, /vstx-heading-text/g) === 1 && count(second, /class="vstx-button"/g) === 1 && count(second, /class="vstx-callout"/g) === 1 && count(second, /class="vstx-figure"/g) === 1 && count(second, /class="vstx-status"/g) === 1);

// v30 look: bold (700) at most, because Arial at 900 turns into Arial Black; rounded components
report("nothing heavier than bold", !/font-weight:\s*900/.test(first), (first.match(/[^;"]*font-weight:\s*900/) || [""])[0]);
report("button, KPI-style pills and signature panel have rounded corners", /class="vstx-button"[\s\S]*?border-radius:6px/.test(first) && /class="vstx-status"[^>]*border-radius:10px/.test(first) && /vstx-signature[\s\S]*?border-radius:6px 0 0 6px/.test(first));

// Plain section titles: a short unpunctuated line between other text becomes a heading, only when switched on
const headingsIn = (html) => [...html.matchAll(/vstx-heading-text"[^>]*>([^<]*)</g)].map((m) => m[1]);
const pastedTitles = email(
  div("Hi all,"), div("Here is the plan for the rollout."), div("What you need"), `<ul><li style="${blue}">Windows</li><li style="${blue}">Short item</li></ul>`,
  div("How to get it"), div("In the terminal, run:"), div("/setup install tracker"), div("See below"), div("Good morning"), div("Project notes (new)"),
  div("Read the notes first."), div('<a href="https://example.com/a">Q3 dashboard</a>'), div("Reply with any questions."), div("This line has far too many words in it"),
  div("More text after it."), div("Questions or problems: message me."), div("Thanks")
);
const titlesOn = load({ plainTitles: true });
report("plain titles: off by default", !/vstx-heading/.test(run(pastedTitles)));
const plainHeadings = headingsIn(titlesOn.applyVersatexStyle(pastedTitles));
report("plain titles: section titles become headings", plainHeadings.join("|") === "What you need|How to get it|Project notes (new)", plainHeadings.join("|"));
report("plain titles: first and last lines stay text (a name greeting, a typed name)",
  headingsIn(titlesOn.applyVersatexStyle(email(div("Priya"), div("The numbers are in."), div("Jordan Reyes")))).length === 0);
const pastedReply = email(div("Hi Sam,"), div(""), div("Thanks for this."), div("Two small things I noticed:"), div(""), div("the export now skips empty rows"), div(""), div("Thanks for flagging the repository."), div("Thanks"));
report("plain titles: a reply without section titles is untouched", headingsIn(titlesOn.applyVersatexStyle(pastedReply)).length === 0);
const titled = titlesOn.styleEmail(pastedTitles).html;
report("plain titles: every word kept and re-send stable", titlesOn.keepsWords(pastedTitles, titled) && titlesOn.styleEmail(titled).html === titlesOn.styleEmail(titlesOn.styleEmail(titled).html).html);

// Feature switches
const off = load({ headings: false, callouts: false, figures: false, buttons: false, statuses: false }).applyVersatexStyle(listEmail);
report("all switches off -> plain styling only", !/vstx-(heading|callout|figure|button|status)/.test(off) && off.includes("vstx-signature"));
const onlyFigures = load({ headings: false, callouts: false, figures: true, buttons: false, statuses: false }).applyVersatexStyle(listEmail);
report("single switch respected", onlyFigures.includes("vstx-figure") && !onlyFigures.includes("vstx-heading") && !onlyFigures.includes("vstx-status"));

// Quoted part is never smart-formatted
out = run(email(div("Thanks"), '<div id="appendonsend"></div><hr><div id="divRplyFwdMsg">From: X</div>', div("<b>Old Heading</b>"), div("Action: old item 5%")));
const quoted = out.slice(out.indexOf("appendonsend"));
report("quoted part untouched", !/vstx-/.test(quoted));

const fs = require("node:fs");
const vm = require("node:vm");

const context = { Office: { actions: { associate() {} }, context: {} }, Promise, String };
vm.createContext(context);
vm.runInContext(fs.readFileSync(process.argv[2], "utf8"), context);
const report = (name, ok, detail = "") => console.log(`${ok ? "PASS" : "FAIL"} | ${name}${detail ? " | " + detail : ""}`);
const count = (s, re) => (s.match(re) || []).length;
const blue = "font-family:Aptos;font-size:12pt;color:rgb(0,111,201)";
const email = (own, quoted = "") => `<html><head><style>P {margin-top:0;margin-bottom:0;}</style></head><body>${own}${quoted}</body></html>`;
const ownPart = (html) => html.slice(0, html.search(/id="?appendonsend|<\/body>/i));

const lists = email(
  `<ul><li style="${blue}"><div>Alpha</div></li><ul><li style="${blue}">Alpha child</li></ul><li style="${blue}">Beta</li></ul>` +
    `<ol><li style="${blue}">Step one</li><ul><li style="${blue}">Detail</li></ul><li style="${blue}">Step two</li></ol>`
);
let out = context.applyVersatexStyle(lists);
const own = ownPart(out);
report("no <ul>/<ol>/<li> left in own part", !/<(ul|ol|li)\b/i.test(own));
report("one hexagon per bullet item", count(own, /&#11042;/g) === 4, String(count(own, /&#11042;/g)));
report("every hexagon in its own cell", count(own, /<td width="22"[^>]*><span[^>]*>&#11042;<\/span><\/td>/g) === 4);
report("number badges kept", /<span[^>]*>1<\/span><\/td>/.test(own) && /<span[^>]*>2<\/span><\/td>/.test(own));
report("all text kept", ["Alpha", "Alpha child", "Beta", "Step one", "Detail", "Step two"].every((t) => own.includes(">" + t)));
report("list text wrapped in ink", !/<td[^>]*padding:0 0 4px 6px">[^<]/.test(own));
report("list tables carry vstx-list", count(own, /class="vstx-list"/g) === 4);

const resent = context.applyVersatexStyle(out);
report("re-send leaves lists alone", resent === context.applyVersatexStyle(resent) && count(resent, /&#11042;/g) === 4);

const pastedStyled = email(`<ul><li><span style="color:#58595b">&#11042;</span>&nbsp;&nbsp;Copied item</li></ul>`);
report("hexagon pasted from an older email isn't doubled", count(context.applyVersatexStyle(pastedStyled), /&#11042;/g) === 1);

const reply = email(`<div style="${blue}">Thanks</div>`, `<div id="appendonsend"></div><hr><div id="divRplyFwdMsg">From: X</div><ul><li>Quoted</li></ul>`);
out = context.applyVersatexStyle(reply);
report("quoted list untouched", out.includes("<ul><li>Quoted</li></ul>"));

out = context.applyVersatexStyle(email(`<div style="${blue}">Hi</div>`));
const bar = /<table id="vstx-bar"[\s\S]*?<\/table>/.exec(out)[0];
report("bar cells sized for Word", count(bar, /mso-line-height-rule:exactly/g) === 4 && /height="6"[^>]*font-size:6px;line-height:6px/.test(bar));
report("bar spacing is a row, not a margin", !/vstx-bar[^>]*margin/.test(bar) && /height="24"/.test(bar));

out = context.applyVersatexStyle(
  email(`<div style="${blue}">Action: one</div><div style="${blue}">Note: two</div><div style="${blue}"><a href="https://app.powerbi.com/r/1">https://app.powerbi.com/r/1</a></div><div style="${blue}"><b>Plan</b></div>`)
);
report("callouts spaced by a row inside the table", count(out, /<table class="vstx-callout"[^>]*><tr><td colspan="2" height="8"/g) === 2);
report("button spaced by a wrapper", /<div class="vstx-space"[^>]*><table class="vstx-button"/.test(out));
report("no component table relies on margin", !/<table class="vstx-(?:callout|button|list)"[^>]*margin/.test(out));
report("heading rule sized for Word", /vstx-heading[\s\S]*?height="3"[^>]*font-size:3px;line-height:3px;mso-line-height-rule:exactly/.test(out));

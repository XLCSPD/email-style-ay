// A.Y. edition: the optional personal reply signature (short name + tagline).
const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync(process.argv[2], "utf8");
const report = (name, ok, detail = "") => console.log(`${ok ? "PASS" : "FAIL"} | ${name}${detail ? " | " + detail : ""}`);
const html = (own, quoted = "") => `<html><head></head><body><div style="font-family:Aptos;font-size:12pt">${own}</div>${quoted}</body></html>`;
const quote = (inner) => `<div id="appendonsend"></div><hr><div id="divRplyFwdMsg">From: someone</div>${inner}`;
const teamsBlock = '<div><span>________________________________________________________________________________</span></div><div class="me-email-text">Join the meeting</div>';
const fullSig = '<table id="x_vstx-signature"><tr><td><span style="font-size:16px">Ayodele Sasore</span></td></tr></table>';
const personalSig = '<table id="x_vstx-signature" class="x_vstx-signature-short"><tr><td><span style="font-size:15px">A.Y.</span></td></tr></table>';
const TAGLINE = "The only disability in life is a bad attitude";
const BASE = { name: "Ayodele Sasore", title: "CITO", phone: "", brand: "defoxx" };
const own = (out) => out.slice(0, out.indexOf("appendonsend") === -1 ? out.length : out.indexOf("appendonsend"));
const hasFullPanel = (part) => /<table id="vstx-signature" role="presentation" width="100%"/.test(part);
const hasPersonal = (part) => part.includes("vstx-signature-personal");

function engine(signature, features) {
  const roaming = { vstxSignature: signature, vstxFeatures: features };
  const Office = { actions: { associate() {} }, context: { roamingSettings: { get: (key) => roaming[key] } } };
  const context = { Office, Promise, String };
  vm.createContext(context);
  vm.runInContext(source, context);
  return context;
}

const mine = { ...BASE, replyName: "A.Y.", replyTagline: TAGLINE };
let c = engine(mine);

const fresh = c.applyVersatexStyle(html("Hi team"));
report("new email still gets the full panel with name and title", hasFullPanel(fresh) && !hasPersonal(fresh) && fresh.includes(">Ayodele Sasore<") && fresh.includes(">CITO<"));

const signed = own(c.applyVersatexStyle(html("Thanks", quote(fullSig))));
report("reply in a thread I signed gets the personal signature", hasPersonal(signed) && !hasFullPanel(signed));
report("personal signature shows the short name, not the full name", signed.includes(">A.Y.<") && !signed.includes(">Ayodele Sasore<") && !signed.includes(">CITO<"));
report("personal signature carries the tagline under the name", new RegExp("A\\.Y\\.</span><br><span[^>]*>" + TAGLINE + "<").test(signed));
report("default look: blue italic handwritten tagline, no yellow rule", /Comic Sans MS[^"]*font-style:italic;color:#0000ff/.test(signed) && !/vstx-signature-personal[\s\S]*border-left/.test(signed));

const later = own(c.applyVersatexStyle(html("Ok", quote(personalSig))));
report("a thread carrying only my personal signature still counts as signed", hasPersonal(later));

const firstReply = own(c.applyVersatexStyle(html("Thanks", quote("<div>Their message</div>"))));
report("first reply to someone else's thread keeps the full panel by default", hasFullPanel(firstReply));

const always = engine({ ...mine, replyWhen: "always" });
report("'every reply': first reply gets the personal signature", hasPersonal(own(always.applyVersatexStyle(html("Thanks", quote("<div>Their message</div>"))))));
report("'every reply': a new email still gets the full panel", hasFullPanel(always.applyVersatexStyle(html("Hi"))));
const invite = always.applyVersatexStyle(html("Agenda below", teamsBlock));
report("'every reply': a meeting invite with a Teams block is not a reply", hasFullPanel(invite) && !hasPersonal(invite));

const branded = own(engine({ ...mine, replyLook: "brand" }).applyVersatexStyle(html("Thanks", quote(fullSig))));
report("brand look: yellow rule, grey italic, no handwriting font", /border-left:3px solid #fdc00f/.test(branded) && /font-style:italic;color:#58595b/.test(branded) && !branded.includes("Comic Sans"));

const standard = own(engine(BASE).applyVersatexStyle(html("Thanks", quote(fullSig))));
report("no reply fields saved -> the standard one-line signature", standard.includes("vstx-signature-short") && !hasPersonal(standard) && standard.includes(">CITO<"));

const taglineOnly = own(engine({ ...BASE, replyTagline: TAGLINE }).applyVersatexStyle(html("Thanks", quote(fullSig))));
report("tagline without a short name uses the full name", hasPersonal(taglineOnly) && taglineOnly.includes(">Ayodele Sasore<"));

const off = own(engine(mine, { shortReplies: false }).applyVersatexStyle(html("Thanks", quote(fullSig))));
report("short signatures switched off -> full panel on replies", hasFullPanel(off) && !hasPersonal(off));

const once = c.applyVersatexStyle(html("Thanks", quote(fullSig)));
const twice = c.applyVersatexStyle(once);
report("re-send is stable and keeps one signature", twice === once && (own(twice).match(/id="vstx-signature"/g) || []).length === 1);

const nasty = own(engine({ ...BASE, replyName: "<b>x</b>", replyTagline: '"><img src=x onerror=1>' }).applyVersatexStyle(html("Thanks", quote(fullSig))));
report("reply fields are escaped", !/<b>x<\/b>|<img src=x/.test(nasty));

const junk = engine({ ...BASE, replyName: 5, replyTagline: null, replyLook: "weird", replyWhen: {} });
report("malformed saved values fall back safely", !hasPersonal(own(junk.applyVersatexStyle(html("Thanks", quote(fullSig))))));

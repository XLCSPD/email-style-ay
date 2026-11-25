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

// ---- Per-email choice: [simple] / [formal] in the subject ----
const ok = (value) => ({ status: "succeeded", value });
function mailbox(subject, body) {
  const state = { subject, body, notices: [] };
  state.item = {
    subject: { getAsync: (cb) => cb(ok(state.subject)), setAsync: (v, cb) => { state.subject = v; cb(ok()); } },
    body: {
      getTypeAsync: (cb) => cb(ok("html")),
      getAsync: (_t, cb) => cb(ok(state.body)),
      setAsync: (v, _o, cb) => { state.body = v; cb(ok()); },
    },
    getAttachmentsAsync: (cb) => cb(ok([])),
    notificationMessages: { replaceAsync: (_k, m, cb) => { state.notices.push(m.message); cb(ok()); } },
  };
  return state;
}
function live(state, signature = mine) {
  const roaming = { vstxSignature: signature };
  const Office = {
    actions: { associate() {} },
    AsyncResultStatus: { Succeeded: "succeeded" },
    CoercionType: { Html: "html", Text: "text" },
    MailboxEnums: { SendModeOverride: { PromptUser: "promptUser" }, ItemNotificationMessageType: { InformationalMessage: "info" } },
    context: { roamingSettings: { get: (key) => roaming[key] }, requirements: { isSetSupported: () => true }, mailbox: { item: state.item } },
  };
  const context = { Office, Promise, String };
  vm.createContext(context);
  vm.runInContext(source, context);
  return context;
}
const send = (state, signature) => new Promise((resolve) => live(state, signature).onMessageSendHandler({ completed: resolve }));
const preview = (state) => new Promise((resolve) => live(state).previewStyle({ completed: resolve }));

(async () => {
  let s = mailbox("[simple] Lunch Friday", html("Are you free?"));
  let r = await send(s);
  report("[simple] on a new email -> personal signature", r.allowEvent === true && hasPersonal(s.body) && !hasFullPanel(s.body));
  report("[simple] is removed from the subject", s.subject === "Lunch Friday");

  s = mailbox("RE: Budget [FORMAL]", html("Approved.", quote(fullSig)));
  r = await send(s);
  report("[formal] on a reply in a signed thread -> full panel", r.allowEvent === true && hasFullPanel(own(s.body)) && !hasPersonal(own(s.body)));
  report("[formal] is removed from the subject, any case", s.subject === "RE: Budget");

  s = mailbox("Lunch Friday", html("Are you free?"));
  await send(s);
  report("no tag -> defaults unchanged (new email, full panel)", hasFullPanel(s.body));

  s = mailbox("[simple] Lunch", html("Are you free?"));
  await send(s, BASE);
  report("[simple] with no personal signature set -> the standard one-line signature", s.body.includes("vstx-signature-short") && !hasPersonal(s.body) && s.body.includes(">CITO<"));

  s = mailbox("Lunch Friday", html("Are you free?"));
  await preview(s);
  const hadFull = hasFullPanel(s.body);
  s.subject = "[simple] Lunch Friday";
  await send(s);
  report("a full signature already in the draft is swapped by [simple]", hadFull && hasPersonal(s.body) && !hasFullPanel(s.body) && (s.body.match(/id="vstx-signature"/g) || []).length === 1);
  report("the swap keeps the text and the brand bar", s.body.includes("Are you free?") && (s.body.match(/id="vstx-bar"/g) || []).length === 1);

  s = mailbox("[simple] Lunch Friday", html("Are you free?"));
  await preview(s);
  report("Apply email style honours the tag and leaves the subject alone", hasPersonal(s.body) && s.subject === "[simple] Lunch Friday");
  const before = s.body;
  await send(s);
  report("sending after that preview keeps the same signature", hasPersonal(s.body) && (s.body.match(/id="vstx-signature"/g) || []).length === 1 && s.subject === "Lunch Friday" && s.body.includes("Are you free?") && before.includes("Are you free?"));

  s = mailbox("[plain] [simple] Note", html("As typed."));
  await send(s);
  report("[plain] still wins: nothing styled, tags cleared", !s.body.includes("vstx-signature") && s.subject === "Note");

  s = mailbox("[simple] [review] Draft", html("Check me."));
  r = await send(s);
  const second = await send(s);
  report("[simple] with [review]: held once, then sent with the personal signature", r.allowEvent === false && second.allowEvent === true && hasPersonal(s.body) && s.subject === "Draft");
})();

// ---- Quoted message found in other mail apps, so the signature sits under the reply, not under the thread ----
{
  const wrapBody = (ownText, quotedHtml) => `<html><head></head><body><div style="font-family:Aptos;font-size:12pt">${ownText}</div><div><br></div>${quotedHtml}</body></html>`;
  const header = '<b>From: </b>Lawali, Mahamadou &lt;m@x.com&gt;<br><b>Date: </b>Friday, November 20, 2026 at 9:48 AM<br><b>To: </b>Sasore, Ayodele &lt;a@x.com&gt;<br><b>Subject: </b>Re: Catalog';
  const earlier = '<div>Earlier message body QUOTEDTEXT</div>';
  const cases = {
    "new Outlook for Mac container": `<div id="mail-editor-reference-message-container"><div class="ms-outlook-mobile-reference-message skipProofing" style="padding:3.0pt 0in 0in 0in;border-top:solid #B5C4DF 1.0pt">${header}</div>${earlier}</div>`,
    "Mac/classic bordered header box": `<div style="border:none;border-top:solid #B5C4DF 1.0pt;padding:3.0pt 0in 0in 0in"><p class="MsoNormal"><b><span style="font-size:12.0pt;color:black">From: </span></b><span>Lawali</span><br><b>Date: </b>Friday<br><b>To: </b>A.Y.<br><b>Subject: </b>Re: Catalog</p></div>${earlier}`,
    "bare header lines, no known marker": `<div><p><b><span>From:</span></b><span> Lawali</span></p><p><b>Sent:</b> Friday</p><p><b>To:</b> A.Y.</p><p><b>Subject:</b> Re: Catalog</p></div>${earlier}`,
    "Original Message divider": `<div>-----Original Message-----</div><div>From: Lawali<br>Sent: Friday<br>To: A.Y.<br>Subject: Re</div>${earlier}`,
    "Gmail-style forward": `<div class="gmail_quote"><div>---------- Forwarded message ---------<br>From: <b>Lawali</b><br>Date: Fri<br>Subject: Re<br>To: A.Y.</div>${earlier}</div>`,
  };
  const eng = engine({ ...mine, replyWhen: "always" });
  for (const [name, quotedHtml] of Object.entries(cases)) {
    const input = wrapBody("My reply REPLYTEXT", quotedHtml);
    const out = eng.applyVersatexStyle(input);
    const sigAt = out.indexOf('id="vstx-signature"');
    const replyAt = out.indexOf("REPLYTEXT");
    const fromAt = out.indexOf("From:");
    const quotedAt = out.indexOf("QUOTEDTEXT");
    report(`${name}: signature sits between the reply and the quoted message`, replyAt !== -1 && replyAt < sigAt && sigAt < fromAt && fromAt < quotedAt, `${replyAt}/${sigAt}/${fromAt}/${quotedAt}`);
    report(`${name}: quoted message is left exactly as it was`, out.endsWith(input.slice(input.indexOf(quotedHtml))));
    report(`${name}: counts as a reply for the personal signature`, hasPersonal(out));
  }
  const trip = eng.applyVersatexStyle(wrapBody("Flight details. From: Cincinnati. Date: Friday. To: Albany.", ""));
  report("From/Date/To typed in my own text is not mistaken for a quoted message", hasFullPanel(trip) && trip.indexOf("Albany") < trip.indexOf('id="vstx-signature"'));
  const attr = eng.applyVersatexStyle(wrapBody('See <a href="https://x.com/?q=From: a Sent: b To: c Subject: d" title="From: a Sent: b To: c Subject: d">link</a> AFTERLINK', ""));
  report("header words inside a link's attributes are ignored", attr.indexOf("AFTERLINK") < attr.indexOf('id="vstx-signature"'));
}

// ---- Toolbar buttons: Formal signature / Simple sign-off ----
(async () => {
  const manifest = fs.readFileSync(require("node:path").join(require("node:path").dirname(process.argv[2]), "manifest.xml"), "utf8");
  const wired = [...manifest.matchAll(/<Control xsi:type="Button" id="([^"]+)">[\s\S]*?<Action xsi:type="\w+">\s*<(?:FunctionName|SourceLocation)[^>]*>?([^<]*)/g)].map((m) => m[1] + ":" + m[2]);
  report("manifest has both buttons on emails and on meeting invites", ["vstxFormalButton:useFormalSignature", "vstxSimpleButton:useSimpleSignature", "vstxMeetingFormalButton:useFormalSignature", "vstxMeetingSimpleButton:useSimpleSignature"].every((b) => wired.includes(b)), wired.join(" "));
  report("both button functions are registered", /associate\("useFormalSignature"/.test(source) && /associate\("useSimpleSignature"/.test(source));
  report("every label and tooltip the buttons name is defined", [...manifest.matchAll(/resid="([^"]+)"/g)].every((m) => manifest.includes('id="' + m[1] + '"')));

  const press = (state, name) => new Promise((resolve) => live(state)[name]({ completed: resolve }));
  const count = (body) => (body.match(/id="vstx-signature"/g) || []).length;

  let s = mailbox("Lunch Friday", html("Are you free?"));
  await press(s, "useSimpleSignature");
  report("Simple sign-off on a new email -> personal signature, subject untouched", hasPersonal(s.body) && !hasFullPanel(s.body) && s.subject === "Lunch Friday" && /Simple sign-off applied/.test(s.notices.join()));
  let r = await send(s);
  report("Send keeps the signature the button chose", r.allowEvent === true && hasPersonal(s.body) && !hasFullPanel(s.body) && count(s.body) === 1);

  await press(s, "useFormalSignature");
  report("Formal signature swaps it back, still one signature", hasFullPanel(s.body) && !hasPersonal(s.body) && count(s.body) === 1 && s.body.includes("Are you free?"));
  await press(s, "useSimpleSignature");
  await press(s, "useSimpleSignature");
  report("pressing a button twice changes nothing more", hasPersonal(s.body) && count(s.body) === 1 && (s.body.match(/id="vstx-bar"/g) || []).length === 1);

  s = mailbox("RE: Budget", html("Approved.", quote(fullSig)));
  await press(s, "useFormalSignature");
  await send(s);
  report("Formal signature on a reply in a signed thread survives Send", hasFullPanel(own(s.body)) && !hasPersonal(own(s.body)) && count(own(s.body)) === 1);

  s = mailbox("RE: Budget", html("Approved.", quote(fullSig)));
  await send(s);
  report("no button pressed -> the default still applies (reply gets the personal signature)", hasPersonal(own(s.body)));
})();

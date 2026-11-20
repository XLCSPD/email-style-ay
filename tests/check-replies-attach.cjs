const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync(process.argv[2], "utf8");
const report = (name, ok, detail = "") => console.log(`${ok ? "PASS" : "FAIL"} | ${name}${detail ? " | " + detail : ""}`);
const ok = (value) => ({ status: "succeeded", value });
const failed = (message) => ({ status: "failed", error: { message } });
const blue = "font-family:Aptos;font-size:12pt;color:rgb(0,111,201)";
const html = (own, quoted = "") => `<html><head></head><body><div style="${blue}">${own}</div>${quoted}</body></html>`;
const quote = (inner) => `<div id="appendonsend"></div><hr><div id="divRplyFwdMsg">From: someone</div>${inner}`;
const mySignature = '<table id="x_vstx-signature"><tr><td><span style="font-size:16px">Reyes</span></td></tr></table>';
const theirSignature = '<table id="x_vstx-signature"><tr><td><span style="font-size:16px">Edwin</span></td></tr></table>';
const hasFullPanel = (part) => /<table id="vstx-signature" role="presentation"/.test(part);

const MY_SIGNATURE = { vstxSignature: { name: "Reyes", title: "Analytics & AI Solutions", phone: "513-555-0142" } };
function engine(roaming = MY_SIGNATURE, mailboxState = null, supported = () => true) {
  const Office = {
    actions: { associate() {} },
    AsyncResultStatus: { Succeeded: "succeeded" },
    CoercionType: { Html: "html", Text: "text" },
    MailboxEnums: { SendModeOverride: { PromptUser: "promptUser" } },
    context: {
      roamingSettings: { get: (key) => roaming[key] },
      requirements: { isSetSupported: (_name, version) => supported(version) },
      mailbox: mailboxState ? { item: mailboxState.item } : undefined,
    },
  };
  const context = { Office, Promise, String };
  vm.createContext(context);
  vm.runInContext(source, context);
  return context;
}

// ---- Short signature on replies ----
let c = engine();
const fresh = c.applyVersatexStyle(html("Hi team"));
report("new email gets the full panel", hasFullPanel(fresh) && !fresh.includes("vstx-signature-short"));

const myThread = c.applyVersatexStyle(html("Thanks, will do", quote(`<div>Earlier</div>${mySignature}`)));
const myOwn = myThread.slice(0, myThread.indexOf("appendonsend"));
report("reply in a thread you signed gets the short signature", myOwn.includes("vstx-signature-short") && !hasFullPanel(myOwn));
report("short signature: name, title, phone on one line", /Reyes<\/span><span[^>]*>&nbsp;&nbsp;\|&nbsp;&nbsp;<\/span><span[^>]*>Analytics &amp; AI Solutions<\/span>.*513&#8209;555&#8209;0142/.test(myOwn));
report("short signature keeps the yellow edge", /vstx-signature-short[\s\S]*border-left:3px solid #fdc00f/.test(myOwn));

const theirThread = c.applyVersatexStyle(html("Thanks", quote(theirSignature)));
report("someone else's signature in the thread doesn't count", hasFullPanel(theirThread.slice(0, theirThread.indexOf("appendonsend"))));

const noThread = c.applyVersatexStyle(html("Thanks", quote("<div>Plain earlier message, no signature</div>")));
report("first reply in a thread gets the full panel", hasFullPanel(noThread.slice(0, noThread.indexOf("appendonsend"))));

const off = engine({ ...MY_SIGNATURE, vstxFeatures: { shortReplies: false } }).applyVersatexStyle(html("Thanks", quote(mySignature)));
report("switch off -> full panel on replies", hasFullPanel(off.slice(0, off.indexOf("appendonsend"))));

const again = c.applyVersatexStyle(myThread);
report("short-signature reply is stable on re-send", again === myThread || c.applyVersatexStyle(again) === again);
report("re-send keeps exactly one signature", (again.slice(0, again.indexOf("appendonsend")).match(/id="vstx-signature"/g) || []).length === 1);

const bare = engine({ vstxSignature: { name: "Reyes", title: "", phone: "" } }).applyVersatexStyle(html("Ok", quote(mySignature)));
report("empty title/phone -> no separators", bare.includes("vstx-signature-short") && !/vstx-signature-short[\s\S]*?&nbsp;\|&nbsp;[\s\S]*?appendonsend/.test(bare));

// ---- Attachment check ----
function mailbox({ subject = "Q3 update", body, type = "html", attachments = [], attachmentError = false }) {
  const state = { subject, body, writes: 0 };
  state.item = {
    subject: { getAsync: (cb) => cb(ok(state.subject)), setAsync: (v, cb) => { state.subject = v; cb(ok()); } },
    body: {
      getTypeAsync: (cb) => cb(ok(type)),
      getAsync: (_t, cb) => cb(ok(state.body)),
      setAsync: (v, _o, cb) => { state.body = v; state.writes++; cb(ok()); },
    },
    getAttachmentsAsync: (cb) => cb(attachmentError ? failed("boom") : ok(attachments)),
  };
  return state;
}

async function send(state, roaming = {}, supported = () => true) {
  const ctx = engine(roaming, state, supported);
  return new Promise((resolve) => ctx.onMessageSendHandler({ completed: resolve }));
}

(async () => {
  let s = mailbox({ body: html("Please see the attached deck.") });
  let r = await send(s);
  report("mentions attachment, none attached -> Send anyway prompt", r.allowEvent === false && r.sendModeOverride === "promptUser" && /attachment/.test(r.errorMessage));
  report("email is already styled when the prompt shows", s.body.includes("vstx-signature"));

  r = await send(mailbox({ body: html("Please see the attached deck."), attachments: [{ name: "deck.pptx", isInline: false }] }));
  report("file attached -> sends", r.allowEvent === true);

  r = await send(mailbox({ body: html("Please see the attached deck."), attachments: [{ name: "image001.png", isInline: true }] }));
  report("only an inline picture -> still reminds", r.allowEvent === false);

  r = await send(mailbox({ body: html("Sounds good.", quote("<div>See attached the contract</div>")) }));
  report("'attached' only in the quoted thread -> sends", r.allowEvent === true);

  r = await send(mailbox({ body: html('Attached is the <a href="https://contoso.sharepoint.com/:x:/r/sites/f/Book.xlsx">workbook</a>') }));
  report("shared SharePoint link counts as attached", r.allowEvent === true);

  s = mailbox({ subject: "[plain] Contract", body: html("Contract enclosed.") });
  r = await send(s);
  report("[plain] email still checked", r.allowEvent === false && s.writes === 0 && s.subject === "Contract");

  r = await send(mailbox({ type: "text", body: "Report attached.\n\nThanks" }));
  report("plain-text email checked", r.allowEvent === false);

  r = await send(mailbox({ type: "text", body: "Thanks!\n\n-----Original Message-----\nFrom: X\nSee attached" }));
  report("plain-text quoted 'attached' ignored", r.allowEvent === true);

  r = await send(mailbox({ body: html("Report attached.") }), { vstxFeatures: { attachmentCheck: false } });
  report("switch off -> no reminder", r.allowEvent === true);

  r = await send(mailbox({ body: html("Report attached.") }), {}, (version) => version !== "1.14");
  report("Outlook without Send anyway support -> no reminder", r.allowEvent === true);

  r = await send(mailbox({ body: html("Report attached."), attachmentError: true }));
  report("attachment lookup fails -> email still goes", r.allowEvent === true);

  r = await send(mailbox({ subject: "[review] Q3", body: html("Report attached.") }));
  report("[review] + missing attachment -> attachment reminder wins", r.allowEvent === false && /attachment/.test(r.errorMessage));

  r = await send(mailbox({ body: html("Did you get my attachments?"), attachments: [] }));
  report("plural 'attachments' recognised", r.allowEvent === false);

  r = await send(mailbox({ body: html("The attaché arrives Monday; Attachmate is the vendor.") }));
  report("lookalike words don't trigger", r.allowEvent === true);
})();

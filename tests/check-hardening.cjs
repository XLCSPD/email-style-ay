// Edge cases and invariants: Send must always finish exactly once, and styling must never lose, reorder
// or endanger what the sender wrote, whatever the email, the settings or Outlook throw at it.
const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync(process.argv[2], "utf8");
const report = (name, ok, detail = "") => console.log(`${ok ? "PASS" : "FAIL"} | ${name}${detail ? " | " + String(detail).slice(0, 400) : ""}`);
process.on("unhandledRejection", (error) => report("no error escapes the engine unhandled", false, error && error.message));

// ---- A fake Outlook whose every call can succeed, fail or throw ----

function outlook(options = {}) {
  const calls = { completed: [], bodySet: [], subjectSet: [], notices: [] };
  const fault = (name) => (options.faults || {})[name];
  const respond = (name, callback, value) => {
    if (fault(name) === "throw") throw new Error(`${name} threw`);
    if (fault(name) === "fail") return callback({ status: "failed", error: { message: `${name} failed` } });
    callback({ status: "succeeded", value });
  };
  let body = options.body === undefined ? "<html><body><div>Hello team</div></body></html>" : options.body;
  let subject = options.subject === undefined ? "Weekly update" : options.subject;
  const item = {
    itemType: options.itemType || "message",
    subject: {
      getAsync: (callback) => respond("subject.getAsync", callback, subject),
      setAsync: (value, callback) => respond("subject.setAsync", (result) => {
        if (result.status === "succeeded") { subject = value; calls.subjectSet.push(value); }
        callback(result);
      }),
    },
    body: {
      getTypeAsync: (callback) => respond("body.getTypeAsync", callback, options.bodyType || "html"),
      getAsync: (type, callback) => respond("body.getAsync", callback, body),
      setAsync: (value, _options, callback) => respond("body.setAsync", (result) => {
        if (result.status === "succeeded") { body = value; calls.bodySet.push(value); }
        callback(result);
      }),
    },
    getAttachmentsAsync: (callback) => respond("getAttachmentsAsync", callback, options.attachments || []),
    notificationMessages: {
      replaceAsync: (id, notice, callback) => respond("replaceAsync", (result) => {
        if (result.status === "succeeded") calls.notices.push({ id, ...notice });
        callback(result);
      }),
    },
  };
  const level = options.mailbox === undefined ? "1.14" : options.mailbox;
  const minor = (version) => Number(String(version).split(".")[1] || 0);
  const settings = options.settings || {};
  const Office = {
    actions: { associate() {} },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    CoercionType: { Html: "html", Text: "text" },
    MailboxEnums: {
      SendModeOverride: { PromptUser: "promptUser" },
      ItemType: { Message: "message", Appointment: "appointment" },
      ItemNotificationMessageType: { InsightMessage: "insightMessage", InformationalMessage: "informationalMessage" },
      ActionType: { ShowTaskPane: "showTaskPane" },
    },
    context: {
      mailbox: { item, userProfile: { displayName: "Reyes, Jordan", emailAddress: options.address || "jordan.reyes@example.com" } },
      roamingSettings: { get: (key) => settings[key] },
      requirements: { isSetSupported: (set, version) => set === "Mailbox" && minor(version) <= minor(level) },
    },
  };
  const context = { Office, Promise, String };
  vm.createContext(context);
  vm.runInContext(source, context);
  const event = { completed: (outcome) => calls.completed.push(outcome || {}) };
  return { context, item, event, calls, body: () => body, subject: () => subject };
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 5));
const engine = outlook().context;
const brands = Boolean(engine.SIGNATURE_BRANDS);

async function send(options) {
  const mail = outlook(options);
  mail.context.onMessageSendHandler(mail.event);
  await settle();
  return mail;
}

// ---- Fixed inputs and a seeded generator of awkward emails ----

const wrap = (inner) => `<html><head><style>p{margin:0}</style></head><body dir="ltr">${inner}</body></html>`;
const div = (text) => `<div style="font-family:Aptos,Calibri,sans-serif;font-size:12pt;color:rgb(0,0,0)">${text}</div>`;
const quotedVersatex = `<div id="appendonsend"></div><hr style="display:inline-block;width:98%"><div id="divRplyFwdMsg"><b>From:</b> Sam Lee &lt;sam.lee@example.com&gt;</div><div><table id="vstx-bar"><tr><td>bar</td></tr></table><div>Earlier note about the vendor feed</div><table role="presentation"><tr><td><table id="vstx-signature" role="presentation"><tr><td><img src="https://versatex-tools.github.io/vstx-email-style/assets/vstx-logo-v2.png" alt="Versatex"></td><td>Sam Lee</td></tr></table></td></tr></table></div>`;

function random(seed) {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const WORDS = ["budget", "Größe", "naïve", "日本語のテキスト", "Привет", "مرحبا بالعالم", "שלום", "🚀", "Q3", "$1,200", "12.5%", "€3.4M", "vendor", "O'Brien", "A&amp;B", "&lt;tag&gt;", "x&nbsp;y", "&#8217;quoted&#8217;", "​zero", "longword".repeat(40)];
const PIECES = [
  (r, w) => div(w()),
  (r, w) => div(`<b>${w()} summary</b>`),
  (r, w) => div(`${w()} [At risk] ${w()}`),
  (r, w) => div(`${w()} [On track]`),
  (r, w) => div(`KPI: Revenue $1.2M +5% | Margin 31% -2 pts`),
  (r, w) => div(`9:00 - 9:30 Kickoff ${w()}`) + div(`9:30 - 10:15 Review ${w()}`),
  (r, w) => div(`Action: ${w()} by Friday`),
  (r, w) => div(`Risk: ${w()}`),
  (r, w) => div(`<a href="https://example.com/report?id=${Math.floor(r() * 99)}">${w()} report</a>`),
  (r, w) => div(`See <a href="https://example.com/a">${w()}</a> and <a href="https://example.org/b">https://example.org/b</a>`),
  (r, w) => div(`<a href="https://evil.example.net/login">https://example.com/login</a>`),
  (r, w) => div(`<a href="mailto:jordan.reyes@example.com" id="OWAAM1A2B3C">@Jordan Reyes</a> please check ${w()}`),
  (r, w) => div(`<a href="javascript:void(0)">${w()}</a>`),
  (r, w) => div(`<img src="cid:image001.png@01DA0000.12345678" width="120" height="40" alt="chart"> ${w()}`),
  (r, w) => div(`<img src="data:image/png;base64,iVBORw0KGgo=" alt="dot">`),
  (r, w) => `<ul><li>${w()}</li><li>${w()}<ul><li>${w()}</li><li></li></ul></li></ul>`,
  (r, w) => `<ol start="4"><li>${w()}</li><li>${w()}</li></ol>`,
  (r, w) => `<ol type="a"><li>${w()}</li><ul><li>${w()}</li></ul></ol>`,
  (r, w) => `<table border="1"><tr><th colspan="2">${w()}</th></tr><tr><td rowspan="2">${w()}</td><td>${w()}</td></tr><tr><td><table><tr><td>${w()}</td></tr></table></td></tr></table>`,
  (r, w) => `<table><tr><td></td></tr></table>`,
  (r, w) => `<p class="MsoNormal"><span lang="EN-US">${w()}<o:p></o:p></span></p><!--[if gte mso 9]><xml><o:OfficeDocumentSettings></o:OfficeDocumentSettings></xml><![endif]-->`,
  (r, w) => `<!-- note ${w()} --><div>${w()}</div>`,
  (r, w) => `<pre style="font-family:Consolas">EVALUATE TOPN(5, Sales)\n  ${w()}</pre>`,
  (r, w) => `<blockquote>${w()}</blockquote>`,
  (r, w) => `<h2>${w()}</h2>`,
  (r, w) => `<font face="Arial" color="#ff0000">${w()}</font>`,
  (r, w) => `<div><span style="font-size:20pt;color:rgb(209,52,56)">${w()}</span></div>`,
  (r, w) => `<div>${w()}<br><br><br>${w()}</div>`,
  (r, w) => `<div>${w()}`,
  (r, w) => `</div><span>${w()}</span>`,
  (r, w) => `<div dir="rtl">${w()}</div>`,
  (r, w) => `<p>${w()}</p><p>&nbsp;</p>`,
  (r, w) => `<div style='font-family:"Segoe UI";color:#0f6cbd'>${w()}</div>`,
];

function generate(seed) {
  const r = random(seed);
  const w = () => WORDS[Math.floor(r() * WORDS.length)] + " w" + Math.floor(r() * 1000);
  const parts = [div("Hi team,")];
  const count = 2 + Math.floor(r() * 8);
  for (let i = 0; i < count; i++) parts.push(PIECES[Math.floor(r() * PIECES.length)](r, w));
  parts.push(div("Thanks,"));
  const quote = r() < 0.3 ? quotedVersatex : "";
  return wrap(parts.join("") + quote);
}

// ---- Independent checks (not the engine's own word counter) ----

const ENTITIES = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " " };
function words(html) {
  const text = html
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/<(style|script|head|xml)\b[\s\S]*?<\/\1>/gi, " ")
    .replace(/<[^>]*>/g, " ")
    .replace(/&(#x?[0-9a-f]+|\w+);/gi, (m, code) => {
      if (code[0] !== "#") return ENTITIES[code.toLowerCase()] || " ";
      const n = code[1].toLowerCase() === "x" ? parseInt(code.slice(2), 16) : parseInt(code.slice(1), 10);
      return n === 11042 || n === 8209 ? " " : String.fromCodePoint(n);
    })
    .replace(/(?:https?:\/\/|www\.)\S+/gi, " ")
    .toLowerCase();
  return text.split(/[^\p{L}\p{N}]+/u).filter((word) => word && word !== "kpi" && word !== "kpis");
}
function missingWords(before, after) {
  const available = new Map();
  for (const word of words(after)) available.set(word, (available.get(word) || 0) + 1);
  const missing = [];
  for (const word of words(before)) {
    const left = available.get(word) || 0;
    if (left === 0) missing.push(word);
    else available.set(word, left - 1);
  }
  return missing;
}
const count = (html, pattern) => (html.match(pattern) || []).length;
const hrefs = (html) => [...html.matchAll(/\bhref\s*=\s*"([^"]*)"/gi)].map((m) => m[1]).sort();
const subset = (small, big) => {
  const pool = [...big];
  return small.every((x) => { const at = pool.indexOf(x); if (at === -1) return false; pool.splice(at, 1); return true; });
};
const balance = (html, tag) => count(html, new RegExp(`<${tag}\\b`, "gi")) - count(html, new RegExp(`</${tag}>`, "gi"));

function invariants(input, ctx) {
  const problems = [];
  let output;
  const started = Date.now();
  try {
    output = ctx.applyVersatexStyle(input);
  } catch (error) {
    return [`threw: ${error.message}`];
  }
  const took = Date.now() - started;
  if (typeof output !== "string") return ["did not return text"];
  if (took > 1500) problems.push(`slow: ${took} ms`);
  const quoteAt = ctx.findQuoteStart(input);
  const quoted = input.slice(quoteAt);
  if (!output.endsWith(quoted)) problems.push("quoted thread changed");
  const ownIn = input.slice(0, quoteAt);
  const ownOut = output.slice(0, output.length - quoted.length);
  const lost = missingWords(ownIn, ownOut);
  if (lost.length) problems.push(`words lost: ${lost.slice(0, 6).join(", ")}`);
  const again = ctx.applyVersatexStyle(output);
  if (again !== output) {
    let at = 0;
    while (again[at] === output[at]) at++;
    problems.push(`not stable when styled again: …${output.slice(Math.max(0, at - 60), at + 40)}… became …${again.slice(Math.max(0, at - 60), at + 40)}…`);
  }
  if (count(ownOut, /id="vstx-signature"/g) !== 1) problems.push(`${count(ownOut, /id="vstx-signature"/g)} signature panels`);
  if (count(ownOut, /id="vstx-bar"/g) !== 1) problems.push(`${count(ownOut, /id="vstx-bar"/g)} brand bars`);
  for (const [label, pattern] of [["script", /<script\b/gi], ["javascript: link", /javascript:/gi], ["event handler", /\son[a-z]+\s*=/gi]]) {
    if (count(ownOut, pattern) > count(ownIn, pattern)) problems.push(`added ${label}`);
  }
  if (!subset(hrefs(ownIn), hrefs(ownOut))) problems.push("a link address changed or disappeared");
  for (const mention of ownIn.match(/<a\b[^>]*OWAAM[^>]*>/gi) || []) if (!ownOut.includes(mention)) problems.push("mention link altered");
  for (const cid of ownIn.match(/src="cid:[^"]+"/gi) || []) if (!ownOut.includes(cid)) problems.push("inline image lost");
  for (const tag of ["table", "tr", "td"]) if (balance(ownOut, tag) !== balance(ownIn, tag)) problems.push(`unbalanced <${tag}>`);
  for (const junk of ["undefined", "NaN", "[object Object]"]) if (count(ownOut, new RegExp(junk.replace(/[[\]]/g, "\\$&"), "g")) > count(ownIn, new RegExp(junk.replace(/[[\]]/g, "\\$&"), "g"))) problems.push(`printed "${junk}"`);
  return problems;
}

(async () => {
  // ---- 1. Send always finishes, exactly once ----

  const steps = ["subject.getAsync", "body.getTypeAsync", "body.getAsync", "body.setAsync", "subject.setAsync", "getAttachmentsAsync"];
  const sendProblems = [];
  for (const step of steps) for (const how of ["fail", "throw"]) for (const mailbox of ["1.14", "1.8"]) for (const subject of ["Weekly update", "[review] Weekly update", "[plain] Weekly update"]) {
    const mail = await send({ faults: { [step]: how }, mailbox, subject, body: wrap(div("Please see attached the vendor report")) });
    if (mail.calls.completed.length !== 1) sendProblems.push(`${step} ${how}, Mailbox ${mailbox}, "${subject}": completed ${mail.calls.completed.length}x`);
  }
  report("Send finishes exactly once whichever Outlook call fails or throws (72 combinations)", sendProblems.length === 0, sendProblems.join("; "));

  const stylingFaults = ["subject.getAsync", "body.getTypeAsync", "body.getAsync", "body.setAsync", "subject.setAsync"];
  const offers = [];
  const lets = [];
  for (const step of stylingFaults) {
    const newer = await send({ faults: { [step]: "fail" }, subject: "[review] Update" });
    const done = newer.calls.completed[0] || {};
    if (!(done.allowEvent === false && done.sendModeOverride === "promptUser" && /^Couldn't apply the (?:email|Versatex) style to this email: /.test(done.errorMessage || ""))) offers.push(step);
    const older = await send({ faults: { [step]: "fail" }, subject: "[review] Update", mailbox: "1.13" });
    if ((older.calls.completed[0] || {}).allowEvent !== true) lets.push(step);
  }
  report("a styling failure offers Send anyway where Outlook supports it", offers.length === 0, offers.join(", "));
  report("a styling failure lets the email go where Send anyway isn't supported", lets.length === 0, lets.join(", "));

  const lookupFails = await send({ faults: { getAttachmentsAsync: "fail" }, body: wrap(div("The file is attached")) });
  report("a failed attachment lookup never blocks Send", lookupFails.calls.completed[0]?.allowEvent === true && lookupFails.calls.bodySet.length === 1);

  for (const [label, value] of [["undefined", undefined], ["null", null], ["a number", 42], ["empty text", ""]]) {
    const mail = outlook({ body: "placeholder" });
    mail.item.body.getAsync = (type, callback) => callback({ status: "succeeded", value });
    mail.context.onMessageSendHandler(mail.event);
    await settle();
    report(`Outlook returning ${label} as the body still finishes Send once`, mail.calls.completed.length === 1, JSON.stringify(mail.calls.completed));
  }

  const plainBody = wrap(div("Hello"));
  const plain = await send({ subject: "[PLAIN] Hello [Review]", body: plainBody });
  report("[plain] in any case sends the body untouched and clears both tags", plain.calls.bodySet.length === 0 && plain.subject() === "Hello" && plain.calls.completed[0]?.allowEvent === true, `subject "${plain.subject()}"`);
  const onlyTag = await send({ subject: "[plain]" });
  report("a subject that is only a tag ends up empty, and Send goes through", onlyTag.subject() === "" && onlyTag.calls.completed[0]?.allowEvent === true);
  const nullSubject = await send({ subject: null });
  report("a missing subject is treated as empty", nullSubject.calls.completed.length === 1 && nullSubject.calls.bodySet.length === 1);

  const review = await send({ subject: "Plan [review]" });
  const reviewDone = review.calls.completed[0] || {};
  const again = await send({ subject: review.subject(), body: review.body() });
  report("[review]: first Send stops to show the styled email, second Send goes", reviewDone.allowEvent === false && !reviewDone.sendModeOverride && again.calls.completed[0]?.allowEvent === true && again.calls.bodySet.length === 0,
    `${JSON.stringify(reviewDone)} then ${JSON.stringify(again.calls.completed[0])}`);

  const plainText = await send({ bodyType: "text", body: "Hello\nsee attached" });
  report("a plain-text email is never rewritten", plainText.calls.bodySet.length === 0 && plainText.calls.completed.length === 1);

  const inlineOnly = await send({ body: wrap(div("Chart attached below")), attachments: [{ isInline: true }] });
  const realFile = await send({ body: wrap(div("Deck attached")), attachments: [{ isInline: false }] });
  const missing = await send({ body: wrap(div("Deck attached")), attachments: [] });
  report("attachment reminder: an inline picture doesn't count as the attachment", inlineOnly.calls.completed[0]?.allowEvent === false);
  report("attachment reminder: a real attachment lets Send through", realFile.calls.completed[0]?.allowEvent === true);
  report("attachment reminder: the email is styled before the reminder shows", missing.calls.bodySet.length === 1 && missing.calls.completed[0]?.sendModeOverride === "promptUser");
  const quotedMention = await send({ body: wrap(div("Thanks, looks good") + `<div id="appendonsend"></div><div>Please see attached</div>`), attachments: [] });
  report("attachment reminder ignores 'attached' in the quoted thread", quotedMention.calls.completed[0]?.allowEvent === true);

  // ---- 2. The other two entry points finish exactly once too ----

  const previewProblems = [];
  for (const step of ["body.getTypeAsync", "body.getAsync", "body.setAsync", "replaceAsync"]) for (const how of ["fail", "throw"]) {
    const mail = outlook({ faults: { [step]: how } });
    mail.context.previewStyle(mail.event);
    await settle();
    if (mail.calls.completed.length !== 1) previewProblems.push(`${step} ${how}: ${mail.calls.completed.length}x`);
  }
  report("Apply email style finishes exactly once whatever fails", previewProblems.length === 0, previewProblems.join("; "));

  const composeProblems = [];
  for (const saved of [undefined, { name: "Jordan Reyes" }]) for (const mailbox of ["1.14", "1.9"]) for (const how of [undefined, "fail", "throw"]) for (const itemType of ["message", "appointment"]) {
    const mail = outlook({ settings: { vstxSignature: saved }, mailbox, faults: { replaceAsync: how }, itemType });
    mail.context.onNewComposeHandler(mail.event);
    await settle();
    if (mail.calls.completed.length !== 1) composeProblems.push(`saved ${Boolean(saved)}, Mailbox ${mailbox}, ${how || "ok"}, ${itemType}: ${mail.calls.completed.length}x`);
  }
  report("new email / new meeting prompt finishes exactly once in all 24 cases", composeProblems.length === 0, composeProblems.join("; "));
  const invite = outlook({ itemType: "appointment" });
  invite.context.onNewComposeHandler(invite.event);
  await settle();
  report("the meeting prompt opens the meeting settings button", invite.calls.notices[0]?.actions?.[0]?.commandId === "vstxMeetingSettingsButton");

  // ---- 3. Styling invariants on 400 generated emails ----

  const failures = [];
  const kinds = {};
  for (let seed = 1; seed <= 400; seed++) {
    const problems = invariants(generate(seed), engine);
    if (problems.length) failures.push(`seed ${seed}: ${problems.join(", ")}`);
    for (const problem of problems) {
      const kind = problem.replace(/:.*$/, "");
      kinds[kind] = kinds[kind] || { count: 0, first: seed };
      kinds[kind].count++;
    }
  }
  report("400 generated emails: no word lost, thread untouched, stable, one panel, no new risks", failures.length === 0,
    `${failures.length} failing; ` + Object.entries(kinds).map(([kind, k]) => `${kind} x${k.count} (first: seed ${k.first})`).join("; ") + (failures.length ? ` || ${failures[0]}` : ""));

  const fixed = {
    "empty text": "",
    "empty body": wrap(""),
    "whitespace only": wrap("   \n "),
    "line breaks only": wrap("<br><br>"),
    "no html or body tag": "<div>Just a fragment</div>",
    "upper-case BODY tag": "<HTML><BODY class=\"x\"><DIV>Shout</DIV></BODY></HTML>",
    "two closing body tags": "<html><body><div>One</div></body><div>Two</div></body></html>",
    "byte-order mark and null byte": "﻿<html><body><div>Start\u0000end</div></body></html>",
    "unclosed tags": wrap("<div><b>Bold <i>both <div>deeper"),
    "stray closing tags": wrap("</td></tr></table></ul><div>After the mess</div>"),
    "script and style in the body": wrap(`<style>.x{color:red}</style><script>var x = "<div>fake</div>";</script>${div("Visible words")}`),
    "comment with a tag inside": wrap(`<!-- <div>hidden</div> -->${div("Shown words")}`),
    "conditional comments": wrap(`<!--[if !mso]><!--><div>Only outside Outlook</div><!--<![endif]-->${div("Everyone sees this")}`),
    "attribute with > inside": wrap(`<div title="a > b">Arrow title</div>`),
    "single-quoted and unquoted attributes": wrap(`<div style='color:red' class=x>Mixed quoting</div>`),
    "very long word": wrap(div("x".repeat(20000))),
    "very long link": wrap(div(`<a href="https://example.com/${"p".repeat(5000)}">long link</a>`)),
    "emoji, CJK and right-to-left": wrap(div("🚀 発売 準備完了 مرحبا بالعالم שלום עולם ✅")),
    "text that looks like markup": wrap(div("Use &lt;div&gt; and a &amp;&amp; b, plus &quot;quotes&quot;")),
    "status tags in odd case": wrap(div("[AT RISK] [on TRACK] [Done] [blocked]")),
    "reply to a Versatex-styled email": wrap(div("Replying with thanks") + quotedVersatex),
    "forward with an hr divider": wrap(div("FYI below") + `<hr><div>Forwarded words</div>`),
    "list-only email": "<html><body><ul><li>Only item</li></ul></body></html>",
    "table-only email": "<html><body><table><tr><td>Only cell</td></tr></table></body></html>",
  };
  const fixedFailures = Object.entries(fixed).map(([name, html]) => [name, invariants(html, engine)]).filter(([, problems]) => problems.length);
  report(`${Object.keys(fixed).length} hand-picked awkward emails pass the same invariants`, fixedFailures.length === 0, fixedFailures.map(([n, p]) => `${n}: ${p.join(", ")}`).join("; "));

  const timed = (html) => { const t = Date.now(); engine.applyVersatexStyle(html); return Date.now() - t; };
  const bigList = wrap(`<ul>${Array.from({ length: 3000 }, (_, i) => `<li>Item ${i} [On track] $${i},000</li>`).join("")}</ul>`);
  const bigBody = wrap(Array.from({ length: 6000 }, (_, i) => div(`Line ${i} with KPI: Sales $${i}K +${i % 9}% and Action: follow up`)).join(""));
  const deep = wrap("<div>".repeat(3000) + "Deep text" + "</div>".repeat(3000));
  const bigTable = wrap(`<table>${Array.from({ length: 800 }, (_, i) => `<tr><td>${i}</td><td>Cell ${i}</td><td>[Done]</td></tr>`).join("")}</table>`);
  const times = { "3,000-item list": timed(bigList), [`${Math.round(bigBody.length / 1024)} KB email`]: timed(bigBody), "3,000 nested divs": timed(deep), "800-row table": timed(bigTable) };
  report("big and deep emails (up to about 850 KB) style in under 2 seconds each", Object.values(times).every((ms) => ms < 2000), Object.entries(times).map(([k, v]) => `${k} ${v} ms`).join(", "));
  report("big and deep emails keep every word", [bigList, deep, bigTable].every((html) => missingWords(html, engine.applyVersatexStyle(html)).length === 0));

  // ---- 4. Signature details can't break the email ----

  const payload = `"><img src=x onerror=alert(1)><script>alert(2)</script>&`;
  const injected = outlook({ settings: { vstxSignature: { name: payload, title: payload, phone: payload } } }).context.applyVersatexStyle(wrap(div("Hi")));
  report("markup typed into name, title or phone is shown as text, never run", !/<img src=x|<script>/i.test(injected) && /&lt;script&gt;/.test(injected) && /&quot;&gt;&lt;img/.test(injected));

  const corrupt = [["a string", "Jordan"], ["a number", 42], ["a list", ["Jordan"]], ["null", null], ["numbers inside", { name: 42, title: 7, phone: false }], ["a 5,000-character name (the pane allows 60)", { name: "J".repeat(5000) }], ["a prototype key", JSON.parse('{"__proto__":{"name":"x"},"name":"Jordan"}')]];
  const corruptProblems = [];
  for (const [label, saved] of corrupt) {
    try {
      const ctx = outlook({ settings: { vstxSignature: saved, vstxFeatures: saved } }).context;
      const out = ctx.applyVersatexStyle(wrap(div("Hi")));
      if (!/id="vstx-signature"/.test(out) || /undefined|\[object Object\]/.test(out)) corruptProblems.push(`${label}: bad panel`);
    } catch (error) {
      corruptProblems.push(`${label}: ${error.message}`);
    }
  }
  report("damaged saved settings fall back to defaults instead of failing", corruptProblems.length === 0, corruptProblems.join("; "));
  const featuresOff = outlook({ settings: { vstxFeatures: { headings: "no", statuses: 0 } } }).context.applyVersatexStyle(wrap(div("Feed [At risk]")));
  report("feature switches that aren't true/false keep their defaults", /At risk/.test(featuresOff) && /vstx-status|border-radius:10px/.test(featuresOff));

  const tricky = outlook({ settings: { vstxSignature: { name: "J. Reyes (PMO) [x]+?*", title: "", phone: "" } } }).context;
  const firstReply = tricky.applyVersatexStyle(wrap(div("First")));
  let secondReply;
  try { secondReply = tricky.applyVersatexStyle(wrap(div("Second") + '<div id="appendonsend"></div>' + firstReply.replace(/^[\s\S]*?<body[^>]*>|<\/body>[\s\S]*$/g, ""))); } catch (error) { secondReply = error.message; }
  report("a name full of regex symbols still gets the short reply signature", typeof secondReply === "string" && /J\. Reyes \(PMO\)/.test(secondReply), secondReply && secondReply.slice(0, 80));

  // ---- 5. Findings of the v35 review: each once lost or altered something ----

  const styled = (html, options) => (options ? outlook(options).context : engine).applyVersatexStyle(html);
  const ownOf = (html) => html.slice(0, html.search(/<table\b[^>]*\bid="vstx-signature"/));
  const linksIn = (html) => [...html.matchAll(/href="([^"]+)"/g)].map((m) => m[1]);
  const keepsLinks = (inner) => { const out = ownOf(styled(wrap(div("Hi team,") + div(inner) + div("Thanks,")))); return linksIn(inner.replace(/&amp;/g, "&")).every((href) => linksIn(out).includes(href)); };
  report("two links on one line both survive (words, a line break, addresses or a mention between)", [
    `<a href="https://a.example/report">Report</a> and <a href="https://b.example/deck">Deck</a>`,
    `<a href="https://a.example/r">https://a.example/r</a><br><a href="https://b.example/d">https://b.example/d</a>`,
    `<a href="https://a.example/r">https://a.example/r</a> 第三季度 <a href="https://b.example/zh">资料</a>`,
    `<a href="https://a.example/r">Report</a> <a href="mailto:jordan.reyes@example.com" id="OWAAM42">@Jordan Reyes</a>`,
  ].every(keepsLinks));
  const boldKeeps = [
    `<b>Agenda and <a href="https://example.com/pre-read">pre-read</a></b>`,
    `<b><a href="mailto:jordan.reyes@example.com" id="OWAAM7">@Jordan Reyes</a> please review</b>`,
    `<b><img src="cid:status@01DA" width="16" height="16"> Launch status</b>`,
    `KPI: Revenue <a href="https://example.com/rev">$1.2M</a> +6%`,
    `<img src="cid:xlsx@01DA" width="16" height="16"> <a href="https://example.sharepoint.com/:x:/r/b.xlsx">Budget</a>`,
  ].map((inner) => { const out = ownOf(styled(wrap(div("Hi team,") + div(inner) + div("Thanks,")))); return [inner, (inner.match(/href="[^"]+"|src="cid:[^"]+"|OWAAM\d+/g) || []).every((part) => out.includes(part))]; });
  report("headings, KPI tiles and buttons never drop a link, @mention or picture", boldKeeps.every(([, ok]) => ok), boldKeeps.filter(([, ok]) => !ok).map(([i]) => i).join(" | "));
  const boldList = ownOf(styled(wrap(div("Hi") + `<div>\n<ul>\n<li><b>Alpha</b></li>\n<li><b>Beta</b></li>\n</ul>\n</div>` + `<div>\n<table><tr><td><b>Region</b></td><td><b>Owner</b></td></tr></table>\n</div>` + div("Thanks"))));
  report("an all-bold list or table alone in a block stays a list or table", !/vstx-heading-text[^>]*>\s*(?:Alpha|Region)/.test(boldList) && /vstx-list/.test(boldList) && /Region[\s\S]*Owner/.test(boldList));
  const pasted = styled(wrap(div("Hi Sam, see Pat's note:") + `<table id="x_vstx-bar"><tr><td>bar</td></tr></table><div>Pat's words</div><table id="x_vstx-signature"><tr><td>Pat Lee</td></tr></table>` + div("Let me know")));
  const typedId = styled(wrap(div("The id is vstx-signature in our CSS") + "<table><tr><td>x</td></tr></table>" + div("End")));
  report("a pasted styled email or the words 'vstx-signature' don't stop your own frame", /<table id="vstx-bar"/.test(pasted) && /<table id="vstx-signature"/.test(pasted) && /<table id="vstx-signature"/.test(typedId));
  const word = styled(wrap(`<p class=MsoNormal style='font-family:"Courier New";color:#1F3864'>EVALUATE TOPN(5)</p><p class=MsoNormal style='font-size:12.0pt;mso-fareast-font-family:"Times New Roman"'>Hello team</p>`));
  const wordStyles = [...word.matchAll(/<p class=MsoNormal style="([^"]*)"/g)].map((m) => m[1]);
  report("Word's single-quoted styles keep every declaration", wordStyles.length === 2 && /^font-family:'Courier New';/.test(wordStyles[0]) && /mso-fareast-font-family:'Times New Roman'/.test(wordStyles[1]) && /^EVALUATE|>EVALUATE/.test(word.slice(word.indexOf(wordStyles[0]) + wordStyles[0].length + 2)), wordStyles.join(" || "));
  const entities = ownOf(styled(wrap(div("Hi") + div("<b>Jordan&rsquo;s Caf&eacute; plan</b>") + div("KPI: Caf&eacute; sales $4K +2%") + div("Thanks"))));
  report("character references show as characters in headings and tiles", entities.includes("Jordan’s Café plan") && entities.includes("CAFÉ SALES") && !/&amp;(?:rsquo|eacute);/i.test(entities));
  const spoofs = ["bank.example/pay", "login.bank.example", "\u200bhttps://bank.example/pay", "(https://bank.example/pay)", "HTTPS ://bank.example/pay"].map((text) => {
    const out = ownOf(styled(wrap(div("Hi") + div(`<a href="https://evil.example/collect">${text}</a>`) + div("Thanks"))));
    const label = (/class="vstx-button"[\s\S]*?<a [^>]*>([\s\S]*?)&nbsp;&#8594;<\/a>/.exec(out) || [])[1];
    return [text, !label || !/bank\.example/i.test(label)];
  });
  report("a button never shows one site and opens another", spoofs.every(([, ok]) => ok), spoofs.filter(([, ok]) => !ok).map(([t]) => t).join(", "));
  const honest = ownOf(styled(wrap(div("Hi") + div(`<a href="https://example.com/q3">example.com/q3</a>`) + div("Thanks"))));
  report("a link naming its own site keeps its text on the button", /class="vstx-button"[\s\S]*?>example\.com\/q3&nbsp;&#8594;/.test(honest));
  const cssWords = await send({ body: "<html><head><style>p{background-attachment:fixed}</style></head><body><div>Hello</div></body></html>", attachments: [] });
  report("CSS such as background-attachment doesn't set off the attachment reminder", cssWords.calls.completed[0]?.allowEvent === true);
  const nameOf = (displayName) => { const mail = outlook(); mail.context.Office.context.mailbox.userProfile.displayName = displayName; return mail.context.signatureDetails().name; };
  report("'Last, First (note)' reads 'First Last (note)'", nameOf("Reyes, Jordan (Contractor)") === "Jordan Reyes (Contractor)" && nameOf("Reyes, Jordan - Finance") === "Jordan Reyes - Finance" && nameOf("Reyes, Jordan") === "Jordan Reyes");

  const odd = [];
  const undefinedResult = outlook();
  undefinedResult.item.body.getAsync = (type, callback) => Promise.resolve().then(() => callback(undefined)).catch((error) => odd.push(`callback threw: ${error.message}`));
  undefinedResult.context.onMessageSendHandler(undefinedResult.event);
  await settle();
  if (undefinedResult.calls.completed.length !== 1 || /undefined/.test(undefinedResult.calls.completed[0]?.errorMessage || "")) odd.push("empty async result");
  const throwing = outlook();
  let attempts = 0;
  throwing.context.onMessageSendHandler({ completed: () => { attempts++; throw new Error("rejected"); } });
  await settle();
  if (attempts !== 1) odd.push(`completed threw: called ${attempts}x`);
  const unreadable = outlook();
  unreadable.context.Office.context.roamingSettings.get = () => { throw new Error("settings unavailable"); };
  unreadable.context.onNewComposeHandler(unreadable.event);
  await settle();
  if (unreadable.calls.completed.length !== 1) odd.push("compose with unreadable settings");
  report("odd Office behaviour still finishes exactly once with a readable message", odd.length === 0, odd.join("; "));

  // ---- 6. Company by address (D.E. Foxx edition only) ----

  if (brands) {
    const brandOf = (address) => outlook({ address }).context.signatureDetails().brand;
    const cases = [
      ["jordan.reyes@versatexmsp.com", "versatex"], ["JORDAN@VERSATEXMSP.COM", "versatex"], ["jordan+news@versatexmsp.com", "versatex"],
      ["jordan@defoxx.com", "defoxx"], ["jordan@versatexmsp.com.example.net", "defoxx"], ["jordan@notversatexmsp.com", "defoxx"],
      ["jordan@versatexmsp.co", "defoxx"], ["jordan@vеrsatexmsp.com", "defoxx"], ["no-at-sign", "defoxx"], ["", "defoxx"], [null, "defoxx"],
    ];
    const wrong = cases.filter(([address, want]) => brandOf(address) !== want).map(([address, want]) => `${address} -> ${brandOf(address)} (want ${want})`);
    report("company by address: exact domain, any case, lookalikes and blanks fall back to D.E. Foxx", wrong.length === 0, wrong.join("; "));
    const crossReply = outlook({ address: "jordan@defoxx.com" }).context.applyVersatexStyle(wrap(div("Thanks Sam") + quotedVersatex));
    const crossOwn = crossReply.slice(0, crossReply.indexOf('<div id="appendonsend">'));
    report("replying to a Versatex email from D.E. Foxx adds a D.E. Foxx panel and leaves the quoted Versatex one alone",
      /defoxx-logo\.png/.test(crossOwn) && crossReply.endsWith(quotedVersatex + "</body></html>"));
  }
})();

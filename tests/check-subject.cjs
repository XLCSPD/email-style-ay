const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync(process.argv[2], "utf8");
const ok = (value) => ({ status: "succeeded", value });

function mailbox(subject) {
  const state = { subject, body: '<html><body><div style="color:rgb(0,111,201);font-size:12pt">Hello</div></body></html>', bodyWrites: 0, subjectWrites: 0 };
  state.item = {
    subject: {
      getAsync: (cb) => cb(ok(state.subject)),
      setAsync: (value, cb) => { state.subject = value; state.subjectWrites++; cb(ok()); },
    },
    body: {
      getTypeAsync: (cb) => cb(ok("html")),
      getAsync: (_t, cb) => cb(ok(state.body)),
      setAsync: (html, _o, cb) => { state.body = html; state.bodyWrites++; cb(ok()); },
    },
  };
  return state;
}

function send(state) {
  const Office = {
    actions: { associate() {} },
    AsyncResultStatus: { Succeeded: "succeeded" },
    CoercionType: { Html: "html", Text: "text" },
    MailboxEnums: { SendModeOverride: { PromptUser: "promptUser" } },
    context: { mailbox: { item: state.item }, requirements: { isSetSupported: () => true } },
  };
  const context = { Office, Promise, String };
  vm.createContext(context);
  vm.runInContext(source, context);
  return new Promise((resolve) => context.onMessageSendHandler({ completed: resolve }));
}

(async () => {
  const review = mailbox("Re: [review] Q3 update");
  const first = await send(review);
  const subjectAfterFirst = review.subject;
  const second = await send(review);
  console.log(
    "[review] | 1st send allowed:", first.allowEvent,
    "| subject now:", JSON.stringify(subjectAfterFirst),
    "| styled:", review.body.includes("vstx-signature"),
    "| 2nd send allowed:", second.allowEvent
  );

  const plain = mailbox("[plain] Contract for signature");
  const result = await send(plain);
  console.log(
    "[plain]  | send allowed:", result.allowEvent,
    "| subject now:", JSON.stringify(plain.subject),
    "| body touched:", plain.bodyWrites > 0
  );

  const normal = mailbox("Weekly KPI digest");
  const normalResult = await send(normal);
  console.log(
    "normal   | send allowed:", normalResult.allowEvent,
    "| subject rewritten:", normal.subjectWrites > 0,
    "| styled:", normal.body.includes("vstx-signature")
  );

  const mixed = mailbox("RE: [Review]   Budget  [PLAIN] draft");
  await send(mixed);
  console.log("mixed    | subject now:", JSON.stringify(mixed.subject), "| body touched:", mixed.bodyWrites > 0);
})();

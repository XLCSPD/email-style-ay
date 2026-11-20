const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync(process.argv[2], "utf8");
let stored = null;
let bodyType = "html";
let writes = 0;
const item = {
  body: {
    getTypeAsync: (cb) => cb({ status: "succeeded", value: bodyType }),
    getAsync: (_type, cb) => cb({ status: "succeeded", value: stored }),
    setAsync: (html, _opts, cb) => { stored = html; writes++; cb({ status: "succeeded" }); },
  },
};
const Office = {
  actions: { associate() {} },
  AsyncResultStatus: { Succeeded: "succeeded" },
  CoercionType: { Html: "html", Text: "text" },
  context: { mailbox: { item } },
};
const context = { Office, Promise, String };
vm.createContext(context);
vm.runInContext(source, context);

const count = (s, re) => (s.match(re) || []).length;
const blueLine = (t) => `<div class="elementToProof" style="font-family:Aptos,Calibri,sans-serif;font-size:12pt;color:rgb(15,108,189)">${t}</div>`;

(async () => {
  // Reply whose quoted part already carries an older signature.
  stored =
    `<html><body>${blueLine("My reply text")}<ul><li><div>Point one</div></li></ul>` +
    '<div id="appendonsend"></div><hr><div id="divRplyFwdMsg"><b>From:</b> me</div><div>Old text</div><table id="x_vstx-signature"><tr><td>old sig</td></tr></table></body></html>';
  const first = await context.restyleCurrentItem();

  // The user adds a new line and a new bullet after the review pause.
  stored = stored.replace("My reply text</div>", `My reply text</div>${blueLine("Added after review")}`);
  stored = stored.replace("Point one</div></li>", "Point one</div></li><li><div>New point</div></li>");
  const second = await context.restyleCurrentItem();
  const third = await context.restyleCurrentItem();

  const own = stored.slice(0, stored.indexOf("appendonsend"));
  const addedLine = /<div[^>]*style="[^"]*"[^>]*>Added after review/.exec(own);
  console.log("first send:", first, "| second send:", second, "| third send:", third);
  console.log(
    "bars:", count(own, /id="vstx-bar"/g),
    "| signatures in own part:", count(own, /id="vstx-signature"/g),
    "| hexagons:", count(own, /&#11042;/g),
    "| added line restyled:", !!addedLine && /color:#231f20/.test(addedLine[0]) && !/rgb\(15,108,189\)\s*;?\s*"/.test(addedLine[0])
  );

  bodyType = "text";
  writes = 0;
  const plain = await context.restyleCurrentItem();
  console.log("plain-text email:", plain, "| body rewritten:", writes > 0);
})();

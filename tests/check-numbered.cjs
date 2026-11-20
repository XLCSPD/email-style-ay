const fs = require("node:fs");
const vm = require("node:vm");

const context = { Office: { actions: { associate() {} } }, Promise, String };
vm.createContext(context);
vm.runInContext(fs.readFileSync(process.argv[2], "utf8"), context);

const li = (t) => `<li style="font-family:Aptos;font-size:12pt"><div role="presentation" class="elementToProof">${t}</div></li>`;
const cases = {
  owaNumbered: `<html><body><div>Test the new add-in</div><ol data-editing-info="x">${li("Plan for the launch")}${li("Update the slides")}</ol><div>Thanks</div></body></html>`,
  startAt3: `<html><body><ol start="3">${li("Third")}${li("Fourth")}</ol></body></html>`,
  nested: `<html><body><ol>${li("Outer")}<li>Has child<ol>${li("Inner")}</ol></li></ol></body></html>`,
  replyWithQuotedList: `<html><body><div>Mine</div><div id="appendonsend"></div><hr><div id="divRplyFwdMsg">From</div><ol>${li("Quoted item")}</ol></body></html>`,
};

for (const [name, html] of Object.entries(cases)) {
  const out = context.applyVersatexStyle(html);
  const own = out.slice(0, out.search(/appendonsend/) === -1 ? out.length : out.search(/appendonsend/));
  const numbers = [...own.matchAll(/text-align:center"><span[^>]*>(\d+)<\/span><\/td>/g)].map((m) => m[1]);
  console.log(
    name.padEnd(20),
    "| own <ol> left:", /<ol\b/i.test(own),
    "| numbers:", numbers.join(",") || "-",
    "| yellow badges:", (own.match(/bgcolor="#fdc00f"/g) || []).length,
    "| quoted <ol> intact:", name === "replyWithQuotedList" ? /<ol>/.test(out.slice(out.search(/appendonsend/))) : "n/a"
  );
}

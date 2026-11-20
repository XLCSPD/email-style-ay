const fs = require("node:fs");
const vm = require("node:vm");

const context = { Office: { actions: { associate() {} } }, Promise, String };
vm.createContext(context);
vm.runInContext(fs.readFileSync(process.argv[2], "utf8"), context);

const samples = {
  owaReply:
    '<html><body><div style="font-size:12pt">Hi</div><ul><li><div class="x">One</div></li><li>Two</li></ul>' +
    '<div id="appendonsend"></div><hr><div id="divRplyFwdMsg"><b>From:</b> S</div><div style="font-family:Calibri">Quoted text</div></body></html>',
  classicReply:
    "<html><body lang=EN-US><div class=WordSection1><p class=MsoNormal><span style='font-size:11.0pt'>Hi</span></p>" +
    "<div style='border:none;border-top:solid #E1E1E1 1.0pt'><p class=MsoNormal><b>From:</b> S</p></div><p class=MsoNormal>Quoted</p></div></body></html>",
};

for (const [name, html] of Object.entries(samples)) {
  const out = context.applyVersatexStyle(html);
  const sigStart = out.indexOf('id="vstx-signature"');
  const afterSignature = out.slice(out.indexOf("</table>", sigStart) + "</table>".length);
  const quoteAt = out.search(/appendonsend|border-top:solid #E1E1E1/);
  console.log(
    name,
    "| bar right after <body>:", /<body[^>]*><table id="vstx-bar" role="presentation"/.test(out),
    "| signature before quote:", sigStart > 0 && sigStart < quoteAt,
    "| quote untouched:", !afterSignature.includes("Roboto"),
    "| hexagons:", (out.match(/&#11042;/g) || []).length,
    "| hexagon inside li's div:", /<div class="x" style="[^"]*"><span style="color:#/.test(out)
  );
}

const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const context = { Office: { actions: { associate() {} }, context: { roamingSettings: { get: (k) => (k === "vstxSignature" ? { name: "J.Reyes", title: "Analytics & AI Solutions | Versatex", phone: "513-555-0142" } : undefined) } } }, Promise, String };
vm.createContext(context);
vm.runInContext(fs.readFileSync(process.argv[2], "utf8"), context);
const st = "font-family:Aptos,Calibri,sans-serif;font-size:12pt;color:rgb(0,0,0)";
const d = (t) => `<div class="elementToProof" style="${st}">${t || "<br>"}</div>`;
const li = (t) => `<li style="${st}"><div role="presentation">${t}</div></li>`;
const body = [
  d("Good morning, team,"), d(""),
  d("<b>Decisions needed</b>"),
  `<div><ol style="margin-top:0"><li style="${st}">Should we adopt the new reporting tool? It needs read access to our sales data. Before approval:</li></ol></div>`,
  `<div><ul><ul style="list-style-type:circle">${li("sign the vendor's terms of service;")}${li("confirm which data the tool can read;")}${li("run a two-week trial with one team.")}</ul></ul></div>`,
  `<div><ol start="2"><li style="${st}">Should we run the pilot, and with what budget?</li></ol></div>`,
  `<div><ul><ul style="list-style-type:circle">${li("Licences for the pilot team cost about $14–16 per person; the full team would be about $41–43 per person. A $60 cap per person covers both.")}${li("The goal is 90 of 100 reports moved by March. Possible, but not likely.")}</ul></ul></div>`,
  d("<b>Spent so far:</b> About $40 on test licences. Training is free."),
  d(""), d("Thanks,"),
].join("");
const styled = context.applyVersatexStyle(`<html><head><meta charset="utf-8"></head><body>${body}</body></html>`);
const inner = styled.replace(/^[\s\S]*?<body[^>]*>|<\/body>[\s\S]*$/g, "");
const out = path.join(__dirname, "preview");
fs.mkdirSync(out, { recursive: true });
fs.writeFileSync(path.join(out, "pasted.html"), `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Pasted email</title><style>body{margin:0;background:#e9e9e9}.frame{max-width:760px;margin:24px auto;background:#fff;padding:24px 28px}</style></head><body><div class="frame">${inner}</div></body></html>`);
fs.writeFileSync(path.join(out, "pasted-word.html"), styled);
console.log("lists left:", (styled.match(/<(ul|ol|li)\b/gi) || []).length, "| words kept:", context.keepsWords(`<html><body>${body}</body></html>`, styled));

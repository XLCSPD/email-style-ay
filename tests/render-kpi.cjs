const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const context = { Office: { actions: { associate() {} }, context: {} }, Promise, String };
vm.createContext(context);
vm.runInContext(fs.readFileSync(process.argv[2], "utf8"), context);
const blue = "font-family:Aptos,Calibri,sans-serif;font-size:12pt;color:rgb(0,111,201)";
const div = (inner) => `<div class="elementToProof" style="${blue}">${inner || "<br>"}</div>`;
const body = [
  div("Hi team,"), div(""),
  div("Here is where Q3 landed:"),
  div("KPI: Spend $4.1M -8% vs plan [On track] | Savings $620K +12% [On track] | Off-contract spend 14% +3 pts [At risk]"),
  div(""),
  div("And the headline numbers without a verdict:"),
  div("KPI: Revenue $1.2M +6% | Margin 18% | Headcount 120"),
  div(""),
  div("KPIs: Suppliers 1,240 | PO cycle 3.2 days -0.4 | Compliance 92% +5% | Invoices 18,400 | EMEA -4% [Off track]"),
  div(""),
  div("Thanks,"),
].join("");
const styled = context.applyVersatexStyle(`<html><head><style>P {margin-top:0;margin-bottom:0;}</style></head><body>${body}</body></html>`);
const inner = styled.replace(/^[\s\S]*?<body[^>]*>|<\/body>[\s\S]*$/g, "");
const out = path.join(__dirname, "preview");
fs.mkdirSync(out, { recursive: true });
fs.writeFileSync(path.join(out, "kpi.html"), `<!DOCTYPE html><html><head><meta charset="utf-8"><title>KPI preview</title><style>body{margin:0;background:#e9e9e9}.frame{max-width:760px;margin:24px auto;background:#fff;padding:24px 28px}</style></head><body><div class="frame">${inner}</div></body></html>`);
fs.writeFileSync(path.join(out, "kpi-phone.html"), `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>KPI phone</title><style>body{margin:0;padding:16px;background:#fff}</style></head><body>${inner}</body></html>`);
fs.writeFileSync(path.join(out, "kpi-word.html"), styled);
console.log("ok", (styled.match(/vstx-kpi-tile/g) || []).length, "tiles");

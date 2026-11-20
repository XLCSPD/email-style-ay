const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const context = { Office: { actions: { associate() {} }, context: {} }, Promise, String };
vm.createContext(context);
vm.runInContext(fs.readFileSync(process.argv[2], "utf8"), context);

const blue = "font-family:Aptos,Calibri,sans-serif;font-size:12pt;color:rgb(0,111,201)";
const div = (inner) => `<div class="elementToProof" style="${blue}">${inner || "<br>"}</div>`;
const li = (inner) => `<li style="${blue}">${inner}</li>`;

const body = [
  div("Hi team,"),
  div(""),
  div("Quick update on the Q3 spend analytics rollout."),
  div(""),
  div("<b>Where we are</b>"),
  div("Revenue influenced is $1.2M, up 6% on plan. EMEA is behind at -4% while the vendor feed is fixed."),
  `<ul>${li("Dashboard migration [On track]")}${li("Vendor data feed [At risk]")}${li("SSO for the portal [Done]")}</ul>`,
  div(""),
  div("Decision: we keep the weekly refresh until the feed is stable."),
  div("Action: Mark to confirm the vendor SLA by Friday."),
  div("Note: figures are preliminary until month-end close."),
  div(""),
  div("<b>Next steps:</b>"),
  `<ol>${li("Validate the EMEA numbers")}${li("Share the deck with finance")}</ol>`,
  div(""),
  div('<a href="https://app.powerbi.com/groups/me/reports/123">https://app.powerbi.com/groups/me/reports/123</a>'),
  div(""),
  div("Thanks,"),
].join("");

fs.mkdirSync(path.join(path.dirname(process.argv[1]), "preview"), { recursive: true });
const html = `<html><head><style>P {margin-top:0;margin-bottom:0;}</style></head><body>${body}</body></html>`;
const styled = context.applyVersatexStyle(html);
const page = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Smart formatting preview</title>
<style>body{margin:0;background:#e9e9e9;font-family:Segoe UI,Arial}.frame{max-width:760px;margin:24px auto;background:#fff;padding:24px 28px;box-shadow:0 1px 4px rgba(0,0,0,.15)}</style>
</head><body><div class="frame">${styled.replace(/^[\s\S]*?<body[^>]*>|<\/body>[\s\S]*$/g, "")}</div></body></html>`;
fs.writeFileSync(path.join(path.dirname(process.argv[1]), "preview", "smart.html"), page);
const bodyOnly = styled.replace(/^[\s\S]*?<body[^>]*>|<\/body>[\s\S]*$/g, "");
fs.writeFileSync(path.join(path.dirname(process.argv[1]), "preview", "smart-phone.html"), `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Phone preview</title><style>body{margin:0;padding:16px;background:#fff}</style></head><body>${bodyOnly}</body></html>`);
fs.writeFileSync(path.join(path.dirname(process.argv[1]), "preview", "smart-word.html"), styled);
console.log("written", page.length);

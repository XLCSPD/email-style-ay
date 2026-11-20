const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync(process.argv[2], "utf8");
function load(roaming) {
  const mailbox = { userProfile: { displayName: "Reyes, Jordan", emailAddress: "jreyes@example.com" } };
  const Office = { actions: { associate() {} }, context: roaming ? { roamingSettings: roaming, mailbox } : { mailbox } };
  const context = { Office, Promise, String };
  vm.createContext(context);
  vm.runInContext(source, context);
  return context;
}
const roamingWith = (value) => ({ get: (key) => (key === "vstxSignature" ? value : undefined) });
const between = (s, a, b) => s.slice(s.indexOf(a), s.indexOf(b, s.indexOf(a)));

// #3 Signature settings
const plainBody = '<html><body><div style="color:rgb(0,111,201)">Hi</div></body></html>';
const sigOf = (ctx) => between(ctx.applyVersatexStyle(plainBody), 'id="vstx-signature"', "</table>");
const defaults = sigOf(load(null));
const custom = sigOf(load(roamingWith({ name: "J. Reyes", title: "Director, Analytics", phone: "513-555-0100" })));
const noTitle = sigOf(load(roamingWith({ name: "Reyes", title: "", phone: "513-555-0142" })));
const emptyName = sigOf(load(roamingWith({ name: "  ", title: "X", phone: "" })));
console.log("settings | defaults from Outlook profile:", />Jordan Reyes</.test(defaults) && !/Analytics/.test(defaults) && !/513/.test(defaults));
console.log("settings | saved values used:", /J\. Reyes/.test(custom) && /Director, Analytics/.test(custom) && /513&#8209;555&#8209;0100/.test(custom));
console.log("settings | empty title hidden:", !/Analytics/.test(noTitle) && (noTitle.match(/<br>/g) || []).length === 1);
console.log("settings | blank name falls back to profile:", />Jordan Reyes</.test(emptyName) && !/513/.test(emptyName));

// #5 Tables
const ctx = load(null);
const cell = (t, extra = "") => `<td style="border:1px solid rgb(171,171,171);padding:4px${extra}"><div class="elementToProof" style="color:rgb(0,111,201);font-size:12pt">${t}</div></td>`;
const inserted = `<table style="border-collapse:collapse"><tbody><tr>${cell("Region")}${cell("Sales")}</tr><tr>${cell("EMEA")}${cell("1.2M")}</tr><tr>${cell("APAC")}${cell("0.9M")}</tr></tbody></table>`;
const pasted = `<table class="MsoTable"><tr><td style="background:#C6EFCE;color:#006100">+12%</td></tr></table>`;
const tableEmail = `<html><body><div style="color:rgb(0,111,201)">Numbers:</div>${inserted}${pasted}</body></html>`;
const first = ctx.styleEmail(tableEmail).html;
const ownTable = between(first, "vstx-table", "</table>");
const headerRow = between(ownTable, "<tr", "</tr>");
const rows = ownTable.split("<tr").slice(1);
console.log("tables   | inserted table branded:", first.includes('class="vstx-table"'));
console.log("tables   | header black w/ yellow rule:", /bgcolor="#231f20"/.test(headerRow) && /border-bottom:3px solid #fdc00f/.test(headerRow));
console.log("tables   | header text white:", /color:#ffffff[^>]*>Region/.test(headerRow) || /color:#ffffff"?>?Region/.test(headerRow) || /<div[^>]*color:#ffffff[^>]*>Region/.test(headerRow));
console.log("tables   | stripes:", /bgcolor="#ffffff"/.test(rows[1]) && /bgcolor="#f4f4f4"/.test(rows[2]));
console.log("tables   | body text black:", /<div[^>]*color:#231f20[^>]*>EMEA/.test(first));
const pastedOut = between(first, "MsoTable", "</table>");
console.log("tables   | pasted table keeps its colours:", !pastedOut.includes("vstx-table") && /background:#C6EFCE/.test(pastedOut) && /color:#006100/.test(pastedOut) && !/bgcolor/.test(pastedOut));
const signatureAt = first.indexOf("vstx-signature");
const tableFrom = first.lastIndexOf("<table", signatureAt);
const second = ctx.styleEmail(first).html;
const third = ctx.styleEmail(second).html;
console.log("tables   | re-send stable:", second === third, "| one bar:", (second.match(/vstx-bar/g) || []).length === 1, "| signature after table:", tableFrom > first.indexOf("vstx-table"));

// #6 Meeting invite with a Teams block
const teams =
  '<div style="width:100%"><span style="white-space:nowrap;color:#5F5F5F;opacity:.36;">________________________________________________________________________________</span></div>' +
  '<div class="me-email-text" style="color:#242424;font-family:\'Segoe UI\',sans-serif"><div style="margin-bottom:24px;overflow:hidden;white-space:nowrap;"><span class="me-email-headline" style="font-size:24px;font-weight:700;color:#242424;">Microsoft Teams meeting</span></div>' +
  '<a href="https://teams.microsoft.com/l/meetup-join/x" style="font-size:14px;color:#5B5FC7">Join the meeting now</a></div>';
const invite = `<html><body><div style="color:rgb(0,111,201);font-size:12pt">Agenda: Q3 review</div><ul><li>Numbers</li><li>Next steps</li></ul>${teams}</body></html>`;
const styledInvite = ctx.applyVersatexStyle(invite);
const teamsAt = styledInvite.indexOf("________");
const sigAt = styledInvite.indexOf("vstx-signature");
console.log("invite   | signature above Teams block:", sigAt > 0 && sigAt < teamsAt);
console.log("invite   | Teams block untouched:", styledInvite.endsWith(teams + "</body></html>"));
console.log("invite   | agenda styled:", /<div[^>]*color:#231f20[^>]*>Agenda/.test(styledInvite), "| hexagons:", (styledInvite.match(/&#11042;/g) || []).length);
const teamsOnly = ctx.applyVersatexStyle(`<html><body>${teams}</body></html>`);
console.log("invite   | Teams-only invite gets bar + signature before Teams:", teamsOnly.indexOf("vstx-bar") < teamsOnly.indexOf("vstx-signature") && teamsOnly.indexOf("vstx-signature") < teamsOnly.indexOf("________"));

// The signature panel follows the sender's company; everything else is the same for both companies.
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(process.argv[2], "utf8");
const site = process.argv[3];
const report = (name, ok, detail = "") => console.log(`${ok ? "PASS" : "FAIL"} | ${name}${detail ? " | " + detail : ""}`);

function engine({ address, saved } = {}) {
  const mailbox = address === undefined ? undefined : { userProfile: { displayName: "Reyes, Jordan", emailAddress: address } };
  const context = {
    Office: { actions: { associate() {} }, context: { mailbox, roamingSettings: { get: (key) => (key === "vstxSignature" ? saved : undefined) } } },
    Promise,
    String,
  };
  vm.createContext(context);
  vm.runInContext(source, context);
  return context;
}

const blue = "font-family:Aptos;font-size:12pt;color:rgb(0,111,201)";
const email = `<html><body><div style="${blue}">Hi team,</div><ul><li style="${blue}">Feed [At risk]</li><li style="${blue}">Deck [Done]</li></ul><div style="${blue}">Thanks</div></body></html>`;
const panelOf = (html) => html.slice(html.lastIndexOf("<table", html.indexOf('id="vstx-signature"')));
const withoutPanel = (html) => html.slice(0, html.lastIndexOf("<table", html.indexOf('id="vstx-signature"')));
const isDefoxx = (panel) => /defoxx-logo\.png" width="70" height="54" alt="D\.E\. Foxx &amp; Associates"/.test(panel) && /defoxx-x\.png" width="49" height="49"/.test(panel) && !/vstx-/.test(panel.replace(/vstx-signature/g, ""));
const isVersatex = (panel) => /vstx-logo-v2\.png" width="42" height="54" alt="Versatex"/.test(panel) && /vstx-hexagons\.png" width="52" height="49"/.test(panel) && !/defoxx-(logo|x)\.png/.test(panel);

const versatex = panelOf(engine({ address: "jordan.reyes@versatexmsp.com" }).applyVersatexStyle(email));
report("versatexmsp.com sender gets the Versatex panel", isVersatex(versatex));
report("domain match ignores case", isVersatex(panelOf(engine({ address: "Jordan.Reyes@VersatexMSP.com " }).applyVersatexStyle(email))));
report("defoxx.com sender gets the D.E. Foxx panel", isDefoxx(panelOf(engine({ address: "jreyes@defoxx.com" }).applyVersatexStyle(email))));
report("any other address gets the D.E. Foxx panel", isDefoxx(panelOf(engine({ address: "jreyes@example.com" }).applyVersatexStyle(email))));
report("no Outlook profile -> D.E. Foxx panel", isDefoxx(panelOf(engine().applyVersatexStyle(email))));
report("a lookalike domain doesn't count", isDefoxx(panelOf(engine({ address: "jreyes@versatexmsp.com.example" }).applyVersatexStyle(email))));

const savedVersatex = { name: "Jordan Reyes", title: "", phone: "", brand: "versatex" };
report("a saved company wins over the address", isVersatex(panelOf(engine({ address: "jreyes@defoxx.com", saved: savedVersatex }).applyVersatexStyle(email))));
const strays = ["acme", "constructor", "__proto__", "toString", 42, null];
report("an unknown saved company falls back to the address",
  strays.every((brand) => isVersatex(panelOf(engine({ address: "jreyes@versatexmsp.com", saved: { name: "Jordan Reyes", brand } }).applyVersatexStyle(email)))));
report("signature details report the company", engine({ address: "a@versatexmsp.com" }).signatureDetails().brand === "versatex" && engine({ address: "a@defoxx.com" }).signatureDetails().brand === "defoxx");

const defoxxEmail = engine({ address: "jreyes@defoxx.com" }).applyVersatexStyle(email);
const versatexEmail = engine({ address: "jreyes@versatexmsp.com" }).applyVersatexStyle(email);
report("everything above the signature is identical for both companies", withoutPanel(defoxxEmail) === withoutPanel(versatexEmail));
report("both companies keep hexagon bullets and status dots", (defoxxEmail.match(/&#11042;/g) || []).length === 4);
report("D.E. Foxx panel keeps the yellow edge and rounded corners",
  /border-left:3px solid #fdc00f/.test(panelOf(defoxxEmail)) && /border-radius:6px 0 0 6px/.test(panelOf(defoxxEmail)) && /border-radius:0 6px 6px 0/.test(panelOf(defoxxEmail)));
const shortFor = (address) => engine({ address }).shortSignatureHtml({ name: "Jordan Reyes", title: "Analyst", phone: "1" });
report("the one-line reply signature is the same for both companies", shortFor("a@defoxx.com") === shortFor("a@versatexmsp.com"));

const ctx = engine();
const pngSize = (file) => {
  const bytes = fs.readFileSync(file);
  return { width: bytes.readUInt32BE(16), height: bytes.readUInt32BE(20) };
};
const images = Object.entries(ctx.SIGNATURE_BRANDS).flatMap(([key, brand]) => [["logo", key, brand.logo], ["motif", key, brand.motif]]);
const assetBase = /var ASSET_BASE = "([^"]+)";/.exec(source)[1];
const problems = images.flatMap(([part, key, image]) => {
  if (!image.url.startsWith(assetBase)) return [`${key} ${part} is not on the add-in's site`];
  const file = path.join(site, "assets", image.url.slice(assetBase.length));
  if (!fs.existsSync(file)) return [`${key} ${part} missing: ${path.basename(file)}`];
  const size = pngSize(file);
  const drift = Math.abs(size.width / size.height - image.width / image.height) / (image.width / image.height);
  return drift > 0.03 ? [`${key} ${part} is ${size.width}x${size.height}, drawn at ${image.width}x${image.height}`] : [];
});
report("every signature image exists on the site, in proportion", problems.length === 0, problems.join("; "));
const newImages = images.filter(([, key]) => key === "defoxx").map(([, , image]) => image);
report("D.E. Foxx images are sharp on high-density screens",
  newImages.every((image) => pngSize(path.join(site, "assets", image.url.slice(assetBase.length))).width >= image.width * 2));

const pane = fs.readFileSync(path.join(site, "taskpane.html"), "utf8");
const brandSelect = /<select id="brand"[\s\S]*?<\/select>/.exec(pane)[0];
const options = [...brandSelect.matchAll(/<option value="([^"]+)"/g)].map((m) => m[1]).sort();
report("the pane offers exactly the engine's companies", options.join(",") === Object.keys(ctx.SIGNATURE_BRANDS).sort().join(","), options.join(","));
const headerLogos = [...pane.matchAll(/src: "(assets\/[^"]+)"/g)].map((m) => m[1]);
report("every header logo in the pane exists", headerLogos.length === options.length && headerLogos.every((file) => fs.existsSync(path.join(site, file))), headerLogos.join(","));

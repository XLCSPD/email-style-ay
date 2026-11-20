const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const src = process.argv[2];
const pane = fs.readFileSync(path.join(src, "taskpane.html"), "utf8");
const commands = fs.readFileSync(path.join(src, "commands.html"), "utf8");
const engine = fs.readFileSync(path.join(src, "launchevent.js"), "utf8");
const report = (name, ok, detail = "") => console.log(`${ok ? "PASS" : "FAIL"} | ${name}${detail ? " | " + detail : ""}`);

const inline = [...pane.matchAll(/<script type="text\/javascript">([\s\S]*?)<\/script>/g)].map((m) => m[1]);
const paneScript = inline[inline.length - 1];

// A small stand-in for the page: every element with an id, its src and its plain text.
function element(id) {
  const listeners = {};
  const attributes = {};
  const tag = new RegExp(`<[a-z]+[^>]*\\bid="${id}"[^>]*>([^<]*)`).exec(pane);
  const srcAttr = tag && /\bsrc="([^"]*)"/.exec(tag[0]);
  return {
    id, value: "", checked: false, disabled: false, hidden: /\shidden[\s>]/.test(tag ? tag[0] : ""), tabIndex: 0,
    textContent: tag ? tag[1] : "", innerHTML: "", src: srcAttr ? srcAttr[1] : "", clientWidth: 0, style: {}, focused: false,
    setAttribute(name, value) { attributes[name] = String(value); },
    getAttribute(name) { return name in attributes ? attributes[name] : null; },
    focus() { this.focused = true; },
    addEventListener(type, fn) { (listeners[type] = listeners[type] || []).push(fn); },
    fire(type, extra) { (listeners[type] || []).forEach((fn) => fn({ preventDefault() {}, ...extra })); },
  };
}

function open(stored, options = {}) {
  const ids = [...pane.matchAll(/\bid="([^"]+)"/g)].map((m) => m[1]);
  const elements = Object.fromEntries(ids.map((id) => [id, element(id)]));
  Object.assign(elements["preview-frame"], { clientWidth: options.paneWidth || 0 });
  const root = element("html-root");
  const store = { ...stored };
  const pending = [];
  let saves = 0;
  const context = {
    String, Object, Promise, Math,
    VSTX_SETTINGS_PANE: true,
    setTimeout: () => 1,
    clearTimeout: () => {},
    window: { addEventListener() {} },
    document: { getElementById: (id) => elements[id] || null, documentElement: root },
    Office: {
      AsyncResultStatus: { Succeeded: "succeeded" },
      actions: { associate() { throw new Error("pane must not associate actions"); } },
      onReady(fn) { fn(); },
      context: {
        mailbox: { userProfile: { displayName: "Reyes, Jordan", emailAddress: options.email || "jreyes@example.com" } },
        roamingSettings: {
          get: (key) => store[key],
          set: (key, value) => { store[key] = JSON.parse(JSON.stringify(value)); },
          saveAsync: (cb) => {
            saves++;
            if (options.deferSaves) pending.push(cb);
            else cb({ status: options.failSaves ? "failed" : "succeeded", error: { message: "offline" } });
          },
        },
      },
    },
  };
  if (options.theme === "throws") {
    Object.defineProperty(context.Office.context, "officeTheme", { get() { throw new Error("not supported"); } });
  } else if (options.theme) {
    context.Office.context.officeTheme = options.theme;
  }
  vm.createContext(context);
  vm.runInContext(engine, context);
  vm.runInContext(paneScript, context);
  const settle = () => pending.shift()({ status: "succeeded" });
  return { elements, store, root, saves: () => saves, pending, settle, context };
}

const keys = Object.keys(open({}).context.FEATURE_DEFAULTS);
report("a switch for every feature", keys.every((key) => new RegExp(`class="switch"[^>]*id="${key}"`).test(pane)), keys.join(","));
report("every switch has a label the save message can name", keys.every((key) => new RegExp(`id="${key}-name" for="${key}">[^<]+<`).test(pane)));
report("pane and Send handler load the same engine version", /launchevent\.js\?v=(\d+)/.exec(pane)[1] === /launchevent\.js\?v=(\d+)/.exec(commands)[1]);

let p = open({});
report("fresh install: switches match the defaults (mention tags off)", keys.every((key) => p.elements[key].checked === p.context.FEATURE_DEFAULTS[key]) && !p.elements.mentions.checked);
report("fresh install: name from the Outlook profile, flipped to First Last", p.elements.name.value === "Jordan Reyes", p.elements.name.value);
report("fresh install: no one else's title or phone", p.elements.title.value === "" && p.elements.phone.value === "");
report("fresh install: welcome banner shown", p.elements.welcome.hidden === false);
report("opens on the Signature tab", !p.elements["panel-signature"].hidden && p.elements["panel-formatting"].hidden && p.elements["panel-help"].hidden);

// Tabs
p.elements["tab-formatting"].fire("click");
report("tab click shows its panel only", p.elements["panel-signature"].hidden && !p.elements["panel-formatting"].hidden && p.elements["tab-formatting"].getAttribute("aria-selected") === "true" && p.elements["tab-signature"].getAttribute("aria-selected") === "false");
p.elements["tab-formatting"].fire("keydown", { key: "End" });
report("End key moves to the last tab and focuses it", !p.elements["panel-help"].hidden && p.elements["tab-help"].focused && p.elements["tab-help"].tabIndex === 0 && p.elements["tab-formatting"].tabIndex === -1);
p.elements["tab-help"].fire("keydown", { key: "ArrowRight" });
report("ArrowRight wraps around to the first tab", !p.elements["panel-signature"].hidden);

// Switches save as soon as they're flipped
p = open({ vstxFeatures: { headings: false, figures: false }, vstxSignature: { name: "Reyes", title: "", phone: "1" } });
report("saved switches restored", !p.elements.headings.checked && !p.elements.figures.checked && p.elements.callouts.checked && p.elements.buttons.checked);
report("saved signature restored", p.elements.title.value === "" && p.elements.phone.value === "1");
p.elements.name.value = "Someone Else";
p.elements.name.fire("input");
p.elements.statuses.checked = false;
p.elements.statuses.fire("change");
report("flipping a switch saves it at once", p.saves() === 1 && p.store.vstxFeatures.statuses === false && p.store.vstxFeatures.headings === false && p.store.vstxFeatures.callouts === true);
report("flipping a switch doesn't save a half-typed signature", p.store.vstxSignature.name === "Reyes");
report("switch message names the feature", p.elements.status.textContent === "Status tags off. Saved for your next email." && p.elements.status.getAttribute("data-show") === "true", p.elements.status.textContent);

p = open({}, { deferSaves: true });
p.elements.kpis.checked = false;
p.elements.kpis.fire("change");
p.elements.agenda.checked = false;
p.elements.agenda.fire("change");
p.elements.figures.checked = false;
p.elements.figures.fire("change");
report("one save at a time while flipping quickly", p.saves() === 1);
p.settle();
report("changes made mid-save are saved right after", p.saves() === 2 && p.store.vstxFeatures.kpis === false && p.store.vstxFeatures.agenda === false && p.store.vstxFeatures.figures === false);
p.settle();
report("last message wins", /^Key figures off\./.test(p.elements.status.textContent), p.elements.status.textContent);

p = open({}, { deferSaves: true });
p.elements.kpis.checked = false;
p.elements.kpis.fire("change");
p.elements["signature-form"].fire("submit");
p.elements.agenda.checked = false;
p.elements.agenda.fire("change");
p.settle();
p.settle();
report("a signature save queued behind switches still clears the welcome", p.elements.welcome.hidden === true && p.store.vstxSignature.name === "Jordan Reyes");

p = open({}, { failSaves: true });
p.elements.kpis.checked = false;
p.elements.kpis.fire("change");
report("a failed save says so and stays on screen", /Couldn't save: offline/.test(p.elements.status.textContent) && p.elements.status.getAttribute("data-tone") === "error");

// Signature
p = open({});
p.elements.title.value = "Head of Analytics";
p.elements.title.fire("input");
p.elements["signature-form"].fire("submit");
report("Save writes the signature once", p.saves() === 1 && p.store.vstxSignature.name === "Jordan Reyes" && p.store.vstxSignature.title === "Head of Analytics");
report("signature message", /Signature saved\. Your next email uses it\./.test(p.elements.status.textContent), p.elements.status.textContent);
report("welcome banner goes once the signature is saved", p.elements.welcome.hidden === true);

p = open({ vstxFeatures: Object.fromEntries(keys.map((key) => [key, false])) });
report("engine reads what the pane saved", Object.values(p.context.featureSettings()).every((on) => on === false));

p = open({});
p.elements.name.value = "  ";
p.elements.name.fire("input");
p.elements["signature-form"].fire("submit");
report("empty name blocks save", p.elements.save.disabled && p.saves() === 0);

p = open({});
report("short signature preview rendered", /vstx-signature-short/.test(p.elements["preview-short"].innerHTML) && /Reyes/.test(p.elements["preview-short"].innerHTML));
p.elements.title.value = "Head of Analytics";
p.elements.title.fire("input");
report("previews follow typing", /Head of Analytics/.test(p.elements["preview-short"].innerHTML) && /Head of Analytics/.test(p.elements.preview.innerHTML));

p = open({ vstxSignature: { name: "J.Reyes", title: "Analytics", phone: "1" } });
report("set up: welcome banner hidden", p.elements.welcome.hidden === true);

// Company: picked from the address until the sender chooses, and saved with the signature
const headerIs = (p, file) => p.elements["brand-logo"].src === "assets/" + file;
p = open({});
report("fresh install: company from the address (D.E. Foxx)", p.elements.brand.value === "defoxx" && headerIs(p, "defoxx-horizontal-reversed.png") && /defoxx-logo\.png/.test(p.elements.preview.innerHTML));
p = open({}, { email: "jordan.reyes@versatexmsp.com" });
report("fresh install at versatexmsp.com: Versatex company, header and preview", p.elements.brand.value === "versatex" && headerIs(p, "vstx-horizontal-light.png") && /vstx-logo-v2\.png/.test(p.elements.preview.innerHTML));
p = open({});
p.elements.brand.value = "versatex";
p.elements.brand.fire("change");
report("changing the company redraws the preview and the header", /vstx-logo-v2\.png/.test(p.elements.preview.innerHTML) && headerIs(p, "vstx-horizontal-light.png") && p.elements["brand-logo"].alt === "Versatex");
report("changing the company alone doesn't save", p.saves() === 0);
p.elements["signature-form"].fire("submit");
report("Save stores the company with the signature", p.store.vstxSignature.brand === "versatex" && p.store.vstxSignature.name === "Jordan Reyes");
p = open({ vstxSignature: { name: "Jordan Reyes", title: "", phone: "", brand: "versatex" } });
report("a saved company wins over the address", p.elements.brand.value === "versatex" && headerIs(p, "vstx-horizontal-light.png"));

// Signature preview drawn at phone width, scaled to the pane
p = open({}, { paneWidth: 288 });
report("narrow pane: preview drawn at 360px and scaled down", p.elements.preview.style.width === "360px" && p.elements.preview.style.zoom === "0.8", JSON.stringify(p.elements.preview.style));
p = open({}, { paneWidth: 420 });
report("wide pane: preview at full size", p.elements.preview.style.width === "420px" && p.elements.preview.style.zoom === "1", JSON.stringify(p.elements.preview.style));

// Live examples come from the styling engine, whatever the switches say
const COMPONENTS = { headings: "vstx-heading", plainTitles: "vstx-heading", callouts: "vstx-callout", figures: "vstx-figure", buttons: "vstx-button", statuses: "vstx-status", kpis: "vstx-kpi", mentions: "vstx-mention", agenda: "vstx-agenda" };
p = open({ vstxFeatures: Object.fromEntries(keys.map((key) => [key, false])) });
const missing = Object.keys(COMPONENTS).filter((key) => !p.elements["example-" + key].innerHTML.includes(COMPONENTS[key]) || !p.elements["typed-" + key].innerHTML);
report("every example is drawn by the engine, even with its switch off", missing.length === 0, missing.join(","));
report("examples can't be clicked or tabbed into", Object.keys(COMPONENTS).every((key) => new RegExp(`id="example-${key}" inert`).test(pane) && new RegExp(`id="typed-${key}" inert`).test(pane)));
report("callout example uses the light design", /vstx-callout-label/.test(p.elements["example-callouts"].innerHTML) && !/bgcolor="#231f20"/.test(p.elements["example-callouts"].innerHTML));

// Version label
p = open({});
report("header shows the running version", p.elements.version.textContent === "Active · v" + /launchevent\.js\?v=(\d+)/.exec(pane)[1], p.elements.version.textContent);

// Theme
report("Outlook dark theme -> dark pane", open({}, { theme: { isDarkTheme: true } }).root.getAttribute("data-theme") === "dark");
report("Outlook light theme -> light pane", open({}, { theme: { isDarkTheme: false } }).root.getAttribute("data-theme") === "light");
report("no isDarkTheme: judged from the background colour", open({}, { theme: { bodyBackgroundColor: "#1f1f1f" } }).root.getAttribute("data-theme") === "dark" && open({}, { theme: { bodyBackgroundColor: "#FFFFFF" } }).root.getAttribute("data-theme") === "light");
p = open({});
report("no theme from Outlook: the system setting decides", p.root.getAttribute("data-theme") === null);
p = open({}, { theme: "throws" });
report("theme lookup failing doesn't break the pane", p.root.getAttribute("data-theme") === null && p.elements.name.value === "Jordan Reyes");

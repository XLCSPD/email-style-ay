const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync(process.argv[2], "utf8");
const manifest = fs.readFileSync(process.argv[3], "utf8");
const report = (name, ok, detail = "") => console.log(`${ok ? "PASS" : "FAIL"} | ${name}${detail ? " | " + detail : ""}`);

function open({ saved, itemType = "message", supported = () => true }) {
  const shown = [];
  const Office = {
    actions: { associate() {} },
    AsyncResultStatus: { Succeeded: "succeeded" },
    MailboxEnums: {
      ItemType: { Message: "message", Appointment: "appointment" },
      ItemNotificationMessageType: { InsightMessage: "insightMessage" },
      ActionType: { ShowTaskPane: "showTaskPane" },
    },
    context: {
      roamingSettings: { get: (key) => (key === "vstxSignature" ? saved : undefined) },
      requirements: { isSetSupported: (_set, version) => supported(version) },
      mailbox: {
        userProfile: { displayName: "Ng, Alex" },
        item: { itemType, notificationMessages: { replaceAsync: (key, details, cb) => { shown.push({ key, details }); cb({ status: "succeeded" }); } } },
      },
    },
  };
  const context = { Office, Promise, String };
  vm.createContext(context);
  vm.runInContext(source, context);
  let completed = 0;
  context.onNewComposeHandler({ completed: () => completed++ });
  return { shown, completed, context };
}

let r = open({ saved: undefined });
const note = r.shown[0] && r.shown[0].details;
report("not set up -> prompt shown once", r.shown.length === 1 && r.completed === 1);
report("prompt is an insight with one 'Open settings' action", note && note.type === "insightMessage" && note.actions.length === 1 && note.actions[0].actionText === "Open settings");
report("action opens the message settings button", note && note.actions[0].actionType === "showTaskPane" && note.actions[0].commandId === "vstxSettingsButton");
report("message fits Outlook's 150-character limit", note && note.message.length <= 150, String(note && note.message.length));

r = open({ saved: undefined, itemType: "appointment" });
report("new invite -> meeting settings button", r.shown[0] && r.shown[0].details.actions[0].commandId === "vstxMeetingSettingsButton");

r = open({ saved: { name: "Alex Ng", title: "", phone: "" } });
report("set up -> no prompt", r.shown.length === 0 && r.completed === 1);

r = open({ saved: undefined, supported: (v) => v !== "1.10" });
report("older Outlook -> no prompt, no error", r.shown.length === 0 && r.completed === 1);

report("new user's signature uses their Outlook name", /Alex Ng/.test(r.context.applyVersatexStyle("<html><body><div>Hi</div></body></html>")));

const buttons = [...manifest.matchAll(/<Control xsi:type="Button" id="([^"]+)">[\s\S]*?<Action xsi:type="(\w+)"/g)].map((m) => m[1] + ":" + m[2]);
report("both settings buttons exist and open the task pane", buttons.includes("vstxSettingsButton:ShowTaskpane") && buttons.includes("vstxMeetingSettingsButton:ShowTaskpane"), buttons.join(" "));
report("manifest wires the new-email events to the handler", /OnNewMessageCompose" FunctionName="onNewComposeHandler"/.test(manifest) && /OnNewAppointmentOrganizer" FunctionName="onNewComposeHandler"/.test(manifest));
report("handler is registered in the add-in", /Office\.actions\.associate\("onNewComposeHandler"/.test(source));
report("prompt icon exists in the manifest", /<bt:Image id="Icon.16x16"/.test(manifest));

{
  const assetBase = /var ASSET_BASE = "([^"]+)assets\/";/.exec(source);
  const site = assetBase && assetBase[1];
  const urls = [...manifest.matchAll(/DefaultValue="(https:\/\/[^"]+)"/g)].map((m) => m[1]);
  const elsewhere = urls.filter((url) => !url.startsWith(site));
  report("manifest and email images come from the same site", Boolean(site) && urls.length > 0 && elsewhere.length === 0, elsewhere.join(" "));
  report("manifest AppDomain is that site's origin", Boolean(site) && manifest.includes(`<AppDomain>${new URL(site).origin}</AppDomain>`));
}

{
  const thrower = fs.readFileSync(process.argv[2], "utf8");
  const Office = {
    actions: { associate() {} },
    MailboxEnums: { ItemType: { Appointment: "appointment" }, ItemNotificationMessageType: { InsightMessage: "insightMessage" }, ActionType: { ShowTaskPane: "showTaskPane" } },
    context: {
      roamingSettings: { get: () => undefined },
      requirements: { isSetSupported: () => true },
      mailbox: { item: { itemType: "message", notificationMessages: { replaceAsync: () => { throw new Error("actions not supported"); } } } },
    },
  };
  const ctx = { Office, Promise, String };
  vm.createContext(ctx);
  vm.runInContext(thrower, ctx);
  let completed = 0;
  ctx.onNewComposeHandler({ completed: () => completed++ });
  report("Outlook rejecting the prompt still completes the event", completed === 1);
}

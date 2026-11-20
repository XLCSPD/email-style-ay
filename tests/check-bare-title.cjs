const fs = require("node:fs"), vm = require("node:vm");
const c = { Office: { actions: { associate() {} } }, Promise, String }; vm.createContext(c);
vm.runInContext(fs.readFileSync(process.argv[2], "utf8"), c);
const out = c.applyVersatexStyle('<html><body><ol><li style="color:rgb(0,111,201);font-size:12pt"><strong>Alpha Sector Operations</strong></li><ul><li>Sub A</li></ul></ol></body></html>');
const m = /<div style="([^"]*)"><strong>Alpha Sector Operations/.exec(out);
console.log("bare title wrapped in black paragraph:", !!m && /color:#231f20/.test(m[1]));
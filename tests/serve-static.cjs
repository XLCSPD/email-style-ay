const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");

const [root, port] = [path.resolve(process.argv[2]), Number(process.argv[3] || 3100)];
const types = { ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".png": "image/png", ".woff2": "font/woff2", ".txt": "text/plain" };

http
  .createServer((req, res) => {
    const file = path.join(root, decodeURIComponent(new URL(req.url, "http://localhost").pathname));
    if (!file.startsWith(root)) return res.writeHead(403).end();
    fs.readFile(file, (err, data) => {
      if (err) return res.writeHead(404).end();
      res.writeHead(200, { "Content-Type": types[path.extname(file)] || "application/octet-stream", "Cache-Control": "no-store" });
      res.end(data);
    });
  })
  .listen(port, "127.0.0.1", () => console.log(`serving ${root} on http://localhost:${port}`));

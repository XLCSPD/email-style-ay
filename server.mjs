import { createServer } from "node:https";
import { readFile } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import devCerts from "office-addin-dev-certs";

const PORT = 3000;
const ROOT = join(fileURLToPath(new URL(".", import.meta.url)), "site");
const CONTENT_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".png": "image/png",
};

const httpsOptions = await devCerts.getHttpsServerOptions();

createServer(httpsOptions, async (req, res) => {
  const path = normalize(decodeURIComponent(new URL(req.url, "https://localhost").pathname));
  const file = join(ROOT, path === "/" ? "commands.html" : path);
  if (!file.startsWith(ROOT)) {
    res.writeHead(403).end();
    return;
  }
  try {
    const body = await readFile(file);
    res.writeHead(200, {
      "Content-Type": CONTENT_TYPES[extname(file)] ?? "application/octet-stream",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*",
    });
    res.end(body);
    console.log(`200 ${path}`);
  } catch {
    res.writeHead(404).end();
    console.log(`404 ${path}`);
  }
}).listen(PORT, () => console.log(`Serving ${ROOT} at https://localhost:${PORT}`));

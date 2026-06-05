import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PUBLIC_DIR = path.resolve(__dirname, "public");
const EXCEL_FILE_RE = /^NANG SUAT.*\.xlsx$/i;

function listExcelFiles(dir) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir).filter((fileName) => EXCEL_FILE_RE.test(fileName));
}

function resolveExcelFile(fileName) {
  for (const dir of [PUBLIC_DIR, __dirname]) {
    const filePath = path.resolve(dir, fileName);
    if (fs.existsSync(filePath)) return filePath;
  }
  return null;
}

function serveExcelFiles() {
  return {
    name: "serve-excel-files",

    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const pathname = decodeURIComponent((req.url || "").split("?")[0]);
        const fileName = pathname.replace(/^\//, "");

        if (!EXCEL_FILE_RE.test(fileName) || fileName.includes("/") || fileName.includes("\\")) {
          next();
          return;
        }

        const filePath = resolveExcelFile(fileName);
        if (!filePath) {
          res.statusCode = 404;
          res.end(`${fileName} not found`);
          return;
        }

        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Cache-Control", "no-store");
        fs.createReadStream(filePath).pipe(res);
      });

      for (const dir of [PUBLIC_DIR, __dirname]) {
        listExcelFiles(dir).forEach((fileName) => {
          server.watcher.add(path.resolve(dir, fileName));
        });
      }

      server.watcher.on("change", (changedPath) => {
        if (EXCEL_FILE_RE.test(path.basename(changedPath))) {
          server.ws.send({ type: "full-reload", path: "*" });
        }
      });
    },

    generateBundle() {
      const files = listExcelFiles(__dirname).filter(
        (fileName) => !fs.existsSync(path.resolve(PUBLIC_DIR, fileName))
      );

      for (const fileName of files) {
        const filePath = path.resolve(__dirname, fileName);
        if (!fs.existsSync(filePath)) continue;
        this.emitFile({
          type: "asset",
          fileName,
          source: fs.readFileSync(filePath),
        });
      }
    },
  };
}

export default defineConfig({
  plugins: [react(), serveExcelFiles()],
  base: "./",
});

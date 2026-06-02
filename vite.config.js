import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const excelFileName = "NANG SUAT.xlsx";
const excelFilePath = path.resolve(__dirname, excelFileName);

function serveExcelFromRoot() {
  return {
    name: "serve-excel-from-root",
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const pathname = decodeURIComponent((req.url || "").split("?")[0]);

        if (pathname !== `/${excelFileName}`) {
          next();
          return;
        }

        if (!fs.existsSync(excelFilePath)) {
          res.statusCode = 404;
          res.end(`${excelFileName} not found`);
          return;
        }

        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Cache-Control", "no-store");
        fs.createReadStream(excelFilePath).pipe(res);
      });

      server.watcher.add(excelFilePath);
      server.watcher.on("change", (changedPath) => {
        if (path.resolve(changedPath) === excelFilePath) {
          server.ws.send({ type: "full-reload", path: "*" });
        }
      });
    },
    generateBundle() {
      if (!fs.existsSync(excelFilePath)) return;

      this.emitFile({
        type: "asset",
        fileName: excelFileName,
        source: fs.readFileSync(excelFilePath),
      });
    },
  };
}

export default defineConfig({
  plugins: [react(), serveExcelFromRoot()],
  base: "./",
});

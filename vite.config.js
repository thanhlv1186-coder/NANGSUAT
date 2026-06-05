import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function serveExcelFiles() {
  // Lấy tất cả file Excel trong thư mục gốc (NANG SUAT*.xlsx)
  function getExcelFiles() {
    return fs.readdirSync(__dirname).filter(f => /^NANG SUAT.*\.xlsx$/i.test(f));
  }

  return {
    name: "serve-excel-files",

    // Dev server: serve tất cả file Excel từ root
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const pathname = decodeURIComponent((req.url || "").split("?")[0]);
        const fileName = pathname.replace(/^\//, "");

        if (!/^NANG SUAT.*\.xlsx$/i.test(fileName)) {
          next();
          return;
        }

        const filePath = path.resolve(__dirname, fileName);
        if (!fs.existsSync(filePath)) {
          res.statusCode = 404;
          res.end(`${fileName} not found`);
          return;
        }

        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Cache-Control", "no-store");
        fs.createReadStream(filePath).pipe(res);
      });

      // Watch tất cả file Excel
      getExcelFiles().forEach(f => {
        const filePath = path.resolve(__dirname, f);
        server.watcher.add(filePath);
      });
      server.watcher.on("change", (changedPath) => {
        if (/NANG SUAT.*\.xlsx$/i.test(changedPath)) {
          server.ws.send({ type: "full-reload", path: "*" });
        }
      });
    },

    // Build: copy tất cả file Excel vào dist
    generateBundle() {
      const files = getExcelFiles();
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

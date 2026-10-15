import fs from "node:fs";
import { defineConfig } from "vite";
import { viteSingleFile } from "vite-plugin-singlefile";

const isDev = process.env.NODE_ENV === "development";

const pkg = JSON.parse(fs.readFileSync(new URL("./package.json", import.meta.url), "utf8"));

export default defineConfig({
  define: { __APP_VERSION__: JSON.stringify(pkg.version) },
  plugins: [viteSingleFile()],
  build: {
    sourcemap: isDev ? "inline" : undefined,
    minify: !isDev,
    rollupOptions: { input: "mcp-app.html" },
    outDir: "dist",
    emptyOutDir: false,
  },
});

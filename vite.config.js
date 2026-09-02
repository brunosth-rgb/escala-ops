import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  build: { outDir: "dist", emptyOutDir: true },
  // durante "npm run dev", manda as chamadas de dados para o servidor
  server: { proxy: { "/api": "http://localhost:8080" } },
});

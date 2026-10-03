import express from "express";
import cors from "cors";
import path from "path";
import dotenv from "dotenv";
import { apiRouter } from "./src/server/routes.ts";

dotenv.config();

const app = express();
const port = 3000;
const isProd = process.env.NODE_ENV === "production";

app.set("trust proxy", true);
app.use(cors());
app.use(express.json({ limit: "20mb" }));
app.use(express.urlencoded({ extended: true, limit: "20mb" }));

// Mount API router
app.use("/api", apiRouter);

// Health check endpoint for external pingers / keep-alive services
app.get("/healthz", (_req, res) => {
  res.json({ status: "ok", timestamp: new Date().toISOString(), uptime: process.uptime() });
});

// Self-ping worker to prevent free tier spin-down (e.g. Render, Railway, Fly)
function startKeepAliveWorker(serverPort: number) {
  const PING_INTERVAL = 4 * 60 * 1000; // 4 minutes (Render sleeps at 15m of inactivity)
  setInterval(async () => {
    try {
      const endpoints = [
        process.env.RENDER_EXTERNAL_URL ? `${process.env.RENDER_EXTERNAL_URL}/api/health` : null,
        process.env.APP_URL ? `${process.env.APP_URL}/api/health` : null,
        `http://127.0.0.1:${serverPort}/api/health`,
      ].filter(Boolean) as string[];

      for (const url of endpoints) {
        try {
          await fetch(url, { signal: AbortSignal.timeout(5000) });
        } catch {
          // quiet fallback
        }
      }
    } catch {
      // Ignored - quiet keep-alive ping
    }
  }, PING_INTERVAL);
}

async function startServer() {
  if (!isProd) {
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: {
        middlewareMode: true,
        hmr: process.env.DISABLE_HMR !== "true",
      },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.resolve(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (_req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(port, "0.0.0.0", () => {
    console.log(`Server listening on http://0.0.0.0:${port} (${isProd ? "production" : "development"})`);
    startKeepAliveWorker(port);
  });
}

startServer().catch((err) => {
  console.error("Failed to start server:", err);
  process.exit(1);
});

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import express from "express";
import path from "path";
import app from "./api/index";

const PORT = 3000;

// Serve frontend assets in development and production
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
    console.log("Vite Development Middleware bound successfully.");
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
    console.log("Production static distribution bound successfully.");
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Presentation master server online on port ${PORT}`);
  });
}

startServer();

export default app;

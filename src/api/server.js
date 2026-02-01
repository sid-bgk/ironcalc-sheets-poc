import express from "express";
import { router, initWorkbook } from "./routes.js";

const app = express();

// Middleware
app.use(express.json());

// Routes
app.use(router);

// Error handler for malformed JSON
app.use((err, req, res, next) => {
  if (err instanceof SyntaxError && err.status === 400 && "body" in err) {
    console.log(`[server] Malformed JSON in request`);
    return res.status(400).json({
      error: {
        code: "INVALID_JSON",
        message: "Request body contains invalid JSON",
      },
    });
  }
  next(err);
});

const PORT = process.env.PORT || 3000;
const EXCEL_PATH =
  process.env.EXCEL_FILE_PATH || "./DSCR_NoArrayFormulas_DEV_CLEAN.xlsx";

/**
 * Start the Express server
 */
export function startServer() {
  // Load workbook before starting server
  initWorkbook(EXCEL_PATH);

  app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
    console.log(`POST /api/v1/calculate/dscr`);
  });
}

export { app, PORT };

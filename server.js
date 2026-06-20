const express = require("express");
const cors    = require("cors");
const net     = require("net");
const fs      = require("fs");
const os      = require("os");
const path    = require("path");
const { execSync, execFile } = require("child_process");

const app     = express();
const PORT    = process.env.PORT || 3001;
const VERSION = "1.1.0";

// SumatraPDF path — set SUMATRA_PATH env var if not on system PATH
const SUMATRA = process.env.SUMATRA_PATH || "SumatraPDF";

app.use(cors());
app.use(express.json({ limit: "2mb" }));

// ── Puppeteer browser (singleton, reused across requests) ─────────────────────

let browser = null;

async function getBrowser() {
  if (browser) return browser;
  const puppeteer = require("puppeteer");
  browser = await puppeteer.launch({
    headless: "new",
    args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-gpu"],
  });
  browser.on("disconnected", () => {
    console.log("Browser disconnected — will relaunch on next print");
    browser = null;
  });
  console.log("Puppeteer browser ready");
  return browser;
}

// Pre-warm on startup
getBrowser().catch((err) => console.error("Browser pre-warm failed:", err.message));

// ── Health ────────────────────────────────────────────────────────────────────

app.get("/health", (_req, res) => {
  res.json({ ok: true, version: VERSION });
});

// ── List printers ─────────────────────────────────────────────────────────────

app.get("/printers", (_req, res) => {
  try {
    const buf = execSync(
      'powershell -NoProfile -Command "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Get-Printer | Select-Object -ExpandProperty Name"',
      { timeout: 5000, encoding: "buffer" }
    );
    const names = buf
      .toString("utf8")
      .split(/\r?\n/)
      .map((s) => s.trim())
      .filter(Boolean);
    res.json({ printers: names });
  } catch (err) {
    console.error("Could not list printers:", err.message);
    res.json({ printers: [], warning: err.message });
  }
});

// ── Print ─────────────────────────────────────────────────────────────────────

app.post("/print", async (req, res) => {
  const { printerName, html } = req.body || {};

  if (!printerName)
    return res.status(400).json({ success: false, error: "printerName is required" });
  if (!html)
    return res.status(400).json({ success: false, error: "html is required" });

  const ts = new Date().toISOString();
  console.log(`\n╔══ PRINT JOB ${ts} ══╗`);
  console.log(`  Printer : ${printerName}`);

  const pdfFile = path.join(os.tmpdir(), `mrh_${Date.now()}.pdf`);

  try {
    if (printerName.toUpperCase().startsWith("TCP:")) {
      // TCP ESC/POS path — not supported with HTML, log a warning
      return res.status(400).json({ success: false, error: "TCP printers not supported with HTML receipts" });
    }

    // 1. Render HTML → PDF via Puppeteer
    const b    = await getBrowser();
    const page = await b.newPage();
    try {
      await page.setContent(html, { waitUntil: "networkidle0", timeout: 10000 });
      const pdf = await page.pdf({
        width:           "80mm",
        printBackground: true,
        margin:          { top: "0", bottom: "0", left: "0", right: "0" },
      });
      fs.writeFileSync(pdfFile, pdf);
    } finally {
      await page.close();
    }

    // 2. Print PDF via SumatraPDF
    await printPdf(pdfFile, printerName);

    console.log(`  Result  : SUCCESS`);
    console.log(`╚${"═".repeat(50)}╝`);
    res.json({ success: true });
  } catch (err) {
    console.error(`  Result  : FAILED — ${err.message}`);
    console.log(`╚${"═".repeat(50)}╝`);
    res.status(500).json({ success: false, error: err.message });
  } finally {
    try { fs.unlinkSync(pdfFile); } catch {}
  }
});

// ── SumatraPDF print helper ───────────────────────────────────────────────────

function printPdf(pdfPath, printerName) {
  return new Promise((resolve, reject) => {
    // -print-to targets a specific Windows printer by name
    // -print-settings "fit" scales content to fit the paper
    // -silent suppresses any UI dialogs
    execFile(
      SUMATRA,
      ["-print-to", printerName, "-print-settings", "fit,portrait", "-silent", pdfPath],
      { timeout: 20000 },
      (err) => {
        if (err) reject(new Error(`SumatraPDF: ${err.message}`));
        else resolve();
      }
    );
  });
}

// ── Start ─────────────────────────────────────────────────────────────────────

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Mr. Healthy Print Server v${VERSION}`);
  console.log(`Listening on http://0.0.0.0:${PORT}`);
  console.log(`SumatraPDF path: ${SUMATRA}`);
  console.log("Press Ctrl+C to stop.\n");
});

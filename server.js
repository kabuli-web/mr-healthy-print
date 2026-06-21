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


// Handle CORS + Chrome Private Network Access preflight
app.options("*", (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Allow-Private-Network", "true");
  res.sendStatus(204);
});
app.use(cors());
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Private-Network", "true");
  next();
});
app.use(express.json({ limit: "2mb" }));

// ── Puppeteer browser (singleton, reused across requests) ─────────────────────

let browser = null;

// Common Chrome/Edge install paths on Windows — first one found is used.
// Override by setting CHROME_PATH env var.
const CHROME_CANDIDATES = [
  process.env.CHROME_PATH,
  "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
  "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
  "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
].filter(Boolean);

function findChrome() {
  for (const p of CHROME_CANDIDATES) {
    try { if (fs.existsSync(p)) return p; } catch {}
  }
  return null; // fall back to puppeteer's bundled browser
}

async function getBrowser() {
  if (browser) return browser;
  const puppeteer  = require("puppeteer");
  const execPath   = findChrome();
  const launchOpts = {
    headless: "new",
    args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-gpu"],
  };
  if (execPath) {
    launchOpts.executablePath = execPath;
    console.log(`Using browser: ${execPath}`);
  } else {
    console.log("Using Puppeteer bundled browser");
  }
  browser = await puppeteer.launch(launchOpts);
  browser.on("disconnected", () => {
    console.log("Browser disconnected — will relaunch on next print");
    browser = null;
  });
  console.log("Puppeteer browser ready");
  return browser;
}

// Pre-warm browser on startup
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

  const t0 = Date.now();
  console.log(`\n╔══ PRINT JOB ${new Date(t0).toISOString()} ══╗`);
  console.log(`  Printer : ${printerName}`);

  const pdfFile = path.join(os.tmpdir(), `mrh_${t0}.png`);

  try {
    if (printerName.toUpperCase().startsWith("TCP:")) {
      return res.status(400).json({ success: false, error: "TCP printers not supported with HTML receipts" });
    }

    // 1. Render HTML → PNG via Puppeteer
    const b    = await getBrowser();
    const page = await b.newPage();
    const t1 = Date.now();
    console.log(`  newPage : ${t1 - t0}ms`);
    try {
      await page.setViewport({ width: 302, height: 800, deviceScaleFactor: 3 });
      await page.setContent(html, { waitUntil: "domcontentloaded", timeout: 10000 });
      const t2 = Date.now();
      console.log(`  setContent : ${t2 - t1}ms`);
      const png = await page.screenshot({ fullPage: true, type: "png" });
      const t3 = Date.now();
      console.log(`  screenshot : ${t3 - t2}ms  (${png.length} bytes)`);
      fs.writeFileSync(pdfFile, png);
    } finally {
      await page.close();
    }

    // 2. Print PNG via GDI+ PowerShell
    const t4 = Date.now();
    await printPng(pdfFile, printerName);
    const t5 = Date.now();
    console.log(`  powershell : ${t5 - t4}ms`);

    console.log(`  TOTAL : ${t5 - t0}ms  ✓`);
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

// ── Persistent PowerShell print worker ───────────────────────────────────────
// System.Drawing is loaded ONCE at startup. Each print command is sent via
// stdin as a JSON line; the worker responds with {"ok":true} or {"ok":false,...}.
// This eliminates the ~1.5s PowerShell cold-start per job.

const PS_WORKER_SCRIPT = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

while ($true) {
    $line = [Console]::ReadLine()
    if ($null -eq $line) { break }
    $line = $line.Trim()
    if (-not $line) { continue }
    try {
        $cmd = $line | ConvertFrom-Json
        $img = [System.Drawing.Image]::FromFile($cmd.pngPath)
        $doc = New-Object System.Drawing.Printing.PrintDocument
        $doc.PrinterSettings.PrinterName = $cmd.printerName
        if (-not $doc.PrinterSettings.IsValid) {
            $img.Dispose(); $doc.Dispose()
            throw "Printer not found: $($cmd.printerName)"
        }
        $doc.DefaultPageSettings.PaperSize = New-Object System.Drawing.Printing.PaperSize('Receipt', 315, 2000)
        $doc.DefaultPageSettings.Margins   = New-Object System.Drawing.Printing.Margins(0, 0, 0, 0)
        $doc.DefaultPageSettings.Landscape = $false
        $printImg = $img
        $doc.Add_PrintPage({
            param($s, $e)
            $pw    = [float]$e.Graphics.VisibleClipBounds.Width
            $scale = $pw / [float]$printImg.Width
            $ph    = [float]$printImg.Height * $scale
            $e.Graphics.DrawImage($printImg, [System.Drawing.RectangleF]::new(0, 0, $pw, $ph))
            $e.HasMorePages = $false
        })
        $doc.Print()
        $doc.Dispose()
        $img.Dispose()
        [Console]::WriteLine('{"ok":true}')
    } catch {
        $msg = $_.Exception.Message -replace '\\', '\\\\' -replace '"', '\\"'
        [Console]::WriteLine('{"ok":false,"error":"' + $msg + '"}')
    }
    [Console]::Out.Flush()
}
`;

const { spawn } = require("child_process");

class PsWorker {
  constructor() {
    this.proc    = null;
    this.queue   = [];   // { resolve, reject }[]
    this.outBuf  = "";
    this.scriptPath = path.join(os.tmpdir(), "mrh_psworker.ps1");
  }

  start() {
    const bom = Buffer.from([0xef, 0xbb, 0xbf]);
    fs.writeFileSync(this.scriptPath, Buffer.concat([bom, Buffer.from(PS_WORKER_SCRIPT, "utf8")]));

    this.proc = spawn(
      "powershell",
      ["-NoProfile", "-STA", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", this.scriptPath],
      { stdio: ["pipe", "pipe", "pipe"] }
    );

    this.proc.stdout.setEncoding("utf8");
    this.proc.stdout.on("data", (chunk) => {
      this.outBuf += chunk;
      const lines = this.outBuf.split("\n");
      this.outBuf = lines.pop();
      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed) continue;
        const pending = this.queue.shift();
        if (!pending) { console.warn("[PS worker] unexpected output:", trimmed); continue; }
        try {
          const result = JSON.parse(trimmed);
          if (result.ok) pending.resolve();
          else           pending.reject(new Error(result.error || "print failed"));
        } catch {
          pending.reject(new Error("PS worker bad response: " + trimmed));
        }
      }
    });

    this.proc.stderr.setEncoding("utf8");
    this.proc.stderr.on("data", (d) => console.error("[PS worker stderr]", d.trim()));

    this.proc.on("error", (err) => {
      console.error("[PS worker] spawn error:", err.message);
      this.proc = null;
      const pending = [...this.queue];
      this.queue = [];
      for (const p of pending) p.reject(err);
    });

    this.proc.on("exit", (code) => {
      console.warn(`[PS worker] exited (code ${code}) — will restart on next job`);
      this.proc = null;
      const pending = [...this.queue];
      this.queue = [];
      for (const p of pending) p.reject(new Error("PS worker exited unexpectedly"));
    });

    console.log("[PS worker] started — System.Drawing loading...");
  }

  print(pngPath, printerName) {
    if (!this.proc) {
      console.log("[PS worker] restarting...");
      this.start();
    }
    return new Promise((resolve, reject) => {
      this.queue.push({ resolve, reject });
      const cmd = JSON.stringify({ pngPath, printerName });
      this.proc.stdin.write(cmd + "\n", "utf8");
    });
  }
}

const psWorker = new PsWorker();

function printPng(pngPath, printerName) {
  return psWorker.print(pngPath, printerName);
}

// ── Start ─────────────────────────────────────────────────────────────────────

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Mr. Healthy Print Server v${VERSION}`);
  console.log(`Listening on http://0.0.0.0:${PORT}`);
  psWorker.start(); // pre-warm: loads System.Drawing once

  console.log("Press Ctrl+C to stop.\n");
});

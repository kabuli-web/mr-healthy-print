const express = require("express");
const cors = require("cors");
const net = require("net");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execSync, execFileSync } = require("child_process");

const app = express();
const PORT = process.env.PORT || 3001;
const VERSION = "1.0.0";

app.use(cors());
app.use(express.json());

// ── Health ──────────────────────────────────────────────────────────────────

app.get("/health", (_req, res) => {
  res.json({ ok: true, version: VERSION });
});

// ── List printers ───────────────────────────────────────────────────────────

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

// ── Print order ─────────────────────────────────────────────────────────────

app.post("/print", async (req, res) => {
  const { printerName, order } = req.body || {};

  if (!printerName)
    return res.status(400).json({ success: false, error: "printerName is required" });
  if (!order)
    return res.status(400).json({ success: false, error: "order is required" });

  // ── Log the job ──
  const ts = new Date().toISOString();
  console.log(`\n╔══ PRINT JOB ${ts} ══╗`);
  console.log(`  Printer  : ${printerName}`);
  console.log(`  Order    : ${order.orderNumber}`);
  console.log(`  Customer : ${order.customerName} / ${order.customerPhone}`);
  console.log(`  Branch   : ${order.branchName || "-"}`);
  console.log(`  Type     : ${order.dineType || "-"} / ${order.fulfillmentType || "-"}`);
  console.log(`  Items    :`);
  for (const item of order.items || []) {
    console.log(`    ${item.quantity}x ${item.name}`);
    for (const a of item.selectedAddons || []) {
      const qty = a.quantity ?? 1;
      console.log(`       + ${a.option?.name ?? ""}${qty > 1 ? ` x${qty}` : ""}`);
    }
  }
  if (order.note) console.log(`  Note     : ${order.note}`);

  try {
    const lines = buildReceiptLines(order);

    if (printerName.toUpperCase().startsWith("TCP:")) {
      const parts = printerName.split(":");
      await sendTcp(parts[1], parseInt(parts[2]) || 9100, lines);
    } else {
      await sendWindowsPrinter(printerName, lines);
    }

    console.log(`  Result   : SUCCESS`);
    console.log(`╚${"═".repeat(50)}╝`);
    res.json({ success: true });
  } catch (err) {
    console.error(`  Result   : FAILED — ${err.message}`);
    console.log(`╚${"═".repeat(50)}╝`);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ── Receipt lines builder ────────────────────────────────────────────────────

function buildReceiptLines(order) {
  const now = new Date().toLocaleString("ar-SA", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: true,
    timeZone: "Asia/Riyadh",
  });

  const lines = [];
  const add = (text, opts = {}) => lines.push({ text, ...opts });

  add("مستر صحي", { bold: true, center: true, large: true });
  add("---");
  add(`طلب رقم: ${String(order.orderNumber || "").slice(-6)}`);
  add(`الوقت: ${now}`);
  add("---");
  add(`العميل: ${order.customerName || ""}`);
  if (order.customerPhone) add(`الجوال: ${order.customerPhone}`);
  if (order.branchName)    add(`الفرع: ${order.branchName}`);
  if (order.dineType)      add(order.dineType === "dine_in" ? "نوع الطلب: محلي" : "نوع الطلب: سفري");
  if (order.fulfillmentType === "delivery") add("طريقة التسليم: توصيل");
  add("---");

  for (const item of order.items || []) {
    add(`${item.quantity}x  ${item.name}`, { bold: true });
    for (const addon of item.selectedAddons || []) {
      const qty = addon.quantity ?? 1;
      add(`    - ${addon.option?.name ?? ""}${qty > 1 ? ` x${qty}` : ""}`);
    }
  }

  add("---");

  if (order.note) {
    add(`ملاحظة: ${order.note}`, { bold: true });
    add("---");
  }

  return lines;
}

// ── Windows named-printer via GDI+ (PowerShell) ─────────────────────────────
// Windows handles Arabic shaping + RTL — no encoding tricks needed.

function sendWindowsPrinter(printerName, lines) {
  const id = Date.now();
  const jsonFile = path.join(os.tmpdir(), `mrh_${id}.json`);
  const psFile   = path.join(os.tmpdir(), `mrh_${id}.ps1`);

  // Write JSON with UTF-8 BOM so PowerShell reads it correctly
  const bom = Buffer.from([0xef, 0xbb, 0xbf]);
  fs.writeFileSync(jsonFile, Buffer.concat([bom, Buffer.from(JSON.stringify(lines), "utf8")]));

  const safeJson    = jsonFile.replace(/\\/g, "\\\\");
  const safePrinter = printerName.replace(/'/g, "''");

  const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$lines = [System.IO.File]::ReadAllText('${safeJson}', [System.Text.Encoding]::UTF8) | ConvertFrom-Json

$doc = New-Object System.Drawing.Printing.PrintDocument
$doc.PrinterSettings.PrinterName = '${safePrinter}'

if (-not $doc.PrinterSettings.IsValid) {
    throw "Printer not found: '${safePrinter}'"
}

# 80 mm wide = ~227 points (at 72 dpi) but PrintDocument uses hundredths of an inch
# 80 mm = 3.15 inch = 315 units; height 2000 = 20 inch (thermal scrolls)
$doc.DefaultPageSettings.PaperSize    = New-Object System.Drawing.Printing.PaperSize('Receipt', 315, 2000)
$doc.DefaultPageSettings.Margins      = New-Object System.Drawing.Printing.Margins(15, 15, 10, 10)
$doc.DefaultPageSettings.Landscape    = $false

$doc.Add_PrintPage({
    param($s, $e)

    $normalFont = New-Object System.Drawing.Font('Tahoma', 9)
    $boldFont   = New-Object System.Drawing.Font('Tahoma', 9,  [System.Drawing.FontStyle]::Bold)
    $titleFont  = New-Object System.Drawing.Font('Tahoma', 13, [System.Drawing.FontStyle]::Bold)

    $rtl = New-Object System.Drawing.StringFormat
    $rtl.FormatFlags   = [System.Drawing.StringFormatFlags]::DirectionRightToLeft
    $rtl.Alignment     = [System.Drawing.StringAlignment]::Near
    $rtl.LineAlignment = [System.Drawing.StringAlignment]::Center

    $center = New-Object System.Drawing.StringFormat
    $center.Alignment     = [System.Drawing.StringAlignment]::Center
    $center.LineAlignment = [System.Drawing.StringAlignment]::Center

    $pageW = [float]$e.MarginBounds.Width
    $x     = [float]$e.MarginBounds.Left
    $y     = [float]$e.MarginBounds.Top

    foreach ($line in $lines) {
        if ($line.text -eq '---') {
            $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black, 0.5)
            $e.Graphics.DrawLine($pen, $x, ($y + 5), ($x + $pageW), ($y + 5))
            $y += 14
            continue
        }

        $font = if ($line.large) { $titleFont } elseif ($line.bold) { $boldFont } else { $normalFont }
        $fmt  = if ($line.center) { $center } else { $rtl }
        $rect = New-Object System.Drawing.RectangleF($x, $y, $pageW, 28)
        $e.Graphics.DrawString($line.text, $font, [System.Drawing.Brushes]::Black, $rect, $fmt)
        $y += if ($line.large) { 26 } else { 20 }
    }
})

$doc.Print()
$doc.Dispose()
`;

  const bom2 = Buffer.from([0xef, 0xbb, 0xbf]);
  fs.writeFileSync(psFile, Buffer.concat([bom2, Buffer.from(script, "utf8")]));

  try {
    execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", psFile], {
      timeout: 15000,
    });
  } finally {
    try { fs.unlinkSync(jsonFile); } catch {}
    try { fs.unlinkSync(psFile);   } catch {}
  }
}

// ── TCP raw ESC/POS sender ────────────────────────────────────────────────────
// Sends UTF-8 bytes — requires the printer to have UTF-8 mode enabled.

function sendTcp(host, port, lines) {
  const ESC = 0x1b;
  const GS  = 0x1d;
  const parts = [Buffer.from([ESC, 0x40])]; // initialize

  for (const line of lines) {
    if (line.text === "---") {
      parts.push(Buffer.from("--------------------------------\n", "ascii"));
      continue;
    }
    if (line.bold)   parts.push(Buffer.from([ESC, 0x45, 0x01]));
    if (line.center) parts.push(Buffer.from([ESC, 0x61, 0x01]));
    parts.push(Buffer.from(line.text + "\n", "utf8"));
    if (line.center) parts.push(Buffer.from([ESC, 0x61, 0x00]));
    if (line.bold)   parts.push(Buffer.from([ESC, 0x45, 0x00]));
  }

  parts.push(Buffer.from([GS, 0x56, 0x41, 0x10])); // full cut

  const buffer = Buffer.concat(parts);

  return new Promise((resolve, reject) => {
    const socket = net.createConnection({ host, port });
    const timer = setTimeout(() => {
      socket.destroy();
      reject(new Error(`TCP timeout connecting to ${host}:${port}`));
    }, 5000);
    socket.on("connect", () => {
      socket.write(buffer, () => {
        clearTimeout(timer);
        socket.end();
        resolve();
      });
    });
    socket.on("error", (err) => {
      clearTimeout(timer);
      reject(err);
    });
  });
}

// ── Start ───────────────────────────────────────────────────────────────────

app.listen(PORT, "127.0.0.1", () => {
  console.log(`Mr. Healthy Print Server v${VERSION}`);
  console.log(`Listening on http://127.0.0.1:${PORT}`);
  console.log("Press Ctrl+C to stop.\n");
});

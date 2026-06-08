const express = require("express");
const cors = require("cors");
const net = require("net");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execSync, execFile } = require("child_process");

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

// ── Receipt builders ─────────────────────────────────────────────────────────

function buildReceiptLines(order) {
  if (order.receiptType === "kitchen") return buildKitchenReceipt(order);
  return buildPrimaryReceipt(order);
}

// Full receipt — primary printer
function buildPrimaryReceipt(order) {
  const lines = [];
  const add    = (text, opts = {}) => lines.push({ text, ...opts });
  const row    = (qty, name, opts = {}) => lines.push({ tableRow: true, qty, name, ...opts });
  const dine   = order.dineType === "dine_in" ? "محلي" : "سفري";
  const orderNum = String(order.orderNumber || "").slice(-5);

  add("مستر صحي", { bold: true, center: true, large: true });
  if (order.branchName) add(order.branchName, { center: true });
  add("---");
  add(`رقم الطلب: ${orderNum}`, { bold: true });
  add(`التاريخ: ${order.printedAt || ""}`);
  add(`نوع الطلب: ${dine}${order.fulfillmentType === "delivery" ? " — توصيل" : ""}`, { bold: true });
  add("---");
  add(`العميل: ${order.customerName || ""}`);
  if (order.customerPhone) add(`الجوال: ${order.customerPhone}`);
  add("---");

  // Table header
  row("الكمية", "الوجبة", { bold: true });
  add("---");

  for (const item of order.items || []) {
    row(String(item.quantity), item.name, { bold: true });
    if (item.nameEn) lines.push({ text: item.nameEn, small: true, ltr: true, indent: true });

    for (const addon of item.selectedAddons || []) {
      const qty = addon.quantity ?? 1;
      const addonName = addon.option?.name ?? "";
      const addonNameEn = addon.option?.nameEn ?? "";
      const freeQty = addon.freeQuantity ?? 0;
      const paidQty = addon.paidQuantity ?? qty;
      let label = `  - ${addonName}${qty > 1 ? ` ×${qty}` : ""}`;
      if (freeQty > 0 && paidQty === 0) label += " (مجاني)";
      else if (freeQty > 0) label += ` (مجاني ${freeQty})`;
      add(label, { small: true });
      if (addonNameEn) add(`    ${addonNameEn}`, { small: true, ltr: true });
    }
  }

  add("---");
  if (order.note) {
    add(`ملاحظة: ${order.note}`, { bold: true });
    add("---");
  }

  return lines;
}

// Compact receipt — kitchen section printer
function buildKitchenReceipt(order) {
  const lines = [];
  const add  = (text, opts = {}) => lines.push({ text, ...opts });
  const row  = (qty, name, opts = {}) => lines.push({ tableRow: true, qty, name, ...opts });
  const dine = order.dineType === "dine_in" ? "محلي 🍽️" : "سفري 🥡";
  const orderNum = String(order.orderNumber || "").slice(-5);

  add(orderNum, { bold: true, center: true, large: true });
  add(dine, { bold: true, center: true });
  add(order.printedAt || "", { center: true, small: true });
  add(`${order.customerName || ""}`, { center: true });
  add("---");

  row("الكمية", "الوجبة", { bold: true });
  add("---");

  for (const item of order.items || []) {
    row(String(item.quantity), item.name, { bold: true });
    if (item.nameEn) add(item.nameEn, { small: true, ltr: true });

    for (const addon of item.selectedAddons || []) {
      const qty = addon.quantity ?? 1;
      const addonName = addon.option?.name ?? "";
      add(`  - ${addonName}${qty > 1 ? ` ×${qty}` : ""}`, { small: true });
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
    $smallFont  = New-Object System.Drawing.Font('Tahoma', 8)
    $boldFont   = New-Object System.Drawing.Font('Tahoma', 9,  [System.Drawing.FontStyle]::Bold)
    $titleFont  = New-Object System.Drawing.Font('Tahoma', 14, [System.Drawing.FontStyle]::Bold)

    $rtl = New-Object System.Drawing.StringFormat
    $rtl.FormatFlags   = [System.Drawing.StringFormatFlags]::DirectionRightToLeft
    $rtl.Alignment     = [System.Drawing.StringAlignment]::Near
    $rtl.LineAlignment = [System.Drawing.StringAlignment]::Center

    $center = New-Object System.Drawing.StringFormat
    $center.Alignment     = [System.Drawing.StringAlignment]::Center
    $center.LineAlignment = [System.Drawing.StringAlignment]::Center

    $ltr = New-Object System.Drawing.StringFormat
    $ltr.Alignment     = [System.Drawing.StringAlignment]::Near
    $ltr.LineAlignment = [System.Drawing.StringAlignment]::Center

    $pageW = [float]$e.MarginBounds.Width
    $x     = [float]$e.MarginBounds.Left
    $y     = [float]$e.MarginBounds.Top

    $qtyColW = [float]36

    foreach ($line in $lines) {
        if ($line.text -eq '---') {
            $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black, 0.5)
            $e.Graphics.DrawLine($pen, $x, ($y + 4), ($x + $pageW), ($y + 4))
            $y += 13
            continue
        }

        if ($line.tableRow -eq $true) {
            $font2   = if ($line.bold) { $boldFont } else { $normalFont }
            $lineH   = if ($line.bold) { 22 } else { 20 }
            $qtyRect  = [System.Drawing.RectangleF]::new($x, $y, $qtyColW, $lineH)
            $nameRect = [System.Drawing.RectangleF]::new(($x + $qtyColW + 4), $y, ($pageW - $qtyColW - 4), $lineH)
            $e.Graphics.DrawString($line.qty,  $font2, [System.Drawing.Brushes]::Black, $qtyRect,  $center)
            $e.Graphics.DrawString($line.name, $font2, [System.Drawing.Brushes]::Black, $nameRect, $rtl)
            $y += $lineH
            continue
        }

        $font = if ($line.large) { $titleFont } elseif ($line.bold) { $boldFont } elseif ($line.small) { $smallFont } else { $normalFont }
        $fmt  = if ($line.center) { $center } elseif ($line.ltr) { $ltr } else { $rtl }
        $rect = New-Object System.Drawing.RectangleF($x, $y, $pageW, 28)
        $e.Graphics.DrawString($line.text, $font, [System.Drawing.Brushes]::Black, $rect, $fmt)
        $y += if ($line.large) { 28 } elseif ($line.small) { 16 } else { 20 }
    }
})

$doc.Print()
$doc.Dispose()
`;

  const bom2 = Buffer.from([0xef, 0xbb, 0xbf]);
  fs.writeFileSync(psFile, Buffer.concat([bom2, Buffer.from(script, "utf8")]));

  // -STA: Single Threaded Apartment — required for GDI/System.Drawing in headless PS
  // execFile (async) so multiple print jobs run in parallel, not sequentially
  return new Promise((resolve, reject) => {
    execFile(
      "powershell",
      ["-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", psFile],
      { timeout: 20000 },
      (err) => {
        try { fs.unlinkSync(jsonFile); } catch {}
        try { fs.unlinkSync(psFile);   } catch {}
        if (err) reject(err);
        else resolve();
      }
    );
  });
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

    if (line.tableRow) {
      if (line.bold) parts.push(Buffer.from([ESC, 0x45, 0x01]));
      // qty left-aligned (fixed 4 chars), name right side
      const qty  = String(line.qty  || "").padEnd(5);
      const name = String(line.name || "");
      parts.push(Buffer.from(`${qty} ${name}\n`, "utf8"));
      if (line.bold) parts.push(Buffer.from([ESC, 0x45, 0x00]));
      continue;
    }

    if (line.bold)   parts.push(Buffer.from([ESC, 0x45, 0x01]));
    if (line.large)  parts.push(Buffer.from([GS,  0x21, 0x11])); // double width+height
    if (line.center) parts.push(Buffer.from([ESC, 0x61, 0x01]));
    else if (line.ltr) parts.push(Buffer.from([ESC, 0x61, 0x00]));
    parts.push(Buffer.from(line.text + "\n", "utf8"));
    if (line.large)  parts.push(Buffer.from([GS,  0x21, 0x00])); // reset size
    if (line.center || line.ltr) parts.push(Buffer.from([ESC, 0x61, 0x02]));
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

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Mr. Healthy Print Server v${VERSION}`);
  console.log(`Listening on http://0.0.0.0:${PORT} (all network interfaces)`);
  console.log("Press Ctrl+C to stop.\n");
});

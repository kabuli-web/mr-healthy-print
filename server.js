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

  try {
    const buffer = buildReceiptBuffer(order);

    if (printerName.toUpperCase().startsWith("TCP:")) {
      const parts = printerName.split(":");
      await sendTcp(parts[1], parseInt(parts[2]) || 9100, buffer);
    } else {
      await sendWindowsPrinter(printerName, buffer);
    }

    res.json({ success: true });
  } catch (err) {
    console.error(`Print error [${printerName}]:`, err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ── Arabic helper ────────────────────────────────────────────────────────────
// Unicode Arabic is in abstract/base form; PC864 expects visual presentation
// forms. Reshape converts each character to the right isolated/initial/medial/
// final form based on context, then iconv-lite encodes the result to CP864 bytes.

const iconv = require("iconv-lite");
const { reshape } = require("arabic-reshaper");

function ar(text) {
  return iconv.encode(reshape(text), "CP864");
}

// ── Receipt builder (returns raw ESC/POS Buffer) ─────────────────────────────

function buildReceiptBuffer(order) {
  const { ThermalPrinter, PrinterTypes, CharacterSet } = require("node-thermal-printer");

  // Dummy TCP interface — we only call getBuffer(), never execute()
  const printer = new ThermalPrinter({
    type: PrinterTypes.EPSON,
    interface: "tcp://127.0.0.1:19999",
    characterSet: CharacterSet.PC864_ARABIC,
    removeSpecialCharacters: false,
    lineCharacter: "-",
    width: 42,
  });

  const now = new Date().toLocaleString("ar-SA", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: true,
    timeZone: "Asia/Riyadh",
  });

  printer.alignCenter();
  printer.bold(true);
  printer.append(ar("مستر صحي")); printer.newLine();
  printer.bold(false);
  printer.drawLine();

  printer.alignRight();
  printer.append(ar(`طلب رقم: ${String(order.orderNumber || "").slice(-6)}`)); printer.newLine();
  printer.append(ar(`الوقت: ${now}`)); printer.newLine();
  printer.drawLine();

  printer.append(ar(`العميل: ${order.customerName || ""}`)); printer.newLine();
  if (order.customerPhone) { printer.append(ar(`الجوال: ${order.customerPhone}`)); printer.newLine(); }
  if (order.branchName)    { printer.append(ar(`الفرع: ${order.branchName}`));      printer.newLine(); }
  if (order.dineType)      { printer.append(ar(order.dineType === "dine_in" ? "محلي" : "سفري")); printer.newLine(); }
  if (order.fulfillmentType === "delivery") { printer.append(ar("توصيل")); printer.newLine(); }

  printer.drawLine();

  for (const item of order.items || []) {
    printer.bold(true);
    printer.append(ar(`${item.quantity}x ${item.name}`)); printer.newLine();
    printer.bold(false);
    if (item.selectedAddons?.length) {
      for (const addon of item.selectedAddons) {
        const qty = addon.quantity ?? 1;
        printer.append(ar(`  - ${addon.option?.name ?? ""}${qty > 1 ? ` x${qty}` : ""}`)); printer.newLine();
      }
    }
  }

  printer.drawLine();

  if (order.note) {
    printer.bold(true);
    printer.append(ar(`ملاحظة: ${order.note}`)); printer.newLine();
    printer.bold(false);
    printer.drawLine();
  }

  printer.cut();
  return printer.getBuffer();
}

// ── TCP sender (raw socket, no extra packages) ───────────────────────────────

function sendTcp(host, port, buffer) {
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

// ── Windows named-printer sender (PowerShell P/Invoke, no native packages) ───

function sendWindowsPrinter(printerName, buffer) {
  const id = Date.now();
  const binFile = path.join(os.tmpdir(), `mrh_${id}.bin`);
  const psFile = path.join(os.tmpdir(), `mrh_${id}.ps1`);

  fs.writeFileSync(binFile, buffer);

  // Single-quote escape for PowerShell string literals
  const safeBin = binFile.replace(/\\/g, "\\\\");
  const safeName = printerName.replace(/'/g, "''");

  const script = `
$ErrorActionPreference = 'Stop'
$bytes = [System.IO.File]::ReadAllBytes('${safeBin}')
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class RawPrint {
    [DllImport("winspool.drv", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern bool OpenPrinter(string n, out IntPtr h, IntPtr d);
    [DllImport("winspool.drv", SetLastError=true)]
    public static extern bool ClosePrinter(IntPtr h);
    [DllImport("winspool.drv", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern int StartDocPrinter(IntPtr h, int lvl, ref DOCINFO info);
    [DllImport("winspool.drv", SetLastError=true)]
    public static extern bool EndDocPrinter(IntPtr h);
    [DllImport("winspool.drv", SetLastError=true)]
    public static extern bool StartPagePrinter(IntPtr h);
    [DllImport("winspool.drv", SetLastError=true)]
    public static extern bool EndPagePrinter(IntPtr h);
    [DllImport("winspool.drv", SetLastError=true)]
    public static extern bool WritePrinter(IntPtr h, byte[] buf, int len, out int written);
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
    public struct DOCINFO { public string pDocName; public string pOutputFile; public string pDataType; }
}
"@
$h = [IntPtr]::Zero
if (-not [RawPrint]::OpenPrinter('${safeName}', [ref]$h, [IntPtr]::Zero)) {
    throw "Cannot open printer '${safeName}' - check the name in Windows Settings > Devices"
}
try {
    $doc = New-Object RawPrint+DOCINFO
    $doc.pDocName  = "Mr Healthy Order"
    $doc.pDataType = "RAW"
    $job = [RawPrint]::StartDocPrinter($h, 1, [ref]$doc)
    if ($job -le 0) { throw "StartDocPrinter failed" }
    [RawPrint]::StartPagePrinter($h) | Out-Null
    $w = 0
    [RawPrint]::WritePrinter($h, $bytes, $bytes.Length, [ref]$w) | Out-Null
    [RawPrint]::EndPagePrinter($h) | Out-Null
    [RawPrint]::EndDocPrinter($h) | Out-Null
} finally {
    [RawPrint]::ClosePrinter($h) | Out-Null
}
`;

  // UTF-8 BOM so PowerShell 5.1 reads the file as UTF-8 (default is ANSI)
  const bom = Buffer.from([0xef, 0xbb, 0xbf]);
  fs.writeFileSync(psFile, Buffer.concat([bom, Buffer.from(script, "utf8")]));

  try {
    execFileSync("powershell", [
      "-NoProfile",
      "-ExecutionPolicy", "Bypass",
      "-File", psFile,
    ], { timeout: 10000 });
  } finally {
    try { fs.unlinkSync(binFile); } catch {}
    try { fs.unlinkSync(psFile); } catch {}
  }
}

// ── Start ───────────────────────────────────────────────────────────────────

app.listen(PORT, "127.0.0.1", () => {
  console.log(`Mr. Healthy Print Server v${VERSION}`);
  console.log(`Listening on http://127.0.0.1:${PORT}`);
  console.log("Press Ctrl+C to stop.");
});

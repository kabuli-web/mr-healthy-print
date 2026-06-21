"use strict";

require("dotenv").config({ path: require("path").join(__dirname, ".env") });

const express    = require("express");
const cors       = require("cors");
const fs         = require("fs");
const os         = require("os");
const path       = require("path");
const { spawn, execSync } = require("child_process");

// ── Firebase ──────────────────────────────────────────────────────────────────

const admin = require("firebase-admin");

let serviceAccount;
try {
  serviceAccount = process.env.FIREBASE_SERVICE_ACCOUNT
    ? JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT)
    : require("./serviceAccount.json");
} catch {
  console.error("ERROR: Firebase credentials not found.");
  console.error("  Place serviceAccount.json next to server.js, or set FIREBASE_SERVICE_ACCOUNT env var.");
  process.exit(1);
}

admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
const db = admin.firestore();

const BRANCH_ID = process.env.BRANCH_ID;
if (!BRANCH_ID) {
  console.error("ERROR: BRANCH_ID is not set. Add it to .env or set the environment variable.");
  process.exit(1);
}

// ── Express ───────────────────────────────────────────────────────────────────

const app     = express();
const PORT    = process.env.PORT || 3001;
const VERSION = "2.0.0";

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

// ── Puppeteer browser singleton ───────────────────────────────────────────────

let browser = null;

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
  return null;
}

async function getBrowser() {
  if (browser) return browser;
  const puppeteer  = require("puppeteer");
  const execPath   = findChrome();
  const launchOpts = { headless: "new", args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-gpu"] };
  if (execPath) { launchOpts.executablePath = execPath; console.log(`Browser: ${execPath}`); }
  browser = await puppeteer.launch(launchOpts);
  browser.on("disconnected", () => { browser = null; });
  console.log("Puppeteer ready");
  return browser;
}

// ── Persistent PowerShell print worker ───────────────────────────────────────

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

class PsWorker {
  constructor() {
    this.proc       = null;
    this.queue      = [];
    this.outBuf     = "";
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
        if (!pending) { console.warn("[PS] unexpected output:", trimmed); continue; }
        try {
          const result = JSON.parse(trimmed);
          if (result.ok) pending.resolve();
          else           pending.reject(new Error(result.error || "print failed"));
        } catch {
          pending.reject(new Error("PS bad response: " + trimmed));
        }
      }
    });

    this.proc.stderr.setEncoding("utf8");
    this.proc.stderr.on("data", (d) => console.error("[PS stderr]", d.trim()));

    this.proc.on("error", (err) => {
      console.error("[PS] spawn error:", err.message);
      this.proc = null;
      const pending = [...this.queue]; this.queue = [];
      for (const p of pending) p.reject(err);
    });

    this.proc.on("exit", (code) => {
      console.warn(`[PS] exited (${code}) — will restart on next job`);
      this.proc = null;
      const pending = [...this.queue]; this.queue = [];
      for (const p of pending) p.reject(new Error("PS worker exited"));
    });

    console.log("[PS] worker started");
  }

  print(pngPath, printerName) {
    if (!this.proc) this.start();
    return new Promise((resolve, reject) => {
      this.queue.push({ resolve, reject });
      this.proc.stdin.write(JSON.stringify({ pngPath, printerName }) + "\n", "utf8");
    });
  }
}

const psWorker = new PsWorker();

// ── HTML builder (mirrors printIntegration.ts exactly) ────────────────────────

function esc(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

const BASE_CSS = `
* { margin: 0; padding: 0; box-sizing: border-box; }
html, body { width: 100%; margin: 0 auto; }
body {
  font-family: Tahoma, Arial, sans-serif;
  font-size: 14px;
  color: #000;
  padding: 3mm 5mm 6mm;
}
.center { text-align: center; }
.bold   { font-weight: bold; }
.title  { font-size: 22px; font-weight: bold; text-align: center; margin-bottom: 2px; }
.order-box {
  border: 2px solid #000;
  font-size: 28px;
  font-weight: bold;
  text-align: center;
  padding: 3px 6px;
  margin: 4px 0;
}
.meta { text-align: center; font-size: 12px; margin: 2px 0; }
.info { margin: 2px 0; font-size: 13px; }
hr { border: none; border-top: 1px solid #000; margin: 4px 0; }
table.bordered {
  width: 100%;
  border-collapse: collapse;
  border: 1.5px solid #000;
  margin: 4px 0;
}
table.bordered th, table.bordered td {
  border: 1px solid #000;
  padding: 4px 5px;
  vertical-align: top;
}
table.bordered th { font-weight: bold; text-align: center; font-size: 13px; }
.col-qty   { width: 30px; text-align: center; white-space: nowrap; }
.col-total { width: 46px; text-align: left; direction: ltr; white-space: nowrap; }
.col-price { width: 40px; text-align: left; direction: ltr; white-space: nowrap; }
.col-name  { text-align: right; word-break: break-word; }
.meal-name { font-weight: bold; }
.addon-cell {
  font-size: 14px;
  padding-right: 10px !important;
  color: #000;
}
table.totals { width: 100%; border-collapse: collapse; margin-top: 3px; }
table.totals td { padding: 2px 3px; font-size: 13px; }
table.totals .total-label { text-align: right; }
table.totals .total-value { text-align: left; direction: ltr; width: 50px; font-weight: bold; }
table.totals .grand td { font-size: 15px; font-weight: bold; border-top: 2px solid #000; padding-top: 4px; }
.note { font-weight: bold; margin-top: 5px; font-size: 13px; }
`;

function buildPrimaryReceiptHtml(order) {
  const num  = String(order.orderNumber).slice(-5);
  const dine = order.dineType === "dine_in" ? "محلي" : "سفري";
  const pay  = order.paymentMethod === "subscription" ? "رصيد" : "نقدي";
  const isSubscription = order.paymentMethod === "subscription";

  let mealTotal = 0, addonCash = 0, itemRows = "";

  for (const item of order.items) {
    const linePrice  = item.unitPrice * item.quantity;
    mealTotal       += linePrice;
    const priceCell  = isSubscription ? "مشمول" : `${item.unitPrice} ر`;
    const totalCell  = isSubscription ? "مشمول" : `${linePrice} ر`;

    itemRows += `
      <tr>
        <td class="col-qty" style="text-align:center">${esc(item.quantity)}</td>
        <td class="col-name"><span class="meal-name">${esc(item.name)}</span></td>
        <td class="col-price">${priceCell}</td>
        <td class="col-total">${totalCell}</td>
      </tr>`;

    for (const addon of (item.selectedAddons ?? [])) {
      const qty        = addon.quantity ?? 1;
      const addonPrice = addon.totalPrice ?? 0;
      const freeQty    = addon.freeQuantity ?? 0;
      const paidQty    = addon.paidQuantity ?? qty;
      if (addonPrice > 0) addonCash += addonPrice * item.quantity;
      let addonTotal = "";
      if (addonPrice > 0)                    addonTotal = `+${addonPrice} ر`;
      else if (freeQty > 0 && paidQty === 0) addonTotal = "مجاني";
      else if (freeQty > 0)                  addonTotal = `مجاني ${freeQty}`;
      const nameAr   = esc(addon.option?.name);
      const nameEn   = addon.option?.nameEn ? esc(addon.option.nameEn) : "";
      const nameCell = nameEn
        ? `<div>${nameAr}</div><div style="font-size:13px;color:#444">${nameEn}</div>`
        : nameAr;
      const nb = "border:none";
      itemRows += `
      <tr>
        <td class="col-qty" style="text-align:center;${nb}">${qty}</td>
        <td colspan="2" class="addon-cell" style="${nb}">${nameCell}</td>
        <td class="col-total" style="${nb}">${addonTotal}</td>
      </tr>`;
    }
  }

  let totalSection = "";
  if (isSubscription) {
    totalSection = `
      <table class="totals">
        <tr><td class="total-label">الوجبات</td><td class="total-value">مشمول</td></tr>
        ${addonCash > 0 ? `<tr><td class="total-label">إضافات</td><td class="total-value">${addonCash} ر</td></tr>` : ""}
        <tr class="grand"><td class="total-label">المبلغ المطلوب</td><td class="total-value">${addonCash > 0 ? `${addonCash} ر` : "صفر"}</td></tr>
      </table>`;
  } else {
    const fullTotal = mealTotal + addonCash;
    totalSection = `
      <table class="totals">
        <tr class="grand"><td class="total-label">الإجمالي</td><td class="total-value">${fullTotal} ر</td></tr>
      </table>`;
  }

  const noteHtml = order.note
    ? `<div class="note" style="margin-top:6px;border:1px solid #000;padding:4px 6px">ملاحظة: ${esc(order.note)}</div>`
    : "";

  return `<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head><meta charset="UTF-8"><style>${BASE_CSS}</style></head>
<body>
  <div class="title">مستر صحي</div>
  ${order.branchName ? `<div class="center" style="font-size:13px">${esc(order.branchName)}</div>` : ""}
  <hr>
  <div class="order-box"># ${esc(num)}</div>
  <div class="meta">${esc(order.printedAt)}</div>
  <div class="meta bold">${esc(dine)} &nbsp;|&nbsp; ${esc(pay)}</div>
  <hr>
  <div class="info">العميل: <strong>${esc(order.customerName)}</strong></div>
  ${order.customerPhone ? `<div class="info">الجوال: ${esc(order.customerPhone)}</div>` : ""}
  <table class="bordered">
    <thead>
      <tr>
        <th class="col-qty">كمية</th>
        <th class="col-name">الصنف</th>
        <th class="col-price">السعر</th>
        <th class="col-total">الإجمالي</th>
      </tr>
    </thead>
    <tbody>${itemRows}</tbody>
  </table>
  ${totalSection}
  ${noteHtml}
</body>
</html>`;
}

function buildKitchenReceiptHtml(order) {
  const num    = String(order.orderNumber).slice(-5);
  const dineAr = order.dineType === "dine_in" ? "محلي" : "سفري";
  const dineEn = order.dineType === "dine_in" ? "DINE-IN" : "TAKE-AWAY";

  let itemRows = "";
  for (const item of order.items) {
    const arName = esc(item.name);
    const enName = item.nameEn ? esc(item.nameEn) : "";
    itemRows += `
      <tr>
        <td class="col-qty" style="text-align:center;font-size:16px;font-weight:bold">${esc(item.quantity)}</td>
        <td style="text-align:right;word-break:break-word">
          <span style="font-weight:bold;font-size:15px">${arName}</span>${enName ? `<br><span style="font-size:14px;color:#333">${enName}</span>` : ""}
        </td>
      </tr>`;

    for (const a of (item.selectedAddons ?? [])) {
      const qty    = a.quantity ?? 1;
      const nameAr = esc(a.option?.name || a.option?.nameEn);
      const nameEn = a.option?.name && a.option?.nameEn ? esc(a.option.nameEn) : "";
      const nameCell = nameEn
        ? `<div>${nameAr}</div><div style="font-size:13px;color:#444">${nameEn}</div>`
        : nameAr;
      const nb = "border:none";
      itemRows += `
      <tr>
        <td class="col-qty" style="text-align:center;${nb}">${qty}</td>
        <td class="addon-cell" style="text-align:right;${nb}">${nameCell}</td>
      </tr>`;
    }
  }

  const noteHtml = order.note
    ? `<div class="note" style="margin-top:6px;border:1px solid #000;padding:4px 6px;text-align:right">ملاحظة / NOTE: ${esc(order.note)}</div>`
    : "";

  return `<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head><meta charset="UTF-8"><style>${BASE_CSS}</style></head>
<body>
  <div class="order-box"># ${esc(num)}</div>
  <div class="meta bold">${dineAr} &nbsp;|&nbsp; ${dineEn}</div>
  <div class="meta">${esc(order.printedAt)}</div>
  <table class="bordered">
    <thead>
      <tr>
        <th class="col-qty">كمية</th>
        <th style="text-align:right">الصنف / ITEM</th>
      </tr>
    </thead>
    <tbody>${itemRows}</tbody>
  </table>
  ${noteHtml}
</body>
</html>`;
}

// ── Metadata cache ────────────────────────────────────────────────────────────

const CACHE_TTL = 24 * 60 * 60 * 1000; // 24 hours
let metaCache = null, metaCacheExpiry = 0;

async function getMeta() {
  if (metaCache && Date.now() < metaCacheExpiry) return metaCache;

  const [mealsSnap, assignmentsSnap, branchesSnap] = await Promise.all([
    db.collection("meals").get(),
    db.collection("settings").doc("printerAssignments").get(),
    db.collection("branches").get(),
  ]);

  const mealMeta = {};        // mealId → { nameEn, categoryId }
  const optionNameEn = {};    // optionId → nameEn

  for (const doc of mealsSnap.docs) {
    const m = doc.data();
    mealMeta[doc.id] = { nameEn: m.nameEn, categoryId: m.categoryId };
    for (const section of (m.addonSections ?? [])) {
      for (const opt of (section.options ?? [])) {
        if (opt.id && opt.nameEn) optionNameEn[opt.id] = opt.nameEn;
      }
    }
  }

  const branchNames = {};
  for (const doc of branchesSnap.docs) {
    branchNames[doc.id] = doc.data().name ?? "";
  }

  const assignmentsData = assignmentsSnap.exists ? assignmentsSnap.data() : {};
  const assignments = assignmentsData.branches ?? {};

  metaCache = { mealMeta, optionNameEn, branchNames, assignments };
  metaCacheExpiry = Date.now() + CACHE_TTL;
  return metaCache;
}

// ── Print pipeline ────────────────────────────────────────────────────────────

async function renderAndPrint(html, printerName) {
  const pngPath = path.join(os.tmpdir(), `mrh_${Date.now()}_${Math.random().toString(36).slice(2)}.png`);
  const t0 = Date.now();
  try {
    const b    = await getBrowser();
    const page = await b.newPage();
    try {
      await page.setViewport({ width: 302, height: 800, deviceScaleFactor: 3 });
      await page.setContent(html, { waitUntil: "domcontentloaded", timeout: 10000 });
      const png = await page.screenshot({ fullPage: true, type: "png" });
      fs.writeFileSync(pngPath, png);
    } finally {
      await page.close();
    }
    const t1 = Date.now();
    await psWorker.print(pngPath, printerName);
    console.log(`  printed to "${printerName}" in ${Date.now() - t0}ms (puppeteer: ${t1 - t0}ms, ps: ${Date.now() - t1}ms)`);
  } finally {
    try { fs.unlinkSync(pngPath); } catch {}
  }
}

function buildPrintedAt() {
  return new Date().toLocaleString("en-GB", {
    timeZone: "Asia/Riyadh",
    day: "2-digit", month: "2-digit", year: "numeric",
    hour: "2-digit", minute: "2-digit", hour12: false,
  });
}

async function printOrder(orderDoc) {
  const order = typeof orderDoc.data === "function"
    ? { id: orderDoc.id, ...orderDoc.data() }
    : orderDoc;

  const { mealMeta, optionNameEn, branchNames, assignments } = await getMeta();

  const branchId   = order.fulfillment?.branchId ?? order.branchId ?? "";
  const branchName = branchNames[branchId] ?? "";
  const branchCfg  = assignments[branchId];
  const printedAt  = buildPrintedAt();

  const buildData = (items) => ({
    orderNumber:    order.orderNumber,
    customerName:   order.customer?.name ?? "",
    customerPhone:  order.customer?.phone ?? "",
    branchName,
    printedAt,
    dineType:       order.dineType,
    paymentMethod:  order.payment?.method,
    note:           order.note,
    items: items.map(i => ({
      name:      i.name,
      nameEn:    mealMeta[i.mealId]?.nameEn,
      quantity:  i.quantity,
      unitPrice: i.finalPrice ?? i.price,
      selectedAddons: (i.selectedAddons ?? []).map(a => ({
        ...a,
        option: { ...a.option, nameEn: a.option?.id ? optionNameEn[a.option.id] : undefined },
      })),
    })),
  });

  const prints = [];

  // Primary receipt
  if (branchCfg?.primaryPrinter) {
    prints.push(renderAndPrint(buildPrimaryReceiptHtml(buildData(order.items)), branchCfg.primaryPrinter));
  }

  // Kitchen receipts — group items by category→printer
  const categoryAssignments = branchCfg?.categories ?? {};
  const byPrinter = new Map();
  for (const item of order.items) {
    const catId   = mealMeta[item.mealId]?.categoryId;
    const printer = catId ? categoryAssignments[catId] : undefined;
    if (!printer) continue;
    if (!byPrinter.has(printer)) byPrinter.set(printer, []);
    byPrinter.get(printer).push(item);
  }
  for (const [printer, items] of byPrinter.entries()) {
    prints.push(renderAndPrint(buildKitchenReceiptHtml(buildData(items)), printer));
  }

  if (prints.length === 0) console.log(`  no printers assigned for order ${order.orderNumber}`);
  await Promise.allSettled(prints);
}

// ── Printed orders tracking (survives restarts) ───────────────────────────────

const PRINTED_FILE = path.join(__dirname, "printed_orders.json");

function loadPrintedOrders() {
  try { return new Set(JSON.parse(fs.readFileSync(PRINTED_FILE, "utf8"))); } catch { return new Set(); }
}

function savePrintedOrders(set) {
  try { fs.writeFileSync(PRINTED_FILE, JSON.stringify([...set])); } catch {}
}

const printedOrders = loadPrintedOrders();

// ── Heartbeat ─────────────────────────────────────────────────────────────────

function startHeartbeat() {
  const write = () => {
    db.collection("serverStatus").doc(BRANCH_ID).set({
      lastSeen: admin.firestore.FieldValue.serverTimestamp(),
      version: VERSION,
      branchId: BRANCH_ID,
    }).catch(err => console.error("[heartbeat] failed:", err.message));
  };
  write();
  setInterval(write, 30_000);
}

// ── Firestore listeners ───────────────────────────────────────────────────────

const STARTUP_TIME = Date.now();
let lastSnapshotAt = Date.now();

function startWatchdog() {
  const STALE_MS = 10 * 60 * 1000; // 10 minutes
  setInterval(() => {
    const stale = Date.now() - lastSnapshotAt;
    if (stale > STALE_MS) {
      console.error(`[watchdog] no snapshot in ${Math.round(stale / 60000)}min — restarting`);
      process.exit(0);
    }
  }, 60_000);
  console.log("[watchdog] started (10min stale threshold)");
}

function startListeners() {
  console.log(`\nWatching orders for branch: ${BRANCH_ID}`);

  // Auto-print: orders that become accepted/preparing
  db.collection("orders")
    .where("fulfillment.branchId", "==", BRANCH_ID)
    .where("status", "in", ["accepted", "preparing"])
    .onSnapshot((snap) => {
      lastSnapshotAt = Date.now();
      for (const change of snap.docChanges()) {
        if (change.type !== "added") continue;
        const orderId = change.doc.id;
        if (printedOrders.has(orderId)) continue;
        printedOrders.add(orderId);
        savePrintedOrders(printedOrders);
        const data = change.doc.data();
        console.log(`\n[AUTO] printing order ${data.orderNumber}`);
        printOrder({ id: orderId, ...data }).catch(err =>
          console.error(`[AUTO] failed to print ${data.orderNumber}:`, err.message)
        );
      }
    }, (err) => console.error("[orders listener error]", err.message));

  // Manual reprint: printJobs written by admin — skip jobs older than 60s
  db.collection("printJobs")
    .where("branchId", "==", BRANCH_ID)
    .where("status", "==", "pending")
    .onSnapshot((snap) => {
      lastSnapshotAt = Date.now();
      for (const change of snap.docChanges()) {
        if (change.type !== "added") continue;
        const jobId = change.doc.id;
        const job   = change.doc.data();

        const jobTime = job.requestedAt?.toMillis?.() ?? 0;
        const ageMs   = Date.now() - jobTime;
        if (ageMs > 60_000) {
          console.log(`[JOB] skipping stale job (${Math.round(ageMs / 1000)}s old)`);
          db.collection("printJobs").doc(jobId).update({
            status: "failed",
            error: "Job was too old when server processed it",
            completedAt: admin.firestore.FieldValue.serverTimestamp(),
          }).catch(() => {});
          continue;
        }

        console.log(`\n[JOB] reprint request for order ${job.orderId}`);
        db.collection("orders").doc(job.orderId).get().then(async (orderSnap) => {
          if (!orderSnap.exists) throw new Error("order not found");
          await printOrder({ id: orderSnap.id, ...orderSnap.data() });
          await db.collection("printJobs").doc(jobId).update({
            status: "done",
            completedAt: admin.firestore.FieldValue.serverTimestamp(),
          });
          console.log(`[JOB] done — ${job.orderId}`);
        }).catch(async (err) => {
          console.error(`[JOB] failed — ${err.message}`);
          await db.collection("printJobs").doc(jobId).update({
            status: "failed",
            error: err.message,
            completedAt: admin.firestore.FieldValue.serverTimestamp(),
          }).catch(() => {});
        });
      }
    }, (err) => console.error("[printJobs listener error]", err.message));

  // Restart command listener
  db.collection("serverCommands").doc(BRANCH_ID).onSnapshot((snap) => {
    if (!snap.exists) return;
    const data = snap.data();
    if (data.command === "restart" && data.status === "pending") {
      console.log("[command] restart received — restarting...");
      snap.ref.update({ status: "acknowledged" }).finally(() => process.exit(0));
    }
  }, (err) => console.error("[commands listener error]", err.message));
}

// ── HTTP endpoints ────────────────────────────────────────────────────────────

app.get("/health", (_req, res) => {
  res.json({ ok: true, version: VERSION, branchId: BRANCH_ID });
});

app.get("/printers", (_req, res) => {
  try {
    const buf   = execSync(
      'powershell -NoProfile -Command "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Get-Printer | Select-Object -ExpandProperty Name"',
      { timeout: 5000, encoding: "buffer" }
    );
    const names = buf.toString("utf8").split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    res.json({ printers: names });
  } catch (err) {
    res.json({ printers: [], warning: err.message });
  }
});

app.get("/jobs", async (_req, res) => {
  try {
    const snap = await db.collection("printJobs")
      .where("branchId", "==", BRANCH_ID)
      .orderBy("requestedAt", "desc")
      .limit(30)
      .get();
    const jobs = snap.docs.map(d => {
      const j = d.data();
      return {
        id: d.id,
        orderId: j.orderId,
        status: j.status,
        error: j.error ?? null,
        requestedAt: j.requestedAt?.toMillis?.() ?? null,
        completedAt: j.completedAt?.toMillis?.() ?? null,
      };
    });
    res.json({ jobs });
  } catch (err) {
    res.json({ jobs: [], error: err.message });
  }
});

app.post("/restart", (_req, res) => {
  res.json({ ok: true });
  setTimeout(() => process.exit(0), 200);
});

app.get("/", (_req, res) => {
  res.setHeader("Content-Type", "text/html; charset=utf-8");
  res.send(DASHBOARD_HTML);
});

const DASHBOARD_HTML = `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>مستر صحي — خادم الطباعة</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Tahoma, Arial, sans-serif; background: #f0f4f0; color: #1a2e1a; direction: rtl; }
  header { background: #003628; color: #fff; padding: 16px 24px; display: flex; align-items: center; gap: 12px; }
  header h1 { font-size: 20px; }
  header .ver { font-size: 12px; opacity: 0.7; margin-right: auto; }
  .dot { width: 12px; height: 12px; border-radius: 50%; background: #4ade80; flex-shrink: 0; box-shadow: 0 0 6px #4ade80; }
  main { max-width: 900px; margin: 0 auto; padding: 20px 16px; display: grid; gap: 16px; }
  .card { background: #fff; border-radius: 8px; padding: 16px; box-shadow: 0 1px 4px rgba(0,0,0,.08); }
  .card h2 { font-size: 15px; color: #003628; margin-bottom: 12px; border-bottom: 1px solid #e5e7eb; padding-bottom: 8px; }
  .info-row { display: flex; gap: 24px; flex-wrap: wrap; }
  .info-item { font-size: 13px; }
  .info-item span { color: #6b7280; }
  .info-item strong { display: block; font-size: 16px; margin-top: 2px; }
  .printer-list { display: flex; flex-wrap: wrap; gap: 8px; }
  .printer-chip { background: #e8f5e9; color: #1b5e20; padding: 4px 12px; border-radius: 20px; font-size: 13px; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th { background: #f9fafb; padding: 8px 10px; text-align: right; font-weight: 600; color: #374151; border-bottom: 2px solid #e5e7eb; }
  td { padding: 8px 10px; border-bottom: 1px solid #f3f4f6; vertical-align: top; }
  tr:last-child td { border-bottom: none; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 12px; font-weight: bold; }
  .badge-done { background: #dcfce7; color: #166534; }
  .badge-pending { background: #fef9c3; color: #92400e; }
  .badge-failed { background: #fee2e2; color: #991b1b; }
  .error-text { color: #dc2626; font-size: 12px; margin-top: 3px; }
  .btn-restart { background: #dc2626; color: #fff; border: none; padding: 10px 22px; border-radius: 6px; font-size: 14px; font-family: inherit; cursor: pointer; }
  .btn-restart:hover { background: #b91c1c; }
  .btn-restart:disabled { opacity: 0.5; cursor: default; }
  #restart-msg { margin-top: 8px; font-size: 13px; color: #dc2626; }
  .refresh-note { font-size: 12px; color: #9ca3af; margin-top: 8px; }
  #uptime { font-size: 13px; color: #6b7280; }
</style>
</head>
<body>
<header>
  <div class="dot"></div>
  <h1>مستر صحي — خادم الطباعة</h1>
  <div class="ver" id="ver-branch">جاري التحميل...</div>
</header>
<main>
  <div class="card">
    <h2>حالة الخادم</h2>
    <div class="info-row" id="status-row">
      <div class="info-item"><span>الحالة</span><strong style="color:#16a34a">يعمل ✓</strong></div>
      <div class="info-item"><span>الفرع</span><strong id="s-branch">—</strong></div>
      <div class="info-item"><span>الإصدار</span><strong id="s-version">—</strong></div>
      <div class="info-item"><span id="uptime-label">وقت التشغيل</span><strong id="s-uptime">—</strong></div>
    </div>
  </div>

  <div class="card">
    <h2>الطابعات المتاحة</h2>
    <div class="printer-list" id="printer-list">جاري التحميل...</div>
  </div>

  <div class="card">
    <h2>آخر طلبات الطباعة اليدوية</h2>
    <table>
      <thead><tr><th>الطلب</th><th>الحالة</th><th>وقت الطلب</th><th>وقت الإنجاز</th></tr></thead>
      <tbody id="jobs-body"><tr><td colspan="4" style="color:#9ca3af;text-align:center">جاري التحميل...</td></tr></tbody>
    </table>
    <div class="refresh-note">يتحدث تلقائياً كل 15 ثانية</div>
  </div>

  <div class="card">
    <h2>إعادة تشغيل الخادم</h2>
    <p style="font-size:13px;color:#6b7280;margin-bottom:12px">سيتوقف الخادم ويعاد تشغيله تلقائياً عبر NSSM. قد تستغرق العملية 5-10 ثوانٍ.</p>
    <button class="btn-restart" id="restart-btn" onclick="doRestart()">إعادة التشغيل</button>
    <div id="restart-msg"></div>
  </div>
</main>

<script>
  const startTime = Date.now();

  function fmt(ms) {
    if (!ms) return '—';
    const d = new Date(ms);
    return d.toLocaleTimeString('ar-SA', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
  }

  function fmtUptime(ms) {
    const s = Math.floor(ms / 1000);
    const h = Math.floor(s / 3600), m = Math.floor((s % 3600) / 60), sec = s % 60;
    return (h ? h + 'س ' : '') + (m ? m + 'د ' : '') + sec + 'ث';
  }

  function badgeHtml(status) {
    const cls = status === 'done' ? 'badge-done' : status === 'pending' ? 'badge-pending' : 'badge-failed';
    const label = status === 'done' ? 'تم' : status === 'pending' ? 'قيد الانتظار' : 'فشل';
    return '<span class="badge ' + cls + '">' + label + '</span>';
  }

  async function loadHealth() {
    try {
      const r = await fetch('/health');
      const d = await r.json();
      document.getElementById('s-branch').textContent = d.branchId || '—';
      document.getElementById('s-version').textContent = 'v' + (d.version || '—');
      document.getElementById('ver-branch').textContent = 'فرع: ' + (d.branchId || '—') + '  |  v' + (d.version || '');
    } catch {}
  }

  async function loadPrinters() {
    try {
      const r = await fetch('/printers');
      const d = await r.json();
      const el = document.getElementById('printer-list');
      if (!d.printers || !d.printers.length) { el.textContent = 'لا توجد طابعات'; return; }
      el.innerHTML = d.printers.map(p => '<div class="printer-chip">' + p + '</div>').join('');
    } catch {}
  }

  async function loadJobs() {
    try {
      const r = await fetch('/jobs');
      const d = await r.json();
      const tbody = document.getElementById('jobs-body');
      if (!d.jobs || !d.jobs.length) {
        tbody.innerHTML = '<tr><td colspan="4" style="color:#9ca3af;text-align:center">لا توجد طلبات طباعة بعد</td></tr>';
        return;
      }
      tbody.innerHTML = d.jobs.map(j => {
        const errRow = j.error ? '<div class="error-text">' + j.error + '</div>' : '';
        return '<tr><td>' + (j.orderId || j.id).slice(-8) + errRow + '</td><td>' + badgeHtml(j.status) + '</td><td>' + fmt(j.requestedAt) + '</td><td>' + fmt(j.completedAt) + '</td></tr>';
      }).join('');
    } catch {}
  }

  function updateUptime() {
    document.getElementById('s-uptime').textContent = fmtUptime(Date.now() - startTime);
  }

  async function doRestart() {
    const btn = document.getElementById('restart-btn');
    const msg = document.getElementById('restart-msg');
    btn.disabled = true;
    msg.textContent = 'جاري إرسال أمر إعادة التشغيل...';
    try {
      await fetch('/restart', { method: 'POST' });
      msg.textContent = 'تم إرسال الأمر. سيعاد تشغيل الخادم خلال لحظات. أعد تحديث الصفحة بعد 10 ثوانٍ.';
      setTimeout(() => location.reload(), 8000);
    } catch {
      msg.textContent = 'تعذر الاتصال بالخادم.';
      btn.disabled = false;
    }
  }

  async function refresh() {
    await Promise.all([loadPrinters(), loadJobs()]);
  }

  loadHealth();
  refresh();
  setInterval(refresh, 15000);
  setInterval(updateUptime, 1000);
</script>
</body>
</html>`;

// ── Start ─────────────────────────────────────────────────────────────────────

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Mr. Healthy Print Server v${VERSION}`);
  console.log(`Listening on http://0.0.0.0:${PORT}`);
  getBrowser().catch(err => console.error("Browser pre-warm failed:", err.message));
  psWorker.start();
  startHeartbeat();
  startListeners();
  startWatchdog();
});

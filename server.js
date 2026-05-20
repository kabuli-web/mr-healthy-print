const express = require("express");
const cors = require("cors");

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

app.get("/printers", async (_req, res) => {
  try {
    const { ThermalPrinter, PrinterTypes } = require("node-thermal-printer");
    // node-thermal-printer exposes a static list via a dummy printer instance
    const printer = new ThermalPrinter({
      type: PrinterTypes.EPSON,
      interface: "printer:dummy",
    });
    const printers = await printer.getPrinters();
    const names = (printers || []).map((p) => (typeof p === "string" ? p : p.name || p.displayName || ""));
    res.json({ printers: names.filter(Boolean) });
  } catch (err) {
    // Fallback: return empty list — client can still type printer name manually
    console.error("Could not list printers:", err.message);
    res.json({ printers: [], warning: err.message });
  }
});

// ── Print order ─────────────────────────────────────────────────────────────

app.post("/print", async (req, res) => {
  const { printerName, order } = req.body || {};

  if (!printerName) return res.status(400).json({ success: false, error: "printerName is required" });
  if (!order) return res.status(400).json({ success: false, error: "order is required" });

  try {
    await printReceipt(printerName, order);
    res.json({ success: true });
  } catch (err) {
    console.error(`Print error [${printerName}]:`, err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ── Receipt builder ─────────────────────────────────────────────────────────

async function printReceipt(printerName, order) {
  const { ThermalPrinter, PrinterTypes, CharacterSet } = require("node-thermal-printer");

  // Support both USB (printer:Name) and TCP (TCP:IP:PORT)
  let iface;
  if (printerName.toUpperCase().startsWith("TCP:")) {
    // TCP:192.168.1.10:9100
    const parts = printerName.split(":");
    iface = `tcp://${parts[1]}:${parts[2] || 9100}`;
  } else {
    iface = `printer:${printerName}`;
  }

  const printer = new ThermalPrinter({
    type: PrinterTypes.EPSON,
    interface: iface,
    characterSet: CharacterSet.PC864_ARABIC,
    removeSpecialCharacters: false,
    lineCharacter: "-",
    width: 42,
  });

  const isConnected = await printer.isPrinterConnected();
  if (!isConnected) throw new Error(`Printer not connected: ${iface}`);

  const now = new Date().toLocaleString("ar-SA", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: true,
    timeZone: "Asia/Riyadh",
  });

  printer.alignCenter();
  printer.bold(true);
  printer.println("مستر صحي");
  printer.bold(false);
  printer.drawLine();

  printer.alignRight();
  printer.println(`طلب رقم: ${String(order.orderNumber || "").slice(-6)}`);
  printer.println(`الوقت: ${now}`);
  printer.drawLine();

  printer.println(`العميل: ${order.customerName || ""}`);
  if (order.customerPhone) printer.println(`الجوال: ${order.customerPhone}`);
  if (order.branchName) printer.println(`الفرع: ${order.branchName}`);

  if (order.dineType) {
    printer.println(order.dineType === "dine_in" ? "محلي" : "سفري");
  }
  if (order.fulfillmentType === "delivery") {
    printer.println("توصيل");
  }

  printer.drawLine();

  for (const item of order.items || []) {
    printer.bold(true);
    printer.println(`${item.quantity}x ${item.name}`);
    printer.bold(false);
    if (item.selectedAddons?.length) {
      for (const addon of item.selectedAddons) {
        const qty = addon.quantity ?? 1;
        printer.println(`  - ${addon.option?.name ?? ""}${qty > 1 ? ` x${qty}` : ""}`);
      }
    }
  }

  printer.drawLine();

  if (order.note) {
    printer.bold(true);
    printer.println(`ملاحظة: ${order.note}`);
    printer.bold(false);
    printer.drawLine();
  }

  printer.cut();
  await printer.execute();
  printer.clear();
}

// ── Start ───────────────────────────────────────────────────────────────────

app.listen(PORT, "127.0.0.1", () => {
  console.log(`Mr. Healthy Print Server v${VERSION}`);
  console.log(`Listening on http://127.0.0.1:${PORT}`);
  console.log("Press Ctrl+C to stop.");
});

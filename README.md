# Mr. Healthy Print Server

Lightweight HTTP server that exposes local thermal printers to the admin web app.

## Setup

```bash
npm install
node server.js
```

Server starts on `http://127.0.0.1:3001`.

## Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/health` | Health check |
| GET | `/printers` | List installed Windows printers |
| POST | `/print` | Print an order receipt |

## POST /print

```json
{
  "printerName": "Kitchen Printer 1",
  "order": {
    "orderNumber": "ORD-001234",
    "customerName": "محمد",
    "customerPhone": "0501234567",
    "branchName": "الرياض",
    "items": [{ "name": "دجاج مشوي", "quantity": 2, "selectedAddons": [] }],
    "note": "بدون ثوم",
    "dineType": "dine_in",
    "fulfillmentType": "branch"
  }
}
```

For TCP/network printers use `"printerName": "TCP:192.168.1.10:9100"`.

## Running on startup (Windows)

Create a shortcut to `start.bat` and place it in the Windows Startup folder (`shell:startup`):

```bat
@echo off
cd /d %~dp0
node server.js
pause
```

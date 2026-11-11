const { app, BrowserWindow, ipcMain, shell, dialog } = require("electron");
const path = require("path");
const fs = require("fs");
const https = require("https");

/* ---------- data locations ---------- */
const DATA_DIR = app.getPath("userData");
const DATA_FILE = path.join(DATA_DIR, "boekhouding.json");
const SETTINGS_FILE = path.join(DATA_DIR, "instellingen.json");

function readJSON(file, fallback) {
  try { return JSON.parse(fs.readFileSync(file, "utf8")); }
  catch { return fallback; }
}
function writeJSON(file, data) {
  fs.writeFileSync(file, JSON.stringify(data, null, 2), "utf8");
}

/* ---------- window ---------- */
let win;
function createWindow() {
  win = new BrowserWindow({
    width: 1280,
    height: 850,
    minWidth: 900,
    minHeight: 640,
    backgroundColor: "#F5F2EB",
    title: "Euro Superstore Boekhouding",
    icon: path.join(__dirname, "build", "icon.ico"),
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });
  win.removeMenu();
  win.loadFile(path.join(__dirname, "renderer", "index.html"));

  // open external links in the real browser
  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: "deny" };
  });
}

app.whenReady().then(() => {
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

/* ---------- IPC: storage ---------- */
ipcMain.handle("data:load", () => readJSON(DATA_FILE, []));
ipcMain.handle("data:save", (_e, entries) => { writeJSON(DATA_FILE, entries); return true; });

ipcMain.handle("settings:load", () => readJSON(SETTINGS_FILE, { apiKey: "" }));
ipcMain.handle("settings:save", (_e, s) => { writeJSON(SETTINGS_FILE, s); return true; });

ipcMain.handle("data:folder", () => { shell.openPath(DATA_DIR); return DATA_DIR; });

/* backup export to a chosen folder */
ipcMain.handle("data:saveFile", async (_e, { name, content, base64 }) => {
  const res = await dialog.showSaveDialog(win, { defaultPath: name });
  if (res.canceled || !res.filePath) return false;
  if (base64) fs.writeFileSync(res.filePath, Buffer.from(content, "base64"));
  else fs.writeFileSync(res.filePath, content, "utf8");
  return true;
});

/* ---------- IPC: invoice scan via Anthropic API ---------- */
const SYS_INVOICE = `Je leest Belgische aankoopfacturen. Geef ALLEEN geldige JSON terug, zonder uitleg, zonder markdown.
Velden:
{
 "datum":"YYYY-MM-DD",
 "leverancier":"naam",
 "totaal_incl_btw": number,
 "groep6_excl": number,
 "groep6_btw": number,
 "groep21_excl": number,
 "groep21_btw": number,
 "groep0_excl": number,
 "korting": number,
 "betaalmethode":"CASH|BC|TR|CR.NOTE"
}
Regels: gebruik het BTW-overzicht (Tax overview) onderaan de factuur. S6=6%, S21=21%, E0=0%. Negatieve bedragen behouden. Ontbreekt een veld: 0 of "".`;

const SYS_ZREPORT = `Je leest een Z-rapport (dagtotaal-ticket) van een kassa in een Belgische winkel.
Het ticket toont de dagomzet per departement of per BTW-tarief. Geef ALLEEN geldige JSON terug, zonder uitleg, zonder markdown.
Velden:
{
 "datum":"YYYY-MM-DD",
 "groep1_6pct_incl": number,
 "groep2_21pct_incl": number,
 "groep3_0pct_incl": number,
 "totaal": number
}
Regels:
- De bedragen zijn INCLUSIEF BTW, precies zoals ze op het ticket staan. Reken niets om.
- Departement 1 (DPT 1) hoort bij 6%, departement 2 bij 21%, departement 3 bij 0%. Staat er in plaats daarvan een BTW-percentage bij elke regel, volg dan dat percentage.
- Een eventueel departement 4 negeer je volledig.
- "totaal" is het dagtotaal van het ticket (TOTAAL / TL SALE / Total).
- Ontbreekt een bedrag, zet dan 0. Verzin nooit cijfers.`;

ipcMain.handle("scan:invoice", async (_e, { base64, mediaType, isPdf, mode }) => {
  const settings = readJSON(SETTINGS_FILE, { apiKey: "" });
  const key = (settings.apiKey || "").trim();
  if (!key) return { ok: false, error: "Geen API-sleutel ingesteld. Ga naar Instellingen." };

  const isZ = mode === "zreport";
  const sys = isZ ? SYS_ZREPORT : SYS_INVOICE;

  const block = isPdf
    ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } }
    : { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } };

  const payload = JSON.stringify({
    model: "claude-sonnet-4-6",
    max_tokens: 1000,
    system: sys,
    messages: [{ role: "user", content: [block, { type: "text", text: isZ ? "Lees dit Z-rapport en geef de JSON." : "Lees deze factuur en geef de JSON." }] }]
  });

  return await new Promise((resolve) => {
    const req = https.request({
      hostname: "api.anthropic.com",
      path: "/v1/messages",
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
        "Content-Length": Buffer.byteLength(payload)
      }
    }, (res) => {
      let body = "";
      res.on("data", c => body += c);
      res.on("end", () => {
        if (res.statusCode !== 200) {
          let msg = "De service antwoordde met code " + res.statusCode + ".";
          try { const j = JSON.parse(body); if (j.error && j.error.message) msg = j.error.message; } catch {}
          if (res.statusCode === 401) msg = "API-sleutel ongeldig. Controleer hem in Instellingen.";
          return resolve({ ok: false, error: msg });
        }
        try {
          const data = JSON.parse(body);
          const txt = (data.content || []).map(c => c.type === "text" ? c.text : "").join("").trim();
          resolve({ ok: true, text: txt });
        } catch (err) {
          resolve({ ok: false, error: "Antwoord kon niet gelezen worden." });
        }
      });
    });
    req.on("error", (err) => resolve({ ok: false, error: "Geen internetverbinding? (" + err.message + ")" }));
    req.write(payload);
    req.end();
  });
});

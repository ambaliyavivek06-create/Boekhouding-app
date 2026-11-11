/* =========================================================
   Excel-import voor Euro Superstore Boekhouding

   Leest twee soorten werkmappen:
     1. "SALE PURCHASE <jaar>.xlsx"
          CASH REGISTER -> verkoop (Z per dag, met BTW-verdeling)
          PURCHASE      -> aankoop
          EXPENSES      -> kosten
     2. "Kasboek<jaar>.xlsx"
          maandtabs     -> verkoop (Z per dag, alleen totaal)

   BTW-groepen volgen het bestand:
     groep 1 = 6%,  groep 2 = 21%,  groep 3 = 0%
   De BTW wordt overgenomen zoals ze in het bestand staat, niet
   herberekend. Het factuurtotaal is heilig: een verschil tussen de
   groepen en het totaal komt in "korting" terecht.

   Bankverkeer uit het kasboek (BC/MC, KBC, ...) wordt overgeslagen.
   ========================================================= */

(function (root) {

  const SKIP_NAMES = ["bc/mc", "bcmc", "bc / mc", "kbc", "mc/bc", "bank"];

  const MONTH_SHEETS = {
    january: 1, januari: 1, jan: 1,
    february: 2, februari: 2, feb: 2,
    march: 3, maart: 3, mar: 3,
    april: 4, apr: 4,
    may: 5, mei: 5,
    june: 6, juni: 6, jun: 6,
    july: 7, juli: 7, jul: 7,
    august: 8, augustus: 8, aug: 8,
    september: 9, sept: 9, sep: 9,
    october: 10, oktober: 10, oct: 10, okt: 10,
    november: 11, nov: 11,
    december: 12, dec: 12
  };

  /* ---------- kleine helpers ---------- */

  function toNum(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return isFinite(v) ? v : 0;
    const s = String(v).replace(/\s|€/g, "").replace(/\.(?=\d{3}\b)/g, "").replace(",", ".");
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }
  function isNumLike(v) {
    if (typeof v === "number") return isFinite(v);
    if (typeof v !== "string" || !v.trim()) return false;
    return !isNaN(parseFloat(v.replace(",", ".")));
  }
  function r2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
  function clean(v) { return String(v === null || v === undefined ? "" : v).trim(); }

  function isoDate(year, month, day) {
    const y = Number(year), m = Number(month), d = Number(day);
    if (!y || m < 1 || m > 12 || d < 1 || d > 31) return null;
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    const p = n => String(n).padStart(2, "0");
    return `${y}-${p(m)}-${p(d)}`;
  }

  function normPay(v) {
    const s = clean(v).toUpperCase().replace(/\./g, "").replace(/\s/g, "");
    if (!s) return "CASH";
    if (s.indexOf("CASH") === 0) return "CASH";
    if (s.indexOf("CRNOTE") === 0 || s.indexOf("CREDIT") === 0) return "CR.NOTE";
    if (s === "BC" || s.indexOf("BC") === 0) return "BC";
    if (s === "TR" || s.indexOf("TRANS") === 0 || s.indexOf("TR") === 0) return "TR";
    return "CASH";
  }

  function sheetRows(XLSX, wb, name) {
    const ws = wb.Sheets[name];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }

  function findSheet(wb, needle) {
    const want = needle.toLowerCase().replace(/\s/g, "");
    return wb.SheetNames.find(n => n.toLowerCase().replace(/\s/g, "") === want) ||
           wb.SheetNames.find(n => n.toLowerCase().replace(/\s/g, "").indexOf(want) !== -1) ||
           null;
  }

  function yearFromName(filename) {
    const m = String(filename || "").match(/(20\d\d)/);
    return m ? parseInt(m[1], 10) : null;
  }

  /* ---------- Een boeking opbouwen uit basis + BTW ----------
     groepen zijn EXCLUSIEF btw; btwTotaal is zoals in het bestand.
     Het totaal van de factuur stuurt de korting.                */

  function buildEntry(opts) {
    const b6 = r2(opts.b6), b21 = r2(opts.b21), b0 = r2(opts.b0);
    const btwTotal = r2(opts.btwTotal);

    // BTW opsplitsen over 6% en 21%, met de totale BTW als anker
    let t6 = 0, t21 = 0;
    if (b6 && b21) {
      t6 = r2(b6 * 0.06);
      t21 = r2(btwTotal - t6);
      if (t21 < 0) { t21 = 0; t6 = btwTotal; }
    } else if (b6) {
      t6 = btwTotal;
    } else if (b21) {
      t21 = btwTotal;
    } else if (btwTotal) {
      t21 = btwTotal; // btw zonder groep: op 21% zetten
    }

    // factuurtotaal is leidend; verschil in korting
    let korting = r2(opts.korting || 0);
    if (opts.total) {
      korting = r2(opts.total - (b6 + b21 + b0 + t6 + t21));
    }

    return {
      type: opts.type,
      datum: opts.datum,
      leverancier: opts.leverancier || "",
      betaalmethode: opts.betaalmethode || "CASH",
      b6: b6, t6: r2(t6), b21: b21, t21: r2(t21), b0: b0,
      korting: korting,
      bron: opts.bron || "import",
      onverdeeld: !!opts.onverdeeld
    };
  }

  /* ---------- CASH REGISTER -> verkoop ----------
     A=maand  C=dag  D=DPT1(6%)  E=DPT2(21%)  F=DPT3(0%)  H=TL SALE
     DPT-bedragen zijn INCLUSIEF btw.                              */

  function parseCashRegister(rows, year, out) {
    let n = 0;
    for (const row of rows) {
      if (!row) continue;
      const month = row[0], day = row[2], tl = row[7];
      if (!isNumLike(month) || !isNumLike(day) || !isNumLike(tl)) continue;
      const m = Math.round(toNum(month)), d = Math.round(toNum(day)), total = r2(toNum(tl));
      if (m < 1 || m > 12 || d < 1 || d > 31 || !total) continue;
      const datum = isoDate(year, m, d);
      if (!datum) { out.overgeslagen.push({ reden: "ongeldige datum", info: `maand ${m} dag ${d}` }); continue; }

      const i6 = r2(toNum(row[3]));   // incl. 6%
      const i21 = r2(toNum(row[4]));  // incl. 21%
      const i0 = r2(toNum(row[5]));   // 0%

      // incl -> excl + btw, zo dat basis + btw exact het inclusieve bedrag geeft
      const b6 = r2(i6 / 1.06), t6 = r2(i6 - r2(i6 / 1.06));
      const b21 = r2(i21 / 1.21), t21 = r2(i21 - r2(i21 / 1.21));

      const heeftSplit = !!(i6 || i21 || i0);
      const e = {
        type: "verkoop", datum: datum, leverancier: "Z-rapport",
        betaalmethode: "CASH",
        b6: heeftSplit ? b6 : 0, t6: heeftSplit ? t6 : 0,
        b21: heeftSplit ? b21 : 0, t21: heeftSplit ? t21 : 0,
        b0: heeftSplit ? i0 : total,
        korting: 0, bron: "import-z", onverdeeld: !heeftSplit
      };
      // totaal exact laten kloppen met TL SALE
      const calc = r2(e.b6 + e.t6 + e.b21 + e.t21 + e.b0);
      if (calc !== total) e.korting = r2(total - calc);

      out.entries.push(e);
      n++;
    }
    return n;
  }

  /* ---------- PURCHASE -> aankoop ----------
     A=maand B=dag C=naam D=bedrag E=btw F=betaalwijze
     G=groep1(6%) H=groep2(21%) I=groep3(0%) J=kort K=totaal      */

  function parsePurchase(rows, year, out) {
    let n = 0;
    for (const row of rows) {
      if (!row) continue;
      const month = row[0], day = row[1], amount = row[3];
      if (!isNumLike(month) || !isNumLike(day) || !isNumLike(amount)) continue;
      const m = Math.round(toNum(month)), d = Math.round(toNum(day));
      if (m < 1 || m > 12 || d < 1 || d > 31) continue;
      const amt = r2(toNum(amount));
      if (!amt) continue;

      const naam = clean(row[2]);
      if (SKIP_NAMES.indexOf(naam.toLowerCase()) !== -1) {
        out.overgeslagen.push({ reden: "bankverkeer", info: naam });
        continue;
      }
      const datum = isoDate(year, m, d);
      if (!datum) { out.overgeslagen.push({ reden: "ongeldige datum", info: `${naam} maand ${m} dag ${d}` }); continue; }

      let total = r2(toNum(row[10]));
      const btw = r2(toNum(row[4]));
      if (!total) total = r2(amt + btw);

      out.entries.push(buildEntry({
        type: "aankoop", datum: datum, leverancier: naam,
        betaalmethode: normPay(row[5]),
        b6: toNum(row[6]), b21: toNum(row[7]), b0: toNum(row[8]),
        btwTotal: btw, korting: toNum(row[9]), total: total,
        bron: "import-aankoop"
      }));
      n++;
    }
    return n;
  }

  /* ---------- EXPENSES -> kosten ----------
     A=maand B=dag C=naam D=bedrag E=btw F=betaalwijze G=totaal
     I=groep1(6%) J=groep2(21%) K=groep3(0%)                      */

  function parseExpenses(rows, year, out) {
    let n = 0;
    for (const row of rows) {
      if (!row) continue;
      const month = row[0], day = row[1], amount = row[3];
      if (!isNumLike(month) || !isNumLike(day) || !isNumLike(amount)) continue;
      const m = Math.round(toNum(month)), d = Math.round(toNum(day));
      if (m < 1 || m > 12 || d < 1 || d > 31) continue;
      const amt = r2(toNum(amount));
      if (!amt) continue;

      const naam = clean(row[2]);
      if (SKIP_NAMES.indexOf(naam.toLowerCase()) !== -1) {
        out.overgeslagen.push({ reden: "bankverkeer", info: naam });
        continue;
      }
      const datum = isoDate(year, m, d);
      if (!datum) { out.overgeslagen.push({ reden: "ongeldige datum", info: `${naam} maand ${m} dag ${d}` }); continue; }

      const btw = r2(toNum(row[4]));
      let total = r2(toNum(row[6]));
      if (!total) total = r2(amt + btw);

      out.entries.push(buildEntry({
        type: "kosten", datum: datum, leverancier: naam,
        betaalmethode: normPay(row[5]),
        b6: toNum(row[8]), b21: toNum(row[9]), b0: toNum(row[10]),
        btwTotal: btw, total: total,
        bron: "import-kosten"
      }));
      n++;
    }
    return n;
  }

  /* ---------- Kasboek maandtab -> verkoop (alleen totaal) ----------
     A=dag  C=Amount Deposited. De uitgavenkant (D/E/F) wordt niet
     ingelezen: die staat al in PURCHASE/EXPENSES.                   */

  function parseKasboekMonth(rows, year, month, out) {
    let n = 0;
    for (const row of rows) {
      if (!row) continue;
      const day = row[0], dep = row[2];
      if (!isNumLike(day) || !isNumLike(dep)) continue;
      const d = Math.round(toNum(day)), bedrag = r2(toNum(dep));
      if (d < 1 || d > 31 || !bedrag) continue;
      const datum = isoDate(year, month, d);
      if (!datum) continue;
      out.entries.push({
        type: "verkoop", datum: datum, leverancier: "Z-rapport",
        betaalmethode: "CASH",
        b6: 0, t6: 0, b21: 0, t21: 0, b0: bedrag, korting: 0,
        bron: "import-kasboek", onverdeeld: true
      });
      n++;
    }
    return n;
  }

  /* ---------- hoofdfunctie ---------- */

  /* ---------- Algemene verkooplijst ----------
     Voor een Excel die geen kasboek of SALE PURCHASE is: bijvoorbeeld een
     dagexport uit de kassa. We zoeken zelf een koprij met een datumkolom en
     een bedragkolom, en lezen de regels eronder.                          */

  function isDate(v) {
    // niet "instanceof Date": dat faalt zodra de xlsx-bibliotheek
    // in een andere context draait dan deze code
    return !!v && typeof v === "object" &&
           Object.prototype.toString.call(v) === "[object Date]" &&
           !isNaN(v.getTime());
  }

  function excelDate(v) {
    if (isDate(v)) {
      const p = n => String(n).padStart(2, "0");
      return `${v.getFullYear()}-${p(v.getMonth() + 1)}-${p(v.getDate())}`;
    }
    if (typeof v === "number" && v > 20000 && v < 80000) {
      // Excel telt dagen vanaf 30-12-1899
      const ms = Math.round((v - 25569) * 86400 * 1000);
      const d = new Date(ms);
      if (!isNaN(d)) {
        const p = n => String(n).padStart(2, "0");
        return `${d.getUTCFullYear()}-${p(d.getUTCMonth() + 1)}-${p(d.getUTCDate())}`;
      }
    }
    const s = clean(v);
    if (!s) return null;
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return isoDate(m[1], m[2], m[3]);
    m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
    if (m) {
      let y = parseInt(m[3], 10);
      if (y < 100) y += 2000;
      return isoDate(y, m[2], m[1]);   // dag/maand/jaar
    }
    return null;
  }

  const RE_DATE = /^(datum|date|dag|day)\b/i;
  const RE_TOTAL = /(totaal|total|bedrag|amount|omzet|sale|verkoop|tl\s*sale)/i;
  const RE_G6 = /(6\s*%|dpt\.?\s*#?\s*1\b|groep\s*1\b|afd\.?\s*1\b)/i;
  const RE_G21 = /(21\s*%|dpt\.?\s*#?\s*2\b|groep\s*2\b|afd\.?\s*2\b)/i;
  const RE_G0 = /(0\s*%|dpt\.?\s*#?\s*3\b|groep\s*3\b|afd\.?\s*3\b)/i;

  function findHeader(rows) {
    for (let i = 0; i < Math.min(rows.length, 40); i++) {
      const row = rows[i];
      if (!row) continue;
      let cDate = -1, cTotal = -1, c6 = -1, c21 = -1, c0 = -1;
      for (let c = 0; c < row.length; c++) {
        const t = clean(row[c]);
        if (!t) continue;
        if (cDate < 0 && RE_DATE.test(t)) { cDate = c; continue; }
        if (c6 < 0 && RE_G6.test(t)) { c6 = c; continue; }
        if (c21 < 0 && RE_G21.test(t)) { c21 = c; continue; }
        if (c0 < 0 && RE_G0.test(t)) { c0 = c; continue; }
        if (cTotal < 0 && RE_TOTAL.test(t)) { cTotal = c; continue; }
      }
      if (cDate >= 0 && (cTotal >= 0 || c6 >= 0 || c21 >= 0 || c0 >= 0)) {
        return { row: i, cDate, cTotal, c6, c21, c0 };
      }
    }
    return null;
  }

  function parseGenericSales(XLSX, wb, year, out) {
    let n = 0;
    for (const sheetName of wb.SheetNames) {
      const rows = sheetRows(XLSX, wb, sheetName);
      const h = findHeader(rows);
      if (!h) continue;

      for (let i = h.row + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        let datum = excelDate(row[h.cDate]);
        if (!datum && isNumLike(row[h.cDate])) {
          // alleen een dagnummer: dan moet de maand uit de tabnaam komen
          const d = Math.round(toNum(row[h.cDate]));
          const key = sheetName.trim().toLowerCase();
          const m = MONTH_SHEETS[key] || (/^\d{1,2}$/.test(key) ? parseInt(key, 10) : 0);
          if (m >= 1 && m <= 12 && d >= 1 && d <= 31) datum = isoDate(year, m, d);
        }
        if (!datum) continue;

        const i6 = h.c6 >= 0 ? r2(toNum(row[h.c6])) : 0;
        const i21 = h.c21 >= 0 ? r2(toNum(row[h.c21])) : 0;
        const i0 = h.c0 >= 0 ? r2(toNum(row[h.c0])) : 0;
        let total = h.cTotal >= 0 ? r2(toNum(row[h.cTotal])) : 0;
        const somGroepen = r2(i6 + i21 + i0);
        if (!total) total = somGroepen;
        if (!total) continue;

        const heeftSplit = !!somGroepen;
        const b6 = r2(i6 / 1.06), b21 = r2(i21 / 1.21);
        const e = {
          type: "verkoop", datum: datum, leverancier: "Z-rapport",
          betaalmethode: "CASH",
          b6: heeftSplit ? b6 : 0, t6: heeftSplit ? r2(i6 - b6) : 0,
          b21: heeftSplit ? b21 : 0, t21: heeftSplit ? r2(i21 - b21) : 0,
          b0: heeftSplit ? i0 : total,
          korting: 0, bron: "import-z", onverdeeld: !heeftSplit
        };
        const calc = r2(e.b6 + e.t6 + e.b21 + e.t21 + e.b0);
        if (calc !== total) e.korting = r2(total - calc);

        out.entries.push(e);
        n++;
      }
      if (n) break;   // eerste bruikbare tabblad volstaat
    }
    return n;
  }

  /* ---------- enkel de verkoop uit een bestand halen ----------
     Gebruikt bij Verkoop (Z): kasboek, SALE PURCHASE of een gewone
     lijst met datum- en bedragkolommen.                           */

  function parseSalesOnly(XLSX, data, filename, yearOverride) {
    const wb = XLSX.read(data, { type: "array", cellDates: true });
    const out = { entries: [], overgeslagen: [], perSoort: {}, bestandstype: null, jaar: null, waarschuwingen: [] };
    const year = yearOverride || yearFromName(filename) || new Date().getFullYear();
    out.jaar = year;

    const sCash = findSheet(wb, "CASH REGISTER");
    if (sCash) {
      out.bestandstype = "SALE PURCHASE";
      out.perSoort.verkoop = parseCashRegister(sheetRows(XLSX, wb, sCash), year, out);
      return out;
    }

    const monthSheets = monthSheetsOf(wb);
    if (monthSheets.length >= 2 && !findSheet(wb, "PURCHASE")) {
      out.bestandstype = "Kasboek";
      let n = 0;
      for (const ms of monthSheets) n += parseKasboekMonth(sheetRows(XLSX, wb, ms.name), year, ms.month, out);
      out.perSoort.verkoop = n;
      out.waarschuwingen.push("Dit kasboek bevat de dagomzet zonder BTW-verdeling.");
      return out;
    }

    const n = parseGenericSales(XLSX, wb, year, out);
    if (n) {
      out.bestandstype = "Verkooplijst";
      out.perSoort.verkoop = n;
      if (out.entries.some(e => e.onverdeeld))
        out.waarschuwingen.push("Er is geen kolom per BTW-groep gevonden, dus alleen het dagtotaal is ingelezen.");
    } else {
      out.waarschuwingen.push("Geen verkoopregels gevonden. Zorg dat er een kolom met een datum staat en een kolom met het bedrag. Tabbladen in dit bestand: " + wb.SheetNames.join(", "));
    }
    return out;
  }

  function monthSheetsOf(wb) {
    const list = [];
    for (const name of wb.SheetNames) {
      const key = name.trim().toLowerCase();
      if (MONTH_SHEETS[key]) list.push({ name: name, month: MONTH_SHEETS[key] });
      else if (/^\d{1,2}$/.test(key)) {
        const m = parseInt(key, 10);
        if (m >= 1 && m <= 12) list.push({ name: name, month: m });
      }
    }
    return list;
  }

  function parseWorkbook(XLSX, data, filename, yearOverride) {
    const wb = XLSX.read(data, { type: "array", cellDates: true });
    const out = { entries: [], overgeslagen: [], perSoort: {}, bestandstype: null, jaar: null, waarschuwingen: [] };

    const year = yearOverride || yearFromName(filename) || new Date().getFullYear();
    out.jaar = year;

    const sCash = findSheet(wb, "CASH REGISTER");
    const sPur = findSheet(wb, "PURCHASE");
    const sExp = findSheet(wb, "EXPENSES");

    // maandtabs van een kasboek opsporen
    const monthSheets = monthSheetsOf(wb);

    const isSalePurchase = !!(sCash || sPur || sExp);
    // een kasboek heeft maandtabs maar geen PURCHASE/CASH REGISTER
    const isKasboek = !isSalePurchase && monthSheets.length >= 2;

    if (isSalePurchase) {
      out.bestandstype = "SALE PURCHASE";
      if (sCash) out.perSoort.verkoop = parseCashRegister(sheetRows(XLSX, wb, sCash), year, out);
      if (sPur) out.perSoort.aankoop = parsePurchase(sheetRows(XLSX, wb, sPur), year, out);
      if (sExp) out.perSoort.kosten = parseExpenses(sheetRows(XLSX, wb, sExp), year, out);
      if (!sCash) out.waarschuwingen.push("Geen tabblad CASH REGISTER gevonden, dus geen verkoop ingelezen.");
      if (!sPur) out.waarschuwingen.push("Geen tabblad PURCHASE gevonden, dus geen aankopen ingelezen.");
    } else if (isKasboek) {
      out.bestandstype = "Kasboek";
      let n = 0;
      for (const ms of monthSheets) n += parseKasboekMonth(sheetRows(XLSX, wb, ms.name), year, ms.month, out);
      out.perSoort.verkoop = n;
      out.waarschuwingen.push("Een kasboek bevat de dagomzet zonder BTW-verdeling. Importeer ook SALE PURCHASE voor de verdeling over 6%, 21% en 0%.");
    } else {
      out.bestandstype = null;
      out.waarschuwingen.push("Dit bestand lijkt geen kasboek of SALE PURCHASE te zijn. Gevonden tabbladen: " + wb.SheetNames.join(", "));
    }

    return out;
  }

  const api = {
    parseWorkbook: parseWorkbook,
    parseSalesOnly: parseSalesOnly,
    _helpers: { toNum: toNum, r2: r2, normPay: normPay, isoDate: isoDate, buildEntry: buildEntry, excelDate: excelDate }
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.BoekImport = api;

})(typeof self !== "undefined" ? self : this);

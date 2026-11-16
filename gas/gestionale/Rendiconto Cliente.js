// ═════════════════════════════════════════════════════════════
// PURCHASE SUMMARY — DEKKA Client Statement
// Template: DEKKA_Cristian_Bertolini__Purchase_Statement.xlsx
//           + manuale Cinalli Elizabeth 12/06-12/07/2026
// ═════════════════════════════════════════════════════════════
// Aggiungi in onOpen() prima di .addToUi():
//   .addSeparator()
//   .addItem("Genera Rendiconto Cliente", "generaRendiconto")
// ═════════════════════════════════════════════════════════════

const RENDICONTI_FOLDER_ID = "166ROQ_zkTE3xZDX6rwPHBYIt-ZSluAtL";

// Colonne Summary sheet
const S_COL_W = { A:15, B:285, C:90, D:105, E:105, F:15 };

// Colori pixel-matched dal template
const C = {
  TEAL    : "#4DA6C8",
  TEAL2   : "#50A6C8",
  WHITE   : "#FFFFFF",
  GRAY_R  : "#F2F2F2",
  GRAY_ED : "#EDEDED",
  DARK_SEP: "#D9D9D9",
  CCCCCC  : "#CCCCCC",
  SUMMARY : "#FAFAFA",
  TEXT    : "#1A1A1A",
  SUBTEXT : "#555555",
  FOOTER  : "#444444",
  LINK    : "#1155CC",
};

// Colonne Order Detail sheet
const D_COL_W = { A:15, B:33, C:81, D:87, E:346, F:33, G:39, H:87, I:15 };
const D_H = { top:5.25, title:13.5, sep:3.75, gap:7.5, header:21.75, row:12.0, thin:2, total:15.75 };

// ─────────────────────────────────────────────────────────────
// Entry point
// ─────────────────────────────────────────────────────────────
function generaRendiconto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. Selezione cliente
  const clientSheets = ss.getSheets()
    .filter(s => s.getName() !== "Dashboard").map(s => s.getName());
  if (!clientSheets.length) { ui.alert("No clients available."); return; }

  const cR = ui.prompt("Purchase Statement — Client",
    "Available clients:\n• " + clientSheets.join("\n• ") + "\n\nClient name:", ui.ButtonSet.OK_CANCEL);
  if (cR.getSelectedButton() !== ui.Button.OK) return;
  const clientName = cR.getResponseText().trim();
  if (!clientSheets.includes(clientName)) { ui.alert('Client not found: "' + clientName + '".'); return; }

  // 2. Periodo
  const iR = ui.prompt("Purchase Statement — Period", "Period START date (dd/mm/yyyy):", ui.ButtonSet.OK_CANCEL);
  if (iR.getSelectedButton() !== ui.Button.OK) return;
  const dataInizio = parseDate_(iR.getResponseText().trim());
  if (!dataInizio) { ui.alert("Invalid date. Use dd/mm/yyyy."); return; }

  const fR = ui.prompt("Purchase Statement — Period", "Period END date (dd/mm/yyyy):", ui.ButtonSet.OK_CANCEL);
  if (fR.getSelectedButton() !== ui.Button.OK) return;
  const dataFineRaw = parseDate_(fR.getResponseText().trim());
  if (!dataFineRaw) { ui.alert("Invalid date. Use dd/mm/yyyy."); return; }
  const dataFine = clampMonthEnd_(dataFineRaw);
  dataFine.setHours(23, 59, 59);
  if (dataInizio > dataFine) { ui.alert("Start date must precede end date."); return; }

  // 3. Saldo iniziale — letto dalla Dashboard automaticamente
  //    e mostrato all'utente per conferma/modifica
  const saldoInizialeDash = leggiSaldoDashboard_(ss, clientName);
  const saldoInizialeResp = ui.prompt(
    "Purchase Statement — Opening Balance",
    "Opening account balance for this period (JPY).\n" +
    "Current balance from Dashboard: " + (saldoInizialeDash !== null ? "¥ " + Math.round(saldoInizialeDash).toLocaleString('ja-JP') : "not available") + "\n\n" +
    "Confirm or enter a different value (numbers only):",
    ui.ButtonSet.OK_CANCEL
  );
  if (saldoInizialeResp.getSelectedButton() !== ui.Button.OK) return;
  const saldoInizialeRaw = saldoInizialeResp.getResponseText().trim().replace(/[^0-9.\-]/g, "");
  const saldoIniziale = saldoInizialeRaw !== "" ? parseFloat(saldoInizialeRaw) : (saldoInizialeDash || 0);

  // 4. Fondi ricevuti
  const fnR = ui.prompt("Purchase Statement — Funds Received",
    "Total procurement funds received from client in this period (JPY, numbers only):", ui.ButtonSet.OK_CANCEL);
  if (fnR.getSelectedButton() !== ui.Button.OK) return;
  const fondiRicevuti = parseFloat(fnR.getResponseText().trim().replace(/[^0-9.]/g, "")) || 0;

  // 5. Rimborso (opzionale)
  const rfR = ui.prompt("Purchase Statement — Refund (optional)",
    "Refund amount to credit to client in this period (JPY, numbers only).\nLeave empty or enter 0 if not applicable:",
    ui.ButtonSet.OK_CANCEL);
  if (rfR.getSelectedButton() !== ui.Button.OK) return;
  const rimborso = parseFloat(rfR.getResponseText().trim().replace(/[^0-9.]/g, "")) || 0;

  // 6. Descrizione rimborso (solo se rimborso > 0)
  let rimborsoDesc = "";
  if (rimborso > 0) {
    const rdR = ui.prompt("Purchase Statement — Refund Description",
      "Brief description for the refund line (e.g. \"Refund for lost item\"):",
      ui.ButtonSet.OK_CANCEL);
    if (rdR.getSelectedButton() !== ui.Button.OK) return;
    rimborsoDesc = rdR.getResponseText().trim() || "Refund";
  }

  // 7. Allegato dettaglio
  const aR = ui.alert("Purchase Statement — Annex",
    "Include the order detail sheet?", ui.ButtonSet.YES_NO);
  const withDetail = (aR === ui.Button.YES);

  // 8. Raccolta ordini
  const cSheet = ss.getSheetByName(clientName);
  const lastRow = cSheet.getLastRow();
  const STATI   = ["PAGATO", "RICEVUTO", "SPEDITO"];
  const ordini  = [];

  if (lastRow >= 3) {
    const dati = cSheet.getRange(3, G_COL.DATA_RICEZIONE, lastRow - 2,
      G_COL.IMPORTI - G_COL.DATA_RICEZIONE + 1).getValues();
    for (const r of dati) {
      const dr    = r[0];
      const stato = r[G_COL.STATO    - G_COL.DATA_RICEZIONE];
      const lk    = r[G_COL.LINK     - G_COL.DATA_RICEZIONE];
      const qt    = r[G_COL.QTA      - G_COL.DATA_RICEZIONE];
      const im    = r[G_COL.IMPORTI  - G_COL.DATA_RICEZIONE];
      if (!dr || !lk || !STATI.includes(stato)) continue;
      const d = new Date(dr);
      if (d < dataInizio || d > dataFine) continue;
      const qn = parseInt(qt) || 1;
      ordini.push({
        data   : Utilities.formatDate(d, "Asia/Tokyo", "dd/MM/yyyy"),
        stato, link: lk, qta: qn, unita: Math.ceil(qn / 25), importo: im || 0
      });
    }
  }
  if (!ordini.length) { ui.alert("No purchases found."); return; }

  const totAcq  = ordini.reduce((a, o) => a + o.importo, 0);
  const totUnit = ordini.reduce((a, o) => a + o.unita,   0);

  // Saldo finale = apertura + fondi ricevuti + rimborso - acquisti
  const saldoFinale = saldoIniziale + fondiRicevuti - rimborso - totAcq;

  // 9. Piano dalla Dashboard
  const dash = ss.getSheetByName("Dashboard");
  let piano = "";
  if (dash && dash.getLastRow() >= 2) {
    const dv = dash.getRange(2, D_COL.CLIENTE, dash.getLastRow() - 1, 2).getValues();
    for (const r of dv) { if (r[0] === clientName) { piano = r[1] || ""; break; } }
  }

  // 10. Dati anagrafici dalle PropertiesService
  const nomeDoc   = cognomeNome_(clientName);
  const piva      = getClientPiva_(clientName);
  const indirizzo = getClientIndirizzo_(clientName);

  // 11. Creazione file
  const titolo = "Purchase Statement - " + fmtDate_(dataInizio) + " - " + fmtDate_(dataFine);
  const nuovoSS = SpreadsheetApp.create(titolo);

  try {
    const file        = DriveApp.getFileById(nuovoSS.getId());
    const rootFolder  = DriveApp.getFolderById(RENDICONTI_FOLDER_ID);
    const clientFolder = getOrCreateSubfolder_(rootFolder, clientName);
    clientFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  } catch(e) { Logger.log("Folder error: " + e.message); }

  const f1 = nuovoSS.getActiveSheet();
  f1.setName("Summary");
  buildSummary_(f1, nomeDoc, piva, indirizzo, piano,
    dataInizio, dataFine,
    saldoIniziale, fondiRicevuti, rimborso, rimborsoDesc,
    totAcq, saldoFinale, ordini.length, totUnit);

  if (withDetail) {
    buildDetail_(nuovoSS.insertSheet("Order Detail"), nomeDoc, dataInizio, dataFine, ordini);
  }

  SpreadsheetApp.flush();
  ui.alert(
    "✅ Statement generated!\n\n" +
    "Client: "          + nomeDoc + "\n" +
    "Period: "          + fmtDate_(dataInizio) + " – " + fmtDate_(dataFine) + "\n" +
    "Opening balance: ¥" + Math.round(saldoIniziale).toLocaleString('ja-JP') + "\n" +
    "Orders: "          + ordini.length + " (" + totUnit + " plan units)\n" +
    "Total purchases: ¥" + Math.round(totAcq).toLocaleString('ja-JP') + "\n" +
    "Ending balance: ¥"  + Math.round(saldoFinale).toLocaleString('ja-JP') + "\n\n" +
    "URL: " + nuovoSS.getUrl()
  );
}

// ─────────────────────────────────────────────────────────────
// Legge il saldo numerico dalla Dashboard per un cliente
// Gestisce sia il formato "¥1,234" che numeri puri
// ─────────────────────────────────────────────────────────────
function leggiSaldoDashboard_(ss, clientName) {
  const dash = ss.getSheetByName("Dashboard");
  if (!dash || dash.getLastRow() < 2) return null;
  const dati = dash.getRange(2, D_COL.CLIENTE, dash.getLastRow() - 1,
    D_COL.SALDO - D_COL.CLIENTE + 1).getValues();
  for (const r of dati) {
    if (r[0] !== clientName) continue;
    const raw = r[D_COL.SALDO - D_COL.CLIENTE];
    if (raw === "" || raw === null) return null;
    // Se è già un numero
    if (typeof raw === "number") return raw;
    // Se è una stringa tipo "¥1,234"
    const n = parseFloat(raw.toString().replace(/[^0-9.\-]/g, ""));
    return isNaN(n) ? null : n;
  }
  return null;
}

// ─────────────────────────────────────────────────────────────
// SUMMARY sheet builder
//
// Struttura righe (dal template + modifiche manuali Cinalli):
//  1  h=6       margin
//  2  h=6       margin
//  3  h=15.75   B3:C3 "Statement ref…" | D3:E3 data
//  4  h=9.75    empty
//  5  h=27.75   DEKKA (bold 26, right)
//  6  h=15.75   by TARABINI (right)
//  7  h=13.5    empty
//  8  h=19.5    nome cliente (bold 12, left)
//  9  h=19.5    📋 VAT/Tax ID
// 10  h=19.5    📍 indirizzo
// 11  h=19.5    DEKKA Plan
// 12  h=19.5    Reference period
// 13  h=3.75    banda teal
// 14  h=13.5    empty
// 15  h=7.5     empty
// 16  h=21.75   header tabella (Description | Amount)
// 17  h=19.5    Available balance — Opening account balance
// 18  h=19.5    Procurement funds received
// 19  h=19.5    Refund (solo se > 0, altrimenti riga vuota)
// 20  h=19.5    Total purchases executed
// 21  h=19.5    Ending available balance
// 22-26 h=19.5  righe vuote (riempimento)
// 27  h=6       banda scura #D9D9D9
// 28  h=13.5    empty
// 29  h=19.5    Subtotal | valore
// 30  h=19.5    Tax (0%) | ¥0
// 31  h=24      Total | valore (bold, #D9D9D9)
// 32  h=15.75   empty
// 33  h=15.75   Notes label
// 34  h=auto    Notes body
// 35  h=13.5    empty
// 36  h=13.5    empty
// 37  h=3.75    banda grigia #CCCCCC
// 38  h=15.75   footer
// 39  h=15.75   empty
// 40  h=6       margin
// ─────────────────────────────────────────────────────────────
function buildSummary_(sheet, nome, piva, indirizzo, piano,
    di, df,
    saldoIniziale, fondiRicevuti, rimborso, rimborsoDesc,
    totAcq, saldoFinale, nLink, nUnit) {

  // Larghezze colonne
  sheet.setColumnWidth(1, S_COL_W.A);
  sheet.setColumnWidth(2, S_COL_W.B);
  sheet.setColumnWidth(3, S_COL_W.C);
  sheet.setColumnWidth(4, S_COL_W.D);
  sheet.setColumnWidth(5, S_COL_W.E);
  sheet.setColumnWidth(6, S_COL_W.F);
  if (sheet.getMaxColumns() > 6) sheet.hideColumns(7, sheet.getMaxColumns() - 6);

  // Altezze righe (array 1-indexed, null = auto)
  const SH = [0,6,6,15.75,9.75,27.75,15.75,13.5,19.5,19.5,19.5,19.5,19.5,3.75,13.5,7.5,
              21.75,19.5,19.5,19.5,19.5,19.5,19.5,19.5,19.5,19.5,19.5,
              6,13.5,19.5,19.5,24,15.75,15.75,null,13.5,13.5,3.75,15.75,15.75,6];
  for (let i = 1; i < SH.length; i++) {
    if (SH[i] !== null) sheet.setRowHeight(i, SH[i]);
  }

  const set = (r, c, v, o) => { const x = sheet.getRange(r, c); x.setValue(v); ap_(x, o); };
  const mrg = (r, c1, c2, v, o) => {
    const x = sheet.getRange(r, c1, 1, c2 - c1 + 1); x.merge(); x.setValue(v); ap_(x, o);
  };

  // row 3: ref | data
  mrg(3, 2, 3, "Statement ref. — / " + new Date().getFullYear(),
    { fg: C.SUBTEXT, sz: 9, ha: "left", va: "center" });
  mrg(3, 4, 5, Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd"),
    { fg: C.SUBTEXT, sz: 9, ha: "right", va: "center" });

  // row 5: DEKKA
  mrg(5, 2, 5, "DEKKA", { fg: C.TEXT, bold: true, sz: 26, ha: "right", va: "center" });

  // row 6: by TARABINI
  mrg(6, 2, 5, "by TARABINI", { fg: C.SUBTEXT, sz: 10, ha: "right", va: "center" });

  // row 8: nome cliente
  mrg(8, 2, 5, nome, { fg: C.TEXT, bold: true, sz: 12, ha: "left", va: "center" });

  // row 9: VAT
  mrg(9, 2, 5, "VAT/Tax ID: " + (piva || "—"), { fg: C.SUBTEXT, sz: 8, ha: "left", va: "center" });

  // row 10: indirizzo
  mrg(10, 2, 5, (indirizzo || "—"), { fg: C.SUBTEXT, sz: 8, ha: "left", va: "center", wrap: true });

  // row 11: piano
  mrg(11, 2, 5, "DEKKA Plan: " + (piano || "—"), { fg: C.SUBTEXT, sz: 8, va: "center" });

  // row 12: periodo
  mrg(12, 2, 5, "Reference period: " + fmtDate_(di) + " — " + fmtDate_(df),
    { fg: C.SUBTEXT, sz: 8, va: "center" });

  // row 13: banda teal
  sheet.getRange(13, 2, 1, 4).setBackground(C.TEAL);

  // row 16: header tabella
  mrg(16, 2, 4, "Description", { bg: C.TEAL, fg: C.WHITE, bold: true, sz: 10, ha: "left",  va: "center" });
  set(16, 5,    "Amount",      { bg: C.TEAL, fg: C.WHITE, bold: true, sz: 10, ha: "right", va: "center" });

  // ── 5 righe dati ─────────────────────────────────────────
  // Palette alternata: riga 17 = GRAY_R, 18 = WHITE, 19 = GRAY_R, 20 = WHITE, 21 = GRAY_R
  const bgs = [C.GRAY_R, C.WHITE, C.GRAY_R, C.WHITE, C.GRAY_R];

  // row 17: Opening balance
  mrg(17, 2, 4, "Available balance — Opening account balance",
    { bg: bgs[0], fg: C.TEXT, sz: 9, ha: "left", va: "center" });
  set(17, 5, saldoIniziale,
    { bg: bgs[0], fg: C.TEXT, sz: 9, ha: "right", va: "center", nf: '"¥ "#,##0' });

  // row 18: Funds received
  mrg(18, 2, 4, "Procurement funds received — advance payments from client",
    { bg: bgs[1], fg: C.TEXT, sz: 9, ha: "left", va: "center" });
  set(18, 5, fondiRicevuti,
    { bg: bgs[1], fg: C.TEXT, sz: 9, ha: "right", va: "center", nf: '"¥ "#,##0' });

  // row 19: Refund (opzionale) — se 0 la riga rimane vuota ma mantiene il colore
  if (rimborso > 0) {
    mrg(19, 2, 4, rimborsoDesc || "Refund",
      { bg: bgs[2], fg: C.TEXT, sz: 9, ha: "left", va: "center" });
    set(19, 5, rimborso,
      { bg: bgs[2], fg: C.TEXT, sz: 9, ha: "right", va: "center", nf: '"¥ "#,##0' });
  } else {
    mrg(19, 2, 4, "", { bg: bgs[2] });
    set(19, 5, "", { bg: bgs[2] });
  }

  // row 20: Total purchases
  mrg(20, 2, 4,
    "Total purchases executed — " + nLink + " order" + (nLink !== 1 ? "s" : "") +
    "  ·  " + nUnit + " plan unit" + (nUnit !== 1 ? "s" : ""),
    { bg: bgs[3], fg: C.TEXT, sz: 9, ha: "left", va: "center" });
  set(20, 5, totAcq,
    { bg: bgs[3], fg: C.TEXT, sz: 9, ha: "right", va: "center", nf: '"¥ "#,##0' });

  // row 21: Ending balance
  mrg(21, 2, 4, "Ending available balance",
    { bg: bgs[4], fg: C.TEXT, sz: 9, ha: "left", va: "center" });
  set(21, 5, saldoFinale,
    { bg: bgs[4], fg: saldoFinale < 0 ? "#C0392B" : C.TEXT, sz: 9,
      ha: "right", va: "center", nf: '"¥ "#,##0' });

  // rows 22-26: righe vuote di riempimento
  const fillBgs = [C.WHITE, C.GRAY_R, C.WHITE, C.GRAY_R, C.WHITE];
  for (let i = 0; i < 5; i++) {
    mrg(22 + i, 2, 4, "", { bg: fillBgs[i] });
    set(22 + i, 5, "", { bg: fillBgs[i] });
  }

  // row 27: banda scura
  sheet.getRange(27, 2, 1, 4).setBackground(C.DARK_SEP);

  // row 29: Subtotal (= total purchases, non ending balance)
  set(29, 4, "Subtotal",
    { fg: C.TEXT, sz: 9, ha: "right", va: "center", borderTop: true });
  set(29, 5, totAcq,
    { bg: C.GRAY_R, fg: C.TEXT, sz: 9, ha: "right", va: "center",
      nf: '"¥ "#,##0', borderTop: true });

  // row 30: Tax
  set(30, 4, "Tax (0%)\nexport of services",
    { fg: C.TEXT, sz: 9, ha: "right", va: "center", wrap: true });
  set(30, 5, 0,
    { bg: C.GRAY_R, fg: C.TEXT, sz: 9, ha: "right", va: "center", nf: '"¥ "#,##0' });

  // row 31: Total
  set(31, 4, "Total",
    { bg: C.DARK_SEP, fg: C.TEXT, bold: true, sz: 11, ha: "right", va: "center", borderBottom: true });
  set(31, 5, totAcq,
    { bg: C.DARK_SEP, fg: C.TEXT, bold: true, sz: 11, ha: "right", va: "center",
      nf: '"¥ "#,##0', borderBottom: true });

  // row 33: Notes label
  mrg(33, 2, 5, "Notes", { bold: true, sz: 9, va: "center" });

  // row 34: Notes body
  mrg(34, 2, 5,
    "This document constitutes an informational purchase statement for goods acquired through " +
    "the DEKKA procurement service during the reference period, acting as agent " +
    "(mandatario senza rappresentanza) on behalf of the client. All amounts are denominated " +
    "in Japanese Yen (JPY). No VAT applies — transactions are outside the scope of VAT as " +
    "purchases executed abroad. Procurement funds advanced by the client are applied exclusively " +
    "to the purchases indicated and are not held, safeguarded, or managed as third-party funds.",
    { fg: C.SUBTEXT, sz: 7, ha: "left", va: "top", wrap: true });

  // row 37: separatore footer
  sheet.getRange(37, 2, 1, 4).setBackground(C.CCCCCC);

  // row 38: footer
  mrg(38, 2, 3, "Higashi Sapporo, Shiroishi-ku, Sapporo, Hokkaido, Japan",
    { fg: C.FOOTER, sz: 8, ha: "left", va: "center" });
  mrg(38, 4, 5, "T3810427739547",
    { fg: C.FOOTER, sz: 8, ha: "left", va: "center" });

  sheet.setFrozenRows(0);
  Logger.log("Summary built.");
}

// ─────────────────────────────────────────────────────────────
// ORDER DETAIL sheet builder
// ─────────────────────────────────────────────────────────────
function buildDetail_(sheet, nome, di, df, ordini) {
  sheet.setColumnWidth(1, D_COL_W.A);
  sheet.setColumnWidth(2, D_COL_W.B);
  sheet.setColumnWidth(3, D_COL_W.C);
  sheet.setColumnWidth(4, D_COL_W.D);
  sheet.setColumnWidth(5, D_COL_W.E);
  sheet.setColumnWidth(6, D_COL_W.F);
  sheet.setColumnWidth(7, D_COL_W.G);
  sheet.setColumnWidth(8, D_COL_W.H);
  sheet.setColumnWidth(9, D_COL_W.I);
  if (sheet.getMaxColumns() >= 10) sheet.hideColumns(10, sheet.getMaxColumns() - 9);

  const set = (r, c, v, o) => { const x = sheet.getRange(r, c); x.setValue(v); ap_(x, o); };
  const mrg = (r, c1, c2, v, o) => {
    const x = sheet.getRange(r, c1, 1, c2 - c1 + 1);
    if (c1 !== c2) x.merge(); x.setValue(v); ap_(x, o);
  };

  sheet.setRowHeight(1, D_H.top);

  sheet.setRowHeight(2, D_H.title);
  mrg(2, 2, 8,
    "DEKKA — Purchase Detail: " + nome + "  |  " + fmtDate_(di) + " – " + fmtDate_(df),
    { fg: C.TEXT, bold: true, sz: 12, ha: "left", va: "bottom" });

  sheet.setRowHeight(4, D_H.sep);
  sheet.getRange(4, 2, 1, 7).setBackground(C.TEAL2);

  sheet.setRowHeight(5, D_H.title);
  sheet.setRowHeight(6, D_H.gap);

  sheet.setRowHeight(7, D_H.header);
  const hLbl = ["#", "Date", "Status", "Source link", "Qty", "Units", "Amount"];
  const hCol = [2, 3, 4, 5, 6, 7, 8];
  const hAln = ["center", "center", "center", "left", "center", "center", "right"];
  for (let i = 0; i < hLbl.length; i++) {
    set(7, hCol[i], hLbl[i],
      { bg: C.TEAL2, fg: C.WHITE, bold: true, sz: 10, ha: hAln[i], va: "center" });
  }

  const firstDataRow = 8;
  for (let i = 0; i < ordini.length; i++) {
    const o  = ordini[i];
    const r  = firstDataRow + i;
    const bg = (i % 2 === 0) ? C.WHITE : C.GRAY_ED;
    sheet.setRowHeight(r, D_H.row);
    set(r, 2, i + 1,    { bg, fg: C.SUBTEXT, sz: 9, ha: "center", va: "bottom" });
    set(r, 3, o.data,   { bg, fg: C.TEXT,    sz: 9, ha: "center", va: "bottom" });
    set(r, 4, o.stato,  { bg, fg: C.TEXT,    sz: 9, ha: "center", va: "bottom" });
    set(r, 5, o.link,   { bg, fg: C.LINK,    sz: 9, ha: "left",   va: "bottom", clip: true });
    set(r, 6, o.qta,    { bg, fg: C.TEXT,    sz: 9, ha: "center", va: "bottom" });
    set(r, 7, o.unita,  { bg, fg: C.SUBTEXT, sz: 9, ha: "center", va: "bottom" });
    set(r, 8, o.importo,{ bg, fg: C.TEXT,    sz: 9, ha: "right",  va: "bottom", nf: '[$¥-411]#,##0' });
  }

  const lastDataRow = firstDataRow + ordini.length - 1;
  const sepRow      = lastDataRow + 1;
  const totRow      = sepRow + 1;

  sheet.setRowHeight(sepRow, D_H.thin);
  sheet.getRange(sepRow, 2, 1, 7).setBackground(C.SUBTEXT);

  sheet.setRowHeight(totRow, D_H.total);
  const totQta  = ordini.reduce((a, o) => a + o.qta,   0);
  const totUnit = ordini.reduce((a, o) => a + o.unita, 0);
  mrg(totRow, 2, 4, "TOTAL",
    { bg: C.SUMMARY, fg: C.TEXT, bold: true, sz: 9, ha: "right", va: "bottom" });
  set(totRow, 5, "", { bg: C.SUMMARY });
  set(totRow, 6, totQta,
    { bg: C.SUMMARY, fg: C.TEXT, bold: true, sz: 9, ha: "center", va: "bottom" });
  set(totRow, 7, totUnit,
    { bg: C.SUMMARY, fg: C.TEXT, bold: true, sz: 9, ha: "center", va: "bottom" });
  set(totRow, 8, "=SUM(H" + firstDataRow + ":H" + lastDataRow + ")",
    { bg: C.SUMMARY, fg: C.TEXT, bold: true, sz: 9, ha: "right", va: "bottom",
      nf: '[$¥-411]#,##0' });

  sheet.setFrozenRows(0);
  Logger.log("Detail built — " + ordini.length + " rows.");
}

// ─────────────────────────────────────────────────────────────
// Applica opzioni di formattazione a un Range
// ─────────────────────────────────────────────────────────────
function ap_(range, o) {
  if (!o) return;
  range.setFontFamily("Quicksand");
  if (o.bg   !== undefined) range.setBackground(o.bg);
  if (o.fg   !== undefined && o.fg) range.setFontColor(o.fg);
  if (o.bold !== undefined) range.setFontWeight(o.bold ? "bold" : "normal");
  if (o.sz   !== undefined) range.setFontSize(o.sz);
  if (o.ha   !== undefined) range.setHorizontalAlignment(o.ha);
  if (o.va   !== undefined) range.setVerticalAlignment(o.va);
  if (o.wrap !== undefined) range.setWrap(o.wrap);
  if (o.clip) range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  if (o.nf   !== undefined && o.nf) range.setNumberFormat(o.nf);
  if (o.borderTop)
    range.setBorder(true, null, null, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  if (o.borderBottom)
    range.setBorder(null, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
}

// ─────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────

// Clamp data al giorno valido del mese (es. 31/04 → 30/04)
function clampMonthEnd_(date) {
  if (!date) return null;
  const last = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
  if (date.getDate() > last) return new Date(date.getFullYear(), date.getMonth(), last);
  return date;
}

// Parse dd/mm/yyyy o dd-mm-yyyy
function parseDate_(str) {
  if (!str) return null;
  const p = str.split(/[\/\-]/);
  if (p.length !== 3) return null;
  const g = parseInt(p[0]), m = parseInt(p[1]), a = parseInt(p[2]);
  if (!g || !m || !a || a < 2000) return null;
  const d = new Date(a, m - 1, 1);
  d.setDate(g);
  return isNaN(d.getTime()) ? null : d;
}

// Formatta Date come dd/MM/yyyy (fuso Tokyo)
function fmtDate_(date) {
  if (!date) return "—";
  return Utilities.formatDate(date, "Asia/Tokyo", "dd/MM/yyyy");
}

// Cerca sottocartella Drive per nome; la crea se non esiste
function getOrCreateSubfolder_(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(name);
}

// "Mario Rossi" → "Rossi Mario" (solo per visualizzazione)
function cognomeNome_(fullName) {
  if (!fullName) return fullName;
  const parts = fullName.trim().split(/\s+/);
  if (parts.length < 2) return fullName;
  return parts[parts.length - 1] + " " + parts.slice(0, parts.length - 1).join(" ");
}
// ─────────────────────────────────────────────────────────────
// STATI
// ─────────────────────────────────────────────────────────────

// Stati che annullano il link: esclusi da utilizzo piano e ordini attivi,
// resettano i check sul foglio client.
const STATI_ANNULLATI = ["CANCELLATO", "RESO", "SOLD OUT", "SMARRITO"];

// UTILIZZO PIANO: un link consuma il piano nel momento in cui viene inviato
// (pagina "Piani e Tariffe"). Contano quindi tutti i link della finestra
// del piano con uno stato valorizzato, compresi NUOVO, esclusi solo
// gli stati annullati (STATI_ANNULLATI).

// ─────────────────────────────────────────────────────────────
// Helpers PIANI — normalizzazione, prompt e menu a tendina
// (piani e massimali definiti in PIANI, file costanti)
// ─────────────────────────────────────────────────────────────

// Riporta qualsiasi valore alla chiave piano (es. " basic " → "BASIC")
function normalizzaPiano_(valore) {
  return (valore || "").toString().trim().toUpperCase();
}

// Elenco piani per i prompt (es. "Trial / Starter / Basic / Pro")
function pianiPromptStr_() {
  return PIANI_LISTA.map(k => PIANI[k].label).join(" / ");
}

// Nome piano da mostrare al cliente (es. "BASIC" → "Basic").
// Valori non riconosciuti vengono restituiti invariati.
function pianoLabel_(valore) {
  const key = normalizzaPiano_(valore);
  return PIANI[key] ? PIANI[key].label : (valore || "");
}

// Menu a tendina dei piani nella colonna PIANO della Dashboard.
//
// Se il menu esiste già NON viene mai riscritto: Apps Script non può
// impostare i colori delle opzioni (chip) del menu a tendina, quindi
// riscriverlo cancellerebbe i colori impostati a mano.
// Lo script si limita a:
//   • crearlo solo se nella colonna non c'è ancora nessun menu;
//   • segnalare nei log se le opzioni non coincidono con PIANI_LISTA,
//     così da aggiornarle a mano (データ → データの入力規則).
function applyPianoValidation_(dashboard) {
  const numRows = dashboard.getMaxRows() - 1;
  if (numRows < 1) return;

  const range    = dashboard.getRange(2, D_COL.PIANO, numRows, 1);
  const esistente = dashboard.getRange(2, D_COL.PIANO).getDataValidation();

  if (!esistente) {
    range.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(PIANI_LISTA, true)
        .setAllowInvalid(false)
        .build()
    );
    Logger.log("Menu a tendina PIANO creato: " + PIANI_LISTA.join(", "));
    return;
  }

  // Menu già presente: solo verifica, nessuna modifica
  if (esistente.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    const valori = (esistente.getCriteriaValues()[0] || []).map(v => normalizzaPiano_(v));
    const mancanti = PIANI_LISTA.filter(p => !valori.includes(p));
    const extra    = valori.filter(v => v && !PIANI_LISTA.includes(v));
    if (mancanti.length || extra.length) {
      Logger.log("⚠️ Menu a tendina PIANO da aggiornare a mano — " +
                 (mancanti.length ? "aggiungere: " + mancanti.join(", ") + ". " : "") +
                 (extra.length    ? "rimuovere: "  + extra.join(", ")    + "."  : ""));
    }
  } else {
    Logger.log("⚠️ La colonna PIANO ha un menu non a elenco: verificare le opzioni (" +
               PIANI_LISTA.join(", ") + ").");
  }
}

// ─────────────────────────────────────────────────────────────
// Colori della colonna PIANO in Dashboard (formattazione condizionale)
//
// La Dashboard originale colorava la colonna PIANO con regole di
// formattazione condizionale "il testo è esattamente Trial / Basic /
// Core / Pro". Con i nuovi nomi (STARTER, ecc.) le regole vecchie non
// coprivano più tutti i piani e le celle restavano senza colore.
//
// Questa funzione, chiamata a ogni "Aggiorna Dashboard":
//   • rimuove SOLO le regole che agiscono esclusivamente sulla
//     colonna PIANO (le altre regole del foglio restano intatte);
//   • ricrea una regola per ogni piano di PIANI con i colori bg/fg,
//     sull'intera colonna da riga 2 all'ultima riga del foglio;
//   • confronta i valori senza distinzione maiuscole/minuscole
//     ("STARTER", "Starter" e " starter " ricevono lo stesso colore).
// I colori si cambiano solo in PIANI (file costanti).
// ─────────────────────────────────────────────────────────────
function applyPianoFormatting_(dashboard) {
  const numRows = dashboard.getMaxRows() - 1;
  if (numRows < 1) return;

  const colLetter  = String.fromCharCode(64 + D_COL.PIANO);   // "C"
  const pianoRange = dashboard.getRange(2, D_COL.PIANO, numRows, 1);

  const altreRegole = dashboard.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().every(r =>
      r.getColumn() === D_COL.PIANO && r.getNumColumns() === 1
    )
  );

  const regolePiani = PIANI_LISTA.map(key =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=UPPER(TRIM($' + colLetter + '2))="' + key + '"')
      .setBackground(PIANI[key].bg)
      .setFontColor(PIANI[key].fg)
      .setRanges([pianoRange])
      .build()
  );

  dashboard.setConditionalFormatRules([...altreRegole, ...regolePiani]);
}

// ─────────────────────────────────────────────────────────────
// Helpers PropertiesService — SheetID e JarID
// ─────────────────────────────────────────────────────────────
function saveClientSheetId_(clientName, sheetId) {
  PropertiesService.getScriptProperties().setProperty('client_' + clientName, sheetId);
}
function getClientSheetId_(clientName) {
  const raw = PropertiesService.getScriptProperties().getProperty('client_' + clientName);
  if (!raw) return null;
  const m = raw.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : raw;
}
function deleteClientSheetId_(clientName) {
  PropertiesService.getScriptProperties().deleteProperty('client_' + clientName);
}
function saveClientJarId_(clientName, jarId) {
  PropertiesService.getScriptProperties().setProperty('jar_' + clientName, jarId);
}
function getClientJarId_(clientName) {
  return PropertiesService.getScriptProperties().getProperty('jar_' + clientName);
}
function deleteClientJarId_(clientName) {
  PropertiesService.getScriptProperties().deleteProperty('jar_' + clientName);
}

// ─────────────────────────────────────────────────────────────
// Helpers PropertiesService — dati anagrafici cliente
// ─────────────────────────────────────────────────────────────
function saveClientPiva_(clientName, piva) {
  PropertiesService.getScriptProperties().setProperty('piva_' + clientName, piva);
}
function getClientPiva_(clientName) {
  return PropertiesService.getScriptProperties().getProperty('piva_' + clientName) || "";
}
function deleteClientPiva_(clientName) {
  PropertiesService.getScriptProperties().deleteProperty('piva_' + clientName);
}
function saveClientIndirizzo_(clientName, indirizzo) {
  PropertiesService.getScriptProperties().setProperty('indirizzo_' + clientName, indirizzo);
}
function getClientIndirizzo_(clientName) {
  return PropertiesService.getScriptProperties().getProperty('indirizzo_' + clientName) || "";
}
function deleteClientIndirizzo_(clientName) {
  PropertiesService.getScriptProperties().deleteProperty('indirizzo_' + clientName);
}
function saveClientEmail_(clientName, email) {
  PropertiesService.getScriptProperties().setProperty('email_' + clientName, email);
}
function getClientEmail_(clientName) {
  return PropertiesService.getScriptProperties().getProperty('email_' + clientName) || "";
}
function deleteClientEmail_(clientName) {
  PropertiesService.getScriptProperties().deleteProperty('email_' + clientName);
}

// ─────────────────────────────────────────────────────────────
// Helpers PropertiesService — data inizio piano
// ─────────────────────────────────────────────────────────────
function savePianoStart_(clientName, dateString) {
  // dateString formato ISO: "YYYY-MM-DD"
  PropertiesService.getScriptProperties().setProperty('pianoStart_' + clientName, dateString);
}
function getPianoStart_(clientName) {
  return PropertiesService.getScriptProperties().getProperty('pianoStart_' + clientName) || null;
}
function deletePianoStart_(clientName) {
  PropertiesService.getScriptProperties().deleteProperty('pianoStart_' + clientName);
}

// ─────────────────────────────────────────────────────────────
// Calcola data scadenza (dataInizio + 30 giorni)
// Restituisce un oggetto Date o null
// ─────────────────────────────────────────────────────────────
function calcolaScadenza_(clientName) {
  const startStr = getPianoStart_(clientName);
  if (!startStr) return null;
  const start = new Date(startStr);
  if (isNaN(start.getTime())) return null;
  const scadenza = new Date(start);
  scadenza.setDate(scadenza.getDate() + 30);
  return scadenza;
}

// ─────────────────────────────────────────────────────────────
// Formatta la scadenza come stringa leggibile
// ─────────────────────────────────────────────────────────────
function formatScadenza_(scadenza) {
  if (!scadenza) return "—";
  const dd = String(scadenza.getDate()).padStart(2, "0");
  const mm = String(scadenza.getMonth() + 1).padStart(2, "0");
  const yyyy = scadenza.getFullYear();
  return dd + "/" + mm + "/" + yyyy;
}

// ─────────────────────────────────────────────────────────────
// Restituisce i giorni rimanenti alla scadenza (negativo = scaduto)
// ─────────────────────────────────────────────────────────────
function giorniAllaScadenza_(scadenza) {
  if (!scadenza) return null;
  const oggi = new Date();
  oggi.setHours(0, 0, 0, 0);
  const sc = new Date(scadenza);
  sc.setHours(0, 0, 0, 0);
  return Math.round((sc - oggi) / (1000 * 60 * 60 * 24));
}

// ─────────────────────────────────────────────────────────────
// Applica colore alla cella SCADENZA in Dashboard
//   > 7 giorni  → verde
//   3-7 giorni  → arancio
//   < 3 giorni  → rosso
//   scaduto     → grigio/barrato
//   non impostata → nessun colore
// ─────────────────────────────────────────────────────────────
function coloraScadenza_(cell, giorni) {
  if (giorni === null) {
    cell.setBackground("#FFFFFF").setFontColor("#AAAAAA");
    return;
  }
  if (giorni > 7) {
    cell.setBackground("#D9EAD3").setFontColor("#1E5128"); // verde
  } else if (giorni >= 3) {
    cell.setBackground("#FFF2CC").setFontColor("#7D6608"); // arancio
  } else if (giorni >= 0) {
    cell.setBackground("#F4CCCC").setFontColor("#C0392A"); // rosso
  } else {
    cell.setBackground("#EFEFEF").setFontColor("#999999"); // grigio — scaduto
  }
}

// ─────────────────────────────────────────────────────────────
// Menu personalizzato
// ─────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ DEKKA Admin")
    .addItem("Aggiungi nuovo cliente", "addNewClient")
    .addItem("Rimuovi cliente", "removeClient")
    .addSeparator()
    .addItem("Aggiorna Dashboard", "refreshDashboard_")
    .addItem("Aggiorna solo Saldi (forza push)", "refreshSaldiManuale")
    .addItem("Applica formattazione a tutti i clienti", "applyFormattingToAllClients_")
    .addSeparator()
    .addItem("Imposta data inizio piano", "impostaPianoStart")
    .addItem("Termina piano in anticipo", "terminaPianoAnticipato")
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("🛡️ Copertura Valore")
      .addItem("Aggiungi sezione CV ai client esistenti", "setupCVTuttiIClient")
      .addItem("Migra schede gestionale (aggiungi col. H e I)", "migrateShedeGestionale")
      .addItem("Migra schede mese client esistenti", "migrateShedeMeseClienti")
      .addSeparator()
      .addItem("⚠️ Rollback CV — rimuovi da client specifico", "rollbackCVCliente"))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("🔧 Configurazione Wise")
      .addItem("Imposta API Token", "setWiseApiToken")
      .addItem("DEBUG: Mostra tutti i balance", "debugAllWiseBalances")
      .addItem("Lista Jar disponibili", "listAvailableJars")
      .addItem("Mappa Jar a cliente esistente", "mapJarToExistingClient"))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("⏰ Aggiornamento Automatico")
      .addItem("✅ Attiva aggiornamento automatico", "installAutoRefreshTrigger")
      .addItem("❌ Disattiva aggiornamento automatico", "removeAutoRefreshTrigger")
      .addItem("📋 Visualizza trigger attivi", "viewActiveTriggers"))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("🗑️ Pulizia Automatica")
      .addItem("✅ Attiva pulizia automatica", "enableCleanupTrigger")
      .addItem("❌ Disattiva pulizia automatica", "removeCleanupTrigger")
      .addItem("📋 Visualizza stato pulizia", "viewCleanupTrigger"))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("🔗 Controllo Doppioni")
      .addItem("✅ Attiva trigger riallineamento (una tantum)", "setupTriggerRiallineamento")
      .addItem("🕐 Backfill storico link (una tantum)", "backfillRegistroStorico"))
    .addSeparator()
    .addItem("Genera Rendiconto Cliente", "generaRendiconto")
    .addItem("Modifica dati anagrafici cliente", "modificaDatiAnagrafici")
    .addSeparator()
    .addItem("Migra schede mese → STATO (una tantum)", "migraSchemeMeseVersoSTATO_")
    .addToUi();
}

// ─────────────────────────────────────────────────────────────
// Imposta la data di inizio piano per un cliente
// Voce di menu: "Imposta data inizio piano"
// ─────────────────────────────────────────────────────────────
function impostaPianoStart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const clientSheets = ss.getSheets()
    .filter(s => s.getName() !== "Dashboard")
    .map(s => s.getName());

  if (!clientSheets.length) { ui.alert("Nessun cliente disponibile."); return; }

  const clientResp = ui.prompt(
    "Imposta data inizio piano",
    "Clienti disponibili:\n• " + clientSheets.join("\n• ") + "\n\nNome del cliente:",
    ui.ButtonSet.OK_CANCEL
  );
  if (clientResp.getSelectedButton() !== ui.Button.OK) return;
  const clientName = clientResp.getResponseText().trim();
  if (!clientSheets.includes(clientName)) { ui.alert("Cliente non trovato."); return; }

  const attuale = getPianoStart_(clientName);
  const attualeStr = attuale
    ? "Attuale: " + formatScadenza_(calcolaScadenza_(clientName)) + " (inizio: " + attuale + ")\n\n"
    : "Nessuna data impostata.\n\n";

  const dateResp = ui.prompt(
    "Data inizio piano — " + clientName,
    attualeStr +
    "Inserisci la data di inizio del piano attivo nel formato GG/MM/AAAA\n" +
    "(es. 14/06/2026 per un piano che scade il 14/07/2026):",
    ui.ButtonSet.OK_CANCEL
  );
  if (dateResp.getSelectedButton() !== ui.Button.OK) return;

  const raw = dateResp.getResponseText().trim();
  const parts = raw.split("/");
  if (parts.length !== 3) { ui.alert("Formato non valido. Usa GG/MM/AAAA."); return; }

  const [dd, mm, yyyy] = parts.map(Number);
  if (isNaN(dd) || isNaN(mm) || isNaN(yyyy) || mm < 1 || mm > 12 || dd < 1 || dd > 31) {
    ui.alert("Data non valida."); return;
  }

  // Salva in formato ISO per robustezza
  const isoStr = yyyy + "-" + String(mm).padStart(2, "0") + "-" + String(dd).padStart(2, "0");
  savePianoStart_(clientName, isoStr);

  const scadenza = calcolaScadenza_(clientName);
  refreshDashboard_();

  ui.alert(
    "✅ Data inizio piano impostata per \"" + clientName + "\"\n\n" +
    "Inizio: " + raw + "\n" +
    "Scadenza: " + formatScadenza_(scadenza)
  );
}

// ─────────────────────────────────────────────────────────────
// Termina anticipatamente il piano/pass attivo di un cliente,
// con possibilità facoltativa di avviarne subito uno nuovo.
// Se non si avvia un nuovo piano, il cliente resta nello stato
// "nessun piano attivo": Dashboard e MANGIA LINK mostrano i
// valori di default già previsti dal sistema (piano vuoto,
// massimale "—", utilizzo "0 / —", scadenza "—").
// Voce di menu: "Termina piano in anticipo"
// ─────────────────────────────────────────────────────────────
function terminaPianoAnticipato() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const ui        = SpreadsheetApp.getUi();
  const dashboard = ss.getSheetByName("Dashboard");

  const clientSheets = ss.getSheets()
    .filter(s => s.getName() !== "Dashboard")
    .map(s => s.getName());

  if (!clientSheets.length) { ui.alert("Nessun cliente disponibile."); return; }

  const clientResp = ui.prompt(
    "Termina piano in anticipo",
    "Clienti disponibili:\n• " + clientSheets.join("\n• ") + "\n\nNome del cliente:",
    ui.ButtonSet.OK_CANCEL
  );
  if (clientResp.getSelectedButton() !== ui.Button.OK) return;
  const clientName = clientResp.getResponseText().trim();
  if (!clientSheets.includes(clientName)) { ui.alert("Cliente non trovato."); return; }

  const pianoStartStr   = getPianoStart_(clientName);
  const scadenzaAttuale = calcolaScadenza_(clientName);
  const giorniAttuali   = giorniAllaScadenza_(scadenzaAttuale);

  if (!pianoStartStr) {
    ui.alert("Il cliente \"" + clientName + "\" non ha attualmente una data di inizio piano impostata: non c'è nulla da terminare.");
    return;
  }

  const rawStart = pianoStartStr.split("-").reverse().join("/");
  const infoPianoAttuale =
    "Piano attivo per \"" + clientName + "\":\n" +
    "Inizio: " + rawStart + "\n" +
    "Scadenza naturale: " + formatScadenza_(scadenzaAttuale) +
    (giorniAttuali !== null
      ? "  (" + (giorniAttuali >= 0 ? giorniAttuali + " giorni rimanenti" : "scaduto") + ")"
      : "") +
    "\n\nVuoi terminare questo piano in anticipo, con effetto immediato?";

  const confirm = ui.alert("Conferma terminazione anticipata", infoPianoAttuale, ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  // ── 1. Termina il piano attivo ──────────────────────────────
  deletePianoStart_(clientName);

  let dashRow = null;
  const dashLastRow = dashboard.getLastRow();
  if (dashLastRow >= 2) {
    const nomi = dashboard.getRange(2, D_COL.CLIENTE, dashLastRow - 1, 1).getValues();
    for (let i = 0; i < nomi.length; i++) {
      if (nomi[i][0] === clientName) { dashRow = i + 2; break; }
    }
  }
  if (dashRow) {
    dashboard.getRange(dashRow, D_COL.PIANO).setValue("");
  }

  // ── 2. Chiede se avviare subito un nuovo piano (facoltativo) ─
  const vuoleNuovoPiano = ui.alert(
    "Nuovo piano",
    "Il piano precedente è stato terminato.\n\nVuoi avviare subito un nuovo piano per \"" + clientName + "\"?",
    ui.ButtonSet.YES_NO
  );

  if (vuoleNuovoPiano !== ui.Button.YES) {
    // Nessun nuovo piano: il cliente resta senza piano attivo,
    // con i valori di default (nessun massimale, scadenza "—").
    refreshDashboard_();
    ui.alert(
      "✅ Piano terminato in anticipo per \"" + clientName + "\".\n\n" +
      "Nessun nuovo piano avviato: il cliente risulta ora senza piano attivo (valori di default)."
    );
    return;
  }

  // ── 3. Nuovo piano (facoltativo, scelto Sì) ─────────────────
  const pianoResp = ui.prompt(
    "Nuovo piano — " + clientName,
    "Inserisci il nuovo piano (" + pianiPromptStr_() + "):",
    ui.ButtonSet.OK_CANCEL
  );
  if (pianoResp.getSelectedButton() !== ui.Button.OK) {
    refreshDashboard_();
    ui.alert(
      "✅ Piano precedente terminato per \"" + clientName + "\".\n\n" +
      "Nessun nuovo piano avviato: il cliente risulta ora senza piano attivo (valori di default)."
    );
    return;
  }
  const nuovoPiano = normalizzaPiano_(pianoResp.getResponseText());
  if (!PIANI_LISTA.includes(nuovoPiano)) {
    ui.alert(
      "Piano non valido. Nessun nuovo piano avviato — \"" + clientName + "\" risulta senza piano attivo."
    );
    refreshDashboard_();
    return;
  }

  const dataResp = ui.prompt(
    "Data inizio nuovo piano",
    "Inserisci la data di inizio nel formato GG/MM/AAAA\n(lascia vuoto per usare la data odierna):",
    ui.ButtonSet.OK_CANCEL
  );
  if (dataResp.getSelectedButton() !== ui.Button.OK) {
    refreshDashboard_();
    ui.alert(
      "✅ Piano precedente terminato per \"" + clientName + "\".\n\n" +
      "Nessun nuovo piano avviato: il cliente risulta ora senza piano attivo (valori di default)."
    );
    return;
  }

  const rawDate = dataResp.getResponseText().trim();
  let nuovoStartIso;

  if (!rawDate) {
    const oggi = new Date();
    nuovoStartIso = oggi.getFullYear() + "-" + String(oggi.getMonth() + 1).padStart(2, "0") + "-" + String(oggi.getDate()).padStart(2, "0");
  } else {
    const parts = rawDate.split("/");
    if (parts.length !== 3) {
      ui.alert("Formato data non valido. Nessun nuovo piano avviato — il cliente risulta senza piano attivo.");
      refreshDashboard_();
      return;
    }
    const [dd, mm, yyyy] = parts.map(Number);
    if (isNaN(dd) || isNaN(mm) || isNaN(yyyy) || mm < 1 || mm > 12 || dd < 1 || dd > 31) {
      ui.alert("Data non valida. Nessun nuovo piano avviato — il cliente risulta senza piano attivo.");
      refreshDashboard_();
      return;
    }
    nuovoStartIso = yyyy + "-" + String(mm).padStart(2, "0") + "-" + String(dd).padStart(2, "0");
  }

  savePianoStart_(clientName, nuovoStartIso);
  if (dashRow) dashboard.getRange(dashRow, D_COL.PIANO).setValue(nuovoPiano);

  refreshDashboard_();

  const nuovaScadenza = calcolaScadenza_(clientName);
  ui.alert(
    "✅ Piano precedente terminato in anticipo.\n\n" +
    "Nuovo piano avviato per \"" + clientName + "\": " + pianoLabel_(nuovoPiano) +
    " (massimale " + PIANI_MAX[nuovoPiano] + " link)\n" +
    "Inizio: " + nuovoStartIso.split("-").reverse().join("/") + "\n" +
    "Scadenza: " + formatScadenza_(nuovaScadenza)
  );
}

// ─────────────────────────────────────────────────────────────
// Trigger onEdit installabile
// ─────────────────────────────────────────────────────────────
function onEditInstallable(e) {
  const sheet     = e.range.getSheet();
  const sheetName = sheet.getName();
  const startCol  = e.range.getColumn();
  const endCol    = e.range.getColumn() + e.range.getNumColumns() - 1;
  const startRow  = e.range.getRow();
  const endRow    = e.range.getRow() + e.range.getNumRows() - 1;

  if (sheetName === "Dashboard") {
    if (startCol <= D_COL.PIANO && endCol >= D_COL.PIANO) {
      refreshDashboard_();
    }
    return;
  }

  if (endRow < 3) return;
  const dataStartRow = Math.max(startRow, 3);

  const coinvolgeData       = startCol <= G_COL.DATA_RICEZIONE && endCol >= G_COL.DATA_RICEZIONE;
  const coinvolgeStato      = startCol <= G_COL.STATO          && endCol >= G_COL.STATO;
  const coinvolgeSpedizione = startCol <= G_COL.SPEDIZIONE     && endCol >= G_COL.SPEDIZIONE;
  const coinvolgeQta        = startCol <= G_COL.QTA            && endCol >= G_COL.QTA;
  const coinvolgeImporti    = startCol <= G_COL.IMPORTI        && endCol >= G_COL.IMPORTI;

  if (!coinvolgeData && !coinvolgeStato && !coinvolgeSpedizione && !coinvolgeQta && !coinvolgeImporti) return;

  // Valore precedente della cella: Google lo fornisce solo per modifiche di singola cella
  const unaSolaCella = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1;
  const valorePrecedenteImporti = unaSolaCella ? e.oldValue : undefined;

  let dashboardDaAggiornare = false;

  for (let row = dataStartRow; row <= endRow; row++) {
    if (coinvolgeData) {
      const dataRic = sheet.getRange(row, G_COL.DATA_RICEZIONE).getValue();
      const link    = sheet.getRange(row, G_COL.LINK).getValue();
      const price   = sheet.getRange(row, G_COL.IMPORTI).getValue();
      if (dataRic && link) {
        sendNewOrderNotification_(sheetName, link, price, dataRic);
      }
    }
    if (coinvolgeStato) {
      const nuovoStato = sheet.getRange(row, G_COL.STATO).getValue();
      if (nuovoStato) {
        handleStatoChange_(sheet, row, nuovoStato);
        dashboardDaAggiornare = true;
      }
    }
    if (coinvolgeSpedizione) {
      const valoreSpedizione = sheet.getRange(row, G_COL.SPEDIZIONE).getValue();
      handleSpedizioneChange_(sheet, row, valoreSpedizione);
      dashboardDaAggiornare = true;
    }
    if (coinvolgeQta) {
      const valoreQta = sheet.getRange(row, G_COL.QTA).getValue();
      handleQtaChange_(sheet, row, valoreQta);
      dashboardDaAggiornare = true;
    }
    if (coinvolgeImporti) {
      const valoreImporti = sheet.getRange(row, G_COL.IMPORTI).getValue();
      handleImportiChange_(sheet, row, valoreImporti, valorePrecedenteImporti);
    }
  }

  if (dashboardDaAggiornare) refreshDashboard_();
}

// ─────────────────────────────────────────────────────────────
// Trigger onChange
// ─────────────────────────────────────────────────────────────
function onChangeInstallable(e) {
  if (!e || !e.changeType) return;
  if (e.changeType === "REMOVE_SHEET") refreshDashboard_();
}

// ─────────────────────────────────────────────────────────────
// Ricava la riga sul client cercando il link
// ─────────────────────────────────────────────────────────────
function findClientRow_(clientSheet, link) {
  const lastRow = clientSheet.getLastRow();
  if (lastRow < 3) return null;
  const links = clientSheet.getRange(3, C_COL.LINK, lastRow - 2, 1).getValues();
  for (let i = 0; i < links.length; i++) {
    if (links[i][0] === link) return i + 3;
  }
  return null;
}

// ─────────────────────────────────────────────────────────────
// Verifica se il link è una pagina di catalogo SNKRDUNK
// (es. https://snkrdunk.com/apparels/881423): per questi link
// il cliente inserisce il prezzo base e l'amministratore
// aggiunge poi le spese di spedizione.
// ─────────────────────────────────────────────────────────────
function isSnkrdunkCatalogLink_(link) {
  return /^https:\/\/snkrdunk\.com\/[a-z]+\/\d+$/i.test(link);
}

// ─────────────────────────────────────────────────────────────
// Gestisce cambio STATO
// ─────────────────────────────────────────────────────────────
function handleStatoChange_(sheet, row, nuovoStato) {
  if (!nuovoStato) return;

  const clientName = sheet.getName();
  const clientId   = getClientSheetId_(clientName);
  const link       = sheet.getRange(row, G_COL.LINK).getValue();
  const dataRic    = sheet.getRange(row, G_COL.DATA_RICEZIONE).getValue();

  if (!clientId || !link || !dataRic) return;

  try {
    const clientSS    = SpreadsheetApp.openById(clientId);
    const meseSheet   = MESI_IT[new Date(dataRic).getMonth()];
    const clientSheet = clientSS.getSheetByName(meseSheet);
    if (!clientSheet) { Logger.log("Scheda mese non trovata: " + meseSheet); return; }

    const clientRow = findClientRow_(clientSheet, link);
    if (!clientRow) { Logger.log("Link non trovato sul client: " + link); return; }

    clientSheet.getRange(clientRow, C_COL.STATO).setValue(nuovoStato);
    if (STATI_ANNULLATI.includes(nuovoStato)) {
      clientSheet.getRange(clientRow, C_COL.SPEDIZIONE).setValue("");
    }
    clientSheet.getRange(clientRow, C_COL.LINK).clearNote();

    Logger.log("Stato aggiornato sul client: " + nuovoStato + " — riga " + clientRow);

  } catch (err) {
    Logger.log("Errore aggiornamento client: " + err.message);
  }
}

// ─────────────────────────────────────────────────────────────
// Gestisce inserimento data spedizione
// ─────────────────────────────────────────────────────────────
function handleSpedizioneChange_(sheet, row, dataSpedizione) {
  Logger.log("handleSpedizioneChange_ chiamata — row: " + row + " valore: " + dataSpedizione);

  const clientName   = sheet.getName();
  const clientId     = getClientSheetId_(clientName);
  const link         = sheet.getRange(row, G_COL.LINK).getValue();
  const dataRic      = sheet.getRange(row, G_COL.DATA_RICEZIONE).getValue();
  const statoPrecKey = 'statoPrec_' + clientName + '_' + row;

  if (!clientId || !link || !dataRic) { Logger.log("Uscita anticipata — dati mancanti"); return; }

  try {
    const clientSS    = SpreadsheetApp.openById(clientId);
    const meseSheet   = MESI_IT[new Date(dataRic).getMonth()];
    const clientSheet = clientSS.getSheetByName(meseSheet);
    if (!clientSheet) return;

    const clientRow = findClientRow_(clientSheet, link);
    if (!clientRow) { Logger.log("Link non trovato sul client: " + link); return; }

    if (!dataSpedizione || dataSpedizione === "") {
      const statoPrecedente = PropertiesService.getScriptProperties()
                                .getProperty(statoPrecKey) || "NUOVO";
      sheet.getRange(row, G_COL.STATO).setValue(statoPrecedente);
      PropertiesService.getScriptProperties().deleteProperty(statoPrecKey);
      clientSheet.getRange(clientRow, C_COL.SPEDIZIONE).setValue("");
      clientSheet.getRange(clientRow, C_COL.STATO).setValue(statoPrecedente);
      Logger.log("Data rimossa — stato ripristinato a: " + statoPrecedente);
    } else {
      const statoAttuale = sheet.getRange(row, G_COL.STATO).getValue();
      if (statoAttuale !== "SPEDITO") {
        PropertiesService.getScriptProperties().setProperty(statoPrecKey, statoAttuale);
        sheet.getRange(row, G_COL.STATO).setValue("SPEDITO");
        clientSheet.getRange(clientRow, C_COL.STATO).setValue("SPEDITO");
      }
      const dataFormattata = Utilities.formatDate(new Date(dataSpedizione), "Asia/Tokyo", "dd/MM/yyyy");
      clientSheet.getRange(clientRow, C_COL.SPEDIZIONE).setValue(dataFormattata);
      Logger.log("Data spedizione copiata — stato salvato: " + statoAttuale);
    }

  } catch (err) {
    Logger.log("Errore copia data spedizione: " + err.message);
  }
}

// ─────────────────────────────────────────────────────────────
// Gestisce cambio Q.TA
// ─────────────────────────────────────────────────────────────
function handleQtaChange_(sheet, row, valoreQta) {
  const clientName = sheet.getName();
  const clientId   = getClientSheetId_(clientName);
  const link       = sheet.getRange(row, G_COL.LINK).getValue();
  const dataRic    = sheet.getRange(row, G_COL.DATA_RICEZIONE).getValue();
  if (!clientId || !link || !dataRic) return;

  try {
    const clientSS    = SpreadsheetApp.openById(clientId);
    const meseSheet   = MESI_IT[new Date(dataRic).getMonth()];
    const clientSheet = clientSS.getSheetByName(meseSheet);
    if (!clientSheet) return;
    const clientRow = findClientRow_(clientSheet, link);
    if (!clientRow) return;
    clientSheet.getRange(clientRow, C_COL.QTA).setValue(
      valoreQta !== "" && valoreQta !== null ? valoreQta : ""
    );
  } catch (err) {
    Logger.log("Errore sincronizzazione Q.TA: " + err.message);
  }
}

// ─────────────────────────────────────────────────────────────
// Converte un valore di cella in numero (se possibile).
// Restituisce: numero, stringa (testo non numerico, es. "N/D")
// oppure null se vuoto.
// ─────────────────────────────────────────────────────────────
function parseImporto_(v) {
  if (v === "" || v === null || v === undefined) return null;
  if (typeof v === "number") return v;
  const n = Number(String(v).replace(/[¥,\s]/g, ""));
  return isNaN(n) ? String(v) : n;
}

// ─────────────────────────────────────────────────────────────
// Gestisce cambio IMPORTI
//
// Quando l'importo viene modificato a mano nel gestionale
// (tipicamente dall'amministratore, per aggiungere le spese di
// spedizione ai link di catalogo SNKRDUNK), il nuovo valore viene
// riportato sul client del cliente. Sulla cella del client e su
// quella del gestionale viene lasciata una nota con data/ora e con
// l'importo originale prima della modifica, così la variazione è
// visibile e tracciata da entrambe le parti.
//
// L'importo originale è letto da e.oldValue (modifica di singola
// cella); per modifiche multiple (es. incolla di più celle) si usa
// il valore che la cella del client conteneva prima dell'aggiornamento.
//
// Se l'importo viene cancellato, viene svuotata anche la cella sul
// client e vengono eliminate le note residue su entrambe le celle.
// ─────────────────────────────────────────────────────────────
function handleImportiChange_(sheet, row, valoreImporti, valorePrecedente) {
  const clientName = sheet.getName();
  const clientId   = getClientSheetId_(clientName);
  const link       = sheet.getRange(row, G_COL.LINK).getValue();
  const dataRic    = sheet.getRange(row, G_COL.DATA_RICEZIONE).getValue();
  if (!clientId || !link || !dataRic) return;

  try {
    const clientSS    = SpreadsheetApp.openById(clientId);
    const meseSheet   = MESI_IT[new Date(dataRic).getMonth()];
    const clientSheet = clientSS.getSheetByName(meseSheet);
    if (!clientSheet) return;
    const clientRow = findClientRow_(clientSheet, link);
    if (!clientRow) return;

    const hasValue   = valoreImporti !== "" && valoreImporti !== null;
    const clientCell = clientSheet.getRange(clientRow, C_COL.IMPORTI);

    // Importo originale: prima e.oldValue, altrimenti il valore attuale sul client
    // (da leggere PRIMA di sovrascriverlo)
    const valoreClientPrima = clientCell.getValue();
    const originale = parseImporto_(
      valorePrecedente !== undefined ? valorePrecedente : valoreClientPrima
    );

    clientCell.setValue(hasValue ? valoreImporti : "");

    if (!hasValue) {
      // Importo cancellato: elimina anche le note residue
      clientCell.clearNote();
      sheet.getRange(row, G_COL.IMPORTI).clearNote();
    }

    // Comunica la modifica al cliente e lascia traccia sul gestionale
    if (hasValue && typeof valoreImporti === "number") {
      const now       = Utilities.formatDate(new Date(), "Asia/Tokyo", "dd/MM/yyyy HH:mm");
      const inclSpese = isSnkrdunkCatalogLink_(link) ? ", spese di spedizione incluse" : "";
      let nota        = "Importo aggiornato dall'amministratore il " + now + inclSpese + ".";

      // Aggiunge l'importo originale solo se esiste ed è diverso da quello nuovo
      if (originale !== null && originale !== valoreImporti) {
        const originaleStr = (typeof originale === "number") ? formatSaldoJPY_(originale) : originale;
        nota += "\nImporto originale: " + originaleStr + ".";
      }

      clientCell.setNote(nota);
      sheet.getRange(row, G_COL.IMPORTI).setNote(nota);
    }

    Logger.log("Importo sincronizzato sul client — " + clientName + " — riga " + clientRow + " — valore: " + valoreImporti);

  } catch (err) {
    Logger.log("Errore sincronizzazione IMPORTI: " + err.message);
  }
}


// ═════════════════════════════════════════════════════════════
// COPERTURA VALORE — funzioni di setup e migrazione
// ═════════════════════════════════════════════════════════════

// ── 1. Aggiunge sezione CV al MANGIA LINK di tutti i client ──
function setupCVTuttiIClient() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const clientSheets = ss.getSheets().filter(s => s.getName() !== "Dashboard");

  let ok = 0, skip = 0, err = 0;

  for (const sheet of clientSheets) {
    const clientName = sheet.getName();
    const clientId   = getClientSheetId_(clientName);
    if (!clientId) { Logger.log("SheetID mancante per: " + clientName); err++; continue; }

    try {
      const clientSS = SpreadsheetApp.openById(clientId);
      setupCVSectionOnSS_(clientSS, clientName);
      ok++;
    } catch (e) {
      Logger.log("Errore setup CV per " + clientName + ": " + e.message);
      err++;
    }
  }

  ui.alert(
    "✅ Setup CV completato\n\n" +
    "Configurati: " + ok + "\n" +
    "Saltati (già configurati): " + skip + "\n" +
    (err > 0 ? "Errori: " + err + " (controlla i log)" : "")
  );
}

function setupCVSectionOnSS_(ss, clientName) {
  const sheet = ss.getSheetByName("MANGIA LINK");
  if (!sheet) { Logger.log("MANGIA LINK non trovata per: " + clientName); return; }

  // Controlla marker
  if (sheet.getRange("A3").getValue() === "CV_SETUP_OK") {
    Logger.log("CV già configurata per: " + clientName + " — skip");
    return;
  }

  const CV_VALORI_LOCAL = ["—", "Base", "Plus", "Max"];
  const CV_CELLE_LOCAL  = { PIANO: "B3", NOTA: "B4", LABEL_NOTA: "C4" };

  sheet.insertRowsAfter(2, 3);

  sheet.getRange("A3").setValue("CV_SETUP_OK");

  const cellPiano = sheet.getRange(CV_CELLE_LOCAL.PIANO);
  cellPiano.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(CV_VALORI_LOCAL, false)
      .setAllowInvalid(false).build()
  );
  cellPiano.setValue("—");
  cellPiano.setBackground("#F8F9FA");
  cellPiano.setFontFamily("Quicksand");
  cellPiano.setFontWeight("bold");
  cellPiano.setHorizontalAlignment("center");
  cellPiano.setNote("Seleziona il piano Copertura Valore prima di inviare il link.\nLascia '—' se non vuoi attivare il servizio per questo link.");

  sheet.getRange("C3").setValue("Copertura Valore (opzionale)");
  sheet.getRange("C3").setFontColor("#888888").setFontFamily("Quicksand").setFontSize(10).setVerticalAlignment("middle");

  const cellNota = sheet.getRange(CV_CELLE_LOCAL.NOTA);
  cellNota.setBackground("#F8F9FA").setFontFamily("Quicksand").setHorizontalAlignment("left");
  cellNota.setNote("Indica quale carta vuoi coprire.\nLascia vuoto se il link contiene una sola carta.");

  sheet.getRange(CV_CELLE_LOCAL.LABEL_NOTA)
    .setValue("Carta da coprire — lascia vuoto se il link ha una sola carta")
    .setFontColor("#888888").setFontFamily("Quicksand").setFontSize(10).setVerticalAlignment("middle");

  sheet.getRange("B5:U5").setBackground("#EEEEEE");
  sheet.setColumnWidth(1, 2);
  sheet.setColumnWidth(2, 160);

  Logger.log("setupCVSection_ completata per: " + clientName);
}

// ── 2. Migra schede gestionale (aggiunge H e I a schede cliente) ──
function migrateShedeGestionale() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const clientSheets = ss.getSheets().filter(s => s.getName() !== "Dashboard");

  let ok = 0, skip = 0;

  for (const sheet of clientSheets) {
    const headerH = sheet.getRange(1, G_COL.COV_PIANO).getValue();
    if (headerH === "COV") { skip++; continue; }

    // Merge righe 1-2 per H e I come le altre colonne
    sheet.getRange(1, G_COL.COV_PIANO, 2, 1).merge();
    sheet.getRange(1, G_COL.COV_NOTA,  2, 1).merge();

    sheet.getRange(1, G_COL.COV_PIANO, 2, 2)
      .setBackground("#6D9EEB")
      .setFontColor("#FFFFFF")
      .setFontFamily("Quicksand")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    sheet.getRange(1, G_COL.COV_PIANO).setValue("COV");
    sheet.getRange(1, G_COL.COV_NOTA).setValue("NOTA CV");

    sheet.setColumnWidth(G_COL.COV_PIANO, 70);
    sheet.setColumnWidth(G_COL.COV_NOTA,  140);

    sheet.getRange(3, G_COL.COV_PIANO, 998, 1)
      .setHorizontalAlignment("center")
      .setFontFamily("Quicksand")
      .setVerticalAlignment("middle");
    sheet.getRange(3, G_COL.COV_NOTA, 998, 1)
      .setHorizontalAlignment("left")
      .setFontFamily("Quicksand")
      .setVerticalAlignment("middle")
      .setFontColor("#555555");

    // Formattazione condizionale colonna H
    applyGestionaleCovFormatting_(sheet);

    ok++;
    Logger.log("Colonne CV aggiunte alla scheda gestionale: " + sheet.getName());
  }

  ui.alert(
    "✅ Migrazione gestionale completata\n\n" +
    "Schede aggiornate: " + ok + "\n" +
    "Già configurate (skip): " + skip
  );
}

function applyGestionaleCovFormatting_(sheet) {
  const covRange    = sheet.getRange(3, G_COL.COV_PIANO, 998, 1);
  const existingRules = sheet.getConditionalFormatRules();
  const newRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Base")
      .setBackground("#D9EAD3").setFontColor("#1E5128")
      .setRanges([covRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Plus")
      .setBackground("#FFF2CC").setFontColor("#7D6608")
      .setRanges([covRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Max")
      .setBackground("#F4CCCC").setFontColor("#7B0000")
      .setRanges([covRange]).build(),
  ];
  sheet.setConditionalFormatRules([...existingRules, ...newRules]);
}

// ── 3. Migra schede mese di tutti i client esistenti ──────────
function migrateShedeMeseClienti() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const clientSheets = ss.getSheets().filter(s => s.getName() !== "Dashboard");

  let clientOk = 0, clientErr = 0;

  for (const sheet of clientSheets) {
    const clientName = sheet.getName();
    const clientId   = getClientSheetId_(clientName);
    if (!clientId) { Logger.log("SheetID mancante per: " + clientName); clientErr++; continue; }

    try {
      const clientSS = SpreadsheetApp.openById(clientId);
      migrateShedeMeseOnSS_(clientSS, clientName);
      clientOk++;
    } catch (e) {
      Logger.log("Errore migrazione schede mese per " + clientName + ": " + e.message);
      clientErr++;
    }
  }

  ui.alert(
    "✅ Migrazione schede mese completata\n\n" +
    "Client aggiornati: " + clientOk + "\n" +
    (clientErr > 0 ? "Errori: " + clientErr + " (controlla i log)" : "")
  );
}

function migrateShedeMeseOnSS_(ss, clientName) {
  const CV_COLORI_LOCAL = {
    "Base": { bg: "#D9EAD3", fg: "#1E5128" },
    "Plus": { bg: "#FFF2CC", fg: "#7D6608" },
    "Max":  { bg: "#F4CCCC", fg: "#7B0000" },
  };

  for (const mese of MESI_IT) {
    const sheet = ss.getSheetByName(mese);
    if (!sheet) continue;

    const headerH = sheet.getRange(1, 8).getValue();
    if (headerH === "COV") {
      Logger.log("CV già presente in " + mese + " per " + clientName + " — skip");
      continue;
    }

    sheet.getRange(1, 8, 2, 1).merge();
    sheet.getRange(1, 9, 2, 1).merge();

    sheet.getRange(1, 8, 2, 2)
      .setBackground("#6D9EEB")
      .setFontColor("#FFFFFF")
      .setFontFamily("Quicksand")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    sheet.getRange(1, 8).setValue("COV");
    sheet.getRange(1, 9).setValue("NOTA CV");

    sheet.setColumnWidth(8, 70);
    sheet.setColumnWidth(9, 120);

    sheet.getRange(3, 8, 998, 1)
      .setHorizontalAlignment("center")
      .setFontFamily("Quicksand")
      .setVerticalAlignment("middle");
    sheet.getRange(3, 9, 998, 1)
      .setHorizontalAlignment("left")
      .setFontFamily("Quicksand")
      .setVerticalAlignment("middle")
      .setFontColor("#555555");

    const covRange      = sheet.getRange("H3:H1000");
    const existingRules = sheet.getConditionalFormatRules();
    const newRules = Object.entries(CV_COLORI_LOCAL).map(([val, col]) =>
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(val)
        .setBackground(col.bg).setFontColor(col.fg)
        .setRanges([covRange]).build()
    );
    sheet.setConditionalFormatRules([...existingRules, ...newRules]);

    Logger.log("Colonne CV aggiunte a " + mese + " per " + clientName);
  }
}

// ── ROLLBACK — rimuove CV da un client specifico (reset completo) ──
function rollbackCVCliente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const clientSheets = ss.getSheets()
    .filter(s => s.getName() !== "Dashboard")
    .map(s => s.getName());

  if (!clientSheets.length) { ui.alert("Nessun cliente disponibile."); return; }

  const clientResp = ui.prompt(
    "Rollback Copertura Valore",
    "Seleziona il cliente su cui rimuovere completamente il servizio CV:\n\n• " +
    clientSheets.join("\n• ") + "\n\nNome del cliente:",
    ui.ButtonSet.OK_CANCEL
  );
  if (clientResp.getSelectedButton() !== ui.Button.OK) return;
  const clientName = clientResp.getResponseText().trim();
  if (!clientSheets.includes(clientName)) { ui.alert("Cliente non trovato."); return; }

  const clientId = getClientSheetId_(clientName);
  if (!clientId) {
    ui.alert("SheetID non trovato per \"" + clientName + "\".\nImpossibile procedere.");
    return;
  }

  const confirm = ui.alert(
    "Conferma rollback",
    "Stai per rimuovere il servizio CV dal client di \"" + clientName + "\".\n\n" +
    "Verranno eliminati:\n" +
    "• Sezione CV dal MANGIA LINK (righe B3/B4/B5)\n" +
    "• Colonne H e I da tutte le schede mese del client\n" +
    "• Colonne H e I dalla scheda cliente nel gestionale\n\n" +
    "I dati CV storici andranno persi.\n\nContinuare?",
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  const errori = [];

  // 1. Rollback MANGIA LINK e schede mese sul client
  try {
    const clientSS = SpreadsheetApp.openById(clientId);
    rollbackCVOnClientSS_(clientSS, clientName, errori);
  } catch (e) {
    errori.push("Errore accesso al client: " + e.message);
  }

  // 2. Rollback scheda cliente nel gestionale
  try {
    const gestSheet = ss.getSheetByName(clientName);
    if (gestSheet) {
      rollbackCVGestionale_(gestSheet, clientName, errori);
    } else {
      errori.push("Scheda gestionale non trovata per: " + clientName);
    }
  } catch (e) {
    errori.push("Errore rollback gestionale: " + e.message);
  }

  const esito = errori.length === 0
    ? "✅ Rollback completato per \"" + clientName + "\".\nIl client è tornato alle condizioni precedenti."
    : "Rollback completato con errori:\n\n" + errori.join("\n");
  ui.alert(esito);
  Logger.log("Rollback CV per " + clientName + (errori.length ? " — errori: " + errori.join(" | ") : " — OK"));
}

function rollbackCVOnClientSS_(ss, clientName, errori) {

  // ── MANGIA LINK ───────────────────────────────────────────
  const mangiaLink = ss.getSheetByName("MANGIA LINK");
  if (!mangiaLink) {
    errori.push("MANGIA LINK non trovata per: " + clientName);
  } else {
    const marker = mangiaLink.getRange("A3").getValue();
    if (marker !== "CV_SETUP_OK") {
      Logger.log("CV non configurata su MANGIA LINK per " + clientName + " — skip righe");
    } else {
      // Elimina le 3 righe in ordine decrescente per non spostare gli indici
      mangiaLink.deleteRow(5);
      mangiaLink.deleteRow(4);
      mangiaLink.deleteRow(3);
      // Pulisce eventuale residuo nella cella A3 (ora tornata la riga originale)
      mangiaLink.getRange("A3").clearContent();
      Logger.log("Righe CV rimosse da MANGIA LINK per: " + clientName);
    }
  }

  // ── SCHEDE MESE ───────────────────────────────────────────
  for (const mese of MESI_IT) {
    const sheet = ss.getSheetByName(mese);
    if (!sheet) continue;

    const headerH = sheet.getRange(1, 8).getValue();
    if (headerH !== "COV") {
      Logger.log("Colonne CV non presenti in " + mese + " — skip");
      continue;
    }

    try {
      // Rimuove solo le regole condizionali che coinvolgono la colonna H
      const rulesFiltrate = sheet.getConditionalFormatRules().filter(rule =>
        !rule.getRanges().some(r => r.getColumn() === 8)
      );
      sheet.setConditionalFormatRules(rulesFiltrate);

      // Elimina in ordine inverso per non spostare gli indici
      sheet.deleteColumn(9); // I — NOTA CV
      sheet.deleteColumn(8); // H — COV
      Logger.log("Colonne CV rimosse da " + mese + " per " + clientName);
    } catch (e) {
      errori.push("Errore rimozione colonne in " + mese + ": " + e.message);
    }
  }
}

function rollbackCVGestionale_(sheet, clientName, errori) {
  const headerH = sheet.getRange(1, G_COL.COV_PIANO).getValue();
  if (headerH !== "COV") {
    Logger.log("Colonne CV non presenti nel gestionale per " + clientName + " — skip");
    return;
  }

  try {
    // Rimuove solo le regole condizionali che coinvolgono COV_PIANO (col. 8)
    const rulesFiltrate = sheet.getConditionalFormatRules().filter(rule =>
      !rule.getRanges().some(r => r.getColumn() === G_COL.COV_PIANO)
    );
    sheet.setConditionalFormatRules(rulesFiltrate);

    // Elimina in ordine inverso
    sheet.deleteColumn(G_COL.COV_NOTA);  // I = 9
    sheet.deleteColumn(G_COL.COV_PIANO); // H = 8
    Logger.log("Colonne CV rimosse dalla scheda gestionale di: " + clientName);
  } catch (e) {
    errori.push("Errore rimozione colonne gestionale: " + e.message);
  }
}


// ═════════════════════════════════════════════════════════════
// AGGIUNGI NUOVO CLIENTE
// ═════════════════════════════════════════════════════════════

function addNewClient() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const nameResponse = ui.prompt("Nuovo cliente", "Inserisci il nome del cliente (es. Mario Rossi):", ui.ButtonSet.OK_CANCEL);
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  const clientName = nameResponse.getResponseText().trim();
  if (!clientName) { ui.alert("Nome non valido."); return; }
  if (ss.getSheetByName(clientName)) { ui.alert("Esiste già una scheda con questo nome."); return; }

  const pianoResponse = ui.prompt("Piano DEKKA", "Inserisci il piano (" + pianiPromptStr_() + "):", ui.ButtonSet.OK_CANCEL);
  if (pianoResponse.getSelectedButton() !== ui.Button.OK) return;
  const piano = normalizzaPiano_(pianoResponse.getResponseText());
  if (!PIANI_LISTA.includes(piano)) { ui.alert("Piano non valido. Piani disponibili: " + pianiPromptStr_() + "."); return; }

  // Data inizio piano
  const startResp = ui.prompt(
    "Data inizio piano",
    "Inserisci la data di inizio del piano nel formato GG/MM/AAAA\n(lascia vuoto per impostarla in seguito):",
    ui.ButtonSet.OK_CANCEL
  );
  if (startResp.getSelectedButton() !== ui.Button.OK) return;
  const startRaw = startResp.getResponseText().trim();
  let pianoStartIso = null;
  if (startRaw) {
    const parts = startRaw.split("/");
    if (parts.length === 3) {
      const [dd, mm, yyyy] = parts.map(Number);
      if (!isNaN(dd) && !isNaN(mm) && !isNaN(yyyy)) {
        pianoStartIso = yyyy + "-" + String(mm).padStart(2, "0") + "-" + String(dd).padStart(2, "0");
      }
    }
    if (!pianoStartIso) { ui.alert("Formato data non valido — continuo senza impostare la data."); }
  }

  const idResponse = ui.prompt("SheetID client cliente", "Incolla lo Spreadsheet ID del foglio DEKKA del cliente:", ui.ButtonSet.OK_CANCEL);
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  const clientSheetId = idResponse.getResponseText().trim();
  if (!clientSheetId) { ui.alert("SheetID non valido."); return; }

  const pivaResponse = ui.prompt("P.IVA / Codice Fiscale", "Inserisci la P.IVA o Codice Fiscale (lascia vuoto per saltare):", ui.ButtonSet.OK_CANCEL);
  if (pivaResponse.getSelectedButton() !== ui.Button.OK) return;
  const piva = pivaResponse.getResponseText().trim();

  const indirizzoResponse = ui.prompt("Indirizzo di spedizione", "Inserisci l'indirizzo di spedizione (lascia vuoto per saltare):", ui.ButtonSet.OK_CANCEL);
  if (indirizzoResponse.getSelectedButton() !== ui.Button.OK) return;
  const indirizzo = indirizzoResponse.getResponseText().trim();

  const emailResponse = ui.prompt("Email cliente", "Inserisci l'email del cliente (lascia vuoto per saltare):", ui.ButtonSet.OK_CANCEL);
  if (emailResponse.getSelectedButton() !== ui.Button.OK) return;
  const email = emailResponse.getResponseText().trim();

  ui.alert("Recupero Jar Wise in corso...");
  const jars = getWiseJars_();
  let selectedJarId = null;

  if (jars.length > 0) {
    const jarResponse = ui.prompt(
      "Seleziona Jar Wise",
      "Jar disponibili:\n• " + jars.map(j => j.name + " (ID: " + j.id + ")").join("\n• ") +
      "\n\nIncolla l'ID del Jar (oppure lascia vuoto):",
      ui.ButtonSet.OK_CANCEL
    );
    if (jarResponse.getSelectedButton() === ui.Button.OK) {
      selectedJarId = jarResponse.getResponseText().trim();
      if (selectedJarId && !jars.find(j => j.id.toString() === selectedJarId)) {
        ui.alert("⚠️ ID Jar non valido. Continuo senza mappare il Jar.");
        selectedJarId = null;
      }
    }
  }

  const newSheet = ss.insertSheet(clientName);
  ss.moveActiveSheet(ss.getSheets().length);
  setupClientSheet_(newSheet);

  saveClientSheetId_(clientName, clientSheetId);
  if (selectedJarId)  saveClientJarId_(clientName, selectedJarId);
  if (piva)           saveClientPiva_(clientName, piva);
  if (indirizzo)      saveClientIndirizzo_(clientName, indirizzo);
  if (email)          saveClientEmail_(clientName, email);
  if (pianoStartIso)  savePianoStart_(clientName, pianoStartIso);

  SpreadsheetApp.flush();

  const dashboard  = ss.getSheetByName("Dashboard");
  const newDashRow = dashboard.getLastRow() + 1;
  dashboard.getRange(newDashRow, D_COL.CLIENTE).setValue(clientName);
  dashboard.getRange(newDashRow, D_COL.PIANO).setValue(piano);
  if (piva)      dashboard.getRange(newDashRow, D_COL.PIVA).setValue(piva);
  if (indirizzo) dashboard.getRange(newDashRow, D_COL.INDIRIZZO).setValue(indirizzo);
  if (email)     dashboard.getRange(newDashRow, D_COL.EMAIL).setValue(email);

  refreshDashboard_();
  ss.setActiveSheet(newSheet);

  const jarMsg   = selectedJarId  ? "\n✅ Jar Wise mappato."            : "\n⚠️ Jar Wise non mappato.";
  const startMsg = pianoStartIso  ? "\n📅 Scadenza: " + formatScadenza_(calcolaScadenza_(clientName)) : "\n⚠️ Data inizio piano non impostata.";
  ui.alert("✅ Cliente \"" + clientName + "\" aggiunto con piano " + pianoLabel_(piano) +
           " (massimale " + PIANI_MAX[piano] + " link)." + jarMsg + startMsg);
}

// ─────────────────────────────────────────────────────────────
// Rimuove scheda cliente
// ─────────────────────────────────────────────────────────────
function removeClient() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const clientSheets = ss.getSheets().filter(s => s.getName() !== "Dashboard").map(s => s.getName());
  if (!clientSheets.length) { ui.alert("Nessun cliente da rimuovere."); return; }

  const response = ui.prompt(
    "Rimuovi cliente",
    "Inserisci il nome esatto del cliente:\n\n• " + clientSheets.join("\n• "),
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const clientName = response.getResponseText().trim();
  if (!clientName) { ui.alert("Nome non valido."); return; }

  const clientSheet = ss.getSheetByName(clientName);
  if (!clientSheet) { ui.alert("Nessuna scheda trovata per \"" + clientName + "\"."); return; }

  const confirm = ui.alert(
    "Conferma eliminazione",
    "Eliminare il cliente \"" + clientName + "\"?\n\nNon è reversibile.",
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  ss.deleteSheet(clientSheet);
  deleteClientSheetId_(clientName);
  deleteClientJarId_(clientName);
  deleteSaldoPrecedente_(clientName);
  deleteClientPiva_(clientName);
  deleteClientIndirizzo_(clientName);
  deleteClientEmail_(clientName);
  deletePianoStart_(clientName);

  SpreadsheetApp.flush();
  refreshDashboard_();
  ui.alert("✅ Cliente \"" + clientName + "\" rimosso.");
}

// ─────────────────────────────────────────────────────────────
// Setup scheda cliente (nuovi clienti — include H e I)
// ─────────────────────────────────────────────────────────────
function setupClientSheet_(sheet) {
  const colWidths = { 1:30, 2:90, 3:90, 4:90, 5:240, 6:45, 7:90, 8:70, 9:140 };
  for (const [col, width] of Object.entries(colWidths)) {
    sheet.setColumnWidth(Number(col), width);
  }

  for (const col of [2, 3, 4, 5, 6, 7, 8, 9]) {
    sheet.getRange(1, col, 2, 1).merge();
  }

  const headers = {
    2:"DATA\nRICEZIONE", 3:"STATO", 4:"SPEDIZIONE",
    5:"LINK", 6:"Q.TA", 7:"IMPORTI", 8:"COV", 9:"NOTA CV"
  };

  sheet.getRange(1, 2, 2, 8)
    .setBackground("#6D9EEB").setFontColor("#FFFFFF").setFontFamily("Quicksand")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

  for (const [col, label] of Object.entries(headers)) {
    sheet.getRange(1, Number(col)).setValue(label);
  }

  const numRows = 998, numCols = 8;
  const backgrounds = [];
  for (let r = 0; r < numRows; r++) {
    backgrounds.push(new Array(numCols).fill(r % 2 === 0 ? "#FFFFFF" : "#FFFFF0"));
  }
  sheet.getRange(3, 2, numRows, numCols).setBackgrounds(backgrounds);

  sheet.getRange("B3:I1000")
    .setFontWeight("normal").setFontFamily("Quicksand").setVerticalAlignment("middle");
  sheet.getRange("B3:B1000").setNumberFormat("yyyy/mm/dd").setHorizontalAlignment("center").setFontColor("#000000");
  sheet.getRange("C3:C1000").setHorizontalAlignment("center").setFontColor("#000000");
  sheet.getRange("D3:D1000").setHorizontalAlignment("center").setFontColor("#000000");
  sheet.getRange("E3:E1000").setHorizontalAlignment("left").setFontColor("#1155CC").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange("F3:F1000").setHorizontalAlignment("center").setFontColor("#000000");
  sheet.getRange("G3:G1000").setHorizontalAlignment("right").setNumberFormat("#,##0").setFontColor("#000000");
  sheet.getRange("H3:H1000").setHorizontalAlignment("center").setFontColor("#000000");
  sheet.getRange("I3:I1000").setHorizontalAlignment("left").setFontColor("#555555");

  sheet.getRange("D3:D1000").setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build()
  );
  sheet.getRange("C3:C1000").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["NUOVO","PAGATO","RICEVUTO","SPEDITO","CANCELLATO","RESO","SOLD OUT","SMARRITO"], false)
      .setAllowInvalid(false).build()
  );

  const statoRange = sheet.getRange("C3:C1000");
  const covRange   = sheet.getRange("H3:H1000");

  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("NUOVO")
      .setBackground("#efefef").setFontColor("#555555").setRanges([statoRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("PAGATO")
      .setBackground("#fff2cc").setFontColor("#7d6608").setRanges([statoRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("RICEVUTO")
      .setBackground("#cfe2ff").setFontColor("#0d47a1").setRanges([statoRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("SPEDITO")
      .setBackground("#d9ead3").setFontColor("#1e5128").setRanges([statoRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("CANCELLATO")
      .setBackground("#f4cccc").setFontColor("#7b0000").setRanges([statoRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("RESO")
      .setBackground("#ead1dc").setFontColor("#4a1030").setRanges([statoRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("SOLD OUT")
      .setBackground("#f0e6ff").setFontColor("#4a0072").setRanges([statoRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("SMARRITO")
      .setBackground("#f5e6d3").setFontColor("#7d4400").setRanges([statoRange]).build(),
    // CV
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Base")
      .setBackground("#D9EAD3").setFontColor("#1E5128").setRanges([covRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Plus")
      .setBackground("#FFF2CC").setFontColor("#7D6608").setRanges([covRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Max")
      .setBackground("#F4CCCC").setFontColor("#7B0000").setRanges([covRange]).build(),
  ]);

  sheet.setFrozenRows(2);
  Logger.log("setupClientSheet_ completato: " + sheet.getName());
}

// ─────────────────────────────────────────────────────────────
// Applica formattazione a tutte le schede cliente esistenti
// ─────────────────────────────────────────────────────────────
function applyFormattingToAllClients_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientSheets = ss.getSheets().filter(s => s.getName() !== "Dashboard");
  let count = 0;

  for (const sheet of clientSheets) {
    sheet.getRange("B3:I1000")
      .setFontWeight("normal").setFontFamily("Quicksand").setVerticalAlignment("middle");
    sheet.getRange("B3:B1000").setNumberFormat("yyyy/mm/dd").setHorizontalAlignment("center").setFontColor("#000000");
    sheet.getRange("C3:C1000").setHorizontalAlignment("center").setFontColor("#000000");
    sheet.getRange("D3:D1000").setHorizontalAlignment("center").setFontColor("#000000");
    sheet.getRange("E3:E1000").setHorizontalAlignment("left").setFontColor("#1155CC").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sheet.getRange("F3:F1000").setHorizontalAlignment("center").setFontColor("#000000");
    sheet.getRange("G3:G1000").setHorizontalAlignment("right").setNumberFormat("#,##0").setFontColor("#000000");
    sheet.getRange("H3:H1000").setHorizontalAlignment("center").setFontColor("#000000");
    sheet.getRange("I3:I1000").setHorizontalAlignment("left").setFontColor("#555555");

    sheet.getRange("C3:C1000").setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["NUOVO","PAGATO","RICEVUTO","SPEDITO","CANCELLATO","RESO","SOLD OUT","SMARRITO"], false)
        .setAllowInvalid(false).build()
    );

    const statoRange = sheet.getRange("C3:C1000");
    const covRange   = sheet.getRange("H3:H1000");
    sheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("NUOVO")
        .setBackground("#efefef").setFontColor("#555555").setRanges([statoRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("PAGATO")
        .setBackground("#fff2cc").setFontColor("#7d6608").setRanges([statoRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("RICEVUTO")
        .setBackground("#cfe2ff").setFontColor("#0d47a1").setRanges([statoRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("SPEDITO")
        .setBackground("#d9ead3").setFontColor("#1e5128").setRanges([statoRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("CANCELLATO")
        .setBackground("#f4cccc").setFontColor("#7b0000").setRanges([statoRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("RESO")
        .setBackground("#ead1dc").setFontColor("#4a1030").setRanges([statoRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("SOLD OUT")
        .setBackground("#f0e6ff").setFontColor("#4a0072").setRanges([statoRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("SMARRITO")
        .setBackground("#f5e6d3").setFontColor("#7d4400").setRanges([statoRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Base")
        .setBackground("#D9EAD3").setFontColor("#1E5128").setRanges([covRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Plus")
        .setBackground("#FFF2CC").setFontColor("#7D6608").setRanges([covRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Max")
        .setBackground("#F4CCCC").setFontColor("#7B0000").setRanges([covRange]).build(),
    ]);
    count++;
  }

  SpreadsheetApp.getUi().alert("✅ Formattazione aggiornata su " + count + " schede cliente.");
}


// ═════════════════════════════════════════════════════════════
// DASHBOARD
// ═════════════════════════════════════════════════════════════

function formatSaldoJPY_(valore) {
  const n = Math.round(valore);
  return "¥" + n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function setupDashboardHeaders_() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName("Dashboard");
  if (!dashboard) return;

  const headers = {
    [D_COL.CLIENTE]   : "CLIENTE",
    [D_COL.PIANO]     : "PIANO",
    [D_COL.UTILIZZO]  : "UTILIZZO PIANO",
    [D_COL.ATTIVI]    : "ORDINI ATTIVI",
    [D_COL.SALDO]     : "SALDO",
    [D_COL.PIVA]      : "P.IVA / C.F.",
    [D_COL.INDIRIZZO] : "INDIRIZZO SPEDIZIONE",
    [D_COL.EMAIL]     : "EMAIL",
    [D_COL.SCADENZA]  : "SCADENZA PIANO",
  };

  for (const [col, label] of Object.entries(headers)) {
    const cell = dashboard.getRange(1, Number(col));
    if (!cell.getValue()) cell.setValue(label);
  }

  // Imposta larghezza colonna SCADENZA se non già impostata
  dashboard.setColumnWidth(D_COL.SCADENZA, 140);

  // Menu a tendina PIANO allineato ai piani ufficiali (PIANI)
  applyPianoValidation_(dashboard);

  // Colori della colonna PIANO (palette originale, nomi piani attuali)
  applyPianoFormatting_(dashboard);
}

function refreshDashboard_() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName("Dashboard");
  if (!dashboard) return;

  setupDashboardHeaders_();

  const sheets       = ss.getSheets();
  const meseCorrente = MESI_IT[new Date().getMonth()];
  const annoCorrente = new Date().getFullYear();

  const pianiEsistenti = {}, pivaEsistenti = {}, indirizziEsistenti = {}, emailEsistenti = {};
  const dashLastRow = dashboard.getLastRow();

  if (dashLastRow >= 2) {
    const dashValues = dashboard.getRange(
      2, D_COL.CLIENTE, dashLastRow - 1, D_COL.EMAIL - D_COL.CLIENTE + 1
    ).getValues();
    for (const r of dashValues) {
      if (!r[0]) continue;
      const nome = r[0];
      pianiEsistenti[nome]     = r[D_COL.PIANO     - D_COL.CLIENTE] || "";
      pivaEsistenti[nome]      = r[D_COL.PIVA      - D_COL.CLIENTE] || "";
      indirizziEsistenti[nome] = r[D_COL.INDIRIZZO - D_COL.CLIENTE] || "";
      emailEsistenti[nome]     = r[D_COL.EMAIL     - D_COL.CLIENTE] || "";
    }
  }

  const saldiManuali = {};
  if (dashLastRow >= 2) {
    const saldiValues = dashboard.getRange(
      2, D_COL.CLIENTE, dashLastRow - 1, D_COL.SALDO - D_COL.CLIENTE + 1
    ).getValues();
    for (const r of saldiValues) {
      const nome  = r[0];
      const saldo = r[D_COL.SALDO - D_COL.CLIENTE];
      if (nome === "Cristian Bertolini" && saldo !== "" && saldo !== null) {
        saldiManuali[nome] = saldo;
      }
    }
  }

  const clientSheets = sheets.filter(s => s.getName() !== "Dashboard");
  const nuoviDati    = [];

  for (const sheet of clientSheets) {
    const name      = sheet.getName();
    const piano     = normalizzaPiano_(pianiEsistenti[name]);
    const massimale = PIANI_MAX[piano] || null;
    const lastRow   = sheet.getLastRow();

    // Piano compilato ma non presente tra quelli ufficiali (es. vecchio "CORE")
    if (piano && !massimale) {
      Logger.log("⚠️ Piano non riconosciuto per " + name + ": \"" + piano +
                 "\" — selezionare un piano valido in Dashboard (" + pianiPromptStr_() + ").");
    }

    let utilizzo = "0" + (massimale ? " / " + massimale : " / —");
    let ordiniAttivi = 0;

    // Finestra del piano: da pianoStart a pianoStart+30gg.
    // Se non impostata, fallback al mese/anno corrente.
    const pianoStartStr = getPianoStart_(name);
    let finestraInizio  = null;
    let finestraFine    = null;
    if (pianoStartStr) {
      // "YYYY-MM-DD" letto come data locale (new Date("YYYY-MM-DD")
      // la interpreterebbe in UTC, con possibile slittamento di un giorno)
      const p = pianoStartStr.split("-").map(Number);
      if (p.length === 3 && p.every(n => !isNaN(n))) {
        finestraInizio = new Date(p[0], p[1] - 1, p[2], 0, 0, 0, 0);
        finestraFine   = new Date(p[0], p[1] - 1, p[2] + 30, 23, 59, 59, 999);
      }
    }

    if (lastRow >= 3) {
      const rangeValues = sheet.getRange(
        3, G_COL.DATA_RICEZIONE, lastRow - 2, G_COL.QTA - G_COL.DATA_RICEZIONE + 1
      ).getValues();

      // Ordini attivi = data presente, stato non vuoto, non annullato e non spedito
      // (= ancora in lavorazione presso DEKKA)
      const statiNonAttivi = [...STATI_ANNULLATI, "SPEDITO"];
      ordiniAttivi = rangeValues.filter(r =>
        r[0] !== "" && r[0] !== null && r[0] !== false &&
        r[1] !== "" && r[1] !== null && r[1] !== false &&
        !statiNonAttivi.includes(r[1])
      ).length;

      // Utilizzo piano = tutti i link inviati nella finestra del piano
      // attivo (dalla data di attivazione), in qualsiasi stato compreso
      // NUOVO. Esclusi solo gli stati annullati: se un ordine passa a
      // CANCELLATO/RESO/SOLD OUT/SMARRITO sparisce dal contatore.
      // Stessa regola del controllo Trial negli script dei clienti.
      let utilizzoTotale = 0;
      for (const r of rangeValues) {
        if (!r[0] || r[0] === false) continue;                           // B = data
        if (!r[G_COL.LINK - G_COL.DATA_RICEZIONE]) continue;             // E = link
        const stato = (r[G_COL.STATO - G_COL.DATA_RICEZIONE] || "").toString().trim().toUpperCase();
        if (!stato || STATI_ANNULLATI.includes(stato)) continue;

        const d = new Date(r[0]);
        if (isNaN(d.getTime())) {
          Logger.log("⚠️ Data non valida in scheda " + name + ": \"" + r[0] + "\" — riga esclusa dal conteggio");
          continue;
        }
        d.setHours(0, 0, 0, 0);

        if (finestraInizio && finestraFine) {
          // Finestra del piano impostata
          if (d < finestraInizio || d > finestraFine) continue;
        } else {
          // Fallback: mese e anno correnti
          if (MESI_IT[d.getMonth()] !== meseCorrente || d.getFullYear() !== annoCorrente) continue;
        }

        const qta = parseInt(r[G_COL.QTA - G_COL.DATA_RICEZIONE]) || 1;  // F = Q.TA (vuota = 1)
        utilizzoTotale += Math.ceil(qta / 25);
      }
      utilizzo = utilizzoTotale + (massimale ? " / " + massimale : " / —");
    }

    let saldo = null;
    if (name === "Cristian Bertolini") {
      saldo = saldiManuali[name] !== undefined ? saldiManuali[name] : null;
    } else {
      const jarId = getClientJarId_(name);
      if (jarId) {
        const balance = getJarBalance_(jarId);
        saldo = formatSaldoJPY_(balance !== null ? balance : 0);
      } else {
        saldo = "—";
      }
    }

    const piva      = getClientPiva_(name)     || pivaEsistenti[name]      || "";
    const indirizzo = getClientIndirizzo_(name) || indirizziEsistenti[name] || "";
    const email     = getClientEmail_(name)     || emailEsistenti[name]     || "";

    // Scadenza piano
    const scadenza = calcolaScadenza_(name);
    const giorni   = giorniAllaScadenza_(scadenza);
    const scadenzaStr = scadenza ? formatScadenza_(scadenza) + (giorni !== null ? "  (" + (giorni >= 0 ? giorni + "g)" : "scaduto)") : "") : "—";

    nuoviDati.push({
      row: [name, pianiEsistenti[name] || "", utilizzo, ordiniAttivi, saldo, piva, indirizzo, email],
      scadenzaStr,
      giorni,
    });
  }

  if (dashLastRow >= 2) {
    dashboard.getRange(2, D_COL.CLIENTE, dashLastRow - 1, D_COL.SCADENZA - D_COL.CLIENTE + 1).clearContent();
  }

  for (let i = 0; i < nuoviDati.length; i++) {
    const dashRow = i + 2;
    const d = nuoviDati[i].row;
    dashboard.getRange(dashRow, D_COL.CLIENTE).setValue(d[0]);
    if (d[1]) dashboard.getRange(dashRow, D_COL.PIANO).setValue(d[1]);
    dashboard.getRange(dashRow, D_COL.UTILIZZO).setValue(d[2]);
    dashboard.getRange(dashRow, D_COL.ATTIVI).setValue(d[3]);
    if (d[4] !== null) dashboard.getRange(dashRow, D_COL.SALDO).setValue(d[4]);
    if (d[5]) dashboard.getRange(dashRow, D_COL.PIVA).setValue(d[5]);
    if (d[6]) dashboard.getRange(dashRow, D_COL.INDIRIZZO).setValue(d[6]);
    if (d[7]) dashboard.getRange(dashRow, D_COL.EMAIL).setValue(d[7]);

    // Colonna SCADENZA con colore
    const scadenzaCell = dashboard.getRange(dashRow, D_COL.SCADENZA);
    scadenzaCell.setValue(nuoviDati[i].scadenzaStr);
    scadenzaCell.setFontFamily("Quicksand").setHorizontalAlignment("center").setFontSize(10);
    coloraScadenza_(scadenzaCell, nuoviDati[i].giorni);
  }

  const nuovoLastRow         = nuoviDati.length + 1;
  const totaleRowsAggiornato = dashboard.getLastRow();
  if (totaleRowsAggiornato > nuovoLastRow) {
    dashboard.deleteRows(nuovoLastRow + 1, totaleRowsAggiornato - nuovoLastRow);
  }

  // Aggiorna header SCADENZA (potrebbe non esserci se dashboard esistente)
  const headerScad = dashboard.getRange(1, D_COL.SCADENZA);
  if (!headerScad.getValue()) headerScad.setValue("SCADENZA PIANO");

  pushDashboardToClients_(nuoviDati.map(d => d.row));
}

function pushDashboardToClients_(nuoviDati) {
  for (const row of nuoviDati) {
    const clientName = row[0];
    const piano      = pianoLabel_(row[1]);   // es. "BASIC" → "Basic" sul client
    const utilizzo   = row[2];                // es. "12 / 100" — massimale del piano selezionato
    const clientId   = getClientSheetId_(clientName);
    if (!clientId) continue;

    try {
      const clientSS   = SpreadsheetApp.openById(clientId);
      const mangiaLink = clientSS.getSheetByName("MANGIA LINK");
      if (!mangiaLink) continue;

      // Le celle B10/C10 sono ora B13/C13 dopo l'inserimento delle 3 righe CV
      // Usiamo le celle nominate per robustezza, leggendo la posizione dal marker
      const pianoCell    = findPianoCellAfterSetup_(mangiaLink);
      const utilizzoCell = pianoCell ? mangiaLink.getRange(pianoCell.getRow(), pianoCell.getColumn() + 1) : null;

      if (pianoCell)    pianoCell.setValue(piano);
      if (utilizzoCell) utilizzoCell.setValue(utilizzo);

      Logger.log("Dashboard pushata a " + clientName + " — Piano: " + piano + " Utilizzo: " + utilizzo);
    } catch (err) {
      Logger.log("Errore push dashboard a " + clientName + ": " + err.message);
    }

    // Push scadenza piano (data + countdown) — separato dal try sopra
    // per non bloccare l'aggiornamento di piano/utilizzo in caso di errore
    pushScadenzaToClient_(clientName);
  }
}

// Trova la cella PIANO DEKKA nel MANGIA LINK tenendo conto
// che dopo il setup CV le righe sono slittate di 3
function findPianoCellAfterSetup_(mangiaLink) {
  try {
    // Controlla se CV è stato configurato (marker in A3)
    const marker = mangiaLink.getRange("A3").getValue();
    if (marker === "CV_SETUP_OK") {
      // Righe slittate: piano era B10, ora è B13
      return mangiaLink.getRange("B13");
    } else {
      // Setup CV non ancora fatto — posizione originale
      return mangiaLink.getRange("B10");
    }
  } catch (e) {
    Logger.log("Errore findPianoCellAfterSetup_: " + e.message);
    return mangiaLink.getRange("B10");
  }
}


// ═════════════════════════════════════════════════════════════
// SALDI WISE
// ═════════════════════════════════════════════════════════════

// ─────────────────────────────────────────────────────────────
// Aggiorna i saldi Wise in Dashboard e li pusha ai client.
//
// Le righe della Dashboard vengono individuate per NOME cliente
// (mappa nameToRow dalla colonna CLIENTE), mai per posizione: così
// un riordino delle schede non può scrivere il saldo di un cliente
// sulla riga di un altro.
//
// forzaPush: true solo se chiamata da refreshSaldiManuale. Quando la
// funzione è lanciata da un trigger temporale, Apps Script passa come
// primo argomento l'oggetto evento: viene ignorato (nessun push forzato).
// ─────────────────────────────────────────────────────────────
function refreshSaldiOnly_(forzaPush) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName("Dashboard");
  if (!dashboard) return;

  forzaPush = (forzaPush === true);

  // Mappa nome cliente → riga Dashboard
  const nameToRow   = {};
  const dashLastRow = dashboard.getLastRow();
  if (dashLastRow >= 2) {
    const nomi = dashboard.getRange(2, D_COL.CLIENTE, dashLastRow - 1, 1).getValues();
    nomi.forEach((r, idx) => { if (r[0]) nameToRow[r[0]] = idx + 2; });
  }

  const clientSheets = ss.getSheets().filter(s => s.getName() !== "Dashboard");

  for (let i = 0; i < clientSheets.length; i++) {
    const clientName = clientSheets[i].getName();
    const dashRow    = nameToRow[clientName];

    if (!dashRow) {
      Logger.log("Saldi: cliente non presente in Dashboard, saltato — " + clientName +
                 " (eseguire \"Aggiorna Dashboard\")");
      continue;
    }

    if (clientName === "Cristian Bertolini") {
      const saldoManuale = dashboard.getRange(dashRow, D_COL.SALDO).getValue();
      if (saldoManuale !== "" && saldoManuale !== null) {
        const clientId = getClientSheetId_(clientName);
        if (clientId) {
          try {
            const mangiaLink = SpreadsheetApp.openById(clientId).getSheetByName("MANGIA LINK");
            if (mangiaLink) {
              const saldoCell = findSaldoCellAfterSetup_(mangiaLink);
              const dataSaldoCell = findDataSaldoCellAfterSetup_(mangiaLink);
              saldoCell.setValue(saldoManuale);

              const saldoPrecKey = "saldoPrec_CristianBertolini_manuale";
              const saldoAttuale = saldoManuale.toString();
              const saldoPrecStr = PropertiesService.getScriptProperties().getProperty(saldoPrecKey);

              if (saldoAttuale !== saldoPrecStr) {
                const oggi = Utilities.formatDate(new Date(), "Europe/Rome", "dd/MM/yyyy");
                dataSaldoCell.setValue(oggi);
                PropertiesService.getScriptProperties().setProperty(saldoPrecKey, saldoAttuale);
              }
            }
          } catch (err) { Logger.log("Errore push saldo manuale: " + err.message); }
        }
      }
      continue;
    }

    const jarId = getClientJarId_(clientName);
    let saldo = "—", saldoNumerico = 0;
    if (jarId) {
      const balance = getJarBalance_(jarId);
      if (balance !== null && balance !== 0) {
        saldoNumerico = Math.round(balance);
        saldo         = formatSaldoJPY_(saldoNumerico);
      } else {
        saldo = "¥0";
      }
    }

    dashboard.getRange(dashRow, D_COL.SALDO).setValue(saldo);

    if (jarId) {
      const saldoPrecedente = getSaldoPrecedente_(clientName);
      if (saldoNumerico !== saldoPrecedente || forzaPush) {
        const clientId = getClientSheetId_(clientName);
        if (clientId) {
          try {
            const mangiaLink = SpreadsheetApp.openById(clientId).getSheetByName("MANGIA LINK");
            if (mangiaLink) {
              findSaldoCellAfterSetup_(mangiaLink).setValue(saldo);
            }
          } catch (err) { Logger.log("Errore push saldo: " + err.message); }
        }
        saveSaldoPrecedente_(clientName, saldoNumerico);
      }
    }
  }
}

// Localizza la cella saldo H9 (ora H12 se CV configurato)
function findSaldoCellAfterSetup_(mangiaLink) {
  const marker = mangiaLink.getRange("A3").getValue();
  return marker === "CV_SETUP_OK"
    ? mangiaLink.getRange("H12")
    : mangiaLink.getRange("H9");
}

// Localizza la cella data aggiornamento saldo H8 (ora H11 se CV configurato)
function findDataSaldoCellAfterSetup_(mangiaLink) {
  const marker = mangiaLink.getRange("A3").getValue();
  return marker === "CV_SETUP_OK"
    ? mangiaLink.getRange("H11")
    : mangiaLink.getRange("H8");
}

// Localizza la cella label "Scadenza piano" L9 (ora L12 se CV configurato)
function findScadenzaLabelCellAfterSetup_(mangiaLink) {
  const marker = mangiaLink.getRange("A3").getValue();
  return marker === "CV_SETUP_OK"
    ? mangiaLink.getRange("L12")
    : mangiaLink.getRange("L9");
}

// Localizza la cella valore scadenza L10 (ora L13 se CV configurato)
function findScadenzaValoreCellAfterSetup_(mangiaLink) {
  const marker = mangiaLink.getRange("A3").getValue();
  return marker === "CV_SETUP_OK"
    ? mangiaLink.getRange("L13")
    : mangiaLink.getRange("L10");
}

// ─────────────────────────────────────────────────────────────
// Pusha la scadenza piano (data + countdown) al MANGIA LINK
// del cliente. Chiamata da refreshDashboard_ e da impostaPianoStart.
// ─────────────────────────────────────────────────────────────
function pushScadenzaToClient_(clientName) {
  const clientId = getClientSheetId_(clientName);
  if (!clientId) return;

  try {
    const clientSS   = SpreadsheetApp.openById(clientId);
    const mangiaLink = clientSS.getSheetByName("MANGIA LINK");
    if (!mangiaLink) return;

    const scadenza = calcolaScadenza_(clientName);
    const giorni   = giorniAllaScadenza_(scadenza);

    let valoreStr;
    if (!scadenza) {
      valoreStr = "—";
    } else if (giorni === null) {
      valoreStr = formatScadenza_(scadenza);
    } else if (giorni >= 0) {
      valoreStr = formatScadenza_(scadenza) + " · " + giorni + "gg";
    } else {
      valoreStr = formatScadenza_(scadenza) + " · scaduto";
    }

    const labelCell  = findScadenzaLabelCellAfterSetup_(mangiaLink);
    const valoreCell = findScadenzaValoreCellAfterSetup_(mangiaLink);

    labelCell.setValue("Scadenza piano");
    valoreCell.setValue(valoreStr);

    Logger.log("Scadenza pushata a " + clientName + " — " + valoreStr);
  } catch (err) {
    Logger.log("Errore push scadenza a " + clientName + ": " + err.message);
  }
}

function refreshSaldiManuale() {
  refreshSaldiOnly_(true);
  SpreadsheetApp.getUi().alert("✅ Saldi aggiornati e pushati forzatamente a tutti i client!");
}

function saveSaldoPrecedente_(clientName, saldo) {
  PropertiesService.getScriptProperties().setProperty('saldoPrec_' + clientName, saldo.toString());
}
function getSaldoPrecedente_(clientName) {
  const saldo = PropertiesService.getScriptProperties().getProperty('saldoPrec_' + clientName);
  return saldo ? parseFloat(saldo) : null;
}
function deleteSaldoPrecedente_(clientName) {
  PropertiesService.getScriptProperties().deleteProperty('saldoPrec_' + clientName);
}


// ═════════════════════════════════════════════════════════════
// TRIGGER AUTOMATICO SALDI
// ═════════════════════════════════════════════════════════════

function installAutoRefreshTrigger() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Aggiornamento Automatico Saldi",
    "Scegli l'intervallo:\n\n1 = Ogni 15 minuti\n2 = Ogni 30 minuti\n3 = Ogni 1 ora\n" +
    "4 = Ogni 6 ore\n5 = Ogni 12 ore\n6 = Una volta al giorno (ore 9:00)\n\nInserisci il numero (1-6):",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const scelta = response.getResponseText().trim();
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "refreshSaldiOnly_") ScriptApp.deleteTrigger(t);
  }

  let descrizione = "";
  switch(scelta) {
    case "1": ScriptApp.newTrigger("refreshSaldiOnly_").timeBased().everyMinutes(15).create(); descrizione = "ogni 15 minuti"; break;
    case "2": ScriptApp.newTrigger("refreshSaldiOnly_").timeBased().everyMinutes(30).create(); descrizione = "ogni 30 minuti"; break;
    case "3": ScriptApp.newTrigger("refreshSaldiOnly_").timeBased().everyHours(1).create();   descrizione = "ogni 1 ora"; break;
    case "4": ScriptApp.newTrigger("refreshSaldiOnly_").timeBased().everyHours(6).create();   descrizione = "ogni 6 ore"; break;
    case "5": ScriptApp.newTrigger("refreshSaldiOnly_").timeBased().everyHours(12).create();  descrizione = "ogni 12 ore"; break;
    case "6": ScriptApp.newTrigger("refreshSaldiOnly_").timeBased().everyDays(1).atHour(9).create(); descrizione = "ogni giorno alle 9:00"; break;
    default:  ui.alert("⚠️ Scelta non valida."); return;
  }

  ui.alert("✅ Aggiornamento automatico attivato — " + descrizione + ".");
}

function removeAutoRefreshTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = false;
  for (const t of triggers) {
    if (t.getHandlerFunction() === "refreshSaldiOnly_") { ScriptApp.deleteTrigger(t); removed = true; }
  }
  SpreadsheetApp.getUi().alert(removed ? "✅ Aggiornamento automatico disattivato." : "⚠️ Nessun trigger attivo.");
}

function viewActiveTriggers() {
  const ui = SpreadsheetApp.getUi();
  const refreshTriggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "refreshSaldiOnly_");
  if (!refreshTriggers.length) { ui.alert("⚠️ Nessun trigger automatico attivo."); return; }
  let info = "═══ TRIGGER ATTIVI ═══\n\n";
  refreshTriggers.forEach((t, idx) => {
    info += (idx+1) + ". Funzione: refreshSaldiOnly_\n   Tipo: " + t.getEventType() + "\n\n";
  });
  ui.alert(info);
}


// ═════════════════════════════════════════════════════════════
// INSTALLAZIONE TRIGGER
// ═════════════════════════════════════════════════════════════

function installTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "onEditInstallable") ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger("onEditInstallable").forSpreadsheet(ss).onEdit().create();
  Logger.log("Trigger onEdit installato.");
}

function installOnChangeTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "onChangeInstallable") ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger("onChangeInstallable").forSpreadsheet(ss).onChange().create();
  Logger.log("Trigger onChange installato.");
}


// ═════════════════════════════════════════════════════════════
// WISE API
// ═════════════════════════════════════════════════════════════

function setWiseApiToken() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Wise API Token", "Incolla il tuo Wise API Token (read-only):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('WISE_API_TOKEN', response.getResponseText().trim());
    ui.alert("✅ Token salvato.");
  }
}

function getWiseProfileId_() {
  const token = PropertiesService.getScriptProperties().getProperty('WISE_API_TOKEN');
  if (!token) return null;
  try {
    const response = UrlFetchApp.fetch("https://api.wise.com/v1/profiles", {
      method: "get", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true
    });
    const profiles = JSON.parse(response.getContentText());
    for (const profile of profiles) {
      if (profile.type === "business") return profile.id;
    }
    return null;
  } catch (err) { Logger.log("Errore recupero Profile ID: " + err.message); return null; }
}

function getWiseJars_() {
  const token     = PropertiesService.getScriptProperties().getProperty('WISE_API_TOKEN');
  const profileId = getWiseProfileId_();
  if (!token || !profileId) return [];
  try {
    const response = UrlFetchApp.fetch(
      "https://api.wise.com/v4/profiles/" + profileId + "/balances?types=STANDARD,SAVINGS",
      { method: "get", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true }
    );
    const data = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200 || !Array.isArray(data)) return [];
    return data.filter(b => b.type === "SAVINGS" && b.currency === "JPY").map(jar => ({
      id: jar.id, name: jar.name || "Jar senza nome",
      currency: "JPY", amount: { value: jar.amount.value }, type: "SAVINGS"
    }));
  } catch (err) { Logger.log("Errore recupero Jars: " + err.message); return []; }
}

function getJarBalance_(jarId) {
  const token     = PropertiesService.getScriptProperties().getProperty('WISE_API_TOKEN');
  const profileId = getWiseProfileId_();
  if (!token || !profileId || !jarId) return 0;
  try {
    const response = UrlFetchApp.fetch(
      "https://api.wise.com/v4/profiles/" + profileId + "/balances?types=STANDARD,SAVINGS",
      { method: "get", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true }
    );
    const data = JSON.parse(response.getContentText());
    if (!Array.isArray(data)) return 0;
    const jar = data.find(b => b.id.toString() === jarId.toString());
    return (jar && jar.amount) ? jar.amount.value : 0;
  } catch (err) { Logger.log("Errore recupero saldo Jar " + jarId + ": " + err.message); return 0; }
}

function debugAllWiseBalances() {
  const ui        = SpreadsheetApp.getUi();
  const token     = PropertiesService.getScriptProperties().getProperty('WISE_API_TOKEN');
  const profileId = getWiseProfileId_();
  if (!token || !profileId) { ui.alert("Token o Profile ID mancante."); return; }
  try {
    const response   = UrlFetchApp.fetch(
      "https://api.wise.com/v4/profiles/" + profileId + "/balances?types=STANDARD,SAVINGS",
      { method: "get", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true }
    );
    const statusCode = response.getResponseCode();
    const data       = JSON.parse(response.getContentText());
    if (statusCode !== 200) { ui.alert("Errore API (Status " + statusCode + ")"); return; }
    if (!Array.isArray(data)) { ui.alert("Response non è un array."); return; }

    let info = "═══ BALANCE WISE (" + data.length + ") ═══\n\n";
    data.filter(b => b.type === "STANDARD").forEach((b,i) => {
      info += (i+1)+". "+b.currency+" (ID: "+b.id+")\n   Saldo: "+b.amount.value+"\n\n";
    });
    data.filter(b => b.type === "SAVINGS").forEach((j,i) => {
      info += (i+1)+". "+(j.name||"Jar")+"\n   ID: "+j.id+"\n   Saldo: "+j.amount.value+" JPY\n\n";
    });
    ui.alert(info.length > 1800 ? "Output troppo lungo — controlla i log." : info);
  } catch (err) { ui.alert("Errore: " + err.message); }
}

function listAvailableJars() {
  const ui   = SpreadsheetApp.getUi();
  const jars = getWiseJars_();
  if (!jars.length) { ui.alert("Nessun Jar trovato."); return; }
  let jarList = "Jar disponibili:\n\n";
  jars.forEach(j => {
    jarList += "• " + j.name + "\n  ID: " + j.id + "\n  Saldo: ¥" +
               Math.round(j.amount.value).toLocaleString('ja-JP') + "\n\n";
  });
  ui.alert(jarList);
}

function mapJarToExistingClient() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const clientSheets = ss.getSheets().filter(s => s.getName() !== "Dashboard").map(s => s.getName());
  if (!clientSheets.length) { ui.alert("Nessun cliente disponibile."); return; }

  const clientResponse = ui.prompt(
    "Seleziona cliente",
    "Clienti:\n• " + clientSheets.join("\n• ") + "\n\nNome cliente:",
    ui.ButtonSet.OK_CANCEL
  );
  if (clientResponse.getSelectedButton() !== ui.Button.OK) return;
  const clientName = clientResponse.getResponseText().trim();
  if (!clientSheets.includes(clientName)) { ui.alert("Cliente non trovato."); return; }

  const jars = getWiseJars_();
  if (!jars.length) { ui.alert("Nessun Jar trovato."); return; }

  const jarResponse = ui.prompt(
    "Seleziona Jar",
    "Jar:\n• " + jars.map(j => j.name + " (ID: " + j.id + ")").join("\n• ") + "\n\nID Jar:",
    ui.ButtonSet.OK_CANCEL
  );
  if (jarResponse.getSelectedButton() !== ui.Button.OK) return;
  const jarId = jarResponse.getResponseText().trim();
  if (!jars.find(j => j.id.toString() === jarId)) { ui.alert("ID Jar non valido."); return; }

  saveClientJarId_(clientName, jarId);
  ui.alert("✅ Jar mappato a \"" + clientName + "\".");
  refreshDashboard_();
}


// ═════════════════════════════════════════════════════════════
// PULIZIA AUTOMATICA
// ═════════════════════════════════════════════════════════════

function viewCleanupTrigger() {
  const ui = SpreadsheetApp.getUi();
  const cleanupTriggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "cleanOldRecords");
  if (!cleanupTriggers.length) {
    ui.alert("⚠️ Pulizia automatica NON attiva.");
    return;
  }
  ui.alert("═══ PULIZIA AUTOMATICA ATTIVA ═══\n\nFunzione: cleanOldRecords\nFrequenza: 1° di ogni mese alle 3:00\nElimina: Record più vecchi di 3 anni");
}

function removeCleanupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = false;
  for (const t of triggers) {
    if (t.getHandlerFunction() === "cleanOldRecords") { ScriptApp.deleteTrigger(t); removed = true; }
  }
  SpreadsheetApp.getUi().alert(removed ? "✅ Pulizia automatica disattivata." : "⚠️ Nessun trigger di pulizia attivo.");
}

function enableCleanupTrigger() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    "Conferma attivazione",
    "Attivare la pulizia automatica?\n\nIl 1° di ogni mese alle 3:00 verranno eliminati i record più vecchi di 3 anni.\n\n⚠️ Operazione irreversibile.",
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;
  installCleanupTrigger();
  ui.alert("✅ Pulizia automatica attivata.");
}


// ═════════════════════════════════════════════════════════════
// MODIFICA DATI ANAGRAFICI
// ═════════════════════════════════════════════════════════════

function modificaDatiAnagrafici() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const clientSheets = ss.getSheets().filter(s => s.getName() !== "Dashboard").map(s => s.getName());
  if (!clientSheets.length) { ui.alert("Nessun cliente disponibile."); return; }

  const clientResp = ui.prompt(
    "Modifica dati anagrafici",
    "Clienti:\n• " + clientSheets.join("\n• ") + "\n\nNome del cliente:",
    ui.ButtonSet.OK_CANCEL
  );
  if (clientResp.getSelectedButton() !== ui.Button.OK) return;
  const clientName = clientResp.getResponseText().trim();
  if (!clientSheets.includes(clientName)) { ui.alert("Cliente non trovato."); return; }

  const pivaAttuale      = getClientPiva_(clientName);
  const indirizzoAttuale = getClientIndirizzo_(clientName);
  const emailAttuale     = getClientEmail_(clientName);

  const pivaResp = ui.prompt("P.IVA",
    "Attuale: " + (pivaAttuale || "—") + "\n\nNuovo valore (vuoto = non modificare):", ui.ButtonSet.OK_CANCEL);
  if (pivaResp.getSelectedButton() !== ui.Button.OK) return;
  const nuovaPiva = pivaResp.getResponseText().trim();

  const indirizzoResp = ui.prompt("Indirizzo",
    "Attuale: " + (indirizzoAttuale || "—") + "\n\nNuovo valore (vuoto = non modificare):", ui.ButtonSet.OK_CANCEL);
  if (indirizzoResp.getSelectedButton() !== ui.Button.OK) return;
  const nuovoIndirizzo = indirizzoResp.getResponseText().trim();

  const emailResp = ui.prompt("Email",
    "Attuale: " + (emailAttuale || "—") + "\n\nNuovo valore (vuoto = non modificare):", ui.ButtonSet.OK_CANCEL);
  if (emailResp.getSelectedButton() !== ui.Button.OK) return;
  const nuovaEmail = emailResp.getResponseText().trim();

  if (nuovaPiva)      saveClientPiva_(clientName, nuovaPiva);
  if (nuovoIndirizzo) saveClientIndirizzo_(clientName, nuovoIndirizzo);
  if (nuovaEmail)     saveClientEmail_(clientName, nuovaEmail);

  refreshDashboard_();

  ui.alert(
    "✅ Dati aggiornati per \"" + clientName + "\"\n\n" +
    "P.IVA: "     + (nuovaPiva      || pivaAttuale      || "—") + "\n" +
    "Indirizzo: " + (nuovoIndirizzo || indirizzoAttuale || "—") + "\n" +
    "Email: "     + (nuovaEmail     || emailAttuale     || "—")
  );
}


// ═════════════════════════════════════════════════════════════
// NOTIFICA NUOVO ORDINE (chiamata dall'onEdit del gestionale
// quando viene inserita la DATA RICEZIONE)
// ═════════════════════════════════════════════════════════════

function sendNewOrderNotification_(clientName, link, price, dataRic) {
  // Questa funzione è il punto di notifica lato gestionale.
  // Puoi implementarla se vuoi una email anche al momento
  // della presa in carico lato DEKKA.
  Logger.log("sendNewOrderNotification_ chiamata per " + clientName + " — " + link);
}


// ─────────────────────────────────────────────────────────────
// TRIGGER RIALLINEAMENTO DOPPIONI
// Da eseguire UNA SOLA VOLTA dopo il deploy (menu ⚙️ DEKKA Admin
// → Controllo Doppioni → Attiva trigger riallineamento).
// ─────────────────────────────────────────────────────────────
function setupTriggerRiallineamento() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'riallineaDoppioni')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('riallineaDoppioni')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('Trigger riallineaDoppioni installato (ogni 5 min).');
}


// ═════════════════════════════════════════════════════════════
// MIGRAZIONE SCHEDE MESE → COLONNA STATO (una tantum)
//
// Sostituisce le 2 colonne ACQUISTO (B) e RICEZIONE (C) con una
// sola colonna STATO identica a quella del gestionale.
// Elimina la colonna C (RICEZIONE), rinomina B in STATO,
// applica dropdown e formattazione condizionale.
// Il valore STATO viene letto dal gestionale (fonte di verità).
// ═════════════════════════════════════════════════════════════
function migraSchemeMeseVersoSTATO_() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    "Migrazione schede mese → STATO",
    "Questa operazione modifica le schede mese di tutti i clienti:\n\n" +
    "• Rimuove le colonne ACQUISTO e RICEZIONE (checkbox)\n" +
    "• Aggiunge una colonna STATO identica a quella del gestionale\n\n" +
    "⚠️ Eseguire dopo aver aggiornato gli script Mangia Link.\n" +
    "L'operazione è irreversibile. Continuare?",
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  const clientSheets = ss.getSheets().filter(s => s.getName() !== "Dashboard");
  let clientOk = 0, clientErr = 0;

  for (const gestSheet of clientSheets) {
    const clientName = gestSheet.getName();
    const clientId   = getClientSheetId_(clientName);
    if (!clientId) {
      Logger.log("SheetID mancante per: " + clientName);
      clientErr++;
      continue;
    }

    try {
      const clientSS = SpreadsheetApp.openById(clientId);
      _migraMeseClienteSS_(gestSheet, clientSS, clientName);
      clientOk++;
    } catch (e) {
      Logger.log("Errore migrazione " + clientName + ": " + e.message);
      clientErr++;
    }
  }

  ui.alert(
    "✅ Migrazione completata\n\n" +
    "Client aggiornati: " + clientOk + "\n" +
    (clientErr > 0 ? "Errori: " + clientErr + " (controlla i log)" : "")
  );
}

function _migraMeseClienteSS_(gestSheet, clientSS, clientName) {
  const STATI_LISTA = ["NUOVO","PAGATO","RICEVUTO","SPEDITO","CANCELLATO","RESO","SOLD OUT","SMARRITO"];
  const STATO_COLORS = [
    { v: "NUOVO",     bg: "#efefef", fg: "#555555" },
    { v: "PAGATO",    bg: "#fff2cc", fg: "#7d6608" },
    { v: "RICEVUTO",  bg: "#cfe2ff", fg: "#0d47a1" },
    { v: "SPEDITO",   bg: "#d9ead3", fg: "#1e5128" },
    { v: "CANCELLATO",bg: "#f4cccc", fg: "#7b0000" },
    { v: "RESO",      bg: "#ead1dc", fg: "#4a1030" },
    { v: "SOLD OUT",  bg: "#f0e6ff", fg: "#4a0072" },
    { v: "SMARRITO",  bg: "#f5e6d3", fg: "#7d4400" },
  ];
  const CV_COLORS = [
    { v: "Base", bg: "#D9EAD3", fg: "#1E5128" },
    { v: "Plus", bg: "#FFF2CC", fg: "#7D6608" },
    { v: "Max",  bg: "#F4CCCC", fg: "#7B0000" },
  ];

  // Costruisce lookup link → stato dal gestionale (fonte di verità)
  const linkToStato = {};
  const lr = gestSheet.getLastRow();
  if (lr >= 3) {
    const gestLinks = gestSheet.getRange(3, G_COL.LINK,  lr - 2, 1).getValues();
    const gestStati = gestSheet.getRange(3, G_COL.STATO, lr - 2, 1).getValues();
    for (let i = 0; i < gestLinks.length; i++) {
      const lk = gestLinks[i][0];
      if (lk) linkToStato[lk] = gestStati[i][0] || "NUOVO";
    }
  }

  for (const mese of MESI_IT) {
    const sheet = clientSS.getSheetByName(mese);
    if (!sheet) continue;

    const headerB = sheet.getRange(1, 2).getValue();
    if (headerB === "STATO") {
      Logger.log("Già migrata: " + mese + " per " + clientName + " — skip");
      continue;
    }

    // Leggi link dalla colonna E (col 5) PRIMA di eliminare la colonna C
    const sheetLr = sheet.getLastRow();
    const statoVals = [];
    if (sheetLr >= 3) {
      const links = sheet.getRange(3, 5, sheetLr - 2, 1).getValues();
      for (let i = 0; i < links.length; i++) {
        const lk = links[i][0];
        statoVals.push([lk ? (linkToStato[lk] || "NUOVO") : ""]);
      }
    }

    // Elimina colonna C (RICEZIONE) — shift: D→C, E→D, F→E, G→F, H→G, I→H
    sheet.deleteColumn(3);

    // Aggiorna header B → STATO
    sheet.getRange(1, 2).setValue("STATO");

    // Larghezze colonne
    sheet.setColumnWidth(2, 90);   // STATO
    sheet.setColumnWidth(3, 90);   // SPEDIZIONE
    sheet.setColumnWidth(4, 240);  // LINK

    // Validazione e formattazione colonna B (STATO)
    sheet.getRange("B3:B1000")
      .clearDataValidations()
      .setHorizontalAlignment("center")
      .setFontColor("#000000");
    sheet.getRange("B3:B1000").setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(STATI_LISTA, false)
        .setAllowInvalid(false).build()
    );

    // Scrivi valori STATO
    if (statoVals.length > 0) {
      sheet.getRange(3, 2, statoVals.length, 1).setValues(statoVals);
    }

    // Aggiorna formattazione colonna D (LINK, era E)
    sheet.getRange("D3:D1000")
      .setHorizontalAlignment("left")
      .setFontColor("#1155CC")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // Ricostruisce CF: STATO (B) + CV (G, era H)
    const statoRange = sheet.getRange("B3:B1000");
    const covRange   = sheet.getRange("G3:G1000");
    const newRules = [
      ...STATO_COLORS.map(c =>
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(c.v).setBackground(c.bg).setFontColor(c.fg)
          .setRanges([statoRange]).build()
      ),
      ...CV_COLORS.map(c =>
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(c.v).setBackground(c.bg).setFontColor(c.fg)
          .setRanges([covRange]).build()
      ),
    ];
    sheet.setConditionalFormatRules(newRules);

    Logger.log("Migrata scheda " + mese + " per " + clientName);
  }
}
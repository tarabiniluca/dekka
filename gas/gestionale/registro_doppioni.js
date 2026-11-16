// ================================================================
// registro_doppioni.js — DEKKA gestionale
// Controllo automatico link doppi tra clienti.
// ================================================================

// -----------------------------------------------------------------
// COSTANTI
// -----------------------------------------------------------------
const REG_SHEET_NAME      = 'DEKKA_Registro';
const NOTE_PREFIX         = '⚠ DEKKA: ';
const BG_PRIORITA         = '#93C47D';   // verde chiaro Google Sheets
const BG_NO_ACQUISTO      = '#E06666';   // rosso chiaro Google Sheets
const PROP_ORIG_BG_PREFIX = 'dekka_orig_bg_';

// Header accettati per la colonna link (case-insensitive)
const LINK_COLUMN_HEADERS = ['link', 'url', 'link annuncio', 'annuncio', 'link prodotto'];

// Schede del gestionale che NON sono tab cliente (da NON scansionare)
const EXCLUDED_SHEETS = ['Dashboard', REG_SHEET_NAME];

// -----------------------------------------------------------------
// NORMALIZZAZIONE URL → codice annuncio univoco
// Ritorna null se l'URL non è di una piattaforma supportata
// o è un link catalogo SNKRDUNK (senza /used/).
// -----------------------------------------------------------------
function normalizzaCodiceAnnuncio(url) {
  if (!url || typeof url !== 'string') return null;
  url = url.trim().split('?')[0].replace(/\/+$/, '');  // strip query + slash finale

  // PayPayフリマ — es. https://paypayfleamarket.yahoo.co.jp/items/z1234567890
  let m = url.match(/paypay(?:fleamarket)?\.(?:yahoo\.co\.jp|me)\/(?:.*\/)?items?\/(z[A-Za-z0-9]+)/i);
  if (m) return 'ppf_' + m[1].toLowerCase();

  // Mercari JP — es. https://jp.mercari.com/item/m12345678901234
  m = url.match(/mercari\.com\/(?:.*\/)?item\/(m[A-Za-z0-9]+)/i);
  if (m) return 'mer_' + m[1].toLowerCase();

  // SNKRDUNK usato — es. https://snkrdunk.com/apparels/12345/used/67890
  // ESCLUSO: /apparels/ID senza /used/ (catalogo, non una copia specifica)
  m = url.match(/snkrdunk\.com\/(?:.*\/)?apparels\/([^/?#]+)\/used\/([A-Za-z0-9]+)/i);
  if (m) return 'snkr_' + m[1].toLowerCase() + '_' + m[2].toLowerCase();

  // SNKRDUNK catalogo puro → escluso
  if (url.match(/snkrdunk\.com\/(?:.*\/)?apparels\/[^/?#]+\/?$/i)) return null;

  return null;
}

// -----------------------------------------------------------------
// REGISTRO — scheda nascosta nel gestionale
// Colonne: A=clientName  B=codiceNorm  C=urlOriginale  D=timestampISO
// -----------------------------------------------------------------
function getRegistroSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(REG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(REG_SHEET_NAME);
    sh.hideSheet();
    sh.getRange(1, 1, 1, 4).setValues([['clientName', 'codiceNorm', 'urlOriginale', 'timestampISO']]);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 220);
    sh.setColumnWidth(3, 400);
    sh.setColumnWidth(4, 180);
  }
  return sh;
}

// -----------------------------------------------------------------
// REGISTRA LINK — chiamata dal Mangia Link via SpreadsheetApp.openById
// -----------------------------------------------------------------
function registraLink(clientName, url) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('registraLink: lock timeout per ' + clientName);
    return null;
  }
  try {
    const codice = normalizzaCodiceAnnuncio(url);
    if (!codice) return { codice: null };

    const timestamp = new Date().toISOString();
    const sh = getRegistroSheet_();
    const data = sh.getDataRange().getValues();

    let rigaEsistente = -1;
    let priorityClient = null;
    let priorityTs = null;

    for (let i = 1; i < data.length; i++) {
      const [cn, cod, , ts] = data[i];
      if (cod !== codice) continue;
      if (cn === clientName) {
        rigaEsistente = i + 1;  // 1-based
      } else {
        // Doppione di un altro cliente — priorità al timestamp più vecchio
        if (!priorityTs || ts < priorityTs) {
          priorityClient = cn;
          priorityTs = ts;
        }
      }
    }

    if (rigaEsistente > 0) {
      sh.getRange(rigaEsistente, 1, 1, 4).setValues([[clientName, codice, url, timestamp]]);
    } else {
      sh.appendRow([clientName, codice, url, timestamp]);
    }

    return { codice, timestamp, isDoppio: !!priorityClient, priorityClient, priorityTs };
  } finally {
    lock.releaseLock();
  }
}

// -----------------------------------------------------------------
// RIMUOVI LINK — chiamata dal Mangia Link quando il link viene
//               cancellato o sostituito con un altro URL
// -----------------------------------------------------------------
function rimuoviLink(clientName, url) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('rimuoviLink: lock timeout per ' + clientName);
    return;
  }
  try {
    const codice = normalizzaCodiceAnnuncio(url);
    if (!codice) return;

    const sh = getRegistroSheet_();
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === clientName && data[i][1] === codice) {
        sh.deleteRow(i + 1);
      }
    }

    // Riallinea subito le celle che coinvolgono questo codice
    _aggiornaPerCodice_(codice, sh);
  } finally {
    lock.releaseLock();
  }
}

// -----------------------------------------------------------------
// RIALLINEAMENTO GLOBALE — trigger ogni ~5 minuti (sicurezza)
// -----------------------------------------------------------------
function riallineaDoppioni() {
  const sh = getRegistroSheet_();
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return;

  // --- FASE 1: valida e pulisci il registro ---
  for (let i = data.length - 1; i >= 1; i--) {
    const [cn, , url] = data[i];
    const sheet = getSchedaCliente_(cn);
    if (!sheet) continue;
    const col = getColonnaLink_(sheet);
    if (!col) continue;
    if (!cercaUrlInScheda_(sheet, col, url)) {
      sh.deleteRow(i + 1);
      data.splice(i, 1);  // aggiorna array in-memory
    }
  }

  // --- FASE 2: raggruppa per codice e aggiorna celle ---
  const perCodice = {};
  for (let i = 1; i < data.length; i++) {
    const [cn, cod, url, ts] = data[i];
    if (!cod) continue;
    if (!perCodice[cod]) perCodice[cod] = [];
    perCodice[cod].push({ clientName: cn, url, ts });
  }

  for (const codice in perCodice) {
    aggiornaPerCodiceConDati_(codice, perCodice[codice]);
  }
}

// -----------------------------------------------------------------
// AGGIORNA UN SINGOLO CODICE (usato anche da rimuoviLink)
// -----------------------------------------------------------------
function _aggiornaPerCodice_(codice, sh) {
  if (!sh) sh = getRegistroSheet_();
  const data = sh.getDataRange().getValues();
  const voci = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === codice) {
      voci.push({ clientName: data[i][0], url: data[i][2], ts: data[i][3] });
    }
  }
  aggiornaPerCodiceConDati_(codice, voci);
}

function aggiornaPerCodiceConDati_(codice, voci) {
  if (voci.length === 0) return;

  if (voci.length === 1) {
    // Nessun doppione — rimuovi eventuale marcatura residua
    rimuoviMarcaturaUrl_(voci[0].clientName, voci[0].url);
    return;
  }

  // Ordina per timestamp — il primo ha la priorità
  voci.sort((a, b) => (a.ts < b.ts ? -1 : a.ts > b.ts ? 1 : 0));
  const prioritario = voci[0];
  const dataOra = formatDataOra_(prioritario.ts);

  voci.forEach((v, idx) => {
    const sheet = getSchedaCliente_(v.clientName);
    if (!sheet) return;
    const col = getColonnaLink_(sheet);
    if (!col) return;
    const riga = cercaUrlInScheda_(sheet, col, v.url);
    if (!riga) return;

    if (idx === 0) {
      const altriClienti = voci.slice(1).map(x => x.clientName).join(', ');
      const nota = NOTE_PREFIX + 'Doppio – priorità a questo cliente. Inviato anche da ' +
        altriClienti + ', ' + dataOra;
      marcaCella_(sheet, riga, col, BG_PRIORITA, nota);
    } else {
      const nota = NOTE_PREFIX + 'NON acquistare – inviato prima da ' +
        prioritario.clientName + ', ' + dataOra;
      marcaCella_(sheet, riga, col, BG_NO_ACQUISTO, nota);
    }
  });
}

// -----------------------------------------------------------------
// MARCA CELLA — applica sfondo + nota DEKKA, preserva note manuali
// -----------------------------------------------------------------
function marcaCella_(sheet, riga, colonnaLink, colore, notaDekka) {
  const cella = sheet.getRange(riga, colonnaLink);

  // Avvisa se ci sono CF rules che potrebbero prevalere sullo sfondo
  verificaCFConflitto_(sheet, colonnaLink);

  const noteAtt = cella.getNote() || '';
  const propKey = PROP_ORIG_BG_PREFIX + sheet.getName() + '_' + riga;
  const props = PropertiesService.getScriptProperties();

  // Salva origBg in PropertiesService se non ancora presente
  if (!props.getProperty(propKey)) {
    const mBg = noteAtt.match(/\[origBg:([^\]]+)\]/);
    if (mBg) {
      // Migra da nota (scritta da Mangia Link) a PropertiesService
      props.setProperty(propKey, mBg[1]);
    } else if (!noteAtt.includes(NOTE_PREFIX)) {
      // Prima marcatura assoluta — leggi dalla cella
      props.setProperty(propKey, cella.getBackground() || '#ffffff');
    }
  }

  cella.setBackground(colore);

  // Costruisce nota: mantieni manuale, imposta riga DEKKA senza tag [origBg:]
  const noteManuale = noteAtt.includes(NOTE_PREFIX)
    ? noteAtt.split('\n').filter(r => !r.startsWith(NOTE_PREFIX)).join('\n').trim()
    : noteAtt.replace(/\[origBg:[^\]]+\]/g, '').trim();
  const notaFinale = noteManuale ? noteManuale + '\n' + notaDekka : notaDekka;
  cella.setNote(notaFinale);
}

// -----------------------------------------------------------------
// RIMUOVI MARCATURA — ripristina colore originale, elimina nota DEKKA
// -----------------------------------------------------------------
function rimuoviMarcaturaUrl_(clientName, url) {
  const sheet = getSchedaCliente_(clientName);
  if (!sheet) return;
  const col = getColonnaLink_(sheet);
  if (!col) return;
  const riga = cercaUrlInScheda_(sheet, col, url);
  if (!riga) return;
  rimuoviMarcaturaCella_(sheet, riga, col);
}

function rimuoviMarcaturaCella_(sheet, riga, colonnaLink) {
  const cella = sheet.getRange(riga, colonnaLink);
  const nota = cella.getNote() || '';
  if (!nota.includes(NOTE_PREFIX)) return;

  const propKey = PROP_ORIG_BG_PREFIX + sheet.getName() + '_' + riga;
  const props = PropertiesService.getScriptProperties();
  let origBg = props.getProperty(propKey);

  // Fallback: origBg codificato nella nota da Mangia Link
  if (!origBg) {
    const m = nota.match(/\[origBg:([^\]]+)\]/);
    if (m) origBg = m[1];
  }

  cella.setBackground(origBg || null);
  if (origBg) props.deleteProperty(propKey);

  const nuovaNota = nota.split('\n')
    .filter(r => !r.startsWith(NOTE_PREFIX))
    .join('\n')
    .trim();
  cella.setNote(nuovaNota || '');
}

// -----------------------------------------------------------------
// HELPERS
// -----------------------------------------------------------------
function getSchedaCliente_(clientName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(clientName);
  if (!sh) Logger.log('getSchedaCliente_: scheda non trovata: ' + clientName);
  return sh || null;
}

function cercaUrlInScheda_(sheet, colonnaLink, url) {
  const lr = sheet.getLastRow();
  if (lr < 2) return null;
  const vals = sheet.getRange(2, colonnaLink, lr - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === url.trim()) return i + 2;  // 1-based
  }
  return null;
}

// Cache per evitare lettura header ripetuta nella stessa esecuzione
const colonnaLinkCache_ = {};
function getColonnaLink_(sheet) {
  const name = sheet.getName();
  if (colonnaLinkCache_[name] !== undefined) return colonnaLinkCache_[name];

  const lc = sheet.getLastColumn();
  if (lc < 1) { colonnaLinkCache_[name] = null; return null; }

  const headers = sheet.getRange(1, 1, 1, lc).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i]).toLowerCase().trim();
    if (LINK_COLUMN_HEADERS.indexOf(h) >= 0) {
      colonnaLinkCache_[name] = i + 1;
      return i + 1;
    }
  }
  Logger.log('getColonnaLink_: colonna link non trovata in "' + name + '". ' +
    'Headers trovati: ' + headers.join(', '));
  colonnaLinkCache_[name] = null;
  return null;
}

function formatDataOra_(isoString) {
  if (!isoString) return '—';
  const d = new Date(isoString);
  const p = n => String(n).padStart(2, '0');
  return p(d.getDate()) + '/' + p(d.getMonth() + 1) + ' ' +
    p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
}

// Avvisa in Logger se ci sono CF rules sulla colonna link (eseguito una volta per scheda)
const cfCheckFatto_ = {};
function verificaCFConflitto_(sheet, colonnaLink) {
  const name = sheet.getName();
  if (cfCheckFatto_[name]) return;
  cfCheckFatto_[name] = true;

  const rules = sheet.getConditionalFormatRules();
  const conflict = rules.some(r =>
    r.getRanges().some(range =>
      range.getColumn() <= colonnaLink && colonnaLink <= range.getLastColumn()
    )
  );
  if (conflict) {
    Logger.log(
      '[DEKKA] ⚠ La scheda "' + name + '" ha formattazione condizionale sulla colonna link (' + colonnaLink + '). ' +
      'Se il colore doppione non è visibile, aggiungere una regola CF dedicata DEKKA con priorità massima.'
    );
  }
}

// -----------------------------------------------------------------
// onEdit — ripristina la cella e riallinea gli altri clienti quando
// un URL viene rimosso o sostituito manualmente nel gestionale.
// Trigger semplice: si attiva automaticamente senza installazione.
// -----------------------------------------------------------------
function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // Solo schede cliente
  if (EXCLUDED_SHEETS.indexOf(sheetName) >= 0) return;

  // Solo colonna LINK, una riga alla volta
  const col = getColonnaLink_(sheet);
  if (!col || range.getColumn() !== col || range.getNumRows() !== 1) return;

  const oldUrl = (e.oldValue !== undefined && e.oldValue !== null)
    ? String(e.oldValue).trim() : '';
  if (!oldUrl) return;

  const codice = normalizzaCodiceAnnuncio(oldUrl);
  if (!codice) return;

  // 1. Ripristina subito la cella modificata
  rimuoviMarcaturaCella_(sheet, range.getRow(), col);

  // 2. Rimuovi dal registro e riallinea gli altri clienti coinvolti
  rimuoviLink(sheetName, oldUrl);
}

// -----------------------------------------------------------------
// BACKFILL UNA TANTUM — importa link già presenti nel gestionale
// Eseguire una sola volta dopo il deploy iniziale.
// -----------------------------------------------------------------
function backfillRegistroStorico() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getRegistroSheet_();
  const dataReg = sh.getDataRange().getValues();

  // Set dei codici già registrati per evitare duplicati
  const registrati = new Set();
  for (let i = 1; i < dataReg.length; i++) {
    registrati.add(dataReg[i][0] + '|' + dataReg[i][1]);
  }

  const sheets = ss.getSheets().filter(s => EXCLUDED_SHEETS.indexOf(s.getName()) < 0);

  sheets.forEach(sheet => {
    const nome = sheet.getName();
    const col = getColonnaLink_(sheet);
    if (!col) return;

    const lr = sheet.getLastRow();
    if (lr < 2) return;
    const vals = sheet.getRange(2, col, lr - 1, 1).getValues();

    vals.forEach(row => {
      const url = String(row[0]).trim();
      if (!url) return;
      const codice = normalizzaCodiceAnnuncio(url);
      if (!codice) return;

      const key = nome + '|' + codice;
      if (registrati.has(key)) return;

      // Timestamp storico: data attuale alle 00:00:00
      const tsStorico = new Date();
      tsStorico.setHours(0, 0, 0, 0);
      sh.appendRow([nome, codice, url, tsStorico.toISOString()]);
      registrati.add(key);
    });
  });

  Logger.log('backfillRegistroStorico: completato. ' +
    'Verifica manualmente i doppioni con timestamp identico (stessa data).');
}

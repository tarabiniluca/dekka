// ================================================================
// integrazione_doppioni.js — DEKKA Mangia Link · Dennys Marzola
// Registra/rimuove link nel registro doppioni del gestionale e
// applica immediatamente la colorazione doppioni alle celle.
//
// Usa GESTIONALE_ID e CLIENT_NAME già definiti in questo script.
// ================================================================

const NOTE_PREFIX_ML      = '⚠ DEKKA: ';
const BG_PRIORITA_ML      = '#93C47D';
const BG_NO_ACQUISTO_ML   = '#E06666';
const LINK_COL_GESTIONALE = 5;   // colonna E — LINK nel gestionale

// -----------------------------------------------------------------
// Registra un link nel registro doppioni e colora immediatamente le
// celle coinvolte nel gestionale.
// Chiamare dopo pushToGestionale_().
// -----------------------------------------------------------------
function registraInGestionale_(url) {
  if (!url) return;
  try {
    const gestionale = SpreadsheetApp.openById(GESTIONALE_ID);
    const reg = gestionale.getSheetByName('DEKKA_Registro');
    if (!reg) {
      Logger.log('registraInGestionale_: scheda DEKKA_Registro non trovata nel gestionale.');
      return;
    }

    const codice = normalizzaCodiceAnnuncioML_(url);
    if (!codice) return;

    const timestamp = new Date().toISOString();
    const data = reg.getDataRange().getValues();

    let rigaEsistente = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === CLIENT_NAME && data[i][1] === codice) {
        rigaEsistente = i + 1;
        break;
      }
    }

    if (rigaEsistente > 0) {
      reg.getRange(rigaEsistente, 1, 1, 4).setValues([[CLIENT_NAME, codice, url, timestamp]]);
    } else {
      reg.appendRow([CLIENT_NAME, codice, url, timestamp]);
    }

    // Colorazione immediata — rilegge registro dopo la scrittura
    const dataAgg = reg.getDataRange().getValues();
    const voci = [];
    for (let i = 1; i < dataAgg.length; i++) {
      if (dataAgg[i][1] === codice) {
        voci.push({ clientName: dataAgg[i][0], url: dataAgg[i][2], ts: dataAgg[i][3] });
      }
    }
    _coloraImmediato_(gestionale, voci);

    Logger.log('registraInGestionale_: registrato codice=' + codice + ' client=' + CLIENT_NAME);
  } catch (e) {
    Logger.log('registraInGestionale_: ' + e.message);
  }
}

// -----------------------------------------------------------------
// Rimuove dal registro la voce corrispondente a un URL e aggiorna
// immediatamente le celle dei clienti rimanenti.
// Chiamare prima di rimuovere o sostituire un link.
// -----------------------------------------------------------------
function rimuoviDaGestionale_(url) {
  if (!url) return;
  try {
    const gestionale = SpreadsheetApp.openById(GESTIONALE_ID);
    const reg = gestionale.getSheetByName('DEKKA_Registro');
    if (!reg) return;

    const codice = normalizzaCodiceAnnuncioML_(url);
    if (!codice) return;

    // Rimuovi la riga del registro
    const data = reg.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === CLIENT_NAME && data[i][1] === codice) {
        reg.deleteRow(i + 1);
        data.splice(i, 1);
      }
    }

    // Rimuovi subito la marcatura dalla cella di questo cliente
    _rimuoviMarcaturaImm_(gestionale, CLIENT_NAME, url);

    // Riallinea i clienti rimanenti per questo codice
    const voci = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === codice) {
        voci.push({ clientName: data[i][0], url: data[i][2], ts: data[i][3] });
      }
    }
    _coloraImmediato_(gestionale, voci);
  } catch (e) {
    Logger.log('rimuoviDaGestionale_: ' + e.message);
  }
}

// -----------------------------------------------------------------
// Colorazione immediata — aggiorna tutte le celle di un gruppo di voci
// (stesso codice annuncio). Se ≤1 voce, rimuove marcature residue.
// -----------------------------------------------------------------
function _coloraImmediato_(gestionale, voci) {
  if (voci.length === 0) return;

  if (voci.length === 1) {
    _rimuoviMarcaturaImm_(gestionale, voci[0].clientName, voci[0].url);
    return;
  }

  voci.sort((a, b) => (a.ts < b.ts ? -1 : a.ts > b.ts ? 1 : 0));
  const prio = voci[0];
  const dataOra = _formatDataOraImm_(prio.ts);

  voci.forEach((v, idx) => {
    const sheet = gestionale.getSheetByName(v.clientName);
    if (!sheet) return;
    const riga = _cercaUrlImm_(sheet, v.url);
    if (!riga) return;

    if (idx === 0) {
      const altri = voci.slice(1).map(x => x.clientName).join(', ');
      _marcaCellaImm_(sheet, riga, LINK_COL_GESTIONALE, BG_PRIORITA_ML,
        NOTE_PREFIX_ML + 'Doppio – priorità a questo cliente. Inviato anche da ' +
        altri + ', ' + dataOra);
    } else {
      _marcaCellaImm_(sheet, riga, LINK_COL_GESTIONALE, BG_NO_ACQUISTO_ML,
        NOTE_PREFIX_ML + 'NON acquistare – inviato prima da ' +
        prio.clientName + ', ' + dataOra);
    }
  });
}

// Applica sfondo + nota DEKKA a una cella; salva origBg nella nota
// come [origBg:#xxxxxx] — leggibile dal trigger del gestionale per la
// migrazione a PropertiesService.
function _marcaCellaImm_(sheet, riga, col, colore, notaDekka) {
  const cella = sheet.getRange(riga, col);
  const noteAtt = cella.getNote() || '';

  // Leggi origBg: da nota precedente (già marcata) o dalla cella corrente
  let origBg;
  const mBg = noteAtt.match(/\[origBg:([^\]]+)\]/);
  if (mBg) {
    origBg = mBg[1];
  } else {
    origBg = cella.getBackground() || '#ffffff';
  }

  cella.setBackground(colore);

  // Costruisce nota: mantieni manuale, aggiorna riga DEKKA con tag origBg
  const notaManuale = noteAtt.includes(NOTE_PREFIX_ML)
    ? noteAtt.split('\n').filter(r => !r.startsWith(NOTE_PREFIX_ML)).join('\n').trim()
    : noteAtt.replace(/\[origBg:[^\]]+\]/g, '').trim();
  const notaFinale = (notaManuale ? notaManuale + '\n' : '') +
    notaDekka + ' [origBg:' + origBg + ']';
  cella.setNote(notaFinale);
}

// Ripristina la cella rimuovendo marcatura DEKKA e colore originale
function _rimuoviMarcaturaImm_(gestionale, clientName, url) {
  const sheet = gestionale.getSheetByName(clientName);
  if (!sheet) return;
  const riga = _cercaUrlImm_(sheet, url);
  if (!riga) return;

  const cella = sheet.getRange(riga, LINK_COL_GESTIONALE);
  const nota = cella.getNote() || '';
  if (!nota.includes(NOTE_PREFIX_ML)) return;

  const m = nota.match(/\[origBg:([^\]]+)\]/);
  cella.setBackground(m ? m[1] : null);

  const nuovaNota = nota.split('\n')
    .filter(r => !r.startsWith(NOTE_PREFIX_ML))
    .join('\n').trim();
  cella.setNote(nuovaNota || '');
}

// Cerca URL nella colonna LINK di una scheda cliente (riga 1-based)
function _cercaUrlImm_(sheet, url) {
  const lr = sheet.getLastRow();
  if (lr < 2) return null;
  const vals = sheet.getRange(2, LINK_COL_GESTIONALE, lr - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === url.trim()) return i + 2;
  }
  return null;
}

function _formatDataOraImm_(isoString) {
  if (!isoString) return '—';
  const d = new Date(isoString);
  const p = n => String(n).padStart(2, '0');
  return p(d.getDate()) + '/' + p(d.getMonth() + 1) + ' ' +
    p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
}

// -----------------------------------------------------------------
// Copia locale di normalizzaCodiceAnnuncio — deve restare sincronizzata
// con la versione in registro_doppioni.js del gestionale.
// -----------------------------------------------------------------
function normalizzaCodiceAnnuncioML_(url) {
  if (!url || typeof url !== 'string') return null;
  url = url.trim().split('?')[0].replace(/\/+$/, '');

  let m = url.match(/paypay(?:fleamarket)?\.(?:yahoo\.co\.jp|me)\/(?:.*\/)?items?\/(z[A-Za-z0-9]+)/i);
  if (m) return 'ppf_' + m[1].toLowerCase();

  m = url.match(/mercari\.com\/(?:.*\/)?item\/(m[A-Za-z0-9]+)/i);
  if (m) return 'mer_' + m[1].toLowerCase();

  m = url.match(/snkrdunk\.com\/(?:.*\/)?apparels\/([^/?#]+)\/used\/([A-Za-z0-9]+)/i);
  if (m) return 'snkr_' + m[1].toLowerCase() + '_' + m[2].toLowerCase();

  if (url.match(/snkrdunk\.com\/(?:.*\/)?apparels\/[^/?#]+\/?$/i)) return null;

  return null;
}

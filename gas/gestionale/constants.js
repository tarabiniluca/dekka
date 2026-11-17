// ─────────────────────────────────────────────────────────────
// COLONNE — Gestionale (schede cliente)
// ─────────────────────────────────────────────────────────────
const G_COL = {
  DATA_RICEZIONE : 2,  // B
  STATO          : 3,  // C
  SPEDIZIONE     : 4,  // D
  LINK           : 5,  // E
  QTA            : 6,  // F
  IMPORTI        : 7,  // G
  COV_PIANO      : 8,  // H  ← Copertura Valore: piano attivato
  COV_NOTA       : 9,  // I  ← Copertura Valore: specifica carta
};

// ─────────────────────────────────────────────────────────────
// COLONNE — Client cliente (schede mese)
// ─────────────────────────────────────────────────────────────
const C_COL = {
  STATO      : 2,  // B
  SPEDIZIONE : 3,  // C
  LINK       : 4,  // D
  QTA        : 5,  // E
  IMPORTI    : 6,  // F
  COV_PIANO  : 7,  // G  ← Copertura Valore: piano attivato
  COV_NOTA   : 8,  // H  ← Copertura Valore: specifica carta
};

// ─────────────────────────────────────────────────────────────
// COLONNE — Dashboard
// ─────────────────────────────────────────────────────────────
const D_COL = {
  CLIENTE   : 2,  // B
  PIANO     : 3,  // C
  UTILIZZO  : 4,  // D
  ATTIVI    : 5,  // E
  SALDO     : 6,  // F
  PIVA      : 7,  // G
  INDIRIZZO : 8,  // H
  EMAIL     : 9,  // I
  SCADENZA  : 10, // J  ← Data scadenza piano attivo
};

// ─────────────────────────────────────────────────────────────
// MESI italiani
// ─────────────────────────────────────────────────────────────
const MESI_IT = [
  "Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
  "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"
];

// ─────────────────────────────────────────────────────────────
// PIANI DEKKA — allineati alla pagina ufficiale "Piani e Tariffe"
//
// Unica fonte di verità per piani e massimali: menu a tendina
// della Dashboard, prompt dei menu, calcolo utilizzo e push al
// client leggono tutti da qui.
// Chiave = valore salvato in Dashboard (maiuscolo).
// label  = nome mostrato al cliente nel MANGIA LINK.
// bg/fg  = colori della cella PIANO in Dashboard (formattazione
//          condizionale). Palette originale della Dashboard:
//          Trial grigio, Basic giallo, Pro verde invariati;
//          Starter eredita il blu del vecchio piano Core.
// ─────────────────────────────────────────────────────────────
const PIANI = {
  TRIAL   : { label: "Trial",   massimale: 100, eccedenza: null, canone: 0,      bg: "#EFEFEF", fg: "#555555" }, // non sforabile, non rinnovabile
  STARTER : { label: "Starter", massimale: 40,  eccedenza: 850,  canone: 34000,  bg: "#CFE2FF", fg: "#0D47A1" },
  BASIC   : { label: "Basic",   massimale: 100, eccedenza: 650,  canone: 70000,  bg: "#FFF2CC", fg: "#7D6608" },
  PRO     : { label: "Pro",     massimale: 180, eccedenza: 550,  canone: 108000, bg: "#D9EAD3", fg: "#1E5128" },
};

// Valori validi per il menu a tendina PIANO in Dashboard (ordine di visualizzazione)
const PIANI_LISTA = Object.keys(PIANI);

// Massimali per piano (derivati da PIANI — non modificare a mano)
const PIANI_MAX = Object.fromEntries(PIANI_LISTA.map(k => [k, PIANI[k].massimale]));

// ─────────────────────────────────────────────────────────────
// COPERTURA VALORE — piani e costi (¥ per carta)
// ─────────────────────────────────────────────────────────────
const CV_PIANI = {
  "—"    : 0,
  "Base" : 300,
  "Plus" : 500,
  "Max"  : 900,
};

// Valori validi per il dropdown CV nel MANGIA LINK del cliente
const CV_VALORI = ["—", "Base", "Plus", "Max"];

// ─────────────────────────────────────────────────────────────
// COPERTURA VALORE — celle nel foglio MANGIA LINK del cliente
// ─────────────────────────────────────────────────────────────
const CV_CELLE = {
  PIANO      : "B3",   // Dropdown selezione piano CV
  NOTA       : "B4",   // Testo libero specifica carta
  LABEL_NOTA : "C4",   // Label disclaimer (scritta da setup, non modificabile)
  RESET      : ["B3", "B4"],  // Celle da azzerare dopo ogni invio
};

// ─────────────────────────────────────────────────────────────
// COPERTURA VALORE — formattazione condizionale colonna H
// ─────────────────────────────────────────────────────────────
const CV_COLORI = {
  "—"    : { bg: "#EFEFEF", fg: "#555555" },
  "Base" : { bg: "#D9EAD3", fg: "#1E5128" },
  "Plus" : { bg: "#FFF2CC", fg: "#7D6608" },
  "Max"  : { bg: "#F4CCCC", fg: "#7B0000" },
};
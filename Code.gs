/**
 * StoneTroopers — API du planning (Google Apps Script)
 * À coller dans le Google Sheet du planning : Extensions > Apps Script.
 * Puis Déployer > Nouveau déploiement > Application Web
 *   - Exécuter en tant que : Moi
 *   - Qui a accès : Tout le monde
 * L'URL obtenue (…/exec) se colle dans config.js sur GitHub.
 */
const SHEET_PLANNING = 'Planning';
const SHEET_VACANCES = 'Vacances';
const PLAYERS = ['Suettania', 'Adélie', 'Océane', 'Jean-Pierre', 'Pascal'];
const FIRST_ROW = 3;          // première ligne de match
const FIRST_PLAYER_COL = 9;   // colonne I
const VALUES = ['', 'Oui', 'Incertain', 'Non'];

function doGet() {
  return json_(readAll_());
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const row = Number(body.row);
    const idx = PLAYERS.indexOf(body.player);
    const value = String(body.value || '');
    if (idx < 0) return json_({ ok: false, error: 'joueur inconnu' });
    if (VALUES.indexOf(value) < 0) return json_({ ok: false, error: 'valeur invalide' });
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PLANNING);
    if (!(row >= FIRST_ROW && row <= sh.getLastRow()) || !sh.getRange(row, 1).getValue())
      return json_({ ok: false, error: 'match introuvable' });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sh.getRange(row, FIRST_PLAYER_COL + idx).setValue(value);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function readAll_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const sh = ss.getSheetByName(SHEET_PLANNING);
  const n = sh.getLastRow() - FIRST_ROW + 1;
  const matches = [];
  if (n > 0) {
    const range = sh.getRange(FIRST_ROW, 1, n, FIRST_PLAYER_COL - 1 + PLAYERS.length);
    const vals = range.getValues();
    const disp = range.getDisplayValues();
    vals.forEach(function (r, i) {
      if (!(r[0] instanceof Date)) return;
      const dispo = {};
      PLAYERS.forEach(function (p, j) { dispo[p] = String(r[FIRST_PLAYER_COL - 1 + j] || '').trim(); });
      matches.push({
        row: FIRST_ROW + i,
        date: Utilities.formatDate(r[0], tz, 'yyyy-MM-dd'),
        heure: ((disp[i][2] || '').match(/^\d{1,2}[:h]\d{2}/) || [disp[i][2]])[0],
        piste: String(r[3] || ''),
        adversaire: String(r[4] || ''),
        dispo: dispo
      });
    });
  }
  const vacances = [];
  const vs = ss.getSheetByName(SHEET_VACANCES);
  if (vs && vs.getLastRow() >= 3) {
    vs.getRange(3, 1, vs.getLastRow() - 2, 4).getValues().forEach(function (r) {
      if (r[0] && r[1] instanceof Date && r[2] instanceof Date) {
        vacances.push({
          joueur: String(r[0]),
          debut: Utilities.formatDate(r[1], tz, 'yyyy-MM-dd'),
          fin: Utilities.formatDate(r[2], tz, 'yyyy-MM-dd'),
          commentaire: String(r[3] || '')
        });
      }
    });
  }
  return { matches: matches, vacances: vacances };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

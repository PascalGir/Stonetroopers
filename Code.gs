/**
 * StoneTroopers — API de la page web (Google Apps Script)
 *
 * Installation : dans le Google Sheet « StoneTroopers_Base », Extensions > Apps Script,
 * coller ce fichier, enregistrer, puis Déployer > Nouveau déploiement > Application Web
 *   - Exécuter en tant que : Moi
 *   - Qui a accès : Tout le monde
 * L'URL obtenue (…/exec) se colle dans config.js sur GitHub.
 *
 * Après chaque modification de ce script : Déployer > Gérer les déploiements > ✏️ > Nouvelle version.
 */

// Laisser vide = tout le monde peut gérer les matchs, joueurs et vacances.
// Mettre un code (ex. 'skip2026') = ces actions demandent ce code sur la page.
const CODE_CAPITAINE = '';

const ONGLET_MATCHS = 'Matchs';
const ONGLET_JOUEURS = 'Joueurs';
const ONGLET_VACANCES = 'Vacances';
const COLS_MATCH = ['ID', 'Date', 'Heure', 'Piste', 'Adversaire', 'Score nous', 'Score eux', 'Notes'];
const REPONSES = ['', 'Oui', 'Incertain', 'Non'];

/* ---------- Lecture ---------- */

function doGet() {
  try {
    return json_(lireTout_());
  } catch (err) {
    return json_({ error: String(err) });
  }
}

function lireTout_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  const joueurs = lignes_(ONGLET_JOUEURS).map(function (o) {
    return { nom: String(o['Nom']).trim(), actif: String(o['Actif']).trim() !== 'Non', poste: String(o['Poste'] || '') };
  }).filter(function (j) { return j.nom; });

  const mSheet = feuille_(ONGLET_MATCHS);
  const entetes = entetes_(mSheet);
  const matches = [];
  if (mSheet.getLastRow() >= 2) {
    const range = mSheet.getRange(2, 1, mSheet.getLastRow() - 1, entetes.length);
    const vals = range.getValues();
    const disp = range.getDisplayValues();
    vals.forEach(function (r, i) {
      const o = {}, d = {};
      entetes.forEach(function (h, k) { o[h] = r[k]; d[h] = disp[i][k]; });
      if (!o['ID'] || !(o['Date'] instanceof Date)) return;
      const dispo = {};
      joueurs.forEach(function (j) { if (j.nom in o) dispo[j.nom] = String(o[j.nom] || '').trim(); });
      matches.push({
        id: String(o['ID']),
        date: Utilities.formatDate(o['Date'], tz, 'yyyy-MM-dd'),
        heure: heure_(d['Heure']),
        piste: String(o['Piste'] || ''),
        adversaire: String(o['Adversaire'] || ''),
        scoreNous: d['Score nous'] || '',
        scoreEux: d['Score eux'] || '',
        notes: String(o['Notes'] || ''),
        dispo: dispo
      });
    });
  }

  const vacances = lignes_(ONGLET_VACANCES).filter(function (o) {
    return o['ID'] && o['Joueur'] && o['Début'] instanceof Date && o['Fin'] instanceof Date;
  }).map(function (o) {
    return {
      id: String(o['ID']), joueur: String(o['Joueur']),
      debut: Utilities.formatDate(o['Début'], tz, 'yyyy-MM-dd'),
      fin: Utilities.formatDate(o['Fin'], tz, 'yyyy-MM-dd'),
      commentaire: String(o['Commentaire'] || '')
    };
  });

  return { joueurs: joueurs, matches: matches, vacances: vacances, capitaine: CODE_CAPITAINE !== '' };
}

/* ---------- Écriture ---------- */

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const b = JSON.parse(e.postData.contents);
    lock.waitLock(10000);
    const res = action_(b);
    SpreadsheetApp.flush();
    return json_(Object.assign({ ok: true }, res || {}));
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function action_(b) {
  switch (b.action) {
    case 'setDispo': return setDispo_(b);
    case 'addVacances': return addVacances_(b);
    case 'deleteVacances': return deleteVacances_(b);
    case 'saveMatch': capitaine_(b); return saveMatch_(b);
    case 'deleteMatch': capitaine_(b); return deleteMatch_(b);
    case 'savePlayer': capitaine_(b); return savePlayer_(b);
    default: throw new Error('action inconnue');
  }
}

function setDispo_(b) {
  if (REPONSES.indexOf(b.value) < 0) throw new Error('réponse invalide');
  const sh = feuille_(ONGLET_MATCHS);
  const col = entetes_(sh).indexOf(b.joueur) + 1;
  if (col <= COLS_MATCH.length) throw new Error('joueur inconnu');
  const row = ligneParId_(sh, b.id);
  sh.getRange(row, col).setValue(b.value);
}

function saveMatch_(b) {
  const m = b.match || {};
  if (!/^\d{4}-\d{2}-\d{2}$/.test(m.date || '')) throw new Error('date invalide');
  if (!m.adversaire) throw new Error('adversaire manquant');
  const sh = feuille_(ONGLET_MATCHS);
  const h = entetes_(sh);
  const p = m.date.split('-').map(Number);
  const champs = {
    'Date': new Date(p[0], p[1] - 1, p[2]),
    'Heure': String(m.heure || ''),
    'Piste': String(m.piste || ''),
    'Adversaire': String(m.adversaire),
    'Score nous': m.scoreNous === undefined ? undefined : String(m.scoreNous),
    'Score eux': m.scoreEux === undefined ? undefined : String(m.scoreEux),
    'Notes': m.notes === undefined ? undefined : String(m.notes)
  };
  let row, id = m.id;
  if (id) {
    row = ligneParId_(sh, id);
  } else {
    id = nouvelId_(sh, 'M');
    row = sh.getLastRow() + 1;
    sh.getRange(row, h.indexOf('ID') + 1).setValue(id);
  }
  Object.keys(champs).forEach(function (k) {
    if (champs[k] === undefined) return;
    const c = sh.getRange(row, h.indexOf(k) + 1);
    if (k === 'Heure') c.setNumberFormat('@');
    c.setValue(champs[k]);
  });
  sh.getRange(row, h.indexOf('Date') + 1).setNumberFormat('dd/mm/yyyy');
  trier_(sh);
  return { id: id };
}

function deleteMatch_(b) {
  const sh = feuille_(ONGLET_MATCHS);
  sh.deleteRow(ligneParId_(sh, b.id));
}

function addVacances_(b) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(b.debut || '') || !/^\d{4}-\d{2}-\d{2}$/.test(b.fin || '')) throw new Error('dates invalides');
  if (b.fin < b.debut) throw new Error('la fin est avant le début');
  const joueurs = lignes_(ONGLET_JOUEURS).map(function (o) { return String(o['Nom']).trim(); });
  if (joueurs.indexOf(b.joueur) < 0) throw new Error('joueur inconnu');
  const sh = feuille_(ONGLET_VACANCES);
  const id = nouvelId_(sh, 'V');
  const d = function (s) { const p = s.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); };
  sh.appendRow([id, b.joueur, d(b.debut), d(b.fin), String(b.commentaire || '')]);
  sh.getRange(sh.getLastRow(), 3, 1, 2).setNumberFormat('dd/mm/yyyy');
  return { id: id };
}

function deleteVacances_(b) {
  const sh = feuille_(ONGLET_VACANCES);
  sh.deleteRow(ligneParId_(sh, b.id));
}

function savePlayer_(b) {
  const nom = String(b.nom || '').trim();
  if (!nom) throw new Error('nom manquant');
  if (COLS_MATCH.indexOf(nom) >= 0) throw new Error('nom réservé');
  const js = feuille_(ONGLET_JOUEURS);
  const noms = js.getLastRow() >= 2 ? js.getRange(2, 1, js.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  const i = noms.indexOf(nom);
  const actif = b.actif === false ? 'Non' : 'Oui';
  if (i >= 0) {
    js.getRange(i + 2, 2, 1, 2).setValues([[actif, String(b.poste || '')]]);
  } else {
    js.appendRow([nom, actif, String(b.poste || '')]);
    const ms = feuille_(ONGLET_MATCHS);
    if (entetes_(ms).indexOf(nom) < 0) {
      const col = ms.getLastColumn() + 1;
      ms.getRange(1, col).setValue(nom)
        .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1d4e89').setHorizontalAlignment('center');
    }
  }
}

/* ---------- Outils ---------- */

function capitaine_(b) {
  if (CODE_CAPITAINE && b.code !== CODE_CAPITAINE) throw new Error('code capitaine incorrect');
}
function feuille_(nom) {
  const sh = SpreadsheetApp.getActive().getSheetByName(nom);
  if (!sh) throw new Error('onglet « ' + nom + ' » introuvable');
  return sh;
}
function entetes_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
}
function lignes_(nom) {
  const sh = feuille_(nom);
  if (sh.getLastRow() < 2) return [];
  const h = entetes_(sh);
  return sh.getRange(2, 1, sh.getLastRow() - 1, h.length).getValues().map(function (r) {
    const o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o;
  });
}
function ligneParId_(sh, id) {
  if (!id) throw new Error('identifiant manquant');
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return String(r[0]); });
  const i = ids.indexOf(String(id));
  if (i < 1) throw new Error('élément introuvable (' + id + ')');
  return i + 1;
}
function nouvelId_(sh, prefixe) {
  let max = 0;
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
      const m = String(r[0]).match(/(\d+)$/); if (m) max = Math.max(max, Number(m[1]));
    });
  }
  return prefixe + ('00' + (max + 1)).slice(-3);
}
function trier_(sh) {
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).sort({ column: 2, ascending: true });
}
function heure_(s) {
  const m = String(s || '').match(/(\d{1,2})\s*[:h]\s*(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : String(s || '');
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

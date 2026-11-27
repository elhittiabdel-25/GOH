/**
 * Gate of Hope — API de l'application + synchronisation avec Google Sheets.
 *
 * Ce script est attaché au classeur « Collectes Gate of Hope ».
 * Le classeur reste LA référence : l'application lit et écrit directement dedans.
 *
 * Onglets gérés par ce script (les onglets PROJETS et PROJETS.OLD ne sont jamais modifiés) :
 *   PROJETS CLAUDE      collectes en cours (une ligne = une collecte)
 *   PROJETS.OLD CLAUDE  collectes clôturées (historique)
 *   DONS                journal de tous les dons reçus
 *   MEMBRES             personnes autorisées à utiliser l'application
 *   JOURNAL APP         traces des actions faites depuis l'application
 *   MOUVEMENTS          argent qui change de main (transfert entre membres, remise au bénéficiaire)
 *
 * Suivi de l'argent (« qui a quoi ») : chaque don indique qui a reçu l'argent (colonne
 * « Qui a l'argent » de DONS), puis les MOUVEMENTS le font passer d'une personne à l'autre.
 * Ces colonnes sont ajoutées à droite : les formules existantes ne les lisent pas.
 */

var CFG = {
  ACTIVE: 'PROJETS CLAUDE',
  ARCHIVE: 'PROJETS.OLD CLAUDE',
  DONS: 'DONS',
  MEMBRES: 'MEMBRES',
  JOURNAL: 'JOURNAL APP',
  MOUVEMENTS: 'MOUVEMENTS',
  SRC_ACTIVE: 'PROJETS',
  SRC_ARCHIVE: 'PROJETS.OLD',
  TOKEN_DAYS: 120,
  VERSION: '1.1.0',
  SCHEMA: 2          // 2 = suivi de l'argent (colonnes « Responsable », « Qui a l'argent », onglet MOUVEMENTS)
};

var H_CASE = ['ID', 'Types de collectes', 'Collectes en cours', 'Budget global nécessaire (Dhs)',
  'Montant collecté (Dhs)', 'Reste à collecter (Dhs)', 'Situation actuelle', 'Cycle', 'Ouverte le'];
var H_ACTIVE = H_CASE.concat(['Mise à jour le', 'Responsable']);
var H_ARCHIVE = H_CASE.concat(['Clôturée le', 'Responsable']);
var H_DONS = ['Réf', 'Date du don', 'ID collecte', 'Cycle', 'Montant (Dhs)', 'Donateur',
  'Mode de paiement', 'Note', 'Saisi par', 'Saisi le', 'Annulé', "Qui a l'argent"];
var H_MVT = ['Réf', 'Date', 'ID collecte', 'Cycle', 'Montant (Dhs)', "De (qui avait l'argent)",
  'Vers (qui le reçoit)', 'Note', 'Saisi par', 'Saisi le', 'Annulé'];
var H_MEMBRES = ['Nom', 'Email', 'Rôle (admin / membre)', "Code d'accès", 'Actif (OUI / NON)'];
var H_JOURNAL = ['Horodatage', 'Utilisateur', 'Action', 'Détails'];

// Index (base 0) des colonnes
var C = { ID: 0, CAT: 1, DESC: 2, BUDGET: 3, COLL: 4, RESTE: 5, SIT: 6, CYCLE: 7, OPEN: 8, LAST: 9, RESP: 10 };
var D = { REF: 0, DATE: 1, ID: 2, CYCLE: 3, AMT: 4, DONOR: 5, MODE: 6, NOTE: 7, BY: 8, AT: 9, CANCEL: 10, HOLDER: 11 };
var MV = { REF: 0, DATE: 1, ID: 2, CYCLE: 3, AMT: 4, FROM: 5, TO: 6, NOTE: 7, BY: 8, AT: 9, CANCEL: 10 };

var UNASSIGNED = 'Non attribué';   // dons enregistrés sans préciser qui a l'argent (dont l'historique)
var BENEF = 'Bénéficiaire';         // « Vers » d'un mouvement : argent remis / dépensé pour le cas
var M = { NAME: 0, EMAIL: 1, ROLE: 2, CODE: 3, ACTIVE: 4 };

var MODE_OPENING = 'Reprise Excel';
var MODE_ADJUST = 'Ajustement feuille';
var MODES = ['Espèces', 'Virement', 'Chèque', 'Carte', 'Autre', MODE_OPENING, MODE_ADJUST];

var PREFIXES = [
  ['cas medicaux', 'MED'], ['medical', 'MED'], ['frais mensuels', 'FML'], ['loyer', 'FML'],
  ['mosquee', 'MOS'], ['forage', 'PUI'], ['puits', 'PUI'], ['maison de vie', 'MDV'],
  ['gaza', 'GAZ'], ['maraude', 'MAR'], ['ramadan', 'RAM'], ['aid', 'RAM'],
  ['reconstruction', 'LOG'], ['logement', 'LOG']
];

/* ====================================================================== */
/*  Menu dans Google Sheets                                                */
/* ====================================================================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Gate of Hope')
    .addItem('Clôturer la collecte sélectionnée', 'menuCloturer')
    .addItem('Vérifier / réparer les onglets', 'menuVerifier')
    .addSeparator()
    .addItem('Installer (première fois uniquement)', 'installer')
    .addToUi();
}

/** Déclencheur simple : chaque modification manuelle est remise en forme (ID, cycle, formules). */
function onEdit(e) {
  try {
    var sh = e && e.range && e.range.getSheet();
    if (!sh) return;
    var name = sh.getName();
    if ([CFG.ACTIVE, CFG.DONS, CFG.MEMBRES].indexOf(name) < 0) return;
    var who = '';
    try { who = Session.getActiveUser().getEmail(); } catch (err) { /* ignore */ }
    var opts = { user: who || 'Feuille Google' };
    if (name === CFG.ACTIVE && e.range.getColumn() === C.ID + 1 && e.range.getRow() > 1) {
      opts.idEditedRows = [];
      for (var r = 0; r < e.range.getNumRows(); r++) opts.idEditedRows.push(e.range.getRow() + r);
    }
    withLock_(function () { sync_(opts); });
  } catch (err) {
    console.error(err);
  }
}

function menuVerifier() {
  var res = withLock_(function () { return sync_({ user: currentUserEmail_(), forceUpgrade: true }); });
  SpreadsheetApp.getUi().alert('Vérification terminée.\n\n' + (res.notes.length ? res.notes.join('\n') : 'Tout est en ordre.'));
}

function menuCloturer() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== CFG.ACTIVE) {
    ui.alert('Placez-vous sur une ligne de l\'onglet « ' + CFG.ACTIVE + ' » puis relancez.');
    return;
  }
  var row = sh.getActiveRange().getRow();
  if (row < 2) { ui.alert('Sélectionnez une ligne de collecte.'); return; }
  var vals = sh.getRange(row, 1, 1, H_ACTIVE.length).getValues()[0];
  var id = String(vals[C.ID] || '').trim();
  if (!id) { ui.alert('Cette ligne n\'a pas d\'ID.'); return; }
  var resp = ui.prompt('Clôturer ' + id,
    String(vals[C.DESC]).slice(0, 200) + '\n\nSituation finale (laisser vide pour « Collecte clôturée ») :',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var res = withLock_(function () {
    sync_({ user: currentUserEmail_() });
    return closeCase_(id, resp.getResponseText(), currentUserEmail_() || 'Feuille Google');
  });
  ui.alert(id + ' a été déplacée vers « ' + CFG.ARCHIVE + ' ».' + (res.held && res.held.length
    ? '\n\nAttention, argent encore détenu pour ce cas :\n' + res.held.map(function (h) { return '• ' + h.name + ' : ' + fmt_(h.amount) + ' Dhs'; }).join('\n')
    : ''));
}

/* ====================================================================== */
/*  Installation                                                           */
/* ====================================================================== */

/**
 * Crée et remplit les onglets à partir de PROJETS et PROJETS.OLD.
 * Ne touche jamais aux onglets d'origine. Refuse de s'exécuter si les onglets
 * de l'application contiennent déjà des données (pour ne rien écraser).
 */
function installer() {
  var ss = SpreadsheetApp.getActive();
  var report = withLock_(function () { return install_(ss); });
  var msg = 'Installation terminée.\n\n' + report.join('\n');
  console.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* exécuté depuis l'éditeur */ }
  return msg;
}

function install_(ss) {
  [CFG.ACTIVE, CFG.ARCHIVE, CFG.DONS].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() > 1) {
      throw new Error('L\'onglet « ' + n + ' » contient déjà des données : installation annulée pour ne rien écraser.');
    }
  });

  var src = readSource_(ss.getSheetByName(CFG.SRC_ACTIVE), true);
  var old = readSource_(ss.getSheetByName(CFG.SRC_ARCHIVE), false);
  var today = startOfDay_(new Date());
  var now = new Date();
  var unmatched = 0;

  function resolveId(r) {
    if (r.id) return r.id;
    var k = keyHash_(r.desc);
    if (typeof MIGRATION_MAP !== 'undefined' && MIGRATION_MAP[k]) return MIGRATION_MAP[k];
    unmatched++;
    return '';
  }
  src.forEach(function (r) { r.id = resolveId(r); });
  old.forEach(function (r) { r.id = resolveId(r); });

  // IDs manquants (lignes ajoutées depuis l'analyse) : nouveaux IDs
  var used = {};
  src.concat(old).forEach(function (r) { if (r.id) used[r.id] = true; });
  function freshId(cat) {
    var p = prefixFor_(cat), n = 1, id;
    do { id = p + '-' + ('00' + n).slice(-3); n++; } while (used[id]);
    used[id] = true;
    return id;
  }
  // lignes sans correspondance : les textes identiques (en cours ou historique) partagent un même nouvel ID
  var newByKey = {};
  src.concat(old).forEach(function (r) {
    if (r.id) return;
    var k = normKey_(r.desc);
    r.id = newByKey[k] || (newByKey[k] = freshId(r.cat));
  });

  // Cycles : dans l'historique, la ligne la plus basse est la plus ancienne (cycle 1)
  var count = {};
  for (var i = old.length - 1; i >= 0; i--) {
    count[old[i].id] = (count[old[i].id] || 0) + 1;
    old[i].cycle = count[old[i].id];
  }
  src.forEach(function (r) {
    count[r.id] = (count[r.id] || 0) + 1;
    r.cycle = count[r.id];
  });

  var shA = ensureSheet_(ss, CFG.ACTIVE, H_ACTIVE);
  var shO = ensureSheet_(ss, CFG.ARCHIVE, H_ARCHIVE);
  var shD = ensureSheet_(ss, CFG.DONS, H_DONS);
  var shM = ensureSheet_(ss, CFG.MEMBRES, H_MEMBRES);
  var shJ = ensureSheet_(ss, CFG.JOURNAL, H_JOURNAL);
  ensureSheet_(ss, CFG.MOUVEMENTS, H_MVT);

  if (old.length) {
    shO.getRange(2, 1, old.length, H_ARCHIVE.length).setValues(old.map(function (r) {
      return [r.id, r.cat, r.desc, r.budget, r.coll, r.reste, r.sit, r.cycle, '', '', ''];
    }));
  }
  if (src.length) {
    shA.getRange(2, 1, src.length, H_ACTIVE.length).setValues(src.map(function (r) {
      return [r.id, r.cat, r.desc, r.budget, '', '', r.sit, r.cycle, today, now, ''];
    }));
    src.forEach(function (r, i) { setCaseFormulas_(shA, i + 2); });
    var dons = src.filter(function (r) { return num_(r.coll) > 0; }).map(function (r) {
      return [newRef_(), today, r.id, r.cycle, num_(r.coll), 'Solde repris de l\'Excel', MODE_OPENING,
        'Montant déjà collecté avant la mise en service de l\'application', 'Installation', now, '', ''];
    });
    if (dons.length) shD.getRange(2, 1, dons.length, H_DONS.length).setValues(dons);
  }

  // Premier admin : la personne qui installe
  var me = currentUserEmail_();
  var code = '';
  if (shM.getLastRow() < 2) {
    code = randomCode_();
    shM.getRange(2, 1, 1, H_MEMBRES.length).setValues([['Administrateur', me, 'admin', code, 'OUI']]);
  }
  secret_(); // crée la clé de signature

  formatSheets_(ss);
  formatMoney_(ss);
  PropertiesService.getScriptProperties().setProperty('SCHEMA', String(CFG.SCHEMA));
  var report = [
    src.length + ' collectes en cours copiées dans « ' + CFG.ACTIVE + ' »',
    old.length + ' lignes d\'historique copiées dans « ' + CFG.ARCHIVE + ' »',
    'Soldes déjà collectés repris dans « ' + CFG.DONS + ' »',
    unmatched ? unmatched + ' ligne(s) sans correspondance connue : un nouvel ID leur a été attribué' : 'Toutes les lignes ont retrouvé leur ID',
    code ? 'Votre accès admin : ' + me + ' / code ' + code + ' (voir onglet MEMBRES)' : ''
  ].filter(String);
  shJ.appendRow([now, me, 'installation', report.join(' | ')]);
  return report;
}

function readSource_(sh, hasIdCol) {
  if (!sh) return [];
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  var head = vals[0].map(normKey_);
  function col(word, dflt) {
    for (var i = 0; i < head.length; i++) if (head[i].indexOf(word) >= 0) return i;
    return dflt;
  }
  var off = hasIdCol ? 1 : 0;
  var ci = {
    id: hasIdCol ? (head.indexOf('id') >= 0 ? head.indexOf('id') : 0) : -1,
    cat: col('types', off), desc: col('collectes en cours', off + 1), budget: col('budget', off + 2),
    coll: col('montant collecte', off + 3), reste: col('reste', off + 4), sit: col('situation', off + 5)
  };
  var out = [], cat = '';
  for (var r = 1; r < vals.length; r++) {
    var v = vals[r];
    if (String(v[ci.cat] || '').trim()) cat = String(v[ci.cat]).replace(/\s+$/g, '').trim();
    var desc = String(v[ci.desc] || '').trim();
    if (!desc) continue;
    out.push({
      id: ci.id >= 0 ? String(v[ci.id] || '').trim().toUpperCase() : '',
      cat: cat, desc: desc,
      budget: v[ci.budget] === '' || v[ci.budget] === null ? '' : v[ci.budget],
      coll: v[ci.coll], reste: v[ci.reste], sit: String(v[ci.sit] || '').trim()
    });
  }
  return out;
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sh;
}

function formatSheets_(ss) {
  try {
    [[CFG.ACTIVE, [90, 170, 420, 110, 110, 110, 260, 60, 100, 140, 130]],
     [CFG.ARCHIVE, [90, 170, 420, 110, 110, 110, 260, 60, 100, 100, 130]],
     [CFG.DONS, [130, 100, 90, 60, 100, 170, 130, 220, 170, 140, 70, 140]],
     [CFG.MEMBRES, [170, 230, 150, 110, 120]],
     [CFG.JOURNAL, [150, 200, 140, 500]]].forEach(function (p) {
      var sh = ss.getSheetByName(p[0]);
      if (!sh) return;
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, p[1].length).setFontWeight('bold').setBackground('#0f5c55').setFontColor('#ffffff').setWrap(true);
      p[1].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    });
    var a = ss.getSheetByName(CFG.ACTIVE), o = ss.getSheetByName(CFG.ARCHIVE), d = ss.getSheetByName(CFG.DONS);
    [a, o].forEach(function (sh) {
      sh.getRange('C2:C').setWrap(true);
      sh.getRange('G2:G').setWrap(true);
      sh.getRange('D2:F').setNumberFormat('#,##0');
      sh.getRange('I2:J').setNumberFormat('dd/mm/yyyy');
    });
    a.getRange('J2:J').setNumberFormat('dd/mm/yyyy hh:mm');
    d.getRange('B2:B').setNumberFormat('dd/mm/yyyy');
    d.getRange('E2:E').setNumberFormat('#,##0');
    d.getRange('J2:J').setNumberFormat('dd/mm/yyyy hh:mm');
    d.getRange('G2:G').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(MODES, true).setAllowInvalid(true).build());
    d.getRange('K2:K').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['OUI', ''], true).setAllowInvalid(true).build());
    var m = ss.getSheetByName(CFG.MEMBRES);
    m.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['admin', 'membre'], true).build());
    m.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['OUI', 'NON'], true).build());
  } catch (e) {
    console.warn('Mise en forme partielle : ' + e);
  }
}

/* ====================================================================== */
/*  Synchronisation feuille <-> application                                */
/* ====================================================================== */

/**
 * Remet les onglets en cohérence après une modification manuelle :
 *  - attribue ID, cycle et date d'ouverture aux nouvelles lignes de PROJETS CLAUDE
 *  - remet les formules « Montant collecté » / « Reste »
 *    (un montant tapé à la main devient une ligne d'ajustement dans DONS)
 *  - complète les dons saisis à la main dans DONS (réf, cycle, date de saisie)
 *  - génère un code d'accès pour les nouveaux membres
 */
function sync_(opts) {
  opts = opts || {};
  var ss = SpreadsheetApp.getActive();
  var shA = ss.getSheetByName(CFG.ACTIVE), shO = ss.getSheetByName(CFG.ARCHIVE), shD = ss.getSheetByName(CFG.DONS);
  if (!shA || !shO || !shD) throw new Error('Onglets manquants : lancez d\'abord « Installer ».');
  var notes = upgrade_(ss, opts.forceUpgrade);

  var A = readRows_(shA, H_ACTIVE.length);
  var O = readRows_(shO, H_ARCHIVE.length);
  var Dn = readRows_(shD, H_DONS.length);
  var formulas = A.length ? shA.getRange(2, C.COLL + 1, A.length, 2).getFormulas() : [];
  var now = new Date(), today = startOfDay_(now);
  var idEdited = {};
  (opts.idEditedRows || []).forEach(function (r) { idEdited[r] = true; });

  // cycles connus (historique + lignes actives qui ont déjà un cycle et n'ont pas changé d'ID)
  var maxCycle = {};
  function bump(id, c) { c = num_(c); if (id && c > (maxCycle[id] || 0)) maxCycle[id] = c; }
  O.forEach(function (r) { bump(cleanId_(r[C.ID]), r[C.CYCLE]); });
  A.forEach(function (r, i) { if (!idEdited[i + 2]) bump(cleanId_(r[C.ID]), r[C.CYCLE]); });

  var usedIds = {};
  O.concat(A).forEach(function (r) { var id = cleanId_(r[C.ID]); if (id) usedIds[id] = true; });

  var lastCat = '', sums = sumDonations_(Dn), adjustments = [];
  A.forEach(function (r, i) {
    var row = i + 2, changed = false;
    var desc = String(r[C.DESC] || '').trim();
    if (String(r[C.CAT] || '').trim()) lastCat = String(r[C.CAT]).trim();
    if (!desc) return;
    if (!String(r[C.CAT] || '').trim() && lastCat) { r[C.CAT] = lastCat; changed = true; }
    var id = cleanId_(r[C.ID]);
    if (id !== r[C.ID]) { r[C.ID] = id; changed = true; }
    if (!id) {
      id = nextId_(prefixFor_(r[C.CAT]), usedIds);
      r[C.ID] = id; changed = true;
      notes.push('Nouvel ID ' + id + ' attribué (ligne ' + row + ')');
    }
    if (!num_(r[C.CYCLE]) || idEdited[row]) {
      var c = (maxCycle[id] || 0) + 1;
      if (num_(r[C.CYCLE]) !== c) { r[C.CYCLE] = c; changed = true; }
      maxCycle[id] = c;
    }
    if (!r[C.OPEN]) { r[C.OPEN] = today; changed = true; }
    // Montant collecté saisi à la main => ajustement dans DONS
    var f = formulas[i] || ['', ''];
    if (String(f[0]).indexOf('DONS') < 0) {
      var typed = r[C.COLL];
      var current = sums[id + '#' + r[C.CYCLE]] || 0;
      if (typed !== '' && typed !== null && !isNaN(Number(typed)) && Math.abs(Number(typed) - current) > 0.001) {
        var delta = Number(typed) - current;
        adjustments.push([newRef_(), today, id, r[C.CYCLE], delta, '', MODE_ADJUST,
          'Montant collecté modifié directement dans la feuille (' + fmt_(current) + ' → ' + fmt_(typed) + ')',
          opts.user || 'Feuille Google', now, '', '']);
        sums[id + '#' + r[C.CYCLE]] = Number(typed);
        notes.push(id + ' : montant tapé à la main transformé en ajustement de ' + fmt_(delta) + ' Dhs');
      }
      r.needsFormula = true;
    } else if (String(f[0]).indexOf('$A' + row + argSep_()) < 0 || String(f[1]).indexOf('$D' + row + '=""' + argSep_()) < 0) {
      r.needsFormula = true;
    }
    if (changed || r.needsFormula) r[C.LAST] = now;
    if (changed) r.dirty = true;
  });

  A.forEach(function (r, i) {
    if (r.dirty) {
      shA.getRange(i + 2, 1, 1, 4).setValues([[r[C.ID], r[C.CAT], r[C.DESC], r[C.BUDGET]]]);
      shA.getRange(i + 2, C.SIT + 1, 1, 4).setValues([[r[C.SIT], r[C.CYCLE], r[C.OPEN], r[C.LAST]]]);
    }
    if (r.needsFormula) {
      setCaseFormulas_(shA, i + 2);
      if (!r.dirty) shA.getRange(i + 2, C.LAST + 1).setValue(r[C.LAST]);
    }
  });

  // Dons saisis à la main
  var activeCycle = {};
  A.forEach(function (r) { var id = cleanId_(r[C.ID]); if (id) activeCycle[id] = num_(r[C.CYCLE]); });
  Dn.forEach(function (r, i) {
    var id = cleanId_(r[D.ID]);
    if (!id && r[D.AMT] === '') return;
    var changed = false;
    if (id !== r[D.ID]) { r[D.ID] = id; changed = true; }
    if (!r[D.REF]) { r[D.REF] = newRef_(); changed = true; }
    if (!num_(r[D.CYCLE]) && id) {
      if (activeCycle[id]) { r[D.CYCLE] = activeCycle[id]; changed = true; }
      else if (maxCycle[id]) {
        r[D.CYCLE] = maxCycle[id]; changed = true;
        notes.push('Don ligne ' + (i + 2) + ' : ' + id + ' n\'est pas en cours, rattaché à son dernier cycle');
      } else notes.push('Don ligne ' + (i + 2) + ' : ID « ' + id + ' » inconnu');
    }
    if (!r[D.DATE]) { r[D.DATE] = today; changed = true; }
    if (!r[D.AT]) { r[D.AT] = now; changed = true; }
    if (!r[D.BY]) { r[D.BY] = opts.user || 'Feuille Google'; changed = true; }
    if (changed) shD.getRange(i + 2, 1, 1, H_DONS.length).setValues([r.slice(0, H_DONS.length)]);
  });
  if (adjustments.length) {
    shD.getRange(shD.getLastRow() + 1, 1, adjustments.length, H_DONS.length).setValues(adjustments);
  }

  // Membres sans code
  var shM = ss.getSheetByName(CFG.MEMBRES);
  if (shM) {
    readRows_(shM, H_MEMBRES.length).forEach(function (r, i) {
      if (String(r[M.EMAIL] || '').trim() && !String(r[M.CODE] || '').trim()) {
        var upd = [randomCode_(), r[M.ACTIVE] || 'OUI'];
        shM.getRange(i + 2, M.CODE + 1, 1, 2).setValues([upd]);
        if (!r[M.ROLE]) shM.getRange(i + 2, M.ROLE + 1).setValue('membre');
      }
    });
  }
  return { notes: notes };
}

function setCaseFormulas_(sh, row) {
  var x = argSep_();
  sh.getRange(row, C.COLL + 1, 1, 2).setFormulas([[
    '=SUMIFS(DONS!$E:$E' + x + 'DONS!$C:$C' + x + '$A' + row + x + 'DONS!$D:$D' + x + '$H' + row + x + 'DONS!$K:$K' + x + '"<>OUI")',
    '=IF(OR($D' + row + '=""' + x + '$E' + row + '>=$D' + row + ')' + x + '"-"' + x + '$D' + row + '-$E' + row + ')'
  ]]);
}

/**
 * Séparateur d'arguments des formules : « , » (anglais) ou « ; » (français et autres langues
 * à virgule décimale). Détecté une fois par paramètre régional du classeur, en testant une
 * formule dans une cellule libre de JOURNAL APP, puis mémorisé.
 */
var ARG_SEP_CACHE_ = null;
function argSep_() {
  if (ARG_SEP_CACHE_) return ARG_SEP_CACHE_;
  var ss = SpreadsheetApp.getActive();
  var locale = '';
  try { locale = ss.getSpreadsheetLocale(); } catch (e) { /* ignore */ }
  var props = PropertiesService.getScriptProperties();
  var key = 'ARG_SEP_' + locale;
  var sep = props.getProperty(key);
  if (!sep) {
    var sh = ss.getSheetByName(CFG.JOURNAL) || ss.getSheets()[0];
    var cell = sh.getRange(1, 26); // colonne Z, en dehors des données
    cell.setFormula('=SUM(1,2)');
    SpreadsheetApp.flush();
    sep = String(cell.getDisplayValue()) === '3' ? ',' : ';';
    cell.clearContent();
    props.setProperty(key, sep);
  }
  ARG_SEP_CACHE_ = sep;
  return sep;
}

/* ====================================================================== */
/*  API web (appelée par l'application)                                    */
/* ====================================================================== */

function doGet() {
  return json_({ ok: true, service: 'Gate of Hope API', version: CFG.VERSION });
}

function doPost(e) {
  var out;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = { ok: true, data: route_(body) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err), code: err && err.code || 'ERROR' };
  }
  return json_(out);
}

function route_(b) {
  var action = String(b.action || '');
  if (action === 'login') return login_(b.email, b.code);
  var user = auth_(b.token);
  var admin = function () { if (user.role !== 'admin') throw apiError_('Réservé aux administrateurs.', 'FORBIDDEN'); };
  switch (action) {
    case 'me': return { user: user };
    case 'list': return withLock_(function () { sync_({ user: user.email }); return list_(user); });
    case 'donations': admin(); return donations_(b.id);
    case 'recentDonations': admin(); return recentDonations_(b.limit || 60);
    case 'addDonation': admin(); return withLock_(function () { sync_({ user: user.email }); return addDonation_(b, user); });
    case 'cancelDonation': admin(); return withLock_(function () { return cancelDonation_(b.ref, b.reason, user); });
    case 'movements': admin(); return movements_(b.id);
    case 'addMovement': admin(); return withLock_(function () { return addMovement_(b, user); });
    case 'cancelMovement': admin(); return withLock_(function () { return cancelMovement_(b.ref, b.reason, user); });
    case 'createCase': admin(); return withLock_(function () { sync_({ user: user.email }); return createCase_(b, user); });
    case 'updateCase': admin(); return withLock_(function () { sync_({ user: user.email }); return updateCase_(b, user); });
    case 'closeCase': admin(); return withLock_(function () {
      sync_({ user: user.email });
      var r = closeCase_(cleanId_(b.id), b.situation, user.email);
      return r;
    });
    default: throw apiError_('Action inconnue : ' + action, 'BAD_REQUEST');
  }
}

/* ---------- Authentification ---------- */

function login_(email, code) {
  email = String(email || '').trim().toLowerCase();
  code = String(code || '').trim();
  if (!email || !code) throw apiError_('Email et code requis.', 'BAD_REQUEST');
  var cache = CacheService.getScriptCache();
  var key = 'login:' + email;
  var tries = Number(cache.get(key) || 0);
  if (tries >= 8) throw apiError_('Trop de tentatives. Réessayez dans 15 minutes.', 'RATE_LIMIT');
  var m = findMember_(email);
  if (!m || m.code !== code || !m.active) {
    cache.put(key, String(tries + 1), 900);
    throw apiError_('Email ou code incorrect.', 'BAD_LOGIN');
  }
  cache.remove(key);
  var payload = { e: email, x: Date.now() + CFG.TOKEN_DAYS * 864e5 };
  var p = Utilities.base64EncodeWebSafe(JSON.stringify(payload));
  var token = p + '.' + sign_(p);
  log_(email, 'connexion', '');
  return { token: token, user: { email: email, name: m.name, role: m.role } };
}

function auth_(token) {
  var parts = String(token || '').split('.');
  if (parts.length !== 2 || sign_(parts[0]) !== parts[1]) throw apiError_('Session invalide, reconnectez-vous.', 'AUTH');
  var payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
  if (!payload.x || payload.x < Date.now()) throw apiError_('Session expirée, reconnectez-vous.', 'AUTH');
  var m = findMember_(payload.e);
  if (!m || !m.active) throw apiError_('Accès retiré.', 'AUTH');
  return { email: m.email, name: m.name, role: m.role };
}

function findMember_(email) {
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.MEMBRES);
  if (!sh) return null;
  var rows = readRows_(sh, H_MEMBRES.length);
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[M.EMAIL] || '').trim().toLowerCase() === String(email).toLowerCase()) {
      return {
        email: String(r[M.EMAIL]).trim().toLowerCase(),
        name: String(r[M.NAME] || '').trim() || String(r[M.EMAIL]).split('@')[0],
        role: String(r[M.ROLE] || '').trim().toLowerCase() === 'admin' ? 'admin' : 'membre',
        code: String(r[M.CODE] || '').trim(),
        active: String(r[M.ACTIVE] || 'OUI').trim().toUpperCase() !== 'NON'
      };
    }
  }
  return null;
}

function secret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('TOKEN_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('TOKEN_SECRET', s); }
  return s;
}

function sign_(s) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, secret_()));
}

/* ---------- Lecture ---------- */

function list_(user) {
  var ss = SpreadsheetApp.getActive();
  var A = readRows_(ss.getSheetByName(CFG.ACTIVE), H_ACTIVE.length);
  var O = readRows_(ss.getSheetByName(CFG.ARCHIVE), H_ARCHIVE.length);
  var Dn = readRows_(ss.getSheetByName(CFG.DONS), H_DONS.length);
  var sums = sumDonations_(Dn);
  var lastDon = {};
  Dn.forEach(function (r) {
    if (isCancelled_(r)) return;
    var k = cleanId_(r[D.ID]) + '#' + num_(r[D.CYCLE]);
    var t = toDate_(r[D.AT]);
    if (t && (!lastDon[k] || t > lastDon[k])) lastDon[k] = t;
  });

  var cases = A.filter(function (r) { return String(r[C.DESC] || '').trim() && r[C.ID]; }).map(function (r) {
    var id = cleanId_(r[C.ID]), cyc = num_(r[C.CYCLE]);
    var budget = r[C.BUDGET] === '' ? null : num_(r[C.BUDGET]);
    var collected = sums[id + '#' + cyc] || 0;
    return {
      id: id, cycle: cyc, category: String(r[C.CAT] || '').trim(), description: String(r[C.DESC]).trim(),
      budget: budget, collected: collected,
      remaining: budget === null ? null : Math.max(0, budget - collected),
      situation: String(r[C.SIT] || '').trim(), responsable: String(r[C.RESP] || '').trim(),
      openedAt: iso_(r[C.OPEN]), updatedAt: iso_(maxDate_(toDate_(r[C.LAST]), lastDon[id + '#' + cyc]))
    };
  });

  var closed = O.filter(function (r) { return String(r[C.DESC] || '').trim(); }).map(function (r) {
    return {
      id: cleanId_(r[C.ID]), cycle: num_(r[C.CYCLE]), category: String(r[C.CAT] || '').trim(),
      description: String(r[C.DESC]).trim(),
      budget: r[C.BUDGET] === '' ? null : num_(r[C.BUDGET]), collected: num_(r[C.COLL]),
      situation: String(r[C.SIT] || '').trim(), responsable: String(r[C.RESP] || '').trim(),
      openedAt: iso_(r[C.OPEN]), closedAt: iso_(r[C.LAST])
    };
  });

  var monthStart = new Date(); monthStart.setDate(1); monthStart = startOfDay_(monthStart);
  var month = 0, monthCount = 0;
  Dn.forEach(function (r) {
    if (isCancelled_(r) || r[D.MODE] === MODE_OPENING || r[D.MODE] === MODE_ADJUST) return;
    var d = toDate_(r[D.DATE]);
    if (d && d >= monthStart) { month += num_(r[D.AMT]); monthCount++; }
  });
  var stats = {
    activeCount: cases.length, closedCount: closed.length,
    collected: cases.reduce(function (s, c) { return s + c.collected; }, 0),
    remaining: cases.reduce(function (s, c) { return s + (c.remaining || 0); }, 0),
    monthCollected: month, monthCount: monthCount
  };
  var cats = {};
  cases.concat(closed).forEach(function (c) { if (c.category) cats[c.category] = true; });
  var out = { user: user, cases: cases, closed: closed, categories: Object.keys(cats), stats: stats, serverTime: iso_(new Date()) };
  if (user) {
    var mem = listMembers_();
    out.money = money_(user, mem, Dn, cases, closed);
    if (user.role === 'admin') out.members = mem.filter(function (m) { return m.active; }).map(function (m) { return { name: m.name, email: m.email }; });
  }
  return out;
}

function donations_(id) {
  id = cleanId_(id);
  var Dn = readRows_(SpreadsheetApp.getActive().getSheetByName(CFG.DONS), H_DONS.length);
  return Dn.filter(function (r) { return cleanId_(r[D.ID]) === id; }).map(donObj_).reverse();
}

function recentDonations_(limit) {
  var Dn = readRows_(SpreadsheetApp.getActive().getSheetByName(CFG.DONS), H_DONS.length);
  return Dn.filter(function (r) { return r[D.MODE] !== MODE_OPENING && r[D.ID]; })
    .map(donObj_)
    .sort(function (a, b) { return String(b.enteredAt).localeCompare(String(a.enteredAt)); })
    .slice(0, Math.min(Number(limit) || 60, 300));
}

function donObj_(r) {
  return {
    ref: String(r[D.REF]), date: iso_(r[D.DATE]), id: cleanId_(r[D.ID]), cycle: num_(r[D.CYCLE]),
    amount: num_(r[D.AMT]), donor: String(r[D.DONOR] || ''), mode: String(r[D.MODE] || ''),
    note: String(r[D.NOTE] || ''), by: String(r[D.BY] || ''), enteredAt: iso_(r[D.AT]), cancelled: isCancelled_(r),
    holder: String(r[D.HOLDER] || '').trim()
  };
}

/* ---------- Écriture ---------- */

function findActive_(id) {
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.ACTIVE);
  var rows = readRows_(sh, H_ACTIVE.length);
  for (var i = 0; i < rows.length; i++) if (cleanId_(rows[i][C.ID]) === id) return { sh: sh, row: i + 2, r: rows[i] };
  return null;
}

function addDonation_(b, user) {
  var id = cleanId_(b.id);
  var amount = Number(String(b.amount).replace(/\s/g, '').replace(',', '.'));
  if (!(amount > 0)) throw apiError_('Montant invalide.', 'BAD_REQUEST');
  var f = findActive_(id);
  if (!f) throw apiError_('Collecte ' + id + ' introuvable parmi les collectes en cours.', 'NOT_FOUND');
  var date = b.date ? parseDay_(b.date) : startOfDay_(new Date());
  var mode = MODES.indexOf(b.mode) >= 0 && b.mode !== MODE_OPENING ? b.mode : (b.mode ? 'Autre' : 'Espèces');
  // Qui a l'argent : la personne choisie, sinon celle qui saisit le don
  var holder = b.holder ? memberName_(b.holder, listMembers_(), true) : user.name;
  var row = [newRef_(), date, id, num_(f.r[C.CYCLE]), amount, String(b.donor || '').trim().slice(0, 120) || 'Anonyme',
    mode, String(b.note || '').trim().slice(0, 500), user.email, new Date(), '', holder];
  var shD = SpreadsheetApp.getActive().getSheetByName(CFG.DONS);
  shD.getRange(shD.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  f.sh.getRange(f.row, C.LAST + 1).setValue(new Date());
  log_(user.email, 'don', id + ' +' + fmt_(amount) + ' Dhs, chez ' + holder + ' (' + row[0] + ')');
  return { ref: row[0], case: caseById_(id) };
}

function cancelDonation_(ref, reason, user) {
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.DONS);
  var rows = readRows_(sh, H_DONS.length);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][D.REF]) === String(ref)) {
      if (rows[i][D.MODE] === MODE_OPENING) throw apiError_('Le solde repris de l\'Excel ne peut pas être annulé ici : corrigez-le dans la feuille.', 'BAD_REQUEST');
      sh.getRange(i + 2, D.CANCEL + 1).setValue('OUI');
      var note = String(rows[i][D.NOTE] || '');
      sh.getRange(i + 2, D.NOTE + 1).setValue((note ? note + ' | ' : '') + 'Annulé par ' + user.email + (reason ? ' : ' + reason : ''));
      log_(user.email, 'annulation don', ref + ' (' + rows[i][D.ID] + ', ' + fmt_(rows[i][D.AMT]) + ' Dhs)');
      var f = findActive_(cleanId_(rows[i][D.ID]));
      if (f) f.sh.getRange(f.row, C.LAST + 1).setValue(new Date());
      return { ref: ref };
    }
  }
  throw apiError_('Don introuvable.', 'NOT_FOUND');
}

function createCase_(b, user) {
  var ss = SpreadsheetApp.getActive();
  var shA = ss.getSheetByName(CFG.ACTIVE), shO = ss.getSheetByName(CFG.ARCHIVE);
  var cat = String(b.category || '').trim(), desc = String(b.description || '').trim();
  if (!cat || !desc) throw apiError_('Type et description obligatoires.', 'BAD_REQUEST');
  var A = readRows_(shA, H_ACTIVE.length), O = readRows_(shO, H_ARCHIVE.length);
  var id = cleanId_(b.reuseId), cycle = 1;
  if (id) {
    if (A.some(function (r) { return cleanId_(r[C.ID]) === id; })) throw apiError_(id + ' est déjà en cours.', 'BAD_REQUEST');
    var known = false;
    O.forEach(function (r) { if (cleanId_(r[C.ID]) === id) { known = true; cycle = Math.max(cycle, num_(r[C.CYCLE]) + 1); } });
    if (!known) throw apiError_('ID ' + id + ' inconnu dans l\'historique.', 'BAD_REQUEST');
  } else {
    var used = {};
    A.concat(O).forEach(function (r) { var x = cleanId_(r[C.ID]); if (x) used[x] = true; });
    id = nextId_(prefixFor_(cat), used);
  }
  var budget = b.budget === '' || b.budget === null || b.budget === undefined ? '' : Number(b.budget);
  if (budget !== '' && !(budget >= 0)) throw apiError_('Budget invalide.', 'BAD_REQUEST');
  // insertion après la dernière ligne de la même catégorie (sinon en bas)
  var at = A.length + 2;
  for (var i = A.length - 1; i >= 0; i--) {
    if (String(A[i][C.CAT]).trim().toLowerCase() === cat.toLowerCase()) { at = i + 3; break; }
  }
  var resp = b.responsable ? memberName_(b.responsable, listMembers_(), true) : (user && user.name) || '';
  if (at <= A.length + 1) shA.insertRowBefore(at);
  var now = new Date();
  shA.getRange(at, 1, 1, H_ACTIVE.length).setValues([[id, cat, desc, budget, '', '', String(b.situation || '').trim(), cycle, startOfDay_(now), now, resp]]);
  // les formules des lignes décalées sont ajustées automatiquement par Sheets ; on pose celles de la nouvelle ligne
  setCaseFormulas_(shA, at);
  log_(user.email, 'nouvelle collecte', id + ' (cycle ' + cycle + ') ' + desc.slice(0, 80));
  return { case: caseById_(id) };
}

function updateCase_(b, user) {
  var id = cleanId_(b.id);
  var f = findActive_(id);
  if (!f) throw apiError_('Collecte introuvable.', 'NOT_FOUND');
  var r = f.r, changes = [];
  function set(col, key, val) {
    if (b[key] === undefined) return;
    if (String(r[col]) !== String(val)) { changes.push(key); r[col] = val; }
  }
  if (b.category !== undefined && !String(b.category).trim()) throw apiError_('Type obligatoire.', 'BAD_REQUEST');
  if (b.description !== undefined && !String(b.description).trim()) throw apiError_('Description obligatoire.', 'BAD_REQUEST');
  set(C.CAT, 'category', String(b.category || '').trim());
  set(C.DESC, 'description', String(b.description || '').trim());
  if (b.budget !== undefined) {
    var bud = b.budget === '' || b.budget === null ? '' : Number(b.budget);
    if (bud !== '' && !(bud >= 0)) throw apiError_('Budget invalide.', 'BAD_REQUEST');
    set(C.BUDGET, 'budget', bud);
  }
  set(C.SIT, 'situation', String(b.situation || '').trim());
  if (b.responsable !== undefined) set(C.RESP, 'responsable', b.responsable ? memberName_(b.responsable, listMembers_(), true) : '');
  if (changes.length) {
    f.sh.getRange(f.row, 1, 1, 4).setValues([[r[C.ID], r[C.CAT], r[C.DESC], r[C.BUDGET]]]);
    f.sh.getRange(f.row, C.SIT + 1).setValue(r[C.SIT]);
    f.sh.getRange(f.row, C.RESP + 1).setValue(r[C.RESP]);
    f.sh.getRange(f.row, C.LAST + 1).setValue(new Date());
    log_(user.email, 'modification', id + ' : ' + changes.join(', '));
  }
  return { case: caseById_(id) };
}

/** Déplace une collecte en cours vers PROJETS.OLD CLAUDE (valeurs figées). */
function closeCase_(id, situation, who) {
  var ss = SpreadsheetApp.getActive();
  var f = findActive_(id);
  if (!f) throw apiError_('Collecte ' + id + ' introuvable parmi les collectes en cours.', 'NOT_FOUND');
  var r = f.r;
  var Dn = readRows_(ss.getSheetByName(CFG.DONS), H_DONS.length);
  var collected = sumDonations_(Dn)[id + '#' + num_(r[C.CYCLE])] || 0;
  var budget = r[C.BUDGET] === '' ? '' : num_(r[C.BUDGET]);
  var reste = budget === '' || collected >= budget ? '-' : budget - collected;
  var sit = String(situation || '').trim() || (String(r[C.SIT] || '').trim() ? String(r[C.SIT]).trim() + ' => Collecte clôturée' : 'Collecte clôturée');
  var shO = ss.getSheetByName(CFG.ARCHIVE);
  shO.insertRowBefore(2);
  shO.getRange(2, 1, 1, H_ARCHIVE.length).setValues([[id, r[C.CAT], r[C.DESC], budget, collected, reste, sit,
    num_(r[C.CYCLE]), r[C.OPEN] || '', startOfDay_(new Date()), r[C.RESP] || '']]);
  shO.getRange(2, 1, 1, H_ARCHIVE.length).setFontWeight('normal').setBackground(null).setFontColor(null);
  if (f.sh.getMaxRows() <= 2) f.sh.getRange(f.row, 1, 1, H_ACTIVE.length).clearContent();
  else f.sh.deleteRow(f.row);
  log_(who, 'clôture', id + ' cycle ' + num_(r[C.CYCLE]) + ' — ' + fmt_(collected) + ' Dhs collectés');
  // l'argent encore détenu reste visible dans le profil de chacun (avertissement seulement)
  var held = [];
  try {
    var bal = balances_(Dn, readRows_(ss.getSheetByName(CFG.MOUVEMENTS), H_MVT.length), listMembers_()).bal;
    var k = id + '#' + num_(r[C.CYCLE]);
    Object.keys(bal).forEach(function (h) { if ((bal[h][k] || 0) > 0.005) held.push({ name: h, amount: round2_(bal[h][k]) }); });
  } catch (e) { console.warn(e); }
  return { id: id, collected: collected, held: held };
}

function caseById_(id) {
  var l = list_(null);
  for (var i = 0; i < l.cases.length; i++) if (l.cases[i].id === id) return l.cases[i];
  return null;
}

/* ====================================================================== */
/*  Suivi de l'argent : qui a quoi, pour quel cas                          */
/* ====================================================================== */

/**
 * Mise à niveau d'une installation existante (v1.0) : ajoute à DROITE les colonnes
 * « Responsable » et « Qui a l'argent » et crée l'onglet MOUVEMENTS. Aucune donnée
 * existante n'est déplacée ; les formules ne lisent pas ces nouvelles colonnes.
 * Ne s'exécute qu'une fois (sauf via le menu « Vérifier / réparer »).
 */
function upgrade_(ss, force) {
  var props = PropertiesService.getScriptProperties();
  if (!force && Number(props.getProperty('SCHEMA') || 0) >= CFG.SCHEMA) return [];
  var notes = [];
  [[CFG.ACTIVE, H_ACTIVE], [CFG.ARCHIVE, H_ARCHIVE], [CFG.DONS, H_DONS]].forEach(function (p) {
    var sh = ss.getSheetByName(p[0]);
    if (!sh) return;
    var col = p[1].length, want = p[1][col - 1];
    var cell = sh.getRange(1, col), cur = String(cell.getValues()[0][0] || '').trim();
    if (!cur) {
      cell.setValue(want);
      notes.push('Colonne « ' + want + ' » ajoutée dans « ' + p[0] + ' »');
    } else if (cur !== want) {
      notes.push('Attention : la colonne ' + col + ' de « ' + p[0] + ' » contient « ' + cur + ' » au lieu de « ' + want + ' »');
    }
  });
  if (!ss.getSheetByName(CFG.MOUVEMENTS)) {
    ensureSheet_(ss, CFG.MOUVEMENTS, H_MVT);
    notes.push('Onglet « ' + CFG.MOUVEMENTS + ' » créé');
  }
  formatMoney_(ss);
  props.setProperty('SCHEMA', String(CFG.SCHEMA));
  return notes;
}

function formatMoney_(ss) {
  try {
    var mem = ss.getSheetByName(CFG.MEMBRES);
    var names = mem ? SpreadsheetApp.newDataValidation().requireValueInRange(mem.getRange('A2:A'), true).setAllowInvalid(true).build() : null;
    var head = function (sh, col) {
      sh.getRange(1, col).setFontWeight('bold').setBackground('#0f5c55').setFontColor('#ffffff').setWrap(true);
      sh.setColumnWidth(col, 140);
    };
    var a = ss.getSheetByName(CFG.ACTIVE), o = ss.getSheetByName(CFG.ARCHIVE), d = ss.getSheetByName(CFG.DONS);
    if (a) { head(a, C.RESP + 1); if (names) a.getRange('K2:K').setDataValidation(names); }
    if (o) head(o, C.RESP + 1);
    if (d) { head(d, D.HOLDER + 1); if (names) d.getRange('L2:L').setDataValidation(names); }
    var mv = ss.getSheetByName(CFG.MOUVEMENTS);
    if (mv) {
      mv.setFrozenRows(1);
      mv.getRange(1, 1, 1, H_MVT.length).setFontWeight('bold').setBackground('#0f5c55').setFontColor('#ffffff').setWrap(true);
      [130, 100, 90, 60, 100, 150, 150, 220, 170, 140, 70].forEach(function (w, i) { mv.setColumnWidth(i + 1, w); });
      mv.getRange('B2:B').setNumberFormat('dd/mm/yyyy');
      mv.getRange('E2:E').setNumberFormat('#,##0');
      mv.getRange('J2:J').setNumberFormat('dd/mm/yyyy hh:mm');
      mv.getRange('K2:K').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['OUI', ''], true).setAllowInvalid(true).build());
    }
  } catch (e) {
    console.warn('Mise en forme partielle : ' + e);
  }
}

function listMembers_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.MEMBRES);
  return readRows_(sh, H_MEMBRES.length).filter(function (r) { return String(r[M.EMAIL] || r[M.NAME] || '').trim(); }).map(function (r) {
    var email = String(r[M.EMAIL] || '').trim().toLowerCase();
    return {
      name: String(r[M.NAME] || '').trim() || email.split('@')[0], email: email,
      active: String(r[M.ACTIVE] || 'OUI').trim().toUpperCase() !== 'NON'
    };
  });
}

/**
 * Retrouve le nom d'un membre à partir de son nom ou de son email (sans tenir compte des
 * majuscules / accents). Vide => « Non attribué ». Inconnu => tel quel, ou erreur si strict.
 */
function memberName_(v, mem, strict) {
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s) { if (strict) throw apiError_('Indiquez qui a l\'argent.', 'BAD_REQUEST'); return UNASSIGNED; }
  var k = normKey_(s);
  for (var i = 0; i < mem.length; i++) {
    if (normKey_(mem[i].name) === k || mem[i].email === s.toLowerCase()) return mem[i].name;
  }
  if (k === normKey_(UNASSIGNED)) return UNASSIGNED;
  if (strict) throw apiError_('« ' + s + ' » n\'est pas dans l\'onglet ' + CFG.MEMBRES + '.', 'BAD_REQUEST');
  return s;
}

function isBenef_(name) { return normKey_(name).indexOf('beneficiaire') === 0; }
function isCancelledMv_(r) { return String(r[MV.CANCEL] || '').trim().toUpperCase() === 'OUI'; }
function round2_(n) { return Math.round(n * 100) / 100; }

/** Soldes : bal[personne][« ID#cycle »] et given[« ID#cycle »] (remis au bénéficiaire). */
function balances_(Dn, Mv, mem) {
  var bal = {}, given = {};
  function add(h, k, a) { bal[h] = bal[h] || {}; bal[h][k] = (bal[h][k] || 0) + a; }
  Dn.forEach(function (r) {
    if (isCancelled_(r)) return;
    var id = cleanId_(r[D.ID]);
    if (!id || r[D.AMT] === '') return;
    add(memberName_(r[D.HOLDER], mem), id + '#' + num_(r[D.CYCLE]), num_(r[D.AMT]));
  });
  Mv.forEach(function (r) {
    if (isCancelledMv_(r)) return;
    var id = cleanId_(r[MV.ID]);
    if (!id) return;
    var k = id + '#' + num_(r[MV.CYCLE]), a = num_(r[MV.AMT]);
    add(memberName_(r[MV.FROM], mem), k, -a);
    var to = String(r[MV.TO] || '').trim();
    if (isBenef_(to)) given[k] = (given[k] || 0) + a;
    else add(memberName_(to, mem), k, a);
  });
  return { bal: bal, given: given };
}

/** Bloc « money » renvoyé avec la liste : tout pour un admin, seulement son propre argent pour un membre. */
function money_(user, mem, Dn, cases, closed) {
  var Mv = readRows_(SpreadsheetApp.getActive().getSheetByName(CFG.MOUVEMENTS), H_MVT.length);
  var b = balances_(Dn, Mv, mem);
  var info = {};
  closed.forEach(function (c) { info[c.id + '#' + c.cycle] = c; });
  cases.forEach(function (c) { info[c.id + '#' + c.cycle] = c; });
  var me = memberName_(user.email, mem);
  var admin = user.role === 'admin';
  var byCase = {};
  function caseInfo(k) {
    var c = info[k], p = k.split('#');
    return { id: p[0], cycle: Number(p[1]), description: c ? c.description : '', responsable: c ? c.responsable : '', closed: !c || !!c.closedAt };
  }
  var holders = Object.keys(b.bal).map(function (h) {
    var lines = Object.keys(b.bal[h]).filter(function (k) { return Math.abs(b.bal[h][k]) > 0.005; }).map(function (k) {
      var ci = caseInfo(k);
      ci.amount = round2_(b.bal[h][k]);
      if (admin) {
        byCase[k] = byCase[k] || Object.assign(caseInfo(k), { held: [], given: 0 });
        byCase[k].held.push({ name: h, amount: ci.amount });
      }
      return ci;
    });
    var total = round2_(lines.reduce(function (s, l) { return s + l.amount; }, 0));
    return { name: h, total: total, lines: lines };
  }).filter(function (h) { return h.lines.length && (admin || h.name === me); });
  if (admin) {
    Object.keys(b.given).forEach(function (k) {
      if (b.given[k] <= 0.005) return;
      byCase[k] = byCase[k] || Object.assign(caseInfo(k), { held: [], given: 0 });
      byCase[k].given = round2_(b.given[k]);
    });
  }
  holders.sort(function (x, y) {
    if (x.name === UNASSIGNED) return 1;
    if (y.name === UNASSIGNED) return -1;
    return y.total - x.total;
  });
  return { me: me, unassigned: UNASSIGNED, beneficiary: BENEF, holders: holders, byCase: admin ? byCase : null };
}

function movements_(id) {
  id = cleanId_(id);
  var Mv = readRows_(SpreadsheetApp.getActive().getSheetByName(CFG.MOUVEMENTS), H_MVT.length);
  return Mv.filter(function (r) { return cleanId_(r[MV.ID]) === id; }).map(function (r) {
    return {
      ref: String(r[MV.REF]), date: iso_(r[MV.DATE]), id: cleanId_(r[MV.ID]), cycle: num_(r[MV.CYCLE]),
      amount: num_(r[MV.AMT]), from: String(r[MV.FROM] || '').trim() || UNASSIGNED, to: String(r[MV.TO] || '').trim(),
      note: String(r[MV.NOTE] || ''), by: String(r[MV.BY] || ''), enteredAt: iso_(r[MV.AT]), cancelled: isCancelledMv_(r)
    };
  }).reverse();
}

/** Argent qui change de main : d'un membre à un autre, ou remis au bénéficiaire. */
function addMovement_(b, user) {
  var ss = SpreadsheetApp.getActive();
  var mem = listMembers_();
  var id = cleanId_(b.id);
  var amount = Number(String(b.amount).replace(/\s/g, '').replace(',', '.'));
  if (!(amount > 0)) throw apiError_('Montant invalide.', 'BAD_REQUEST');
  var from = normKey_(b.from) === normKey_(UNASSIGNED) ? UNASSIGNED : memberName_(b.from, mem, true);
  var to = isBenef_(b.to) ? BENEF : memberName_(b.to, mem, true);
  if (from === to) throw apiError_('La personne qui donne et celle qui reçoit sont les mêmes.', 'BAD_REQUEST');

  // cycle : celui indiqué, sinon la collecte en cours, sinon la dernière clôturée
  var cycle = num_(b.cycle);
  if (!cycle) {
    var f = findActive_(id);
    if (f) cycle = num_(f.r[C.CYCLE]);
    else readRows_(ss.getSheetByName(CFG.ARCHIVE), H_ARCHIVE.length).forEach(function (r) {
      if (cleanId_(r[C.ID]) === id) cycle = Math.max(cycle, num_(r[C.CYCLE]));
    });
  }
  if (!id || !cycle) throw apiError_('Collecte ' + id + ' introuvable.', 'NOT_FOUND');

  var shMv = ss.getSheetByName(CFG.MOUVEMENTS);
  if (!shMv) shMv = ensureSheet_(ss, CFG.MOUVEMENTS, H_MVT);
  var Dn = readRows_(ss.getSheetByName(CFG.DONS), H_DONS.length);
  var bal = balances_(Dn, readRows_(shMv, H_MVT.length), mem).bal;
  var k = id + '#' + cycle, has = (bal[from] && bal[from][k]) || 0;
  if (amount > has + 0.005) {
    throw apiError_(from + ' ne détient que ' + fmt_(Math.max(0, has)) + ' Dhs pour ' + id + '.', 'BAD_REQUEST');
  }
  var date = b.date ? parseDay_(b.date) : startOfDay_(new Date());
  var row = [newRef_('MVT'), date, id, cycle, amount, from, to, String(b.note || '').trim().slice(0, 500), user.email, new Date(), ''];
  shMv.getRange(shMv.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  log_(user.email, to === BENEF ? 'remise bénéficiaire' : 'transfert',
    id + ' : ' + fmt_(amount) + ' Dhs de ' + from + ' vers ' + to + ' (' + row[0] + ')');
  return { ref: row[0] };
}

function cancelMovement_(ref, reason, user) {
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.MOUVEMENTS);
  var rows = readRows_(sh, H_MVT.length);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][MV.REF]) === String(ref)) {
      sh.getRange(i + 2, MV.CANCEL + 1).setValue('OUI');
      var note = String(rows[i][MV.NOTE] || '');
      sh.getRange(i + 2, MV.NOTE + 1).setValue((note ? note + ' | ' : '') + 'Annulé par ' + user.email + (reason ? ' : ' + reason : ''));
      log_(user.email, 'annulation mouvement', ref + ' (' + rows[i][MV.ID] + ', ' + fmt_(rows[i][MV.AMT]) + ' Dhs)');
      return { ref: ref };
    }
  }
  throw apiError_('Mouvement introuvable.', 'NOT_FOUND');
}

/* ====================================================================== */
/*  Utilitaires                                                            */
/* ====================================================================== */

function readRows_(sh, width) {
  if (!sh) return [];
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, width).getValues();
}

function sumDonations_(Dn) {
  var s = {};
  Dn.forEach(function (r) {
    if (isCancelled_(r)) return;
    var id = cleanId_(r[D.ID]);
    if (!id) return;
    var k = id + '#' + num_(r[D.CYCLE]);
    s[k] = (s[k] || 0) + num_(r[D.AMT]);
  });
  return s;
}

function isCancelled_(r) { return String(r[D.CANCEL] || '').trim().toUpperCase() === 'OUI'; }

function nextId_(prefix, used) {
  var max = 0;
  Object.keys(used).forEach(function (id) {
    var m = id.match(/^([A-Z]+)-(\d+)$/);
    if (m && m[1] === prefix) max = Math.max(max, Number(m[2]));
  });
  var id = prefix + '-' + ('00' + (max + 1)).slice(-3);
  used[id] = true;
  return id;
}

function prefixFor_(cat) {
  var n = normKey_(cat);
  for (var i = 0; i < PREFIXES.length; i++) if (n.indexOf(PREFIXES[i][0]) >= 0) return PREFIXES[i][1];
  return 'GOH';
}

/** Même normalisation que celle utilisée pour calculer MIGRATION_MAP. */
function normKey_(s) {
  return String(s === null || s === undefined ? '' : s).normalize('NFD').toLowerCase()
    .replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();
}

/** Empreinte courte (16 caractères hexadécimaux du SHA-256) du texte normalisé. */
function keyHash_(s) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, normKey_(s), Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').slice(0, 16);
}

function cleanId_(v) { return String(v === null || v === undefined ? '' : v).trim().toUpperCase(); }

function num_(v) {
  if (typeof v === 'number') return v;
  var n = Number(String(v === null || v === undefined ? '' : v).replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function fmt_(n) { return String(Math.round(num_(n) * 100) / 100); }

function toDate_(v) {
  if (v instanceof Date) return v;
  if (!v) return null;
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function maxDate_(a, b) { return !a ? b : !b ? a : (a > b ? a : b); }

function iso_(v) { var d = toDate_(v); return d ? d.toISOString() : null; }

function startOfDay_(d) { var x = new Date(d); x.setHours(0, 0, 0, 0); return x; }

function parseDay_(s) {
  var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) throw apiError_('Date invalide.', 'BAD_REQUEST');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function newRef_(prefix) {
  var d = new Date();
  var stamp = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
  return (prefix || 'DON') + '-' + stamp + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
}

function randomCode_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = '';
  for (var i = 0; i < 8; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

function currentUserEmail_() {
  try { return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail(); } catch (e) { return ''; }
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { SpreadsheetApp.flush(); lock.releaseLock(); }
}

function log_(user, action, details) {
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.JOURNAL);
  if (sh) sh.appendRow([new Date(), user || '', action, details || '']);
}

function apiError_(msg, code) { var e = new Error(msg); e.code = code; return e; }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

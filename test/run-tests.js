// Tests de bout en bout du script Apps Script sur un classeur émulé.
//   node test/run-tests.js                      -> données fictives (test/fixture-demo.json)
//   FIXTURE=chemin/vers/classeur.json node test/run-tests.js   -> vraies données exportées
'use strict';
const fs = require('fs');
const path = require('path');
const assert = require('assert');
const { createContext } = require('./gas-mock');

const fixture = JSON.parse(fs.readFileSync(process.env.FIXTURE || path.join(__dirname, 'fixture-demo.json'), 'utf8'));
const expected = process.env.EXPECTED ? JSON.parse(fs.readFileSync(process.env.EXPECTED, 'utf8')) : null;
const g = createContext(fixture, { quiet: true, locale: process.env.LOCALE || 'fr_FR' });
console.log('Paramètre régional du classeur : ' + g.__ss.locale);
const sheet = (n) => g.__ss.getSheetByName(n);
const rows = (n) => { const s = sheet(n); return s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 11).getValues() : []; };
let passed = 0;
const t = (name, fn) => { fn(); passed++; console.log('  ✓ ' + name); };
const ok = (res) => { if (!res.ok) throw new Error(res.error); return res.data; };

const srcActive = fixture['PROJETS'].slice(1).filter((r) => String(r[2]).trim());
const srcOld = fixture['PROJETS.OLD'].slice(1).filter((r) => String(r[1]).trim());

console.log('Installation');
const report = g.installer();
t('copie des collectes en cours', () => assert.strictEqual(rows('PROJETS CLAUDE').length, srcActive.length));
t('copie de l\'historique', () => assert.strictEqual(rows('PROJETS.OLD CLAUDE').length, srcOld.length));
t('montants collectés identiques à PROJETS (formules)', () => {
  rows('PROJETS CLAUDE').forEach((r, i) => assert.strictEqual(Number(r[4]), Number(srcActive[i][4]) || 0, r[0]));
});
t('reste à collecter identique', () => {
  rows('PROJETS CLAUDE').forEach((r, i) => {
    const want = srcActive[i][5] === '-' || srcActive[i][5] === '' ? '-' : Number(srcActive[i][5]);
    assert.strictEqual(r[5], want, r[0]);
  });
});
t('catégories recopiées sur chaque ligne', () => rows('PROJETS CLAUDE').forEach((r) => assert.ok(r[1])));
t('IDs uniques parmi les collectes en cours', () => {
  const ids = rows('PROJETS CLAUDE').map((r) => r[0]);
  assert.strictEqual(new Set(ids).size, ids.length);
});
t('un cas récurrent garde le même ID et un cycle supérieur', () => {
  const act = rows('PROJETS CLAUDE'), old = rows('PROJETS.OLD CLAUDE');
  const norm = (s) => String(s).trim();
  const a = act.find((r) => old.some((o) => norm(o[2]) === norm(r[2])));
  assert.ok(a, 'aucun cas récurrent trouvé');
  const hist = old.filter((o) => o[0] === a[0]);
  assert.ok(hist.length >= 1);
  assert.strictEqual(a[7], Math.max(...hist.map((h) => h[7])) + 1);
});
if (expected) {
  t('IDs identiques à l\'analyse (mapping)', () => {
    assert.ok(/Toutes les lignes ont retrouvé leur ID/.test(report), report);
    rows('PROJETS CLAUDE').forEach((r, i) => assert.strictEqual(r[0], expected.ident['A' + (i + 2)]));
  });
}
console.log(report.split('\n').map((l) => '    ' + l).join('\n'));

console.log('Connexion');
const m = sheet('MEMBRES').getRange(2, 1, 1, 5).getValues()[0];
let admin;
t('code admin généré', () => assert.ok(/^[A-Z0-9]{8}$/.test(m[3])));
t('mauvais code refusé', () => assert.strictEqual(g.api({ action: 'login', email: m[1], code: 'XXXX' }).ok, false));
t('connexion admin', () => { admin = ok(g.api({ action: 'login', email: m[1].toUpperCase(), code: m[3] })); assert.strictEqual(admin.user.role, 'admin'); });
t('jeton falsifié refusé', () => assert.strictEqual(g.api({ action: 'list', token: admin.token.slice(0, -2) + 'xx' }).code, 'AUTH'));
const T = admin.token;

console.log('Application');
let list = ok(g.api({ action: 'list', token: T }));
const first = list.cases[0];
t('liste : montants = feuille', () => {
  rows('PROJETS CLAUDE').forEach((r) => assert.strictEqual(list.cases.find((c) => c.id === r[0]).collected, Number(r[4])));
});
t('ajout d\'un don', () => {
  const res = ok(g.api({ action: 'addDonation', token: T, id: first.id, amount: '1 250', donor: 'Test', mode: 'Virement', date: '2026-11-10' }));
  assert.strictEqual(res.case.collected, first.collected + 1250);
  const r = rows('PROJETS CLAUDE').find((x) => x[0] === first.id);
  assert.strictEqual(r[4], first.collected + 1250, 'formule Excel mise à jour');
});
t('montant invalide refusé', () => assert.strictEqual(g.api({ action: 'addDonation', token: T, id: first.id, amount: '-5' }).ok, false));
t('dons de la collecte', () => {
  const d = ok(g.api({ action: 'donations', token: T, id: first.id }));
  assert.strictEqual(d[0].amount, 1250);
});
t('annulation d\'un don', () => {
  const d = ok(g.api({ action: 'donations', token: T, id: first.id }));
  ok(g.api({ action: 'cancelDonation', token: T, ref: d[0].ref, reason: 'erreur' }));
  list = ok(g.api({ action: 'list', token: T }));
  assert.strictEqual(list.cases.find((c) => c.id === first.id).collected, first.collected);
});

console.log('Coexistence avec la feuille');
t('montant tapé à la main => ajustement + formule remise', () => {
  const shA = sheet('PROJETS CLAUDE');
  shA.getRange(2, 5).setValue(first.collected + 300);
  g.onEdit({ range: shA.getRange(2, 5) });
  assert.ok(String(shA.getRange(2, 5).getFormulas()[0][0]).startsWith('=SUMIFS'));
  rows('PROJETS CLAUDE').forEach((r) => assert.ok(!String(r[4]).startsWith('#') && !String(r[5]).startsWith('#'), 'erreur de formule ' + r[0]));
  assert.strictEqual(shA.getRange(2, 5).getValues()[0][0], first.collected + 300);
  const last = rows('DONS').pop();
  assert.strictEqual(last[6], 'Ajustement feuille');
  assert.strictEqual(last[4], 300);
});
t('nouvelle ligne ajoutée dans la feuille => ID, cycle, formules', () => {
  const shA = sheet('PROJETS CLAUDE');
  const r = shA.getLastRow() + 1;
  shA.getRange(r, 2, 1, 3).setValues([['Cas médicaux', 'Nouveau cas saisi dans la feuille', 4000]]);
  g.onEdit({ range: shA.getRange(r, 3) });
  const v = shA.getRange(r, 1, 1, 10).getValues()[0];
  assert.ok(/^MED-\d{3}$/.test(v[0]), v[0]);
  assert.strictEqual(v[7], 1);
  assert.strictEqual(v[5], 4000);
  list = ok(g.api({ action: 'list', token: T }));
  assert.ok(list.cases.some((c) => c.id === v[0]));
});
t('don saisi à la main dans DONS => complété et compté', () => {
  const shD = sheet('DONS');
  const r = shD.getLastRow() + 1;
  shD.getRange(r, 2, 1, 4).setValues([[new Date(), first.id.toLowerCase(), '', 200]]);
  g.onEdit({ range: shD.getRange(r, 5) });
  const v = shD.getRange(r, 1, 1, 11).getValues()[0];
  assert.ok(v[0].startsWith('DON-'));
  assert.strictEqual(v[2], first.id);
  assert.strictEqual(v[3], first.cycle);
  list = ok(g.api({ action: 'list', token: T }));
  assert.strictEqual(list.cases.find((c) => c.id === first.id).collected, first.collected + 500);
});

t('anciennes formules à virgules (installation v1.0.0) réparées automatiquement', () => {
  if (g.__ss.sep === ',') return;
  const shA = sheet('PROJETS CLAUDE');
  shA.getRange(3, 5, 1, 2).setFormulas([['=SUMIFS(DONS!$E:$E,DONS!$C:$C,$A3,DONS!$D:$D,$H3,DONS!$K:$K,"<>OUI")', '=IF(OR($D3="",$E3>=$D3),"-",$D3-$E3)']]);
  assert.strictEqual(shA.getRange(3, 5).getValues()[0][0], '#ERROR!');
  ok(g.api({ action: 'list', token: T }));
  assert.ok(!String(shA.getRange(3, 5).getValues()[0][0]).startsWith('#'));
  assert.ok(!String(shA.getRange(3, 6).getValues()[0][0]).startsWith('#'));
});

console.log('Cycle de vie');
let reopened;
t('clôture => déplacée vers PROJETS.OLD CLAUDE', () => {
  const before = rows('PROJETS.OLD CLAUDE').length;
  ok(g.api({ action: 'closeCase', token: T, id: first.id, situation: 'Collecte clôturée (test)' }));
  assert.strictEqual(rows('PROJETS.OLD CLAUDE').length, before + 1);
  const top = rows('PROJETS.OLD CLAUDE')[0];
  assert.strictEqual(top[0], first.id);
  assert.strictEqual(top[4], first.collected + 500);
  assert.ok(!rows('PROJETS CLAUDE').some((r) => r[0] === first.id));
});
t('formules des autres lignes toujours justes après suppression', () => {
  const shA = sheet('PROJETS CLAUDE');
  for (let r = 2; r <= shA.getLastRow(); r++) assert.ok(shA.getRange(r, 5).getFormulas()[0][0].includes('$A' + r + g.__ss.sep), 'ligne ' + r);
  list = ok(g.api({ action: 'list', token: T }));
  rows('PROJETS CLAUDE').forEach((r) => assert.strictEqual(list.cases.find((c) => c.id === r[0]).collected, Number(r[4])));
});
t('réouverture d\'un cas récurrent => même ID, nouveau cycle, compteur à zéro', () => {
  reopened = ok(g.api({ action: 'createCase', token: T, reuseId: first.id, category: first.category, description: first.description, budget: 2500 })).case;
  assert.strictEqual(reopened.id, first.id);
  assert.strictEqual(reopened.cycle, first.cycle + 1);
  assert.strictEqual(reopened.collected, 0);
});
t('nouvelle collecte => nouvel ID par catégorie, insérée dans sa catégorie', () => {
  const c = ok(g.api({ action: 'createCase', token: T, category: 'Mosquées', description: 'Nouvelle mosquée (test)', budget: 1000 })).case;
  assert.ok(/^MOS-\d{3}$/.test(c.id));
  const shA = sheet('PROJETS CLAUDE');
  for (let r = 2; r <= shA.getLastRow(); r++) assert.ok(shA.getRange(r, 5).getFormulas()[0][0].includes('$A' + r + g.__ss.sep), 'ligne ' + r);
  list = ok(g.api({ action: 'list', token: T }));
  rows('PROJETS CLAUDE').forEach((r) => assert.strictEqual(list.cases.find((x) => x.id === r[0]).collected, Number(r[4])));
});
t('modification d\'une collecte', () => {
  const c = ok(g.api({ action: 'updateCase', token: T, id: reopened.id, budget: 3000, situation: 'Mise à jour' })).case;
  assert.strictEqual(c.budget, 3000);
  assert.strictEqual(c.remaining, 3000);
});

console.log('Droits');
t('un membre voit les cartes mais ne peut pas ajouter de don', () => {
  sheet('MEMBRES').appendRow(['Membre test', 'membre@example.org', 'membre', '', '']);
  g.onEdit({ range: sheet('MEMBRES').getRange(3, 2) });
  const code = sheet('MEMBRES').getRange(3, 4).getValues()[0][3 - 3];
  const mem = ok(g.api({ action: 'login', email: 'membre@example.org', code }));
  assert.strictEqual(mem.user.role, 'membre');
  assert.ok(ok(g.api({ action: 'list', token: mem.token })).cases.length > 0);
  assert.strictEqual(g.api({ action: 'addDonation', token: mem.token, id: reopened.id, amount: 10 }).code, 'FORBIDDEN');
  assert.strictEqual(g.api({ action: 'donations', token: mem.token, id: reopened.id }).code, 'FORBIDDEN');
  sheet('MEMBRES').getRange(3, 5).setValue('NON');
  assert.strictEqual(g.api({ action: 'list', token: mem.token }).code, 'AUTH');
});
t('installation refusée si déjà faite', () => assert.throws(() => g.installer(), /déjà des données/));
t('onglets d\'origine inchangés', () => {
  assert.deepStrictEqual(sheet('PROJETS').getDataRange().getValues().length, fixture['PROJETS'].length);
  assert.deepStrictEqual(sheet('PROJETS.OLD').getDataRange().getValues().length, fixture['PROJETS.OLD'].length);
});
console.log(`\n${passed} tests réussis`);

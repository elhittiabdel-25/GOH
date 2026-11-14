// Enregistre la vidéo-guide verticale (1080x1920) à partir du studio.
//   node test/video/server.js 8090 &   puis   node test/video/record.js <dossier_sortie> [--preview]
'use strict';
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const OUT = process.argv[2] || path.join(__dirname, 'out');
const PREVIEW = process.argv.includes('--preview');
const BASE = 'http://localhost:8090/';
fs.mkdirSync(OUT, { recursive: true });

(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({
    viewport: { width: 1080, height: 1920 }, locale: 'fr-FR', timezoneId: 'Africa/Casablanca',
    ...(PREVIEW ? {} : { recordVideo: { dir: OUT, size: { width: 1080, height: 1920 } } })
  });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  page.on('dialog', (d) => (d.type() === 'prompt' ? d.accept('Erreur de saisie') : d.accept()));
  await page.goto(BASE);
  await page.waitForTimeout(1500);
  const app = page.frameLocator('#phone');
  const appFrame = () => page.frames().find((f) => f.url().includes('/app/'));
  const S = (fn, ...args) => page.evaluate(([fn, args]) => window.studio[fn](...args), [fn, args]);
  const wait = (ms) => page.waitForTimeout(PREVIEW ? Math.min(ms, 300) : ms);
  let shot = 0;
  const snap = async (name) => { if (PREVIEW) await page.screenshot({ path: path.join(OUT, `${String(++shot).padStart(2, '0')}-${name}.png`) }); };

  async function say(html, extra = 0) {
    await S('say', html);
    const len = html.replace(/<[^>]+>/g, '').length;
    await wait(Math.max(2600, 1200 + len * 52) + extra);
  }
  async function tap(sel, opts = {}) {
    const loc = typeof sel === 'string' ? app.locator(sel).first() : sel;
    await loc.scrollIntoViewIfNeeded();
    const b = await loc.boundingBox();
    if (b) await S('tap', b.x + b.width / 2, b.y + b.height / 2);
    await wait(350);
    await loc.click(opts);
    await wait(500);
  }
  async function typeIn(sel, text) {
    await tap(sel);
    await app.locator(sel).pressSequentially(String(text), { delay: PREVIEW ? 0 : 55 });
    await wait(400);
  }
  async function chapter(n, title, sub) {
    await S('chapter', `${n} · ${title}`);
    await S('say', `<b>${title}</b>${sub ? `<small>${sub}</small>` : ''}`);
    await S('card', `<div class="n">Étape ${n}</div><h1>${title}</h1>${sub ? `<p>${sub}</p>` : ''}`);
    await wait(3400);
    await S('hideCard');
    await wait(700);
  }
  async function appRefresh() { await tap('#refreshBtn'); await wait(1300); }
  async function appScroll(y) { await appFrame().evaluate((y) => window.scrollTo({ top: y, behavior: 'smooth' }), y); await wait(900); }

  /* ---------------- Introduction ---------------- */
  await S('card', `<div class="n">Guide vidéo</div><h1>Suivre les collectes de Gate of Hope</h1>
    <p>Ce que fait l'application, comment l'utiliser, et comment elle reste synchronisée avec le Google Sheet.</p>`);
  await wait(5200);
  await S('hideCard'); await wait(700);

  /* ---------------- 1. Le principe ---------------- */
  await chapter(1, 'Le principe', 'Un seul document de référence : le Google Sheet.');
  await S('focus', 'both');
  await say('En haut, le <b>Google Sheet</b>. En bas, l\'<b>application</b> sur le téléphone.<small>Ce sont deux fenêtres sur les mêmes données.</small>');
  await S('focus', 'sheet');
  await say('L\'application <b>lit et écrit directement</b> dans le Google Sheet.<small>Il n\'y a pas d\'autre base de données : le fichier reste la référence.</small>');
  await say('<b>PROJETS CLAUDE</b> : les collectes en cours.<small>Chaque collecte a un <b>ID</b> permanent (MED-001, FML-002…) et un numéro de <b>cycle</b>.</small>');
  await S('select', 'E', 2);
  await say('« Montant collecté » et « Reste » sont des <b>formules</b> calculées à partir des dons.');
  await S('tab', 'DONS'); await wait(600);
  await say('<b>DONS</b> : une ligne par don reçu.<small>Date, ID de la collecte, montant, donateur, mode de paiement…</small>');
  await S('tab', 'PROJETS.OLD CLAUDE'); await wait(600);
  await say('<b>PROJETS.OLD CLAUDE</b> : les collectes clôturées, avec leur ID.<small>Un cas qui revient garde le même ID.</small>');
  await S('tab', 'MEMBRES'); await wait(600);
  await say('<b>MEMBRES</b> : les personnes autorisées, leur rôle (admin ou membre) et leur code d\'accès.');
  await S('tab', 'JOURNAL APP'); await wait(600);
  await say('<b>JOURNAL APP</b> : la trace de chaque action faite depuis l\'application.');
  await say('Les onglets d\'origine <b>PROJETS</b> et <b>PROJETS.OLD</b> ne sont <b>jamais modifiés</b>.', -400);
  await snap('principe');

  /* ---------------- 2. Se connecter ---------------- */
  await chapter(2, 'Se connecter', 'Email + code d\'accès personnel.');
  await S('tab', 'MEMBRES'); await S('select', 'D', 2);
  await S('focus', 'sheet');
  await say('Le code d\'accès de chaque personne est dans l\'onglet <b>MEMBRES</b>.<small>Ajoutez une ligne (nom, email, rôle) : le code est créé tout seul.</small>');
  await S('focus', 'app');
  await say('Sur le téléphone, ouvrez l\'adresse de l\'application et connectez-vous.', -1200);
  await typeIn('#email', 'admin@exemple.org');
  await typeIn('#code', 'K7P2Q9XA');
  await tap('#loginBtn');
  await app.locator('.card').first().waitFor();
  await wait(800);
  await say('Installez-la sur l\'écran d\'accueil :<small><b>iPhone</b> : Partager → « Sur l\'écran d\'accueil ».<br><b>Android</b> : menu ⋮ → « Installer l\'application ».</small>');
  await snap('connexion');

  /* ---------------- 3. Les cartes ---------------- */
  await chapter(3, 'Les cartes des collectes', 'Ce que voient tous les membres.');
  await S('focus', 'app'); await S('tab', 'PROJETS CLAUDE');
  await say('En haut : le total collecté, le reste à collecter et les dons du mois.');
  await say('Chaque collecte est une <b>carte</b> : barre de progression, montant collecté, budget et reste.');
  await tap('[data-cat="Cas médicaux"]');
  await say('Filtrez par type de collecte, ou cherchez un texte ou un ID.', -800);
  await tap('[data-open="MED-001"]');
  await wait(800);
  await say('Touchez une carte pour voir le détail.<small>Ici MED-001 en est à sa <b>3e collecte</b> : les précédentes apparaissent dans l\'historique.</small>');
  await appFrame().evaluate(() => { const s = document.querySelector('.sheet-inner'); s.scrollTo({ top: s.scrollHeight, behavior: 'smooth' }); });
  await wait(2600);
  await tap('[data-close]');
  await tap('[data-cat=""]');
  await say('Les <b>membres</b> voient uniquement les cartes. Les <b>admins</b> peuvent en plus saisir des dons et gérer les collectes.');
  await snap('cartes');

  /* ---------------- 4. Don depuis l'application ---------------- */
  await chapter(4, 'Ajouter un don depuis l\'application', 'Réservé aux administrateurs.');
  await S('focus', 'app');
  await say('Un don arrive ? Touchez <span class="k">+ Don reçu</span>.', -800);
  await tap('#fab');
  await app.locator('#dCase').selectOption('FML-001'); await wait(900);
  await say('Choisissez la collecte, puis le montant.', -1000);
  await tap('[data-q="500"]');
  await app.locator('#dMode').selectOption('Virement'); await wait(600);
  await typeIn('#dDonor', 'Famille B.');
  await say('Donateur et note sont facultatifs (« Anonyme » par défaut).', -900);
  await tap('#dSave');
  await wait(1500);
  await snap('don-app');
  await S('focus', 'sheet');
  await S('tab', 'DONS'); await wait(900);
  await say('Dans le Google Sheet, une <b>nouvelle ligne</b> apparaît aussitôt dans <b>DONS</b>.');
  await S('tab', 'PROJETS CLAUDE'); await S('select', 'E', 2); await wait(900);
  await say('Et le montant collecté de <b>FML-001</b> est mis à jour : 4 200 → <b>4 700 Dhs</b>.');
  await snap('don-feuille');

  /* ---------------- 5. Don saisi dans la feuille ---------------- */
  await chapter(5, 'Saisir un don directement dans le Google Sheet', 'Pour ceux qui préfèrent le fichier.');
  await S('tab', 'DONS'); await S('focus', 'sheet');
  await say('Ajoutez simplement une ligne dans <b>DONS</b> : la date, l\'ID et le montant.', -900);
  const dRow = await nextEmptyRow('DONS');
  await S('type', 'B', dRow, '14/11/2026');
  await S('type', 'C', dRow, 'MED-002');
  await S('type', 'E', dRow, '1000');
  await S('type', 'F', dRow, 'Mme K.');
  await wait(800);
  await say('Le reste se complète tout seul : <b>cycle</b>, référence, « saisi par »…');
  await S('tab', 'PROJETS CLAUDE'); await wait(900);
  await say('Le montant de <b>MED-002</b> est recalculé automatiquement.');
  await S('focus', 'app');
  await say('L\'application se met à jour toute seule, ou tout de suite avec <span class="k">↻</span>.', -1000);
  await appRefresh();
  await appScroll(560);
  await say('La carte affiche <b>+1 000 Dhs</b> : les membres voient ce qui a changé depuis leur dernière visite.');
  await snap('don-feuille-app');

  /* ---------------- 6. Nouvelle collecte dans la feuille ---------------- */
  await chapter(6, 'Créer une collecte dans le Google Sheet', 'Une ligne de plus suffit.');
  await S('focus', 'sheet'); await S('tab', 'PROJETS CLAUDE');
  const cRow = await nextEmptyRow('PROJETS CLAUDE');
  await say('Ajoutez une ligne : description et budget.<small>Si le type est vide, il reprend celui de la ligne du dessus.</small>', -1200);
  await S('type', 'C', cRow, 'Achat de tapis pour une mosquée (exemple fictif)', 40);
  await S('type', 'D', cRow, '8000');
  await wait(900);
  await S('select', 'A', cRow); await wait(600);
  await say('L\'<b>ID</b> (MOS-002), le <b>cycle</b>, la date et les formules sont ajoutés automatiquement.');
  await S('focus', 'app');
  await appRefresh();
  await appScroll(5000);
  await say('Elle apparaît dans l\'application avec l\'étiquette <b>Nouvelle</b>.');
  await snap('collecte-feuille');

  /* ---------------- 7. Montant tapé à la main ---------------- */
  await chapter(7, 'Modifier un montant à la main', 'Rien ne se perd.');
  await S('focus', 'sheet'); await S('tab', 'PROJETS CLAUDE');
  await say('Habitude de taper le montant collecté directement ? C\'est possible.', -1000);
  await S('type', 'E', 3, '8000');
  await wait(900);
  await say('La différence (+500) devient une ligne <b>« Ajustement feuille »</b> dans DONS, et la formule est remise en place.');
  await S('tab', 'DONS'); await wait(1600);
  await say('Ainsi les totaux restent justes et l\'historique est complet.', -800);
  await snap('ajustement');

  /* ---------------- 8. Nouvelle collecte depuis l'application ---------------- */
  await chapter(8, 'Créer une collecte depuis l\'application', 'Réservé aux administrateurs.');
  await S('focus', 'app'); await S('tab', 'PROJETS CLAUDE');
  await appScroll(0);
  await tap('#newCase');
  await app.locator('#cCat').selectOption('Cas médicaux'); await wait(700);
  await typeIn('#cDesc', 'Bilan sanguin pour une jeune maman (exemple fictif)');
  await typeIn('#cBud', '3500');
  await say('Type, description, budget (vide s\'il n\'y a pas d\'objectif fixe) et situation.', -900);
  await tap('#cSave');
  await wait(1500);
  await S('focus', 'sheet'); await wait(600);
  await say('La ligne est insérée dans le Google Sheet, <b>avec les autres cas médicaux</b>, et reçoit l\'ID <b>MED-004</b>.');
  await snap('collecte-app');

  /* ---------------- 9. Clôturer ---------------- */
  await chapter(9, 'Clôturer une collecte', 'Elle part dans l\'historique.');
  await S('focus', 'app');
  await tap('[data-close]');
  await tap('[data-open="MED-002"]');
  await wait(800);
  await tap('#aClose');
  await say('Depuis la carte : <span class="k">Clôturer</span>, puis indiquez la situation finale.', -900);
  await tap('#xSave');
  await wait(1500);
  await S('focus', 'sheet'); await S('tab', 'PROJETS CLAUDE'); await wait(900);
  await say('La ligne quitte <b>PROJETS CLAUDE</b>…', -1200);
  await S('tab', 'PROJETS.OLD CLAUDE'); await wait(900);
  await say('… et arrive en haut de <b>PROJETS.OLD CLAUDE</b>, avec le montant final et la date de clôture.');
  await say('Dans le Google Sheet : sélectionnez la ligne puis menu <b>Gate of Hope → Clôturer la collecte sélectionnée</b>.');
  await snap('cloture');

  /* ---------------- 10. Cas récurrent ---------------- */
  await chapter(10, 'Un cas qui revient', 'Même ID, nouveau cycle.');
  await S('focus', 'app');
  await tap('[data-tab="closed"]');
  await typeIn('#q', 'rééducation');
  await tap('[data-closed="0"]');
  await say('Dans « Clôturées », ouvrez le cas puis <span class="k">Relancer une collecte</span>.', -900);
  await tap('#reopen');
  await app.locator('#cBud').fill(''); await typeIn('#cBud', '3000');
  await tap('#cSave');
  await wait(1600);
  await say('Il garde son ID <b>MED-003</b> : c\'est sa <b>collecte n°2</b>, qui repart de zéro.');
  await S('focus', 'sheet'); await S('tab', 'PROJETS CLAUDE'); await wait(900);
  await say('Dans le Google Sheet, même principe : tapez l\'<b>ID existant</b> en colonne A d\'une nouvelle ligne.');
  await snap('recurrent');

  /* ---------------- 11. Corriger une erreur ---------------- */
  await chapter(11, 'Corriger une erreur de saisie', 'On annule, on ne supprime pas.');
  await S('focus', 'app');
  await tap('[data-close]');
  await tap('[data-tab="dons"]');
  await wait(900);
  await say('Onglet <b>Dons</b> : tous les dons saisis, avec qui les a saisis.', -1000);
  await tap(app.locator('.row', { hasText: 'Famille B.' }).locator('[data-cancel]'));
  await wait(1500);
  await S('focus', 'sheet'); await S('tab', 'DONS'); await wait(1200);
  await say('Le don reste dans DONS, marqué <b>Annulé = OUI</b> : il n\'est plus compté.<small>Dans la feuille, tapez simplement OUI dans la colonne « Annulé ».</small>');
  await snap('annulation');

  /* ---------------- Récapitulatif ---------------- */
  await S('chapter', 'Récapitulatif');
  await S('card', `<div class="n">À retenir</div><h1>Les deux mondes coexistent</h1><ul>
    <li><b>Application</b> : dons, nouvelles collectes, clôtures, en quelques secondes.</li>
    <li><b>Google Sheet</b> : on peut toujours ajouter des lignes ou modifier à la main.</li>
    <li>Tout ce qui est fait d'un côté apparaît de l'autre.</li></ul>`);
  await wait(9000);
  await S('card', `<div class="n">Règles d'or</div><h1>Pour garder des données justes</h1><ul>
    <li>Ne pas changer l'ID d'une collecte existante.</li>
    <li>Ne pas supprimer de ligne dans DONS : mettre « OUI » dans Annulé.</li>
    <li>Clôturer avec le menu Gate of Hope ou l'application.</li>
    <li>PROJETS et PROJETS.OLD restent intacts.</li></ul>`);
  await wait(10000);
  await S('card', `<div class="n">Gate of Hope</div><h1>Merci !</h1><p>Une question ? Contactez un administrateur.</p>`);
  await wait(3500);

  async function nextEmptyRow(name) {
    return page.evaluate(async (name) => {
      const d = await (await fetch('/dev/sheet?name=' + encodeURIComponent(name))).json();
      return d.values.length + 1;
    }, name);
  }

  const video = page.video();
  await ctx.close();
  if (video) console.log('VIDEO', await video.path());
  console.log('errors:', errors);
  await browser.close();
})();

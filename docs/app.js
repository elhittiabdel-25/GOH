/* Gate of Hope — application de suivi des collectes (PWA, sans dépendances). */
(function () {
  'use strict';

  var CFG = window.GOH_CONFIG || {};
  var REFRESH_MS = (CFG.REFRESH_SECONDS || 60) * 1000;

  /* ---------------- Stockage local (préférences + dernier état connu) ---------------- */
  var LS = {
    get: function (k, d) { try { var v = localStorage.getItem('goh:' + k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set: function (k, v) { try { localStorage.setItem('goh:' + k, JSON.stringify(v)); } catch (e) { /* stockage indisponible */ } },
    del: function (k) { try { localStorage.removeItem('goh:' + k); } catch (e) { /* idem */ } }
  };

  var state = {
    session: LS.get('session', null),
    data: LS.get('data', null),
    tab: 'open', cat: '', q: '', closedLimit: 40,
    loading: false, offline: false, lastSync: LS.get('lastSync', null),
    seen: LS.get('seen', null),
    recent: null, installPrompt: null
  };

  /* ---------------- Utilitaires ---------------- */
  var $app = document.getElementById('app');
  var $sheet = document.getElementById('sheet');
  var $sheetInner = $sheet.querySelector('.sheet-inner');
  var nf = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 2 });

  function esc(s) {
    return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function dh(n) { return nf.format(Math.round((n || 0) * 100) / 100) + ' Dhs'; }
  function isAdmin() { return state.session && state.session.user && state.session.user.role === 'admin'; }
  function key(c) { return c.id + '#' + c.cycle; }
  function pct(c) { return c.budget ? Math.min(100, Math.round(c.collected / c.budget * 100)) : null; }
  function todayStr() {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function fmtDate(iso) {
    if (!iso) return '';
    return new Date(iso).toLocaleDateString('fr-FR', { day: 'numeric', month: 'short', year: 'numeric' });
  }
  function ago(iso) {
    if (!iso) return '';
    var s = (Date.now() - new Date(iso).getTime()) / 1000;
    if (s < 60) return 'à l\'instant';
    if (s < 3600) return 'il y a ' + Math.floor(s / 60) + ' min';
    if (s < 86400) return 'il y a ' + Math.floor(s / 3600) + ' h';
    if (s < 172800) return 'hier';
    if (s < 30 * 86400) return 'il y a ' + Math.floor(s / 86400) + ' jours';
    return 'le ' + fmtDate(iso);
  }
  function norm(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase(); }
  function short(s, n) { s = String(s || ''); return s.length > n ? s.slice(0, n - 1) + '…' : s; }

  var toastTimer;
  function toast(msg) {
    var t = document.getElementById('toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove('show'); }, 3200);
  }

  var ICON = {
    refresh: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-2.64-6.36"/><path d="M21 3v6h-6"/></svg>',
    menu: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="8" r="4"/><path d="M4 21c1.5-4 4.5-6 8-6s6.5 2 8 6"/></svg>',
    close: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>',
    plus: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>',
    search: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>'
  };

  /* ---------------- API ---------------- */
  function api(action, payload) {
    if (!CFG.API_URL) return Promise.reject(new Error('Application non configurée : renseignez API_URL dans config.js.'));
    var body = Object.assign({ action: action, token: state.session && state.session.token }, payload || {});
    return fetch(CFG.API_URL, {
      method: 'POST', redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    }).catch(function () {
      var e = new Error('Pas de connexion internet.'); e.code = 'OFFLINE'; throw e;
    }).then(function (res) {
      return res.json().catch(function () { throw new Error('Réponse inattendue du serveur (' + res.status + ').'); });
    }).then(function (j) {
      if (!j.ok) {
        var e = new Error(j.error || 'Erreur inconnue'); e.code = j.code;
        if (j.code === 'AUTH' && action !== 'login') { logout(); toast(e.message); }
        throw e;
      }
      return j.data;
    });
  }

  function refresh(quiet) {
    if (!state.session || state.loading) return Promise.resolve();
    state.loading = true;
    renderRefreshBtn();
    return api('list').then(function (d) {
      state.data = d;
      state.offline = false;
      state.lastSync = new Date().toISOString();
      if (d.user) state.session.user = Object.assign(state.session.user, d.user);
      if (!state.seen) saveSeen();
      LS.set('data', d); LS.set('lastSync', state.lastSync); LS.set('session', state.session);
      if (state.tab === 'dons' && isAdmin()) loadRecent();
    }).catch(function (e) {
      if (e.code === 'OFFLINE') state.offline = true;
      else if (!quiet) toast(e.message);
    }).then(function () {
      state.loading = false;
      if (state.session) render();
    });
  }

  function loadRecent() {
    return api('recentDonations', { limit: 80 }).then(function (r) { state.recent = r; if (state.tab === 'dons') render(); })
      .catch(function (e) { toast(e.message); });
  }

  /** Mémorise les montants vus, pour signaler ce qui a changé à la prochaine ouverture. */
  function saveSeen() {
    if (!state.data) return;
    var s = {};
    state.data.cases.forEach(function (c) { s[key(c)] = c.collected; });
    state.seen = s;
    LS.set('seen', s);
  }

  function logout() {
    state.session = null; state.data = null; state.seen = null; state.recent = null;
    ['session', 'data', 'seen', 'lastSync'].forEach(LS.del);
    closeSheet();
    render();
  }

  /* ---------------- Rendu ---------------- */
  function render() {
    if (!state.session) return renderLogin();
    renderMain();
  }

  function renderLogin(err) {
    $app.innerHTML =
      '<div class="login"><form class="login-box" id="loginForm" autocomplete="on">' +
      '<img src="icons/icon-192.png" alt="">' +
      '<h1>Gate of Hope</h1><p>Suivi des collectes</p>' +
      (CFG.API_URL ? '' : '<div class="banner">Application pas encore reliée au Google Sheet (API_URL vide dans config.js).</div>') +
      '<label for="email">Email</label><input class="field" id="email" type="email" autocomplete="username" required inputmode="email" value="' + esc(LS.get('lastEmail', '')) + '">' +
      '<label for="code">Code d\'accès</label><input class="field" id="code" type="password" autocomplete="current-password" required autocapitalize="characters">' +
      '<div class="err" id="loginErr">' + esc(err || '') + '</div>' +
      '<div class="form-foot"><button class="btn primary block" id="loginBtn">Se connecter</button></div>' +
      '<p class="small">Le code vous est communiqué par un administrateur.</p>' +
      '</form></div>';
    document.getElementById('loginForm').onsubmit = function (ev) {
      ev.preventDefault();
      var btn = document.getElementById('loginBtn');
      var email = document.getElementById('email').value.trim();
      btn.disabled = true; btn.textContent = 'Connexion…';
      api('login', { email: email, code: document.getElementById('code').value.trim().toUpperCase() })
        .then(function (s) {
          LS.set('lastEmail', email);
          state.session = s; LS.set('session', s);
          render(); refresh();
        })
        .catch(function (e) {
          btn.disabled = false; btn.textContent = 'Se connecter';
          document.getElementById('loginErr').textContent = e.message;
        });
    };
  }

  function renderRefreshBtn() {
    var b = document.getElementById('refreshBtn');
    if (b) b.classList.toggle('spin', state.loading);
  }

  function renderMain() {
    var d = state.data;
    var user = state.session.user || {};
    var scrollY = window.scrollY;
    var html = '<header class="top"><div class="wrap">' +
      '<img src="icons/icon-192.png" alt="">' +
      '<h1>Gate of Hope<small>' + esc(user.name || '') + (isAdmin() ? ' · admin' : '') + '</small></h1>' +
      '<button class="icon-btn' + (state.loading ? ' spin' : '') + '" id="refreshBtn" aria-label="Actualiser">' + ICON.refresh + '</button>' +
      '<button class="icon-btn" id="menuBtn" aria-label="Mon compte">' + ICON.menu + '</button>' +
      '</div></header><main class="wrap">';

    if (state.offline) html += '<div class="banner">Hors connexion : affichage des dernières données reçues' + (state.lastSync ? ' (' + ago(state.lastSync) + ')' : '') + '.</div>';
    html += installBanner();

    if (!d) {
      html += '<div class="list" style="margin-top:16px"><div class="skeleton"></div><div class="skeleton"></div><div class="skeleton"></div></div></main>';
      $app.innerHTML = html; bindHeader(); return;
    }

    var s = d.stats;
    html += '<section class="stats" aria-label="Résumé">' +
      '<div class="stat"><b>' + nf.format(Math.round(s.collected)) + '</b><span>Dhs collectés</span></div>' +
      '<div class="stat"><b>' + nf.format(Math.round(s.remaining)) + '</b><span>Dhs à collecter</span></div>' +
      '<div class="stat"><b>' + nf.format(Math.round(s.monthCollected)) + '</b><span>Dhs reçus ce mois</span></div>' +
      '</section>';

    html += '<nav class="tabs" role="tablist">' +
      tabBtn('open', 'En cours (' + d.cases.length + ')') +
      tabBtn('closed', 'Clôturées (' + d.closed.length + ')') +
      (isAdmin() ? tabBtn('dons', 'Dons') : '') + '</nav>';

    if (state.tab === 'dons') {
      html += renderRecent();
    } else {
      var src = state.tab === 'open' ? d.cases : d.closed;
      var cats = [];
      src.forEach(function (c) { if (c.category && cats.indexOf(c.category) < 0) cats.push(c.category); });
      html += '<div class="search">' + ICON.search + '<input id="q" type="search" placeholder="Rechercher (texte ou ID)…" value="' + esc(state.q) + '" aria-label="Rechercher"></div>';
      html += '<div class="chips">' + chip('', 'Toutes') + cats.map(function (c) { return chip(c, c); }).join('') + '</div>';
      html += state.tab === 'open' ? renderOpen(filter(src)) : renderClosed(filter(src));
    }
    html += '<p class="foot-note">' + (state.lastSync ? 'Données du Google Sheet, actualisées ' + ago(state.lastSync) : '') + '</p></main>';
    if (isAdmin() && state.tab !== 'closed') html += '<button class="fab" id="fab">' + ICON.plus + 'Don reçu</button>';

    $app.innerHTML = html;
    window.scrollTo(0, scrollY);
    bindHeader();
    bindMain();
  }

  function tabBtn(id, label) {
    return '<button role="tab" data-tab="' + id + '" aria-selected="' + (state.tab === id) + '">' + esc(label) + '</button>';
  }
  function chip(val, label) {
    return '<button class="chip" data-cat="' + esc(val) + '" aria-pressed="' + (state.cat === val) + '">' + esc(label) + '</button>';
  }

  function filter(list) {
    var q = norm(state.q).trim();
    return list.filter(function (c) {
      if (state.cat && c.category !== state.cat) return false;
      if (!q) return true;
      return norm(c.id + ' ' + c.description + ' ' + c.situation + ' ' + c.category).indexOf(q) >= 0;
    });
  }

  function renderOpen(list) {
    if (!list.length) return '<div class="empty">Aucune collecte en cours' + (state.q || state.cat ? ' pour ce filtre' : '') + '.</div>';
    var html = '<div class="list">', lastCat = null;
    list.forEach(function (c) {
      if (!state.cat && c.category !== lastCat) { html += '<div class="group-title">' + esc(c.category) + '</div>'; lastCat = c.category; }
      html += caseCard(c);
    });
    return html + '</div>';
  }

  function renderClosed(list) {
    if (!list.length) return '<div class="empty">Aucune collecte clôturée pour ce filtre.</div>';
    var shown = list.slice(0, state.closedLimit);
    var html = '<div class="list">' + shown.map(closedCard).join('');
    if (list.length > shown.length) html += '<button class="btn block" id="more">Afficher plus (' + (list.length - shown.length) + ')</button>';
    return html + '</div>';
  }

  function progress(c) {
    var p = pct(c);
    if (p === null) return '<div class="bar open" aria-hidden="true"><i style="width:100%;opacity:.35"></i></div>';
    return '<div class="bar' + (p >= 100 ? ' full' : '') + '" role="progressbar" aria-valuenow="' + p + '" aria-valuemin="0" aria-valuemax="100"><i style="width:' + p + '%"></i></div>';
  }

  function caseCard(c) {
    var prev = state.seen ? state.seen[key(c)] : undefined;
    var delta = prev === undefined ? (state.seen ? c.collected : 0) : c.collected - prev;
    var isNew = state.seen && prev === undefined;
    var p = pct(c);
    return '<button class="card' + (delta > 0 || isNew ? ' changed' : '') + '" data-open="' + esc(c.id) + '">' +
      '<div class="card-head"><span class="tag id">' + esc(c.id) + '</span>' +
      (isNew ? '<span class="tag new">Nouvelle</span>' : delta > 0 ? '<span class="tag new">+' + dh(delta) + '</span>' : '') +
      (p !== null && p >= 100 ? '<span class="tag done">Objectif atteint</span>' : '') +
      '<span class="muted small" style="margin-left:auto">' + esc(ago(c.updatedAt)) + '</span></div>' +
      '<p class="desc">' + esc(c.description) + '</p>' + progress(c) +
      '<div class="nums"><span><b>' + dh(c.collected) + '</b>' + (c.budget ? ' <span class="muted">/ ' + dh(c.budget) + '</span>' : '') + '</span>' +
      '<span class="muted">' + (c.budget ? (c.remaining > 0 ? 'Reste ' + dh(c.remaining) : c.collected > c.budget ? 'Dépassé de ' + dh(c.collected - c.budget) : 'Objectif atteint') : 'Sans objectif fixe') + '</span></div>' +
      (c.situation ? '<p class="situation">' + esc(short(c.situation, 160)) + '</p>' : '') +
      '</button>';
  }

  function closedCard(c, i) {
    return '<button class="card" data-closed="' + i + '">' +
      '<div class="card-head"><span class="tag id">' + esc(c.id) + '</span><span class="tag">' + esc(c.category) + '</span>' +
      '<span class="tag done">Clôturée</span>' +
      (c.closedAt ? '<span class="muted small" style="margin-left:auto">' + esc(fmtDate(c.closedAt)) + '</span>' : '') + '</div>' +
      '<p class="desc">' + esc(c.description) + '</p>' +
      '<div class="nums"><span><b>' + dh(c.collected) + '</b>' + (c.budget ? ' <span class="muted">/ ' + dh(c.budget) + '</span>' : '') + '</span></div>' +
      (c.situation ? '<p class="situation">' + esc(short(c.situation, 140)) + '</p>' : '') + '</button>';
  }

  function renderRecent() {
    if (!state.recent) { loadRecent(); return '<div class="list"><div class="skeleton"></div></div>'; }
    if (!state.recent.length) return '<div class="empty">Aucun don enregistré depuis l\'application pour l\'instant.</div>';
    var byId = {};
    state.data.cases.concat(state.data.closed).forEach(function (c) { byId[c.id] = byId[c.id] || c; });
    return '<div class="rows" style="margin-bottom:110px">' + state.recent.map(function (d) {
      return donationRow(d, byId[d.id], true);
    }).join('') + '</div>';
  }

  function donationRow(d, c, showCase) {
    var opening = d.mode === 'Reprise Excel';
    return '<div class="row' + (d.cancelled ? ' cancelled' : '') + '"><div class="grow">' +
      '<div>' + (showCase ? '<span class="tag id">' + esc(d.id) + '</span> ' : '') + esc(d.donor || 'Anonyme') + '</div>' +
      (showCase && c ? '<div class="small muted">' + esc(short(c.description, 70)) + '</div>' : '') +
      '<div class="small muted">' + esc(fmtDate(d.date)) + ' · ' + esc(d.mode) + (d.by && !opening ? ' · ' + esc(d.by.split('@')[0]) : '') +
      (d.note ? ' · ' + esc(short(d.note, 80)) : '') + '</div></div>' +
      '<div style="text-align:right"><div class="amt">' + (d.amount > 0 ? '+' : '') + dh(d.amount) + '</div>' +
      (!d.cancelled && !opening ? '<button class="link" data-cancel="' + esc(d.ref) + '">Annuler</button>' : d.cancelled ? '<span class="small muted">annulé</span>' : '') +
      '</div></div>';
  }

  function installBanner() {
    var standalone = window.matchMedia('(display-mode: standalone)').matches || navigator.standalone;
    if (standalone || LS.get('hideInstall', false)) return '';
    var ios = /iphone|ipad|ipod/i.test(navigator.userAgent);
    if (state.installPrompt) return '<div class="banner">Installez l\'application sur votre téléphone. <button class="link" style="color:inherit;font-weight:700" id="installBtn">Installer</button> <button class="link" id="hideInstall">Plus tard</button></div>';
    if (ios) return '<div class="banner">Pour installer l\'app : touchez <b>Partager</b> puis <b>« Sur l\'écran d\'accueil »</b>. <button class="link" id="hideInstall">OK</button></div>';
    return '';
  }

  function bindHeader() {
    document.getElementById('refreshBtn').onclick = function () { refresh(); };
    document.getElementById('menuBtn').onclick = openMenu;
    var ib = document.getElementById('installBtn');
    if (ib) ib.onclick = function () {
      state.installPrompt.prompt();
      state.installPrompt.userChoice.then(function () { state.installPrompt = null; render(); });
    };
    var hb = document.getElementById('hideInstall');
    if (hb) hb.onclick = function () { LS.set('hideInstall', true); render(); };
  }

  function bindMain() {
    $app.querySelectorAll('[data-tab]').forEach(function (b) {
      b.onclick = function () { state.tab = b.dataset.tab; state.cat = ''; state.closedLimit = 40; render(); window.scrollTo(0, 0); };
    });
    $app.querySelectorAll('[data-cat]').forEach(function (b) {
      b.onclick = function () { state.cat = b.dataset.cat; render(); };
    });
    var q = document.getElementById('q');
    if (q) {
      q.oninput = function () {
        state.q = q.value;
        var pos = q.selectionStart;
        render();
        var nq = document.getElementById('q'); nq.focus(); try { nq.setSelectionRange(pos, pos); } catch (e) { /* */ }
      };
    }
    $app.querySelectorAll('[data-open]').forEach(function (b) {
      b.onclick = function () { openCase(b.dataset.open); };
    });
    var closedList = filter(state.data.closed);
    $app.querySelectorAll('[data-closed]').forEach(function (b) {
      b.onclick = function () { openClosed(closedList[Number(b.dataset.closed)]); };
    });
    var more = document.getElementById('more');
    if (more) more.onclick = function () { state.closedLimit += 60; render(); };
    var fab = document.getElementById('fab');
    if (fab) fab.onclick = function () { donationForm(null); };
    bindCancel($app);
  }

  function bindCancel(root) {
    root.querySelectorAll('[data-cancel]').forEach(function (b) {
      b.onclick = function (ev) {
        ev.stopPropagation();
        var reason = window.prompt('Annuler ce don ? Indiquez la raison (le don reste visible dans le Google Sheet, marqué « Annulé ») :', 'Erreur de saisie');
        if (reason === null) return;
        b.disabled = true;
        api('cancelDonation', { ref: b.dataset.cancel, reason: reason }).then(function () {
          toast('Don annulé');
          state.recent = null;
          var open = $sheet.dataset.case;
          return refresh(true).then(function () { if (open && $sheet.open) openCase(open); });
        }).catch(function (e) { b.disabled = false; toast(e.message); });
      };
    });
  }

  /* ---------------- Feuille / modale ---------------- */
  function openSheet(html, caseId) {
    $sheetInner.innerHTML = '<div class="grab"></div>' + html;
    $sheet.dataset.case = caseId || '';
    if (!$sheet.open) {
      if ($sheet.showModal) $sheet.showModal(); else $sheet.setAttribute('open', '');
    }
    $sheetInner.scrollTop = 0;
    $sheetInner.querySelectorAll('[data-close]').forEach(function (b) { b.onclick = closeSheet; });
  }
  function closeSheet() {
    if ($sheet.open) { if ($sheet.close) $sheet.close(); else $sheet.removeAttribute('open'); }
    $sheet.dataset.case = '';
  }
  $sheet.addEventListener('click', function (e) { if (e.target === $sheet) closeSheet(); });

  function head(title, sub) {
    return '<div class="sheet-head"><div style="flex:1;min-width:0">' + (sub || '') + '<h2>' + esc(title) + '</h2></div>' +
      '<button class="icon-btn" data-close aria-label="Fermer">' + ICON.close + '</button></div>';
  }

  function findCase(id) {
    return state.data && state.data.cases.filter(function (c) { return c.id === id; })[0];
  }
  function historyOf(id) {
    return state.data.closed.filter(function (c) { return c.id === id; }).sort(function (a, b) { return b.cycle - a.cycle; });
  }

  function historyHtml(id, exceptCycle) {
    var h = historyOf(id).filter(function (c) { return c.cycle !== exceptCycle; });
    if (!h.length) return '';
    return '<h3>Collectes précédentes pour ' + esc(id) + ' (' + h.length + ')</h3><div class="rows">' + h.map(function (c) {
      return '<div class="row"><div class="grow"><div class="small"><b>Collecte n°' + c.cycle + '</b>' + (c.closedAt ? ' · clôturée le ' + esc(fmtDate(c.closedAt)) : '') + '</div>' +
        '<div class="small muted">' + esc(short(c.situation || c.description, 110)) + '</div></div>' +
        '<div class="amt">' + dh(c.collected) + '</div></div>';
    }).join('') + '</div>';
  }

  function openCase(id) {
    var c = findCase(id);
    if (!c) { closeSheet(); toast('Cette collecte n\'est plus en cours.'); return; }
    // la carte a été consultée : on mémorise son montant
    if (state.seen) { state.seen[key(c)] = c.collected; LS.set('seen', state.seen); }
    var p = pct(c);
    var html = head(c.category, '<div class="card-head"><span class="tag id">' + esc(c.id) + '</span><span class="tag">Collecte n°' + c.cycle + '</span>' +
      (p !== null && p >= 100 ? '<span class="tag done">Objectif atteint</span>' : '') + '</div>') +
      '<p class="full-desc">' + esc(c.description) + '</p>' + progress(c) +
      '<div class="kpis"><div><b>' + dh(c.collected) + '</b><span>Collecté</span></div>' +
      '<div><b>' + (c.budget ? dh(c.budget) : '—') + '</b><span>Budget</span></div>' +
      '<div><b>' + (c.budget ? dh(c.remaining) : '—') + '</b><span>Reste</span></div></div>' +
      (c.situation ? '<p class="situation">' + esc(c.situation) + '</p>' : '') +
      '<p class="small muted">' + (c.openedAt ? 'Ouverte le ' + esc(fmtDate(c.openedAt)) + ' · ' : '') + 'Mise à jour ' + esc(ago(c.updatedAt)) + '</p>';
    if (isAdmin()) {
      html += '<div class="actions"><button class="btn primary" id="aDon">' + ICON.plus + ' Ajouter un don</button>' +
        '<button class="btn" id="aEdit">Modifier</button><button class="btn danger" id="aClose">Clôturer</button></div>' +
        '<h3>Dons de cette collecte</h3><div id="dons"><div class="skeleton" style="height:60px"></div></div>';
    }
    html += historyHtml(c.id, c.cycle);
    openSheet(html, c.id);
    render();
    if (isAdmin()) {
      document.getElementById('aDon').onclick = function () { donationForm(c.id); };
      document.getElementById('aEdit').onclick = function () { caseForm(c); };
      document.getElementById('aClose').onclick = function () { closeForm(c); };
      api('donations', { id: c.id }).then(function (list) {
        var box = document.getElementById('dons');
        if (!box || $sheet.dataset.case !== c.id) return;
        var cur = list.filter(function (d) { return d.cycle === c.cycle; });
        box.innerHTML = cur.length ? '<div class="rows">' + cur.map(function (d) { return donationRow(d, c, false); }).join('') + '</div>'
          : '<p class="muted small">Aucun don pour l\'instant.</p>';
        bindCancel(box);
      }).catch(function (e) { var box = document.getElementById('dons'); if (box) box.innerHTML = '<p class="err">' + esc(e.message) + '</p>'; });
    }
  }

  function openClosed(c) {
    if (!c) return;
    var html = head(c.category, '<div class="card-head"><span class="tag id">' + esc(c.id) + '</span><span class="tag">Collecte n°' + c.cycle + '</span><span class="tag done">Clôturée</span></div>') +
      '<p class="full-desc">' + esc(c.description) + '</p>' +
      '<div class="kpis"><div><b>' + dh(c.collected) + '</b><span>Collecté</span></div>' +
      '<div><b>' + (c.budget ? dh(c.budget) : '—') + '</b><span>Budget</span></div>' +
      '<div><b>' + (c.closedAt ? esc(fmtDate(c.closedAt)) : '—') + '</b><span>Clôturée le</span></div></div>' +
      (c.situation ? '<p class="situation">' + esc(c.situation) + '</p>' : '');
    var active = findCase(c.id);
    if (active) html += '<div class="banner">' + esc(c.id) + ' a une collecte en cours (n°' + active.cycle + '). <button class="link" style="color:inherit;font-weight:700" id="goActive">Voir</button></div>';
    else if (isAdmin()) html += '<div class="actions"><button class="btn primary" id="reopen">Relancer une collecte pour ' + esc(c.id) + '</button></div>';
    html += historyHtml(c.id, c.cycle);
    openSheet(html);
    var ga = document.getElementById('goActive');
    if (ga) ga.onclick = function () { openCase(c.id); };
    var ro = document.getElementById('reopen');
    if (ro) ro.onclick = function () { caseForm(null, c); };
  }

  function openMenu() {
    var u = state.session.user || {};
    var html = head('Mon compte') +
      '<p>' + esc(u.name) + '<br><span class="muted small">' + esc(u.email) + ' · ' + (u.role === 'admin' ? 'administrateur' : 'membre') + '</span></p>' +
      (state.installPrompt ? '<button class="btn block" id="mInstall" style="margin-bottom:8px">Installer l\'application</button>' : '') +
      '<button class="btn block" id="mRefresh" style="margin-bottom:8px">Actualiser les données</button>' +
      '<button class="btn danger block" id="mLogout">Se déconnecter</button>' +
      '<p class="hint">Les données viennent directement du Google Sheet de Gate of Hope. Toute modification faite dans la feuille apparaît ici à la prochaine actualisation (automatique toutes les ' + Math.round(REFRESH_MS / 1000) + ' s).</p>';
    openSheet(html);
    document.getElementById('mLogout').onclick = function () { if (window.confirm('Se déconnecter ?')) logout(); };
    document.getElementById('mRefresh').onclick = function () { closeSheet(); refresh(); };
    var mi = document.getElementById('mInstall');
    if (mi) mi.onclick = function () { closeSheet(); state.installPrompt.prompt(); };
  }

  /* ---------------- Formulaires ---------------- */
  function busy(btn, on, label) { btn.disabled = on; if (label) btn.textContent = label; }

  function donationForm(caseId) {
    var cases = state.data.cases;
    if (!cases.length) { toast('Aucune collecte en cours.'); return; }
    var sel = caseId || '';
    var html = head('Don reçu') + '<form id="fDon" novalidate>' +
      '<label for="dCase">Collecte</label>' +
      '<select class="field" id="dCase" required><option value="">— Choisir la collecte —</option>' +
      cases.map(function (c) { return '<option value="' + esc(c.id) + '"' + (c.id === sel ? ' selected' : '') + '>' + esc(c.id + ' — ' + short(c.description, 60)) + '</option>'; }).join('') +
      '</select><div class="hint" id="dCaseInfo"></div>' +
      '<label for="dAmt">Montant (Dhs)</label><input class="field amount-input" id="dAmt" inputmode="decimal" autocomplete="off" placeholder="0" required>' +
      '<div class="quick">' + [100, 200, 500, 1000, 2000].map(function (v) { return '<button type="button" data-q="' + v + '">' + nf.format(v) + '</button>'; }).join('') + '</div>' +
      '<div class="two"><div><label for="dDate">Date</label><input class="field" id="dDate" type="date" value="' + todayStr() + '" max="' + todayStr() + '"></div>' +
      '<div><label for="dMode">Mode</label><select class="field" id="dMode">' + ['Espèces', 'Virement', 'Chèque', 'Carte', 'Autre'].map(function (m) { return '<option>' + m + '</option>'; }).join('') + '</select></div></div>' +
      '<label for="dDonor">Donateur <span class="opt">(facultatif)</span></label><input class="field" id="dDonor" placeholder="Anonyme" autocomplete="off">' +
      '<label for="dNote">Note <span class="opt">(facultatif)</span></label><input class="field" id="dNote" autocomplete="off">' +
      '<div class="err" id="dErr"></div>' +
      '<div class="form-foot"><button class="btn primary block" id="dSave">Enregistrer le don</button>' +
      '<button type="button" class="btn block" data-close>Annuler</button></div></form>';
    openSheet(html, caseId);
    var $case = document.getElementById('dCase'), $amt = document.getElementById('dAmt');
    function info() {
      var c = findCase($case.value);
      document.getElementById('dCaseInfo').textContent = c ? 'Déjà collecté : ' + dh(c.collected) + (c.budget ? ' sur ' + dh(c.budget) + ' — reste ' + dh(c.remaining) : '') : '';
    }
    $case.onchange = info; info();
    $sheetInner.querySelectorAll('[data-q]').forEach(function (b) { b.onclick = function () { $amt.value = b.dataset.q; }; });
    if (sel) setTimeout(function () { $amt.focus(); }, 250);
    document.getElementById('fDon').onsubmit = function (ev) {
      ev.preventDefault();
      var err = document.getElementById('dErr'), btn = document.getElementById('dSave');
      var amount = Number(String($amt.value).replace(/\s/g, '').replace(',', '.'));
      if (!$case.value) { err.textContent = 'Choisissez la collecte.'; return; }
      if (!(amount > 0)) { err.textContent = 'Indiquez un montant valide.'; return; }
      var c = findCase($case.value);
      if (!window.confirm('Enregistrer ' + dh(amount) + ' pour ' + c.id + ' ?')) return;
      busy(btn, true, 'Enregistrement…'); err.textContent = '';
      api('addDonation', {
        id: $case.value, amount: amount, date: document.getElementById('dDate').value,
        mode: document.getElementById('dMode').value, donor: document.getElementById('dDonor').value,
        note: document.getElementById('dNote').value
      }).then(function (r) {
        toast('Don de ' + dh(amount) + ' enregistré pour ' + c.id);
        if (r.case) replaceCase(r.case);
        if (state.seen) { state.seen[key(r.case || c)] = (r.case || c).collected; LS.set('seen', state.seen); }
        state.recent = null;
        closeSheet(); render(); refresh(true);
      }).catch(function (e) { busy(btn, false, 'Enregistrer le don'); err.textContent = e.message; });
    };
  }

  function replaceCase(nc) {
    var list = state.data.cases;
    for (var i = 0; i < list.length; i++) if (list[i].id === nc.id) { list[i] = nc; return; }
    list.push(nc);
  }

  function categoryField(current) {
    var cats = state.data.categories.slice();
    var known = !current || cats.indexOf(current) >= 0;
    return '<label for="cCat">Type de collecte</label><select class="field" id="cCat">' +
      '<option value="">— Choisir —</option>' + cats.map(function (k) { return '<option' + (k === current ? ' selected' : '') + '>' + esc(k) + '</option>'; }).join('') +
      '<option value="__new"' + (known ? '' : ' selected') + '>Autre (nouveau type)…</option></select>' +
      '<input class="field" id="cCatNew" style="margin-top:8px;' + (known ? 'display:none' : '') + '" placeholder="Nom du nouveau type" value="' + (known ? '' : esc(current)) + '">';
  }

  /** Création (c = null), modification (c = collecte en cours), relance d'un cas récurrent (from = collecte clôturée). */
  function caseForm(c, from) {
    var editing = !!c, base = c || from || {};
    var reuse = from ? from.id : '';
    var html = head(editing ? 'Modifier ' + c.id : 'Nouvelle collecte') + '<form id="fCase" novalidate>';
    if (!editing) {
      html += '<label>Cas récurrent ? <span class="opt">(même bénéficiaire qu\'une collecte clôturée)</span></label>' +
        '<input class="field" id="cReuseQ" placeholder="Rechercher dans l\'historique (texte ou ID)…" autocomplete="off" value="' + esc(reuse) + '">' +
        '<div class="pick" id="cPick" style="display:none"></div>' +
        '<div class="hint" id="cReuseInfo">' + (reuse ? 'Réutilise l\'ID ' + esc(reuse) + ' (nouvelle collecte n°' + (Math.max.apply(null, historyOf(reuse).map(function (h) { return h.cycle; })) + 1) + ')' : 'Laissez vide pour un nouveau cas : un ID sera créé automatiquement.') + '</div>';
    }
    html += categoryField(base.category) +
      '<label for="cDesc">Description</label><textarea class="field" id="cDesc" required>' + esc(base.description || '') + '</textarea>' +
      '<label for="cBud">Budget global nécessaire (Dhs) <span class="opt">(vide si pas d\'objectif fixe)</span></label>' +
      '<input class="field" id="cBud" inputmode="decimal" value="' + (editing && c.budget !== null ? esc(c.budget) : '') + '">' +
      '<label for="cSit">Situation actuelle <span class="opt">(facultatif)</span></label><textarea class="field" id="cSit" style="min-height:70px">' + esc(editing ? c.situation : '') + '</textarea>' +
      '<div class="err" id="cErr"></div><div class="form-foot"><button class="btn primary block" id="cSave">' + (editing ? 'Enregistrer' : 'Créer la collecte') + '</button>' +
      '<button type="button" class="btn block" data-close>Annuler</button></div></form>';
    openSheet(html, editing ? c.id : '');

    var $cat = document.getElementById('cCat'), $catNew = document.getElementById('cCatNew');
    $cat.onchange = function () { $catNew.style.display = $cat.value === '__new' ? '' : 'none'; if ($cat.value === '__new') $catNew.focus(); };

    if (!editing) {
      var $q = document.getElementById('cReuseQ'), $pick = document.getElementById('cPick'), $info = document.getElementById('cReuseInfo');
      var unique = {};
      state.data.closed.forEach(function (h) { if (!findCase(h.id) && (!unique[h.id] || h.cycle > unique[h.id].cycle)) unique[h.id] = h; });
      var pool = Object.keys(unique).map(function (k) { return unique[k]; });
      $q.oninput = function () {
        reuse = '';
        $info.textContent = 'Laissez vide pour un nouveau cas : un ID sera créé automatiquement.';
        var q = norm($q.value).trim();
        if (q.length < 2) { $pick.style.display = 'none'; return; }
        var res = pool.filter(function (h) { return norm(h.id + ' ' + h.description).indexOf(q) >= 0; }).slice(0, 30);
        $pick.style.display = res.length ? '' : 'none';
        $pick.innerHTML = res.map(function (h) {
          return '<button type="button" data-id="' + esc(h.id) + '"><b>' + esc(h.id) + '</b> · ' + esc(short(h.description, 90)) + '</button>';
        }).join('');
        $pick.querySelectorAll('[data-id]').forEach(function (b) {
          b.onclick = function () {
            var h = unique[b.dataset.id];
            reuse = h.id; $q.value = h.id; $pick.style.display = 'none';
            $info.textContent = 'Réutilise l\'ID ' + h.id + ' (nouvelle collecte n°' + (h.cycle + 1) + '). Vous pouvez adapter le texte et le budget.';
            document.getElementById('cDesc').value = h.description;
            if (state.data.categories.indexOf(h.category) >= 0) { $cat.value = h.category; $catNew.style.display = 'none'; }
            else { $cat.value = '__new'; $catNew.style.display = ''; $catNew.value = h.category; }
          };
        });
      };
    }

    document.getElementById('fCase').onsubmit = function (ev) {
      ev.preventDefault();
      var err = document.getElementById('cErr'), btn = document.getElementById('cSave');
      var cat = $cat.value === '__new' ? $catNew.value.trim() : $cat.value;
      var desc = document.getElementById('cDesc').value.trim();
      var budRaw = document.getElementById('cBud').value.replace(/\s/g, '').replace(',', '.');
      if (!cat) { err.textContent = 'Choisissez le type de collecte.'; return; }
      if (!desc) { err.textContent = 'La description est obligatoire.'; return; }
      if (budRaw && !(Number(budRaw) >= 0)) { err.textContent = 'Budget invalide.'; return; }
      var payload = { category: cat, description: desc, budget: budRaw === '' ? '' : Number(budRaw), situation: document.getElementById('cSit').value.trim() };
      busy(btn, true, 'Enregistrement…'); err.textContent = '';
      var p = editing ? api('updateCase', Object.assign({ id: c.id }, payload)) : api('createCase', Object.assign({ reuseId: reuse }, payload));
      p.then(function (r) {
        toast(editing ? 'Collecte mise à jour' : 'Collecte ' + r.case.id + ' créée');
        if (r.case) replaceCase(r.case);
        if (!editing && state.seen && r.case) { state.seen[key(r.case)] = r.case.collected; LS.set('seen', state.seen); }
        state.tab = 'open';
        refresh(true).then(function () { openCase(r.case.id); });
      }).catch(function (e) { busy(btn, false, editing ? 'Enregistrer' : 'Créer la collecte'); err.textContent = e.message; });
    };
  }

  function closeForm(c) {
    var html = head('Clôturer ' + c.id) +
      '<p class="muted small">La collecte sera déplacée vers l\'onglet « PROJETS.OLD CLAUDE » du Google Sheet avec un montant final de <b>' + dh(c.collected) + '</b>. Si ce cas revient, vous pourrez la relancer avec le même ID.</p>' +
      '<form id="fClose"><label for="xSit">Situation finale</label><textarea class="field" id="xSit">' +
      esc(c.situation ? c.situation + ' => Collecte clôturée' : 'Collecte clôturée') + '</textarea>' +
      '<div class="err" id="xErr"></div><div class="form-foot"><button class="btn primary block" id="xSave" style="background:var(--danger);border-color:var(--danger)">Clôturer la collecte</button>' +
      '<button type="button" class="btn block" data-close>Annuler</button></div></form>';
    openSheet(html, c.id);
    document.getElementById('fClose').onsubmit = function (ev) {
      ev.preventDefault();
      var btn = document.getElementById('xSave');
      busy(btn, true, 'Clôture…');
      api('closeCase', { id: c.id, situation: document.getElementById('xSit').value }).then(function () {
        toast(c.id + ' clôturée et archivée');
        closeSheet(); refresh();
      }).catch(function (e) { busy(btn, false, 'Clôturer la collecte'); document.getElementById('xErr').textContent = e.message; });
    };
  }

  /* ---------------- Démarrage ---------------- */
  window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); state.installPrompt = e; if (state.session) render(); });
  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState === 'hidden') saveSeen();
    else refresh(true);
  });
  window.addEventListener('online', function () { refresh(true); });
  setInterval(function () { if (document.visibilityState === 'visible' && !$sheet.open) refresh(true); }, REFRESH_MS);

  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(function () { /* facultatif */ });
  }

  render();
  if (state.session) refresh(true);
})();

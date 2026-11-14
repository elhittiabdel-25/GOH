// Serveur du « studio » vidéo : l'application (/app/), l'API émulée, et une vue du classeur émulé.
//   node test/video/server.js [port]
'use strict';
const http = require('http');
const fs = require('fs');
const path = require('path');
const { createContext } = require('../gas-mock');

const port = Number(process.argv[2] || 8090);
const fixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixture-video.json'), 'utf8'));
const g = createContext(fixture, { quiet: true, locale: 'fr_FR', ownerEmail: 'admin@exemple.org' });
g.installer();
const ss = g.__ss;
const mem = ss.getSheetByName('MEMBRES');
mem.getRange(2, 1, 1, 5).setValues([['Fatima (admin)', 'admin@exemple.org', 'admin', 'K7P2Q9XA', 'OUI']]);
mem.appendRow(['Youssef', 'youssef@exemple.org', 'membre', 'M4R8T2LB', 'OUI']);
mem.appendRow(['Salma', 'salma@exemple.org', 'membre', 'H9D3W6NE', 'OUI']);

const docs = path.join(__dirname, '..', '..', 'docs');
const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.webmanifest': 'application/manifest+json' };

function sheetJson(name) {
  const sh = ss.getSheetByName(name);
  const n = Math.max(sh.getLastRow(), 1), w = Math.max(sh.getLastColumn(), 1);
  const rng = sh.getRange(1, 1, n, w);
  const conv = (v) => (Object.prototype.toString.call(v) === '[object Date]' ? { d: v.toISOString() } : v);
  return { values: rng.getValues().map((r) => r.map(conv)), formulas: rng.getFormulas() };
}

function body(req) {
  return new Promise((ok) => { let b = ''; req.on('data', (c) => (b += c)); req.on('end', () => ok(b)); });
}

http.createServer(async (req, res) => {
  const url = req.url.split('?')[0];
  const send = (code, type, data) => { res.writeHead(code, { 'Content-Type': type, 'Cache-Control': 'no-store' }); res.end(data); };
  if (req.method === 'POST' && url === '/api') {
    const out = g.doPost({ postData: { contents: await body(req) } }).getContent();
    return setTimeout(() => send(200, 'application/json', out), 250);
  }
  if (url === '/dev/sheet') {
    const name = decodeURIComponent(new URL(req.url, 'http://x').searchParams.get('name'));
    return send(200, 'application/json', JSON.stringify(sheetJson(name)));
  }
  if (req.method === 'POST' && url === '/dev/edit') {
    const e = JSON.parse(await body(req));
    const sh = ss.getSheetByName(e.sheet);
    let v = e.value;
    if (typeof v === 'string' && /^\d{2}\/\d{2}\/\d{4}$/.test(v)) { const [d, m, y] = v.split('/'); v = new Date(+y, m - 1, +d); }
    else if (typeof v === 'string' && v !== '' && !isNaN(Number(v))) v = Number(v);
    sh.getRange(e.row, e.col).setValue(v);
    g.onEdit({ range: sh.getRange(e.row, e.col) });
    return send(200, 'application/json', '{"ok":true}');
  }
  if (url === '/app/config.js') return send(200, 'text/javascript', "window.GOH_CONFIG = { API_URL: '/api', REFRESH_SECONDS: 600 };");
  if (url === '/' || url === '/studio.html') return send(200, types['.html'], fs.readFileSync(path.join(__dirname, 'studio.html')));
  if (url.startsWith('/app/')) {
    const file = path.join(docs, decodeURIComponent(url.slice(5)) || 'index.html');
    if (file.startsWith(docs) && fs.existsSync(file) && fs.statSync(file).isFile()) return send(200, types[path.extname(file)] || 'application/octet-stream', fs.readFileSync(file));
  }
  send(404, 'text/plain', '404');
}).listen(port, () => console.log('studio http://localhost:' + port));

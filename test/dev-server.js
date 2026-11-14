// Serveur local : sert l'application (docs/) et simule l'API Apps Script sur un classeur émulé.
//   node test/dev-server.js [port]     (FIXTURE=... pour charger un autre classeur)
'use strict';
const http = require('http');
const fs = require('fs');
const path = require('path');
const { createContext } = require('./gas-mock');

const port = Number(process.argv[2] || 8080);
const fixture = JSON.parse(fs.readFileSync(process.env.FIXTURE || path.join(__dirname, 'fixture-demo.json'), 'utf8'));
const g = createContext(fixture, { quiet: true });
console.log(g.installer());
const ss = g.__ss.getSheetByName('MEMBRES');
ss.appendRow(['Membre démo', 'membre@example.org', 'membre', 'MEMBRE01', 'OUI']);
const admin = ss.getRange(2, 1, 1, 5).getValues()[0];
console.log(`Admin : ${admin[1]} / ${admin[3]}   Membre : membre@example.org / MEMBRE01`);

const root = path.join(__dirname, '..', 'docs');
const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.webmanifest': 'application/manifest+json' };
http.createServer((req, res) => {
  if (req.method === 'POST' && req.url === '/api') {
    let body = '';
    req.on('data', (c) => (body += c));
    req.on('end', () => {
      const out = g.doPost({ postData: { contents: body } }).getContent();
      setTimeout(() => { res.writeHead(200, { 'Content-Type': 'application/json' }); res.end(out); }, 150);
    });
    return;
  }
  if (req.url.startsWith('/config.js')) {
    res.writeHead(200, { 'Content-Type': 'text/javascript' });
    return res.end("window.GOH_CONFIG = { API_URL: '/api', REFRESH_SECONDS: 60 };");
  }
  const file = path.join(root, decodeURIComponent(req.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!file.startsWith(root) || !fs.existsSync(file)) { res.writeHead(404); return res.end('404'); }
  res.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream' });
  fs.createReadStream(file).pipe(res);
}).listen(port, () => console.log(`http://localhost:${port}`));

// 本機測試伺服器：提供 public/ 靜態檔，/api/parse 在 MOCK=1 時回傳 fixture，否則呼叫真正的 handler
const http = require('http');
const fs = require('fs');
const path = require('path');
const handler = require('../api/parse.js');
const PUB = path.join(__dirname, '../public');
const fixtures = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixture-results.json'), 'utf8'));
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css' };
let mockIdx = 0;

http.createServer((req, res) => {
  res.status = (c) => { res.statusCode = c; return res; };
  res.json = (o) => { res.setHeader('Content-Type', 'application/json'); res.end(JSON.stringify(o)); };
  if (req.url.startsWith('/api/parse')) {
    let body = '';
    req.on('data', (c) => (body += c));
    req.on('end', async () => {
      req.body = body ? JSON.parse(body) : {};
      if (process.env.MOCK === '1') {
        if (req.headers['x-app-password'] !== process.env.APP_PASSWORD) return res.status(401).json({ error: '密碼錯誤' });
        // 依圖片大小判斷是哪張（型錄圖比較大）
        const big = req.body.image.length > 200000;
        const fx = fixtures.find((f) => (big ? f.kind === 'catalog' : f.kind === 'orders'));
        mockIdx++;
        return setTimeout(() => res.status(200).json({ result: fx }), 500);
      }
      handler(req, res);
    });
    return;
  }
  const file = path.join(PUB, decodeURIComponent(req.url.split('?')[0]) === '/' ? 'index.html' : decodeURIComponent(req.url.split('?')[0]));
  if (!file.startsWith(PUB) || !fs.existsSync(file)) { res.statusCode = 404; return res.end('not found'); }
  res.setHeader('Content-Type', TYPES[path.extname(file)] || 'application/octet-stream');
  fs.createReadStream(file).pipe(res);
}).listen(process.env.PORT || 5173, () => console.log('http://localhost:' + (process.env.PORT || 5173)));

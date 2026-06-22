const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 7771;
const BASE = 'C:/Users/25114/WorkBuddy/20260622144609';

http.createServer((req, res) => {
  let f = path.join(BASE, req.url === '/' ? '/ai-mindmap.html' : req.url);
  try {
    const d = fs.readFileSync(f);
    const ext = path.extname(f);
    const mime = ext === '.html' ? 'text/html; charset=utf-8' : 'text/plain';
    res.writeHead(200, { 'Content-Type': mime });
    res.end(d);
  } catch (e) {
    res.writeHead(404);
    res.end('Not found');
  }
}).listen(PORT, () => {
  console.log('Server at http://localhost:' + PORT);
});

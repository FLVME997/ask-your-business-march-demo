const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const PUBLIC_DIR = path.join(__dirname, 'public');

const types = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.pdf': 'application/pdf',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

function send(res, status, body, type='text/plain; charset=utf-8') {
  res.writeHead(status, { 'Content-Type': type, 'Cache-Control': 'no-store' });
  res.end(body);
}

const server = http.createServer((req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    if (url.pathname === '/health') {
      return send(res, 200, JSON.stringify({ ok: true, app: 'Ask Your Business Multi-Month Import Demo' }), 'application/json; charset=utf-8');
    }
    let safePath = decodeURIComponent(url.pathname);
    if (safePath === '/') safePath = '/index.html';
    safePath = safePath.replace(/\.\./g, '');
    const filePath = path.join(PUBLIC_DIR, safePath);
    if (!filePath.startsWith(PUBLIC_DIR)) return send(res, 403, 'Forbidden');
    fs.readFile(filePath, (err, content) => {
      if (err) return send(res, 404, 'Not found');
      send(res, 200, content, types[path.extname(filePath).toLowerCase()] || 'application/octet-stream');
    });
  } catch (error) {
    send(res, 500, 'Server error');
  }
});

server.listen(PORT, () => console.log(`Ask Your Business multi-month demo running on port ${PORT}`));

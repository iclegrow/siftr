/**
 * Siftr Review Server — zero-dependency Node.js HTTP server
 * Serves the interactive learning UI and provides a REST API
 * for reading/saving triage classification corrections.
 *
 * Usage:
 *   node server.js <json-file-path> [--port 8473]
 *
 * API:
 *   GET  /              → serves the review UI (public/index.html)
 *   GET  /api/data      → returns the current triage JSON
 *   POST /api/save      → saves user corrections back to the JSON file
 *   POST /api/shutdown   → graceful shutdown
 */

const http = require('http');
const fs = require('fs');
const path = require('path');

// ---------------------------------------------------------------------------
// CLI args
// ---------------------------------------------------------------------------
const args = process.argv.slice(2);
let dataFilePath = null;
let port = 8473;

for (let i = 0; i < args.length; i++) {
  if (args[i] === '--port' && args[i + 1]) {
    port = parseInt(args[i + 1], 10);
    i++;
  } else if (!dataFilePath) {
    dataFilePath = args[i];
  }
}

if (!dataFilePath) {
  console.error('Usage: node server.js <json-file-path> [--port 8473]');
  process.exit(1);
}

dataFilePath = path.resolve(dataFilePath);

if (!fs.existsSync(dataFilePath)) {
  console.error(`Data file not found: ${dataFilePath}`);
  process.exit(1);
}

// ---------------------------------------------------------------------------
// Static file serving
// ---------------------------------------------------------------------------
const MIME_TYPES = {
  '.html': 'text/html',
  '.css': 'text/css',
  '.js': 'application/javascript',
  '.json': 'application/json',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
};

function serveStatic(res, filePath) {
  const ext = path.extname(filePath);
  const mime = MIME_TYPES[ext] || 'application/octet-stream';

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('Not found');
      return;
    }
    res.writeHead(200, { 'Content-Type': mime });
    res.end(data);
  });
}

// ---------------------------------------------------------------------------
// JSON helpers
// ---------------------------------------------------------------------------
function readData() {
  let raw = fs.readFileSync(dataFilePath, 'utf-8');
  // Strip UTF-8 BOM if present (PowerShell often writes one)
  if (raw.charCodeAt(0) === 0xFEFF) raw = raw.slice(1);
  return JSON.parse(raw);
}

function writeData(data) {
  fs.writeFileSync(dataFilePath, JSON.stringify(data, null, 2), 'utf-8');
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => {
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString()));
      } catch (e) {
        reject(e);
      }
    });
    req.on('error', reject);
  });
}

// ---------------------------------------------------------------------------
// Server
// ---------------------------------------------------------------------------
const publicDir = path.join(__dirname, 'public');

const server = http.createServer(async (req, res) => {
  // CORS (in case browser needs it for localhost)
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const url = new URL(req.url, `http://localhost:${port}`);

  // --- API routes ---
  if (url.pathname === '/api/data' && req.method === 'GET') {
    try {
      const data = readData();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(data));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  if (url.pathname === '/api/save' && req.method === 'POST') {
    try {
      const body = await readBody(req);
      // body is expected to be { emails: [...] } with updated overrides/notes
      const current = readData();
      current.emails = body.emails;
      current.lastReviewed = new Date().toISOString();
      writeData(current);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true, saved: current.emails.length }));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  if (url.pathname === '/api/shutdown' && req.method === 'POST') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: true, message: 'Shutting down' }));
    setTimeout(() => {
      server.close(() => process.exit(0));
    }, 200);
    return;
  }

  // --- Static files ---
  if (url.pathname === '/' || url.pathname === '/index.html') {
    serveStatic(res, path.join(publicDir, 'index.html'));
    return;
  }

  // Serve other static files from public/
  const safePath = path.normalize(url.pathname).replace(/^(\.\.[\/\\])+/, '');
  serveStatic(res, path.join(publicDir, safePath));
});

server.listen(port, '127.0.0.1', () => {
  console.log(`\n  🔍 Siftr Review Server running at http://localhost:${port}`);
  console.log(`  📄 Data file: ${dataFilePath}`);
  console.log(`  💡 Open the URL above in your browser to review classifications`);
  console.log(`  ⏹️  Press Ctrl+C or POST /api/shutdown to stop\n`);
});

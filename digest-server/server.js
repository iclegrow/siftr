/**
 * Siftr Digest Server — zero-dependency Node.js HTTP server
 * Serves the digest UI and provides a REST API for reading/saving
 * mark-read state on triaged emails.
 *
 * Usage:
 *   node server.js <json-file-path> [--port 8474]
 *
 * API:
 *   GET  /              → serves the digest UI (public/index.html)
 *   GET  /api/data      → returns the current digest JSON
 *   POST /api/save      → saves mark-read state back to the JSON file
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
let port = 8474;

for (let i = 0; i < args.length; i++) {
  if (args[i] === '--port' && args[i + 1]) {
    port = parseInt(args[i + 1], 10);
    i++;
  } else if (!dataFilePath) {
    dataFilePath = args[i];
  }
}

if (!dataFilePath) {
  console.error('Usage: node server.js <json-file-path> [--port 8474]');
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
// User identity from org cache
// ---------------------------------------------------------------------------
function getUserFirstName() {
  try {
    // Try to read the org cache which has manager info — the user is the manager's direct
    const orgCachePath = path.join(
      process.env.USERPROFILE || '',
      'OneDrive - Microsoft', 'AI-Tools', 'siftr_personal', 'org-cache.json'
    );
    let raw = fs.readFileSync(orgCachePath, 'utf-8');
    if (raw.charCodeAt(0) === 0xFEFF) raw = raw.slice(1);
    const cache = JSON.parse(raw);
    // The cache has manager.name (e.g. "Manager Person") — look for the user via peers
    // Since the cache doesn't store the user's name, check for a userName field
    if (cache.userName) return cache.userName.split(' ')[0];
  } catch { /* ignore */ }
  // Fallback: derive from Windows USERNAME (e.g. user.person → User)
  // Map of known usernames for this environment
  const knownUsers = { user.person: 'User' };
  const username = (process.env.USERNAME || '').toLowerCase();
  return knownUsers[username] || username.charAt(0).toUpperCase() + username.slice(1);
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
      data.userName = getUserFirstName();
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
      const current = readData();
      current.emails = body.emails;
      current.lastUpdated = new Date().toISOString();
      writeData(current);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true, saved: current.emails.length }));
    } catch (e) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }

  if (url.pathname === '/api/open-email' && req.method === 'POST') {
    try {
      const body = await readBody(req);
      const msgId = body.internetMessageId;
      if (!msgId) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'internetMessageId required' }));
        return;
      }
      // Write a temp PS1 script to avoid quote-escaping issues
      const tmpPs1 = path.join(require('os').tmpdir(), 'siftr-open-email.ps1');
      const safeMsgId = msgId.replace(/'/g, "''");
      const psScript = [
        `$ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')`,
        `$ns = $ol.GetNamespace('MAPI')`,
        `$inbox = $ns.GetDefaultFolder(6)`,
        `$filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x1035001F"" = '${safeMsgId}'"`,
        `$found = $inbox.Items.Find($filter)`,
        `if (-not $found) { foreach ($sub in $inbox.Folders) { $found = $sub.Items.Find($filter); if ($found) { break } } }`,
        `if ($found) { $found.Display(); Write-Output 'opened' } else { Write-Output 'not_found' }`,
      ].join('\n');
      fs.writeFileSync(tmpPs1, psScript, 'utf-8');
      const { execSync } = require('child_process');
      const result = execSync(`powershell -NoProfile -ExecutionPolicy Bypass -File "${tmpPs1}"`,
        { encoding: 'utf-8', timeout: 15000 }).trim();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: true, result }));
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

  const safePath = path.normalize(url.pathname).replace(/^(\.\.[\/\\])+/, '');
  serveStatic(res, path.join(publicDir, safePath));
});

server.listen(port, '127.0.0.1', () => {
  console.log(`\n  📬 Siftr Digest Server running at http://localhost:${port}`);
  console.log(`  📄 Data file: ${dataFilePath}`);
  console.log(`  💡 Open the URL above in your browser to scan your inbox digest`);
  console.log(`  ⏹️  Press Ctrl+C or POST /api/shutdown to stop\n`);
});

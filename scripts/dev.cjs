const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { createRuntime } = require('./gas-runtime.cjs');
const root = path.join(__dirname, '..');
const runtime = createRuntime();
const allowed = new Set(['getBootstrap', 'createTransaction', 'updateTransaction', 'deleteTransaction', 'listTransactions', 'getDashboard', 'savePaymentMethod', 'saveCategory']);
const shim = `<script>
document.addEventListener('DOMContentLoaded', () => { document.getElementById('dev-banner').hidden = false; });
window.google = { script: { get run() {
  let success = () => {}, failure = () => {};
  const runner = new Proxy({}, { get(_target, method) {
    if (method === 'withSuccessHandler') return handler => { success = handler; return runner; };
    if (method === 'withFailureHandler') return handler => { failure = handler; return runner; };
    return argument => fetch('/rpc', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ method, argument }) })
      .then(response => { if (!response.ok) throw new Error('RPC failed'); return response.json(); }).then(success).catch(failure);
  } }); return runner;
} } };
</script>`;
const server = http.createServer(async (request, response) => {
  response.setHeader('Cache-Control', 'no-store');
  response.setHeader('X-Content-Type-Options', 'nosniff');
  if (request.method === 'GET' && request.url === '/') {
    let html = fs.readFileSync(path.join(root, 'Index.html'), 'utf8').replace(/<\?!= include_\('([^']+)'\); \?>/g, (_, name) => fs.readFileSync(path.join(root, name + '.html'), 'utf8'));
    html = html.replace('<head>', '<head><meta name="viewport" content="width=device-width, initial-scale=1">' + shim);
    response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }); response.end(html); return;
  }
  if (request.method === 'POST' && request.url === '/rpc') {
    try {
      let body = '';
      for await (const chunk of request) { body += chunk; if (body.length > 100000) { response.writeHead(413); response.end(); return; } }
      const { method, argument } = JSON.parse(body);
      if (!allowed.has(method)) { response.writeHead(400); response.end(); return; }
      const result = runtime.invoke(method, argument);
      response.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' }); response.end(JSON.stringify(result));
    } catch (_) { response.writeHead(500); response.end(); }
    return;
  }
  response.writeHead(404); response.end();
});
const port = Number(process.env.PORT || 3000);
server.listen(port, '127.0.0.1', () => console.log(`Local validation server on port ${server.address().port}; in-memory Apps Script doubles, not Google Sheets. Ctrl+C to stop.`));

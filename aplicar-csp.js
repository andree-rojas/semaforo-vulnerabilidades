// Calcula los hashes sha256 de los <script> y <style> inline y escribe la CSP en la etiqueta <meta>.
// Uso: node aplicar-csp.js <index.html>   (volver a ejecutar después de cada cambio en el código)
const fs = require('fs');
const crypto = require('crypto');
const ruta = process.argv[2];
let html = fs.readFileSync(ruta, 'utf8').replace(/\r\n/g, '\n');
const hash = (t) => `'sha256-${crypto.createHash('sha256').update(t, 'utf8').digest('base64')}'`;
const scripts = [...html.matchAll(/<script id="[^"]+">([\s\S]*?)<\/script>/g)].map((m) => hash(m[1]));
const estilos = [...html.matchAll(/<style>([\s\S]*?)<\/style>/g)].map((m) => hash(m[1]));
const csp = [
  "default-src 'none'",
  `script-src https://cdn.sheetjs.com ${scripts.join(' ')}`,
  `style-src ${estilos.join(' ')}`,
  "img-src 'none'", "connect-src 'none'", "font-src 'none'", "object-src 'none'",
  "base-uri 'none'", "form-action 'none'"
].join('; ');
html = html.replace(/(<meta http-equiv="Content-Security-Policy" content=")[^"]*(")/, `$1${csp}$2`);
fs.writeFileSync(ruta, html, 'utf8');
console.log(csp);

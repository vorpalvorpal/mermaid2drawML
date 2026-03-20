#!/usr/bin/env node
//
// mermaid_bridge.js  (ES Module — requires "type":"module" in package.json)
//
// Called by the mermaid2drawml R package via processx.
// Reads a mermaid diagram file, outputs a single JSON object to stdout:
//   { "ast": <parsed AST from @mermaid-js/parser>, "svg": "<svg string>" }
//
// Usage:
//   node mermaid_bridge.js <path-to-diagram.mmd>
//
// Requirements (installed by setup_mermaid2drawml()):
//   @mermaid-js/mermaid-cli  (ESM)
//   @mermaid-js/parser       (ESM)
//
// Bare specifiers are resolved relative to this file's location, which is
// the same directory as node_modules/ after setup_mermaid2drawml() copies
// everything into place.
//

import fs   from 'fs';
import path from 'path';
import os   from 'os';

const mmdPath = process.argv[2];
if (!mmdPath) {
  process.stderr.write('Usage: node mermaid_bridge.js <diagram.mmd>\n');
  process.exit(1);
}

const code = fs.readFileSync(mmdPath, 'utf8').trim();

// ── 1. Parse AST ──────────────────────────────────────────────────────────
// Note: @mermaid-js/parser v0.3.x supports only a subset of diagram types
// (architecture, gitGraph, info, packet, pie). For flowcharts and others,
// the promise rejects with "Unknown diagram type" — that's expected and the
// SVG from mermaid-cli remains the authoritative source.
let ast = null;
try {
  const { parse } = await import('@mermaid-js/parser');
  // First token of the diagram code is the diagram type
  const diagType = code.split(/\s+/)[0].toLowerCase();
  ast = await parse(diagType, code);   // parse() returns a Promise
} catch (e) {
  // Unsupported diagram type or parse error — carry on; SVG is the primary output
  ast = { _parseError: e.message };
}

// ── 2. Render SVG via mermaid-cli ─────────────────────────────────────────
const { run } = await import('@mermaid-js/mermaid-cli');

const tmpSvg = path.join(os.tmpdir(), `mermaid2drawml_${process.pid}_${Date.now()}.svg`);
try {
  await run(mmdPath, tmpSvg, {
    outputFormat: 'svg',
    quiet: true,
    parseMMDOptions: { suppressErrors: true }
  });
} catch (e) {
  process.stderr.write('SVG render error: ' + e.message + '\n');
  process.exit(1);
}

const svg = fs.readFileSync(tmpSvg, 'utf8');
try { fs.unlinkSync(tmpSvg); } catch (_) {}

// ── 3. Output ──────────────────────────────────────────────────────────────
process.stdout.write(JSON.stringify({ ast, svg }));

#!/usr/bin/env node
//
// mermaid_bridge.js
//
// Called by the mermaid2drawml R package via processx.
// Reads a mermaid diagram file, outputs a single JSON object to stdout:
//   { "ast": <parsed AST from @mermaid-js/parser>, "svg": "<svg string>" }
//
// Usage:
//   node mermaid_bridge.js <path-to-diagram.mmd>
//
// Requirements (installed by setup_mermaid2drawml()):
//   @mermaid-js/mermaid-cli
//   @mermaid-js/parser
//

'use strict';

const fs   = require('fs');
const path = require('path');
const os   = require('os');

// Resolve deps relative to this script so they work whether called from the
// package library or from inst/node during development.
const nodeDir = __dirname;
function req(pkg) {
  return require(path.join(nodeDir, 'node_modules', pkg));
}

async function main() {
  const mmdPath = process.argv[2];
  if (!mmdPath) {
    process.stderr.write('Usage: node mermaid_bridge.js <diagram.mmd>\n');
    process.exit(1);
  }

  const code = fs.readFileSync(mmdPath, 'utf8').trim();

  // ── 1. Parse AST ──────────────────────────────────────────────────────────
  let ast = null;
  try {
    const { parse } = req('@mermaid-js/parser');
    // First token of the diagram code is the diagram type
    const diagType = code.split(/\s+/)[0].toLowerCase();
    ast = parse(diagType, code);
  } catch (e) {
    // If parsing fails (unsupported diagram type etc.) carry on; SVG is primary
    ast = { _parseError: e.message };
  }

  // ── 2. Render SVG via mermaid-cli ─────────────────────────────────────────
  const { run } = req('@mermaid-js/mermaid-cli');

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
}

main().catch(e => {
  process.stderr.write(e.stack || e.message || String(e));
  process.exit(1);
});

#!/usr/bin/env node
// Regenerate the bridge fixtures: render every tests/testthat/fixtures/*.mmd
// in ONE run of inst/node/mermaid_bridge.js and write <name>.json beside each
// .mmd, holding that diagram's bridge result without its `svg` field.
//
// Usage (from the package root):
//   node dev/make-fixtures.mjs [--config <config.json>]
//
// M2D_NODE_DIR defaults to the R user data directory that
// setup_mermaid2drawml() installs into; on macOS PUPPETEER_EXECUTABLE_PATH
// defaults to Google Chrome. Exits 1, naming the diagrams, if any fixture
// fails to render (the others are still written).
//
// Layout: one top-level key per line, and one prim / db node / db edge per
// line, so a diff shows which primitives changed. Key order is the bridge's.

import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const bridge = path.join(root, 'inst', 'node', 'mermaid_bridge.js');
const fixtures = path.join(root, 'tests', 'testthat', 'fixtures');

const defaultDepsDir = () => {
  const home = os.homedir();
  if (process.platform === 'darwin')
    return path.join(home, 'Library', 'Application Support', 'org.R-project.R', 'R', 'mermaid2drawml');
  if (process.platform === 'win32')
    return path.join(process.env.APPDATA || path.join(home, 'AppData', 'Roaming'), 'R', 'data', 'R', 'mermaid2drawml');
  return path.join(process.env.XDG_DATA_HOME || path.join(home, '.local', 'share'), 'R', 'mermaid2drawml');
};
const macChrome = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const childEnv = {
  ...process.env,
  M2D_NODE_DIR: process.env.M2D_NODE_DIR || defaultDepsDir(),
  ...(!process.env.PUPPETEER_EXECUTABLE_PATH && process.platform === 'darwin' && fs.existsSync(macChrome)
    ? { PUPPETEER_EXECUTABLE_PATH: macChrome } : {}),
};

const lines = (xs) => (xs.length === 0 ? '[]' : `[\n${xs.map((x) => JSON.stringify(x)).join(',\n')}\n]`);
const formatDb = (db) => (db === null ? 'null'
  : `{\n"nodes": ${lines(db.nodes)},\n"edges": ${lines(db.edges)}\n}`);
const formatValue = (k, v) => (k === 'prims' ? lines(v) : k === 'db' ? formatDb(v) : JSON.stringify(v));
const format = (res) => `{\n${Object.entries(res)
  .map(([k, v]) => `${JSON.stringify(k)}: ${formatValue(k, v)}`).join(',\n')}\n}\n`;

const withoutSvg = ({ svg, ...rest }) => rest;

const main = (argv) => {
  const names = fs.readdirSync(fixtures).filter((f) => f.endsWith('.mmd')).sort();
  if (names.length === 0) throw new Error(`no .mmd files in ${fixtures}`);
  const r = spawnSync(process.execPath, [bridge, ...argv, ...names.map((n) => path.join(fixtures, n))], {
    env: childEnv, encoding: 'utf8', maxBuffer: 1024 * 1024 * 1024, stdio: ['ignore', 'pipe', 'inherit'],
  });
  if (r.status !== 0) throw new Error(`bridge exited ${r.status}${r.error ? `: ${r.error.message}` : ''}`);
  const results = JSON.parse(r.stdout);
  if (results.length !== names.length) throw new Error(`bridge returned ${results.length} results for ${names.length} inputs`);
  const failed = names.flatMap((name, i) => {
    const res = results[i];
    if (!res.ok) return [`${name}: ${res.error.message.split('\n')[0]}`];
    const out = path.join(fixtures, name.replace(/\.mmd$/, '.json'));
    fs.writeFileSync(out, format(withoutSvg(res)));
    process.stderr.write(`wrote ${path.relative(root, out)} (${res.prims.length} prims)\n`);
    return [];
  });
  if (failed.length > 0) {
    process.stderr.write(`\n${failed.length} fixture(s) did not render:\n  ${failed.join('\n  ')}\n`);
    process.exitCode = 1;
  }
};

try {
  main(process.argv.slice(2));
} catch (e) {
  process.stderr.write(`make-fixtures: ${e.message}\n`);
  process.exitCode = 1;
}

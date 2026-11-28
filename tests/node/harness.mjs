// Shared set-up for the bridge tests (inst/node/mermaid_bridge.js): where the
// Node dependencies and the browser are, how to start the bridge, and how to
// read what it prints. Each bridge run starts headless Chrome, so test files
// share one run between many tests.
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
export const bridge = path.resolve(here, '..', '..', 'inst', 'node', 'mermaid_bridge.js');

function defaultDepsDir() {
  const home = os.homedir();
  if (process.platform === 'darwin')
    return path.join(home, 'Library', 'Application Support', 'org.R-project.R', 'R', 'mermaid2drawml');
  if (process.platform === 'win32')
    return path.join(process.env.APPDATA || path.join(home, 'AppData', 'Roaming'), 'R', 'data', 'R', 'mermaid2drawml');
  return path.join(process.env.XDG_DATA_HOME || path.join(home, '.local', 'share'), 'R', 'mermaid2drawml');
}
export const depsDir = process.env.M2D_NODE_DIR || defaultDepsDir();
export const haveDeps = fs.existsSync(path.join(depsDir, 'node_modules', '@mermaid-js', 'mermaid-cli'));
export const skipReason = haveDeps ? false
  : `mermaid-cli not found under ${depsDir}/node_modules; set M2D_NODE_DIR or install the dependencies`;

export const childEnv = { ...process.env, M2D_NODE_DIR: depsDir };
const macChrome = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
if (!childEnv.PUPPETEER_EXECUTABLE_PATH && process.platform === 'darwin' && fs.existsSync(macChrome))
  childEnv.PUPPETEER_EXECUTABLE_PATH = macChrome;

export function runBridge(args) {
  return spawnSync(process.execPath, [bridge, ...args], {
    env: childEnv, encoding: 'utf8', timeout: 120000, maxBuffer: 256 * 1024 * 1024,
  });
}
export function parseOut(r) {
  assert.equal(r.status, 0, `bridge exited ${r.status}; stderr: ${(r.stderr || '').slice(-2000)}`);
  return JSON.parse(r.stdout);
}

// The Font Awesome icon packs the bridge registers with Mermaid (V9) are
// optional in the Node directory; tests that need them skip without them.
const iconPacks = ['fa6-solid', 'fa6-regular', 'fa6-brands'];
const missingIcons = iconPacks.filter((n) => !fs.existsSync(path.join(depsDir, 'node_modules', '@iconify-json', n)));
export const iconSkipReason = skipReason || (missingIcons.length === 0 ? false
  : `icon packs @iconify-json/${missingIcons.join(', ')} not installed under ${depsDir}/node_modules`);

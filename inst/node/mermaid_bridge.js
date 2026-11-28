#!/usr/bin/env node
// mermaid_bridge.js: render Mermaid diagrams in headless Chrome and report
// what the browser drew, so that R never has to guess styles, transforms or
// shape identity from SVG text. ES module ("type": "module" in package.json).
//
// Usage
//   node mermaid_bridge.js [--config <config.json>] <in1.mmd> [<in2.mmd> ...]
//
//   Dependencies (puppeteer, mermaid, @mermaid-js/mermaid-cli and
//   @mermaid-js/mermaid-zenuml) are found under $M2D_NODE_DIR/node_modules, or
//   next to this script when M2D_NODE_DIR is unset. Chrome comes from
//   $PUPPETEER_EXECUTABLE_PATH, else puppeteer's own. --config names a JSON
//   object spread into mermaid.initialize({ startOnLoad: false, ...config }).
//   One browser for all inputs, one fresh 1600x1200 page (scale 1) per input.
//
//   stdout: a JSON array, one result per input, in input order; nothing else.
//   Exit 0 whenever the browser ran (failed diagrams say so in their result);
//   2 for bad arguments or unreadable inputs, 1 for missing dependencies or a
//   browser failure, with a message on stderr.
//
// Result (bridge contract 3)
//   failure: { ok: false, error: { message } }   (syntax error, render error,
//            or "timed out after 60 s"; never Mermaid's error picture)
//   success: { ok: true, bridge: 2, mermaid: "<version>", type, role,
//              viewBox: [x, y, w, h], title, desc, svg, prims, markers, db }
//     type     diagram.type from mermaidAPI.getDiagramFromText
//     role     the root svg's aria-roledescription
//     title, desc  text of the root <title>/<desc> (accTitle/accDescr) or null
//     db       { nodes, edges }: JSON-safe copy of diagram.db.getData(), or
//              null when the diagram's db has no getData. It comes from a
//              second parse in a fresh mermaid bundle with Math.random reseeded
//              (as it is for the render), so node domIds and random ids match
//              the ids in the svg.
//     markers  { "<marker id>": { viewBox: [x, y, w, h] | null, refX, refY,
//              w, h (markerWidth/Height, default 3), units ("strokeWidth" |
//              "userSpaceOnUse"), orient ("auto" | "auto-start-reverse" | degrees),
//              prims: [shape prims] } }, only for markers some prim refers to.
//              A marker prim has kind, tag, cls, geom, m, bbox, fill, fill_a,
//              stroke, stroke_a, stroke_w, dash; m and bbox are in marker
//              content coordinates (only transforms inside the marker are
//              composed) and stroke_w is in those units. A paint of
//              context-fill / context-stroke is the string "context".
//   All coordinates are in the root svg's user space, after every transform.
//   Ids have the render id prefix ("m2d_" / "m2d-") removed. Numbers are
//   rounded: coordinates 3 dp, matrix a-d 6 dp, alphas 4 dp.
//
// Prims: one per drawn element, in paint (document) order. Nothing from defs,
// marker, clipPath, mask, symbol, pattern, style, title, desc, metadata; no
// display:none, non-visible, zero-opacity, unpainted or zero-size element.
//   every prim  i, kind ("rect" | "ellipse" | "line" | "poly" | "path" |
//               "text" | "image" | "other"), tag, id, cls, data {data-* attrs},
//               n (DOM serial: index in document order among all elements
//               under the root svg, root = 0; unique per element),
//               anc [{ n, id, cls, data?, href? }] (g/a/svg/switch ancestors,
//               outermost first, root excluded; unnamed ones carry n alone,
//               with id and cls empty), bbox [x, y, w, h]
//   shapes      m [a, b, c, d, e, f] (own user space -> root),
//               geom: rect {x, y, w, h, rx, ry} | ellipse {cx, cy, rx, ry} |
//                     line {x1, y1, x2, y2} | poly {points, closed} | path {d},
//               fill "#RRGGBB" | null, fill_a, fill_approx (gradient/pattern),
//               stroke, stroke_a, stroke_w (root units), dash [..] | null,
//               marker_start, marker_end (marker ids) | null.
//               When the element or an ancestor has a computed filter that is
//               exactly one brightness(v), fill and stroke colours (and HTML
//               box colours) are multiplied by the product of the v, clamped;
//               alphas and text colours are untouched.
//               HTML boxes with a background inside a foreignObject are rect
//               prims with html: true and identity m, placed just before that
//               foreignObject's text prim.
//   text        html (foreignObject or SVG <text>), anchor ("start" |
//               "middle" | "end"), rot (degrees, clockwise), bbox, rows
//               [{ bbox, runs: [{ t, bold, italic, underline, color, size,
//               family, font }] }]. For rotated text every box is the
//               unrotated box, centred where the rotated text is.
//   image       href, m          other   (tag and bbox only)

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const BRIDGE_VERSION = 3;
const SVG_ID = 'm2d';
const PAGE_TIMEOUT_MS = 60000;
const RANDOM_SEED = 20240101;
const VIEWPORT = { width: 1600, height: 1200, deviceScaleFactor: 1 };
const USAGE = 'usage: node mermaid_bridge.js [--config <config.json>] <in1.mmd> [<in2.mmd> ...]';

// ---- Node driver ------------------------------------------------------------

const fail = (message, exitCode) => Object.assign(new Error(message), { exitCode });

const parseArgs = (args, acc = { config: null, inputs: [] }) => {
  if (args.length === 0) return acc;
  const [a, ...rest] = args;
  if (a === '--config') {
    if (rest.length === 0) throw fail('--config needs a file name', 2);
    return parseArgs(rest.slice(1), { ...acc, config: rest[0] });
  }
  if (a.startsWith('--config=')) return parseArgs(rest, { ...acc, config: a.slice(9) });
  if (a.startsWith('--')) throw fail(`unknown option ${a}`, 2);
  return parseArgs(rest, { ...acc, inputs: [...acc.inputs, a] });
};

const readText = (file, what) => {
  try {
    return fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, '');
  } catch (e) {
    throw fail(`cannot read ${what} ${file}: ${e.message}`, 2);
  }
};

const readConfig = (file) => {
  if (file === null) return {};
  const text = readText(file, 'config');
  const parsed = (() => {
    try { return JSON.parse(text); } catch (e) { throw fail(`config ${file} is not JSON: ${e.message}`, 2); }
  })();
  if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed))
    throw fail(`config ${file} must hold a JSON object`, 2);
  return parsed;
};

const resolveDeps = (dir) => {
  const nm = path.join(dir, 'node_modules');
  const files = {
    html: path.join(nm, '@mermaid-js', 'mermaid-cli', 'dist', 'index.html'),
    mermaid: path.join(nm, 'mermaid', 'dist', 'mermaid.js'),
    zenuml: path.join(nm, '@mermaid-js', 'mermaid-zenuml', 'dist', 'mermaid-zenuml.js'),
    mermaidPkg: path.join(nm, 'mermaid', 'package.json'),
  };
  const missing = Object.values(files).filter((f) => !fs.existsSync(f));
  if (missing.length > 0)
    throw fail(`Node dependencies not found under ${nm} (missing ${missing.join(', ')}); ` +
      'set M2D_NODE_DIR to the directory holding node_modules/', 1);
  const puppeteer = (() => {
    try {
      const mod = createRequire(path.join(dir, 'package.json'))('puppeteer');
      return mod.default ?? mod;
    } catch (e) { throw fail(`cannot load puppeteer from ${nm}: ${e.message}`, 1); }
  })();
  const mermaidVersion = JSON.parse(fs.readFileSync(files.mermaidPkg, 'utf8')).version;
  return { puppeteer, mermaidVersion, files };
};

// Resolve with the promise, or with onTimeout() after ms.
const withTimeout = (promise, ms, onTimeout) => {
  const timer = { id: null };
  const timeout = new Promise((resolve) => { timer.id = setTimeout(() => resolve(onTimeout()), ms); });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer.id));
};

const mapSeries = (xs, f) =>
  xs.reduce(async (acc, x, i) => [...(await acc), await f(x, i)], Promise.resolve([]));

const failure = (message) => ({ ok: false, error: { message } });

// Runs in the page before any script. Mermaid draws some ids from Math.random
// (state dividers, gitGraph commits) and roughjs seeds from it. A seeded
// generator whose state can be reset, also through references captured at
// bundle load, makes output repeatable and lets the db parse reproduce the
// render's random ids.
function installSeededRandom(seed) {
  const state = { s: seed >>> 0 };
  Math.random = () => {
    state.s = (state.s + 0x6D2B79F5) >>> 0;
    const t1 = Math.imul(state.s ^ (state.s >>> 15), 1 | state.s);
    const t2 = (t1 + Math.imul(t1 ^ (t1 >>> 7), 61 | t1)) ^ t1;
    return ((t2 ^ (t2 >>> 14)) >>> 0) / 4294967296;
  };
  globalThis.__m2dReseed = () => { state.s = seed >>> 0; };
}

const renderOnPage = async (page, deps, code, config) => {
  await page.setViewport(VIEWPORT);
  await page.evaluateOnNewDocument(installSeededRandom, RANDOM_SEED);
  await page.goto(pathToFileURL(deps.files.html).href);
  await page.addScriptTag({ path: deps.files.mermaid });
  await page.addScriptTag({ path: deps.files.zenuml });
  const res = await page.evaluate(pageMain, { code, config, svgId: SVG_ID });
  if (!res.ok) return res;
  await page.addScriptTag({ path: deps.files.mermaid });
  const { type, db } = await page.evaluate(pageDiagram, { code, config });
  const { ok, ...rest } = res;
  return { ok, bridge: BRIDGE_VERSION, mermaid: deps.mermaidVersion, type, ...rest, db };
};

const renderOne = async (browser, deps, code, config, label) => {
  const page = await browser.newPage();
  page.on('console', (msg) => process.stderr.write(`[${label}] ${msg.text()}\n`));
  page.on('pageerror', (err) => process.stderr.write(`[${label}] page error: ${err.message}\n`));
  try {
    return await withTimeout(renderOnPage(page, deps, code, config), PAGE_TIMEOUT_MS,
      () => failure(`timed out after ${PAGE_TIMEOUT_MS / 1000} s`));
  } catch (e) {
    return failure(`bridge: ${e.message}`);
  } finally {
    await withTimeout(page.close().catch(() => null), 5000, () => null);
  }
};

const main = async (argv, env) => {
  const args = parseArgs(argv);
  if (args.inputs.length === 0) throw fail(`no input files\n${USAGE}`, 2);
  const config = readConfig(args.config);
  const codes = args.inputs.map((f) => readText(f, 'input'));
  const here = path.dirname(fileURLToPath(import.meta.url));
  const deps = resolveDeps(env.M2D_NODE_DIR || here);

  const browser = await deps.puppeteer.launch({
    headless: 'shell',
    executablePath: env.PUPPETEER_EXECUTABLE_PATH || undefined,
  }).catch((e) => { throw fail(`cannot start the browser: ${e.message}`, 1); });
  const results = await (async () => {
    try {
      return await mapSeries(codes, (code, i) =>
        renderOne(browser, deps, code, config, path.basename(args.inputs[i])));
    } finally {
      await browser.close();
    }
  })();
  process.stdout.write(JSON.stringify(results) + '\n');
};

// ---- In-page walker -----------------------------------------------------------
// Runs inside the page (serialised by page.evaluate), so everything it needs
// is defined within it. It knows SVG, CSS and HTML, not diagram types.

async function pageMain({ code, config, svgId }) {
  const errMessage = (e) => (e && (e.message || e.str)) || String(e);

  // -- numbers and matrices (m = [a, b, c, d, e, f]) --
  const rnd = (k) => (x) => Math.round(x * k) / k;
  const r3 = rnd(1e3), r4 = rnd(1e4), r6 = rnd(1e6);
  const num = (s, dflt) => { const v = parseFloat(s); return Number.isFinite(v) ? v : dflt; };
  const fromDOM = (M) => [M.a, M.b, M.c, M.d, M.e, M.f];
  const mul = ([a, b, c, d, e, f], [A, B, C, D, E, F]) =>
    [a * A + c * B, b * A + d * B, a * C + c * D, b * C + d * D, a * E + c * F + e, b * E + d * F + f];
  const inv = ([a, b, c, d, e, f]) => {
    const det = a * d - b * c;
    return [d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det];
  };
  const apply = ([a, b, c, d, e, f], [x, y]) => [a * x + c * y + e, b * x + d * y + f];
  const scaleOf = ([a, b, c, d]) => Math.sqrt(Math.abs(a * d - b * c));
  const rotOf = ([a, b]) => { const deg = r3(Math.atan2(b, a) * 180 / Math.PI); return deg === -180 ? 180 : deg || 0; };
  const roundM = ([a, b, c, d, e, f]) => [r6(a), r6(b), r6(c), r6(d), r3(e), r3(f)];

  // -- boxes [x, y, w, h] --
  const boxOfPoints = (pts) => {
    const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
    const x = Math.min(...xs), y = Math.min(...ys);
    return [x, y, Math.max(...xs) - x, Math.max(...ys) - y];
  };
  const corners = ([x, y, w, h]) => [[x, y], [x + w, y], [x, y + h], [x + w, y + h]];
  const boxThrough = (m, box) => boxOfPoints(corners(box).map((p) => apply(m, p)));
  const unionBoxes = (boxes) => boxOfPoints(boxes.flatMap(corners));
  // Unrotated box: centre mapped through m, sides scaled by m's axis scales.
  const placeBox = (m, [x, y, w, h]) => {
    const [cx, cy] = apply(m, [x + w / 2, y + h / 2]);
    const sx = Math.hypot(m[0], m[1]), sy = Math.hypot(m[2], m[3]);
    return [cx - (w * sx) / 2, cy - (h * sy) / 2, w * sx, h * sy];
  };
  const roundBox = (b) => b.map(r3);
  const rectBox = (r) => [r.left, r.top, r.width, r.height];

  // -- colours and paint --
  const ctx2d = document.createElement('canvas').getContext('2d');
  const canonColour = (s) => {
    ctx2d.fillStyle = '#010203'; ctx2d.fillStyle = s; const a = ctx2d.fillStyle;
    ctx2d.fillStyle = '#040506'; ctx2d.fillStyle = s; const b = ctx2d.fillStyle;
    return a === b ? a : null;
  };
  const hex2 = (v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0');
  const parseColour = (s) => {
    if (!s || s === 'none') return null;
    const c = /^(rgba?\(|#)/.test(s.trim()) ? s.trim() : canonColour(s.trim());
    if (!c) return null;
    const h = /^#([0-9a-f]{6})$/i.exec(c);
    if (h) return { hex: `#${h[1].toUpperCase()}`, a: 1 };
    const m = /^rgba?\(([^)]*)\)$/.exec(c);
    if (!m) return null;
    const parts = m[1].split(/[\s,/]+/).filter(Boolean);
    const chan = (p) => (p.endsWith('%') ? (num(p, 0) * 255) / 100 : num(p, 0));
    const alpha = parts.length > 3 ? (parts[3].endsWith('%') ? num(parts[3], 100) / 100 : num(parts[3], 1)) : 1;
    return { hex: `#${parts.slice(0, 3).map(chan).map(hex2).join('')}`.toUpperCase(), a: alpha };
  };
  const NO_PAINT = { colour: null, a: 0, approx: false };
  const urlId = (s) => { const m = /url\(\s*["']?#([^"')]+)["']?\s*\)/.exec(s || ''); return m ? m[1] : null; };
  const gradientStop = (el, depth = 0) => {
    if (!el || depth > 4) return null;
    const stop = el.querySelector('stop');
    if (stop) {
      const cs = getComputedStyle(stop);
      const c = parseColour(cs.stopColor);
      return c ? { hex: c.hex, a: c.a * num(cs.stopOpacity, 1) } : null;
    }
    const href = el.getAttribute('href') || el.getAttribute('xlink:href');
    return href && href.startsWith('#') ? gradientStop(document.getElementById(href.slice(1)), depth + 1) : null;
  };
  // Paint from a computed fill/stroke value; opacity = property opacity x effective opacity.
  const paintOf = (value, opacity) => {
    const id = urlId(value);
    if (id) {
      const ref = document.getElementById(id);
      if (!ref) return paintOf(value.replace(/^\s*url\([^)]*\)\s*/, '') || 'none', opacity);
      const stop = /gradient$/i.test(ref.tagName) ? gradientStop(ref) : null;
      const c = stop || { hex: '#808080', a: 1 };
      return c.a * opacity > 0 ? { colour: c.hex, a: r4(c.a * opacity), approx: true } : NO_PAINT;
    }
    const c = parseColour(value);
    return c && c.a * opacity > 0 ? { colour: c.hex, a: r4(c.a * opacity), approx: false } : NO_PAINT;
  };
  // Marker content may paint with context-fill / context-stroke.
  const CONTEXT_PAINT = { colour: 'context', a: 1, approx: false };
  const BRIGHT = /^brightness\(\s*(-?[\d.]+)(%?)\s*\)$/;
  const brightnessOf = (cs) => {
    const m = BRIGHT.exec((cs.filter || '').trim());
    return m ? num(m[1], 1) / (m[2] ? 100 : 1) : 1;
  };
  const brighten = (paint, k) => {
    if (k === 1 || !/^#[0-9A-F]{6}$/.test(paint.colour || '')) return paint;
    const c = [1, 3, 5].map((i) => parseInt(paint.colour.slice(i, i + 2), 16) * k);
    return { ...paint, colour: `#${c.map(hex2).join('')}`.toUpperCase() };
  };
  const parseDash = (s, scale) => {
    if (!s || s === 'none') return null;
    const v = s.split(/[\s,]+/).filter(Boolean).map((p) => parseFloat(p));
    // "1, 0" (Mermaid's markers) has no gaps: a solid line.
    const gapless = v.length % 2 === 0 && v.every((x, k) => k % 2 === 0 || x === 0);
    return v.length === 0 || v.some((x) => !Number.isFinite(x) || x < 0) || v.every((x) => x === 0) || gapless
      ? null : v.map((x) => r3(x * scale));
  };

  // -- fonts --
  const GENERIC = {
    'sans-serif': 'Arial', serif: 'Times New Roman', monospace: 'Courier New', cursive: 'Comic Sans MS',
    fantasy: 'Impact', 'system-ui': 'Segoe UI', 'ui-sans-serif': 'Arial', 'ui-serif': 'Times New Roman',
    'ui-monospace': 'Courier New', 'ui-rounded': 'Arial', math: 'Cambria Math', emoji: 'Segoe UI Emoji',
  };
  const splitFamilies = (s) =>
    (s.match(/"[^"]*"|'[^']*'|[^,]+/g) || []).map((f) => f.trim().replace(/^["']|["']$/g, '').trim()).filter(Boolean);
  const PROBE = 'mmmmmmmmmmlli1WQ@#';
  const widthIn = (font) => { ctx2d.font = font; return ctx2d.measureText(PROBE).width; };
  const isInstalled = (name) =>
    ['monospace', 'serif', 'sans-serif'].some((g) => widthIn(`72px "${name}", ${g}`) !== widthIn(`72px ${g}`));
  const fontCache = new Map();
  const concreteFont = (family) => {
    if (!fontCache.has(family)) {
      const pick = splitFamilies(family).map((f) => GENERIC[f.toLowerCase()] || (isInstalled(f) ? f : null)).find(Boolean);
      fontCache.set(family, pick || 'Arial');
    }
    return fontCache.get(family);
  };

  // -- element facts --
  const stripId = (id) => (id || '').replace(new RegExp(`^${svgId}[_-]`), '');
  // Long values (Mermaid's base64 `data-points` on edges) identify nothing
  // and would dominate the output, so they are left out.
  const dataOf = (el) => Object.fromEntries([...el.attributes]
    .filter((a) => a.name.startsWith('data-') && a.value.length <= 120)
    .map((a) => [a.name.slice(5), a.value]));
  const clsOf = (el) => el.getAttribute('class') || '';
  const ancEntry = (el, ctx) => {
    const id = stripId(el.getAttribute('id')), cls = clsOf(el), data = dataOf(el);
    const href = el.tagName.toLowerCase() === 'a' ? (el.getAttribute('href') || el.getAttribute('xlink:href') || '') : '';
    const hasData = Object.keys(data).length > 0;
    return { n: ctx.serial.get(el), id, cls, ...(hasData ? { data } : {}), ...(href ? { href } : {}) };
  };
  const head = (kind, tag, el, ctx, bbox) =>
    ({ kind, tag, n: ctx.serial.get(el), id: stripId(el.getAttribute ? el.getAttribute('id') : el.id), cls: clsOf(el), data: dataOf(el), anc: ctx.anc, bbox: roundBox(bbox) });
  const safeBBox = (el) => { try { const b = el.getBBox(); return [b.x, b.y, b.width, b.height]; } catch (e) { return null; } };

  // -- the root --
  const container = document.getElementById('container');
  const rootInvOf = (root) => inv(fromDOM(root.getScreenCTM()));
  const matrixOf = (rootInv, el) => { const M = el.getScreenCTM(); return M ? mul(rootInv, fromDOM(M)) : null; };

  // -- shapes --
  const GEOM = {
    rect: (el, cs, bb) => {
      if (!(bb[2] > 0 && bb[3] > 0)) return null;
      const rx0 = num(cs.rx, NaN), ry0 = num(cs.ry, NaN);
      const rx = Number.isFinite(rx0) ? rx0 : Number.isFinite(ry0) ? ry0 : 0;
      const ry = Number.isFinite(ry0) ? ry0 : rx;
      return { x: r3(bb[0]), y: r3(bb[1]), w: r3(bb[2]), h: r3(bb[3]), rx: r3(Math.min(rx, bb[2] / 2)), ry: r3(Math.min(ry, bb[3] / 2)) };
    },
    ellipse: (el, cs, bb) => (bb[2] > 0 && bb[3] > 0
      ? { cx: r3(bb[0] + bb[2] / 2), cy: r3(bb[1] + bb[3] / 2), rx: r3(bb[2] / 2), ry: r3(bb[3] / 2) } : null),
    line: (el) => ({ x1: r3(el.x1.baseVal.value), y1: r3(el.y1.baseVal.value), x2: r3(el.x2.baseVal.value), y2: r3(el.y2.baseVal.value) }),
    poly: (el) => {
      const points = [...el.points].map((p) => [r3(p.x), r3(p.y)]);
      return points.length >= 2 ? { points, closed: el.tagName.toLowerCase() === 'polygon' } : null;
    },
    path: (el) => { const d = (el.getAttribute('d') || '').trim(); return d ? { d } : null; },
  };
  const SHAPE_KIND = { rect: 'rect', circle: 'ellipse', ellipse: 'ellipse', line: 'line', polygon: 'poly', polyline: 'poly', path: 'path' };
  const MARKABLE = new Set(['line', 'poly', 'path']);

  const shapePrim = (el, tag, cs, ctx) => {
    const kind = SHAPE_KIND[tag];
    const bb = safeBBox(el);
    const m = ctx.mOf(el);
    if (!bb || !m) return [];
    const geom = GEOM[kind](el, cs, bb);
    if (!geom) return [];
    const scale = scaleOf(m);
    const shapePaint = (v, o) => (/^context-(fill|stroke)$/.test(v) && ctx.inMarker ? CONTEXT_PAINT : brighten(paintOf(v, o), ctx.br));
    const fill = kind === 'line' ? NO_PAINT : shapePaint(cs.fill, num(cs.fillOpacity, 1) * ctx.op);
    const width = num(cs.strokeWidth, 0);
    const stroke = width > 0 ? shapePaint(cs.stroke, num(cs.strokeOpacity, 1) * ctx.op) : NO_PAINT;
    if (fill.colour === null && stroke.colour === null) return [];
    const marker = (v) => (MARKABLE.has(kind) && v && v !== 'none' ? stripId(urlId(v)) : null);
    return [{
      ...head(kind, tag, el, ctx, boxThrough(m, bb)),
      m: roundM(m), geom,
      fill: fill.colour, fill_a: fill.a, fill_approx: fill.approx,
      stroke: stroke.colour, stroke_a: stroke.a,
      stroke_w: stroke.colour ? r3(width * scale) : 0,
      dash: stroke.colour ? parseDash(cs.strokeDasharray, scale) : null,
      marker_start: marker(cs.markerStart), marker_end: marker(cs.markerEnd),
    }];
  };

  // -- text --
  const textNodesIn = (el, skipTags) => [...el.childNodes].flatMap((n) => {
    if (n.nodeType === Node.TEXT_NODE) return [n];
    return n.nodeType === Node.ELEMENT_NODE && !skipTags.has(n.tagName.toLowerCase()) ? textNodesIn(n, skipTags) : [];
  });
  const isUnderlined = (el, stop) =>
    !!el && (getComputedStyle(el).textDecorationLine.includes('underline') || (el !== stop && isUnderlined(el.parentElement, stop)));
  const styleOf = (el, html, scale, stop) => {
    const cs = getComputedStyle(el);
    const colour = (html ? parseColour(cs.color) : parseColour(cs.fill) || parseColour(cs.color)) || { hex: '#000000' };
    return {
      bold: num(cs.fontWeight, 400) >= 600 || cs.fontWeight === 'bold',
      italic: /italic|oblique/.test(cs.fontStyle),
      underline: isUnderlined(el, stop),
      color: colour.hex,
      size: r3(num(cs.fontSize, 16) * scale),
      family: cs.fontFamily,
      font: concreteFont(cs.fontFamily),
    };
  };
  const sameStyle = (a, b) => ['bold', 'italic', 'underline', 'color', 'size', 'family', 'font'].every((k) => a[k] === b[k]);
  const PRESERVE = /^(pre|pre-wrap|break-spaces)$/;

  // Words of a text container: { t, style, box (in the container's local frame), space }.
  const wordsOf = (el, html, frame, toLocal) => {
    const nodes = textNodesIn(el, html ? new Set() : new Set(['title', 'desc']));
    const offsets = nodes.reduce((acc, n) => [...acc, acc[acc.length - 1] + n.data.length], [0]);
    const full = nodes.map((n) => n.data).join('');
    const scale = scaleOf(frame);
    return nodes.flatMap((node, k) => {
      const parent = node.parentElement;
      const pcs = getComputedStyle(parent);
      if (pcs.display === 'none' || pcs.visibility !== 'visible') return [];
      const style = styleOf(parent, html, scale, el);
      const re = PRESERVE.test(pcs.whiteSpace) ? /[^\n]+/g : /\S+/g;
      return [...node.data.matchAll(re)].flatMap((mt) => {
        const range = document.createRange();
        range.setStart(node, mt.index);
        range.setEnd(node, mt.index + mt[0].length);
        const rects = [...range.getClientRects()].filter((r) => r.width > 0 || r.height > 0);
        if (rects.length === 0) return [];
        const g = offsets[k] + mt.index;
        return [{
          t: mt[0], style,
          box: unionBoxes(rects.map((r) => boxThrough(toLocal, rectBox(r)))),
          space: g > 0 && /\s/.test(full[g - 1]),
        }];
      });
    });
  };
  // Group words into rows: a word whose vertical centre lies inside the current row joins it.
  const groupRows = (words) => words.reduce((rows, w) => {
    const last = rows[rows.length - 1];
    const cy = w.box[1] + w.box[3] / 2;
    const joins = last && cy >= last.box[1] && cy <= last.box[1] + last.box[3];
    return joins
      ? [...rows.slice(0, -1), { box: unionBoxes([last.box, w.box]), words: [...last.words, w] }]
      : [...rows, { box: w.box, words: [w] }];
  }, []).sort((a, b) => a.box[1] - b.box[1]);
  // A space separates words with source whitespace between them, or a visible gap.
  const needsSpace = (prev, w) => w.space || w.box[0] - (prev.box[0] + prev.box[2]) > 0.25 * w.style.size;
  const mergeRuns = (words) => words.reduce((runs, w, k) => {
    const last = runs[runs.length - 1];
    const sep = k > 0 && needsSpace(words[k - 1], w) ? ' ' : '';
    if (last && sameStyle(last, w.style)) return [...runs.slice(0, -1), { ...last, t: last.t + sep + w.t }];
    const before = last && sep ? [...runs.slice(0, -1), { ...last, t: last.t + sep }] : runs;
    return [...before, { t: w.t, ...w.style }];
  }, []);
  const ALIGN = { left: 'start', start: 'start', center: 'middle', right: 'end', end: 'end', justify: 'start', '-webkit-center': 'middle' };
  const blockOf = (el, stop) =>
    (el === stop || !el.parentElement || !/^inline/.test(getComputedStyle(el).display) ? el : blockOf(el.parentElement, stop));
  const htmlAnchor = (fo) => {
    const first = textNodesIn(fo, new Set()).find((n) => /\S/.test(n.data));
    const block = first ? blockOf(first.parentElement, fo) : fo;
    return ALIGN[getComputedStyle(block).textAlign] || 'start';
  };
  const foLocalBox = (fo) => [fo.x.baseVal.value, fo.y.baseVal.value, fo.width.baseVal.value, fo.height.baseVal.value];
  const textPrim = (el, tag, ctx, html, frame) => {
    // client space -> root (rootInv) -> the element's own user space (inverse frame)
    const words = wordsOf(el, html, frame, mul(inv(frame), ctx.rootInv));
    if (words.length === 0) return [];
    // Client rects are axis-aligned, so mapping them back into the element's
    // frame recovers word boxes exactly only at multiples of 90 degrees. At
    // other angles: one row, boxed by the element's own local box.
    const rot = rotOf(frame);
    const local = rot % 90 === 0 ? groupRows(words)
      : [{ box: (html ? foLocalBox(el) : safeBBox(el)) || unionBoxes(words.map((w) => w.box)), words }];
    const anchor = getComputedStyle(el).textAnchor;
    return [{
      ...head('text', tag, el, ctx, placeBox(frame, unionBoxes(local.map((r) => r.box)))),
      html,
      anchor: html ? htmlAnchor(el) : (['start', 'middle', 'end'].includes(anchor) ? anchor : 'start'),
      rot,
      rows: local.map((r) => ({ bbox: roundBox(placeBox(frame, r.box)), runs: mergeRuns(r.words) })),
    }];
  };

  // -- foreignObject: background boxes, then the text --
  const htmlOpacity = (el, stop) => (el === stop || !el ? 1 : num(getComputedStyle(el).opacity, 1) * htmlOpacity(el.parentElement, stop));
  const htmlRect = (el, fo, ctx, scale) => {
    const cs = getComputedStyle(el);
    if (cs.display === 'none' || cs.visibility !== 'visible') return [];
    const op = ctx.op * htmlOpacity(el, fo);
    const fill = brighten(paintOf(cs.backgroundColor, op), ctx.br);
    if (fill.colour === null) return [];
    const r = el.getBoundingClientRect();
    if (!(r.width > 0 && r.height > 0)) return [];
    const box = boxThrough(ctx.rootInv, rectBox(r));
    const bw = num(cs.borderTopWidth, 0);
    const border = bw > 0 && cs.borderTopStyle !== 'none' && cs.borderTopStyle !== 'hidden' ? brighten(paintOf(cs.borderTopColor, op), ctx.br) : NO_PAINT;
    const rad = r3(num(cs.borderTopLeftRadius, 0) * scale);
    return [{
      ...head('rect', el.tagName.toLowerCase(), el, ctx, box),
      m: [1, 0, 0, 1, 0, 0],
      geom: { x: r3(box[0]), y: r3(box[1]), w: r3(box[2]), h: r3(box[3]), rx: rad, ry: rad },
      fill: fill.colour, fill_a: fill.a, fill_approx: false,
      stroke: border.colour, stroke_a: border.a, stroke_w: border.colour ? r3(bw * scale) : 0,
      dash: null, marker_start: null, marker_end: null,
      html: true,
    }];
  };
  const foreignPrims = (fo, ctx) => {
    const frame = matrixOf(ctx.rootInv, fo);
    if (!frame) return [];
    const boxes = [...fo.querySelectorAll('*')].flatMap((el) => htmlRect(el, fo, ctx, scaleOf(frame)));
    return [...boxes, ...textPrim(fo, 'foreignobject', ctx, true, frame)];
  };

  const imagePrim = (el, ctx) => {
    const bb = safeBBox(el), m = matrixOf(ctx.rootInv, el);
    if (!bb || !m || !(bb[2] > 0 && bb[3] > 0)) return [];
    const href = el.getAttribute('href') || el.getAttribute('xlink:href') || '';
    return [{ ...head('image', 'image', el, ctx, boxThrough(m, bb)), href, m: roundM(m) }];
  };
  const otherPrim = (el, tag, ctx) => {
    if (!(el instanceof SVGGraphicsElement)) return [];
    const bb = safeBBox(el), m = matrixOf(ctx.rootInv, el);
    if (!bb || !m || !(bb[2] > 0 || bb[3] > 0)) return [];
    return [head('other', tag, el, ctx, boxThrough(m, bb))];
  };

  // -- the walk --
  const SKIP = new Set(['defs', 'marker', 'clippath', 'mask', 'symbol', 'pattern', 'style', 'title', 'desc',
    'metadata', 'script', 'lineargradient', 'radialgradient', 'filter']);
  const CONTAINERS = new Set(['g', 'a', 'svg', 'switch']);
  const emit = (el, tag, cs, ctx) => {
    if (SHAPE_KIND[tag]) return shapePrim(el, tag, cs, ctx);
    if (tag === 'text') { const f = matrixOf(ctx.rootInv, el); return f ? textPrim(el, 'text', ctx, false, f) : []; }
    if (tag === 'foreignobject') return foreignPrims(el, ctx);
    if (tag === 'image') return imagePrim(el, ctx);
    return otherPrim(el, tag, ctx);
  };
  const walk = (el, ctx) => {
    const tag = el.tagName.toLowerCase();
    if (SKIP.has(tag)) return [];
    const cs = getComputedStyle(el);
    if (cs.display === 'none') return [];
    const op = ctx.op * num(cs.opacity, 1);
    if (op <= 0) return [];
    const br = ctx.br * brightnessOf(cs);
    if (CONTAINERS.has(tag)) {
      const inner = { ...ctx, op, br, anc: [...ctx.anc, ancEntry(el, ctx)] };
      return [...el.children].flatMap((ch) => walk(ch, inner));
    }
    return cs.visibility === 'visible' ? emit(el, tag, cs, { ...ctx, op, br }) : [];
  };

  // -- marker definitions --
  const MARKER_SHAPES = new Set(['rect', 'ellipse', 'line', 'poly', 'path']);
  const ownMatrix = (el) => {
    const t = el.transform && el.transform.baseVal && el.transform.baseVal.consolidate();
    return t ? fromDOM(t.matrix) : [1, 0, 0, 1, 0, 0];
  };
  // Matrix of an element in marker content coordinates: only transforms below the <marker>.
  const markerMatrixOf = (marker) => {
    const f = (el) => (el === marker ? [1, 0, 0, 1, 0, 0] : mul(f(el.parentElement), ownMatrix(el)));
    return f;
  };
  const markerOrient = (s) => (s === 'auto' || s === 'auto-start-reverse' ? s : num(s, 0));
  const markerPrim = (p) => ({
    kind: p.kind, tag: p.tag, cls: p.cls, geom: p.geom, m: p.m, bbox: p.bbox,
    fill: p.fill, fill_a: p.fill_a, fill_approx: p.fill_approx,
    stroke: p.stroke, stroke_a: p.stroke_a, stroke_w: p.stroke_w, dash: p.dash,
  });
  const markerDef = (marker, serial) => {
    const attr = (name) => marker.getAttribute(name);
    const vb = (attr('viewBox') || '').trim().split(/[\s,]+/).map(Number);
    const ctx = { rootInv: null, op: 1, br: 1, anc: [], serial, inMarker: true, mOf: markerMatrixOf(marker) };
    return {
      viewBox: vb.length === 4 && vb.every(Number.isFinite) ? vb.map(r3) : null,
      refX: r3(num(attr('refX'), 0)), refY: r3(num(attr('refY'), 0)),
      w: r3(num(attr('markerWidth'), 3)), h: r3(num(attr('markerHeight'), 3)),
      units: attr('markerUnits') === 'userSpaceOnUse' ? 'userSpaceOnUse' : 'strokeWidth',
      orient: markerOrient(attr('orient')),
      prims: [...marker.children].flatMap((ch) => walk(ch, ctx)).filter((p) => MARKER_SHAPES.has(p.kind)).map(markerPrim),
    };
  };
  const markersFor = (root, prims, serial) => {
    const used = new Set(prims.flatMap((p) => [p.marker_start, p.marker_end]).filter(Boolean));
    return Object.fromEntries([...root.querySelectorAll('marker')]
      .map((el) => [stripId(el.getAttribute('id')), el])
      .filter(([id]) => used.has(id))
      .map(([id, el]) => [id, markerDef(el, serial)]));
  };

  // -- render --
  const { mermaid, 'mermaid-zenuml': zenuml, elkLayouts } = globalThis;
  await Promise.all(Array.from(document.fonts, (font) => font.load()));
  await mermaid.registerExternalDiagrams([zenuml]);
  mermaid.registerLayoutLoaders(elkLayouts);
  mermaid.initialize({ startOnLoad: false, ...config });
  globalThis.__m2dReseed();
  const rendered = await (async () => {
    try {
      const { svg } = await mermaid.render(svgId, code, container);
      return { svg };
    } catch (e) {
      return { error: errMessage(e) };
    }
  })();
  if (rendered.error !== undefined) return { ok: false, error: { message: rendered.error } };
  container.innerHTML = rendered.svg;
  await document.fonts.ready;
  const root = container.querySelector('svg');
  if (!root) return { ok: false, error: { message: 'Mermaid produced no svg element' } };

  try {
    const vb = root.viewBox && root.viewBox.baseVal && root.viewBox.baseVal.width > 0
      ? [root.viewBox.baseVal.x, root.viewBox.baseVal.y, root.viewBox.baseVal.width, root.viewBox.baseVal.height]
      : [0, 0, root.width.baseVal.value, root.height.baseVal.value];
    const childText = (tag) => { const c = [...root.children].find((n) => n.tagName.toLowerCase() === tag); return c ? c.textContent : null; };
    const serial = new Map([...root.querySelectorAll('*')].map((el, k) => [el, k + 1]).concat([[root, 0]]));
    const rootInv = rootInvOf(root);
    const ctx = { rootInv, op: 1, br: 1, anc: [], serial, mOf: (el) => matrixOf(rootInv, el) };
    const prims = [...root.children].flatMap((ch) => walk(ch, ctx)).map((p, i) => ({ i, ...p }));
    return {
      ok: true,
      role: root.getAttribute('aria-roledescription'),
      viewBox: vb.map(r3),
      title: childText('title'),
      desc: childText('desc'),
      svg: rendered.svg,
      prims,
      markers: markersFor(root, prims, serial),
    };
  } catch (e) {
    return { ok: false, error: { message: `bridge walker: ${errMessage(e)}${e && e.stack ? `\n${e.stack}` : ''}` } };
  }
}

// Parse the diagram again for its type and its db. Run it in a freshly loaded
// mermaid bundle: some diagram dbs keep module-level id counters (classDiagram's
// "classId-Animal-0"), and a fresh bundle makes them count from where the
// render's own parse started, so db domIds match the ids in the svg.
async function pageDiagram({ code, config }) {
  const { mermaid, 'mermaid-zenuml': zenuml } = globalThis;
  const jsonSafe = (x, stack) => {
    if (x === null || typeof x === 'string' || typeof x === 'boolean') return x;
    if (typeof x === 'number') return Number.isFinite(x) ? x : null;
    if (typeof x !== 'object' || stack.includes(x) || stack.length > 12 || x instanceof Node) return undefined;
    const inner = [...stack, x];
    if (Array.isArray(x) || x instanceof Set) return [...x].map((v) => { const s = jsonSafe(v, inner); return s === undefined ? null : s; });
    const entries = x instanceof Map ? [...x.entries()] : Object.entries(x);
    return Object.fromEntries(entries.map(([k, v]) => [String(k), jsonSafe(v, inner)]).filter(([, v]) => v !== undefined));
  };
  const dbOf = (diagram) => {
    try {
      if (!diagram.db || typeof diagram.db.getData !== 'function') return null;
      const d = diagram.db.getData();
      if (!d || (!Array.isArray(d.nodes) && !Array.isArray(d.edges))) return null;
      return { nodes: jsonSafe(d.nodes || [], []) || [], edges: jsonSafe(d.edges || [], []) || [] };
    } catch (e) {
      return null;
    }
  };

  try {
    await mermaid.registerExternalDiagrams([zenuml]);
    mermaid.initialize({ startOnLoad: false, ...config });
    globalThis.__m2dReseed();
    const diagram = await mermaid.mermaidAPI.getDiagramFromText(code);
    return { type: diagram.type, db: dbOf(diagram) };
  } catch (e) {
    return { type: null, db: null };
  }
}

// ---- entry point ------------------------------------------------------------

main(process.argv.slice(2), process.env).catch((e) => {
  process.stderr.write(`mermaid_bridge: ${e.message}\n`);
  process.exitCode = e.exitCode || 1;
});

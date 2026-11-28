// Tests for bridge contract 3 (inst/node/mermaid_bridge.js): marker
// definitions (A1), DOM serials `n` (A2) and the brightness filter (A3).
// Two bridge launches in total: one for all sample diagrams, one for the
// timeline with its brightness filter switched off through themeCSS.
import { describe, it, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const bridge = path.resolve(here, '..', '..', 'inst', 'node', 'mermaid_bridge.js');

function defaultDepsDir() {
  const home = os.homedir();
  if (process.platform === 'darwin')
    return path.join(home, 'Library', 'Application Support', 'org.R-project.R', 'R', 'mermaid2drawml');
  if (process.platform === 'win32')
    return path.join(process.env.APPDATA || path.join(home, 'AppData', 'Roaming'), 'R', 'data', 'R', 'mermaid2drawml');
  return path.join(process.env.XDG_DATA_HOME || path.join(home, '.local', 'share'), 'R', 'mermaid2drawml');
}
const depsDir = process.env.M2D_NODE_DIR || defaultDepsDir();
const haveDeps = fs.existsSync(path.join(depsDir, 'node_modules', '@mermaid-js', 'mermaid-cli'));
const skipReason = haveDeps ? false
  : `mermaid-cli not found under ${depsDir}/node_modules; set M2D_NODE_DIR or install the dependencies`;

const childEnv = { ...process.env, M2D_NODE_DIR: depsDir };
const macChrome = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
if (!childEnv.PUPPETEER_EXECUTABLE_PATH && process.platform === 'darwin' && fs.existsSync(macChrome))
  childEnv.PUPPETEER_EXECUTABLE_PATH = macChrome;

function runBridge(args) {
  return spawnSync(process.execPath, [bridge, ...args], {
    env: childEnv, encoding: 'utf8', timeout: 120000, maxBuffer: 256 * 1024 * 1024,
  });
}
function parseOut(r) {
  assert.equal(r.status, 0, `bridge exited ${r.status}; stderr: ${(r.stderr || '').slice(-2000)}`);
  return JSON.parse(r.stdout);
}

// ---- diagram sources -------------------------------------------------------
const FLOW = `flowchart TB
  subgraph S1[One]
    A --> B
  end
  subgraph S2[Two]
    C --x D
    D --o E
  end
  B --> C
`;
const CLASS = `classDiagram
  A <|-- B
  C o-- D
  E *-- F
`;
const ER = `erDiagram
  A ||--o{ B : has
`;
const SEQ = `sequenceDiagram
  autonumber
  Alice->>Bob: Hi
  Bob-->>Alice: Yo
`;
const NOMARK = `pie title Pets
  "Dogs" : 5
  "Cats" : 3
`;
// Mermaid's stylesheet gives .eventWrapper { filter: brightness(120%) }.
const TIMELINE = `timeline
  title T
  2002 : LinkedIn
  2004 : Facebook : Google
`;
const NEUTRAL_CFG = { themeCSS: '.eventWrapper{filter:none !important;}' };

// ---- helpers ---------------------------------------------------------------
const isNum = (x) => typeof x === 'number' && Number.isFinite(x);
const isInt = Number.isInteger;
const isShape = (p) => ['rect', 'ellipse', 'line', 'poly', 'path'].includes(p.kind);
const usedMarkers = (res) => [...new Set(res.prims.flatMap((p) => [p.marker_start, p.marker_end]).filter(Boolean))].sort();
const hex = (c) => [1, 3, 5].map((k) => parseInt(c.slice(k, k + 2), 16));
const bright = (c, f) => hex(c).map((v) => Math.min(255, Math.round(v * f)));
const near1 = (a, b) => a.every((v, k) => Math.abs(v - b[k]) <= 1);
const unionBox = (boxes) => {
  const x0 = Math.min(...boxes.map((b) => b[0])), y0 = Math.min(...boxes.map((b) => b[1]));
  const x1 = Math.max(...boxes.map((b) => b[0] + b[2])), y1 = Math.max(...boxes.map((b) => b[1] + b[3]));
  return [x0, y0, x1 - x0, y1 - y0];
};
const isPaint = (v) => v === null || v === 'context' || /^#[0-9A-Fa-f]{6}$/.test(v);

describe('bridge contract 3', { skip: skipReason }, () => {
  let tmp, flow, cls, er, seq, nomark, tl, tlNeutral;
  const all = () => [flow, cls, er, seq, nomark, tl];

  before(() => {
    tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'm2d-bridge3-'));
    const write = (name, src) => {
      const f = path.join(tmp, name + '.mmd');
      fs.writeFileSync(f, src);
      return f;
    };
    const files = [
      write('flow', FLOW), write('cls', CLASS), write('er', ER),
      write('seq', SEQ), write('nomark', NOMARK), write('tl', TIMELINE),
    ];
    [flow, cls, er, seq, nomark, tl] = parseOut(runBridge(files));

    const cfg = path.join(tmp, 'neutral.json');
    fs.writeFileSync(cfg, JSON.stringify(NEUTRAL_CFG));
    [tlNeutral] = parseOut(runBridge(['--config', cfg, files[5]]));
  });
  after(() => { if (tmp) fs.rmSync(tmp, { recursive: true, force: true }); });

  describe('version', () => {
    it('reports bridge 3 on every successful result', () => {
      for (const r of [...all(), tlNeutral]) {
        assert.equal(r.ok, true);
        assert.equal(r.bridge, 3);
      }
    });
  });

  describe('A1 marker definitions', () => {
    it('has a markers object on every successful result', () => {
      for (const r of all()) {
        assert.equal(typeof r.markers, 'object', r.type);
        assert.ok(r.markers !== null && !Array.isArray(r.markers), r.type);
      }
    });
    it('keys markers by exactly the ids the prims refer to', () => {
      for (const r of all()) assert.deepEqual(Object.keys(r.markers).sort(), usedMarkers(r), r.type);
    });
    it('uses prefix-stripped ids for the keys', () => {
      for (const r of all())
        for (const k of Object.keys(r.markers)) assert.ok(!/^m2d[-_]/.test(k), k);
      assert.ok('flowchart-v2-pointEnd' in flow.markers);
    });
    it('gives a diagram without markers an empty object', () => {
      assert.equal(usedMarkers(nomark).length, 0);
      assert.deepEqual(nomark.markers, {});
    });
    it('lists the three flowchart arrow heads, the class ends, the ER ends and the sequence ones', () => {
      assert.deepEqual(Object.keys(flow.markers).sort(),
        ['flowchart-v2-circleEnd', 'flowchart-v2-crossEnd', 'flowchart-v2-pointEnd']);
      assert.deepEqual(Object.keys(cls.markers).sort(),
        ['class-aggregationStart', 'class-compositionStart', 'class-extensionStart']);
      assert.deepEqual(Object.keys(er.markers).sort(), ['er-onlyOneStart', 'er-zeroOrMoreEnd']);
      assert.ok('sequencenumber' in seq.markers);
      assert.ok('arrowhead' in seq.markers);
    });
    it('gives each marker the documented fields with the right types', () => {
      for (const r of all())
        for (const [id, m] of Object.entries(r.markers)) {
          const w = `${r.type}/${id}`;
          assert.ok(m.viewBox === null || (Array.isArray(m.viewBox) && m.viewBox.length === 4 && m.viewBox.every(isNum)), w);
          for (const f of ['refX', 'refY', 'w', 'h']) assert.ok(isNum(m[f]), `${w} ${f}`);
          assert.ok(m.w > 0 && m.h > 0, w);
          assert.ok(['strokeWidth', 'userSpaceOnUse'].includes(m.units), `${w} units ${m.units}`);
          assert.ok(m.orient === 'auto' || m.orient === 'auto-start-reverse' || isNum(m.orient), `${w} orient ${m.orient}`);
          assert.ok(Array.isArray(m.prims), w);
        }
    });
    it('reports the markerUnits and orient the svg declares', () => {
      assert.equal(flow.markers['flowchart-v2-pointEnd'].units, 'userSpaceOnUse');
      assert.equal(flow.markers['flowchart-v2-pointEnd'].orient, 'auto');
      assert.equal(cls.markers['class-extensionStart'].units, 'strokeWidth'); // attribute absent: default
      assert.equal(seq.markers.arrowhead.units, 'userSpaceOnUse');
      assert.equal(seq.markers.arrowhead.orient, 'auto-start-reverse');
    });
    it('reports viewBox, refX, refY, w and h from the marker element', () => {
      const p = flow.markers['flowchart-v2-pointEnd'];
      assert.deepEqual(p.viewBox, [0, 0, 10, 10]);
      assert.equal(p.refX, 5); assert.equal(p.refY, 5);
      assert.equal(p.w, 8); assert.equal(p.h, 8);
      const c = cls.markers['class-extensionStart'];
      assert.equal(c.viewBox, null); // no viewBox attribute in the svg
      assert.equal(c.refX, 18); assert.equal(c.refY, 7);
      assert.equal(c.w, 190); assert.equal(c.h, 240);
    });
    it('has drawn children for every known marker', () => {
      for (const r of all())
        for (const [id, m] of Object.entries(r.markers)) assert.ok(m.prims.length > 0, `${r.type}/${id}`);
    });
    it('gives marker prims the shape fields and none of the document-only ones', () => {
      for (const r of all())
        for (const [id, m] of Object.entries(r.markers))
          for (const p of m.prims) {
            const w = `${r.type}/${id}`;
            assert.ok(isShape(p), `${w} kind ${p.kind}`);
            assert.equal(typeof p.geom, 'object', w);
            assert.ok(Array.isArray(p.m) && p.m.length === 6 && p.m.every(isNum), w);
            assert.ok(Array.isArray(p.bbox) && p.bbox.length === 4 && p.bbox.every(isNum), w);
            assert.ok(isPaint(p.fill), `${w} fill ${p.fill}`);
            assert.ok(isPaint(p.stroke), `${w} stroke ${p.stroke}`);
            assert.ok(isNum(p.fill_a) && isNum(p.stroke_a) && isNum(p.stroke_w), w);
            for (const f of ['id', 'anc', 'i', 'marker_start', 'marker_end']) assert.ok(!(f in p), `${w} has ${f}`);
          }
    });
    it('draws the flowchart point marker inside its viewBox, with refX in range', () => {
      const m = flow.markers['flowchart-v2-pointEnd'];
      const [vx, vy, vw, vh] = m.viewBox;
      const [bx, by, bw, bh] = unionBox(m.prims.map((p) => p.bbox));
      assert.ok(bx >= vx - 1 && by >= vy - 1 && bx + bw <= vx + vw + 1 && by + bh <= vy + vh + 1,
        `bbox ${[bx, by, bw, bh]} vs viewBox ${m.viewBox}`);
      assert.ok(m.refX >= vx && m.refX <= vx + vw);
      assert.ok(m.refY >= vy && m.refY <= vy + vh);
    });
    it('puts marker geometry in marker content coordinates (path transform is identity here)', () => {
      const [p] = flow.markers['flowchart-v2-pointEnd'].prims;
      assert.deepEqual(p.m.map((v) => Math.round(v * 1000) / 1000), [1, 0, 0, 1, 0, 0]);
      assert.equal(p.kind, 'path');
    });
    it('fills the flowchart point marker', () => {
      const prims = flow.markers['flowchart-v2-pointEnd'].prims;
      assert.ok(prims.every((p) => p.fill !== null && p.fill_a > 0));
      assert.equal(prims[0].fill, '#333333'); // `.marker{fill:#333333}` in Mermaid's stylesheet
    });
    it('strokes the flowchart cross marker and the circle marker is an ellipse', () => {
      const cross = flow.markers['flowchart-v2-crossEnd'].prims;
      assert.ok(cross.every((p) => p.stroke !== null && p.stroke_w > 0));
      const circ = flow.markers['flowchart-v2-circleEnd'].prims;
      assert.ok(circ.every((p) => p.kind === 'ellipse'));
    });
    it('draws the sequence number marker as a circle', () => {
      const m = seq.markers.sequencenumber;
      assert.equal(m.prims.length, 1);
      assert.equal(m.prims[0].kind, 'ellipse');
      assert.equal(m.prims[0].geom.cx, 15);
      assert.equal(m.prims[0].geom.cy, 15);
      assert.equal(m.prims[0].geom.rx, 6);
      assert.ok(m.prims[0].fill !== null); // `#sequencenumber{fill:#333}`
    });
    it('reports the class diagram hollow markers as stroked with no visible fill', () => {
      // Stylesheet: `.extension, .aggregation { fill: transparent !important; stroke: #333333 !important }`
      for (const id of ['class-extensionStart', 'class-aggregationStart'])
        for (const p of cls.markers[id].prims) {
          assert.equal(p.stroke, '#333333', id);
          assert.ok(p.fill === null || p.fill_a === 0, `${id} fill ${p.fill} a=${p.fill_a}`);
        }
    });
    it('reports the class composition marker as filled and stroked', () => {
      // Stylesheet: `.composition { fill: #333333 !important; stroke: #333333 !important }`
      for (const p of cls.markers['class-compositionStart'].prims) {
        assert.equal(p.fill, '#333333');
        assert.equal(p.stroke, '#333333');
      }
    });
    it('reports ER markers as stroked, with unfilled paths and a white-filled circle', () => {
      // Stylesheet: `.marker { fill: none !important; stroke: #333333 !important }`; the circle has fill="white".
      for (const [id, m] of Object.entries(er.markers)) {
        for (const p of m.prims) {
          assert.equal(p.stroke, '#333333', id);
          if (p.kind === 'path') assert.ok(p.fill === null || p.fill_a === 0, `${id} path fill ${p.fill}`);
        }
      }
      const circles = er.markers['er-zeroOrMoreEnd'].prims.filter((p) => p.kind === 'ellipse');
      assert.equal(circles.length, 1);
      assert.equal(circles[0].fill, '#FFFFFF');
    });
  });

  describe('A2 DOM serials', () => {
    const entries = (r) => {
      const e = r.prims.flatMap((p) => p.anc || []);
      assert.ok(e.length > 0 && e.every((a) => isInt(a.n)), `${r.type}: anc entries lack integer n`);
      return e;
    };
    it('gives every prim and every anc entry an integer n', () => {
      for (const r of all()) {
        for (const p of r.prims) {
          assert.ok(isInt(p.n), `${r.type} prim ${p.i} n=${p.n}`);
          for (const a of p.anc || []) assert.ok(isInt(a.n), `${r.type} prim ${p.i} anc ${a.cls} n=${a.n}`);
        }
      }
    });
    it('keeps anc n strictly increasing outermost to innermost, all below the prim n', () => {
      for (const r of all())
        for (const p of r.prims) {
          const ns = (p.anc || []).map((a) => a.n);
          for (let k = 1; k < ns.length; k++) assert.ok(ns[k] > ns[k - 1], `${r.type} prim ${p.i} anc n ${ns}`);
          if (ns.length) assert.ok(ns[ns.length - 1] < p.n, `${r.type} prim ${p.i}: anc ${ns} vs own ${p.n}`);
        }
    });
    it('orders prims by n the way it orders them by i (document order)', () => {
      for (const r of all()) {
        // An HTML background box is emitted before its text prim but lies
        // inside the text's foreignObject, so only shapes are compared.
        const shapes = r.prims.filter((p) => isShape(p) && !p.html);
        for (let k = 1; k < shapes.length; k++)
          assert.ok(shapes[k].n > shapes[k - 1].n, `${r.type}: shapes ${k}`);
      }
    });
    it('uses one n per element: the same n always means the same id and class', () => {
      for (const r of all()) {
        const seen = new Map();
        for (const a of entries(r)) {
          const key = `${a.id}|${a.cls}`;
          if (seen.has(a.n)) assert.equal(seen.get(a.n), key, `${r.type} n=${a.n}`);
          else seen.set(a.n, key);
        }
      }
    });
    it('gives the same element the same n from every prim below it', () => {
      for (const r of all()) {
        const byId = new Map();
        for (const a of entries(r).filter((a) => a.id)) {
          if (byId.has(a.id)) assert.equal(byId.get(a.id), a.n, `${r.type} ${a.id}`);
          else byId.set(a.id, a.n);
        }
      }
    });
    it('shares the ancestor n between prims under the same innermost element', () => {
      // A node group holds a shape and its label: two prims, one parent.
      entries(flow);
      const parents = new Map();
      for (const p of flow.prims) {
        const a = (p.anc || []).at(-1);
        if (a) parents.set(a.n, (parents.get(a.n) || 0) + 1);
      }
      assert.ok([...parents.values()].some((c) => c >= 2), 'no ancestor is shared by two prims');
    });
    it('tells two sibling subgraph clusters with the same class apart', () => {
      const cluster = (id) => entries(flow).find((a) => a.id === id && a.cls === 'cluster');
      const s1 = cluster('S1'), s2 = cluster('S2');
      assert.ok(s1 && s2);
      assert.notEqual(s1.n, s2.n);
      // and a cluster's label group differs from its cluster
      const labels = entries(flow).filter((a) => a.cls === 'cluster-label');
      assert.equal(new Set(labels.map((a) => a.n)).size, 2);
    });
    it('gives sibling elements with identical id and class different n (class diagram nodes)', () => {
      const nodes = entries(cls).filter((a) => /\bnode\b/.test(a.cls));
      assert.ok(nodes.length > 1);
      const ids = new Set(nodes.map((a) => a.n));
      const keys = new Set(nodes.map((a) => `${a.id}|${a.cls}`));
      assert.equal(ids.size, keys.size);
    });
  });

  describe('A3 brightness filter', () => {
    const inWrapper = (p) => (p.anc || []).some((a) => /\beventWrapper\b/.test(a.cls || ''));
    it('uses a timeline whose stylesheet really has filter: brightness(120%) on .eventWrapper', () => {
      assert.match(tl.svg, /\.eventWrapper\{filter:brightness\(120%\);?\}/);
      assert.doesNotMatch(tl.svg, /filter:none/);
      assert.match(tlNeutral.svg, /\.eventWrapper\{filter:none ?!important;?\}/); // the later rule wins
      assert.equal(tl.prims.length, tlNeutral.prims.length);
    });
    it('multiplies fill and stroke under the filtered element by 1.2, clamped, within 1 per channel', () => {
      let checked = 0;
      tl.prims.forEach((p, k) => {
        const base = tlNeutral.prims[k];
        if (!isShape(p) || !inWrapper(p)) return;
        for (const f of ['fill', 'stroke']) {
          if (base[f] === null) { assert.equal(p[f], null); continue; }
          assert.ok(near1(hex(p[f]), bright(base[f], 1.2)),
            `prim ${k} ${f}: base ${base[f]} -> ${p[f]}, expected ${bright(base[f], 1.2)}`);
          checked++;
        }
      });
      assert.ok(checked >= 2, `only ${checked} colours checked`);
    });
    it('actually changes a colour that is not already saturated', () => {
      const changed = tl.prims.some((p, k) => isShape(p) && inWrapper(p) &&
        ((p.fill && p.fill !== tlNeutral.prims[k].fill) || (p.stroke && p.stroke !== tlNeutral.prims[k].stroke)));
      assert.ok(changed);
    });
    it('leaves alphas, fill_approx and stroke width unchanged', () => {
      tl.prims.forEach((p, k) => {
        if (!isShape(p)) return;
        const b = tlNeutral.prims[k];
        for (const f of ['fill_a', 'stroke_a', 'fill_approx', 'stroke_w']) assert.equal(p[f], b[f], `prim ${k} ${f}`);
      });
    });
    it('leaves shapes outside the filtered element exactly as without the filter', () => {
      let outside = 0;
      tl.prims.forEach((p, k) => {
        if (!isShape(p) || inWrapper(p)) return;
        assert.equal(p.fill, tlNeutral.prims[k].fill, `prim ${k}`);
        assert.equal(p.stroke, tlNeutral.prims[k].stroke, `prim ${k}`);
        outside++;
      });
      assert.ok(outside > 0);
    });
  });
});

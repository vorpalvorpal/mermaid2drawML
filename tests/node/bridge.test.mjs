// Tests for inst/node/mermaid_bridge.js (bridge contract v2).
// Each bridge launch starts headless Chrome, so all sample diagrams are
// rendered in one invocation and every test asserts on that shared result.
import { describe, it, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { skipReason, runBridge, parseOut } from './harness.mjs';

// ---- diagram sources -------------------------------------------------------
const FLOW = `flowchart TB
  subgraph S1[Group one]
    A[Start here] --> B{Decide<br/>now}
  end
  B -- yes --> C[(Database)]
  B -. no .-> D>Asym]
  E["\`A **bold** word\`"] ==> C
  F["2 * 3 * 4 __init__"] ~~~ A
  style A fill:lightblue,stroke:#f00,stroke-width:4px
  style D fill:none
  style E fill:#ff000080
`;
const BAD = 'this is definitely not a mermaid diagram !!!\n';
const SEQ = `sequenceDiagram
  participant Alice
  participant Bob
  Alice->>Bob: Hello Bob
  Note over Alice,Bob: A note
`;
const GANTT = `gantt
  title Plan
  dateFormat YYYY-MM-DD
  section Work
  Task one :a1, 2024-01-01, 5d
  Task two :after a1, 4d
`;
const PIE = `pie title Pets
  "Dogs" : 5
  "Cats" : 3
  "Rats" : 2
`;
const XY = `xychart-beta
  title "Sales"
  x-axis [a, b, c]
  y-axis "Revenue" 0 --> 100
  bar [10, 50, 80]
`;
const ACC = `flowchart LR
  accTitle: My title
  accDescr: My description
  X --> Y
`;
const CFG_FLOW = `flowchart TB
  A[Start here] --> B{Decide<br/>now}
`;

// ---- helpers ---------------------------------------------------------------
const rowText = (row) => row.runs.map((r) => r.t).join('');
const primText = (p) => p.rows.map(rowText).join(' ');
const isText = (p) => p.kind === 'text';
const isShape = (p) => ['rect', 'ellipse', 'line', 'poly', 'path'].includes(p.kind);
const hasAnc = (p, id) => p.id === id || (p.anc || []).some((a) => a.id === id);
const ancCls = (p, re) => (p.anc || []).some((a) => re.test(a.cls || ''));
const textPrim = (res, s) => res.prims.find((p) => isText(p) && primText(p) === s);
const nodeShapes = (res, domId) =>
  res.prims.filter((p) => isShape(p) && !p.html && hasAnc(p, domId));
const dbNode = (res, id) => res.db.nodes.find((n) => n.id === id);
const inside = (inner, outer, tol = 0.5) =>
  inner[0] >= outer[0] - tol && inner[1] >= outer[1] - tol &&
  inner[0] + inner[2] <= outer[0] + outer[2] + tol &&
  inner[1] + inner[3] <= outer[1] + outer[3] + tol;
const near = (a, b, tol = 0.5) => Math.abs(a - b) <= tol;
const isNum = (x) => typeof x === 'number' && Number.isFinite(x);

describe('mermaid bridge', { skip: skipReason }, () => {
  let tmp, results, flow, seq, gantt, pie, xy, acc, cfgRes, cfgFlow, cfgOut;
  let badArgsNoInputs, badArgsMissing;

  before(() => {
    tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'm2d-bridge-'));
    const files = {};
    const write = (name, src) => {
      files[name] = path.join(tmp, name + '.mmd');
      fs.writeFileSync(files[name], src);
    };
    write('flow', FLOW); write('bad', BAD); write('seq', SEQ); write('gantt', GANTT);
    write('pie', PIE); write('xy', XY); write('acc', ACC); write('cfgflow', CFG_FLOW);
    const order = ['flow', 'bad', 'seq', 'gantt', 'pie', 'xy', 'acc'];

    // Bad arguments need no browser.
    badArgsNoInputs = runBridge([]);
    badArgsMissing = runBridge([path.join(tmp, 'does-not-exist.mmd')]);

    // One launch for all the samples.
    const r = runBridge(order.map((k) => files[k]));
    results = { raw: r, stdout: r.stdout, status: r.status, stderr: r.stderr };
    const out = parseOut(r);
    results.out = out;
    [flow, , seq, gantt, pie, xy, acc] = out;
    results.badRes = out[1];

    // One launch with --config.
    const cfgPath = path.join(tmp, 'config.json');
    fs.writeFileSync(cfgPath, JSON.stringify({ theme: 'dark', htmlLabels: false }));
    cfgOut = parseOut(runBridge(['--config', cfgPath, files.cfgflow]));
    cfgFlow = cfgOut[0];
  });
  after(() => { if (tmp) fs.rmSync(tmp, { recursive: true, force: true }); });

  describe('invocation and result envelope', () => {
    it('prints one JSON array with one result per input, in order', () => {
      assert.ok(Array.isArray(results.out));
      assert.equal(results.out.length, 7);
      assert.equal(results.out[0].type, 'flowchart-v2');
      assert.equal(results.out[2].type, 'sequence');
      assert.equal(results.out[4].type, 'pie');
    });
    it('reports a syntax error as ok:false with a message while others succeed', () => {
      assert.equal(results.badRes.ok, false);
      assert.equal(typeof results.badRes.error.message, 'string');
      assert.ok(results.badRes.error.message.length > 0);
      for (const i of [0, 2, 3, 4, 5, 6]) assert.equal(results.out[i].ok, true, `input ${i}`);
    });
    it('exits 0 even when one diagram fails', () => {
      assert.equal(results.status, 0);
    });
    it('exits non-zero with stderr text when called with no inputs', () => {
      assert.notEqual(badArgsNoInputs.status, 0);
      assert.ok(badArgsNoInputs.stderr.trim().length > 0);
    });
    it('exits non-zero with stderr text when an input file is missing', () => {
      assert.notEqual(badArgsMissing.status, 0);
      assert.ok(badArgsMissing.stderr.trim().length > 0);
    });
  });

  describe('flowchart basics', () => {
    it('has type, role, bridge version, mermaid version and svg string', () => {
      assert.equal(flow.ok, true);
      assert.equal(flow.type, 'flowchart-v2');
      assert.equal(flow.role, 'flowchart-v2');
      assert.equal(flow.bridge, 4);
      assert.match(flow.mermaid, /^\d+\.\d+\.\d+/);
      assert.ok(flow.svg.startsWith('<svg'));
    });
    it('has a numeric four-element viewBox', () => {
      assert.equal(flow.viewBox.length, 4);
      assert.ok(flow.viewBox.every(isNum));
      assert.ok(flow.viewBox[2] > 0 && flow.viewBox[3] > 0);
    });
    it('returns accTitle and accDescr as title and desc, and null otherwise', () => {
      assert.equal(acc.title, 'My title');
      assert.equal(acc.desc, 'My description');
      assert.equal(flow.title, null);
      assert.equal(flow.desc, null);
    });
  });

  describe('db for a flowchart', () => {
    it('lists nodes with id, domId and shape', () => {
      const A = dbNode(flow, 'A'), B = dbNode(flow, 'B'), C = dbNode(flow, 'C');
      assert.equal(A.shape, 'squareRect');
      assert.equal(B.shape, 'diamond');
      assert.equal(C.shape, 'cylinder');
      assert.match(A.domId, /^flowchart-A-\d+$/);
    });
    it('records subgraph membership and isGroup', () => {
      assert.equal(dbNode(flow, 'A').parentId, 'S1');
      assert.equal(dbNode(flow, 'S1').isGroup, true);
    });
    it('lists edges with start, end and arrowTypeEnd', () => {
      const e = flow.db.edges.find((x) => x.start === 'A' && x.end === 'B');
      assert.ok(e, 'edge A->B');
      assert.equal(e.arrowTypeEnd, 'arrow_point');
    });
    it('is JSON-safe: no functions or config/other keys on nodes', () => {
      for (const n of flow.db.nodes) {
        assert.ok(!('config' in n) && !('other' in n));
      }
    });
  });

  describe('computed paint', () => {
    it('resolves a named CSS colour fill, hex stroke and stroke width from style', () => {
      const A = dbNode(flow, 'A');
      const s = nodeShapes(flow, A.domId).find((p) => p.fill === '#ADD8E6');
      assert.ok(s, 'shape prim of A with fill #ADD8E6');
      assert.equal(s.stroke, '#FF0000');
      assert.ok(near(s.stroke_w, 4));
    });
    it('returns fill null (not black) for fill:none with a visible stroke', () => {
      const D = dbNode(flow, 'D');
      const shapes = nodeShapes(flow, D.domId);
      assert.ok(shapes.length > 0);
      const s = shapes[0];
      assert.equal(s.fill, null);
      assert.notEqual(s.stroke, null);
    });
    it('carries fill alpha from an 8-digit hex colour', () => {
      const E = dbNode(flow, 'E');
      const s = nodeShapes(flow, E.domId).find((p) => p.fill === '#FF0000');
      assert.ok(s, 'E shape with red fill');
      assert.ok(near(s.fill_a, 0.5, 0.05), `fill_a ${s.fill_a}`);
    });
    it('uses the default theme fill for the subgraph cluster rect', () => {
      const c = flow.prims.find((p) => p.kind === 'rect' && !p.html && (p.cls.includes('cluster') || ancCls(p, /\bcluster\b/)));
      assert.ok(c, 'cluster rect');
      assert.equal(c.fill, '#FFFFDE');
    });
    it('reports the true cylinder height, not its straight side', () => {
      const C = dbNode(flow, 'C');
      const shapes = nodeShapes(flow, C.domId);
      assert.ok(shapes.length > 0);
      const h = Math.max(...shapes.map((p) => p.bbox[3]));
      assert.ok(h > 60, `cylinder bbox height ${h}`);
    });
    it('gives default node paint with a null fill_approx-style flag false', () => {
      const B = dbNode(flow, 'B');
      const s = nodeShapes(flow, B.domId)[0];
      assert.equal(s.fill, '#ECECFF');
      assert.equal(s.fill_approx, false);
    });
  });

  describe('edges', () => {
    const links = () => flow.prims.filter((p) => p.kind === 'path' && /flowchart-link/.test(p.cls));
    it('emits one path prim per visible edge with marker_end pointEnd', () => {
      const l = links();
      assert.equal(l.length, 4, 'A->B, B->C, B-.->D, E==>C; the ~~~ edge is invisible');
      for (const p of l) assert.equal(p.marker_end, 'flowchart-v2-pointEnd');
    });
    it('draws edges in #333333', () => {
      for (const p of links()) assert.equal(p.stroke, '#333333');
    });
    it('gives the dotted edge a dash array', () => {
      const p = links().find((x) => /edge-pattern-dotted/.test(x.cls));
      assert.ok(p, 'dotted edge');
      assert.ok(Array.isArray(p.dash) && p.dash.length > 0 && p.dash.every((d) => d > 0));
    });
    it('gives solid edges a null dash', () => {
      const p = links().find((x) => /edge-pattern-solid/.test(x.cls));
      assert.equal(p.dash, null);
    });
    it('gives the thick edge stroke_w 3.5', () => {
      const p = links().find((x) => /edge-thickness-thick/.test(x.cls));
      assert.ok(p, 'thick edge');
      assert.ok(near(p.stroke_w, 3.5));
    });
    it('emits no prim for the invisible ~~~ edge', () => {
      assert.ok(!flow.prims.some((p) => /invisible/.test(p.cls)));
    });
  });

  describe('text', () => {
    it('splits a <br/> label into two rows', () => {
      const t = flow.prims.find((p) => isText(p) && p.rows.length === 2 && rowText(p.rows[0]) === 'Decide');
      assert.ok(t, 'Decide / now text prim');
      assert.equal(rowText(t.rows[1]), 'now');
    });
    it('returns a one-line label as one row whose runs concatenate exactly', () => {
      const t = textPrim(flow, 'Start here');
      assert.ok(t);
      assert.equal(t.rows.length, 1);
      assert.equal(rowText(t.rows[0]), 'Start here');
    });
    it('marks markdown bold as a bold run and reads as plain text', () => {
      const t = flow.prims.find((p) => isText(p) && primText(p) === 'A bold word');
      assert.ok(t, 'markdown label');
      const bold = t.rows.flatMap((r) => r.runs).filter((r) => r.bold);
      assert.equal(bold.length, 1);
      assert.equal(bold[0].t.trim(), 'bold');
    });
    it('preserves literal asterisks and underscores character for character', () => {
      assert.ok(textPrim(flow, '2 * 3 * 4 __init__'));
    });
    it('puts colour, size and font on every run', () => {
      const t = textPrim(flow, 'Start here');
      for (const r of t.rows[0].runs) {
        assert.equal(r.color, '#333333');
        assert.ok(isNum(r.size));
        assert.ok(near(r.size, 16, 0.01));
        assert.equal(typeof r.family, 'string'); assert.ok(r.family.length > 0);
        assert.equal(typeof r.font, 'string'); assert.ok(r.font.length > 0);
        assert.equal(r.bold, false); assert.equal(r.italic, false); assert.equal(r.underline, false);
      }
    });
    it('keeps a node label inside its node shape bbox', () => {
      const t = textPrim(flow, 'Start here');
      const shape = nodeShapes(flow, dbNode(flow, 'A').domId)[0];
      assert.ok(inside(t.bbox, shape.bbox), `${t.bbox} in ${shape.bbox}`);
    });
    it('gives every text prim an anchor, rot 0 and rows with bbox', () => {
      for (const p of flow.prims.filter(isText)) {
        assert.ok(['start', 'middle', 'end'].includes(p.anchor));
        assert.equal(p.rot, 0);
        assert.equal(p.bbox.length, 4);
        for (const r of p.rows) assert.equal(r.bbox.length, 4);
      }
    });
    it('flowchart text is HTML by default', () => {
      assert.equal(textPrim(flow, 'Start here').html, true);
    });
  });

  describe('edge label background', () => {
    it('emits an HTML rect with a fill immediately before the label text, containing it', () => {
      const k = flow.prims.findIndex((p) => isText(p) && primText(p) === 'yes');
      assert.ok(k > 0, 'yes label');
      const t = flow.prims[k], r = flow.prims[k - 1];
      assert.equal(r.kind, 'rect');
      assert.equal(r.html, true);
      assert.notEqual(r.fill, null);
      assert.ok(inside(t.bbox, r.bbox), `${t.bbox} in ${r.bbox}`);
    });
  });

  describe('hygiene', () => {
    const all = () => [flow, seq, gantt, pie, xy, acc].flatMap((r) => r.prims);
    it('strips the m2d prefix from ids, ancestor ids and marker ids', () => {
      for (const p of all()) {
        assert.ok(!p.id.startsWith('m2d'), p.id);
        for (const a of p.anc || []) assert.ok(!a.id.startsWith('m2d'), a.id);
        if (p.marker_end) assert.ok(!p.marker_end.startsWith('m2d'), p.marker_end);
        if (p.marker_start) assert.ok(!p.marker_start.startsWith('m2d'), p.marker_start);
      }
    });
    it('emits nothing from defs or markers', () => {
      for (const p of all()) {
        assert.ok(!(p.anc || []).some((a) => /pointEnd|pointStart|crossEnd|circleEnd/.test(a.id)), 'marker ancestor');
        assert.ok(!/pointEnd/.test(p.id));
      }
    });
    it('emits no zero-size rect or ellipse', () => {
      for (const p of all().filter((q) => q.kind === 'rect' || q.kind === 'ellipse')) {
        assert.ok(p.bbox[2] > 0 && p.bbox[3] > 0, JSON.stringify(p.bbox));
      }
    });
    it('emits no invisible shape', () => {
      for (const p of all().filter(isShape)) assert.ok(p.fill !== null || p.stroke !== null);
    });
    it('numbers prims 0..n-1 in order', () => {
      for (const r of [flow, seq, gantt, pie, xy, acc]) r.prims.forEach((p, i) => assert.equal(p.i, i));
    });
    it('gives every shape prim a 6-number m, 4-number bbox and geom of the right form', () => {
      for (const p of all().filter(isShape)) {
        assert.equal(p.m.length, 6); assert.ok(p.m.every(isNum));
        assert.equal(p.bbox.length, 4); assert.ok(p.bbox.every(isNum));
        const g = p.geom;
        switch (p.kind) {
          case 'rect': for (const k of ['x', 'y', 'w', 'h']) assert.ok(isNum(g[k]), k); break;
          case 'ellipse': for (const k of ['cx', 'cy', 'rx', 'ry']) assert.ok(isNum(g[k]), k); break;
          case 'line': for (const k of ['x1', 'y1', 'x2', 'y2']) assert.ok(isNum(g[k]), k); break;
          case 'poly':
            assert.ok(Array.isArray(g.points) && g.points.every((q) => q.length === 2));
            assert.equal(typeof g.closed, 'boolean'); break;
          case 'path': assert.equal(typeof g.d, 'string'); assert.ok(g.d.length > 0); break;
        }
      }
    });
    it('gives every prim kind, tag, cls, data, anc and a 4-number bbox', () => {
      for (const p of all()) {
        assert.equal(typeof p.kind, 'string'); assert.equal(typeof p.tag, 'string');
        assert.equal(typeof p.cls, 'string'); assert.equal(typeof p.data, 'object');
        assert.ok(Array.isArray(p.anc)); assert.equal(p.bbox.length, 4);
      }
    });
  });

  describe('sequence diagram', () => {
    it('has type sequence and SVG (non-HTML) text', () => {
      assert.equal(seq.type, 'sequence');
      const texts = seq.prims.filter(isText);
      assert.ok(texts.length > 0);
      for (const t of texts) assert.equal(t.html, false);
      assert.ok(textPrim(seq, 'Hello Bob'));
    });
    it('takes the actor box fill from CSS', () => {
      const a = seq.prims.find((p) => p.kind === 'rect' && /\bactor\b/.test(p.cls));
      assert.ok(a, 'actor rect');
      assert.notEqual(a.fill, null);
    });
    it('takes the message line stroke from CSS and carries the arrowhead marker', () => {
      const l = seq.prims.find((p) => p.kind === 'line' && /messageLine0/.test(p.cls));
      assert.ok(l, 'message line');
      assert.notEqual(l.stroke, null);
      assert.equal(l.marker_end, 'arrowhead');
    });
  });

  describe('gantt chart', () => {
    it('has task rects with fills that exist only in CSS', () => {
      const tasks = gantt.prims.filter((p) => p.kind === 'rect' && /task/.test(p.cls));
      assert.ok(tasks.length >= 2);
      for (const t of tasks) assert.notEqual(t.fill, null);
    });
    it('has text prims including the task names', () => {
      assert.ok(gantt.prims.some((p) => isText(p) && /Task one/.test(primText(p))));
    });
  });

  describe('pie chart', () => {
    it('has a filled path per slice', () => {
      const slices = pie.prims.filter((p) => p.kind === 'path' && p.fill !== null);
      assert.ok(slices.length >= 3, `${slices.length} filled paths`);
    });
    it('has a title text', () => {
      assert.ok(textPrim(pie, 'Pets'));
    });
  });

  describe('rotated text', () => {
    it('reports a vertical axis title with rot -90 and an unrotated wide bbox', () => {
      const t = xy.prims.find((p) => isText(p) && /Revenue/.test(primText(p)));
      assert.ok(t, 'y-axis title');
      assert.equal(Math.abs(t.rot), 90);
      assert.ok(t.bbox[2] > t.bbox[3], `bbox ${t.bbox} should be wider than tall`);
    });
  });

  describe('--config', () => {
    it('applies htmlLabels:false so text prims are SVG text', () => {
      const texts = cfgFlow.prims.filter(isText);
      assert.ok(texts.length > 0);
      for (const t of texts) assert.equal(t.html, false);
    });
    it('still yields two rows for a <br/> label with htmlLabels false', () => {
      const t = cfgFlow.prims.find((p) => isText(p) && p.rows.length === 2 && rowText(p.rows[0]) === 'Decide');
      assert.ok(t, 'Decide / now');
      assert.equal(rowText(t.rows[1]), 'now');
    });
    it('applies the dark theme: node fill is no longer #ECECFF', () => {
      const s = nodeShapes(cfgFlow, dbNode(cfgFlow, 'A').domId).find((p) => p.fill !== null);
      assert.ok(s);
      assert.notEqual(s.fill, '#ECECFF');
    });
  });
});

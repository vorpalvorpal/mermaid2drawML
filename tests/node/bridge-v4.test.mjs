// Tests for bridge contract 4 (inst/node/mermaid_bridge.js): units resolved in
// the page (V1), text alpha (V2), gradients (V3), image data (V4), error kinds
// and --timeout (P2/V5), run independence (V6), opt-in svg (V7) and static /
// abstract class members (P11).
// Bridge launches: one for all sample diagrams, one each for flow alone and
// class alone (V6), one with --svg, one with --timeout 1, one with a themeCSS
// that paints a radial gradient, and one bad-argument call (no browser).
import { describe, it, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { skipReason, iconSkipReason, runBridge, parseOut } from './harness.mjs';

// ---- diagram sources -------------------------------------------------------
// Node ids that differ from their DOM ids: "B-1" is drawn as flowchart-B-1-1,
// "c_d" as flowchart-c_d-2. Two parallel edges A -> B-1, one labelled, and an
// edge to a subgraph.
const FLOW = `flowchart TB
  A[Alpha]
  B-1[Beta]
  c_d[Gamma]
  subgraph Outer
    subgraph Inner
      B-1
    end
    A
  end
  A -->|go| B-1
  A --> B-1
  c_d --> Outer
`;
const ELK = `---
config:
  layout: elk
---
${FLOW}`;
const CLASS = `classDiagram
  class A {
    +normal()
    +stat()$
    +abs()*
  }
  A <|-- B
  B --> C : uses
`;
const STATE = `stateDiagram-v2
  [*] --> One
  One --> Two : go
  Two --> [*]
  state Outer {
    [*] --> In
  }
  Two --> Outer
`;
const ER = `erDiagram
  CUSTOMER ||--o{ ORDER : places
  ORDER ||--|{ ITEM : contains
`;
const SEQ = `sequenceDiagram
  Alice->>Bob: Hello Bob
`;
const PIE = `pie title Pets
  "Dogs" : 5
  "Cats" : 3
`;
const GANTT = `gantt
  dateFormat YYYY-MM-DD
  section W
  T1 :a1, 2024-01-01, 20d
  T2 :after a1, 15d
`;
// Sankey links are <path stroke="url(#linearGradient-N)"> with userSpaceOnUse
// x1/x2 (y defaulting to 0), inside <g class="link"> under <g class="links"
// stroke-opacity="0.5">.
const SANKEY = `sankey-beta

Agricultural 'waste',Bio-conversion,124.729
Bio-conversion,Liquid,0.597
Bio-conversion,Solid,26.862
`;
// A valid 2x2 RGBA PNG.
const PNG_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4nGP4z8DwHwyBNBAw/AcAR8oI+ItOQ4UAAAAASUVORK5CYII=';
const IMG_NODE = `flowchart LR
  A@{ img: "data:image/png;base64,${PNG_B64}", label: "pic", pos: "b", w: 40, h: 40 }
`;
// Mermaid itself throws "The source image cannot be decoded." for an image
// NODE whose URL is broken; the diagram never reaches the bridge's walk.
const IMG_NODE_BROKEN = `flowchart LR
  B@{ img: "https://invalid.invalid/x.png", label: "bad", pos: "b", w: 40, h: 40 }
`;
// <img> in an HTML label: Mermaid waits for load or error and carries on, so a
// broken link still renders.
const IMG_HTML = `flowchart LR
  A["<img src='https://invalid.invalid/x.png' width='30' height='30'> hello"]
  B["<img src='data:image/png;base64,${PNG_B64}' width='30' height='30'> ok"]
`;
const ZEN = `zenuml
  Alice->Bob: Hi
`;
// Label icons (solid, brand, regular) and icon nodes. Mermaid 11.13 draws a
// label icon as <svg class="label-icon" viewBox=".."><path fill="currentColor">
// inside the label's <p>, before the text, once a pack is registered; an icon
// node's icon is <g style="color:..."><g><svg width="48" height="48"><path>.
const ICONS = `flowchart LR
  A[fa:fa-user Person]
  B[fab:fa-github Code]
  E[far:fa-star Star]
  C@{ icon: "fa:user", label: "User" }
  D@{ icon: "fab:github", label: "Gh", form: "square" }
`;
const BAD = 'this is definitely not a mermaid diagram !!!\n';
// Radial gradient: Mermaid draws none, so one is declared in an HTML label
// (DOMPurify keeps <svg><defs><radialGradient>) and referenced from themeCSS.
const RADIAL = `flowchart LR
  A[Alpha]
  B["<svg width='1' height='1'><defs><radialGradient id='rg1' cx='50%' cy='50%' r='50%'><stop offset='0%' stop-color='#ff0000'/><stop offset='100%' stop-color='#0000ff' stop-opacity='0.5'/></radialGradient></defs></svg> hi"]
  A --> B
`;
const RADIAL_CFG = { themeCSS: '.node.default rect.label-container{fill:url(#rg1) !important;}' };

// ---- helpers ---------------------------------------------------------------
const isNum = (x) => typeof x === 'number' && Number.isFinite(x);
const near = (a, b, tol) => Math.abs(a - b) <= tol;
const HEX = /^#[0-9A-Fa-f]{6}$/;
const rowText = (row) => row.runs.map((r) => r.t).join('');
const primText = (p) => (p.rows || []).map(rowText).join(' ');
const isText = (p) => p.kind === 'text';
const runs = (res) => res.prims.filter(isText).flatMap((p) => p.rows.flatMap((r) => r.runs));
const textPrim = (res, s) => res.prims.find((p) => isText(p) && primText(p) === s);
// Every unit attached to a prim: on the prim itself or on any ancestor entry.
const unitsOf = (p) => [p.unit, ...(p.anc || []).map((a) => a.unit)].filter(Boolean);
const unitOf = (p, kind) => unitsOf(p).find((u) => u.kind === kind);
const unitsOfKind = (res, kind) => res.prims.map((p) => unitOf(p, kind)).filter(Boolean);
const ids = (us) => [...new Set(us.map((u) => u.id))].sort();
const dbNodes = (res) => res.db.nodes.filter((n) => !n.isGroup);
const dbGroups = (res) => res.db.nodes.filter((n) => n.isGroup);
const sortedIds = (xs) => xs.map((x) => x.id).sort();
const hasToken = (p, re) => (p.anc || []).some((a) => re.test(a.cls || ''));
const pngSize = (dataUri) => {
  const buf = Buffer.from(dataUri.slice(dataUri.indexOf(',') + 1), 'base64');
  return { sig: [...buf.subarray(0, 8)], w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
};
const PNG_SIG = [137, 80, 78, 71, 13, 10, 26, 10];

describe('bridge contract 4', { skip: skipReason }, () => {
  let tmp, flow, elk, cls, state, er, seq, pie, gantt, sankey;
  let imgNode, imgNodeBroken, imgHtml, zen, bad, icons;
  let flowAlone, clsAlone, clsSvg, svgRun, timedOut, timedOutRun, radial, badFlag;
  const unitDiagrams = () => [['flowchart', flow], ['elk', elk], ['class', cls], ['state', state], ['er', er]];

  before(() => {
    tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'm2d-bridge4-'));
    const write = (name, src) => {
      const f = path.join(tmp, name + '.mmd');
      fs.writeFileSync(f, src);
      return f;
    };
    const f = {
      flow: write('flow', FLOW), elk: write('elk', ELK), cls: write('cls', CLASS),
      state: write('state', STATE), er: write('er', ER), seq: write('seq', SEQ),
      pie: write('pie', PIE), gantt: write('gantt', GANTT), sankey: write('sankey', SANKEY),
      imgNode: write('imgnode', IMG_NODE), imgBroken: write('imgbroken', IMG_NODE_BROKEN),
      imgHtml: write('imghtml', IMG_HTML), zen: write('zen', ZEN), bad: write('bad', BAD),
      radial: write('radial', RADIAL), icons: write('icons', ICONS),
    };

    // One launch for every sample diagram; flow first, class late (V6).
    [flow, elk, state, er, seq, pie, gantt, sankey, imgNode, imgNodeBroken, imgHtml, zen, bad, icons, cls] =
      parseOut(runBridge([f.flow, f.elk, f.state, f.er, f.seq, f.pie, f.gantt, f.sankey,
        f.imgNode, f.imgBroken, f.imgHtml, f.zen, f.bad, f.icons, f.cls]));

    [flowAlone] = parseOut(runBridge([f.flow]));
    [clsAlone] = parseOut(runBridge([f.cls]));
    // These two use flags the version 3 bridge rejects: keep the raw run so a
    // failure is reported by the tests that need it, not by this hook.
    svgRun = runBridge(['--svg', f.cls]);
    if (svgRun.status === 0) [clsSvg] = JSON.parse(svgRun.stdout);
    timedOutRun = runBridge(['--timeout', '1', f.flow]);
    if (timedOutRun.status === 0) [timedOut] = JSON.parse(timedOutRun.stdout);
    const cfg = path.join(tmp, 'radial.json');
    fs.writeFileSync(cfg, JSON.stringify(RADIAL_CFG));
    [radial] = parseOut(runBridge(['--config', cfg, f.radial]));
    badFlag = runBridge(['--no-such-flag', f.flow]);
  });
  after(() => { if (tmp) fs.rmSync(tmp, { recursive: true, force: true }); });

  describe('version and compatibility', () => {
    it('reports bridge 4 on every successful result', () => {
      for (const r of [flow, elk, cls, state, er, seq, pie, gantt, sankey, imgNode, imgHtml, zen, radial]) {
        assert.equal(r.ok, true, r.type);
        assert.equal(r.bridge, 4);
      }
    });
    it('keeps the version 3 fields the R side reads', () => {
      for (const r of [flow, cls, gantt]) {
        assert.equal(typeof r.markers, 'object');
        assert.ok(Array.isArray(r.viewBox) && r.viewBox.length === 4);
        for (const p of r.prims) {
          assert.ok(Number.isInteger(p.i) && Number.isInteger(p.n), 'i and n');
          assert.equal(typeof p.kind, 'string');
          assert.ok(Array.isArray(p.bbox) && p.bbox.length === 4 && p.bbox.every(isNum));
          for (const a of p.anc) {
            assert.ok(Number.isInteger(a.n));
            assert.equal(typeof a.id, 'string');
            assert.equal(typeof a.cls, 'string');
          }
        }
      }
      const shape = flow.prims.find((p) => p.kind === 'rect');
      assert.ok('fill' in shape && 'fill_a' in shape && 'stroke' in shape && 'm' in shape && 'geom' in shape);
      assert.ok(Array.isArray(flow.db.nodes) && Array.isArray(flow.db.edges));
    });
  });

  describe('V1 units (flowchart, dagre)', () => {
    it('gives every node prim an ancestor unit with the id written in the diagram', () => {
      for (const [id, label] of [['A', 'Alpha'], ['B-1', 'Beta'], ['c_d', 'Gamma']]) {
        const t = textPrim(flow, label);
        assert.ok(t, label);
        const u = unitOf(t, 'node');
        assert.ok(u, `${label} has a node unit`);
        assert.equal(u.id, id);
        const own = flow.prims.filter((p) => unitOf(p, 'node')?.id === id);
        assert.ok(own.length >= 2, `${id}: shape and text both carry the unit`);
        assert.ok(own.some((p) => p.kind === 'rect'));
      }
    });
    it('puts a node unit on every prim inside a DOM node group', () => {
      for (const [name, r] of unitDiagrams()) {
        const inNode = r.prims.filter((p) => hasToken(p, /\bnode\b/));
        assert.ok(inNode.length > 0, name);
        for (const p of inNode) assert.ok(unitOf(p, 'node'), `${name}: prim ${p.i} in a .node has no node unit`);
      }
    });
    it('node unit ids are exactly the non-group db node ids, never DOM ids', () => {
      assert.deepEqual(ids(unitsOfKind(flow, 'node')), ['A', 'B-1', 'c_d']);
      assert.deepEqual(ids(unitsOfKind(flow, 'node')), sortedIds(dbNodes(flow)));
      assert.ok(unitsOfKind(flow, 'node').every((u) => !/^flowchart-/.test(u.id)));
    });
    it('gives subgraph backgrounds and titles a subgraph unit with the subgraph id', () => {
      assert.deepEqual(ids(unitsOfKind(flow, 'subgraph')), ['Inner', 'Outer']);
      assert.deepEqual(ids(unitsOfKind(flow, 'subgraph')), sortedIds(dbGroups(flow)));
      for (const title of ['Outer', 'Inner']) {
        assert.equal(unitOf(textPrim(flow, title), 'subgraph')?.id, title);
      }
      for (const id of ['Outer', 'Inner']) {
        const bg = flow.prims.find((p) => p.kind === 'rect' && unitOf(p, 'subgraph')?.id === id);
        assert.ok(bg, `${id} background rect`);
        assert.equal(unitOf(bg, 'node'), undefined, `${id} is not also a node`);
      }
    });
    it('does not report a node as a subgraph or the reverse', () => {
      for (const p of flow.prims) {
        const k = unitsOf(p).map((u) => u.kind);
        assert.ok(!(k.includes('node') && k.includes('subgraph')) || p.anc.length > 6, `prim ${p.i}`);
      }
      assert.equal(unitOf(textPrim(flow, 'Beta'), 'subgraph'), undefined);
    });
    it('gives each edge path a unit with the db edge id, start and end', () => {
      const edgeUnits = unitsOfKind(flow, 'edge');
      assert.deepEqual(ids(edgeUnits), sortedIds(flow.db.edges));
      for (const e of flow.db.edges) {
        const p = flow.prims.find((q) => q.kind === 'path' && unitOf(q, 'edge')?.id === e.id);
        assert.ok(p, `path prim for edge ${e.id}`);
        const u = unitOf(p, 'edge');
        assert.equal(u.start, e.start);
        assert.equal(u.end, e.end);
      }
    });
    it('resolves an edge to a subgraph and gives parallel edges different ids', () => {
      const toSub = unitsOfKind(flow, 'edge').find((u) => u.start === 'c_d');
      assert.equal(toSub.end, 'Outer');
      const parallel = unitsOfKind(flow, 'edge').filter((u) => u.start === 'A' && u.end === 'B-1');
      assert.equal(ids(parallel).length, 2);
    });
    it('gives the label "go" an edge_label unit matching its edge', () => {
      const label = textPrim(flow, 'go');
      const u = unitOf(label, 'edge_label');
      assert.ok(u, 'edge_label unit');
      const e = flow.db.edges.find((x) => x.label === 'go');
      assert.deepEqual([u.id, u.start, u.end], [e.id, e.start, e.end]);
      const edge = unitsOfKind(flow, 'edge').find((x) => x.id === e.id);
      assert.deepEqual([edge.start, edge.end], [u.start, u.end]);
      // the label's background boxes belong to the same label unit
      const sameLabel = label.anc.find((a) => a.unit?.kind === 'edge_label');
      const boxes = flow.prims.filter((p) => p.kind === 'rect' && p.anc.some((a) => a.n === sameLabel.n));
      assert.ok(boxes.length > 0);
      for (const b of boxes) assert.equal(unitOf(b, 'edge_label')?.id, e.id);
    });
  });

  describe('V1 units (ELK layout)', () => {
    it('resolves nodes, subgraphs, edges and the label against db ids', () => {
      assert.equal(elk.ok, true);
      assert.deepEqual(ids(unitsOfKind(elk, 'node')), ['A', 'B-1', 'c_d']);
      assert.deepEqual(ids(unitsOfKind(elk, 'subgraph')), ['Inner', 'Outer']);
      // ELK draws edge paths with ids like "L_A_B-1_0_0" (db id + "_0"):
      // the unit id is still the db id.
      assert.deepEqual(ids(unitsOfKind(elk, 'edge')), sortedIds(elk.db.edges));
      for (const e of elk.db.edges) {
        const u = unitsOfKind(elk, 'edge').find((x) => x.id === e.id);
        assert.deepEqual([u.start, u.end], [e.start, e.end]);
      }
      const u = unitOf(textPrim(elk, 'go'), 'edge_label');
      const e = elk.db.edges.find((x) => x.label === 'go');
      assert.deepEqual([u.id, u.start, u.end], [e.id, e.start, e.end]);
      assert.equal(unitOf(textPrim(elk, 'Beta'), 'node').id, 'B-1');
    });
  });

  describe('V1 units (class, state and ER diagrams)', () => {
    it('class: node units for each class, edge units for relations, label unit', () => {
      assert.deepEqual(ids(unitsOfKind(cls, 'node')), ['A', 'B', 'C']);
      assert.deepEqual(ids(unitsOfKind(cls, 'edge')), sortedIds(cls.db.edges));
      assert.equal(unitOf(textPrim(cls, '+stat()'), 'node').id, 'A');
      const e = cls.db.edges.find((x) => x.label === 'uses');
      const u = unitOf(textPrim(cls, 'uses'), 'edge_label');
      assert.deepEqual([u.id, u.start, u.end], [e.id, e.start, e.end]);
    });
    it('state: node units for each state, edge units for transitions, label unit', () => {
      assert.deepEqual(ids(unitsOfKind(state, 'node')), sortedIds(dbNodes(state)));
      assert.ok(ids(unitsOfKind(state, 'node')).includes('One'));
      assert.ok(ids(unitsOfKind(state, 'node')).includes('Two'));
      assert.deepEqual(ids(unitsOfKind(state, 'edge')), sortedIds(state.db.edges));
      assert.equal(unitOf(textPrim(state, 'One'), 'node').id, 'One');
      const e = state.db.edges.find((x) => x.label === 'go');
      assert.equal(unitOf(textPrim(state, 'go'), 'edge_label').id, e.id);
    });
    it('state: a composite state is a subgraph unit', () => {
      // A composite state's element carries the state's id, not its domId.
      assert.deepEqual(ids(unitsOfKind(state, 'subgraph')), sortedIds(dbGroups(state)));
      assert.ok(ids(unitsOfKind(state, 'subgraph')).includes('Outer'));
      // Mermaid draws the inner states beside the composite's element, not
      // inside it: which composite holds a state is the db's parentId.
      assert.equal(unitOf(textPrim(state, 'In'), 'node').id, 'In');
      assert.equal(state.db.nodes.find((n) => n.id === 'In').parentId, 'Outer');
      assert.equal(unitOf(textPrim(state, 'Outer'), 'subgraph').id, 'Outer');
    });
    it('ER: node units for each entity, edge units for relationships, label unit', () => {
      assert.deepEqual(ids(unitsOfKind(er, 'node')), sortedIds(dbNodes(er)));
      assert.equal(unitOf(textPrim(er, 'ORDER'), 'node').id, er.db.nodes.find((n) => /ORDER/.test(n.id)).id);
      assert.deepEqual(ids(unitsOfKind(er, 'edge')), sortedIds(er.db.edges));
      const e = er.db.edges.find((x) => x.label === 'places');
      const u = unitOf(textPrim(er, 'places'), 'edge_label');
      assert.deepEqual([u.id, u.start, u.end], [e.id, e.start, e.end]);
    });
  });

  describe('V1 units (diagram types without nodes and edges)', () => {
    it('sequence, pie, gantt and sankey carry no unit anywhere and still succeed', () => {
      for (const r of [seq, pie, gantt, sankey]) {
        assert.equal(r.ok, true, r.type);
        assert.ok(r.prims.length > 0, r.type);
        for (const p of r.prims) {
          assert.equal(p.unit, undefined, `${r.type} prim ${p.i}`);
          for (const a of p.anc) assert.equal(a.unit, undefined, `${r.type} anc ${a.n}`);
        }
      }
    });
  });

  describe('V2 text alpha', () => {
    it('gives every text run a numeric alpha between 0 and 1', () => {
      for (const r of [flow, elk, cls, state, er, seq, pie, gantt, sankey, imgHtml, zen]) {
        const rs = runs(r);
        assert.ok(rs.length > 0, r.type);
        for (const run of rs) assert.ok(isNum(run.alpha) && run.alpha >= 0 && run.alpha <= 1, `${r.type}: ${JSON.stringify(run.alpha)}`);
      }
    });
    it('gives ordinary flowchart labels alpha 1', () => {
      for (const label of ['Alpha', 'Beta', 'go']) assert.equal(textPrim(flow, label).rows[0].runs[0].alpha, 1);
    });
    it('gives gantt axis tick labels alpha 0.8 (.grid .tick { opacity: 0.8 })', () => {
      const ticks = gantt.prims.filter((p) => isText(p) && /^2024-\d\d-\d\d$/.test(primText(p)));
      assert.ok(ticks.length >= 5);
      for (const t of ticks) for (const run of t.rows[0].runs) assert.ok(near(run.alpha, 0.8, 0.02), `alpha ${run.alpha}`);
      const title = textPrim(gantt, 'W');
      assert.equal(title.rows[0].runs[0].alpha, 1);
    });
  });

  describe('V3 gradients', () => {
    const links = () => sankey.prims.filter((p) => p.kind === 'path' && p.stroke_w > 0);
    it('reports a linear stroke_grad on every sankey link', () => {
      assert.equal(links().length, 3);
      for (const p of links()) {
        assert.equal(p.stroke_grad?.type, 'linear');
        assert.ok(Array.isArray(p.stroke_grad.p1) && p.stroke_grad.p1.length === 2 && p.stroke_grad.p1.every(isNum));
        assert.ok(Array.isArray(p.stroke_grad.p2) && p.stroke_grad.p2.length === 2 && p.stroke_grad.p2.every(isNum));
      }
    });
    it('places p1 and p2 at the left and right ends of the link (root user space)', () => {
      for (const p of links()) {
        const [x, , w] = p.bbox;
        assert.ok(near(p.stroke_grad.p1[0], x, 2), `p1.x ${p.stroke_grad.p1[0]} vs ${x}`);
        assert.ok(near(p.stroke_grad.p2[0], x + w, 2), `p2.x ${p.stroke_grad.p2[0]} vs ${x + w}`);
      }
    });
    it('reports at least two stops with ascending offsets, hex colours and alphas', () => {
      for (const p of links()) {
        const st = p.stroke_grad.stops;
        assert.ok(st.length >= 2);
        st.forEach((s, k) => {
          assert.ok(isNum(s.o) && s.o >= 0 && s.o <= 1);
          assert.match(s.c, HEX);
          assert.ok(isNum(s.a) && s.a >= 0 && s.a <= 1);
          if (k > 0) assert.ok(s.o >= st[k - 1].o);
        });
        assert.equal(st[0].o, 0);
        assert.equal(st.at(-1).o, 1);
      }
    });
    it('keeps the first stop colour as stroke, and gives the three links their own colours', () => {
      for (const p of links()) assert.equal(p.stroke.toUpperCase(), p.stroke_grad.stops[0].c.toUpperCase());
      assert.equal(links()[0].stroke_grad.stops.at(-1).c.toUpperCase(), '#F28E2C');
      assert.equal(links()[0].fill, null);
    });
    it('leaves prims without gradient paint without fill_grad / stroke_grad', () => {
      for (const p of sankey.prims.filter((q) => q.kind === 'rect')) {
        assert.ok(p.fill_grad == null && p.stroke_grad == null);
      }
      for (const r of [flow, cls, gantt]) {
        for (const p of r.prims) assert.ok(p.fill_grad == null && p.stroke_grad == null, `${r.type} prim ${p.i}`);
      }
      for (const p of links()) assert.ok(p.fill_grad == null);
    });
    it('reports a radial fill_grad (objectBoundingBox units resolved) with fill_approx kept', () => {
      const boxes = radial.prims.filter((p) => p.kind === 'rect' && p.fill_grad);
      assert.ok(boxes.length >= 1, 'a rect with a radial fill_grad');
      for (const p of boxes) {
        const g = p.fill_grad;
        assert.equal(g.type, 'radial');
        assert.equal(p.fill_approx, true);
        assert.equal(p.fill.toUpperCase(), '#FF0000');
        const [x, y, w, h] = p.bbox;
        assert.ok(Array.isArray(g.c) && g.c.length === 2 && g.c.every(isNum));
        assert.ok(near(g.c[0], x + w / 2, 1) && near(g.c[1], y + h / 2, 1), `centre ${g.c} in ${p.bbox}`);
        assert.ok(isNum(g.r) && g.r >= Math.min(w, h) / 2 - 1 && g.r <= Math.max(w, h) / 2 + 1, `r ${g.r}`);
        assert.equal(g.stops.length, 2);
        assert.deepEqual(g.stops.map((s) => s.c.toUpperCase()), ['#FF0000', '#0000FF']);
        assert.deepEqual(g.stops.map((s) => s.o), [0, 1]);
        assert.equal(g.stops[0].a, 1);
        assert.ok(near(g.stops[1].a, 0.5, 0.01));
        assert.ok(p.stroke_grad == null);
      }
    });
  });

  describe('V4 images', () => {
    it('reports an image node as an image prim with PNG data, bbox and m', () => {
      const img = imgNode.prims.find((p) => p.kind === 'image');
      assert.ok(img, 'image prim');
      // The picture is `png` (ruling B1): every prim already has `data` for
      // its data-* attributes.
      assert.equal(typeof img.png, 'string', `image prim png is ${JSON.stringify(img.png)}`);
      assert.match(img.png, /^data:image\/png;base64,/);
      const { sig, w, h } = pngSize(img.png);
      assert.deepEqual(sig, PNG_SIG);
      assert.deepEqual([w, h], [2, 2], 'drawn at natural size');
      assert.ok(img.bbox.length === 4 && img.bbox.every(isNum));
      assert.ok(img.m.length === 6 && img.m.every(isNum));
      assert.equal(typeof img.href, 'string');
    });
    it('reports <img> in an HTML label as image prims with html: true', () => {
      const imgs = imgHtml.prims.filter((p) => p.kind === 'image');
      assert.equal(imgs.length, 2);
      for (const p of imgs) {
        assert.equal(p.html, true);
        assert.ok(p.bbox.length === 4 && p.bbox.every(isNum));
        assert.ok(p.m.length === 6 && p.m.every(isNum));
      }
    });
    it('gives png: null for a broken <img> link without failing the diagram', () => {
      assert.equal(imgHtml.ok, true);
      const imgs = imgHtml.prims.filter((p) => p.kind === 'image');
      assert.equal(imgs.filter((p) => p.png === null).length, 1, JSON.stringify(imgs.map((p) => p.png)));
      const good = imgs.filter((p) => typeof p.png === 'string');
      assert.equal(good.length, 1);
      assert.match(good[0].png, /^data:image\/png;base64,/);
      assert.deepEqual(pngSize(good[0].png).sig, PNG_SIG);
    });
    it('guard: an image node with a broken URL fails inside Mermaid, as a result not a crash', () => {
      assert.equal(imgNodeBroken.ok, false);
      assert.ok(['syntax', 'bridge', 'timeout'].includes(imgNodeBroken.error.kind), String(imgNodeBroken.error.kind));
      assert.match(imgNodeBroken.error.message, /cannot be decoded/);
    });
  });

  describe('P2 / V5 error kinds and --timeout', () => {
    it('reports a syntax error as ok:false with error.kind "syntax"', () => {
      assert.equal(bad.ok, false);
      assert.equal(bad.error.kind, 'syntax');
      assert.ok(bad.error.message.length > 0);
    });
    it('reports --timeout 1 as error.kind "timeout" and still exits 0 with JSON', () => {
      assert.equal(timedOutRun.status, 0, `--timeout: exit ${timedOutRun.status}; stderr: ${(timedOutRun.stderr || '').slice(-300)}`);
      assert.equal(timedOut.ok, false);
      assert.equal(timedOut.error.kind, 'timeout');
      assert.equal(typeof timedOut.error.message, 'string');
    });
    it('guard: an unknown flag still exits 2 with a message on stderr', () => {
      assert.equal(badFlag.status, 2);
      assert.ok(badFlag.stderr.trim().length > 0);
    });
  });

  describe('V6 run independence', () => {
    it('renders a flowchart identically alone and in a long run (ids and positions)', () => {
      assert.deepEqual(flow, flowAlone);
    });
    it('renders a class diagram identically alone and late in a long run', () => {
      assert.deepEqual(cls, clsAlone);
    });
    it('still renders a zenuml diagram', () => {
      assert.equal(zen.ok, true, zen.ok ? '' : zen.error.message);
      assert.ok(zen.prims.length > 0);
    });
  });

  describe('V7 svg is opt-in', () => {
    it('has no svg field by default, on successes or failures', () => {
      for (const r of [flow, cls, gantt, bad, imgNodeBroken]) assert.ok(!('svg' in r), r.ok ? r.type : 'failure');
    });
    it('returns svg as a string starting with <svg under --svg', () => {
      assert.equal(svgRun.status, 0, `--svg: exit ${svgRun.status}; stderr: ${(svgRun.stderr || '').slice(-300)}`);
      assert.equal(typeof clsSvg.svg, 'string');
      assert.ok(clsSvg.svg.startsWith('<svg'));
      assert.equal(clsSvg.bridge, 4);
    });
  });

  describe('P11 static and abstract class members', () => {
    const run = (res, text) => res.prims.filter(isText).flatMap((p) => p.rows.flatMap((r) => r.runs)).find((r) => r.t === text);
    it('marks the static method +stat() underline', () => {
      assert.equal(run(cls, '+stat()').underline, true);
      assert.equal(run(cls, '+stat()').italic, false);
    });
    it('marks the abstract method +abs() italic', () => {
      assert.equal(run(cls, '+abs()').italic, true);
      assert.equal(run(cls, '+abs()').underline, false);
    });
    it('leaves +normal() plain', () => {
      assert.equal(run(cls, '+normal()').underline, false);
      assert.equal(run(cls, '+normal()').italic, false);
    });
    it('guard: Mermaid marks them with inline styles on the label group', () => {
      assert.equal(svgRun.status, 0, `--svg: exit ${svgRun.status}`);
      assert.match(clsSvg.svg, /<g class="label" style="text-decoration:underline;"/);
      assert.match(clsSvg.svg, /<g class="label" style="font-style:italic;"/);
    });
  });

  describe('V9 icon packs', { skip: iconSkipReason }, () => {
    const inside = (b, o, tol = 0.5) =>
      b[0] >= o[0] - tol && b[1] >= o[1] - tol && b[0] + b[2] <= o[0] + o[2] + tol && b[1] + b[3] <= o[1] + o[3] + tol;
    const inlineOf = (res, id) => res.prims.filter((p) => p.inline === true && unitOf(p, 'node')?.id === id);
    const outlineOf = (res, id) => res.prims.filter((p) => p.kind === 'rect' && /label-container/.test(p.cls) && unitOf(p, 'node')?.id === id);

    it('renders the icon sample', () => {
      assert.equal(icons.ok, true);
    });
    for (const [id, text] of [['A', 'Person'], ['B', 'Code'], ['E', 'Star']]) {
      it(`label icon ${id}: inline shape prims left of "${text}", inside the node, in the node unit`, () => {
        const mine = inlineOf(icons, id);
        assert.ok(mine.length >= 1, `an inline prim in node ${id}`);
        const t = textPrim(icons, text);
        assert.equal(unitOf(t, 'node').id, id);
        const [outline] = outlineOf(icons, id);
        for (const p of mine) {
          assert.ok(['path', 'rect', 'ellipse', 'poly', 'line'].includes(p.kind), p.kind);
          assert.ok(p.bbox.every(isNum) && p.bbox[2] > 0 && p.bbox[3] > 0);
          assert.ok(inside(p.bbox, outline.bbox), `icon ${p.bbox} inside outline ${outline.bbox}`);
          assert.ok(p.bbox[0] + p.bbox[2] <= t.bbox[0] + 1, `icon right edge ${p.bbox[0] + p.bbox[2]} left of text at ${t.bbox[0]}`);
          const mid = p.bbox[1] + p.bbox[3] / 2;
          assert.ok(mid >= t.bbox[1] - 2 && mid <= t.bbox[1] + t.bbox[3] + 2, 'vertically level with the text');
          assert.match(p.fill, HEX);
        }
      });
    }
    it('icon node C: inline icon path above its label, centred, in the node unit', () => {
      const mine = inlineOf(icons, 'C');
      assert.ok(mine.length >= 1);
      const t = textPrim(icons, 'User');
      for (const p of mine) {
        assert.equal(p.kind, 'path');
        assert.ok(p.bbox[1] + p.bbox[3] <= t.bbox[1] + 1, 'above the label');
        assert.ok(near(p.bbox[0] + p.bbox[2] / 2, t.bbox[0] + t.bbox[2] / 2, 4), 'centred on the label');
        assert.ok(p.bbox[2] > 10 && p.bbox[3] > 10);
        assert.match(p.fill, HEX);
      }
    });
    it('icon node D (square form): icon is inline, the frame is not', () => {
      assert.ok(inlineOf(icons, 'D').length >= 1);
      const frame = icons.prims.filter((p) => unitOf(p, 'node')?.id === 'D' && p.kind === 'path' && p.inline !== true);
      assert.ok(frame.length >= 1, 'frame path without inline');
    });
    it('text rows are exactly the label: no icon source text, "?" or private-use glyph', () => {
      const rows = (id) => icons.prims.filter((p) => isText(p) && unitOf(p, 'node')?.id === id).flatMap((p) => p.rows.map(rowText));
      assert.deepEqual(rows('A'), ['Person']);
      assert.deepEqual(rows('B'), ['Code']);
      assert.deepEqual(rows('E'), ['Star']);
      assert.deepEqual(rows('C'), ['User']);
      assert.deepEqual(rows('D'), ['Gh']);
      for (const run of runs(icons)) assert.doesNotMatch(run.t, /fa:|fab:|far:|fa-|\?|[\uE000-\uF8FF]/);
    });
    it('gives ordinary outline prims no inline flag (or false)', () => {
      for (const id of ['A', 'B', 'E']) for (const p of outlineOf(icons, id)) assert.ok(!p.inline, `${id} outline`);
      for (const p of flow.prims) assert.ok(!p.inline, `flowchart prim ${p.i}`);
      for (const p of icons.prims.filter((q) => isText(q))) assert.ok(!p.inline);
    });
  });
});

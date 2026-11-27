# mermaid2drawml 0.3.0

## Conversion

* Any Mermaid diagram type is converted, not only flowcharts: sequence,
  class, state, entity-relationship, Gantt, pie, mind map, timeline, XY and
  quadrant charts and the others all come out as editable Word shapes.

* Colours, line widths, dashes and fonts now come from the styles the browser
  computed, so themes, `classDef`, `style` and `linkStyle` are honoured:
  named colours such as `fill:lightblue`, `fill:none`, transparency,
  subgraph theme fills, and per-node stroke width and font size.

* Node outlines use Mermaid's own stroke width instead of a fixed 1.5 pt.
  Lines thinner than 0.75 pt are drawn at 0.75 pt so that they stay visible;
  the new `min_stroke_pt` argument sets that floor.

* Labels are placed inside their shapes with Mermaid's own line breaks and
  are never re-wrapped by Word. This replaces the separate text-overlay
  shapes of earlier versions: click a box and edit its label.

* Shapes Word has no preset for (double circle, lined and tagged documents,
  odd and flag shapes, stacked rectangles and others) are drawn from their
  real outline instead of an approximation.

* Cylinders have their true height. The horizontal (`h-cyl`) and lined
  (`lin-cyl`) cylinders are no longer swapped.

* Literal `*` and `_` in labels are kept, and edges to node ids that contain
  underscores are no longer lost.

* The drawing's alternative text comes from the diagram's `accTitle` /
  `accDescr`, or from the new `alt_text` argument.

## Interface

* `mermaid_to_word()` and `body_add_mermaid()` gain `mermaid_config`, a list
  of Mermaid configuration (for example `list(theme = "forest")`).

* A Mermaid syntax error now reports Mermaid's own message (error class
  `mermaid2drawml_syntax_error`).

* `build_diagram_xml()` no longer uses `ast`, `margin_in`, `default_fill`,
  `default_stroke`, `default_text_color` or `stroke_width_pt`; they are
  ignored with a warning. Any other unknown argument is an error.

* The result of `mermaid_to_word()` holds `scene` (what was drawn, one row
  per Word shape) and `type` instead of `svg_data` and `ast`. Its print
  method shows the diagram type, the number of shapes and the ids used.

* `render_mermaid()` returns a `mermaid_render` object: the shapes the
  browser drew, with Mermaid's node and edge data, instead of SVG text.

## Page layout

* A landscape or custom-margin diagram no longer changes the page size and
  margins of the pages before it: the earlier content is closed in its own
  section first.

* `orientation = "auto"` turns the page landscape only when the diagram is
  wide and also too wide for a portrait page at its drawn size.

* `page_margins_in` is validated; a bad value is an error instead of being
  silently ignored.

## Setup

* The Node.js packages are pinned and checked: if the installed versions do
  not match, rendering asks you to run `setup_mermaid2drawml()` again. Re-run
  it only when that check asks. `@mermaid-js/parser` is no longer needed, and
  the bridge script runs from the installed package, so it can never be out
  of date.

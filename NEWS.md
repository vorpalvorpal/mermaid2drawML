# mermaid2drawml 0.3.0

After upgrading, run `setup_mermaid2drawml()` again. It installs the icon
packs Mermaid needs to draw icons, and rendering asks for it when the
installed Node.js packages do not match the pinned ones.

## New features

* Any Mermaid diagram type is converted, not only flowcharts: sequence,
  class, state, entity-relationship, Gantt, pie, mind map, timeline, XY and
  quadrant charts and the others all come out as editable Word shapes.

* Colours, line widths, dashes and fonts come from the styles the browser
  computed, so themes, `classDef`, `style` and `linkStyle` are honoured:
  named colours such as `fill:lightblue`, `fill:none`, transparency,
  subgraph theme fills, and per-node stroke width and font size.

* Text transparency is kept: transparent text (a colour with alpha, or text
  under an element with opacity, such as Gantt axis ticks) is written as Word
  text-fill transparency, shown by Word 2010 and later.

* Gradient fills and lines (Sankey links, gradients from `classDef`) are Word
  gradient fills, linear or radial, running the way the SVG ran them. Word's
  radial gradient only approximates SVG's.

* Hyperlinks (`click X href`; `http`, `https` and `mailto` only) are carried
  over. With `links = "shape"` (the default) the shape is the click target.
  With `links = "field"` the text sits in a `HYPERLINK` field, which needs
  nothing outside the XML but links only shapes that hold text. `links =
  "none"` leaves links out.

* Pictures are embedded: image nodes, C4 persons and images in labels, when
  the browser could read them. A remote image needs network access at render
  time. A picture that could not be read, or any picture with
  `embed_images = FALSE`, is left out with a `mermaid2drawml_unsupported`
  warning.

* Icons (`fa:`, `fas:`, `far:` and `fab:` icons, icon nodes, architecture
  service icons) are editable shapes named `mermaid:icon:<node id>`, grouped
  inside their node and no longer changing the node's outline or colours.
  They need the icon packs installed by `setup_mermaid2drawml()`; without
  them Mermaid draws a placeholder.

* The hand-drawn look's hatching is a Word pattern fill in the hatch's
  direction and colour, instead of a flat fill. It is an approximation:
  Word's pattern, not Mermaid's strokes.

* Several diagrams can render in one launch of Node and Chrome, which is most
  of the time a single render takes (a few seconds): give `render_mermaid()`
  or `mermaid_to_word()` a list of sources. A list of N diagrams costs about
  one launch instead of N. A character vector is still one diagram. The
  result is a list; a diagram that fails holds its error condition in its
  slot, and `mermaid_to_word()` numbers each diagram's shapes on from the
  previous one's `next_id`. `body_add_mermaid()` takes one diagram.

* Labels sit inside their shapes with Mermaid's own line breaks and are never
  re-wrapped by Word. This replaces the separate text-overlay shapes of
  earlier versions: click a box and edit its label. Titles and other centred
  labels stay centred when edited; member lists stay left-aligned.

* Node outlines use Mermaid's own stroke width. Lines thinner than 0.75 pt
  are drawn at 0.75 pt so that they stay visible; the new `min_stroke_pt`
  argument sets that floor.

* Shapes Word has no preset for (double circle, lined and tagged documents,
  odd and flag shapes, stacked rectangles and others) are drawn from their
  real outline. Inverted trapezoids and the other shapes that needed a
  vertical flip are drawn the same way, so their text is no longer upside
  down. Simple paths become Word presets: circles, rounded rectangles,
  straight lines, and the four-line frames of sequence diagrams as
  rectangles.

* Boxes of class, ER, state and requirement diagrams are one shape each. Class
  and requirement boxes are stacks of compartments (title, attributes,
  methods), each holding its own text.

* Line ends reach their nodes and take the size of Mermaid's markers.
  Markers Word has no line end for (crow's feet, hollow triangles and
  diamonds, crosses, sequence numbers, and circles or diamonds larger than
  Word's largest end) are small shapes grouped with their line. A mark
  Mermaid draws from several parts (a legend entry, an actor) is one group.

* Shape names: nodes, subgraphs, edges and edge labels are named from
  Mermaid's own node and edge data (`mermaid:node:<id>`,
  `mermaid:subgraph:<id>`, `mermaid:edge:<a>-<b>`, `mermaid:edge_label:<a>:<b>`),
  and composite states have their subgraph names. Block and sankey diagrams
  have no such data, so their shapes are named generically (`mermaid:shape`,
  `mermaid:group:...`). Icons are `mermaid:icon:<id>` and pictures
  `mermaid:image[:<id>]`.

* The drawing's alternative text comes from the diagram's `accTitle` /
  `accDescr`, or from the new `alt_text` argument. Line breaks and tabs in it
  are kept.

* A Mermaid syntax error reports Mermaid's own message.

## Interface changes

* `mermaid_to_word()` returns `scene` (what was drawn, one row per Word shape
  in `scene$items`) and `type` instead of `svg_data` and `ast`, and also
  `links` and `images`. Its print method shows the diagram type, the number
  of shapes and the ids used.

* `build_diagram_xml()` and `mermaid_to_word()` return `links` (`rel_id`,
  `url`) and `images` (`rel_id`, `bytes`) beside `xml` and `next_id`. The XML
  refers to them by placeholder ids, so when they have rows it is not
  complete on its own. Use `body_add_mermaid()`, which adds them to the
  document; or pass `links = "field"` and `embed_images = FALSE` for
  self-contained XML; or register the relationships yourself (see
  `?build_diagram_xml`).

* `render_mermaid()` returns a `mermaid_render` object (the shapes the
  browser drew, with Mermaid's node and edge data) instead of SVG text. The
  SVG is kept only with `render_mermaid(svg = TRUE)`, for looking at a
  diagram whose conversion seems wrong.

* `render_mermaid(timeout = )` limits the whole render, Node and Chrome
  start-up included.

* `max_scale`, `page_width_in` and `page_height_in` are named arguments of
  `mermaid_to_word()` and `body_add_mermaid()`. Passing them by name through
  `...` still works.

* `mermaid_to_word()` and `body_add_mermaid()` gain `mermaid_config`, a list
  of Mermaid configuration (for example `list(theme = "forest")`). An empty
  list sends no configuration; `NULL` members are kept as JSON `null`, and
  numbers keep their full precision.

* The package's conditions are classed and carry data. A syntax error is
  `mermaid2drawml_syntax_error` with `mermaid_message`; a render that runs out
  of time is `mermaid2drawml_timeout` with `timeout` (seconds); any other
  failure of the bridge or browser is `mermaid2drawml_render_error`. The
  `mermaid2drawml_unsupported` warning has `tags` and the
  `mermaid2drawml_small_text` warning has `points`.

* `build_diagram_xml()` no longer uses `ast`, `margin_in`, `default_fill`,
  `default_stroke`, `default_text_color` or `stroke_width_pt`; they are
  ignored with a warning. Any other unknown argument is an error.

* The Node.js packages are pinned and checked. `@mermaid-js/parser` is no
  longer needed, and the bridge script runs from the installed package, so it
  can never be out of date.

## Page layout

* A landscape or custom-margin diagram no longer changes the page size and
  margins of the pages before it: the earlier content is closed in its own
  section first.

* `orientation = "auto"` turns the page landscape only when the diagram is
  wide and also too wide for a portrait page at the size it is drawn
  (`max_scale` included). With `page_width_in` or `page_height_in` given, the
  diagram is fitted to that space and the page stays portrait.

* `page_margins_in` is validated; a bad value is an error instead of being
  silently ignored.

## Fixes

* A hand-drawn node with no fill (`fill:none`, `fill:transparent`) keeps its
  outline instead of being taken for a hatched node.

* `render_mermaid(timeout =)` is per diagram: a list of diagrams is allowed
  that many seconds for each.

* `build_diagram_xml()` takes `NULL` for `page_width_in` and `page_height_in`
  as the default, and a value that is not one finite positive number is an
  error naming the argument.

* The warnings for an old Node.js and an incomplete installation have the
  classes `mermaid2drawml_old_node` and `mermaid2drawml_incomplete_setup`.

* Cylinders have their true height. The horizontal (`h-cyl`) and lined
  (`lin-cyl`) cylinders are no longer swapped.

* Literal `*` and `_` in labels are kept, and edges to node ids that contain
  underscores are no longer lost.

* A line that repeats its last point, or a curve whose end control point sits
  on its end, gets its arrowhead pointing along the line, as in SVG.

* Diagrams larger than about 225 000 px no longer lose their positions and
  sizes (they overflowed 32-bit integers). A coordinate that is not a finite
  number is an error naming the shape.

* Figures a few pixels across (a 3 px diamond or pentagon) are no longer
  taken for ellipses, and a rectangle with only one corner radius is no
  longer rounded.

* A stray moveto at the end of a path no longer stretches its bounding box.

* The bridge's output is read as UTF-8 on every platform, so non-ASCII labels
  survive on Windows.

## Performance

* Building the drawing is about 3.5 times faster and writing the DrawingML
  about twice as fast. Large diagrams are no longer slow: a diagram of 2000
  grouped marks took over two minutes and now takes seconds.

## Not supported or deferred

* Glued edges (a moved node dragging its edges along) are deferred, and may
  be rejected. Word reroutes only its own preset connectors on a drawing
  canvas; using them would mean giving up Mermaid's edge routes and the group
  that holds a diagram together. Not planned unless a workable form is found.

* A browser kept alive between calls is deferred. Rendering several diagrams
  in one call is the supported way to save launches.

* Text is not re-wrapped: labels keep Mermaid's line breaks.

* A picture is stretched to its box: one the browser would crop or letterbox
  comes out distorted (flowchart image nodes are stretched by Mermaid too, so
  they are unaffected).

* An image node whose picture cannot be loaded fails the whole diagram with
  Mermaid's own error; only images inside labels are left out with a warning.

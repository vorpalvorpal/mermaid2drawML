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

* Boxes of class, ER, state and requirement diagrams are one shape each,
  not a fill and an outline drawn twice. Class and requirement boxes are
  stacks of compartments (title, attributes, methods), each holding its
  own text.

* Line ends reach their nodes and take the size of Mermaid's markers.
  Markers Word has no line end for (crow's feet, hollow triangles and
  diamonds, crosses, sequence numbers, and circles or diamonds larger than
  Word's largest end) are drawn as small shapes, grouped with their line.

* A mark Mermaid draws from several parts (a legend entry, an actor, an
  icon) is one group.

* Simple paths become Word presets: circles, rounded rectangles, straight
  lines, and the four-line frames of sequence diagrams as rectangles.

* Titles and other centred labels stay centred when edited in Word; member
  lists stay left-aligned.

* Building the scene is about 3.5 times faster, and large diagrams are no
  longer slow: no step takes time that grows with the square of the number
  of shapes or groups (a diagram of 2000 grouped marks took over two minutes
  and now takes seconds). Writing the DrawingML is about twice as fast.

* Which part of the drawing is which node, subgraph, edge or edge label now
  comes from Mermaid's own data, matched in the browser, instead of being
  guessed from class names and element ids. Diagram types without node and
  edge data (block, sankey) are drawn as loose, scope-grouped shapes.

* Cylinders have their true height. The horizontal (`h-cyl`) and lined
  (`lin-cyl`) cylinders are no longer swapped.

* Literal `*` and `_` in labels are kept, and edges to node ids that contain
  underscores are no longer lost.

* Inverted trapezoids and the other shapes that needed a vertical flip are
  drawn from their real outline, so their text is no longer upside down in
  Word.

* A line that repeats its last point, or a curve whose end control point
  sits on its end, gets its arrowhead pointing along the line, as in SVG.

* Diagrams larger than about 225 000 px no longer lose their positions and
  sizes (they overflowed 32-bit integers); a coordinate that is not a finite
  number is an error naming the shape.

* Figures a few pixels across (a 3 px diamond or pentagon) are no longer
  taken for ellipses, and a rectangle with only one corner radius is no
  longer rounded.

* A stray moveto at the end of a path no longer stretches its bounding box.

* Line breaks and tabs in the diagram's alternative text are kept.

* The drawing's alternative text comes from the diagram's `accTitle` /
  `accDescr`, or from the new `alt_text` argument.

* Transparent text (a colour with alpha, or text under an element with
  opacity, such as Gantt axis ticks) keeps its transparency in Word 2010
  and later.

* Gradient fills and lines (Sankey links, gradients from `classDef`) are
  Word gradient fills, linear or radial, running the way the SVG ran them,
  instead of their first colour.

* Pictures are embedded: an image node, a C4 person, an `<img>` in a label.
  `body_add_mermaid()` adds them to the document; `embed_images = FALSE`
  leaves them out with the `mermaid2drawml_unsupported` warning, as does a
  picture the browser could not load.

* Icons (`fa:` icons in labels, icon nodes, architecture service icons) are
  shapes of their own, `mermaid:icon:<node id>`, grouped inside their node;
  they no longer change the node's outline or colours.

* Links (`click X href`, web and mail addresses only) make the node's shapes
  click targets in Word. `links = "field"` writes them as `HYPERLINK`
  fields around the text instead, and `links = "none"` leaves them out.

* The hand-drawn look's hatched fills are Word line patterns in the hatch's
  colour and direction, instead of a flat fill.

## Interface

* `build_diagram_xml()` returns `links` and `images` beside `xml` and
  `next_id`, and `mermaid_to_word()` returns them too: the hyperlinks and
  pictures the XML refers to by placeholder ids. When they have rows, XML
  inserted with `officer::body_add_xml()` needs them added to the document
  (see `?build_diagram_xml`); `body_add_mermaid()` does this itself.

* `mermaid_to_word()` and `body_add_mermaid()` gain `mermaid_config`, a list
  of Mermaid configuration (for example `list(theme = "forest")`).

* A Mermaid syntax error now reports Mermaid's own message (error class
  `mermaid2drawml_syntax_error`). A render that runs out of time is
  `mermaid2drawml_timeout`, and any other failure of the bridge or browser
  is `mermaid2drawml_render_error`.

* `render_mermaid(timeout = )` limits the whole render; the bridge's own
  limit for the page is that less a start-up allowance (at least 5 seconds).

* Several diagrams render in one launch of Node and Chrome, which is most of
  the time a render takes: give `render_mermaid()` or `mermaid_to_word()` a
  list of sources (a character vector is still one diagram). They return a
  list; a diagram that fails holds its error condition in its slot, and
  `mermaid_to_word()` numbers each diagram's shapes on from the previous
  one's `next_id`.

* `mermaid_to_word()` and `body_add_mermaid()` take `max_scale`,
  `page_width_in` and `page_height_in` as named arguments (they were passed
  through `...` before, which still works by name).

* The package's conditions carry their data as fields:
  `mermaid2drawml_unsupported` has `tags`, `mermaid2drawml_small_text` has
  `points`, `mermaid2drawml_syntax_error` has `mermaid_message` (Mermaid's
  own text) and `mermaid2drawml_timeout` has `timeout` (seconds).

* `render_mermaid()` no longer returns the SVG text unless asked:
  `render_mermaid(svg = TRUE)` keeps it in `svg`, for looking at a diagram
  whose conversion seems wrong.

* An empty `mermaid_config` sends no configuration; `NULL` members are kept
  as JSON `null`, and numbers keep their full precision.

* The bridge's output is read as UTF-8 on every platform, so non-ASCII
  labels survive on Windows.

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
  wide and also too wide for a portrait page at the size it is drawn
  (`max_scale` included). With `page_width_in` or `page_height_in` given,
  the diagram is fitted to that space and the page stays portrait.

* `page_margins_in` is validated; a bad value is an error instead of being
  silently ignored.

## Setup

* The Node.js packages are pinned and checked: if the installed versions do
  not match, rendering asks you to run `setup_mermaid2drawml()` again. Re-run
  it only when that check asks. `@mermaid-js/parser` is no longer needed, and
  the bridge script runs from the installed package, so it can never be out
  of date.

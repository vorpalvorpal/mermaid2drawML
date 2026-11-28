# mermaid2drawml

mermaid2drawml converts [Mermaid](https://mermaid.js.org/) diagrams into
Microsoft Word shapes you can edit. Unlike a picture, the result can be
selected, moved, restyled and relabelled inside Word, and each label sits
inside its shape.

Mermaid lays out and draws the diagram in a headless browser. The browser
reports every shape it drew, with its computed colours, line styles, fonts
and text layout, so the Word drawing matches what you see in a browser,
including themes, `classDef`, `style` and `linkStyle`. Diagrams go into
documents through [officer](https://davidgohel.github.io/officer/).

## Setup

The package needs Node.js (>= 18) and npm. After installing it, run this
once to install the Node.js packages:

```r
mermaid2drawml::setup_mermaid2drawml()
```

The packages include the Font Awesome icon packs Mermaid needs to draw icons.
If you are upgrading from an earlier version, run `setup_mermaid2drawml()`
again: without the icon packs, Mermaid draws a placeholder instead of an icon.
Rendering tells you when the installed packages do not match the ones this
version pins.

## Usage

```r
library(officer)
library(mermaid2drawml)

read_docx() |>
  body_add_mermaid("flowchart LR\n  A[Input] --> B[Transform] --> C[Output]") |>
  print(target = "diagram.docx")
```

`mermaid_to_word()` returns the DrawingML XML (`xml`), a table of what was
drawn (`scene$items`, one row per Word shape), and the hyperlinks and pictures
the XML refers to (`links`, `images`). The vignette
(`vignette("mermaid2drawml")`) covers sizing, fonts, R Markdown, several
diagrams in one document, handling problems, and building documents yourself.

## Diagram types

Every type Mermaid draws is converted.

Good:

- **flowchart**: all node shapes (a Word preset where Word has the same
  figure, otherwise the exact outline), all edge kinds, subgraphs,
  `classDef` / `style` / `linkStyle`, markdown labels, the ELK layout and
  the hand-drawn look (hatching is an approximation; see below).
- **sequence**, **gantt**, **state**, **gitGraph**, **journey**, **block**,
  **packet**, **radar**, **treemap**, **xychart**. Block diagrams have no node
  data, so their shapes have generic names.

Good, with caveats:

- **class**: each box is a stack of compartments (title, attributes,
  methods) holding their text; UML line ends (hollow triangles, diamonds)
  are small shapes grouped with their line.
- **ER**: each entity is a name compartment over a body compartment; the
  attribute cells are separate texts, grouped with their entity; crow's
  feet are small shapes.
- **requirement**: each box is two compartments holding their text; the
  "contains" end is a small shape.
- **pie**: slices are drawn from their outline; the title, legend and
  slice labels are text boxes of their own.
- **quadrant**: point and axis labels are text boxes of their own.
- **timeline**, **mindmap**: most boxes are drawn from their outline
  rather than a Word preset.
- **kanban**: a card with several texts keeps them as separate text boxes.
- **sankey**: links keep their colour ramp as a gradient. Shapes have generic
  names, as the diagram has no node data.
- **architecture**: icons are editable shapes, but need the icon packs that
  `setup_mermaid2drawml()` installs.
- **C4**: person images are embedded as pictures when the browser could read
  them, and the texts of a box are separate text boxes over it.

## Also supported

- **Text transparency.** Text that is not opaque (a colour with alpha, or text
  under an element with opacity, such as Gantt axis ticks) is written as Word
  text-fill transparency. Word 2010 and later show it.
- **Gradients.** Gradients on fills and lines, linear and radial, become Word
  gradient fills. Sankey bands keep their colour ramp. Word's radial gradient
  only approximates SVG's, so a radial one can look a little different.
- **Hyperlinks.** `click X href` links are carried over. Only `http`, `https`
  and `mailto` targets are kept. `links = "shape"` (the default) makes the
  shape the click target; it needs a relationship in the document, which
  `body_add_mermaid()` registers. `links = "field"` wraps the text in a
  `HYPERLINK` field instead; it is self-contained XML, but links only shapes
  that hold text. `links = "none"` leaves links out.
- **Pictures.** Image nodes and images in labels are embedded as pictures when
  the browser could read them. A remote image needs network access at render
  time; if it could not be read, it is left out with a
  `mermaid2drawml_unsupported` warning. `embed_images = FALSE` leaves all
  pictures out.
- **Icons.** `fa:`, `fas:`, `far:` and `fab:` icons and icon nodes are drawn as
  editable shapes named `mermaid:icon:<node id>`. This needs the icon packs
  that `setup_mermaid2drawml()` installs; without them Mermaid draws a
  placeholder.
- **Hand-drawn look.** Hatching becomes a Word pattern fill in the hatch's
  direction and colour. It is an approximation: Word's pattern, not Mermaid's
  strokes.

## Several diagrams

Starting Node and Chrome dominates the time one render takes (a few seconds).
Give `render_mermaid()` or `mermaid_to_word()` a list of diagrams and they are
rendered in one launch, so N diagrams cost about one launch instead of N. A
character vector is still one diagram. A diagram that fails leaves its
condition in its slot of the result. `body_add_mermaid()` takes one diagram.

```r
ds <- mermaid_to_word(list("pie\n  \"a\" : 1", "flowchart LR\n  A --> B"))
```

## Building documents yourself

`mermaid_to_word()` and `build_diagram_xml()` return `links` (`rel_id`, `url`)
and `images` (`rel_id`, `bytes`) beside `xml` and `next_id`. The XML refers to
them by placeholder ids, so when they have rows the XML alone is not
complete. Either:

- use `body_add_mermaid()`, which adds them and gives a complete document;
- pass `links = "field"` and `embed_images = FALSE` for self-contained XML; or
- register the relationships yourself (see `?build_diagram_xml`).

## Not supported or deferred

- **Glued edges** (a moved node dragging its edges along) are deferred, and
  may be rejected. Word reroutes only its own preset connectors on a drawing
  canvas. Using them would mean giving up Mermaid's edge routes and the group
  that holds a diagram together. They are not planned unless a workable form
  is found. Until then, grouping keeps a node with its parts, so move groups.
- **A browser kept alive between calls** is deferred. Rendering several
  diagrams in one call is the supported way to save launches.
- **Re-wrapping text.** Labels keep Mermaid's line breaks; Word does not
  re-wrap them.

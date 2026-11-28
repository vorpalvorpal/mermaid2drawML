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

## Usage

```r
library(officer)
library(mermaid2drawml)

read_docx() |>
  body_add_mermaid("flowchart LR\n  A[Input] --> B[Transform] --> C[Output]") |>
  print(target = "diagram.docx")
```

`mermaid_to_word()` returns the DrawingML XML and a table of what was drawn,
one row per Word shape. The vignette (`vignette("mermaid2drawml")`) covers
sizing, fonts, R Markdown and several diagrams in one document.

## Diagram types

Every type Mermaid draws is converted.

Good:

- **flowchart**: all node shapes (a Word preset where Word has the same
  figure, otherwise the exact outline), all edge kinds, subgraphs,
  `classDef` / `style` / `linkStyle`, markdown labels, the ELK layout and
  the hand-drawn look (without its hatching).
- **sequence**, **gantt**, **state**, **gitGraph**, **journey**, **block**,
  **packet**, **radar**, **treemap**, **xychart**.

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
- **sankey**: links have flat colours instead of gradients.
- **architecture**: icons are groups of shapes.
- **C4**: person images are left out, and the texts of a box are separate
  text boxes over it.

## Limits

- Embedded images and icon fonts are left out, with a warning.
- Hyperlinks (`click`) are not carried over.
- Text transparency is ignored; gradients become flat colours.
- The hand-drawn look loses its hatching.
- Lines are not glued to shapes: moving a node in Word does not drag its
  edges. Grouping keeps a node with its parts, so move groups.
- Text is not re-wrapped: labels keep Mermaid's line breaks.

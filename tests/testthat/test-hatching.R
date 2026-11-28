# Hand-drawn hatching: in Mermaid's hand-drawn look a fill is a stroke-only
# zigzag path (rough.js hachure) under the rough outline. Word has no
# hachure, so the node's one shape gets a preset pattern fill instead of a
# flat colour: items carry fill_pattern (a DrawingML preset name, NA for
# none), the emitter writes a:pattFill with the hatch colour as foreground
# and nothing behind it. The preset follows the hatch direction. A solid
# hand-drawn fill stays solid, and the classic look never gets a pattern.

node_a <- function(r) item_named(build_scene(r)$items, "mermaid:node:A")

node_a_xml <- function(r) shape_named(dml_doc(build_diagram_xml(r)$xml), "mermaid:node:A")

# A pattern fill with nothing behind the hatch: a fully transparent a:bgClr
# (the schema needs no a:bgClr, but Word reads one).
expect_no_background <- function(patt, info = NULL) {
  testthat::expect_true(!is.null(patt) && patt$has_bg && isTRUE(patt$bg_alpha == 0),
                        info = info)
}

test_that("every item has a fill_pattern column, NA for an ordinary fill", {
  items <- generic_items(list(prim_rect(10, 10, 50, 20)))
  expect_true("fill_pattern" %in% names(items))
  expect_type(items$fill_pattern, "character")
  expect_true(is.na(items$fill_pattern))
})

test_that("a hand-drawn rect with a hachure fill is one shape with a pattern and its outline", {
  r <- handdrawn_node_render(angle = -41, hatch = "#EEEEEE", outline = "#999999")
  items <- build_scene(r)$items
  a <- item_named(items, "mermaid:node:A")
  expect_equal(nrow(a), 1L)
  expect_equal(a$kind, "shape")
  expect_equal(a$prst, "rect")
  expect_equal(a$fill_pattern, "ltUpDiag")
  expect_equal(a$stroke, "999999")
  expect_equal(a$stroke_w, 1.3)
  expect_identical(a$label, "Label")
  expect_equal(sum(items$kind != "group"), 1L)
})

test_that("in the XML the hatch is a:pattFill with the hatch colour in front", {
  r <- handdrawn_node_render(angle = -41, hatch = "#EEEEEE", outline = "#999999")
  doc <- dml_doc(build_diagram_xml(r)$xml)
  node <- shape_named(doc, "mermaid:node:A")
  sp <- shape_sppr(node)
  expect_true(is_missing_node(xfirst(sp, "./a:solidFill")))
  patt <- patt_of(sp)
  expect_false(is.null(patt))
  expect_equal(patt$prst, "ltUpDiag")
  expect_equal(patt$fg, "EEEEEE")
  expect_no_background(patt)
  ln <- shape_ln(node)
  expect_equal(ln$kind, "solid")
  expect_equal(ln$color, "999999")
  expect_length(schema_order_problems(doc), 0L)
})

test_that("the preset follows the direction of the hatch strokes", {
  # Angles in SVG terms (y down): negative runs up to the right.
  cases <- c(
    "-41" = "ltUpDiag",  # rough.js's default hachure angle
    "-49" = "ltUpDiag",  # what flowchart-handdrawn's hatches measure
    "135" = "ltUpDiag",  # the same strokes drawn the other way
    "45" = "ltDnDiag", "-135" = "ltDnDiag",
    "0" = "ltHorz", "8" = "ltHorz", "180" = "ltHorz",
    "90" = "ltVert", "-80" = "ltVert"
  )
  for (ang in names(cases)) {
    r <- handdrawn_node_render(angle = as.numeric(ang))
    expect_equal(node_a(r)$fill_pattern, cases[[ang]], info = ang)
    expect_equal(patt_of(shape_sppr(node_a_xml(r)))$prst, cases[[ang]], info = ang)
  }
})

test_that("a solid hand-drawn fill stays a solid fill", {
  fill <- rough_rect(20, 20, 100, 60, fill = "#EEEEEE", stroke = NULL)
  r <- handdrawn_node_render(fill_prim = fill, outline = "#999999")
  a <- node_a(r)
  expect_true(is.na(a$fill_pattern))
  expect_equal(a$fill, "EEEEEE")
  sp <- shape_sppr(node_a_xml(r))
  expect_null(patt_of(sp))
  expect_equal(paint_of(sp)[c("kind", "color")], list(kind = "solid", color = "EEEEEE"))
})

test_that("the classic look never gets a pattern", {
  r <- handdrawn_node_render(look = "classic")
  expect_true(is.na(node_a(r)$fill_pattern))
  expect_null(patt_of(shape_sppr(node_a_xml(r))))
})

# -- flowchart-handdrawn ----------------------------------------------------------

test_that("flowchart-handdrawn: every node and the subgraph are hatched with their fill colour", {
  items <- fixture_scene("flowchart-handdrawn")$items
  expected <- tibble::tribble(
    ~name,                    ~fg,      ~stroke,
    "mermaid:node:A",         "EEEEEE", "999999",
    "mermaid:node:B",         "EEEEEE", "999999",
    "mermaid:node:C",         "EEEEEE", "999999",
    "mermaid:node:D",         "EEEEEE", "999999",
    "mermaid:node:E",         "EEEEEE", "999999",
    "mermaid:node:F",         "EEEEEE", "999999",
    "mermaid:subgraph_bg:G",  "FCFCFC", "707070"
  )
  doc <- fixture_doc("flowchart-handdrawn")
  for (k in seq_len(nrow(expected))) {
    nm <- expected$name[[k]]
    it <- item_named(items, nm)
    expect_equal(it$fill_pattern, "ltUpDiag", info = nm)
    expect_equal(it$stroke, expected$stroke[[k]], info = nm)
    node <- shape_named(doc, nm)
    patt <- patt_of(shape_sppr(node))
    expect_equal(patt$prst, "ltUpDiag", info = nm)
    expect_equal(patt$fg, expected$fg[[k]], info = nm)
    expect_no_background(patt, info = nm)
    expect_equal(shape_ln(node)$color, expected$stroke[[k]], info = nm)
  }
  # Edge labels have solid backgrounds and keep them.
  labels <- items[startsWith(items$name, "mermaid:edge_label"), ]
  expect_equal(labels$fill_pattern, rep(NA_character_, nrow(labels)))
  expect_length(schema_order_problems(doc), 0L)
})

test_that("no diagram in the classic look gets a pattern fill", {
  for (name in setdiff(all_fixture_names(), "flowchart-handdrawn")) {
    items <- fixture_scene(name)$items
    expect_equal(items$fill_pattern, rep(NA_character_, nrow(items)), info = name)
    expect_length(xfind(fixture_doc(name), "//a:pattFill"), 0L)
  }
})

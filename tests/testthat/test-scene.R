# build_scene(): the scene object, generic conversion of prims (spec 3a) and
# putting text inside shapes (spec 3b). Hand-built renders give exact boxes;
# the loop at the end runs every bridge fixture.

scene_cols <- c("name", "kind", "parent", "x", "y", "w", "h", "rot", "prst",
                "fill", "stroke", "fill_a", "stroke_a", "stroke_w", "head",
                "tail", "label")
end_kinds <- c("none", "triangle", "arrow", "stealth", "diamond", "oval")

shapes_of <- function(items) items[items$kind == "shape", ]
texts_of <- function(items) items[items$kind == "text", ]

# -- The scene object ----------------------------------------------------------

test_that("build_scene() returns a mermaid_scene with the documented fields", {
  r <- generic_render(list(prim_rect(10, 10, 50, 20)), viewBox = c(-5, -6, 300, 120),
                      title = "My title", desc = "My description")
  sc <- build_scene(r)
  expect_s3_class(sc, "mermaid_scene")
  expect_equal(sc$type, "sequence")
  expect_equal(as.numeric(sc$viewbox), c(-5, -6, 300, 120))
  expect_equal(sc$title, "My title")
  expect_equal(sc$desc, "My description")
  expect_type(sc$unsupported, "character")
  expect_length(sc$unsupported, 0L)
  expect_s3_class(sc$items, "tbl_df")
  expect_true(all(scene_cols %in% names(sc$items)))
})

test_that("a render with no prims gives an empty items tibble with every column", {
  items <- generic_items(list())
  expect_s3_class(items, "tbl_df")
  expect_equal(nrow(items), 0L)
  expect_true(all(scene_cols %in% names(items)))
})

test_that("items keep paint order", {
  items <- generic_items(list(
    prim_rect(0, 0, 10, 10, fill = "#110000"),
    prim_ellipse(50, 50, 5),
    prim_rect(100, 0, 10, 10, fill = "#330000")
  ))
  expect_equal(items$prst, c("rect", "ellipse", "rect"))
  expect_equal(items$fill, c("110000", "FFFFFF", "330000"))
})

# -- 3a: rect, ellipse, line ---------------------------------------------------

test_that("a rect prim becomes a preset rect item with its paint copied", {
  items <- generic_items(list(
    prim_rect(10, 20, 100, 40, fill = "#ABCDEF", stroke = "#123456", stroke_w = 2,
              fill_a = 0.5, stroke_a = 0.25)
  ))
  expect_equal(nrow(items), 1L)
  it <- items[1, ]
  expect_equal(it$kind, "shape")
  expect_equal(it$prst, "rect")
  expect_near(c(it$x, it$y, it$w, it$h), c(10, 20, 100, 40), tol = 1e-9)
  expect_equal(it$rot, 0)
  expect_true(is.na(it$parent))
  expect_equal(it$fill, "ABCDEF")
  expect_equal(it$stroke, "123456")
  expect_equal(it$fill_a, 0.5)
  expect_equal(it$stroke_a, 0.25)
  expect_equal(it$stroke_w, 2)
  expect_equal(it$head, "none")
  expect_equal(it$tail, "none")
  expect_identical(it$label, "")
})

test_that("a missing fill or stroke is NA, not a colour", {
  items <- generic_items(list(prim_rect(10, 20, 100, 40, fill = NULL, stroke = NULL)))
  expect_true(is.na(items$fill))
  expect_true(is.na(items$stroke))
})

test_that("a rect with rx > 0 is a roundRect", {
  items <- generic_items(list(prim_rect(10, 20, 100, 40, rx = 8)))
  expect_equal(items$prst, "roundRect")
})

test_that("an upright matrix gives the box directly (translate, scale, flip)", {
  items <- generic_items(list(
    prim_rect(0, 0, 50, 20, m = c(1, 0, 0, 1, 30, 40)),
    prim_rect(0, 0, 20, 10, m = c(2, 0, 0, 2, 10, 10)),
    prim_rect(0, 0, 30, 10, m = c(-1, 0, 0, 1, 100, 0))   # a mirror is upright
  ), viewBox = c(0, 0, 200, 200))
  expect_equal(unlist(items[1, c("x", "y", "w", "h")]), c(x = 30, y = 40, w = 50, h = 20))
  expect_equal(unlist(items[2, c("x", "y", "w", "h")]), c(x = 10, y = 10, w = 40, h = 20))
  expect_equal(unlist(items[3, c("x", "y", "w", "h")]), c(x = 70, y = 0, w = 30, h = 10))
  expect_equal(items$rot, c(0, 0, 0))
})

test_that("a rotated matrix gives the unrotated box plus rot", {
  # x' = -y + 100, y' = x + 50: the 60 x 20 rect is turned 90 degrees
  # clockwise about its centre (90, 80).
  items <- generic_items(list(prim_rect(0, 0, 60, 20, m = c(0, 1, -1, 0, 100, 50))),
                         viewBox = c(0, 0, 200, 200))
  expect_near(items$rot, 90, tol = 1e-6)
  expect_near(unlist(items[1, c("x", "y", "w", "h")]), c(x = 60, y = 70, w = 60, h = 20),
               tol = 1e-6)
})

test_that("an ellipse prim becomes a preset ellipse with its bounding box", {
  items <- generic_items(list(prim_ellipse(50, 40, rx = 30, ry = 20)))
  expect_equal(items$kind, "shape")
  expect_equal(items$prst, "ellipse")
  expect_equal(unlist(items[1, c("x", "y", "w", "h")]), c(x = 20, y = 20, w = 60, h = 40))
})

test_that("a line prim becomes a line item with preset line", {
  items <- generic_items(list(prim_line(100, 80, 20, 10, stroke = "#00FF00", stroke_w = 3)))
  expect_equal(items$kind, "line")
  expect_equal(items$prst, "line")
  expect_equal(unlist(items[1, c("x", "y", "w", "h")]), c(x = 20, y = 10, w = 80, h = 70))
  expect_equal(items$stroke, "00FF00")
  expect_equal(items$stroke_w, 3)
  expect_true(is.na(items$fill))
})

# -- 3a: poly and path ---------------------------------------------------------

test_that("a rectangle-shaped path is a preset rect, in either direction", {
  for (d in c("M10,10 L110,10 L110,60 L10,60 Z",
              "M10,10 L10,60 L110,60 L110,10 Z",
              "M10,10 L110,10 L110,60 L10,60 L10,10 Z",
              "M10 10 H110 V60 H10 Z")) {
    items <- generic_items(list(prim_path(d, bbox = c(10, 10, 100, 50))))
    expect_equal(items$prst, "rect", info = d)
    expect_equal(unlist(items[1, c("x", "y", "w", "h")]), c(x = 10, y = 10, w = 100, h = 50),
                 info = d)
  }
})

test_that("a rectangle-shaped closed poly is a preset rect", {
  items <- generic_items(list(prim_poly(list(c(0, 0), c(80, 0), c(80, 30), c(0, 30)),
                                        m = c(1, 0, 0, 1, 20, 10))))
  expect_equal(items$prst, "rect")
  expect_equal(unlist(items[1, c("x", "y", "w", "h")]), c(x = 20, y = 10, w = 80, h = 30))
})

test_that("a four-point poly that is not a rectangle is custom geometry (no guessing from point counts)", {
  diamond <- prim_poly(list(c(50, 0), c(100, 25), c(50, 50), c(0, 25)))
  items <- generic_items(list(diamond))
  expect_equal(items$kind, "shape")
  expect_true(is.na(items$prst))
})

test_that("a closed curved path is custom geometry with its true extent", {
  # The curve from (100, 0) back to (0, 0) bulges to y = 37.5, short of its
  # control points at y = 50.
  p <- prim_path("M0,0 L100,0 C100,50 0,50 0,0 Z", m = c(1, 0, 0, 1, 20, 30),
                 bbox = c(20, 30, 100, 37.5))
  items <- generic_items(list(p))
  expect_equal(items$kind, "shape")
  expect_true(is.na(items$prst))
  expect_near(unlist(items[1, c("x", "y", "w", "h")]),
               c(x = 20, y = 30, w = 100, h = 37.5), tol = 1e-6)
  expect_equal(items$fill, "FFFFFF")
})

test_that("an open path with a fill is filled, as SVG fills it", {
  # SVG fills an open subpath as if it were closed but strokes only the drawn
  # segments, so this stays custom geometry: a preset rect would gain a side.
  p <- prim_path("M0,0 L100,0 L100,50 L0,50", fill = "#FF0000", stroke = "#0000FF")
  items <- generic_items(list(p))
  expect_equal(items$kind, "shape")
  expect_equal(items$fill, "FF0000")
  expect_equal(items$stroke, "0000FF")
  expect_true(is.na(items$prst))
})

test_that("an open rectangle path with a fill and no stroke is a preset rect", {
  # Mermaid draws many boxes as an unclosed fill path plus a separate outline.
  p <- prim_path("M10,10 L110,10 L110,60 L10,60", fill = "#FF0000", stroke = NULL)
  items <- generic_items(list(p))
  expect_equal(items$prst, "rect")
  expect_equal(items$fill, "FF0000")
  expect_true(is.na(items$stroke))
})

test_that("a fill-only shape and the stroke-only outline over it become one shape", {
  fill <- prim_path("M10,10 L110,10 L110,60 L10,60", fill = "#ECECFF", stroke = NULL)
  line <- prim_path("M10,10 L110,10 L110,60 L10,60 L10,10", fill = NULL,
                    stroke = "#9370DB", stroke_w = 1.3)
  items <- generic_items(list(fill, line))
  expect_equal(nrow(items), 1L)
  expect_equal(items$prst, "rect")
  expect_equal(items$fill, "ECECFF")
  expect_equal(items$stroke, "9370DB")
  expect_equal(items$stroke_w, 1.3)
  expect_equal(unlist(items[c("x", "y", "w", "h")]),
               c(x = 10, y = 10, w = 100, h = 50), tolerance = 1e-6)
})

test_that("a fill and an outline with different boxes stay two shapes", {
  fill <- prim_rect(10, 10, 100, 50, fill = "#ECECFF", stroke = NULL)
  line <- prim_rect(10, 10, 100, 20, fill = NULL, stroke = "#9370DB")
  expect_equal(nrow(generic_items(list(fill, line))), 2L)
})

test_that("stacked HTML boxes with the same box collapse into one, alphas composited", {
  # Mermaid paints an edge label's background three times (div, span, p).
  bg <- function(a) prim_rect(50, 20, 40, 24, fill = "#E8E8E8", fill_a = a,
                              stroke = NULL, tag = "div", html = TRUE)
  items <- generic_items(list(bg(0.5), bg(0.8), bg(0.8)))
  expect_equal(nrow(items), 1L)
  expect_equal(items$fill, "E8E8E8")
  # 1 - (1 - 0.5) * (1 - 0.8) * (1 - 0.8) = 0.98
  expect_equal(items$fill_a, 0.98, tolerance = 1e-6)
})

test_that("items wholly outside the viewBox are dropped", {
  # Mermaid's gantt "today" marker can sit tens of thousands of px off-canvas.
  items <- generic_items(list(
    prim_rect(10, 10, 40, 20),
    prim_line(5000, 0, 5000, 100, stroke = "#FF0000"),
    prim_text("far", x = -900, y = 10, w = 30, h = 12),
    prim_rect(190, 90, 40, 20)          # straddles the edge: kept
  ), viewBox = c(0, 0, 200, 100))
  expect_equal(nrow(items), 2L)
  expect_true(all(items$kind == "shape"))
})

# -- 3a: markers ---------------------------------------------------------------

test_that("marker_start is the head (path start) and marker_end the tail (path end)", {
  items <- generic_items(list(prim_line(10, 10, 100, 10,
                                        marker_start = "class-aggregationStart",
                                        marker_end = "class-dependencyEnd")))
  expect_equal(items$head, "diamond")
  expect_equal(items$tail, "arrow")
})

test_that("a path or line without markers has no ends", {
  items <- generic_items(list(prim_line(10, 10, 100, 10),
                              prim_path("M0,0 L50,50", fill = NULL)))
  expect_equal(items$head, c("none", "none"))
  expect_equal(items$tail, c("none", "none"))
})

test_that("line_end_kind() maps each documented marker id pattern", {
  # step 2: C1 corrected the table; with no marker definition (as here) an
  # end drawn as a shape falls back to its nearest filled Word end.
  expected <- c(
    "flowchart-v2-pointEnd" = "triangle",
    "flowchart-v2-pointStart" = "triangle",
    "flowchart-v2-pointEnd__f00" = "triangle",
    "block-pointEnd" = "triangle",
    "arrowhead" = "triangle",
    "filled-head" = "stealth",                     # step 2: C1
    "stateDiagram-barbEnd" = "stealth",            # step 2: C1
    "class-extensionEnd" = "triangle",
    "class-extensionStart" = "triangle",
    "requirement-requirement_arrowEnd" = "arrow",  # step 2: C1
    "class-dependencyEnd" = "arrow",
    "my-stickEnd" = "arrow",
    "flowchart-v2-circleEnd" = "oval",
    "class-lollipopStart" = "oval",
    "flowchart-v2-crossEnd" = "none",              # step 2: C1 (no X end in Word)
    "crosshead" = "none",                          # step 2: C1
    "class-aggregationStart" = "diamond",
    "class-compositionStart" = "diamond",
    "sequencenumber" = "none",
    "er-onlyOneStart" = "none",
    "er-zeroOrOneEnd" = "none",
    "er-oneOrMoreEnd" = "none",
    "er-oneOrMoreStart" = "none",
    "er-zeroOrMoreEnd" = "none",
    "something-unheard-of" = "triangle"
  )
  got <- vapply(names(expected), function(id) line_end_kind(id), character(1))
  expect_equal(got, expected)
  expect_true(all(got %in% end_kinds))
})

# -- 3a: text, images, other ---------------------------------------------------

test_that("a lone text prim becomes a free text item at its own box", {
  items <- generic_items(list(prim_text(c("one", "two"), x = 30, y = 40, w = 50, h = 40,
                                        anchor = "start")))
  expect_equal(items$kind, "text")
  expect_true(is.na(items$prst))
  expect_equal(unlist(items[1, c("x", "y", "w", "h")]), c(x = 30, y = 40, w = 50, h = 40))
  expect_identical(items$label, "one\ntwo")
})

test_that("rotated text carries its rotation", {
  items <- generic_items(list(prim_text("Revenue", x = 10, y = 40, w = 60, h = 20, rot = -90)))
  expect_equal(items$kind, "text")
  expect_equal(items$rot, -90)
  expect_equal(unlist(items[1, c("x", "y", "w", "h")]), c(x = 10, y = 40, w = 60, h = 20))
})

test_that("image and other prims are not drawn but listed as unsupported", {
  sc <- build_scene(generic_render(list(
    prim_rect(0, 0, 50, 50),
    prim_image(60, 0, 30, 30),
    prim_other("use", 100, 0, 20, 20)
  )))
  expect_equal(nrow(sc$items), 1L)
  expect_true(all(c("image", "use") %in% sc$unsupported))
})

test_that("flowchart-full lists its image as unsupported; flowchart-paint has none", {
  expect_true("image" %in% fixture_scene("flowchart-full")$unsupported)
  expect_length(fixture_scene("flowchart-paint")$unsupported, 0L)
})

# -- 3b: putting text inside shapes --------------------------------------------

test_that("text centred in a rect merges into it", {
  items <- generic_items(list(
    prim_rect(20, 20, 100, 40),
    prim_text("Hi", x = 50, y = 30, w = 40, h = 20)
  ))
  expect_equal(nrow(items), 1L)
  expect_equal(items$kind, "shape")
  expect_identical(items$label, "Hi")
})

test_that("text at the top left of a larger rect still merges", {
  items <- generic_items(list(
    prim_rect(0, 0, 200, 100),
    prim_text("Title", x = 10, y = 5, w = 60, h = 20, anchor = "start")
  ))
  expect_equal(nrow(items), 1L)
  expect_identical(items$label, "Title")
})

test_that("two texts in one rect both stay free text boxes", {
  items <- generic_items(list(
    prim_rect(0, 0, 200, 100),
    prim_text("first", x = 10, y = 10, w = 50, h = 20),
    prim_text("second", x = 10, y = 50, w = 50, h = 20)
  ))
  expect_identical(shapes_of(items)$label, "")
  expect_setequal(texts_of(items)$label, c("first", "second"))
})

test_that("a text only partly inside a rect does not merge", {
  # Centre (90, 20) is inside, but only 40 of the 60 px width (67 %) is.
  items <- generic_items(list(
    prim_rect(0, 0, 100, 50),
    prim_text("spill", x = 60, y = 10, w = 60, h = 20)
  ))
  expect_identical(shapes_of(items)$label, "")
  expect_identical(texts_of(items)$label, "spill")
})

test_that("text more than 80 % inside a rect does merge", {
  # 50 of 60 px inside (83 %), centre inside.
  items <- generic_items(list(
    prim_rect(0, 0, 100, 50),
    prim_text("most", x = 50, y = 10, w = 60, h = 20)
  ))
  expect_identical(shapes_of(items)$label, "most")
  expect_equal(nrow(texts_of(items)), 0L)
})

test_that("the smallest containing shape wins", {
  items <- generic_items(list(
    prim_rect(0, 0, 200, 100, fill = "#AA0000"),
    prim_rect(50, 25, 100, 50, fill = "#00AA00"),
    prim_text("T", x = 80, y = 40, w = 40, h = 20)
  ))
  expect_equal(nrow(items), 2L)
  expect_identical(items$label[items$fill == "00AA00"], "T")
  expect_identical(items$label[items$fill == "AA0000"], "")
})

test_that("a shape painted after the text is not a candidate", {
  items <- generic_items(list(
    prim_text("under", x = 50, y = 30, w = 40, h = 20),
    prim_rect(20, 20, 100, 40)
  ))
  expect_identical(shapes_of(items)$label, "")
  expect_identical(texts_of(items)$label, "under")
})

test_that("text over an open path or a line does not merge", {
  items <- generic_items(list(
    prim_path("M0,0 L100,0 L100,50 L0,50", fill = NULL),
    prim_text("open", x = 30, y = 15, w = 40, h = 20)
  ))
  expect_identical(texts_of(items)$label, "open")
  expect_identical(shapes_of(items)$label, "")

  items <- generic_items(list(
    prim_line(0, 25, 100, 25),
    prim_text("on a line", x = 30, y = 15, w = 40, h = 20)
  ))
  expect_identical(texts_of(items)$label, "on a line")
  expect_identical(items$label[items$kind == "line"], "")
})

test_that("text merges into a stroked-only or a filled-only closed shape", {
  items <- generic_items(list(
    prim_rect(0, 0, 100, 50, fill = NULL),
    prim_text("outline", x = 30, y = 15, w = 40, h = 20),
    prim_rect(0, 60, 100, 30, stroke = NULL),
    prim_text("fill", x = 30, y = 65, w = 40, h = 20)
  ), viewBox = c(0, 0, 200, 100))
  expect_equal(nrow(texts_of(items)), 0L)
  expect_setequal(shapes_of(items)$label, c("outline", "fill"))
})

test_that("rotated text does not merge into an upright shape and keeps rot", {
  items <- generic_items(list(
    prim_rect(20, 20, 100, 40),
    prim_text("tilt", x = 50, y = 30, w = 40, h = 20, rot = 30)
  ))
  expect_identical(shapes_of(items)$label, "")
  expect_equal(texts_of(items)$rot, 30)
})

test_that("text merges into an ellipse or a closed custom shape too", {
  items <- generic_items(list(
    prim_ellipse(50, 50, 40, 30),
    prim_text("oval", x = 35, y = 42, w = 30, h = 16),
    prim_poly(list(c(150, 10), c(190, 50), c(150, 90), c(110, 50))),
    prim_text("diam", x = 135, y = 42, w = 30, h = 16)
  ))
  expect_equal(nrow(texts_of(items)), 0L)
  expect_equal(items$label, c("oval", "diam"))
})

test_that("a two-row text merges with both rows in its label", {
  items <- generic_items(list(
    prim_rect(0, 0, 100, 60),
    prim_text(c("Decide", "now"), x = 25, y = 10, w = 50, h = 40)
  ))
  expect_identical(items$label, "Decide\nnow")
})

# -- Every fixture -------------------------------------------------------------

for (nm in all_fixture_names()) {
  test_that(paste0("build_scene() gives a well-formed items tibble: ", nm), {
    sc <- fixture_scene(nm)
    it <- sc$items
    expect_s3_class(sc, "mermaid_scene")
    expect_true(all(scene_cols %in% names(it)))
    expect_gt(nrow(it), 0L)
    expect_true(all(it$kind %in% c("group", "shape", "line", "text")))

    drawn <- it[it$kind != "group", ]
    for (col in c("x", "y", "w", "h")) {
      expect_false(anyNA(drawn[[col]]), info = paste(nm, col))
    }
    expect_true(all(drawn$w >= 0 & drawn$h >= 0))
    expect_false(anyNA(it$rot))
    expect_true(all(is.na(it$fill) | grepl("^[0-9A-Fa-f]{6}$", it$fill)))
    expect_true(all(is.na(it$stroke) | grepl("^[0-9A-Fa-f]{6}$", it$stroke)))
    expect_true(all(it$head %in% end_kinds))
    expect_true(all(it$tail %in% end_kinds))
    expect_false(anyNA(it$label))

    # Every parent is a group item of this scene.
    parents <- unique(stats::na.omit(it$parent))
    expect_true(all(parents %in% it$name[it$kind == "group"]))

    # No label is lost: every text prim's text is in some item's label.
    texts <- fixture_prims(nm, "text")
    for (p in texts) {
      txt <- paste(prim_row_texts(p), collapse = "\n")
      if (!nzchar(trimws(txt))) next
      expect_true(any(grepl(txt, it$label, fixed = TRUE)),
                  info = paste0(nm, ": text prim ", p$i, " '", txt, "'"))
    }
  })
}

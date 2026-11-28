# build_scene(), from prims to pieces: what each kind of prim becomes (a
# preset or custom geometry, its box, rotation and paint), the recognisers
# that give a path a preset, and what is dropped or left out as
# unsupported. Hand-built renders give exact boxes. How pieces fold into
# one shape is in test-fold.R, where texts go in test-text-placement.R.

# -- The scene object -----------------------------------------------------------

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

# -- Rects, ellipses and lines --------------------------------------------------

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

test_that("a rect is rounded only when both radii are positive", {
  expect_equal(generic_items(list(prim_rect(10, 20, 100, 40, rx = 8, ry = 0)))$prst, "rect")
  expect_equal(generic_items(list(prim_rect(10, 20, 100, 40, rx = 8, ry = 6)))$prst, "roundRect")
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

# -- Polygons and paths ---------------------------------------------------------

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

# -- Recognisers: lines, ellipses and rounded rectangles ------------------------

test_that("a path of one M and one L with no fill is a line item, flipped as drawn", {
  items <- generic_items(list(prim_path("M110,70 L10,20", fill = NULL, stroke = "#333333",
                                        marker_end = "arrowhead")))
  expect_equal(nrow(items), 1L)
  expect_equal(items$kind, "line")
  expect_equal(items$prst, "line")
  expect_near(box_of(items), c(10, 20, 100, 50), tol = 1e-9)
  expect_true(items$flip_h)
  expect_true(items$flip_v)
  expect_equal(items$tail, "triangle")     # the marker is kept
  expect_false(items$closed)
})

test_that("a two-point path with a fill, or a three-point one, is not a line", {
  filled <- generic_items(list(prim_path("M10,20 L110,70", fill = "#FF0000", stroke = "#333333")))
  expect_equal(filled$kind, "shape")
  bent <- generic_items(list(prim_path("M10,20 L60,40 L110,70", fill = NULL)))
  expect_equal(bent$kind, "shape")
})

# A circle or ellipse as four cubic quarter arcs (k = 0.5523).
cubic_ellipse <- function(cx, cy, rx, ry, k = 0.5523) {
  sprintf(paste("M%s,%s C%s,%s %s,%s %s,%s C%s,%s %s,%s %s,%s",
                "C%s,%s %s,%s %s,%s C%s,%s %s,%s %s,%s Z"),
          cx + rx, cy,
          cx + rx, cy + k * ry, cx + k * rx, cy + ry, cx, cy + ry,
          cx - k * rx, cy + ry, cx - rx, cy + k * ry, cx - rx, cy,
          cx - rx, cy - k * ry, cx - k * rx, cy - ry, cx, cy - ry,
          cx + k * rx, cy - ry, cx + rx, cy - k * ry, cx + rx, cy)
}

test_that("a closed path of curves on the inscribed ellipse is preset ellipse", {
  items <- generic_items(list(prim_path(cubic_ellipse(50, 50, 40, 40), bbox = c(10, 10, 80, 80))))
  expect_equal(items$prst, "ellipse")
  expect_near(box_of(items), c(10, 10, 80, 80), tol = 0.01)

  items <- generic_items(list(prim_path(cubic_ellipse(100, 50, 60, 30), bbox = c(40, 20, 120, 60))))
  expect_equal(items$prst, "ellipse")
})

test_that("a circle drawn with two arcs is preset ellipse", {
  items <- generic_items(list(prim_path("M10,50 A40,40 0 1 1 90,50 A40,40 0 1 1 10,50 Z",
                                        bbox = c(10, 10, 80, 80))))
  expect_equal(items$prst, "ellipse")
})

test_that("a superellipse (end points on the ellipse, midpoints outside) stays custom", {
  # control factor 0.9: the quarter midpoints bulge ~7 px past a 40 px circle
  items <- generic_items(list(prim_path(cubic_ellipse(50, 50, 40, 40, k = 0.9),
                                        bbox = c(10, 10, 80, 80))))
  expect_true(is.na(items$prst))
})

test_that("a closed path with fewer than four segments is not an ellipse", {
  # two cubics whose ends and midpoints all lie on the inscribed ellipse
  items <- generic_items(list(prim_path("M10,50 C10,-3.333 90,-3.333 90,50 C90,103.333 10,103.333 10,50 Z",
                                        bbox = c(10, 10, 80, 80))))
  expect_true(is.na(items$prst))
})

rounded_path <- function(r = c(8, 8, 8, 8), x = 20, y = 20, w = 100, h = 40) {
  # corners: top right, bottom right, bottom left, top left
  sprintf(paste("M%s,%s L%s,%s A%s,%s 0 0 1 %s,%s L%s,%s A%s,%s 0 0 1 %s,%s",
                "L%s,%s A%s,%s 0 0 1 %s,%s L%s,%s A%s,%s 0 0 1 %s,%s Z"),
          x + r[4], y, x + w - r[1], y, r[1], r[1], x + w, y + r[1],
          x + w, y + h - r[2], r[2], r[2], x + w - r[2], y + h,
          x + r[3], y + h, r[3], r[3], x, y + h - r[3],
          x, y + r[4], r[4], r[4], x + r[4], y)
}

test_that("a path with four equal round corners is preset roundRect with its radius", {
  items <- generic_items(list(prim_path(rounded_path(), bbox = c(20, 20, 100, 40))))
  expect_equal(items$prst, "roundRect")
  expect_equal(unname(items$adj[[1]]["adj"]), 20000)   # round(100000 * 8 / 40)
  expect_near(box_of(items), c(20, 20, 100, 40), tol = 0.01)
})

test_that("a stadium path is a roundRect at full rounding", {
  pill <- prim_path("M40,20 L100,20 A20,20 0 0 1 100,60 L40,60 A20,20 0 0 1 40,20 Z",
                    bbox = c(20, 20, 100, 40))
  items <- generic_items(list(pill))
  expect_equal(items$prst, "roundRect")
  expect_equal(unname(items$adj[[1]]["adj"]), 50000)
})

test_that("a rounded rectangle with unequal corner radii stays custom", {
  items <- generic_items(list(prim_path(rounded_path(c(8, 8, 8, 14)), bbox = c(20, 20, 100, 40))))
  expect_true(is.na(items$prst))
})

test_that("a rectangle with snipped (straight) corners stays custom", {
  # Every end point and every segment midpoint must lie on the rounded
  # rectangle (as in the ellipse rule). The chamfer's midpoint is inside the
  # corner circle's inner square; its end points are all on the bbox boundary.
  snip <- prim_path("M28,20 L112,20 L120,28 L120,52 L112,60 L28,60 L20,52 L20,28 Z")
  items <- generic_items(list(snip))
  expect_true(is.na(items$prst))
})

path_prst <- function(d, size) {
  generic_items(list(prim_path(d, bbox = c(0, 0, size, size))))$prst
}

test_that("a 3 px diamond and a 4 px pentagon stay custom geometry", {
  expect_true(is.na(path_prst("M1.5,0 L3,1.5 L1.5,3 L0,1.5 Z", 3)))
  r <- 2
  ang <- pi / 2 + (0:4) * 2 * pi / 5
  pts <- cbind(r + r * cos(ang), r - r * sin(ang))
  d <- paste0("M", paste(sprintf("%s,%s", pts[, 1], pts[, 2]), collapse = " L"), " Z")
  expect_true(is.na(path_prst(d, 4)))
})

test_that("a 3 px circle drawn as arcs or cubics is still an ellipse", {
  expect_equal(path_prst(cubic_ellipse(1.5, 1.5, 1.5, 1.5), 3), "ellipse")
  arcs <- paste("M3,1.5 A1.5,1.5 0 0 1 1.5,3 A1.5,1.5 0 0 1 0,1.5",
                "A1.5,1.5 0 0 1 1.5,0 A1.5,1.5 0 0 1 3,1.5 Z")
  expect_equal(path_prst(arcs, 3), "ellipse")
})

test_that("a 20 px diamond stays custom and a 100 px circle is an ellipse", {
  expect_true(is.na(path_prst("M10,0 L20,10 L10,20 L0,10 Z", 20)))
  expect_equal(path_prst(cubic_ellipse(50, 50, 50, 50), 100), "ellipse")
})

# -- A fill needs an area -------------------------------------------------------

test_that("a filled single-segment path of zero height has no fill and is a line connector", {
  p <- prim_path("M10,50 L110,50", fill = "#ECECFF", stroke = "#9370DB")
  items <- generic_items(list(p))
  expect_equal(nrow(items), 1L)
  expect_equal(items$kind, "line")
  expect_equal(items$prst, "line")
  expect_true(is.na(items$fill))
  expect_equal(items$fill_a, 0)
  expect_equal(items$stroke, "9370DB")
  expect_false(items$closed)
  expect_near(box_of(items), c(10, 50, 100, 0), tol = 1e-9)

  wsp <- all_wsp(generic_doc(list(p)))
  expect_length(wsp, 1L)
  expect_false(is_missing_node(xfirst(wsp[[1]], "./wps:cNvCnPr")))
  expect_true(is_missing_node(xfirst(wsp[[1]], "./wps:spPr/a:solidFill")))
})

test_that("a filled single-segment path of zero width likewise", {
  p <- prim_path("M50,10 L50,90", fill = "#FF0000", stroke = "#333333")
  items <- generic_items(list(p))
  expect_equal(nrow(items), 1L)
  expect_equal(items$kind, "line")
  expect_true(is.na(items$fill))
  expect_near(box_of(items), c(50, 10, 0, 80), tol = 1e-9)

  wsp <- all_wsp(generic_doc(list(p)))
  expect_false(is_missing_node(xfirst(wsp[[1]], "./wps:cNvCnPr")))
  expect_true(is_missing_node(xfirst(wsp[[1]], "./wps:spPr/a:solidFill")))
})

test_that("Mermaid's divider (a filled zero-height path of straight cubics) is an unfilled line", {
  # A path whose box is a segment is that segment, whatever its commands.
  # Mermaid's class, ER and requirement dividers are two straight cubic
  # subpaths (requirement's unfilled); the compartment rule needs them as
  # `line` items.
  items <- generic_items(list(mermaid_divider(10, 110, 50)))
  expect_equal(nrow(items), 1L)
  expect_equal(items$kind, "line")
  expect_true(is.na(items$fill))
  expect_near(box_of(items), c(10, 50, 100, 0), tol = 1e-6)

  unfilled <- generic_items(list(mermaid_divider(10, 110, 50, fill = NULL)))
  expect_equal(unfilled$kind, "line")
})

test_that("a path with a thin but non-zero box keeps its fill", {
  p <- prim_path("M10,50 L110,50 L110,50.5 L10,50.5", fill = "#FF0000", stroke = NULL)
  items <- generic_items(list(p))
  expect_equal(nrow(items), 1L)
  expect_equal(items$kind, "shape")
  expect_equal(items$fill, "FF0000")
  expect_near(items$h, 0.5, tol = 1e-9)
})

# A straight edge (a path of zero width or height, whatever its commands) is
# a line item; it keeps its direction, its marker end and the move of the end
# to the marker's tip.
test_that("a straight multi-point path with an end marker is a line that keeps its end", {
  id <- "flowchart-v2-pointEnd"
  run <- function(d) {
    p <- prim_path(d, stroke = "#333333", marker_end = id)
    generic_items(list(p), viewBox = c(0, 0, 200, 200),
                  markers = stats::setNames(list(marker_arrow_point()), id))
  }
  # the 8 px point: tip 5 content units beyond refX at scale 0.8 -> +4 px
  rtl <- run("M110,50 L60,50 L10,50")
  expect_equal(rtl$kind, "line")
  expect_equal(rtl$tail, "triangle")
  expect_true(rtl$flip_h)
  expect_near(item_ends(rtl)$start, c(110, 50), tol = 1e-9)
  expect_near(item_ends(rtl)$end, c(6, 50), tol = 1e-9)

  btt <- run("M50,110 L50,60 L50,10")
  expect_equal(btt$kind, "line")
  expect_equal(btt$tail, "triangle")
  expect_true(btt$flip_v)
  expect_near(item_ends(btt)$start, c(50, 110), tol = 1e-9)
  expect_near(item_ends(btt)$end, c(50, 6), tol = 1e-9)
})

# -- Texts ----------------------------------------------------------------------

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

# -- What is not drawn ----------------------------------------------------------

test_that("a line or open path of zero size with no marker is not drawn", {
  expect_equal(nrow(generic_items(list(prim_line(50, 50, 50, 50)))), 0L)
  expect_equal(nrow(generic_items(list(prim_line(50, 50, 50.005, 50.005)))), 0L)
  expect_equal(nrow(generic_items(list(prim_path("M50,50 L50,50", fill = NULL)))), 0L)
})

test_that("a stroke just over 0.01 px, or long in one direction, is drawn", {
  expect_equal(nrow(generic_items(list(prim_line(50, 50, 50.02, 50)))), 1L)
  expect_equal(nrow(generic_items(list(prim_line(50, 10, 50, 90)))), 1L)
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

test_that("an image without a picture and other prims are not drawn but listed as unsupported", {
  sc <- build_scene(generic_render(list(
    prim_rect(0, 0, 50, 50),
    prim_image(60, 0, 30, 30, png = NULL),
    prim_other("use", 100, 0, 20, 20)
  )))
  expect_equal(nrow(sc$items), 1L)
  expect_true(all(c("image", "use") %in% sc$unsupported))
})

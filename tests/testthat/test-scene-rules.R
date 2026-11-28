# Step 2 scene rules (contract-step2.md, section B): scope (B1), pairing by
# box (B2), recognisers (B3), zero-size strokes (B4), text stacks (B5),
# placement per axis (B6), scope groups (B7) and four-line rectangles (B8).
# Hand-built renders only; the fixture assertions are in
# test-fixtures-step2.R.

shapes_in <- function(items) items[items$kind == "shape", ]
groups_in <- function(items) items[items$kind == "group", ]
box_of <- function(it) c(it$x, it$y, it$w, it$h)

# The outline of a box drawn as Mermaid 11 draws a stroke: straight cubic
# segments and no Z (so it is not a preset rect by itself).
curvy_rect <- function(x, y, w, h, ...) {
  r <- x + w
  b <- y + h
  d <- sprintf(paste("M%s,%s C%s,%s %s,%s %s,%s C%s,%s %s,%s %s,%s",
                     "C%s,%s %s,%s %s,%s C%s,%s %s,%s %s,%s"),
               x, y, x + w / 3, y, x + 2 * w / 3, y, r, y,
               r, y + h / 3, r, y + 2 * h / 3, r, b,
               r - w / 3, b, r - 2 * w / 3, b, x, b,
               x, b - h / 3, x, b - 2 * h / 3, x, y)
  prim_path(d, bbox = c(x, y, w, h), ...)
}

# A fill-only box path, as Mermaid draws the fill under the outline.
fill_rect_path <- function(x, y, w, h, fill = "#ECECFF", stroke = NULL, ...) {
  prim_path(sprintf("M%s,%s L%s,%s L%s,%s L%s,%s", x, y, x + w, y, x + w, y + h, x, y + h),
            bbox = c(x, y, w, h), fill = fill, stroke = stroke, ...)
}

# The text body of the one item that has one.
the_text <- function(items) {
  k <- which(!purrr::map_lgl(items$text, is.null))
  testthat::expect_length(k, 1L)
  if (length(k) == 0L) return(list())
  items$text[[k[[1]]]]
}

# -- B1/B2: same box in the same scope is one shape ----------------------------

test_that("a fill-only path and a stroked path with the same box in one group become one shape", {
  items <- generic_items(in_g(list(
    fill_rect_path(10, 10, 100, 50, fill = "#ECECFF"),
    curvy_rect(10, 10, 100, 50, fill = NULL, stroke = "#9370DB", stroke_w = 1.3)
  ), n = 5))
  s <- shapes_in(items)
  expect_equal(nrow(s), 1L)
  expect_equal(s$prst, "rect")           # the preset of the first piece that has one
  expect_equal(s$fill, "ECECFF")
  expect_equal(s$stroke, "9370DB")
  expect_equal(s$stroke_w, 1.3)
  expect_near(box_of(s), c(10, 10, 100, 50), tol = 1e-6)
})

test_that("two pieces that both fill and stroke with the same box become one shape", {
  # Mermaid's class boxes: a straight path and a rough path, both painted.
  items <- generic_items(in_g(list(
    fill_rect_path(10, 10, 100, 50, fill = "#AAAAAA", stroke = "#111111"),
    curvy_rect(10, 10, 100, 50, fill = "#BBBBBB", stroke = "#0000FF", stroke_w = 2)
  ), n = 5))
  s <- shapes_in(items)
  expect_equal(nrow(s), 1L)
  expect_equal(s$fill, "AAAAAA")         # first piece that has a fill
  expect_equal(s$stroke, "0000FF")       # last piece that has a stroke
  expect_equal(s$stroke_w, 2)
})

test_that("three pieces with one box fold into one: first fill, last stroke and its dash", {
  items <- generic_items(in_g(list(
    prim_rect(10, 10, 100, 50, fill = "#111111", fill_a = 0.5, stroke = "#222222"),
    prim_rect(10, 10, 100, 50, fill = "#333333", stroke = NULL),
    prim_rect(10, 10, 100, 50, fill = NULL, stroke = "#444444", stroke_w = 3,
              stroke_a = 0.25, dash = c(2, 2))
  ), n = 5))
  s <- shapes_in(items)
  expect_equal(nrow(s), 1L)
  expect_equal(s$fill, "111111")
  expect_equal(s$fill_a, 0.5)
  expect_equal(s$stroke, "444444")
  expect_equal(s$stroke_w, 3)
  expect_equal(s$stroke_a, 0.25)
  expect_equal(s$dash[[1]], c(2, 2))
})

test_that("when the first piece has no fill, the fill comes from a later one", {
  items <- generic_items(in_g(list(
    prim_rect(10, 10, 100, 50, fill = NULL, stroke = "#222222", id = "outline"),
    prim_rect(10, 10, 100, 50, fill = "#333333", stroke = NULL, id = "under")
  ), n = 5))
  s <- shapes_in(items)
  expect_equal(nrow(s), 1L)
  expect_equal(s$fill, "333333")
  expect_equal(s$stroke, "222222")
})

test_that("the folded shape takes the name and paint position of the first piece", {
  items <- generic_items(in_g(list(
    prim_rect(10, 10, 100, 50, fill = NULL, stroke = "#222222", id = "first"),
    prim_rect(130, 10, 40, 40, fill = "#00AA00", id = "between"),
    prim_rect(10, 10, 100, 50, fill = "#333333", stroke = NULL, id = "last")
  ), n = 5))
  s <- shapes_in(items)
  expect_equal(nrow(s), 2L)
  expect_equal(s$name, c("mermaid:shape:first", "mermaid:shape:between"))
  expect_equal(s$fill[[1]], "333333")
})

test_that("boxes within max(0.5 px, 1 % of the larger side) pair; just outside they do not", {
  # 200 x 100: tolerance 2 px
  pair <- function(dx, w = 200, h = 100) {
    items <- generic_items(in_g(list(
      prim_rect(10, 10, w, h, fill = "#ECECFF", stroke = NULL),
      prim_rect(10 + dx, 10, w, h, fill = NULL, stroke = "#9370DB")
    ), n = 5), viewBox = c(0, 0, 300, 200))
    nrow(shapes_in(items))
  }
  expect_equal(pair(1.9), 1L)
  expect_equal(pair(2.1), 2L)
  # 20 x 10: 1 % is 0.2, so the 0.5 px floor applies
  expect_equal(pair(0.45, 20, 10), 1L)
  expect_equal(pair(0.6, 20, 10), 2L)
})

test_that("pieces with the same box in different groups do not pair", {
  a <- in_g(list(prim_rect(10, 10, 100, 50, fill = "#ECECFF", stroke = NULL)), n = 5)
  b <- in_g(list(prim_rect(10, 10, 100, 50, fill = NULL, stroke = "#9370DB")), n = 6)
  expect_equal(nrow(shapes_in(generic_items(c(a, b)))), 2L)
  # a child of the root and a child of a group are in different scopes too
  root <- list(prim_rect(10, 10, 100, 50, fill = NULL, stroke = "#9370DB"))
  expect_equal(nrow(shapes_in(generic_items(c(a, root)))), 2L)
})

test_that("hand-built pieces without n share the root scope and pair as before", {
  items <- generic_items(list(
    fill_rect_path(10, 10, 100, 50),
    curvy_rect(10, 10, 100, 50, fill = NULL, stroke = "#9370DB")
  ))
  expect_equal(nrow(shapes_in(items)), 1L)
})

test_that("pieces with the same box but different rotations do not pair", {
  # x' = -y + 100, y' = x + 50 turns 60 x 20 by 90 degrees: box 60, 70, 60, 20.
  items <- generic_items(in_g(list(
    prim_rect(0, 0, 60, 20, m = c(0, 1, -1, 0, 100, 50), fill = "#ECECFF", stroke = NULL),
    prim_rect(60, 70, 60, 20, fill = NULL, stroke = "#9370DB")
  ), n = 5), viewBox = c(0, 0, 200, 200))
  expect_equal(nrow(shapes_in(items)), 2L)
})

test_that("a piece with a marker never pairs", {
  items <- generic_items(in_g(list(
    prim_rect(10, 10, 100, 50, fill = "#ECECFF", stroke = NULL),
    prim_path("M10,10 L110,10 L110,60 L10,60", fill = NULL, stroke = "#333333",
              marker_end = "arrowhead")
  ), n = 5))
  expect_equal(sum(items$kind != "group"), 2L)
})

test_that("two custom outlines fold into one closed path from the first filled piece", {
  diamond_fill <- prim_poly(list(c(60, 10), c(110, 35), c(60, 60), c(10, 35)),
                            fill = "#ECECFF", stroke = NULL)
  diamond_line <- prim_path("M60,10 L110,35 L60,60 L10,35", fill = NULL, stroke = "#9370DB")
  items <- generic_items(in_g(list(diamond_line, diamond_fill), n = 5))
  s <- shapes_in(items)
  expect_equal(nrow(s), 1L)
  expect_true(is.na(s$prst))
  expect_true(s$closed)
  path <- s$path[[1]]
  expect_equal(path$cmd[[nrow(path)]], "Z")
  expect_equal(sum(path$cmd == "M"), sum(path$cmd == "Z"))   # every subpath closed
  expect_equal(s$fill, "ECECFF")
  expect_equal(s$stroke, "9370DB")
})

test_that("with no filled piece the geometry is the first piece's path, closed", {
  first <- prim_path("M60,10 L110,35 L60,60 L10,35", fill = NULL, stroke = "#111111")
  second <- prim_path("M10,35 L60,10 L110,35 L60,60 Z", fill = NULL, stroke = "#222222")
  items <- generic_items(in_g(list(first, second), n = 5))
  s <- shapes_in(items)
  expect_equal(nrow(s), 1L)
  path <- s$path[[1]]
  pts <- path[path$cmd != "Z", ]
  expect_equal(c(pts$x[[1]], pts$y[[1]]), c(60, 10))
  expect_equal(path$cmd[[nrow(path)]], "Z")
  expect_true(s$closed)
  expect_equal(s$stroke, "222222")
})

# -- B3: recognisers -------------------------------------------------------------

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
  # SPEC? B3 says "every point"; this reads it as end points AND segment
  # midpoints (as in the ellipse rule). The chamfer's midpoint is inside the
  # corner circle's inner square; its end points are all on the bbox boundary.
  snip <- prim_path("M28,20 L112,20 L120,28 L120,52 L112,60 L28,60 L20,52 L20,28 Z")
  items <- generic_items(list(snip))
  expect_true(is.na(items$prst))
})

# -- B4: zero-size strokes -------------------------------------------------------

test_that("a line or open path of zero size with no marker is not drawn", {
  expect_equal(nrow(generic_items(list(prim_line(50, 50, 50, 50)))), 0L)
  expect_equal(nrow(generic_items(list(prim_line(50, 50, 50.005, 50.005)))), 0L)
  expect_equal(nrow(generic_items(list(prim_path("M50,50 L50,50", fill = NULL)))), 0L)
})

test_that("a stroke just over 0.01 px, or long in one direction, is drawn", {
  expect_equal(nrow(generic_items(list(prim_line(50, 50, 50.02, 50)))), 1L)
  expect_equal(nrow(generic_items(list(prim_line(50, 10, 50, 90)))), 1L)
})

# -- B5: several texts in one shape ----------------------------------------------

test_that("centred texts stacked in a shape become one text body, rows in y order", {
  items <- generic_items(list(
    prim_rect(0, 0, 200, 100),
    prim_text("Sub", x = 80, y = 42, w = 40, h = 16),        # painted first
    prim_text("Title", x = 70, y = 20, w = 60, h = 20)
  ))
  expect_equal(nrow(items), 1L)
  expect_identical(items$label, "Title\nSub")
  tx <- the_text(items)
  expect_length(tx$rows, 2L)
  expect_near(tx$box, c(70, 20, 60, 38), tol = 1e-9)
  expect_equal(tx$align, "center")
})

test_that("left-aligned and right-aligned stacks merge too", {
  left <- generic_items(list(
    prim_rect(0, 0, 200, 100),
    prim_text("name", x = 10, y = 10, w = 40, h = 20, anchor = "start"),
    prim_text("a longer row", x = 10, y = 34, w = 90, h = 20, anchor = "start")
  ))
  expect_equal(nrow(left), 1L)
  expect_identical(left$label, "name\na longer row")

  right <- generic_items(list(
    prim_rect(0, 0, 200, 100),
    prim_text("x", x = 170, y = 10, w = 20, h = 20, anchor = "end"),
    prim_text("wider", x = 140, y = 34, w = 50, h = 20, anchor = "end")
  ))
  expect_equal(nrow(right), 1L)
})

test_that("the gap limit is 0.6 x the taller neighbouring row", {
  stacked <- function(gap, h1 = 20, h2 = 20) {
    items <- generic_items(list(
      prim_rect(0, 0, 200, 100),
      prim_text("a", x = 80, y = 10, w = 40, h = h1),
      prim_text("b", x = 80, y = 10 + h1 + gap, w = 40, h = h2)
    ))
    nrow(items) == 1L
  }
  expect_true(stacked(11.9))      # 0.6 x 20 = 12
  expect_false(stacked(12.1))
  expect_true(stacked(15, h1 = 30, h2 = 10))   # 0.6 x 30 = 18, not 0.6 x 10
})

test_that("texts with mixed horizontal alignment stay free", {
  items <- generic_items(list(
    prim_rect(0, 0, 200, 100),
    prim_text("centred", x = 70, y = 10, w = 60, h = 20),
    prim_text("left", x = 10, y = 34, w = 40, h = 20, anchor = "start")
  ))
  expect_identical(items$label[items$kind == "shape"], "")
  expect_equal(sum(items$kind == "text"), 2L)
})

test_that("texts overlapping vertically by more than 25 % of the smaller height stay free; by less they stack", {
  # round 2: H3 raised the limit from 1 px to 25 % of the smaller height (5 px here)
  overlap <- function(o) {
    generic_items(list(
      prim_rect(0, 0, 200, 100),
      prim_text("a", x = 80, y = 10, w = 40, h = 20),
      prim_text("b", x = 80, y = 30 - o, w = 40, h = 20)
    ))
  }
  expect_equal(sum(overlap(5.5)$kind == "text"), 2L)
  expect_equal(nrow(overlap(0.5)), 1L)
})

# -- B6: placement per axis ------------------------------------------------------

# The text body of a 200 x 100 rect (tolerances: 4 px across, 3 px down)
# holding one text at box c(x, y, w, h).
placed <- function(x, y, w = 40, h = 20, anchor = "middle", shape = c(0, 0, 200, 100),
                   viewBox = c(0, 0, 400, 200)) {
  items <- generic_items(list(
    prim_rect(shape[1], shape[2], shape[3], shape[4]),
    # ruling: SVG text, which I5 leaves to B6 (a one-row HTML label takes
    # the nearer side)
    prim_text("t", x = x, y = y, w = w, h = h, anchor = anchor, html = FALSE)
  ), viewBox = viewBox)
  testthat::expect_equal(nrow(items), 1L)
  items$text[[1]]
}

test_that("a text centred on both axes: placement centre, center / center, zero insets", {
  tx <- placed(80, 40)
  expect_equal(tx$placement, "centre")
  expect_equal(tx$align, "center")
  expect_equal(tx$valign, "center")
  expect_equal(tx$insets, c(0, 0, 0, 0))
})

test_that("horizontal centring tolerance is 0.03 x the width, between 1 and 4 px", {
  # 200 wide: 4 px
  expect_equal(placed(80 + 3.9, 40)$align, "center")
  expect_equal(placed(80 + 3.9, 40)$insets[c(1, 3)], c(0, 0))
  off <- placed(80 + 4.1, 40)
  expect_equal(off$placement, "inset")
  # anchor middle, right of centre: l = 2 x 4.1, r = 0
  expect_near(off$insets[c(1, 3)], c(8.2, 0), tol = 1e-9)
  expect_equal(off$align, "center")
  # 20 wide: 0.6 -> 1 px
  expect_equal(placed(5 + 0.9, 40, w = 10, shape = c(0, 0, 20, 100))$align, "center")
  expect_near(placed(5 + 1.1, 40, w = 10, shape = c(0, 0, 20, 100))$insets[[1]], 2.2,
              tol = 1e-9)
  # 400 wide: 12 -> 4 px
  expect_equal(placed(180 + 4.5, 40, shape = c(0, 0, 400, 100))$placement, "inset")
})

test_that("off centre horizontally: one inset, on the side the anchor names", {
  start <- placed(10, 40, w = 60, anchor = "start")
  expect_equal(start$align, "left")
  expect_near(start$insets[c(1, 3)], c(10, 0), tol = 1e-9)

  end <- placed(130, 40, w = 60, anchor = "end")
  expect_equal(end$align, "right")
  expect_near(end$insets[c(1, 3)], c(0, 10), tol = 1e-9)

  # anchor middle, left of centre: r = 2 x 10
  mid <- placed(70, 40, w = 40, anchor = "middle")
  expect_equal(mid$align, "center")
  expect_near(mid$insets[c(1, 3)], c(0, 20), tol = 1e-9)
})

test_that("vertical placement: centre within 0.03 x the height, else top or bottom with one inset", {
  # 100 high: 3 px
  expect_equal(placed(80, 40 + 2.9)$valign, "center")
  expect_equal(placed(80, 40 + 2.9)$insets[c(2, 4)], c(0, 0))
  low <- placed(80, 40 + 3.1)
  expect_equal(low$valign, "bottom")
  expect_near(low$insets[c(2, 4)], c(0, 100 - 63.1), tol = 1e-9)

  top <- placed(80, 10)
  expect_equal(top$valign, "top")
  expect_near(top$insets[c(2, 4)], c(10, 0), tol = 1e-9)
  expect_equal(top$align, "center")        # still centred across
  expect_equal(top$placement, "inset")

  bottom <- placed(80, 70)
  expect_equal(bottom$valign, "bottom")
  expect_near(bottom$insets[c(2, 4)], c(0, 10), tol = 1e-9)
})

test_that("a negative inset is clamped to 0", {
  # 50 of 60 px inside, anchor end: r would be 100 - 110 = -10
  tx <- placed(50, 10, w = 60, anchor = "end", shape = c(0, 0, 100, 50))
  expect_equal(tx$align, "right")
  expect_equal(tx$insets[c(1, 3)], c(0, 0))
  expect_true(all(tx$insets >= 0))
})

test_that("a free text has valign center and placement free", {
  items <- generic_items(list(prim_text("free", x = 10, y = 10, w = 40, h = 20)))
  expect_equal(items$text[[1]]$placement, "free")
  expect_equal(items$text[[1]]$valign, "center")
})

# -- B7: scope groups --------------------------------------------------------------

legend <- function(n, y = 10, cls = "legend", id = "") {
  in_g(list(prim_rect(10, y, 18, 18, fill = "#ECECFF"),
            prim_text("Dogs", x = 32, y = y, w = 40, h = 18, anchor = "start")),
       n = n, cls = cls, id = id)
}

# round 2: H4b measures "the whole diagram" against the union of all items,
# so a scope tested alone would be the whole diagram; this puts something
# else on the canvas.
outside <- function() list(prim_rect(150, 60, 40, 30, fill = "#00AA00"))

test_that("loose items sharing a scope become a group named after its first class", {
  items <- generic_items(c(legend(5, cls = "legend item"), outside()))   # round 2: H4b
  g <- groups_in(items)
  expect_equal(nrow(g), 1L)
  expect_equal(g$name, "mermaid:group:legend")
  kids <- items[items$kind != "group" & !items$fill %in% "00AA00", ]
  expect_equal(nrow(kids), 2L)
  expect_true(all(kids$parent == "mermaid:group:legend"))
  expect_near(box_of(g), c(10, 10, 62, 18), tol = 1e-9)
  expect_equal(which(items$kind == "group"), 1L)     # the group row comes first
})

test_that("a scope group is named after the element id first, and 'g' when unnamed", {
  # round 2: H4b, with something outside the scope
  expect_equal(groups_in(generic_items(c(legend(5, id = "key", cls = "legend"), outside())))$name,
               "mermaid:group:key")
  expect_equal(groups_in(generic_items(c(legend(5, cls = ""), outside())))$name, "mermaid:group:g")
})

test_that("a scope that would hold one item makes no group", {
  # the text merges into its rect: one item
  items <- generic_items(in_g(list(prim_rect(10, 10, 80, 40),
                                   prim_text("in", x = 30, y = 20, w = 40, h = 20)), n = 5))
  expect_equal(nrow(groups_in(items)), 0L)
  expect_equal(nrow(items), 1L)
  expect_true(is.na(items$parent))
})

test_that("a scope with more than 40 drawn items is not grouped; 40 is", {
  dots <- function(k) {
    # round 2: H4b, a rect below the dots so that they are not the whole diagram
    c(in_g(purrr::map(seq_len(k), function(i) prim_rect(5 * i, 10, 2, 2)), n = 5, cls = "dots"),
      list(prim_rect(5, 80, 10, 10)))
  }
  expect_equal(nrow(groups_in(generic_items(dots(40), viewBox = c(0, 0, 400, 100)))), 1L)
  expect_equal(nrow(groups_in(generic_items(dots(41), viewBox = c(0, 0, 400, 100)))), 0L)
})

test_that("a scope covering 90 % of the diagram both ways is the whole diagram: no group", {
  # round 2: H4b measures against the union of all items, not the viewBox
  whole <- in_g(list(prim_rect(0, 0, 190, 95), prim_rect(50, 50, 10, 10)), n = 5, cls = "all")
  expect_equal(nrow(groups_in(generic_items(whole))), 0L)
  # all of the width but 80 % of the height (a rect below it): grouped
  wide <- in_g(list(prim_rect(0, 0, 190, 80), prim_rect(50, 50, 10, 10)), n = 5, cls = "all")
  expect_equal(nrow(groups_in(generic_items(c(wide, list(prim_rect(0, 90, 10, 10)))))), 1L)
})

test_that("groups nest when their scopes nest", {
  inner <- in_g(list(prim_rect(10, 10, 10, 10), prim_rect(30, 10, 10, 10)), n = 3, cls = "inner")
  outer <- in_g(c(inner, list(prim_rect(10, 40, 10, 10), prim_rect(30, 40, 10, 10))),
                n = 2, cls = "outer")
  items <- generic_items(c(outer, outside()))     # round 2: H4b
  g <- groups_in(items)
  expect_setequal(g$name, c("mermaid:group:outer", "mermaid:group:inner"))
  expect_true(is.na(item_named(items, "mermaid:group:outer")$parent))
  expect_equal(item_named(items, "mermaid:group:inner")$parent, "mermaid:group:outer")
  drawn <- items[items$kind != "group", ]
  expect_equal(sum(drawn$parent %in% "mermaid:group:inner"), 2L)
  expect_equal(sum(drawn$parent %in% "mermaid:group:outer"), 2L)
})

test_that("sibling scopes with the same class get unique names in document order", {
  # SPEC? "#<k> suffixes in document order": this accepts the first one plain
  # or suffixed, and checks only uniqueness, the pattern and the order.
  items <- generic_items(c(legend(5, y = 10), legend(8, y = 40), legend(11, y = 70)))
  g <- groups_in(items)
  expect_equal(nrow(g), 3L)
  expect_false(anyDuplicated(g$name) > 0L)
  expect_true(all(grepl("^mermaid:group:legend(#[0-9]+)?$", g$name)))
  expect_equal(g$y, c(10, 40, 70))          # in document order
  k <- suppressWarnings(as.integer(sub(".*#", "", g$name)))
  k[is.na(k)] <- 0L
  expect_true(all(diff(k) > 0))
})

test_that("items of a unit are not regrouped by scope", {
  anc <- list(anc_g(1, cls = "root"), anc_g(2, cls = "nodes"),
              anc_g(3, cls = "default", unit = unit_node("X")))
  prims <- list(
    prim_ellipse(60, 50, 40, anc = anc),
    prim_ellipse(60, 50, 35, anc = anc),
    prim_text("Ring", x = 45, y = 42, w = 30, h = 16, anc = c(anc, list(anc_g(4, cls = "label"))))
  )
  db <- mini_db(list(db_node("X", shape = "doublecircle")))
  items <- build_scene(mini_render(prims, db = db))$items
  expect_false(any(grepl("^mermaid:group:", items$name)))
  expect_equal(item_named(items, "mermaid:node:X")$kind, "group")
})

test_that("a text chooses a shape of its own scope before a smaller one elsewhere", {
  items <- generic_items(c(
    list(prim_rect(40, 30, 60, 30, fill = "#AA0000")),                    # root
    in_g(list(prim_rect(10, 10, 180, 80, fill = "#00AA00"),
              prim_text("T", x = 50, y = 35, w = 40, h = 20)), n = 5)
  ))
  expect_identical(items$label[items$fill %in% "00AA00"], "T")
  expect_identical(items$label[items$fill %in% "AA0000"], "")
})

test_that("a text with no shape in its scope chooses among all shapes", {
  items <- generic_items(c(
    list(prim_rect(10, 10, 180, 80, fill = "#AA0000")),
    in_g(list(prim_text("T", x = 50, y = 35, w = 40, h = 20)), n = 5)
  ))
  expect_equal(nrow(items), 1L)
  expect_identical(items$label, "T")
})

# -- B8: four lines that draw a rectangle ----------------------------------------

frame_lines <- function(x = 10, y = 10, w = 100, h = 50, stroke = c("#DACEF3", "#DACEF3",
                        "#DACEF3", "#DACEF3"), dash = list(NULL, NULL, NULL, NULL),
                        marker = list(NULL, NULL, NULL, NULL), jitter = 0) {
  # drawn as Mermaid's loop frame: top, right, bottom (left to right), left
  ends <- list(c(x, y, x + w, y), c(x + w, y, x + w + jitter, y + h),
               c(x, y + h, x + w, y + h), c(x, y, x, y + h))
  purrr::pmap(list(ends, stroke, dash, marker), function(e, s, d, mk) {
    prim_line(e[1], e[2], e[3], e[4], stroke = s, stroke_w = 2, dash = d, marker_end = mk)
  })
}

test_that("four lines that are the sides of a box become one unfilled rect", {
  items <- generic_items(frame_lines())
  expect_equal(nrow(items), 1L)
  expect_equal(items$kind, "shape")
  expect_equal(items$prst, "rect")
  expect_near(box_of(items), c(10, 10, 100, 50), tol = 1e-9)
  expect_true(is.na(items$fill))
  expect_equal(items$stroke, "DACEF3")
  expect_equal(items$stroke_w, 2)
})

test_that("the four-line rect keeps the lines' dash and the first line's paint position", {
  dashed <- generic_items(frame_lines(dash = rep(list(c(3, 3)), 4)))
  expect_equal(dashed$dash[[1]], c(3, 3))

  lines <- frame_lines()
  items <- generic_items(c(lines[1], list(prim_rect(150, 10, 30, 30, fill = "#00AA00")),
                           lines[2:4]))
  expect_equal(items$prst, c("rect", "rect"))
  expect_true(is.na(items$fill[[1]]))        # the frame first, where its first line was
  expect_equal(items$fill[[2]], "00AA00")
})

test_that("ends meeting within 1 px still make the rect; 1.5 px apart they do not", {
  expect_equal(nrow(generic_items(frame_lines(jitter = 0.9))), 1L)
  expect_equal(nrow(generic_items(frame_lines(jitter = 1.5))), 4L)
})

test_that("four sides with different paint, dash or a marker stay lines", {
  expect_equal(nrow(generic_items(frame_lines(stroke = c("#DACEF3", "#DACEF3", "#FF0000",
                                                         "#DACEF3")))), 4L)
  expect_equal(nrow(generic_items(frame_lines(dash = list(NULL, c(3, 3), NULL, NULL)))), 4L)
  expect_equal(nrow(generic_items(frame_lines(marker = list(NULL, NULL, "arrowhead", NULL)))), 4L)
})

test_that("four sides split between two scopes stay lines", {
  l <- frame_lines()
  items <- generic_items(c(in_g(l[1:2], n = 5), in_g(l[3:4], n = 6)))
  expect_equal(sum(items$kind == "line"), 4L)
})

test_that("a fifth line across the frame stays a line beside the rect", {
  divider <- prim_line(10, 35, 110, 35, stroke = "#DACEF3", stroke_w = 2, dash = c(2, 2))
  items <- generic_items(c(frame_lines(), list(divider)))
  expect_equal(sum(items$kind == "shape" & items$prst %in% "rect"), 1L)
  expect_equal(sum(items$kind == "line"), 1L)
})

# -- rulings on B2 and B5 ------------------------------------------------------

test_that("an outer circle and a sector with the same box stay two shapes", {
  # ruling: a preset pairs with a custom path only when the path's end points
  # and midpoints lie on the preset's figure (the pie's outer circle and its
  # largest slice share a box, not a figure).
  sector <- prim_path("M200,15 A185,185 0 1 1 15,200 L200,200 Z", bbox = c(15, 15, 370, 370),
                      fill = "#ECECFF", stroke = "#000000", stroke_w = 2)
  ring <- prim_ellipse(200, 200, 186, fill = NULL, stroke = "#000000", stroke_w = 2)
  items <- generic_items(in_g(list(ring, sector), n = 5), viewBox = c(0, 0, 400, 400))
  s <- shapes_in(items)
  expect_equal(nrow(s), 2L)
  expect_equal(s$prst[[1]], "ellipse")
  expect_true(is.na(s$prst[[2]]))
  expect_equal(s$fill[[2]], "ECECFF")
})

test_that("a stack in the corner of an ellipse's box, outside the ellipse, stays free", {
  # ruling: a stack merges only when its union box lies inside the
  # container's figure (for an ellipse, all four corners inside it).
  corner <- generic_items(list(
    prim_ellipse(100, 100, 90),
    prim_text("Alice", x = 150, y = 15, w = 35, h = 18, anchor = "start"),
    prim_text("Bob", x = 150, y = 35, w = 25, h = 18, anchor = "start")
  ))
  expect_equal(sum(corner$kind == "text"), 2L)
  expect_identical(corner$label[corner$kind == "shape"], "")

  centre <- generic_items(list(
    prim_ellipse(100, 100, 90),
    prim_text("Alice", x = 80, y = 80, w = 40, h = 18),
    prim_text("Bob", x = 85, y = 100, w = 30, h = 18)
  ))
  expect_equal(nrow(centre), 1L)
  expect_identical(centre$label, "Alice\nBob")
})

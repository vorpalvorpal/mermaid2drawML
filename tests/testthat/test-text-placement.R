# Texts in shapes: which shape a text joins, how texts stacked in one shape
# become one text body, and where the body sits in its shape (alignment,
# vertical anchor and insets on each axis).

# -- Which shape a text joins ---------------------------------------------------

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

# -- Stacks of texts ------------------------------------------------------------

# The text body of the one item that has one.
the_text <- function(items) {
  k <- which(!purrr::map_lgl(items$text, is.null))
  testthat::expect_length(k, 1L)
  if (length(k) == 0L) return(list())
  items$text[[k[[1]]]]
}

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

# Texts stack when they overlap vertically by at most 25 % of the smaller
# height.
test_that("texts overlapping by up to 25 % of the smaller height stack; by more they stay free", {
  overlap <- function(o, h1 = 20, h2 = 20) {
    generic_items(list(
      prim_rect(0, 0, 200, 100),
      prim_text("a", x = 80, y = 10, w = 40, h = h1),
      prim_text("b", x = 85, y = 10 + h1 - o, w = 30, h = h2)
    ))
  }
  stacked <- overlap(4)                     # 20 % of 20
  expect_equal(nrow(stacked), 1L)
  expect_identical(stacked$label, "a\nb")
  expect_equal(sum(overlap(6)$kind == "text"), 2L)   # 30 % of 20
  # the smaller of the two heights: 25 % of 10 is 2.5
  expect_equal(nrow(overlap(2, h1 = 40, h2 = 10)), 1L)    # 20 % of 10
  expect_equal(sum(overlap(3, h1 = 40, h2 = 10)$kind == "text"), 2L)   # 30 % of 10
  # close to the limit: 0.5 px and 5.5 px (27.5 %) of 20
  expect_equal(nrow(overlap(0.5)), 1L)
  expect_equal(sum(overlap(5.5)$kind == "text"), 2L)
})

test_that("a stack in the corner of an ellipse's box, outside the ellipse, stays free", {
  # A stack merges only when its union box lies inside the
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

# -- Placement on each axis -----------------------------------------------------

# The text body of a 200 x 100 rect (tolerances: 4 px across, 3 px down)
# holding one text at box c(x, y, w, h).
placed <- function(x, y, w = 40, h = 20, anchor = "middle", shape = c(0, 0, 200, 100),
                   viewBox = c(0, 0, 400, 200)) {
  items <- generic_items(list(
    prim_rect(shape[1], shape[2], shape[3], shape[4]),
    # SVG text, placed by its anchor (a one-row HTML label takes the nearer
    # side instead)
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

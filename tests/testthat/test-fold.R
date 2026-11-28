# fold_pieces(): the rules that make several pieces one shape. Stacked HTML
# boxes collapse into one, pieces that draw one box in one scope pair, four
# lines that draw a rectangle become one, and a box crossed by dividers of
# its own stroke is split into stacked compartments.

# -- HTML boxes collapse --------------------------------------------------------

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

# -- Pieces with one box pair into one shape ------------------------------------

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

# A fill-only box path, as Mermaid draws the fill under the outline.
fill_rect_path <- function(x, y, w, h, fill = "#ECECFF", stroke = NULL, ...) {
  prim_path(sprintf("M%s,%s L%s,%s L%s,%s L%s,%s", x, y, x + w, y, x + w, y + h, x, y + h),
            bbox = c(x, y, w, h), fill = fill, stroke = stroke, ...)
}

test_that("a fill-only path and a stroked path with the same box in one group become one shape", {
  items <- generic_items(in_g(list(
    fill_rect_path(10, 10, 100, 50, fill = "#ECECFF"),
    rough_rect(10, 10, 100, 50, fill = NULL, stroke = "#9370DB", stroke_w = 1.3)
  ), n = 5))
  s <- shapes_of(items)
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
    rough_rect(10, 10, 100, 50, fill = "#BBBBBB", stroke = "#0000FF", stroke_w = 2)
  ), n = 5))
  s <- shapes_of(items)
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
  s <- shapes_of(items)
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
  s <- shapes_of(items)
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
  s <- shapes_of(items)
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
    nrow(shapes_of(items))
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
  expect_equal(nrow(shapes_of(generic_items(c(a, b)))), 2L)
  # a child of the root and a child of a group are in different scopes too
  root <- list(prim_rect(10, 10, 100, 50, fill = NULL, stroke = "#9370DB"))
  expect_equal(nrow(shapes_of(generic_items(c(a, root)))), 2L)
})

test_that("hand-built pieces without n share the root scope and pair as before", {
  items <- generic_items(list(
    fill_rect_path(10, 10, 100, 50),
    rough_rect(10, 10, 100, 50, fill = NULL, stroke = "#9370DB")
  ))
  expect_equal(nrow(shapes_of(items)), 1L)
})

test_that("pieces with the same box but different rotations do not pair", {
  # x' = -y + 100, y' = x + 50 turns 60 x 20 by 90 degrees: box 60, 70, 60, 20.
  items <- generic_items(in_g(list(
    prim_rect(0, 0, 60, 20, m = c(0, 1, -1, 0, 100, 50), fill = "#ECECFF", stroke = NULL),
    prim_rect(60, 70, 60, 20, fill = NULL, stroke = "#9370DB")
  ), n = 5), viewBox = c(0, 0, 200, 200))
  expect_equal(nrow(shapes_of(items)), 2L)
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
  s <- shapes_of(items)
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
  s <- shapes_of(items)
  expect_equal(nrow(s), 1L)
  path <- s$path[[1]]
  pts <- path[path$cmd != "Z", ]
  expect_equal(c(pts$x[[1]], pts$y[[1]]), c(60, 10))
  expect_equal(path$cmd[[nrow(path)]], "Z")
  expect_true(s$closed)
  expect_equal(s$stroke, "222222")
})

test_that("an outer circle and a sector with the same box stay two shapes", {
  # A preset pairs with a custom path only when the path's end points and
  # midpoints lie on the preset's figure (the pie's outer circle and its
  # largest slice share a box, not a figure).
  sector <- prim_path("M200,15 A185,185 0 1 1 15,200 L200,200 Z", bbox = c(15, 15, 370, 370),
                      fill = "#ECECFF", stroke = "#000000", stroke_w = 2)
  ring <- prim_ellipse(200, 200, 186, fill = NULL, stroke = "#000000", stroke_w = 2)
  items <- generic_items(in_g(list(ring, sector), n = 5), viewBox = c(0, 0, 400, 400))
  s <- shapes_of(items)
  expect_equal(nrow(s), 2L)
  expect_equal(s$prst[[1]], "ellipse")
  expect_true(is.na(s$prst[[2]]))
  expect_equal(s$fill[[2]], "ECECFF")
})

# -- Four lines that draw a rectangle -------------------------------------------

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

# -- Compartments ---------------------------------------------------------------

# Only a box with no text of its own is split. Compartments are made after
# pairing and before texts merge, when no hand-built shape has a text yet, so
# that condition has no hand-built case here.

# A 200 x 120 box at (10, 10) and the lines drawn across it, in one scope.
class_box <- function(..., id = "box", fill = "#ECECFF", fill_a = 1, stroke = "#9370DB",
                      stroke_a = 1, stroke_w = 1, dash = NULL) {
  prim_rect(10, 10, 200, 120, id = id, fill = fill, fill_a = fill_a, stroke = stroke,
            stroke_a = stroke_a, stroke_w = stroke_w, dash = dash, ...)
}

divider <- function(y, x0 = 10, x1 = 210, stroke = "#9370DB", stroke_w = 1, dash = NULL, ...) {
  prim_line(x0, y, x1, y, stroke = stroke, stroke_w = stroke_w, dash = dash, ...)
}

# How many shapes a box with these other prims becomes.
n_shapes <- function(prims, viewBox = c(0, 0, 300, 200)) {
  nrow(shapes_of(generic_items(prims, viewBox = viewBox)))
}

test_that("a rect with one divider becomes two stacked rects and the divider is gone", {
  items <- generic_items(in_g(list(class_box(), divider(50)), n = 5),
                         viewBox = c(0, 0, 300, 200))
  s <- shapes_of(items)
  expect_equal(nrow(s), 2L)
  expect_equal(sum(items$kind == "line"), 0L)
  expect_true(all(s$prst == "rect"))
  expect_near(box_of(s[1, ]), c(10, 10, 200, 40), tol = 1e-6)
  expect_near(box_of(s[2, ]), c(10, 50, 200, 80), tol = 1e-6)
})

test_that("two dividers make three rects, top to bottom, with the box's name and paint", {
  # dividers drawn in either order
  items <- generic_items(in_g(list(
    class_box(fill_a = 0.5, stroke_a = 0.8, stroke_w = 1.5, dash = c(4, 2)),
    divider(90, stroke_w = 1.5, dash = c(4, 2)),
    divider(50, stroke_w = 1.5, dash = c(4, 2))
  ), n = 5), viewBox = c(0, 0, 300, 200))
  s <- shapes_of(items)
  expect_equal(nrow(s), 3L)
  expect_equal(nrow(items), 3L)
  expect_near(rbind(box_of(s[1, ]), box_of(s[2, ]), box_of(s[3, ])),
              rbind(c(10, 10, 200, 40), c(10, 50, 200, 40), c(10, 90, 200, 40)), tol = 1e-6)
  expect_equal(s$name, rep("mermaid:shape:box", 3))
  expect_equal(s$prst, rep("rect", 3))
  expect_equal(s$rot, rep(0, 3))
  expect_equal(s$fill, rep("ECECFF", 3))
  expect_equal(s$fill_a, rep(0.5, 3))
  expect_equal(s$stroke, rep("9370DB", 3))
  expect_equal(s$stroke_a, rep(0.8, 3))
  expect_equal(s$stroke_w, rep(1.5, 3))
  expect_equal(s$dash, rep(list(c(4, 2)), 3))
  expect_true(all(s$closed))
})

test_that("the compartments take the box's paint position", {
  items <- generic_items(in_g(list(
    prim_rect(230, 10, 20, 20, fill = "#00AA00", id = "before"),
    class_box(),
    prim_rect(230, 40, 20, 20, fill = "#AA0000", id = "after"),
    divider(50), divider(90)                   # Mermaid draws dividers last
  ), n = 5), viewBox = c(0, 0, 300, 200))
  drawn <- items[items$kind != "group", ]
  expect_equal(drawn$name, c("mermaid:shape:before", rep("mermaid:shape:box", 3),
                             "mermaid:shape:after"))
  expect_equal(drawn$y[2:4], c(10, 50, 90))
})

test_that("each compartment takes the stack of texts inside it", {
  items <- generic_items(in_g(list(
    class_box(),
    prim_text("«abstract»", x = 70, y = 14, w = 80, h = 14),
    prim_text("Animal", x = 85, y = 30, w = 50, h = 16),
    prim_text("+String name", x = 20, y = 54, w = 90, h = 14, anchor = "start"),
    prim_text("-int age", x = 20, y = 70, w = 50, h = 14, anchor = "start"),
    prim_text("+eat() void", x = 20, y = 94, w = 80, h = 14, anchor = "start"),
    prim_text("+sleep() void", x = 20, y = 110, w = 95, h = 14, anchor = "start"),
    divider(50), divider(90)
  ), n = 5), viewBox = c(0, 0, 300, 200))
  expect_equal(sum(items$kind == "text"), 0L)
  s <- shapes_of(items)
  expect_equal(nrow(s), 3L)
  if (nrow(s) != 3L) return()
  s <- s[order(s$y), ]
  expect_identical(s$label, c("«abstract»\nAnimal", "+String name\n-int age",
                              "+eat() void\n+sleep() void"))
  top <- s$text[[1]]
  expect_length(top$rows, 2L)
  expect_equal(top$align, "center")
  expect_equal(top$placement, "centre")       # union 14..46 in 10..50
  for (k in 2:3) {
    tx <- s$text[[k]]
    expect_equal(length(tx$rows), 2L, info = k)
    expect_equal(tx$align, "left", info = k)
    expect_near(tx$insets[[1]], 10, tol = 1e-9, info = k)  # 20 - 10
  }
})

test_that("a compartment with nothing inside it has no text", {
  items <- generic_items(in_g(list(
    class_box(),
    prim_text("Dog", x = 95, y = 22, w = 30, h = 16),
    divider(50), divider(90)
  ), n = 5), viewBox = c(0, 0, 300, 200))
  s <- shapes_of(items)
  expect_equal(nrow(s), 3L)
  if (nrow(s) != 3L) return()
  s <- s[order(s$y), ]
  expect_identical(s$label, c("Dog", "", ""))
  expect_null(s$text[[2]])
  expect_null(s$text[[3]])
  expect_equal(sum(items$kind == "text"), 0L)
})

test_that("a divider in a descendant scope of the box splits it", {
  items <- generic_items(in_g(c(list(class_box()), in_g(list(divider(50)), n = 6)), n = 5),
                         viewBox = c(0, 0, 300, 200))
  expect_equal(nrow(shapes_of(items)), 2L)
  expect_equal(sum(items$kind == "line"), 0L)
})

test_that("Mermaid's class box: paired outline, dividers in sibling <g>s, texts in compartments", {
  # A divider splits a box when it is in the box's scope or a descendant of it.
  # Mermaid puts the box paths in <g class="basic label-container"> and each
  # divider in its own <g class="divider">, siblings inside the node's <g>, so
  # the scope starts at the element that holds the shape's own element. The
  # box is also a filled and stroked open M-L-L-L path paired with a rough
  # outline; only preset rects are split, so the pair must be recognised as
  # a rect.
  box <- in_g(list(
    prim_path("M10,10 L210,10 L210,130 L10,130", fill = "#ECECFF", stroke = "#9370DB"),
    rough_rect(10, 10, 200, 120, fill = "#ECECFF", stroke = "#9370DB")
  ), n = 101, cls = "basic label-container")
  texts <- c(
    in_g(list(prim_text("Animal", x = 85, y = 22, w = 50, h = 16)), n = 105, cls = "label"),
    in_g(list(prim_text("+String name", x = 20, y = 54, w = 90, h = 14)), n = 111, cls = "label"),
    in_g(list(prim_text("-int age", x = 20, y = 70, w = 50, h = 14)), n = 112, cls = "label")
  )
  dividers <- c(in_g(list(mermaid_divider(10, 210, 50)), n = 138, cls = "divider"),
                in_g(list(mermaid_divider(10, 210, 90)), n = 140, cls = "divider"))
  items <- generic_items(in_g(c(box, texts, dividers), n = 100, cls = "node default",
                              unit = unit_node("Animal")),
                         viewBox = c(0, 0, 300, 200))
  s <- shapes_of(items)
  expect_equal(nrow(s), 3L)
  expect_true(all(s$prst %in% "rect"))
  expect_equal(sum(items$kind %in% c("line", "text")), 0L)
  if (nrow(s) == 3L) {
    s <- s[order(s$y), ]
    expect_near(s$y, c(10, 50, 90), tol = 1e-6)
    expect_identical(s$label, c("Animal", "+String name\n-int age", ""))
    expect_true(all(s$fill == "ECECFF" & s$stroke == "9370DB"))
  }
})

test_that("a divider that reaches only one edge, or a vertical line, splits nothing", {
  one_edge <- generic_items(in_g(list(class_box(), divider(50, x1 = 150)), n = 5),
                            viewBox = c(0, 0, 300, 200))
  expect_equal(nrow(shapes_of(one_edge)), 1L)
  expect_equal(sum(one_edge$kind == "line"), 1L)
  expect_near(box_of(shapes_of(one_edge)), c(10, 10, 200, 120), tol = 1e-6)

  other_edge <- generic_items(in_g(list(class_box(), divider(50, x0 = 60)), n = 5),
                              viewBox = c(0, 0, 300, 200))
  expect_equal(nrow(shapes_of(other_edge)), 1L)

  vertical <- generic_items(in_g(list(class_box(),
                                      prim_line(110, 10, 110, 130, stroke = "#9370DB")), n = 5),
                            viewBox = c(0, 0, 300, 200))
  expect_equal(nrow(shapes_of(vertical)), 1L)
  expect_equal(sum(vertical$kind == "line"), 1L)
})

test_that("a divider within 1 px of the top or bottom splits nothing", {
  expect_equal(n_shapes(in_g(list(class_box(), divider(10.5)), n = 5)), 1L)
  expect_equal(n_shapes(in_g(list(class_box(), divider(129.5)), n = 5)), 1L)
  expect_equal(n_shapes(in_g(list(class_box(), divider(11.5)), n = 5)), 2L)
})

test_that("a line with a marker is not a divider", {
  items <- generic_items(in_g(list(class_box(), divider(50, marker_end = "arrowhead")), n = 5),
                         viewBox = c(0, 0, 300, 200))
  expect_equal(nrow(shapes_of(items)), 1L)
  ln <- items[items$kind == "line", ]
  expect_equal(nrow(ln), 1L)
  expect_equal(ln$tail, "triangle")
})

test_that("a line in an unrelated scope is not a divider", {
  box <- in_g(in_g(list(class_box()), n = 5), n = 4)
  line <- in_g(in_g(list(divider(50)), n = 7), n = 6)
  items <- generic_items(c(box, line), viewBox = c(0, 0, 300, 200))
  expect_equal(nrow(shapes_of(items)), 1L)
  expect_equal(sum(items$kind == "line"), 1L)
})

test_that("a rotated rect, a roundRect or an ellipse is not split", {
  # 200 x 120 turned 30 degrees about (110, 70): its unrotated box is
  # c(10, 10, 200, 120), and the line spans that box's width.
  a <- pi / 6
  m <- c(cos(a), sin(a), -sin(a), cos(a),
         110 - (cos(a) * 110 - sin(a) * 70), 70 - (sin(a) * 110 + cos(a) * 70))
  turned <- prim_rect(10, 10, 200, 120, m = m, fill = "#ECECFF", stroke = "#9370DB")
  items <- generic_items(in_g(list(turned, divider(70)), n = 5), viewBox = c(0, 0, 300, 200))
  expect_equal(nrow(shapes_of(items)), 1L)
  expect_equal(sum(items$kind == "line"), 1L)

  rounded <- generic_items(in_g(list(class_box(rx = 8), divider(50)), n = 5),
                           viewBox = c(0, 0, 300, 200))
  expect_equal(shapes_of(rounded)$prst, "roundRect")
  expect_equal(sum(rounded$kind == "line"), 1L)

  oval <- generic_items(in_g(list(prim_ellipse(110, 70, 100, 60, fill = "#ECECFF",
                                               stroke = "#9370DB"),
                                  divider(70)), n = 5), viewBox = c(0, 0, 300, 200))
  expect_equal(shapes_of(oval)$prst, "ellipse")
  expect_equal(sum(oval$kind == "line"), 1L)
})

test_that("divider ends within max(1 px, 1 % of the width) of the edges; just outside they are not", {
  # 200 wide: 2 px
  expect_equal(n_shapes(in_g(list(class_box(), divider(50, x0 = 11.9, x1 = 208.1)), n = 5)), 2L)
  expect_equal(n_shapes(in_g(list(class_box(), divider(50, x0 = 8.1, x1 = 211.9)), n = 5)), 2L)
  expect_equal(n_shapes(in_g(list(class_box(), divider(50, x0 = 12.1)), n = 5)), 1L)
  expect_equal(n_shapes(in_g(list(class_box(), divider(50, x1 = 207.9)), n = 5)), 1L)
  # 50 wide: 1 % is 0.5, so the 1 px floor applies
  small <- function(x0, x1) {
    n_shapes(in_g(list(prim_rect(10, 10, 50, 60, fill = "#ECECFF", stroke = "#9370DB"),
                       divider(40, x0 = x0, x1 = x1)), n = 5))
  }
  expect_equal(small(10.9, 59.1), 2L)
  expect_equal(small(11.1, 60), 1L)
})

test_that("a line across a box drawn in another stroke is not a divider", {
  # The stacked boxes' shared borders draw the dividers, which holds only when
  # a divider has the box's stroke: the same colour, width and dash. Mermaid's
  # sequence `else` line ([3, 3]) across the alt frame ([2, 2]) and the
  # quadrant chart's 1 px centre line across its 2 px border would
  # otherwise vanish into a frame split in two.
  expect_equal(n_shapes(in_g(list(class_box(dash = c(2, 2)), divider(50, dash = c(3, 3))),
                             n = 5)), 1L)
  expect_equal(n_shapes(in_g(list(class_box(stroke_w = 2), divider(50, stroke_w = 1)), n = 5)), 1L)
  expect_equal(n_shapes(in_g(list(class_box(), divider(50, stroke = "#FF0000")), n = 5)), 1L)
})

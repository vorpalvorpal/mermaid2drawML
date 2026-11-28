# Step 2, second round (contract-step2.md, section H): a fill needs an area
# (H1), compartments (H2), stacks that touch (H3), two scope group
# exclusions (H4), centred ends too big for Word (H6) and marker parts
# (H7). Hand-built renders first, then the fixture assertions of H5.

shapes_of2 <- function(items) items[items$kind == "shape", ]
groups_of2 <- function(items) items[items$kind == "group", ]
box4_2 <- function(it) c(it$x, it$y, it$w, it$h)
kids_of2 <- function(items, name) items[!is.na(items$parent) & items$parent == name, ]
marker_rows2 <- function(items, id = NULL) {
  pre <- if (is.null(id)) "mermaid:marker:" else paste0("mermaid:marker:", id)
  items[!is.na(items$name) & startsWith(items$name, pre), ]
}

# Where an item starts and ends: a line from its box and flips, a custom
# shape from its path.
ends_of2 <- function(it) {
  if (it$kind == "line") {
    x <- c(it$x, it$x + it$w)
    y <- c(it$y, it$y + it$h)
    if (isTRUE(it$flip_h)) x <- rev(x)
    if (isTRUE(it$flip_v)) y <- rev(y)
    return(list(start = c(x[[1]], y[[1]]), end = c(x[[2]], y[[2]])))
  }
  p <- it$path[[1]]
  p <- p[p$cmd != "Z", ]
  list(start = c(p$x[[1]], p$y[[1]]), end = c(p$x[[nrow(p)]], p$y[[nrow(p)]]))
}

box_gap2 <- function(a, b) {
  dx <- max(0, a[1] - (b[1] + b[3]), b[1] - (a[1] + a[3]))
  dy <- max(0, a[2] - (b[2] + b[4]), b[2] - (a[2] + a[4]))
  sqrt(dx^2 + dy^2)
}

# The outline of a box as Mermaid draws its rough stroke: straight cubics.
rough_rect <- function(x, y, w, h, ...) {
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

# Mermaid's class divider: two straight cubic subpaths of zero height,
# filled and stroked as the box is.
mermaid_divider <- function(x0, x1, y, fill = "#ECECFF", stroke = "#9370DB", ...) {
  d <- sprintf("M%s,%s C%s,%s %s,%s %s,%s M%s,%s C%s,%s %s,%s %s,%s",
               x0, y, x0 + (x1 - x0) / 3, y, x0 + 2 * (x1 - x0) / 3, y, x1, y,
               x0, y, x0 + (x1 - x0) / 4, y, x0 + 3 * (x1 - x0) / 4, y, x1, y)
  prim_path(d, bbox = c(x0, y, x1 - x0, 0), fill = fill, stroke = stroke, ...)
}

# -- H1: a fill needs an area ------------------------------------------------------

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
  expect_near(box4_2(items), c(10, 50, 100, 0), tol = 1e-9)

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
  expect_near(box4_2(items), c(50, 10, 0, 80), tol = 1e-9)

  wsp <- all_wsp(generic_doc(list(p)))
  expect_false(is_missing_node(xfirst(wsp[[1]], "./wps:cNvCnPr")))
  expect_true(is_missing_node(xfirst(wsp[[1]], "./wps:spPr/a:solidFill")))
})

test_that("Mermaid's divider (a filled zero-height path of straight cubics) is an unfilled line", {
  # SPEC? H1 says "a single segment becomes a line item", but Mermaid's
  # class, ER and requirement dividers are two straight cubic subpaths
  # (requirement's unfilled). H2 needs them as `line` items, so this reads
  # "a path whose box is a segment is that segment".
  items <- generic_items(list(mermaid_divider(10, 110, 50)))
  expect_equal(nrow(items), 1L)
  expect_equal(items$kind, "line")
  expect_true(is.na(items$fill))
  expect_near(box4_2(items), c(10, 50, 100, 0), tol = 1e-6)

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

# ruling: I1 makes a straight edge (a path of zero width or height, whatever
# its commands) a line item; it keeps its direction, its marker end and the
# C2 end move.
test_that("ruling: a straight multi-point path with an end marker is a line that keeps its end", {
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
  expect_near(ends_of2(rtl)$start, c(110, 50), tol = 1e-9)
  expect_near(ends_of2(rtl)$end, c(6, 50), tol = 1e-9)

  btt <- run("M50,110 L50,60 L50,10")
  expect_equal(btt$kind, "line")
  expect_equal(btt$tail, "triangle")
  expect_true(btt$flip_v)
  expect_near(ends_of2(btt)$start, c(50, 110), tol = 1e-9)
  expect_near(ends_of2(btt)$end, c(50, 6), tol = 1e-9)
})

test_that("ruling: flowchart-paint A-B is a line whose triangle tail ends at B", {
  items <- fixture_scene("flowchart-paint")$items
  e <- item_named(items, "mermaid:edge:A-B")
  expect_equal(nrow(e), 1L)
  if (nrow(e) != 1L) return()
  expect_equal(e$kind, "line")
  expect_equal(e$tail, "triangle")
  b <- item_named(items, "mermaid:node:B")
  expect_equal(nrow(b), 1L)
  if (nrow(b) != 1L) return()
  ends <- ends_of2(e)
  bb <- box4_2(b)
  # the flips put the tail end at B, not the head end
  expect_lt(box_gap2(c(ends$end, 0, 0), bb), 1)
  expect_gt(box_gap2(c(ends$start, 0, 0), bb), box_gap2(c(ends$end, 0, 0), bb))
})

# -- H2: compartments --------------------------------------------------------------

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
  nrow(shapes_of2(generic_items(prims, viewBox = viewBox)))
}

test_that("a rect with one divider becomes two stacked rects and the divider is gone", {
  items <- generic_items(in_g(list(class_box(), divider(50)), n = 5),
                         viewBox = c(0, 0, 300, 200))
  s <- shapes_of2(items)
  expect_equal(nrow(s), 2L)
  expect_equal(sum(items$kind == "line"), 0L)
  expect_true(all(s$prst == "rect"))
  expect_near(box4_2(s[1, ]), c(10, 10, 200, 40), tol = 1e-6)
  expect_near(box4_2(s[2, ]), c(10, 50, 200, 80), tol = 1e-6)
})

test_that("two dividers make three rects, top to bottom, with the box's name and paint", {
  # dividers drawn in either order
  items <- generic_items(in_g(list(
    class_box(fill_a = 0.5, stroke_a = 0.8, stroke_w = 1.5, dash = c(4, 2)),
    divider(90, stroke_w = 1.5, dash = c(4, 2)),
    divider(50, stroke_w = 1.5, dash = c(4, 2))
  ), n = 5), viewBox = c(0, 0, 300, 200))
  s <- shapes_of2(items)
  expect_equal(nrow(s), 3L)
  expect_equal(nrow(items), 3L)
  expect_near(rbind(box4_2(s[1, ]), box4_2(s[2, ]), box4_2(s[3, ])),
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
  s <- shapes_of2(items)
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
    expect_length(tx$rows, 2L)
    expect_equal(tx$align, "left", info = k)
    expect_near(tx$insets[[1]], 10, tol = 1e-9)  # 20 - 10
  }
})

test_that("a compartment with nothing inside it has no text", {
  items <- generic_items(in_g(list(
    class_box(),
    prim_text("Dog", x = 95, y = 22, w = 30, h = 16),
    divider(50), divider(90)
  ), n = 5), viewBox = c(0, 0, 300, 200))
  s <- shapes_of2(items)
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
  expect_equal(nrow(shapes_of2(items)), 2L)
  expect_equal(sum(items$kind == "line"), 0L)
})

test_that("Mermaid's class box: paired outline, dividers in sibling <g>s, texts in compartments", {
  # SPEC? H2 says "the shape's scope or a descendant of it". Mermaid puts the
  # box paths in <g class="basic label-container"> and each divider in its own
  # <g class="divider">, siblings inside the node's <g>; H5 needs these split,
  # so this reads the scope chain as starting at the element that holds the
  # shape's own element. The box is also a fill+stroke open M-L-L-L path
  # paired with a rough outline (B2): today custom geometry, while H2 splits
  # only preset rects, so the paired closed path must be recognised as a rect.
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
  s <- shapes_of2(items)
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
  expect_equal(nrow(shapes_of2(one_edge)), 1L)
  expect_equal(sum(one_edge$kind == "line"), 1L)
  expect_near(box4_2(shapes_of2(one_edge)), c(10, 10, 200, 120), tol = 1e-6)

  other_edge <- generic_items(in_g(list(class_box(), divider(50, x0 = 60)), n = 5),
                              viewBox = c(0, 0, 300, 200))
  expect_equal(nrow(shapes_of2(other_edge)), 1L)

  vertical <- generic_items(in_g(list(class_box(),
                                      prim_line(110, 10, 110, 130, stroke = "#9370DB")), n = 5),
                            viewBox = c(0, 0, 300, 200))
  expect_equal(nrow(shapes_of2(vertical)), 1L)
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
  expect_equal(nrow(shapes_of2(items)), 1L)
  ln <- items[items$kind == "line", ]
  expect_equal(nrow(ln), 1L)
  expect_equal(ln$tail, "triangle")
})

test_that("a line in an unrelated scope is not a divider", {
  box <- in_g(in_g(list(class_box()), n = 5), n = 4)
  line <- in_g(in_g(list(divider(50)), n = 7), n = 6)
  items <- generic_items(c(box, line), viewBox = c(0, 0, 300, 200))
  expect_equal(nrow(shapes_of2(items)), 1L)
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
  expect_equal(nrow(shapes_of2(items)), 1L)
  expect_equal(sum(items$kind == "line"), 1L)

  rounded <- generic_items(in_g(list(class_box(rx = 8), divider(50)), n = 5),
                           viewBox = c(0, 0, 300, 200))
  expect_equal(shapes_of2(rounded)$prst, "roundRect")
  expect_equal(sum(rounded$kind == "line"), 1L)

  oval <- generic_items(in_g(list(prim_ellipse(110, 70, 100, 60, fill = "#ECECFF",
                                               stroke = "#9370DB"),
                                  divider(70)), n = 5), viewBox = c(0, 0, 300, 200))
  expect_equal(shapes_of2(oval)$prst, "ellipse")
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
  # SPEC? H2 sets no paint condition, but says the stacked boxes' shared
  # borders draw the dividers, which holds only when the divider has the
  # box's stroke. Read as: same stroke colour, width and dash. Mermaid's
  # sequence `else` line ([3, 3]) across the alt frame ([2, 2]) and the
  # quadrant chart's 1 px centre line across its 2 px border would
  # otherwise vanish into a frame split in two.
  expect_equal(n_shapes(in_g(list(class_box(dash = c(2, 2)), divider(50, dash = c(3, 3))),
                             n = 5)), 1L)
  expect_equal(n_shapes(in_g(list(class_box(stroke_w = 2), divider(50, stroke_w = 1)), n = 5)), 1L)
  expect_equal(n_shapes(in_g(list(class_box(), divider(50, stroke = "#FF0000")), n = 5)), 1L)
})

# SPEC? H2: "no text of its own". Compartments are made after pairing and
# before texts merge, when no generic shape has a text yet, so there is no
# hand-built case to test; read as "not a shape a unit rule gave a text".

# -- H3: stacks may touch ------------------------------------------------------------

test_that("two texts overlapping by 20 % of the smaller height stack; by 30 % they do not", {
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
})

# -- H4: scope groups, two exclusions ------------------------------------------------

# Something drawn outside the scope under test, so that the scope is not
# the whole diagram (H4b).
elsewhere <- function() list(prim_rect(150, 60, 40, 30, fill = "#00AA00"))

test_that("a scope holding only free texts is not grouped", {
  labels <- in_g(list(prim_text("1", x = 10, y = 10, w = 10, h = 12),
                      prim_text("*", x = 60, y = 40, w = 10, h = 12)), n = 5, cls = "edgeLabels")
  items <- generic_items(c(labels, elsewhere()))
  expect_equal(nrow(groups_of2(items)), 0L)
  expect_equal(sum(items$kind == "text"), 2L)
})

test_that("a scope holding a shape and a text is still grouped", {
  legend <- in_g(list(prim_rect(10, 10, 18, 18, fill = "#ECECFF"),
                      prim_text("Dogs", x = 32, y = 10, w = 40, h = 18, anchor = "start")),
                 n = 5, cls = "legend")
  items <- generic_items(c(legend, elsewhere()))
  g <- groups_of2(items)
  expect_equal(g$name, "mermaid:group:legend")
  expect_equal(nrow(kids_of2(items, "mermaid:group:legend")), 2L)
})

test_that("a scope covering 90 % of the union of all items both ways is not grouped, whatever the viewBox", {
  whole <- in_g(list(prim_rect(10, 10, 100, 50), prim_rect(50, 30, 60, 30)), n = 5, cls = "all")
  # union of all items: 10..115 x 10..65, the scope 100 x 50 of 105 x 55
  items <- generic_items(c(whole, list(prim_rect(105, 55, 10, 10))),
                         viewBox = c(0, 0, 1000, 1000))
  expect_equal(nrow(groups_of2(items)), 0L)
})

test_that("a scope covering 90 % of the union in one direction only is grouped", {
  part <- in_g(list(prim_rect(10, 10, 100, 50), prim_rect(50, 30, 60, 30)), n = 5, cls = "all")
  # union 10..110 x 10..110: the scope is all of its width, half its height
  items <- generic_items(c(part, list(prim_rect(50, 100, 10, 10))),
                         viewBox = c(0, 0, 1000, 1000))
  expect_equal(groups_of2(items)$name, "mermaid:group:all")
})

# -- H6: a centred end too big for Word is a shape ---------------------------------------

h6_line <- function(..., x1 = 10, y1 = 50, x2 = 110, y2 = 50, stroke_w = 1) {
  prim_line(x1, y1, x2, y2, id = "msg", stroke = "#333333", stroke_w = stroke_w, ...)
}
h6_items <- function(line, id, def, viewBox = c(0, 0, 200, 200)) {
  generic_items(list(line), viewBox = viewBox, markers = stats::setNames(list(def), id))
}
line_of <- function(items) items[items$kind == "line", ]

test_that("the flowchart circle end (11 px) on a 1 px line is a marker shape, the end not moved", {
  id <- "flowchart-v2-circleEnd"
  items <- h6_items(h6_line(marker_end = id), id, marker_circle_end())
  ln <- line_of(items)
  expect_equal(nrow(ln), 1L)
  expect_equal(ln$tail, "none")
  expect_near(ends_of2(ln)$end, c(110, 50), tol = 1e-9)
  s <- marker_rows2(items, id)
  expect_equal(nrow(s), 1L)
  if (nrow(s) != 1L) return()
  expect_equal(s$kind, "shape")
  expect_equal(s$prst, "ellipse")
  # content 0..10 at scale 1.1, refX 11 on the end: 97.9 .. 108.9
  expect_near(box4_2(s), c(97.9, 44.5, 11, 11), tol = 0.01)
  expect_equal(s$fill, "333333")
  # grouped with its line (C3)
  expect_false(is.na(s$parent))
  expect_equal(ln$parent, s$parent)
})

test_that("on a line under 1 px the limit is still 5 px", {
  id <- "flowchart-v2-circleEnd"
  items <- h6_items(h6_line(stroke_w = 0.5, marker_end = id), id, marker_circle_end())
  expect_equal(line_of(items)$tail, "none")
  expect_equal(nrow(marker_rows2(items, id)), 1L)
})

test_that("the same circle on a 3 px line (11 <= 15) stays a Word oval, the end at its centre", {
  # userSpaceOnUse: the marker is 11 px whatever the stroke
  id <- "flowchart-v2-circleEnd"
  items <- h6_items(h6_line(stroke_w = 3, marker_end = id), id, marker_circle_end())
  ln <- line_of(items)
  expect_equal(ln$tail, "oval")
  expect_near(ends_of2(ln)$end, c(103.4, 50), tol = 1e-9)    # (5 - 11) x 1.1
  expect_equal(nrow(marker_rows2(items)), 0L)
})

test_that("the limit is 5 x the stroke width: 10.5 < 11 is a shape, 11.5 is a Word end", {
  id <- "flowchart-v2-circleEnd"
  thin <- h6_items(h6_line(stroke_w = 2.1, marker_end = id), id, marker_circle_end())
  expect_equal(line_of(thin)$tail, "none")
  expect_equal(nrow(marker_rows2(thin, id)), 1L)
  thick <- h6_items(h6_line(stroke_w = 2.3, marker_end = id), id, marker_circle_end())
  expect_equal(line_of(thick)$tail, "oval")
  expect_equal(nrow(marker_rows2(thick)), 0L)
})

test_that("the class composition diamond (17 px) on a 1 px line is a shape at the start", {
  id <- "class-compositionStart"
  items <- h6_items(h6_line(x1 = 30, x2 = 130, marker_start = id), id, marker_composition_start())
  ln <- line_of(items)
  expect_equal(ln$head, "none")
  expect_near(ends_of2(ln)$start, c(30, 50), tol = 1e-9)
  s <- marker_rows2(items, id)
  expect_equal(nrow(s), 1L)
  if (nrow(s) != 1L) return()
  # plain auto at a start: content x 1..18 runs back from refX 18 on the end
  expect_near(box4_2(s), c(13, 44, 17, 12), tol = 0.01)
  expect_equal(s$fill, "333333")
})

test_that("with strokeWidth units the diamond grows with the line: a shape on a 3 px line too", {
  # 17 x 3 = 51 > 5 x 3 = 15
  id <- "class-compositionStart"
  items <- h6_items(h6_line(x1 = 80, x2 = 180, stroke_w = 3, marker_start = id), id,
                    marker_composition_start())
  expect_equal(line_of(items)$head, "none")
  s <- marker_rows2(items, id)
  expect_equal(nrow(s), 1L)
  if (nrow(s) != 1L) return()
  expect_near(box4_2(s), c(29, 32, 51, 36), tol = 0.01)
  expect_equal(s$stroke_w, 3)
})

test_that("a composition diamond in userSpaceOnUse units on a 4 px line stays a Word diamond", {
  # 17 <= 20; centre offset back toward the node: refX - centre = 18 - 9.5
  id <- "class-compositionStart"
  def <- marker_composition_start()
  def$units <- "userSpaceOnUse"
  items <- h6_items(h6_line(x1 = 30, x2 = 130, stroke_w = 4, marker_start = id), id, def)
  ln <- line_of(items)
  expect_equal(ln$head, "diamond")
  expect_near(ends_of2(ln)$start, c(21.5, 50), tol = 1e-9)
  expect_equal(nrow(marker_rows2(items)), 0L)
})

test_that("a big triangle, stealth or arrow end stays a Word end", {
  # 40 px on a 1 px line: viewBox 0..10 at scale 4, tip 5 beyond refX -> +20 px
  def <- marker_arrow_point(w = 40)
  for (case in list(c("flowchart-v2-pointEnd", "triangle"), c("filled-head", "stealth"),
                    c("class-dependencyEnd", "arrow"))) {
    items <- h6_items(h6_line(marker_end = case[[1]]), case[[1]], def)
    ln <- line_of(items)
    expect_equal(ln$tail, case[[2]], info = case[[1]])
    expect_near(ends_of2(ln)$end, c(130, 50), tol = 1e-9)
    expect_equal(nrow(marker_rows2(items)), 0L, info = case[[1]])
  }
})

# -- H7: marker parts ------------------------------------------------------------------

test_that("a marker part that is a line prim is an item of kind shape", {
  id <- "requirement-requirement_containsEnd"
  def <- marker_def(NULL, refX = 0, refY = 10, w = 20, h = 20, units = "strokeWidth",
                    prims = list(prim_ellipse(10, 10, 9, fill = NULL, stroke = "#333333"),
                                 prim_line(1, 10, 19, 10, stroke = "#333333"),
                                 prim_line(10, 1, 10, 19, stroke = "#333333")))
  items <- h6_items(h6_line(marker_end = id), id, def)
  s <- marker_rows2(items, id)
  expect_equal(nrow(s), 3L)
  expect_equal(s$kind, rep("shape", 3))
  expect_equal(sum(items$kind == "line"), 1L)        # only the edge itself
})

two_dots <- function() {
  marker_def(NULL, refX = 6, refY = 5, w = 12, h = 10, units = "userSpaceOnUse",
             prims = list(prim_ellipse(3, 5, 2, fill = "#333333", stroke = NULL),
                          prim_ellipse(9, 5, 2, fill = "#333333", stroke = NULL)))
}

test_that("a zero-length marked stroke with two marker parts is one group of two shapes", {
  dot <- prim_line(50, 50, 50, 50, id = "dot", stroke = "#333333", marker_end = "my-pair")
  items <- generic_items(list(dot), markers = list("my-pair" = two_dots()))
  expect_equal(sum(items$kind == "line"), 0L)
  g <- groups_of2(items)
  expect_equal(nrow(g), 1L)
  # SPEC? "grouped under the stroke's name": read as C3's group named after
  # the line item, mermaid:edge_group:<item name without "mermaid:">.
  expect_equal(g$name, "mermaid:edge_group:line:dot")
  kids <- kids_of2(items, g$name)
  expect_equal(kids$kind, c("shape", "shape"))
  expect_equal(kids$name, rep("mermaid:marker:my-pair", 2))
  expect_equal(nrow(items), 3L)
})

test_that("a zero-length marked stroke with one marker part is one ungrouped shape", {
  one <- two_dots()
  one$prims <- one$prims[1]
  dot <- prim_line(50, 50, 50, 50, id = "dot", stroke = "#333333", marker_end = "my-dot")
  items <- generic_items(list(dot), markers = list("my-dot" = one))
  expect_equal(nrow(items), 1L)
  expect_equal(items$kind, "shape")
  expect_equal(items$name, "mermaid:marker:my-dot")
  expect_true(is.na(items$parent))
})

# -- H5: fixtures --------------------------------------------------------------------

test_that("classDiagram: Animal is three stacked rects holding its title, attributes and methods", {
  items <- fixture_scene("classDiagram")$items
  u <- item_named(items, "mermaid:node:Animal")
  expect_equal(nrow(u), 1L)
  sub <- subtree_rows(items, "mermaid:node:Animal")
  s <- sub[sub$kind == "shape", ]
  expect_equal(nrow(s), 3L)
  expect_true(all(s$prst %in% "rect"))
  if (nrow(s) != 3L) return()
  s <- s[order(s$y), ]
  # the box 334.31, 8, 165.03 x 216; dividers at y 80 and 152
  expect_near(cbind(s$x, s$y, s$w, s$h),
              cbind(rep(334.31, 3), c(8, 80, 152), rep(165.03, 3), rep(72, 3)), tol = 0.02)
  expect_identical(s$label, c("«abstract»\nAnimal", "+String name\n-int age",
                              "+eat() : void\n+sleep() : void"))
  top <- s$text[[1]]
  expect_length(top$rows, 2L)
  expect_equal(top$align, "center")
  # SPEC? the bridge reports anchor "middle" for these HTML rows, so B6 alone
  # would give align "center" with an inset (centring the shorter row in
  # Word); H5 says the attribute and method rows are left-aligned.
  expect_equal(s$text[[2]]$align, "left")
  expect_equal(s$text[[3]]$align, "left")
  expect_length(s$text[[2]]$rows, 2L)
  expect_length(s$text[[3]]$rows, 2L)
})

test_that("classDiagram: no free text and no divider line is left in any class node", {
  items <- fixture_scene("classDiagram")$items
  nodes <- items$name[items$kind == "group" & startsWith(items$name, "mermaid:node:")]
  expect_gte(length(nodes), 8L)
  for (nm in nodes) {
    sub <- subtree_rows(items, nm)
    expect_equal(sum(sub$kind == "text"), 0L, info = nm)
    drawn <- sub[sub$kind != "group", ]
    expect_equal(sum(drawn$kind == "line" | drawn$h <= 0.01), 0L, info = nm)
  }
})

test_that("requirementDiagram: each box is two stacked rects with their texts", {
  items <- fixture_scene("requirementDiagram")$items
  want <- list(
    test_req = list(box = rbind(c(8, 194, 157.73, 68), c(8, 262, 157.73, 116)),
                    label = c("<<Requirement>>\ntest_req",
                              "ID: 1\nText: the test text.\nRisk: High\nVerification: Test")),
    test_entity = list(box = rbind(c(17.82, 8, 138.09, 68), c(17.82, 76, 138.09, 44)),
                       label = c("<<Element>>\ntest_entity", "Type: simulation"))
  )
  for (id in names(want)) {
    sub <- subtree_rows(items, paste0("mermaid:node:", id))
    s <- sub[sub$kind == "shape", ]
    expect_equal(nrow(s), 2L, info = id)
    expect_equal(sum(sub$kind %in% c("text", "line")), 0L, info = id)
    if (nrow(s) != 2L) next
    s <- s[order(s$y), ]
    expect_true(all(s$prst %in% "rect"), info = id)
    expect_near(cbind(s$x, s$y, s$w, s$h), want[[id]]$box, tol = 0.02)
    expect_identical(s$label, want[[id]]$label, info = id)
  }
})

test_that("treemap-beta: every leaf's label and value are one text in the leaf's rect", {
  items <- fixture_scene("treemap-beta")$items
  prims <- fixture_prims("treemap-beta")
  scope_of <- function(p) p$anc[[length(p$anc)]]$n
  leaves <- purrr::keep(prims, function(p) identical(p$cls, "treemapLeaf"))
  expect_equal(length(leaves), 4L)
  for (leaf in leaves) {
    texts <- purrr::keep(prims, function(p) {
      p$kind == "text" && identical(scope_of(p), scope_of(leaf))
    })
    label <- paste(unlist(purrr::map(texts, prim_row_texts)), collapse = "\n")
    b <- as.numeric(unlist(leaf$bbox))
    hit <- items[items$kind == "shape" & abs(items$x - b[1]) < 0.5 & abs(items$y - b[2]) < 0.5 &
                   abs(items$w - b[3]) < 0.5 & abs(items$h - b[4]) < 0.5, ]
    expect_equal(nrow(hit), 1L, info = label)
    if (nrow(hit) == 1L) expect_identical(hit$label, label)
  }
})

test_that("classDiagram has no mermaid:group:edgeLabels", {
  items <- fixture_scene("classDiagram")$items
  expect_false(any(grepl("^mermaid:group:edgeLabels", items$name)))
})

test_that("pie: no group holds every other item", {
  items <- fixture_scene("pie")$items
  drawn <- sum(items$kind != "group")
  for (g in items$name[items$kind == "group"]) {
    sub <- subtree_rows(items, g)
    expect_lt(sum(sub$kind != "group"), drawn, label = g)
  }
})

test_that("flowchart-full: the circle end of a-f is a marker shape; the edge ends where Mermaid's did", {
  items <- fixture_scene("flowchart-full")$items
  e <- item_named(items, "mermaid:edge:a-f")
  expect_equal(nrow(e), 1L)
  if (nrow(e) != 1L) return()
  expect_equal(c(e$head, e$tail), c("none", "none"))
  expect_equal(e$parent, "mermaid:edge_group:a-f")
  marks <- kids_of2(items, "mermaid:edge_group:a-f")
  marks <- marks[marks$name %in% "mermaid:marker:flowchart-v2-circleEnd", ]
  expect_equal(nrow(marks), 1L)
  if (nrow(marks) != 1L) return()
  expect_equal(marks$prst, "ellipse")
  expect_false(is.na(marks$fill))
  p <- purrr::detect(fixture_prims("flowchart-full", "path"), function(p) identical(p$id, "L_a_f_0"))
  path <- path_transform(path_parse(p$geom$d), prim_matrix(p))
  path <- path[path$cmd != "Z", ]
  drawn_end <- c(path$x[[nrow(path)]], path$y[[nrow(path)]])
  expect_near(ends_of2(e)$end, drawn_end, tol = 0.01)        # not moved
  expect_lt(box_gap2(c(drawn_end, 0, 0), box4_2(marks)), 1.5)
})

test_that("classDiagram: the Dog-Tail composition is a filled diamond shape touching the Dog box", {
  items <- fixture_scene("classDiagram")$items
  e <- item_named(items, "mermaid:edge:Dog-Tail")
  expect_equal(nrow(e), 1L)
  if (nrow(e) != 1L) return()
  expect_equal(e$head, "none")
  g <- "mermaid:edge_group:Dog-Tail"
  expect_equal(e$parent, g)
  marks <- kids_of2(items, g)
  marks <- marks[marks$name %in% "mermaid:marker:class-compositionStart", ]
  expect_equal(nrow(marks), 1L)
  if (nrow(marks) != 1L) return()
  expect_equal(marks$fill, "333333")
  path <- marks$path[[1]]
  corners <- if (is.null(path)) 0L else sum(path$cmd %in% c("M", "L"))
  expect_true(corners == 4L || marks$prst %in% "diamond")
  dog <- item_named(items, "mermaid:node:Dog")
  expect_equal(nrow(dog), 1L)
  if (nrow(dog) == 1L) expect_lt(box_gap2(box4_2(marks), box4_2(dog)), 1.5)
  p <- purrr::detect(fixture_prims("classDiagram", "path"), function(p) identical(p$id, "id_Dog_Tail_4"))
  path <- path_transform(path_parse(p$geom$d), prim_matrix(p))
  expect_near(ends_of2(e)$start, c(path$x[[1]], path$y[[1]]), tol = 0.01)   # not moved
})

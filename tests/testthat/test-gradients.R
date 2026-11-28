# Gradients: a fill or stroke painted with an SVG linear or radial gradient
# is written as a DrawingML gradient fill (a:gradFill), as the shape's fill
# or inside its a:ln, instead of the flat colour of its first stop. The
# bridge gives the gradient in root user space; the item carries it mapped
# into its own box (fill_grad / stroke_grad, NULL when there is none).

# The box every hand-built gradient below is drawn in.
gbox <- c(20, 10, 100, 50)

# Two stops, red to blue, opaque unless given.
red_blue <- function(a1 = 1, a2 = 1) {
  list(grad_stop(0, "#FF0000", a1), grad_stop(1, "#0000FF", a2))
}

grad_rect <- function(grad, ...) prim_rect(gbox[1], gbox[2], gbox[3], gbox[4], fill_grad = grad, ...)

the_wsp <- function(doc) {
  s <- all_wsp(doc)
  testthat::expect_length(s, 1L)
  s[[1]]
}

fill_grad_xml <- function(prim, ...) grad_of(shape_sppr(the_wsp(generic_doc(list(prim), ...))))

# DrawingML's linear angle: clockwise from the x axis, in 60000ths of a
# degree. SVG's y axis points down, so atan2(dy, dx) on SVG points already
# turns clockwise.
lin_ang <- function(p1, p2) {
  ((atan2(p2[2] - p1[2], p2[1] - p1[1]) * 180 / pi) %% 360) * 60000
}

# Where offset `o` of the gradient p1 -> p2 falls along the box's extent in
# that direction (0 at the box's first corner, 100000 at its last), clamped
# to the box.
lin_pos <- function(box, p1, p2, o) {
  d <- p2 - p1
  len <- sqrt(sum(d^2))
  u <- d / len
  corners <- rbind(c(box[1], box[2]), c(box[1] + box[3], box[2]),
                   c(box[1], box[2] + box[4]), c(box[1] + box[3], box[2] + box[4]))
  proj <- as.vector(corners %*% u)
  s <- sum(p1 * u) + o * len
  pmin(1, pmax(0, (s - min(proj)) / (max(proj) - min(proj)))) * 100000
}

# -- Scene ---------------------------------------------------------------------

test_that("every item has fill_grad and stroke_grad, NULL for flat paint", {
  items <- generic_items(list(prim_rect(10, 10, 50, 20)))
  expect_true(all(c("fill_grad", "stroke_grad") %in% names(items)))
  expect_type(items$fill_grad, "list")
  expect_type(items$stroke_grad, "list")
  expect_null(items$fill_grad[[1]])
  expect_null(items$stroke_grad[[1]])
})

test_that("a gradient fill reaches the item as fill_grad, and only there", {
  g <- linear_grad(c(20, 0), c(120, 0), red_blue())
  items <- generic_items(list(grad_rect(g)))
  expect_equal(nrow(items), 1L)
  expect_false(is.null(items$fill_grad[[1]]))
  expect_null(items$stroke_grad[[1]])
})

test_that("a gradient stroke reaches the item as stroke_grad", {
  g <- linear_grad(c(20, 10), c(120, 60), red_blue())
  items <- generic_items(list(prim_path("M20 10 L120 60", fill = NULL, stroke_grad = g)))
  expect_null(items$fill_grad[[1]])
  expect_false(is.null(items$stroke_grad[[1]]))
})

# -- Linear gradients in the XML -------------------------------------------------

test_that("a horizontal gradient spanning the box is a:gradFill with stops at 0 and 100 %", {
  p1 <- c(20, 0)
  p2 <- c(120, 0)
  doc <- generic_doc(list(grad_rect(linear_grad(p1, p2, red_blue()))))
  sp <- shape_sppr(the_wsp(doc))
  expect_true(is_missing_node(xfirst(sp, "./a:solidFill")))
  g <- grad_of(sp)
  expect_false(is.null(g))
  expect_equal(g$stops$color, c("FF0000", "0000FF"))
  expect_near(g$stops$pos, lin_pos(gbox, p1, p2, c(0, 1)), tol = 1)
  expect_near(g$ang, lin_ang(p1, p2), tol = 1)
  expect_equal(g$scaled, "0")
  expect_true(is.na(g$path))
  expect_length(schema_order_problems(doc), 0L)
})

test_that("the angle follows the direction p1 -> p2: right, down, left, up and diagonals", {
  # Each case: p1, p2 in root space (SVG: y down, so p1 above p2 points down).
  cases <- list(
    right = list(c(20, 35), c(120, 35)),
    down = list(c(70, 10), c(70, 60)),
    left = list(c(120, 35), c(20, 35)),
    up = list(c(70, 60), c(70, 10)),
    down_right = list(c(20, 10), c(120, 60)),
    up_right = list(c(20, 60), c(120, 10)),
    down_left = list(c(120, 10), c(20, 60))
  )
  for (nm in names(cases)) {
    p1 <- cases[[nm]][[1]]
    p2 <- cases[[nm]][[2]]
    g <- fill_grad_xml(grad_rect(linear_grad(p1, p2, red_blue())))
    expect_near(g$ang, lin_ang(p1, p2), tol = 1, info = nm)
    expect_near(g$stops$pos, lin_pos(gbox, p1, p2, c(0, 1)), tol = 1, info = nm)
  }
  # The cardinal directions, spelled out: 0, 90, 180 and 270 degrees.
  expect_equal(lin_ang(c(0, 0), c(1, 0)), 0)
  expect_equal(lin_ang(c(0, 0), c(0, 1)), 90 * 60000)
  expect_equal(lin_ang(c(1, 0), c(0, 0)), 180 * 60000)
  expect_equal(lin_ang(c(0, 1), c(0, 0)), 270 * 60000)
})

test_that("a gradient vector shorter than the box has its stops re-expressed along the box", {
  # The vector spans x 45..95 of a box spanning 20..120: offsets 0, 0.5 and 1
  # fall at 25 %, 50 % and 75 % of the box.
  p1 <- c(45, 30)
  p2 <- c(95, 30)
  stops <- list(grad_stop(0, "#FF0000"), grad_stop(0.5, "#00FF00"), grad_stop(1, "#0000FF"))
  g <- fill_grad_xml(grad_rect(linear_grad(p1, p2, stops)))
  expect_equal(g$stops$color, c("FF0000", "00FF00", "0000FF"))
  expect_near(g$stops$pos, lin_pos(gbox, p1, p2, c(0, 0.5, 1)), tol = 1)
  expect_near(g$stops$pos, c(25000, 50000, 75000), tol = 1)
})

test_that("stops beyond the box are clamped to 0 and 100 %", {
  stops <- list(grad_stop(0, "#FF0000"), grad_stop(0.5, "#00FF00"), grad_stop(1, "#0000FF"))
  for (v in list(list(c(-30, 30), c(70, 30)), list(c(70, 30), c(170, 30)),
                 list(c(-80, 30), c(220, 30)))) {
    g <- fill_grad_xml(grad_rect(linear_grad(v[[1]], v[[2]], stops)))
    expect_near(g$stops$pos, lin_pos(gbox, v[[1]], v[[2]], c(0, 0.5, 1)), tol = 1)
    expect_true(all(g$stops$pos >= 0 & g$stops$pos <= 100000))
  }
})

test_that("a diagonal gradient on a box that is not square uses the box's extent along it", {
  p1 <- c(20, 10)
  p2 <- c(70, 60)
  g <- fill_grad_xml(grad_rect(linear_grad(p1, p2, red_blue())))
  expect_near(g$ang, 45 * 60000, tol = 1)
  expect_near(g$stops$pos, lin_pos(gbox, p1, p2, c(0, 1)), tol = 1)
})

test_that("stop alphas are written as a:alpha", {
  g <- fill_grad_xml(grad_rect(linear_grad(c(20, 0), c(120, 0), red_blue(a1 = 0.25, a2 = 1))))
  expect_near(g$stops$alpha[[1]], 0.25 * 100000, tol = 1)
  # An opaque stop writes no a:alpha.
  expect_true(is.na(g$stops$alpha[[2]]))
})

test_that("a gradient whose stops are all one colour is a solid fill", {
  stops <- list(grad_stop(0, "#FF0000", 0.6), grad_stop(1, "#FF0000", 0.6))
  sp <- shape_sppr(the_wsp(generic_doc(list(grad_rect(linear_grad(c(20, 0), c(120, 0), stops))))))
  expect_null(grad_of(sp))
  f <- paint_of(sp)
  expect_equal(f$kind, "solid")
  expect_equal(f$color, "FF0000")
  expect_near(f$alpha, 0.6 * 100000, tol = 1)
})

test_that("scaling the drawing to the page does not change the angle or the stops", {
  p1 <- c(20, 10)
  p2 <- c(120, 60)
  r <- generic_render(list(grad_rect(linear_grad(p1, p2, red_blue()))))
  for (ms in c(1, Inf)) {
    g <- grad_of(shape_sppr(the_wsp(dml_doc(build_diagram_xml(r, max_scale = ms)$xml))))
    expect_near(g$ang, lin_ang(p1, p2), tol = 1, info = format(ms))
    expect_near(g$stops$pos, lin_pos(gbox, p1, p2, c(0, 1)), tol = 1, info = format(ms))
  }
})

# -- Strokes -------------------------------------------------------------------

test_that("a gradient stroke is the line's a:gradFill", {
  p1 <- c(20, 10)
  p2 <- c(120, 60)
  doc <- generic_doc(list(prim_path("M20 10 L120 60", fill = NULL, stroke_w = 3,
                                    stroke_grad = linear_grad(p1, p2, red_blue(0.5, 0.5)))))
  node <- the_wsp(doc)
  ln <- xfirst(node, "./wps:spPr/a:ln")
  expect_false(is_missing_node(ln))
  expect_true(is_missing_node(xfirst(ln, "./a:solidFill")))
  g <- grad_of(ln)
  expect_false(is.null(g))
  expect_equal(g$stops$color, c("FF0000", "0000FF"))
  expect_near(g$stops$alpha, c(50000, 50000), tol = 1)
  expect_near(g$ang, lin_ang(p1, p2), tol = 1)
  expect_near(g$stops$pos, lin_pos(c(20, 10, 100, 50), p1, p2, c(0, 1)), tol = 1)
  expect_null(grad_of(shape_sppr(node)))
  expect_equal(xnum(ln, "w"), 3 * 9525)
  expect_length(schema_order_problems(doc), 0L)
})

# -- Flipped items ---------------------------------------------------------------
#
# Word mirrors a shape's fill with the shape, so a flipped item's gradient is
# written in its own unflipped frame: flip_h mirrors the x direction (an
# angle theta becomes 180 - theta), flip_v the y direction (theta becomes
# -theta), both together turn it by 180. Drawn, it then runs p1 -> p2 as in
# the SVG. Mirroring leaves the stop positions as they were.
flipped_ang <- function(p1, p2, flip_h, flip_v) {
  theta <- atan2(p2[2] - p1[2], p2[1] - p1[1]) * 180 / pi
  if (flip_h) theta <- 180 - theta
  if (flip_v) theta <- -theta
  (theta %% 360) * 60000
}

test_that("a gradient stroke on a flipped line still starts at the stroke's start", {
  # Each line runs from p1 to p2, and so does its gradient.
  cases <- list(
    right_to_left = list(c(120, 10), c(20, 60)),   # flip_h
    bottom_to_top = list(c(20, 60), c(120, 10)),   # flip_v
    both = list(c(120, 60), c(20, 10))             # flip_h and flip_v
  )
  seen <- list(h = FALSE, v = FALSE)
  for (nm in names(cases)) {
    p1 <- cases[[nm]][[1]]
    p2 <- cases[[nm]][[2]]
    prim <- prim_path(sprintf("M%s %s L%s %s", p1[1], p1[2], p2[1], p2[2]), fill = NULL,
                      stroke_w = 3, stroke_grad = linear_grad(p1, p2, red_blue()))
    it <- generic_items(list(prim))
    expect_equal(it$kind, "line", info = nm)
    seen$h <- seen$h || it$flip_h
    seen$v <- seen$v || it$flip_v
    g <- grad_of(xfirst(the_wsp(generic_doc(list(prim))), "./wps:spPr/a:ln"))
    expect_false(is.null(g), info = nm)
    expect_near(g$ang, flipped_ang(p1, p2, it$flip_h, it$flip_v), tol = 1, info = nm)
    expect_near(g$stops$pos, lin_pos(c(20, 10, 100, 50), p1, p2, c(0, 1)), tol = 1, info = nm)
    expect_equal(g$stops$color, c("FF0000", "0000FF"), info = nm)
  }
  # The cases really are flipped both ways.
  expect_true(seen$h)
  expect_true(seen$v)
  expect_equal(flipped_ang(c(120, 10), c(20, 60), TRUE, FALSE),
               ((180 - atan2(50, -100) * 180 / pi) %% 360) * 60000)
})

test_that("a gradient fill on a horizontally flipped node is mirrored the same way", {
  # A lean-left node is Word's parallelogram flipped horizontally.
  p1 <- c(20, 0)
  p2 <- c(120, 0)
  shape <- prim_poly(list(c(20, 20), c(100, 20), c(120, 70), c(40, 70)),
                     fill_grad = linear_grad(p1, p2, red_blue()))
  r <- one_node_render("lean_left", shape)
  node <- item_named(build_scene(r)$items, "mermaid:node:X")
  expect_true(node$flip_h)
  g <- grad_of(shape_sppr(shape_named(dml_doc(build_diagram_xml(r)$xml), "mermaid:node:X")))
  expect_false(is.null(g))
  expect_near(g$ang, flipped_ang(p1, p2, node$flip_h, node$flip_v), tol = 1)
  expect_near(g$ang, 180 * 60000, tol = 1)
  expect_near(g$stops$pos, lin_pos(c(20, 20, 100, 50), p1, p2, c(0, 1)), tol = 1)
})

# -- Radial gradients ------------------------------------------------------------

test_that("a radial gradient is a:path circle with its centre as a:fillToRect", {
  # fillToRect insets, from each side, as shares of the box (100000 = all).
  for (centre in list(c(70, 35), c(45, 22.5))) {
    doc <- generic_doc(list(grad_rect(radial_grad(centre, 40, red_blue()))))
    g <- grad_of(shape_sppr(the_wsp(doc)))
    expect_false(is.null(g))
    expect_equal(g$path, "circle")
    expect_true(is.na(g$ang))
    expect_equal(g$stops$color, c("FF0000", "0000FF"))
    expected <- c(l = (centre[1] - gbox[1]) / gbox[3], t = (centre[2] - gbox[2]) / gbox[4],
                  r = (gbox[1] + gbox[3] - centre[1]) / gbox[3],
                  b = (gbox[2] + gbox[4] - centre[2]) / gbox[4]) * 100000
    expect_near(g$fill_to_rect, expected, tol = 1)
    expect_length(schema_order_problems(doc), 0L)
  }
})

# -- Paired pieces -----------------------------------------------------------------

test_that("a gradient fill under a flat outline: one shape, the gradient stays the fill", {
  # Mermaid draws many boxes twice: a fill-only piece under a stroke-only one.
  g <- linear_grad(c(20, 0), c(120, 0), red_blue())
  prims <- list(grad_rect(g, stroke = NULL),
                prim_rect(gbox[1], gbox[2], gbox[3], gbox[4], fill = NULL, stroke = "#00AA00"))
  items <- generic_items(prims)
  expect_equal(nrow(items), 1L)
  expect_false(is.null(items$fill_grad[[1]]))
  expect_null(items$stroke_grad[[1]])
  expect_equal(items$stroke, "00AA00")

  node <- the_wsp(generic_doc(prims))
  expect_false(is.null(grad_of(shape_sppr(node))))
  ln <- shape_ln(node)
  expect_equal(ln$kind, "solid")
  expect_equal(ln$color, "00AA00")
})

test_that("a flat fill under a gradient outline: the gradient stays the line's", {
  g <- linear_grad(c(20, 0), c(120, 0), red_blue())
  prims <- list(prim_rect(gbox[1], gbox[2], gbox[3], gbox[4], fill = "#FFFF00", stroke = NULL),
                prim_rect(gbox[1], gbox[2], gbox[3], gbox[4], fill = NULL, stroke_grad = g))
  items <- generic_items(prims)
  expect_equal(nrow(items), 1L)
  expect_null(items$fill_grad[[1]])
  expect_false(is.null(items$stroke_grad[[1]]))
  expect_equal(items$fill, "FFFF00")

  node <- the_wsp(generic_doc(prims))
  expect_equal(shape_fill(node)[c("kind", "color")], list(kind = "solid", color = "FFFF00"))
  expect_null(grad_of(shape_sppr(node)))
  expect_false(is.null(grad_of(xfirst(node, "./wps:spPr/a:ln"))))
})

test_that("a flowchart node filled with a gradient keeps it on its one shape", {
  g <- linear_grad(c(20, 0), c(120, 0), red_blue())
  r <- one_node_render("squareRect", prim_rect(20, 20, 100, 50, fill_grad = g))
  items <- build_scene(r)$items
  node <- item_named(items, "mermaid:node:X")
  expect_equal(node$prst, "rect")
  expect_false(is.null(node$fill_grad[[1]]))
  doc <- dml_doc(build_diagram_xml(r)$xml)
  gx <- grad_of(shape_sppr(shape_named(doc, "mermaid:node:X")))
  expect_equal(gx$stops$color, c("FF0000", "0000FF"))
  expect_near(gx$stops$pos, lin_pos(c(20, 20, 100, 50), c(20, 0), c(120, 0), c(0, 1)), tol = 1)
})

# -- sankey-beta: link bands stroked with gradients ------------------------------

test_that("sankey-beta: its four link bands carry their gradients as strokes", {
  prims <- purrr::keep(fixture_prims("sankey-beta"), function(p) !is.null(p$stroke_grad))
  expect_length(prims, 4L)
  items <- fixture_scene("sankey-beta")$items
  has_sg <- !purrr::map_lgl(items$stroke_grad, is.null)
  expect_equal(sum(has_sg), 4L)
  expect_true(all(items$kind[has_sg] %in% c("line", "shape")))
  expect_true(all(purrr::map_lgl(items$fill_grad[has_sg], is.null)))
  expect_true(all(is.na(items$fill[has_sg])))
})

test_that("sankey-beta XML: each band's a:ln is a horizontal gradient of its two colours", {
  bands <- purrr::keep(fixture_prims("sankey-beta"), function(p) !is.null(p$stroke_grad))
  expected <- purrr::map_chr(bands, function(p) {
    paste(sub("^#", "", purrr::map_chr(p$stroke_grad$stops, "c")), collapse = ">")
  })
  doc <- fixture_doc("sankey-beta")
  lns <- xfind(doc, "//wps:wsp/wps:spPr/a:ln[a:gradFill]")
  expect_length(lns, 4L)
  got <- purrr::map(as.list(lns), grad_of)
  expect_setequal(purrr::map_chr(got, function(g) paste(g$stops$color, collapse = ">")),
                  expected)
  for (g in got) {
    # The vector runs left to right across each band's whole width.
    expect_near(g$ang, 0, tol = 1)
    expect_near(g$stops$pos, c(0, 100000), tol = 100)
    expect_near(g$stops$alpha, c(50000, 50000), tol = 1)
  }
  expect_length(xfind(doc, "//wps:wsp/wps:spPr/a:gradFill"), 0L)
  expect_length(schema_order_problems(doc), 0L)
})

test_that("diagrams without gradients write no a:gradFill", {
  for (name in setdiff(all_fixture_names(), "sankey-beta")) {
    expect_length(xfind(fixture_doc(name), "//a:gradFill"), 0L)
  }
})

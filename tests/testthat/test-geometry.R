# Path geometry (R/geometry.R): matrices, parsing SVG path data, arcs as
# cubics, transforms, bounding boxes, closure and rectangle recognition.

# Compare a bbox / rect result ignoring any names.
expect_box <- function(object, expected, tolerance = 1e-6) {
  expect_equal(unname(as.numeric(object)), expected, tolerance = tolerance)
}

pp <- function(d) path_parse(d)

# ---- matrices ------------------------------------------------------------

test_that("mat_apply applies x' = a x + c y + e, y' = b x + d y + f", {
  m <- c(2, 1, 3, 4, 10, 20)
  # x' = 2*1 + 3*2 + 10 = 18 ; y' = 1*1 + 4*2 + 20 = 29
  r <- mat_apply(m, 1, 2)
  expect_equal(r$x, 18)
  expect_equal(r$y, 29)
})

test_that("mat_apply is vectorised", {
  r <- mat_apply(c(1, 0, 0, 1, 5, -5), c(0, 1, 2), c(0, 1, 2))
  expect_equal(r$x, c(5, 6, 7))
  expect_equal(r$y, c(-5, -4, -3))
})

test_that("mat_scale is sqrt(abs(det))", {
  expect_equal(mat_scale(c(1, 0, 0, 1, 7, 8)), 1)
  expect_equal(mat_scale(c(2, 0, 0, 8, 0, 0)), 4)    # sqrt(16)
  expect_equal(mat_scale(c(-2, 0, 0, 2, 0, 0)), 2)   # flip: abs(det)
  expect_equal(mat_scale(c(0, 3, -3, 0, 0, 0)), 3)   # rotation x3: det = 9
})

test_that("mat_rotation gives degrees clockwise via atan2(b, a)", {
  expect_equal(mat_rotation(c(1, 0, 0, 1, 0, 0)), 0)
  expect_equal(mat_rotation(c(0, 1, -1, 0, 0, 0)), 90)
  expect_equal(mat_rotation(c(0, -1, 1, 0, 0, 0)), -90)
  r <- sqrt(2) / 2
  expect_equal(mat_rotation(c(r, r, -r, r, 0, 0)), 45, tolerance = 1e-9)
})

test_that("mat_is_upright accepts translate, scale and flips", {
  expect_true(mat_is_upright(c(1, 0, 0, 1, 5, 6)))
  expect_true(mat_is_upright(c(2, 0, 0, 3, 0, 0)))
  expect_true(mat_is_upright(c(-1, 0, 0, 1, 0, 0)))
  expect_true(mat_is_upright(c(1, 0, 0, -1, 0, 0)))
  expect_true(mat_is_upright(c(1, 1e-9, -1e-9, 1, 0, 0)))
})

test_that("mat_is_upright rejects rotation and skew", {
  expect_false(mat_is_upright(c(0, 1, -1, 0, 0, 0)))
  r <- sqrt(2) / 2
  expect_false(mat_is_upright(c(r, r, -r, r, 0, 0)))
  expect_false(mat_is_upright(c(1, 0, 0.5, 1, 0, 0)))  # horizontal skew
  expect_false(mat_is_upright(c(1, 0.5, 0, 1, 0, 0)))  # vertical skew
})

# ---- path_parse: shape of the result ------------------------------------

test_that("path_parse returns a tibble with the documented columns", {
  p <- pp("M 10 20 L 30 40")
  expect_s3_class(p, "tbl_df")
  expect_named(p, c("cmd", "x1", "y1", "x2", "y2", "x", "y"))
  expect_equal(p$cmd, c("M", "L"))
  expect_equal(p$x, c(10, 30))
  expect_equal(p$y, c(20, 40))
})

test_that("empty and garbage input give a zero-row tibble with the columns", {
  for (d in c("", "   ", "hello world", "#$%")) {
    p <- pp(d)
    expect_s3_class(p, "tbl_df")
    expect_equal(nrow(p), 0L, info = paste0("'", d, "'"))
    expect_named(p, path_cols, info = paste0("'", d, "'"))
  }
})

test_that("M, L and Z rows have NA control points; Z has NA x and y", {
  p <- pp("M0 0 L10 0 Z")
  expect_true(all(is.na(p$x1[p$cmd != "C"])))
  expect_true(all(is.na(p$y2[p$cmd != "C"])))
  expect_true(is.na(p$x[3]) && is.na(p$y[3]))
})

# ---- path_parse: commands ------------------------------------------------

test_that("relative moveto/lineto are made absolute", {
  p <- pp("m 10 10 l 5 0 l 0 5")
  expect_equal(p$cmd, c("M", "L", "L"))
  expect_equal(p$x, c(10, 15, 15))
  expect_equal(p$y, c(10, 10, 15))
})

test_that("extra coordinate pairs after M are implicit linetos", {
  p <- pp("M0 0 10 0 10 10")
  expect_equal(p$cmd, c("M", "L", "L"))
  expect_equal(p$x, c(0, 10, 10))
  expect_equal(p$y, c(0, 0, 10))
})

test_that("extra pairs after relative m are relative linetos", {
  p <- pp("m1 1 2 0 0 2")
  expect_equal(p$cmd, c("M", "L", "L"))
  expect_equal(p$x, c(1, 3, 3))   # 1, 1+2, 3+0
  expect_equal(p$y, c(1, 1, 3))   # 1, 1+0, 1+2
})

test_that("repeated coordinates after L repeat the lineto", {
  p <- pp("M0 0 L 1 1 2 2 3 3")
  expect_equal(p$cmd, c("M", "L", "L", "L"))
  expect_equal(p$x, c(0, 1, 2, 3))
})

test_that("H, V, h and v are lines", {
  p <- pp("M0 0 H10 V5 h-4 v-2")
  expect_equal(p$cmd, c("M", "L", "L", "L", "L"))
  expect_equal(p$x, c(0, 10, 10, 6, 6))
  expect_equal(p$y, c(0, 0, 5, 5, 3))
})

test_that("absolute and relative cubics", {
  p <- pp("M0 0 C1 2 3 4 5 6")
  expect_equal(p$cmd, c("M", "C"))
  expect_equal(unlist(p[2, c("x1", "y1", "x2", "y2", "x", "y")]),
               c(x1 = 1, y1 = 2, x2 = 3, y2 = 4, x = 5, y = 6))

  q <- pp("M10 10 c1 2 3 4 5 6")
  expect_equal(unlist(q[2, c("x1", "y1", "x2", "y2", "x", "y")]),
               c(x1 = 11, y1 = 12, x2 = 13, y2 = 14, x = 15, y = 16))
})

test_that("S after C reflects the previous second control point", {
  # previous c2 = (10,10), current = (10,0): reflection = 2*(10,0)-(10,10) = (10,-10)
  p <- pp("M0 0 C0 10 10 10 10 0 S 20 -10 20 0")
  expect_equal(p$cmd, c("M", "C", "C"))
  expect_equal(unlist(p[3, c("x1", "y1", "x2", "y2", "x", "y")]),
               c(x1 = 10, y1 = -10, x2 = 20, y2 = -10, x = 20, y = 0))
})

test_that("relative s after C reflects and offsets from the current point", {
  p <- pp("M0 0 C0 10 10 10 10 0 s10 -10 10 0")
  expect_equal(unlist(p[3, c("x1", "y1", "x2", "y2", "x", "y")]),
               c(x1 = 10, y1 = -10, x2 = 20, y2 = -10, x = 20, y = 0))
})

test_that("S after a non-curve uses the current point as first control", {
  p <- pp("M0 0 L10 0 S 20 10 30 0")
  expect_equal(p$cmd, c("M", "L", "C"))
  expect_equal(unlist(p[3, c("x1", "y1", "x2", "y2", "x", "y")]),
               c(x1 = 10, y1 = 0, x2 = 20, y2 = 10, x = 30, y = 0))
})

test_that("Q is raised to a cubic", {
  # P0=(0,0) Q=(30,30) P2=(60,0)
  # c1 = P0 + 2/3 (Q-P0) = (20,20); c2 = P2 + 2/3 (Q-P2) = (60-20, 0+20) = (40,20)
  p <- pp("M0 0 Q 30 30 60 0")
  expect_equal(p$cmd, c("M", "C"))
  expect_equal(unlist(p[2, c("x1", "y1", "x2", "y2", "x", "y")]),
               c(x1 = 20, y1 = 20, x2 = 40, y2 = 20, x = 60, y = 0))
})

test_that("relative q is raised to a cubic from the current point", {
  p <- pp("M10 10 q 30 30 60 0")
  # absolute Q = (40,40), P2 = (70,10); c1 = (10+20, 10+20); c2 = (70-20, 10+20)
  expect_equal(unlist(p[2, c("x1", "y1", "x2", "y2", "x", "y")]),
               c(x1 = 30, y1 = 30, x2 = 50, y2 = 30, x = 70, y = 10))
})

test_that("T after Q reflects the quadratic control point", {
  # first Q ends at (60,0) with control (30,30): reflected control = (90,-30)
  # P0=(60,0), P2=(120,0): c1 = (60+2/3*30, 0+2/3*-30) = (80,-20)
  #                        c2 = (120+2/3*(90-120), 0+2/3*-30) = (100,-20)
  p <- pp("M0 0 Q30 30 60 0 T 120 0")
  expect_equal(p$cmd, c("M", "C", "C"))
  expect_equal(unlist(p[3, c("x1", "y1", "x2", "y2", "x", "y")]),
               c(x1 = 80, y1 = -20, x2 = 100, y2 = -20, x = 120, y = 0))
})

test_that("T after a non-Q uses the current point as its control", {
  # control = P0 = (10,0): c1 = P0 = (10,0); c2 = P2 + 2/3 (P0 - P2) = 20 - 20/3
  p <- pp("M0 0 L10 0 T 20 0")
  expect_equal(p$cmd, c("M", "L", "C"))
  expect_equal(unlist(p[3, c("x1", "y1", "x2", "y2", "x", "y")]),
               c(x1 = 10, y1 = 0, x2 = 20 - 20 / 3, y2 = 0, x = 20, y = 0))
})

test_that("only M, L, C and Z appear and all coordinates are absolute", {
  d <- "m 5 5 h 10 v 10 q 5 5 10 0 t 10 0 s 5 5 10 0 a 5 5 0 0 1 10 0 z"
  p <- pp(d)
  expect_true(all(p$cmd %in% c("M", "L", "C", "Z")))
  # the path never goes below x=5 or y=5 and ends at the closed start point
  expect_equal(p$cmd[nrow(p)], "Z")
  expect_true(all(p$x[!is.na(p$x)] >= 5 - 1e-9))
})

# ---- Z behaviour -----------------------------------------------------------

test_that("Z returns the current point to the subpath start for a relative m", {
  # after Z the current point is (5,5); m 1 1 -> (6,6); l 2 0 -> (8,6)
  p <- pp("M5 5 L15 5 L15 15 Z m 1 1 l 2 0")
  expect_equal(p$cmd, c("M", "L", "L", "Z", "M", "L"))
  expect_equal(c(p$x[5], p$y[5]), c(6, 6))
  expect_equal(c(p$x[6], p$y[6]), c(8, 6))
})

test_that("a relative l straight after Z is measured from the subpath start", {
  p <- pp("M5 5 L15 5 Z l 2 0")
  expect_equal(p$cmd, c("M", "L", "Z", "L"))
  expect_equal(c(p$x[4], p$y[4]), c(7, 5))
})

test_that("a lowercase z closes like Z", {
  p <- pp("M0 0 L1 0 L1 1 z")
  expect_equal(p$cmd, c("M", "L", "L", "Z"))
})

test_that("two subpaths are both kept", {
  p <- pp("M0 0 L10 0 L10 10 Z M20 20 L30 20 L30 30 Z")
  expect_equal(p$cmd, c("M", "L", "L", "Z", "M", "L", "L", "Z"))
  expect_equal(p$x[5], 20)
})

# ---- arcs ------------------------------------------------------------------

test_that("a half-circle arc becomes cubics and bulges up for sweep 1", {
  p <- pp("M 0 0 A 10 10 0 0 1 20 0")
  expect_true(all(p$cmd %in% c("M", "C")))
  expect_gte(sum(p$cmd == "C"), 2L)   # at most 90 degrees per segment
  b <- unname(path_bbox(p))
  expect_equal(b[1], 0, tolerance = 0.01)
  expect_equal(b[2], -10, tolerance = 0.01)
  expect_equal(b[3], 20, tolerance = 0.01)
  expect_equal(b[4], 10, tolerance = 0.01)
  # arc ends exactly at the requested point
  last <- p[nrow(p), ]
  expect_equal(c(last$x, last$y), c(20, 0), tolerance = 1e-6)
})

test_that("sweep flag 0 bulges the other way (down on screen)", {
  b <- unname(path_bbox(pp("M 0 0 A 10 10 0 0 0 20 0")))
  expect_equal(b, c(0, 0, 20, 10), tolerance = 0.01)
})

test_that("two arcs make a full ellipse", {
  # centre (10,0), rx 10, ry 5
  p <- pp("M0 0 a10 5 0 1 1 20 0 a10 5 0 1 1 -20 0")
  expect_true(all(p$cmd %in% c("M", "C")))
  expect_equal(unname(path_bbox(p)), c(0, -5, 20, 10), tolerance = 0.01)
})

test_that("a relative arc is measured from the current point", {
  b <- unname(path_bbox(pp("M100 50 a 10 10 0 0 1 20 0")))
  expect_equal(b, c(100, 40, 20, 10), tolerance = 0.01)
})

test_that("an arc whose radii are too small is scaled up to reach its end", {
  # radius 1 cannot span 20: radius becomes 10 (half circle)
  b <- unname(path_bbox(pp("M0 0 A 1 1 0 0 1 20 0")))
  expect_equal(b, c(0, -10, 20, 10), tolerance = 0.01)
})

test_that("an arc with a zero radius is a straight line", {
  p <- pp("M0 0 A 0 5 0 0 1 20 10")
  expect_equal(p$cmd, c("M", "L"))
  expect_equal(c(p$x[2], p$y[2]), c(20, 10))
  q <- pp("M0 0 a 5 0 0 0 1 20 10")
  expect_equal(q$cmd, c("M", "L"))
})

test_that("an arc to its own start point draws nothing extra", {
  p <- pp("M5 5 A 10 10 0 0 1 5 5")
  expect_false("C" %in% p$cmd)
})

# ---- packed number syntax --------------------------------------------------

test_that("numbers packed with minus signs are split", {
  p <- pp("M1-2")
  expect_equal(c(p$x, p$y), c(1, -2))
  q <- pp("M1-2.5.5-3")
  expect_equal(q$cmd, c("M", "L"))
  expect_equal(c(q$x, q$y), c(1, 0.5, -2.5, -3))
})

test_that("numbers packed with leading decimal points are split", {
  p <- pp("M.5.5")
  expect_equal(c(p$x, p$y), c(0.5, 0.5))
})

test_that("exponent notation is read", {
  p <- pp("M0 0 l1e-3 2E1")
  expect_equal(c(p$x[2], p$y[2]), c(0.001, 20))
})

test_that("arc flags can be written without separators", {
  # a5,5 0 01-10,0 : rx 5 ry 5 rot 0 large 0 sweep 1 dx -10 dy 0
  p <- pp("M0 0 a5,5 0 01-10,0")
  expect_true(all(p$cmd %in% c("M", "C")))
  last <- p[nrow(p), ]
  expect_equal(c(last$x, last$y), c(-10, 0), tolerance = 1e-6)
  # sweep 1 from right to left on screen is clockwise: passes below the chord
  expect_equal(unname(path_bbox(p)), c(-10, 0, 10, 5), tolerance = 0.01)
})

test_that("commas and newlines are separators", {
  p <- pp("M 1,2\nL 3,4\n")
  expect_equal(p$x, c(1, 3))
  expect_equal(p$y, c(2, 4))
})

# ---- path_transform --------------------------------------------------------

test_that("path_transform translates every point including control points", {
  p <- pp("M0 0 C1 2 3 4 5 6 Z")
  t <- path_transform(p, c(1, 0, 0, 1, 10, 20))
  expect_equal(t$cmd, p$cmd)
  expect_equal(unlist(t[2, c("x1", "y1", "x2", "y2", "x", "y")]),
               c(x1 = 11, y1 = 22, x2 = 13, y2 = 24, x = 15, y = 26))
  expect_equal(c(t$x[1], t$y[1]), c(10, 20))
  expect_true(is.na(t$x[3]))
})

test_that("path_transform scales", {
  t <- path_transform(pp("M1 1 L2 3"), c(2, 0, 0, 3, 0, 0))
  expect_equal(t$x, c(2, 4))
  expect_equal(t$y, c(3, 9))
})

test_that("path_transform rotates by 90 degrees", {
  # matrix (0,1,-1,0): x' = -y, y' = x
  t <- path_transform(pp("M10 0 L10 5"), c(0, 1, -1, 0, 0, 0))
  expect_equal(t$x, c(0, -5))
  expect_equal(t$y, c(10, 10))
})

test_that("path_transform keeps the columns and NA control points", {
  t <- path_transform(pp("M0 0 L1 1"), c(1, 0, 0, 1, 1, 1))
  expect_named(t, path_cols)
  expect_true(all(is.na(t$x1)))
})

# ---- path_bbox -------------------------------------------------------------

test_that("path_bbox of straight segments is the point extent", {
  expect_box(path_bbox(pp("M10 20 L30 5 L25 40")), c(10, 5, 20, 35))
})

test_that("path_bbox of a cubic is the true extent, not the control hull", {
  # y(t) = 3t(1-t)^2*100 + 3t^2(1-t)*100 = 300 t (1-t), maximum 75 at t = 0.5
  b <- unname(path_bbox(pp("M0,0 C0,100 100,100 100,0")))
  expect_equal(b, c(0, 0, 100, 75), tolerance = 1e-3)
})

test_that("path_bbox includes extrema of cubics in both directions", {
  # symmetric S-ish cubic: x peaks inside (0..100) at 75 beyond end x=0
  b <- unname(path_bbox(pp("M0,0 C100,0 100,100 0,100")))
  # x(t) = 300 t (1-t)^2 *... max of 3t(1-t)^2*100 + 3t^2(1-t)*100 = 300 t(1-t) = 75
  expect_equal(b, c(0, 0, 75, 100), tolerance = 1e-3)
})

test_that("path_bbox covers several subpaths", {
  b <- unname(path_bbox(pp("M0 0 L10 10 M50 -5 L60 0")))
  expect_equal(b, c(0, -5, 60, 15))
})

test_that("path_bbox of the cylinder in flowchart-paint matches the browser bbox", {
  render <- fixture_render("flowchart-paint")
  is_cyl <- function(p) {
    p$kind == "path" && any(purrr::map_lgl(p$anc, ~ identical(.x$id, "flowchart-C-3"))) &&
      grepl("a40.25", p$geom$d, fixed = TRUE)
  }
  prim <- purrr::detect(render$prims, is_cyl)
  expect_false(is.null(prim))
  p <- path_transform(path_parse(prim$geom$d), unlist(prim$m))
  expect_equal(unname(path_bbox(p)), unlist(prim$bbox), tolerance = 0.05)
})

# ---- path_closed -----------------------------------------------------------

test_that("path_closed is TRUE when every subpath ends with Z", {
  expect_true(path_closed(pp("M0 0 L1 0 L1 1 Z")))
  expect_true(path_closed(pp("M0 0 L1 0 L1 1 Z M5 5 L6 5 L6 6 Z")))
})

test_that("path_closed is FALSE when a subpath is open", {
  expect_false(path_closed(pp("M0 0 L1 0 L1 1")))
  expect_false(path_closed(pp("M0 0 L1 0 L1 1 Z M5 5 L6 5")))
})

test_that("path_closed is TRUE when the path returns to its start point", {
  expect_true(path_closed(pp("M0 0 L1 0 L1 1 L0 0")))
})

# ---- path_as_rect ----------------------------------------------------------

test_that("path_as_rect recognises a clockwise rectangle with H/V", {
  expect_box(path_as_rect(pp("M10 20 H110 V70 H10 Z")), c(10, 20, 100, 50))
})

test_that("path_as_rect recognises an anticlockwise rectangle", {
  expect_box(path_as_rect(pp("M10 20 V70 H110 V20 Z")), c(10, 20, 100, 50))
})

test_that("path_as_rect recognises corners given with L", {
  expect_box(path_as_rect(pp("M0 0 L40 0 L40 30 L0 30 Z")), c(0, 0, 40, 30))
})

test_that("path_as_rect accepts the start point repeated before Z", {
  expect_box(path_as_rect(pp("M10 20 L110 20 L110 70 L10 70 L10 20 Z")),
             c(10, 20, 100, 50))
})

test_that("path_as_rect accepts a rectangle that starts at another corner", {
  expect_box(path_as_rect(pp("M110 70 L10 70 L10 20 L110 20 Z")),
             c(10, 20, 100, 50))
})

test_that("path_as_rect works on a poly_path rectangle", {
  r <- poly_path(list(c(0, 0), c(8, 0), c(8, 4), c(0, 4)), closed = TRUE)
  expect_box(path_as_rect(r), c(0, 0, 8, 4))
})

test_that("path_as_rect rejects a rotated quadrilateral", {
  expect_null(path_as_rect(pp("M50 0 L100 50 L50 100 L0 50 Z")))
})

test_that("path_as_rect rejects a slanted quadrilateral", {
  expect_null(path_as_rect(pp("M0 0 L100 0 L110 50 L10 50 Z")))
})

test_that("path_as_rect rejects a five-point polygon", {
  expect_null(path_as_rect(pp("M0 0 L100 0 L100 50 L50 80 L0 50 Z")))
})

test_that("path_as_rect rejects an open path", {
  expect_null(path_as_rect(pp("M0 0 L10 0 L10 10 L0 10")))
})

test_that("path_as_rect rejects a path with a curve", {
  expect_null(path_as_rect(pp("M0 0 H10 C10 5 10 5 10 10 H0 Z")))
})

test_that("path_as_rect rejects two subpaths", {
  expect_null(path_as_rect(pp("M0 0 H10 V10 H0 Z M20 0 H30 V10 H20 Z")))
})

test_that("path_as_rect honours its tolerance", {
  slightly_off <- pp("M0 0 L100 0 L100.005 50 L0 50 Z")
  expect_false(is.null(path_as_rect(slightly_off, tol = 0.01)))
  expect_null(path_as_rect(pp("M0 0 L100 0 L100.5 50 L0 50 Z"), tol = 0.01))
})

# ---- poly_path -------------------------------------------------------------

test_that("poly_path builds M, L... and a Z when closed", {
  p <- poly_path(list(c(0, 0), c(10, 0), c(10, 5)), closed = TRUE)
  expect_named(p, path_cols)
  expect_equal(p$cmd, c("M", "L", "L", "Z"))
  expect_equal(p$x[1:3], c(0, 10, 10))
  expect_equal(p$y[1:3], c(0, 0, 5))
})

test_that("poly_path leaves an open polyline without Z", {
  p <- poly_path(list(c(0, 0), c(10, 0), c(10, 5)), closed = FALSE)
  expect_equal(p$cmd, c("M", "L", "L"))
  expect_false(path_closed(p))
})

test_that("poly_path accepts points as parsed from the bridge JSON", {
  # jsonlite::fromJSON(simplifyVector = FALSE) gives a list of list(x, y)
  pts <- jsonlite::fromJSON("[[1,2],[3,4],[5,6]]", simplifyVector = FALSE)
  p <- poly_path(pts, closed = FALSE)
  expect_equal(p$x, c(1, 3, 5))
  expect_equal(p$y, c(2, 4, 6))
})

# -- path_bbox and bare movetos -------------------------------------------------

test_that("a trailing or leading bare moveto does not count in the bounding box", {
  expect_equal(path_bbox(path_parse("M10 10 L50 10 L50 50 M 190 190")), c(10, 10, 40, 40))
  expect_equal(path_bbox(path_parse("M 5 5 M10 10 L50 10 L50 50")), c(10, 10, 40, 40))
  expect_equal(path_bbox(path_parse("M10 10 L50 10 L50 50 M 190 190 M 200 200")),
               c(10, 10, 40, 40))
})

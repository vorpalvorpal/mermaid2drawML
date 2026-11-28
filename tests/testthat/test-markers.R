# Line ends (R/markers.R): which Word end each Mermaid marker gives, how the
# path end moves so that Word's end sits where Mermaid's marker was, the
# size classes, and the markers drawn as shapes of their own (those Word
# has no end for, and centred ends too big for Word's largest), grouped
# with their line. Hand-built renders with `markers` (helper-scene.R:
# marker_def() and Mermaid's own definitions).

mk_items <- function(prims, markers = NULL, viewBox = c(0, 0, 200, 100)) {
  generic_items(prims, viewBox = viewBox, markers = markers)
}

# The points of a custom path, Z rows dropped, as a 2-column matrix.
path_points <- function(path) {
  p <- path[path$cmd != "Z", ]
  cbind(p$x, p$y)
}

# The same for a (possibly missing) marker shape row.
marker_points <- function(s) {
  if (nrow(s) == 0L || is.null(s$path[[1]])) return(matrix(numeric(), 0L, 2L))
  path_points(s$path[[1]])
}

a_line <- function(..., x1 = 10, y1 = 50, x2 = 110, y2 = 50, stroke = "#333333", stroke_w = 1) {
  prim_line(x1, y1, x2, y2, stroke = stroke, stroke_w = stroke_w, ...)
}

# -- The marker table -----------------------------------------------------------

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
  # With no marker definition (as here) an end that is drawn as a shape when
  # it has one falls back to its nearest filled Word end.
  expected <- c(
    "flowchart-v2-pointEnd" = "triangle",
    "flowchart-v2-pointStart" = "triangle",
    "flowchart-v2-pointEnd__f00" = "triangle",
    "block-pointEnd" = "triangle",
    "arrowhead" = "triangle",
    "filled-head" = "stealth",
    "stateDiagram-barbEnd" = "stealth",
    "class-extensionEnd" = "triangle",
    "class-extensionStart" = "triangle",
    "requirement-requirement_arrowEnd" = "arrow",
    "class-dependencyEnd" = "arrow",
    "my-stickEnd" = "arrow",
    "flowchart-v2-circleEnd" = "oval",
    "class-lollipopStart" = "oval",
    "flowchart-v2-crossEnd" = "none",              # Word has no X end
    "crosshead" = "none",
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

test_that("the marker table has a logical as_shape column", {
  expect_true("as_shape" %in% names(.line_end_table))
  expect_type(.line_end_table$as_shape, "logical")
})

test_that("without definitions, each marker id gives its Word end or its fallback", {
  expected <- c(
    "sequencenumber" = "none",
    "er-onlyOneStart" = "none", "er-zeroOrOneEnd" = "none", "er-oneOrMoreEnd" = "none",
    "er-zeroOrMoreStart" = "none", "er-mdParentStart" = "none",
    "flowchart-v2-crossEnd" = "none", "crosshead" = "none",
    "class-extensionEnd" = "triangle",          # hollow: nearest filled end
    "class-aggregationStart" = "diamond",       # hollow: nearest filled end
    "class-compositionStart" = "diamond",
    "flowchart-v2-circleEnd" = "oval", "class-lollipopEnd" = "oval",
    "class-dependencyEnd" = "arrow", "stickTopArrowHead" = "arrow",
    "requirement-requirement_arrowEnd" = "arrow",
    "filled-head" = "stealth", "stateDiagram-barbEnd" = "stealth",
    "flowchart-v2-pointEnd" = "triangle", "arrowhead" = "triangle",
    "solidTopArrowHead" = "triangle", "something-unheard-of" = "triangle",
    # requirement_contains has no fallback: its table end
    "requirement-requirement_containsStart" = "none"
  )
  for (id in names(expected)) {
    items <- mk_items(list(a_line(marker_end = id)))
    expect_equal(lines_of(items)$tail, expected[[id]], info = id)
    expect_equal(nrow(marker_rows(items)), 0L, info = id)
  }
})

hollow_diamond <- function() {
  marker_def(NULL, refX = 18, refY = 7, w = 190, h = 240, units = "strokeWidth",
             prims = list(prim_path("M 18,7 L9,13 L1,7 L9,1 Z", bbox = c(1, 1, 17, 12),
                                    fill = NULL, stroke = "#333333")))
}

hollow_circle <- function() {
  marker_def(NULL, refX = 1, refY = 7, w = 190, h = 240, units = "strokeWidth",
             prims = list(prim_ellipse(7, 7, 6, fill = NULL, stroke = "#333333")))
}

contains_marker <- function() {
  marker_def(NULL, refX = 0, refY = 10, w = 20, h = 20, units = "strokeWidth",
             prims = list(prim_ellipse(10, 10, 9, fill = NULL, stroke = "#333333"),
                          prim_line(1, 10, 19, 10, stroke = "#333333"),
                          prim_line(10, 1, 10, 19, stroke = "#333333")))
}

test_that("with a definition, a marker Word cannot draw leaves no Word end and becomes shapes", {
  defs <- list(
    "class-extensionEnd" = marker_extension_end(),
    "class-aggregationEnd" = hollow_diamond(),
    "flowchart-v2-crossEnd" = marker_cross_end(),
    "class-lollipopEnd" = hollow_circle(),
    "er-oneOrMoreEnd" = marker_er_one_or_more_end(),
    "requirement-requirement_containsEnd" = contains_marker()
  )
  for (id in names(defs)) {
    items <- mk_items(list(a_line(marker_end = id)), markers = defs[id])
    expect_equal(lines_of(items)$tail, "none", info = id)
    expect_gte(nrow(markers_named(items, id)), 1L, label = paste(id, "marker shapes"))
  }
})

test_that("with a definition, Word-drawable ends stay Word ends and add no shapes", {
  # An oval or diamond end longer than 5 x the stroke is drawn as a shape,
  # so this uses a 4 px line and a userSpaceOnUse diamond (17 <= 20;
  # with strokeWidth units it would be 68 px).
  filled_diamond <- marker_composition_start()
  filled_diamond$units <- "userSpaceOnUse"
  defs <- list(
    "class-compositionEnd" = list("diamond", filled_diamond),
    "flowchart-v2-circleEnd" = list("oval", marker_circle_end()),
    "flowchart-v2-pointEnd" = list("triangle", marker_arrow_point()),
    "class-dependencyEnd" = list("arrow", marker_arrow_point()),
    "filled-head" = list("stealth", marker_arrow_point())
  )
  for (id in names(defs)) {
    items <- mk_items(list(a_line(stroke_w = 4, marker_end = id)),
                      markers = stats::setNames(list(defs[[id]][[2]]), id))
    expect_equal(lines_of(items)$tail, defs[[id]][[1]], info = id)
    expect_equal(nrow(marker_rows(items)), 0L, info = id)
  }
})

test_that("an unknown id is a triangle when its marker is one filled polygon, else shapes", {
  one_poly <- list("my-tip" = marker_arrow_point())
  items <- mk_items(list(a_line(marker_end = "my-tip")), markers = one_poly)
  expect_equal(lines_of(items)$tail, "triangle")
  expect_equal(nrow(marker_rows(items)), 0L)

  open_stroke <- list("my-tip" = marker_def(c(0, 0, 10, 10), refX = 5, refY = 5, w = 8, h = 8,
                                            units = "userSpaceOnUse",
                                            prims = list(prim_path("M0,0 L10,5 L0,10", fill = NULL,
                                                                   stroke = "#333333"))))
  items <- mk_items(list(a_line(marker_end = "my-tip")), markers = open_stroke)
  expect_equal(lines_of(items)$tail, "none")
  expect_equal(nrow(markers_named(items, "my-tip")), 1L)

  two <- marker_arrow_point()
  two$prims <- c(two$prims, list(marker_prim(prim_ellipse(5, 5, 2, fill = "#333333"))))
  items <- mk_items(list(a_line(marker_end = "my-tip")), markers = list("my-tip" = two))
  expect_equal(lines_of(items)$tail, "none")
  expect_equal(nrow(markers_named(items, "my-tip")), 2L)
})

# -- The end moves to the marker ------------------------------------------------

test_that("an arrow point end moves the line end to the marker's tip (+4 px for the flowchart point)", {
  # viewBox 0..10 in 8 px: scale 0.8; tip at content x 10, refX 5 -> +4 px
  items <- mk_items(list(a_line(marker_end = "flowchart-v2-pointEnd")),
                    markers = list("flowchart-v2-pointEnd" = marker_arrow_point()))
  ln <- lines_of(items)
  expect_equal(ln$tail, "triangle")
  e <- item_ends(ln)
  expect_near(e$start, c(10, 50), tol = 1e-9)
  expect_near(e$end, c(114, 50), tol = 1e-9)
})

test_that("the tip offset follows the end tangent: vertical and reversed lines", {
  m <- list("flowchart-v2-pointEnd" = marker_arrow_point())
  down <- lines_of(mk_items(list(a_line(x1 = 50, y1 = 10, x2 = 50, y2 = 90,
                                        marker_end = "flowchart-v2-pointEnd")), m))
  expect_near(item_ends(down)$end, c(50, 94), tol = 1e-9)
  left <- lines_of(mk_items(list(a_line(x1 = 110, x2 = 10, marker_end = "flowchart-v2-pointEnd")), m))
  expect_near(item_ends(left)$end, c(6, 50), tol = 1e-9)
  expect_true(left$flip_h)
})

test_that("an oval end is centred on the marker: the circle end shortens the line", {
  # circleEnd: viewBox 0..10 in 11 px (scale 1.1), centre 5, refX 11:
  # (5 - 11) x 1.1 = -6.6 px.
  # The scale is the uniform viewBox scale of Mermaid's 11 px circleEnd. On a
  # 3 px line (11 <= 15) it is a Word end; at 1 px the circle is a shape.
  items <- mk_items(list(a_line(stroke_w = 3, marker_end = "flowchart-v2-circleEnd")),
                    markers = list("flowchart-v2-circleEnd" = marker_circle_end()))
  ln <- lines_of(items)
  expect_equal(ln$tail, "oval")
  expect_near(item_ends(ln)$end, c(103.4, 50), tol = 1e-9)
})

test_that("a start marker with plain auto: offset back toward the node is refX - min (tip kinds)", {
  # pointStart: refX 4.5, content 0..10, scale 0.8 -> 3.6 px
  items <- mk_items(list(a_line(marker_start = "flowchart-v2-pointStart")),
                    markers = list("flowchart-v2-pointStart" = marker_arrow_point_start()))
  ln <- lines_of(items)
  expect_equal(ln$head, "triangle")
  expect_near(item_ends(ln)$start, c(6.4, 50), tol = 1e-9)
  expect_near(item_ends(ln)$end, c(110, 50), tol = 1e-9)
})

test_that("a start marker with plain auto, centred kind: refX - centre (class composition)", {
  # content x 1..18 (centre 9.5), refX 18, scale 1: 8.5 px
  # userSpaceOnUse units on a 4 px line (17 <= 20); Mermaid's own
  # strokeWidth diamond is always longer than 5 x the stroke, so a shape.
  def <- marker_composition_start()
  def$units <- "userSpaceOnUse"
  items <- mk_items(list(a_line(x1 = 30, x2 = 130, stroke_w = 4,
                                marker_start = "class-compositionStart")),
                    markers = list("class-compositionStart" = def))
  ln <- lines_of(items)
  expect_equal(ln$head, "diamond")
  expect_near(item_ends(ln)$start, c(21.5, 50), tol = 1e-9)
})

test_that("a start marker with auto-start-reverse is mirrored: the end formula applies", {
  def <- marker_arrow_point(orient = "auto-start-reverse")
  items <- mk_items(list(a_line(marker_start = "arrowhead")), markers = list(arrowhead = def))
  expect_near(item_ends(lines_of(items))$start, c(6, 50), tol = 1e-9)
})

test_that("a marker without a viewBox maps content at the units scale alone (sequence arrowhead)", {
  # content x -1..10, refX 7.9, userSpaceOnUse: +2.1 px; w, h are only a clip box
  def <- marker_def(NULL, refX = 7.9, refY = 5, w = 12, h = 12, units = "userSpaceOnUse",
                    orient = "auto-start-reverse",
                    prims = list(prim_path("M -1 0 L 10 5 L 0 10 z", bbox = c(-1, 0, 11, 10),
                                           fill = "#333333", stroke = "#333333")))
  items <- mk_items(list(a_line(x1 = 82, y1 = 111, x2 = 271, y2 = 111, stroke_w = 1.5,
                                marker_end = "arrowhead")),
                    markers = list(arrowhead = def), viewBox = c(0, 0, 400, 200))
  ln <- lines_of(items)
  expect_near(item_ends(ln)$end, c(273.1, 111), tol = 1e-9)
  # length 11 / 1.5 = 7.3 and width 10 / 1.5 = 6.7: both large
  expect_equal(ln$tail_len, "lg")
  expect_equal(ln$tail_w, "lg")
})

test_that("strokeWidth units scale the marker by the stroke width", {
  # 3 x 3 at stroke 2 = 6 px for viewBox 0..10: content scale 0.6, offset 3 px
  def <- marker_arrow_point(w = 3, units = "strokeWidth")
  items <- mk_items(list(a_line(stroke_w = 2, marker_end = "flowchart-v2-pointEnd")),
                    markers = list("flowchart-v2-pointEnd" = def))
  ln <- lines_of(items)
  expect_near(item_ends(ln)$end, c(113, 50), tol = 1e-9)
  expect_equal(ln$tail_len, "sm")    # 6 / 2 = 3
  expect_equal(ln$tail_w, "sm")
})

test_that("an offset of 0.25 px or less leaves the end alone; a larger negative one shortens", {
  end_x <- function(refX) {
    def <- marker_arrow_point(refX = refX)
    ln <- lines_of(mk_items(list(a_line(marker_end = "flowchart-v2-pointEnd")),
                            markers = list("flowchart-v2-pointEnd" = def)))
    item_ends(ln)$end[[1]]
  }
  expect_equal(end_x(9.7), 110)              # +0.24
  expect_equal(end_x(10.3), 110)             # -0.24
  expect_near(end_x(9.6), 110.32, tol = 1e-9)
  expect_near(end_x(10.5), 109.6, tol = 1e-9)
})

test_that("the last segment of a path moves; a cubic's second control point moves with it", {
  m <- list("flowchart-v2-pointEnd" = marker_arrow_point())
  bent <- mk_items(list(prim_path("M10,50 L60,50 L60,90", fill = NULL, stroke = "#333333",
                                  marker_end = "flowchart-v2-pointEnd")), m)
  p <- bent$path[[1]]
  expect_near(c(p$x[[3]], p$y[[3]]), c(60, 94), tol = 1e-9)
  expect_near(c(p$x[[2]], p$y[[2]]), c(60, 50), tol = 1e-9)
  expect_near(c(p$x[[1]], p$y[[1]]), c(10, 50), tol = 1e-9)

  curve <- mk_items(list(prim_path("M10,10 C50,10 90,50 90,90", bbox = c(10, 10, 80, 80),
                                   fill = NULL, stroke = "#333333",
                                   marker_end = "flowchart-v2-pointEnd")), m)
  p <- curve$path[[1]]
  last <- p[nrow(p), ]
  expect_equal(last$cmd, "C")
  expect_near(c(last$x, last$y), c(90, 94), tol = 1e-9)
  expect_near(c(last$x2, last$y2), c(90, 54), tol = 1e-9)
  expect_near(c(last$x1, last$y1), c(50, 10), tol = 1e-9)
})

test_that("without a definition there is no offset and the size is med", {
  ln <- lines_of(mk_items(list(a_line(marker_end = "flowchart-v2-pointEnd"))))
  expect_near(item_ends(ln)$end, c(110, 50), tol = 1e-9)
  expect_equal(ln$tail_len, "med")
  expect_equal(ln$tail_w, "med")
})

test_that("a repeated end point does not change the end frame", {
  plain <- path_end_frame(path_parse("M0 0 L100 100"), "end")
  expect_equal(path_end_frame(path_parse("M0 0 L100 100 L100 100"), "end"), plain)
  expect_equal(path_end_frame(path_parse("M0 0 L100 100 L100 100 L100 100"), "end"), plain)
  expect_equal(plain$dir, rep(sqrt(0.5), 2))
})

test_that("a cubic whose end control point is its end points along its other control", {
  f <- path_end_frame(path_parse("M0 0 L50 0 C80 10 100 20 100 20"), "end")
  expect_equal(f$point, c(100, 20))
  expect_equal(f$dir, c(20, 10) / sqrt(500))
  # both control points on the end: back to the previous end point
  g <- path_end_frame(path_parse("M0 0 L50 0 C100 20 100 20 100 20"), "end")
  expect_equal(g$dir, c(50, 20) / sqrt(2900))
})

test_that("a zero-length first segment does not change the start frame", {
  plain <- path_end_frame(path_parse("M0 0 L100 0"), "start")
  expect_equal(path_end_frame(path_parse("M0 0 L0 0 L100 0"), "start"), plain)
  expect_equal(plain$dir, c(1, 0))
})

# The line item(s) and marker shapes of one drawn path, for comparing two paths.
degenerate_pair <- function(d_plain, d_repeated, marker_id, def) {
  draw <- function(d) {
    mk_items <- generic_items(
      list(prim_path(d, fill = NULL, stroke = "#333333", marker_end = marker_id)),
      viewBox = c(0, 0, 200, 100), markers = stats::setNames(list(def), marker_id))
    mk_items[mk_items$kind != "group", ]
  }
  list(plain = draw(d_plain), repeated = draw(d_repeated))
}

test_that("a path that repeats its end point gets the same end and marker as the plain path", {
  defs <- list("class-extensionEnd" = marker_extension_end(),
               "flowchart-v2-pointEnd" = marker_arrow_point())
  for (id in names(defs)) {
    pair <- degenerate_pair("M10,10 L60,20 L110,70", "M10,10 L60,20 L110,70 L110,70",
                            id, defs[[id]])
    is_marker <- function(items) !is.na(items$name) & startsWith(items$name, "mermaid:marker:")
    expect_equal(sum(is_marker(pair$repeated)), sum(is_marker(pair$plain)), info = id)
    box <- function(items) unlist(items[!is_marker(items), c("x", "y", "w", "h")][1, ])
    expect_equal(box(pair$repeated), box(pair$plain), info = id)
    pts <- function(items) {
      m <- items[is_marker(items), ]
      if (nrow(m) == 0L) return(matrix(numeric(), 0L, 2L))
      unique(round(do.call(rbind, lapply(m$path, function(p) outline_points(p))), 6))
    }
    expect_equal(pts(pair$repeated), pts(pair$plain), info = id)
    expect_equal(pair$repeated$tail[!is_marker(pair$repeated)][[1]],
                 pair$plain$tail[!is_marker(pair$plain)][[1]], info = id)
  }
})

# -- Sizes ----------------------------------------------------------------------

test_that("every item has head_len, head_w, tail_len and tail_w, default med", {
  items <- mk_items(list(prim_rect(10, 10, 40, 20), a_line()))
  for (col in c("head_len", "head_w", "tail_len", "tail_w")) {
    expect_true(col %in% names(items), info = col)
    expect_equal(items[[col]], c("med", "med"), info = col)
  }
})

test_that("size classes: r < 3.5 sm, r <= 6 med, above lg", {
  size_of <- function(L, stroke_w = 1) {
    ln <- lines_of(mk_items(list(a_line(stroke_w = stroke_w, marker_end = "flowchart-v2-pointEnd")),
                            markers = list("flowchart-v2-pointEnd" = marker_arrow_point(w = L))))
    c(ln$tail_len, ln$tail_w)
  }
  expect_equal(size_of(3.49), c("sm", "sm"))
  expect_equal(size_of(3.5), c("med", "med"))
  expect_equal(size_of(6), c("med", "med"))
  expect_equal(size_of(6.01), c("lg", "lg"))
  # thin lines count as 1 px: 5 / max(0.5, 1) = 5
  expect_equal(size_of(5, stroke_w = 0.5), c("med", "med"))
})

test_that("length along the path and width across it are classed separately", {
  def <- marker_def(c(0, 0, 10, 4), refX = 5, refY = 2, w = 10, h = 4, units = "userSpaceOnUse",
                    prims = list(prim_path("M0 0 L10 2 L0 4 z", fill = "#333333")))
  size_of <- function(stroke_w) {
    ln <- lines_of(mk_items(list(a_line(stroke_w = stroke_w, marker_end = "my-pointEnd")),
                            markers = list("my-pointEnd" = def)))
    c(len = ln$tail_len, w = ln$tail_w)
  }
  expect_equal(size_of(1), c(len = "lg", w = "med"))    # 10 and 4
  expect_equal(size_of(2), c(len = "med", w = "sm"))    # 5 and 2
})

test_that("a start marker sets head_len and head_w, leaving the tail at med", {
  ln <- lines_of(mk_items(list(a_line(marker_start = "flowchart-v2-pointStart")),
                          markers = list("flowchart-v2-pointStart" = marker_arrow_point_start())))
  expect_equal(c(ln$head_len, ln$head_w), c("lg", "lg"))   # 8 px at 1 px
  expect_equal(c(ln$tail_len, ln$tail_w), c("med", "med"))
})

# -- Markers drawn as shapes ----------------------------------------------------

# The items of one marked line, with the definition `def` for marker `id`.
marked_items <- function(line, def = marker_extension_end(), id = "class-extensionEnd",
                         viewBox = c(0, 0, 200, 200)) {
  mk_items(list(line), markers = stats::setNames(list(def), id), viewBox = viewBox)
}

test_that("a marker shape is instanced at a horizontal end: refX, refY on the end", {
  items <- marked_items(a_line(marker_end = "class-extensionEnd"))
  s <- markers_named(items, "class-extensionEnd")
  expect_equal(nrow(s), 1L)
  expect_equal(s$kind, "shape")
  expect_near(c(s$x, s$y, s$w, s$h), c(110, 44, 17, 12), tol = 0.01)
  expect_near(marker_points(s), rbind(c(110, 44), c(110, 56), c(127, 50)), tol = 0.01)
  expect_true(is.na(s$fill))                  # hollow
  expect_equal(s$stroke, "333333")
  expect_equal(s$stroke_w, 1)
  expect_equal(lines_of(items)$tail, "none")
  expect_near(item_ends(lines_of(items))$end, c(110, 50), tol = 1e-9)   # no Word-end offset
})

test_that("a marker shape is rotated to a vertical end", {
  s <- markers_named(marked_items(a_line(x1 = 50, y1 = 10, x2 = 50, y2 = 90,
                                      marker_end = "class-extensionEnd")), "class-extensionEnd")
  expect_equal(nrow(s), 1L)
  expect_near(c(s$x, s$y, s$w, s$h), c(44, 90, 12, 17), tol = 0.01)
  expect_near(marker_points(s), rbind(c(56, 90), c(44, 90), c(50, 107)), tol = 0.01)
})

test_that("a marker shape is rotated to a 45 degree end", {
  s <- markers_named(marked_items(a_line(x1 = 10, y1 = 10, x2 = 80, y2 = 80,
                                      marker_end = "class-extensionEnd")), "class-extensionEnd")
  expect_equal(nrow(s), 1L)
  r <- 6 / sqrt(2)
  t <- 17 / sqrt(2)
  expect_near(marker_points(s),
              rbind(c(80 + r, 80 - r), c(80 - r, 80 + r), c(80 + t, 80 + t)), tol = 0.01)
  expect_near(c(s$x, s$y, s$w, s$h), c(80 - r, 80 - r, r + t, r + t), tol = 0.01)
})

test_that("at a start, plain auto keeps the path direction and auto-start-reverse turns it round", {
  line <- a_line(x1 = 30, x2 = 130, marker_start = "my-extensionStart")
  auto <- markers_named(marked_items(line, marker_extension_end(orient = "auto"), "my-extensionStart"),
                        "my-extensionStart")
  expect_near(c(auto$x, auto$y, auto$w, auto$h), c(30, 44, 17, 12), tol = 0.01)
  rev <- markers_named(marked_items(line, marker_extension_end(orient = "auto-start-reverse"),
                                 "my-extensionStart"), "my-extensionStart")
  expect_near(c(rev$x, rev$y, rev$w, rev$h), c(13, 44, 17, 12), tol = 0.01)
  expect_near(marker_points(rev), rbind(c(30, 56), c(30, 44), c(13, 50)), tol = 0.01)
})

test_that("Mermaid's class extensionStart (refX 18, plain auto) points back at the start node", {
  s <- markers_named(marked_items(a_line(x1 = 30, x2 = 130, marker_start = "class-extensionStart"),
                               marker_extension_start(), "class-extensionStart"),
                     "class-extensionStart")
  expect_near(c(s$x, s$y, s$w, s$h), c(13, 44, 17, 12), tol = 0.01)
  expect_near(marker_points(s), rbind(c(13, 50), c(30, 56), c(30, 44)), tol = 0.01)
})

test_that("a numeric orient is used as is", {
  def <- marker_extension_end()
  def$orient <- 90
  s <- markers_named(marked_items(a_line(marker_end = "class-extensionEnd"), def), "class-extensionEnd")
  expect_near(c(s$x, s$y, s$w, s$h), c(104, 50, 12, 17), tol = 0.01)
  expect_near(marker_points(s), rbind(c(116, 50), c(104, 50), c(110, 67)), tol = 0.01)
})

test_that("a marker shape scales with strokeWidth units, its stroke width too", {
  s <- markers_named(marked_items(a_line(stroke_w = 2, marker_end = "class-extensionEnd")),
                     "class-extensionEnd")
  expect_near(c(s$x, s$y, s$w, s$h), c(110, 38, 34, 24), tol = 0.01)
  expect_equal(s$stroke_w, 2)
})

test_that("context paint takes the edge's stroke colour", {
  def <- marker_extension_end(fill = "context", stroke = "context")
  s <- markers_named(marked_items(a_line(stroke = "#FF0000", marker_end = "class-extensionEnd"), def),
                     "class-extensionEnd")
  expect_equal(s$fill, "FF0000")
  expect_equal(s$stroke, "FF0000")
})

test_that("the flowchart circle end (11 px) on a 1 px line is a marker shape, the end not moved", {
  id <- "flowchart-v2-circleEnd"
  items <- marked_items(a_line(id = "msg", marker_end = id), marker_circle_end(), id)
  ln <- lines_of(items)
  expect_equal(nrow(ln), 1L)
  expect_equal(ln$tail, "none")
  expect_near(item_ends(ln)$end, c(110, 50), tol = 1e-9)
  s <- markers_named(items, id)
  expect_equal(nrow(s), 1L)
  if (nrow(s) != 1L) return()
  expect_equal(s$kind, "shape")
  expect_equal(s$prst, "ellipse")
  # content 0..10 at scale 1.1, refX 11 on the end: 97.9 .. 108.9
  expect_near(box_of(s), c(97.9, 44.5, 11, 11), tol = 0.01)
  expect_equal(s$fill, "333333")
  # grouped with its line
  expect_false(is.na(s$parent))
  expect_equal(ln$parent, s$parent)
})

test_that("on a line under 1 px the limit is still 5 px", {
  id <- "flowchart-v2-circleEnd"
  items <- marked_items(a_line(id = "msg", stroke_w = 0.5, marker_end = id), marker_circle_end(), id)
  expect_equal(lines_of(items)$tail, "none")
  expect_equal(nrow(markers_named(items, id)), 1L)
})

test_that("the same circle on a 3 px line (11 <= 15) stays a Word oval, the end at its centre", {
  # userSpaceOnUse: the marker is 11 px whatever the stroke
  id <- "flowchart-v2-circleEnd"
  items <- marked_items(a_line(id = "msg", stroke_w = 3, marker_end = id), marker_circle_end(), id)
  ln <- lines_of(items)
  expect_equal(ln$tail, "oval")
  expect_near(item_ends(ln)$end, c(103.4, 50), tol = 1e-9)    # (5 - 11) x 1.1
  expect_equal(nrow(marker_rows(items)), 0L)
})

test_that("the limit is 5 x the stroke width: 10.5 < 11 is a shape, 11.5 is a Word end", {
  id <- "flowchart-v2-circleEnd"
  thin <- marked_items(a_line(id = "msg", stroke_w = 2.1, marker_end = id), marker_circle_end(), id)
  expect_equal(lines_of(thin)$tail, "none")
  expect_equal(nrow(markers_named(thin, id)), 1L)
  thick <- marked_items(a_line(id = "msg", stroke_w = 2.3, marker_end = id), marker_circle_end(), id)
  expect_equal(lines_of(thick)$tail, "oval")
  expect_equal(nrow(marker_rows(thick)), 0L)
})

test_that("the class composition diamond (17 px) on a 1 px line is a shape at the start", {
  id <- "class-compositionStart"
  items <- marked_items(a_line(id = "msg", x1 = 30, x2 = 130, marker_start = id), marker_composition_start(), id)
  ln <- lines_of(items)
  expect_equal(ln$head, "none")
  expect_near(item_ends(ln)$start, c(30, 50), tol = 1e-9)
  s <- markers_named(items, id)
  expect_equal(nrow(s), 1L)
  if (nrow(s) != 1L) return()
  # plain auto at a start: content x 1..18 runs back from refX 18 on the end
  expect_near(box_of(s), c(13, 44, 17, 12), tol = 0.01)
  expect_equal(s$fill, "333333")
})

test_that("with strokeWidth units the diamond grows with the line: a shape on a 3 px line too", {
  # 17 x 3 = 51 > 5 x 3 = 15
  id <- "class-compositionStart"
  items <- marked_items(a_line(id = "msg", x1 = 80, x2 = 180, stroke_w = 3, marker_start = id),
                        marker_composition_start(), id)
  expect_equal(lines_of(items)$head, "none")
  s <- markers_named(items, id)
  expect_equal(nrow(s), 1L)
  if (nrow(s) != 1L) return()
  expect_near(box_of(s), c(29, 32, 51, 36), tol = 0.01)
  expect_equal(s$stroke_w, 3)
})

test_that("a composition diamond in userSpaceOnUse units on a 4 px line stays a Word diamond", {
  # 17 <= 20; centre offset back toward the node: refX - centre = 18 - 9.5
  id <- "class-compositionStart"
  def <- marker_composition_start()
  def$units <- "userSpaceOnUse"
  items <- marked_items(a_line(id = "msg", x1 = 30, x2 = 130, stroke_w = 4, marker_start = id), def, id)
  ln <- lines_of(items)
  expect_equal(ln$head, "diamond")
  expect_near(item_ends(ln)$start, c(21.5, 50), tol = 1e-9)
  expect_equal(nrow(marker_rows(items)), 0L)
})

test_that("a big triangle, stealth or arrow end stays a Word end", {
  # 40 px on a 1 px line: viewBox 0..10 at scale 4, tip 5 beyond refX -> +20 px
  def <- marker_arrow_point(w = 40)
  for (case in list(c("flowchart-v2-pointEnd", "triangle"), c("filled-head", "stealth"),
                    c("class-dependencyEnd", "arrow"))) {
    items <- marked_items(a_line(id = "msg", marker_end = case[[1]]), def, case[[1]])
    ln <- lines_of(items)
    expect_equal(ln$tail, case[[2]], info = case[[1]])
    expect_near(item_ends(ln)$end, c(130, 50), tol = 1e-9, info = case[[1]])
    expect_equal(nrow(marker_rows(items)), 0L, info = case[[1]])
  }
})

test_that("a marker part that is a line prim is an item of kind shape", {
  id <- "requirement-requirement_containsEnd"
  def <- marker_def(NULL, refX = 0, refY = 10, w = 20, h = 20, units = "strokeWidth",
                    prims = list(prim_ellipse(10, 10, 9, fill = NULL, stroke = "#333333"),
                                 prim_line(1, 10, 19, 10, stroke = "#333333"),
                                 prim_line(10, 1, 10, 19, stroke = "#333333")))
  items <- marked_items(a_line(id = "msg", marker_end = id), def, id)
  s <- markers_named(items, id)
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
  g <- groups_of(items)
  expect_equal(nrow(g), 1L)
  # The group is named after the line item: mermaid:edge_group:<item name
  # without "mermaid:">.
  expect_equal(g$name, "mermaid:edge_group:line:dot")
  kids <- children_of(items, g$name)
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

# -- A line and its marker shapes are one group ---------------------------------

test_that("a line and its marker shapes are grouped, the shapes right after the line", {
  items <- mk_items(list(a_line(id = "msg", marker_end = "class-extensionEnd"),
                         prim_rect(150, 10, 30, 30, fill = "#00AA00")),
                    markers = list("class-extensionEnd" = marker_extension_end()),
                    viewBox = c(0, 0, 200, 200))
  g <- items[items$kind == "group", ]
  expect_equal(nrow(g), 1L)
  # The group is named after the line item: any mermaid:edge_group: name
  # that carries the line's id.
  expect_match(g$name, "^mermaid:edge_group:.*msg")
  kids <- items[!is.na(items$parent) & items$parent %in% g$name, ]
  expect_equal(kids$kind, c("line", "shape"))
  expect_equal(kids$name[2], "mermaid:marker:class-extensionEnd")
  expect_equal(which(items$kind == "group"), 1L)     # before the green rect
  expect_true(is.na(items$parent[items$fill %in% "00AA00"]))
})

test_that("an edge unit and its marker shapes are grouped as mermaid:edge_group:<start>-<end>", {
  nodes <- c(node_prims("flowchart-a-0", prim_rect(10, 10, 60, 30)),
             node_prims("flowchart-b-1", prim_rect(130, 10, 60, 30)))
  edge <- prim_edge("M70,25 L118,25", "L_a_b_0", marker_end = "flowchart-v2-crossEnd")
  db <- mini_db(list(db_node("a", dom_id = "flowchart-a-0"), db_node("b", dom_id = "flowchart-b-1")),
                list(db_edge("a", "b")))
  items <- build_scene(mini_render(c(list(edge), nodes), db = db,
                                   markers = list("flowchart-v2-crossEnd" = marker_cross_end())))$items
  g <- item_named(items, "mermaid:edge_group:a-b")
  expect_equal(nrow(g), 1L)
  expect_equal(g$kind, "group")
  kids <- items[!is.na(items$parent) & items$parent == "mermaid:edge_group:a-b", ]
  expect_equal(kids$name, c("mermaid:edge:a-b", "mermaid:marker:flowchart-v2-crossEnd"))
  expect_equal(item_named(items, "mermaid:edge:a-b")$tail, "none")
})

test_that("both ends drawn as shapes: one group with the line and both markers", {
  items <- mk_items(list(a_line(id = "rel", marker_start = "er-onlyOneStart",
                                marker_end = "er-oneOrMoreEnd", x1 = 50, x2 = 150)),
                    markers = list("er-onlyOneStart" = marker_er_only_one_start(),
                                   "er-oneOrMoreEnd" = marker_er_one_or_more_end()),
                    viewBox = c(0, 0, 220, 100))
  g <- items[items$kind == "group", ]
  expect_equal(nrow(g), 1L)
  kids <- items[!is.na(items$parent) & items$parent %in% g$name, ]
  expect_equal(kids$kind[1], "line")
  expect_setequal(kids$name[-1], c("mermaid:marker:er-onlyOneStart", "mermaid:marker:er-oneOrMoreEnd"))
  ln <- lines_of(items)
  expect_equal(c(ln$head, ln$tail), c("none", "none"))
})

test_that("a zero-length marked line is only its marker shape, which takes the text inside it", {
  # Mermaid's sequence number: a 0 x 0 line with stroke width 2, the
  # sequencenumber circle (r 6, strokeWidth units -> r 12) and a white "1".
  prims <- list(
    prim_line(75, 111, 75, 111, stroke = "#DACEF3", stroke_w = 2, marker_start = "sequencenumber"),
    prim_text("1", x = 71.7, y = 104, w = 6.7, h = 14, html = FALSE, color = "#FFFFFF")
  )
  items <- mk_items(prims, markers = list(sequencenumber = marker_seq_number()),
                    viewBox = c(0, 0, 200, 200))
  expect_equal(nrow(items), 1L)
  expect_equal(sum(items$kind %in% c("line", "group", "text")), 0L)
  s <- markers_named(items, "sequencenumber")
  expect_equal(nrow(s), 1L)
  expect_equal(s$prst, "ellipse")
  expect_near(c(s$x, s$y, s$w, s$h), c(63, 99, 24, 24), tol = 0.01)
  expect_equal(s$fill, "333333")
  expect_identical(s$label, "1")
})

# -- In the XML -----------------------------------------------------------------

end_ln <- function(prim, markers = NULL) {
  doc <- dml_doc(build_diagram_xml(generic_render(list(prim), markers = markers))$xml)
  shape_ln(all_wsp(doc)[[1]])
}

test_that("a:tailEnd and a:headEnd take w and len from the item's size classes", {
  m <- list("flowchart-v2-pointEnd" = marker_arrow_point())
  ln <- end_ln(a_line(marker_end = "flowchart-v2-pointEnd"), m)
  expect_equal(ln$tail, "triangle")
  expect_equal(c(ln$tail_w, ln$tail_len), c("lg", "lg"))       # 8 px at 1 px
  ln <- end_ln(a_line(stroke_w = 3.5, marker_end = "flowchart-v2-pointEnd"), m)
  expect_equal(c(ln$tail_w, ln$tail_len), c("sm", "sm"))       # 8 / 3.5 = 2.3

  ln <- end_ln(a_line(marker_start = "flowchart-v2-pointStart"),
               list("flowchart-v2-pointStart" = marker_arrow_point_start()))
  expect_equal(c(ln$head_w, ln$head_len), c("lg", "lg"))

  def <- marker_def(c(0, 0, 10, 4), refX = 5, refY = 2, w = 10, h = 4, units = "userSpaceOnUse",
                    prims = list(prim_path("M0 0 L10 2 L0 4 z", fill = "#333333")))
  ln <- end_ln(a_line(marker_end = "my-pointEnd"), list("my-pointEnd" = def))
  expect_equal(c(w = ln$tail_w, len = ln$tail_len), c(w = "med", len = "lg"))
})

test_that("without a definition the ends are written w and len med", {
  ln <- end_ln(a_line(marker_end = "flowchart-v2-pointEnd"))
  expect_equal(c(ln$tail_w, ln$tail_len), c("med", "med"))
})

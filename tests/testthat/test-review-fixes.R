# Regression tests for the second review's bugs. Hand-built renders and
# mocked bridge calls (helper-xml.R, helper-render.R, test-render.R's idioms).

# ---- a mocked bridge (the same idiom as test-render.R) ---------------------------

pie_stdout <- function() {
  paste0("[", paste(readLines(fixture_path("pie"), warn = FALSE), collapse = "\n"), "]")
}

# Mock node/browser lookup and processx::run; returns an environment that
# records the call.
local_bridge <- function(result = NULL, env = parent.frame()) {
  result <- result %||% list(status = 0L, stdout = pie_stdout(), stderr = "",
                             timeout = FALSE)
  st <- rlang::env(call = NULL)
  node_dir <- withr::local_tempdir(.local_envir = env)
  testthat::local_mocked_bindings(
    check_mermaid2drawml = function() invisible(TRUE),
    .find_browser = function(...) NA_character_,
    .node_dir = function() node_dir,
    .env = env
  )
  testthat::local_mocked_bindings(
    run = function(command, args = character(), ...) {
      st$call <- list(command = command, args = args, ...)
      result
    },
    .package = "processx",
    .env = env
  )
  st
}

# ---- a shape that holds text is never flipped vertically --------------------

has_text <- function(items) !purrr::map_lgl(items$text, is.null)

# The distinct corner points of an item's custom path (Z rows dropped).
outline_points <- function(path) {
  p <- path[path$cmd != "Z", ]
  unique(round(cbind(x = p$x, y = p$y), 6))
}

sort_points <- function(m) m[order(m[, 1], m[, 2]), , drop = FALSE]

test_that("no item with text is flipped vertically in any fixture", {
  for (nm in all_fixture_names()) {
    items <- fixture_scene(nm)$items
    bad <- has_text(items) & items$flip_v
    expect_false(any(bad), info = nm)
  }
})

test_that("an inverted trapezoid is drawn from its outline, upright, with its text", {
  inv <- list(c(30, 60), c(110, 60), c(120, 20), c(20, 20))
  for (shape in c("inv_trapezoid", "trap-t", "manual", "trapezoid-top", "inv-trapezoid")) {
    r <- one_node_render(shape, prim_poly(inv))
    items <- build_scene(r)$items
    it <- items[has_text(items), ]
    expect_equal(nrow(it), 1L, info = shape)
    expect_false(it$flip_v, info = shape)
    expect_true(is.na(it$prst), info = shape)
    path <- it$path[[1]]
    expect_false(is.null(path), info = shape)
    if (is.null(path)) next
    expect_equal(sort_points(outline_points(path)),
                 sort_points(do.call(rbind, lapply(inv, function(p) c(x = p[[1]], y = p[[2]])))),
                 ignore_attr = TRUE, info = shape)
  }
})

test_that("in the XML no vertically flipped shape holds text", {
  inv <- list(c(30, 60), c(110, 60), c(120, 20), c(20, 20))
  doc <- dml_doc(build_diagram_xml(one_node_render("inv_trapezoid", prim_poly(inv)))$xml)
  expect_length(xfind(doc, "//wps:wsp[wps:txbx]/wps:spPr/a:xfrm[@flipV]"), 0L)
  node <- shape_named(doc, "mermaid:node:X")
  expect_true(shape_is_custom(node))
  expect_equal(shape_paragraph_texts(node), "Label")

  full <- fixture_doc("flowchart-full")
  expect_length(xfind(full, "//wps:wsp[wps:txbx]/wps:spPr/a:xfrm[@flipV]"), 0L)
})

test_that("a horizontally flipped registry shape still uses its preset", {
  left <- list(c(30, 60), c(120, 60), c(110, 20), c(20, 20))
  doc <- dml_doc(build_diagram_xml(one_node_render("lean_left", prim_poly(left)))$xml)
  node <- shape_named(doc, "mermaid:node:X")
  expect_equal(shape_prst(node), "parallelogram")
  expect_true(shape_xfrm(node)$flipH)
  expect_false(shape_xfrm(node)$flipV)
})

# ---- timeout and failure kinds ----------------------------------------------

failure_stdout <- function(error) {
  as.character(jsonlite::toJSON(list(list(ok = FALSE, error = error)), auto_unbox = TRUE))
}

mock_failure <- function(error, env = parent.frame()) {
  local_bridge(list(status = 0L, stdout = failure_stdout(error), stderr = "",
                         timeout = FALSE), env = env)
}

test_that("a bridge failure of kind timeout is a mermaid2drawml_timeout", {
  mock_failure(list(kind = "timeout", message = "page took too long"))
  expect_error(render_mermaid("flowchart TD\n A --> B"),
               class = "mermaid2drawml_timeout")
})

test_that("a bridge failure of kind bridge is a render error, not a syntax error", {
  mock_failure(list(kind = "bridge", message = "Chrome crashed"))
  cnd <- tryCatch(render_mermaid("flowchart TD\n A --> B"), error = function(e) e)
  expect_s3_class(cnd, "mermaid2drawml_render_error")
  expect_false(inherits(cnd, "mermaid2drawml_syntax_error"))
})

test_that("a syntax failure, or one with no kind (an old bridge), is a syntax error", {
  mock_failure(list(kind = "syntax", message = "Parse error on line 2"))
  expect_error(render_mermaid("flowchart TD\n A -->"),
               class = "mermaid2drawml_syntax_error", regexp = "Parse error")
  mock_failure(list(message = "Parse error on line 2"))
  expect_error(render_mermaid("flowchart TD\n A -->"),
               class = "mermaid2drawml_syntax_error", regexp = "Parse error")
})

# The value given with --timeout, in milliseconds.
timeout_arg <- function(st) {
  i <- match("--timeout", st$call$args)
  if (is.na(i)) NA_real_ else as.numeric(st$call$args[[i + 1L]])
}

test_that("the bridge's in-page timeout is shorter than the whole-render timeout", {
  st <- local_bridge()
  render_mermaid("flowchart TD\n A --> B", timeout = 120)
  ms <- timeout_arg(st)
  expect_false(is.na(ms))
  expect_gt(ms, 0)
  expect_lt(ms, 120000)
  expect_gte(ms, 5000)
})

test_that("a very short timeout still leaves the bridge at least five seconds", {
  st <- local_bridge()
  render_mermaid("flowchart TD\n A --> B", timeout = 3)
  expect_gte(timeout_arg(st), 5000)
})

# ---- the direction at a path end ---------------------------------------------

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
    is_marker <- function(items) startsWith(dplyr::coalesce(items$name, ""), "mermaid:marker:")
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

# ---- large and non-finite coordinates ----------------------------------------

test_that("coordinates beyond the 32-bit EMU range are written in full", {
  far <- 300000L
  prims <- list(
    prim_rect(far, far, 80, 40),
    prim_path(sprintf("M%d,%d L%d,%d L%d,%d Z", far, far, far + 60, far, far + 30, far + 50),
              fill = "#FF0000")
  )
  # The canvas must hold the shapes (items off the canvas are dropped) and the
  # page must be big enough that the diagram is not shrunk.
  doc <- generic_doc(prims, viewBox = c(0, 0, far + 200, far + 100),
                     page_width_in = 1e4, page_height_in = 1e4)
  offs <- xfind(doc, "//a:off")
  exts <- xfind(doc, "//a:ext")
  expect_gt(length(offs), 0L)
  expect_false(anyNA(xattr(offs, "x")))
  expect_false(anyNA(xattr(offs, "y")))
  expect_false(anyNA(xattr(exts, "cx")))
  expect_false(anyNA(xattr(exts, "cy")))
  vals <- xml2::xml_text(xfind(doc, "//@*"))
  expect_false(any(vals == "NA"))
  whole <- c(xattr(offs, "x"), xattr(offs, "y"), xattr(exts, "cx"), xattr(exts, "cy"))
  expect_true(all(grepl("^-?[0-9]+$", whole)))
  # the far shape is where it should be: 3e5 px at 9525 EMU per px
  expect_true(any(abs(as.numeric(xattr(offs, "x")) - far * 9525) < 10))
})

test_that("a non-finite coordinate reaching the emitter is a render error", {
  scene <- build_scene(generic_render(list(prim_rect(10, 10, 80, 40))))
  for (bad in c(Inf, NaN, NA_real_)) {
    s <- scene
    s$items$x[[1]] <- bad
    expect_error(build_diagram_xml(s), class = "mermaid2drawml_render_error",
                 info = as.character(bad))
  }
})

# ---- tiny figures are not ellipses -------------------------------------------

cubic_ellipse_path <- function(cx, cy, r, k = 0.5523) {
  sprintf(paste("M%s,%s C%s,%s %s,%s %s,%s C%s,%s %s,%s %s,%s",
                "C%s,%s %s,%s %s,%s C%s,%s %s,%s %s,%s Z"),
          cx + r, cy,
          cx + r, cy + k * r, cx + k * r, cy + r, cx, cy + r,
          cx - k * r, cy + r, cx - r, cy + k * r, cx - r, cy,
          cx - r, cy - k * r, cx - k * r, cy - r, cx, cy - r,
          cx + k * r, cy - r, cx + r, cy - k * r, cx + r, cy)
}

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
  expect_equal(path_prst(cubic_ellipse_path(1.5, 1.5, 1.5), 3), "ellipse")
  arcs <- paste("M3,1.5 A1.5,1.5 0 0 1 1.5,3 A1.5,1.5 0 0 1 0,1.5",
                "A1.5,1.5 0 0 1 1.5,0 A1.5,1.5 0 0 1 3,1.5 Z")
  expect_equal(path_prst(arcs, 3), "ellipse")
})

test_that("a 20 px diamond stays custom and a 100 px circle is an ellipse", {
  expect_true(is.na(path_prst("M10,0 L20,10 L10,20 L0,10 Z", 20)))
  expect_equal(path_prst(cubic_ellipse_path(50, 50, 50), 100), "ellipse")
})

# ---- rounded rectangles ------------------------------------------------------

test_that("a rect is rounded only when both radii are positive", {
  expect_equal(generic_items(list(prim_rect(10, 20, 100, 40, rx = 8, ry = 0)))$prst, "rect")
  expect_equal(generic_items(list(prim_rect(10, 20, 100, 40, rx = 8, ry = 6)))$prst, "roundRect")
})

# ---- bounding boxes ignore bare movetos --------------------------------------

test_that("a trailing or leading bare moveto does not count in the bounding box", {
  expect_equal(path_bbox(path_parse("M10 10 L50 10 L50 50 M 190 190")), c(10, 10, 40, 40))
  expect_equal(path_bbox(path_parse("M 5 5 M10 10 L50 10 L50 50")), c(10, 10, 40, 40))
  expect_equal(path_bbox(path_parse("M10 10 L50 10 L50 50 M 190 190 M 200 200")),
               c(10, 10, 40, 40))
})

# ---- white space in attribute values -----------------------------------------

test_that("line breaks and tabs in alt text survive a round trip through XML", {
  alt <- "first line\nsecond line\twith a tab"
  r <- generic_render(list(prim_rect(10, 10, 50, 30)))
  doc <- dml_doc(build_diagram_xml(r, alt_text = alt)$xml)
  expect_equal(xattr(doc_pr(doc), "descr"), alt)
})

# ---- mermaid_config ------------------------------------------------------------

test_that("an empty config sends no --config", {
  for (cfg in list(NULL, list())) {
    st <- local_bridge()
    render_mermaid("flowchart TD\n A --> B", config = cfg)
    expect_false("--config" %in% st$call$args)
  }
})

# Run render_mermaid() with a mocked bridge, keeping the --config file's text.
config_text <- function(config) {
  st <- local_bridge()
  seen <- rlang::env(text = NULL)
  testthat::local_mocked_bindings(
    run = function(command, args = character(), ...) {
      i <- match("--config", args)
      seen$text <- paste(readLines(args[[i + 1L]], warn = FALSE, encoding = "UTF-8"),
                         collapse = "")
      list(status = 0L, stdout = pie_stdout(), stderr = "", timeout = FALSE)
    },
    .package = "processx"
  )
  render_mermaid("flowchart TD\n A --> B", config = config)
  seen$text
}

test_that("config writes null members as null, numbers in full and scalars as scalars", {
  txt <- config_text(list(theme = "dark", flowchart = list(padding = NULL),
                          themeVariables = list(fontSize = 16.123456789)))
  expect_match(txt, "\"theme\":\\s*\"dark\"")
  expect_match(txt, "\"padding\":\\s*null")
  expect_match(txt, "16.123456789", fixed = TRUE)
  parsed <- jsonlite::fromJSON(txt, simplifyVector = FALSE)
  expect_identical(parsed$theme, "dark")
  expect_true("padding" %in% names(parsed$flowchart))
  expect_null(parsed$flowchart$padding)
  expect_equal(parsed$themeVariables$fontSize, 16.123456789, tolerance = 1e-12)
})

# ---- UTF-8 output ------------------------------------------------------------

test_that("the bridge's output is read as UTF-8", {
  st <- local_bridge()
  render_mermaid("flowchart TD\n A --> B")
  expect_equal(st$call$encoding, "UTF-8")
})

test_that("non-ASCII label text comes back intact", {
  label <- "Zürich — 東京"
  raw <- list(ok = TRUE, bridge = 2L, mermaid = "11.13.0", type = "flowchart-v2",
              role = "flowchart-v2", viewBox = list(0, 0, 200, 100),
              prims = list(c(list(i = 0L), prim_text(label, 10, 10, 80, 20))), db = NULL)
  out <- paste0("[", jsonlite::toJSON(raw, auto_unbox = TRUE, null = "null", digits = NA), "]")
  local_bridge(list(status = 0L, stdout = out, stderr = "", timeout = FALSE))
  r <- render_mermaid("flowchart TD\n A --> B")
  expect_equal(r$prims[[1]]$rows[[1]]$runs[[1]]$t, label)
})

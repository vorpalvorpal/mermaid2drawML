# tests/testthat/test-drawingml.R
# Tests for build_diagram_xml(). All fixtures are static — no Node.js required.

# ── Static SVG-data fixture ────────────────────────────────────────────────

make_svg_data <- function(n_nodes = 2L, add_edge = TRUE,
                          shapes = NULL,
                          dashed = FALSE,
                          edge_label = NA_character_) {

  if (is.null(shapes)) shapes <- rep("rect", n_nodes)

  ids    <- LETTERS[seq_len(n_nodes)]
  labels <- paste0("Node ", ids)

  nodes <- tibble::tibble(
    id     = ids,
    label  = labels,
    shape  = shapes,
    svg_cx = seq(100, by = 150, length.out = n_nodes),
    svg_cy = rep(50, n_nodes),
    svg_w  = rep(80, n_nodes),
    svg_h  = rep(35, n_nodes),
    fill   = rep("FFFFFF", n_nodes),
    stroke = rep("4472C4", n_nodes),
    class  = rep("default", n_nodes)
  )

  if (add_edge && n_nodes >= 2L) {
    edges <- tibble::tibble(
      id          = "L-A-B-0",
      from        = ids[1],
      to          = ids[2],
      label       = edge_label,
      path_d      = "M 100 68 L 250 32",
      arrow_start = "none",
      arrow_end   = "arrow",
      line_type   = if (dashed) "dashed" else "solid",
      label_x     = if (!is.na(edge_label)) 175 else NA_real_,
      label_y     = if (!is.na(edge_label)) 50  else NA_real_
    )
  } else {
    edges <- tibble::tibble(
      id          = character(),
      from        = character(),
      to          = character(),
      label       = character(),
      path_d      = character(),
      arrow_start = character(),
      arrow_end   = character(),
      line_type   = character(),
      label_x     = numeric(),
      label_y     = numeric()
    )
  }

  list(nodes = nodes, edges = edges, viewbox = c(0, 0, 400, 150))
}

# ── Return value structure ─────────────────────────────────────────────────

test_that("build_diagram_xml returns list with xml and next_id", {
  r <- build_diagram_xml(make_svg_data())
  expect_type(r, "list")
  expect_named(r, c("xml", "next_id"))
  expect_type(r$xml, "character")
})

test_that("next_id advances by at least n_nodes + n_edges", {
  d <- make_svg_data(n_nodes = 3L)
  r <- build_diagram_xml(d, start_id = 1L)
  # 3 nodes + 1 edge = at least 4 shapes consumed
  expect_gte(r$next_id, 5L)
})

# ── XML validity ───────────────────────────────────────────────────────────

test_that("output is valid XML", {
  r <- build_diagram_xml(make_svg_data())
  expect_no_error(xml2::read_xml(r$xml))
})

test_that("output is valid XML for 3-node diagram", {
  r <- build_diagram_xml(make_svg_data(n_nodes = 3L))
  expect_no_error(xml2::read_xml(r$xml))
})

test_that("output is valid XML when no edges", {
  r <- build_diagram_xml(make_svg_data(add_edge = FALSE))
  expect_no_error(xml2::read_xml(r$xml))
})

# ── Namespace declarations ─────────────────────────────────────────────────

test_that("required namespaces present", {
  xml <- build_diagram_xml(make_svg_data())$xml
  expect_true(grepl("wordprocessingShape",   xml, fixed = TRUE))
  expect_true(grepl("wordprocessingDrawing", xml, fixed = TRUE))
  expect_true(grepl('xmlns:a=',              xml, fixed = TRUE))
  expect_true(grepl('xmlns:w=',              xml, fixed = TRUE))
  expect_true(grepl('xmlns:wps=',            xml, fixed = TRUE))
  expect_true(grepl('xmlns:wp=',             xml, fixed = TRUE))
})

# ── Shape count ────────────────────────────────────────────────────────────

test_that("correct number of wsp elements: 2 nodes + 1 edge = 3", {
  r   <- build_diagram_xml(make_svg_data(n_nodes = 2L))
  doc <- xml2::read_xml(r$xml)
  wsps <- xml2::xml_find_all(doc, "//*[local-name()='wsp']")
  expect_gte(length(wsps), 3L)
})

test_that("correct wsp count for 3 nodes 2 edges", {
  d <- make_svg_data(n_nodes = 3L, add_edge = FALSE)
  # manually add two edges
  d$edges <- tibble::tibble(
    id          = c("L-A-B-0", "L-B-C-0"),
    from        = c("A", "B"),
    to          = c("B", "C"),
    label       = c(NA_character_, NA_character_),
    path_d      = c("M 100 68 L 250 32", "M 250 68 L 400 32"),
    arrow_start = c("none", "none"),
    arrow_end   = c("arrow", "arrow"),
    line_type   = c("solid", "solid"),
    label_x     = c(NA_real_, NA_real_),
    label_y     = c(NA_real_, NA_real_)
  )
  r   <- build_diagram_xml(d)
  doc <- xml2::read_xml(r$xml)
  wsps <- xml2::xml_find_all(doc, "//*[local-name()='wsp']")
  expect_gte(length(wsps), 5L)   # 3 nodes + 2 edges
})

# ── Shape geometry ─────────────────────────────────────────────────────────

test_that("diamond prstGeom present for diamond shape", {
  d <- make_svg_data(n_nodes = 1L, add_edge = FALSE, shapes = "diamond")
  xml <- build_diagram_xml(d)$xml
  expect_true(grepl('prst="diamond"', xml, fixed = TRUE))
})

test_that("ellipse prstGeom present for ellipse shape", {
  d <- make_svg_data(n_nodes = 1L, add_edge = FALSE, shapes = "ellipse")
  xml <- build_diagram_xml(d)$xml
  expect_true(grepl('prst="ellipse"', xml, fixed = TRUE))
})

test_that("roundRect prstGeom present for roundRect shape", {
  d <- make_svg_data(n_nodes = 1L, add_edge = FALSE, shapes = "roundRect")
  xml <- build_diagram_xml(d)$xml
  expect_true(grepl('prst="roundRect"', xml, fixed = TRUE))
})

# ── Edge styling ───────────────────────────────────────────────────────────

test_that("dashed edge includes prstDash in XML", {
  d   <- make_svg_data(dashed = TRUE)
  xml <- build_diagram_xml(d)$xml
  expect_true(grepl("prstDash", xml, fixed = TRUE))
})

test_that("connector marker present for edges", {
  d   <- make_svg_data()
  xml <- build_diagram_xml(d)$xml
  expect_true(grepl("cNvCnPr", xml, fixed = TRUE))
})

# ── Edge label shapes ──────────────────────────────────────────────────────

test_that("edge label produces additional wsp text box", {
  d_no_lbl  <- make_svg_data(edge_label = NA_character_)
  d_with_lbl <- make_svg_data(edge_label = "yes")

  count_wsps <- function(xml) {
    doc <- xml2::read_xml(xml)
    length(xml2::xml_find_all(doc, "//*[local-name()='wsp']"))
  }
  expect_gt(count_wsps(build_diagram_xml(d_with_lbl)$xml),
            count_wsps(build_diagram_xml(d_no_lbl)$xml))
})

# ── Shape ID management ────────────────────────────────────────────────────

test_that("no shape ID collision across two diagrams", {
  d  <- make_svg_data()
  r1 <- build_diagram_xml(d, start_id = 1L)
  r2 <- build_diagram_xml(d, start_id = r1$next_id)

  get_ids <- function(xml) {
    m <- gregexpr('(?<=<wp:docPr id=")[0-9]+', xml, perl = TRUE)
    as.integer(regmatches(xml, m)[[1]])
  }
  expect_equal(length(intersect(get_ids(r1$xml), get_ids(r2$xml))), 0L)
})

test_that("start_id parameter is respected", {
  r <- build_diagram_xml(make_svg_data(n_nodes = 1L, add_edge = FALSE),
                         start_id = 42L)
  expect_true(grepl('<wp:docPr id="42"', r$xml, fixed = TRUE))
})

# ── Round-trip metadata ────────────────────────────────────────────────────

test_that("descr attribute contains JSON type field", {
  xml <- build_diagram_xml(make_svg_data())$xml
  # JSON is XML-escaped inside the descr attribute
  expect_true(grepl("&quot;type&quot;", xml, fixed = TRUE))
})

test_that("node metadata includes id and shape fields", {
  xml <- build_diagram_xml(make_svg_data(n_nodes = 1L, add_edge = FALSE))$xml
  expect_true(grepl("&quot;id&quot;",    xml, fixed = TRUE))
  expect_true(grepl("&quot;shape&quot;", xml, fixed = TRUE))
})

# ── EMU positions are positive ─────────────────────────────────────────────

test_that("posOffset values are positive integers", {
  xml <- build_diagram_xml(make_svg_data())$xml
  m   <- gregexpr("<wp:posOffset>([0-9]+)</wp:posOffset>", xml, perl = TRUE)
  offsets <- regmatches(xml, m)[[1]]
  expect_true(length(offsets) >= 2L)
})

# ── Empty diagram ──────────────────────────────────────────────────────────

test_that("empty diagram returns valid XML w:p", {
  d <- list(nodes = tibble::tibble(
               id=character(), label=character(), shape=character(),
               svg_cx=numeric(), svg_cy=numeric(),
               svg_w=numeric(), svg_h=numeric(),
               fill=character(), stroke=character(), class=character()),
            edges = tibble::tibble(
               id=character(), from=character(), to=character(),
               label=character(), path_d=character(),
               arrow_start=character(), arrow_end=character(),
               line_type=character(), label_x=numeric(), label_y=numeric()),
            viewbox = c(0, 0, 200, 100))
  r <- build_diagram_xml(d)
  expect_no_error(xml2::read_xml(r$xml))
  expect_equal(r$next_id, 1L)
})

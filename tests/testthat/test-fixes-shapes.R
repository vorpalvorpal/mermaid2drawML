# F7 (stadium / rounded radius), F9 (invisible links), F10 (ellipse text overlay).

# ── F7 ─────────────────────────────────────────────────────────────────────

test_that("F7: parse gives roundRect with svg_rx = h/2 for a stadium, rx for rounded", {
  nodes <- fix_parse("base-segoe-stadium-bold-invisible")$nodes
  expect_true("svg_rx" %in% names(nodes))
  a <- nodes[nodes$id == "A", ]
  expect_equal(a$shape, "roundRect")
  expect_equal(a$svg_rx, a$svg_h / 2, tolerance = 1 / a$svg_h)   # within 1 px
  d <- nodes[nodes$id == "D", ]
  expect_equal(d$shape, "roundRect")
  expect_equal(d$svg_rx, 5)
  # plain rect and circle are not rounded
  expect_true(is.na(nodes$svg_rx[nodes$id == "B"]))
  expect_true(is.na(nodes$svg_rx[nodes$id == "C"]))
})

test_that("F7: the default-theme stadium also carries its radius", {
  nodes <- fix_parse("default-theme")$nodes
  a <- nodes[nodes$id == "A", ]
  expect_equal(a$svg_rx, a$svg_h / 2, tolerance = 1 / a$svg_h)
})

adj_of <- function(doc, name) {
  shp <- fix_shape(doc, name)
  geom <- xml2::xml_find_first(shp, ".//*[local-name()='prstGeom']")
  list(prst = xml2::xml_attr(geom, "prst"),
       gd = xml2::xml_attr(
         xml2::xml_find_first(geom, "./*[local-name()='avLst']/*[local-name()='gd']"),
         "fmla"),
       name = xml2::xml_attr(
         xml2::xml_find_first(geom, "./*[local-name()='avLst']/*[local-name()='gd']"),
         "name"))
}

test_that("F7: a stadium is a pill: roundRect with adj 50000", {
  doc <- fix_xml(build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))$xml)
  a <- adj_of(doc, "mermaid:node:A")
  expect_equal(a$prst, "roundRect")
  expect_equal(a$name, "adj")
  expect_equal(a$gd, "val 50000")
})

test_that("F7: a rounded rect gets adj = round(100000 * rx / min(w, h))", {
  doc <- fix_xml(build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))$xml)
  d <- adj_of(doc, "mermaid:node:D")           # rx 5, 93.8 x 54
  expect_equal(d$prst, "roundRect")
  expect_equal(d$gd, paste0("val ", round(100000 * 5 / 54)))
})

test_that("F7: adj is capped at 50000", {
  dat <- fix_svg_data(shapes = c("roundRect", "rect"))
  dat$nodes$svg_rx <- c(500, NA_real_)
  doc <- fix_xml(build_diagram_xml(dat)$xml)
  expect_equal(adj_of(doc, "mermaid:node:A")$gd, "val 50000")
})

# ── F9 ─────────────────────────────────────────────────────────────────────

test_that("F9: parse drops or marks the invisible link A ~~~ C", {
  edges <- fix_parse("base-segoe-stadium-bold-invisible")$edges
  ac <- edges[edges$from == "A" & edges$to == "C", ]
  expect_true(nrow(ac) == 0L || all(ac$line_type == "invisible"))
  visible <- edges[edges$line_type != "invisible", ]
  expect_setequal(paste(visible$from, visible$to), c("A B", "B C", "B D"))
})

test_that("F9: the invisible link yields no connector shape in the output", {
  doc <- fix_xml(build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))$xml)
  conns <- fix_shapes_named(doc, "mermaid:edge:")
  expect_length(conns, 3L)
  expect_false(any(xml2::xml_attr(
    xml2::xml_find_all(doc, "//*[local-name()='cNvPr']"), "name") == "mermaid:edge:A-C"))
})

test_that("F9: a synthetic invisible edge gets no shape; the visible one does", {
  edges <- rbind(fix_edges("A", "B"), fix_edges("A", "C", line_type = "invisible"))
  dat <- fix_svg_data(shapes = c("rect", "rect", "rect"), edges = edges)
  doc <- fix_xml(build_diagram_xml(dat)$xml)
  expect_length(fix_shapes_named(doc, "mermaid:edge:"), 1L)
})

# ── F10 ────────────────────────────────────────────────────────────────────

test_that("F10: a circle is a group: ellipse visual + transparent rect text overlay", {
  doc <- fix_xml(build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))$xml)
  grp <- xml2::xml_find_first(
    doc, "//*[local-name()='grpSp'][*[local-name()='cNvPr'][@name='mermaid:node:C']]")
  expect_false(inherits(grp, "xml_missing"))
  wsps <- xml2::xml_find_all(grp, "./*[local-name()='wsp']")
  expect_length(wsps, 2L)
  prst <- vapply(wsps, function(w) xml2::xml_attr(
    xml2::xml_find_first(w, ".//*[local-name()='prstGeom']"), "prst"), character(1))
  expect_setequal(prst, c("ellipse", "rect"))
  txt <- wsps[prst == "rect"][[1]]
  expect_false(inherits(
    xml2::xml_find_first(txt, "./*[local-name()='spPr']/*[local-name()='noFill']"),
    "xml_missing"))
  expect_false(inherits(
    xml2::xml_find_first(txt, ".//*[local-name()='txbxContent']"), "xml_missing"))
  vis <- wsps[prst == "ellipse"][[1]]
  expect_true(inherits(
    xml2::xml_find_first(vis, ".//*[local-name()='txbxContent']"), "xml_missing"))
})

test_that("F10: the circle's label is in the overlay text box", {
  doc <- fix_xml(build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))$xml)
  grp <- xml2::xml_find_first(
    doc, "//*[local-name()='grpSp'][*[local-name()='cNvPr'][@name='mermaid:node:C']]")
  expect_equal(trimws(xml2::xml_text(grp)), "Circle")
})

test_that("F10: a synthetic ellipse node also gets the group", {
  dat <- fix_svg_data(shapes = c("ellipse", "rect"))
  doc <- fix_xml(build_diagram_xml(dat)$xml)
  grp <- xml2::xml_find_first(
    doc, "//*[local-name()='grpSp'][*[local-name()='cNvPr'][@name='mermaid:node:A']]")
  expect_false(inherits(grp, "xml_missing"))
})

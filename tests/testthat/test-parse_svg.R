# tests/testthat/test-parse_svg.R
# Tests for SVG parsing helpers. These run on static SVG strings so they do
# NOT require Node.js / mermaid-cli to be installed.

# ── Minimal SVG fixture ────────────────────────────────────────────────────

minimal_svg <- function(inner = "", width = 200, height = 100,
                         viewBox = "0 0 200 100") {
  paste0(
    '<svg xmlns="http://www.w3.org/2000/svg" ',
    'width="', width, '" height="', height, '" ',
    'viewBox="', viewBox, '">',
    inner,
    '</svg>'
  )
}

node_g <- function(id, transform = 'translate(100,50)',
                   shape_el = '<rect x="-40" y="-17.5" width="80" height="35"/>',
                   label = "Hello", class = "node default") {
  paste0(
    '<g id="', id, '" class="', class, '" transform="', transform, '">',
    shape_el,
    '<g class="label"><foreignObject width="80" height="35">',
    label, '</foreignObject></g>',
    '</g>'
  )
}

# ── viewBox parsing ────────────────────────────────────────────────────────

test_that("parse_viewbox extracts four numbers", {
  svg  <- minimal_svg()
  doc  <- xml2::read_xml(svg)
  xml2::xml_ns_strip(doc)
  vb   <- parse_viewbox(doc)
  expect_equal(vb, c(0, 0, 200, 100))
})

test_that("parse_viewbox falls back to width/height when viewBox absent", {
  svg  <- '<svg xmlns="http://www.w3.org/2000/svg" width="300" height="150"></svg>'
  doc  <- xml2::read_xml(svg)
  xml2::xml_ns_strip(doc)
  vb   <- parse_viewbox(doc)
  expect_equal(vb[3], 300)
  expect_equal(vb[4], 150)
})

# ── Translate parsing ──────────────────────────────────────────────────────

test_that("parse_translate extracts x and y", {
  expect_equal(parse_translate('translate(100, 50)'), c(100, 50))
  expect_equal(parse_translate('translate(0,0)'),     c(0, 0))
})

test_that("parse_translate returns zeros for missing transform", {
  expect_equal(parse_translate(""), c(0, 0))
})

# ── Node geometry ──────────────────────────────────────────────────────────

test_that("rect node parsed to shape=rect", {
  svg  <- minimal_svg(node_g("flowchart-A-1"))
  data <- parse_mermaid_svg(svg)
  expect_equal(nrow(data$nodes), 1L)
  expect_equal(data$nodes$shape, "rect")
})

test_that("rounded rect detected via rx attribute", {
  el   <- '<rect x="-40" y="-17.5" width="80" height="35" rx="5" ry="5"/>'
  svg  <- minimal_svg(node_g("flowchart-B-1", shape_el = el))
  data <- parse_mermaid_svg(svg)
  expect_equal(data$nodes$shape, "roundRect")
})

test_that("polygon with 4 points → diamond", {
  el   <- '<polygon points="50,0 0,-25 -50,0 0,25"/>'
  svg  <- minimal_svg(node_g("flowchart-C-1", shape_el = el))
  data <- parse_mermaid_svg(svg)
  expect_equal(data$nodes$shape, "diamond")
})

test_that("ellipse node parsed correctly", {
  el   <- '<ellipse rx="30" ry="20"/>'
  svg  <- minimal_svg(node_g("flowchart-D-1", shape_el = el))
  data <- parse_mermaid_svg(svg)
  expect_equal(data$nodes$shape, "ellipse")
  expect_equal(data$nodes$svg_w, 60)
  expect_equal(data$nodes$svg_h, 40)
})

test_that("node cx/cy set from translate", {
  svg  <- minimal_svg(node_g("flowchart-A-1", transform = 'translate(80,60)'))
  data <- parse_mermaid_svg(svg)
  expect_equal(data$nodes$svg_cx, 80)
  expect_equal(data$nodes$svg_cy, 60)
})

test_that("node label extracted from foreignObject", {
  svg  <- minimal_svg(node_g("flowchart-A-1", label = "My Label"))
  data <- parse_mermaid_svg(svg)
  expect_true(grepl("My Label", data$nodes$label))
})

test_that("mermaid id extracted from flowchart-X-N pattern", {
  svg  <- minimal_svg(node_g("flowchart-StartNode-3"))
  data <- parse_mermaid_svg(svg)
  expect_equal(data$nodes$id, "StartNode")
})

# ── Empty SVG ─────────────────────────────────────────────────────────────

test_that("empty SVG returns empty node/edge tibbles", {
  data <- parse_mermaid_svg(minimal_svg())
  expect_equal(nrow(data$nodes), 0L)
  expect_equal(nrow(data$edges), 0L)
})

# ── SVG path parser ────────────────────────────────────────────────────────

test_that("parse_svg_path handles M and L", {
  cmds <- parse_svg_path("M 10 20 L 30 40")
  expect_length(cmds, 2L)
  expect_equal(cmds[[1]]$cmd, "M")
  expect_equal(cmds[[2]]$cmd, "L")
  expect_equal(cmds[[2]]$x, 30)
  expect_equal(cmds[[2]]$y, 40)
})

test_that("parse_svg_path handles relative m and l", {
  cmds <- parse_svg_path("M 10 10 l 5 5")
  expect_equal(cmds[[2]]$cmd, "L")
  expect_equal(cmds[[2]]$x, 15)
  expect_equal(cmds[[2]]$y, 15)
})

test_that("parse_svg_path handles cubic bezier C", {
  cmds <- parse_svg_path("M 0 0 C 10 5 20 5 30 0")
  bz   <- cmds[[2]]
  expect_equal(bz$cmd, "C")
  expect_equal(bz$x1, 10); expect_equal(bz$y1, 5)
  expect_equal(bz$x,  30); expect_equal(bz$y,  0)
})

test_that("parse_svg_path handles Z close", {
  cmds <- parse_svg_path("M 0 0 L 10 10 Z")
  expect_equal(cmds[[3]]$cmd, "Z")
})

test_that("svg_path_to_custgeom returns valid custGeom XML", {
  result <- svg_path_to_custgeom("M 0 0 L 100 0 L 100 50 L 0 50 Z", scale = 9525)
  expect_false(is.null(result))
  expect_true(grepl("custGeom", result$xml))
  expect_true(grepl("moveTo",   result$xml))
  expect_true(grepl("lnTo",     result$xml))
  expect_gt(result$w_emu, 0L)
  expect_gt(result$h_emu, 0L)
})

test_that("svg_path_to_custgeom custGeom XML is valid XML", {
  result <- svg_path_to_custgeom("M 10 10 C 20 5 30 5 40 10", scale = 9525)
  expect_false(is.null(result))
  # Wrap in a root element to make it parseable standalone
  wrapped <- paste0('<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
                    result$xml, '</root>')
  expect_no_error(xml2::read_xml(wrapped))
})

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

# Mermaid draws trapezoids and parallelograms as 4-point polygons too. The
# point lists below are copied from mermaid 11.17 output (y grows downwards,
# so y = 0 is the bottom edge).

test_that("polygon with narrow top edge → trapezoid", {
  # T[/"text"\]
  el   <- '<polygon points="-31.5,0 121.83,0 90.33,-63 0,-63"/>'
  svg  <- minimal_svg(node_g("flowchart-T-1", shape_el = el))
  data <- parse_mermaid_svg(svg)
  expect_equal(data$nodes$shape, "trapezoid")
  expect_equal(data$nodes$svg_w, 153.33)
  expect_equal(data$nodes$svg_h, 63)
  expect_equal(data$nodes$svg_inset, 31.5)   # run of the slanted sides
})

test_that("polygon with wide top edge → manualOp (alt trapezoid)", {
  # T[\"text"/]
  el   <- '<polygon points="0,0 129.47,0 168.47,-78 -39,-78"/>'
  svg  <- minimal_svg(node_g("flowchart-T-1", shape_el = el))
  data <- parse_mermaid_svg(svg)
  expect_equal(data$nodes$shape, "manualOp")
})

test_that("polygon with top edge shifted right → leanR", {
  # P[/text/]
  el   <- '<polygon points="-19.5,0 87.7,0 107.2,-39 0,-39"/>'
  svg  <- minimal_svg(node_g("flowchart-P-1", shape_el = el))
  data <- parse_mermaid_svg(svg)
  expect_equal(data$nodes$shape, "leanR")
})

test_that("polygon with top edge shifted left → leanL", {
  # P[\text\]
  el   <- '<polygon points="0,0 99,0 79.5,-39 -19.5,-39"/>'
  svg  <- minimal_svg(node_g("flowchart-P-1", shape_el = el))
  data <- parse_mermaid_svg(svg)
  expect_equal(data$nodes$shape, "leanL")
  expect_equal(data$nodes$svg_inset, 19.5)
})

test_that("classify_quad does not depend on vertex order", {
  trap <- data.frame(x = c(0, 90, 120, -30), y = c(-60, -60, 0, 0))
  expect_equal(classify_quad(trap)$shape, "trapezoid")
  expect_equal(classify_quad(trap[c(3, 1, 4, 2), ])$shape, "trapezoid")
  expect_equal(classify_quad(trap)$inset, 30)
  square <- data.frame(x = c(0, 80, 80, 0), y = c(0, 0, -40, -40))
  expect_equal(classify_quad(square), list(shape = "rect", inset = NA_real_))
})

test_that("nodes without slanted sides carry no inset", {
  svg  <- minimal_svg(node_g("flowchart-A-1"))
  expect_true(is.na(parse_mermaid_svg(svg)$nodes$svg_inset))
  el   <- '<polygon points="50,0 0,-25 -50,0 0,25"/>'
  svg  <- minimal_svg(node_g("flowchart-C-1", shape_el = el))
  expect_true(is.na(parse_mermaid_svg(svg)$nodes$svg_inset))
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

# ── Element ids and edge labels ────────────────────────────────────────────

# A two-node diagram with one labelled and one unlabelled edge, laid out the
# way mermaid 11 writes it. `prefix` is what mermaid 11.17 puts in front of
# every element id (the id of the root <svg> plus "-"); "" gives the bare ids
# of earlier versions. `label_body` is the content of the labelled edge's
# <g class="label">: SVG text by default, a <foreignObject> for HTML labels.
edge_svg <- function(prefix = "my-svg-",
                     label_body = paste0(
                       '<g><rect class="background" x="-30" y="-1" width="60" height="28"/>',
                       '<text y="-10.1"><tspan class="row">',
                       '<tspan>Self</tspan><tspan> haul</tspan></tspan></text></g>')) {
  edge_path <- function(id, d) paste0(
    '<path d="', d, '" id="', prefix, id, '" class="flowchart-link" ',
    'data-id="', id, '" marker-end="url(#my-svg_flowchart-v2-pointEnd)"/>')
  paste0(
    '<svg id="my-svg" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 200">',
    '<style>#my-svg .edgeLabel{background-color:#E8E8E8;text-align:center;}</style>',
    '<g class="root">',
    '<g class="edgePaths">',
    edge_path("L_A_B_0", "M140,50L260,50"),
    edge_path("L_B_A_0", "M260,60L140,60"),
    '</g>',
    '<g class="edgeLabels">',
    '<g class="edgeLabel" transform="translate(200, 50)">',
    '<g class="label" data-id="L_A_B_0" transform="translate(0, -13)">',
    label_body, '</g></g>',
    '<g class="edgeLabel"><g class="label" data-id="L_B_A_0" ',
    'transform="translate(0, 0)"><text><tspan class="row"/></text></g></g>',
    '</g>',
    '<g class="nodes">',
    node_g(paste0(prefix, "flowchart-A-0"), transform = "translate(100,50)"),
    node_g(paste0(prefix, "flowchart-B-1"), transform = "translate(300,50)"),
    '</g></g></svg>'
  )
}

test_that("svg-id prefix is stripped from node and edge ids", {
  data <- parse_mermaid_svg(edge_svg())
  expect_equal(data$nodes$id, c("A", "B"))
  expect_equal(data$edges$id, c("L_A_B_0", "L_B_A_0"))
  expect_equal(data$edges$from, c("A", "B"))
  expect_equal(data$edges$to,   c("B", "A"))
})

test_that("ids without the svg-id prefix parse the same way", {
  expect_equal(parse_mermaid_svg(edge_svg(prefix = "")),
               parse_mermaid_svg(edge_svg()))
})

test_that("strip_svg_id_prefix leaves other ids alone", {
  doc <- xml2::read_xml('<svg id="s"><g id="s-a"/><g id="t-a"/><g id="sa"/></svg>')
  strip_svg_id_prefix(doc)
  expect_equal(xml2::xml_attr(xml2::xml_children(doc), "id"), c("a", "t-a", "sa"))

  no_id <- xml2::read_xml('<svg><g id="s-a"/></svg>')
  expect_no_error(strip_svg_id_prefix(no_id))
  expect_equal(xml2::xml_attr(xml2::xml_children(no_id), "id"), "s-a")
})

test_that("@{ shape } overrides match nodes with prefixed ids", {
  data <- parse_mermaid_svg(edge_svg(), source = "flowchart LR\n  A@{ shape: cyl }\n  A --> B")
  expect_false(data$nodes$shape[data$nodes$id == "A"] == "rect")
  expect_equal(data$nodes$shape[data$nodes$id == "B"], "rect")
})

test_that("edge label is found through data-id, with position and size", {
  edges <- parse_mermaid_svg(edge_svg())$edges
  expect_equal(edges$label,   c("Self haul", NA))
  expect_equal(edges$label_x, c(200, NA))
  expect_equal(edges$label_y, c(50, NA))
  expect_equal(edges$label_w, c(60, NA))
  expect_equal(edges$label_h, c(28, NA))
})

test_that("HTML edge label keeps line breaks and takes its size from foreignObject", {
  html <- paste0('<foreignObject width="74" height="48"><div class="labelBkg">',
                 '<span class="edgeLabel"><p>two words<br/>two lines</p></span>',
                 '</div></foreignObject>')
  edges <- parse_mermaid_svg(edge_svg(label_body = html))$edges
  expect_equal(edges$label[1],   "two words\ntwo lines")
  expect_equal(edges$label_w[1], 74)
  expect_equal(edges$label_h[1], 48)
})

test_that("edge label background colour is read from the stylesheet", {
  expect_equal(parse_mermaid_svg(edge_svg())$style$edge_label_bg, "E8E8E8")
  expect_equal(parse_mermaid_svg(minimal_svg())$style$edge_label_bg, "FFFFFF")
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

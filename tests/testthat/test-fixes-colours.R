# F11: theme colours (hsl parsing, node_fill / node_stroke, build fallbacks).

near_hex <- function(actual, expected, tol = 1) {
  expect_false(is.na(actual))
  expect_lte(max(abs(fix_rgb(actual) - fix_rgb(expected))), tol)
}

test_that("F11: parse_css_colour understands hsl() and hsla()", {
  expect_equal(parse_css_colour("hsl(0,100%,50%)"), "FF0000")
  expect_equal(parse_css_colour("hsl(240, 100%, 96.27%)"), "ECECFF")
  near_hex(parse_css_colour("hsl(80, 100%, 96.2745098039%)"), "F9FFEC")
  expect_equal(parse_css_colour("hsl(180, 0%, 100%)"), "FFFFFF")
  near_hex(parse_css_colour("hsla(120, 100%, 25%, 0.5)"), "008000")
  near_hex(parse_css_colour("HSL(0, 0%, 50%)"), "808080")
})

test_that("F11: existing colour forms are unchanged", {
  expect_equal(parse_css_colour("#abc"), "AABBCC")
  expect_equal(parse_css_colour("rgb(255, 0, 16)"), "FF0010")
  expect_true(is.na(parse_css_colour("transparent")))
})

test_that("F11: extract_svg_style returns node_fill / node_stroke from the theme rule", {
  s <- fix_style("default-theme")
  expect_equal(s$node_fill, "ECECFF")
  expect_equal(s$node_stroke, "9370DB")
  b <- fix_style("base-segoe-stadium-bold-invisible")
  expect_equal(b$node_fill, "FFFFFF")
  expect_equal(b$node_stroke, "005490")
})

test_that("F11: node_fill / node_stroke are NA when the svg has no such rule", {
  s <- extract_svg_style(fix_doc(fix_svg_with_style("#my-svg{font-size:16px;}")))
  expect_true("node_fill" %in% names(s))
  expect_true(is.na(s$node_fill))
  expect_true(is.na(s$node_stroke))
})

test_that("F11: hsl() cluster colours are parsed", {
  s <- fix_style("base-segoe-subgraph")
  expect_equal(s$cluster_fill, "FFFFFF")
  near_hex(s$cluster_stroke, "E6E6E6", tol = 1)   # hsl(180, 0%, 90%)
})

node_fill_stroke <- function(doc, name) {
  # The wsp that draws the node (a grouped node's group shares its name).
  shp <- xml2::xml_find_first(doc, sprintf(
    "//*[local-name()='wsp'][*[local-name()='cNvPr'][@name='%s']]", name))
  sp <- xml2::xml_find_first(shp, "./*[local-name()='spPr']")
  fill <- xml2::xml_attr(xml2::xml_find_first(
    sp, "./*[local-name()='solidFill']/*[local-name()='srgbClr']"), "val")
  nofill <- !inherits(xml2::xml_find_first(sp, "./*[local-name()='noFill']"), "xml_missing")
  stroke <- xml2::xml_attr(xml2::xml_find_first(
    sp, "./*[local-name()='ln']/*[local-name()='solidFill']/*[local-name()='srgbClr']"), "val")
  list(fill = fill, nofill = nofill, stroke = stroke)
}

test_that("F11: a node with no fill/stroke of its own takes the theme colours", {
  dat <- fix_svg_data(style = fix_style_list(node_fill = "ECECFF", node_stroke = "9370DB"))
  dat$nodes$fill <- NA_character_
  dat$nodes$stroke <- NA_character_
  doc <- fix_xml(build_diagram_xml(dat)$xml)
  a <- node_fill_stroke(doc, "mermaid:node:A")
  expect_equal(a$fill, "ECECFF")
  expect_equal(a$stroke, "9370DB")
})

test_that("F11: an explicit node colour is not overridden by the theme", {
  dat <- fix_svg_data(style = fix_style_list(node_fill = "ECECFF", node_stroke = "9370DB"),
                      fill = "00FF00", stroke = "0000FF")
  doc <- fix_xml(build_diagram_xml(dat)$xml)
  a <- node_fill_stroke(doc, "mermaid:node:A")
  expect_equal(a$fill, "00FF00")
  expect_equal(a$stroke, "0000FF")
})

test_that("F11: diamond (visual wsp of a group) also takes the theme colours", {
  dat <- fix_svg_data(shapes = c("diamond", "rect"),
                      style = fix_style_list(node_fill = "ECECFF", node_stroke = "9370DB"))
  dat$nodes$fill <- NA_character_
  dat$nodes$stroke <- NA_character_
  doc <- fix_xml(build_diagram_xml(dat)$xml)
  grp <- xml2::xml_find_first(
    doc, "//*[local-name()='grpSp'][*[local-name()='cNvPr'][@name='mermaid:node:A']]")
  vis <- xml2::xml_find_first(grp, "./*[local-name()='wsp'][1]/*[local-name()='spPr']")
  expect_equal(xml2::xml_attr(xml2::xml_find_first(
    vis, "./*[local-name()='solidFill']/*[local-name()='srgbClr']"), "val"), "ECECFF")
  expect_equal(xml2::xml_attr(xml2::xml_find_first(
    vis, "./*[local-name()='ln']/*[local-name()='solidFill']/*[local-name()='srgbClr']"), "val"),
    "9370DB")
})

test_that("F11: base theme fixture: unstyled nodes get blue borders, transparent stays transparent", {
  doc <- fix_xml(build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))$xml)
  b <- node_fill_stroke(doc, "mermaid:node:B")
  expect_equal(b$stroke, "005490")
  expect_equal(b$fill, "FFFFFF")
  d <- node_fill_stroke(doc, "mermaid:node:D")     # classDef fill:transparent,stroke:#000000
  expect_true(d$nofill)
  expect_true(is.na(d$fill))
  expect_equal(d$stroke, "000000")
})

test_that("F11: default theme fixture: unstyled nodes are ECECFF / 9370DB", {
  doc <- fix_xml(build_diagram_xml(fix_parse("default-theme"))$xml)
  c_ <- node_fill_stroke(doc, "mermaid:node:C")      # rect, no classDef
  expect_equal(c_$fill, "ECECFF")
  expect_equal(c_$stroke, "9370DB")
})

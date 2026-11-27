# F1: the Word font is the theme's declared font-family, not --mermaid-font-family.

test_that("F1: extract_svg_style takes font_family from the svg root rule (Segoe UI)", {
  expect_equal(fix_style("base-segoe-stadium-bold-invisible")$font_family, "Segoe UI")
  expect_equal(fix_style("base-segoe-subgraph")$font_family, "Segoe UI")
})

test_that("F1: default theme (quoted lower-case stack) gives canonical Trebuchet MS", {
  expect_equal(fix_style("default-theme")$font_family, "Trebuchet MS")
})

test_that("F1: first family wins, quotes are removed, known names get canonical case", {
  doc <- fix_doc(fix_svg_with_style(
    '#my-svg{font-family:"georgia",serif;font-size:16px;fill:#333;}'))
  expect_equal(extract_svg_style(doc)$font_family, "Georgia")

  doc <- fix_doc(fix_svg_with_style(
    "#my-svg{font-family:'Myriad Pro',Arial,sans-serif;font-size:16px;}"))
  expect_equal(extract_svg_style(doc)$font_family, "Myriad Pro")
})

test_that("F1: falls back to --mermaid-font-family, then to Trebuchet MS", {
  doc <- fix_doc(fix_svg_with_style(
    '#my-svg :root{--mermaid-font-family:"verdana",arial,sans-serif;}#my-svg{font-size:16px;}'))
  expect_equal(extract_svg_style(doc)$font_family, "Verdana")

  doc <- fix_doc(fix_svg_with_style("#my-svg{font-size:16px;}"))
  expect_equal(extract_svg_style(doc)$font_family, "Trebuchet MS")
})

test_that("F1: build_diagram_xml writes the theme font into every rFonts attribute", {
  res <- build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))
  doc <- fix_xml(res$xml)
  fonts <- xml2::xml_find_all(doc, "//*[local-name()='rFonts']")
  expect_gt(length(fonts), 0L)
  for (attr in c("ascii", "hAnsi", "cs", "eastAsia")) {
    vals <- vapply(fonts, function(f) {
      a <- xml2::xml_attr(f, paste0("w:", attr), fix_ns(doc))
      if (is.na(a)) NA_character_ else a
    }, character(1))
    expect_true(all(vals == "Segoe UI"), info = attr)
  }
  expect_false(grepl("Trebuchet", res$xml, fixed = TRUE))
})

test_that("F1: word_font_family still overrides the theme font (incl. eastAsia)", {
  res <- build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"),
                           word_font_family = "Calibri")
  doc <- fix_xml(res$xml)
  fonts <- xml2::xml_find_all(doc, "//*[local-name()='rFonts']")
  east <- xml2::xml_attr(fonts, "w:eastAsia", fix_ns(doc))
  expect_true(all(east == "Calibri"))
  expect_false(grepl("Segoe UI", res$xml, fixed = TRUE))
})

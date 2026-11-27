# Shared helpers for the test-fixes-*.R files (faults F1-F12).
# All offline: real mermaid-cli 11.12 SVGs live in tests/testthat/fixtures/.

fix_read <- function(name) {
  paste(readLines(test_path("fixtures", name), warn = FALSE), collapse = "\n")
}

# Parse a fixture (<name>.svg + <name>.mmd) the way mermaid_to_word() does.
fix_parse <- function(name) {
  parse_mermaid_svg(fix_read(paste0(name, ".svg")), ast = NULL,
                    source = fix_read(paste0(name, ".mmd")))
}

# The svg document as extract_svg_style() receives it (namespace stripped).
fix_doc <- function(svg) {
  doc <- xml2::read_xml(svg)
  xml2::xml_ns_strip(doc)
  doc
}

fix_style <- function(name) extract_svg_style(fix_doc(fix_read(paste0(name, ".svg"))))

# Minimal synthetic svg with a given <style> body.
fix_svg_with_style <- function(css) {
  paste0('<svg xmlns="http://www.w3.org/2000/svg" id="my-svg" ',
         'viewBox="0 0 100 100"><style>', css, '</style><g/></svg>')
}

# Parse the XML returned by build_diagram_xml() / label_txbx_xml() fragments.
fix_xml <- function(xml) xml2::read_xml(xml)
fix_frag <- function(frag) {
  xml2::read_xml(paste0("<w:root ", .dml_namespaces, ">", frag, "</w:root>"))
}
fix_ns <- function(doc) xml2::xml_ns(doc)

# Find the cNvPr (any prefix) with this name, and return the element that owns it.
fix_shape <- function(doc, name) {
  xml2::xml_find_first(
    doc, sprintf("//*[local-name()='cNvPr'][@name='%s']/..", name))
}

fix_shapes_named <- function(doc, prefix) {
  xml2::xml_find_all(
    doc, sprintf("//*[local-name()='cNvPr'][starts-with(@name,'%s')]/..", prefix))
}

# A small synthetic svg_data (no parsing). `style` may be extended per test.
fix_svg_data <- function(shapes = c("rect", "rect"), vb = c(0, 0, 400, 150),
                         edges = NULL, style = NULL, fill = "FFFFFF",
                         stroke = "4472C4") {
  n <- length(shapes)
  nodes <- tibble::tibble(
    id = LETTERS[seq_len(n)], label = paste("Node", LETTERS[seq_len(n)]),
    shape = shapes,
    svg_cx = seq(60, by = 120, length.out = n), svg_cy = rep(50, n),
    svg_w = rep(80, n), svg_h = rep(35, n),
    txt_w = rep(60, n), txt_h = rep(24, n),
    fill = rep(fill, n), stroke = rep(stroke, n), class = rep("default", n)
  )
  if (is.null(edges)) edges <- fix_edges("A", "B")
  out <- list(nodes = nodes, edges = edges, viewbox = vb)
  if (!is.null(style)) out$style <- style
  out
}

fix_edges <- function(from, to, line_type = "solid", label = NA_character_) {
  tibble::tibble(
    id = paste0("L-", from, "-", to, "-0"), from = from, to = to,
    label = label, path_d = "M 100 68 L 250 32",
    arrow_start = "none", arrow_end = "arrow", line_type = line_type,
    label_x = NA_real_, label_y = NA_real_
  )
}

fix_style_list <- function(...) {
  utils::modifyList(
    list(font_size_px = 14, font_family = "Arial", edge_stroke = "333333",
         edge_sw_px = 2, edge_label_bg = "FFFFFF"),
    list(...))
}

# "F9FFEC" -> c(249, 255, 236)
fix_rgb <- function(hex) strtoi(substring(hex, c(1, 3, 5), c(2, 4, 6)), 16L)

# F2 (inline drawing), F3 (unique shape ids), F6 (max_scale).

# ── F2 ─────────────────────────────────────────────────────────────────────

test_that("F2: the drawing is wp:inline, with no anchor/wrap/position elements", {
  res <- build_diagram_xml(fix_svg_data(), start_id = 7L)
  doc <- fix_xml(res$xml)
  n <- function(xp) length(xml2::xml_find_all(doc, xp))
  expect_equal(n("//*[local-name()='inline']"), 1L)
  expect_equal(n("//*[local-name()='anchor']"), 0L)
  expect_equal(n("//*[local-name()='wrapNone']"), 0L)
  expect_equal(n("//*[local-name()='positionV']"), 0L)
  expect_equal(n("//*[local-name()='positionH']"), 0L)
  expect_equal(n("//*[local-name()='simplePos']"), 0L)
})

test_that("F2: wp:extent equals the group size and docPr carries start_id", {
  res <- build_diagram_xml(fix_svg_data(), start_id = 7L)
  doc <- fix_xml(res$xml)
  ext <- xml2::xml_find_first(doc, "//*[local-name()='inline']/*[local-name()='extent']")
  grp <- xml2::xml_find_first(doc, "//*[local-name()='wgp']/*[local-name()='grpSpPr']/*[local-name()='xfrm']/*[local-name()='ext']")
  expect_false(inherits(ext, "xml_missing"))
  expect_equal(xml2::xml_attr(ext, "cx"), xml2::xml_attr(grp, "cx"))
  expect_equal(xml2::xml_attr(ext, "cy"), xml2::xml_attr(grp, "cy"))
  dp <- xml2::xml_find_first(doc, "//*[local-name()='inline']/*[local-name()='docPr']")
  expect_equal(xml2::xml_attr(dp, "id"), "7")
})

test_that("F2: align becomes the paragraph w:jc", {
  for (al in c("center", "left", "right")) {
    doc <- fix_xml(build_diagram_xml(fix_svg_data(), align = al)$xml)
    jc <- xml2::xml_attr(
      xml2::xml_find_first(doc, "/*[local-name()='p']/*[local-name()='pPr']/*[local-name()='jc']"),
      "w:val", fix_ns(doc))
    expect_equal(jc, al)
  }
})

test_that("F2: the holder paragraph has explicit single spacing", {
  doc <- fix_xml(build_diagram_xml(fix_svg_data())$xml)
  sp <- xml2::xml_find_first(
    doc, "/*[local-name()='p']/*[local-name()='pPr']/*[local-name()='spacing']")
  expect_false(inherits(sp, "xml_missing"))
  ns <- fix_ns(doc)
  expect_equal(xml2::xml_attr(sp, "w:before", ns), "0")
  expect_equal(xml2::xml_attr(sp, "w:after", ns), "0")
  expect_equal(xml2::xml_attr(sp, "w:line", ns), "240")
  expect_equal(xml2::xml_attr(sp, "w:lineRule", ns), "auto")
})

test_that("F2: the mc:AlternateContent / Requires=wpg wrapper stays", {
  doc <- fix_xml(build_diagram_xml(fix_svg_data())$xml)
  ch <- xml2::xml_find_first(
    doc, "//*[local-name()='AlternateContent']/*[local-name()='Choice'][@Requires='wpg']")
  expect_false(inherits(ch, "xml_missing"))
  expect_false(inherits(
    xml2::xml_find_first(ch, ".//*[local-name()='drawing']/*[local-name()='inline']"),
    "xml_missing"))
})

# ── F3 ─────────────────────────────────────────────────────────────────────

doc_with_ids <- function(xml_pieces) {
  doc <- officer::read_docx()
  for (p in xml_pieces) doc <- officer::body_add_xml(doc, p)
  doc
}

test_that("F3: .next_free_shape_id is 1 for an empty document, and an integer", {
  id <- .next_free_shape_id(officer::read_docx())
  expect_type(id, "integer")
  expect_equal(id, 1L)
})

test_that("F3: .next_free_shape_id is one more than the largest docPr / any cNvPr id", {
  piece <- function(...) {
    paste0("<w:p ", .dml_namespaces, ">", paste0(..., collapse = ""), "</w:p>")
  }
  doc <- doc_with_ids(list(
    piece('<wp:docPr id="5" name="d"/>'),
    piece('<wpg:cNvPr id="9" name="g"/>', '<wps:cNvPr id="3" name="s"/>')
  ))
  expect_equal(.next_free_shape_id(doc), 10L)

  doc2 <- doc_with_ids(list(
    piece('<wp:docPr id="42" name="d"/>', '<wps:cNvPr id="9" name="s"/>')
  ))
  expect_equal(.next_free_shape_id(doc2), 43L)
})

test_that("F3: two body_add_mermaid() calls with default start_id share no ids (offline)", {
  svg <- fix_read("default-theme.svg")
  testthat::local_mocked_bindings(
    render_mermaid = function(mermaid_code, timeout = 120L) list(ast = NULL, svg = svg)
  )
  doc <- officer::read_docx()
  doc <- body_add_mermaid(doc, "flowchart TB\n A --> B")
  doc <- body_add_mermaid(doc, "flowchart TB\n A --> B")

  body <- officer::docx_body_xml(doc)
  ids <- xml2::xml_attr(
    xml2::xml_find_all(body, "//*[local-name()='docPr' or local-name()='cNvPr']"), "id")
  expect_gt(length(ids), 10L)
  expect_false(anyDuplicated(ids) > 0L)
})

test_that("F3: an explicit start_id is still honoured (offline)", {
  svg <- fix_read("default-theme.svg")
  testthat::local_mocked_bindings(
    render_mermaid = function(mermaid_code, timeout = 120L) list(ast = NULL, svg = svg)
  )
  doc <- body_add_mermaid(officer::read_docx(), "flowchart TB\n A --> B", start_id = 500L)
  body <- officer::docx_body_xml(doc)
  dp <- xml2::xml_attr(xml2::xml_find_all(body, "//*[local-name()='docPr']"), "id")
  expect_true("500" %in% dp)
})

# ── F6 ─────────────────────────────────────────────────────────────────────

extent_cx <- function(res) {
  doc <- fix_xml(res$xml)
  as.numeric(xml2::xml_attr(
    xml2::xml_find_first(doc, "//*[local-name()='wgp']/*[local-name()='grpSpPr']/*[local-name()='xfrm']/*[local-name()='ext']"), "cx"))
}

test_that("F6: by default a small diagram is drawn at natural size (9525 EMU per px)", {
  # 400 px wide would be blown up to ~5.7M EMU to fill the page without max_scale.
  expect_equal(extent_cx(build_diagram_xml(fix_svg_data())), 400 * 9525, tolerance = 1e-3)
})

test_that("F6: max_scale = Inf restores fill-the-page; smaller max_scale shrinks", {
  big <- extent_cx(build_diagram_xml(fix_svg_data(), max_scale = Inf))
  expect_gt(big, 400 * 9525 * 1.2)
  expect_lte(big, inches_to_emu(6.27) + 1000)
  half <- extent_cx(build_diagram_xml(fix_svg_data(), max_scale = 0.5))
  expect_equal(half, 400 * 9525 * 0.5, tolerance = 1e-3)
})

test_that("F6: a diagram wider than the page is still shrunk to fit", {
  res <- build_diagram_xml(fix_svg_data(vb = c(0, 0, 2000, 150)))
  expect_lte(extent_cx(res), inches_to_emu(6.27) + 1000)
})

test_that("F6: default text is 12 pt (24 half-points) for 16 px at natural size", {
  res <- build_diagram_xml(fix_svg_data(style = fix_style_list(font_size_px = 16)))
  sz <- xml2::xml_attr(xml2::xml_find_all(fix_xml(res$xml), "//*[local-name()='sz']"),
                       "w:val", fix_ns(fix_xml(res$xml)))
  expect_true(all(sz == "24"))
})

test_that("F6: max_scale passes through mermaid_to_word() and body_add_mermaid() (offline)", {
  svg <- fix_read("default-theme.svg")
  testthat::local_mocked_bindings(
    render_mermaid = function(mermaid_code, timeout = 120L) list(ast = NULL, svg = svg)
  )
  natural <- extent_cx(mermaid_to_word("flowchart TB\n A --> B"))
  filled  <- extent_cx(mermaid_to_word("flowchart TB\n A --> B", max_scale = Inf))
  expect_gt(filled, natural)

  doc <- body_add_mermaid(officer::read_docx(), "flowchart TB\n A --> B", max_scale = Inf)
  body <- officer::docx_body_xml(doc)
  cx <- as.numeric(xml2::xml_attr(
    xml2::xml_find_first(body, "//*[local-name()='wgp']/*[local-name()='grpSpPr']/*[local-name()='xfrm']/*[local-name()='ext']"), "cx"))
  expect_equal(cx, filled)
})

# F4 (paragraph spacing in text boxes), F8 (bold / italic runs), F5 (text size).

label_ctx <- function() list(font_family = "Arial", font_size_hp = 20L)

paras_of <- function(frag) {
  doc <- fix_frag(frag)
  list(doc = doc, ns = fix_ns(doc),
       p = xml2::xml_find_all(doc, "//*[local-name()='txbxContent']/*[local-name()='p']"))
}

run_info <- function(p, ns) {
  rs <- xml2::xml_find_all(p, "./*[local-name()='r']")
  tibble::tibble(
    text = vapply(rs, function(r)
      xml2::xml_text(xml2::xml_find_first(r, ".//*[local-name()='t']")), character(1)),
    bold = vapply(rs, function(r)
      !inherits(xml2::xml_find_first(r, "./*[local-name()='rPr']/*[local-name()='b']"), "xml_missing"),
      logical(1)),
    italic = vapply(rs, function(r)
      !inherits(xml2::xml_find_first(r, "./*[local-name()='rPr']/*[local-name()='i']"), "xml_missing"),
      logical(1))
  )
}

expect_reset_ppr <- function(p, ns) {
  sp <- xml2::xml_find_first(p, "./*[local-name()='pPr']/*[local-name()='spacing']")
  ind <- xml2::xml_find_first(p, "./*[local-name()='pPr']/*[local-name()='ind']")
  jc <- xml2::xml_find_first(p, "./*[local-name()='pPr']/*[local-name()='jc']")
  expect_false(inherits(sp, "xml_missing"))
  expect_false(inherits(ind, "xml_missing"))
  expect_false(inherits(jc, "xml_missing"))
  expect_equal(xml2::xml_attr(sp, "w:before", ns), "0")
  expect_equal(xml2::xml_attr(sp, "w:after", ns), "0")
  expect_equal(xml2::xml_attr(sp, "w:line", ns), "240")
  expect_equal(xml2::xml_attr(sp, "w:lineRule", ns), "auto")
  expect_equal(xml2::xml_attr(ind, "w:left", ns), "0")
  expect_equal(xml2::xml_attr(ind, "w:right", ns), "0")
  expect_equal(xml2::xml_attr(ind, "w:firstLine", ns), "0")
}

# ── F4 ─────────────────────────────────────────────────────────────────────

test_that("F4: label_txbx_xml paragraphs carry spacing and indent resets with the jc", {
  r <- paras_of(label_txbx_xml("one<br/>two", "000000", "center", label_ctx()))
  expect_length(r$p, 2L)
  for (p in r$p) expect_reset_ppr(p, r$ns)
})

test_that("F4: every text-box paragraph in a whole diagram has the resets", {
  data <- fix_parse("base-segoe-subgraph")          # nodes + subgraph label
  data2 <- fix_parse("base-segoe-stadium-bold-invisible")  # multi-line node + edge label
  for (d in list(data, data2)) {
    doc <- fix_xml(build_diagram_xml(d)$xml)
    ps <- xml2::xml_find_all(doc, "//*[local-name()='txbxContent']/*[local-name()='p']")
    expect_gt(length(ps), 3L)
    for (p in ps) expect_reset_ppr(p, fix_ns(doc))
  }
})

test_that("F4: edge-label and subgraph-label boxes are included", {
  doc <- fix_xml(build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))$xml)
  el <- fix_shape(doc, "mermaid:edge_label:B:C")
  expect_false(inherits(el, "xml_missing"))
  p <- xml2::xml_find_first(el, ".//*[local-name()='txbxContent']/*[local-name()='p']")
  expect_reset_ppr(p, fix_ns(doc))

  doc2 <- fix_xml(build_diagram_xml(fix_parse("base-segoe-subgraph"))$xml)
  sg <- fix_shape(doc2, "mermaid:subgraph:Calls")
  expect_false(inherits(sg, "xml_missing"))
  p2 <- xml2::xml_find_first(sg, ".//*[local-name()='txbxContent']/*[local-name()='p']")
  expect_reset_ppr(p2, fix_ns(doc2))
})

# ── F8 ─────────────────────────────────────────────────────────────────────

test_that("F8: <b> and <i> split a line into runs with w:b / w:i", {
  r <- paras_of(label_txbx_xml("Call <b>000</b> now<br/>then <i>wait</i> for it",
                               "000000", "center", label_ctx()))
  expect_length(r$p, 2L)
  r1 <- run_info(r$p[[1]], r$ns)
  expect_equal(r1$text, c("Call ", "000", " now"))
  expect_equal(r1$bold, c(FALSE, TRUE, FALSE))
  expect_false(any(r1$italic))
  r2 <- run_info(r$p[[2]], r$ns)
  expect_equal(r2$text, c("then ", "wait", " for it"))
  expect_equal(r2$italic, c(FALSE, TRUE, FALSE))
  expect_false(any(r2$bold))
})

test_that("F8: plain labels give one run per line; other tags stripped; entities kept", {
  r <- paras_of(label_txbx_xml("plain one<br/>A &amp; <u>B</u> &lt;x&gt;",
                               "000000", "center", label_ctx()))
  expect_length(r$p, 2L)
  r1 <- run_info(r$p[[1]], r$ns)
  expect_equal(r1$text, "plain one")
  r2 <- run_info(r$p[[2]], r$ns)
  expect_equal(paste(r2$text, collapse = ""), "A & B <x>")
  expect_false(any(r2$bold | r2$italic))
})

test_that("F8: nested/adjacent bold and italic are handled", {
  r <- paras_of(label_txbx_xml("<b>all bold</b>", "000000", "center", label_ctx()))
  r1 <- run_info(r$p[[1]], r$ns)
  expect_equal(paste(r1$text, collapse = ""), "all bold")
  expect_true(all(r1$bold))
})

test_that("F8: node label bold/italic survive parse + build (B in the fixture)", {
  doc <- fix_xml(build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))$xml)
  ns <- fix_ns(doc)
  b <- fix_shape(doc, "mermaid:node:B")
  ps <- xml2::xml_find_all(b, ".//*[local-name()='txbxContent']/*[local-name()='p']")
  expect_length(ps, 2L)
  r1 <- run_info(ps[[1]], ns)
  expect_equal(r1$text, c("Call ", "000", " now"))
  expect_equal(r1$bold, c(FALSE, TRUE, FALSE))
  r2 <- run_info(ps[[2]], ns)
  expect_true(any(r2$italic & r2$text == "wait"))
  expect_false(any(r2$italic & r2$text != "wait"))
})

test_that("F8: edge label bold survives parse + build (<b>Yes</b>, go)", {
  doc <- fix_xml(build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))$xml)
  el <- fix_shape(doc, "mermaid:edge_label:B:C")
  p <- xml2::xml_find_first(el, ".//*[local-name()='txbxContent']/*[local-name()='p']")
  r1 <- run_info(p, fix_ns(doc))
  expect_equal(r1$text, c("Yes", ", go"))
  expect_equal(r1$bold, c(TRUE, FALSE))
})

# ── F5 ─────────────────────────────────────────────────────────────────────

wide <- function(d, width) { d$viewbox[3] <- width; d }
all_sz <- function(xml) {
  doc <- fix_xml(xml)
  xml2::xml_attr(xml2::xml_find_all(doc, "//*[local-name()='sz']"), "w:val", fix_ns(doc))
}

test_that("F5: proportional text below 7 pt is kept (no floor of 14 half-points) and warns", {
  d <- fix_svg_data(vb = c(0, 0, 1600, 150))   # scale ~3583 EMU/px -> 14 px = 7.9 hp
  expect_warning(res <- build_diagram_xml(d), class = "mermaid2drawml_small_text")
  # F14: rounded down, so 7.9 half-points -> 7
  expect_true(all(all_sz(res$xml) == "7"))
})

test_that("F5: the warning gives the size in points and suggests remedies", {
  d <- fix_svg_data(vb = c(0, 0, 1600, 150))
  cnd <- tryCatch(build_diagram_xml(d), mermaid2drawml_small_text = function(c) c)
  expect_s3_class(cnd, "mermaid2drawml_small_text")
  msg <- conditionMessage(cnd)
  expect_match(msg, "3.5", fixed = TRUE)   # 7 half-points (F14 rounds down)
  expect_match(msg, "pt|point", ignore.case = TRUE)
  expect_match(msg, "landscape", ignore.case = TRUE)
})

test_that("F5: font_scale is applied to the unfloored size", {
  d <- fix_svg_data(vb = c(0, 0, 1600, 150))
  res <- suppressWarnings(build_diagram_xml(d, font_scale = 0.5))
  # 7.9 x 0.5 = 3.95 half-points, rounded down (F14) -> 3
  expect_true(all(all_sz(res$xml) == "3"))
})

test_that("F5: no warning when the text is 7 pt or larger", {
  expect_no_warning(build_diagram_xml(fix_svg_data()),
                    class = "mermaid2drawml_small_text")
  # 14 half-points exactly (7 pt) is not 'below 14'
  expect_no_warning(
    build_diagram_xml(fix_svg_data(vb = c(0, 0, 1600, 150),
                                   style = fix_style_list(font_size_px = 28))),
    class = "mermaid2drawml_small_text")
})

test_that("F5: subgraph labels follow the same rule (no floor), also from classDef size", {
  d <- wide(fix_parse("base-segoe-subgraph"), 1800)   # 16 px -> ~8 hp
  expect_warning(res <- build_diagram_xml(d), class = "mermaid2drawml_small_text")
  expect_true(all(all_sz(res$xml) == "8"))

  d$subgraphs$font_size_px <- 16
  expect_warning(res2 <- build_diagram_xml(d), class = "mermaid2drawml_small_text")
  expect_true(all(all_sz(res2$xml) == "8"))
})

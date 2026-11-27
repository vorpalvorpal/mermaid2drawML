# Second round: F13 (roundRect text-overlay group), F14 (font size rounds down).

# ── F13 ────────────────────────────────────────────────────────────────────

node_group <- function(doc, id) {
  xml2::xml_find_first(
    doc, sprintf("//*[local-name()='grpSp'][*[local-name()='cNvPr'][@name='mermaid:node:%s']]", id))
}
prst_of <- function(w) xml2::xml_attr(
  xml2::xml_find_first(w, ".//*[local-name()='prstGeom']"), "prst")
ext_of <- function(w) {
  e <- xml2::xml_find_first(w, "./*[local-name()='spPr']/*[local-name()='xfrm']/*[local-name()='ext']")
  as.numeric(c(xml2::xml_attr(e, "cx"), xml2::xml_attr(e, "cy")))
}
check_roundrect_group <- function(doc, id, adj = NULL) {
  grp <- node_group(doc, id)
  expect_false(inherits(grp, "xml_missing"))
  wsps <- xml2::xml_find_all(grp, "./*[local-name()='wsp']")
  expect_length(wsps, 2L)
  prst <- vapply(wsps, prst_of, character(1))
  expect_setequal(prst, c("roundRect", "rect"))
  vis <- wsps[prst == "roundRect"][[1]]
  txt <- wsps[prst == "rect"][[1]]
  expect_true(inherits(
    xml2::xml_find_first(vis, ".//*[local-name()='txbxContent']"), "xml_missing"))
  expect_false(inherits(
    xml2::xml_find_first(txt, ".//*[local-name()='txbxContent']"), "xml_missing"))
  expect_false(inherits(
    xml2::xml_find_first(txt, "./*[local-name()='spPr']/*[local-name()='noFill']"),
    "xml_missing"))
  if (!is.null(adj)) {
    gd <- xml2::xml_attr(xml2::xml_find_first(
      vis, ".//*[local-name()='avLst']/*[local-name()='gd']"), "fmla")
    expect_equal(gd, adj)
  }
  list(grp = grp, vis = vis, txt = txt)
}

test_that("F13: a parsed stadium is a group: roundRect visual (adj 50000) + rect overlay", {
  doc <- fix_xml(build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))$xml)
  g <- check_roundrect_group(doc, "A", adj = "val 50000")
  expect_true(nzchar(trimws(xml2::xml_text(g$txt))))
})

test_that("F13: the stadium overlay is 1.15 x foreignObject size, within the shape", {
  dat <- fix_parse("base-segoe-stadium-bold-invisible")
  a <- dat$nodes[dat$nodes$id == "A", ]
  doc <- fix_xml(build_diagram_xml(dat)$xml)
  g <- check_roundrect_group(doc, "A")
  vis <- ext_of(g$vis); txt <- ext_of(g$txt)
  expect_lte(txt[1], vis[1] + 1); expect_lte(txt[2], vis[2] + 1)
  # width ratio to the visual equals min(1, 1.15 * txt_w / svg_w)
  expect_equal(txt[1] / vis[1], min(1, 1.15 * a$txt_w / a$svg_w), tolerance = 0.01)
})

test_that("F13: a rounded rect (D) is also a group, keeping its adj", {
  doc <- fix_xml(build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))$xml)
  check_roundrect_group(doc, "D", adj = paste0("val ", round(100000 * 5 / 54)))
})

test_that("F13: the stadium label text sits in the overlay", {
  doc <- fix_xml(build_diagram_xml(fix_parse("base-segoe-stadium-bold-invisible"))$xml)
  g <- check_roundrect_group(doc, "A")
  expect_equal(trimws(xml2::xml_text(g$txt)), trimws(xml2::xml_text(g$grp)))
})

test_that("F13: a synthetic roundRect with svg_rx gets the group; overlay capped at shape size", {
  dat <- fix_svg_data(shapes = c("roundRect", "rect"))
  dat$nodes$svg_rx <- c(10, NA_real_)
  dat$nodes$txt_w[1] <- 78                       # 78 * 1.15 > svg_w 80
  doc <- fix_xml(build_diagram_xml(dat)$xml)
  g <- check_roundrect_group(doc, "A")
  vis <- ext_of(g$vis); txt <- ext_of(g$txt)
  expect_lte(txt[1], vis[1] + 1)
  expect_equal(txt[1], vis[1], tolerance = 0.01)
  expect_false(inherits(node_group(doc, "A"), "xml_missing"))
  # the plain rect stays a single shape, not a group
  expect_true(inherits(node_group(doc, "B"), "xml_missing"))
})

# ── F14 ────────────────────────────────────────────────────────────────────

test_that("F14: .font_hp rounds down", {
  expect_equal(.font_hp(16, 9525, 1), 24L)
  expect_equal(.font_hp(14.6, 6350, 1), 14L)
  expect_equal(.font_hp(14.99, 6350, 1), 14L)
  expect_equal(.font_hp(15, 6350, 1), 15L)
  expect_equal(.font_hp(14, 9525, 1), 21L)
})

test_that("F14: an exact size survives float noise; font_scale applies; minimum 2", {
  expect_equal(.font_hp(24, 6350, 1), 24L)
  expect_equal(.font_hp(0.1 * 3 * 80, 6350, 1), 24L)   # 24.000000000000004-ish
  expect_equal(.font_hp(12, 6350, 2), 24L)
  expect_equal(.font_hp(1, 6350, 1), 2L)
  expect_equal(.font_hp(2.9, 6350, 1), 2L)
})

test_that("F14: w:sz in the output is floored for a fractional proportional size", {
  # 14.5 px at natural scale (9525) = 21.75 half-points: round gives 22, floor 21
  dat <- fix_svg_data(style = fix_style_list(font_size_px = 14.5))
  doc <- fix_xml(build_diagram_xml(dat)$xml)
  sz <- unique(xml2::xml_attr(
    xml2::xml_find_all(doc, "//*[local-name()='sz']"), "val"))
  expect_true("21" %in% sz)
  expect_false("22" %in% sz)
})

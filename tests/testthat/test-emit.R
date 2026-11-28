# build_diagram_xml() (spec 4): the inline wrapper, scale and placement, ids,
# shape and line markup, alt text, deprecated arguments, warnings; then the
# code-review behaviours on flowchart-paint, a sweep of every fixture and a
# few checks per diagram type.

PAGE_W_EMU <- 6.27 * 914400   # 5733288
PAGE_H_EMU <- 9.69 * 914400   # 8860536

one_shape_doc <- function(prim, ..., viewBox = c(0, 0, 200, 100)) {
  generic_doc(list(prim), ..., viewBox = viewBox)
}

the_shape <- function(doc) {
  s <- all_wsp(doc)
  testthat::expect_length(s, 1L)
  s[[1]]
}

int_attr_ok <- function(v) all(grepl("^-?[0-9]+$", v))

# -- Return value and wrapper markup -------------------------------------------

test_that("build_diagram_xml() returns list(xml, next_id) with one parseable w:p", {
  res <- build_diagram_xml(generic_render(list(prim_rect(10, 10, 50, 20))))
  expect_type(res, "list")
  expect_true(all(c("xml", "next_id") %in% names(res)))
  expect_type(res$xml, "character")
  expect_length(res$xml, 1L)
  expect_true(rlang::is_integerish(res$next_id))
  doc <- dml_doc(res$xml)
  expect_equal(xml2::xml_name(xml2::xml_root(doc), ns = dml_ns), "w:p")
})

test_that("the required namespaces are declared", {
  doc <- dml_doc(build_diagram_xml(generic_render(list(prim_rect(10, 10, 50, 20))))$xml)
  declared <- unname(xml2::xml_ns(doc))
  for (p in c("w", "wp", "wps", "wpg", "a", "mc")) {
    expect_true(dml_ns[[p]] %in% declared, info = p)
  }
})

test_that("the drawing is wp:inline inside mc:AlternateContent / mc:Choice Requires=wpg", {
  doc <- generic_doc(list(prim_rect(10, 10, 50, 20)))
  inline <- xfind(doc, "/w:p/w:r/mc:AlternateContent/mc:Choice[@Requires='wpg']/w:drawing/wp:inline")
  expect_length(inline, 1L)
  expect_length(xfind(doc, "//mc:AlternateContent/mc:Fallback"), 1L)
  for (absent in c("wp:anchor", "wp:simplePos", "wp:positionH", "wp:positionV",
                   "wp:wrapNone", "wp:wrapSquare", "wp:wrapTopAndBottom", "wp:posOffset")) {
    expect_length(xfind(doc, paste0("//", absent)), 0L)
  }
  gd <- xfind(doc, paste0("//wp:inline/a:graphic/a:graphicData",
                          "[@uri='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup']",
                          "/wpg:wgp"))
  expect_length(gd, 1L)
})

test_that("the root wpg:wgp has no cNvPr and every nested wpg:grpSp has one", {
  doc <- fixture_doc("flowchart-full")
  expect_length(xfind(doc, "//wpg:wgp/wpg:cNvPr"), 0L)
  groups <- xfind(doc, "//wpg:grpSp")
  expect_gt(length(groups), 1L)
  for (g in as.list(groups)) expect_false(is_missing_node(xfirst(g, "./wpg:cNvPr")))
  expect_length(xfind(doc, "//wpg:wgp//wpg:wgp"), 0L)
})

test_that("wp:extent equals the root group size", {
  doc <- fixture_doc("flowchart-paint")
  expect_equal(inline_extent(doc), root_ext(doc))
})

test_that("align becomes the holder paragraph's w:jc, and the holder has single spacing", {
  r <- generic_render(list(prim_rect(10, 10, 50, 20)))
  for (al in c("center", "left", "right")) {
    doc <- dml_doc(build_diagram_xml(r, align = al)$xml)
    expect_equal(xattr(xfirst(doc, "/w:p/w:pPr/w:jc"), "w:val"), al)
  }
  sp <- xfirst(dml_doc(build_diagram_xml(r)$xml), "/w:p/w:pPr/w:spacing")
  expect_equal(xattr(sp, "w:before"), "0")
  expect_equal(xattr(sp, "w:after"), "0")
  expect_equal(xattr(sp, "w:line"), "240")
  expect_equal(xattr(sp, "w:lineRule"), "auto")
  expect_equal(xattr(xfirst(dml_doc(build_diagram_xml(r)$xml), "/w:p/w:pPr/w:jc"), "w:val"),
               "center")
  expect_error(build_diagram_xml(r, align = "middle"))
})

test_that("a scene gives the same XML as the render it came from", {
  r <- fixture_render("flowchart-paint")
  expect_identical(build_diagram_xml(build_scene(r))$xml, build_diagram_xml(r)$xml)
})

test_that("an empty or zero-size diagram still yields a valid w:p", {
  res <- build_diagram_xml(generic_render(list()), start_id = 5L)
  doc <- dml_doc(res$xml)
  expect_equal(xml2::xml_name(xml2::xml_root(doc), ns = dml_ns), "w:p")
  expect_equal(xattr(doc_pr(doc), "id"), "5")
  expect_equal(res$next_id, 6L)

  res0 <- build_diagram_xml(generic_render(list(), viewBox = c(0, 0, 0, 0)))
  doc0 <- dml_doc(res0$xml)
  expect_length(xfind(doc0, "//wp:inline"), 1L)
  expect_true(int_attr_ok(xattr(xfind(doc0, "//wp:extent"), "cx")))
})

# -- Scale and placement -------------------------------------------------------

test_that("by default a small diagram is drawn at natural size (9525 EMU per px)", {
  doc <- generic_doc(list(prim_rect(10, 20, 100, 40)))
  expect_equal(unname(inline_extent(doc)), c(200, 100) * 9525)
  s <- the_shape(doc)
  expect_equal(unname(shape_off(s)), c(10, 20) * 9525)
  expect_equal(unname(shape_ext(s)), c(100, 40) * 9525)
})

test_that("max_scale = Inf fills the page; a smaller max_scale shrinks", {
  r <- generic_render(list(prim_rect(10, 20, 100, 40)))
  big <- dml_doc(build_diagram_xml(r, max_scale = Inf)$xml)
  s <- PAGE_W_EMU / 200
  expect_near(unname(inline_extent(big)), c(200, 100) * s, tol = 1)
  expect_near(unname(shape_off(the_shape(big))), c(10, 20) * s, tol = 1)

  half <- dml_doc(build_diagram_xml(r, max_scale = 0.5)$xml)
  expect_near(unname(inline_extent(half)), c(200, 100) * 9525 * 0.5, tol = 1)
})

test_that("a diagram wider than the page is shrunk to fit", {
  doc <- suppressWarnings(generic_doc(list(prim_rect(1000, 0, 500, 50)),
                                      viewBox = c(0, 0, 2000, 100)))
  s <- PAGE_W_EMU / 2000
  expect_near(unname(inline_extent(doc)), c(2000, 100) * s, tol = 1)
  expect_near(unname(shape_off(the_shape(doc))), c(1000, 0) * s, tol = 1)
  expect_near(unname(shape_ext(the_shape(doc))), c(500, 50) * s, tol = 1)
})

test_that("a diagram taller than the page is shrunk to the page height", {
  doc <- suppressWarnings(generic_doc(list(prim_rect(0, 0, 50, 100)),
                                      viewBox = c(0, 0, 100, 2000)))
  s <- PAGE_H_EMU / 2000
  expect_near(unname(inline_extent(doc)), c(100, 2000) * s, tol = 1)
})

test_that("page_width_in and page_height_in set the page", {
  r <- generic_render(list(prim_rect(10, 20, 100, 40)))
  doc <- dml_doc(build_diagram_xml(r, page_width_in = 3, max_scale = Inf)$xml)
  expect_near(inline_extent(doc)[["cx"]], 3 * 914400, tol = 1)
  doc <- dml_doc(build_diagram_xml(r, page_height_in = 1, max_scale = Inf)$xml)
  expect_near(inline_extent(doc)[["cy"]], 914400, tol = 1)
})

test_that("items are placed relative to the viewBox origin", {
  doc <- generic_doc(list(prim_rect(-50, -20, 40, 20), prim_rect(0, 0, 10, 10)),
                     viewBox = c(-50, -20, 200, 100))
  shapes <- all_wsp(doc)
  expect_equal(unname(shape_off(shapes[[1]])), c(0, 0))
  expect_equal(unname(shape_off(shapes[[2]])), c(50, 20) * 9525)
})

test_that("offsets and sizes are integers and sizes are never below 1 EMU", {
  doc <- generic_doc(list(prim_rect(10.3333, 20.6667, 100.1111, 40.4444),
                          prim_line(10, 50, 150, 50)))
  offs <- xfind(doc, "//a:xfrm/a:off")
  exts <- xfind(doc, "//a:xfrm/a:ext")
  expect_true(int_attr_ok(c(xattr(offs, "x"), xattr(offs, "y"))))
  expect_true(int_attr_ok(c(xattr(exts, "cx"), xattr(exts, "cy"))))
  expect_true(all(as.numeric(c(xattr(exts, "cx"), xattr(exts, "cy"))) >= 1))
})

test_that("nested groups keep diagram coordinates for their children", {
  doc <- fixture_doc("flowchart-paint")   # natural size
  a <- shape_named(doc, "mermaid:node:A")
  expect_true(length(xfind(a, "ancestor::wpg:grpSp")) >= 1L)
  expect_near(unname(abs_box(a)), c(254.227, 137, 132.25, 54) * 9525, tol = 3)

  full <- fixture_doc("flowchart-full")   # shrunk: 2088.957 EMU per px
  s3 <- shape_named(full, "mermaid:node:s3")
  expect_gte(length(xfind(s3, "ancestor::wpg:grpSp")), 2L)
  s <- min(PAGE_W_EMU / fixture_render("flowchart-full")$viewBox[3], 9525)
  expect_near(unname(abs_box(s3)), c(382.75, 63, 74.875, 54) * s, tol = 4)
})

# -- Ids -----------------------------------------------------------------------

test_that("docPr gets start_id and every cNvPr the next ids in document order", {
  for (nm in c("flowchart-paint", "classDiagram")) {
    res <- fixture_xml(nm, start_id = 7L)
    ids <- all_ids(dml_doc(res$xml))
    expect_equal(ids[1], 7L, info = nm)
    expect_equal(ids, seq.int(7L, length.out = length(ids)), info = nm)
    expect_equal(res$next_id, 7L + length(ids), info = nm)
    expect_equal(xattr(doc_pr(dml_doc(res$xml)), "id"), "7")
  }
})

test_that("two diagrams chained by next_id share no ids", {
  r1 <- fixture_xml("flowchart-paint", start_id = 1L)
  r2 <- fixture_xml("default-theme", start_id = r1$next_id)
  ids1 <- all_ids(dml_doc(r1$xml))
  ids2 <- all_ids(dml_doc(r2$xml))
  expect_length(intersect(ids1, ids2), 0L)
  expect_equal(min(ids2), r1$next_id)
})

test_that("docPr is named mermaid:diagram", {
  expect_equal(xattr(doc_pr(fixture_doc("flowchart-paint")), "name"), "mermaid:diagram")
})

# -- Shape markup --------------------------------------------------------------

test_that("a preset shape has prstGeom with an avLst", {
  s <- the_shape(one_shape_doc(prim_rect(10, 20, 100, 40)))
  expect_equal(shape_prst(s), "rect")
  expect_false(is_missing_node(xfirst(s, "./wps:spPr/a:prstGeom/a:avLst")))
  expect_false(shape_is_custom(s))
})

test_that("a rect with rx gives roundRect adj = round(100000 * rx / min(w, h)), capped at 50000", {
  s <- the_shape(one_shape_doc(prim_rect(10, 20, 100, 40, rx = 8)))
  expect_equal(shape_prst(s), "roundRect")
  expect_equal(shape_adj(s)[["adj"]], 20000)
  s <- the_shape(one_shape_doc(prim_rect(10, 20, 100, 40, rx = 7)))
  expect_equal(shape_adj(s)[["adj"]], round(100000 * 7 / 40))
  s <- the_shape(one_shape_doc(prim_rect(10, 20, 100, 40, rx = 30)))
  expect_equal(shape_adj(s)[["adj"]], 50000)
})

test_that("a closed curved path is custGeom with one path of the box size", {
  p <- prim_path("M0,0 L100,0 C100,50 0,50 0,0 Z", m = c(1, 0, 0, 1, 20, 30),
                 bbox = c(20, 30, 100, 37.5))
  s <- the_shape(one_shape_doc(p))
  expect_true(shape_is_custom(s))
  expect_true(is.na(shape_prst(s)))
  paths <- xfind(s, "./wps:spPr/a:custGeom/a:pathLst/a:path")
  expect_length(paths, 1L)
  expect_near(xnum(paths, "w"), 100 * 9525, tol = 1)
  expect_near(xnum(paths, "h"), 37.5 * 9525, tol = 1)
  expect_near(unname(shape_ext(s)), c(100, 37.5) * 9525, tol = 1)
  expect_near(unname(shape_off(s)), c(20, 30) * 9525, tol = 1)

  pt <- function(xp) {
    n <- xfind(paths[[1]], xp)
    cbind(xnum(n, "x"), xnum(n, "y"))
  }
  expect_near(pt("./a:moveTo/a:pt")[1, ], c(0, 0), tol = 1)
  expect_near(pt("./a:lnTo/a:pt")[1, ], c(100 * 9525, 0), tol = 1)
  bez <- pt("./a:cubicBezTo/a:pt")
  expect_equal(nrow(bez), 3L)
  expect_near(bez[3, ], c(0, 0), tol = 1)
  expect_length(xfind(paths[[1]], "./a:close"), 1L)
})

test_that("arcs are written as cubic Beziers", {
  p <- prim_path("M10,50 A40,40 0 0 1 90,50 Z", bbox = c(10, 10, 80, 40))
  s <- the_shape(one_shape_doc(p))
  expect_gte(length(xfind(s, ".//a:path/a:cubicBezTo")), 2L)
  expect_length(xfind(s, ".//a:path/a:arcTo"), 0L)
})

test_that("fill: solidFill with alpha below 1, noFill for none", {
  s <- the_shape(one_shape_doc(prim_rect(10, 20, 100, 40, fill = "#ABCDEF", fill_a = 0.25)))
  f <- shape_fill(s)
  expect_equal(f$kind, "solid")
  expect_equal(f$color, "ABCDEF")
  expect_equal(f$alpha, 25000)

  s <- the_shape(one_shape_doc(prim_rect(10, 20, 100, 40, fill = "#ABCDEF", fill_a = 1)))
  expect_true(is.na(shape_fill(s)$alpha))
  expect_length(xfind(s, "./wps:spPr/a:solidFill//a:alpha"), 0L)

  s <- the_shape(one_shape_doc(prim_rect(10, 20, 100, 40, fill = NULL)))
  expect_equal(shape_fill(s)$kind, "none")
})

test_that("line: width in EMU, colour, alpha; noFill when there is no stroke", {
  s <- the_shape(one_shape_doc(prim_rect(10, 20, 100, 40, stroke = "#123456", stroke_w = 2,
                                         stroke_a = 0.6)))
  ln <- shape_ln(s)
  expect_equal(ln$w, 2 * 9525)
  expect_equal(ln$color, "123456")
  expect_equal(ln$alpha, 60000)

  s <- the_shape(one_shape_doc(prim_rect(10, 20, 100, 40, stroke = NULL)))
  ln <- shape_ln(s)
  expect_true(ln$present)
  expect_equal(ln$kind, "none")
})

test_that("min_stroke_pt floors thin lines", {
  thin <- prim_rect(10, 20, 100, 40, stroke_w = 0.5)
  expect_equal(shape_ln(the_shape(one_shape_doc(thin)))$w, 0.75 * 12700)
  expect_near(shape_ln(the_shape(one_shape_doc(thin, min_stroke_pt = 0)))$w, 0.5 * 9525,
               tol = 1)
  expect_equal(shape_ln(the_shape(one_shape_doc(thin, min_stroke_pt = 2)))$w, 2 * 12700)
  # a 1 px line on a shrunk diagram would be 2866 EMU; it is floored too
  wide <- suppressWarnings(one_shape_doc(prim_rect(10, 20, 1000, 40, stroke_w = 1),
                                         viewBox = c(0, 0, 2000, 100)))
  expect_equal(shape_ln(the_shape(wide))$w, 0.75 * 12700)
})

test_that("the dash preset follows the ratio of the first dash to the stroke width", {
  dash_of <- function(dash, w = 1) {
    shape_ln(the_shape(one_shape_doc(prim_line(10, 50, 150, 50, dash = dash, stroke_w = w))))$dash
  }
  expect_equal(dash_of(1), "sysDot")
  expect_equal(dash_of(1.5), "sysDot")
  expect_equal(dash_of(c(2, 2)), "sysDash")
  expect_equal(dash_of(3.5), "sysDash")
  expect_equal(dash_of(c(5, 5)), "dash")
  expect_equal(dash_of(6), "dash")
  expect_equal(dash_of(c(9, 5)), "lgDash")
  expect_equal(dash_of(c(6, 3), w = 2), "sysDash")   # ratio 3
  expect_true(is.na(dash_of(NULL)))
})

test_that("line ends: headEnd for the start, tailEnd for the end, w and len med", {
  s <- the_shape(one_shape_doc(prim_line(10, 50, 150, 50,
                                         marker_start = "class-aggregationStart",
                                         marker_end = "flowchart-v2-pointEnd")))
  ln <- shape_ln(s)
  expect_equal(ln$head, "diamond")
  expect_equal(ln$tail, "triangle")
  expect_equal(c(ln$head_w, ln$head_len, ln$tail_w, ln$tail_len), rep("med", 4))

  s <- the_shape(one_shape_doc(prim_line(10, 50, 150, 50, marker_end = "class-dependencyEnd")))
  ln <- shape_ln(s)
  expect_equal(ln$head, "none")
  expect_equal(ln$tail, "arrow")
})

test_that("a line is preset line, flipped so that its start stays its start", {
  flips <- function(x1, y1, x2, y2) {
    xf <- shape_xfrm(the_shape(one_shape_doc(prim_line(x1, y1, x2, y2,
                                                       marker_end = "arrowhead"))))
    c(h = xf$flipH, v = xf$flipV)
  }
  s <- the_shape(one_shape_doc(prim_line(20, 10, 100, 80)))
  expect_equal(shape_prst(s), "line")
  expect_equal(unname(shape_off(s)), c(20, 10) * 9525)
  expect_equal(unname(shape_ext(s)), c(80, 70) * 9525)
  expect_equal(flips(20, 10, 100, 80), c(h = FALSE, v = FALSE))   # down-right
  expect_equal(flips(100, 80, 20, 10), c(h = TRUE, v = TRUE))     # up-left
  expect_equal(flips(20, 80, 100, 10), c(h = FALSE, v = TRUE))    # up-right
  expect_equal(flips(100, 10, 20, 80), c(h = TRUE, v = FALSE))    # down-left
  expect_equal(flips(100, 50, 20, 50)[["h"]], TRUE)               # leftwards
})

test_that("rotation is written in 60000ths of a degree, 0 to 21599999", {
  rot_of <- function(prim) shape_xfrm(the_shape(one_shape_doc(prim, viewBox = c(0, 0, 200, 200))))$rot
  expect_equal(rot_of(prim_rect(0, 0, 60, 20, m = c(0, 1, -1, 0, 100, 50))), 90 * 60000)
  expect_equal(rot_of(prim_text("t", x = 50, y = 50, w = 40, h = 20, rot = 30)), 30 * 60000)
  expect_equal(rot_of(prim_text("t", x = 50, y = 50, w = 40, h = 20, rot = -90)), 270 * 60000)
  r <- shape_xfrm(the_shape(one_shape_doc(prim_rect(0, 0, 60, 20, m = c(0, 1, -1, 0, 100, 50)),
                                          viewBox = c(0, 0, 200, 200))))
  expect_near(c(r$x, r$y, r$cx, r$cy), c(60, 70, 60, 20) * 9525, tol = 1)
})

test_that("closed shapes use cNvSpPr, lines and open paths cNvCnPr, free text txBox", {
  doc <- generic_doc(list(
    prim_rect(10, 10, 40, 20),
    prim_line(60, 10, 100, 30),
    prim_path("M110,10 L150,30 L190,10", fill = NULL),
    prim_text("free", x = 10, y = 60, w = 40, h = 20)
  ))
  s <- all_wsp(doc)
  expect_length(s, 4L)
  expect_false(is_missing_node(xfirst(s[[1]], "./wps:cNvSpPr")))
  expect_false(xattr(xfirst(s[[1]], "./wps:cNvSpPr"), "txBox") %in% "1")
  expect_false(is_missing_node(xfirst(s[[2]], "./wps:cNvCnPr")))
  expect_false(is_missing_node(xfirst(s[[3]], "./wps:cNvCnPr")))
  expect_equal(xattr(xfirst(s[[4]], "./wps:cNvSpPr"), "txBox"), "1")
})

test_that("an open path that the browser fills is written filled", {
  s <- the_shape(one_shape_doc(prim_path("M10,10 L100,10 L100,60", fill = "#FF0000")))
  expect_equal(shape_fill(s)$color, "FF0000")
})

# -- Alt text and names --------------------------------------------------------

test_that("docPr descr is alt_text, else desc, else title, else empty; title is the title", {
  prims <- list(prim_rect(10, 10, 50, 20))
  dp <- function(r, ...) doc_pr(dml_doc(build_diagram_xml(r, ...)$xml))

  both <- generic_render(prims, title = "The title", desc = "The description")
  expect_equal(xattr(dp(both), "descr"), "The description")
  expect_equal(xattr(dp(both), "title"), "The title")
  expect_equal(xattr(dp(both, alt_text = "Given alt text"), "descr"), "Given alt text")

  title_only <- generic_render(prims, title = "Only a title")
  expect_equal(xattr(dp(title_only), "descr"), "Only a title")

  neither <- generic_render(prims)
  expect_equal(xattr(dp(neither), "descr"), "")
  expect_true(xattr(dp(neither), "title") %in% c(NA, ""))
})

test_that("alt text, titles, names and labels with & < \" are escaped", {
  odd <- "a&b<\"c\">"
  r <- generic_render(list(prim_rect(10, 10, 50, 20)), title = odd, desc = odd)
  doc <- dml_doc(build_diagram_xml(r, alt_text = paste("alt", odd))$xml)
  expect_equal(xattr(doc_pr(doc), "title"), odd)
  expect_equal(xattr(doc_pr(doc), "descr"), paste("alt", odd))

  dom <- "flowchart-x-0"
  label <- "Tom & \"Jerry\" <3>"
  prims <- node_prims(dom, prim_rect(20, 20, 100, 40), label = label)
  db <- mini_db(list(db_node(odd, dom_id = dom, label = label)))
  res <- build_diagram_xml(mini_render(prims, db = db))
  doc <- dml_doc(res$xml)
  node <- shape_named(doc, paste0("mermaid:node:", odd))
  expect_false(is_missing_node(node))
  expect_equal(shape_paragraph_texts(node), label)
})

# -- Deprecated and unknown arguments ------------------------------------------

test_that("removed arguments are ignored with one deprecation warning naming them", {
  r <- generic_render(list(prim_rect(10, 10, 50, 20)))
  seen <- rlang::env(n = 0L, msg = character())
  res <- withCallingHandlers(
    build_diagram_xml(r, stroke_width_pt = 2, default_fill = "FFFFFF"),
    mermaid2drawml_deprecated = function(c) {
      seen$n <- seen$n + 1L
      seen$msg <- c(seen$msg, conditionMessage(c))
      rlang::cnd_muffle(c)
    }
  )
  expect_equal(seen$n, 1L)
  expect_match(seen$msg, "stroke_width_pt", fixed = TRUE)
  expect_match(seen$msg, "default_fill", fixed = TRUE)
  expect_no_error(dml_doc(res$xml))
  expect_identical(res$xml, build_diagram_xml(r)$xml)
})

test_that("each removed argument is accepted with a deprecation warning", {
  r <- generic_render(list(prim_rect(10, 10, 50, 20)))
  removed <- list(ast = list(), margin_in = 1, default_fill = "FFFFFF",
                  default_stroke = "000000", default_text_color = "000000",
                  stroke_width_pt = 1)
  for (arg in names(removed)) {
    expect_warning(do.call(build_diagram_xml, c(list(r), removed[arg])),
                   class = "mermaid2drawml_deprecated", label = arg)
  }
})

test_that("an unknown argument is an error", {
  r <- generic_render(list(prim_rect(10, 10, 50, 20)))
  expect_error(build_diagram_xml(r, no_such_option = 1), "no_such_option")
})

# -- Unsupported prims ---------------------------------------------------------

test_that("unsupported prims raise one warning naming what was left out", {
  r <- generic_render(list(prim_rect(0, 0, 50, 50), prim_image(60, 0, 30, 30),
                           prim_image(100, 0, 30, 30), prim_other("use", 140, 0, 20, 20)))
  seen <- rlang::env(n = 0L, msg = character())
  withCallingHandlers(build_diagram_xml(r), mermaid2drawml_unsupported = function(c) {
    seen$n <- seen$n + 1L
    seen$msg <- c(seen$msg, conditionMessage(c))
    rlang::cnd_muffle(c)
  })
  expect_equal(seen$n, 1L)
  expect_match(seen$msg, "image", fixed = TRUE)
  expect_match(seen$msg, "use", fixed = TRUE)
})

test_that("flowchart-full warns that its image was left out", {
  expect_warning(
    withCallingHandlers(build_diagram_xml(fixture_render("flowchart-full")),
                        mermaid2drawml_small_text = function(c) rlang::cnd_muffle(c)),
    class = "mermaid2drawml_unsupported")
})

test_that("a diagram with nothing unsupported does not warn", {
  expect_no_warning(build_diagram_xml(fixture_render("flowchart-paint")))
})

# -- flowchart-paint: the code-review behaviours, in the XML -------------------

test_that("flowchart-paint XML: node A is ADD8E6 with a 4 px FF0000 outline", {
  a <- shape_named(fixture_doc("flowchart-paint"), "mermaid:node:A")
  expect_equal(shape_fill(a)[c("kind", "color")], list(kind = "solid", color = "ADD8E6"))
  ln <- shape_ln(a)
  expect_equal(ln$color, "FF0000")
  expect_equal(ln$w, 4 * 9525)
})

test_that("flowchart-paint XML: D has noFill and a visible line", {
  d <- shape_named(fixture_doc("flowchart-paint"), "mermaid:node:D")
  outlines <- purrr::keep(as.list(xfind(d, "descendant-or-self::wps:wsp")),
                          function(n) !identical(xattr(xfirst(n, "./wps:cNvSpPr"), "txBox"), "1"))
  expect_gte(length(outlines), 1L)
  for (o in outlines) {
    expect_equal(shape_fill(o)$kind, "none")
    expect_equal(shape_ln(o)$kind, "solid")
  }
})

test_that("flowchart-paint XML: E's fill carries alpha 50000", {
  e <- shape_named(fixture_doc("flowchart-paint"), "mermaid:node:E")
  f <- shape_fill(e)
  expect_equal(f$color, "FF0000")
  expect_near(f$alpha, 50000, tol = 100)
})

test_that("flowchart-paint XML: the subgraph background keeps the theme fill FFFFDE", {
  bg <- shape_named(fixture_doc("flowchart-paint"), "mermaid:subgraph_bg:S1")
  expect_equal(shape_fill(bg)[c("kind", "color")], list(kind = "solid", color = "FFFFDE"))
})

test_that("flowchart-paint XML: the cylinder is can at its true height", {
  cc <- shape_named(fixture_doc("flowchart-paint"), "mermaid:node:C")
  expect_equal(shape_prst(cc), "can")
  # 68.38 is given to two decimals (68.3796 in the fixture): half a unit of
  # its last digit is 0.005 px.
  expect_near(unname(shape_ext(cc)), c(80.5, 68.38) * 9525, tol = 0.005 * 9525)
})

test_that("flowchart-paint XML: the odd node D is custom geometry, not flowChartMultidocument", {
  d <- shape_named(fixture_doc("flowchart-paint"), "mermaid:node:D")
  prsts <- xattr(xfind(d, "descendant-or-self::wps:wsp/wps:spPr/a:prstGeom"), "prst")
  expect_false("flowChartMultidocument" %in% prsts)
  expect_gte(length(xfind(d, "descendant-or-self::wps:wsp/wps:spPr/a:custGeom")), 1L)
})

test_that("flowchart-paint XML: one edge shape per drawn link, styled from the prims", {
  doc <- fixture_doc("flowchart-paint")
  edges <- shapes_named(doc, "mermaid:edge:")
  expect_length(edges, 4L)
  expect_equal(count_named(doc, "mermaid:edge:F-A") + count_named(doc, "mermaid:edge:A-F"), 0)
  for (e in as.list(edges)) {
    ln <- shape_ln(e)
    nm <- shape_name(e)
    expect_equal(ln$color, "333333", info = nm)
    expect_equal(shape_fill(e)$kind, "none", info = nm)
    expect_equal(ln$tail, "triangle", info = nm)
    # step 2: C2, the 8 px flowchart point is "lg" on a 1 px edge, "sm" at 3.5 px
    size <- if (identical(nm, "mermaid:edge:E-C")) "sm" else "lg"
    expect_equal(c(ln$tail_w, ln$tail_len), c(size, size), info = nm)
    expect_true(shape_is_custom(e), info = nm)
    expect_false(is_missing_node(xfirst(e, "./wps:cNvCnPr")), info = nm)
  }
  dash <- function(name) shape_ln(shape_named(doc, name))$dash
  expect_equal(dash("mermaid:edge:B-D"), "sysDash")       # dotted: dash 2 at width 1
  expect_true(is.na(dash("mermaid:edge:A-B")))
  expect_true(is.na(dash("mermaid:edge:B-C")))
  expect_true(is.na(dash("mermaid:edge:E-C")))
  expect_near(shape_ln(shape_named(doc, "mermaid:edge:E-C"))$w, 3.5 * 9525, tol = 1)
})

test_that("flowchart-paint XML: the edge label is a filled shape with its text inside", {
  el <- shape_named(fixture_doc("flowchart-paint"), "mermaid:edge_label:B:C")
  expect_true(is_wsp_node(el))
  f <- shape_fill(el)
  expect_equal(f$color, "E8E8E8")
  # three stacked backgrounds at 0.5, 0.8, 0.8: 1 - 0.5 * 0.2 * 0.2 = 0.98
  expect_near(f$alpha, 98000, tol = 100)
  expect_equal(shape_paragraph_texts(el), "yes")
})

test_that("flowchart-paint XML: subgraph S1 holds its background, A, B and A-B, not B-C", {
  doc <- fixture_doc("flowchart-paint")
  g <- shape_named(doc, "mermaid:subgraph:S1")
  expect_equal(xml2::xml_name(g), "grpSp")
  inside <- descendant_names(g)
  expect_true(all(c("mermaid:subgraph_bg:S1", "mermaid:node:A", "mermaid:node:B",
                    "mermaid:edge:A-B") %in% inside))
  expect_false(any(c("mermaid:edge:B-C", "mermaid:node:C", "mermaid:edge_label:B:C") %in% inside))
  first_child <- xfind(g, "./wps:wsp | ./wpg:grpSp")[[1]]
  expect_equal(shape_name(first_child), "mermaid:subgraph_bg:S1")
})

test_that("flowchart-full XML: subgraph S2 is a group nested in S1", {
  doc <- fixture_doc("flowchart-full")
  s2 <- shape_named(doc, "mermaid:subgraph:S2")
  expect_equal(xml2::xml_name(s2), "grpSp")
  expect_true("mermaid:subgraph:S2" %in% descendant_names(shape_named(doc, "mermaid:subgraph:S1")))
  expect_true(all(c("mermaid:node:s3", "mermaid:node:s4", "mermaid:edge:s3-s4") %in%
                    descendant_names(s2)))
  expect_false("mermaid:edge:s2-s3" %in% descendant_names(s2))
})

# -- Every fixture -------------------------------------------------------------

for (nm in all_fixture_names()) {
  test_that(paste0("build_diagram_xml() gives well-formed, consistent XML: ", nm), {
    r <- fixture_render(nm)
    res <- NULL
    expect_no_warning(res <- quietly_expected(build_diagram_xml(r, start_id = 3L)))
    doc <- NULL
    expect_no_error(doc <- dml_doc(res$xml))

    # ids: unique, consecutive from start_id, next_id one past the last
    ids <- all_ids(doc)
    expect_false(anyDuplicated(ids) > 0L)
    expect_equal(ids, seq.int(3L, length.out = length(ids)))
    expect_equal(res$next_id, 3L + length(ids))

    # wrapper
    expect_length(xfind(doc, "//wp:inline"), 1L)
    expect_equal(inline_extent(doc), root_ext(doc))
    for (g in as.list(xfind(doc, "//wpg:grpSp"))) {
      expect_false(is_missing_node(xfirst(g, "./wpg:cNvPr")))
    }

    # offsets and extents: integers, sizes at least 1 EMU
    offs <- xfind(doc, "//a:xfrm/a:off")
    exts <- xfind(doc, "//a:xfrm/a:ext")
    expect_true(int_attr_ok(c(xattr(offs, "x"), xattr(offs, "y"))))
    expect_true(int_attr_ok(c(xattr(exts, "cx"), xattr(exts, "cy"))))
    expect_true(all(as.numeric(c(xattr(exts, "cx"), xattr(exts, "cy"))) >= 1))

    # every shape lies within the drawing (its rotated footprint, for rotated ones)
    ext <- root_ext(doc)
    tol <- 4 * ext[["cx"]] / r$viewBox[3]          # 4 px
    for (s in as.list(all_wsp(doc))) {
      fp <- rotated_footprint(abs_box(s), shape_xfrm(s)$rot)
      ok <- fp[["x"]] >= -tol && fp[["y"]] >= -tol &&
        fp[["x"]] + fp[["w"]] <= ext[["cx"]] + tol && fp[["y"]] + fp[["h"]] <= ext[["cy"]] + tol
      expect_true(ok, info = paste0(nm, ": ", shape_name(s), " at ",
                                    paste(round(fp), collapse = ", ")))
    }

    # no label is lost: every text row is a paragraph somewhere
    rows <- unlist(purrr::map(fixture_prims(nm, "text"), prim_row_texts))
    rows <- rows[nzchar(trimws(rows))]
    paras <- all_paragraph_texts(doc)
    for (t in unique(rows)) {
      expect_gte(sum(paras == t), sum(rows == t), label = paste0(nm, ": '", t, "'"))
    }
  })
}

test_that("items wholly outside the viewBox are not drawn (gantt's today line)", {
  # The today marker is at x = 63553 in a 1584 px wide chart; the browser
  # clips it, and so must Word.
  doc <- fixture_doc("gantt")
  right_edge <- root_ext(doc)[["cx"]]
  lefts <- purrr::map_dbl(as.list(all_wsp(doc)), function(s) abs_box(s)[["x"]])
  expect_false(any(lefts > right_edge))
})

# -- Per diagram type ----------------------------------------------------------

test_that("sequenceDiagram: participant boxes hold their names; messages have line ends", {
  items <- fixture_scene("sequenceDiagram")$items
  for (who in c("Alice", "Bob")) {
    boxes <- items[items$kind == "shape" & items$label == who, ]
    expect_gte(nrow(boxes), 2L, label = paste(who, "boxes"))   # top and bottom
    expect_true(all(boxes$prst == "roundRect"), info = who)
  }
  msgs <- items[items$kind == "line" & items$tail != "none", ]
  # step 2: C1, filled-head is stealth and crosshead no Word end (a shape)
  expect_gte(nrow(msgs), 7L)
  expect_true(all(c("triangle", "stealth") %in% msgs$tail))   # arrowhead, filled-head
  expect_false("diamond" %in% msgs$tail)
  doc <- fixture_doc("sequenceDiagram")
  tails <- xfind(doc, "//a:ln/a:tailEnd[@type='triangle']")
  expect_gte(length(tails), 6L)
})

test_that("gantt: task bars are shapes with their label merged or as text beside them", {
  items <- fixture_scene("gantt")$items
  tasks <- c("D3D3D3" = "Done task", "BFC7FF" = "Active", "FF0000" = "Crit", "8A90DD" = "Future")
  for (fill in names(tasks)) {
    bars <- items[items$kind == "shape" & items$fill %in% fill & items$rot == 0, ]
    expect_gte(nrow(bars), 1L)
    label <- tasks[[fill]]
    expect_true(label %in% bars$label || label %in% items$label[items$kind == "text"],
                info = label)
  }
  # the milestone is a rect rotated 45 degrees (matrix scale 0.8)
  ms <- items[items$kind == "shape" & abs(items$rot - 45) < 1e-3, ]
  expect_equal(nrow(ms), 1L)
  expect_near(c(ms$x, ms$y, ms$w, ms$h), c(1501, 148, 16, 16), tol = 1e-2)
})

test_that("pie: the slices are filled custom-geometry shapes", {
  items <- fixture_scene("pie")$items
  slices <- items[items$kind == "shape" & items$fill %in% c("ECECFF", "FFFFDE", "B5FF20") &
                    is.na(items$prst), ]
  expect_equal(nrow(slices), 3L)
  expect_near(slices$fill_a, rep(0.7, 3), tol = 1e-3)

  doc <- fixture_doc("pie")
  custom_fills <- purrr::map_chr(
    purrr::keep(as.list(all_wsp(doc)), shape_is_custom),
    function(s) paste(shape_fill(s)$color, shape_fill(s)$alpha))
  for (col in c("ECECFF", "FFFFDE", "B5FF20")) {
    expect_true(paste(col, 70000) %in% custom_fills, info = col)
  }
})

test_that("xychart: the rotated axis title carries its rotation", {
  items <- fixture_scene("xychart-beta")$items
  title <- items[items$label == "Revenue", ]
  expect_equal(nrow(title), 1L)
  expect_equal(title$rot, -90)

  doc <- fixture_doc("xychart-beta")
  box <- purrr::keep(as.list(all_wsp(doc)),
                     function(s) identical(shape_paragraph_texts(s), "Revenue"))
  expect_length(box, 1L)
  expect_equal(shape_xfrm(box[[1]])$rot, 270 * 60000)
})

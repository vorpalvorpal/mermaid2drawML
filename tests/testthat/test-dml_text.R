# Text in the DrawingML (spec 4, "Text"): paragraphs, runs, fonts, sizes,
# body properties, and the small-text warning.

one_text_doc <- function(..., viewBox = c(0, 0, 200, 100), xml_args = list()) {
  r <- generic_render(list(prim_text(...)), viewBox = viewBox)
  dml_doc(do.call(build_diagram_xml, c(list(r), xml_args))$xml)
}

the_text_box <- function(doc) {
  boxes <- all_wsp(doc)
  testthat::expect_length(boxes, 1L)
  boxes[[1]]
}

# -- font_display_name() -------------------------------------------------------

test_that("font_display_name() gives the capitalisation Word shows", {
  expect_equal(font_display_name("trebuchet ms"), "Trebuchet MS")
  expect_equal(font_display_name("segoe ui"), "Segoe UI")
  expect_equal(font_display_name("arial"), "Arial")
  expect_equal(font_display_name("times new roman"), "Times New Roman")
  expect_equal(font_display_name("pt sans"), "PT Sans")
  expect_equal(font_display_name("noto sans sc"), "Noto Sans SC")
})

test_that("font_display_name() leaves a name with an upper-case letter alone", {
  expect_equal(font_display_name("Myriad Pro"), "Myriad Pro")
  expect_equal(font_display_name("ArialMT"), "ArialMT")
  expect_equal(font_display_name("Segoe UI"), "Segoe UI")
})

# -- Fonts ---------------------------------------------------------------------

font_attrs <- function(doc) {
  f <- xfind(doc, "//w:rPr/w:rFonts")
  list(n = length(f),
       values = unlist(lapply(c("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"),
                              function(a) xattr(f, a))))
}

test_that("every rFonts attribute is the theme font: Segoe UI fixtures", {
  for (nm in c("base-segoe-subgraph", "base-segoe-stadium-bold-invisible")) {
    fa <- font_attrs(fixture_doc(nm))
    expect_gt(fa$n, 0L)
    expect_true(all(fa$values == "Segoe UI"), info = nm)
  }
})

test_that("every rFonts attribute is the theme font: Trebuchet MS for the default theme", {
  fa <- font_attrs(fixture_doc("default-theme"))
  expect_gt(fa$n, 0L)
  expect_true(all(fa$values == "Trebuchet MS"))
})

test_that("word_font_family overrides all four rFonts attributes", {
  for (nm in c("default-theme", "base-segoe-subgraph")) {
    res <- fixture_xml(nm, word_font_family = "Myriad Pro")
    fa <- font_attrs(dml_doc(res$xml))
    expect_gt(fa$n, 0L)
    expect_true(all(fa$values == "Myriad Pro"), info = nm)
    expect_false(grepl("Trebuchet|Segoe UI", res$xml), info = nm)
  }
})

test_that("the run's font wins, else the first family, with Word's capitalisation", {
  run_font <- function(font, family) {
    doc <- one_text_doc(rows = list(list(text_run("x", font = font, family = family))),
                        x = 10, y = 10, w = 40, h = 20)
    unique(unlist(shape_runs(the_text_box(doc))[c("ascii", "hAnsi", "eastAsia", "cs")]))
  }
  expect_equal(run_font("segoe ui", "\"segoe ui\", sans-serif"), "Segoe UI")
  expect_equal(run_font(NULL, "\"georgia\", serif"), "Georgia")
  expect_equal(run_font(NULL, "'Open Sans', sans-serif"), "Open Sans")
  expect_equal(run_font("ArialMT", "ArialMT, Arial"), "ArialMT")
})

# -- Rows, runs and styling ----------------------------------------------------

test_that("each row is one paragraph and each run one w:r", {
  doc <- one_text_doc(rows = list(
    list(text_run("Call "), text_run("000", bold = TRUE), text_run(" now")),
    list(text_run("then "), text_run("wait", italic = TRUE), text_run(" for it"))
  ), x = 10, y = 10, w = 120, h = 40)
  box <- the_text_box(doc)
  expect_equal(shape_paragraph_texts(box), c("Call 000 now", "then wait for it"))
  runs <- shape_runs(box)
  expect_equal(runs$text, c("Call ", "000", " now", "then ", "wait", " for it"))
  expect_equal(runs$para, c(1, 1, 1, 2, 2, 2))
  expect_equal(runs$bold, c(FALSE, TRUE, FALSE, FALSE, FALSE, FALSE))
  expect_equal(runs$italic, c(FALSE, FALSE, FALSE, FALSE, TRUE, FALSE))
})

test_that("underline and colour reach the run properties", {
  doc <- one_text_doc(rows = list(list(text_run("plain", color = "#333333"),
                                       text_run("under", underline = TRUE, color = "#FF0000"))),
                      x = 10, y = 10, w = 80, h = 20)
  runs <- shape_runs(the_text_box(doc))
  expect_equal(runs$underline, c(FALSE, TRUE))
  expect_equal(runs$color, c("333333", "FF0000"))
})

test_that("text is emitted exactly as the bridge gives it (no markdown stripping, no entity decoding)", {
  raw <- "a &amp; b **c** __d__ <i>e</i> 2 * 3 \"q\""
  doc <- one_text_doc(raw, x = 10, y = 10, w = 180, h = 20)
  expect_equal(shape_paragraph_texts(the_text_box(doc)), raw)
})

test_that("w:t keeps leading and trailing spaces", {
  doc <- one_text_doc(rows = list(list(text_run("A "), text_run("bold ", bold = TRUE),
                                       text_run("word"))),
                      x = 10, y = 10, w = 100, h = 20)
  box <- the_text_box(doc)
  ts <- xfind(box, ".//w:t")
  expect_true(all(xml2::xml_attr(ts, "space") == "preserve"))
  expect_equal(shape_runs(box)$text, c("A ", "bold ", "word"))
})

test_that("node labels keep their bold and italic runs (base-segoe fixture)", {
  doc <- fixture_doc("base-segoe-stadium-bold-invisible")
  b <- shape_named(doc, "mermaid:node:B")
  expect_equal(shape_paragraph_texts(b), c("Call 000 now", "then wait for **help**"))
  runs <- shape_runs(b)
  expect_equal(trimws(runs$text[runs$bold]), "000")
  expect_equal(trimws(runs$text[runs$italic]), "wait")

  el <- shape_named(doc, "mermaid:edge_label:B:C")
  er <- shape_runs(el)
  expect_equal(er$text, c("Yes", ", go"))
  expect_equal(er$bold, c(TRUE, FALSE))
})

test_that("flowchart-paint labels: verbatim text, two rows, markdown bold run", {
  doc <- fixture_doc("flowchart-paint")
  expect_equal(shape_paragraph_texts(shape_named(doc, "mermaid:node:F")), "2 * 3 * 4 __init__")
  expect_equal(shape_paragraph_texts(shape_named(doc, "mermaid:node:B")), c("Decide", "now"))
  e <- shape_named(doc, "mermaid:node:E")
  expect_equal(shape_paragraph_texts(e), "A bold word")
  runs <- shape_runs(e)
  expect_equal(trimws(runs$text[runs$bold]), "bold")
})

test_that("a classDef text colour reaches w:color", {
  doc <- fixture_doc("base-segoe-subgraph")
  for (id in c("S", "E")) {
    expect_true(all(shape_runs(shape_named(doc, paste0("mermaid:node:", id)))$color == "FFFFFF"),
                info = id)
  }
  expect_true(all(shape_runs(shape_named(doc, "mermaid:node:A"))$color == "333333"))
})

# -- Paragraph properties ------------------------------------------------------

test_that("every text paragraph has single spacing, zero indents and a jc", {
  for (nm in c("flowchart-paint", "base-segoe-subgraph",
               "base-segoe-stadium-bold-invisible", "default-theme")) {
    paras <- xfind(fixture_doc(nm), "//w:txbxContent/w:p")
    expect_gt(length(paras), 3L)
    for (p in as.list(paras)) expect_reset_ppr(p, info = nm)
  }
})

test_that("edge-label and subgraph-title paragraphs are included", {
  doc <- fixture_doc("base-segoe-stadium-bold-invisible")
  p <- shape_paragraphs(shape_named(doc, "mermaid:edge_label:B:C"))
  expect_length(p, 1L)
  expect_reset_ppr(p[[1]])

  doc2 <- fixture_doc("base-segoe-subgraph")
  bg <- shape_named(doc2, "mermaid:subgraph_bg:Calls")
  expect_equal(shape_paragraph_texts(bg), "Phone the authorities in order")
  expect_reset_ppr(shape_paragraphs(bg)[[1]])
})

test_that("free text paragraphs are aligned from the text anchor", {
  jc_for <- function(anchor) {
    doc <- one_text_doc("x", x = 10, y = 10, w = 40, h = 20, anchor = anchor)
    paragraph_jc(shape_paragraphs(the_text_box(doc))[[1]])
  }
  expect_equal(jc_for("start"), "left")
  expect_equal(jc_for("middle"), "center")
  expect_equal(jc_for("end"), "right")
})

# -- Sizes ---------------------------------------------------------------------

all_sz <- function(doc) xnum(xfind(doc, "//w:rPr/w:sz"), "w:val")
all_szcs <- function(doc) xnum(xfind(doc, "//w:rPr/w:szCs"), "w:val")

test_that("16 px text at natural size is 12 pt (24 half-points), in w:sz and w:szCs", {
  doc <- one_text_doc("Hello", x = 10, y = 10, w = 60, h = 20, size = 16)
  expect_equal(all_sz(doc), 24)
  expect_equal(all_szcs(doc), 24)

  doc <- fixture_doc("flowchart-paint")   # all 16 px, drawn at natural size
  expect_true(all(all_sz(doc) == 24))
  expect_true(all(all_szcs(doc) == 24))
})

test_that("a fractional proportional size rounds down", {
  # 14.5 px * 9525 / 6350 = 21.75 half-points; 15 px = 22.5
  expect_equal(all_sz(one_text_doc("x", x = 10, y = 10, w = 40, h = 20, size = 14.5)), 21)
  expect_equal(all_sz(one_text_doc("x", x = 10, y = 10, w = 40, h = 20, size = 15)), 22)
})

test_that("an exact size survives float noise", {
  # 0.3 in page over 69.12 px: 16 px is 9.999999999999998 half-points in
  # floating point, exactly 10 on paper.
  r <- generic_render(list(prim_text("x", x = 5, y = 2, w = 20, h = 6, size = 16)),
                      viewBox = c(0, 0, 69.12, 10))
  expect_warning(res <- build_diagram_xml(r, page_width_in = 0.3),
                 class = "mermaid2drawml_small_text")
  expect_equal(all_sz(dml_doc(res$xml)), 10)
})

test_that("font_scale applies to the size", {
  expect_equal(all_sz(one_text_doc("x", x = 10, y = 10, w = 40, h = 20, size = 16,
                                   xml_args = list(font_scale = 1.5))), 36)
  expect_equal(all_sz(one_text_doc("x", x = 10, y = 10, w = 40, h = 20, size = 16,
                                   xml_args = list(font_scale = 1.25))), 30)
})

test_that("the size is at least 2 half-points", {
  r <- generic_render(list(prim_text("x", x = 10, y = 10, w = 40, h = 20, size = 1)))
  expect_warning(res <- build_diagram_xml(r), class = "mermaid2drawml_small_text")
  expect_equal(all_sz(dml_doc(res$xml)), 2)
})

test_that("text shrinks with a diagram that is shrunk to fit the page", {
  # 2000 px on a 6.27 in page: 2866.644 EMU per px, so 16 px = 7.22 half-points.
  r <- generic_render(list(prim_rect(0, 0, 2000, 100),
                           prim_text("wide", x = 980, y = 40, w = 40, h = 20, size = 16)),
                      viewBox = c(0, 0, 2000, 100))
  res <- suppressWarnings(build_diagram_xml(r))
  expect_equal(all_sz(dml_doc(res$xml)), 7)
})

# -- The small-text warning ----------------------------------------------------

wide_render <- function(size = 16, width = 2000) {
  generic_render(list(prim_text("A label of some length", x = 10, y = 40, w = 160, h = 20,
                                size = size)),
                 viewBox = c(0, 0, width, 100))
}

test_that("main text below 7 pt warns, giving the size in points and the remedies", {
  expect_warning(build_diagram_xml(wide_render()), class = "mermaid2drawml_small_text")
  cnd <- tryCatch(build_diagram_xml(wide_render()),
                  mermaid2drawml_small_text = function(c) c)
  msg <- conditionMessage(cnd)
  expect_match(msg, "3.5", fixed = TRUE)                 # 7 half-points
  expect_match(msg, "pt|point", ignore.case = TRUE)
  expect_match(msg, "landscape", ignore.case = TRUE)
})

test_that("the small-text warning is raised once per diagram", {
  r <- generic_render(list(
    prim_text("first label", x = 10, y = 10, w = 100, h = 20),
    prim_text("second label", x = 10, y = 40, w = 100, h = 20),
    prim_text("third label", x = 10, y = 70, w = 100, h = 20)
  ), viewBox = c(0, 0, 2000, 100))
  seen <- rlang::env(n = 0L)
  withCallingHandlers(build_diagram_xml(r), mermaid2drawml_small_text = function(c) {
    seen$n <- seen$n + 1L
    rlang::cnd_muffle(c)
  })
  expect_equal(seen$n, 1L)
})

test_that("no small-text warning at 7 pt or above", {
  # 1000 px wide: 5733.288 EMU per px, 16 px = 14.4 half-points -> 14 (7 pt).
  expect_no_warning(build_diagram_xml(wide_render(width = 1000)),
                    class = "mermaid2drawml_small_text")
  expect_no_warning(build_diagram_xml(wide_render(width = 200)),
                    class = "mermaid2drawml_small_text")
})

test_that("the warning follows the main text size, the one carried by the most characters", {
  big_main <- generic_render(list(
    prim_text("a long label carried in sixteen pixel text", x = 5, y = 10, w = 190, h = 20,
              size = 16),
    prim_text("x", x = 5, y = 60, w = 10, h = 10, size = 6)       # 9 half-points
  ))
  expect_no_warning(build_diagram_xml(big_main), class = "mermaid2drawml_small_text")

  small_main <- generic_render(list(
    prim_text("a long label carried in eight pixel text", x = 5, y = 10, w = 190, h = 20,
              size = 8),                                             # 12 half-points
    prim_text("X", x = 5, y = 60, w = 10, h = 20, size = 16)
  ))
  expect_warning(build_diagram_xml(small_main), class = "mermaid2drawml_small_text")
})

# -- Body properties and text boxes --------------------------------------------

test_that("a free text box: no fill, no line, its own box, wrap none, zero insets", {
  doc <- one_text_doc("Free", x = 30, y = 40, w = 50, h = 20)
  box <- the_text_box(doc)
  expect_equal(xattr(xfirst(box, "./wps:cNvSpPr"), "txBox"), "1")
  expect_equal(shape_fill(box)$kind, "none")
  expect_equal(shape_ln(box)$kind, "none")
  expect_near(unname(shape_off(box)), c(30, 40) * 9525, tol = 2)
  expect_near(unname(shape_ext(box)), c(50, 20) * 9525, tol = 2)
  bp <- body_pr(box)
  expect_equal(bp[["wrap"]], "none")
  for (ins in c("lIns", "tIns", "rIns", "bIns")) expect_equal(bp[[ins]], "0", info = ins)
  expect_equal(bp[["anchor"]], "ctr")
})

test_that("centred merged text: centre anchor, centred paragraphs, zero insets", {
  doc <- generic_doc(list(prim_rect(20, 20, 100, 40),
                          prim_text("Hi", x = 50, y = 30, w = 40, h = 20, anchor = "start")))
  box <- the_text_box(doc)
  expect_equal(shape_paragraph_texts(box), "Hi")
  bp <- body_pr(box)
  expect_equal(bp[["anchor"]], "ctr")
  expect_equal(bp[["wrap"]], "none")
  for (ins in c("lIns", "tIns", "rIns", "bIns")) expect_equal(bp[[ins]], "0", info = ins)
  expect_equal(paragraph_jc(shape_paragraphs(box)[[1]]), "center")
})

test_that("off-centre merged text: insets are the gaps to the shape, alignment from the anchor", {
  # rect 0..200 x 0..100, text 10..70 x 5..25, at natural scale (9525 EMU/px)
  doc <- generic_doc(list(prim_rect(0, 0, 200, 100),
                          prim_text("Title", x = 10, y = 5, w = 60, h = 20, anchor = "start")))
  box <- the_text_box(doc)
  bp <- body_pr(box)
  ins <- as.numeric(bp[c("lIns", "tIns", "rIns", "bIns")])
  expect_near(ins, c(10, 5, 130, 75) * 9525, tol = 2)
  expect_equal(bp[["wrap"]], "none")
  expect_equal(paragraph_jc(shape_paragraphs(box)[[1]]), "left")
})

test_that("a gap that is negative is clamped to a zero inset", {
  # 50 of 60 px inside the rect: merged, but overhanging on the right.
  doc <- generic_doc(list(prim_rect(0, 0, 100, 50),
                          prim_text("most", x = 50, y = 10, w = 60, h = 20, anchor = "end")))
  bp <- body_pr(the_text_box(doc))
  expect_equal(as.numeric(bp[["rIns"]]), 0)
  expect_near(as.numeric(bp[["lIns"]]), 50 * 9525, tol = 2)
  expect_equal(paragraph_jc(shape_paragraphs(the_text_box(doc))[[1]]), "right")
})

test_that("flowchart-paint: the subgraph title sits at the top of its background", {
  bg <- shape_named(fixture_doc("flowchart-paint"), "mermaid:subgraph_bg:S1")
  expect_equal(shape_paragraph_texts(bg), "Group one")
  bp <- body_pr(bg)
  # cluster 189.203, 112 (232.273 x 280.672); title 268.262, 114 (74.156 x 19)
  expect_near(as.numeric(bp[["tIns"]]), 2 * 9525, tol = 2)
  expect_near(as.numeric(bp[["lIns"]]), (268.262 - 189.203) * 9525, tol = 2)
  expect_near(as.numeric(bp[["bIns"]]), (392.672 - 133) * 9525, tol = 2)
})

test_that("node labels centred in their shapes have zero insets (flowchart-paint)", {
  doc <- fixture_doc("flowchart-paint")
  # (C's label sits below the cylinder's top ellipse, so it is inset instead.)
  for (id in c("A", "B", "E", "F")) {
    bp <- body_pr(shape_named(doc, paste0("mermaid:node:", id)))
    expect_equal(bp[["anchor"]], "ctr", info = id)
    for (ins in c("lIns", "tIns", "rIns", "bIns")) expect_equal(bp[[ins]], "0", info = id)
  }
})

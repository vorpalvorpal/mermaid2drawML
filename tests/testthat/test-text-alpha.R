# Text transparency: a run the browser drew with a colour alpha (or under an
# element with opacity) keeps its plain w:color, and Word 2010's w14:textFill
# carries the transparency, which Word writes as (1 - alpha) * 100000.

# The runs of the one text paragraph set of a hand-built diagram.
alpha_runs <- function(runs) {
  doc <- generic_doc(list(prim_text(rows = list(runs), x = 20, y = 20, w = 120, h = 20)))
  rows <- run_rows(xfind(doc, "//w:txbxContent/w:p"))
  rows[nzchar(rows$text), ]
}

# Word's text alpha is a transparency, in thousandths of a percent.
word_transparency <- function(alpha) round((1 - alpha) * 100000)

test_that("a run with alpha 0.8 keeps its colour and gains a text fill of transparency 20000", {
  r <- alpha_runs(list(text_run("faded", color = "#336699", alpha = 0.8)))
  expect_equal(nrow(r), 1L)
  expect_equal(r$color, "336699")
  expect_true(r$text_fill)
  expect_equal(r$text_fill_color, "336699")
  expect_near(r$text_fill_alpha, word_transparency(0.8), tol = 1)
  expect_equal(r$text_fill_alpha, 20000)
})

test_that("an opaque run writes no text fill", {
  r <- alpha_runs(list(text_run("solid", color = "#336699", alpha = 1)))
  expect_false(r$text_fill)
  expect_equal(r$color, "336699")
})

test_that("the transparency follows the alpha for any value", {
  for (a in c(0.37, 0.5, 0.05)) {
    r <- alpha_runs(list(text_run("x", color = "#000000", alpha = a)))
    expect_near(r$text_fill_alpha, word_transparency(a), tol = 1, info = format(a))
  }
})

test_that("only the transparent run of a paragraph carries a text fill", {
  r <- alpha_runs(list(text_run("one ", color = "#111111"),
                       text_run("two", color = "#222222", alpha = 0.5, bold = TRUE)))
  expect_equal(r$text, c("one ", "two"))
  expect_equal(r$text_fill, c(FALSE, TRUE))
  expect_equal(r$text_fill_color[[2]], "222222")
  expect_true(r$bold[[2]])
})

test_that("the text fill is the last run property, in schema order, with w14 declared", {
  doc <- generic_doc(list(prim_text(
    rows = list(list(text_run("styled", bold = TRUE, italic = TRUE, underline = TRUE,
                              color = "#AA0000", alpha = 0.6))),
    x = 20, y = 20, w = 120, h = 20)))
  expect_true(dml_ns[["w14"]] %in% unname(xml2::xml_ns(doc)))
  rpr <- xfirst(doc, "//w:txbxContent/w:p/w:r[w:t]/w:rPr")
  kids <- purrr::map_chr(as.list(xml2::xml_children(rpr)), schema_name)
  expect_equal(utils::tail(kids, 1L), "w14:textFill")
  expect_true(all(c("w:color", "w:u", "w:sz") %in% kids))
  fill <- xfirst(rpr, "./w14:textFill/w14:solidFill/w14:srgbClr")
  expect_equal(xattr(fill, "w14:val"), "AA0000")
  expect_length(schema_order_problems(doc), 0L)
})

test_that("text inside a shape carries its transparency too", {
  r <- one_node_render("squareRect", prim_rect(20, 20, 100, 50))
  # The label's run, made 70 % opaque as an element opacity would make it.
  label <- which(purrr::map_chr(r$prims, "kind") == "text")
  r$prims[[label]]$rows[[1]]$runs[[1]]$alpha <- 0.7
  node <- shape_named(dml_doc(build_diagram_xml(r)$xml), "mermaid:node:X")
  runs <- shape_runs(node)
  expect_equal(runs$text, "Label")
  expect_near(runs$text_fill_alpha, word_transparency(0.7), tol = 1)
})

test_that("gantt: the axis ticks (alpha 0.8) are transparent and nothing else is", {
  faded <- unlist(purrr::map(fixture_prims("gantt", "text"), function(p) {
    runs <- unlist(purrr::map(p$rows, "runs"), recursive = FALSE)
    purrr::map_chr(purrr::keep(runs, function(r) r$alpha < 1), "t")
  }))
  expect_length(faded, 13L)
  rows <- run_rows(xfind(fixture_doc("gantt"), "//w:txbxContent/w:p"))
  rows <- rows[nzchar(rows$text), ]
  expect_setequal(rows$text[rows$text_fill], faded)
  expect_true(all(rows$text_fill_alpha[rows$text_fill] == word_transparency(0.8)))
  expect_true(all(rows$text_fill_color[rows$text_fill] == rows$color[rows$text_fill]))
  expect_false(any(rows$text_fill[!rows$text %in% faded]))
})

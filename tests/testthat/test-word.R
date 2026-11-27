# The public wrappers (spec 5), offline: render_mermaid() is mocked to return
# a bridge fixture. Page layout, orientation and sections are in test-page.R.

code <- "flowchart TD\n  A --> B"

# Mock render_mermaid() with a fixture and record the arguments it was given.
local_capturing_render <- function(fixture = "default-theme", env = parent.frame()) {
  seen <- rlang::env(calls = list())
  render <- fixture_render(fixture)
  testthat::local_mocked_bindings(
    render_mermaid = function(mermaid_code, timeout = 120L, config = NULL, ...) {
      seen$calls <- c(seen$calls, list(list(mermaid_code = mermaid_code, timeout = timeout,
                                            config = config)))
      render
    },
    .env = env
  )
  seen
}

body_doc <- function(rdocx) officer::docx_body_xml(rdocx)

# -- Signatures ----------------------------------------------------------------

test_that("the wrappers keep their signatures, with mermaid_config just before ...", {
  expect_equal(names(formals(mermaid_to_word)),
               c("mermaid_code", "start_id", "timeout", "orientation", "auto_aspect",
                 "page_margins_in", "mermaid_config", "..."))
  expect_equal(names(formals(body_add_mermaid)),
               c("doc", "mermaid_code", "start_id", "timeout", "orientation", "auto_aspect",
                 "page_margins_in", "mermaid_config", "..."))
  expect_null(formals(mermaid_to_word)$mermaid_config)
  expect_null(formals(body_add_mermaid)$mermaid_config)
})

# -- mermaid_to_word() ---------------------------------------------------------

test_that("mermaid_to_word() returns a mermaid_word_diagram with the documented fields", {
  local_mock_render("default-theme")
  d <- mermaid_to_word(code)
  expect_s3_class(d, "mermaid_word_diagram")
  expect_true(all(c("xml", "scene", "type", "next_id", "orientation", "margins") %in% names(d)))
  expect_false("svg_data" %in% names(d))
  expect_type(d$xml, "character")
  expect_no_error(dml_doc(d$xml))
  expect_s3_class(d$scene, "mermaid_scene")
  expect_equal(d$type, "flowchart-v2")
  expect_true(rlang::is_integerish(d$next_id))
  expect_equal(d$orientation, "portrait")
  expect_null(d$margins)
})

test_that("mermaid_to_word() passes start_id and returns the next free id", {
  local_mock_render("default-theme")
  d <- mermaid_to_word(code, start_id = 10L)
  ids <- all_ids(dml_doc(d$xml))
  expect_equal(ids[1], 10L)
  expect_equal(d$next_id, 10L + length(ids))
})

test_that("mermaid_to_word(code, start_id =, word_font_family =) works as other code calls it", {
  local_mock_render("default-theme")
  d <- mermaid_to_word(code, start_id = 3L, word_font_family = "Myriad Pro")
  doc <- dml_doc(d$xml)
  fonts <- xfind(doc, "//w:rPr/w:rFonts")
  expect_gt(length(fonts), 0L)
  expect_true(all(xattr(fonts, "w:ascii") == "Myriad Pro"))
  expect_equal(xattr(doc_pr(doc), "id"), "3")
})

test_that("mermaid_config is passed to render_mermaid(config = )", {
  seen <- local_capturing_render()
  mermaid_to_word(code, mermaid_config = list(theme = "forest"))
  expect_length(seen$calls, 1L)
  expect_equal(seen$calls[[1]]$config, list(theme = "forest"))
  expect_equal(seen$calls[[1]]$mermaid_code, code)

  mermaid_to_word(code)
  expect_null(seen$calls[[2]]$config)
})

test_that("timeout is passed to render_mermaid()", {
  seen <- local_capturing_render()
  mermaid_to_word(code, timeout = 7L)
  expect_equal(seen$calls[[1]]$timeout, 7L)
})

test_that("arguments in ... reach build_diagram_xml()", {
  local_mock_render("default-theme")
  natural <- mermaid_to_word(code)
  filled <- mermaid_to_word(code, max_scale = Inf)
  expect_gt(root_ext(dml_doc(filled$xml))[["cx"]], root_ext(dml_doc(natural$xml))[["cx"]])

  left <- mermaid_to_word(code, align = "left")
  expect_equal(xattr(xfirst(dml_doc(left$xml), "/w:p/w:pPr/w:jc"), "w:val"), "left")

  big <- mermaid_to_word(code, font_scale = 1.5)
  expect_true(all(xnum(xfind(dml_doc(big$xml), "//w:rPr/w:sz"), "w:val") == 36))
})

test_that("an unknown argument in ... is an error; a removed one only warns", {
  local_mock_render("default-theme")
  expect_error(mermaid_to_word(code, no_such_option = 1), "no_such_option")
  expect_warning(mermaid_to_word(code, default_fill = "FFFFFF"),
                 class = "mermaid2drawml_deprecated")
})

test_that("custom margins do not trigger a deprecation warning", {
  local_mock_render("default-theme")
  expect_no_warning(mermaid_to_word(code, page_margins_in = 0.5),
                    class = "mermaid2drawml_deprecated")
})

test_that("the scene in the result is the one the XML was built from", {
  local_mock_render("flowchart-paint")
  d <- mermaid_to_word(code)
  expect_identical(build_diagram_xml(d$scene)$xml, d$xml)
})

# -- format() and print() ------------------------------------------------------

test_that("format() names the type, shape count, id range and orientation", {
  local_mock_render("default-theme")
  d <- mermaid_to_word(code, start_id = 5L)
  txt <- paste(format(d), collapse = "\n")
  expect_type(format(d), "character")
  n_shapes <- sum(d$scene$items$kind != "group")
  expect_match(txt, "flowchart-v2", fixed = TRUE)
  expect_match(txt, paste0("\\b", n_shapes, "\\b"))
  expect_match(txt, paste0("\\b5\\s*(-|\u2013|to)+\\s*", d$next_id - 1L, "\\b"))
  expect_false(grepl("\\b1\\s*(-|\u2013)", txt))
  expect_match(txt, "portrait", fixed = TRUE)
})

test_that("format() shows the margins only when they were set", {
  local_mock_render("default-theme")
  plain <- paste(format(mermaid_to_word(code)), collapse = "\n")
  with_margins <- paste(format(mermaid_to_word(code, page_margins_in = 0.5)), collapse = "\n")
  expect_false(grepl("margin", plain, ignore.case = TRUE))
  expect_match(with_margins, "margin", ignore.case = TRUE)
  expect_match(with_margins, "0.5", fixed = TRUE)
})

test_that("print() shows the summary and returns the diagram invisibly", {
  local_mock_render("default-theme")
  d <- mermaid_to_word(code, start_id = 5L)
  expect_snapshot(print(d))
  vis <- NULL
  suppressMessages(utils::capture.output(vis <- withVisible(print(d))))
  expect_false(vis$visible)
  expect_identical(vis$value, d)
})

test_that("print() with custom margins", {
  local_mock_render("default-theme")
  expect_snapshot(print(mermaid_to_word(code, page_margins_in = c(top = 0.5, left = 0.75))))
})

# -- body_add_mermaid() --------------------------------------------------------

test_that("body_add_mermaid() returns an rdocx whose body holds the drawing", {
  local_mock_render("default-theme")
  doc <- body_add_mermaid(officer::read_docx(), code)
  expect_s3_class(doc, "rdocx")
  body <- body_doc(doc)
  expect_length(xfind(body, "//wp:inline/wp:docPr[@name='mermaid:diagram']"), 1L)
  expect_length(xfind(body, "//wp:inline//wpg:wgp"), 1L)
  expect_gt(length(xfind(body, "//wps:wsp")), 3L)
  expect_gt(length(xfind(body, "//wps:cNvPr[starts-with(@name, 'mermaid:node:')]")), 0L)
})

test_that("two body_add_mermaid() calls on one document share no ids", {
  local_mock_render("default-theme")
  doc <- officer::read_docx()
  doc <- body_add_mermaid(doc, code)
  doc <- body_add_mermaid(doc, code)
  ids <- all_ids(body_doc(doc))
  expect_gt(length(ids), 10L)
  expect_false(anyDuplicated(ids) > 0L)
  expect_length(xfind(body_doc(doc), "//wp:inline"), 2L)
})

test_that("body_add_mermaid() passes mermaid_config and ... through", {
  seen <- local_capturing_render()
  doc <- body_add_mermaid(officer::read_docx(), code, mermaid_config = list(theme = "dark"),
                          max_scale = Inf)
  expect_equal(seen$calls[[1]]$config, list(theme = "dark"))
  cx <- xnum(xfirst(body_doc(doc), "//wpg:wgp/wpg:grpSpPr/a:xfrm/a:ext"), "cx")
  natural <- root_ext(dml_doc(mermaid_to_word(code)$xml))[["cx"]]
  expect_gt(cx, natural)
})

test_that("a saved document with a diagram opens as a docx", {
  local_mock_render("flowchart-paint")
  path <- withr::local_tempfile(fileext = ".docx")
  doc <- body_add_mermaid(officer::read_docx(), code)
  expect_no_error(print(doc, target = path))
  expect_gt(file.size(path), 5000)
  expect_s3_class(officer::read_docx(path), "rdocx")
})

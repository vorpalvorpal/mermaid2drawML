# The only tests that start Node and a browser. Each begins with
# skip_if_no_mermaid(), so a skipped run renders nothing.

.live <- rlang::env()

live_flowchart_code <- c(
  "flowchart TD",
  "  A[Start] --> B[End]",
  "  style A fill:#ff0000"
)

# Rendered at most once per run, and only after the skip guard has passed.
live_flowchart <- function() {
  skip_if_no_mermaid()
  if (is.null(.live$flowchart)) {
    .live$flowchart <- render_mermaid(live_flowchart_code)
  }
  .live$flowchart
}

test_that("live: a small flowchart renders to a mermaid_render", {
  r <- live_flowchart()
  expect_s3_class(r, "mermaid_render")
  expect_equal(r$type, "flowchart-v2")
  expect_gt(length(r$prims), 0L)
  expect_length(r$viewBox, 4L)
})

test_that("live: node A's shape carries the fill from the style statement", {
  r <- live_flowchart()
  in_node_a <- function(p) {
    any(purrr::map_lgl(p$anc, ~ identical(.x$id, "flowchart-A-0")))
  }
  shapes <- purrr::keep(r$prims, ~ in_node_a(.x) && !is.null(.x$fill))
  expect_gt(length(shapes), 0L)
  fills <- toupper(purrr::map_chr(shapes, "fill"))
  expect_true("#FF0000" %in% fills)
})

test_that("live: a Mermaid syntax error raises mermaid2drawml_syntax_error", {
  skip_if_no_mermaid()
  expect_error(render_mermaid("flowchart TD\n  A -->"),
               class = "mermaid2drawml_syntax_error")
})

test_that("live: config = list(theme = 'dark') changes the default node fill", {
  skip_if_no_mermaid()
  r <- render_mermaid("flowchart TD\n  A --> B", config = list(theme = "dark"))
  fills <- toupper(purrr::compact(purrr::map(r$prims, "fill")))
  expect_gt(length(fills), 0L)
  expect_false("#ECECFF" %in% unlist(fills))
})

# The Word end-to-end tests reuse the single render above through a mock, so the
# package's real conversion and officer integration run without a second browser.

test_that("live: body_add_mermaid writes a wpg:wgp group into document.xml", {
  r <- live_flowchart()
  local_mock_render(r)
  doc <- officer::read_docx()
  doc <- body_add_mermaid(doc, live_flowchart_code)
  target <- withr::local_tempfile(fileext = ".docx")
  print(doc, target = target)
  expect_true(file.exists(target))

  unz <- withr::local_tempdir()
  utils::unzip(target, files = "word/document.xml", exdir = unz)
  xml <- xml2::read_xml(file.path(unz, "word", "document.xml"))
  groups <- xml2::xml_find_all(
    xml, "//*[local-name()='wgp' and namespace-uri()='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup']"
  )
  expect_gte(length(groups), 1L)
})

test_that("live: mermaid_to_word returns a mermaid_word_diagram with advanced ids", {
  r <- live_flowchart()
  local_mock_render(r)
  d <- mermaid_to_word(live_flowchart_code, start_id = 5L)
  expect_s3_class(d, "mermaid_word_diagram")
  expect_gt(d$next_id, 5L)
})

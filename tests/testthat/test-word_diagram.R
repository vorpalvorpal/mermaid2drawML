test_that("mermaid_to_word returns correct class", {
  d <- mermaid_to_word("flowchart TB\n A[Start]-->B[End]")
  expect_s3_class(d, "mermaid_word_diagram")
  expect_named(d, c("xml", "graph", "next_id"))
})

test_that("print method", {
  d <- mermaid_to_word("flowchart TB\n A-->B")
  expect_output(print(d), "Mermaid Word Diagram")
  expect_invisible(print(d))
})

test_that("body_add_mermaid returns rdocx", {
  doc    <- officer::read_docx()
  result <- body_add_mermaid(doc, "flowchart TB\n A[Start]-->B[End]")
  expect_s3_class(result, "rdocx")
})

test_that("round-trip to docx", {
  tmp <- tempfile(fileext = ".docx")
  on.exit(unlink(tmp))
  doc <- officer::read_docx() |>
    body_add_mermaid("flowchart TB\n A[Start]-->B[End]")
  expect_no_error(print(doc, target = tmp))
  expect_true(file.exists(tmp))
  expect_gt(file.size(tmp), 5000L)
})

test_that("two diagrams no ID collision", {
  d1 <- mermaid_to_word("flowchart TB\n A-->B", start_id = 1L)
  d2 <- mermaid_to_word("flowchart TB\n C-->D", start_id = d1$next_id)
  get_ids <- function(xml) {
    m <- gregexpr('(?<=<wp:docPr id=")[0-9]+', xml, perl = TRUE)
    as.integer(regmatches(xml, m)[[1]])
  }
  expect_equal(length(intersect(get_ids(d1$xml), get_ids(d2$xml))), 0L)
})

test_that("EMP-SMART produces valid docx", {
  candidates <- c(
    file.path(dirname(dirname(testthat::test_path())), "examples", "EMP-SMART.mermaid"),
    file.path(testthat::test_path(), "..", "..", "examples", "EMP-SMART.mermaid"),
    "/Users/rjs/Documents/mermaid2drawML/examples/EMP-SMART.mermaid"
  )
  path <- Find(file.exists, candidates)
  skip_if(is.null(path), "EMP-SMART.mermaid not found")
  code <- paste(readLines(path, warn = FALSE), collapse = "\n")
  tmp  <- tempfile(fileext = ".docx")
  on.exit(unlink(tmp))
  doc  <- officer::read_docx() |> body_add_mermaid(code)
  expect_no_error(print(doc, target = tmp))
  expect_gt(file.size(tmp), 5000L)
})

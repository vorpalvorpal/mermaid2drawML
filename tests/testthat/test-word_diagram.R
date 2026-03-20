# tests/testthat/test-word_diagram.R
# Integration tests for mermaid_to_word() and body_add_mermaid().
# All tests that call Node.js are guarded with skip_if_not(nzchar(Sys.which("node"))).

node_available <- function() {
  # Check that node is on PATH and that setup_mermaid2drawml() has been run
  nzchar(Sys.which("node")) &&
    tryCatch({ check_mermaid2drawml(); TRUE }, error = function(e) FALSE)
}

# ── mermaid_to_word ────────────────────────────────────────────────────────

test_that("mermaid_to_word returns correct class", {
  skip_if_not(node_available(), "node not on PATH")
  d <- mermaid_to_word("flowchart TB\n  A[Start] --> B[End]")
  expect_s3_class(d, "mermaid_word_diagram")
  expect_named(d, c("xml", "svg_data", "ast", "next_id"))
})

test_that("mermaid_to_word xml is a non-empty character string", {
  skip_if_not(node_available(), "node not on PATH")
  d <- mermaid_to_word("flowchart TB\n  A --> B")
  expect_type(d$xml, "character")
  expect_gt(nchar(d$xml), 100L)
})

test_that("mermaid_to_word svg_data has nodes and edges tibbles", {
  skip_if_not(node_available(), "node not on PATH")
  d <- mermaid_to_word("flowchart TB\n  A[Start] --> B[End]")
  expect_true(is.data.frame(d$svg_data$nodes))
  expect_true(is.data.frame(d$svg_data$edges))
  expect_gte(nrow(d$svg_data$nodes), 2L)
  expect_gte(nrow(d$svg_data$edges), 1L)
})

test_that("mermaid_to_word next_id is positive integer", {
  skip_if_not(node_available(), "node not on PATH")
  d <- mermaid_to_word("flowchart TB\n  A --> B")
  expect_type(d$next_id, "integer")
  expect_gt(d$next_id, 1L)
})

# ── print method ───────────────────────────────────────────────────────────

test_that("print.mermaid_word_diagram shows summary and is invisible", {
  skip_if_not(node_available(), "node not on PATH")
  d <- mermaid_to_word("flowchart TB\n  A --> B")
  expect_output(print(d), "Mermaid Word Diagram")
  expect_output(print(d), "Nodes")
  expect_invisible(print(d))
})

# ── body_add_mermaid ───────────────────────────────────────────────────────

test_that("body_add_mermaid returns rdocx", {
  skip_if_not(node_available(), "node not on PATH")
  doc    <- officer::read_docx()
  result <- body_add_mermaid(doc, "flowchart TB\n  A[Start] --> B[End]")
  expect_s3_class(result, "rdocx")
})

test_that("round-trip to docx: file exists and is non-trivial", {
  skip_if_not(node_available(), "node not on PATH")
  tmp <- tempfile(fileext = ".docx")
  on.exit(unlink(tmp), add = TRUE)
  doc <- officer::read_docx() |>
    body_add_mermaid("flowchart TB\n  A[Start] --> B[End]")
  expect_no_error(print(doc, target = tmp))
  expect_true(file.exists(tmp))
  expect_gt(file.size(tmp), 5000L)
})

# ── No ID collisions across diagrams ──────────────────────────────────────

test_that("two diagrams in one doc have non-overlapping shape IDs", {
  skip_if_not(node_available(), "node not on PATH")
  d1 <- mermaid_to_word("flowchart TB\n  A --> B", start_id = 1L)
  d2 <- mermaid_to_word("flowchart TB\n  C --> D", start_id = d1$next_id)

  get_ids <- function(xml) {
    m <- gregexpr('(?<=<wp:docPr id=")[0-9]+', xml, perl = TRUE)
    as.integer(regmatches(xml, m)[[1]])
  }
  expect_equal(length(intersect(get_ids(d1$xml), get_ids(d2$xml))), 0L)
})

# ── Example file smoke test ────────────────────────────────────────────────

test_that("EMP-SMART.mermaid produces valid non-empty docx", {
  skip_if_not(node_available(), "node not on PATH")
  candidates <- c(
    file.path(dirname(dirname(testthat::test_path())),
              "examples", "EMP-SMART.mermaid"),
    "/Users/rjs/Documents/mermaid2drawML/examples/EMP-SMART.mermaid"
  )
  path <- Find(file.exists, candidates)
  skip_if(is.null(path), "EMP-SMART.mermaid not found")

  code <- paste(readLines(path, warn = FALSE), collapse = "\n")
  tmp  <- tempfile(fileext = ".docx")
  on.exit(unlink(tmp), add = TRUE)
  doc  <- officer::read_docx() |> body_add_mermaid(code)
  expect_no_error(print(doc, target = tmp))
  expect_gt(file.size(tmp), 5000L)
})

# ── Output XML passes well-formedness check ───────────────────────────────

test_that("generated XML is well-formed", {
  skip_if_not(node_available(), "node not on PATH")
  d <- mermaid_to_word("flowchart TB\n  A[Start] --> B{Decision}\n  B -->|Yes| C[Done]")
  expect_no_error(xml2::read_xml(d$xml))
})

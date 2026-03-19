test_that("returns mermaid_graph class", {
  g <- parse_mermaid("flowchart TB\n  A[Start] --> B[End]")
  expect_s3_class(g, "mermaid_graph")
  expect_named(g, c("direction", "nodes", "edges", "subgraphs", "classes"))
})

test_that("direction parsed", {
  expect_equal(parse_mermaid("flowchart TB\n A-->B")$direction, "TB")
  expect_equal(parse_mermaid("flowchart LR\n A-->B")$direction, "LR")
  expect_equal(parse_mermaid("graph BT\n A-->B")$direction,     "BT")
  expect_equal(parse_mermaid("flowchart TD\n A-->B")$direction, "TB")
})

test_that("nodes tibble columns", {
  g <- parse_mermaid("flowchart TB\n A[Hello]")
  expect_true(all(c("id","label","shape","fill","stroke","color","class") %in% names(g$nodes)))
})

test_that("EMP-SMART: 6 nodes 5 edges", {
  candidates <- c(
    file.path(dirname(dirname(testthat::test_path())), "examples", "EMP-SMART.mermaid"),
    file.path(testthat::test_path(), "..", "..", "examples", "EMP-SMART.mermaid"),
    "/Users/rjs/Documents/mermaid2drawML/examples/EMP-SMART.mermaid"
  )
  path <- Find(file.exists, candidates)
  skip_if(is.null(path), "EMP-SMART.mermaid not found")
  g <- parse_mermaid(paste(readLines(path, warn = FALSE), collapse = "\n"))
  expect_equal(g$direction, "TB")
  expect_equal(nrow(g$nodes), 6L)
  expect_equal(nrow(g$edges), 5L)
})

test_that("Infrastructure-CSM1: subgraphs and classDefs", {
  candidates <- c(
    file.path(dirname(dirname(testthat::test_path())), "examples", "Infrastructure-CSM1.mermaid"),
    file.path(testthat::test_path(), "..", "..", "examples", "Infrastructure-CSM1.mermaid"),
    "/Users/rjs/Documents/mermaid2drawML/examples/Infrastructure-CSM1.mermaid"
  )
  path <- Find(file.exists, candidates)
  skip_if(is.null(path), "Infrastructure-CSM1.mermaid not found")
  g <- parse_mermaid(paste(readLines(path, warn = FALSE), collapse = "\n"))
  expect_gte(nrow(g$subgraphs), 5L)
  expect_gte(nrow(g$classes),   5L)
})

test_that("node shape detection", {
  g <- parse_mermaid("flowchart TB\n A{{hex}}\n B{dia}\n C((circ))\n D[rect]\n E(rnd)")
  shapes <- setNames(g$nodes$shape, g$nodes$id)
  expect_equal(shapes[["A"]], "hexagon")
  expect_equal(shapes[["B"]], "diamond")
  expect_equal(shapes[["C"]], "ellipse")
  expect_equal(shapes[["D"]], "rect")
  expect_equal(shapes[["E"]], "roundRect")
})

test_that("arrow edge types", {
  g <- parse_mermaid("flowchart TB\n A --> B\n C --- D\n E -.-> F")
  expect_equal(nrow(g$edges), 3L)
  ab <- g$edges[g$edges$from == "A" & g$edges$to == "B", ]
  expect_equal(ab$arrow_end,  "arrow")
  cd <- g$edges[g$edges$from == "C" & g$edges$to == "D", ]
  expect_equal(cd$arrow_end,  "none")
  ef <- g$edges[g$edges$from == "E" & g$edges$to == "F", ]
  expect_equal(ef$line_type, "dashed")
})

test_that("edge labels: pipe syntax", {
  g <- parse_mermaid("flowchart TB\n A -->|yes| B\n C -->|no| D")
  expect_true("yes" %in% g$edges$label)
  expect_true("no"  %in% g$edges$label)
})

test_that("classDef colours applied to nodes", {
  code <- "flowchart TB\n  A[Node]:::myClass\n  classDef myClass fill:#FF0000,stroke:#000000"
  g <- parse_mermaid(code)
  a_node <- g$nodes[g$nodes$id == "A", ]
  expect_equal(a_node$fill[[1]], "FF0000")
})

test_that("style placeholder emits warning", {
  code <- "graph LR\n  A[x]:::@@{bg}@@"
  expect_warning(parse_mermaid(code), regexp = "placeholder|supported", ignore.case = TRUE)
})

test_that("empty input returns zero-row tibbles", {
  g <- parse_mermaid("")
  expect_equal(nrow(g$nodes), 0L)
  expect_equal(nrow(g$edges), 0L)
})

test_that("comments stripped", {
  g <- parse_mermaid("flowchart TB\n  %% this is a comment\n  A --> B")
  expect_equal(nrow(g$nodes), 2L)
})

test_that("print method is invisible and has output", {
  g <- parse_mermaid("flowchart TB\n A-->B")
  expect_output(print(g), "Mermaid graph")
  expect_invisible(print(g))
})

test_that("chained edges A->B->C produce 2 edge records", {
  g <- parse_mermaid("flowchart TB\n A --> B --> C")
  expect_equal(nrow(g$edges), 2L)
  expect_true(any(g$edges$from == "A" & g$edges$to == "B"))
  expect_true(any(g$edges$from == "B" & g$edges$to == "C"))
})

test_that("subgraphs have correct structure", {
  code <- "flowchart TB\n  subgraph Outer\n    A-->B\n  end"
  g <- parse_mermaid(code)
  expect_gte(nrow(g$subgraphs), 1L)
  expect_true("Outer" %in% g$subgraphs$id)
})

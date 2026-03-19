test_that("TB layout: y increases with depth", {
  g <- parse_mermaid("flowchart TB\n A[a]-->B[b]\n B-->C[c]") |> calculate_layout()
  ids <- g$nodes$id
  ya <- g$nodes$y[ids == "A"]; yb <- g$nodes$y[ids == "B"]; yc <- g$nodes$y[ids == "C"]
  expect_lt(ya, yb); expect_lt(yb, yc)
})

test_that("LR layout: x increases with depth", {
  g <- parse_mermaid("flowchart LR\n A-->B\n B-->C") |> calculate_layout()
  xa <- g$nodes$x[g$nodes$id=="A"]; xb <- g$nodes$x[g$nodes$id=="B"]; xc <- g$nodes$x[g$nodes$id=="C"]
  expect_lt(xa, xb); expect_lt(xb, xc)
})

test_that("all nodes get positive dimensions", {
  g <- parse_mermaid("flowchart TB\n A-->B\n B-->C") |> calculate_layout()
  expect_true(all(g$nodes$width  > 0))
  expect_true(all(g$nodes$height > 0))
  expect_true(all(!is.na(g$nodes$x)))
  expect_true(all(!is.na(g$nodes$y)))
})

test_that("edge coordinates are numeric", {
  g <- parse_mermaid("flowchart TB\n A-->B") |> calculate_layout()
  expect_type(g$edges$x1, "double")
  expect_type(g$edges$y1, "double")
})

test_that("single node no error", {
  g <- parse_mermaid("flowchart TB\n A[alone]") |> calculate_layout()
  expect_equal(nrow(g$nodes), 1L)
  expect_false(is.na(g$nodes$x))
})

test_that("empty graph no error", {
  g <- parse_mermaid("") |> calculate_layout()
  expect_equal(nrow(g$nodes), 0L)
})

test_that("disconnected nodes get grid layout", {
  g <- parse_mermaid("flowchart TB\n A[a]\n B[b]\n C[c]") |> calculate_layout()
  expect_equal(nrow(g$nodes), 3L)
  expect_true(all(!is.na(g$nodes$x)))
})

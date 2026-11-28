# Several diagrams in one bridge run: a list of Mermaid sources is rendered
# with one launch of Node and Chrome, which is most of the time a render
# takes. The bridge is mocked (local_mock_bridge() in helper-render.R).

code_pie <- "pie\n  \"a\" : 1"
code_flow <- "flowchart TD\n  A --> B"

test_that("a list of two sources runs node once with two input files", {
  st <- local_mock_bridge(bridge_ok(bridge_array(fixture_json("pie"),
                                                 fixture_json("default-theme"))))
  res <- render_mermaid(list(code_pie, code_flow))
  expect_equal(st$runs, 1L)
  expect_equal(st$inputs, c(code_pie, code_flow))
  expect_type(res, "list")
  expect_length(res, 2L)
  expect_s3_class(res[[1]], "mermaid_render")
  expect_s3_class(res[[2]], "mermaid_render")
  # In the order given.
  expect_equal(res[[1]]$type, fixture_render("pie")$type)
  expect_equal(res[[2]]$type, fixture_render("default-theme")$type)
})

test_that("each element of the list may be a vector of lines, joined into one diagram", {
  st <- local_mock_bridge(bridge_ok(bridge_array(fixture_json("pie"),
                                                 fixture_json("default-theme"))))
  render_mermaid(list(c("pie", "  \"a\" : 1"), c("flowchart TD", "  A --> B")))
  expect_equal(st$inputs, c(code_pie, code_flow))
})

test_that("a character vector of two lines is still one diagram", {
  st <- local_mock_bridge()
  res <- render_mermaid(c("pie", "  \"a\" : 1"))
  expect_equal(st$runs, 1L)
  expect_length(st$inputs, 1L)
  expect_equal(st$inputs, code_pie)
  expect_s3_class(res, "mermaid_render")
})

test_that("a failed diagram is a condition in its slot, and the others still render", {
  bad <- list(kind = "syntax", message = "Parse error on line 2: unexpected boom")
  local_mock_bridge(bridge_ok(bridge_array(fixture_json("pie"), failure_json(bad))))
  expect_no_error(res <- render_mermaid(list(code_pie, "flowchart TD\n A -->")))
  expect_length(res, 2L)
  expect_s3_class(res[[1]], "mermaid_render")
  expect_s3_class(res[[2]], "mermaid2drawml_syntax_error")
  expect_s3_class(res[[2]], "error")
  expect_match(conditionMessage(res[[2]]), "unexpected boom", fixed = TRUE)
})

test_that("a list of one source gives a list of one render", {
  st <- local_mock_bridge()
  res <- render_mermaid(list(code_pie))
  expect_equal(st$runs, 1L)
  expect_type(res, "list")
  expect_length(res, 1L)
  expect_s3_class(res[[1]], "mermaid_render")
})

test_that("mermaid_to_word() on a list threads the ids from one diagram to the next", {
  st <- local_mock_bridge(bridge_ok(bridge_array(fixture_json("pie"),
                                                 fixture_json("default-theme"))))
  ds <- quietly_expected(mermaid_to_word(list(code_pie, code_flow), start_id = 10))
  expect_equal(st$runs, 1L)
  expect_type(ds, "list")
  expect_length(ds, 2L)
  expect_s3_class(ds[[1]], "mermaid_word_diagram")
  expect_s3_class(ds[[2]], "mermaid_word_diagram")
  expect_equal(ds[[1]]$type, fixture_render("pie")$type)
  expect_equal(ds[[2]]$type, fixture_render("default-theme")$type)

  expect_equal(ds[[1]]$start_id, 10L)
  expect_equal(ds[[2]]$start_id, ds[[1]]$next_id)
  ids1 <- all_ids(dml_doc(ds[[1]]$xml))
  ids2 <- all_ids(dml_doc(ds[[2]]$xml))
  expect_equal(ids1[[1]], 10L)
  expect_equal(ids2[[1]], ds[[1]]$next_id)
  expect_equal(ds[[2]]$next_id, ds[[1]]$next_id + length(ids2))
  expect_false(anyDuplicated(c(ids1, ids2)) > 0L)
})

test_that("mermaid_to_word() on a list of one gives a list of one diagram", {
  local_mock_bridge()
  ds <- quietly_expected(mermaid_to_word(list(code_pie), start_id = 3))
  expect_type(ds, "list")
  expect_length(ds, 1L)
  expect_s3_class(ds[[1]], "mermaid_word_diagram")
  expect_equal(ds[[1]]$start_id, 3L)
})

# -- The timeout is per diagram ---------------------------------------------------

# `timeout` is the time one diagram gets. A list of diagrams shares one bridge
# run, so the process as a whole is allowed more time the more diagrams it has,
# while the bridge's own per-diagram limit (--timeout) is the same as for one.
bridge_timeout_arg <- function(st) {
  args <- st$call$args
  args[[match("--timeout", args) + 1L]]
}

test_that("the process timeout for one diagram, and a list of one, is the timeout", {
  st1 <- local_mock_bridge()
  render_mermaid(code_pie, timeout = 100)
  expect_equal(st1$call$timeout, 100)
  st2 <- local_mock_bridge()
  render_mermaid(list(code_pie), timeout = 100)
  expect_equal(st2$call$timeout, 100)
})

test_that("the process timeout for a list grows with the number of diagrams", {
  st1 <- local_mock_bridge()
  render_mermaid(code_pie, timeout = 100)
  one_arg <- bridge_timeout_arg(st1)
  st3 <- local_mock_bridge(bridge_ok(bridge_array(rep(fixture_json("pie"), 3L))))
  render_mermaid(list(code_pie, code_pie, code_pie), timeout = 100)
  expect_gte(st3$call$timeout, 3 * 100)
  expect_equal(bridge_timeout_arg(st3), one_arg)
})

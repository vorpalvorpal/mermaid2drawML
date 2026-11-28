# The package's warnings and errors carry what they report as data, so that
# calling code can act on them without parsing the message.

test_that("the unsupported warning carries the tags it left out", {
  r <- generic_render(list(prim_rect(0, 0, 50, 50), prim_image(60, 0, 30, 30),
                           prim_image(100, 0, 30, 30), prim_other("use", 140, 0, 20, 20)))
  cnd <- rlang::catch_cnd(build_diagram_xml(r), classes = "mermaid2drawml_unsupported")
  expect_s3_class(cnd, "mermaid2drawml_unsupported")
  expect_type(cnd$tags, "character")
  expect_setequal(cnd$tags, c("image", "use"))
})

test_that("the small-text warning carries the size in points", {
  # 16 px text on a 2000 px wide diagram fitted to the page: 7 half-points.
  r <- generic_render(list(prim_text("A label of some length", x = 10, y = 40, w = 160,
                                     h = 20, size = 16)),
                      viewBox = c(0, 0, 2000, 100))
  cnd <- rlang::catch_cnd(build_diagram_xml(r), classes = "mermaid2drawml_small_text")
  expect_s3_class(cnd, "mermaid2drawml_small_text")
  expect_type(cnd$points, "double")
  expect_equal(cnd$points, 3.5)
})

test_that("a syntax error carries Mermaid's own message", {
  msg <- "Parse error on line 2: unexpected boom"
  local_mock_bridge(bridge_ok(failure_stdout(list(kind = "syntax", message = msg))))
  cnd <- tryCatch(render_mermaid("flowchart TD\n A -->"),
                  mermaid2drawml_syntax_error = function(e) e)
  expect_s3_class(cnd, "mermaid2drawml_syntax_error")
  expect_identical(cnd$mermaid_message, msg)
})

test_that("a syntax error from an older bridge, with no kind, carries the message too", {
  local_mock_bridge(bridge_ok(failure_stdout(list(message = "Parse error on line 3"))))
  cnd <- tryCatch(render_mermaid("flowchart TD\n A -->"),
                  mermaid2drawml_syntax_error = function(e) e)
  expect_identical(cnd$mermaid_message, "Parse error on line 3")
})

test_that("a render that runs out of time carries the timeout in seconds", {
  local_mock_bridge(list(status = NA_integer_, stdout = "", stderr = "", timeout = TRUE))
  cnd <- tryCatch(render_mermaid("flowchart TD\n A --> B", timeout = 5L),
                  mermaid2drawml_timeout = function(e) e)
  expect_s3_class(cnd, "mermaid2drawml_timeout")
  expect_equal(cnd$timeout, 5)
})

test_that("a timeout the bridge reports carries the render's timeout in seconds", {
  # reading: the field holds the `timeout` the caller gave (seconds), not the
  # shorter in-page limit the bridge was handed.
  local_mock_bridge(bridge_ok(failure_stdout(list(kind = "timeout",
                                                  message = "page took too long"))))
  cnd <- tryCatch(render_mermaid("flowchart TD\n A --> B", timeout = 60),
                  mermaid2drawml_timeout = function(e) e)
  expect_s3_class(cnd, "mermaid2drawml_timeout")
  expect_equal(cnd$timeout, 60)
})

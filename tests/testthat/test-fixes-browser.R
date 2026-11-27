# F12: finding a Chrome for puppeteer.

test_that("F12: PUPPETEER_EXECUTABLE_PATH wins when it exists", {
  f <- withr::local_tempfile()
  writeLines("x", f)
  other <- withr::local_tempfile()
  writeLines("y", other)
  withr::local_envvar(PUPPETEER_EXECUTABLE_PATH = f)
  expect_equal(.find_browser(candidates = c(other)), f)
})

test_that("F12: a non-existent env path is ignored in favour of a candidate", {
  cand <- withr::local_tempfile()
  writeLines("y", cand)
  withr::local_envvar(PUPPETEER_EXECUTABLE_PATH = tempfile("nope"))
  expect_equal(.find_browser(candidates = c(tempfile("none"), cand)), cand)
})

test_that("F12: the first existing candidate is returned", {
  a <- withr::local_tempfile(); writeLines("a", a)
  b <- withr::local_tempfile(); writeLines("b", b)
  withr::local_envvar(PUPPETEER_EXECUTABLE_PATH = NA)
  expect_equal(.find_browser(candidates = c(tempfile("none"), a, b)), a)
})

test_that("F12: NA when nothing exists", {
  withr::local_envvar(PUPPETEER_EXECUTABLE_PATH = NA)
  res <- .find_browser(candidates = c(tempfile("none"), tempfile("none2")))
  expect_true(length(res) == 1L && is.na(res))
})

test_that("F12: the default candidates include the usual macOS Chrome path (when present)", {
  mac <- "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
  skip_if_not(file.exists(mac), "Google Chrome not installed in /Applications")
  withr::local_envvar(PUPPETEER_EXECUTABLE_PATH = NA)
  expect_equal(.find_browser(), mac)
})

mock_run_capture <- function() {
  st <- rlang::env(args = NULL)
  list(
    st = st,
    run = function(command, args = character(), ...) {
      st$args <- list(command = command, args = args, ...)
      list(status = 0L, stdout = '{"svg":"<svg/>","ast":null}', stderr = "")
    }
  )
}

test_that("F12: render_mermaid passes a found browser to node's environment", {
  cap <- mock_run_capture()
  testthat::local_mocked_bindings(
    check_mermaid2drawml = function() invisible(TRUE),
    .find_browser = function(...) "/fake/chrome"
  )
  testthat::local_mocked_bindings(run = cap$run, .package = "processx")
  render_mermaid("flowchart TB\n A --> B")
  env <- cap$st$args$env
  expect_false(is.null(env))
  expect_equal(unname(env["PUPPETEER_EXECUTABLE_PATH"]), "/fake/chrome")
})

test_that("F12: render_mermaid leaves the environment alone when no browser is found", {
  cap <- mock_run_capture()
  testthat::local_mocked_bindings(
    check_mermaid2drawml = function() invisible(TRUE),
    .find_browser = function(...) NA_character_
  )
  testthat::local_mocked_bindings(run = cap$run, .package = "processx")
  render_mermaid("flowchart TB\n A --> B")
  env <- cap$st$args$env
  expect_false("PUPPETEER_EXECUTABLE_PATH" %in% names(env))
})

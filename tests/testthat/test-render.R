# render_mermaid() and the bridge's result, offline: the external process
# is mocked (local_mock_bridge() in helper-render.R). Also .find_browser()
# and read_render().

# -- .find_browser() ------------------------------------------------------------

test_that("PUPPETEER_EXECUTABLE_PATH wins when it exists", {
  f <- withr::local_tempfile()
  writeLines("x", f)
  other <- withr::local_tempfile()
  writeLines("y", other)
  withr::local_envvar(PUPPETEER_EXECUTABLE_PATH = f)
  expect_equal(.find_browser(candidates = c(other)), f)
})

test_that("a non-existent env path is ignored in favour of a candidate", {
  cand <- withr::local_tempfile()
  writeLines("y", cand)
  withr::local_envvar(PUPPETEER_EXECUTABLE_PATH = tempfile("nope"))
  expect_equal(.find_browser(candidates = c(tempfile("none"), cand)), cand)
})

test_that("the first existing browser candidate is returned", {
  a <- withr::local_tempfile(); writeLines("a", a)
  b <- withr::local_tempfile(); writeLines("b", b)
  withr::local_envvar(PUPPETEER_EXECUTABLE_PATH = NA)
  expect_equal(.find_browser(candidates = c(tempfile("none"), a, b)), a)
})

test_that(".find_browser returns NA when nothing exists", {
  withr::local_envvar(PUPPETEER_EXECUTABLE_PATH = NA)
  res <- .find_browser(candidates = c(tempfile("none"), tempfile("none2")))
  expect_true(length(res) == 1L && is.na(res))
})

test_that("the default candidates include the usual macOS Chrome path", {
  mac <- "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
  skip_if_not(file.exists(mac), "Google Chrome not installed in /Applications")
  withr::local_envvar(PUPPETEER_EXECUTABLE_PATH = NA)
  expect_equal(.find_browser(), mac)
})

# -- read_render() and as_mermaid_render() --------------------------------------

test_that("read_render returns a validated mermaid_render", {
  r <- read_render(fixture_path("flowchart-paint"))
  expect_s3_class(r, "mermaid_render")
  expect_true(all(c("type", "role", "viewBox", "prims", "db", "title",
                    "desc", "mermaid") %in% names(r)))
  expect_type(r$type, "character")
  expect_type(r$role, "character")
  expect_type(r$viewBox, "double")
  expect_length(r$viewBox, 4L)
  expect_type(r$prims, "list")
  expect_gt(length(r$prims), 0L)
  expect_true(all(purrr::map_lgl(r$prims, is.list)))
  expect_type(r$mermaid, "character")
})

test_that("read_render keeps db as a list with nodes and edges when present", {
  r <- fixture_render("flowchart-paint")
  expect_type(r$db, "list")
  expect_true(all(c("nodes", "edges") %in% names(r$db)))
})

test_that("fixture_render caches: the same object comes back", {
  expect_identical(fixture_render("pie"), fixture_render("pie"))
})

test_that("as_mermaid_render rejects a result without prims", {
  bad <- list(ok = TRUE, type = "x", role = "x", viewBox = c(0, 0, 1, 1))
  expect_error(as_mermaid_render(bad), class = "mermaid2drawml_render_error")
})

test_that("as_mermaid_render rejects a result without a viewBox", {
  bad <- list(ok = TRUE, type = "x", role = "x", prims = list())
  expect_error(as_mermaid_render(bad), class = "mermaid2drawml_render_error")
})

test_that("as_mermaid_render rejects a viewBox that is not four numbers", {
  bad <- list(ok = TRUE, type = "x", role = "x", prims = list(),
              viewBox = c(0, 0, 1))
  expect_error(as_mermaid_render(bad), class = "mermaid2drawml_render_error")
})

test_that("as_mermaid_render rejects something that is not a list", {
  expect_error(as_mermaid_render("nope"), class = "mermaid2drawml_render_error")
})

# -- render_mermaid() with a mocked processx::run -------------------------------

# local_mock_bridge() and fixture_stdout() are in helper-render.R.
test_that("render_mermaid returns a mermaid_render built from the bridge array", {
  local_mock_bridge()
  r <- render_mermaid("pie\n  \"a\" : 1")
  expect_s3_class(r, "mermaid_render")
  expect_equal(r$type, fixture_render("pie")$type)
  expect_length(r$viewBox, 4L)
})

test_that("render_mermaid runs node on the bridge inside the package", {
  st <- local_mock_bridge()
  render_mermaid("flowchart TD\n A --> B")
  expect_equal(st$call$command, "node")
  bridge <- st$call$args[[1]]
  expect_match(bridge, "mermaid_bridge\\.js$")
  expect_true(file.exists(bridge))
  node_dir <- normalizePath(st$node_dir, mustWork = FALSE)
  expect_false(startsWith(normalizePath(bridge), node_dir))
})

test_that("the bridge environment carries M2D_NODE_DIR", {
  st <- local_mock_bridge()
  render_mermaid("flowchart TD\n A --> B")
  env <- st$call$env
  expect_false(is.null(env))
  expect_equal(unname(env["M2D_NODE_DIR"]), st$node_dir)
})

test_that("a found browser is passed as PUPPETEER_EXECUTABLE_PATH", {
  st <- local_mock_bridge(browser = "/fake/chrome")
  render_mermaid("flowchart TD\n A --> B")
  expect_equal(unname(st$call$env["PUPPETEER_EXECUTABLE_PATH"]), "/fake/chrome")
})

test_that("no PUPPETEER_EXECUTABLE_PATH is set when no browser is found", {
  st <- local_mock_bridge(browser = NA_character_)
  render_mermaid("flowchart TD\n A --> B")
  expect_false("PUPPETEER_EXECUTABLE_PATH" %in% names(st$call$env))
  expect_true("M2D_NODE_DIR" %in% names(st$call$env))
})

test_that("the last argument is a file holding the Mermaid code during the call", {
  st <- local_mock_bridge()
  render_mermaid("flowchart TD\n A --> B")
  expect_true(st$input_existed)
  expect_equal(st$input, "flowchart TD\n A --> B")
})

test_that("a vector of lines is collapsed with newlines", {
  st <- local_mock_bridge()
  render_mermaid(c("flowchart TD", " A --> B", " B --> C"))
  expect_equal(st$input, "flowchart TD\n A --> B\n B --> C")
})

test_that("an empty config sends no --config", {
  for (cfg in list(NULL, list())) {
    st <- local_mock_bridge()
    render_mermaid("flowchart TD\n A --> B", config = cfg)
    expect_false("--config" %in% st$call$args, info = deparse(cfg))
  }
})

test_that("config is written as JSON and passed with --config", {
  st <- local_mock_bridge()
  render_mermaid("flowchart TD\n A --> B", config = list(theme = "forest"))
  expect_true("--config" %in% st$call$args)
  expect_equal(st$config, list(theme = "forest"))
  # the Mermaid file is still the last argument
  expect_true(st$input_existed)
  expect_equal(st$input, "flowchart TD\n A --> B")
})

# Run render_mermaid() with a mocked bridge, keeping the --config file's text.
config_text <- function(config) {
  st <- local_mock_bridge()
  seen <- rlang::env(text = NULL)
  testthat::local_mocked_bindings(
    run = function(command, args = character(), ...) {
      i <- match("--config", args)
      seen$text <- paste(readLines(args[[i + 1L]], warn = FALSE, encoding = "UTF-8"),
                         collapse = "")
      list(status = 0L, stdout = fixture_stdout("pie"), stderr = "", timeout = FALSE)
    },
    .package = "processx"
  )
  render_mermaid("flowchart TD\n A --> B", config = config)
  seen$text
}

test_that("config writes null members as null, numbers in full and scalars as scalars", {
  txt <- config_text(list(theme = "dark", flowchart = list(padding = NULL),
                          themeVariables = list(fontSize = 16.123456789)))
  expect_match(txt, "\"theme\":\\s*\"dark\"")
  expect_match(txt, "\"padding\":\\s*null")
  expect_match(txt, "16.123456789", fixed = TRUE)
  parsed <- jsonlite::fromJSON(txt, simplifyVector = FALSE)
  expect_identical(parsed$theme, "dark")
  expect_true("padding" %in% names(parsed$flowchart))
  expect_null(parsed$flowchart$padding)
  expect_equal(parsed$themeVariables$fontSize, 16.123456789, tolerance = 1e-12)
})

test_that("cleanup_tree and the timeout are passed to processx", {
  st <- local_mock_bridge()
  render_mermaid("flowchart TD\n A --> B", timeout = 33L)
  expect_true(st$call$cleanup_tree)
  expect_equal(st$call$timeout, 33L)
})

test_that("the temporary input file is removed after the call", {
  st <- local_mock_bridge()
  render_mermaid("flowchart TD\n A --> B")
  expect_false(file.exists(utils::tail(st$call$args, 1L)))
})

test_that("the bridge's output is read as UTF-8", {
  st <- local_mock_bridge()
  render_mermaid("flowchart TD\n A --> B")
  expect_equal(st$call$encoding, "UTF-8")
})

test_that("non-ASCII label text comes back intact", {
  label <- "Zürich — 東京"
  raw <- list(ok = TRUE, bridge = 2L, mermaid = "11.13.0", type = "flowchart-v2",
              role = "flowchart-v2", viewBox = list(0, 0, 200, 100),
              prims = list(c(list(i = 0L), prim_text(label, 10, 10, 80, 20))), db = NULL)
  out <- paste0("[", jsonlite::toJSON(raw, auto_unbox = TRUE, null = "null", digits = NA), "]")
  local_mock_bridge(list(status = 0L, stdout = out, stderr = "", timeout = FALSE))
  r <- render_mermaid("flowchart TD\n A --> B")
  expect_equal(r$prims[[1]]$rows[[1]]$runs[[1]]$t, label)
})

# -- Failures -------------------------------------------------------------------

test_that("a non-zero exit status is a render error that shows stderr", {
  local_mock_bridge(list(status = 1L, stdout = "",
                         stderr = "Cannot find package puppeteer", timeout = FALSE))
  expect_error(render_mermaid("flowchart TD\n A --> B"),
               class = "mermaid2drawml_render_error",
               regexp = "Cannot find package puppeteer")
})

mock_failure <- function(error, env = parent.frame()) {
  local_mock_bridge(list(status = 0L, stdout = failure_stdout(error), stderr = "",
                         timeout = FALSE), env = env)
}

test_that("a bridge failure of kind bridge is a render error, not a syntax error", {
  mock_failure(list(kind = "bridge", message = "Chrome crashed"))
  cnd <- tryCatch(render_mermaid("flowchart TD\n A --> B"), error = function(e) e)
  expect_s3_class(cnd, "mermaid2drawml_render_error")
  expect_false(inherits(cnd, "mermaid2drawml_syntax_error"))
})

test_that("a syntax failure, or one with no kind (an old bridge), is a syntax error", {
  mock_failure(list(kind = "syntax", message = "Parse error on line 2"))
  expect_error(render_mermaid("flowchart TD\n A -->"),
               class = "mermaid2drawml_syntax_error", regexp = "Parse error")
  mock_failure(list(message = "Parse error on line 2"))
  expect_error(render_mermaid("flowchart TD\n A -->"),
               class = "mermaid2drawml_syntax_error", regexp = "Parse error")
})

test_that("empty stdout is a render error", {
  local_mock_bridge(list(status = 0L, stdout = "", stderr = "", timeout = FALSE))
  expect_error(render_mermaid("flowchart TD\n A --> B"),
               class = "mermaid2drawml_render_error")
})

test_that("stdout that is not JSON is a render error", {
  local_mock_bridge(list(status = 0L, stdout = "Warning: not json at all",
                         stderr = "", timeout = FALSE))
  expect_error(render_mermaid("flowchart TD\n A --> B"),
               class = "mermaid2drawml_render_error")
})

# -- Timeouts -------------------------------------------------------------------

test_that("the bridge's in-page timeout is shorter than the whole-render timeout", {
  st <- local_mock_bridge()
  render_mermaid("flowchart TD\n A --> B", timeout = 120)
  ms <- timeout_arg(st)
  expect_false(is.na(ms))
  expect_gt(ms, 0)
  expect_lt(ms, 120000)
  expect_gte(ms, 5000)
})

test_that("a very short timeout still leaves the bridge at least five seconds", {
  st <- local_mock_bridge()
  render_mermaid("flowchart TD\n A --> B", timeout = 3)
  expect_gte(timeout_arg(st), 5000)
})

# `timeout` is the time one diagram gets. A list of diagrams shares one bridge
# run, so the process as a whole is allowed more time the more diagrams it has,
# while the bridge's own per-diagram limit (--timeout) is the same as for one.
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
  one_arg <- timeout_arg(st1)
  st3 <- local_mock_bridge(bridge_ok(bridge_array(rep(fixture_json("pie"), 3L))))
  render_mermaid(list(code_pie, code_pie, code_pie), timeout = 100)
  expect_gte(st3$call$timeout, 3 * 100)
  expect_equal(timeout_arg(st3), one_arg)
})

# -- A list of diagrams in one bridge run ---------------------------------------

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

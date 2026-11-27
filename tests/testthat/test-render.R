# Offline tests of R/render.R and R/setup.R: the external process is mocked.

# ---- .find_browser() (moved from test-fixes-browser.R) ---------------------

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

# ---- read_render() / as_mermaid_render() -----------------------------------

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

# ---- render_mermaid() with a mocked processx::run ---------------------------

fixture_stdout <- function(name) {
  paste0("[", paste(readLines(fixture_path(name), warn = FALSE), collapse = "\n"), "]")
}

# Install mocks for node/browser lookup and processx::run. Returns an
# environment that records the call, including the contents of the input file
# and of any --config file at the time of the call.
local_mock_bridge <- function(result = NULL, browser = NA_character_,
                              node_dir = withr::local_tempdir(.local_envir = env),
                              env = parent.frame()) {
  result <- result %||% list(status = 0L, stdout = fixture_stdout("pie"),
                             stderr = "", timeout = FALSE)
  st <- rlang::env(call = NULL, input = NULL, input_existed = NA,
                   config = NULL, node_dir = node_dir)
  testthat::local_mocked_bindings(
    check_mermaid2drawml = function() invisible(TRUE),
    .find_browser = function(...) browser,
    .node_dir = function() node_dir,
    .env = env
  )
  testthat::local_mocked_bindings(
    run = function(command, args = character(), ...) {
      st$call <- list(command = command, args = args, ...)
      last <- utils::tail(args, 1L)
      st$input_existed <- file.exists(last)
      if (st$input_existed) st$input <- paste(readLines(last, warn = FALSE), collapse = "\n")
      i <- match("--config", args)
      if (!is.na(i) && file.exists(args[i + 1L])) {
        st$config <- jsonlite::fromJSON(args[i + 1L], simplifyVector = FALSE)
      }
      result
    },
    .package = "processx",
    .env = env
  )
  st
}

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

test_that("no --config argument is passed when config is NULL", {
  st <- local_mock_bridge()
  render_mermaid("flowchart TD\n A --> B")
  expect_false("--config" %in% st$call$args)
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

# ---- render_mermaid(): failures ---------------------------------------------

test_that("a Mermaid syntax error becomes a mermaid2drawml_syntax_error", {
  out <- '[{"ok":false,"error":{"message":"Parse error on line 2: unexpected boom"}}]'
  local_mock_bridge(list(status = 0L, stdout = out, stderr = "", timeout = FALSE))
  expect_error(render_mermaid("flowchart TD\n A -->"),
               class = "mermaid2drawml_syntax_error",
               regexp = "unexpected boom")
})

test_that("a non-zero exit status is a render error that shows stderr", {
  local_mock_bridge(list(status = 1L, stdout = "",
                         stderr = "Cannot find package puppeteer", timeout = FALSE))
  expect_error(render_mermaid("flowchart TD\n A --> B"),
               class = "mermaid2drawml_render_error",
               regexp = "Cannot find package puppeteer")
})

test_that("a processx timeout is a mermaid2drawml_timeout", {
  local_mock_bridge(list(status = NA_integer_, stdout = "", stderr = "",
                         timeout = TRUE))
  expect_error(render_mermaid("flowchart TD\n A --> B", timeout = 5L),
               class = "mermaid2drawml_timeout")
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

# ---- check_mermaid2drawml() ---------------------------------------------------

pinned_versions <- function() {
  pj <- jsonlite::fromJSON(system.file("node", "package.json",
                                       package = "mermaid2drawml"))
  list(mermaid = pj$dependencies[["mermaid"]],
       cli = pj$dependencies[["@mermaid-js/mermaid-cli"]])
}

write_pkg_version <- function(dir, pkg, version) {
  d <- file.path(dir, "node_modules", pkg)
  dir.create(d, recursive = TRUE, showWarnings = FALSE)
  writeLines(sprintf('{"name": "%s", "version": "%s"}', pkg, version),
             file.path(d, "package.json"))
}

fake_node_dir <- function(mermaid = pinned_versions()$mermaid,
                          cli = pinned_versions()$cli,
                          env = parent.frame()) {
  dir <- withr::local_tempdir(.local_envir = env)
  if (!is.null(mermaid)) write_pkg_version(dir, "mermaid", mermaid)
  if (!is.null(cli)) write_pkg_version(dir, "@mermaid-js/mermaid-cli", cli)
  dir
}

# Mock .node_dir() and a counting check_node(); returns the counter.
local_mock_node <- function(dir, env = parent.frame()) {
  st <- rlang::env(dir = dir, calls = 0L)
  testthat::local_mocked_bindings(
    .node_dir = function() st$dir,
    check_node = function() {
      st$calls <- st$calls + 1L
      invisible("node")
    },
    .env = env
  )
  st
}

test_that("the pinned versions in inst/node/package.json are readable", {
  v <- pinned_versions()
  expect_match(v$mermaid, "^[0-9]+\\.[0-9]+\\.[0-9]+")
  expect_match(v$cli, "^[0-9]+\\.[0-9]+\\.[0-9]+")
})

test_that("check_mermaid2drawml fails with a setup error when packages are missing", {
  dir <- withr::local_tempdir()
  local_mock_node(dir)
  expect_error(check_mermaid2drawml(), class = "mermaid2drawml_setup",
               regexp = "setup_mermaid2drawml()", fixed = FALSE)
})

test_that("a different mermaid version is a setup error", {
  dir <- fake_node_dir(mermaid = "0.0.1")
  local_mock_node(dir)
  expect_error(check_mermaid2drawml(), class = "mermaid2drawml_setup",
               regexp = "setup_mermaid2drawml()")
})

test_that("a different mermaid-cli version is a setup error", {
  dir <- fake_node_dir(cli = "0.0.1")
  local_mock_node(dir)
  expect_error(check_mermaid2drawml(), class = "mermaid2drawml_setup",
               regexp = "setup_mermaid2drawml()")
})

test_that("only one of the two packages being installed is a setup error", {
  dir <- fake_node_dir(cli = NULL)
  local_mock_node(dir)
  expect_error(check_mermaid2drawml(), class = "mermaid2drawml_setup")
})

test_that("check_mermaid2drawml succeeds when both versions match", {
  dir <- fake_node_dir()
  local_mock_node(dir)
  expect_no_error(check_mermaid2drawml())
})

test_that("a successful check is cached for the session", {
  dir <- fake_node_dir()
  st <- local_mock_node(dir)
  check_mermaid2drawml()
  n <- st$calls
  expect_gte(n, 1L)
  check_mermaid2drawml()
  check_mermaid2drawml()
  expect_equal(st$calls, n)
})

test_that("the cache is keyed by Node directory: another directory is checked afresh", {
  dir1 <- fake_node_dir()
  st <- local_mock_node(dir1)
  check_mermaid2drawml()
  n <- st$calls

  dir2 <- fake_node_dir()
  st$dir <- dir2
  check_mermaid2drawml()
  expect_gt(st$calls, n)

  # and a broken second directory still fails even though the first passed
  dir3 <- withr::local_tempdir()
  st$dir <- dir3
  expect_error(check_mermaid2drawml(), class = "mermaid2drawml_setup")
})

test_that("a failed check is not cached", {
  dir <- fake_node_dir(mermaid = "0.0.1")
  st <- local_mock_node(dir)
  expect_error(check_mermaid2drawml(), class = "mermaid2drawml_setup")
  write_pkg_version(dir, "mermaid", pinned_versions()$mermaid)
  expect_no_error(check_mermaid2drawml())
})

# ---- setup_mermaid2drawml() ------------------------------------------------

local_mock_setup <- function(status = 0L, env = parent.frame()) {
  parent <- withr::local_tempdir(.local_envir = env)
  node_dir <- file.path(parent, "node")
  st <- rlang::env(call = NULL, node_dir = node_dir)
  testthat::local_mocked_bindings(
    check_node = function() invisible("node"),
    .node_dir = function() node_dir,
    .env = env
  )
  testthat::local_mocked_bindings(
    run = function(command, args = character(), ...) {
      st$call <- list(command = command, args = args, ...)
      list(status = status, stdout = "", stderr = "npm ERR! boom", timeout = FALSE)
    },
    .package = "processx",
    .env = env
  )
  st
}

quietly <- function(expr) suppressWarnings(suppressMessages(expr))

test_that("setup copies package.json to the Node directory but not the bridge", {
  st <- local_mock_setup()
  quietly(setup_mermaid2drawml())
  expect_true(file.exists(file.path(st$node_dir, "package.json")))
  expect_false(file.exists(file.path(st$node_dir, "mermaid_bridge.js")))
  expect_equal(
    readLines(file.path(st$node_dir, "package.json")),
    readLines(system.file("node", "package.json", package = "mermaid2drawml"))
  )
})

test_that("setup runs npm install on the Node directory", {
  skip_if(.Platform$OS.type == "windows", "npm is npm.cmd on Windows")
  st <- local_mock_setup()
  quietly(setup_mermaid2drawml())
  expect_equal(st$call$command, "npm")
  expect_true("install" %in% st$call$args)
  expect_true(st$node_dir %in% st$call$args)
  expect_false("--prefer-online" %in% st$call$args)
})

test_that("upgrade = TRUE adds --prefer-online", {
  st <- local_mock_setup()
  quietly(setup_mermaid2drawml(upgrade = TRUE))
  expect_true("--prefer-online" %in% st$call$args)
})

test_that("setup raises a setup error when npm exits non-zero", {
  local_mock_setup(status = 1L)
  expect_error(quietly(setup_mermaid2drawml()), class = "mermaid2drawml_setup")
})

# check_mermaid2drawml() and setup_mermaid2drawml(), offline: the Node
# directory is a temporary one and npm and node are mocked.

# -- check_mermaid2drawml() -----------------------------------------------------

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
  # Every other pinned package (the icon packs) at its pinned version, so
  # that only the two above decide the check.
  pins <- jsonlite::fromJSON(system.file("node", "package.json",
                                         package = "mermaid2drawml"))$dependencies
  others <- setdiff(names(pins), c("mermaid", "@mermaid-js/mermaid-cli"))
  for (pkg in others) write_pkg_version(dir, pkg, pins[[pkg]])
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

# -- setup_mermaid2drawml() -----------------------------------------------------

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

# -- Setup warnings carry a class -----------------------------------------------

test_that("an old Node.js is a warning of class mermaid2drawml_old_node", {
  skip_if(!nzchar(Sys.which("node")), "node is not on PATH")
  testthat::local_mocked_bindings(
    run = function(command, args = character(), ...) {
      list(status = 0L, stdout = "v16.20.0\n", stderr = "", timeout = FALSE)
    },
    .package = "processx"
  )
  expect_warning(check_node(), class = "mermaid2drawml_old_node")
})

test_that("an incomplete installation is a warning of class mermaid2drawml_incomplete_setup", {
  local_mock_setup()
  expect_warning(suppressMessages(setup_mermaid2drawml()),
                 class = "mermaid2drawml_incomplete_setup")
})

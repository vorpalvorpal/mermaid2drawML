# Test helpers for renders: the bridge fixtures and what is built from them
# (cached once per session), mocked rendering and a mocked bridge.

# -- Fixtures --------------------------------------------------------------------

.test_cache <- rlang::env()

# `make()`, evaluated once per session for each `key`.
cached <- function(key, make) {
  if (is.null(.test_cache[[key]])) .test_cache[[key]] <- make()
  .test_cache[[key]]
}

fixture_path <- function(name) {
  testthat::test_path("fixtures", paste0(name, ".json"))
}

all_fixture_names <- function() {
  sub("\\.json$", "", list.files(testthat::test_path("fixtures"), pattern = "\\.json$"))
}

fixture_render <- function(name) {
  cached(paste0("render:", name), function() read_render(fixture_path(name)))
}

# Evaluate `expr`, muffling only the warnings a real diagram is expected to
# raise: text below 7 pt on a large diagram, and unsupported prims (images).
quietly_expected <- function(expr) {
  withCallingHandlers(
    expr,
    mermaid2drawml_small_text = function(cnd) rlang::cnd_muffle(cnd),
    mermaid2drawml_unsupported = function(cnd) rlang::cnd_muffle(cnd)
  )
}

fixture_scene <- function(name) {
  cached(paste0("scene:", name), function() quietly_expected(build_scene(fixture_render(name))))
}

# build_diagram_xml() on a fixture's scene (list(xml, next_id, ...)), cached
# per arguments.
fixture_xml <- function(name, ...) {
  args <- list(...)
  cached(paste0("xml:", name, ":", rlang::hash(args)), function() {
    quietly_expected(rlang::exec(build_diagram_xml, fixture_scene(name), !!!args))
  })
}

fixture_doc <- function(name, ...) {
  args <- list(...)
  cached(paste0("doc:", name, ":", rlang::hash(args)), function() {
    dml_doc(rlang::exec(fixture_xml, name, !!!args)$xml)
  })
}

# The bridge prims of one kind from a fixture.
fixture_prims <- function(name, kind = NULL) {
  p <- fixture_render(name)$prims
  if (is.null(kind)) p else purrr::keep(p, function(x) identical(x$kind, kind))
}

# The text of a text prim's rows (runs concatenated), one string per row.
prim_row_texts <- function(p) {
  vapply(p$rows, function(row) paste(vapply(row$runs, function(r) r$t, ""), collapse = ""),
         character(1))
}

# Run `check(name)` for every fixture, each in a test of its own named
# "<what>: <fixture>", so that a failure names the fixture.
for_each_fixture <- function(what, check) {
  for (name in all_fixture_names()) {
    testthat::test_that(paste0(what, ": ", name), check(name))
  }
}

# -- Rendering -------------------------------------------------------------------

skip_if_no_mermaid <- function() {
  testthat::skip_on_cran()
  testthat::skip_if(!nzchar(Sys.which("node")), "node is not on PATH")
  ok <- tryCatch({
    check_mermaid2drawml()
    TRUE
  }, error = function(e) FALSE)
  testthat::skip_if(!ok, "mermaid2drawml Node packages are not installed")
}

# Make render_mermaid() return `render` (a fixture name or a mermaid_render).
local_mock_render <- function(render, env = parent.frame()) {
  if (is.character(render)) render <- fixture_render(render)
  testthat::local_mocked_bindings(
    render_mermaid = function(...) render,
    .env = env
  )
}

# Two small diagrams' sources, for the calls that render a list of them.
code_pie <- "pie\n  \"a\" : 1"
code_flow <- "flowchart TD\n  A --> B"

# -- A mocked bridge -----------------------------------------------------------

# The bridge's stdout for results given as JSON object texts: a JSON array.
bridge_array <- function(...) paste0("[", paste(c(...), collapse = ",\n"), "]")

# One fixture as the text of a bridge result object.
fixture_json <- function(name) read_text_file(fixture_path(name))

# The bridge's stdout for one fixture.
fixture_stdout <- function(name) bridge_array(fixture_json(name))

# A failed diagram as the text of a bridge result object.
failure_json <- function(error) {
  as.character(jsonlite::toJSON(list(ok = FALSE, error = error), auto_unbox = TRUE))
}

# The bridge's stdout for one failure.
failure_stdout <- function(error) bridge_array(failure_json(error))

# A successful processx::run() result with `stdout`.
bridge_ok <- function(stdout) list(status = 0L, stdout = stdout, stderr = "", timeout = FALSE)

read_text_file <- function(path) paste(readLines(path, warn = FALSE), collapse = "\n")

# Install mocks for node/browser lookup and processx::run. Returns an
# environment that records the call: `runs` (how many times node was run),
# `call` (the command, its arguments and the other arguments given to
# processx::run, such as `timeout`), `input` (the last argument's contents)
# and `input_existed`, `inputs` (the contents of every input file, in
# argument order), `config` (the parsed --config file), all read at the time
# of the call, and `node_dir`.
local_mock_bridge <- function(result = NULL, browser = NA_character_,
                              env = parent.frame()) {
  result <- result %||% bridge_ok(fixture_stdout("pie"))
  node_dir <- withr::local_tempdir(.local_envir = env)
  st <- rlang::env(runs = 0L, call = NULL, input = NULL, inputs = character(),
                   input_existed = NA, config = NULL, node_dir = node_dir)
  testthat::local_mocked_bindings(
    check_mermaid2drawml = function() invisible(TRUE),
    .find_browser = function(...) browser,
    .node_dir = function() node_dir,
    .env = env
  )
  testthat::local_mocked_bindings(
    run = function(command, args = character(), ...) {
      st$runs <- st$runs + 1L
      st$call <- list(command = command, args = args, ...)
      last <- utils::tail(args, 1L)
      st$input_existed <- file.exists(last)
      if (st$input_existed) st$input <- read_text_file(last)
      i <- match("--config", args)
      cfg <- if (!is.na(i)) args[[i + 1L]] else character()
      if (length(cfg) > 0L && file.exists(cfg)) {
        st$config <- jsonlite::fromJSON(cfg, simplifyVector = FALSE)
      }
      # Input files: the arguments after the bridge script that are files,
      # other than the config file.
      rest <- args[-1L]
      files <- rest[file.exists(rest) & !rest %in% cfg]
      st$inputs <- vapply(files, read_text_file, character(1), USE.NAMES = FALSE)
      result
    },
    .package = "processx",
    .env = env
  )
  st
}

# The bridge's own per-diagram limit, the value given with --timeout (ms), of
# a call recorded by local_mock_bridge().
timeout_arg <- function(st) {
  i <- match("--timeout", st$call$args)
  if (is.na(i)) NA_real_ else as.numeric(st$call$args[[i + 1L]])
}

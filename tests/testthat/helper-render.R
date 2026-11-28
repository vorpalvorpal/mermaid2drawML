# Shared test helpers: bridge fixtures and mocked rendering.

fixture_path <- function(name) {
  testthat::test_path("fixtures", paste0(name, ".json"))
}

# One parsed copy of each fixture per session.
.fixture_cache <- rlang::env()

fixture_render <- function(name) {
  if (is.null(.fixture_cache[[name]])) {
    .fixture_cache[[name]] <- read_render(fixture_path(name))
  }
  .fixture_cache[[name]]
}

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

# -- A mocked bridge -----------------------------------------------------------

# The bridge's stdout for results given as JSON object texts: a JSON array.
bridge_array <- function(...) paste0("[", paste(c(...), collapse = ",\n"), "]")

# One fixture as the text of a bridge result object.
fixture_json <- function(name) paste(readLines(fixture_path(name), warn = FALSE), collapse = "\n")

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
# `call`, `input` (the last argument's contents) and `input_existed`, `inputs`
# (the contents of every input file, in argument order), and `config` (the
# parsed --config file), all read at the time of the call.
local_mock_bridge <- function(result = NULL, browser = NA_character_,
                              node_dir = withr::local_tempdir(.local_envir = env),
                              env = parent.frame()) {
  result <- result %||% bridge_ok(fixture_stdout("pie"))
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

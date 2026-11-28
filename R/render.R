# R/render.R
# Runs inst/node/mermaid_bridge.js through processx and turns its JSON result
# into a validated `mermaid_render` object.

#' Find a Chrome-like browser for puppeteer
#'
#' Returns `PUPPETEER_EXECUTABLE_PATH` if it is set and exists; otherwise the
#' first existing candidate among the usual Chrome, Chromium and Edge locations
#' on macOS, Windows and Linux (`google-chrome`, `chromium`, `chromium-browser`
#' on `PATH`).
#'
#' @param candidates Character vector of paths to try, in order.
#' @return A path, or `NA_character_` when none exists.
#' @noRd
.find_browser <- function(candidates = .browser_candidates()) {
  env <- Sys.getenv("PUPPETEER_EXECUTABLE_PATH", unset = "")
  if (nzchar(env) && file.exists(env)) return(env)
  found <- Filter(function(p) !is.na(p) && nzchar(p) && file.exists(p), candidates)
  if (length(found) > 0L) found[[1]] else NA_character_
}

.browser_candidates <- function() {
  win_dirs <- Filter(nzchar, c(Sys.getenv("ProgramFiles"),
                               Sys.getenv("ProgramFiles(x86)"),
                               "C:/Program Files", "C:/Program Files (x86)"))
  win <- as.vector(outer(unique(win_dirs),
    c("Google/Chrome/Application/chrome.exe",
      "Microsoft/Edge/Application/msedge.exe"),
    file.path))
  c("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    win,
    unname(Sys.which(c("google-chrome", "chromium", "chromium-browser"))))
}

#' Render Mermaid diagrams in a headless browser
#'
#' Runs the package's Node bridge (`inst/node/mermaid_bridge.js`), which has
#' Mermaid draw the diagram in headless Chrome and reports what the browser
#' drew: positioned shapes with computed paint, laid-out text and Mermaid's
#' own node and edge data. The bridge is run from the installed package, so it
#' always matches the R code reading its output.
#'
#' Starting Node and Chrome is most of the time a render takes, so several
#' diagrams are best rendered in one call: give `mermaid_code` as a list.
#'
#' @param mermaid_code Mermaid source: a character vector, whose lines are
#'   joined with newlines into one diagram; or a list of such vectors, each
#'   rendered as its own diagram in one run of the bridge.
#' @param timeout Seconds the whole render may take, Node and Chrome start-up
#'   included; Chrome can be slow to start the first time. The bridge is
#'   given this less a start-up allowance (never under 5 seconds) as its own
#'   limit for each page.
#' @param config `NULL`, or a named list of Mermaid configuration (for example
#'   `list(theme = "forest")`), written as JSON and given to the bridge.
#' @param svg Logical. `TRUE` also returns the SVG text Mermaid drew, in the
#'   `svg` field, to look at when a conversion seems wrong. Default `FALSE`:
#'   the conversion never needs it, and it can be large.
#' @return For a character vector, an object of class `mermaid_render`: a
#'   list with `type`, `role`, `viewBox` (numeric, length 4), `title`,
#'   `desc`, `svg` (`NULL` unless `svg = TRUE`), `prims` (the drawn
#'   primitives), `markers` (the line-end markers the primitives use, by id),
#'   `db` (Mermaid's node and edge data, or `NULL`) and `mermaid` (the
#'   Mermaid version). For a list, a list of the same length holding one
#'   `mermaid_render` per diagram, in order; a diagram that failed holds the
#'   error condition it would have raised on its own instead.
#' @section Errors:
#' A diagram Mermaid cannot parse or draw is an error of class
#' `mermaid2drawml_syntax_error`, with Mermaid's own message in its
#' `mermaid_message` field; a render that runs out of time is
#' `mermaid2drawml_timeout`, with the `timeout` (seconds) in its field of
#' that name; any other failure of the bridge or the browser is
#' `mermaid2drawml_render_error`. With a list, a failure of the whole run
#' (Node or Chrome failing, the run timing out) is still raised.
#' @export
render_mermaid <- function(mermaid_code, timeout = 120L, config = NULL, svg = FALSE) {
  check_mermaid2drawml()
  several <- is.list(mermaid_code)
  if (several && length(mermaid_code) == 0L) return(list())
  codes <- if (several) mermaid_code else list(mermaid_code)
  inputs <- purrr::map_chr(codes, write_mermaid_file, frame = environment())

  config_args <- if (length(config) > 0L) {
    cfg <- withr::local_tempfile(fileext = ".json")
    # NULL members must stay (as null), and digits = NA keeps full precision.
    jsonlite::write_json(config, cfg, auto_unbox = TRUE, null = "null", digits = NA)
    c("--config", cfg)
  }

  # mermaid-cli fails with "Could not find Chrome" when puppeteer's own browser
  # was never downloaded, so point it at an installed one when there is one.
  browser <- .find_browser()
  result <- processx::run(
    "node",
    c(system.file("node", "mermaid_bridge.js", package = "mermaid2drawml"),
      "--timeout", bridge_timeout_ms(timeout), config_args, if (isTRUE(svg)) "--svg",
      inputs),
    env = c("current",
            M2D_NODE_DIR = .node_dir(),
            if (!is.na(browser)) c(PUPPETEER_EXECUTABLE_PATH = browser)),
    timeout = timeout,
    encoding = "UTF-8",
    cleanup_tree = TRUE,
    error_on_status = FALSE
  )

  results <- parse_bridge_output(result, timeout, length(inputs))
  if (!several) return(bridge_render(results[[1]], timeout))
  purrr::map(results, function(r) {
    tryCatch(bridge_render(r, timeout),
             mermaid2drawml_syntax_error = identity,
             mermaid2drawml_timeout = identity,
             mermaid2drawml_render_error = identity)
  })
}

# One diagram's source in a temporary file that lasts as long as `frame`.
write_mermaid_file <- function(code, frame) {
  mmd <- withr::local_tempfile(fileext = ".mmd", .local_envir = frame)
  writeLines(enc2utf8(paste(code, collapse = "\n")), mmd, useBytes = TRUE)
  mmd
}

# Time (s) Node and Chrome need to start before the page's own limit applies.
.bridge_startup_s <- 20

# The in-page limit given to the bridge (ms): the render's timeout less the
# start-up allowance, but long enough for a small diagram.
bridge_timeout_ms <- function(timeout) {
  max(5000, round((timeout - .bridge_startup_s) * 1000))
}

# The error class for each kind of failure the bridge reports. An older
# bridge reports no kind, and only ever reported Mermaid's own errors.
.failure_classes <- c(syntax = "mermaid2drawml_syntax_error",
                      timeout = "mermaid2drawml_timeout",
                      bridge = "mermaid2drawml_render_error")

failure_kind <- function(kind) {
  if (is.character(kind) && length(kind) == 1L && kind %in% names(.failure_classes)) kind else "syntax"
}

#' The results of one bridge run, one per input
#'
#' The bridge prints a JSON array with one result per input, in input order.
#' A failed diagram is a normal result (`ok: false`) with an error kind; a
#' failure of the run itself is raised here.
#' @param n The number of inputs.
#' @noRd
parse_bridge_output <- function(result, timeout, n = 1L) {
  if (isTRUE(result$timeout)) {
    cli::cli_abort("Rendering did not finish within {timeout} second{?s}.",
                   class = "mermaid2drawml_timeout", timeout = timeout)
  }
  if (!identical(as.integer(result$status), 0L)) {
    cli::cli_abort(
      c("The Mermaid bridge failed (exit status {result$status}).",
        i = if (nzchar(result$stderr)) "{result$stderr}" else "No error output.",
        i = "Check that {.run setup_mermaid2drawml()} has been run."),
      class = "mermaid2drawml_render_error"
    )
  }
  if (!nzchar(trimws(result$stdout))) {
    cli::cli_abort("The Mermaid bridge returned no output.",
                   class = "mermaid2drawml_render_error")
  }
  out <- tryCatch(
    jsonlite::fromJSON(result$stdout, simplifyVector = FALSE),
    error = function(e) {
      cli::cli_abort(
        c("The Mermaid bridge output is not valid JSON.", i = conditionMessage(e)),
        class = "mermaid2drawml_render_error"
      )
    }
  )
  # A single result object, not in an array, is taken as the array of it.
  results <- if (is.list(out) && is.null(names(out))) out else list(out)
  if (length(results) < n) {
    cli::cli_abort(
      "The Mermaid bridge returned {length(results)} result{?s} for {n} diagram{?s}.",
      class = "mermaid2drawml_render_error"
    )
  }
  results[seq_len(n)]
}

#' One bridge result as a `mermaid_render`, or the error its failure is
#' @noRd
bridge_render <- function(x, timeout, call = rlang::caller_env()) {
  if (is.list(x) && identical(x$ok, FALSE)) {
    kind <- failure_kind(x$error$kind)
    message <- as.character(x$error$message %||% "")
    if (kind == "timeout") {
      cli::cli_abort(c("Mermaid could not render the diagram in time.", x = "{message}"),
                     class = "mermaid2drawml_timeout", timeout = timeout,
                     mermaid_message = message, call = call)
    }
    cli::cli_abort(c("Mermaid could not render the diagram.", x = "{message}"),
                   class = .failure_classes[[kind]], mermaid_message = message, call = call)
  }
  as_mermaid_render(x)
}

#' Read one bridge result from a JSON file
#'
#' For fixtures written by `dev/make-fixtures.mjs` (without `svg`).
#' @param path Path to a JSON file holding one bridge result object.
#' @noRd
read_render <- function(path) {
  as_mermaid_render(jsonlite::fromJSON(path, simplifyVector = FALSE))
}

#' Validate a parsed bridge result and give it the `mermaid_render` class
#'
#' The one constructor behind both `render_mermaid()` and `read_render()`.
#' @param x A bridge result object parsed with `simplifyVector = FALSE`.
#' @noRd
as_mermaid_render <- function(x) {
  fail <- function(what) {
    cli::cli_abort(c("The Mermaid bridge result is malformed.", x = what),
                   class = "mermaid2drawml_render_error")
  }
  if (!is.list(x)) fail("It is not a list.")
  if (!is.list(x$prims)) fail("It has no {.field prims}.")
  view_box <- suppressWarnings(as.numeric(unlist(x$viewBox)))
  if (length(view_box) != 4L || anyNA(view_box)) {
    fail("{.field viewBox} is not four numbers.")
  }
  structure(
    list(type = x$type, role = x$role, viewBox = view_box,
         title = x$title, desc = x$desc, svg = x$svg,
         prims = x$prims, markers = x$markers, db = x$db,
         mermaid = x$mermaid),
    class = "mermaid_render"
  )
}

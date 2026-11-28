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

#' Render a Mermaid diagram in a headless browser
#'
#' Runs the package's Node bridge (`inst/node/mermaid_bridge.js`), which has
#' Mermaid draw the diagram in headless Chrome and reports what the browser
#' drew: positioned shapes with computed paint, laid-out text and Mermaid's
#' own node and edge data. The bridge is run from the installed package, so it
#' always matches the R code reading its output.
#'
#' @param mermaid_code Character. Mermaid source; a vector of lines is joined
#'   with newlines.
#' @param timeout Seconds to wait for the Node process. Chrome can be slow to
#'   start the first time.
#' @param config `NULL`, or a named list of Mermaid configuration (for example
#'   `list(theme = "forest")`), written as JSON and given to the bridge.
#' @return An object of class `mermaid_render`: a list with `type`, `role`,
#'   `viewBox` (numeric, length 4), `title`, `desc`, `svg`, `prims` (the drawn
#'   primitives), `markers` (the line-end markers the primitives use, by id),
#'   `db` (Mermaid's node and edge data, or `NULL`) and `mermaid` (the
#'   Mermaid version).
#' @export
render_mermaid <- function(mermaid_code, timeout = 120L, config = NULL) {
  check_mermaid2drawml()

  mmd <- withr::local_tempfile(fileext = ".mmd")
  writeLines(enc2utf8(paste(mermaid_code, collapse = "\n")), mmd, useBytes = TRUE)

  config_args <- if (!is.null(config)) {
    cfg <- withr::local_tempfile(fileext = ".json")
    jsonlite::write_json(config, cfg, auto_unbox = TRUE)
    c("--config", cfg)
  }

  # mermaid-cli fails with "Could not find Chrome" when puppeteer's own browser
  # was never downloaded, so point it at an installed one when there is one.
  browser <- .find_browser()
  result <- processx::run(
    "node",
    c(system.file("node", "mermaid_bridge.js", package = "mermaid2drawml"),
      config_args, mmd),
    env = c("current",
            M2D_NODE_DIR = .node_dir(),
            if (!is.na(browser)) c(PUPPETEER_EXECUTABLE_PATH = browser)),
    timeout = timeout,
    cleanup_tree = TRUE,
    error_on_status = FALSE
  )

  as_mermaid_render(parse_bridge_output(result, timeout))
}

# The bridge prints a JSON array with one result per input; we send one input.
# A Mermaid syntax error is a normal result (`ok: false`), not a failure.
parse_bridge_output <- function(result, timeout) {
  if (isTRUE(result$timeout)) {
    cli::cli_abort(
      "Rendering did not finish within {timeout} second{?s}.",
      class = "mermaid2drawml_timeout"
    )
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
  # A well-formed array holds exactly the one result we asked for.
  first <- if (is.list(out) && is.null(names(out)) && length(out) > 0L) out[[1]] else out
  if (is.list(first) && identical(first$ok, FALSE)) {
    cli::cli_abort(
      c("Mermaid could not render the diagram.", x = "{first$error$message}"),
      class = "mermaid2drawml_syntax_error"
    )
  }
  first
}

#' Read one bridge result from a JSON file
#'
#' For fixtures written by `dev/make-fixtures.mjs`, which omit `svg`.
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

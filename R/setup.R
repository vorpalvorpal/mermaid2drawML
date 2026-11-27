# R/setup.R
# Managing the Node.js packages the bridge needs.
#
# The package ships inst/node/package.json (pinned versions) and the bridge
# script. setup_mermaid2drawml() copies only package.json to a user-local
# directory and runs `npm install` there; the bridge itself is run from the
# installed package, so it cannot go stale.

# Passing checks are remembered per Node directory, so `node --version` is not
# spawned for every diagram. Failures are never stored.
.setup_cache <- rlang::env()

#' The user-local directory holding `node_modules`
#' @noRd
.node_dir <- function() {
  tools::R_user_dir("mermaid2drawml", "data")
}

#' The required Node packages: pinned version beside installed version
#'
#' The single source of truth for setup and check. Pins are read from the
#' package's own `inst/node/package.json`.
#' @return A tibble with `package`, `pinned`, `installed` (`NA` when absent)
#'   and `path` (the package's directory under `node_modules`).
#' @noRd
.node_packages <- function(node_dir = .node_dir()) {
  pj <- jsonlite::read_json(system.file("node", "package.json",
                                        package = "mermaid2drawml"))
  pinned <- unlist(pj$dependencies)
  path <- file.path(node_dir, "node_modules", names(pinned))
  tibble::tibble(
    package = names(pinned),
    pinned = unname(pinned),
    installed = purrr::map_chr(path, installed_version),
    path = path
  )
}

installed_version <- function(dir) {
  file <- file.path(dir, "package.json")
  if (!file.exists(file)) return(NA_character_)
  version <- jsonlite::read_json(file)$version
  if (is.character(version)) version else NA_character_
}

#' Set up Node.js dependencies for mermaid2drawml
#'
#' Installs the Node.js packages that render Mermaid diagrams. Run it once
#' after installing or updating the R package. Node.js (>= 18) and npm must be
#' on `PATH`.
#'
#' The packages, at the versions pinned in the package's `inst/node/package.json`,
#' are installed into a user-local directory
#' (`tools::R_user_dir("mermaid2drawml", "data")`):
#' - `@mermaid-js/mermaid-cli`, which brings Puppeteer and its Chromium
#'   (about 150 MB);
#' - `mermaid`, the diagram library itself.
#'
#' @param upgrade Logical. If `TRUE`, passes `--prefer-online` to npm to
#'   force fresh downloads. Default `FALSE`.
#' @return Invisibly, the path to the Node directory.
#' @export
setup_mermaid2drawml <- function(upgrade = FALSE) {
  check_node()

  node_dir <- .node_dir()
  dir.create(node_dir, recursive = TRUE, showWarnings = FALSE)
  file.copy(system.file("node", "package.json", package = "mermaid2drawml"),
            file.path(node_dir, "package.json"), overwrite = TRUE)

  cli::cli_inform(c(
    "Installing Node.js packages; the first run can take a few minutes.",
    i = "Installing to {.path {node_dir}}"
  ))
  result <- processx::run(
    if (.Platform$OS.type == "windows") "npm.cmd" else "npm",
    c("install", "--prefix", node_dir, if (upgrade) "--prefer-online"),
    echo = TRUE,
    error_on_status = FALSE
  )
  if (!identical(as.integer(result$status), 0L)) {
    cli::cli_abort(
      c("{.code npm install} failed.",
        i = if (nzchar(result$stderr)) "{result$stderr}",
        i = "Make sure npm is installed and on {.envvar PATH}."),
      class = "mermaid2drawml_setup"
    )
  }

  problems <- setup_problems(.node_packages(node_dir))
  if (length(problems) > 0L) {
    cli::cli_warn(c("The installation is incomplete.", problems))
  } else {
    cli::cli_inform(c(v = "mermaid2drawml is ready to use."))
  }
  invisible(node_dir)
}

#' Check that Node.js is available
#' @noRd
check_node <- function() {
  node <- Sys.which("node")
  if (!nzchar(node)) {
    cli::cli_abort(
      c("Node.js is not found on {.envvar PATH}.",
        i = "Install Node.js (>= 18) from {.url https://nodejs.org}.",
        i = "Then restart R and run {.run setup_mermaid2drawml()}."),
      class = "mermaid2drawml_setup"
    )
  }

  version <- tryCatch(
    processx::run("node", "--version", error_on_status = FALSE)$stdout,
    error = function(e) ""
  )
  major <- suppressWarnings(as.integer(stringr::str_extract(version, "[0-9]+")))
  if (!is.na(major) && major < 18L) {
    cli::cli_warn(c(
      "Node.js {trimws(version)} detected; version 18 or later is recommended.",
      i = "Some features may not work."
    ))
  }
  invisible(node)
}

# One named bullet per package that is missing or at the wrong version.
setup_problems <- function(packages) {
  bad <- packages[is.na(packages$installed) | packages$installed != packages$pinned, ]
  msgs <- ifelse(
    is.na(bad$installed),
    sprintf("%s is not installed (need %s).", bad$package, bad$pinned),
    sprintf("%s is %s (need %s).", bad$package, bad$installed, bad$pinned)
  )
  stats::setNames(msgs, rep("x", length(msgs)))
}

#' Check that mermaid2drawml's Node.js setup is complete
#'
#' Verifies that Node.js is available and that the installed Node packages are
#' the versions pinned by this version of mermaid2drawml. Rendering functions
#' call it for you; a passing check is remembered for the session.
#'
#' @return Invisibly, `TRUE`. Raises an error of class `mermaid2drawml_setup`
#'   when something is missing or out of date.
#' @export
check_mermaid2drawml <- function() {
  key <- normalizePath(.node_dir(), mustWork = FALSE)
  if (isTRUE(.setup_cache[[key]])) return(invisible(TRUE))

  check_node()
  problems <- setup_problems(.node_packages())
  if (length(problems) > 0L) {
    cli::cli_abort(
      c("The Node.js packages mermaid2drawml needs are not ready.",
        problems,
        i = "Run {.run setup_mermaid2drawml()} to install them."),
      class = "mermaid2drawml_setup"
    )
  }
  .setup_cache[[key]] <- TRUE
  invisible(TRUE)
}

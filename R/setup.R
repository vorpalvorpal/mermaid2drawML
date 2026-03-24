# R/setup.R
# Functions for managing the Node.js dependencies required by mermaid2drawml.
#
# The package ships inst/node/package.json and inst/node/mermaid_bridge.js.
# setup_mermaid2drawml() copies these to a user-writable cache directory and
# runs `npm install` there, downloading @mermaid-js/mermaid-cli (which bundles
# Chromium via Puppeteer) and @mermaid-js/parser. This only needs to be run
# once per machine (or after updating the package).

# ── Internal path helpers ──────────────────────────────────────────────────

#' Return the path to the user-local node directory
#' @keywords internal
.node_dir <- function() {
  tools::R_user_dir("mermaid2drawml", "data")
}

#' Return the path to mermaid_bridge.js in the node directory
#' @keywords internal
.bridge_js <- function() {
  file.path(.node_dir(), "mermaid_bridge.js")
}

#' Return the path to the locally installed mmdc binary
#' @keywords internal
.mmdc_bin <- function() {
  # npm installs .bin symlinks here
  file.path(.node_dir(), "node_modules", ".bin", "mmdc")
}

# ── Setup ──────────────────────────────────────────────────────────────────

#' Set up Node.js dependencies for mermaid2drawml
#'
#' Downloads and installs the Node.js packages required to convert Mermaid
#' diagrams to Word shapes. Must be run once after installing or updating the
#' R package. Requires Node.js (>= 18) and npm to be on `PATH`.
#'
#' The following npm packages are installed into a user-local cache directory
#' (`tools::R_user_dir("mermaid2drawml", "data")`):
#' - `@mermaid-js/mermaid-cli` — renders Mermaid diagrams to SVG (bundles
#'   Chromium via Puppeteer, ~150 MB download)
#' - `@mermaid-js/parser` — parses Mermaid syntax to a JSON AST
#'
#' @param upgrade Logical. If `TRUE`, passes `--prefer-online` to npm to
#'   force downloading fresh package versions. Default `FALSE`.
#' @return Invisibly returns the path to the node directory.
#' @export
setup_mermaid2drawml <- function(upgrade = FALSE) {
  check_node()

  node_dir <- .node_dir()
  if (!dir.exists(node_dir)) dir.create(node_dir, recursive = TRUE)

  # Copy package.json and bridge script from the installed package
  pkg_node <- system.file("node", package = "mermaid2drawml")
  if (!nzchar(pkg_node)) {
    # Development fallback: look relative to package root
    pkg_node <- file.path(dirname(dirname(sys.frame(1)$filename %||% ".")), "inst", "node")
  }

  file.copy(file.path(pkg_node, "package.json"),       file.path(node_dir, "package.json"),       overwrite = TRUE)
  file.copy(file.path(pkg_node, "mermaid_bridge.js"),  file.path(node_dir, "mermaid_bridge.js"),  overwrite = TRUE)

  message("Installing Node.js dependencies (this may take a few minutes on first run)...")
  message("  Installing to: ", node_dir)

  npm_args <- c("install", "--prefix", node_dir)
  if (upgrade) npm_args <- c(npm_args, "--prefer-online")

  result <- processx::run(
    "npm", npm_args,
    echo = TRUE,
    error_on_status = FALSE
  )

  if (result$status != 0L) {
    rlang::abort(c(
      "npm install failed.",
      i = result$stderr,
      i = "Make sure npm is installed and on PATH."
    ))
  }

  # Verify key package directories exist on disk.
  # (@mermaid-js/parser is ESM-only so require() would give a false failure;
  #  checking the file system is sufficient and more reliable.)
  message("Verifying installation...")
  parser_dir <- file.path(node_dir, "node_modules", "@mermaid-js", "parser")
  mmdc_dir   <- file.path(node_dir, "node_modules", "@mermaid-js", "mermaid-cli")
  bridge_ok  <- file.exists(file.path(node_dir, "mermaid_bridge.js"))

  if (!dir.exists(parser_dir) || !dir.exists(mmdc_dir) || !bridge_ok) {
    rlang::warn(c(
      "Dependency verification had issues.",
      i = if (!dir.exists(parser_dir)) paste("Missing:", parser_dir),
      i = if (!dir.exists(mmdc_dir))   paste("Missing:", mmdc_dir),
      i = if (!bridge_ok)              paste("Missing bridge script:", .bridge_js()),
      i = "Try running setup_mermaid2drawml() again."
    ))
  } else {
    message("mermaid2drawml is ready to use.")
  }

  invisible(node_dir)
}

# ── Checks ─────────────────────────────────────────────────────────────────

#' Check that Node.js is available
#' @keywords internal
check_node <- function() {
  node <- Sys.which("node")
  if (!nzchar(node)) {
    rlang::abort(c(
      "Node.js is not found on PATH.",
      i = "Install Node.js (>= 18) from https://nodejs.org and ensure it is on your PATH.",
      i = "After installing Node.js, restart R and run setup_mermaid2drawml()."
    ))
  }

  # Check version >= 18
  ver_raw <- tryCatch(
    processx::run("node", "--version", error_on_status = FALSE)$stdout,
    error = function(e) ""
  )
  ver_match <- regmatches(ver_raw, regexpr("[0-9]+", ver_raw))
  if (length(ver_match) > 0) {
    major <- as.integer(ver_match[1])
    if (!is.na(major) && major < 18L) {
      rlang::warn(c(
        paste0("Node.js ", trimws(ver_raw), " detected; version >= 18 is recommended."),
        i = "Some features may not work correctly."
      ))
    }
  }

  invisible(node)
}

#' Check that mermaid2drawml's Node.js setup is complete
#'
#' Verifies that `setup_mermaid2drawml()` has been run and the required
#' Node.js packages are available. Called automatically by rendering functions.
#'
#' @return Invisibly returns the path to `mermaid_bridge.js`.
#' @export
check_mermaid2drawml <- function() {
  check_node()

  bridge <- .bridge_js()
  if (!file.exists(bridge)) {
    rlang::abort(c(
      "mermaid2drawml Node.js dependencies are not installed.",
      i = "Run setup_mermaid2drawml() to install them."
    ))
  }

  parser_dir <- file.path(.node_dir(), "node_modules", "@mermaid-js", "parser")
  if (!dir.exists(parser_dir)) {
    rlang::abort(c(
      "Node.js packages are missing.",
      i = "Run setup_mermaid2drawml() to install them."
    ))
  }

  invisible(bridge)
}

# ── Null-coalescing helper (base R) ────────────────────────────────────────
`%||%` <- function(x, y) if (is.null(x)) y else x

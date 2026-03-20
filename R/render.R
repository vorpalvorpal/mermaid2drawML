# R/render.R
# Calls mermaid_bridge.js via processx to get the AST (JSON) and rendered SVG
# for a Mermaid diagram. Both are returned together so Node.js / Puppeteer
# only spin up once per diagram.

#' Render a Mermaid diagram via the Node.js bridge
#'
#' Calls `mermaid_bridge.js`, which runs two Node.js operations in one process:
#' `@mermaid-js/mermaid-cli` renders the diagram to SVG (via Puppeteer/Chromium),
#' and `@mermaid-js/parser` attempts to parse the syntax to a JSON AST.
#'
#' The SVG is the authoritative output and drives all shape generation. The AST
#' is collected opportunistically for future use (see `enrich_from_ast()` in
#' `parse_svg.R`), but is not currently consulted during conversion. For
#' flowcharts — the most common diagram type — `@mermaid-js/parser` v0.3.x does
#' not support the diagram type and returns a `_parseError` sentinel instead.
#'
#' @param mermaid_code Character string of Mermaid diagram syntax.
#' @param timeout Integer. Seconds to wait for the Node.js process. Puppeteer /
#'   Chromium can be slow on first call. Default `120`.
#' @return A named list with:
#'   - `ast`: parsed list from `@mermaid-js/parser`, or
#'     `list(_parseError = "<message>")` for unsupported diagram types
#'   - `svg`: character string of the rendered SVG (always present)
#' @export
render_mermaid <- function(mermaid_code, timeout = 120L) {
  check_mermaid2drawml()

  # Write diagram to a temp file (mermaid-cli expects a file path)
  tmp_mmd <- tempfile(fileext = ".mmd")
  on.exit(unlink(tmp_mmd), add = TRUE)
  writeLines(mermaid_code, tmp_mmd, useBytes = FALSE)

  bridge <- .bridge_js()

  result <- processx::run(
    command         = "node",
    args            = c(bridge, tmp_mmd),
    timeout         = timeout,
    error_on_status = FALSE
  )

  if (result$status != 0L) {
    rlang::abort(c(
      "mermaid_bridge.js failed.",
      i = if (nzchar(result$stderr)) result$stderr else "(no stderr output)",
      i = "Check that setup_mermaid2drawml() has been run successfully."
    ))
  }

  if (!nzchar(result$stdout)) {
    rlang::abort("mermaid_bridge.js returned empty output.")
  }

  out <- tryCatch(
    jsonlite::fromJSON(result$stdout, simplifyVector = FALSE),
    error = function(e) rlang::abort(c("Failed to parse mermaid_bridge.js output as JSON.", i = e$message))
  )

  if (is.null(out$svg) || !nzchar(out$svg)) {
    rlang::abort("mermaid_bridge.js returned no SVG content.")
  }

  out
}

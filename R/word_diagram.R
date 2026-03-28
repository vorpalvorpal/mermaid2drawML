#' Convert a Mermaid diagram to a Word-embeddable object
#'
#' Renders the diagram via `@mermaid-js/mermaid-cli` (for layout and SVG) and
#' `@mermaid-js/parser` (for the AST), then converts the result to DrawingML
#' XML suitable for insertion into a Word document.
#'
#' Requires Node.js and the package's Node.js dependencies. Run
#' [setup_mermaid2drawml()] once after installation.
#'
#' @param mermaid_code A character string of Mermaid flowchart syntax.
#' @param start_id Integer. Starting shape ID. Increment between diagrams in
#'   the same document to avoid ID collisions. Default `1L`.
#' @param timeout Integer. Seconds to allow for the Node.js render. Default 120.
#' @param orientation Page orientation for the output document: `"portrait"`,
#'   `"landscape"`, or `"auto"` (default). In auto mode the SVG aspect ratio is
#'   compared against `auto_aspect`; if width/height exceeds that threshold the
#'   diagram is treated as landscape.
#' @param auto_aspect Numeric. Aspect-ratio threshold used when
#'   `orientation = "auto"`. Diagrams wider than this ratio are landscape.
#'   Default `1.3`.
#' @param ... Additional arguments passed to [build_diagram_xml()], e.g.
#'   `page_width_in`, `margin_in`, `default_stroke`.
#' @return A `mermaid_word_diagram` object with fields:
#'   - `xml`: DrawingML XML string ready for [officer::body_add_xml()]
#'   - `svg_data`: parsed SVG geometry (nodes, edges, viewbox tibbles)
#'   - `ast`: raw output from `@mermaid-js/parser`; a `_parseError` field
#'     indicates an unsupported diagram type (e.g. flowchart). Currently unused
#'     during conversion — see `enrich_from_ast()` in `parse_svg.R`.
#'   - `next_id`: next available shape ID for chaining documents
#'   - `orientation`: resolved orientation (`"portrait"` or `"landscape"`)
#' @examples
#' \dontrun{
#' d <- mermaid_to_word("flowchart TB\n  A[Start] --> B[End]")
#' print(d)
#' }
#' @export
mermaid_to_word <- function(mermaid_code, start_id = 1L, timeout = 120L,
                             orientation = c("auto", "portrait", "landscape"),
                             auto_aspect = 1.3,
                             ...) {
  orientation <- match.arg(orientation)

  rendered <- render_mermaid(mermaid_code, timeout = timeout)
  svg_data <- parse_mermaid_svg(rendered$svg, rendered$ast, source = mermaid_code)

  # Resolve orientation from SVG aspect ratio when "auto"
  use_landscape <- if (orientation == "landscape") {
    TRUE
  } else if (orientation == "portrait") {
    FALSE
  } else {
    vb <- svg_data$viewbox
    isTRUE(vb[4] > 0 && (vb[3] / vb[4]) > auto_aspect)
  }
  resolved_orientation <- if (use_landscape) "landscape" else "portrait"

  # Choose page dimensions based on orientation (content area, inside margins)
  extra_args <- list(...)
  if (use_landscape) {
    if (is.null(extra_args$page_width_in))  extra_args$page_width_in  <- 9.0
    if (is.null(extra_args$page_height_in)) extra_args$page_height_in <- 6.0
  } else {
    if (is.null(extra_args$page_width_in))  extra_args$page_width_in  <- 6.0
    if (is.null(extra_args$page_height_in)) extra_args$page_height_in <- 8.0
  }

  result <- do.call(
    build_diagram_xml,
    c(list(svg_data  = svg_data,
           ast       = rendered$ast,
           start_id  = as.integer(start_id)),
      extra_args)
  )

  structure(
    list(xml         = result$xml,
         svg_data    = svg_data,
         ast         = rendered$ast,
         next_id     = result$next_id,
         orientation = resolved_orientation),
    class = "mermaid_word_diagram"
  )
}

#' Add a Mermaid diagram to a Word document
#'
#' A convenience wrapper around [mermaid_to_word()] that parses, renders, and
#' inserts a Mermaid diagram directly into an `rdocx` document object.
#' When the resolved orientation is landscape a matching section-break property
#' is appended so that Word displays the page in landscape.
#'
#' @param doc An `rdocx` object from [officer::read_docx()].
#' @param mermaid_code A character string of Mermaid flowchart syntax.
#' @param start_id Integer. Starting shape ID. Default `1L`.
#' @param timeout Integer. Seconds for Node.js render. Default 120.
#' @param orientation Page orientation: `"portrait"`, `"landscape"`, or
#'   `"auto"` (default). See [mermaid_to_word()] for details.
#' @param auto_aspect Numeric. Aspect-ratio threshold for auto orientation.
#'   Default `1.3`.
#' @param ... Additional arguments passed to [build_diagram_xml()].
#' @return The modified `rdocx` object (invisibly).
#' @examples
#' \dontrun{
#' library(officer)
#' doc <- read_docx() |>
#'   body_add_mermaid("flowchart LR\n  A[Start] --> B[End]")
#' print(doc, target = "output.docx")
#' }
#' @export
body_add_mermaid <- function(doc, mermaid_code, start_id = 1L,
                              timeout = 120L,
                              orientation = c("auto", "portrait", "landscape"),
                              auto_aspect = 1.3,
                              ...) {
  orientation <- match.arg(orientation)
  diagram <- mermaid_to_word(mermaid_code,
                              start_id    = as.integer(start_id),
                              timeout     = timeout,
                              orientation = orientation,
                              auto_aspect = auto_aspect,
                              ...)
  doc <- officer::body_add_xml(doc, diagram$xml)

  if (diagram$orientation == "landscape") {
    ps  <- officer::prop_section(
      page_size = officer::page_size(orient = "landscape"),
      type      = "continuous"
    )
    doc <- officer::body_end_block_section(doc, officer::block_section(ps))
  }

  invisible(doc)
}

#' @export
print.mermaid_word_diagram <- function(x, ...) {
  nd <- nrow(x$svg_data$nodes)
  ed <- nrow(x$svg_data$edges)
  cat("Mermaid Word Diagram\n")
  cat("  Nodes      :", nd, "\n")
  cat("  Edges      :", ed, "\n")
  cat("  Shape IDs  : 1 -", x$next_id - 1L, "\n")
  cat("  Orientation:", x$orientation, "\n")
  invisible(x)
}

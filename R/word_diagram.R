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
#' @param ... Additional arguments passed to [build_diagram_xml()], e.g.
#'   `page_width_in`, `margin_in`, `default_stroke`.
#' @return A `mermaid_word_diagram` object with fields:
#'   - `xml`: DrawingML XML string ready for [officer::body_add_xml()]
#'   - `svg_data`: parsed SVG geometry (nodes, edges, viewbox tibbles)
#'   - `ast`: raw output from `@mermaid-js/parser`; a `_parseError` field
#'     indicates an unsupported diagram type (e.g. flowchart). Currently unused
#'     during conversion — see `enrich_from_ast()` in `parse_svg.R`.
#'   - `next_id`: next available shape ID for chaining documents
#' @examples
#' \dontrun{
#' d <- mermaid_to_word("flowchart TB\n  A[Start] --> B[End]")
#' print(d)
#' }
#' @export
mermaid_to_word <- function(mermaid_code, start_id = 1L, timeout = 120L, ...) {
  rendered <- render_mermaid(mermaid_code, timeout = timeout)
  svg_data <- parse_mermaid_svg(rendered$svg, rendered$ast)
  result   <- build_diagram_xml(svg_data,
                                ast      = rendered$ast,
                                start_id = as.integer(start_id),
                                ...)
  structure(
    list(xml      = result$xml,
         svg_data = svg_data,
         ast      = rendered$ast,
         next_id  = result$next_id),
    class = "mermaid_word_diagram"
  )
}

#' Add a Mermaid diagram to a Word document
#'
#' A convenience wrapper around [mermaid_to_word()] that parses, renders, and
#' inserts a Mermaid diagram directly into an `rdocx` document object.
#'
#' @param doc An `rdocx` object from [officer::read_docx()].
#' @param mermaid_code A character string of Mermaid flowchart syntax.
#' @param start_id Integer. Starting shape ID. Default `1L`.
#' @param timeout Integer. Seconds for Node.js render. Default 120.
#' @param ... Additional arguments passed to [build_diagram_xml()].
#' @return The modified `rdocx` object (invisibly).
#' @examples
#' \dontrun{
#' library(officer)
#' doc <- read_docx() |>
#'   body_add_mermaid("flowchart TB\n  A[Start] --> B[End]")
#' print(doc, target = "output.docx")
#' }
#' @export
body_add_mermaid <- function(doc, mermaid_code, start_id = 1L,
                              timeout = 120L, ...) {
  diagram <- mermaid_to_word(mermaid_code,
                              start_id = as.integer(start_id),
                              timeout  = timeout, ...)
  officer::body_add_xml(doc, diagram$xml)
}

#' @export
print.mermaid_word_diagram <- function(x, ...) {
  nd <- nrow(x$svg_data$nodes)
  ed <- nrow(x$svg_data$edges)
  cat("Mermaid Word Diagram\n")
  cat("  Nodes    :", nd, "\n")
  cat("  Edges    :", ed, "\n")
  cat("  Shape IDs: 1 -", x$next_id - 1L, "\n")
  invisible(x)
}

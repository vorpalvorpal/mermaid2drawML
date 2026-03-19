#' Convert a Mermaid diagram to a Word-embeddable object
#'
#' Parses a Mermaid flowchart and converts it to an object containing DrawingML
#' XML suitable for insertion into a Word document via [body_add_mermaid()] or
#' [officer::body_add_xml()].
#'
#' @param mermaid_code A character string of Mermaid flowchart syntax.
#' @param start_id Integer starting shape ID. Increment between diagrams in the
#'   same document to avoid ID collisions. Default 1L.
#' @param ... Additional arguments passed to [calculate_layout()].
#' @return A `mermaid_word_diagram` object with fields `xml`, `graph`, `next_id`.
#' @examples
#' d <- mermaid_to_word("flowchart TB\n  A[Start] --> B[End]")
#' print(d)
#' @export
mermaid_to_word <- function(mermaid_code, start_id = 1L, ...) {
  graph  <- parse_mermaid(mermaid_code)
  graph  <- calculate_layout(graph, ...)
  result <- build_diagram_xml(graph, start_id = as.integer(start_id))
  structure(
    list(xml = result$xml, graph = graph, next_id = result$next_id),
    class = "mermaid_word_diagram"
  )
}

#' Add a Mermaid diagram to a Word document
#'
#' A convenience wrapper around [mermaid_to_word()] that inserts the diagram
#' directly into an `rdocx` document from the officer package.
#'
#' @param doc An `rdocx` object from [officer::read_docx()].
#' @param mermaid_code A character string of Mermaid flowchart syntax.
#' @param start_id Integer starting shape ID. Default 1L.
#' @param ... Additional arguments passed to [calculate_layout()].
#' @return The modified `rdocx` object.
#' @examples
#' \dontrun{
#' library(officer)
#' doc <- read_docx() |>
#'   body_add_mermaid("flowchart TB\n  A[Start] --> B[End]")
#' print(doc, target = "output.docx")
#' }
#' @export
body_add_mermaid <- function(doc, mermaid_code, start_id = 1L, ...) {
  diagram <- mermaid_to_word(mermaid_code, start_id = as.integer(start_id), ...)
  officer::body_add_xml(doc, diagram$xml)
}

#' @export
print.mermaid_word_diagram <- function(x, ...) {
  cat("Mermaid Word Diagram\n")
  cat("  Nodes    :", nrow(x$graph$nodes), "\n")
  cat("  Edges    :", nrow(x$graph$edges), "\n")
  cat("  Subgraphs:", nrow(x$graph$subgraphs), "\n")
  cat("  Direction:", x$graph$direction, "\n")
  cat("  Shape IDs: 1 -", x$next_id - 1L, "\n")
  invisible(x)
}

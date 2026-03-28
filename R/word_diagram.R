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
#' @param page_margins_in Page margins in inches. Controls how much of the
#'   physical page is available for the diagram. Can be:
#'   \itemize{
#'     \item `NULL` (default): use built-in defaults (1" margins, US Letter).
#'     \item A single number: the same margin on all four sides.
#'     \item A named numeric vector with any subset of `top`, `bottom`, `left`,
#'       `right`; unspecified sides default to `1`.
#'   }
#'   The content area is computed from a US Letter page (8.5" × 11" portrait,
#'   11" × 8.5" landscape) minus the supplied margins. For A4 or custom sizes
#'   pass `page_width_in` / `page_height_in` / `margin_in` directly via `...`.
#' @param ... Additional arguments passed to [build_diagram_xml()], e.g.
#'   `page_width_in`, `margin_in`, `default_stroke`. These take precedence over
#'   values computed from `page_margins_in`.
#' @return A `mermaid_word_diagram` object with fields:
#'   - `xml`: DrawingML XML string ready for [officer::body_add_xml()]
#'   - `svg_data`: parsed SVG geometry (nodes, edges, viewbox tibbles)
#'   - `ast`: raw output from `@mermaid-js/parser`; a `_parseError` field
#'     indicates an unsupported diagram type (e.g. flowchart). Currently unused
#'     during conversion — see `enrich_from_ast()` in `parse_svg.R`.
#'   - `next_id`: next available shape ID for chaining documents
#'   - `orientation`: resolved orientation (`"portrait"` or `"landscape"`)
#'   - `margins`: normalised margin list, or `NULL` if `page_margins_in` was
#'     not supplied
#' @examples
#' \dontrun{
#' d <- mermaid_to_word("flowchart TB\n  A[Start] --> B[End]")
#' print(d)
#' }
#' @export
mermaid_to_word <- function(mermaid_code, start_id = 1L, timeout = 120L,
                             orientation    = c("auto", "portrait", "landscape"),
                             auto_aspect    = 1.3,
                             page_margins_in = NULL,
                             ...) {
  orientation <- match.arg(orientation)
  margins     <- .normalise_margins(page_margins_in)

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

  # A4 physical dimensions in inches (210 mm × 297 mm)
  a4_short <- 8.27
  a4_long  <- 11.69

  # Build extra_args for build_diagram_xml, respecting explicit overrides in ...
  extra_args <- list(...)

  if (!is.null(margins)) {
    # Compute content area from A4 physical page size minus supplied margins
    phys_w <- if (use_landscape) a4_long  else a4_short
    phys_h <- if (use_landscape) a4_short else a4_long
    if (is.null(extra_args$page_width_in))
      extra_args$page_width_in  <- phys_w - margins$left - margins$right
    if (is.null(extra_args$page_height_in))
      extra_args$page_height_in <- phys_h - margins$top  - margins$bottom
    if (is.null(extra_args$margin_in))
      extra_args$margin_in      <- margins$left
  } else {
    # Built-in defaults: A4 with 1" margins each side
    if (use_landscape) {
      if (is.null(extra_args$page_width_in))  extra_args$page_width_in  <- a4_long  - 2  # 9.69"
      if (is.null(extra_args$page_height_in)) extra_args$page_height_in <- a4_short - 2  # 6.27"
    } else {
      if (is.null(extra_args$page_width_in))  extra_args$page_width_in  <- a4_short - 2  # 6.27"
      if (is.null(extra_args$page_height_in)) extra_args$page_height_in <- a4_long  - 2  # 9.69"
    }
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
         orientation = resolved_orientation,
         margins     = margins),
    class = "mermaid_word_diagram"
  )
}

#' Add a Mermaid diagram to a Word document
#'
#' A convenience wrapper around [mermaid_to_word()] that parses, renders, and
#' inserts a Mermaid diagram directly into an `rdocx` document object.
#'
#' When the resolved orientation is landscape, or when `page_margins_in` is
#' supplied, a matching section-break property is appended so that Word uses
#' the correct page layout.
#'
#' @param doc An `rdocx` object from [officer::read_docx()].
#' @param mermaid_code A character string of Mermaid flowchart syntax.
#' @param start_id Integer. Starting shape ID. Default `1L`.
#' @param timeout Integer. Seconds for Node.js render. Default 120.
#' @param orientation Page orientation: `"portrait"`, `"landscape"`, or
#'   `"auto"` (default). See [mermaid_to_word()] for details.
#' @param auto_aspect Numeric. Aspect-ratio threshold for auto orientation.
#'   Default `1.3`.
#' @param page_margins_in Page margins in inches. See [mermaid_to_word()] for
#'   full details. Default `NULL` (use built-in defaults).
#' @param ... Additional arguments passed to [build_diagram_xml()].
#' @return The modified `rdocx` object (invisibly).
#' @examples
#' \dontrun{
#' library(officer)
#' # Auto orientation, narrow margins
#' doc <- read_docx() |>
#'   body_add_mermaid("flowchart LR\n  A[Start] --> B[End]",
#'                    page_margins_in = 0.5)
#' print(doc, target = "output.docx")
#' }
#' @export
body_add_mermaid <- function(doc, mermaid_code, start_id = 1L,
                              timeout         = 120L,
                              orientation     = c("auto", "portrait", "landscape"),
                              auto_aspect     = 1.3,
                              page_margins_in = NULL,
                              ...) {
  orientation <- match.arg(orientation)
  diagram <- mermaid_to_word(mermaid_code,
                              start_id        = as.integer(start_id),
                              timeout         = timeout,
                              orientation     = orientation,
                              auto_aspect     = auto_aspect,
                              page_margins_in = page_margins_in,
                              ...)
  doc <- officer::body_add_xml(doc, diagram$xml)

  # Apply a section break whenever orientation or margins differ from document
  # defaults (landscape requires it; custom margins always require it)
  needs_section <- diagram$orientation == "landscape" || !is.null(diagram$margins)

  if (needs_section) {
    orient   <- diagram$orientation
    # A4 physical dimensions: 210 mm × 297 mm = 8.27" × 11.69"
    phys_w <- if (orient == "landscape") 11.69 else 8.27
    phys_h <- if (orient == "landscape") 8.27  else 11.69

    ps_args <- list(
      page_size = officer::page_size(
        orient = orient,
        width  = phys_w,
        height = phys_h
      ),
      type = "continuous"
    )

    if (!is.null(diagram$margins)) {
      m <- diagram$margins
      ps_args$page_margins <- officer::page_mar(
        top    = m$top,
        bottom = m$bottom,
        left   = m$left,
        right  = m$right
      )
    }

    ps  <- do.call(officer::prop_section, ps_args)
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
  if (!is.null(x$margins)) {
    m <- x$margins
    cat(sprintf("  Margins    : top=%.2f  bottom=%.2f  left=%.2f  right=%.2f\n",
                m$top, m$bottom, m$left, m$right))
  }
  invisible(x)
}

# ── Internal helpers ──────────────────────────────────────────────────────────

#' Normalise the page_margins_in argument to a named list of four sides
#' @keywords internal
.normalise_margins <- function(page_margins_in) {
  if (is.null(page_margins_in)) return(NULL)

  default_margin <- 1.0

  if (is.numeric(page_margins_in) && length(page_margins_in) == 1L &&
      is.null(names(page_margins_in))) {
    # Scalar: apply to all sides
    v <- page_margins_in
    return(list(top = v, bottom = v, left = v, right = v))
  }

  if (!is.numeric(page_margins_in)) {
    stop("`page_margins_in` must be a numeric scalar or a named numeric vector.",
         call. = FALSE)
  }

  m <- as.list(page_margins_in)
  valid <- c("top", "bottom", "left", "right")
  bad   <- setdiff(names(m), valid)
  if (length(bad) > 0L)
    stop("`page_margins_in` has unknown names: ",
         paste(bad, collapse = ", "),
         ". Valid names are: top, bottom, left, right.",
         call. = FALSE)

  # Fill in any missing sides with the default
  for (s in valid) {
    if (is.null(m[[s]])) m[[s]] <- default_margin
  }

  list(top    = m$top,
       bottom = m$bottom,
       left   = m$left,
       right  = m$right)
}

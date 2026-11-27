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
#'   the same document to avoid ID collisions (see [body_add_mermaid()], which
#'   picks a free one automatically). Default `1L`.
#' @param timeout Integer. Seconds to allow for the Node.js render. Default 120.
#' @param orientation Page orientation for the output document: `"portrait"`,
#'   `"landscape"`, or `"auto"` (default). Auto chooses landscape only when the
#'   diagram is wider than `auto_aspect` (width/height) **and** would not fit
#'   the portrait content width at its drawn size (its width at 96 px per inch,
#'   times `max_scale`). A small wide diagram therefore stays portrait.
#' @param auto_aspect Numeric. Aspect-ratio threshold used when
#'   `orientation = "auto"`. Diagrams wider than this ratio are landscape.
#'   Default `1.3`.
#' @param page_margins_in Page margins in inches. Controls how much of the
#'   physical page is available for the diagram. Can be:
#'   \itemize{
#'     \item `NULL` (default): use built-in defaults (1" margins, A4).
#'     \item A single number: the same margin on all four sides.
#'     \item A named numeric vector with any subset of `top`, `bottom`, `left`,
#'       `right`; unspecified sides default to `1`.
#'   }
#'   The content area is computed from an A4 page (8.27" × 11.69" portrait,
#'   11.69" × 8.27" landscape) minus the supplied margins. For other sizes
#'   pass `page_width_in` / `page_height_in` / `margin_in` directly via `...`.
#' @param ... Additional arguments passed to [build_diagram_xml()], e.g.
#'   `page_width_in`, `default_stroke`, `max_scale` (default `1`: small diagrams
#'   are not enlarged; `Inf` fills the page), `font_scale`, `word_font_family`,
#'   `align`. These take precedence over values computed from
#'   `page_margins_in`. The drawing is placed in line with the text. Text below
#'   7 pt signals a `mermaid2drawml_small_text` warning.
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
  # render_mermaid() accepts a vector of lines; the old SVG parser wants one string.
  mermaid_code <- paste(mermaid_code, collapse = "\n")

  rendered <- render_mermaid(mermaid_code, timeout = timeout)
  svg_data <- parse_mermaid_svg(rendered$svg, rendered$ast, source = mermaid_code)

  extra_args <- list(...)

  # The portrait content width is what "auto" compares the diagram against,
  # so it is needed before the orientation is known.
  resolved_orientation <- .resolve_orientation(
    orientation,
    width_px  = svg_data$viewbox[3],
    height_px = svg_data$viewbox[4],
    auto_aspect = auto_aspect,
    max_scale = extra_args$max_scale %||% 1,
    portrait_width_in = .content_size("portrait", margins)[["width"]]
  )

  size <- .content_size(resolved_orientation, margins)
  layout <- list(page_width_in = size[["width"]], page_height_in = size[["height"]])
  if (!is.null(margins)) layout$margin_in <- margins$left
  extra_args <- utils::modifyList(layout, extra_args)

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
#' supplied, the diagram gets its own section. Anything already in the document
#' is first closed in a section that keeps the document's existing layout, so
#' earlier pages are not reformatted; the diagram then follows, and a section
#' break carrying its page size and margins ends it. No leading break is added
#' when the diagram is the first thing in the document. A portrait diagram with
#' default margins adds no section properties.
#'
#' @param doc An `rdocx` object from [officer::read_docx()].
#' @param mermaid_code A character string of Mermaid flowchart syntax.
#' @param start_id Integer. Starting shape ID. Default `NULL`, meaning one more
#'   than the largest id already used by a `wp:docPr` or any `cNvPr` in the
#'   document body (1 for an empty document), so repeated calls on one document
#'   never produce duplicate ids. An explicit value is used as given.
#' @param timeout Integer. Seconds for Node.js render. Default 120.
#' @param orientation Page orientation: `"portrait"`, `"landscape"`, or
#'   `"auto"` (default). See [mermaid_to_word()] for details.
#' @param auto_aspect Numeric. Aspect-ratio threshold for auto orientation.
#'   Default `1.3`.
#' @param page_margins_in Page margins in inches. See [mermaid_to_word()] for
#'   full details. Default `NULL` (use built-in defaults).
#' @param ... Additional arguments passed to [build_diagram_xml()], e.g.
#'   `max_scale` (default `1`: a diagram is drawn at natural size at most;
#'   `Inf` fills the page), `font_scale`, `word_font_family`, `align`.
#'   The drawing is placed in line with the text, in its own paragraph.
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
body_add_mermaid <- function(doc, mermaid_code, start_id = NULL,
                              timeout         = 120L,
                              orientation     = c("auto", "portrait", "landscape"),
                              auto_aspect     = 1.3,
                              page_margins_in = NULL,
                              ...) {
  orientation <- match.arg(orientation)
  start_id <- start_id %||% .next_free_shape_id(doc)
  diagram <- mermaid_to_word(mermaid_code,
                              start_id        = as.integer(start_id),
                              timeout         = timeout,
                              orientation     = orientation,
                              auto_aspect     = auto_aspect,
                              page_margins_in = page_margins_in,
                              ...)
  needs_section <- diagram$orientation == "landscape" || !is.null(diagram$margins)

  # A section's properties sit at its end, so a diagram with its own layout
  # needs the content before it closed off first; otherwise Word would apply
  # the diagram's page size and margins to those earlier pages as well.
  if (needs_section && .has_body_content(doc)) {
    doc <- officer::body_end_section_continuous(doc)
  }

  doc <- officer::body_add_xml(doc, diagram$xml)

  if (needs_section) {
    doc <- officer::body_end_block_section(
      doc, officer::block_section(.diagram_section(diagram)))
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

#' One more than the largest shape id used in a document body
#'
#' Looks at `wp:docPr` and any `*:cNvPr` ids, so a diagram added next can use
#' ids that clash with nothing already there.
#' @param doc An `rdocx` object.
#' @return An integer; 1 when the body holds no ids.
#' @keywords internal
.next_free_shape_id <- function(doc) {
  ids <- xml2::xml_attr(
    xml2::xml_find_all(officer::docx_body_xml(doc),
                       "//*[local-name()='docPr' or local-name()='cNvPr']"),
    "id")
  ids <- suppressWarnings(as.integer(ids))
  ids <- ids[!is.na(ids)]
  if (length(ids) == 0L) 1L else max(ids) + 1L
}

#' Does the document body hold anything besides its final section properties?
#' @param doc An `rdocx` object.
#' @return A logical scalar.
#' @keywords internal
.has_body_content <- function(doc) {
  kids <- xml2::xml_find_all(
    officer::docx_body_xml(doc),
    "/*/*[local-name()='body']/*[not(local-name()='sectPr')]")
  length(kids) > 0L
}

#' Section properties for the page a diagram sits on
#' @param diagram A `mermaid_word_diagram`.
#' @return An officer `prop_section`.
#' @keywords internal
.diagram_section <- function(diagram) {
  dims <- .page_dims(diagram$orientation)
  args <- list(
    page_size = officer::page_size(orient = diagram$orientation,
                                   width = dims[["width"]],
                                   height = dims[["height"]]),
    type = "continuous")
  m <- diagram$margins
  if (!is.null(m)) {
    args$page_margins <- officer::page_mar(
      top = m$top, bottom = m$bottom, left = m$left, right = m$right)
  }
  do.call(officer::prop_section, args)
}

#' A4 page size in inches
#'
#' The only place the page size is defined.
#' @param orientation `"portrait"` or `"landscape"`.
#' @return Named numeric `c(width, height)`.
#' @keywords internal
.page_dims <- function(orientation) {
  short <- 8.27
  long  <- 11.69
  if (identical(orientation, "landscape")) {
    c(width = long, height = short)
  } else {
    c(width = short, height = long)
  }
}

#' Content area of a page: its size minus the margins
#' @param orientation `"portrait"` or `"landscape"`.
#' @param margins Output of `.normalise_margins()`; `NULL` means 1 inch on
#'   every side.
#' @return Named numeric `c(width, height)` in inches.
#' @keywords internal
.content_size <- function(orientation, margins) {
  m <- margins %||% list(top = 1, bottom = 1, left = 1, right = 1)
  dims <- .page_dims(orientation)
  size <- c(width  = dims[["width"]]  - m$left - m$right,
            height = dims[["height"]] - m$top  - m$bottom)
  if (any(size <= 0)) {
    cli::cli_abort(
      c("The page margins leave no room for content on a {orientation} page.",
        "i" = "Content area would be {round(size[['width']], 2)} x {round(size[['height']], 2)} inches."),
      class = "mermaid2drawml_bad_margins")
  }
  size
}

#' Choose portrait or landscape for a diagram
#'
#' `"auto"` picks landscape only when the diagram is both wide (aspect above
#' `auto_aspect`) and would not fit the portrait content width at its drawn
#' scale; a small wide diagram gains nothing from a landscape page.
#' @param orientation `"portrait"`, `"landscape"` or `"auto"`.
#' @param width_px,height_px Diagram size in CSS pixels (96 per inch).
#' @param auto_aspect Width/height ratio above which a diagram counts as wide.
#' @param max_scale The largest enlargement the diagram is drawn at.
#' @param portrait_width_in Content width of a portrait page, in inches.
#' @return `"portrait"` or `"landscape"`.
#' @keywords internal
.resolve_orientation <- function(orientation, width_px, height_px,
                                 auto_aspect = 1.3, max_scale = 1,
                                 portrait_width_in) {
  if (orientation != "auto") return(orientation)
  if (is.na(height_px) || height_px <= 0) return("portrait")
  wide <- width_px / height_px > auto_aspect
  too_wide_for_portrait <- width_px / 96 * max_scale > portrait_width_in
  if (isTRUE(wide && too_wide_for_portrait)) "landscape" else "portrait"
}

#' Normalise the page_margins_in argument to a named list of four sides
#' @param page_margins_in `NULL`, a single number, or a named numeric vector
#'   or list using any of `top`, `bottom`, `left`, `right`.
#' @return `NULL`, or a list with `top`, `bottom`, `left`, `right` in that
#'   order; missing sides are 1.
#' @keywords internal
.normalise_margins <- function(page_margins_in) {
  if (is.null(page_margins_in)) return(NULL)

  sides <- c("top", "bottom", "left", "right")
  bad <- function(msg) {
    cli::cli_abort(msg, class = "mermaid2drawml_bad_margins", call = NULL)
  }

  values <- unlist(page_margins_in)
  if (!is.numeric(values)) {
    bad("{.arg page_margins_in} must be numeric, not {.cls {class(page_margins_in)}}.")
  }
  nms <- names(page_margins_in)
  if (is.null(nms)) {
    if (length(values) != 1L) {
      bad("{.arg page_margins_in} must be one number or have names {.val {sides}}.")
    }
    nms <- rep("", length(values))
  }
  if (length(values) == 1L && !nzchar(nms)) {
    names(values) <- NULL
    nms <- sides
    values <- rep(values, 4L)
  } else {
    if (!all(nzchar(nms))) {
      bad("Every element of {.arg page_margins_in} must be named, or give a single number.")
    }
    unknown <- setdiff(nms, sides)
    if (length(unknown) > 0L) {
      bad("Unknown margin names {.val {unknown}}; use {.val {sides}}.")
    }
    if (anyDuplicated(nms) > 0L) {
      bad("Margin names must be unique; {.val {unique(nms[duplicated(nms)])}} repeated.")
    }
  }
  if (anyNA(values)) bad("{.arg page_margins_in} must not contain {.val {NA}}.")
  if (any(values < 0)) bad("{.arg page_margins_in} must not be negative.")

  given <- stats::setNames(as.list(unname(values)), nms)
  as.list(stats::setNames(vapply(sides, function(s) given[[s]] %||% 1, numeric(1)),
                          sides))
}

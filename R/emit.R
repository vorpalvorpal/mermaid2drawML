# R/emit.R
# Scene -> DrawingML: the scale, the items as wps:wsp shapes inside nested
# groups, shape ids, warnings. The tag markup itself is in R/dml.R and
# R/dml_text.R; the scene's columns are described at the top of R/scene.R.

#' Build the DrawingML for a Mermaid diagram
#'
#' Turns a rendered diagram into one Word paragraph holding the diagram as an
#' in-line group of editable shapes. Most users call [mermaid_to_word()] or
#' [body_add_mermaid()], which render the diagram and pass their `...` here.
#'
#' The diagram keeps Mermaid's proportions and is shrunk to fit the page if
#' it is larger; text shrinks with it. Main text below 7 pt signals a
#' `mermaid2drawml_small_text` warning, and parts of the diagram that cannot
#' be drawn as shapes (embedded images) a `mermaid2drawml_unsupported` one.
#'
#' @param x A `mermaid_render` from [render_mermaid()], or the scene built
#'   from one (the `scene` field of a [mermaid_to_word()] result).
#' @param start_id Integer. Id of the drawing; its shapes take the following
#'   ids. Use the `next_id` of the previous diagram when one document holds
#'   several.
#' @param page_width_in,page_height_in Numeric. The space, in inches, the
#'   diagram may fill (default: A4 portrait with 1 inch margins).
#' @param font_scale Numeric. Multiplies every text size. Values below 1 help
#'   when `word_font_family` is wider than the font Mermaid measured.
#' @param word_font_family Character. Font for all text. Default `NULL`: the
#'   font the browser used, which is the one the labels were measured in.
#' @param align Horizontal alignment of the drawing: `"center"`, `"left"` or
#'   `"right"`.
#' @param max_scale Numeric. Largest size, as a multiple of Mermaid's own
#'   (96 px per inch, so 16 px text is 12 pt). Default `1` never enlarges a
#'   small diagram; `Inf` fills the page.
#' @param min_stroke_pt Numeric. Thinnest line drawn, in points (default
#'   0.75), so that lines stay visible on a shrunk diagram.
#' @param alt_text Character. Alternative text for the drawing. Default
#'   `NULL`: the diagram's `accDescr`, else its `accTitle`, else none.
#' @param ... The arguments of earlier versions `ast`, `margin_in`,
#'   `default_fill`, `default_stroke`, `default_text_color` and
#'   `stroke_width_pt` are ignored with a warning (colours and line widths now
#'   come from the diagram itself). Anything else is an error.
#' @return A list: `xml`, the `<w:p>` for [officer::body_add_xml()], and
#'   `next_id`, the first id the drawing did not use.
#' @export
build_diagram_xml <- function(x, start_id = 1L, page_width_in = 6.27,
                              page_height_in = 9.69, font_scale = 1,
                              word_font_family = NULL,
                              align = c("center", "left", "right"),
                              max_scale = 1, min_stroke_pt = 0.75,
                              alt_text = NULL, ...) {
  check_dots(rlang::names2(list(...)))
  align <- match.arg(align)
  scene <- as_scene(x)
  warn_unsupported(scene$unsupported)

  vb <- scene$viewbox
  scale <- diagram_scale(vb, page_width_in, page_height_in, max_scale)
  warn_small_text(main_half_points(scene$items, scale, font_scale))

  ctx <- list(
    scale = scale,
    min_stroke = min_stroke_pt * .emu_per_pt,
    style = list(scale = scale, font_scale = font_scale, font_family = word_font_family)
  )
  w <- emu_size(vb[[3]], scale)
  h <- emu_size(vb[[4]], scale)
  group <- dml_root_group(w, h, children_xml(scene$items, NA_character_, vb[1:2], ctx))
  inline <- dml_inline(w, h, group,
                       descr = first_text(alt_text, scene$desc, scene$title) %||% "",
                       title = first_text(scene$title))
  number_ids(dml_paragraph(align, inline, indents = FALSE, namespaces = TRUE),
             as.integer(start_id))
}

# ---- arguments -----------------------------------------------------------------

.removed_args <- c("ast", "margin_in", "default_fill", "default_stroke",
                   "default_text_color", "stroke_width_pt")

# `...` exists only so that calls written for earlier versions keep working.
check_dots <- function(dot_names, call = rlang::caller_env()) {
  label <- ifelse(nzchar(dot_names), dot_names, paste0("..", seq_along(dot_names)))
  unknown <- label[!dot_names %in% .removed_args]
  if (length(unknown) > 0L) {
    cli::cli_abort("Unknown argument{?s} {.arg {unknown}}.",
                   class = "mermaid2drawml_unknown_arg", call = call)
  }
  removed <- intersect(dot_names, .removed_args)
  if (length(removed) > 0L) {
    cli::cli_warn(
      c("{.arg {removed}} {?is/are} no longer used and {?was/were} ignored.",
        i = "Colours, line widths and fonts now come from the diagram itself."),
      class = "mermaid2drawml_deprecated", call = call
    )
  }
  invisible(NULL)
}

as_scene <- function(x, call = rlang::caller_env()) {
  if (inherits(x, "mermaid_scene")) return(x)
  if (inherits(x, "mermaid_render")) return(build_scene(x))
  cli::cli_abort(
    "{.arg x} must be a {.cls mermaid_render} or {.cls mermaid_scene}, not {.obj_type_friendly {x}}.",
    class = "mermaid2drawml_bad_input", call = call
  )
}

# The first argument that is a non-empty string, or NULL.
first_text <- function(...) {
  purrr::detect(list(...), function(v) length(v) == 1L && !is.na(v) && nzchar(v))
}

warn_unsupported <- function(tags) {
  if (length(tags) == 0L) return(invisible(NULL))
  cli::cli_warn(
    "Left out what Word shapes cannot show: {.val {tags}}.",
    class = "mermaid2drawml_unsupported"
  )
}

# ---- scale and units -------------------------------------------------------------

#' EMU per px: fit the page, never above max_scale times natural size
#'
#' A zero-size diagram constrains nothing, so it is drawn at natural size.
#' @noRd
diagram_scale <- function(vb, page_width_in, page_height_in, max_scale) {
  fit <- c(page_width_in * .emu_per_inch / vb[[3]],
           page_height_in * .emu_per_inch / vb[[4]],
           max_scale * .emu_per_px)
  fit <- fit[is.finite(fit) & fit > 0]
  if (length(fit) == 0L) .emu_per_px else min(fit)
}

# Whole EMU as a double: 2^31 EMU is only about 225 000 px, which a large
# diagram can pass.
emu <- function(px, scale) round(px * scale)

# Word rejects a zero extent.
emu_size <- function(px, scale) max(1, emu(px, scale))

# ---- items -> markup ---------------------------------------------------------------

#' The markup of the items directly inside group `parent` (NA: the root),
#' placed relative to `origin` (root px), each group with its own children
#' @noRd
children_xml <- function(items, parent, origin, ctx) {
  here <- if (is.na(parent)) is.na(items$parent) else items$parent %in% parent
  xml <- purrr::pmap_chr(items[here, ], function(...) {
    item_xml(list(...), items, origin, ctx)
  })
  paste(xml, collapse = "")
}

# Every number an item puts in the markup. A non-finite one would be dropped
# or written as "NA", so the item is named in an error instead.
item_numbers <- function(it) {
  p <- it$path
  drawn <- if (!is.null(p)) p[p$cmd != "Z", ]
  curved <- if (!is.null(p)) p[p$cmd == "C", ]
  c(it$x, it$y, it$w, it$h, if (it$kind != "group") c(it$rot, it$stroke_w, it$text$insets),
    drawn$x, drawn$y, curved$x1, curved$y1, curved$x2, curved$y2)
}

check_item_finite <- function(it) {
  num <- item_numbers(it)
  if (!all(is.finite(num))) {
    cli::cli_abort("The item {.val {it$name}} has a position or size that is not a finite number.",
                   class = "mermaid2drawml_render_error")
  }
  it
}

item_xml <- function(it, items, origin, ctx) {
  check_item_finite(it)
  s <- ctx$scale
  if (it$kind == "group") {
    return(dml_grp(it$name, emu(it$x - origin[[1]], s), emu(it$y - origin[[2]], s),
                   emu_size(it$w, s), emu_size(it$h, s),
                   children_xml(items, it$name, c(it$x, it$y), ctx)))
  }
  shape_xml(it, origin, ctx)
}

# A preset from the ratio of the first dash to the stroke width: up to 1.5
# dotted, 3.5 short dashes, 6 dashes, longer ones long dashes.
.dash_presets <- tibble::tibble(upto = c(1.5, 3.5, 6, Inf),
                                prst = c("sysDot", "sysDash", "dash", "lgDash"))

dash_preset <- function(dash, stroke_w) {
  if (length(dash) == 0L || !any(dash > 0) || !isTRUE(stroke_w > 0)) return(NA_character_)
  ratio <- dash[[1]] / stroke_w
  .dash_presets$prst[[findInterval(ratio, .dash_presets$upto, left.open = TRUE) + 1L]]
}

#' The geometry element of a drawn item: its preset, its outline in EMU
#' relative to its box, or a plain rectangle for a free text box
#' @noRd
geometry_xml <- function(it, w, h, s) {
  if (!is.na(it$prst)) return(dml_prst_geom(it$prst, it$adj))
  if (is.null(it$path)) return(dml_prst_geom("rect"))
  p <- it$path
  rel <- function(v, o) (v - o) * s
  local <- tibble::tibble(cmd = p$cmd, x1 = rel(p$x1, it$x), y1 = rel(p$y1, it$y),
                          x2 = rel(p$x2, it$x), y2 = rel(p$y2, it$y),
                          x = rel(p$x, it$x), y = rel(p$y, it$y))
  dml_cust_geom(local, w, h)
}

# A connector cannot hold text or a fill, so an open outline that took a label
# or is filled stays a shape.
shape_role <- function(it) {
  if (it$kind == "text") return("text_box")
  if (isFALSE(it$closed) && is.na(it$fill) && is.null(it$text)) "connector" else "shape"
}

shape_xml <- function(it, origin, ctx) {
  s <- ctx$scale
  w <- emu_size(it$w, s)
  h <- emu_size(it$h, s)
  free_text <- it$kind == "text"
  stroked <- !free_text && !is.na(it$stroke)
  line_w <- if (stroked) max(it$stroke_w * s, ctx$min_stroke) else 0
  sp_pr <- c(
    dml_xfrm(emu(it$x - origin[[1]], s), emu(it$y - origin[[2]], s), w, h,
             rot = it$rot, flip_h = it$flip_h, flip_v = it$flip_v),
    geometry_xml(it, w, h, s),
    dml_fill(if (free_text) NA_character_ else it$fill, it$fill_a),
    dml_ln(line_w, if (stroked) it$stroke else NA_character_, it$stroke_a,
           dash = dash_preset(it$dash, it$stroke_w), head = it$head, tail = it$tail,
           head_size = c(w = it$head_w, len = it$head_len),
           tail_size = c(w = it$tail_w, len = it$tail_len))
  )
  has_text <- !is.null(it$text) && length(it$text$rows) > 0L
  dml_wsp(
    it$name, shape_role(it), sp_pr,
    txbx = if (has_text) dml_txbx(it$text, ctx$style),
    body_pr = if (has_text) {
      dml_body_pr(emu(it$text$insets, s), it$text$valign %||% "center")
    } else {
      xml_tag("wps:bodyPr")
    }
  )
}

# ---- ids ---------------------------------------------------------------------------

#' Number the id marks of a finished drawing in document order
#'
#' The docPr comes first and takes `start_id`, every cNvPr the next integer.
#' @return list(xml, next_id).
#' @noRd
number_ids <- function(xml, start_id) {
  marks <- gregexpr(.id_mark, xml, fixed = TRUE)
  n <- sum(marks[[1]] > 0L)
  regmatches(xml, marks) <- list(as.character(start_id + seq_len(n) - 1L))
  list(xml = xml, next_id = start_id + n)
}

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
#' `mermaid2drawml_small_text` warning, whose `points` field holds that size,
#' and parts of the diagram that cannot be drawn as shapes (`<use>`
#' elements, pictures the browser could not read, pictures left out with
#' `embed_images = FALSE`) a `mermaid2drawml_unsupported` one, whose `tags`
#' field names them.
#'
#' Pictures and click links need parts of the Word document package that a
#' paragraph's XML cannot carry. The XML refers to them by placeholder ids,
#' listed in the result's `images` and `links`, and [body_add_mermaid()]
#' adds them to the document. When either has rows, the XML alone is not
#' complete: before inserting it yourself with [officer::body_add_xml()],
#' replace each image `rel_id` in the XML by the path of a PNG file holding
#' its `bytes` (officer then embeds the file when the document is saved), and
#' add each link as an external hyperlink relationship of the document body
#' (`officer::docx_body_relationship(doc)$add()`), replacing its `rel_id` by
#' that relationship's id. Or use `links = "field"` and
#' `embed_images = FALSE`, which give XML that needs nothing else.
#'
#' @param x A `mermaid_render` from [render_mermaid()], or the scene built
#'   from one (the `scene` field of a [mermaid_to_word()] result).
#' @param start_id Integer. Id of the drawing; its shapes take the following
#'   ids. Use the `next_id` of the previous diagram when one document holds
#'   several.
#' @param page_width_in,page_height_in Numeric, or `NULL` for the default. The
#'   space, in inches, the diagram may fill (default: A4 portrait with 1 inch
#'   margins). Each must be one finite positive number.
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
#' @param links How Mermaid's links (`click X href`; web and mail addresses
#'   only) are written: `"shape"` (default) makes each linked shape a click
#'   target, which needs a relationship in the document (see Details);
#'   `"field"` wraps the linked text in a Word `HYPERLINK` field, which needs
#'   nothing else but links only shapes that hold text; `"none"` leaves links
#'   out.
#' @param embed_images Logical. Draw the pictures the diagram holds (default
#'   `TRUE`), which need parts in the document (see Details); `FALSE` leaves
#'   them out with a `mermaid2drawml_unsupported` warning.
#' @param ... The arguments of earlier versions `ast`, `margin_in`,
#'   `default_fill`, `default_stroke`, `default_text_color` and
#'   `stroke_width_pt` are ignored with a warning (colours and line widths now
#'   come from the diagram itself). Anything else is an error.
#' @return A list: `xml`, the `<w:p>` for [officer::body_add_xml()];
#'   `next_id`, the first id the drawing did not use; `links`, a tibble of
#'   `rel_id` (the placeholder in the XML) and `url`, one row per linked URL;
#'   `images`, a tibble of `rel_id` (the placeholder) and `bytes` (a list of
#'   PNG files as raw vectors), one row per distinct picture. Both have no
#'   rows when the XML needs nothing else.
#' @export
build_diagram_xml <- function(x, start_id = 1L, page_width_in = NULL,
                              page_height_in = NULL, font_scale = 1,
                              word_font_family = NULL,
                              align = c("center", "left", "right"),
                              max_scale = 1, min_stroke_pt = 0.75,
                              alt_text = NULL, links = c("shape", "field", "none"),
                              embed_images = TRUE, ...) {
  check_dots(rlang::names2(list(...)))
  page_width_in <- page_size_in(page_width_in, .default_page_width_in)
  page_height_in <- page_size_in(page_height_in, .default_page_height_in)
  align <- match.arg(align)
  links <- match.arg(links)
  scene <- as_scene(x)
  items <- scene$items
  left_out <- !isTRUE(embed_images) && any(items$kind == "image")
  if (left_out) items <- without_images(items)
  warn_unsupported(unique(c(scene$unsupported, if (left_out) "image")))

  vb <- scene$viewbox
  scale <- diagram_scale(vb, page_width_in, page_height_in, max_scale)
  warn_small_text(main_half_points(items, scale, font_scale))

  start_id <- as.integer(start_id)
  parts <- package_parts(items, links, start_id)
  ctx <- list(
    scale = scale,
    min_stroke = min_stroke_pt * .emu_per_pt,
    style = list(scale = scale, font_scale = font_scale, font_family = word_font_family),
    click_id = parts$click_id, field_href = parts$field_href, image_id = parts$image_id
  )
  w <- emu_size(vb[[3]], scale)
  h <- emu_size(vb[[4]], scale)
  group <- dml_root_group(w, h, children_xml(items, NA_character_, vb[1:2], ctx))
  inline <- dml_inline(w, h, group,
                       descr = first_text(alt_text, scene$desc, scene$title) %||% "",
                       title = first_text(scene$title))
  drawing <- number_ids(dml_paragraph(align, inline, indents = FALSE, namespaces = TRUE),
                        start_id)
  c(drawing, list(links = parts$links, images = parts$images))
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
    class = "mermaid2drawml_unsupported", tags = tags
  )
}

# ---- scale and units ---------------------------------------------------------

# A4 portrait inside 1 inch margins.
.default_page_width_in <- 6.27
.default_page_height_in <- 9.69

#' A page dimension: the default for NULL, else one finite positive number
#' @noRd
page_size_in <- function(value, default, arg = rlang::caller_arg(value),
                         call = rlang::caller_env()) {
  if (is.null(value)) return(default)
  if (!is.numeric(value) || length(value) != 1L || !is.finite(value) || value <= 0) {
    cli::cli_abort(
      "{.arg {arg}} must be one finite positive number of inches, or {.code NULL}.",
      class = "mermaid2drawml_bad_arg", call = call
    )
  }
  value
}

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
  # The rows are split by parent once, so that each group finds its children
  # without scanning every row.
  ctx$children <- rows_by_key(items$parent)
  group_children_xml(items, parent, origin, ctx)
}

group_children_xml <- function(items, parent, origin, ctx) {
  here <- ctx$children(parent)
  xml <- vapply(here, function(r) item_xml(item_row(items, r), items, origin, ctx), character(1))
  paste(xml, collapse = "")
}

# Row `r` of the items as a plain list, list columns unwrapped.
item_row <- function(items, r) lapply(items, `[[`, r)

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
  write <- .item_writers[[it$kind]] %||% shape_xml
  write(it, items, origin, ctx)
}

group_xml <- function(it, items, origin, ctx) {
  s <- ctx$scale
  dml_grp(it$name, emu(it$x - origin[[1]], s), emu(it$y - origin[[2]], s),
          emu_size(it$w, s), emu_size(it$h, s),
          group_children_xml(items, it$name, c(it$x, it$y), ctx))
}

# A picture is placed exactly as a shape is.
picture_xml <- function(it, items, origin, ctx) {
  s <- ctx$scale
  xfrm <- dml_xfrm(emu(it$x - origin[[1]], s), emu(it$y - origin[[2]], s),
                   emu_size(it$w, s), emu_size(it$h, s), rot = it$rot)
  dml_pic(it$name, xfrm, ctx$image_id(it$png), link = ctx$click_id(it$href))
}

# Item kinds with markup of their own; every other kind is a wps:wsp.
.item_writers <- list(group = group_xml, image = picture_xml)

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
  local <- path_map_points(it$path, function(x, y) list(x = (x - it$x) * s, y = (y - it$y) * s))
  dml_cust_geom(local, w, h)
}

# A connector cannot hold text or a fill, so an open outline that took a label
# or is filled stays a shape.
shape_role <- function(it) {
  if (it$kind == "text") return("text_box")
  if (isFALSE(it$closed) && is.na(it$fill) && is.null(it$text)) "connector" else "shape"
}

shape_xml <- function(it, items, origin, ctx) {
  s <- ctx$scale
  w <- emu_size(it$w, s)
  h <- emu_size(it$h, s)
  free_text <- it$kind == "text"
  stroked <- !free_text && !is.na(it$stroke)
  line_w <- if (stroked) max(it$stroke_w * s, ctx$min_stroke) else 0
  frame <- list(box = c(it$x, it$y, it$w, it$h), rot = it$rot, flip_h = it$flip_h,
                flip_v = it$flip_v)
  sp_pr <- c(
    dml_xfrm(emu(it$x - origin[[1]], s), emu(it$y - origin[[2]], s), w, h,
             rot = it$rot, flip_h = it$flip_h, flip_v = it$flip_v),
    geometry_xml(it, w, h, s),
    paint_xml(if (free_text) NA_character_ else it$fill, it$fill_a, it$fill_grad,
              it$fill_pattern, frame),
    dml_ln(line_w,
           if (stroked) paint_xml(it$stroke, it$stroke_a, it$stroke_grad, NA_character_, frame),
           dash = dash_preset(it$dash, it$stroke_w), head = it$head, tail = it$tail,
           head_size = c(w = it$head_w, len = it$head_len),
           tail_size = c(w = it$tail_w, len = it$tail_len))
  )
  has_text <- !is.null(it$text) && length(it$text$rows) > 0L
  dml_wsp(
    it$name, shape_role(it), sp_pr,
    txbx = if (has_text) dml_txbx(it$text, ctx$style, href = ctx$field_href(it$href)),
    body_pr = if (has_text) {
      dml_body_pr(emu(it$text$insets, s), it$text$valign)
    } else {
      xml_tag("wps:bodyPr")
    },
    link = ctx$click_id(it$href)
  )
}

# ---- paint ---------------------------------------------------------------------

#' How a fill or stroke is painted: "none", "solid", "pattern", "linear" or
#' "radial". A gradient of one colour throughout is solid.
#' @noRd
paint_kind <- function(color, grad, pattern) {
  if (is.na(color)) return("none")
  if (!is.na(pattern)) return("pattern")
  if (is.null(grad) || one_colour(grad$stops)) return("solid")
  grad$type
}

one_colour <- function(stops) {
  nrow(vctrs::vec_unique(stops[c("color", "alpha")])) == 1L
}

#' The fill element of a fill or stroke
#' @param frame list(box = c(x, y, w, h), rot, flip_h, flip_v) in root px:
#'   the item a gradient is expressed in.
#' @noRd
paint_xml <- function(color, alpha, grad, pattern, frame) {
  paint <- list(color = color, alpha = alpha, grad = grad, pattern = pattern)
  .paint_writers[[paint_kind(color, grad, pattern)]](paint, frame)
}

.paint_writers <- list(
  none = function(paint, frame) dml_fill(NA_character_),
  solid = function(paint, frame) {
    # A one-colour gradient is drawn in its stops' colour.
    s <- paint$grad$stops
    if (is.null(s)) dml_fill(paint$color, paint$alpha) else dml_fill(s$color[[1]], s$alpha[[1]])
  },
  pattern = function(paint, frame) dml_patt_fill(paint$pattern, paint$color, paint$alpha),
  linear = function(paint, frame) {
    g <- linear_in_frame(paint$grad, frame)
    dml_grad_fill(g$stops, ang = g$ang)
  },
  radial = function(paint, frame) {
    g <- radial_in_frame(paint$grad, frame)
    dml_grad_fill(g$stops, fill_to_rect = g$fill_to_rect)
  }
)

#' A root-space point in an item's own frame: unturned by its rotation and
#' mirrored by its flips about its centre. Word turns and mirrors a shape's
#' fill with the shape, so a gradient given in this frame is drawn where the
#' SVG drew it.
#' @noRd
frame_point <- function(pt, frame) {
  b <- frame$box
  centre <- c(b[[1]] + b[[3]] / 2, b[[2]] + b[[4]] / 2)
  r <- -frame$rot * pi / 180
  v <- pt - centre
  v <- c(v[[1]] * cos(r) - v[[2]] * sin(r), v[[1]] * sin(r) + v[[2]] * cos(r))
  centre + v * ifelse(c(isTRUE(frame$flip_h), isTRUE(frame$flip_v)), -1, 1)
}

#' A linear gradient in an item's frame: its angle (60000ths of a degree
#' clockwise) and its stops re-expressed along the box's extent in that
#' direction, as DrawingML spans a linear gradient over the whole box
#' @noRd
linear_in_frame <- function(g, frame) {
  p1 <- frame_point(g$p1, frame)
  p2 <- frame_point(g$p2, frame)
  d <- p2 - p1
  len <- sqrt(sum(d^2))
  if (!isTRUE(len > 0)) d <- c(1, 0)
  u <- d / sqrt(sum(d^2))
  b <- frame$box
  corners <- rbind(c(b[[1]], b[[2]]), c(b[[1]] + b[[3]], b[[2]]),
                   c(b[[1]], b[[2]] + b[[4]]), c(b[[1]] + b[[3]], b[[2]] + b[[4]]))
  proj <- as.vector(corners %*% u)
  extent <- max(proj) - min(proj)
  along <- sum(p1 * u) + g$stops$o * len
  share <- if (extent > 0) (along - min(proj)) / extent else g$stops$o
  ang <- round((atan2(u[[2]], u[[1]]) * 180 / pi) %% 360 * 60000) %% 21600000
  list(ang = ang, stops = grad_stops(g$stops, pmin(1, pmax(0, share))))
}

#' A radial gradient in an item's frame: its centre as a:fillToRect insets
#' from the box's sides. Word's circular path fill only approximates SVG's
#' circle, so the stops keep their own offsets.
#' @noRd
radial_in_frame <- function(g, frame) {
  c0 <- frame_point(g$c, frame)
  b <- frame$box
  w <- if (b[[3]] > 0) b[[3]] else 1
  h <- if (b[[4]] > 0) b[[4]] else 1
  ftr <- c((c0[[1]] - b[[1]]) / w, (c0[[2]] - b[[2]]) / h,
           (b[[1]] + b[[3]] - c0[[1]]) / w, (b[[2]] + b[[4]] - c0[[2]]) / h) * 100000
  list(fill_to_rect = ftr, stops = grad_stops(g$stops, g$stops$o))
}

grad_stops <- function(stops, share) {
  vctrs::new_data_frame(list(pos = share * 100000, color = stops$color, alpha = stops$alpha))
}

# ---- links and pictures ----------------------------------------------------------
#
# A picture and a click link each need a relationship in the document
# package, which a drawing's XML cannot add. The XML refers to them by
# placeholder ids, listed with what they stand for; body_add_mermaid() puts
# them into the document. Placeholders hold the drawing's start id, so the
# diagrams of one document have different ones, and never look like the
# rId<n> ids officer and Word give.

.placeholder_prefix <- "rIdM2D"

placeholder_ids <- function(start_id, kind, n) {
  paste0(.placeholder_prefix, start_id, kind, seq_len(n))
}

# How each `links` mode writes an item's link: as its shape's click target
# (`click`), as a field around its text (`field`), or not at all.
.link_modes <- list(
  shape = c(click = TRUE, field = FALSE),
  field = c(click = FALSE, field = TRUE),
  none = c(click = FALSE, field = FALSE)
)

#' The package parts a drawing refers to: `links` (rel_id, url) and
#' `images` (rel_id, bytes), one row per distinct URL or picture in
#' document order, and look-ups from an item's href or bytes to what its
#' markup writes (NULL for nothing)
#' @noRd
package_parts <- function(items, links, start_id) {
  mode <- .link_modes[[links]]
  drawn <- items$kind != "group"
  urls <- if (mode[["click"]]) unique(stats::na.omit(items$href[drawn])) else character()
  link_tbl <- tibble::tibble(rel_id = placeholder_ids(start_id, "L", length(urls)), url = urls)
  pictures <- vctrs::vec_unique(items$png[items$kind == "image"])
  image_tbl <- tibble::tibble(rel_id = placeholder_ids(start_id, "I", length(pictures)),
                              bytes = pictures)
  list(
    links = link_tbl, images = image_tbl,
    click_id = function(href) {
      if (!is.na(href) && mode[["click"]]) link_tbl$rel_id[[match(href, link_tbl$url)]]
    },
    field_href = function(href) if (!is.na(href) && mode[["field"]]) href,
    image_id = function(bytes) image_tbl$rel_id[[vctrs::vec_match(list(bytes), pictures)]]
  )
}

#' Items without their pictures, and without the groups left holding
#' nothing drawn
#' @noRd
without_images <- function(items) {
  items <- items[items$kind != "image", ]
  g <- items$kind == "group"
  parent_of <- stats::setNames(items$parent[g], items$name[g])
  live <- holding_groups(unique(stats::na.omit(items$parent[!g])), parent_of)
  items[!g | items$name %in% live, ]
}

# Groups `live` and every group above them.
holding_groups <- function(live, parent_of) {
  up <- setdiff(stats::na.omit(unname(parent_of[live])), live)
  if (length(up) == 0L) live else holding_groups(c(live, up), parent_of)
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

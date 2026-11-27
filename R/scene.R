# R/scene.R
# Bridge prims -> scene items: the generic conversion every diagram type gets
# (spec 3a) and putting text inside shapes (3b). What nodes, subgraphs, edges
# and edge labels are, the flowchart shape registry and the marker table live
# in R/semantics.R; this file knows nothing about diagram types.
#
# ---- The items tibble: contract with the DrawingML emitter -------------------
#
# `scene$items` has one row per Word object. Rows are in document order: depth
# first, so a group row is immediately followed by all of its descendants, and
# siblings are in paint order (back to front). Lengths are root px (the root
# svg's user space, the units of `scene$viewbox`). Colours are "RRGGBB" with no
# "#", or NA for none.
#
# Documented columns (spec 3):
#   name       chr  Word shape name (cNvPr name). Group names are unique; other
#                   names may repeat (generic "mermaid:text", two edges A-B).
#   kind       chr  "group" | "shape" | "line" | "text"
#   parent     chr  `name` of the enclosing group row, NA at the top level
#   x, y, w, h dbl  box. With rot != 0 it is the unrotated box, centred where
#                   the item is drawn. Group rows: the union of the boxes of
#                   their descendants (rotation ignored).
#   rot        dbl  degrees clockwise about the box centre (0 for groups,
#                   lines and custom geometry)
#   prst       chr  preset name ("rect", "roundRect", "ellipse", "line", ...)
#                   or NA for custom geometry, free text and groups
#   fill, stroke         chr  colour or NA
#   fill_a, stroke_a     dbl  alpha 0-1 (0 when the paint is NA)
#   stroke_w   dbl  stroke width, px (0 when there is no stroke)
#   head, tail chr  line end at the path START / END: "none", "triangle",
#                   "arrow", "stealth", "diamond", "oval"
#   label      chr  the item's text, rows joined with "\n"; "" when none
#
# Internal columns:
#   flip_h, flip_v lgl   a:xfrm flipH / flipV. For a line (kind "line", preset
#                   "line", which runs from the box's top left to its bottom
#                   right) the flips put its start (head end) and end (tail
#                   end) where they were drawn: flip_h when the line runs right
#                   to left, flip_v when it runs bottom to top.
#   adj        list named numeric vector of preset adjust values, DrawingML
#                   units, e.g. c(adj = 20000) -> <a:gd name="adj" fmla="val
#                   20000"/>; numeric(0) means none (the preset's defaults).
#   path       list custom geometry only (kind "shape" with prst NA): a tibble
#                   of ABSOLUTE segments in root px, columns cmd ("M", "L",
#                   "C", "Z"), x1, y1, x2, y2 (cubic control points, NA for
#                   M/L/Z), x, y (end point, NA for Z); see R/geometry.R.
#                   Subtract the item's x, y to get a:path coordinates; the
#                   path's own extent is the item's w, h. NULL otherwise.
#   closed     lgl  FALSE for lines and open custom paths (they take the
#                   connector properties, wps:cNvCnPr); TRUE for presets and
#                   closed paths; NA for groups and free text.
#   dash       list numeric dash array in root px, NULL for a solid line
#   text       list NULL when the item has no text. Otherwise a list:
#                   rows      list of tibbles, one per paragraph (Mermaid's own
#                             text rows), one row per run, columns t (chr),
#                             bold, italic, underline (lgl), color (chr
#                             "RRGGBB"), size (dbl, px), font (chr, the face
#                             the browser used, may be ""), family (chr, the
#                             CSS font-family list)
#                   align     "left" | "center" | "right" (w:jc)
#                   placement "centre": merged and centred in its shape
#                             (anchor ctr, centred paragraphs, zero insets);
#                             "inset": merged off centre, `insets` put it where
#                             Mermaid drew it; "free": a text box of its own
#                             (kind "text", zero insets)
#                   insets    numeric c(l, t, r, b), px, >= 0 (zeros unless
#                             placement is "inset")
#                   box       numeric c(x, y, w, h), px: where the browser
#                             drew the text
#                   anchor    the bridge's anchor, "start" | "middle" | "end"
#                   html      lgl, TRUE for foreignObject (HTML) text
# -----------------------------------------------------------------------------

#' Build the scene: what to draw, one row per Word object
#'
#' @param render A `mermaid_render`.
#' @return A `mermaid_scene`: a list with `type`, `viewbox` (numeric 4),
#'   `title`, `desc`, `unsupported` (tags of prims that were not drawn) and
#'   `items` (see the comment block at the top of `R/scene.R`).
#' @noRd
build_scene <- function(render) {
  if (!inherits(render, "mermaid_render")) {
    cli::cli_abort(
      "{.arg render} must be a {.cls mermaid_render}, not {.obj_type_friendly {render}}.",
      class = "mermaid2drawml_bad_input"
    )
  }
  pieces <- scene_pieces(render$prims, render$viewBox)
  structure(
    list(type = render$type, viewbox = render$viewBox, title = render$title,
         desc = render$desc, unsupported = scene_unsupported(render$prims),
         items = scene_items(pieces, render$db)),
    class = "mermaid_scene"
  )
}

# Tags of the prims that are reported but not drawn (images, <use>, ...).
scene_unsupported <- function(prims) {
  skipped <- purrr::keep(prims, function(p) p$kind %in% c("image", "other"))
  unique(purrr::map_chr(skipped, function(p) p$tag %||% p$kind))
}

# ---- prims -> pieces ---------------------------------------------------------
#
# A piece is one drawable prim converted to root space: an items row plus the
# columns the folding and unit steps need: i (the prim's index), pos (paint
# position), src (the prim's kind), html, corner_r (corner radius of a rounded
# rect, or the first arc radius of a path, root px), id, cls, anc, data.

item_cols <- c("name", "kind", "parent", "x", "y", "w", "h", "rot", "prst",
               "fill", "stroke", "fill_a", "stroke_a", "stroke_w", "head",
               "tail", "label", "flip_h", "flip_v", "adj", "path", "closed",
               "dash", "text")

piece_ptype <- function() {
  tibble::tibble(
    i = integer(), pos = numeric(), src = character(), name = character(),
    kind = character(),
    parent = character(), x = numeric(), y = numeric(), w = numeric(),
    h = numeric(), rot = numeric(), prst = character(), fill = character(),
    stroke = character(), fill_a = numeric(), stroke_a = numeric(),
    stroke_w = numeric(), head = character(), tail = character(),
    label = character(), flip_h = logical(), flip_v = logical(), adj = list(),
    path = list(), closed = logical(), dash = list(), text = list(),
    html = logical(), corner_r = numeric(), id = character(), cls = character(),
    anc = list(), data = list()
  )
}

#' Every drawable prim as a piece in root space, without those wholly outside
#' the viewBox
#' @noRd
scene_pieces <- function(prims, viewbox) {
  converters <- list(
    rect = piece_rect, ellipse = piece_ellipse, line = piece_line,
    poly = function(p) piece_outline(p, poly_path(p$geom$points, isTRUE(p$geom$closed))),
    path = function(p) piece_outline(p, path_parse(p$geom$d %||% "")),
    text = piece_text
  )
  drawable <- purrr::keep(prims, function(p) p$kind %in% names(converters))
  pieces <- purrr::list_rbind(
    purrr::map(drawable, function(p) converters[[p$kind]](p)),
    ptype = piece_ptype()
  )
  pieces[box_in_view(pieces, viewbox), ]
}

# Mermaid can park items far off the canvas (the gantt "today" marker); an
# item that only straddles the edge is kept.
box_in_view <- function(b, vb) {
  known <- !is.na(b$x) & !is.na(b$y) & !is.na(b$w) & !is.na(b$h)
  outside <- b$x > vb[[1]] + vb[[3]] | b$x + b$w < vb[[1]] |
    b$y > vb[[2]] + vb[[4]] | b$y + b$h < vb[[2]]
  known & !outside
}

new_piece <- function(p, kind, box, rot = 0, prst = NA_character_,
                      paint = no_paint(), path = NULL, closed = NA,
                      corner_r = NA_real_, flip_h = FALSE, flip_v = FALSE,
                      adj = numeric(), text = NULL, label = "") {
  tibble::tibble(
    i = as.integer(p$i %||% NA_integer_), pos = as.numeric(p$i %||% NA_real_),
    src = as.character(p$kind), name = NA_character_, kind = kind,
    parent = NA_character_,
    x = box[[1]], y = box[[2]], w = box[[3]], h = box[[4]], rot = rot,
    prst = prst, fill = paint$fill, stroke = paint$stroke,
    fill_a = paint$fill_a, stroke_a = paint$stroke_a, stroke_w = paint$stroke_w,
    head = paint$head, tail = paint$tail, label = label,
    flip_h = flip_h, flip_v = flip_v, adj = list(adj), path = list(path),
    closed = closed, dash = list(paint$dash), text = list(text),
    html = isTRUE(p$html), corner_r = corner_r,
    id = as.character(p$id %||% ""), cls = as.character(p$cls %||% ""),
    anc = list(p$anc %||% list()), data = list(p$data %||% list())
  )
}

# "#rrggbb" from the bridge -> "RRGGBB"; a missing paint is NA.
bridge_colour <- function(x) {
  if (is.null(x) || length(x) == 0L || is.na(x[[1]]) || !nzchar(x[[1]])) {
    return(NA_character_)
  }
  toupper(sub("^#", "", x[[1]]))
}

prim_number <- function(x, default = 0) {
  if (is.null(x) || length(x) == 0L) default else as.numeric(x[[1]])
}

prim_matrix <- function(p) {
  m <- as.numeric(unlist(p$m))
  if (length(m) == 6L) m else c(1, 0, 0, 1, 0, 0)
}

no_paint <- function() {
  list(fill = NA_character_, fill_a = 0, stroke = NA_character_, stroke_a = 0,
       stroke_w = 0, dash = NULL, head = "none", tail = "none")
}

prim_paint <- function(p) {
  fill <- bridge_colour(p$fill)
  stroke <- bridge_colour(p$stroke)
  stroked <- !is.na(stroke)
  list(
    fill = fill,
    fill_a = if (is.na(fill)) 0 else prim_number(p$fill_a, 1),
    stroke = stroke,
    stroke_a = if (stroked) prim_number(p$stroke_a, 1) else 0,
    stroke_w = if (stroked) prim_number(p$stroke_w, 1) else 0,
    dash = if (stroked && length(p$dash) > 0L) as.numeric(unlist(p$dash)) else NULL,
    head = line_end_kind(p$marker_start),
    tail = line_end_kind(p$marker_end)
  )
}

# The root-space box of a local rectangle: upright matrices (translate, scale,
# mirror) map it directly; a rotated one gives the unrotated box about the
# mapped centre plus the rotation.
frame_box <- function(x, y, w, h, m) {
  if (mat_is_upright(m)) {
    corners <- mat_apply(m, c(x, x + w), c(y, y + h))
    return(list(box = c(min(corners$x), min(corners$y),
                        abs(diff(corners$x)), abs(diff(corners$y))),
                rot = 0))
  }
  centre <- mat_apply(m, x + w / 2, y + h / 2)
  rw <- w * sqrt(m[[1]]^2 + m[[2]]^2)
  rh <- h * sqrt(m[[3]]^2 + m[[4]]^2)
  list(box = c(centre$x - rw / 2, centre$y - rh / 2, rw, rh),
       rot = mat_rotation(m))
}

#' roundRect adjust for a corner radius: round(100000 r / min(w, h)), at most
#' 50000 (a fully rounded end)
#' @noRd
round_adj <- function(r, box) {
  c(adj = min(50000, round(100000 * r / min(box[[3]], box[[4]]))))
}

piece_rect <- function(p) {
  g <- p$geom
  m <- prim_matrix(p)
  fb <- frame_box(prim_number(g$x), prim_number(g$y), prim_number(g$w),
                  prim_number(g$h), m)
  # SVG takes the missing radius from the other one.
  rx <- prim_number(g$rx)
  r <- (if (rx > 0) rx else prim_number(g$ry)) * mat_scale(m)
  rounded <- r > 0
  new_piece(p, "shape", fb$box, fb$rot,
            prst = if (rounded) "roundRect" else "rect",
            paint = prim_paint(p), closed = TRUE,
            corner_r = if (rounded) r else NA_real_,
            adj = if (rounded) round_adj(r, fb$box) else numeric())
}

piece_ellipse <- function(p) {
  g <- p$geom
  rx <- prim_number(g$rx)
  ry <- prim_number(g$ry)
  fb <- frame_box(prim_number(g$cx) - rx, prim_number(g$cy) - ry, 2 * rx, 2 * ry,
                  prim_matrix(p))
  new_piece(p, "shape", fb$box, fb$rot, prst = "ellipse", paint = prim_paint(p),
            closed = TRUE)
}

piece_line <- function(p) {
  g <- p$geom
  ends <- mat_apply(prim_matrix(p), c(prim_number(g$x1), prim_number(g$x2)),
                    c(prim_number(g$y1), prim_number(g$y2)))
  paint <- prim_paint(p)
  # SVG never fills a line.
  paint$fill <- NA_character_
  paint$fill_a <- 0
  new_piece(p, "line",
            c(min(ends$x), min(ends$y), abs(diff(ends$x)), abs(diff(ends$y))),
            prst = "line", paint = paint, closed = FALSE,
            flip_h = ends$x[[2]] < ends$x[[1]], flip_v = ends$y[[2]] < ends$y[[1]])
}

# An open path with a Z appended, so that a four-corner outline can be
# recognised as the rectangle SVG fills.
close_subpath <- function(path) {
  if (path_closed(path)) path else purrr::list_rbind(list(path, path_tbl("Z")))
}

# Radius of the first arc in path data, in the path's own units, or NA.
first_arc_radius <- function(d) {
  if (is.null(d) || !nzchar(d)) return(NA_real_)
  hit <- stringr::str_match(d, paste0("[Aa][\\s,]*(", .num_re, ")"))[, 2]
  abs(as.numeric(hit))
}

piece_outline <- function(p, local_path) {
  if (nrow(local_path) == 0L) return(NULL)
  m <- prim_matrix(p)
  path <- path_transform(local_path, m)
  paint <- prim_paint(p)
  closed <- path_closed(path)
  # Closing an open outline would add a stroked side, so only an unstroked
  # open rectangle may become the preset.
  rect <- if (closed || is.na(paint$stroke)) path_as_rect(close_subpath(path))
  box <- if (is.null(rect)) path_bbox(path) else rect
  if (anyNA(box)) return(NULL)
  arc_r <- if (identical(p$kind, "path")) first_arc_radius(p$geom$d) * mat_scale(m) else NA_real_
  new_piece(p, "shape", box,
            prst = if (is.null(rect)) NA_character_ else "rect",
            paint = paint, path = path, closed = closed || !is.null(rect),
            corner_r = if (isTRUE(arc_r > 0)) arc_r else NA_real_)
}

anchor_align <- function(anchor) {
  aligns <- c(start = "left", middle = "center", end = "right")
  if (length(anchor) == 1L && anchor %in% names(aligns)) aligns[[anchor]] else "left"
}

text_runs <- function(runs) {
  field <- function(name, default) {
    purrr::map(runs, function(r) r[[name]] %||% default)
  }
  tibble::tibble(
    t = as.character(unlist(field("t", ""))),
    bold = purrr::map_lgl(field("bold", FALSE), isTRUE),
    italic = purrr::map_lgl(field("italic", FALSE), isTRUE),
    underline = purrr::map_lgl(field("underline", FALSE), isTRUE),
    color = purrr::map_chr(field("color", NULL), bridge_colour),
    size = as.numeric(unlist(field("size", NA_real_))),
    font = as.character(unlist(field("font", ""))),
    family = as.character(unlist(field("family", "")))
  )
}

piece_text <- function(p) {
  rows <- purrr::map(p$rows %||% list(), function(r) text_runs(r$runs %||% list()))
  label <- paste(purrr::map_chr(rows, function(r) paste(r$t, collapse = "")),
                 collapse = "\n")
  box <- as.numeric(unlist(p$bbox))
  anchor <- as.character(p$anchor %||% "start")
  text <- list(rows = rows, align = anchor_align(anchor), placement = "free",
               insets = c(0, 0, 0, 0), box = box, anchor = anchor,
               html = isTRUE(p$html))
  new_piece(p, "text", box, rot = prim_number(p$rot), text = text, label = label)
}

# ---- folding a scope's pieces (3a paired outlines, stacked boxes; 3b) ---------

#' Apply the generic rules to the pieces of one scope (a unit, or everything
#' outside any unit), in paint order
#' @noRd
fold_pieces <- function(p) {
  merge_texts(pair_outlines(collapse_html_boxes(p)))
}

box_matrix <- function(p) cbind(p$x, p$y, p$w, p$h)

# Which rows of box matrix `b` agree with the box `one` within `tol` px.
boxes_agree <- function(b, one, tol = 0.5) {
  rowSums(abs(sweep(b, 2L, one)) <= tol) == 4L
}

# Mermaid paints an edge label's background on its div, span and p: one box.
collapse_html_boxes <- function(p) {
  n <- nrow(p)
  if (n < 2L) return(p)
  is_box <- p$html & p$kind == "shape"
  b <- box_matrix(p)
  cur <- seq_len(n)[-1L]
  prev <- cur - 1L
  same <- rowSums(abs(b[cur, , drop = FALSE] - b[prev, , drop = FALSE]) <= 0.5) == 4L
  joins <- c(FALSE, is_box[cur] & is_box[prev] & p$i[cur] == p$i[prev] + 1L & same)
  run <- cumsum(!joins)
  composite <- stats::ave(p$fill_a, run, FUN = function(a) 1 - prod(1 - a))
  p$fill_a <- ifelse(is_box, composite, p$fill_a)
  p[!joins, ]
}

same_geometry_kind <- function(prst, one) {
  ifelse(is.na(prst), "", prst) == ifelse(is.na(one), "", one)
}

# Mermaid 11 draws many boxes twice, a fill-only path and a stroke-only path
# over it. Each stroke-only shape takes the latest unclaimed fill-only shape
# before it with the same box and geometry kind; the pair becomes one shape
# with the outline's geometry, painted where the fill was.
pair_outlines <- function(p) {
  n <- nrow(p)
  is_shape <- p$kind == "shape"
  fill_only <- is_shape & !is.na(p$fill) & is.na(p$stroke)
  stroke_only <- is_shape & is.na(p$fill) & !is.na(p$stroke)
  if (!any(fill_only) || !any(stroke_only)) return(p)
  b <- box_matrix(p)
  claim <- function(partner, j) {
    cand <- which(fill_only & seq_len(n) < j & !seq_len(n) %in% partner)
    ok <- cand[boxes_agree(b[cand, , drop = FALSE], b[j, ]) &
                 same_geometry_kind(p$prst[cand], p$prst[[j]])]
    if (length(ok) > 0L) partner[[j]] <- max(ok)
    partner
  }
  partner <- purrr::reduce(which(stroke_only), claim, .init = rep(NA_integer_, n))
  outline <- which(!is.na(partner))
  if (length(outline) == 0L) return(p)
  filler <- partner[outline]
  p$fill[outline] <- p$fill[filler]
  p$fill_a[outline] <- p$fill_a[filler]
  p$pos[outline] <- p$pos[filler]
  p <- p[setdiff(seq_len(n), filler), ]
  p[order(p$pos), ]
}

# The index (into the shape rows given) of the shape a text merges into, or NA:
# the smallest shape painted before it, at its rotation, holding its centre
# and at least 80 % of its area; a tie goes to the one painted last.
text_container <- function(tb, trot, tpos, sb, srot, spos) {
  cx <- tb[[1]] + tb[[3]] / 2
  cy <- tb[[2]] + tb[[4]] / 2
  holds_centre <- sb[, 1] <= cx & cx <= sb[, 1] + sb[, 3] &
    sb[, 2] <= cy & cy <= sb[, 2] + sb[, 4]
  ow <- pmax(0, pmin(tb[[1]] + tb[[3]], sb[, 1] + sb[, 3]) - pmax(tb[[1]], sb[, 1]))
  oh <- pmax(0, pmin(tb[[2]] + tb[[4]], sb[, 2] + sb[, 4]) - pmax(tb[[2]], sb[, 2]))
  area <- tb[[3]] * tb[[4]]
  covers <- if (area > 0) ow * oh / area >= 0.8 else rep(TRUE, nrow(sb))
  ok <- which(holds_centre & covers & abs(srot - trot) < 1e-6 & spos < tpos)
  if (length(ok) == 0L) return(NA_integer_)
  s_area <- sb[ok, 3] * sb[ok, 4]
  best <- ok[s_area == min(s_area)]
  best[[which.max(spos[best])]]
}

#' Where a text goes inside the shape box it was merged into (3b)
#' @noRd
place_text <- function(text, sbox) {
  tb <- text$box
  off_x <- abs(tb[[1]] + tb[[3]] / 2 - (sbox[[1]] + sbox[[3]] / 2))
  off_y <- abs(tb[[2]] + tb[[4]] / 2 - (sbox[[2]] + sbox[[4]] / 2))
  if (off_x <= max(2, 0.1 * sbox[[3]]) && off_y <= max(2, 0.1 * sbox[[4]])) {
    text$placement <- "centre"
    text$align <- "center"
    text$insets <- c(0, 0, 0, 0)
  } else {
    text$placement <- "inset"
    text$align <- anchor_align(text$anchor)
    text$insets <- pmax(0, c(tb[[1]] - sbox[[1]], tb[[2]] - sbox[[2]],
                             sbox[[1]] + sbox[[3]] - (tb[[1]] + tb[[3]]),
                             sbox[[2]] + sbox[[4]] - (tb[[2]] + tb[[4]])))
  }
  text
}

#' Several texts as one: rows in order, the union of their boxes
#' @noRd
combine_texts <- function(texts) {
  if (length(texts) == 1L) return(texts[[1]])
  boxes <- do.call(rbind, purrr::map(texts, "box"))
  first <- texts[[1]]
  first$rows <- purrr::list_flatten(purrr::map(texts, "rows"))
  first$box <- union_box(boxes[, 1], boxes[, 2], boxes[, 3], boxes[, 4])
  first
}

# Word users expect to click a box and edit its label, so a text that sits in
# a shape becomes that shape's text body. Texts that would share a shape (a
# class box with its member list) all stay text boxes.
merge_texts <- function(p) {
  texts <- which(p$kind == "text")
  shapes <- which(p$kind == "shape" &
                    (!is.na(p$fill) | (p$closed %in% TRUE & !is.na(p$stroke))))
  if (length(texts) == 0L || length(shapes) == 0L) return(p)
  b <- box_matrix(p)
  sb <- b[shapes, , drop = FALSE]
  pick <- purrr::map_int(texts, function(t) {
    k <- text_container(b[t, ], p$rot[[t]], p$pos[[t]], sb, p$rot[shapes], p$pos[shapes])
    if (is.na(k)) NA_integer_ else shapes[[k]]
  })
  shared <- pick %in% pick[duplicated(pick) & !is.na(pick)]
  pick[shared] <- NA_integer_
  merged <- texts[!is.na(pick)]
  into <- pick[!is.na(pick)]
  p$text[into] <- purrr::map2(p$text[merged], into, function(text, s) place_text(text, b[s, ]))
  p$label[into] <- p$label[merged]
  p[setdiff(seq_len(nrow(p)), merged), ]
}

# ---- items -------------------------------------------------------------------

#' Box c(x, y, w, h) holding all the given boxes; NA when there are none
#' @noRd
union_box <- function(x, y, w, h) {
  keep <- !is.na(x) & !is.na(y) & !is.na(w) & !is.na(h)
  if (!any(keep)) return(rep(NA_real_, 4L))
  x0 <- min(x[keep])
  y0 <- min(y[keep])
  c(x0, y0, max(x[keep] + w[keep]) - x0, max(y[keep] + h[keep]) - y0)
}

first_class_token <- function(cls) {
  tok <- stringr::str_extract(cls, "[^\\s]+")
  ifelse(is.na(tok), "", tok)
}

#' Names of items that belong to no unit: mermaid:<kind>:<prim id>, else
#' :<first class token>, else just mermaid:<kind>
#' @noRd
generic_names <- function(p) {
  tag <- ifelse(nzchar(p$id), p$id, first_class_token(p$cls))
  paste0("mermaid:", p$kind, ifelse(nzchar(tag), paste0(":", tag), ""))
}

#' Pieces as items rows: generic names unless named already, item columns only
#' @noRd
to_items <- function(p) {
  p$name <- ifelse(is.na(p$name), generic_names(p), p$name)
  custom <- p$kind == "shape" & is.na(p$prst)
  p$path[!custom] <- list(NULL)
  p[, item_cols]
}

empty_items <- function() piece_ptype()[, item_cols]

group_item <- function(name, parent, box) {
  tibble::tibble(
    name = name, kind = "group", parent = as.character(parent),
    x = box[[1]], y = box[[2]], w = box[[3]], h = box[[4]], rot = 0,
    prst = NA_character_, fill = NA_character_, stroke = NA_character_,
    fill_a = 0, stroke_a = 0, stroke_w = 0, head = "none", tail = "none",
    label = "", flip_h = FALSE, flip_v = FALSE, adj = list(numeric()),
    path = list(NULL), closed = NA, dash = list(NULL), text = list(NULL)
  )
}

#' Put items at the top level of group `parent` (their own nesting is kept)
#' @noRd
reparent <- function(items, parent) {
  items$parent[is.na(items$parent)] <- as.character(parent)
  items
}

#' The box of drawn rows (not groups)
#' @noRd
drawn_box <- function(items) {
  d <- items$kind != "group"
  union_box(items$x[d], items$y[d], items$w[d], items$h[d])
}

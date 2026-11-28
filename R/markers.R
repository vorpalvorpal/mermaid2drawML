# R/markers.R
# Line ends. Mermaid draws arrow heads, crow's feet, hollow triangles and
# the sequence number circles as SVG markers; Word has five line ends of its
# own. This file decides, per marker id, which of the two a marker becomes
# (the id table), and does the arithmetic for both:
#   - a definition from `render$markers` is read once per id into a small
#     normalised structure: reference point, the scale from content to px,
#     the content box (the union of its prims' boxes), its orientation;
#   - a Word end is moved along the path by the signed tip offset, so that
#     it sits where Mermaid's marker was, and sized from the marker;
#   - a marker Word cannot draw (no end of that figure, or a round or
#     diamond end bigger than Word's largest) is instanced at the path end through ONE
#     affine matrix (translate the reference point to the origin, scale,
#     rotate to the tangent, translate to the end) and its prims become
#     ordinary scene pieces, named mermaid:marker:<id>.
# The pieces grouping with their edge is done where items are assembled
# (R/scene.R scope_blocks(), R/semantics.R unit_blocks_edge()).

# ---- the id table ------------------------------------------------------------

# Patterns on the marker id (the bridge has removed the render id prefix),
# tried in order, case-insensitively. `kind` is the Word end; `as_shape` is
# TRUE for markers Word has no end for, which are drawn as shapes when their
# definition is known. Without one, `kind` is the fallback: a hollow marker
# takes the nearest filled Word end. An id matching no row is a triangle,
# unless its definition is anything but one filled polygon (then shapes).
.line_end_table <- tibble::tribble(
  ~pattern,                                                  ~kind,      ~as_shape,
  "sequencenumber",                                          "none",     TRUE,
  "onlyone|zeroorone|oneormore|zeroormore|mdparent|mdchild", "none",     TRUE,
  "requirement_contains",                                    "none",     TRUE,
  "cross",                                                   "none",     TRUE,
  "extension",                                               "triangle", TRUE,
  "aggregation",                                             "diamond",  TRUE,
  "composition",                                             "diamond",  FALSE,
  "lollipop",                                                "oval",     TRUE,
  "circle",                                                  "oval",     FALSE,
  "dependency|stick|requirement_arrow",                      "arrow",    FALSE,
  "filled-head|barb",                                        "stealth",  FALSE,
  "point|arrowhead",                                         "triangle", FALSE
)

marker_row <- function(id) {
  purrr::detect_index(.line_end_table$pattern, function(pat) grepl(pat, tolower(id)))
}

#' What one marker end becomes: the Word end and whether it is drawn as
#' shapes. `spec` is its read definition, or NULL when there is none; with
#' one, `stroke_w` is the width of the line it ends.
#' @noRd
end_plan <- function(id, spec = NULL, stroke_w = 1) {
  hit <- marker_row(id)
  shape_ok <- !is.null(spec)
  kind <- if (hit == 0L) "triangle" else .line_end_table$kind[[hit]]
  as_shape <- if (hit == 0L) shape_ok && !spec$filled_polygon else
    shape_ok && (.line_end_table$as_shape[[hit]] || too_big_for_word(spec, kind, stroke_w))
  list(kind = if (as_shape) "none" else kind, as_shape = as_shape)
}

# Word's largest line end is about this many line widths long.
.word_end_max <- 5

#' TRUE for a round or diamond end, which Word centres on the line end, whose
#' marker is longer along the line than Word's largest end (H6): Mermaid's
#' 11 px flowchart circle on a 1 px line would shrink to a dot. Ends drawn
#' from their tip stay Word ends at any size.
#' @noRd
too_big_for_word <- function(spec, kind, stroke_w) {
  centred <- identical(unname(.end_anchor[kind]), "centre")
  centred && length(spec$prims) > 0L && !anyNA(spec$content) &&
    spec$content[[3]] * marker_scale(spec, stroke_w) > .word_end_max * max(stroke_w, 1)
}

#' The DrawingML line end for a marker id, without its definition
#'
#' @param marker_id Marker id(s) as the bridge reports them, or NULL.
#' @return "none" for no marker; otherwise the kind from the marker table
#'   (for a marker drawn as a shape, its fallback), "triangle" when no
#'   pattern matches.
#' @noRd
line_end_kind <- function(marker_id) {
  if (length(marker_id) == 0L) return("none")
  purrr::map_chr(as.character(unlist(marker_id)), function(id) {
    if (is.na(id) || !nzchar(id)) "none" else end_plan(id)$kind
  })
}

# ---- reading a definition --------------------------------------------------------

#' A marker definition (bridge A1) as what the arithmetic needs
#'
#' @return list(ref = c(refX, refY); fit, the content scale from the viewBox
#'   (1 without one: markerWidth and markerHeight are then only a clip box);
#'   per_stroke, TRUE for strokeWidth units; orient; content, the union of the
#'   prims' boxes c(x, y, w, h) in content units; prims; filled_polygon).
#' @noRd
read_marker <- function(def) {
  vb <- as.numeric(unlist(def$viewBox))
  w <- prim_number(def$w, 3)
  h <- prim_number(def$h, 3)
  fit <- if (length(vb) == 4L && all(vb[3:4] > 0)) min(w / vb[[3]], h / vb[[4]]) else 1
  prims <- def$prims %||% list()
  boxes <- purrr::map(prims, function(p) as.numeric(unlist(p$bbox)))
  boxes <- purrr::keep(boxes, function(b) length(b) == 4L)
  content <- if (length(boxes) == 0L) rep(NA_real_, 4L) else {
    b <- do.call(rbind, boxes)
    union_box(b[, 1], b[, 2], b[, 3], b[, 4])
  }
  one <- if (length(prims) == 1L) prims[[1]]
  list(
    ref = c(prim_number(def$refX), prim_number(def$refY)),
    fit = fit,
    # SVG's default markerUnits is strokeWidth.
    per_stroke = !identical(def$units, "userSpaceOnUse"),
    orient = def$orient %||% "auto",
    content = content,
    prims = prims,
    filled_polygon = !is.null(one) && one$kind %in% c("path", "poly") &&
      !is.null(one$fill) && prim_number(one$fill_a, 1) > 0
  )
}

# Content units to root px for a marker on a line `stroke_w` wide.
marker_scale <- function(spec, stroke_w) {
  spec$fit * (if (spec$per_stroke && stroke_w > 0) stroke_w else 1)
}

# ---- Word ends: offset and size ------------------------------------------------------

# Which point of the marker content Word's end is drawn from: the tip for the
# arrow kinds, the centre for the round and diamond ones.
.end_anchor <- c(triangle = "tip", stealth = "tip", arrow = "tip",
                 oval = "centre", diamond = "centre")

#' How far (px) the path end moves along the path, positive = beyond the end,
#' so that Word's end sits where Mermaid's marker was (spec C2)
#' @noRd
tip_offset <- function(spec, kind, at, scale) {
  anchor <- .end_anchor[kind]
  if (is.na(anchor) || anyNA(spec$content)) return(0)
  x0 <- spec$content[[1]]
  x1 <- x0 + spec$content[[3]]
  ref <- spec$ref[[1]]
  # Content that points along the path (a start with plain auto) reaches
  # back towards the node from its low x side.
  along <- at == "start" && identical(spec$orient, "auto")
  point <- if (anchor == "centre") (x0 + x1) / 2 else if (along) x0 else x1
  (if (along) ref - point else point - ref) * scale
}

# Word's sm / med / lg for a length relative to the stroke width; the 1e-9
# keeps the boundaries where the spec puts them despite rounding.
size_class <- function(px, stroke_w) {
  r <- px / max(stroke_w, 1)
  c("sm", "med", "lg")[[1L + (r >= 3.5 - 1e-9) + (r > 6 + 1e-9)]]
}

#' Word end size classes c(len, w) from the drawn marker content: its length
#' along the path and its width across it
#' @noRd
end_size <- function(spec, scale, stroke_w) {
  if (anyNA(spec$content)) return(c("med", "med"))
  c(size_class(spec$content[[3]] * scale, stroke_w),
    size_class(spec$content[[4]] * scale, stroke_w))
}

# ---- markers as shapes ------------------------------------------------------------

# Degrees clockwise the marker's x axis is turned to: a number is used as is;
# auto follows the path; auto-start-reverse turns round at the start.
marker_angle <- function(orient, dir, at) {
  fixed <- suppressWarnings(as.numeric(orient))
  if (length(fixed) == 1L && !is.na(fixed)) return(fixed)
  base <- if (anyNA(dir)) 0 else atan2(dir[[2]], dir[[1]]) * 180 / pi
  if (at == "start" && identical(orient, "auto-start-reverse")) base + 180 else base
}

#' Marker content to root px: the reference point to the origin, scaled,
#' rotated to the path, moved to the path end
#' @noRd
marker_matrix <- function(spec, scale, angle, point) {
  mat_chain(
    mat_translate(-spec$ref[[1]], -spec$ref[[2]]),
    mat_scaling(scale),
    mat_rotate(angle),
    mat_translate(point[[1]], point[[2]])
  )
}

# A context-fill / context-stroke paint takes the edge's stroke.
context_paint <- function(mp, field, piece) {
  if (!identical(mp[[field]], "context")) return(mp)
  stroked <- !is.na(piece$stroke)
  mp[field] <- list(if (stroked) piece$stroke)
  mp[[paste0(field, "_a")]] <- if (stroked) piece$stroke_a else 0
  mp
}

#' The marker's prims at one end of piece `piece` (a plain list) as scene
#' pieces, through matrix `m`, drawn with the piece's place in the DOM (its
#' scope and units)
#' @noRd
marker_shapes <- function(spec, id, m, scale, piece) {
  prims <- purrr::map(spec$prims, function(mp) {
    mp <- context_paint(context_paint(mp, "fill", piece), "stroke", piece)
    mp$m <- mat_then(prim_matrix(mp), m)
    mp$stroke_w <- prim_number(mp$stroke_w, 1) * scale
    if (length(mp$dash) > 0L) mp$dash <- as.numeric(unlist(mp$dash)) * scale
    mp$i <- piece$i
    mp$anc <- piece$anc
    mp
  })
  shapes <- purrr::list_rbind(purrr::map(prims, prim_piece))
  if (is.null(shapes)) return(NULL)
  shapes$name <- rep(paste0("mermaid:marker:", id), nrow(shapes))
  # Every part of a marker is a shape item (spec C3), a stroke too: it keeps
  # the line preset and flips, but is not a line of the diagram.
  shapes$kind[shapes$kind == "line"] <- "shape"
  shapes
}

# ---- applying markers to pieces -----------------------------------------------------
#
# The per-piece functions below take one piece as a plain list (one value
# per column, list columns unwrapped): assigning into one-row tibbles was
# most of the time this step took.

# A piece's stroke as a path: its own, or a line's two ends.
piece_track <- function(pc) {
  if (pc$kind != "line") return(pc$path)
  e <- line_ends(pc)
  path_tbl(c("M", "L"), x = e[1, c("x0", "x1")], y = e[1, c("y0", "y1")])
}

# Where a piece is drawn along a moved track: its box, flips (a line) and
# path (anything else).
track_geometry <- function(kind, track) {
  if (kind == "line") {
    x <- track$x
    y <- track$y
    return(list(box = c(min(x), min(y), abs(diff(x)), abs(diff(y))),
                flip_h = x[[2]] < x[[1]], flip_v = y[[2]] < y[[1]], path = NULL))
  }
  list(box = path_bbox(track), flip_h = FALSE, flip_v = FALSE, path = track)
}

#' One marked end of a piece: its Word end, size classes c(len, w), how far
#' the end moves, and the marker shapes (NULL for none)
#' @noRd
end_effect <- function(id, at, pc, track, specs) {
  none <- list(kind = pc[[if (at == "start") "head" else "tail"]],
               size = c("med", "med"), shift = c(0, 0), shapes = NULL)
  spec <- if (!is.na(id)) specs[[id]]
  if (is.null(spec)) return(none)
  plan <- end_plan(id, spec, pc$stroke_w)
  frame <- if (is.null(track) || nrow(track) == 0L) {
    list(point = c(NA_real_, NA_real_), dir = c(NA_real_, NA_real_))
  } else path_end_frame(track, at)
  scale <- marker_scale(spec, pc$stroke_w)
  if (plan$as_shape) {
    m <- marker_matrix(spec, scale, marker_angle(spec$orient, frame$dir, at), frame$point)
    shapes <- if (!anyNA(frame$point)) marker_shapes(spec, id, m, scale, pc)
    return(list(kind = "none", size = c("med", "med"), shift = c(0, 0), shapes = shapes))
  }
  off <- tip_offset(spec, plan$kind, at, scale)
  moves <- abs(off) > 0.25 && !anyNA(frame$dir)
  # At the start "beyond the end" is against the direction of travel.
  shift <- if (moves) off * frame$dir * (if (at == "start") -1 else 1) else c(0, 0)
  list(kind = plan$kind, size = end_size(spec, scale, pc$stroke_w), shift = shift,
       shapes = NULL)
}

#' Both ends of a marked piece: list(head, tail) end effects, `track` (the
#' moved stroke, NULL when it stays), `gone` (a stroke of no length, which
#' is drawn only as its marker shapes) and `shapes`
#' @noRd
mark_piece <- function(pc, specs) {
  track <- piece_track(pc)
  head <- end_effect(pc$marker_start, "start", pc, track, specs)
  tail <- end_effect(pc$marker_end, "end", pc, track, specs)
  moved <- any(head$shift != 0) || any(tail$shift != 0)
  gone <- zero_length(pc)
  parts <- purrr::compact(list(head$shapes, tail$shapes))
  shapes <- if (length(parts) > 0L) purrr::list_rbind(parts)
  if (!is.null(shapes)) {
    k <- nrow(shapes)
    # Right after the edge in paint order. A lone shape of a stroke that has
    # gone stands alone (the sequence number's circle); otherwise the shapes
    # are grouped with their edge.
    shapes$pos <- pc$pos + seq_len(k) / (k + 1)
    grouped <- k > 1L || !gone
    shapes$marker_of <- rep(if (grouped) pc$i else NA_integer_, k)
    shapes$edge_group <- rep(if (grouped) edge_group_name(pc) else NA_character_, k)
  }
  list(head = head, tail = tail, gone = gone, shapes = shapes,
       track = if (moved) path_move_end(path_move_end(track, "start", head$shift),
                                        "end", tail$shift))
}

#' For each piece, the row of the edge a grouped marker shape belongs to; NA
#' for other pieces and for the shapes of a stroke that has gone
#' @noRd
marker_host <- function(p) {
  edges <- which(is.na(p$marker_of))
  edges[match_key(p$marker_of, p$i[edges])]
}

# The group of a loose edge and its marker shapes is named after the edge
# item: mermaid:line:msg -> mermaid:edge_group:line:msg.
edge_group_name <- function(pc) {
  sub("^mermaid:", "mermaid:edge_group:", generic_names(pc))
}

# Rows `rows` of `pieces` redrawn along moved tracks.
set_tracks <- function(pieces, rows, tracks) {
  if (length(rows) == 0L) return(pieces)
  geo <- purrr::map2(pieces$kind[rows], tracks, track_geometry)
  box <- do.call(rbind, purrr::map(geo, "box"))
  pieces$x[rows] <- box[, 1]
  pieces$y[rows] <- box[, 2]
  pieces$w[rows] <- box[, 3]
  pieces$h[rows] <- box[, 4]
  pieces$flip_h[rows] <- purrr::map_lgl(geo, "flip_h")
  pieces$flip_v[rows] <- purrr::map_lgl(geo, "flip_v")
  pieces$path[rows] <- purrr::map(geo, "path")
  pieces
}

#' Pieces with their marker ends applied (spec C2, C3): every marked piece
#' moved and sized for its Word ends, followed in paint order by the shapes
#' of the markers Word cannot draw
#'
#' @param markers `render$markers`: definitions by marker id, or NULL.
#' @noRd
place_markers <- function(pieces, markers) {
  marked <- which(!no_marker(pieces))
  if (length(marked) == 0L) return(pieces)
  specs <- purrr::map(markers %||% list(), read_marker)
  fx <- purrr::map(marked, function(k) mark_piece(lapply(pieces, `[[`, k), specs))
  end_col <- function(end, field, k) purrr::map_chr(fx, function(f) f[[end]][[field]][[k]])
  pieces$head[marked] <- end_col("head", "kind", 1L)
  pieces$tail[marked] <- end_col("tail", "kind", 1L)
  pieces$head_len[marked] <- end_col("head", "size", 1L)
  pieces$head_w[marked] <- end_col("head", "size", 2L)
  pieces$tail_len[marked] <- end_col("tail", "size", 1L)
  pieces$tail_w[marked] <- end_col("tail", "size", 2L)
  moved <- !purrr::map_lgl(fx, function(f) is.null(f$track))
  pieces <- set_tracks(pieces, marked[moved], purrr::map(fx[moved], "track"))
  gone <- marked[purrr::map_lgl(fx, "gone")]
  kept <- if (length(gone) > 0L) pieces[-gone, ] else pieces
  out <- purrr::list_rbind(c(list(kept), purrr::map(fx, "shapes")), ptype = piece_ptype())
  out[order(out$pos), ]
}

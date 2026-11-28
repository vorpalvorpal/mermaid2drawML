# R/scene.R
# Bridge prims -> pieces -> folded pieces: the generic conversion every
# diagram type gets, and putting text inside shapes. How pieces become the
# nested, named items below is in R/assemble.R; the units the bridge reports
# and the flowchart shape registry are in R/semantics.R; line ends and
# marker shapes in R/markers.R. This file knows nothing about diagram types.
#
# ---- The items tibble: contract with the DrawingML emitter -------------------
#
# `scene$items` has one row per Word object. Rows are in document order: depth
# first, so a group row is immediately followed by all of its descendants, and
# siblings are in paint order (back to front). Lengths are root px (the root
# svg's user space, the units of `scene$viewbox`). Colours are "RRGGBB" with no
# "#", or NA for none.
#
# Documented columns:
#   name       chr  Word shape name (cNvPr name), always mermaid:<kind> or
#                   mermaid:<kind>:<tag>. Group names are unique (x, x#2,
#                   x#3, ... in document order); other names may repeat
#                   (generic "mermaid:text", two edges A-B). Units are named
#                   after the id the bridge gives them (mermaid:node:<id>,
#                   mermaid:subgraph:<id>, mermaid:edge:<start>-<end>,
#                   mermaid:edge_label:<start>:<end>); free items after their
#                   element (mermaid:<kind>:<id | first class>).
#                   Groups are units or scope groups: loose items drawn
#                   inside one DOM element (a legend entry, an actor figure,
#                   an icon) are grouped as mermaid:group:<element id, else its
#                   first class, else "g">, nested as the elements nest (not for an
#                   element holding only free texts, nor one spanning the
#                   whole drawing). A scope is the element a piece was drawn
#                   in (its innermost ancestor, by the bridge's DOM serial
#                   `n`); see "prims -> pieces".
#                   A marker Word has no line end for (crow's foot, hollow
#                   triangle or diamond, cross, sequence number circle) is
#                   drawn as shapes (kind "shape", its strokes too) named
#                   mermaid:marker:<marker id>, right after its edge in paint
#                   order (R/markers.R). The edge and
#                   its marker shapes are grouped as
#                   mermaid:edge_group:<start>-<end> (an edge unit), or
#                   mermaid:edge_group:<item name without "mermaid:"> (any
#                   other line or path, e.g. mermaid:edge_group:line:msg).
#                   A stroke of no length is only its marker shapes, grouped
#                   under that name when there are several; a single one
#                   stands alone and is folded like any shape, so that the
#                   sequence number's circle takes the number as its text.
#                   A box crossed by full-width horizontal lines of its own
#                   stroke (a class or requirement box, an ER entity) is
#                   several "rect" shapes of the box's name and paint,
#                   stacked top to bottom with no gap (compartments), each
#                   with its own text; the dividing lines are not items.
#                   An icon (a shape drawn inside a nested <svg>: an icon in
#                   a label, an icon node's glyph, an architecture service's
#                   icon) is mermaid:icon:<node or subgraph id>, else
#                   mermaid:icon:icon. It is never part of a unit's outline
#                   or paint and never joins another piece; its paths are
#                   grouped by the scope rule like loose items, inside the
#                   unit's group. A picture is mermaid:image:<id | first
#                   class>, placed like any drawn item.
#   kind       chr  "group" | "shape" | "line" | "text" | "image". A path
#                   whose box has no width or no height is a "line" (never
#                   filled). An "image" is a picture the browser could read
#                   (`png`); it has no paint or text.
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
#   fill_grad, stroke_grad
#              list NULL for flat paint; else the gradient that paints the
#                   fill / stroke (whose colour is then its first stop's):
#                   list(type "linear" | "radial", p1, p2 (linear) or c, r
#                   (radial) as root px, stops: tibble of o (0-1), color,
#                   alpha). It is kept in root space; the emitter expresses
#                   it in the item's own box, rotation and flips.
#   fill_pattern
#              chr  a DrawingML preset pattern (ltUpDiag, ltDnDiag, ltHorz,
#                   ltVert) drawn in the fill colour over nothing, or NA for a
#                   plain fill: Mermaid's hand-drawn hatching
#   href       chr  the web or mail link of the item's nearest <a> ancestor,
#                   NA for none
#   png        list the picture's PNG bytes (raw) for an "image", else NULL
#   stroke_w   dbl  stroke width, px (0 when there is no stroke)
#   head, tail chr  line end at the path START / END: "none", "triangle",
#                   "arrow", "stealth", "diamond", "oval". The path end is
#                   moved so that Word's end sits where Mermaid's marker was
#                   (an arrow's tip, a circle's or diamond's centre); an end
#                   drawn as marker shapes is "none". A circle or diamond
#                   longer than 5 line widths, more than Word's largest end,
#                   is drawn as marker shapes and the path end stays.
#   label      chr  the item's text, rows joined with "\n"; "" when none
#
# Internal columns:
#   head_len, head_w, tail_len, tail_w
#              chr  Word's size of each line end along and across the line,
#                   "sm" | "med" | "lg" (a:headEnd / a:tailEnd len and w),
#                   from the marker's drawn size over the stroke width;
#                   "med" when there is no end or no marker definition
#   flip_h, flip_v lgl   a:xfrm flipH / flipV. For a line (kind "line", preset
#                   "line", which runs from the box's top left to its bottom
#                   right) the flips put its start (head end) and end (tail
#                   end) where they were drawn: flip_h when the line runs right
#                   to left, flip_v when it runs bottom to top. A shape that
#                   holds text is never flipped vertically (Word turns its
#                   text upside down), so flip_v is FALSE for it; a
#                   horizontal flip leaves text alone.
#   adj        list named numeric vector of preset adjust values, DrawingML
#                   units, e.g. c(adj = 20000) -> <a:gd name="adj" fmla="val
#                   20000"/>; numeric(0) means none (the preset's defaults).
#   path       list custom geometry only (kind "shape" with prst NA): a tibble
#                   of ABSOLUTE segments in root px, columns cmd ("M", "L",
#                   "C", "Z"), x1, y1, x2, y2 (cubic control points, NA for
#                   M/L/Z), x, y (end point, NA for Z); see R/geometry.R.
#                   Subtract the item's x, y to get a:path coordinates; the
#                   path's own extent is the item's w, h. NULL otherwise.
#   closed     lgl  FALSE for lines and open custom paths; TRUE for presets and
#                   closed paths; NA for groups and free text. Only an item
#                   with closed FALSE, no fill and no text is written as a
#                   connector (wps:cNvCnPr); a filled open path is a shape.
#   dash       list numeric dash array in root px, NULL for a solid line
#   text       list NULL when the item has no text. Otherwise a list:
#                   rows      list of tibbles, one per paragraph (Mermaid's own
#                             text rows), one row per run, columns t (chr),
#                             bold, italic, underline (lgl), color (chr
#                             "RRGGBB"), alpha (dbl 0-1, the colour's and the
#                             element's opacity), size (dbl, px), font (chr, the face
#                             the browser used, may be ""), family (chr, the
#                             CSS font-family list)
#                   align     "left" | "center" | "right" (w:jc). Several
#                             texts merged into one shape take the relation
#                             their rows share (all centred, one left edge,
#                             one right edge); a one-row HTML label off
#                             centre takes the side it is nearer; otherwise
#                             the anchor gives it
#                   valign    "top" | "center" | "bottom": the body's vertical
#                             anchor (bodyPr anchor t / ctr / b); "center" for
#                             free text
#                   placement "centre": merged and centred on both axes
#                             (zero insets); "inset": merged, off centre on
#                             at least one axis; "free": a text box of its own
#                             (kind "text", zero insets). Each axis is decided
#                             separately: centred (zero insets on that axis),
#                             or one inset on the side `align` / `valign`
#                             names (a centred paragraph off centre has an
#                             inset on the side it leans away from)
#                   insets    numeric c(l, t, r, b), px, >= 0 (zeros unless
#                             placement is "inset"); at most one of l, r and
#                             one of t, b is non-zero
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
  drawn <- scene_pieces(render$prims, render$viewBox, render$markers)
  structure(
    list(type = render$type, viewbox = render$viewBox, title = render$title,
         desc = render$desc, unsupported = scene_unsupported(render$prims),
         items = scene_items(drawn$pieces, drawn$markers, render$db)),
    class = "mermaid_scene"
  )
}

# Tags of the prims that are reported but not drawn (<use>, ...). A picture
# is drawn when the browser could read it; one it could not, SVG <image> or
# HTML <img>, is reported as "image".
scene_unsupported <- function(prims) {
  unread_image <- function(p) identical(p$kind, "image") && is.null(png_data(p$png))
  skipped <- purrr::keep(prims, function(p) identical(p$kind, "other") || unread_image(p))
  unique(purrr::map_chr(skipped, function(p) if (p$kind == "image") "image" else p$tag %||% p$kind))
}

# ---- tolerances ----------------------------------------------------------------
#
# Every slack the rules allow, in px unless named _rel (a share of a length)
# or _deg. Each says what it protects against.
.tolerance <- list(
  # A box side shorter than this is no side: a path with no area is a line,
  # and a stroke this short has no length.
  flat_px = 0.01,
  # Two drawings of one box agree within this: the browser rounds each of
  # an HTML label's div, span and p on its own.
  box_px = 0.5,
  # Mermaid draws many shapes twice (a fill under an outline); the copies
  # agree within box_px or this share of the larger side, since a rough
  # outline strays further on a large shape.
  pair_rel = 0.01,
  # Rotations that are one rotation: matrix arithmetic leaves noise.
  angle_deg = 1e-6,
  # Stroke widths that are one width, for the same reason.
  width_eps = 1e-6,
  # The four lines of a frame meet at its corners within this (line caps,
  # rounding).
  frame_px = 1,
  # A divider reaches both sides of its box within this or divider_rel of
  # the box's width (rough strokes overshoot), and lies more than this inside
  # the box's top and bottom, so that the box's own edge is no divider.
  divider_px = 1,
  divider_rel = 0.01,
  # Dividers this close are one divider drawn twice (an ER entity's).
  cut_px = 1,
  # A text goes into a shape that covers at least this share of its box: a
  # label may overhang its node a little.
  text_cover = 0.8,
  # A text is centred on an axis within this share of the shape's side,
  # but at least centre_min_px (rounding) and at most centre_max_px (a large
  # shape must not swallow a deliberate offset).
  centre_rel = 0.03,
  centre_min_px = 1,
  centre_max_px = 4,
  # Rows of one label: neighbours overlap by at most this share of the
  # smaller height (line boxes overlap a little), gaps stay below
  # stack_gap_rel of the taller one (a larger gap is another label), and
  # rows sharing a left or right edge agree within stack_edge_px.
  stack_overlap_rel = 0.25,
  stack_gap_rel = 0.6,
  stack_edge_px = 1,
  # A path draws a preset when it strays from it by at most figure_rel of its
  # longer side, or for a thin shape up to figure_thin_px, capped at
  # figure_thin_rel of its shorter side (a 3 px diamond is no ellipse); see
  # figure_tolerance().
  figure_rel = 0.01,
  figure_thin_px = 0.5,
  figure_thin_rel = 0.05,
  # A Word line end is moved along its path only when Mermaid's marker sits
  # further than this from the path end; less cannot be seen.
  end_shift_px = 0.25,
  # The points of a slanted outline on its top or bottom edge lie within
  # this or inset_rel of the box height of that edge (rough outlines).
  inset_px = 1,
  inset_rel = 0.05
)

# ---- prims -> pieces ---------------------------------------------------------
#
# A piece is one drawable prim converted to root space: an items row plus the
# columns the folding and unit steps need: i (the prim's index), pos (paint
# position), src (the prim's kind), html, corner_r (corner radius of a rounded
# rect, or the first arc radius of a path, root px), id, cls, anc, data,
# marker_start and marker_end (the marker ids, NA for none), and
#   scope  int   the DOM serial `n` of the prim's innermost ancestor, NA for a
#                child of the root svg (and for hand-built prims without `n`).
#                Pieces of one scope were drawn inside one element: the rules
#                that fold pieces together (pairing, frames) look only within
#                a scope.
#   chain  list  the serials of all its ancestors, outermost first (those
#                that have one); scope groups nest along it.
#   unit   list  the prim's own unit from the bridge (an edge path), NULL
#                for most; the units of its ancestors are in `anc`.
#   inline lgl   TRUE for a shape drawn inside a nested <svg> (an icon).
# Icons and pictures are drawn apart (drawn_apart()): no folding rule looks
# at them.
# The shapes of markers Word has no line end for are not pieces: they are a
# side table beside them (R/markers.R), so that no rule below can fold them
# into other shapes, and they join their stroke when items are assembled.

# Every item column with the value an item has when nothing sets it: the one
# definition that pieces, group rows and registered shapes are all made from.
# List columns hold their value wrapped in a list.
.item_defaults <- list(
  name = NA_character_, kind = NA_character_, parent = NA_character_,
  x = NA_real_, y = NA_real_, w = NA_real_, h = NA_real_, rot = 0,
  prst = NA_character_, fill = NA_character_, stroke = NA_character_,
  fill_a = 0, stroke_a = 0, stroke_w = 0, head = "none", tail = "none",
  label = "", flip_h = FALSE, flip_v = FALSE, adj = list(numeric()),
  path = list(NULL), closed = NA, dash = list(NULL), text = list(NULL),
  head_len = "med", head_w = "med", tail_len = "med", tail_w = "med",
  fill_grad = list(NULL), stroke_grad = list(NULL), fill_pattern = NA_character_,
  href = NA_character_, png = list(NULL)
)

item_cols <- names(.item_defaults)

# The columns a piece has besides the item columns (see "prims -> pieces").
.piece_only_defaults <- list(
  i = NA_integer_, pos = NA_real_, src = NA_character_, html = FALSE,
  corner_r = NA_real_, id = "", cls = "", anc = list(list()), data = list(list()),
  marker_start = NA_character_, marker_end = NA_character_, scope = NA_integer_,
  chain = list(integer()), unit = list(NULL), inline = FALSE
)

.piece_defaults <- c(.item_defaults, .piece_only_defaults)

#' One row of `defaults` with the columns in `...` set (length 1 each, list
#' columns wrapped)
#' @noRd
row_from_defaults <- function(defaults, ...) {
  given <- list(...)
  defaults[names(given)] <- given
  tibble::new_tibble(defaults, nrow = 1L)
}

#' One items row: the defaults of `.item_defaults` with the columns given
#' @noRd
new_item <- function(...) row_from_defaults(.item_defaults, ...)

# Built once: the prototype of every rbind of pieces and of empty items.
.piece_proto <- vctrs::vec_ptype(tibble::new_tibble(.piece_defaults, nrow = 1L))

piece_ptype <- function() .piece_proto

#' Every drawable prim as a piece in root space, without those wholly outside
#' the viewBox, with its marker ends applied (R/markers.R)
#' @param markers `render$markers`, or NULL.
#' @return list(pieces, markers), as place_markers() gives them.
#' @noRd
scene_pieces <- function(prims, viewbox, markers = NULL) {
  drawable <- purrr::keep(prims, function(p) p$kind %in% names(.piece_converters))
  pieces <- purrr::list_rbind(purrr::map(drawable, prim_piece), ptype = piece_ptype())
  place_markers(pieces[box_in_view(pieces, viewbox) & !zero_stroke(pieces), ], markers)
}

zero_length <- function(p) {
  flat <- .tolerance$flat_px
  (p$kind == "line" | p$closed %in% FALSE) & p$w <= flat & p$h <= flat
}

# A stroke of no length draws nothing in Word; one with a marker is kept,
# because the marker is what Mermaid meant to show there.
zero_stroke <- function(p) zero_length(p) & no_marker(p)

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
                      adj = numeric(), text = NULL, label = "", png = NULL) {
  # One piece is made per prim, so this builds the row directly: tibble()'s
  # per-call checks were most of the time build_scene() took.
  row_from_defaults(
    .piece_defaults,
    i = as.integer(p$i %||% NA_integer_), pos = as.numeric(p$i %||% NA_real_),
    src = as.character(p$kind), kind = kind,
    x = box[[1]], y = box[[2]], w = box[[3]], h = box[[4]], rot = rot,
    prst = prst, fill = paint$fill, stroke = paint$stroke,
    fill_a = paint$fill_a, stroke_a = paint$stroke_a, stroke_w = paint$stroke_w,
    fill_grad = list(paint$fill_grad), stroke_grad = list(paint$stroke_grad),
    head = paint$head, tail = paint$tail, label = label,
    flip_h = flip_h, flip_v = flip_v, adj = list(adj), path = list(path),
    closed = closed, dash = list(paint$dash), text = list(text),
    href = anc_href(p$anc), png = list(png),
    html = isTRUE(p$html), corner_r = corner_r,
    id = as.character(p$id %||% ""), cls = as.character(p$cls %||% ""),
    anc = list(p$anc %||% list()), data = list(p$data %||% list()),
    marker_start = json_word(p$marker_start), marker_end = json_word(p$marker_end),
    scope = anc_scope(p$anc), chain = list(anc_serials(p$anc)), unit = list(p$unit),
    inline = isTRUE(p$inline)
  )
}

# Link targets Word may open from a click: web and mail links only (a
# `javascript:` link, a page of the diagram's own site, a data URI are not).
.link_schemes <- "^(https?|mailto):"

#' The href of a prim's nearest `<a>` ancestor, NA when there is none or its
#' target is not a web or mail link
#' @noRd
anc_href <- function(anc) {
  # Base iteration: this runs for every prim.
  hrefs <- vapply(anc %||% list(), function(a) json_chr(a$href), character(1))
  near <- utils::tail(hrefs[!is.na(hrefs)], 1L)
  if (length(near) == 1L && grepl(.link_schemes, near, ignore.case = TRUE)) near else NA_character_
}

#' A field of the bridge's JSON as one string: its first value, NA when it
#' is missing, empty or null
#' @noRd
json_chr <- function(x) {
  if (length(x) == 0L || isTRUE(is.na(x[[1]]))) NA_character_ else as.character(x[[1]])
}

# The same, with "" also NA: an id or paint the bridge reports as empty.
json_word <- function(x) {
  v <- json_chr(x)
  if (!is.na(v) && nzchar(v)) v else NA_character_
}

anc_serials <- function(anc) {
  as.integer(unlist(purrr::map(anc, function(a) a$n), use.names = FALSE))
}

# The innermost ancestor's serial; NA when it has none (a child of the root,
# or a hand-built prim).
anc_scope <- function(anc) {
  if (length(anc) == 0L) return(NA_integer_)
  as.integer(anc[[length(anc)]]$n %||% NA_integer_)
}

# "#rrggbb" from the bridge -> "RRGGBB"; a missing paint is NA.
bridge_colour <- function(x) {
  x <- json_word(x)
  if (is.na(x)) x else toupper(sub("^#", "", x))
}

prim_number <- function(x, default = 0) {
  if (is.null(x) || length(x) == 0L) default else as.numeric(x[[1]])
}

prim_matrix <- function(p) {
  m <- as.numeric(unlist(p$m))
  if (length(m) == 6L) m else c(1, 0, 0, 1, 0, 0)
}

no_paint <- function() {
  list(fill = NA_character_, fill_a = 0, fill_grad = NULL, stroke = NA_character_,
       stroke_a = 0, stroke_grad = NULL, stroke_w = 0, dash = NULL, head = "none",
       tail = "none")
}

prim_paint <- function(p) {
  fill <- bridge_colour(p$fill)
  stroke <- bridge_colour(p$stroke)
  stroked <- !is.na(stroke)
  list(
    fill = fill,
    fill_a = if (is.na(fill)) 0 else prim_number(p$fill_a, 1),
    fill_grad = if (!is.na(fill)) prim_gradient(p$fill_grad),
    stroke = stroke,
    stroke_a = if (stroked) prim_number(p$stroke_a, 1) else 0,
    stroke_grad = if (stroked) prim_gradient(p$stroke_grad),
    stroke_w = if (stroked) prim_number(p$stroke_w, 1) else 0,
    dash = if (stroked && length(p$dash) > 0L) as.numeric(unlist(p$dash)) else NULL,
    head = line_end_kind(p$marker_start),
    tail = line_end_kind(p$marker_end)
  )
}

#' The bridge's gradient as the scene keeps it, or NULL for none
#'
#' list(type "linear" | "radial", p1, p2 (linear) or c, r (radial) in root
#' px, stops: a tibble of o (0-1), color ("RRGGBB") and alpha). It stays in
#' root space: the emitter expresses it in the item's own box, which folding
#' may still change.
#' @noRd
prim_gradient <- function(g) {
  if (is.null(g) || !isTRUE(g$type %in% c("linear", "radial")) || length(g$stops) == 0L) {
    return(NULL)
  }
  point <- function(v) as.numeric(unlist(v))
  stops <- tibble::new_tibble(list(
    o = vapply(g$stops, function(s) prim_number(s$o), numeric(1)),
    color = vapply(g$stops, function(s) bridge_colour(s$c), character(1)),
    alpha = vapply(g$stops, function(s) prim_number(s$a, 1), numeric(1))
  ), nrow = length(g$stops))
  list(type = g$type, p1 = point(g$p1), p2 = point(g$p2), c = point(g$c),
       r = prim_number(g$r, NA_real_), stops = stops)
}

#' A gradient carried through matrix `m` (a marker's content to root space)
#' @noRd
gradient_through <- function(g, m) {
  if (is.null(g)) return(NULL)
  move <- function(pt) {
    if (length(pt) != 2L) return(pt)
    q <- mat_apply(m, pt[[1]], pt[[2]])
    c(q$x, q$y)
  }
  g$p1 <- move(g$p1)
  g$p2 <- move(g$p2)
  g$c <- move(g$c)
  g$r <- g$r * mat_scale(m)
  g
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
  # The bridge has already given a missing radius the other's value, and SVG
  # draws square corners when either one is zero.
  rx <- prim_number(g$rx)
  r <- rx * mat_scale(m)
  rounded <- rx > 0 && prim_number(g$ry) > 0
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
  line_piece(p, ends$x, ends$y, prim_paint(p))
}

# A straight line from (x[1], y[1]) to (x[2], y[2]).
line_piece <- function(p, x, y, paint) {
  # SVG never fills a line.
  paint$fill <- NA_character_
  paint$fill_a <- 0
  paint["fill_grad"] <- list(NULL)
  f <- segment_frame(x, y)
  new_piece(p, "line", f$box, prst = "line", paint = paint, closed = FALSE,
            flip_h = f$flip_h, flip_v = f$flip_v)
}

# Every subpath ended with a Z: the outline of a filled shape, so that a
# four-corner outline can be recognised as the rectangle SVG fills.
close_subpaths <- function(path) {
  if (is.null(path) || nrow(path) == 0L) return(path)
  open_end <- c(path$cmd[-1L] == "M", TRUE) & path$cmd != "Z"
  if (!any(open_end)) return(path)
  closers <- path_tbl(rep("Z", sum(open_end)))
  out <- purrr::list_rbind(list(path, closers))
  # Each Z goes right after the last row of its subpath.
  order_key <- c(seq_len(nrow(path)), which(open_end) + 0.5)
  out[order(order_key), ]
}

# Radius of the first arc in path data, in the path's own units, or NA.
first_arc_radius <- function(d) {
  if (is.null(d) || !nzchar(d)) return(NA_real_)
  hit <- stringr::str_match(d, paste0("[Aa][\\s,]*(", .num_re, ")"))[, 2]
  abs(as.numeric(hit))
}

#' A poly or path prim as a piece: a line when it draws no area, else a
#' shape, the preset it draws when there is one
#' @param arc_r The first arc radius of its path data, in its own units, or
#'   NA; taken as the corner radius when no preset gives one.
#' @noRd
piece_outline <- function(p, local_path, arc_r = NA_real_) {
  if (nrow(local_path) == 0L) return(NULL)
  m <- prim_matrix(p)
  path <- path_transform(local_path, m)
  paint <- prim_paint(p)
  bbox <- path_bbox(path)
  if (anyNA(bbox)) return(NULL)
  # A lone straight stroke is a line, whatever element drew it. So is a path
  # with no area, whatever its commands and reported fill: a fill needs an
  # area, and Mermaid draws its box dividers as filled straight cubics.
  flat <- min(bbox[[3]], bbox[[4]]) <= .tolerance$flat_px
  if (flat || (identical(path$cmd, c("M", "L")) && is.na(paint$fill))) {
    ends <- segment_ends(path, bbox)
    return(line_piece(p, ends$x, ends$y, paint))
  }
  closed <- path_closed(path)
  # Closing an open outline would add a stroked side, so only an unstroked
  # open outline may become a preset.
  figure <- if (closed || is.na(paint$stroke)) closed_figure(close_subpaths(path)) else
    no_figure(bbox)
  arc_r <- arc_r * mat_scale(m)
  new_piece(p, "shape", figure$box, prst = figure$prst, adj = figure$adj,
            paint = paint, path = path, closed = closed || !is.na(figure$prst),
            corner_r = if (!is.na(figure$corner_r)) figure$corner_r else
              if (isTRUE(arc_r > 0)) arc_r else NA_real_)
}

# The two ends of the segment a straight or flat path draws: its box, run
# from where the path starts towards where it ends.
segment_ends <- function(path, box) {
  ends <- path_ends(path)
  n <- length(ends$x)
  x <- c(box[[1]], box[[1]] + box[[3]])
  y <- c(box[[2]], box[[2]] + box[[4]])
  list(x = if (ends$x[[n]] < ends$x[[1]]) rev(x) else x,
       y = if (ends$y[[n]] < ends$y[[1]]) rev(y) else y)
}

no_figure <- function(box) {
  list(box = box, prst = NA_character_, adj = numeric(), corner_r = NA_real_)
}

#' The preset a closed path draws and its box: a four-corner outline is the
#' rect, then the figures of `.path_figures`; prst NA for none
#' @noRd
closed_figure <- function(path) {
  rect <- path_as_rect(path)
  if (!is.null(rect)) return(list(box = rect, prst = "rect", adj = numeric(), corner_r = NA_real_))
  box <- path_bbox(path)
  c(list(box = box), path_figure(path, box))
}

# Word presets that some closed paths draw exactly, tried in order; each
# returns the preset, its adjust and corner radius, or NULL. Pie sectors,
# home plates and rough (hand-drawn) outlines match none and stay custom.
.path_figures <- list(
  ellipse = function(path, box, tol) {
    if (path_on_ellipse(path, box, tol)) {
      list(prst = "ellipse", adj = numeric(), corner_r = NA_real_)
    }
  },
  round_rect = function(path, box, tol) {
    r <- path_corner_radius(path, box, tol)
    if (!is.na(r)) list(prst = "roundRect", adj = round_adj(r, box), corner_r = r)
  }
)

#' How far (px) a path may stray from a preset and still be drawn as it
#'
#' A hundredth of the longer side, or up to half a pixel for a thin shape.
#' The half pixel is capped at a twentieth of the shorter side so that a
#' figure a few pixels across (a 3 px diamond) is not taken for an ellipse.
#' @noRd
figure_tolerance <- function(box) {
  tol <- .tolerance
  max(tol$figure_rel * max(box[[3]], box[[4]]),
      min(tol$figure_thin_px, tol$figure_thin_rel * min(box[[3]], box[[4]])))
}

#' The preset a path draws exactly, or prst NA
#' @noRd
path_figure <- function(path, box) {
  found <- purrr::reduce(.path_figures, function(hit, f) hit %||% f(path, box, figure_tolerance(box)),
                         .init = NULL)
  found %||% list(prst = NA_character_, adj = numeric(), corner_r = NA_real_)
}

anchor_align <- function(anchor) {
  aligns <- c(start = "left", middle = "center", end = "right")
  if (length(anchor) == 1L && anchor %in% names(aligns)) aligns[[anchor]] else "left"
}

text_runs <- function(runs) {
  field <- function(name, default) {
    purrr::map(runs, function(r) r[[name]] %||% default)
  }
  tibble::new_tibble(list(
    t = as.character(unlist(field("t", ""))),
    bold = purrr::map_lgl(field("bold", FALSE), isTRUE),
    italic = purrr::map_lgl(field("italic", FALSE), isTRUE),
    underline = purrr::map_lgl(field("underline", FALSE), isTRUE),
    color = purrr::map_chr(field("color", NULL), bridge_colour),
    alpha = as.numeric(unlist(field("alpha", 1))),
    size = as.numeric(unlist(field("size", NA_real_))),
    font = as.character(unlist(field("font", ""))),
    family = as.character(unlist(field("family", "")))
  ), nrow = length(runs))
}

piece_text <- function(p) {
  rows <- purrr::map(p$rows %||% list(), function(r) text_runs(r$runs %||% list()))
  label <- paste(purrr::map_chr(rows, function(r) paste(r$t, collapse = "")),
                 collapse = "\n")
  box <- as.numeric(unlist(p$bbox))
  anchor <- as.character(p$anchor %||% "start")
  text <- list(rows = rows, align = anchor_align(anchor), valign = "center",
               placement = "free", insets = c(0, 0, 0, 0), box = box,
               anchor = anchor, html = isTRUE(p$html))
  new_piece(p, "text", box, rot = prim_number(p$rot), text = text, label = label)
}

#' The bytes of a `data:image/png;base64,...` URI, NULL for anything else
#' @noRd
png_data <- function(uri) {
  prefix <- "^data:image/png;base64,"
  if (!is.character(uri) || length(uri) != 1L || !grepl(prefix, uri)) return(NULL)
  jsonlite::base64_dec(sub(prefix, "", uri))
}

# A picture is placed like a rect: its own box through its matrix. Without
# the picture's bytes (the browser could not read it) there is nothing to
# draw, and scene_unsupported() reports it.
piece_image <- function(p) {
  bytes <- png_data(p$png)
  if (is.null(bytes)) return(NULL)
  g <- p$geom
  fb <- frame_box(prim_number(g$x), prim_number(g$y), prim_number(g$w),
                  prim_number(g$h), prim_matrix(p))
  new_piece(p, "image", fb$box, fb$rot, closed = TRUE, png = bytes)
}

# Prim kind -> piece; kinds not listed (<use>, ...) are not drawn.
.piece_converters <- list(
  rect = piece_rect, ellipse = piece_ellipse, line = piece_line,
  poly = function(p) piece_outline(p, poly_path(p$geom$points, isTRUE(p$geom$closed))),
  path = function(p) {
    d <- p$geom$d %||% ""
    piece_outline(p, path_parse(d), first_arc_radius(d))
  },
  text = piece_text, image = piece_image
)

#' One prim as a piece, or NULL for a kind that is not drawn
#' @noRd
prim_piece <- function(p) {
  convert <- .piece_converters[[p$kind]]
  if (is.null(convert)) NULL else convert(p)
}

# ---- folding pieces -------------------------------------------------------------

#' Apply the generic rules to the pieces of a unit, or to everything outside
#' any unit, in paint order
#' @noRd
fold_pieces <- function(p) {
  apart <- drawn_apart(p)
  folded <- p[!apart, ] |>
    collapse_html_boxes() |>
    frame_rects() |>
    pair_by_box() |>
    split_compartments() |>
    merge_texts()
  if (!any(apart)) return(folded)
  out <- purrr::list_rbind(list(folded, p[apart, ]))
  out[order(out$pos), ]
}

#' Pieces that no folding rule may touch: icons (shapes inside a nested
#' <svg>, which belong to a label or an icon, not to any outline) and
#' pictures. They are drawn as they are, beside the shapes they sit on.
#' @noRd
drawn_apart <- function(p) p$inline | p$kind == "image"

box_matrix <- function(p) cbind(p$x, p$y, p$w, p$h)

#' Which boxes of `a` are the box of `b` (the same row, or `b`'s only row)
#' on all four of x, y, w, h within `tol` (one value, or one per row)
#' @noRd
same_box <- function(a, b, tol = .tolerance$box_px) {
  a <- matrix(a, ncol = 4L)
  b <- matrix(b, ncol = 4L)
  if (nrow(b) == 1L) b <- b[rep(1L, nrow(a)), , drop = FALSE]
  rowSums(abs(a - b) <= tol) == 4L
}

# Scope as a key that also matches NA (the root) with NA.
scope_key <- function(scope) ifelse(is.na(scope), -1L, scope)

# An edge with a marker keeps its exact stroke, so no rule folds it.
no_marker <- function(p) is.na(p$marker_start) & is.na(p$marker_end)

# Mermaid paints an edge label's background on its div, span and p: one box.
collapse_html_boxes <- function(p) {
  n <- nrow(p)
  if (n < 2L) return(p)
  is_box <- p$html & p$kind == "shape"
  b <- box_matrix(p)
  cur <- seq_len(n)[-1L]
  prev <- cur - 1L
  same <- same_box(b[cur, , drop = FALSE], b[prev, , drop = FALSE])
  joins <- c(FALSE, is_box[cur] & is_box[prev] & p$i[cur] == p$i[prev] + 1L & same)
  run <- cumsum(!joins)
  composite <- stats::ave(p$fill_a, run, FUN = function(a) 1 - prod(1 - a))
  p$fill_a <- ifelse(is_box, composite, p$fill_a)
  p[!joins, ]
}

# -- four lines that draw a box -------------------------------------------------

# Where each line starts and ends (its box and flips), as a matrix x0 y0 x1 y1.
line_ends <- function(p) {
  x0 <- ifelse(p$flip_h, p$x + p$w, p$x)
  y0 <- ifelse(p$flip_v, p$y + p$h, p$y)
  cbind(x0 = x0, y0 = y0, x1 = ifelse(p$flip_h, p$x, p$x + p$w),
        y1 = ifelse(p$flip_v, p$y, p$y + p$h))
}

# TRUE when segment `e` (x0 y0 x1 y1) joins points a and b, in either order,
# within `tol` px.
joins_points <- function(e, a, b, tol = .tolerance$frame_px) {
  near <- function(p, q) max(abs(p - q)) <= tol
  (near(e[1:2], a) && near(e[3:4], b)) || (near(e[1:2], b) && near(e[3:4], a))
}

#' Frames among lines of one scope and paint: the row numbers (into `ends`)
#' of each set of four sides, top, bottom, left, right
#' @noRd
find_frames <- function(ends, tol = .tolerance$frame_px) {
  w <- abs(ends[, "x1"] - ends[, "x0"])
  h <- abs(ends[, "y1"] - ends[, "y0"])
  vertical <- which(w <= tol & h > tol)
  horizontal <- which(h <= tol & w > tol)
  if (length(vertical) < 2L || length(horizontal) < 2L) return(list())
  top_end <- function(k) if (ends[k, "y0"] <= ends[k, "y1"]) ends[k, 1:2] else ends[k, 3:4]
  bottom_end <- function(k) if (ends[k, "y0"] <= ends[k, "y1"]) ends[k, 3:4] else ends[k, 1:2]
  pairs <- utils::combn(vertical, 2L, simplify = FALSE)
  purrr::reduce(pairs, function(found, lr) {
    used <- unlist(found)
    if (any(lr %in% used)) return(found)
    tl <- top_end(lr[[1]])
    tr <- top_end(lr[[2]])
    bl <- bottom_end(lr[[1]])
    br <- bottom_end(lr[[2]])
    free <- setdiff(horizontal, used)
    side <- function(a, b) {
      hit <- purrr::detect(free, function(k) joins_points(ends[k, ], a, b, tol))
      hit %||% NA_integer_
    }
    top <- side(tl, tr)
    bottom <- side(bl, br)
    if (is.na(top) || is.na(bottom) || top == bottom) return(found)
    c(found, list(c(top, bottom, lr)))
  }, .init = list())
}

# Mermaid draws a sequence diagram's loop and alt frames as four lines; in
# Word they are one rectangle that moves as one, at the first line's place.
frame_rects <- function(p) {
  cand <- which(p$kind == "line" & no_marker(p))
  if (length(cand) < 4L) return(p)
  dash <- purrr::map_chr(p$dash[cand], function(d) paste(d, collapse = ","))
  key <- paste(scope_key(p$scope[cand]), p$stroke[cand], p$stroke_a[cand],
               p$stroke_w[cand], dash, sep = "\u001f")
  ends <- line_ends(p)
  frames <- purrr::list_flatten(purrr::map(split(cand, key), function(rows) {
    purrr::map(find_frames(ends[rows, , drop = FALSE]), function(k) rows[k])
  }))
  if (length(frames) == 0L) return(p)
  rects <- purrr::list_rbind(purrr::map(frames, function(rows) frame_rect(p[rows, ])))
  gone <- unlist(frames)
  out <- purrr::list_rbind(list(p[-gone, ], rects))
  out[order(out$pos), ]
}

frame_rect <- function(sides) {
  first <- sides[which.min(sides$pos), ]
  box <- union_box(box_matrix(sides))
  first$x <- box[[1]]
  first$y <- box[[2]]
  first$w <- box[[3]]
  first$h <- box[[4]]
  first$kind <- "shape"
  first$src <- "rect"
  first$prst <- "rect"
  first$closed <- TRUE
  first$flip_h <- FALSE
  first$flip_v <- FALSE
  first
}

# -- one box in one scope is one shape --------------------------------------------

# Mermaid 11 draws many boxes two or three times (a fill-only path under a
# stroke-only one, a straight path under a rough one). Each drawn piece joins
# the latest earlier piece of its scope at the same box and rotation; the
# boxes agree within max(0.5 px, 1 % of the larger side). A preset and a
# custom path must also draw the same figure (a pie's outer circle and its
# largest slice share a box, not a figure); two presets must be the same one.
box_leaders <- function(p, eligible) {
  b <- box_matrix(p)
  tol <- .tolerance
  # Pieces join only within their scope, so each scope is searched alone:
  # the work grows with the pieces of a scope, not with the whole diagram.
  join <- function(leader, j, rows) {
    cand <- rows[leader[rows] == rows & rows < j &
                   abs(p$rot[rows] - p$rot[[j]]) < tol$angle_deg]
    slack <- pmax(tol$box_px, tol$pair_rel * pmax(b[cand, 3], b[cand, 4], b[j, 3], b[j, 4]))
    agree <- same_box(b[cand, , drop = FALSE], b[j, ], slack)
    ok <- cand[agree]
    ok <- ok[purrr::map_lgl(ok, function(i) same_figure(p, i, j, slack[match(i, cand)]))]
    if (length(ok) > 0L) leader[[j]] <- max(ok)
    leader
  }
  by_scope <- split(which(eligible), scope_key(p$scope[eligible]))
  purrr::reduce(by_scope, function(leader, rows) {
    purrr::reduce(rows, join, rows = rows, .init = leader)
  }, .init = seq_len(nrow(p)))
}

# Whether pieces i and j (boxes already agreeing within `tol`) draw one figure.
same_figure <- function(p, i, j, tol) {
  pi <- p$prst[[i]]
  pj <- p$prst[[j]]
  if (is.na(pi) && is.na(pj)) return(TRUE)
  if (!is.na(pi) && !is.na(pj)) return(pi == pj)
  preset <- if (is.na(pi)) j else i
  custom <- if (is.na(pi)) i else j
  path_on_figure(p$path[[custom]], p$prst[[preset]],
                 c(p$x[[preset]], p$y[[preset]], p$w[[preset]], p$h[[preset]]),
                 tol, p$corner_r[[preset]])
}

#' Pieces drawing one box as one: geometry from the first preset (else the
#' first filled piece's path, else the first piece's, closed, and then the
#' preset that closed path draws, if any), fill from the first filled piece,
#' stroke from the last stroked one, name and place from the first
#' @noRd
fold_box <- function(pieces) {
  filled <- which(!is.na(pieces$fill))
  stroked <- which(!is.na(pieces$stroke))
  preset <- which(!is.na(pieces$prst))
  geo <- c(preset, filled, 1L)[[1]]
  out <- pieces[1, ]
  geometry <- c("x", "y", "w", "h", "rot", "prst", "adj", "path", "corner_r",
                "flip_h", "flip_v", "src")
  out[geometry] <- pieces[geo, geometry]
  if (is.na(out$prst)) out <- as_closed_figure(out)
  out$closed <- TRUE
  f <- if (length(filled) > 0L) filled[[1]] else NA_integer_
  k <- if (length(stroked) > 0L) max(stroked) else NA_integer_
  paint <- paint_from(pieces, f, k)
  out[names(paint)] <- paint
  out$href <- first_present(pieces$href)
  out
}

#' The paint columns, as list(column = value), of an item that takes its
#' fill from row `f` and its stroke from row `k` of `pieces` (NA: none). A
#' gradient is part of the paint it belongs to; a stroke's width and dash
#' come with it.
#' @noRd
paint_from <- function(pieces, f, k) {
  list(
    fill = pieces$fill[f], fill_a = if (is.na(f)) 0 else pieces$fill_a[[f]],
    fill_grad = list(if (!is.na(f)) pieces$fill_grad[[f]]),
    stroke = pieces$stroke[k], stroke_a = if (is.na(k)) 0 else pieces$stroke_a[[k]],
    stroke_w = if (is.na(k)) 0 else pieces$stroke_w[[k]],
    stroke_grad = list(if (!is.na(k)) pieces$stroke_grad[[k]]),
    dash = list(if (!is.na(k)) pieces$dash[[k]])
  )
}

# The first value that is not NA, else NA.
first_present <- function(x) {
  x <- x[!is.na(x)]
  if (length(x) == 0L) NA_character_ else x[[1]]
}

# A folded custom piece with its path closed, as the preset that closed path
# draws when there is one: Mermaid's class box is an open four-corner path
# under a rough outline, which together are one rectangle.
as_closed_figure <- function(piece) {
  path <- close_subpaths(piece$path[[1]])
  if (is.null(path) || nrow(path) == 0L) return(piece)
  piece$path <- list(path)
  figure <- closed_figure(path)
  if (is.na(figure$prst)) return(piece)
  piece[c("x", "y", "w", "h")] <- as.list(figure$box)
  piece$prst <- figure$prst
  piece$adj <- list(figure$adj)
  if (!is.na(figure$corner_r)) piece$corner_r <- figure$corner_r
  piece
}

pair_by_box <- function(p) {
  eligible <- p$kind == "shape" & no_marker(p)
  if (sum(eligible) < 2L) return(p)
  # Pieces join only a piece that joined nothing, so the leader is the root.
  leader <- box_leaders(p, eligible)
  if (all(leader == seq_len(nrow(p)))) return(p)
  folded <- purrr::map(split(seq_len(nrow(p)), leader), function(rows) {
    if (length(rows) == 1L) p[rows, ] else fold_box(p[rows, ])
  })
  out <- purrr::list_rbind(folded)
  out[order(out$pos), ]
}

# -- compartments ---------------------------------------------------------------------

# The element that holds a piece's own element: the scope's parent, or the
# scope itself at the top of the chain; NA for a child of the root.
scope_holder <- function(chain) {
  k <- length(chain)
  if (k == 0L) NA_integer_ else chain[[max(1L, k - 1L)]]
}

same_dash <- function(a, b) identical(as.numeric(a), as.numeric(b))

#' The rows (of `lines`) that divide shape `s`: horizontal, reaching both its
#' sides, inside its top and bottom, drawn with its stroke, under the element
#' that holds its own
#' @noRd
dividers_of <- function(p, s, lines) {
  t <- .tolerance
  tol <- max(t$divider_px, t$divider_rel * p$w[[s]])
  right <- p$x[[s]] + p$w[[s]]
  holder <- scope_holder(p$chain[[s]])
  ok <- abs(p$x[lines] - p$x[[s]]) <= tol &
    abs(p$x[lines] + p$w[lines] - right) <= tol &
    p$y[lines] > p$y[[s]] + t$divider_px & p$y[lines] < p$y[[s]] + p$h[[s]] - t$divider_px &
    p$stroke[lines] %in% p$stroke[[s]] &
    abs(p$stroke_w[lines] - p$stroke_w[[s]]) < t$width_eps &
    purrr::map_lgl(p$dash[lines], same_dash, p$dash[[s]]) &
    (is.na(holder) | purrr::map_lgl(p$chain[lines], function(ch) holder %in% ch))
  lines[ok]
}

# Where a box's compartments start and end, top to bottom: a divider drawn
# twice (an ER entity's) cuts once.
compartment_edges <- function(top, bottom, cuts) {
  cuts <- sort(cuts)
  cuts <- cuts[c(TRUE, diff(cuts) > .tolerance$cut_px)]
  c(top, cuts, bottom)
}

# A box crossed by full-width lines of its own stroke (a class box's title,
# attributes and methods) is drawn as stacked boxes, so that each takes its
# own texts and the box moves as Mermaid drew it; the shared borders draw
# the dividers, which go. Only an upright rect preset is split.
split_compartments <- function(p) {
  flat <- .tolerance$flat_px
  lines <- which(p$kind == "line" & no_marker(p) & p$h <= flat & p$w > flat)
  boxes <- which(p$kind == "shape" & p$prst %in% "rect" & abs(p$rot) < .tolerance$angle_deg &
                   no_marker(p))
  if (length(lines) == 0L || length(boxes) == 0L) return(p)
  cuts <- purrr::map(boxes, function(s) dividers_of(p, s, lines))
  split <- lengths(cuts) > 0L
  if (!any(split)) return(p)
  boxes <- boxes[split]
  cuts <- cuts[split]
  edges <- purrr::map2(boxes, cuts, function(s, d) {
    compartment_edges(p$y[[s]], p$y[[s]] + p$h[[s]], p$y[d])
  })
  copies <- rep(1L, nrow(p))
  copies[unique(unlist(cuts))] <- 0L
  copies[boxes] <- lengths(edges) - 1L
  out <- p[rep(seq_len(nrow(p)), copies), ]
  first <- cumsum(copies) - copies
  at <- unlist(purrr::map2(first[boxes], copies[boxes], function(f, k) f + seq_len(k)))
  out$y[at] <- unlist(purrr::map(edges, function(e) e[-length(e)]))
  out$h[at] <- unlist(purrr::map(edges, diff))
  out
}

# -- text into shapes ---------------------------------------------------------------

# Which shape boxes (rows of `sb`) hold text box `tb`: its centre, and at
# least the text_cover share of its area.
holds_text <- function(sb, tb) {
  cx <- tb[[1]] + tb[[3]] / 2
  cy <- tb[[2]] + tb[[4]] / 2
  holds_centre <- sb[, 1] <= cx & cx <= sb[, 1] + sb[, 3] &
    sb[, 2] <= cy & cy <= sb[, 2] + sb[, 4]
  ow <- pmax(0, pmin(tb[[1]] + tb[[3]], sb[, 1] + sb[, 3]) - pmax(tb[[1]], sb[, 1]))
  oh <- pmax(0, pmin(tb[[2]] + tb[[4]], sb[, 2] + sb[, 4]) - pmax(tb[[2]], sb[, 2]))
  area <- tb[[3]] * tb[[4]]
  covers <- if (area > 0) ow * oh / area >= .tolerance$text_cover else rep(TRUE, nrow(sb))
  holds_centre & covers
}

# The index (into the shape rows given) of the shape a text merges into, or NA:
# the smallest shape painted before it, at its rotation, that holds it; a tie
# goes to the one painted last.
text_container <- function(tb, trot, tpos, sb, srot, spos) {
  ok <- which(holds_text(sb, tb) & abs(srot - trot) < .tolerance$angle_deg & spos < tpos)
  if (length(ok) == 0L) return(NA_integer_)
  s_area <- sb[ok, 3] * sb[ok, 4]
  best <- ok[s_area == min(s_area)]
  best[[which.max(spos[best])]]
}

# Tolerance for calling a text centred along an axis of a shape this long.
centre_tol <- function(side) {
  t <- .tolerance
  max(t$centre_min_px, min(t$centre_max_px, t$centre_rel * side))
}

# Off centre across, the one inset that puts the text where it was drawn, by
# the paragraph alignment its anchor gives; a centred paragraph is moved by
# an inset on the side it leans away from.
.across_insets <- list(
  left = function(tb, sb, dx) c(tb[[1]] - sb[[1]], 0),
  right = function(tb, sb, dx) c(0, sb[[1]] + sb[[3]] - (tb[[1]] + tb[[3]])),
  center = function(tb, sb, dx) if (dx > 0) c(2 * dx, 0) else c(0, -2 * dx)
)

# Off centre down, the anchor and the one inset.
.down_insets <- list(
  top = function(tb, sb) c(tb[[2]] - sb[[2]], 0),
  bottom = function(tb, sb) c(0, sb[[2]] + sb[[4]] - (tb[[2]] + tb[[4]]))
)

#' Where a text goes inside the shape box it was merged into: centred, or
#' moved by one inset, on each axis separately
#'
#' @param across For a stack of texts, the relation that made it one
#'   ("center", "left" or "right"), which is then its alignment; NULL
#'   for a single text.
#' @noRd
place_text <- function(text, sbox, across = NULL) {
  tb <- text$box
  dx <- tb[[1]] + tb[[3]] / 2 - (sbox[[1]] + sbox[[3]] / 2)
  dy <- tb[[2]] + tb[[4]] / 2 - (sbox[[2]] + sbox[[4]] / 2)
  centred_y <- abs(dy) <= centre_tol(sbox[[4]])
  # A centred stack is centred whatever its union's offset; a centred
  # paragraph off centre keeps an inset.
  centred_x <- if (is.null(across)) abs(dx) <= centre_tol(sbox[[3]]) else across == "center"
  text$align <- across %||% single_align(text, dx, centred_x)
  text$valign <- if (centred_y) "center" else if (dy < 0) "top" else "bottom"
  lr <- if (centred_x) c(0, 0) else .across_insets[[text$align]](tb, sbox, dx)
  tb_ins <- if (centred_y) c(0, 0) else .down_insets[[text$valign]](tb, sbox)
  text$insets <- pmax(0, c(lr[[1]], tb_ins[[1]], lr[[2]], tb_ins[[2]]))
  text$placement <- if (centred_x && centred_y) "centre" else "inset"
  text
}

# A single text's alignment: centred when it is; else by its anchor, except
# that a one-row HTML label takes the side it is nearer, since the bridge's
# anchor for a shrink-wrapped div says nothing about the shape.
single_align <- function(text, dx, centred) {
  if (centred) return("center")
  if (isTRUE(text$html) && length(text$rows) == 1L) {
    return(if (dx < 0) "left" else "right")
  }
  anchor_align(text$anchor)
}

#' Several texts as one: rows in order, the union of their boxes
#' @noRd
combine_texts <- function(texts) {
  if (length(texts) == 1L) return(texts[[1]])
  boxes <- do.call(rbind, purrr::map(texts, "box"))
  first <- texts[[1]]
  first$rows <- purrr::list_flatten(purrr::map(texts, "rows"))
  first$box <- union_box(boxes)
  first
}

#' Whether text boxes (rows of a matrix, in y order) are rows of one label in
#' shape box `sbox` drawing preset `prst`, and how they line up: neighbours
#' overlap by little, gaps are small against the taller neighbour, the rows
#' share one horizontal relation to the shape, and the whole label is inside
#' the shape's figure
#'
#' @return "center", "left" or "right", the first relation that holds; "none"
#'   when the boxes are not a stack.
#' @noRd
stack_relation <- function(tb, sbox, prst = NA_character_) {
  k <- nrow(tb)
  overlap <- (tb[-k, 2] + tb[-k, 4]) - tb[-1, 2]
  t <- .tolerance
  rows_ok <- all(overlap <= t$stack_overlap_rel * pmin(tb[-k, 4], tb[-1, 4])) &&
    all(-overlap < t$stack_gap_rel * pmax(tb[-k, 4], tb[-1, 4]))
  if (!rows_ok || !stack_inside(union_box(tb), sbox, prst)) {
    return("none")
  }
  cx <- tb[, 1] + tb[, 3] / 2
  holds <- c(center = all(abs(cx - (sbox[[1]] + sbox[[3]] / 2)) <= centre_tol(sbox[[3]])),
             left = diff(range(tb[, 1])) <= t$stack_edge_px,
             right = diff(range(tb[, 1] + tb[, 3])) <= t$stack_edge_px)
  if (any(holds)) names(holds)[holds][[1]] else "none"
}

# A stack inside an ellipse's box may still be outside the ellipse (a legend
# beside a radar chart's outer ring); other figures take the single-text test.
stack_inside <- function(ub, sbox, prst) {
  if (identical(prst, "ellipse")) return(box_in_ellipse(ub, sbox))
  isTRUE(holds_text(matrix(sbox, 1L), ub))
}

# The shape each text picks: among shapes of its own scope first, else among
# all of them.
pick_containers <- function(p, texts, shapes) {
  b <- box_matrix(p)
  key <- scope_key(p$scope)
  purrr::map_int(texts, function(t) {
    pick_in <- function(cands) {
      if (length(cands) == 0L) return(NA_integer_)
      k <- text_container(b[t, ], p$rot[[t]], p$pos[[t]], b[cands, , drop = FALSE],
                          p$rot[cands], p$pos[cands])
      if (is.na(k)) NA_integer_ else cands[[k]]
    }
    own <- pick_in(shapes[key[shapes] == key[[t]]])
    if (is.na(own)) pick_in(shapes) else own
  })
}

# Word users expect to click a box and edit its label, so a text that sits in
# a shape becomes that shape's text body. Several texts in one shape merge
# when they are the rows of one label (a class compartment's members);
# otherwise they all stay text boxes.
merge_texts <- function(p) {
  texts <- which(p$kind == "text")
  shapes <- which(p$kind == "shape" &
                    (!is.na(p$fill) | (p$closed %in% TRUE & !is.na(p$stroke))))
  if (length(texts) == 0L || length(shapes) == 0L) return(p)
  pick <- pick_containers(p, texts, shapes)
  b <- box_matrix(p)
  by_shape <- split(texts[!is.na(pick)], pick[!is.na(pick)])
  by_shape <- purrr::map(by_shape, function(ts) ts[order(p$y[ts])])
  # A single text has no relation (NULL); a stack has the one that made it.
  across <- purrr::imap(by_shape, function(ts, s) {
    k <- as.integer(s)
    if (length(ts) > 1L) stack_relation(b[ts, , drop = FALSE], b[k, ], p$prst[[k]])
  })
  merging <- !purrr::map_lgl(across, identical, "none")
  into <- as.integer(names(by_shape)[merging])
  if (length(into) == 0L) return(p)
  merged <- by_shape[merging]
  p$text[into] <- purrr::pmap(list(merged, into, across[merging]), function(ts, s, rel) {
    place_text(combine_texts(p$text[ts]), b[s, ], rel)
  })
  p$label[into] <- purrr::map_chr(merged, function(ts) paste(p$label[ts], collapse = "\n"))
  # A linked text keeps its link in the shape that takes it.
  p$href[into] <- purrr::map2_chr(into, merged, function(s, ts) first_present(p$href[c(s, ts)]))
  p[setdiff(seq_len(nrow(p)), unlist(merged)), ]
}

#' Box c(x, y, w, h) holding all the boxes (rows of a matrix x, y, w, h)
#' that are known; NA when there are none
#' @noRd
union_box <- function(b) {
  b <- b[rowSums(is.na(b)) == 0L, , drop = FALSE]
  if (nrow(b) == 0L) return(rep(NA_real_, 4L))
  x0 <- min(b[, 1])
  y0 <- min(b[, 2])
  c(x0, y0, max(b[, 1] + b[, 3]) - x0, max(b[, 2] + b[, 4]) - y0)
}

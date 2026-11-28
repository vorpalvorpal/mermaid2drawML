# Test helpers for scenes: hand-built renders (mini_render() and prim_*())
# for exact arithmetic, and readers of the items tibble a scene holds.

# -- Reading items -----------------------------------------------------------------

# The columns every items tibble has (R/scene.R), and every line end kind.
scene_cols <- c("name", "kind", "parent", "x", "y", "w", "h", "rot", "prst",
                "fill", "stroke", "fill_a", "stroke_a", "stroke_w", "head",
                "tail", "label")
end_kinds <- c("none", "triangle", "arrow", "stealth", "diamond", "oval")

shapes_of <- function(items) items[items$kind == "shape", ]
groups_of <- function(items) items[items$kind == "group", ]
texts_of <- function(items) items[items$kind == "text", ]
lines_of <- function(items) items[items$kind == "line", ]
drawn <- function(items) items[items$kind != "group", ]

# Items directly inside group `name`, in order.
children_of <- function(items, name) items[!is.na(items$parent) & items$parent == name, ]

# Rows whose name starts with `prefix` (generic items may have no name).
named_with <- function(items, prefix) {
  items[!is.na(items$name) & startsWith(items$name, prefix), ]
}

# The one item called `name`; fails the test when there is not exactly one.
one_item <- function(items, name) {
  it <- item_named(items, name)
  testthat::expect_equal(nrow(it), 1L, info = name)
  it
}

# The marker shapes of marker `id`, and those of every marker.
markers_named <- function(items, id) item_named(items, paste0("mermaid:marker:", id))
marker_rows <- function(items) named_with(items, "mermaid:marker:")

# Which items hold a text body.
has_text <- function(items) !purrr::map_lgl(items$text, is.null)

# The names of a node, of an edge, and of the group of an edge and its
# marker shapes (edge_group_of(), as the package's own edge_group_name()
# takes an item name).
node_name <- function(id) paste0("mermaid:node:", id)
edge_name <- function(from, to) paste0("mermaid:edge:", from, "-", to)
edge_group_of <- function(from, to) paste0("mermaid:edge_group:", from, "-", to)

box_of <- function(it) c(it$x, it$y, it$w, it$h)

# The gap between two boxes c(x, y, w, h); 0 when they touch or overlap.
box_gap <- function(a, b) {
  dx <- max(0, a[1] - (b[1] + b[3]), b[1] - (a[1] + a[3]))
  dy <- max(0, a[2] - (b[2] + b[4]), b[2] - (a[2] + a[4]))
  sqrt(dx^2 + dy^2)
}

# The distance from a point to a box; 0 inside it.
point_gap <- function(p, b) box_gap(c(p[1], p[2], 0, 0), b)

# Where an item starts and ends: a line from its box and flips (R/scene.R),
# a custom shape from its path.
item_ends <- function(it) {
  if (it$kind == "line") {
    x <- c(it$x, it$x + it$w)
    y <- c(it$y, it$y + it$h)
    if (isTRUE(it$flip_h)) x <- rev(x)
    if (isTRUE(it$flip_v)) y <- rev(y)
    return(list(start = c(x[[1]], y[[1]]), end = c(x[[2]], y[[2]])))
  }
  p <- it$path[[1]]
  p <- p[p$cmd != "Z", ]
  list(start = c(p$x[[1]], p$y[[1]]), end = c(p$x[[nrow(p)]], p$y[[nrow(p)]]))
}

# The distinct points of a custom path (Z rows dropped), as a matrix x, y.
outline_points <- function(path) {
  p <- path[path$cmd != "Z", ]
  unique(round(cbind(x = p$x, y = p$y), 6))
}

item_named <- function(items, name) items[!is.na(items$name) & items$name == name, ]

# The names of group `name` and of every group nested in it.
subtree_groups <- function(items, name) {
  kids <- items$name[!is.na(items$parent) & items$parent == name &
                       items$kind == "group" & !is.na(items$name)]
  c(name, unlist(purrr::map(kids, function(k) subtree_groups(items, k))))
}

# The item called `name` and every item below it (following `parent`; the
# children need not have names of their own), in paint order.
subtree_rows <- function(items, name) {
  groups <- subtree_groups(items, name)
  items[(!is.na(items$name) & items$name == name) |
          (!is.na(items$parent) & items$parent %in% groups), ]
}

# -- Expectations --------------------------------------------------------------

# Every element of `object` within `tol` of `expected` (absolute, unlike the
# relative tolerance of expect_equal()).
expect_near <- function(object, expected, tol, info = NULL) {
  act <- as.numeric(unlist(object))
  exp <- as.numeric(unlist(expected))
  testthat::expect_equal(length(act), length(exp), info = info)
  diff <- if (length(act) == length(exp) && length(act) > 0L) max(abs(act - exp)) else Inf
  testthat::expect_lte(
    diff, tol,
    label = sprintf("largest difference between c(%s) and c(%s)",
                    paste(signif(act, 9), collapse = ", "),
                    paste(signif(exp, 9), collapse = ", ")))
}

# -- Hand-built renders --------------------------------------------------------

# A JSON object (`{}`) even when empty.
json_obj <- function(x = list()) {
  if (length(x) == 0L) structure(list(), names = character()) else x
}

# Map points through an SVG matrix c(a, b, c, d, e, f).
apply_m <- function(m, x, y) {
  list(x = m[1] * x + m[3] * y + m[5], y = m[2] * x + m[4] * y + m[6])
}

bbox_of_points <- function(m, x, y) {
  p <- apply_m(m, x, y)
  c(min(p$x), min(p$y), max(p$x) - min(p$x), max(p$y) - min(p$y))
}

# Every field the bridge reports for a shape prim (bridge header), with
# defaults: white fill, black 1 px stroke, identity matrix.
# `fill_grad` / `stroke_grad`: a linear_grad() or radial_grad(); the flat
# fill / stroke then default to the first stop's colour and alpha, with
# fill_approx / stroke_approx TRUE, as the bridge reports them. `inline`:
# TRUE for a shape drawn inside a nested <svg> (an icon).
prim_shape <- function(kind, tag, geom, bbox, m = c(1, 0, 0, 1, 0, 0),
                       fill = "#FFFFFF", fill_a = 1, stroke = "#000000",
                       stroke_a = 1, stroke_w = 1, dash = NULL,
                       marker_start = NULL, marker_end = NULL,
                       id = "", cls = "", data = list(), anc = list(),
                       html = NULL, n = NULL, fill_grad = NULL, stroke_grad = NULL,
                       inline = NULL) {
  if (!is.null(fill_grad) && missing(fill)) fill <- fill_grad$stops[[1]]$c
  if (!is.null(fill_grad) && missing(fill_a)) fill_a <- fill_grad$stops[[1]]$a
  if (!is.null(stroke_grad) && missing(stroke)) stroke <- stroke_grad$stops[[1]]$c
  if (!is.null(stroke_grad) && missing(stroke_a)) stroke_a <- stroke_grad$stops[[1]]$a
  p <- list(
    kind = kind, tag = tag, id = id, cls = cls, data = json_obj(data),
    anc = anc, bbox = as.list(bbox), m = as.list(m), geom = geom,
    fill = fill, fill_a = if (is.null(fill)) 0 else fill_a,
    fill_approx = !is.null(fill_grad),
    stroke = stroke, stroke_a = if (is.null(stroke)) 0 else stroke_a,
    stroke_approx = !is.null(stroke_grad),
    stroke_w = if (is.null(stroke)) 0 else stroke_w,
    dash = if (is.null(dash)) NULL else as.list(dash),
    marker_start = marker_start, marker_end = marker_end
  )
  if (!is.null(fill_grad)) p$fill_grad <- fill_grad
  if (!is.null(stroke_grad)) p$stroke_grad <- stroke_grad
  if (!is.null(inline)) p$inline <- inline
  if (!is.null(html)) p$html <- html
  if (!is.null(n)) p$n <- as.integer(n)
  p
}

# Gradients as the bridge reports them, points in root user space.
grad_stop <- function(o, c, a = 1) list(o = o, c = c, a = a)
linear_grad <- function(p1, p2, stops) {
  list(type = "linear", p1 = as.list(p1), p2 = as.list(p2), stops = stops)
}
radial_grad <- function(c, r, stops) list(type = "radial", c = as.list(c), r = r, stops = stops)

prim_rect <- function(x, y, w, h, rx = 0, ry = rx, m = c(1, 0, 0, 1, 0, 0),
                      bbox = NULL, tag = "rect", ...) {
  bbox <- bbox %||% bbox_of_points(m, c(x, x + w, x, x + w), c(y, y, y + h, y + h))
  prim_shape("rect", tag, list(x = x, y = y, w = w, h = h, rx = rx, ry = ry),
             bbox = bbox, m = m, ...)
}

prim_ellipse <- function(cx, cy, rx, ry = rx, m = c(1, 0, 0, 1, 0, 0),
                         bbox = NULL, ...) {
  bbox <- bbox %||% bbox_of_points(m, c(cx - rx, cx + rx), c(cy - ry, cy + ry))
  prim_shape("ellipse", "ellipse", list(cx = cx, cy = cy, rx = rx, ry = ry),
             bbox = bbox, m = m, ...)
}

prim_line <- function(x1, y1, x2, y2, m = c(1, 0, 0, 1, 0, 0), bbox = NULL,
                      fill = NULL, ...) {
  bbox <- bbox %||% bbox_of_points(m, c(x1, x2), c(y1, y2))
  prim_shape("line", "line", list(x1 = x1, y1 = y1, x2 = x2, y2 = y2),
             bbox = bbox, m = m, fill = fill, ...)
}

# `points`: a list of c(x, y) pairs.
prim_poly <- function(points, closed = TRUE, m = c(1, 0, 0, 1, 0, 0),
                      bbox = NULL, ...) {
  xs <- vapply(points, `[[`, numeric(1), 1)
  ys <- vapply(points, `[[`, numeric(1), 2)
  bbox <- bbox %||% bbox_of_points(m, xs, ys)
  prim_shape("poly", if (closed) "polygon" else "polyline",
             list(points = purrr::map(points, as.list), closed = closed),
             bbox = bbox, m = m, ...)
}

# `bbox` defaults to the extent of the path's coordinate pairs, which is only
# right for absolute paths made of straight segments; give it for curves.
prim_path <- function(d, bbox = NULL, m = c(1, 0, 0, 1, 0, 0), ...) {
  if (is.null(bbox)) {
    nums <- as.numeric(regmatches(d, gregexpr("-?[0-9.]+(e-?[0-9]+)?", d))[[1]])
    xs <- nums[c(TRUE, FALSE)]
    ys <- nums[c(FALSE, TRUE)]
    bbox <- bbox_of_points(m, xs, ys)
  }
  prim_shape("path", "path", list(d = d), bbox = bbox, m = m, ...)
}

# An image prim. SVG <image> (html FALSE): `geom` its own box x, y, w, h
# before `m` (as a rect prim's), `bbox` the root-space footprint of that box
# through `m`. HTML <img> (html TRUE): tag "img", identity m, geom and bbox
# x, y, w, h in root space. `png`: a
# "data:image/png;base64,..." URI, or NULL when the browser could not read
# the picture.
prim_image <- function(x, y, w, h, href = "https://example.com/x.png", anc = list(),
                       png = NULL, m = c(1, 0, 0, 1, 0, 0), html = FALSE,
                       id = "", cls = "", n = NULL) {
  if (html) m <- c(1, 0, 0, 1, 0, 0)
  p <- list(kind = "image", tag = if (html) "img" else "image", id = id, cls = cls,
            data = json_obj(), anc = anc,
            bbox = as.list(bbox_of_points(m, c(x, x + w, x, x + w), c(y, y, y + h, y + h))),
            geom = list(x = x, y = y, w = w, h = h), href = href, m = as.list(m), png = png)
  if (html) p$html <- TRUE
  if (!is.null(n)) p$n <- as.integer(n)
  p
}

# -- Pictures --------------------------------------------------------------------

# A 1 x 1 PNG (a valid file: signature, IHDR, IDAT and IEND with their CRCs).
png_1x1_uri <- function() {
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="
}

# The 48 x 48 PNG of the C4Context fixture's person image.
png_c4_uri <- function() {
  purrr::detect(fixture_render("C4Context")$prims, function(p) identical(p$kind, "image"))$png
}

# The bytes of a data URI.
png_bytes <- function(uri) jsonlite::base64_dec(sub("^data:image/png;base64,", "", uri))

png_signature <- as.raw(c(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a))

is_png <- function(bytes) length(bytes) > 8L && identical(bytes[1:8], png_signature)

prim_other <- function(tag, x, y, w, h, anc = list()) {
  list(kind = "other", tag = tag, id = "", cls = "", data = json_obj(),
       anc = anc, bbox = list(x, y, w, h))
}

# One text run as the bridge reports it.
text_run <- function(t, bold = FALSE, italic = FALSE, underline = FALSE,
                     color = "#000000", size = 16, font = "arial",
                     family = "arial, sans-serif", alpha = 1) {
  list(t = t, bold = bold, italic = italic, underline = underline,
       color = color, alpha = alpha, size = size, family = family, font = font)
}

# A text prim. `text`: one string per row (one run each, styled by `...`, which
# goes to text_run()); or give `rows` as a list of lists of text_run()s. The
# rows split the box height evenly.
prim_text <- function(text = NULL, x, y, w, h, anchor = "middle", rot = 0,
                      rows = NULL, html = TRUE, anc = list(), id = "", cls = "",
                      n = NULL, ...) {
  rows <- rows %||% purrr::map(text, function(t) list(text_run(t, ...)))
  nr <- length(rows)
  rh <- h / nr
  p <- list(
    kind = "text", tag = if (html) "foreignobject" else "text",
    id = id, cls = cls, data = json_obj(), anc = anc,
    bbox = list(x, y, w, h), html = html, anchor = anchor, rot = rot,
    rows = purrr::imap(rows, function(runs, i) {
      list(bbox = list(x, y + (i - 1) * rh, w, rh), runs = runs)
    })
  )
  if (!is.null(n)) p$n <- as.integer(n)
  p
}

# A bridge success object holding `prims`, through the validating constructor.
# `markers`: a named list of marker_def()s, keyed by marker id.
mini_render <- function(prims = list(), viewBox = c(0, 0, 200, 100), db = NULL,
                        type = "flowchart-v2", title = NULL, desc = NULL,
                        markers = NULL) {
  prims <- purrr::imap(unname(prims), function(p, i) c(list(i = i - 1L), p))
  raw <- list(
    ok = TRUE, mermaid = "11.13.0",
    type = type, role = type,
    viewBox = as.list(viewBox), title = title, desc = desc,
    svg = "<svg xmlns=\"http://www.w3.org/2000/svg\"/>",
    prims = prims, db = db
  )
  if (!is.null(markers)) raw$markers <- markers
  json <- jsonlite::toJSON(raw, auto_unbox = TRUE, null = "null", digits = NA)
  as_mermaid_render(jsonlite::fromJSON(json, simplifyVector = FALSE))
}

# A render of a diagram type without node/edge units, so that only the
# generic conversion of prims and the placing of texts in shapes apply.
generic_render <- function(prims = list(), viewBox = c(0, 0, 200, 100), ...) {
  mini_render(prims, viewBox = viewBox, type = "sequence", ...)
}

generic_items <- function(prims = list(), ...) build_scene(generic_render(prims, ...))$items

generic_doc <- function(prims = list(), ..., viewBox = c(0, 0, 200, 100)) {
  dml_doc(build_diagram_xml(generic_render(prims, viewBox = viewBox), ...)$xml)
}

# -- Flowchart-style units for hand-built renders ------------------------------
#
# The bridge says which element is which diagram unit: an `anc` entry (or,
# for an edge path, the prim itself) carries `unit = list(kind, id[, start,
# end])`, with `id` the unit's id in Mermaid's own data. These helpers set
# `unit` and nothing else that could mark a unit: the elements have neutral
# ids and classes (no `node`, `cluster`, `edgeLabel` or `edgePaths` class, no
# `flowchart-<id>-<n>` id, no `data-id` or `data-edge`), so a test built
# with them passes only when units are read from `unit`.

unit_node <- function(id) list(kind = "node", id = id)
unit_subgraph <- function(id) list(kind = "subgraph", id = id)
unit_edge <- function(id, start, end) list(kind = "edge", id = id, start = start, end = end)
unit_edge_label <- function(id, start, end) {
  list(kind = "edge_label", id = id, start = start, end = end)
}

# The db id of a node drawn as `flowchart-<id>-<n>`, the default unit id of
# the helpers that take a DOM id.
dom_node_id <- function(dom_id) sub("^flowchart-", "", sub("-\\d+$", "", dom_id))

# The two ends of a Mermaid edge id `L_<start>_<end>_<n>` (an end holding an
# underscore must be given explicitly).
edge_ends <- function(edge_id) {
  m <- regmatches(edge_id, regexec("^L_(.+)_(.+?)_\\d+$", edge_id))[[1]]
  if (length(m) == 3L) m[2:3] else c(NA_character_, NA_character_)
}

anc_root <- function() list(list(id = "", cls = "root"))
# `href`: the node is drawn inside an <a> with that href, as Mermaid's
# `click X href` draws it (the <a> just outside the node's element).
anc_node <- function(dom_id, cls = "default", id = dom_node_id(dom_id), href = NULL) {
  c(anc_root(), list(list(id = "", cls = "nodes")),
    if (!is.null(href)) list(anc_link(href)),
    list(list(id = "", cls = cls, unit = unit_node(id))))
}
# The `anc` entry of an <a> element: no id or class, its href.
anc_link <- function(href) list(id = "", cls = "", href = href)
anc_label <- function(anc) c(anc, list(list(id = "", cls = "label")))
anc_edge <- function() c(anc_root(), list(list(id = "", cls = "paths")))
anc_cluster <- function(id) {
  c(anc_root(), list(list(id = "", cls = "clusters"),
                     list(id = "", cls = "frame", data = list(look = "classic"),
                          unit = unit_subgraph(id))))
}
anc_cluster_label <- function(id) c(anc_cluster(id), list(list(id = "", cls = "cluster-label")))
anc_edge_label <- function(edge_id, start = edge_ends(edge_id)[[1]],
                           end = edge_ends(edge_id)[[2]]) {
  c(anc_root(), list(list(id = "", cls = "edgeLabels"),
                     list(id = "", cls = "label-box",
                          unit = unit_edge_label(edge_id, start, end)),
                     list(id = "", cls = "label")))
}

db_node <- function(id, shape = "squareRect", dom_id = paste0("flowchart-", id, "-0"),
                    label = id, parent = NULL, is_group = FALSE, look = "classic") {
  n <- list(id = id, label = label, shape = shape, isGroup = is_group, look = look)
  if (!is.null(dom_id)) n$domId <- dom_id
  if (!is.null(parent)) n$parentId <- parent
  n
}

db_edge <- function(start, end, id = paste0("L_", start, "_", end, "_0")) {
  list(id = id, start = start, end = end)
}

mini_db <- function(nodes = list(), edges = list()) list(nodes = nodes, edges = edges)

# An edge path prim as Mermaid draws it (open, stroked, arrow at the end); the
# bridge puts the edge's `unit` on the path itself.
prim_edge <- function(d, edge_id, marker_end = "flowchart-v2-pointEnd",
                      start = edge_ends(edge_id)[[1]], end = edge_ends(edge_id)[[2]],
                      ...) {
  p <- prim_path(d, anc = anc_edge(), fill = NULL, stroke = "#333333",
                 marker_end = marker_end, ...)
  p$unit <- unit_edge(edge_id, start, end)
  p
}

# -- DOM serials and marker definitions -------------------------------------------

# One `anc` entry with its DOM serial `n` (an unnamed <g> has id and cls "").
# `unit`: a unit_*() list when the element is a diagram unit.
# `href`: the element is an <a> with that href.
anc_g <- function(n, cls = "", id = "", data = NULL, unit = NULL, href = NULL) {
  a <- list(id = id, cls = cls, n = as.integer(n))
  if (!is.null(data)) a$data <- data
  if (!is.null(unit)) a$unit <- unit
  if (!is.null(href)) a$href <- href
  a
}

# Wrap prims in one <g> with serial `n`: its entry goes first in each prim's
# `anc` (outermost first), so in_g(in_g(ps, 3), 2) puts ps in g 3 inside g 2.
in_g <- function(prims, n, cls = "", id = "", data = NULL, unit = NULL, href = NULL) {
  g <- anc_g(n, cls = cls, id = id, data = data, unit = unit, href = href)
  purrr::map(prims, function(p) {
    p$anc <- c(list(g), p$anc %||% list())
    p
  })
}

# A marker's drawn child as the bridge reports it: shape fields only, `m`
# and `bbox` in marker content coordinates, no i, id, anc, data or markers.
marker_prim <- function(p) {
  p[setdiff(names(p), c("i", "id", "anc", "data", "marker_start", "marker_end",
                        "html", "n"))]
}

# One entry of `render$markers`. `prims` are prim_*() shapes in content
# coordinates; `viewBox` NULL for none.
marker_def <- function(viewBox = NULL, refX = 0, refY = 0, w = 3, h = 3,
                       units = "strokeWidth", orient = "auto", prims = list()) {
  list(viewBox = if (is.null(viewBox)) NULL else as.list(viewBox),
       refX = refX, refY = refY, w = w, h = h, units = units, orient = orient,
       prims = purrr::map(prims, marker_prim))
}

# Mermaid 11's own markers (from the survey SVGs), with their computed paint.

# flowchart-v2-pointEnd: a filled arrow point, tip at content x = 10.
marker_arrow_point <- function(w = 8, refX = 5, units = "userSpaceOnUse", orient = "auto") {
  marker_def(c(0, 0, 10, 10), refX = refX, refY = 5, w = w, h = w, units = units,
             orient = orient,
             prims = list(prim_path("M 0 0 L 10 5 L 0 10 z", fill = "#333333",
                                    stroke = "#333333", stroke_w = 1)))
}

# flowchart-v2-pointStart: the point turned round, orient auto.
marker_arrow_point_start <- function() {
  marker_def(c(0, 0, 10, 10), refX = 4.5, refY = 5, w = 8, h = 8,
             units = "userSpaceOnUse", orient = "auto",
             prims = list(prim_path("M 0 5 L 10 10 L 10 0 z", fill = "#333333",
                                    stroke = "#333333", stroke_w = 1)))
}

# flowchart-v2-circleEnd: a filled circle, refX beyond the content.
marker_circle_end <- function() {
  marker_def(c(0, 0, 10, 10), refX = 11, refY = 5, w = 11, h = 11,
             units = "userSpaceOnUse", orient = "auto",
             prims = list(prim_ellipse(5, 5, 5, fill = "#333333", stroke = "#333333")))
}

# flowchart-v2-crossEnd: two stroked strokes, no fill.
marker_cross_end <- function() {
  marker_def(c(0, 0, 11, 11), refX = 12, refY = 5.2, w = 11, h = 11,
             units = "userSpaceOnUse", orient = "auto",
             prims = list(prim_path("M 1,1 l 9,9 M 10,1 l -9,9", bbox = c(1, 1, 9, 9),
                                    fill = NULL, stroke = "#333333", stroke_w = 2)))
}

# class-extensionEnd: a hollow triangle, no viewBox, strokeWidth units;
# refX 1 (the base) sits on the path end and the tip points beyond it.
marker_extension_end <- function(orient = "auto", fill = NULL, stroke = "#333333") {
  marker_def(NULL, refX = 1, refY = 7, w = 20, h = 28, units = "strokeWidth",
             orient = orient,
             prims = list(prim_path("M 1,1 V 13 L18,7 Z", bbox = c(1, 1, 17, 12),
                                    fill = fill, stroke = stroke, stroke_w = 1)))
}

# class-extensionStart: the same triangle drawn turned round, refX 18, for a
# start with plain orient auto (190 x 240 is Mermaid's own clip box).
marker_extension_start <- function() {
  marker_def(NULL, refX = 18, refY = 7, w = 190, h = 240, units = "strokeWidth",
             orient = "auto",
             prims = list(prim_path("M 1,7 L18,13 V 1 Z", bbox = c(1, 1, 17, 12),
                                    fill = NULL, stroke = "#333333", stroke_w = 1)))
}

# class-compositionStart: a filled diamond, content x 1..18, refX 18.
marker_composition_start <- function() {
  marker_def(NULL, refX = 18, refY = 7, w = 190, h = 240, units = "strokeWidth",
             orient = "auto",
             prims = list(prim_path("M 18,7 L9,13 L1,7 L9,1 Z", bbox = c(1, 1, 17, 12),
                                    fill = "#333333", stroke = "#333333", stroke_w = 1)))
}

# sequencenumber: one filled circle, centred on (refX, refY).
marker_seq_number <- function() {
  marker_def(NULL, refX = 15, refY = 15, w = 60, h = 40, units = "strokeWidth",
             orient = "auto",
             prims = list(prim_ellipse(15, 15, 6, fill = "#333333", stroke = NULL)))
}

# er-oneOrMoreEnd: a crow's foot, one stroked open path (a bar and a lens).
marker_er_one_or_more_end <- function() {
  marker_def(NULL, refX = 27, refY = 18, w = 45, h = 36, units = "strokeWidth",
             orient = "auto",
             prims = list(prim_path("M3,9 L3,27 M9,18 Q27,0 45,18 Q27,36 9,18",
                                    bbox = c(3, 9, 42, 18), fill = NULL,
                                    stroke = "#333333", stroke_w = 1)))
}

# er-onlyOneStart: two bars, one stroked open path.
marker_er_only_one_start <- function() {
  marker_def(NULL, refX = 0, refY = 9, w = 18, h = 18, units = "strokeWidth",
             orient = "auto",
             prims = list(prim_path("M9,0 L9,18 M15,0 L15,18", fill = NULL,
                                    stroke = "#333333", stroke_w = 1)))
}

# A one-prim node unit (shape prim + centred label) for a db shape test.
# `id`: the node's db id (the unit id); by default the one in `dom_id`.
node_prims <- function(dom_id, shape_prim, label = "Label", id = dom_node_id(dom_id)) {
  box <- unlist(shape_prim$bbox)
  tw <- min(40, box[3] / 2)
  th <- min(16, box[4] / 2)
  shape_prim$anc <- anc_node(dom_id, id = id)
  list(shape_prim,
       prim_text(label, x = box[1] + (box[3] - tw) / 2, y = box[2] + (box[4] - th) / 2,
                 w = tw, h = th, anc = anc_label(anc_node(dom_id, id = id))))
}

# A render with one flowchart node `id` of db shape `shape`, drawn by `prim`.
one_node_render <- function(shape, prim, id = "X", label = "Label") {
  dom <- paste0("flowchart-", id, "-0")
  mini_render(node_prims(dom, prim, label = label, id = id),
              db = mini_db(list(db_node(id, shape = shape, dom_id = dom, label = label))))
}

# -- The hand-drawn look ---------------------------------------------------------

# The outline of a box as Mermaid draws its rough stroke: straight cubics.
rough_rect <- function(x, y, w, h, ...) {
  r <- x + w
  b <- y + h
  d <- sprintf(paste("M%s,%s C%s,%s %s,%s %s,%s C%s,%s %s,%s %s,%s",
                     "C%s,%s %s,%s %s,%s C%s,%s %s,%s %s,%s"),
               x, y, x + w / 3, y, x + 2 * w / 3, y, r, y,
               r, y + h / 3, r, y + 2 * h / 3, r, b,
               r - w / 3, b, r - 2 * w / 3, b, x, b,
               x, b - h / 3, x, b - 2 * h / 3, x, y)
  prim_path(d, bbox = c(x, y, w, h), ...)
}

# A hand-drawn fill as rough.js draws it: a stroke-only path of parallel
# straight strokes 4 px wide (each a cubic with its controls on the stroke,
# drawn twice as rough.js does) across the box x, y, w, h, 5 px apart, at
# `angle` degrees in SVG terms (clockwise from the x axis, y pointing down,
# so -45 runs up to the right). It opens, as Mermaid's do, with two strokes
# of no length at a corner.
hachure_path <- function(x, y, w, h, angle, stroke) {
  gap <- 5
  a <- angle * pi / 180
  d <- c(cos(a), sin(a))
  nrm <- c(-d[2], d[1])
  corners <- rbind(c(x, y), c(x + w, y), c(x, y + h), c(x + w, y + h))
  offs <- corners %*% nrm
  levels <- seq(min(offs) + gap / 2, max(offs) - gap / 2, by = gap)
  # Clip the line {p : p . nrm = k} to the box (slab method on t along d).
  clip <- function(k) {
    p0 <- k * nrm
    lim <- rbind(c(x, x + w), c(y, y + h))
    t <- c(-Inf, Inf)
    for (ax in 1:2) {
      if (abs(d[ax]) < 1e-12) {
        if (p0[ax] < lim[ax, 1] || p0[ax] > lim[ax, 2]) return(NULL)
      } else {
        tt <- sort((lim[ax, ] - p0[ax]) / d[ax])
        t <- c(max(t[1], tt[1]), min(t[2], tt[2]))
      }
    }
    if (t[2] - t[1] < 1) return(NULL)
    rbind(p0 + t[1] * d, p0 + t[2] * d)
  }
  segs <- purrr::compact(purrr::map(levels, clip))
  fmt <- function(v) format(round(v, 4), scientific = FALSE, trim = TRUE)
  stroke_d <- function(s) {
    a0 <- s[1, ]; a1 <- s[2, ]
    c1 <- a0 + (a1 - a0) / 3; c2 <- a0 + 2 * (a1 - a0) / 3
    sprintf("M%s %s C%s %s, %s %s, %s %s", fmt(a0[1]), fmt(a0[2]), fmt(c1[1]), fmt(c1[2]),
            fmt(c2[1]), fmt(c2[2]), fmt(a1[1]), fmt(a1[2]))
  }
  dot <- sprintf("M%s %s C%s %s, %s %s, %s %s", fmt(x), fmt(y), fmt(x), fmt(y), fmt(x),
                 fmt(y), fmt(x), fmt(y))
  parts <- c(dot, dot, unlist(purrr::map(segs, function(s) rep(stroke_d(s), 2L))))
  prim_path(paste(parts, collapse = " "), bbox = c(x, y, w, h), fill = NULL,
            stroke = stroke, stroke_w = 4)
}

# A hand-drawn flowchart node A in the box 20, 20, 100 x 60: the hachure
# fill and a rough outline in one node element, the label centred, and a db
# that says the look is handDrawn. `fill_prim`: the prim drawing the fill
# (by default the hachure). `outline_only = TRUE` leaves the fill out, as
# for a node styled fill:none.
handdrawn_node_render <- function(angle = -41, hatch = "#EEEEEE", outline = "#999999",
                                  fill_prim = NULL, look = "handDrawn",
                                  shape = "squareRect", outline_only = FALSE) {
  id <- "A"
  box <- c(20, 20, 100, 60)
  fill_prim <- fill_prim %||%
    hachure_path(box[1], box[2], box[3], box[4], angle = angle, stroke = hatch)
  dom <- paste0("flowchart-", id, "-0")
  anc <- anc_node(dom, id = id)
  fill_prim$anc <- anc
  # `outline_only`: an unfilled node, drawn as the outline path alone.
  outline_prim <- rough_rect(box[1], box[2], box[3], box[4], fill = NULL, stroke = outline,
                             stroke_w = 1.3, anc = anc)
  text <- prim_text("Label", x = box[1] + box[3] / 2 - 20, y = box[2] + box[4] / 2 - 8,
                    w = 40, h = 16, anc = anc_label(anc))
  prims <- if (isTRUE(outline_only)) list(outline_prim, text) else list(fill_prim, outline_prim, text)
  mini_render(prims, viewBox = c(0, 0, 200, 120),
              db = mini_db(list(db_node(id, shape = shape, dom_id = dom, look = look))))
}

# Mermaid's divider across a class, ER or requirement box: two straight
# cubic subpaths of zero height, filled and stroked as the box is.
mermaid_divider <- function(x0, x1, y, fill = "#ECECFF", stroke = "#9370DB", ...) {
  d <- sprintf("M%s,%s C%s,%s %s,%s %s,%s M%s,%s C%s,%s %s,%s %s,%s",
               x0, y, x0 + (x1 - x0) / 3, y, x0 + 2 * (x1 - x0) / 3, y, x1, y,
               x0, y, x0 + (x1 - x0) / 4, y, x0 + 3 * (x1 - x0) / 4, y, x1, y)
  prim_path(d, bbox = c(x0, y, x1 - x0, 0), fill = fill, stroke = stroke, ...)
}

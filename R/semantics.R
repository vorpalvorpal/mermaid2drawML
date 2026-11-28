# R/semantics.R
# The diagram units the bridge reports (nodes, subgraphs, edges and edge
# labels of the diagrams whose Mermaid data has them), what Mermaid's own
# node data says about them, and the flowchart shape registry: which node
# shapes are a Word preset, and how a registered node becomes one shape.
# The generic conversion these build on is in R/scene.R; how units and
# loose items nest and are named is in R/assemble.R; line ends are in
# R/markers.R.

# ---- the shape registry ------------------------------------------------------

# Flowchart shapes whose Word preset is the same figure, one row per figure
# with every name a node of it can carry in `db$nodes[[i]]$shape`: Mermaid's
# internal names and every `@{ shape: ... }` short name and alias (from the
# shape definitions in mermaid/dist). Shapes Word has no true preset for are
# left out, so they are drawn from their real outline. No figure is flipped
# vertically: Word would turn its text upside down.
.registry_figures <- tibble::tribble(
  ~prst,                         ~flip_h, ~adj,    ~adj_default, ~shapes,
  # rectangles
  "rect",                        FALSE,   "none",  NA,    c("squareRect", "rect", "proc", "process", "rectangle", "text",
                                                            "fork", "join", "forkJoin"),
  "roundRect",                   FALSE,   "round", NA,    c("roundedRect", "rounded", "event"),
  "roundRect",                   FALSE,   "round", 50000, c("stadium", "terminal", "pill"),
  "snip1Rect",                   TRUE,    "none",  NA,    c("notch-rect", "card", "notched-rectangle"),
  # circles and ellipses
  "ellipse",                     FALSE,   "none",  NA,    c("circle", "circ", "ellipse", "sm-circ", "start",
                                                            "small-circle", "stateStart", "f-circ", "junction",
                                                            "filled-circle"),
  "flowChartSummingJunction",    FALSE,   "none",  NA,    c("cross-circ", "summary", "crossed-circle"),
  # polygons; "choice" is the state diagram's choice pseudo-state
  "diamond",                     FALSE,   "none",  NA,    c("diamond", "choice", "diam", "decision", "question"),
  "hexagon",                     FALSE,   "hex",   NA,    c("hexagon", "hex", "prepare"),
  "parallelogram",               FALSE,   "slant", NA,    c("lean_right", "lean-r", "lean-right", "in-out"),
  "parallelogram",               TRUE,    "slant", NA,    c("lean_left", "lean-l", "lean-left", "out-in"),
  "trapezoid",                   FALSE,   "slant", NA,    c("trapezoid", "trap-b", "priority", "trapezoid-bottom"),
  "flowChartExtract",            FALSE,   "none",  NA,    c("tri", "extract", "triangle"),
  "flowChartMerge",              FALSE,   "none",  NA,    c("flip-tri", "manual-file", "flipped-triangle"),
  "flowChartCollate",            FALSE,   "none",  NA,    c("hourglass", "collate"),
  "flowChartManualInput",        FALSE,   "none",  NA,    c("sl-rect", "manual-input", "sloped-rectangle"),
  # flowchart symbols
  "flowChartPredefinedProcess",  FALSE,   "none",  NA,    c("subroutine", "fr-rect", "subprocess", "subproc",
                                                            "framed-rectangle"),
  "flowChartDocument",           FALSE,   "none",  NA,    c("doc", "document"),
  "flowChartMultidocument",      FALSE,   "none",  NA,    c("docs", "documents", "st-doc", "stacked-document"),
  "flowChartInternalStorage",    FALSE,   "none",  NA,    c("win-pane", "internal-storage", "window-pane"),
  "flowChartDelay",              FALSE,   "none",  NA,    c("delay", "half-rounded-rectangle"),
  "flowChartDisplay",            FALSE,   "none",  NA,    c("curv-trap", "curved-trapezoid", "display"),
  "flowChartOnlineStorage",      FALSE,   "none",  NA,    c("bow-rect", "stored-data", "bow-tie-rectangle"),
  "flowChartPunchedTape",        FALSE,   "none",  NA,    c("flag", "paper-tape"),
  # cylinders
  "can",                         FALSE,   "none",  NA,    c("cylinder", "cyl", "db", "database"),
  "flowChartMagneticDrum",       FALSE,   "none",  NA,    c("h-cyl", "das", "horizontal-cylinder"),
  "flowChartMagneticDisk",       FALSE,   "none",  NA,    c("lin-cyl", "disk", "lined-cylinder"),
  # pictures
  "cloud",                       FALSE,   "none",  NA,    "cloud",
  "lightningBolt",               FALSE,   "none",  NA,    c("bolt", "com-link", "lightning-bolt")
)

#' The shape registry: one row per shape name
#'
#' Columns: `shape`; `prst`, the Word preset; `flip_h`, `flip_v`; `adj`, how
#' the adjust value is measured from the drawing ("none"; "round": the corner
#' radius; "slant": the horizontal inset of a parallelogram or trapezoid;
#' "hex": the hexagon's inset); `adj_default`, the adjust value when the
#' drawing gives no measure (a stadium is fully rounded however it is drawn),
#' NA for the preset's own default.
#' @noRd
shape_registry <- function() {
  figs <- .registry_figures
  rows <- rep(seq_len(nrow(figs)), lengths(figs$shapes))
  tibble::tibble(shape = unlist(figs$shapes), prst = figs$prst[rows],
                 flip_h = figs$flip_h[rows], flip_v = FALSE, adj = figs$adj[rows],
                 adj_default = figs$adj_default[rows])
}

# ---- Mermaid's node data --------------------------------------------------------

db_field <- function(xs, name) purrr::map_chr(xs, function(x) json_chr(x[[name]]))

# What the db says about each node and group, by its id (a unit's id).
db_node_table <- function(db) {
  nodes <- db$nodes %||% list()
  quick_tibble(
    id = db_field(nodes, "id"), shape = db_field(nodes, "shape"),
    parent = db_field(nodes, "parentId"), look = db_field(nodes, "look"),
    is_group = purrr::map_lgl(nodes, function(n) isTRUE(n$isGroup))
  )
}

# ---- units -------------------------------------------------------------------
#
# The bridge says which element is which unit: an `anc` entry, or the prim
# itself for an edge path, carries `unit = {kind, id[, start, end]}`, with
# `id` the unit's id in Mermaid's own data. Nothing here guesses a unit from
# classes or DOM ids.

# The unit a piece belongs to: its own (an edge path), else its innermost
# ancestor's. `sig` tells apart neighbouring units, `hand_drawn` is what the
# unit's element says about the look.
piece_unit <- function(i, unit, anc) {
  if (is_unit(unit)) return(unit_fields(unit, sig = paste0("prim#", i), hand_drawn = FALSE))
  k <- purrr::detect_index(anc, function(a) is_unit(a$unit), .dir = "backward")
  if (k == 0L) return(unit_fields(NULL, sig = "", hand_drawn = FALSE))
  a <- anc[[k]]
  # Two elements may draw units of one id (an edge's two labels); the DOM
  # serial keeps them apart.
  sig <- paste(k, a$n %||% "", a$unit$kind, a$unit$id, sep = "\u001f")
  unit_fields(a$unit, sig = sig,
              hand_drawn = "rough-node" %in% class_tokens(a$cls) ||
                identical(a$data$look, "handDrawn"))
}

# The kinds of unit the scene assembles; another kind is drawn as loose items.
.unit_kinds <- c("node", "subgraph", "edge", "edge_label")

is_unit <- function(unit) isTRUE(unit$kind %in% .unit_kinds)

unit_fields <- function(unit, sig, hand_drawn) {
  list(kind = json_chr(unit[["kind"]]), id = json_chr(unit[["id"]]),
       start = json_chr(unit[["start"]]), end = json_chr(unit[["end"]]),
       sig = sig, hand_drawn = hand_drawn)
}

class_tokens <- function(cls) {
  tok <- strsplit(as.character(cls %||% ""), "\\s+")[[1]]
  tok[nzchar(tok)]
}

#' The unit of every piece, as a tibble aligned with `pieces`
#'
#' A unit is a run of consecutive pieces under the same unit element (the
#' DOM keeps an element's descendants together); a prim that is a unit is a
#' unit of its own.
#' @noRd
piece_units <- function(pieces) {
  info <- purrr::pmap(list(pieces$i, pieces$unit, pieces$anc), piece_unit)
  field <- function(name) purrr::map_chr(info, name)
  u <- quick_tibble(kind = field("kind"), id = field("id"), start = field("start"),
                    end = field("end"), sig = field("sig"),
                    hand_drawn = purrr::map_lgl(info, "hand_drawn"))
  n <- nrow(u)
  run <- cumsum(c(TRUE, u$sig[-1L] != u$sig[-n]))[seq_len(n)]
  u$key <- ifelse(is.na(u$kind), NA_character_, paste0("u", run))
  u
}

#' One row per unit: its kind, paint position, id, ends and name, and what
#' the db says about it (nodes and subgraphs, by their id)
#' @noRd
unit_table <- function(pu, pos, nodes) {
  keyed <- !is.na(pu$key)
  first <- !duplicated(pu$key) & keyed
  u <- pu[first, c("key", "kind", "id", "start", "end")]
  u$pos <- pos[first]
  hand_drawn <- tapply(pu$hand_drawn[keyed], pu$key[keyed], any)
  node_row <- match(u$id, nodes$id, incomparables = NA)
  node_row[!u$kind %in% c("node", "subgraph")] <- NA_integer_
  u$shape <- nodes$shape[node_row]
  u$parent <- nodes$parent[node_row]
  u$hand_drawn <- unname(hand_drawn[u$key]) | nodes$look[node_row] %in% "handDrawn"
  u$name <- mermaid_name(u$kind, unit_tag(u))
  u
}

# A unit's name tag: an edge is <start>-<end>, its label <start>:<end>, and
# anything else, or an edge whose ends are unknown, its id.
unit_tag <- function(u) {
  known <- !is.na(u$start) & !is.na(u$end)
  joint <- c(edge = "-", edge_label = ":")[u$kind]
  ifelse(known & !is.na(joint), paste0(u$start, joint, u$end), u$id)
}

# ---- registered nodes --------------------------------------------------------

# Points that outline a piece: its path's end points, or its box corners.
piece_points <- function(path, x, y, w, h) {
  if (is.null(path)) {
    return(quick_tibble(x = c(x, x + w, x + w, x), y = c(y, y, y + h, y + h)))
  }
  path[path$cmd != "Z", c("x", "y")]
}

# The horizontal inset of a slanted or hexagonal outline, as the preset's
# adjust: the largest gap between the box side and the outline along its top
# and bottom edges, over the shorter box side.
inset_adj <- function(pts, box) {
  tol <- max(.tolerance$inset_px, .tolerance$inset_rel * box[[4]])
  top <- pts$x[pts$y <= box[[2]] + tol]
  bottom <- pts$x[pts$y >= box[[2]] + box[[4]] - tol]
  if (length(top) == 0L || length(bottom) == 0L || min(box[[3]], box[[4]]) <= 0) {
    return(numeric())
  }
  right <- box[[1]] + box[[3]]
  inset <- max(0, min(top) - box[[1]], right - max(top),
               min(bottom) - box[[1]], right - max(bottom))
  c(adj = round(100000 * inset / min(box[[3]], box[[4]])))
}

# The inset adjust of a slanted or hexagonal node, from the points that
# outline its shape pieces.
outline_inset_adj <- function(shapes, box) {
  pts <- purrr::list_rbind(purrr::pmap(
    list(shapes$path, shapes$x, shapes$y, shapes$w, shapes$h), piece_points))
  inset_adj(pts, box)
}

# How each registry `adj` kind is measured from the node's shape pieces.
.adj_measures <- list(
  none = function(shapes, box) numeric(),
  round = function(shapes, box) {
    r <- shapes$corner_r[!is.na(shapes$corner_r)]
    if (length(r) == 0L) numeric() else round_adj(r[[1]], box)
  },
  slant = outline_inset_adj,
  hex = outline_inset_adj
)

#' A node with a registered preset as ONE shape: the box of all its drawn
#' prims, the fill of the largest filled one, the outline of the last stroked
#' one, and all of the unit's text
#' @noRd
registered_item <- function(pu, reg, hand_drawn, name) {
  # An icon in the label is no part of the node's outline or paint.
  shapes <- pu[outline_pieces(pu), ]
  box <- union_box(box_matrix(shapes))

  filled <- which(!is.na(shapes$fill))
  fill_from <- if (length(filled) > 0L) {
    filled[[which.max(shapes$w[filled] * shapes$h[filled])]]
  } else NA_integer_
  stroked <- which(!is.na(shapes$stroke))
  stroke_from <- if (length(stroked) > 0L) max(stroked) else NA_integer_
  # The hand-drawn look paints the fill as hachure strokes, drawn first.
  # Word has no hachure: the hatch's colour becomes the fill, drawn as the
  # line pattern nearest the strokes' direction.
  hachure <- is.na(fill_from) && hand_drawn && nrow(shapes) > 0L

  # A registered rect drawn with rounded corners (a state) keeps them.
  rounded_rect <- reg$prst == "rect" &&
    any(shapes$src == "rect" & !is.na(shapes$corner_r))
  prst <- if (rounded_rect) "roundRect" else reg$prst
  measure <- if (rounded_rect) "round" else reg$adj
  adj <- .adj_measures[[measure]](shapes, box)
  if (length(adj) == 0L && !is.na(reg$adj_default)) adj <- c(adj = reg$adj_default)

  texts <- pu[pu$kind == "text", ]
  text <- if (nrow(texts) > 0L) place_text(combine_texts(texts$text), box)

  paint <- paint_from(shapes, fill_from, stroke_from)
  if (hachure) paint[c("fill", "fill_a")] <- list(shapes$stroke[[1]], shapes$stroke_a[[1]])

  rlang::exec(
    new_item,
    name = name, kind = "shape",
    x = box[[1]], y = box[[2]], w = box[[3]], h = box[[4]],
    prst = prst, !!!paint,
    fill_pattern = if (hachure) hatch_pattern(shapes$path[[1]]) else NA_character_,
    label = paste(texts$label, collapse = "\n"),
    flip_h = reg$flip_h, flip_v = reg$flip_v, adj = list(adj),
    closed = TRUE,
    text = list(text),
    href = first_present(pu$href),
    depth = 0L
  )
}

# ---- hand-drawn hatching -------------------------------------------------------

# Word's light line patterns, by the direction of their lines in degrees
# clockwise from the x axis with y down (as SVG measures), modulo 180: an
# "upward" diagonal runs up to the right. rough.js hatches at -41 degrees
# by default, which is the upward diagonal.
.hatch_presets <- tibble::tibble(
  angle = c(0, 45, 90, 135),
  prst = c("ltHorz", "ltDnDiag", "ltVert", "ltUpDiag")
)

#' The preset pattern nearest the direction of a hatch's strokes, NA when
#' the strokes have no direction to measure
#' @noRd
hatch_pattern <- function(path) {
  ang <- stroke_direction(path)
  if (is.na(ang)) return(NA_character_)
  # Distance between line directions, which repeat every 180 degrees.
  gap <- abs(((.hatch_presets$angle - ang) + 90) %% 180 - 90)
  .hatch_presets$prst[[which.min(gap)]]
}

# Built once: the table is looked up for every node.
.shape_registry <- shape_registry()

registry_row <- function(shape) {
  reg <- .shape_registry
  hit <- match(shape, reg$shape, incomparables = NA)
  if (is.na(hit)) NULL else reg[hit, ]
}


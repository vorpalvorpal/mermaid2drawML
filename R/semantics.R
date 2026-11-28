# R/semantics.R
# The diagram units the bridge reports (nodes, subgraphs, edges and edge
# labels of the diagrams whose Mermaid data has them), what Mermaid's own
# node data says about them, and the flowchart shape registry: which node
# shapes are a Word preset, and how a registered node becomes one shape.
# The generic conversion these build on is in R/scene.R; how units and
# loose items nest and are named is in R/assemble.R; line ends are in
# R/markers.R.

# ---- the shape registry ------------------------------------------------------

#' Flowchart shapes whose Word preset is the same figure
#'
#' One row per name a node can carry in `db$nodes[[i]]$shape`: Mermaid's
#' internal names and every `@{ shape: ... }` short name and alias (from the
#' shape definitions in mermaid/dist). Shapes Word has no true preset for are
#' left out, so they are drawn from their real outline.
#'
#' Columns: `shape`; `prst`, the Word preset; `flip_h`, `flip_v`; `adj`, how
#' the adjust value is measured from the drawing ("none"; "round": the corner
#' radius; "slant": the horizontal inset of a parallelogram or trapezoid;
#' "hex": the hexagon's inset); `adj_default`, the adjust value when the
#' drawing gives no measure (a stadium is fully rounded however it is drawn),
#' NA for the preset's own default.
#' @noRd
shape_registry <- function() {
  tibble::tribble(
    ~shape,                   ~prst,                         ~flip_h, ~flip_v, ~adj,    ~adj_default,
    # rectangles
    "squareRect",             "rect",                        FALSE,   FALSE,   "none",  NA,
    "rect",                   "rect",                        FALSE,   FALSE,   "none",  NA,
    "proc",                   "rect",                        FALSE,   FALSE,   "none",  NA,
    "process",                "rect",                        FALSE,   FALSE,   "none",  NA,
    "rectangle",              "rect",                        FALSE,   FALSE,   "none",  NA,
    "text",                   "rect",                        FALSE,   FALSE,   "none",  NA,
    "fork",                   "rect",                        FALSE,   FALSE,   "none",  NA,
    "join",                   "rect",                        FALSE,   FALSE,   "none",  NA,
    "forkJoin",               "rect",                        FALSE,   FALSE,   "none",  NA,
    "roundedRect",            "roundRect",                   FALSE,   FALSE,   "round", NA,
    "rounded",                "roundRect",                   FALSE,   FALSE,   "round", NA,
    "event",                  "roundRect",                   FALSE,   FALSE,   "round", NA,
    "stadium",                "roundRect",                   FALSE,   FALSE,   "round", 50000,
    "terminal",               "roundRect",                   FALSE,   FALSE,   "round", 50000,
    "pill",                   "roundRect",                   FALSE,   FALSE,   "round", 50000,
    "notch-rect",             "snip1Rect",                   TRUE,    FALSE,   "none",  NA,
    "card",                   "snip1Rect",                   TRUE,    FALSE,   "none",  NA,
    "notched-rectangle",      "snip1Rect",                   TRUE,    FALSE,   "none",  NA,
    # circles and ellipses
    "circle",                 "ellipse",                     FALSE,   FALSE,   "none",  NA,
    "circ",                   "ellipse",                     FALSE,   FALSE,   "none",  NA,
    "ellipse",                "ellipse",                     FALSE,   FALSE,   "none",  NA,
    "sm-circ",                "ellipse",                     FALSE,   FALSE,   "none",  NA,
    "start",                  "ellipse",                     FALSE,   FALSE,   "none",  NA,
    "small-circle",           "ellipse",                     FALSE,   FALSE,   "none",  NA,
    "stateStart",             "ellipse",                     FALSE,   FALSE,   "none",  NA,
    "f-circ",                 "ellipse",                     FALSE,   FALSE,   "none",  NA,
    "junction",               "ellipse",                     FALSE,   FALSE,   "none",  NA,
    "filled-circle",          "ellipse",                     FALSE,   FALSE,   "none",  NA,
    "cross-circ",             "flowChartSummingJunction",    FALSE,   FALSE,   "none",  NA,
    "summary",                "flowChartSummingJunction",    FALSE,   FALSE,   "none",  NA,
    "crossed-circle",         "flowChartSummingJunction",    FALSE,   FALSE,   "none",  NA,
    # polygons
    "diamond",                "diamond",                     FALSE,   FALSE,   "none",  NA,
    # the state diagram's choice pseudo-state
    "choice",                 "diamond",                     FALSE,   FALSE,   "none",  NA,
    "diam",                   "diamond",                     FALSE,   FALSE,   "none",  NA,
    "decision",               "diamond",                     FALSE,   FALSE,   "none",  NA,
    "question",               "diamond",                     FALSE,   FALSE,   "none",  NA,
    "hexagon",                "hexagon",                     FALSE,   FALSE,   "hex",   NA,
    "hex",                    "hexagon",                     FALSE,   FALSE,   "hex",   NA,
    "prepare",                "hexagon",                     FALSE,   FALSE,   "hex",   NA,
    "lean_right",             "parallelogram",               FALSE,   FALSE,   "slant", NA,
    "lean-r",                 "parallelogram",               FALSE,   FALSE,   "slant", NA,
    "lean-right",             "parallelogram",               FALSE,   FALSE,   "slant", NA,
    "in-out",                 "parallelogram",               FALSE,   FALSE,   "slant", NA,
    "lean_left",              "parallelogram",               TRUE,    FALSE,   "slant", NA,
    "lean-l",                 "parallelogram",               TRUE,    FALSE,   "slant", NA,
    "lean-left",              "parallelogram",               TRUE,    FALSE,   "slant", NA,
    "out-in",                 "parallelogram",               TRUE,    FALSE,   "slant", NA,
    "trapezoid",              "trapezoid",                   FALSE,   FALSE,   "slant", NA,
    "trap-b",                 "trapezoid",                   FALSE,   FALSE,   "slant", NA,
    "priority",               "trapezoid",                   FALSE,   FALSE,   "slant", NA,
    "trapezoid-bottom",       "trapezoid",                   FALSE,   FALSE,   "slant", NA,
    "tri",                    "flowChartExtract",            FALSE,   FALSE,   "none",  NA,
    "extract",                "flowChartExtract",            FALSE,   FALSE,   "none",  NA,
    "triangle",               "flowChartExtract",            FALSE,   FALSE,   "none",  NA,
    "flip-tri",               "flowChartMerge",              FALSE,   FALSE,   "none",  NA,
    "manual-file",            "flowChartMerge",              FALSE,   FALSE,   "none",  NA,
    "flipped-triangle",       "flowChartMerge",              FALSE,   FALSE,   "none",  NA,
    "hourglass",              "flowChartCollate",            FALSE,   FALSE,   "none",  NA,
    "collate",                "flowChartCollate",            FALSE,   FALSE,   "none",  NA,
    "sl-rect",                "flowChartManualInput",        FALSE,   FALSE,   "none",  NA,
    "manual-input",           "flowChartManualInput",        FALSE,   FALSE,   "none",  NA,
    "sloped-rectangle",       "flowChartManualInput",        FALSE,   FALSE,   "none",  NA,
    # flowchart symbols
    "subroutine",             "flowChartPredefinedProcess",  FALSE,   FALSE,   "none",  NA,
    "fr-rect",                "flowChartPredefinedProcess",  FALSE,   FALSE,   "none",  NA,
    "subprocess",             "flowChartPredefinedProcess",  FALSE,   FALSE,   "none",  NA,
    "subproc",                "flowChartPredefinedProcess",  FALSE,   FALSE,   "none",  NA,
    "framed-rectangle",       "flowChartPredefinedProcess",  FALSE,   FALSE,   "none",  NA,
    "doc",                    "flowChartDocument",           FALSE,   FALSE,   "none",  NA,
    "document",               "flowChartDocument",           FALSE,   FALSE,   "none",  NA,
    "docs",                   "flowChartMultidocument",      FALSE,   FALSE,   "none",  NA,
    "documents",              "flowChartMultidocument",      FALSE,   FALSE,   "none",  NA,
    "st-doc",                 "flowChartMultidocument",      FALSE,   FALSE,   "none",  NA,
    "stacked-document",       "flowChartMultidocument",      FALSE,   FALSE,   "none",  NA,
    "win-pane",               "flowChartInternalStorage",    FALSE,   FALSE,   "none",  NA,
    "internal-storage",       "flowChartInternalStorage",    FALSE,   FALSE,   "none",  NA,
    "window-pane",            "flowChartInternalStorage",    FALSE,   FALSE,   "none",  NA,
    "delay",                  "flowChartDelay",              FALSE,   FALSE,   "none",  NA,
    "half-rounded-rectangle", "flowChartDelay",              FALSE,   FALSE,   "none",  NA,
    "curv-trap",              "flowChartDisplay",            FALSE,   FALSE,   "none",  NA,
    "curved-trapezoid",       "flowChartDisplay",            FALSE,   FALSE,   "none",  NA,
    "display",                "flowChartDisplay",            FALSE,   FALSE,   "none",  NA,
    "bow-rect",               "flowChartOnlineStorage",      FALSE,   FALSE,   "none",  NA,
    "stored-data",            "flowChartOnlineStorage",      FALSE,   FALSE,   "none",  NA,
    "bow-tie-rectangle",      "flowChartOnlineStorage",      FALSE,   FALSE,   "none",  NA,
    "flag",                   "flowChartPunchedTape",        FALSE,   FALSE,   "none",  NA,
    "paper-tape",             "flowChartPunchedTape",        FALSE,   FALSE,   "none",  NA,
    # cylinders
    "cylinder",               "can",                         FALSE,   FALSE,   "none",  NA,
    "cyl",                    "can",                         FALSE,   FALSE,   "none",  NA,
    "db",                     "can",                         FALSE,   FALSE,   "none",  NA,
    "database",               "can",                         FALSE,   FALSE,   "none",  NA,
    "h-cyl",                  "flowChartMagneticDrum",       FALSE,   FALSE,   "none",  NA,
    "das",                    "flowChartMagneticDrum",       FALSE,   FALSE,   "none",  NA,
    "horizontal-cylinder",    "flowChartMagneticDrum",       FALSE,   FALSE,   "none",  NA,
    "lin-cyl",                "flowChartMagneticDisk",       FALSE,   FALSE,   "none",  NA,
    "disk",                   "flowChartMagneticDisk",       FALSE,   FALSE,   "none",  NA,
    "lined-cylinder",         "flowChartMagneticDisk",       FALSE,   FALSE,   "none",  NA,
    # pictures
    "cloud",                  "cloud",                       FALSE,   FALSE,   "none",  NA,
    "bolt",                   "lightningBolt",               FALSE,   FALSE,   "none",  NA,
    "com-link",               "lightningBolt",               FALSE,   FALSE,   "none",  NA,
    "lightning-bolt",         "lightningBolt",               FALSE,   FALSE,   "none",  NA
  )
}

# ---- Mermaid's node data --------------------------------------------------------

db_field <- function(xs, name) {
  purrr::map_chr(xs, function(x) {
    v <- x[[name]]
    if (is.null(v) || length(v) == 0L) NA_character_ else as.character(v[[1]])
  })
}

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
  field <- function(name) {
    v <- unit[[name]]
    if (is.null(v) || length(v) == 0L || is.na(v[[1]])) NA_character_ else as.character(v[[1]])
  }
  list(kind = field("kind"), id = field("id"), start = field("start"), end = field("end"),
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

# How each registry `adj` kind is measured from the node's shape pieces.
.adj_measures <- list(
  none = function(shapes, box) numeric(),
  round = function(shapes, box) {
    r <- shapes$corner_r[!is.na(shapes$corner_r)]
    if (length(r) == 0L) numeric() else round_adj(r[[1]], box)
  },
  slant = function(shapes, box) {
    pts <- purrr::list_rbind(purrr::pmap(
      list(shapes$path, shapes$x, shapes$y, shapes$w, shapes$h), piece_points))
    inset_adj(pts, box)
  }
)
.adj_measures$hex <- .adj_measures$slant

#' A node with a registered preset as ONE shape: the box of all its drawn
#' prims, the fill of the largest filled one, the outline of the last stroked
#' one, and all of the unit's text
#' @noRd
registered_item <- function(pu, reg, hand_drawn, name) {
  shapes <- pu[pu$kind %in% c("shape", "line") & !pu$html, ]
  box <- union_box(shapes$x, shapes$y, shapes$w, shapes$h)

  filled <- which(!is.na(shapes$fill))
  fill_from <- if (length(filled) > 0L) {
    filled[[which.max(shapes$w[filled] * shapes$h[filled])]]
  } else NA_integer_
  stroked <- which(!is.na(shapes$stroke))
  stroke_from <- if (length(stroked) > 0L) max(stroked) else NA_integer_
  # The hand-drawn look paints the fill as hachure strokes, drawn first.
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

  quick_tibble(
    name = name, kind = "shape", parent = NA_character_,
    x = box[[1]], y = box[[2]], w = box[[3]], h = box[[4]], rot = 0,
    prst = prst,
    fill = if (hachure) shapes$stroke[[1]] else shapes$fill[fill_from],
    stroke = shapes$stroke[stroke_from],
    fill_a = if (hachure) shapes$stroke_a[[1]] else
      if (is.na(fill_from)) 0 else shapes$fill_a[[fill_from]],
    stroke_a = if (is.na(stroke_from)) 0 else shapes$stroke_a[[stroke_from]],
    stroke_w = if (is.na(stroke_from)) 0 else shapes$stroke_w[[stroke_from]],
    head = "none", tail = "none",
    label = paste(texts$label, collapse = "\n"),
    flip_h = reg$flip_h, flip_v = reg$flip_v, adj = list(adj),
    path = list(NULL), closed = TRUE,
    dash = list(if (is.na(stroke_from)) NULL else shapes$dash[[stroke_from]]),
    text = list(text),
    head_len = "med", head_w = "med", tail_len = "med", tail_w = "med", depth = 0L
  )
}

# Built once: the table is looked up for every node.
.shape_registry <- shape_registry()

registry_row <- function(shape) {
  reg <- .shape_registry
  hit <- match(shape, reg$shape, incomparables = NA)
  if (is.na(hit)) NULL else reg[hit, ]
}


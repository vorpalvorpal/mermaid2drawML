# R/semantics.R
# What the dagre-family diagrams (flowchart, class, state, ER, requirement,
# kanban, mindmap) mark in their DOM, and what that means for Word:
#   - units: nodes, subgraphs, edges and edge labels (spec 3c), joined to
#     Mermaid's own node/edge data (`db`) when the bridge has it;
#   - the flowchart shape registry (3d);
#   - the marker table that names line ends.
# The generic conversion these build on is in R/scene.R.

# ---- line ends ---------------------------------------------------------------

# Patterns on the marker id (the bridge has removed the render id prefix),
# tried in order; the first that matches names the end. Word has no X end, so
# a cross is a diamond, as it always was here.
.line_end_table <- tibble::tribble(
  ~pattern,                                                       ~kind,
  "sequencenumber",                                               "none",
  "onlyone|zeroorone|oneormore|zeroormore|mdparent|mdchild",      "none",
  "aggregation|composition",                                      "diamond",
  "cross",                                                        "diamond",
  "circle|lollipop",                                              "oval",
  "dependency|stick",                                             "arrow",
  "point|arrowhead|filled-head|barb|extension|requirement_arrow", "triangle"
)

#' The DrawingML line end for a marker id
#'
#' @param marker_id Marker id(s) as the bridge reports them, or NULL.
#' @return "none" for no marker; otherwise the kind from the marker table,
#'   "triangle" when no pattern matches.
#' @noRd
line_end_kind <- function(marker_id) {
  if (length(marker_id) == 0L) return("none")
  purrr::map_chr(as.character(unlist(marker_id)), function(id) {
    if (is.na(id) || !nzchar(id)) return("none")
    hit <- purrr::detect_index(.line_end_table$pattern,
                               function(pat) grepl(pat, tolower(id)))
    if (hit == 0L) "triangle" else .line_end_table$kind[[hit]]
  })
}

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
    "inv_trapezoid",          "trapezoid",                   FALSE,   TRUE,    "slant", NA,
    "trap-t",                 "trapezoid",                   FALSE,   TRUE,    "slant", NA,
    "manual",                 "trapezoid",                   FALSE,   TRUE,    "slant", NA,
    "trapezoid-top",          "trapezoid",                   FALSE,   TRUE,    "slant", NA,
    "inv-trapezoid",          "trapezoid",                   FALSE,   TRUE,    "slant", NA,
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

# ---- Mermaid's node and edge data ----------------------------------------------

db_field <- function(xs, name) {
  purrr::map_chr(xs, function(x) {
    v <- x[[name]]
    if (is.null(v) || length(v) == 0L) NA_character_ else as.character(v[[1]])
  })
}

db_node_table <- function(db) {
  nodes <- db$nodes %||% list()
  tibble::tibble(
    id = db_field(nodes, "id"), dom_id = db_field(nodes, "domId"),
    shape = db_field(nodes, "shape"), parent = db_field(nodes, "parentId"),
    look = db_field(nodes, "look"),
    is_group = purrr::map_lgl(nodes, function(n) isTRUE(n$isGroup))
  )
}

db_edge_table <- function(db) {
  edges <- db$edges %||% list()
  tibble::tibble(id = db_field(edges, "id"), start = db_field(edges, "start"),
                 end = db_field(edges, "end"))
}

#' Row of `table` whose `key` column matches each value, NA for none
#' @noRd
match_key <- function(values, keys) match(values, keys, incomparables = NA)

#' The db edge an edge id refers to; the ELK layout names its paths
#' `L_A_B_0_0` for db edge `L_A_B_0`.
#' @noRd
match_edge <- function(ref, edges) {
  hit <- match_key(ref, edges$id)
  ifelse(is.na(hit), match_key(sub("_\\d+$", "", ref), edges$id), hit)
}

# A unit's id when the db does not know it.
clean_unit_id <- function(dom_id) sub("^flowchart-", "", sub("-\\d+$", "", dom_id))

# ---- units -------------------------------------------------------------------

# Class tokens that make an ancestor a unit, tried in order; an ancestor whose
# id is a db node's domId is a node too (hand-drawn "rough-node", icon and
# image shapes carry no `node` class).
.unit_classes <- tibble::tribble(
  ~kind,        ~pattern,
  "subgraph",   "(^|-)cluster(-alt)?$",
  "edge_label", "^edgeLabel$",
  "node",       "^node$"
)

class_tokens <- function(cls) {
  tok <- strsplit(as.character(cls %||% ""), "\\s+")[[1]]
  tok[nzchar(tok)]
}

anc_unit_kind <- function(a, node_dom_ids) {
  tokens <- class_tokens(a$cls)
  hit <- purrr::detect_index(.unit_classes$pattern,
                             function(pat) any(grepl(pat, tokens)))
  if (hit > 0L) return(.unit_classes$kind[[hit]])
  id <- as.character(a$id %||% "")
  if (nzchar(id) && id %in% node_dom_ids) "node" else NA_character_
}

anc_signature <- function(a) {
  paste(a$id %||% "", a$cls %||% "", a$data$id %||% "", sep = "\u001f")
}

# A path is a link when it says so or sits under the edge paths group.
is_link <- function(src_path, data, anc) {
  src_path && (identical(data$edge, "true") ||
                 any(purrr::map_lgl(anc, function(a) "edgePaths" %in% class_tokens(a$cls))))
}

# What unit one piece belongs to: its kind, the unit ancestor's id, the
# signature that tells sibling units apart, the edge it refers to, and
# whether the ancestor says the look is hand-drawn.
piece_unit <- function(i, is_path, id, data, anc, node_dom_ids) {
  if (is_link(is_path, data, anc)) {
    ref <- as.character(data$id %||% id)
    return(list(kind = "edge", anc_id = NA_character_, sig = paste0("edge#", i),
                ref = if (nzchar(ref)) ref else NA_character_, hand_drawn = FALSE))
  }
  kinds <- purrr::map_chr(anc, anc_unit_kind, node_dom_ids)
  k <- if (any(!is.na(kinds))) max(which(!is.na(kinds))) else 0L
  if (k == 0L) {
    return(list(kind = NA_character_, anc_id = NA_character_, sig = "",
                ref = NA_character_, hand_drawn = FALSE))
  }
  a <- anc[[k]]
  below <- if (k < length(anc)) anc[[k + 1L]]$data$id else NULL
  ref <- as.character(a$data$id %||% below %||% NA_character_)
  # Sibling edge labels differ only in the edge id below their ancestor.
  sig <- paste(c(purrr::map_chr(anc[seq_len(k)], anc_signature), ref),
               collapse = "\u001e")
  list(kind = kinds[[k]], anc_id = as.character(a$id %||% ""), sig = sig, ref = ref,
       hand_drawn = "rough-node" %in% class_tokens(a$cls) ||
         identical(a$data$look, "handDrawn"))
}

#' The unit of every piece, as a tibble aligned with `pieces`
#'
#' A unit is a run of consecutive pieces under the same unit ancestor (the
#' DOM keeps a group's descendants together); every edge path is a unit of
#' its own.
#' @noRd
piece_units <- function(pieces, nodes) {
  dom_ids <- nodes$dom_id[!nodes$is_group & !is.na(nodes$dom_id)]
  info <- purrr::pmap(
    list(pieces$i, pieces$src == "path", pieces$id, pieces$data, pieces$anc),
    piece_unit, node_dom_ids = dom_ids
  )
  u <- tibble::tibble(
    kind = purrr::map_chr(info, "kind"), anc_id = purrr::map_chr(info, "anc_id"),
    sig = purrr::map_chr(info, "sig"), ref = purrr::map_chr(info, "ref"),
    hand_drawn = purrr::map_lgl(info, "hand_drawn")
  )
  n <- nrow(u)
  run <- cumsum(c(TRUE, u$sig[-1L] != u$sig[-n]))[seq_len(n)]
  u$key <- ifelse(is.na(u$kind), NA_character_, paste0("u", run))
  u
}

#' One row per unit: its kind, paint position, resolved id and name, and what
#' the db says about it
#'
#' Nodes and subgraphs join `db$nodes` on `domId`, then on `id` (groups,
#' and ER, kanban and requirement nodes, have no `domId` in the DOM); with no
#' match the id is the DOM id without Mermaid's `flowchart-` prefix and `-<n>`
#' suffix. Edges and edge labels join `db$edges` through their edge id.
#' @noRd
unit_table <- function(pu, pos, nodes, edges) {
  keyed <- !is.na(pu$key)
  first <- !duplicated(pu$key) & keyed
  u <- pu[first, c("key", "kind", "anc_id", "ref")]
  u$pos <- pos[first]
  u$hand_drawn <- purrr::map_lgl(u$key, function(k) any(pu$hand_drawn[pu$key %in% k]))

  on_dom <- match_key(u$anc_id, nodes$dom_id)
  node_row <- ifelse(is.na(on_dom), match_key(u$anc_id, nodes$id), on_dom)
  node_row[!u$kind %in% c("node", "subgraph")] <- NA_integer_
  edge_row <- match_edge(u$ref, edges)
  edge_row[!u$kind %in% c("edge", "edge_label")] <- NA_integer_

  u$id <- ifelse(is.na(node_row), clean_unit_id(u$anc_id), nodes$id[node_row])
  u$shape <- nodes$shape[node_row]
  u$parent <- nodes$parent[node_row]
  u$hand_drawn <- u$hand_drawn | nodes$look[node_row] %in% "handDrawn"
  u$start <- edges$start[edge_row]
  u$end <- edges$end[edge_row]
  u$name <- unit_names(u)
  u
}

# mermaid:node:<id>, mermaid:subgraph:<id>, mermaid:edge:<start>-<end>,
# mermaid:edge_label:<start>:<end>; an edge the db does not know is named
# after its edge id.
unit_names <- function(u) {
  known <- !is.na(u$start) & !is.na(u$end)
  edge_tag <- ifelse(known, paste0(u$start, "-", u$end), u$ref)
  label_tag <- ifelse(known, paste0(u$start, ":", u$end), u$ref)
  tag <- ifelse(u$kind == "edge", edge_tag,
                ifelse(u$kind == "edge_label", label_tag, u$id))
  prefix <- paste0("mermaid:", u$kind)
  ifelse(!is.na(tag) & nzchar(tag), paste0(prefix, ":", tag), prefix)
}

# ---- registered nodes --------------------------------------------------------

# Points that outline a piece: its path's end points, or its box corners.
piece_points <- function(path, x, y, w, h) {
  if (is.null(path)) {
    return(tibble::tibble(x = c(x, x + w, x + w, x), y = c(y, y, y + h, y + h)))
  }
  path[path$cmd != "Z", c("x", "y")]
}

# The horizontal inset of a slanted or hexagonal outline, as the preset's
# adjust: the largest gap between the box side and the outline along its top
# and bottom edges, over the shorter box side.
inset_adj <- function(pts, box) {
  tol <- max(1, 0.05 * box[[4]])
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

  tibble::tibble(
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
    text = list(text)
  )
}

registry_row <- function(shape) {
  reg <- shape_registry()
  hit <- match_key(shape, reg$shape)
  if (is.na(hit)) NULL else reg[hit, ]
}

# ---- blocks: units and loose items, before nesting -----------------------------
#
# A block is what gets placed at one level of the group tree: a unit's items,
# one loose item, or a subgraph (whose `items` are its background and whose
# members are the blocks with member_of == its id).

new_blocks <- function(kind, pos, items, gid = NA_character_,
                       member_of = NA_character_, start = NA_character_,
                       end = NA_character_, parent = NA_character_) {
  tibble::tibble(kind = kind, pos = pos, items = items, gid = gid,
                 member_of = member_of, start = start, end = end, parent = parent)
}

#' Items of a unit under one name: a single item takes it, several are
#' grouped under it
#' @noRd
name_unit <- function(items, name) {
  if (nrow(items) == 0L) return(items)
  if (nrow(items) == 1L) {
    items$name <- name
    return(items)
  }
  purrr::list_rbind(list(group_item(name, NA_character_, drawn_box(items)),
                         reparent(items, name)))
}

# Each row of folded pieces as a block of its own.
loose_blocks <- function(folded, member_of = NA_character_) {
  if (nrow(folded) == 0L) return(new_blocks(character(), numeric(), list()))
  new_blocks("loose", folded$pos, unname(split(to_items(folded), seq_len(nrow(folded)))),
             member_of = member_of)
}

unit_blocks_node <- function(u, pu) {
  reg <- registry_row(u$shape)
  drawn <- any(pu$kind %in% c("shape", "line") & !pu$html)
  items <- if (!is.null(reg) && drawn) {
    registered_item(pu, reg, u$hand_drawn, u$name)
  } else {
    name_unit(to_items(fold_pieces(pu)), u$name)
  }
  new_blocks("node", u$pos, list(items), parent = u$parent)
}

# A subgraph block holds only its background, so that its other parts (a
# title with no rect behind it, a state's inner box) take their paint order
# among the subgraph's members.
unit_blocks_subgraph <- function(u, pu) {
  reg <- registry_row(u$shape)
  bg_name <- paste0("mermaid:subgraph_bg:", u$id)
  drawn <- any(pu$kind %in% c("shape", "line") & !pu$html)
  if (!is.null(reg) && drawn) {
    bg <- registered_item(pu, reg, u$hand_drawn, bg_name)
    rest <- loose_blocks(pu[0, ])
  } else {
    folded <- fold_pieces(pu)
    first_shape <- match("shape", folded$kind)
    bg <- if (is.na(first_shape)) empty_items() else
      to_items(with_name(folded[first_shape, ], bg_name))
    rest <- loose_blocks(if (is.na(first_shape)) folded else folded[-first_shape, ],
                         member_of = u$id)
  }
  purrr::list_rbind(list(
    new_blocks("subgraph", u$pos, list(bg), gid = u$id, parent = u$parent),
    rest
  ))
}

with_name <- function(p, name) {
  p$name <- name
  p
}

# Edges are never filled and keep their exact path, even a rectangular one.
unit_blocks_edge <- function(u, pu) {
  pu$fill <- NA_character_
  pu$fill_a <- 0
  pu$prst <- NA_character_
  pu$closed <- purrr::map_lgl(pu$path, function(p) !is.null(p) && path_closed(p))
  pu$name <- u$name
  new_blocks("edge", u$pos, list(to_items(pu)), start = u$start, end = u$end)
}

unit_blocks_edge_label <- function(u, pu) {
  new_blocks("edge_label", u$pos, list(name_unit(to_items(fold_pieces(pu)), u$name)),
             start = u$start, end = u$end)
}

.unit_builders <- list(
  node = unit_blocks_node, subgraph = unit_blocks_subgraph,
  edge = unit_blocks_edge, edge_label = unit_blocks_edge_label
)

# ---- membership and nesting ----------------------------------------------------

# The groups above `id`, nearest first (not `id` itself).
group_chain <- function(id, parent_of, seen = character()) {
  if (is.na(id)) return(character())
  up <- unname(parent_of[id])
  if (is.na(up) || up %in% seen) return(character())
  c(up, group_chain(up, parent_of, c(seen, up)))
}

# The deepest group holding both ends of an edge.
deepest_common <- function(start, end, parent_of) {
  common <- intersect(group_chain(start, parent_of), group_chain(end, parent_of))
  if (length(common) == 0L) NA_character_ else common[[1]]
}

#' With a db: nodes and subgraphs sit in their db parent, edges and their
#' labels in the deepest subgraph holding both ends
#' @noRd
db_membership <- function(blocks, nodes) {
  groups <- nodes$id[nodes$is_group]
  parent_of <- stats::setNames(nodes$parent, nodes$id)
  by_parent <- blocks$kind %in% c("node", "subgraph") & blocks$parent %in% groups
  blocks$member_of[by_parent] <- blocks$parent[by_parent]
  linked <- blocks$kind %in% c("edge", "edge_label")
  blocks$member_of[linked] <- purrr::map2_chr(blocks$start[linked], blocks$end[linked],
                                              deepest_common, parent_of = parent_of)
  blocks
}

#' Without a db: each unit sits in the smallest subgraph background holding
#' its centre
#' @noRd
box_membership <- function(blocks) {
  bg <- purrr::map(blocks$items, drawn_box)
  is_cluster <- blocks$kind == "subgraph" & purrr::map_lgl(blocks$items, function(it) nrow(it) > 0L)
  cb <- do.call(rbind, c(list(matrix(numeric(), 0L, 4L)), bg[is_cluster]))
  cluster_ids <- blocks$gid[is_cluster]
  area <- cb[, 3] * cb[, 4]
  inside_of <- function(box, kind) {
    cx <- box[[1]] + box[[3]] / 2
    cy <- box[[2]] + box[[4]] / 2
    own <- if (kind == "subgraph") box[[3]] * box[[4]] else -Inf
    ok <- which(cb[, 1] <= cx & cx <= cb[, 1] + cb[, 3] &
                  cb[, 2] <= cy & cy <= cb[, 2] + cb[, 4] & area > own)
    if (length(ok) == 0L) NA_character_ else cluster_ids[[ok[[which.min(area[ok])]]]]
  }
  open <- is.na(blocks$member_of) & blocks$kind != "loose" &
    !purrr::map_lgl(bg, anyNA)
  blocks$member_of[open] <- purrr::map2_chr(bg[open], blocks$kind[open], inside_of)
  blocks
}

# db groups that hold members but were not drawn (no cluster in the DOM)
# still group them.
add_undrawn_groups <- function(blocks, parent_of) {
  missing <- setdiff(stats::na.omit(blocks$member_of), stats::na.omit(blocks$gid))
  if (length(missing) == 0L) return(blocks)
  first_pos <- purrr::map_dbl(missing, function(g) min(blocks$pos[blocks$member_of %in% g]))
  added <- new_blocks("subgraph", first_pos, rep(list(empty_items()), length(missing)),
                      gid = missing, member_of = unname(parent_of[missing]))
  add_undrawn_groups(purrr::list_rbind(list(blocks, added)), parent_of)
}

# The rows of one level of the group tree (group `gid`, NA for the top), in
# paint order, each subgraph followed by its background and members.
nest_level <- function(gid, gname, blocks, seen = character()) {
  here <- if (is.na(gid)) is.na(blocks$member_of) else blocks$member_of %in% gid
  level <- blocks[here, ]
  level <- level[order(level$pos), ]
  rows <- purrr::pmap(list(level$kind, level$gid, level$items), function(kind, g, items) {
    if (kind != "subgraph") return(reparent(items, gname))
    if (g %in% seen) return(empty_items())
    name <- paste0("mermaid:subgraph:", g)
    inner <- purrr::list_rbind(list(reparent(items, name),
                                    nest_level(g, name, blocks, c(seen, g))))
    if (nrow(inner) == 0L) return(empty_items())
    purrr::list_rbind(list(group_item(name, gname, drawn_box(inner)), inner))
  })
  purrr::list_rbind(c(list(empty_items()), rows))
}

#' Scene items from pieces: units assembled by their rules, everything else
#' by the generic ones, nested into subgraph groups
#' @noRd
scene_items <- function(pieces, db) {
  nodes <- db_node_table(db)
  edges <- db_edge_table(db)
  pu <- piece_units(pieces, nodes)
  units <- unit_table(pu, pieces$pos, nodes, edges)

  in_unit <- !is.na(pu$key)
  by_unit <- split(pieces[in_unit, ], factor(pu$key[in_unit], levels = units$key))
  unit_rows <- split(units, seq_len(nrow(units)))
  blocks <- purrr::list_rbind(c(
    list(new_blocks(character(), numeric(), list())),
    purrr::map2(unit_rows, by_unit, function(u, p) .unit_builders[[u$kind]](u, p)),
    list(loose_blocks(fold_pieces(pieces[!in_unit, ])))
  ))

  has_db <- nrow(nodes) > 0L
  blocks <- if (has_db) db_membership(blocks, nodes) else box_membership(blocks)
  blocks <- add_undrawn_groups(blocks, stats::setNames(nodes$parent, nodes$id))
  nest_level(NA_character_, NA_character_, blocks)
}

# R/assemble.R
# Folded pieces -> the items tibble described at the top of R/scene.R: how
# units and loose items nest into groups and what everything is called.
#
#   - Units (R/semantics.R) become blocks: a node, an edge, an edge label, or
#     a subgraph's background; loose pieces become blocks of their own, those
#     drawn in one DOM element grouped by scope.
#   - Blocks nest into subgraph groups: by Mermaid's own parents when the
#     diagram has a db, else by which subgraph background holds them.
#   - A stroke with marker shapes stands for the group of it and them.
#   - Names: every name is mermaid:<kind>:<tag> (mermaid_name()); group names
#     are made unique once, at the end, in document order.

# ---- names -------------------------------------------------------------------

#' mermaid:<kind>:<tag>, or mermaid:<kind> when the tag is empty or NA
#'
#' The one form of every name the scene gives. Vectorised over both.
#' @noRd
mermaid_name <- function(kind, tag = "") {
  if (length(kind) == 0L || length(tag) == 0L) return(character())
  tag <- ifelse(is.na(tag), "", as.character(tag))
  paste0("mermaid:", kind, ifelse(nzchar(tag), paste0(":", tag), ""))
}

#' The group of a stroke and its marker shapes, named after the stroke's
#' item: mermaid:edge:a-b gives mermaid:edge_group:a-b, mermaid:line:msg
#' gives mermaid:edge_group:line:msg
#' @noRd
edge_group_name <- function(name) {
  unit_edge <- grepl("^mermaid:edge(:|$)", name)
  ifelse(unit_edge, sub("^mermaid:edge", "mermaid:edge_group", name),
         sub("^mermaid:", "mermaid:edge_group:", name))
}

first_class_token <- function(cls) {
  tok <- stringr::str_extract(cls, "[^\\s]+")
  ifelse(is.na(tok), "", tok)
}

# What names an element: its id, else its first class, else "".
element_tag <- function(id, cls) ifelse(nzchar(id), id, first_class_token(cls))

#' Names of items that belong to no unit: mermaid:<kind>:<prim id>, else
#' :<first class token>, else just mermaid:<kind>
#' @noRd
generic_names <- function(p) mermaid_name(p$kind, element_tag(p$id, p$cls))

# x, x#2, x#3, ... in the order given.
unique_names <- function(base) {
  k <- stats::ave(seq_along(base), base, FUN = seq_along)
  ifelse(k == 1L, base, paste0(base, "#", k))
}

# ---- items while they are assembled ------------------------------------------
#
# While items are assembled every row carries `depth`, how far it is nested
# inside the block it belongs to: a group row is followed by its
# descendants, one level deeper. Blocks are joined by stacking them and
# nested by deepening them, so no row needs its group's final name until
# finish_items() links each row to the group row above it and makes the
# group names unique.

#' Pieces as items rows: generic names unless named already, item columns
#' and depth only
#' @noRd
to_items <- function(p) {
  unnamed <- is.na(p$name)
  p$name[unnamed] <- generic_names(p[unnamed, ])
  custom <- p$kind == "shape" & is.na(p$prst)
  p$path[!custom] <- list(NULL)
  p$depth <- rep(0L, nrow(p))
  p[, c(item_cols, "depth")]
}

empty_items <- function() to_items(piece_ptype())

group_item <- function(name, box) {
  new_item(name = name, kind = "group", x = box[[1]], y = box[[2]], w = box[[3]],
           h = box[[4]], depth = 0L)
}

#' Items under a new group `name`, whose box holds their drawn rows
#' @noRd
group_under <- function(items, name) {
  items$depth <- items$depth + 1L
  purrr::list_rbind(list(group_item(name, drawn_box(items)), items))
}

#' The box of drawn rows (not groups)
#' @noRd
drawn_box <- function(items) {
  d <- items$kind != "group"
  union_box(items$x[d], items$y[d], items$w[d], items$h[d])
}

# Blocks of items one after another (NULL entries allowed).
stack_items <- function(blocks) purrr::list_rbind(c(list(empty_items()), blocks))

#' Items as the scene gives them: each row's parent named, group names
#' unique in document order, without `depth`
#' @noRd
finish_items <- function(items) {
  up <- parent_rows(items$depth)
  g <- items$kind == "group"
  items$name[g] <- unique_names(items$name[g])
  items$parent <- items$name[up]
  items[, item_cols]
}

# For each row, the row of its group: the last row above it one level up;
# NA at the top.
parent_rows <- function(depth) {
  rows <- seq_along(depth)
  up <- rep(NA_integer_, length(depth))
  links <- purrr::map(setdiff(unique(depth), 0L), function(d) {
    kids <- rows[depth == d]
    above <- rows[depth == d - 1L]
    list(kids = kids, up = above[findInterval(kids, above)])
  })
  up[unlist(purrr::map(links, "kids"))] <- unlist(purrr::map(links, "up"))
  up
}

# ---- strokes with marker shapes -----------------------------------------------

#' Each piece as the items it stands for: its own row, or, when `markers`
#' holds shapes of it, the group of it and them (only them for a stroke of
#' no length, which they draw)
#'
#' @param items `to_items(p)`, named as the caller wants.
#' @param markers The side table of marker shapes (R/markers.R).
#' @return A list of items tibbles, one per piece.
#' @noRd
piece_entries <- function(items, p, markers) {
  entries <- vctrs::vec_chop(items)
  hosts <- which(p$i %in% markers$host)
  if (length(hosts) == 0L) return(entries)
  shapes_of <- rows_by_key(markers$host)
  entries[hosts] <- purrr::map(hosts, function(h) {
    shapes <- to_items(markers[shapes_of(p$i[[h]]), ])
    own <- if (!zero_length(p[h, ])) entries[[h]]
    group_under(stack_items(list(own, shapes)), edge_group_name(items$name[[h]]))
  })
  entries
}


# ---- scope groups ---------------------------------------------------------------
#
# Loose items (outside every unit) drawn inside one DOM element, such as a
# legend entry, an actor figure or an icon, are one thing to a Word user: they
# become a group `mermaid:group:<id | first class | "g">`, nested as the
# elements nest.

# A scope is grouped when it holds 2 to 40 loose items, not all of them free
# texts (Mermaid's edgeLabels element holds the labels of unrelated edges),
# and is not the whole diagram (it is under 90 % of the union of all items in
# at least one direction).
.scope_group_limits <- list(min = 2L, max = 40L, whole = 0.9)

#' The scopes (DOM serials) of the pieces' chains, and the pieces' chains
#' inverted once: `members` (the pieces drawn inside each scope), `chain_at`
#' (each piece's chain as positions in `scopes`) and `depth` (where each
#' scope sits in a chain)
#' @noRd
scope_chains <- function(p) {
  chains <- p$chain
  serial <- unlist(chains, use.names = FALSE)
  owner <- rep(seq_along(chains), lengths(chains))
  scopes <- unique(serial)
  at <- match(serial, scopes)
  level <- sequence(lengths(chains))
  list(scopes = scopes,
       members = unname(split(owner, factor(at, levels = seq_along(scopes)))),
       chain_at = unname(split(at, factor(owner, levels = seq_along(chains)))),
       depth = level[match(seq_along(scopes), at)])
}

#' Which scopes become groups (logical, along `sc$scopes`), deepest decided
#' first: a scope whose loose items would form fewer than two children
#' (items, or groups already made inside it) makes no group
#' @param all_box The union box of every item of the diagram.
#' @noRd
grouped_scopes <- function(p, sc, all_box) {
  lim <- .scope_group_limits
  boxes <- purrr::map(sc$members, function(m) union_box(p$x[m], p$y[m], p$w[m], p$h[m]))
  whole <- purrr::map_lgl(boxes, function(b) {
    b[[3]] >= lim$whole * all_box[[3]] && b[[4]] >= lim$whole * all_box[[4]]
  })
  only_text <- purrr::map_lgl(sc$members, function(m) all(p$kind[m] == "text"))
  count <- lengths(sc$members)
  eligible <- count >= lim$min & count <= lim$max & !whole & !only_text
  todo <- which(eligible)[order(-sc$depth[eligible])]
  purrr::reduce(todo, function(made, k) {
    child <- purrr::map_chr(sc$members[[k]], function(i) {
      below <- sc$chain_at[[i]][-seq_len(sc$depth[[k]])]
      inner <- below[made[below]]
      if (length(inner) > 0L) paste0("g", inner[[1]]) else paste0("i", i)
    })
    if (length(unique(child)) >= lim$min) replace(made, k, TRUE) else made
  }, .init = rep(FALSE, length(sc$scopes)))
}

# mermaid:group:<id | first class | "g"> of the elements with DOM serials
# `serials`, from the pieces' ancestors.
scope_group_names <- function(p, serials) {
  flat <- unlist(p$anc, recursive = FALSE)
  n <- purrr::map_int(flat, function(a) as.integer(a$n %||% NA_integer_))
  tags <- purrr::map_chr(flat[match(serials, n)], function(a) {
    element_tag(as.character(a$id %||% ""), as.character(a$cls %||% ""))
  })
  mermaid_name("group", ifelse(nzchar(tags), tags, "g"))
}

#' Loose pieces as blocks: each item that no scope group holds, and each
#' outermost scope group with everything in it
#'
#' @param all_box The union box of every item of the diagram.
#' @return list(pos, items), one entry per block.
#' @noRd
scope_blocks <- function(p, markers, all_box) {
  if (nrow(p) == 0L) return(list(pos = numeric(), items = list()))
  entries <- piece_entries(to_items(p), p, markers)
  sc <- scope_chains(p)
  made <- grouped_scopes(p, sc, all_box)
  if (!any(made)) return(list(pos = p$pos, items = entries))

  deepest_made <- function(at) {
    hit <- at[made[at]]
    if (length(hit) == 0L) NA_integer_ else hit[[length(hit)]]
  }
  groups <- which(made)
  holder <- purrr::map_int(sc$chain_at, deepest_made)
  # A group's parent is the deepest group above it in its members' chains.
  up <- purrr::map_int(groups, function(k) {
    deepest_made(sc$chain_at[[sc$members[[k]][[1]]]][seq_len(sc$depth[[k]] - 1L)])
  })
  first_pos <- purrr::map_dbl(sc$members[groups], function(m) min(p$pos[m]))
  names <- scope_group_names(p, sc$scopes[groups])
  items_in <- rows_by_key(holder)
  groups_in <- rows_by_key(up)

  # The rows of group j (an index into `groups`): its members in paint order.
  group_rows <- function(j) {
    own <- items_in(groups[[j]])
    subs <- groups_in(groups[[j]])
    parts <- c(entries[own], purrr::map(subs, group_rows))
    inner <- purrr::list_rbind(parts[order(c(p$pos[own], first_pos[subs]))])
    group_under(inner, names[[j]])
  }
  top <- items_in(NA_integer_)
  top_groups <- groups_in(NA_integer_)
  list(pos = c(p$pos[top], first_pos[top_groups]),
       items = c(entries[top], purrr::map(top_groups, group_rows)))
}

# ---- blocks: units and loose items, before nesting -----------------------------
#
# A block is what gets placed at one level of the group tree: a unit's items,
# one loose item, or a subgraph (whose `items` are its background and whose
# members are the blocks with member_of == its id).

new_blocks <- function(kind, pos, items, gid = NA_character_,
                       member_of = NA_character_, start = NA_character_,
                       end = NA_character_, parent = NA_character_) {
  quick_tibble(kind = as.character(kind), pos = as.numeric(pos), items = items,
               gid = gid, member_of = member_of, start = start, end = end,
               parent = parent)
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
  group_under(items, name)
}

# Loose pieces (folded) as blocks: each piece a block of its own, except
# icons and pictures, which are grouped by the scope rule.
loose_blocks <- function(folded, markers, all_box, member_of = NA_character_) {
  apart <- drawn_apart(folded)
  own <- folded[!apart, ]
  beside <- scope_blocks(folded[apart, ], markers, all_box)
  new_blocks("loose", c(own$pos, beside$pos),
             c(piece_entries(to_items(own), own, markers), beside$items),
             member_of = member_of)
}

#' A unit's folded pieces as its items in paint order: each piece an entry
#' (with its marker shapes), icons and pictures grouped by the scope rule as
#' loose items are
#' @param lead Items standing for the unit's other pieces (a registered
#'   shape), placed at `lead_pos`; NULL to take every piece of `p`.
#' @noRd
unit_items <- function(p, markers, all_box, lead = NULL, lead_pos = NA_real_) {
  apart <- drawn_apart(p)
  own <- p[!apart, ]
  entries <- if (is.null(lead)) piece_entries(to_items(own), own, markers) else list(lead)
  pos <- if (is.null(lead)) own$pos else lead_pos
  beside <- scope_blocks(p[apart, ], markers, all_box)
  blocks <- c(entries, beside$items)
  stack_items(blocks[order(c(pos, beside$pos))])
}

# Units whose icons are named after them; an icon anywhere else is
# mermaid:icon:icon.
.icon_owner_kinds <- c("node", "subgraph")

#' Icons (inline shapes) named mermaid:icon:<tag>
#' @noRd
name_icons <- function(p, tag) {
  p$name[p$inline] <- mermaid_name("icon", tag)
  p
}

icon_tag <- function(u) if (u$kind %in% .icon_owner_kinds) u$id else "icon"

# The pieces that draw a unit's outline: not its text, its HTML label
# boxes, its icons or its pictures.
outline_pieces <- function(pu) pu$kind %in% c("shape", "line") & !pu$html & !pu$inline

# A registered shape that the unit draws: the node becomes that one shape.
registered_for <- function(u, pu) {
  reg <- registry_row(u$shape)
  if (!is.null(reg) && any(outline_pieces(pu))) reg
}

unit_blocks_node <- function(u, pu, markers, all_box) {
  pu <- name_icons(pu, icon_tag(u))
  reg <- registered_for(u, pu)
  items <- if (!is.null(reg)) {
    shape <- registered_item(pu, reg, u$hand_drawn, u$name)
    lead_pos <- min(pu$pos[!drawn_apart(pu)])
    name_unit(unit_items(pu, markers, all_box, lead = shape, lead_pos = lead_pos), u$name)
  } else {
    name_unit(unit_items(fold_pieces(pu), markers, all_box), u$name)
  }
  new_blocks("node", u$pos, list(items), parent = u$parent)
}

# A subgraph block holds only its background, so that its other parts (a
# title with no rect behind it, a state's inner box) take their paint order
# among the subgraph's members.
unit_blocks_subgraph <- function(u, pu, markers, all_box) {
  pu <- name_icons(pu, icon_tag(u))
  reg <- registered_for(u, pu)
  bg_name <- mermaid_name("subgraph_bg", u$id)
  if (!is.null(reg)) {
    bg <- registered_item(pu, reg, u$hand_drawn, bg_name)
    rest <- loose_blocks(pu[drawn_apart(pu), ], markers, all_box, member_of = u$id)
  } else {
    folded <- fold_pieces(pu)
    first_shape <- match(TRUE, folded$kind == "shape" & !drawn_apart(folded))
    bg <- if (is.na(first_shape)) empty_items() else name_unit(to_items(folded[first_shape, ]), bg_name)
    rest <- loose_blocks(if (is.na(first_shape)) folded else folded[-first_shape, ], markers,
                         all_box, member_of = u$id)
  }
  purrr::list_rbind(list(
    new_blocks("subgraph", u$pos, list(bg), gid = u$id, parent = u$parent),
    rest
  ))
}

# Edges are never filled and keep their exact path, even a rectangular one.
unit_blocks_edge <- function(u, pu, markers, all_box) {
  pu$fill <- NA_character_
  pu$fill_a <- 0
  pu$fill_grad <- list(NULL)
  pu$prst[pu$kind != "line"] <- NA_character_
  pu$closed <- purrr::map_lgl(pu$path, function(p) !is.null(p) && path_closed(p))
  pu$name <- u$name
  pu <- name_icons(pu, icon_tag(u))
  new_blocks("edge", u$pos, list(unit_items(pu, markers, all_box)), start = u$start, end = u$end)
}

unit_blocks_edge_label <- function(u, pu, markers, all_box) {
  pu <- name_icons(pu, icon_tag(u))
  items <- name_unit(unit_items(fold_pieces(pu), markers, all_box), u$name)
  new_blocks("edge_label", u$pos, list(items), start = u$start, end = u$end)
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

#' The rows of one level of the group tree (subgraph `gid`, NA for the top),
#' in paint order, each subgraph as a group of its background and members
#' @param members `rows_by_key(blocks$member_of)`.
#' @noRd
nest_level <- function(gid, blocks, members, seen = character()) {
  level <- blocks[members(gid), ]
  level <- level[order(level$pos), ]
  rows <- purrr::pmap(list(level$kind, level$gid, level$items), function(kind, g, items) {
    if (kind != "subgraph") return(items)
    if (g %in% seen) return(NULL)
    inner <- stack_items(list(items, nest_level(g, blocks, members, c(seen, g))))
    if (nrow(inner) > 0L) group_under(inner, mermaid_name("subgraph", g))
  })
  stack_items(rows)
}

#' Scene items from pieces: units assembled by their rules, everything else
#' by the generic ones (and grouped by scope), nested into subgraph groups
#' @param markers The side table of marker shapes (R/markers.R).
#' @noRd
scene_items <- function(pieces, markers, db) {
  nodes <- db_node_table(db)
  pu <- piece_units(pieces)
  units <- unit_table(pu, pieces$pos, nodes)

  in_unit <- !is.na(pu$key)
  unit_rows <- split(which(in_unit), factor(pu$key[in_unit], levels = units$key))
  by_unit <- vctrs::vec_chop(pieces, indices = unname(unit_rows))
  all_box <- union_box(c(pieces$x, markers$x), c(pieces$y, markers$y),
                       c(pieces$w, markers$w), c(pieces$h, markers$h))
  loose_pieces <- name_icons(pieces[!in_unit, ], "icon")
  loose <- scope_blocks(fold_pieces(loose_pieces), markers, all_box)
  blocks <- purrr::list_rbind(c(
    list(new_blocks(character(), numeric(), list())),
    purrr::map2(vctrs::vec_chop(units), by_unit, function(u, p) {
      .unit_builders[[u$kind]](u, p, markers, all_box)
    }),
    list(new_blocks("loose", loose$pos, loose$items))
  ))

  has_db <- nrow(nodes) > 0L
  blocks <- if (has_db) db_membership(blocks, nodes) else box_membership(blocks)
  blocks <- add_undrawn_groups(blocks, stats::setNames(nodes$parent, nodes$id))
  finish_items(nest_level(NA_character_, blocks, rows_by_key(blocks$member_of)))
}

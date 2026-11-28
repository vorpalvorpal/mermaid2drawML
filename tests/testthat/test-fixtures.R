# The bridge fixtures end to end: the checks every fixture must pass, then
# what a Word user sees in each diagram type, in the scene and in the XML.

# Rows of `items` whose box is within `tol` px of `box` on every coordinate.
at_box <- function(items, box, tol = 1) {
  if (nrow(items) == 0L) return(items)
  b <- cbind(items$x, items$y, items$w, items$h)
  items[rowSums(abs(sweep(b, 2L, box)) <= tol) == 4L, ]
}

# TRUE when two of the rows draw the same box (within `tol`).
any_same_box <- function(items, tol = 1) {
  n <- nrow(items)
  if (n < 2L) return(FALSE)
  pairs <- utils::combn(n, 2L)
  any(apply(pairs, 2L, function(ij) {
    all(abs(box_of(items[ij[1], ]) - box_of(items[ij[2], ])) <= tol)
  }))
}

# The one drawn shape of a unit at the unit's own box.
unit_box_shapes <- function(items, name) {
  u <- item_named(items, name)
  if (nrow(u) != 1L) return(items[0, ])
  sub <- subtree_rows(items, name)
  at_box(sub[sub$kind == "shape", ], box_of(u))
}

expect_boxes_are_one_shape <- function(nm, ids) {
  items <- fixture_scene(nm)$items
  for (id in ids) {
    name <- node_name(id)
    expect_equal(nrow(item_named(items, name)), 1L, info = name)
    s <- unit_box_shapes(items, name)
    expect_equal(nrow(s), 1L, info = name)
    expect_true(all(!is.na(s$fill) & !is.na(s$stroke)), info = name)
    sub <- subtree_rows(items, name)
    areas <- sub[sub$kind == "shape" & sub$w > 1 & sub$h > 1, ]
    expect_false(any_same_box(areas), info = paste(name, "draws a box twice"))
  }
}

# A box crossed by dividers is drawn as stacked compartment rects. The
# compartments of unit `name`: full-width shapes of its subtree stacked
# from its top to its bottom, each the tallest one starting where the last
# ended (an ER entity's attribute rows lie over its body compartment).
unit_compartments <- function(items, name) {
  u <- item_named(items, name)
  if (nrow(u) != 1L) return(items[0, ])
  sub <- subtree_rows(items, name)
  full <- sub[sub$kind == "shape" & abs(sub$x - u$x) <= 1 & abs(sub$w - u$w) <= 1, ]
  picked <- integer()
  y <- u$y
  repeat {
    at <- which(abs(full$y - y) <= 1)
    if (length(at) == 0L) break
    k <- at[which.max(full$h[at])]
    picked <- c(picked, k)
    y <- full$y[[k]] + full$h[[k]]
    if (abs(y - (u$y + u$h)) <= 1) return(full[picked, ])
  }
  items[0, ]
}

expect_boxes_are_compartments <- function(nm, ids) {
  items <- fixture_scene(nm)$items
  for (id in ids) {
    name <- node_name(id)
    expect_equal(nrow(item_named(items, name)), 1L, info = name)
    s <- unit_compartments(items, name)
    expect_gte(nrow(s), 1L, label = paste(name, "compartments"))
    expect_true(all(!is.na(s$fill) & !is.na(s$stroke)), info = name)
    sub <- subtree_rows(items, name)
    areas <- sub[sub$kind == "shape" & sub$w > 1 & sub$h > 1, ]
    expect_false(any_same_box(areas), info = paste(name, "draws a box twice"))
  }
}

db_node_ids <- function(nm) {
  nodes <- purrr::discard(fixture_render(nm)$db$nodes, function(n) isTRUE(n$isGroup))
  vapply(nodes, function(n) n$id, character(1))
}

# The edge prim (with its marker ids) of a db edge.
edge_prim <- function(nm, edge_id) {
  hits <- purrr::keep(fixture_prims(nm, "path"), function(p) identical(p$id, edge_id))
  if (length(hits) == 0L) NULL else hits[[1]]
}

# -- Every fixture --------------------------------------------------------------

for_each_fixture("build_scene() gives a well-formed items tibble", function(nm) {
  sc <- fixture_scene(nm)
  it <- sc$items
  expect_s3_class(sc, "mermaid_scene")
  expect_true(all(scene_cols %in% names(it)))
  expect_gt(nrow(it), 0L)
  expect_true(all(it$kind %in% c("group", "shape", "line", "text", "image")))

  drawn <- it[it$kind != "group", ]
  for (col in c("x", "y", "w", "h")) {
    expect_false(anyNA(drawn[[col]]), info = paste(nm, col))
  }
  expect_true(all(drawn$w >= 0 & drawn$h >= 0))
  expect_false(anyNA(it$rot))
  expect_true(all(is.na(it$fill) | grepl("^[0-9A-Fa-f]{6}$", it$fill)))
  expect_true(all(is.na(it$stroke) | grepl("^[0-9A-Fa-f]{6}$", it$stroke)))
  expect_true(all(it$head %in% end_kinds))
  expect_true(all(it$tail %in% end_kinds))
  expect_false(anyNA(it$label))

  # No item that holds text is flipped vertically.
  expect_false(any(has_text(it) & it$flip_v))

  # Every parent is a group item of this scene.
  parents <- unique(stats::na.omit(it$parent))
  expect_true(all(parents %in% it$name[it$kind == "group"]))

  # No label is lost: every text prim's text is in some item's label.
  texts <- fixture_prims(nm, "text")
  for (p in texts) {
    txt <- paste(prim_row_texts(p), collapse = "\n")
    if (!nzchar(trimws(txt))) next
    expect_true(any(grepl(txt, it$label, fixed = TRUE)),
                info = paste0(nm, ": text prim ", p$i, " '", txt, "'"))
  }
})

for_each_fixture("build_diagram_xml() gives well-formed, consistent XML", function(nm) {
  r <- fixture_render(nm)
  res <- NULL
  expect_no_warning(res <- quietly_expected(build_diagram_xml(r, start_id = 3L)))
  doc <- NULL
  expect_no_error(doc <- dml_doc(res$xml))

  # ids: unique, consecutive from start_id, next_id one past the last
  ids <- all_ids(doc)
  expect_false(anyDuplicated(ids) > 0L)
  expect_equal(ids, seq.int(3L, length.out = length(ids)))
  expect_equal(res$next_id, 3L + length(ids))

  # wrapper
  expect_length(xfind(doc, "//wp:inline"), 1L)
  expect_equal(inline_extent(doc), root_ext(doc))
  for (g in as.list(xfind(doc, "//wpg:grpSp"))) {
    expect_false(is_missing_node(xfirst(g, "./wpg:cNvPr")))
  }
  # no shape that holds text is flipped vertically
  expect_length(xfind(doc, "//wps:wsp[wps:txbx]/wps:spPr/a:xfrm[@flipV]"), 0L)

  # offsets and extents: integers, sizes at least 1 EMU
  offs <- xfind(doc, "//a:xfrm/a:off")
  exts <- xfind(doc, "//a:xfrm/a:ext")
  expect_true(int_attr_ok(c(xattr(offs, "x"), xattr(offs, "y"))))
  expect_true(int_attr_ok(c(xattr(exts, "cx"), xattr(exts, "cy"))))
  expect_true(all(as.numeric(c(xattr(exts, "cx"), xattr(exts, "cy"))) >= 1))

  # every shape lies within the drawing (its rotated footprint, for rotated ones)
  ext <- root_ext(doc)
  tol <- 4 * ext[["cx"]] / r$viewBox[3]          # 4 px
  for (s in as.list(all_wsp(doc))) {
    fp <- rotated_footprint(abs_box(s), shape_xfrm(s)$rot)
    ok <- fp[["x"]] >= -tol && fp[["y"]] >= -tol &&
      fp[["x"]] + fp[["w"]] <= ext[["cx"]] + tol && fp[["y"]] + fp[["h"]] <= ext[["cy"]] + tol
    expect_true(ok, info = paste0(nm, ": ", shape_name(s), " at ",
                                  paste(round(fp), collapse = ", ")))
  }

  # no label is lost: every text row is a paragraph somewhere
  rows <- unlist(purrr::map(fixture_prims(nm, "text"), prim_row_texts))
  rows <- rows[nzchar(trimws(rows))]
  paras <- all_paragraph_texts(doc)
  for (t in unique(rows)) {
    expect_gte(sum(paras == t), sum(rows == t), label = paste0(nm, ": '", t, "'"))
  }
})

for_each_fixture("no filled open path is a connector, and every text row is kept", function(nm) {
  it <- fixture_scene(nm)$items
  open_filled <- it[it$closed %in% FALSE & !is.na(it$fill), ]
  expect_true(all(open_filled$kind == "shape" & is.na(open_filled$prst) &
                    purrr::map_lgl(open_filled$text, is.null)), info = nm)
  # such shapes are written as shapes, so no connector has a fill
  doc <- fixture_doc(nm)
  connectors <- as.list(xfind(doc, "//wps:wsp[wps:cNvCnPr]"))
  filled <- purrr::map_lgl(connectors, function(s) identical(shape_fill(s)$kind, "solid"))
  expect_equal(sum(filled), 0L, info = nm)

  prim_rows <- unlist(purrr::map(fixture_prims(nm, "text"), prim_row_texts))
  item_rows <- unlist(purrr::map(it$text, function(t) {
    if (is.null(t)) character() else purrr::map_chr(t$rows, function(r) paste(r$t, collapse = ""))
  }))
  keep <- function(x) sort(x[nzchar(trimws(x))])
  expect_equal(keep(item_rows), keep(prim_rows), info = nm)
})

# -- flowchart-paint ------------------------------------------------------------

test_that("flowchart-paint: node A keeps its style fill, stroke and width", {
  a <- one_item(fixture_scene("flowchart-paint")$items, node_name("A"))
  expect_equal(a$kind, "shape")
  expect_equal(a$prst, "rect")
  expect_equal(a$fill, "ADD8E6")
  expect_equal(a$stroke, "FF0000")
  expect_equal(a$stroke_w, 4)
  expect_identical(a$label, "Start here")
})

test_that("flowchart-paint: `style D fill:none` leaves D unfilled but outlined", {
  items <- fixture_scene("flowchart-paint")$items
  d <- subtree_rows(items, node_name("D"))
  d <- d[d$kind == "shape", ]
  expect_gte(nrow(d), 1L)
  expect_true(all(is.na(d$fill)))
  expect_true(any(!is.na(d$stroke)))
})

test_that("flowchart-paint: node E's fill keeps its 50 % alpha", {
  e <- one_item(fixture_scene("flowchart-paint")$items, node_name("E"))
  expect_equal(e$fill, "FF0000")
  expect_near(e$fill_a, 0.5, tol = 1e-3)
})

test_that("flowchart-paint: the cylinder is preset can at its true height", {
  # Fixture bbox 80.5 x 68.38 px: the straight side alone is 48.8 px.
  cc <- one_item(fixture_scene("flowchart-paint")$items, node_name("C"))
  expect_equal(cc$prst, "can")
  expect_near(c(cc$x, cc$y, cc$w, cc$h), c(114.902, 466.672, 80.5, 68.38), tol = 1e-3)
})

test_that("flowchart-paint: the asymmetric (odd) node is drawn from its outline, not a preset", {
  items <- fixture_scene("flowchart-paint")$items
  sub <- subtree_rows(items, node_name("D"))
  expect_false(any(sub$prst %in% c("flowChartMultidocument", "flowChartPunchedTape",
                                   "flowChartDocument")))
  expect_true(any(sub$kind == "shape" & is.na(sub$prst)))
})

test_that("flowchart-paint: one edge item per drawn link and none for F ~~~ A", {
  items <- fixture_scene("flowchart-paint")$items
  edges <- named_with(items, "mermaid:edge:")
  expect_setequal(edges$name, c(edge_name("A", "B"), edge_name("B", "C"),
                                edge_name("B", "D"), edge_name("E", "C")))
  # a straight edge (a path of zero width) is a line item
  expect_true(all(edges$kind %in% c("shape", "line")))
  expect_true(all(is.na(edges$fill)))
  expect_true(all(edges$stroke == "333333"))
  expect_true(all(edges$tail == "triangle"))
  expect_true(all(edges$head == "none"))
  expect_equal(item_named(items, edge_name("E", "C"))$stroke_w, 3.5)
  expect_equal(item_named(items, edge_name("A", "B"))$stroke_w, 1)
})

test_that("flowchart-paint: labels are kept verbatim, rows split, markdown rendered", {
  items <- fixture_scene("flowchart-paint")$items
  expect_identical(item_named(items, node_name("F"))$label, "2 * 3 * 4 __init__")
  expect_identical(item_named(items, node_name("B"))$label, "Decide\nnow")
  expect_identical(item_named(items, node_name("E"))$label, "A bold word")
})

test_that("flowchart-paint: edge labels are shapes with the label background and text", {
  items <- fixture_scene("flowchart-paint")$items
  yes <- one_item(items, "mermaid:edge_label:B:C")
  expect_equal(yes$kind, "shape")
  expect_identical(yes$label, "yes")
  expect_equal(yes$fill, "E8E8E8")
  expect_lt(yes$fill_a, 1)
  expect_identical(one_item(items, "mermaid:edge_label:B:D")$label, "no")
})

test_that("flowchart-paint: subgraph S1 groups its background, A, B and edge A-B", {
  items <- fixture_scene("flowchart-paint")$items
  g <- one_item(items, "mermaid:subgraph:S1")
  expect_equal(g$kind, "group")
  kids <- children_of(items, "mermaid:subgraph:S1")
  expect_equal(kids$name[1], "mermaid:subgraph_bg:S1")
  expect_true(all(c(node_name("A"), node_name("B"), edge_name("A", "B")) %in% kids$name))

  bg <- one_item(items, "mermaid:subgraph_bg:S1")
  expect_equal(bg$fill, "FFFFDE")
  expect_identical(bg$label, "Group one")

  # B -> C leaves the subgraph, so it and C stay at the top level.
  expect_true(is.na(item_named(items, edge_name("B", "C"))$parent))
  expect_true(is.na(item_named(items, node_name("C"))$parent))
  expect_true(is.na(item_named(items, "mermaid:edge_label:B:C")$parent))
})

test_that("flowchart-paint: A-B is a line whose triangle tail ends at B", {
  items <- fixture_scene("flowchart-paint")$items
  e <- item_named(items, "mermaid:edge:A-B")
  expect_equal(nrow(e), 1L)
  if (nrow(e) != 1L) return()
  expect_equal(e$kind, "line")
  expect_equal(e$tail, "triangle")
  b <- item_named(items, "mermaid:node:B")
  expect_equal(nrow(b), 1L)
  if (nrow(b) != 1L) return()
  ends <- item_ends(e)
  bb <- box_of(b)
  # the flips put the tail end at B, not the head end
  expect_lt(box_gap(c(ends$end, 0, 0), bb), 1)
  expect_gt(box_gap(c(ends$start, 0, 0), bb), box_gap(c(ends$end, 0, 0), bb))
})

test_that("flowchart-paint XML: node A is ADD8E6 with a 4 px FF0000 outline", {
  a <- shape_named(fixture_doc("flowchart-paint"), "mermaid:node:A")
  expect_equal(shape_fill(a)[c("kind", "color")], list(kind = "solid", color = "ADD8E6"))
  ln <- shape_ln(a)
  expect_equal(ln$color, "FF0000")
  expect_equal(ln$w, 4 * 9525)
})

test_that("flowchart-paint XML: D has noFill and a visible line", {
  d <- shape_named(fixture_doc("flowchart-paint"), "mermaid:node:D")
  outlines <- purrr::keep(as.list(xfind(d, "descendant-or-self::wps:wsp")),
                          function(n) !identical(xattr(xfirst(n, "./wps:cNvSpPr"), "txBox"), "1"))
  expect_gte(length(outlines), 1L)
  for (o in outlines) {
    expect_equal(shape_fill(o)$kind, "none", info = shape_name(o))
    expect_equal(shape_ln(o)$kind, "solid", info = shape_name(o))
  }
})

test_that("flowchart-paint XML: E's fill carries alpha 50000", {
  e <- shape_named(fixture_doc("flowchart-paint"), "mermaid:node:E")
  f <- shape_fill(e)
  expect_equal(f$color, "FF0000")
  expect_near(f$alpha, 50000, tol = 100)
})

test_that("flowchart-paint XML: the subgraph background keeps the theme fill FFFFDE", {
  bg <- shape_named(fixture_doc("flowchart-paint"), "mermaid:subgraph_bg:S1")
  expect_equal(shape_fill(bg)[c("kind", "color")], list(kind = "solid", color = "FFFFDE"))
})

test_that("flowchart-paint XML: the cylinder is can at its true height", {
  cc <- shape_named(fixture_doc("flowchart-paint"), "mermaid:node:C")
  expect_equal(shape_prst(cc), "can")
  # 68.38 is given to two decimals (68.3796 in the fixture): half a unit of
  # its last digit is 0.005 px.
  expect_near(unname(shape_ext(cc)), c(80.5, 68.38) * 9525, tol = 0.005 * 9525)
})

test_that("flowchart-paint XML: the odd node D is custom geometry, not flowChartMultidocument", {
  d <- shape_named(fixture_doc("flowchart-paint"), "mermaid:node:D")
  prsts <- xattr(xfind(d, "descendant-or-self::wps:wsp/wps:spPr/a:prstGeom"), "prst")
  expect_false("flowChartMultidocument" %in% prsts)
  expect_gte(length(xfind(d, "descendant-or-self::wps:wsp/wps:spPr/a:custGeom")), 1L)
})

test_that("flowchart-paint XML: one edge shape per drawn link, styled from the prims", {
  doc <- fixture_doc("flowchart-paint")
  edges <- shapes_named(doc, "mermaid:edge:")
  expect_length(edges, 4L)
  expect_equal(count_named(doc, "mermaid:edge:F-A") + count_named(doc, "mermaid:edge:A-F"), 0)
  for (e in as.list(edges)) {
    ln <- shape_ln(e)
    nm <- shape_name(e)
    expect_equal(ln$color, "333333", info = nm)
    expect_equal(shape_fill(e)$kind, "none", info = nm)
    expect_equal(ln$tail, "triangle", info = nm)
    # the 8 px flowchart point is "lg" on a 1 px edge, "sm" on a 3.5 px one
    size <- if (identical(nm, "mermaid:edge:E-C")) "sm" else "lg"
    expect_equal(c(ln$tail_w, ln$tail_len), c(size, size), info = nm)
    # a straight edge is a preset line
    expect_true(shape_is_custom(e) || identical(shape_prst(e), "line"), info = nm)
    expect_false(is_missing_node(xfirst(e, "./wps:cNvCnPr")), info = nm)
  }
  dash <- function(name) shape_ln(shape_named(doc, name))$dash
  expect_equal(dash("mermaid:edge:B-D"), "sysDash")       # dotted: dash 2 at width 1
  expect_true(is.na(dash("mermaid:edge:A-B")))
  expect_true(is.na(dash("mermaid:edge:B-C")))
  expect_true(is.na(dash("mermaid:edge:E-C")))
  expect_near(shape_ln(shape_named(doc, "mermaid:edge:E-C"))$w, 3.5 * 9525, tol = 1)
})

test_that("flowchart-paint XML: the edge label is a filled shape with its text inside", {
  el <- shape_named(fixture_doc("flowchart-paint"), "mermaid:edge_label:B:C")
  expect_true(is_wsp_node(el))
  f <- shape_fill(el)
  expect_equal(f$color, "E8E8E8")
  # three stacked backgrounds at 0.5, 0.8, 0.8: 1 - 0.5 * 0.2 * 0.2 = 0.98
  expect_near(f$alpha, 98000, tol = 100)
  expect_equal(shape_paragraph_texts(el), "yes")
})

test_that("flowchart-paint XML: subgraph S1 holds its background, A, B and A-B, not B-C", {
  doc <- fixture_doc("flowchart-paint")
  g <- shape_named(doc, "mermaid:subgraph:S1")
  expect_equal(xml2::xml_name(g), "grpSp")
  inside <- descendant_names(g)
  expect_true(all(c("mermaid:subgraph_bg:S1", "mermaid:node:A", "mermaid:node:B",
                    "mermaid:edge:A-B") %in% inside))
  expect_false(any(c("mermaid:edge:B-C", "mermaid:node:C", "mermaid:edge_label:B:C") %in% inside))
  first_child <- xfind(g, "./wps:wsp | ./wpg:grpSp")[[1]]
  expect_equal(shape_name(first_child), "mermaid:subgraph_bg:S1")
})

# -- flowchart-full -------------------------------------------------------------

test_that("flowchart-full: every node gives exactly one item named after it", {
  items <- fixture_scene("flowchart-full")$items
  db_nodes <- fixture_render("flowchart-full")$db$nodes
  ids <- vapply(purrr::discard(db_nodes, function(n) isTRUE(n$isGroup)),
                function(n) n$id, character(1))
  counts <- vapply(ids, function(id) sum(items$name == node_name(id), na.rm = TRUE), integer(1))
  expect_true(all(counts == 1L), info = paste(ids[counts != 1L], collapse = ", "))
  expect_true(all(items$kind[items$name %in% node_name(ids)] %in% c("shape", "group")))
})

test_that("flowchart-full: edges are named from db start/end, invisible links have none", {
  items <- fixture_scene("flowchart-full")$items
  edges <- named_with(items, "mermaid:edge:")$name
  expect_length(edges, 44L)                       # 45 links, one of them a ~~~ b
  expect_false(edge_name("a", "i") %in% edges)
  expect_true(edge_name("ic", "im") %in% edges)   # user-defined edge id e1
  expect_true(edge_name("uni", "ic") %in% edges)  # e2
})

test_that("flowchart-full: line ends follow the link markers", {
  items <- fixture_scene("flowchart-full")$items
  ends <- function(from, to) unlist(item_named(items, edge_name(from, to))[c("head", "tail")])
  expect_equal(ends("a", "b"), c(head = "none", tail = "triangle"))
  expect_equal(ends("a", "c"), c(head = "none", tail = "none"))
  # the 11 px circle on a 1 px line is drawn as a marker shape
  expect_equal(ends("a", "f"), c(head = "none", tail = "none"))
  expect_equal(ends("a", "g"), c(head = "none", tail = "none"))   # Word has no cross end
  expect_equal(ends("a", "h"), c(head = "triangle", tail = "triangle"))
})

test_that("flowchart-full: linkStyle, classDef and style paint come from the prims", {
  items <- fixture_scene("flowchart-full")$items
  e <- item_named(items, edge_name("n1", "n2"))
  expect_equal(e$stroke, "FF0000")
  expect_equal(e$stroke_w, 4)
  a <- item_named(items, node_name("a"))
  expect_equal(c(a$fill, a$stroke), c("FF9999", "990000"))
  expect_equal(a$stroke_w, 3)
  cc <- item_named(items, node_name("c"))
  expect_equal(c(cc$fill, cc$stroke), c("99FF99", "009900"))
})

test_that("flowchart-full: nested subgraphs nest, and edges sit in the deepest common one", {
  items <- fixture_scene("flowchart-full")$items
  expect_equal(one_item(items, "mermaid:subgraph:S1")$kind, "group")
  expect_equal(one_item(items, "mermaid:subgraph:S2")$parent, "mermaid:subgraph:S1")
  expect_true(is.na(item_named(items, "mermaid:subgraph:S1")$parent))
  expect_equal(children_of(items, "mermaid:subgraph:S1")$name[1], "mermaid:subgraph_bg:S1")
  expect_equal(children_of(items, "mermaid:subgraph:S2")$name[1], "mermaid:subgraph_bg:S2")

  parent_of <- function(name) item_named(items, name)$parent
  expect_equal(parent_of(node_name("s1")), "mermaid:subgraph:S1")
  expect_equal(parent_of(node_name("s2")), "mermaid:subgraph:S1")
  expect_equal(parent_of(node_name("s3")), "mermaid:subgraph:S2")
  expect_equal(parent_of(node_name("s4")), "mermaid:subgraph:S2")
  expect_equal(parent_of(edge_name("s1", "s2")), "mermaid:subgraph:S1")
  expect_equal(parent_of(edge_name("s3", "s4")), "mermaid:subgraph:S2")
  expect_equal(parent_of(edge_name("s2", "s3")), "mermaid:subgraph:S1")
  expect_true(is.na(parent_of(node_name("n1"))))
  expect_true(is.na(parent_of(edge_name("n1", "n2"))))
})

test_that("flowchart-full: registered shapes get their Word presets", {
  items <- fixture_scene("flowchart-full")$items
  expected <- c(
    n1 = "rect", n2 = "roundRect", n3 = "roundRect", n4 = "flowChartPredefinedProcess",
    n5 = "can", n6 = "ellipse", n8 = "diamond", n9 = "hexagon",
    n10 = "parallelogram", n11 = "parallelogram", n12 = "trapezoid",
    v2 = "flowChartDocument", v4 = "can", v6 = "parallelogram",
    v9 = "flowChartManualInput", v10 = "flowChartMultidocument"
  )
  for (id in names(expected)) {
    it <- item_named(items, node_name(id))
    expect_equal(it$kind, "shape", info = id)
    expect_equal(it$prst, expected[[id]], info = id)
  }
})

test_that("flowchart-full: registered nodes drawn with several paths are one shape with their text", {
  items <- fixture_scene("flowchart-full")$items
  # stadium n3 (two paths), document v2 (two), stacked documents v10 (four)
  for (id in c("n3", "v2", "v10")) {
    it <- one_item(items, node_name(id))
    expect_equal(it$kind, "shape", info = id)
    expect_equal(nrow(children_of(items, node_name(id))), 0L, info = id)
  }
  expect_identical(item_named(items, node_name("n3"))$label, "Stadium")
  expect_identical(item_named(items, node_name("v10"))$label, "Docs")
})

test_that("flowchart-full: shapes with no true Word preset are drawn from their outlines", {
  items <- fixture_scene("flowchart-full")$items
  # double circle: its two circles, not one ellipse preset
  n14 <- one_item(items, node_name("n14"))
  expect_equal(n14$kind, "group")
  rings <- children_of(items, node_name("n14"))
  expect_equal(sum(rings$kind == "shape" & rings$prst %in% "ellipse"), 2L)
  # odd (asymmetric) and tagged rect: custom geometry somewhere in the unit
  for (id in c("n7", "v5")) {
    sub <- subtree_rows(items, node_name(id))
    expect_true(any(sub$kind == "shape" & is.na(sub$prst)), info = id)
    it <- one_item(items, node_name(id))
    expect_true(it$kind == "group" || is.na(it$prst), info = id)
  }
})

test_that("flowchart-full: adjust values and flips match the drawn shapes", {
  doc <- fixture_doc("flowchart-full")
  node <- function(id) shape_named(doc, node_name(id))
  # rounded rect: rx 5 on a 54 px high box
  expect_equal(shape_adj(node("n2"))[["adj"]], round(100000 * 5 / 54))
  expect_equal(shape_adj(node("n3"))[["adj"]], 50000)
  # every slanted node has a 19.5 / 39 (or 27 / 54) inset
  for (id in c("n10", "n11", "n12")) {
    expect_equal(shape_adj(node(id))[["adj"]], 50000, info = id)
  }
  expect_equal(shape_adj(node("n9"))[["adj"]], 25000)   # 9.75 / 39
  expect_false(shape_xfrm(node("n10"))$flipH)
  expect_true(shape_xfrm(node("n11"))$flipH)
  expect_false(shape_xfrm(node("n12"))$flipV)
  # the inverted trapezoid is drawn from its outline, upright
  expect_false(shape_xfrm(node("n13"))$flipV)
  expect_true(shape_is_custom(node("n13")))
})

test_that("flowchart-full draws its image (the bridge read its picture); flowchart-paint has none", {
  expect_false("image" %in% fixture_scene("flowchart-full")$unsupported)
  expect_length(fixture_scene("flowchart-paint")$unsupported, 0L)
})

test_that("flowchart-full XML: subgraph S2 is a group nested in S1", {
  doc <- fixture_doc("flowchart-full")
  s2 <- shape_named(doc, "mermaid:subgraph:S2")
  expect_equal(xml2::xml_name(s2), "grpSp")
  expect_true("mermaid:subgraph:S2" %in% descendant_names(shape_named(doc, "mermaid:subgraph:S1")))
  expect_true(all(c("mermaid:node:s3", "mermaid:node:s4", "mermaid:edge:s3-s4") %in%
                    descendant_names(s2)))
  expect_false("mermaid:edge:s2-s3" %in% descendant_names(s2))
})

test_that("flowchart-full: triangle ends end within 1 px of the node they point at", {
  items <- fixture_scene("flowchart-full")$items
  node_box <- function(id) {
    it <- item_named(items, node_name(id))
    if (nrow(it) == 0L) it <- item_named(items, paste0("mermaid:subgraph:", id))
    if (nrow(it) == 0L) NULL else box_of(it[1, ])
  }
  checked <- 0L
  for (e in fixture_render("flowchart-full")$db$edges) {
    rows <- item_named(items, edge_name(e$start, e$end))
    for (k in seq_len(nrow(rows))) {
      it <- rows[k, ]
      ends <- item_ends(it)
      if (it$tail == "triangle" && !is.null(b <- node_box(e$end))) {
        checked <- checked + 1L
        # Mermaid itself leaves 2 px before a fork/join bar (v7).
        tol <- if (identical(e$end, "v7")) 2.5 else 1
        expect_lte(point_gap(ends$end, b), tol, label = paste(it$name, "tail"))
      }
      if (it$head == "triangle" && !is.null(b <- node_box(e$start))) {
        checked <- checked + 1L
        expect_lte(point_gap(ends$start, b), 1, label = paste(it$name, "head"))
      }
    }
  }
  expect_gt(checked, 30L)
})

test_that("flowchart-full: the odd node is one shape holding its label", {
  n7 <- item_named(fixture_scene("flowchart-full")$items, node_name("n7"))
  expect_equal(nrow(n7), 1L)
  expect_equal(n7$kind, "shape")
  expect_identical(n7$label, "Asymmetric")
})

test_that("flowchart-full: the tag-rect node draws its box and its tag once each", {
  # The corner tag has a box of its own, so the box is one shape holding the
  # label, the tag one more, and nothing is drawn twice.
  items <- fixture_scene("flowchart-full")$items
  sub <- subtree_rows(items, node_name("v5"))
  shapes <- sub[sub$kind == "shape", ]
  expect_lte(nrow(shapes), 2L)
  expect_false(any_same_box(shapes))
  expect_true("Tag" %in% shapes$label)
})

test_that("flowchart-full: subgraph titles are centred across, at the top, no side insets", {
  items <- fixture_scene("flowchart-full")$items
  for (id in c("S1", "S2")) {
    bg <- item_named(items, paste0("mermaid:subgraph_bg:", id))
    expect_equal(nrow(bg), 1L, info = id)
    if (nrow(bg) != 1L) next
    tx <- bg$text[[1]]
    expect_equal(tx$align, "center", info = id)
    expect_equal(tx$valign, "top", info = id)
    expect_equal(tx$insets[c(1, 3)], c(0, 0), info = id)
  }
})

test_that("flowchart-full: the circle end of a-f is a marker shape; the edge ends where Mermaid's did", {
  items <- fixture_scene("flowchart-full")$items
  e <- item_named(items, "mermaid:edge:a-f")
  expect_equal(nrow(e), 1L)
  if (nrow(e) != 1L) return()
  expect_equal(c(e$head, e$tail), c("none", "none"))
  expect_equal(e$parent, "mermaid:edge_group:a-f")
  marks <- children_of(items, "mermaid:edge_group:a-f")
  marks <- marks[marks$name %in% "mermaid:marker:flowchart-v2-circleEnd", ]
  expect_equal(nrow(marks), 1L)
  if (nrow(marks) != 1L) return()
  expect_equal(marks$prst, "ellipse")
  expect_false(is.na(marks$fill))
  p <- purrr::detect(fixture_prims("flowchart-full", "path"), function(p) identical(p$id, "L_a_f_0"))
  path <- path_transform(path_parse(p$geom$d), prim_matrix(p))
  path <- path[path$cmd != "Z", ]
  drawn_end <- c(path$x[[nrow(path)]], path$y[[nrow(path)]])
  expect_near(item_ends(e)$end, drawn_end, tol = 0.01)        # not moved
  expect_lt(box_gap(c(drawn_end, 0, 0), box_of(marks)), 1.5)
})

# -- flowchart-handdrawn --------------------------------------------------------

test_that("flowchart-handdrawn: each registered node is ONE shape with its text inside", {
  items <- fixture_scene("flowchart-handdrawn")$items
  expected <- list(A = c("rect", "Start"), B = c("diamond", "Is it?"), C = c("roundRect", "OK"),
                   D = c("can", "Db"), E = c("ellipse", "End"),
                   F = c("flowChartPredefinedProcess", "Sub"))
  for (id in names(expected)) {
    it <- one_item(items, node_name(id))
    expect_equal(it$kind, "shape", info = id)
    expect_equal(it$prst, expected[[id]][1], info = id)
    expect_identical(it$label, expected[[id]][2], info = id)
    expect_equal(nrow(children_of(items, node_name(id))), 0L, info = id)
  }
})

test_that("flowchart-handdrawn: a node's box is the union of its prims' boxes", {
  # D is drawn by three paths; their union is x 344.458..380.675, y 43.06..98.707.
  d <- item_named(fixture_scene("flowchart-handdrawn")$items, node_name("D"))
  expect_near(c(d$x, d$y, d$w, d$h), c(344.458, 43.06, 36.217, 55.647), tol = 1e-3)
})

# -- base-segoe-stadium-bold-invisible ------------------------------------------

test_that("circle, stadium and rounded nodes hold their own text (no overlay or label shape)", {
  items <- fixture_scene("base-segoe-stadium-bold-invisible")$items
  for (id in c("A", "C", "D")) {
    it <- one_item(items, node_name(id))
    expect_equal(it$kind, "shape", info = id)
    expect_equal(nrow(children_of(items, node_name(id))), 0L, info = id)
  }
  expect_equal(item_named(items, node_name("A"))$prst, "roundRect")
  expect_equal(item_named(items, node_name("C"))$prst, "ellipse")
  expect_equal(item_named(items, node_name("D"))$prst, "roundRect")
  expect_identical(item_named(items, node_name("A"))$label, "Start here")
  expect_identical(item_named(items, node_name("C"))$label, "Circle")
  expect_identical(item_named(items, node_name("D"))$label, "Rounded")
  expect_equal(nrow(named_with(items, "mermaid:label")), 0L)
  expect_equal(sum(items$kind == "text"), 0L)

  doc <- fixture_doc("base-segoe-stadium-bold-invisible")
  for (id in c("A", "C", "D")) {
    node <- shape_named(doc, node_name(id))
    expect_true(is_wsp_node(node), info = id)
    expect_equal(length(shape_paragraphs(node)), 1L, info = id)
  }
})

test_that("a transparent classDef fill gives no fill", {
  d <- item_named(fixture_scene("base-segoe-stadium-bold-invisible")$items, node_name("D"))
  expect_true(is.na(d$fill))
  expect_equal(d$stroke, "000000")
})

# -- classDiagram ---------------------------------------------------------------

test_that("classDiagram: each class is a unit named after its class id", {
  items <- fixture_scene("classDiagram")$items
  for (id in c("Animal", "Dog", "Cat", "Bone", "Tail", "Collar", "Mouse", "Pet", "note0")) {
    one_item(items, node_name(id))
  }
  animal <- subtree_rows(items, node_name("Animal"))
  labels <- paste(animal$label, collapse = "\n")
  for (t in c("Animal", "+String name", "-int age", "+eat() : void")) {
    expect_true(grepl(t, labels, fixed = TRUE), info = t)
  }
  expect_true(edge_name("Dog", "Bone") %in% items$name)
  expect_identical(item_named(items, "mermaid:edge_label:Dog:Bone")$label, "owns")
})

test_that("classDiagram: each class box is its stacked compartments and the note one shape", {
  ids <- db_node_ids("classDiagram")
  expect_boxes_are_compartments("classDiagram", ids)
  notes <- ids[startsWith(ids, "note")]
  expect_gte(length(notes), 1L)
  expect_boxes_are_one_shape("classDiagram", notes)
})

test_that("classDiagram: extension ends are hollow triangle shapes touching the parent class box", {
  items <- fixture_scene("classDiagram")$items
  edges <- fixture_render("classDiagram")$db$edges
  seen <- 0L
  for (e in edges) {
    p <- edge_prim("classDiagram", e$id)
    for (end in c("start", "end")) {
      id <- p[[paste0("marker_", end)]]
      if (is.null(id) || !grepl("extension", id)) next
      seen <- seen + 1L
      g <- edge_group_of(e$start, e$end)
      kids <- children_of(items, g)
      marks <- kids[kids$name %in% paste0("mermaid:marker:", id), ]
      expect_equal(nrow(marks), 1L, info = g)
      if (nrow(marks) != 1L) next
      expect_true(is.na(marks$fill) || marks$fill_a == 0, info = g)       # hollow
      expect_false(is.na(marks$stroke), info = g)
      path <- marks$path[[1]]
      corners <- if (is.null(path)) 0L else sum(path$cmd %in% c("M", "L"))
      expect_true(corners == 3L || marks$prst %in% "triangle", info = g)
      parent <- item_named(items, node_name(if (end == "start") e$start else e$end))
      expect_equal(nrow(parent), 1L, info = g)
      if (nrow(parent) == 1L) expect_lt(box_gap(box_of(marks), box_of(parent)), 1.5, label = paste(g, "marker gap to its node"))
    }
  }
  expect_equal(seen, 3L)     # Animal <|-- Dog, Animal <|-- Cat, Cat ..|> Pet
})

test_that("classDiagram: Animal is three stacked rects holding its title, attributes and methods", {
  items <- fixture_scene("classDiagram")$items
  u <- item_named(items, "mermaid:node:Animal")
  expect_equal(nrow(u), 1L)
  sub <- subtree_rows(items, "mermaid:node:Animal")
  s <- sub[sub$kind == "shape", ]
  expect_equal(nrow(s), 3L)
  expect_true(all(s$prst %in% "rect"))
  if (nrow(s) != 3L) return()
  s <- s[order(s$y), ]
  # the box 334.31, 8, 165.03 x 216; dividers at y 80 and 152
  expect_near(cbind(s$x, s$y, s$w, s$h),
              cbind(rep(334.31, 3), c(8, 80, 152), rep(165.03, 3), rep(72, 3)), tol = 0.02)
  expect_identical(s$label, c("«abstract»\nAnimal", "+String name\n-int age",
                              "+eat() : void\n+sleep() : void"))
  top <- s$text[[1]]
  expect_length(top$rows, 2L)
  expect_equal(top$align, "center")
  # The bridge reports anchor "middle" for these HTML rows, which alone would
  # give align "center" with an inset (centring the shorter row in Word);
  # the attribute and method rows are left-aligned.
  expect_equal(s$text[[2]]$align, "left")
  expect_equal(s$text[[3]]$align, "left")
  expect_length(s$text[[2]]$rows, 2L)
  expect_length(s$text[[3]]$rows, 2L)
})

test_that("classDiagram: no free text and no divider line is left in any class node", {
  items <- fixture_scene("classDiagram")$items
  nodes <- items$name[items$kind == "group" & startsWith(items$name, "mermaid:node:")]
  expect_gte(length(nodes), 8L)
  for (nm in nodes) {
    sub <- subtree_rows(items, nm)
    expect_equal(sum(sub$kind == "text"), 0L, info = nm)
    drawn <- sub[sub$kind != "group", ]
    expect_equal(sum(drawn$kind == "line" | drawn$h <= 0.01), 0L, info = nm)
  }
})

test_that("classDiagram has no mermaid:group:edgeLabels", {
  items <- fixture_scene("classDiagram")$items
  expect_false(any(grepl("^mermaid:group:edgeLabels", items$name)))
})

test_that("classDiagram: the Dog-Tail composition is a filled diamond shape touching the Dog box", {
  # the 17 px diamond on a 1 px line is too big for Word's end: the diagram's
  # one composition marker is a shape
  items <- fixture_scene("classDiagram")$items
  expect_equal(nrow(named_with(items, "mermaid:marker:class-composition")), 1L)
  e <- item_named(items, "mermaid:edge:Dog-Tail")
  expect_equal(nrow(e), 1L)
  if (nrow(e) != 1L) return()
  expect_equal(e$head, "none")
  g <- "mermaid:edge_group:Dog-Tail"
  expect_equal(e$parent, g)
  marks <- children_of(items, g)
  marks <- marks[marks$name %in% "mermaid:marker:class-compositionStart", ]
  expect_equal(nrow(marks), 1L)
  if (nrow(marks) != 1L) return()
  expect_equal(marks$fill, "333333")
  path <- marks$path[[1]]
  corners <- if (is.null(path)) 0L else sum(path$cmd %in% c("M", "L"))
  expect_true(corners == 4L || marks$prst %in% "diamond")
  dog <- item_named(items, "mermaid:node:Dog")
  expect_equal(nrow(dog), 1L)
  if (nrow(dog) == 1L) expect_lt(box_gap(box_of(marks), box_of(dog)), 1.5)
  p <- purrr::detect(fixture_prims("classDiagram", "path"), function(p) identical(p$id, "id_Dog_Tail_4"))
  path <- path_transform(path_parse(p$geom$d), prim_matrix(p))
  expect_near(item_ends(e)$start, c(path$x[[1]], path$y[[1]]), tol = 0.01)   # not moved
})

# -- stateDiagram-v2 ------------------------------------------------------------

test_that("stateDiagram: states are units named after their state ids", {
  items <- fixture_scene("stateDiagram-v2")$items
  for (id in c("Still", "Moving", "Crash", "Idle", "Busy", "Other")) {
    it <- one_item(items, node_name(id))
    labels <- subtree_rows(items, node_name(id))$label
    expect_true(id %in% labels, info = id)
  }
  expect_identical(item_named(items, "mermaid:edge_label:Still:Moving")$label, "push")
})

test_that("stateDiagram-v2: the choice node is a filled diamond shape", {
  ch <- item_named(fixture_scene("stateDiagram-v2")$items, node_name("choice"))
  expect_equal(nrow(ch), 1L)
  expect_equal(ch$kind, "shape")
  expect_equal(ch$prst, "diamond")
  expect_false(is.na(ch$fill))
})

test_that("stateDiagram-v2: the end state is drawn with ellipse presets", {
  items <- fixture_scene("stateDiagram-v2")$items
  sub <- subtree_rows(items, node_name("root_end"))
  shapes <- sub[sub$kind == "shape", ]
  expect_equal(nrow(shapes), 2L)                 # the ring and the dot
  expect_true(all(shapes$prst %in% "ellipse"))
})

test_that("stateDiagram-v2: the note is one shape with its text", {
  note <- item_named(fixture_scene("stateDiagram-v2")$items, node_name("Still----note-11"))
  expect_equal(nrow(note), 1L)
  expect_equal(note$kind, "shape")
  expect_identical(note$label, "a note")
})

# -- erDiagram ------------------------------------------------------------------

test_that("erDiagram: entities are units named after their db id", {
  items <- fixture_scene("erDiagram")$items
  one_item(items, node_name("entity-CUSTOMER-0"))
  one_item(items, node_name("entity-ORDER-1"))
  cust <- subtree_rows(items, node_name("entity-CUSTOMER-0"))
  labels <- paste(cust$label, collapse = "\n")
  for (t in c("CUSTOMER", "custNumber", "PK")) expect_true(grepl(t, labels, fixed = TRUE), info = t)
  expect_true(edge_name("entity-CUSTOMER-0", "entity-ORDER-1") %in% items$name)
})

test_that("erDiagram: every entity box is its stacked compartments, filled and outlined", {
  # an entity with attributes is its name and body compartments
  expect_boxes_are_compartments("erDiagram", db_node_ids("erDiagram"))
})

test_that("erDiagram: every relationship line has marker shapes at its ends, grouped with it", {
  items <- fixture_scene("erDiagram")$items
  for (e in fixture_render("erDiagram")$db$edges) {
    g <- edge_group_of(e$start, e$end)
    expect_equal(nrow(item_named(items, g)), 1L, info = g)
    kids <- children_of(items, g)
    edge <- kids[kids$name %in% edge_name(e$start, e$end), ]
    expect_equal(nrow(edge), 1L, info = g)
    if (nrow(edge) != 1L) next
    expect_equal(c(edge$head, edge$tail), c("none", "none"), info = g)
    ends <- item_ends(edge)
    p <- edge_prim("erDiagram", e$id)
    for (end in c("start", "end")) {
      id <- p[[paste0("marker_", end)]]
      if (is.null(id)) next
      marks <- kids[kids$name %in% paste0("mermaid:marker:", id), ]
      expect_gte(nrow(marks), 1L, label = paste(g, end, "markers"))
      other <- setdiff(c("start", "end"), end)
      for (k in seq_len(nrow(marks))) {
        b <- box_of(marks[k, ])
        # at its own end: nearer it than the other end, and within the marker's reach
        expect_lt(point_gap(ends[[end]], b), point_gap(ends[[other]], b), label = paste(g, end, "marker gap"))
        expect_lt(point_gap(ends[[end]], b), 50, label = paste(g, end, "marker gap"))
      }
    }
  }
})

# -- requirementDiagram ---------------------------------------------------------

test_that("requirementDiagram: each box is its stacked compartments, filled and outlined", {
  expect_boxes_are_compartments("requirementDiagram", db_node_ids("requirementDiagram"))
})

test_that("requirementDiagram: each box is two stacked rects with their texts", {
  items <- fixture_scene("requirementDiagram")$items
  want <- list(
    test_req = list(box = rbind(c(8, 194, 157.73, 68), c(8, 262, 157.73, 116)),
                    label = c("<<Requirement>>\ntest_req",
                              "ID: 1\nText: the test text.\nRisk: High\nVerification: Test")),
    test_entity = list(box = rbind(c(17.82, 8, 138.09, 68), c(17.82, 76, 138.09, 44)),
                       label = c("<<Element>>\ntest_entity", "Type: simulation"))
  )
  for (id in names(want)) {
    sub <- subtree_rows(items, paste0("mermaid:node:", id))
    s <- sub[sub$kind == "shape", ]
    expect_equal(nrow(s), 2L, info = id)
    expect_equal(sum(sub$kind %in% c("text", "line")), 0L, info = id)
    if (nrow(s) != 2L) next
    s <- s[order(s$y), ]
    expect_true(all(s$prst %in% "rect"), info = id)
    expect_near(cbind(s$x, s$y, s$w, s$h), want[[id]]$box, tol = 0.02, info = id)
    expect_identical(s$label, want[[id]]$label, info = id)
  }
})

# -- sequenceDiagram ------------------------------------------------------------

test_that("sequenceDiagram: every sequence number is the text of a filled ellipse", {
  items <- fixture_scene("sequenceDiagram")$items
  nums <- purrr::keep(fixture_prims("sequenceDiagram", "text"),
                      function(p) identical(p$cls, "sequenceNumber"))
  expect_equal(length(nums), 8L)
  for (p in nums) {
    t <- paste(prim_row_texts(p), collapse = "\n")
    hit <- items$kind == "shape" & items$label == t & items$prst %in% "ellipse" & !is.na(items$fill)
    expect_equal(sum(hit), 1L, info = t)
  }
})

test_that("sequenceDiagram: no zero-size item exists", {
  items <- fixture_scene("sequenceDiagram")$items
  drawn <- items[items$kind != "group", ]
  expect_equal(sum(drawn$w <= 0.01 & drawn$h <= 0.01), 0L)
})

test_that("sequenceDiagram: the loop and alt frames are rect shapes", {
  items <- fixture_scene("sequenceDiagram")$items
  shapes <- items[items$kind == "shape" & items$prst %in% "rect" & is.na(items$fill), ]
  for (b in list(loop = c(264, 354, 222, 91), alt = c(264, 455, 222, 182))) {
    f <- at_box(shapes, b, tol = 1)
    expect_equal(nrow(f), 1L, info = paste(b, collapse = ", "))
    expect_equal(f$stroke, "DACEF3", info = paste(b, collapse = ", "))
  }
})

test_that("sequenceDiagram: each actor figure is one group", {
  items <- fixture_scene("sequenceDiagram")$items
  for (top in c(-5, 652)) {
    head <- at_box(items[items$kind == "shape", ], c(60, top, 30, 30), tol = 0.5)
    expect_equal(nrow(head), 1L, info = paste("actor head at y", top))
    if (nrow(head) != 1L) next
    g <- head$parent
    expect_false(is.na(g), info = paste("actor head at y", top))
    if (is.na(g)) next
    expect_true(startsWith(g, "mermaid:group:"), info = g)
    kids <- children_of(items, g)
    expect_equal(sum(kids$kind == "line"), 4L, info = g)        # torso, arms, two legs
    expect_true("User" %in% kids$label, info = g)
    expect_equal(nrow(kids), 6L, info = g)
  }
})

test_that("sequenceDiagram: participant boxes hold their names; messages have line ends", {
  items <- fixture_scene("sequenceDiagram")$items
  for (who in c("Alice", "Bob")) {
    boxes <- items[items$kind == "shape" & items$label == who, ]
    expect_gte(nrow(boxes), 2L, label = paste(who, "boxes"))   # top and bottom
    expect_true(all(boxes$prst == "roundRect"), info = who)
  }
  msgs <- items[items$kind == "line" & items$tail != "none", ]
  # filled-head is stealth and crosshead no Word end (a shape)
  expect_gte(nrow(msgs), 7L)
  expect_true(all(c("triangle", "stealth") %in% msgs$tail))   # arrowhead, filled-head
  expect_false("diamond" %in% msgs$tail)
  doc <- fixture_doc("sequenceDiagram")
  tails <- xfind(doc, "//a:ln/a:tailEnd[@type='triangle']")
  expect_gte(length(tails), 6L)
})

# -- gantt ----------------------------------------------------------------------

test_that("items wholly outside the viewBox are not drawn (gantt's today line)", {
  # The today marker is at x = 63553 in a 1584 px wide chart; the browser
  # clips it, and so must Word.
  doc <- fixture_doc("gantt")
  right_edge <- root_ext(doc)[["cx"]]
  lefts <- purrr::map_dbl(as.list(all_wsp(doc)), function(s) abs_box(s)[["x"]])
  expect_false(any(lefts > right_edge))
})

test_that("gantt: task bars are shapes with their label merged or as text beside them", {
  items <- fixture_scene("gantt")$items
  tasks <- c("D3D3D3" = "Done task", "BFC7FF" = "Active", "FF0000" = "Crit", "8A90DD" = "Future")
  for (fill in names(tasks)) {
    bars <- items[items$kind == "shape" & items$fill %in% fill & items$rot == 0, ]
    expect_gte(nrow(bars), 1L, label = paste("bars filled", fill))
    label <- tasks[[fill]]
    expect_true(label %in% bars$label || label %in% items$label[items$kind == "text"],
                info = label)
  }
  # the milestone is a rect rotated 45 degrees (matrix scale 0.8)
  ms <- items[items$kind == "shape" & abs(items$rot - 45) < 1e-3, ]
  expect_equal(nrow(ms), 1L)
  expect_near(c(ms$x, ms$y, ms$w, ms$h), c(1501, 148, 16, 16), tol = 1e-2)
})

# -- pie ------------------------------------------------------------------------

test_that("pie: the slices are filled custom-geometry shapes", {
  items <- fixture_scene("pie")$items
  slices <- items[items$kind == "shape" & items$fill %in% c("ECECFF", "FFFFDE", "B5FF20") &
                    is.na(items$prst), ]
  expect_equal(nrow(slices), 3L)
  expect_near(slices$fill_a, rep(0.7, 3), tol = 1e-3)

  doc <- fixture_doc("pie")
  custom_fills <- purrr::map_chr(
    purrr::keep(as.list(all_wsp(doc)), shape_is_custom),
    function(s) paste(shape_fill(s)$color, shape_fill(s)$alpha))
  for (col in c("ECECFF", "FFFFDE", "B5FF20")) {
    expect_true(paste(col, 70000) %in% custom_fills, info = col)
  }
})

test_that("pie: each legend entry is a group of its swatch and its text", {
  items <- fixture_scene("pie")$items
  g <- named_with(items[items$kind == "group", ], "mermaid:group:legend")
  expect_equal(nrow(g), 3L)
  for (name in g$name) {
    kids <- subtree_rows(items, name)
    expect_equal(sum(kids$kind == "shape"), 1L, info = name)
    expect_equal(sum(kids$kind == "text"), 1L, info = name)
  }
})

test_that("pie: no group holds every other item", {
  items <- fixture_scene("pie")$items
  drawn <- sum(items$kind != "group")
  for (g in items$name[items$kind == "group"]) {
    sub <- subtree_rows(items, g)
    expect_lt(sum(sub$kind != "group"), drawn, label = g)
  }
})

# -- xychart-beta ---------------------------------------------------------------

test_that("xychart: the rotated axis title carries its rotation", {
  items <- fixture_scene("xychart-beta")$items
  title <- items[items$label == "Revenue", ]
  expect_equal(nrow(title), 1L)
  expect_equal(title$rot, -90)

  doc <- fixture_doc("xychart-beta")
  box <- purrr::keep(as.list(all_wsp(doc)),
                     function(s) identical(shape_paragraph_texts(s), "Revenue"))
  expect_length(box, 1L)
  expect_equal(shape_xfrm(box[[1]])$rot, 270 * 60000)
})

# -- timeline -------------------------------------------------------------------

hex_rgb <- function(h) strtoi(substring(h, c(1, 3, 5), c(2, 4, 6)), 16L)

anc_classes <- function(p) paste(purrr::map_chr(p$anc, function(a) a$cls %||% ""), collapse = " ")

test_that("timeline: event boxes carry the brightened colour (brightness 120 %)", {
  items <- fixture_scene("timeline")$items
  boxes <- fixture_prims("timeline", "path")
  section_of <- function(p) stringr::str_extract(anc_classes(p), "section-+[0-9]+")
  tasks <- purrr::keep(boxes, function(p) grepl("taskWrapper", anc_classes(p)))
  events <- purrr::keep(boxes, function(p) grepl("eventWrapper", anc_classes(p)))
  expect_gt(length(events), 0L)
  for (ev in events) {
    task <- purrr::detect(tasks, function(t) identical(section_of(t), section_of(ev)))
    expect_false(is.null(task), info = section_of(ev))
    if (is.null(task)) next
    base <- hex_rgb(toupper(sub("^#", "", task$fill)))
    want <- pmin(255, round(base * 1.2))
    it <- at_box(items[items$kind == "shape", ], as.numeric(unlist(ev$bbox)), tol = 0.5)
    expect_equal(nrow(it), 1L, info = section_of(ev))
    if (nrow(it) != 1L || is.na(it$fill)) next
    expect_lte(max(abs(hex_rgb(it$fill) - want)), 1,
               label = paste0("event fill ", it$fill, " vs brightened ", task$fill))
  }
})

# -- quadrantChart --------------------------------------------------------------

test_that("quadrantChart: each data point is a group of its dot and its label", {
  items <- fixture_scene("quadrantChart")$items
  g <- items[items$kind == "group" & grepl("^mermaid:group:data-point(#[0-9]+)?$", items$name), ]
  expect_equal(nrow(g), 3L)
  for (name in g$name) {
    kids <- subtree_rows(items, name)
    expect_true(any(kids$prst %in% "ellipse"), info = name)
    expect_true(any(grepl("^Campaign", kids$label)), info = name)
  }
})

# -- architecture-beta ----------------------------------------------------------

test_that("architecture-beta: each service and its icon are one group", {
  items <- fixture_scene("architecture-beta")$items
  services <- c(db = "Database", server = "Server", disk = "Storage")
  for (id in names(services)) {
    name <- paste0("mermaid:group:service-", id)
    g <- item_named(items, name)
    expect_equal(nrow(g), 1L, info = name)
    if (nrow(g) != 1L) next
    sub <- subtree_rows(items, name)
    expect_true(services[[id]] %in% sub$label, info = name)
    expect_gte(sum(sub$kind %in% c("shape", "line")), 3L, label = paste(name, "drawn items"))    # the 80 x 80 tile and its icon
  }
})

# -- treemap-beta ---------------------------------------------------------------

test_that("treemap-beta: every leaf's label and value are one text in the leaf's rect", {
  items <- fixture_scene("treemap-beta")$items
  prims <- fixture_prims("treemap-beta")
  scope_of <- function(p) p$anc[[length(p$anc)]]$n
  leaves <- purrr::keep(prims, function(p) identical(p$cls, "treemapLeaf"))
  expect_equal(length(leaves), 4L)
  for (leaf in leaves) {
    texts <- purrr::keep(prims, function(p) {
      p$kind == "text" && identical(scope_of(p), scope_of(leaf))
    })
    label <- paste(unlist(purrr::map(texts, prim_row_texts)), collapse = "\n")
    b <- as.numeric(unlist(leaf$bbox))
    hit <- items[items$kind == "shape" & abs(items$x - b[1]) < 0.5 & abs(items$y - b[2]) < 0.5 &
                   abs(items$w - b[3]) < 0.5 & abs(items$h - b[4]) < 0.5, ]
    expect_equal(nrow(hit), 1L, info = label)
    if (nrow(hit) == 1L) expect_identical(hit$label, label, info = label)
  }
})

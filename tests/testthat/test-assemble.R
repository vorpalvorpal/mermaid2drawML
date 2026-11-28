# Nesting and names (R/assemble.R): subgraph groups, by Mermaid's own
# parents or by which background holds a unit; edges with their marker
# shapes; scope groups of the loose items drawn in one DOM element; and
# group names, unique in document order.

# -- Subgraphs ------------------------------------------------------------------

test_that("with a db, subgraph members, inner edges and their labels go in the group", {
  cluster <- list(
    prim_rect(5, 5, 190, 60, anc = anc_cluster("S"), fill = "#FFFFDE", stroke = "#AAAA33"),
    prim_text("Sub", x = 85, y = 7, w = 30, h = 16, anchor = "start", anc = anc_cluster_label("S"))
  )
  nodes <- c(node_prims("flowchart-A-0", prim_rect(15, 25, 50, 30)),
             node_prims("flowchart-B-1", prim_rect(135, 25, 50, 30)),
             node_prims("flowchart-C-2", prim_rect(135, 120, 50, 30)))
  edges <- list(prim_edge("M65,40 L135,40", "L_A_B_0"),
                prim_edge("M160,55 L160,120", "L_B_C_0"))
  lbl <- list(
    prim_rect(90, 30, 20, 20, anc = anc_edge_label("L_A_B_0"), html = TRUE,
              fill = "#E8E8E8", stroke = NULL),
    prim_text("in", x = 92, y = 32, w = 16, h = 16, anc = anc_edge_label("L_A_B_0"))
  )
  db <- mini_db(
    list(db_node("S", shape = "rect", dom_id = NULL, is_group = TRUE),
         db_node("A", dom_id = "flowchart-A-0", parent = "S"),
         db_node("B", dom_id = "flowchart-B-1", parent = "S"),
         db_node("C", dom_id = "flowchart-C-2")),
    list(db_edge("A", "B"), db_edge("B", "C"))
  )
  items <- build_scene(mini_render(c(cluster, edges, lbl, nodes), db = db,
                                   viewBox = c(0, 0, 200, 160)))$items
  g <- "mermaid:subgraph:S"
  expect_equal(one_item(items, g)$kind, "group")
  kids <- children_of(items, g)
  expect_equal(kids$name[1], "mermaid:subgraph_bg:S")
  expect_identical(item_named(items, "mermaid:subgraph_bg:S")$label, "Sub")
  expect_true(all(c(node_name("A"), node_name("B"), edge_name("A", "B"),
                    "mermaid:edge_label:A:B") %in% kids$name))
  expect_true(is.na(item_named(items, node_name("C"))$parent))
  expect_true(is.na(item_named(items, edge_name("B", "C"))$parent))
})

test_that("without a db, membership is the smallest cluster containing the unit's centre", {
  cluster <- list(prim_rect(5, 5, 140, 90, anc = anc_cluster("S"), fill = "#FFFFDE"))
  nodes <- c(node_prims("flowchart-A-0", prim_rect(20, 30, 60, 30)),
             node_prims("flowchart-B-1", prim_rect(200, 30, 60, 30)))
  items <- build_scene(mini_render(c(cluster, nodes), db = NULL,
                                   viewBox = c(0, 0, 300, 100)))$items
  expect_equal(item_named(items, node_name("A"))$parent, "mermaid:subgraph:S")
  expect_true(is.na(item_named(items, node_name("B"))$parent))
})

test_that("a subgraph with no visible rect still groups its members, with no background", {
  # Mermaid's `style S fill:none,stroke:none`: the bridge drops the unpainted
  # rect, leaving only the title.
  title <- prim_text("Bare", x = 80, y = 7, w = 40, h = 16, anchor = "start",
                     anc = anc_cluster_label("S"))
  nodes <- c(node_prims("flowchart-A-0", prim_rect(15, 30, 50, 30)),
             node_prims("flowchart-B-1", prim_rect(135, 30, 50, 30)))
  db <- mini_db(list(db_node("S", shape = "rect", dom_id = NULL, is_group = TRUE),
                     db_node("A", dom_id = "flowchart-A-0", parent = "S"),
                     db_node("B", dom_id = "flowchart-B-1", parent = "S")))
  items <- build_scene(mini_render(c(list(title), nodes), db = db))$items
  expect_equal(one_item(items, "mermaid:subgraph:S")$kind, "group")
  expect_equal(nrow(item_named(items, "mermaid:subgraph_bg:S")), 0L)
  expect_equal(item_named(items, node_name("A"))$parent, "mermaid:subgraph:S")
  expect_equal(item_named(items, node_name("B"))$parent, "mermaid:subgraph:S")
})

test_that("a node in a subgraph in a subgraph nests in both", {
  root <- list(anc_g(1, cls = "root"))
  outer <- c(root, list(anc_g(2, unit = unit_subgraph("Outer"))))
  inner <- c(root, list(anc_g(3, unit = unit_subgraph("Inner"))))
  # The node's element sits inside the inner subgraph's: the innermost unit
  # owns the prims.
  a <- c(inner, list(anc_g(4, unit = unit_node("A"))))
  prims <- list(
    prim_rect(5, 5, 290, 140, anc = outer, fill = "#FFFFDE"),
    prim_rect(20, 30, 160, 100, anc = inner, fill = "#DDEEFF"),
    prim_rect(40, 60, 60, 30, anc = a),
    prim_text("A", x = 60, y = 67, w = 20, h = 16, anc = c(a, list(anc_g(5)))),
    prim_rect(200, 60, 60, 30, anc = c(root, list(anc_g(6, unit = unit_node("B")))))
  )
  db <- mini_db(list(db_node("Outer", shape = "rect", dom_id = NULL, is_group = TRUE),
                     db_node("Inner", shape = "rect", dom_id = NULL, is_group = TRUE,
                             parent = "Outer"),
                     db_node("A", parent = "Inner"), db_node("B", parent = "Outer")))
  for (with_db in c(TRUE, FALSE)) {
    items <- build_scene(mini_render(prims, db = if (with_db) db,
                                     viewBox = c(0, 0, 300, 150)))$items
    info <- paste("with db:", with_db)
    expect_identical(one_item(items, node_name("A"))$label, "A", info = info)
    expect_equal(item_named(items, node_name("A"))$parent, "mermaid:subgraph:Inner",
                 info = info)
    expect_equal(item_named(items, "mermaid:subgraph:Inner")$parent, "mermaid:subgraph:Outer",
                 info = info)
    expect_equal(item_named(items, node_name("B"))$parent, "mermaid:subgraph:Outer",
                 info = info)
    expect_true(is.na(item_named(items, "mermaid:subgraph:Outer")$parent), info = info)
    # The inner subgraph's background is its own rect, not the node's.
    bg <- item_named(items, "mermaid:subgraph_bg:Inner")
    expect_equal(c(bg$x, bg$y, bg$w, bg$h), c(20, 30, 160, 100), info = info)
  }
})

# -- Edges with the same ends ---------------------------------------------------

test_that("two edges with the same ends keep their own groups and labels", {
  nodes <- c(node_prims("a", prim_rect(10, 10, 60, 40), id = "A"),
             node_prims("b", prim_rect(130, 10, 60, 40), id = "B"))
  edges <- list(
    prim_edge("M70,20 L118,20", "L_A_B_0", marker_end = "flowchart-v2-crossEnd"),
    prim_edge("M70,40 L118,40", "L_A_B_1", marker_end = "flowchart-v2-crossEnd")
  )
  label <- function(id, text, y) {
    anc <- anc_edge_label(id, "A", "B")
    list(prim_rect(85, y, 30, 12, anc = anc, html = TRUE, fill = "#E8E8E8", stroke = NULL),
         prim_text(text, x = 87, y = y, w = 26, h = 12, anc = anc))
  }
  db <- mini_db(list(db_node("A"), db_node("B")),
                list(db_edge("A", "B", id = "L_A_B_0"), db_edge("A", "B", id = "L_A_B_1")))
  r <- mini_render(c(edges, label("L_A_B_0", "one", 14), label("L_A_B_1", "two", 34), nodes),
                   db = db, markers = list("flowchart-v2-crossEnd" = marker_cross_end()))
  items <- build_scene(r)$items
  groups <- items$name[items$kind == "group" &
                         grepl("^mermaid:edge_group:A-B(#[0-9]+)?$", items$name)]
  expect_length(groups, 2L)
  expect_false(anyDuplicated(groups) > 0L)
  for (g in groups) {
    kids <- children_of(items, g)
    expect_equal(kids$name, c("mermaid:edge:A-B", "mermaid:marker:flowchart-v2-crossEnd"),
                 info = g)
  }
  labels <- item_named(items, "mermaid:edge_label:A:B")
  expect_equal(sort(labels$label), c("one", "two"))
})

# -- Scope groups ---------------------------------------------------------------

legend <- function(n, y = 10, cls = "legend", id = "") {
  in_g(list(prim_rect(10, y, 18, 18, fill = "#ECECFF"),
            prim_text("Dogs", x = 32, y = y, w = 40, h = 18, anchor = "start")),
       n = n, cls = cls, id = id)
}

# "The whole diagram" is the union of all items, so a scope tested alone would
# be the whole diagram and not grouped; this puts something else beside it.
outside <- function() list(prim_rect(150, 60, 40, 30, fill = "#00AA00"))

test_that("loose items sharing a scope become a group named after its first class", {
  items <- generic_items(c(legend(5, cls = "legend item"), outside()))
  g <- groups_of(items)
  expect_equal(nrow(g), 1L)
  expect_equal(g$name, "mermaid:group:legend")
  kids <- items[items$kind != "group" & !items$fill %in% "00AA00", ]
  expect_equal(nrow(kids), 2L)
  expect_true(all(kids$parent == "mermaid:group:legend"))
  expect_near(box_of(g), c(10, 10, 62, 18), tol = 1e-9)
  expect_equal(which(items$kind == "group"), 1L)     # the group row comes first
})

test_that("a scope group is named after the element id first, and 'g' when unnamed", {
  expect_equal(groups_of(generic_items(c(legend(5, id = "key", cls = "legend"), outside())))$name,
               "mermaid:group:key")
  expect_equal(groups_of(generic_items(c(legend(5, cls = ""), outside())))$name, "mermaid:group:g")
})

test_that("a scope that would hold one item makes no group", {
  # the text merges into its rect: one item
  items <- generic_items(in_g(list(prim_rect(10, 10, 80, 40),
                                   prim_text("in", x = 30, y = 20, w = 40, h = 20)), n = 5))
  expect_equal(nrow(groups_of(items)), 0L)
  expect_equal(nrow(items), 1L)
  expect_true(is.na(items$parent))
})

test_that("a scope with more than 40 drawn items is not grouped; 40 is", {
  dots <- function(k) {
    # a rect below the dots, so that they are not the whole diagram
    c(in_g(purrr::map(seq_len(k), function(i) prim_rect(5 * i, 10, 2, 2)), n = 5, cls = "dots"),
      list(prim_rect(5, 80, 10, 10)))
  }
  expect_equal(nrow(groups_of(generic_items(dots(40), viewBox = c(0, 0, 400, 100)))), 1L)
  expect_equal(nrow(groups_of(generic_items(dots(41), viewBox = c(0, 0, 400, 100)))), 0L)
})

test_that("a scope covering 90 % of the diagram both ways is the whole diagram: no group", {
  # measured against the union of all items, not the viewBox
  whole <- in_g(list(prim_rect(0, 0, 190, 95), prim_rect(50, 50, 10, 10)), n = 5, cls = "all")
  expect_equal(nrow(groups_of(generic_items(whole))), 0L)
  # all of the width but 80 % of the height (a rect below it): grouped
  wide <- in_g(list(prim_rect(0, 0, 190, 80), prim_rect(50, 50, 10, 10)), n = 5, cls = "all")
  expect_equal(nrow(groups_of(generic_items(c(wide, list(prim_rect(0, 90, 10, 10)))))), 1L)
})

test_that("groups nest when their scopes nest", {
  inner <- in_g(list(prim_rect(10, 10, 10, 10), prim_rect(30, 10, 10, 10)), n = 3, cls = "inner")
  outer <- in_g(c(inner, list(prim_rect(10, 40, 10, 10), prim_rect(30, 40, 10, 10))),
                n = 2, cls = "outer")
  items <- generic_items(c(outer, outside()))
  g <- groups_of(items)
  expect_setequal(g$name, c("mermaid:group:outer", "mermaid:group:inner"))
  expect_true(is.na(item_named(items, "mermaid:group:outer")$parent))
  expect_equal(item_named(items, "mermaid:group:inner")$parent, "mermaid:group:outer")
  drawn <- items[items$kind != "group", ]
  expect_equal(sum(drawn$parent %in% "mermaid:group:inner"), 2L)
  expect_equal(sum(drawn$parent %in% "mermaid:group:outer"), 2L)
})

test_that("sibling scopes with the same class get unique names in document order", {
  # Names get "#<k>" suffixes in document order. The first may be plain or
  # suffixed, so this checks only uniqueness, the pattern and the order.
  items <- generic_items(c(legend(5, y = 10), legend(8, y = 40), legend(11, y = 70)))
  g <- groups_of(items)
  expect_equal(nrow(g), 3L)
  expect_false(anyDuplicated(g$name) > 0L)
  expect_true(all(grepl("^mermaid:group:legend(#[0-9]+)?$", g$name)))
  expect_equal(g$y, c(10, 40, 70))          # in document order
  k <- suppressWarnings(as.integer(sub(".*#", "", g$name)))
  k[is.na(k)] <- 0L
  expect_true(all(diff(k) > 0))
})

test_that("items of a unit are not regrouped by scope", {
  anc <- list(anc_g(1, cls = "root"), anc_g(2, cls = "nodes"),
              anc_g(3, cls = "default", unit = unit_node("X")))
  prims <- list(
    prim_ellipse(60, 50, 40, anc = anc),
    prim_ellipse(60, 50, 35, anc = anc),
    prim_text("Ring", x = 45, y = 42, w = 30, h = 16, anc = c(anc, list(anc_g(4, cls = "label"))))
  )
  db <- mini_db(list(db_node("X", shape = "doublecircle")))
  items <- build_scene(mini_render(prims, db = db))$items
  expect_false(any(grepl("^mermaid:group:", items$name)))
  expect_equal(item_named(items, "mermaid:node:X")$kind, "group")
})

test_that("a scope holding only free texts is not grouped", {
  labels <- in_g(list(prim_text("1", x = 10, y = 10, w = 10, h = 12),
                      prim_text("*", x = 60, y = 40, w = 10, h = 12)), n = 5, cls = "edgeLabels")
  items <- generic_items(c(labels, outside()))
  expect_equal(nrow(groups_of(items)), 0L)
  expect_equal(sum(items$kind == "text"), 2L)
})

test_that("a scope holding a shape and a text is still grouped", {
  legend <- in_g(list(prim_rect(10, 10, 18, 18, fill = "#ECECFF"),
                      prim_text("Dogs", x = 32, y = 10, w = 40, h = 18, anchor = "start")),
                 n = 5, cls = "legend")
  items <- generic_items(c(legend, outside()))
  g <- groups_of(items)
  expect_equal(g$name, "mermaid:group:legend")
  expect_equal(nrow(children_of(items, "mermaid:group:legend")), 2L)
})

test_that("a scope covering 90 % of the union of all items both ways is not grouped, whatever the viewBox", {
  whole <- in_g(list(prim_rect(10, 10, 100, 50), prim_rect(50, 30, 60, 30)), n = 5, cls = "all")
  # union of all items: 10..115 x 10..65, the scope 100 x 50 of 105 x 55
  items <- generic_items(c(whole, list(prim_rect(105, 55, 10, 10))),
                         viewBox = c(0, 0, 1000, 1000))
  expect_equal(nrow(groups_of(items)), 0L)
})

test_that("a scope covering 90 % of the union in one direction only is grouped", {
  part <- in_g(list(prim_rect(10, 10, 100, 50), prim_rect(50, 30, 60, 30)), n = 5, cls = "all")
  # union 10..110 x 10..110: the scope is all of its width, half its height
  items <- generic_items(c(part, list(prim_rect(50, 100, 10, 10))),
                         viewBox = c(0, 0, 1000, 1000))
  expect_equal(groups_of(items)$name, "mermaid:group:all")
})

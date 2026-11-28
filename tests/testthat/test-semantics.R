# Node, subgraph and edge units for the dagre-family diagrams (spec 3c) and
# the flowchart shape registry (spec 3d). Scene-level checks use the items
# tibble; preset adjust values and flips are read from the XML.

node_name <- function(id) paste0("mermaid:node:", id)
edge_name <- function(from, to) paste0("mermaid:edge:", from, "-", to)

# The one item called `name`; fails the test when there is not exactly one.
one_item <- function(items, name) {
  it <- item_named(items, name)
  testthat::expect_equal(nrow(it), 1L, info = name)
  it
}

# Rows whose name starts with `prefix` (generic items may have no name).
named_like <- function(items, prefix) !is.na(items$name) & startsWith(items$name, prefix)

# Items directly inside group `name`, in order.
children_of <- function(items, name) items[!is.na(items$parent) & items$parent == name, ]

# -- flowchart-paint (behaviours from the code review) -------------------------

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
  edges <- items[named_like(items, "mermaid:edge:"), ]
  expect_setequal(edges$name, c(edge_name("A", "B"), edge_name("B", "C"),
                                edge_name("B", "D"), edge_name("E", "C")))
  expect_true(all(edges$kind == "shape"))
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

# -- flowchart-full ------------------------------------------------------------

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
  edges <- items$name[named_like(items, "mermaid:edge:")]
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
  # round 2: H6, the 11 px circle on a 1 px line is drawn as a marker shape
  expect_equal(ends("a", "f"), c(head = "none", tail = "none"))
  expect_equal(ends("a", "g"), c(head = "none", tail = "none"))   # step 2: C1, a cross is no Word end
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
    n10 = "parallelogram", n11 = "parallelogram", n12 = "trapezoid", n13 = "trapezoid",
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

test_that("flowchart-full: formerly approximated shapes are drawn from their outlines", {
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

# -- flowchart-handdrawn -------------------------------------------------------

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

# -- Circle, stadium and rounded nodes are single shapes -----------------------

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
  expect_false(any(named_like(items, "mermaid:label")))
  expect_equal(sum(items$kind == "text"), 0L)

  doc <- fixture_doc("base-segoe-stadium-bold-invisible")
  for (id in c("A", "C", "D")) {
    node <- shape_named(doc, node_name(id))
    expect_true(is_wsp_node(node), info = id)
    expect_length(shape_paragraphs(node), 1L)
  }
})

test_that("a transparent classDef fill gives no fill", {
  d <- item_named(fixture_scene("base-segoe-stadium-bold-invisible")$items, node_name("D"))
  expect_true(is.na(d$fill))
  expect_equal(d$stroke, "000000")
})

# -- Other dagre-family diagrams -----------------------------------------------

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

test_that("stateDiagram: states are units named after their state ids", {
  items <- fixture_scene("stateDiagram-v2")$items
  for (id in c("Still", "Moving", "Crash", "Idle", "Busy", "Other")) {
    it <- one_item(items, node_name(id))
    labels <- subtree_rows(items, node_name(id))$label
    expect_true(id %in% labels, info = id)
  }
  expect_identical(item_named(items, "mermaid:edge_label:Still:Moving")$label, "push")
})

test_that("erDiagram: entities are units named after their db id", {
  items <- fixture_scene("erDiagram")$items
  one_item(items, node_name("entity-CUSTOMER-0"))
  one_item(items, node_name("entity-ORDER-1"))
  cust <- subtree_rows(items, node_name("entity-CUSTOMER-0"))
  labels <- paste(cust$label, collapse = "\n")
  for (t in c("CUSTOMER", "custNumber", "PK")) expect_true(grepl(t, labels, fixed = TRUE), info = t)
  expect_true(edge_name("entity-CUSTOMER-0", "entity-ORDER-1") %in% items$name)
})

# -- Units in hand-built renders -----------------------------------------------

test_that("a node joins the db on domId and is named after the db id", {
  r <- one_node_render("squareRect", prim_rect(20, 20, 100, 40), id = "Start")
  items <- build_scene(r)$items
  it <- one_item(items, node_name("Start"))
  expect_identical(it$label, "Label")
})

test_that("a node with no domId joins the db on id (ER, kanban, requirement)", {
  dom <- "entity-A-0"
  prims <- node_prims(dom, prim_rect(20, 20, 100, 40), label = "A")
  db <- mini_db(list(db_node("entity-A-0", shape = "erBox", dom_id = NULL, label = "A")))
  items <- build_scene(mini_render(prims, db = db, type = "er"))$items
  one_item(items, node_name("entity-A-0"))
})

test_that("without a db match the id loses its flowchart- prefix and -<n> suffix", {
  prims <- node_prims("flowchart-my_node-12", prim_rect(20, 20, 100, 40))
  items <- build_scene(mini_render(prims, db = NULL))$items
  one_item(items, node_name("my_node"))
})

test_that("ids with underscores keep their edges and edge labels", {
  nodes <- c(node_prims("flowchart-my_node-0", prim_rect(10, 10, 60, 30), "one"),
             node_prims("flowchart-b-1", prim_rect(130, 10, 60, 30), "two"))
  edge <- prim_edge("M70,25 L130,25", "L_my_node_b_0")
  lbl <- list(
    prim_rect(90, 15, 20, 20, anc = anc_edge_label("L_my_node_b_0"), html = TRUE,
              fill = "#E8E8E8", stroke = NULL),
    prim_text("go", x = 92, y = 17, w = 16, h = 16, anc = anc_edge_label("L_my_node_b_0"))
  )
  db <- mini_db(list(db_node("my_node", dom_id = "flowchart-my_node-0"),
                     db_node("b", dom_id = "flowchart-b-1")),
                list(db_edge("my_node", "b")))
  items <- build_scene(mini_render(c(list(edge), nodes, lbl), db = db))$items
  one_item(items, node_name("my_node"))
  e <- one_item(items, edge_name("my_node", "b"))
  expect_equal(e$tail, "triangle")
  l <- one_item(items, "mermaid:edge_label:my_node:b")
  expect_equal(l$kind, "shape")
  expect_identical(l$label, "go")
})

test_that("a path under edgePaths is an edge even without data-edge, and is never filled", {
  p <- prim_path("M70,25 L130,25 L130,60 Z", id = "L_a_b_0", data = list(id = "L_a_b_0"),
                 anc = anc_edge(), fill = "#FF0000", stroke = "#333333")
  nodes <- c(node_prims("flowchart-a-0", prim_rect(10, 10, 60, 30)),
             node_prims("flowchart-b-1", prim_rect(130, 10, 60, 30)))
  db <- mini_db(list(db_node("a", dom_id = "flowchart-a-0"), db_node("b", dom_id = "flowchart-b-1")),
                list(db_edge("a", "b")))
  items <- build_scene(mini_render(c(list(p), nodes), db = db))$items
  e <- one_item(items, edge_name("a", "b"))
  expect_equal(e$kind, "shape")
  expect_true(is.na(e$fill))
})

test_that("an edge label without a background is a free text item with the label name", {
  nodes <- c(node_prims("flowchart-a-0", prim_rect(10, 10, 60, 30)),
             node_prims("flowchart-b-1", prim_rect(130, 10, 60, 30)))
  edge <- prim_edge("M70,25 L130,25", "L_a_b_0")
  txt <- prim_text("bare", x = 85, y = 30, w = 30, h = 16, anc = anc_edge_label("L_a_b_0"))
  db <- mini_db(list(db_node("a", dom_id = "flowchart-a-0"), db_node("b", dom_id = "flowchart-b-1")),
                list(db_edge("a", "b")))
  items <- build_scene(mini_render(c(list(edge), nodes, list(txt)), db = db))$items
  l <- one_item(items, "mermaid:edge_label:a:b")
  expect_equal(l$kind, "text")
  expect_identical(l$label, "bare")
})

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

test_that("a registered node takes its box from all its prims and paint from the filled one", {
  dom <- "flowchart-X-0"
  prims <- list(
    prim_rect(20, 20, 100, 40, fill = NULL, stroke = "#FF0000", stroke_w = 2, anc = anc_node(dom)),
    prim_rect(25, 25, 110, 30, fill = "#00FF00", stroke = NULL, anc = anc_node(dom)),
    prim_text("Both", x = 50, y = 32, w = 40, h = 16, anc = anc_label(anc_node(dom)))
  )
  db <- mini_db(list(db_node("X", shape = "squareRect", dom_id = dom)))
  items <- build_scene(mini_render(prims, db = db))$items
  it <- one_item(items, node_name("X"))
  expect_equal(it$kind, "shape")
  expect_equal(c(it$x, it$y, it$w, it$h), c(20, 20, 115, 40))
  expect_equal(it$fill, "00FF00")
  expect_identical(it$label, "Both")
})

test_that("an unregistered node drawn by one prim is a single custom shape named after it", {
  pentagon <- prim_poly(list(c(70, 10), c(120, 40), c(100, 90), c(40, 90), c(20, 40)))
  items <- build_scene(one_node_render("notAShape", pentagon, label = "Five"))$items
  it <- one_item(items, node_name("X"))
  expect_equal(it$kind, "shape")
  expect_true(is.na(it$prst))
  expect_identical(it$label, "Five")
})

test_that("an unregistered node drawn by several prims is a group named after it", {
  dom <- "flowchart-X-0"
  prims <- list(
    prim_ellipse(60, 50, 40, anc = anc_node(dom)),
    prim_ellipse(60, 50, 35, anc = anc_node(dom)),
    prim_text("Ring", x = 45, y = 42, w = 30, h = 16, anc = anc_label(anc_node(dom)))
  )
  db <- mini_db(list(db_node("X", shape = "doublecircle", dom_id = dom)))
  items <- build_scene(mini_render(prims, db = db))$items
  expect_equal(one_item(items, node_name("X"))$kind, "group")
  expect_equal(nrow(children_of(items, node_name("X"))), 2L)
})

# -- 3d: the shape registry ----------------------------------------------------

test_that("shape_registry() is one tibble with the documented columns", {
  reg <- shape_registry()
  expect_s3_class(reg, "tbl_df")
  expect_true(all(c("shape", "prst", "flip_h", "flip_v", "adj") %in% names(reg)))
  expect_false(anyDuplicated(reg$shape) > 0L)
  expect_type(reg$flip_h, "logical")
  expect_type(reg$flip_v, "logical")
  expect_true(all(reg$adj %in% c("none", "round", "slant", "hex")))
  expect_false(anyNA(reg$prst))
})

test_that("shape_registry() maps Mermaid shape names and aliases to the same Word figure", {
  reg <- shape_registry()
  row_of <- function(s) reg[reg$shape == s, ]
  expected <- c(
    squareRect = "rect", rect = "rect", process = "rect",
    roundedRect = "roundRect", rounded = "roundRect", event = "roundRect",
    stadium = "roundRect", terminal = "roundRect", pill = "roundRect",
    diamond = "diamond", question = "diamond", diam = "diamond", decision = "diamond",
    hexagon = "hexagon", hex = "hexagon",
    circle = "ellipse", circ = "ellipse",
    cylinder = "can", cyl = "can", db = "can", database = "can",
    lean_right = "parallelogram", "lean-r" = "parallelogram", "lean-right" = "parallelogram",
    lean_left = "parallelogram", "lean-l" = "parallelogram", "lean-left" = "parallelogram",
    trapezoid = "trapezoid", "trap-b" = "trapezoid",
    inv_trapezoid = "trapezoid", "trap-t" = "trapezoid", "inv-trapezoid" = "trapezoid",
    subroutine = "flowChartPredefinedProcess", "fr-rect" = "flowChartPredefinedProcess",
    doc = "flowChartDocument", document = "flowChartDocument",
    docs = "flowChartMultidocument",
    "h-cyl" = "flowChartMagneticDrum", das = "flowChartMagneticDrum",
    "horizontal-cylinder" = "flowChartMagneticDrum",
    "lin-cyl" = "flowChartMagneticDisk", disk = "flowChartMagneticDisk",
    "lined-cylinder" = "flowChartMagneticDisk"
  )
  for (s in names(expected)) {
    r <- row_of(s)
    expect_equal(nrow(r), 1L, info = s)
    expect_equal(r$prst, expected[[s]], info = s)
  }
  expect_true(row_of("lean_left")$flip_h)
  expect_false(row_of("lean_right")$flip_h)
  expect_true(row_of("inv_trapezoid")$flip_v)
  expect_false(row_of("trapezoid")$flip_v)
  expect_equal(row_of("roundedRect")$adj, "round")
  expect_equal(row_of("lean_right")$adj, "slant")
  expect_equal(row_of("trapezoid")$adj, "slant")
  expect_equal(row_of("hexagon")$adj, "hex")
  expect_equal(row_of("diamond")$adj, "none")
})

test_that("shape_registry() leaves out the shapes Word has no true preset for", {
  reg <- shape_registry()
  not_registered <- c("doublecircle", "dbl-circ", "double-circle",
                      "odd", "rect_left_inv_arrow",
                      "lin-doc", "lined-document", "lin-rect", "lined-rectangle", "shadedProcess",
                      "tag-doc", "tagged-document", "tag-rect", "tagged-rectangle",
                      "st-rect", "stacked-rectangle", "fr-circ", "framed-circle", "stateEnd")
  expect_false(any(not_registered %in% reg$shape),
               info = paste(intersect(not_registered, reg$shape), collapse = ", "))
})

test_that("unregistered and unknown shape names fall back to exact geometry", {
  pentagon <- prim_poly(list(c(70, 10), c(120, 40), c(100, 90), c(40, 90), c(20, 40)))
  for (s in c("doublecircle", "odd", "lin-doc", "lin-rect", "tag-doc", "tag-rect",
              "st-rect", "fr-circ", "no-such-shape")) {
    doc <- dml_doc(build_diagram_xml(one_node_render(s, pentagon))$xml)
    node <- shape_named(doc, node_name("X"))
    expect_false(is_missing_node(node), info = s)
    expect_true(shape_is_custom(node), info = s)
    expect_true(is.na(shape_prst(node)), info = s)
  }
})

# Registered nodes in hand-built renders: preset, adjust and flips in the XML.
registry_shape <- function(shape, prim) {
  doc <- dml_doc(build_diagram_xml(one_node_render(shape, prim))$xml)
  shape_named(doc, node_name("X"))
}

test_that("registered shapes give their presets whatever prims drew them", {
  box_path <- prim_path("M20,20 L120,20 L120,60 L20,60 Z")
  cases <- list(
    squareRect = list(prim_rect(20, 20, 100, 40), "rect"),
    diamond = list(prim_poly(list(c(70, 10), c(120, 50), c(70, 90), c(20, 50))), "diamond"),
    circle = list(prim_ellipse(70, 50, 30), "ellipse"),
    cylinder = list(prim_path("M20,25 a50,5 0 0,0 100,0 a50,5 0 0,0 -100,0 l0,30 a50,5 0 0,0 100,0 l0,-30",
                              bbox = c(20, 20, 100, 40)), "can"),
    subroutine = list(box_path, "flowChartPredefinedProcess"),
    doc = list(box_path, "flowChartDocument"),
    "h-cyl" = list(box_path, "flowChartMagneticDrum"),
    "lin-cyl" = list(box_path, "flowChartMagneticDisk")
  )
  for (s in names(cases)) {
    node <- registry_shape(s, cases[[s]][[1]])
    expect_true(is_wsp_node(node), info = s)
    expect_equal(shape_prst(node), cases[[s]][[2]], info = s)
    expect_equal(shape_paragraph_texts(node), "Label", info = s)
  }
})

test_that("a rounded rect's adjust comes from its corner radius, capped at 50000", {
  # round(100000 * 8 / 40) = 20000
  n <- registry_shape("roundedRect", prim_rect(20, 20, 100, 40, rx = 8))
  expect_equal(shape_prst(n), "roundRect")
  expect_equal(shape_adj(n)[["adj"]], 20000)
  n <- registry_shape("roundedRect", prim_rect(20, 20, 100, 40, rx = 30))
  expect_equal(shape_adj(n)[["adj"]], 50000)
  # drawn as a path: the radius of its arcs
  arcs <- prim_path(paste("M28,20 L112,20 A8,8 0 0 1 120,28 L120,52 A8,8 0 0 1 112,60",
                          "L28,60 A8,8 0 0 1 20,52 L20,28 A8,8 0 0 1 28,20 Z"),
                    bbox = c(20, 20, 100, 40))
  n <- registry_shape("roundedRect", arcs)
  expect_equal(shape_prst(n), "roundRect")
  expect_equal(shape_adj(n)[["adj"]], 20000)
})

test_that("a stadium is a roundRect with adjust 50000", {
  pill <- prim_path("M40,20 L100,20 A20,20 0 0 1 100,60 L40,60 A20,20 0 0 1 40,20 Z",
                    bbox = c(20, 20, 100, 40))
  n <- registry_shape("stadium", pill)
  expect_equal(shape_prst(n), "roundRect")
  expect_equal(shape_adj(n)[["adj"]], 50000)
})

test_that("parallelograms and trapezoids take their adjust from the drawn inset", {
  # Box 100 x 40 with a 10 px horizontal inset: adj = round(100000 * 10 / 40).
  pts <- list(
    lean_right    = list(c(20, 60), c(110, 60), c(120, 20), c(30, 20)),
    lean_left     = list(c(30, 60), c(120, 60), c(110, 20), c(20, 20)),
    trapezoid     = list(c(20, 60), c(120, 60), c(110, 20), c(30, 20)),
    inv_trapezoid = list(c(30, 60), c(110, 60), c(120, 20), c(20, 20))
  )
  prst <- c(lean_right = "parallelogram", lean_left = "parallelogram",
            trapezoid = "trapezoid", inv_trapezoid = "trapezoid")
  flip_h <- c(lean_right = FALSE, lean_left = TRUE, trapezoid = FALSE, inv_trapezoid = FALSE)
  flip_v <- c(lean_right = FALSE, lean_left = FALSE, trapezoid = FALSE, inv_trapezoid = TRUE)
  for (s in names(pts)) {
    n <- registry_shape(s, prim_poly(pts[[s]]))
    xf <- shape_xfrm(n)
    expect_equal(shape_prst(n), prst[[s]], info = s)
    expect_equal(shape_adj(n)[["adj"]], 25000, info = s)
    expect_equal(xf$flipH, flip_h[[s]], info = s)
    expect_equal(xf$flipV, flip_v[[s]], info = s)
    expect_near(unname(shape_ext(n)), c(100, 40) * 9525, tol = 2, info = s)
  }
})

test_that("a hexagon takes its adjust from the drawn inset", {
  hex <- prim_poly(list(c(30, 20), c(110, 20), c(120, 40), c(110, 60), c(30, 60), c(20, 40)))
  n <- registry_shape("hexagon", hex)
  expect_equal(shape_prst(n), "hexagon")
  expect_equal(shape_adj(n)[["adj"]], 25000)
})

test_that("flowchart-full: adjust values and flips match the drawn shapes", {
  doc <- fixture_doc("flowchart-full")
  node <- function(id) shape_named(doc, node_name(id))
  # rounded rect: rx 5 on a 54 px high box
  expect_equal(shape_adj(node("n2"))[["adj"]], round(100000 * 5 / 54))
  expect_equal(shape_adj(node("n3"))[["adj"]], 50000)
  # every slanted node has a 19.5 / 39 (or 27 / 54) inset
  for (id in c("n10", "n11", "n12", "n13")) {
    expect_equal(shape_adj(node(id))[["adj"]], 50000, info = id)
  }
  expect_equal(shape_adj(node("n9"))[["adj"]], 25000)   # 9.75 / 39
  expect_false(shape_xfrm(node("n10"))$flipH)
  expect_true(shape_xfrm(node("n11"))$flipH)
  expect_false(shape_xfrm(node("n12"))$flipV)
  expect_true(shape_xfrm(node("n13"))$flipV)
})

test_that("no fill or stroke is forced by shape name", {
  # A fork drawn unfilled stays unfilled; a text-shape rect has what its prims say.
  for (s in c("fork", "f-circ", "junction", "text")) {
    prim <- prim_rect(20, 40, 100, 10, fill = NULL, stroke = "#123456")
    doc <- dml_doc(build_diagram_xml(one_node_render(s, prim))$xml)
    shapes <- as.list(xfind(shape_named(doc, node_name("X")), "descendant-or-self::wps:wsp"))
    fills <- purrr::map(shapes, shape_fill)
    black <- purrr::map_lgl(fills, function(f) identical(f$kind, "solid") &&
                              identical(f$color, "000000"))
    expect_false(any(black), info = s)
  }
})

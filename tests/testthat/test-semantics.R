# Diagram units and the flowchart shape registry (R/semantics.R): what the
# bridge's `unit` field makes a node, an edge or an edge label, and how
# registered shapes become Word presets with their adjust values and flips.

# -- Units are what the bridge's `unit` field says ------------------------------

test_that("a node is named after its unit's id, and takes its db shape by that id", {
  r <- one_node_render("squareRect", prim_rect(20, 20, 100, 40), id = "Start")
  items <- build_scene(r)$items
  it <- one_item(items, node_name("Start"))
  expect_identical(it$label, "Label")
  expect_equal(it$prst, "rect")
})

test_that("a node's name is the unit id even when it looks like a DOM id (ER, kanban)", {
  # ER entities, kanban items and requirements have db ids such as
  # entity-A-0; the unit id is that db id, used as it is.
  prims <- node_prims("e", prim_rect(20, 20, 100, 40), label = "A", id = "entity-A-0")
  db <- mini_db(list(db_node("entity-A-0", shape = "erBox", dom_id = NULL, label = "A")))
  items <- build_scene(mini_render(prims, db = db, type = "er"))$items
  one_item(items, node_name("entity-A-0"))
})

test_that("without a db a unit is still named after its unit id", {
  prims <- node_prims("n", prim_rect(20, 20, 100, 40), id = "my_node")
  items <- build_scene(mini_render(prims, db = NULL))$items
  one_item(items, node_name("my_node"))
})

test_that("ids with underscores keep their edges and edge labels", {
  nodes <- c(node_prims("n1", prim_rect(10, 10, 60, 30), "one", id = "my_node"),
             node_prims("n2", prim_rect(130, 10, 60, 30), "two", id = "b"))
  edge <- prim_edge("M70,25 L130,25", "L_my_node_b_0", start = "my_node", end = "b")
  lbl_anc <- anc_edge_label("L_my_node_b_0", start = "my_node", end = "b")
  lbl <- list(
    prim_rect(90, 15, 20, 20, anc = lbl_anc, html = TRUE, fill = "#E8E8E8", stroke = NULL),
    prim_text("go", x = 92, y = 17, w = 16, h = 16, anc = lbl_anc)
  )
  db <- mini_db(list(db_node("my_node"), db_node("b")), list(db_edge("my_node", "b")))
  items <- build_scene(mini_render(c(list(edge), nodes, lbl), db = db))$items
  one_item(items, node_name("my_node"))
  e <- one_item(items, edge_name("my_node", "b"))
  expect_equal(e$tail, "triangle")
  l <- one_item(items, "mermaid:edge_label:my_node:b")
  expect_equal(l$kind, "shape")
  expect_identical(l$label, "go")
})

test_that("an edge unit is never filled, even when its path is closed and filled", {
  p <- prim_path("M70,25 L130,25 L130,60 Z", anc = anc_edge(), fill = "#FF0000",
                 stroke = "#333333")
  p$unit <- unit_edge("L_a_b_0", "a", "b")
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

# Names that only a unit gets.
unit_names_in <- function(items) {
  items$name[!is.na(items$name) &
               grepl("^mermaid:(node|subgraph|subgraph_bg|edge|edge_label|edge_group):",
                     items$name)]
}

# Every text row of every item.
label_rows <- function(items) unlist(strsplit(items$label[nzchar(items$label)], "\n", fixed = TRUE))

test_that("prims under an element whose unit is a node are that node", {
  root <- list(anc_g(1, cls = "root"))
  a <- c(root, list(anc_g(2, unit = unit_node("N1"))))
  b <- c(root, list(anc_g(3, unit = unit_node("N2"))))
  prims <- list(
    prim_rect(10, 20, 60, 30, anc = a),
    prim_text("One", x = 20, y = 27, w = 40, h = 16, anc = c(a, list(anc_g(4)))),
    prim_rect(110, 20, 60, 30, anc = b),
    prim_text("Two", x = 120, y = 27, w = 40, h = 16, anc = c(b, list(anc_g(5))))
  )
  db <- mini_db(list(db_node("N1"), db_node("N2")))
  items <- build_scene(mini_render(prims, db = db))$items
  # Elements alike in all but their unit are still two nodes.
  expect_identical(one_item(items, node_name("N1"))$label, "One")
  expect_identical(one_item(items, node_name("N2"))$label, "Two")
})

test_that("prims under an element whose unit is a subgraph are its background and title", {
  sg <- list(anc_g(1, cls = "root"), anc_g(5, unit = unit_subgraph("S")))
  prims <- c(
    list(prim_rect(5, 5, 190, 90, anc = sg, fill = "#FFFFDE", stroke = "#AAAA33"),
         prim_text("Sub", x = 85, y = 7, w = 30, h = 16, anchor = "start",
                   anc = c(sg, list(anc_g(6))))),
    node_prims("a", prim_rect(20, 30, 60, 30), id = "A")
  )
  db <- mini_db(list(db_node("S", shape = "rect", dom_id = NULL, is_group = TRUE),
                     db_node("A", parent = "S")))
  items <- build_scene(mini_render(prims, db = db))$items
  expect_equal(one_item(items, "mermaid:subgraph:S")$kind, "group")
  kids <- children_of(items, "mermaid:subgraph:S")
  expect_equal(kids$name[1], "mermaid:subgraph_bg:S")
  expect_identical(item_named(items, "mermaid:subgraph_bg:S")$label, "Sub")
  expect_equal(item_named(items, node_name("A"))$parent, "mermaid:subgraph:S")
})

test_that("an edge label unit is named after the ends its unit gives", {
  # No db edges: the name can only come from the unit's start and end.
  lbl <- list(anc_g(1, cls = "root"),
              anc_g(7, unit = unit_edge_label("L_P_Q_0", "P", "Q")))
  prims <- list(
    prim_rect(10, 10, 180, 60, fill = "#00AA00"),
    prim_rect(90, 30, 30, 20, anc = lbl, html = TRUE, fill = "#E8E8E8", stroke = NULL),
    prim_text("lbl", x = 92, y = 32, w = 26, h = 16, anc = c(lbl, list(anc_g(8))))
  )
  items <- build_scene(mini_render(prims, db = mini_db()))$items
  l <- one_item(items, "mermaid:edge_label:P:Q")
  expect_identical(l$label, "lbl")
})

test_that("an edge path whose prim carries an edge unit is mermaid:edge:<start>-<end>", {
  edge <- prim_path("M10,50 L190,50", anc = list(anc_g(1, cls = "root")), fill = NULL,
                    stroke = "#333333")
  edge$unit <- unit_edge("L_P_Q_0", "P", "Q")
  items <- build_scene(mini_render(list(edge, prim_rect(10, 10, 20, 20)), db = NULL))$items
  e <- one_item(items, edge_name("P", "Q"))
  expect_true(is.na(e$fill))
})

test_that("an element with a unit's class or DOM id but no unit is not a unit", {
  # What Mermaid's DOM marks, without the bridge's unit field: the classes
  # node, cluster, edgeLabel and edgePaths, flowchart-<id>-<n> ids, data-id
  # and data-edge. None of these alone makes a unit.
  root <- list(anc_g(1, cls = "root"))
  node_anc <- function(dom, n) c(root, list(anc_g(2, cls = "nodes"),
                                            anc_g(n, id = dom, cls = "node default")))
  cluster <- c(root, list(anc_g(3, cls = "clusters"),
                          anc_g(4, id = "S", cls = "cluster", data = list(look = "classic"))))
  lbl <- c(root, list(anc_g(20, cls = "edgeLabels"), anc_g(21, cls = "edgeLabel"),
                      anc_g(22, cls = "label", data = list(id = "L_A_B_0"))))
  prims <- list(
    prim_rect(5, 5, 290, 90, anc = cluster, fill = "#FFFFDE"),
    prim_path("M70,40 L130,40", id = "L_A_B_0",
              data = list(edge = "true", et = "edge", id = "L_A_B_0"),
              anc = c(root, list(anc_g(10, cls = "edgePaths"))), fill = NULL,
              stroke = "#333333"),
    prim_text("go", x = 90, y = 42, w = 20, h = 16, anc = lbl),
    prim_rect(10, 25, 60, 30, anc = node_anc("flowchart-A-0", 30)),
    prim_text("A", x = 30, y = 32, w = 20, h = 16,
              anc = c(node_anc("flowchart-A-0", 30), list(anc_g(31, cls = "label")))),
    prim_rect(130, 25, 60, 30, anc = node_anc("flowchart-B-1", 40)),
    prim_text("B", x = 150, y = 32, w = 20, h = 16,
              anc = c(node_anc("flowchart-B-1", 40), list(anc_g(41, cls = "label"))))
  )
  db <- mini_db(list(db_node("S", shape = "rect", dom_id = NULL, is_group = TRUE),
                     db_node("A", parent = "S"), db_node("B", parent = "S")),
                list(db_edge("A", "B")))
  items <- build_scene(mini_render(prims, db = db, viewBox = c(0, 0, 300, 100)))$items
  expect_length(unit_names_in(items), 0L)
  # Everything is still drawn, as loose items.
  expect_true(all(c("A", "B", "go") %in% label_rows(items)))
})

test_that("a diagram with no unit anywhere is drawn from its loose items", {
  prims <- c(
    in_g(list(prim_rect(10, 10, 60, 30), prim_text("one", x = 20, y = 17, w = 40, h = 16)),
         n = 2),
    in_g(list(prim_rect(130, 10, 60, 30), prim_text("two", x = 140, y = 17, w = 40, h = 16)),
         n = 3),
    list(prim_path("M70,25 L130,25", fill = NULL, stroke = "#333333"))
  )
  db <- mini_db(list(db_node("A"), db_node("B")), list(db_edge("A", "B")))
  items <- build_scene(mini_render(prims, db = db))$items
  expect_length(unit_names_in(items), 0L)
  expect_true(all(c("one", "two") %in% label_rows(items)))
  expect_equal(sum(items$kind %in% c("shape", "line")), 3L)
})

# -- Registered and unregistered nodes ------------------------------------------

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

# -- The shape registry ---------------------------------------------------------

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
  # No row flips a shape vertically (Word would turn its text upside down), so
  # the inverted trapezoids have no row and are drawn from their outline.
  expect_false(any(.shape_registry$flip_v))
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

# -- Presets, adjust values and flips -------------------------------------------

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
    trapezoid     = list(c(20, 60), c(120, 60), c(110, 20), c(30, 20))
  )
  prst <- c(lean_right = "parallelogram", lean_left = "parallelogram",
            trapezoid = "trapezoid")
  flip_h <- c(lean_right = FALSE, lean_left = TRUE, trapezoid = FALSE)
  flip_v <- c(lean_right = FALSE, lean_left = FALSE, trapezoid = FALSE)
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

sort_points <- function(m) m[order(m[, 1], m[, 2]), , drop = FALSE]

test_that("an inverted trapezoid is drawn from its outline, upright, with its text", {
  inv <- list(c(30, 60), c(110, 60), c(120, 20), c(20, 20))
  for (shape in c("inv_trapezoid", "trap-t", "manual", "trapezoid-top", "inv-trapezoid")) {
    r <- one_node_render(shape, prim_poly(inv))
    items <- build_scene(r)$items
    it <- items[has_text(items), ]
    expect_equal(nrow(it), 1L, info = shape)
    expect_false(it$flip_v, info = shape)
    expect_true(is.na(it$prst), info = shape)
    path <- it$path[[1]]
    expect_false(is.null(path), info = shape)
    if (is.null(path)) next
    expect_equal(sort_points(outline_points(path)),
                 sort_points(do.call(rbind, lapply(inv, function(p) c(x = p[[1]], y = p[[2]])))),
                 ignore_attr = TRUE, info = shape)
  }
})

test_that("in the XML no vertically flipped shape holds text", {
  inv <- list(c(30, 60), c(110, 60), c(120, 20), c(20, 20))
  doc <- dml_doc(build_diagram_xml(one_node_render("inv_trapezoid", prim_poly(inv)))$xml)
  expect_length(xfind(doc, "//wps:wsp[wps:txbx]/wps:spPr/a:xfrm[@flipV]"), 0L)
  node <- shape_named(doc, "mermaid:node:X")
  expect_true(shape_is_custom(node))
  expect_equal(shape_paragraph_texts(node), "Label")
})


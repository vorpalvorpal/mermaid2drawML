# Icons: with an icon pack installed Mermaid draws `fa:fa-user` in a label,
# an icon node's icon and an architecture service's icon as a nested <svg>.
# Its shapes (`inline` prims) are part of a label or icon, never the unit's
# outline: they do not change a node's preset or paint, never pair or merge
# with its shape, and are drawn as shapes of their own named
# mermaid:icon:<node id> (mermaid:icon:icon outside a node), inside the
# node's group. Nothing else is special: like any loose shapes, the paths
# drawn in one element (the icon's <svg>, or a group inside it) form one
# group by the usual scope rule, nested as the elements nest.

# A flowchart node X: its rect, then a label holding an icon (the paths in
# `icon`, drawn inside a nested <svg>) before its text. Every element has its
# DOM serial, as the bridge reports them.
icon_label_render <- function(icon, shape = "squareRect", rect = prim_rect(20, 20, 160, 50),
                              text_box = c(60, 37, 60, 16)) {
  node_g <- anc_g(10, cls = "node default", unit = unit_node("X"))
  label_g <- anc_g(12, cls = "label")
  svg_g <- anc_g(13)
  rect$anc <- list(node_g)
  rect$n <- 11L
  icon <- purrr::imap(icon, function(p, k) {
    p$anc <- list(node_g, label_g, svg_g)
    p$n <- 13L + k
    p
  })
  text <- prim_text("Label", x = text_box[1], y = text_box[2], w = text_box[3], h = text_box[4],
                    anc = list(node_g, label_g), n = 20L)
  mini_render(c(list(rect), icon, list(text)), viewBox = c(0, 0, 200, 100),
              db = mini_db(list(db_node("X", shape = shape))))
}

# A user icon in two paths (head and body), stroked so that an outline taken
# from it would show.
icon_head <- function() {
  prim_path("M30 32 L40 32 L40 42 L30 42 Z", fill = "#333333", stroke = "#FF00FF", inline = TRUE)
}
icon_body <- function() {
  prim_path("M28 44 L42 44 L42 56 L28 56 Z", fill = "#333333", stroke = "#FF00FF", inline = TRUE)
}

icons_of <- function(items) items[startsWith(items$name, "mermaid:icon"), ]

# -- A label icon in a registered node -----------------------------------------------

test_that("a label icon does not change the node's preset or paint", {
  items <- build_scene(icon_label_render(list(icon_head())))$items
  sub <- drawn(subtree_rows(items, "mermaid:node:X"))
  node <- sub[!startsWith(sub$name, "mermaid:icon"), ]
  expect_equal(nrow(node), 1L)
  expect_equal(node$prst, "rect")
  expect_equal(node$fill, "FFFFFF")
  expect_equal(node$stroke, "000000")
  expect_equal(node$stroke_w, 1)
  expect_identical(node$label, "Label")
  expect_near(unlist(node[1, c("x", "y", "w", "h")]), c(20, 20, 160, 50), tol = 1e-6)
})

test_that("a label icon is a shape of its own, mermaid:icon:<node id>, in the node's group", {
  # A registered node with an icon becomes a group named after the
  # node, holding its shape and the icon.
  items <- build_scene(icon_label_render(list(icon_head())))$items
  icon <- icons_of(items)
  expect_equal(nrow(icon), 1L)
  expect_equal(icon$name, "mermaid:icon:X")
  expect_equal(icon$kind, "shape")
  expect_equal(icon$fill, "333333")
  expect_equal(icon$stroke, "FF00FF")
  expect_identical(icon$label, "")
  expect_near(unlist(icon[1, c("x", "y", "w", "h")]), c(30, 32, 10, 10), tol = 1e-6)
  expect_true(icon$name %in% subtree_rows(items, "mermaid:node:X")$name)
})

test_that("the label's text keeps its own box beside the icon", {
  items <- build_scene(icon_label_render(list(icon_head())))$items
  node <- items[items$label == "Label", ]
  expect_equal(nrow(node), 1L)
  expect_near(node$text[[1]]$box, c(60, 37, 60, 16), tol = 1e-6)
})

test_that("an icon with the node's own box still does not pair or merge with it", {
  # An icon's background square that happens to cover the node exactly.
  bg <- prim_rect(20, 20, 160, 50, fill = "#087EBF", stroke = NULL, inline = TRUE)
  items <- build_scene(icon_label_render(list(bg)))$items
  icon <- icons_of(items)
  expect_equal(nrow(icon), 1L)
  expect_equal(icon$fill, "087EBF")
  sub <- drawn(subtree_rows(items, "mermaid:node:X"))
  node <- sub[!startsWith(sub$name, "mermaid:icon"), ]
  expect_equal(node$fill, "FFFFFF")
  expect_equal(node$stroke, "000000")
})

test_that("the icon paths of one label form one group inside the node", {
  items <- build_scene(icon_label_render(list(icon_head(), icon_body())))$items
  icons <- icons_of(items)
  expect_equal(nrow(icons), 2L)
  expect_true(all(icons$name == "mermaid:icon:X"))
  expect_length(unique(icons$parent), 1L)
  group <- item_named(items, icons$parent[1])
  expect_equal(group$kind, "group")
  # The group's name is a scope group's (mermaid:group:...); its tag
  # is not fixed.
  expect_true(startsWith(group$name, "mermaid:group:"))
  expect_equal(sort(drawn(subtree_rows(items, group$name))$name), sort(icons$name))
  expect_true(group$name %in% subtree_rows(items, "mermaid:node:X")$name)
})

test_that("in the XML the icon is a shape in the node's group, the node keeps its paint", {
  doc <- dml_doc(build_diagram_xml(icon_label_render(list(icon_head(), icon_body())))$xml)
  node <- shape_named(doc, "mermaid:node:X")
  expect_true(is_group_node(node))
  icons <- owners_named(doc, "mermaid:icon:X")
  expect_length(icons, 2L)
  for (ic in as.list(icons)) {
    expect_true("mermaid:node:X" %in% xattr(xfind(ic, "ancestor::wpg:grpSp/wpg:cNvPr"), "name"))
    expect_equal(shape_fill(ic)$color, "333333")
  }
  labelled <- purrr::keep(as.list(xfind(node, "descendant-or-self::wps:wsp")),
                          function(s) identical(shape_paragraph_texts(s), "Label"))
  expect_length(labelled, 1L)
  for (shape in labelled) {
    expect_equal(shape_prst(shape), "rect")
    expect_equal(shape_fill(shape)$color, "FFFFFF")
    expect_equal(shape_ln(shape)$color, "000000")
  }
  expect_length(schema_order_problems(doc), 0L)
})

test_that("an icon outside every node is named mermaid:icon:icon", {
  svg <- in_g(in_g(list(icon_head(), icon_body()), n = 21), n = 20, cls = "service")
  items <- generic_items(c(svg, list(prim_rect(100, 10, 80, 60))))
  icons <- icons_of(items)
  expect_equal(nrow(icons), 2L)
  expect_true(all(icons$name == "mermaid:icon:icon"))
  expect_length(unique(icons$parent), 1L)
  expect_equal(item_named(items, icons$parent[1])$kind, "group")
})

# -- flowchart-full: an icon node --------------------------------------------------

test_that("flowchart-full: the icon node's glyph is mermaid:icon:ic in the node's group", {
  items <- fixture_scene("flowchart-full")$items
  icon <- icons_of(items)
  expect_equal(nrow(icon), 1L)
  expect_equal(icon$name, "mermaid:icon:ic")
  expect_equal(icon$parent, "mermaid:node:ic")
  expect_equal(icon$fill, "9370DB")
  expect_near(unlist(icon[1, c("x", "y", "w", "h")]), c(2596.258, 485.5, 42, 48), tol = 1e-3)
  # The node's own shapes and label are unchanged around it.
  sub <- drawn(subtree_rows(items, "mermaid:node:ic"))
  expect_true("Icon" %in% sub$label)
  expect_true("ECECFF" %in% sub$fill[!startsWith(sub$name, "mermaid:icon")])
})

test_that("flowchart-full XML: the icon is a shape inside the icon node's group", {
  doc <- fixture_doc("flowchart-full")
  icon <- owners_named(doc, "mermaid:icon:ic")
  expect_length(icon, 1L)
  expect_equal(xattr(xfind(icon, "parent::wpg:grpSp/wpg:cNvPr"), "name"), "mermaid:node:ic")
  expect_equal(xattr(xfind(icon, "./wps:spPr/a:solidFill/a:srgbClr"), "val"), "9370DB")
})

# -- architecture-beta: service icons ----------------------------------------------------

test_that("architecture-beta: every inline shape is an icon, each icon's paths in one group", {
  prims <- fixture_prims("architecture-beta")
  inline <- purrr::keep(prims, function(p) isTRUE(p$inline))
  expect_length(inline, 37L)
  items <- fixture_scene("architecture-beta")$items
  icons <- icons_of(drawn(items))
  expect_true(all(icons$name == "mermaid:icon:icon"))
  # No inline shape pairs or merges with another piece: one item each.
  expect_equal(nrow(icons), length(inline))
  expect_false(any(grepl("IconifyId", items$name)))
  expect_false(any(startsWith(items$name[items$kind == "text"], "mermaid:icon")))
  # Each icon belongs to a service (or the group box) element; the scope
  # rule puts all of that element's icon paths inside one group that holds
  # nothing else.
  owner_of <- function(p) {
    a <- purrr::detect(p$anc, function(e) e$cls %in% c("architecture-service",
                                                         "architecture-groups"))
    if (nzchar(a$id)) a$id else a$cls
  }
  owners <- purrr::map_chr(inline, owner_of)
  expect_length(unique(owners), 4L)
  for (o in unique(owners)) {
    holder <- paste0("mermaid:group:", o)
    expect_equal(item_named(items, holder)$kind, "group", info = o)
    in_holder <- icons_of(drawn(subtree_rows(items, holder)))
    expect_equal(nrow(in_holder), sum(owners == o), info = o)
    groups <- items$name[items$kind == "group" & items$name %in% subtree_groups(items, holder)]
    only_icons <- purrr::keep(groups, function(g) {
      d <- drawn(subtree_rows(items, g))
      nrow(d) == sum(owners == o) && all(startsWith(d$name, "mermaid:icon"))
    })
    expect_equal(length(only_icons), 1L, info = o)
  }
})

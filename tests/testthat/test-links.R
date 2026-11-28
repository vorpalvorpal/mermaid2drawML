# Hyperlinks: Mermaid's `click X href` draws the node inside an <a>. Items
# take the href of their nearest <a> ancestor (web and mail links only), and
# build_diagram_xml(links = ) writes it: "shape" makes the shape the click
# target (a:hlinkClick, whose relationship body_add_mermaid() adds), "field"
# wraps the text in a Word HYPERLINK field (self-contained XML), "none"
# leaves links out.

# Flowchart nodes in a row, each a rect and a one-row label, drawn inside an
# <a> when `hrefs` gives it a URL (NA for none).
linked_nodes_render <- function(hrefs, labels = NULL) {
  ids <- names(hrefs)
  labels <- labels %|null|% as.list(stats::setNames(ids, ids))
  prims <- purrr::imap(ids, function(id, k) {
    href <- if (is.na(hrefs[[id]])) NULL else hrefs[[id]]
    anc <- anc_node(paste0("flowchart-", id, "-0"), id = id, href = href)
    x <- 10 + (k - 1) * 120
    rows <- labels[[id]]
    list(prim_rect(x, 20, 100, 50, anc = anc),
         prim_text(rows, x = x + 20, y = 45 - 8 * length(rows), w = 60,
                   h = 16 * length(rows), anc = anc_label(anc), bold = TRUE,
                   color = "#AA0000"))
  })
  nodes <- purrr::map(ids, function(id) db_node(id))
  mini_render(purrr::list_flatten(prims), viewBox = c(0, 0, 10 + 120 * length(ids), 100),
              db = mini_db(nodes))
}

# A free text drawn inside an <a>.
linked_text_render <- function(href, text = "Docs") {
  generic_render(in_g(list(prim_text(text, x = 20, y = 20, w = 60, h = 16)), n = 5,
                      href = href))
}

# Where each a:hlinkClick points, by shape name: tibble(name, rel_id).
shape_links <- function(doc) {
  hl <- xfind(doc, "//wps:cNvPr/a:hlinkClick | //wpg:cNvPr/a:hlinkClick | //pic:cNvPr/a:hlinkClick")
  tibble::tibble(name = xml2::xml_attr(xml2::xml_parent(hl), "name"),
                 rel_id = xattr(hl, "r:id"))
}

# The kinds of a paragraph's runs, in order: a field character's type
# (begin, separate, end), "instr" for field code, "text" for shown text.
field_parts <- function(p) {
  purrr::map_chr(as.list(xfind(p, "./w:r")), function(r) {
    fc <- xattr(xfirst(r, "./w:fldChar"), "w:fldCharType")
    if (!is.na(fc)) return(fc)
    if (!is_missing_node(xfirst(r, "./w:instrText"))) return("instr")
    if (!is_missing_node(xfirst(r, "./w:t"))) return("text")
    "other"
  })
}

# The URL of a paragraph's HYPERLINK field code, unquoted: a quote inside it
# may be written \" (Word's field escape) or %22 (URL escape).
field_url <- function(p) {
  instr <- paste(xml2::xml_text(xfind(p, ".//w:instrText")), collapse = "")
  m <- regmatches(instr, regexec("^\\s*HYPERLINK\\s+\"(.*)\"\\s*$", instr))[[1]]
  if (length(m) < 2L) return(NA_character_)
  gsub("%22", "\"", gsub("\\\\\"", "\"", m[[2]]), fixed = TRUE)
}

drawn <- function(items) items[items$kind != "group", ]

# -- Scene: href ---------------------------------------------------------------

test_that("every item has an href column, NA where nothing is linked", {
  items <- generic_items(list(prim_rect(10, 10, 50, 20)))
  expect_true("href" %in% names(items))
  expect_type(items$href, "character")
  expect_true(is.na(items$href))
})

test_that("a node drawn inside an <a> takes its href; its neighbour stays unlinked", {
  items <- build_scene(linked_nodes_render(c(A = "https://example.com/a", B = NA)))$items
  expect_equal(item_named(items, "mermaid:node:A")$href, "https://example.com/a")
  expect_true(is.na(item_named(items, "mermaid:node:B")$href))
})

test_that("a free text inside an <a> takes its href", {
  items <- generic_items(in_g(list(prim_text("Docs", x = 20, y = 20, w = 60, h = 16)), n = 5,
                              href = "https://example.com/docs"))
  expect_equal(items$kind, "text")
  expect_equal(items$href, "https://example.com/docs")
})

test_that("the nearest <a> gives the href", {
  inner <- in_g(list(prim_text("Docs", x = 20, y = 20, w = 60, h = 16)), n = 6,
                href = "https://inner.example/")
  items <- generic_items(in_g(inner, n = 5, href = "https://outer.example/"))
  expect_equal(items$href, "https://inner.example/")
})

test_that("only http, https and mailto links are kept", {
  kept <- c("https://example.com/x", "http://example.com/y", "mailto:someone@example.com")
  dropped <- c("javascript:alert(1)", "JavaScript:void(0)", "docs/page.html", "#section",
               "ftp://example.com/file", "data:text/html,hi", "")
  for (h in kept) {
    expect_equal(build_scene(linked_text_render(h))$items$href, h, info = h)
  }
  for (h in dropped) {
    expect_true(is.na(build_scene(linked_text_render(h))$items$href), info = h)
  }
})

# -- links = "shape" (the default) ---------------------------------------------------

test_that("a linked node's shape is the click target; links lists the placeholder and URL", {
  r <- linked_nodes_render(c(A = "https://example.com/a", B = NA))
  res <- build_diagram_xml(r)
  expect_true(all(c("xml", "next_id", "links") %in% names(res)))
  expect_s3_class(res$links, "tbl_df")
  expect_true(all(c("rel_id", "url") %in% names(res$links)))
  expect_equal(res$links$url, "https://example.com/a")

  doc <- dml_doc(res$xml)
  a <- shape_named(doc, "mermaid:node:A")
  id <- shape_link_id(a)
  expect_false(is.na(id))
  expect_equal(id, res$links$rel_id)
  expect_true(is.na(shape_link_id(shape_named(doc, "mermaid:node:B"))))
  # The link sits inside the shape's cNvPr, first among its children.
  expect_equal(xml2::xml_name(xfind(a, "./wps:cNvPr/*"))[1], "hlinkClick")
  expect_length(schema_order_problems(doc), 0L)
})

test_that("links = \"shape\" is the default", {
  r <- linked_nodes_render(c(A = "https://example.com/a"))
  expect_identical(build_diagram_xml(r)$xml, build_diagram_xml(r, links = "shape")$xml)
})

test_that("one placeholder per distinct URL; every r:id in the XML is in links", {
  r <- linked_nodes_render(c(A = "https://example.com/one", B = "https://example.com/two",
                             C = "https://example.com/one", D = NA))
  res <- build_diagram_xml(r)
  expect_equal(nrow(res$links), 2L)
  expect_setequal(res$links$url, c("https://example.com/one", "https://example.com/two"))
  expect_false(anyDuplicated(res$links$rel_id) > 0L)
  # Placeholders cannot be mistaken for a document's own rIdN ids.
  expect_false(any(grepl("^rId[0-9]+$", res$links$rel_id)))

  doc <- dml_doc(res$xml)
  sl <- shape_links(doc)
  expect_setequal(sl$name, c("mermaid:node:A", "mermaid:node:B", "mermaid:node:C"))
  expect_true(all(r_attributes(doc)$value %in% res$links$rel_id))
  url_of <- function(name) res$links$url[match(sl$rel_id[sl$name == name], res$links$rel_id)]
  expect_equal(url_of("mermaid:node:A"), "https://example.com/one")
  expect_equal(url_of("mermaid:node:C"), "https://example.com/one")
  expect_equal(url_of("mermaid:node:B"), "https://example.com/two")
})

test_that("every drawn item of a linked unit is a click target", {
  # A node drawn as two separate shapes and a label: a group of drawn items.
  anc <- anc_node("flowchart-N-0", id = "N", href = "https://example.com/n")
  prims <- list(prim_rect(20, 20, 100, 50, anc = anc),
                prim_ellipse(160, 45, 12, anc = anc, fill = "#FF0000"),
                prim_text("Node", x = 50, y = 37, w = 40, h = 16, anc = anc_label(anc)))
  r <- mini_render(prims, viewBox = c(0, 0, 200, 100),
                   db = mini_db(list(db_node("N", shape = "notAShapeMermaidHas"))))
  items <- build_scene(r)$items
  sub <- drawn(subtree_rows(items, "mermaid:node:N"))
  expect_gte(nrow(sub), 2L)
  expect_equal(sub$href, rep("https://example.com/n", nrow(sub)))

  res <- build_diagram_xml(r)
  doc <- dml_doc(res$xml)
  wsps <- xfind(shape_named(doc, "mermaid:node:N"), ".//wps:wsp")
  expect_gte(length(wsps), 2L)
  ids <- purrr::map_chr(as.list(wsps), shape_link_id)
  expect_false(anyNA(ids))
  expect_equal(unique(ids), res$links$rel_id)
})

test_that("a free text inside an <a> is a click target", {
  res <- build_diagram_xml(linked_text_render("https://example.com/docs"))
  doc <- dml_doc(res$xml)
  expect_equal(shape_links(doc)$rel_id, res$links$rel_id)
  expect_equal(res$links$url, "https://example.com/docs")
})

test_that("a dropped href and a diagram without links give no link and an empty links", {
  for (r in list(linked_text_render("javascript:alert(1)"),
                 generic_render(list(prim_rect(10, 10, 50, 20))))) {
    res <- build_diagram_xml(r)
    expect_s3_class(res$links, "tbl_df")
    expect_equal(nrow(res$links), 0L)
    expect_true(all(c("rel_id", "url") %in% names(res$links)))
    expect_length(xfind(dml_doc(res$xml), "//a:hlinkClick"), 0L)
  }
})

# -- links = "field" -----------------------------------------------------------

test_that("links = \"field\" wraps the runs of every paragraph in a HYPERLINK field", {
  r <- linked_nodes_render(c(A = "https://example.com/a"),
                           labels = list(A = c("Line one", "Line two")))
  res <- build_diagram_xml(r, links = "field")
  doc <- dml_doc(res$xml)
  expect_length(xfind(doc, "//a:hlinkClick"), 0L)
  expect_equal(nrow(res$links), 0L)
  paras <- shape_paragraphs(shape_named(doc, "mermaid:node:A"))
  expect_length(paras, 2L)
  for (p in as.list(paras)) {
    parts <- paste(field_parts(p), collapse = " ")
    expect_match(parts, "^begin( instr)+ separate( text)+ end$")
    expect_equal(field_url(p), "https://example.com/a")
  }
  expect_equal(shape_paragraph_texts(shape_named(doc, "mermaid:node:A")),
               c("Line one", "Line two"))
  expect_length(schema_order_problems(doc), 0L)
})

test_that("a field link keeps Mermaid's run formatting", {
  r <- linked_nodes_render(c(A = "https://example.com/a"))
  runs_in <- function(mode) {
    doc <- dml_doc(build_diagram_xml(r, links = mode)$xml)
    rr <- shape_runs(shape_named(doc, "mermaid:node:A"))
    rr[nzchar(rr$text), c("text", "bold", "italic", "underline", "color", "sz", "ascii")]
  }
  expect_equal(runs_in("field"), runs_in("none"))
  expect_true(all(runs_in("field")$bold))
  expect_equal(unique(runs_in("field")$color), "AA0000")
})

test_that("a field URL is quoted and XML-escaped", {
  url <- "https://example.com/q?a=1&b=\"two\""
  r <- linked_nodes_render(c(A = url))
  res <- build_diagram_xml(r, links = "field")
  expect_no_error(doc <- dml_doc(res$xml))
  paras <- shape_paragraphs(shape_named(doc, "mermaid:node:A"))
  expect_equal(field_url(paras[[1]]), url)
})

test_that("a free text inside an <a> gets a field link too", {
  doc <- dml_doc(build_diagram_xml(linked_text_render("mailto:someone@example.com"),
                                   links = "field")$xml)
  p <- xfirst(doc, "//w:txbxContent/w:p")
  expect_equal(field_url(p), "mailto:someone@example.com")
})

test_that("in field mode a linked shape without text gets no link", {
  r <- generic_render(in_g(list(prim_rect(20, 20, 60, 30)), n = 5, href = "https://example.com/"))
  res <- build_diagram_xml(r, links = "field")
  doc <- dml_doc(res$xml)
  expect_length(xfind(doc, "//w:instrText"), 0L)
  expect_length(xfind(doc, "//a:hlinkClick"), 0L)
  expect_equal(nrow(res$links), 0L)
})

# -- links = "none" ------------------------------------------------------------

test_that("links = \"none\" leaves links out", {
  r <- linked_nodes_render(c(A = "https://example.com/a"))
  res <- build_diagram_xml(r, links = "none")
  doc <- dml_doc(res$xml)
  expect_length(xfind(doc, "//a:hlinkClick"), 0L)
  expect_length(xfind(doc, "//w:instrText"), 0L)
  expect_equal(nrow(res$links), 0L)
  expect_length(r_attributes(doc)$value, 0L)
})

test_that("links must be one of shape, field or none", {
  r <- linked_nodes_render(c(A = "https://example.com/a"))
  for (mode in c("shape", "field", "none")) expect_no_error(build_diagram_xml(r, links = mode))
  expect_error(build_diagram_xml(r, links = "bogus"))
})

# -- flowchart-full: click ... href ----------------------------------------------

test_that("flowchart-full: nodes a and b carry their click hrefs, nothing else does", {
  items <- drawn(fixture_scene("flowchart-full")$items)
  expect_equal(item_named(items, "mermaid:node:a")$href, "https://example.com/")
  expect_equal(item_named(items, "mermaid:node:b")$href, "https://example.org/")
  others <- items[!items$name %in% c("mermaid:node:a", "mermaid:node:b"), ]
  expect_equal(others$href, rep(NA_character_, nrow(others)))
})

test_that("flowchart-full XML: a and b link to their URLs through links", {
  res <- fixture_xml("flowchart-full")
  doc <- dml_doc(res$xml)
  sl <- shape_links(doc)
  expect_setequal(sl$name, c("mermaid:node:a", "mermaid:node:b"))
  expect_equal(nrow(res$links), 2L)
  url_of <- function(name) res$links$url[match(sl$rel_id[sl$name == name], res$links$rel_id)]
  expect_equal(url_of("mermaid:node:a"), "https://example.com/")
  expect_equal(url_of("mermaid:node:b"), "https://example.org/")
})

# -- The wrappers ----------------------------------------------------------------

test_that("mermaid_to_word() returns links beside xml, and passes links = through", {
  local_mock_render("flowchart-full")
  d <- quietly_expected(mermaid_to_word("graph TD"))
  expect_true("links" %in% names(d))
  expect_setequal(d$links$url, c("https://example.com/", "https://example.org/"))
  at <- r_attributes(dml_doc(d$xml))
  expect_gt(sum(at$element == "hlinkClick"), 0L)
  expect_true(all(at$value[at$element == "hlinkClick"] %in% d$links$rel_id))
  none <- quietly_expected(mermaid_to_word("graph TD", links = "none"))
  expect_equal(nrow(none$links), 0L)
  expect_length(xfind(dml_doc(none$xml), "//a:hlinkClick"), 0L)
})

test_that("mermaid_to_word() on a list returns links per diagram", {
  renders <- list(linked_nodes_render(c(A = "https://one.example/")),
                  linked_nodes_render(c(A = "https://two.example/")))
  testthat::local_mocked_bindings(render_mermaid = function(...) renders)
  ds <- mermaid_to_word(list("graph TD", "graph TD"))
  expect_length(ds, 2L)
  expect_equal(ds[[1]]$links$url, "https://one.example/")
  expect_equal(ds[[2]]$links$url, "https://two.example/")
})

test_that("a saved document resolves every link to an external hyperlink of its URL", {
  local_mock_render("flowchart-full")
  doc <- quietly_expected(body_add_mermaid(officer::read_docx(), "graph TD"))
  pkg <- saved_docx(doc)
  problems <- relationship_problems(pkg)
  expect(length(problems) == 0L, paste(problems, collapse = "\n"))
  sl <- shape_links(pkg$doc)
  expect_setequal(sl$name, c("mermaid:node:a", "mermaid:node:b"))
  target <- function(name) pkg$rels$target[match(sl$rel_id[sl$name == name], pkg$rels$id)]
  expect_equal(target("mermaid:node:a"), "https://example.com/")
  expect_equal(target("mermaid:node:b"), "https://example.org/")
})

test_that("two diagrams with links in one document keep their own URLs", {
  doc <- officer::read_docx()
  local_mock_render(linked_nodes_render(c(A = "https://one.example/")))
  doc <- body_add_mermaid(doc, "graph TD")
  local_mock_render(linked_nodes_render(c(A = "https://two.example/")))
  doc <- body_add_mermaid(doc, "graph TD")
  pkg <- saved_docx(doc)
  problems <- relationship_problems(pkg)
  expect(length(problems) == 0L, paste(problems, collapse = "\n"))
  sl <- shape_links(pkg$doc)
  expect_equal(sl$name, c("mermaid:node:A", "mermaid:node:A"))
  expect_equal(pkg$rels$target[match(sl$rel_id, pkg$rels$id)],
               c("https://one.example/", "https://two.example/"))
})

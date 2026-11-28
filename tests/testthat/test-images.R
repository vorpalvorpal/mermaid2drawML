# Embedded images: an SVG <image> or an HTML <img> whose picture the browser
# could read (`png`, a PNG data URI) becomes an item of kind "image" and a
# pic:pic in the drawing. A picture needs a part in the document package, so
# build_diagram_xml() returns `images` (a placeholder rel_id and the PNG
# bytes) and body_add_mermaid() adds the parts. An image without a picture,
# or with embed_images = FALSE, is left out with the unsupported warning.

EMU_PX <- 9525

# The PNG bytes an image item holds, from whichever list column holds them.
item_png <- function(row) {
  vals <- purrr::map(names(row), function(col) {
    v <- row[[col]]
    if (is.list(v)) v[[1]] else NULL
  })
  hits <- purrr::keep(vals, is.raw)
  if (length(hits) == 1L) hits[[1]] else NULL
}

images_of <- function(items) items[items$kind == "image", ]

# A rotation by `deg` (clockwise on screen) and a translation, as an SVG
# matrix c(a, b, c, d, e, f).
rot_m <- function(deg, e = 0, f = 0, s = 1) {
  r <- deg * pi / 180
  c(s * cos(r), s * sin(r), -s * sin(r), s * cos(r), e, f)
}

muffle_small <- function(expr) {
  withCallingHandlers(expr, mermaid2drawml_small_text = function(c) rlang::cnd_muffle(c))
}

# -- Scene -----------------------------------------------------------------------

test_that("an SVG image with a picture is an image item at its box holding the PNG", {
  sc <- build_scene(generic_render(list(
    prim_image(60, 20, 30, 40, png = png_1x1_uri(), id = "logo"))))
  img <- images_of(sc$items)
  expect_equal(nrow(img), 1L)
  expect_near(unlist(img[1, c("x", "y", "w", "h")]), c(60, 20, 30, 40), tol = 1e-6)
  expect_equal(img$rot, 0)
  expect_equal(img$name, "mermaid:image:logo")
  expect_identical(item_png(img[1, ]), png_bytes(png_1x1_uri()))
  expect_false("image" %in% sc$unsupported)
})

test_that("an image is named after its id, else its first class, else as an image", {
  name_of <- function(...) {
    images_of(generic_items(list(prim_image(10, 10, 20, 20, png = png_1x1_uri(), ...))))$name
  }
  expect_equal(name_of(id = "logo"), "mermaid:image:logo")
  expect_equal(name_of(cls = "pic big"), "mermaid:image:pic")
  # With neither there is no tag, as for any other item.
  expect_equal(name_of(), "mermaid:image")
})

test_that("an HTML <img> with a picture is an image item at its root box", {
  img <- images_of(generic_items(list(
    prim_image(40, 30, 24, 18, png = png_1x1_uri(), html = TRUE))))
  expect_equal(nrow(img), 1L)
  expect_near(unlist(img[1, c("x", "y", "w", "h")]), c(40, 30, 24, 18), tol = 1e-6)
  expect_equal(img$rot, 0)
})

test_that("a rotated image's box and rotation follow from its own box through its matrix", {
  # Own box 10, 5, 40 x 20, turned 45 degrees and moved: like a rect, the
  # item is that box at the matrix's scale, centred where the matrix puts
  # its centre, turned by the matrix's angle atan2(b, a).
  m <- rot_m(45, 100, 50)
  img <- images_of(generic_items(list(prim_image(10, 5, 40, 20, m = m, png = png_1x1_uri())),
                                 viewBox = c(0, 0, 200, 140)))
  centre <- apply_m(m, 10 + 40 / 2, 5 + 20 / 2)
  sx <- sqrt(m[1]^2 + m[2]^2)
  sy <- sqrt(m[3]^2 + m[4]^2)
  expect_near(img$rot, atan2(m[2], m[1]) * 180 / pi, tol = 0.01)
  expect_near(c(img$w, img$h), c(40 * sx, 20 * sy), tol = 0.01)
  expect_near(c(img$x + img$w / 2, img$y + img$h / 2), c(centre$x, centre$y), tol = 0.01)
  # The same box drawn as a rect through the same matrix.
  rect <- generic_items(list(prim_rect(10, 5, 40, 20, m = m)), viewBox = c(0, 0, 200, 140))
  expect_near(unlist(img[1, c("x", "y", "w", "h", "rot")]),
              unlist(rect[1, c("x", "y", "w", "h", "rot")]), tol = 0.01)
})

test_that("a scaled image takes the matrix's scale", {
  m <- c(2, 0, 0, 2, 10, 10)
  img <- images_of(generic_items(list(prim_image(0, 0, 30, 20, m = m, png = png_1x1_uri()))))
  expect_near(unlist(img[1, c("x", "y", "w", "h")]), c(10, 10, 60, 40), tol = 1e-6)
})

test_that("an image without a picture is no item and is listed as unsupported", {
  sc <- build_scene(generic_render(list(prim_rect(0, 0, 50, 50),
                                        prim_image(60, 0, 30, 30, png = NULL))))
  expect_equal(nrow(images_of(sc$items)), 0L)
  expect_true("image" %in% sc$unsupported)
})

test_that("an image in a node is in the node's group", {
  anc <- anc_node("flowchart-P-0", id = "P")
  prims <- list(prim_rect(20, 20, 80, 60, anc = anc),
                prim_image(40, 25, 40, 30, png = png_1x1_uri(), anc = anc),
                prim_text("Photo", x = 40, y = 60, w = 40, h = 16, anc = anc_label(anc)))
  r <- mini_render(prims, db = mini_db(list(db_node("P", shape = "imageSquare"))))
  items <- build_scene(r)$items
  sub <- subtree_rows(items, "mermaid:node:P")
  expect_equal(sum(sub$kind == "image"), 1L)
})

test_that("an image drawn in one element with a shape is in that element's group", {
  prims <- in_g(list(prim_rect(10, 10, 60, 60), prim_image(20, 20, 40, 40, png = png_1x1_uri())),
                n = 7, cls = "person-man")
  items <- generic_items(c(prims, list(prim_rect(120, 10, 60, 60))))
  img <- images_of(items)
  expect_equal(img$parent, "mermaid:group:person-man")
})

# -- XML ------------------------------------------------------------------------

test_that("an image is a pic:pic at its place, its blip a placeholder listed in images", {
  bytes <- png_bytes(png_c4_uri())
  r <- generic_render(list(prim_image(60, 20, 30, 40, png = png_c4_uri(), id = "logo")))
  res <- build_diagram_xml(r)
  expect_true(all(c("xml", "next_id", "links", "images") %in% names(res)))
  expect_s3_class(res$images, "tbl_df")
  expect_true(all(c("rel_id", "bytes") %in% names(res$images)))
  expect_equal(nrow(res$images), 1L)
  expect_type(res$images$bytes, "list")
  expect_identical(res$images$bytes[[1]], bytes)
  expect_false(grepl("^rId[0-9]+$", res$images$rel_id))

  doc <- dml_doc(res$xml)
  pics <- all_pics(doc)
  expect_length(pics, 1L)
  pic <- if (length(pics) > 0L) pics[[1]] else xml_missing_node()
  expect_equal(xml2::xml_name(xml2::xml_parent(pic)), "wgp")
  expect_equal(shape_name(pic), "mermaid:image:logo")
  expect_equal(pic_embed(pic), res$images$rel_id)
  expect_false(is_missing_node(xfirst(pic, "./pic:nvPicPr/pic:cNvPicPr")))
  expect_false(is_missing_node(xfirst(pic, "./pic:blipFill/a:stretch/a:fillRect")))
  expect_equal(xattr(xfirst(pic, "./pic:spPr/a:prstGeom"), "prst"), "rect")
  expect_near(abs_box(pic), c(60, 20, 30, 40) * EMU_PX, tol = 1)
  expect_length(schema_order_problems(doc), 0L)
})

test_that("a picture's id is one of the drawing's ids, and next_id follows them all", {
  r <- generic_render(list(prim_rect(0, 0, 50, 50),
                           prim_image(60, 0, 30, 30, png = png_1x1_uri())))
  res <- build_diagram_xml(r, start_id = 7L)
  ids <- all_ids(dml_doc(res$xml))
  expect_length(ids, 3L)
  expect_false(anyDuplicated(ids) > 0L)
  expect_equal(sort(ids), 7:9)
  expect_equal(res$next_id, 10L)
})

test_that("a picture keeps its place in paint order", {
  r <- generic_render(list(prim_rect(0, 0, 50, 50, id = "under"),
                           prim_image(20, 10, 30, 30, png = png_1x1_uri(), id = "pic"),
                           prim_rect(40, 20, 50, 50, id = "over")))
  doc <- dml_doc(build_diagram_xml(r)$xml)
  members <- xml2::xml_children(root_group(doc))
  members <- members[xml2::xml_name(members) %in% c("wsp", "grpSp", "pic")]
  expect_equal(purrr::map_chr(as.list(members), shape_name),
               c("mermaid:shape:under", "mermaid:image:pic", "mermaid:shape:over"))
})

test_that("a rotated image is a rotated picture", {
  m <- rot_m(45, 100, 50)
  doc <- dml_doc(build_diagram_xml(generic_render(
    list(prim_image(10, 5, 40, 20, m = m, png = png_1x1_uri())), viewBox = c(0, 0, 200, 140)))$xml)
  pics <- all_pics(doc)
  expect_length(pics, 1L)
  pic <- if (length(pics) > 0L) pics[[1]] else xml_missing_node()
  expect_near(shape_xfrm(pic)$rot, atan2(m[2], m[1]) * 180 / pi * 60000, tol = 60)
  centre <- apply_m(m, 10 + 40 / 2, 5 + 20 / 2)
  box <- abs_box(pic)
  expect_near(box[c("w", "h")], c(40, 20) * EMU_PX, tol = 2)
  expect_near(c(box[["x"]] + box[["w"]] / 2, box[["y"]] + box[["h"]] / 2),
              c(centre$x, centre$y) * EMU_PX, tol = 2)
})

test_that("identical pictures share one rel_id; different ones get their own", {
  same <- build_diagram_xml(generic_render(list(
    prim_image(10, 10, 30, 30, png = png_1x1_uri()),
    prim_image(60, 10, 30, 30, png = png_1x1_uri()))))
  expect_equal(nrow(same$images), 1L)
  embeds <- purrr::map_chr(as.list(all_pics(dml_doc(same$xml))), pic_embed)
  expect_equal(embeds, rep(same$images$rel_id, 2L))

  diff <- build_diagram_xml(generic_render(list(
    prim_image(10, 10, 30, 30, png = png_1x1_uri()),
    prim_image(60, 10, 30, 30, png = png_c4_uri()))))
  expect_equal(nrow(diff$images), 2L)
  expect_false(anyDuplicated(diff$images$rel_id) > 0L)
  embeds <- purrr::map_chr(as.list(all_pics(dml_doc(diff$xml))), pic_embed)
  expect_equal(diff$images$bytes[match(embeds, diff$images$rel_id)],
               list(png_bytes(png_1x1_uri()), png_bytes(png_c4_uri())))
})

test_that("image and link placeholders never collide", {
  prims <- c(list(prim_image(10, 10, 30, 30, png = png_1x1_uri())),
             in_g(list(prim_text("Docs", x = 60, y = 20, w = 60, h = 16)), n = 5,
                  href = "https://example.com/"))
  res <- build_diagram_xml(generic_render(prims))
  expect_equal(nrow(res$images), 1L)
  expect_equal(nrow(res$links), 1L)
  expect_length(intersect(res$images$rel_id, res$links$rel_id), 0L)
  expect_true(all(r_attributes(dml_doc(res$xml))$value %in%
                    c(res$images$rel_id, res$links$rel_id)))
})

test_that("embed_images = FALSE leaves pictures out with the unsupported warning", {
  r <- generic_render(list(prim_rect(0, 0, 50, 50),
                           prim_image(60, 0, 30, 30, png = png_1x1_uri())))
  cnd <- rlang::catch_cnd(res <- build_diagram_xml(r, embed_images = FALSE),
                          classes = "mermaid2drawml_unsupported")
  expect_s3_class(cnd, "mermaid2drawml_unsupported")
  expect_true("image" %in% cnd$tags)
  res <- suppressWarnings(build_diagram_xml(r, embed_images = FALSE))
  expect_length(all_pics(dml_doc(res$xml)), 0L)
  expect_equal(nrow(res$images), 0L)
  expect_length(all_wsp(dml_doc(res$xml)), 1L)
})

test_that("embed_images = TRUE is the default", {
  r <- generic_render(list(prim_image(60, 0, 30, 30, png = png_1x1_uri())))
  expect_identical(build_diagram_xml(r)$xml, build_diagram_xml(r, embed_images = TRUE)$xml)
})

test_that("an image without a picture is left out with a warning naming image", {
  r <- generic_render(list(prim_rect(0, 0, 50, 50),
                           prim_image(60, 0, 30, 30, png = NULL, html = TRUE)))
  cnd <- rlang::catch_cnd(build_diagram_xml(r), classes = "mermaid2drawml_unsupported")
  expect_true("image" %in% cnd$tags)
  res <- suppressWarnings(build_diagram_xml(r))
  expect_length(all_pics(dml_doc(res$xml)), 0L)
  expect_equal(nrow(res$images), 0L)
})

test_that("a diagram with pictures and nothing unsupported does not warn", {
  r <- generic_render(list(prim_image(60, 0, 30, 30, png = png_1x1_uri())))
  expect_no_warning(build_diagram_xml(r))
})

test_that("a diagram without images returns an empty images table", {
  res <- build_diagram_xml(generic_render(list(prim_rect(10, 10, 50, 20))))
  expect_s3_class(res$images, "tbl_df")
  expect_equal(nrow(res$images), 0L)
  expect_true(all(c("rel_id", "bytes") %in% names(res$images)))
})

# -- Fixtures ----------------------------------------------------------------------

test_that("C4Context: the person picture is an image in its person group", {
  sc <- fixture_scene("C4Context")
  expect_false("image" %in% sc$unsupported)
  img <- images_of(sc$items)
  expect_equal(nrow(img), 1L)
  expect_equal(img$parent, "mermaid:group:person-man")
  expect_near(unlist(img[1, c("x", "y", "w", "h")]), c(234, 197, 48, 48), tol = 1e-3)
  expect_identical(item_png(img[1, ]), png_bytes(png_c4_uri()))

  res <- fixture_xml("C4Context")
  doc <- dml_doc(res$xml)
  pic <- all_pics(doc)
  expect_length(pic, 1L)
  expect_equal(xattr(xfind(pic, "parent::wpg:grpSp/wpg:cNvPr"), "name"),
               "mermaid:group:person-man")
  expect_equal(res$images$bytes, list(png_bytes(png_c4_uri())))
  expect_length(schema_order_problems(doc), 0L)
})

test_that("C4Context and flowchart-full warn about nothing left out", {
  for (name in c("C4Context", "flowchart-full")) {
    expect_no_warning(muffle_small(build_diagram_xml(fixture_render(name))))
  }
})

test_that("flowchart-full: the image node's picture is in the node's group", {
  items <- fixture_scene("flowchart-full")$items
  sub <- subtree_rows(items, "mermaid:node:im")
  img <- sub[sub$kind == "image", ]
  expect_equal(nrow(img), 1L)
  expect_near(unlist(img[1, c("x", "y", "w", "h")]), c(2587.258, 663, 60, 60), tol = 1e-3)
  expect_true(startsWith(img$name, "mermaid:image"))
  pic <- all_pics(fixture_doc("flowchart-full"))
  expect_length(pic, 1L)
  expect_true("mermaid:node:im" %in% xattr(xfind(pic, "ancestor::wpg:grpSp/wpg:cNvPr"), "name"))
})

# -- The wrappers ----------------------------------------------------------------------

test_that("mermaid_to_word() returns images beside xml, and passes embed_images through", {
  local_mock_render("C4Context")
  d <- quietly_expected(mermaid_to_word("graph TD"))
  expect_true("images" %in% names(d))
  expect_equal(nrow(d$images), 1L)
  expect_true(is_png(d$images$bytes[[1]]))
  embeds <- r_attributes(dml_doc(d$xml))
  expect_equal(embeds$value[embeds$attr == "embed"], d$images$rel_id)
  off <- quietly_expected(mermaid_to_word("graph TD", embed_images = FALSE))
  expect_equal(nrow(off$images), 0L)
  expect_length(all_pics(dml_doc(off$xml)), 0L)
})

test_that("mermaid_to_word() on a list returns images per diagram", {
  renders <- list(generic_render(list(prim_image(10, 10, 30, 30, png = png_1x1_uri()))),
                  generic_render(list(prim_image(10, 10, 30, 30, png = png_c4_uri()))))
  testthat::local_mocked_bindings(render_mermaid = function(...) renders)
  ds <- mermaid_to_word(list("graph TD", "graph TD"))
  expect_equal(ds[[1]]$images$bytes, list(png_bytes(png_1x1_uri())))
  expect_equal(ds[[2]]$images$bytes, list(png_bytes(png_c4_uri())))
})

test_that("a saved document holds each picture as a PNG part its blip points to", {
  local_mock_render("C4Context")
  doc <- quietly_expected(body_add_mermaid(officer::read_docx(), "graph TD"))
  pkg <- saved_docx(doc)
  problems <- relationship_problems(pkg)
  expect(length(problems) == 0L, paste(problems, collapse = "\n"))
  pic <- all_pics(pkg$doc)
  expect_length(pic, 1L)
  target <- pkg$rels$target[match(xattr(xfind(pic, "./pic:blipFill/a:blip"), "r:embed")[1],
                                  pkg$rels$id)]
  expect_identical(rel_part_bytes(pkg, target), png_bytes(png_c4_uri()))
  expect_true(file.exists(file.path(pkg$dir, "word", target)))
  expect_true(startsWith(target, "media/"))
  # The package declares the PNG content type.
  ct <- xml2::read_xml(file.path(pkg$dir, "[Content_Types].xml"))
  defaults <- xml2::xml_find_all(ct, "//*[local-name()='Default']")
  overrides <- xml2::xml_find_all(ct, "//*[local-name()='Override']")
  png_declared <- any(tolower(xml2::xml_attr(defaults, "Extension")) == "png") ||
    any(xml2::xml_attr(overrides, "PartName") == paste0("/word/", target))
  expect_true(png_declared)
})

test_that("two diagrams with different pictures in one document keep their own", {
  doc <- officer::read_docx()
  local_mock_render(generic_render(list(prim_image(10, 10, 30, 30, png = png_1x1_uri()))))
  doc <- body_add_mermaid(doc, "graph TD")
  local_mock_render(generic_render(list(prim_image(10, 10, 30, 30, png = png_c4_uri()))))
  doc <- body_add_mermaid(doc, "graph TD")
  local_mock_render(generic_render(list(prim_image(10, 10, 30, 30, png = png_1x1_uri()))))
  doc <- body_add_mermaid(doc, "graph TD")
  pkg <- saved_docx(doc)
  problems <- relationship_problems(pkg)
  expect(length(problems) == 0L, paste(problems, collapse = "\n"))
  embeds <- purrr::map_chr(as.list(all_pics(pkg$doc)), pic_embed)
  expect_length(embeds, 3L)
  bytes <- purrr::map(pkg$rels$target[match(embeds, pkg$rels$id)], rel_part_bytes, pkg = pkg)
  expect_identical(bytes, list(png_bytes(png_1x1_uri()), png_bytes(png_c4_uri()),
                               png_bytes(png_1x1_uri())))
})

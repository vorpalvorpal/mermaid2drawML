# The whole pipeline into a saved Word document: every fixture through
# body_add_mermaid() (render mocked) into one document, then the package's
# document.xml read back. Word refuses a file whose drawings break the
# schema's child order or reuse an id, so these checks guard the file as Word
# opens it, not one diagram at a time.

# One document holding every fixture, saved, and its parts read back:
# word/document.xml parsed (`doc`), its relationships (`rels`) and what is
# wrong with them (`problems`, read while the unzipped package still exists).
all_fixtures_package <- function() {
  cached("document:all-fixtures", function() {
    doc <- officer::read_docx()
    for (name in all_fixture_names()) {
      local_mock_render(name)
      doc <- quietly_expected(body_add_mermaid(doc, "graph TD"))
      doc <- officer::body_add_par(doc, name)
    }
    pkg <- saved_docx(doc)
    list(doc = pkg$doc, rels = pkg$rels, problems = relationship_problems(pkg))
  })
}

all_fixtures_document <- function() all_fixtures_package()$doc

test_that("a document with every fixture parses and holds a drawing for each", {
  expect_no_error(doc <- all_fixtures_document())
  expect_s3_class(doc, "xml_document")
  expect_length(xfind(doc, "//wp:inline"), length(all_fixture_names()))
})

test_that("every drawing and shape id is unique across the whole document", {
  doc <- all_fixtures_document()
  doc_pr_ids <- xml2::xml_attr(xfind(doc, "//wp:docPr"), "id")
  cnv_ids <- xml2::xml_attr(xml2::xml_find_all(doc, "//*[local-name()='cNvPr']"), "id")
  ids <- c(doc_pr_ids, cnv_ids)
  expect_false(anyNA(ids))
  expect_equal(length(doc_pr_ids), length(all_fixture_names()))
  dups <- unique(ids[duplicated(ids)])
  expect_length(dups, 0L)
})

test_that("children appear in the order the schema requires", {
  doc <- all_fixtures_document()
  problems <- schema_order_problems(doc)
  expect(length(problems) == 0L,
         paste(c("Children out of schema order:", utils::head(problems, 20L)),
               collapse = "\n"))
})

test_that("the order check finds a child out of place and one it does not know", {
  # The check itself: a:ext before a:off, and an unknown child of a:ln.
  bad <- dml_doc(paste0(
    "<w:document xmlns:w='", dml_ns[["w"]], "' xmlns:a='", dml_ns[["a"]], "'>",
    "<a:xfrm><a:ext cx='1' cy='1'/><a:off x='0' y='0'/></a:xfrm>",
    "<a:ln><a:solidFill/><a:bogus/></a:ln>",
    "</w:document>"
  ))
  problems <- schema_order_problems(bad)
  expect_true(any(grepl("a:xfrm: a:off comes after a:ext", problems, fixed = TRUE)))
  expect_true(any(grepl("a:ln: unexpected child a:bogus", problems, fixed = TRUE)))
})

test_that("every r:embed and r:id in the document resolves to a relationship of its kind", {
  # Pictures need an image part, links an external hyperlink; a placeholder
  # left without its relationship would make Word refuse the file.
  pkg <- all_fixtures_package()
  expect(length(pkg$problems) == 0L,
         paste(c("Unresolved relationships:", utils::head(pkg$problems, 20L)), collapse = "\n"))
  # The fixtures hold pictures (C4Context, flowchart-full) and links
  # (flowchart-full), so the check has something to resolve.
  pictures <- sum(purrr::map_int(all_fixture_names(), function(name) {
    sum(purrr::map_lgl(fixture_prims(name, "image"), function(p) !is.null(p$png)))
  }))
  expect_length(all_pics(pkg$doc), pictures)
  expect_gte(length(xfind(pkg$doc, "//a:hlinkClick")), 2L)
})

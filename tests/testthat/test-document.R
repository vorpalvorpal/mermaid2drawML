# The whole pipeline into a saved Word document: every fixture through
# body_add_mermaid() (render mocked) into one document, then the package's
# document.xml read back. Word refuses a file whose drawings break the
# schema's child order or reuse an id, so these checks guard the file as Word
# opens it, not one diagram at a time.

# One document holding every fixture, saved; its word/document.xml, parsed.
all_fixtures_document <- function() {
  cached("document:all-fixtures", function() {
    doc <- officer::read_docx()
    for (name in all_fixture_names()) {
      local_mock_render(name)
      doc <- quietly_expected(body_add_mermaid(doc, "graph TD"))
      doc <- officer::body_add_par(doc, name)
    }
    path <- withr::local_tempfile(fileext = ".docx")
    print(doc, target = path)
    dir <- withr::local_tempdir()
    utils::unzip(path, files = "word/document.xml", exdir = dir)
    xml2::read_xml(file.path(dir, "word", "document.xml"))
  })
}

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

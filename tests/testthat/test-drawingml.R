test_that("build_diagram_xml returns list with xml and next_id", {
  g <- parse_mermaid("flowchart TB\n A[Start]-->B[End]") |> calculate_layout()
  r <- build_diagram_xml(g)
  expect_type(r, "list")
  expect_named(r, c("xml", "next_id"))
  expect_type(r$xml, "character")
})

test_that("xml is valid XML", {
  g <- parse_mermaid("flowchart TB\n A[Start]-->B[End]") |> calculate_layout()
  r <- build_diagram_xml(g)
  expect_no_error(xml2::read_xml(r$xml))
})

test_that("required namespaces present", {
  g <- parse_mermaid("flowchart TB\n A-->B") |> calculate_layout()
  xml <- build_diagram_xml(g)$xml
  expect_true(grepl("wordprocessingShape", xml, fixed = TRUE))
  expect_true(grepl("wordprocessingDrawing", xml, fixed = TRUE))
  expect_true(grepl("xmlns:a=", xml, fixed = TRUE))
  expect_true(grepl("xmlns:w=", xml, fixed = TRUE))
})

test_that("correct number of wsp elements for 3 nodes 2 edges", {
  g <- parse_mermaid("flowchart TB\n A-->B\n B-->C") |> calculate_layout()
  r <- build_diagram_xml(g)
  doc  <- xml2::read_xml(r$xml)
  wsps <- xml2::xml_find_all(doc, "//*[local-name()='wsp']")
  expect_gte(length(wsps), 5L) # 3 nodes + 2 edges
})

test_that("diamond shape in xml", {
  g <- parse_mermaid("flowchart TB\n D{decision}") |> calculate_layout()
  xml <- build_diagram_xml(g)$xml
  expect_true(grepl('prst="diamond"', xml, fixed = TRUE))
})

test_that("dashed edge includes prstDash", {
  g <- parse_mermaid("flowchart TB\n A-.->B") |> calculate_layout()
  xml <- build_diagram_xml(g)$xml
  expect_true(grepl("prstDash", xml, fixed = TRUE))
})

test_that("no shape ID collision across two diagrams", {
  g <- parse_mermaid("flowchart TB\n A-->B") |> calculate_layout()
  r1 <- build_diagram_xml(g, start_id = 1L)
  r2 <- build_diagram_xml(g, start_id = r1$next_id)
  get_ids <- function(xml) {
    m <- gregexpr('(?<=<wp:docPr id=")[0-9]+', xml, perl = TRUE)
    as.integer(regmatches(xml, m)[[1]])
  }
  expect_equal(length(intersect(get_ids(r1$xml), get_ids(r2$xml))), 0L)
})

test_that("metadata json present in descr", {
  g <- parse_mermaid("flowchart TB\n A[Start]-->B[End]") |> calculate_layout()
  xml <- build_diagram_xml(g)$xml
  # JSON is XML-escaped in the descr attribute: "type" becomes &quot;type&quot;
  expect_true(grepl("&quot;type&quot;", xml, fixed = TRUE))
})

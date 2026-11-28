# Step 2 emitter (contract-step2.md D1 and C5): the text body's vertical
# anchor and insets from the per-axis placement (B6), and which items are
# written as connectors.

body_of <- function(prims, viewBox = c(0, 0, 200, 100)) {
  doc <- generic_doc(prims, viewBox = viewBox)
  boxes <- all_wsp(doc)
  testthat::expect_length(boxes, 1L)
  body_pr(boxes[[1]])
}

insets_of <- function(bp) as.numeric(bp[c("lIns", "tIns", "rIns", "bIns")])

# -- D1: anchor and insets -------------------------------------------------------

test_that("a text above the shape's centre is anchored t with only the top inset", {
  # rect 200 x 100, text 10..70 x 5..25, anchor start: l = 10, t = 5
  bp <- body_of(list(prim_rect(0, 0, 200, 100),
                     prim_text("Title", x = 10, y = 5, w = 60, h = 20, anchor = "start")))
  expect_equal(bp[["anchor"]], "t")
  expect_near(insets_of(bp), c(10, 5, 0, 0) * 9525, tol = 2)
})

test_that("a text below the shape's centre is anchored b with only the bottom inset", {
  bp <- body_of(list(prim_rect(0, 0, 200, 100),
                     prim_text("Foot", x = 130, y = 70, w = 60, h = 20, anchor = "end")))
  expect_equal(bp[["anchor"]], "b")
  expect_near(insets_of(bp), c(0, 0, 10, 10) * 9525, tol = 2)
})

test_that("a text centred vertically but not across is anchored ctr with one side inset", {
  # anchor middle, centre 10 px right of the shape's: l = 20
  bp <- body_of(list(prim_rect(0, 0, 200, 100),
                     prim_text("Off", x = 90, y = 40, w = 40, h = 20, anchor = "middle")))
  expect_equal(bp[["anchor"]], "ctr")
  expect_near(insets_of(bp), c(20, 0, 0, 0) * 9525, tol = 2)
})

test_that("a centred text is anchored ctr with zero insets", {
  bp <- body_of(list(prim_rect(0, 0, 200, 100),
                     prim_text("Mid", x = 80, y = 40, w = 40, h = 20)))
  expect_equal(bp[["anchor"]], "ctr")
  expect_equal(insets_of(bp), c(0, 0, 0, 0))
})

# -- C5: connector role ------------------------------------------------------------

role_of <- function(prim) {
  s <- all_wsp(generic_doc(list(prim)))[[1]]
  if (!is_missing_node(xfirst(s, "./wps:cNvCnPr"))) "connector" else "shape"
}

test_that("a filled open path is written as a shape, not a connector", {
  expect_equal(role_of(prim_path("M10,10 L100,10 L100,60", fill = "#FF0000", stroke = "#333333")),
               "shape")
})

test_that("an unfilled open path and a line are connectors; a closed path is a shape", {
  expect_equal(role_of(prim_path("M10,10 L100,10 L100,60", fill = NULL, stroke = "#333333")),
               "connector")
  expect_equal(role_of(prim_line(10, 10, 100, 60)), "connector")
  expect_equal(role_of(prim_path("M10,10 L100,10 L100,60 Z", fill = NULL, stroke = "#333333")),
               "shape")
})

# Step 2 performance (contract-step2.md E). The limits are well above the
# targets (flowchart-full under 0.6 s) so that only a regression fails, not a
# slow machine.

skip_unless_timing <- function() {
  testthat::skip_on_cran()
  testthat::skip_if(!nzchar(Sys.getenv("NOT_CRAN")), "NOT_CRAN is unset")
}

test_that("build_scene() on flowchart-full takes under 1.5 s", {
  skip_unless_timing()
  r <- fixture_render("flowchart-full")
  quietly_expected(build_scene(r))           # warm up (package loading, caches)
  elapsed <- system.time(quietly_expected(build_scene(r)))[["elapsed"]]
  expect_lt(elapsed, 1.5)
})

test_that("path_parse() on a path of 2000 commands takes under 0.5 s", {
  skip_unless_timing()
  pts <- sprintf("L%d,%d", seq_len(2000), (seq_len(2000) %% 7) * 3)
  d <- paste("M0,0", paste(pts, collapse = " "), "Z")
  elapsed <- system.time(p <- path_parse(d))[["elapsed"]]
  expect_equal(nrow(p), 2002L)
  expect_lt(elapsed, 0.5)
})

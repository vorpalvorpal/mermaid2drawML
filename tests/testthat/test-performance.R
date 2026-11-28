# How long building a scene and its XML takes. The limits are well above the
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

# -- Scaling with the number of groups -----------------------------------------
#
# A synthetic diagram of `n` scope groups, each a box with a text inside and
# a small dot, each drawn in its own element: no step of build_scene() may
# take time that grows with the square of the number of items or groups.

scope_grid_render <- function(n, cols = 40L) {
  prims <- purrr::map(seq_len(n), function(k) {
    x <- ((k - 1L) %% cols) * 60
    y <- ((k - 1L) %/% cols) * 40
    in_g(list(prim_rect(x, y, 50, 30, fill = "#ECECFF"),
              prim_text(paste0("n", k), x = x + 5, y = y + 8, w = 30, h = 14),
              prim_ellipse(x + 45, y + 5, 2, fill = "#333333", stroke = NULL)),
         n = k)
  })
  generic_render(purrr::list_flatten(prims),
                 viewBox = c(0, 0, cols * 60, ceiling(n / cols) * 40))
}

# Seconds build_scene() takes on scope_grid_render(n), measured once per
# session. `limit`: give up after that many seconds (the elapsed time is then
# reported as Inf), so that a slow build fails the test without holding up
# the suite for minutes.
scene_seconds <- function(n, limit = Inf) {
  cached(paste0("scene-seconds:", n), function() {
    r <- scope_grid_render(n)
    quietly_expected(build_scene(scope_grid_render(20L)))    # warm up
    t0 <- proc.time()[["elapsed"]]
    done <- tryCatch({
      setTimeLimit(elapsed = limit, transient = TRUE)
      items <- quietly_expected(build_scene(r))$items
      sum(items$kind == "group")
    }, error = function(e) {
      if (grepl("time limit", conditionMessage(e))) NA_integer_ else stop(e)
    })
    setTimeLimit(elapsed = Inf)
    elapsed <- proc.time()[["elapsed"]] - t0
    list(seconds = if (is.na(done)) Inf else elapsed, groups = done)
  })
}

test_that("500 scope groups build in under 4 s", {
  skip_unless_timing()
  s <- scene_seconds(500L, limit = 30)
  expect_equal(s$groups, 500L)
  expect_lt(s$seconds, 4)
})

test_that("2000 scope groups build in under 20 s", {
  skip_unless_timing()
  s <- scene_seconds(2000L, limit = 25)
  expect_lt(s$seconds, 20)
})

test_that("doubling the number of groups less than triples the time", {
  skip_unless_timing()
  small <- scene_seconds(500L, limit = 30)$seconds
  large <- scene_seconds(1000L, limit = 120)$seconds
  expect_lt(large / small, 3)
})

test_that("path_parse() on a path of 2000 commands takes under 0.5 s", {
  skip_unless_timing()
  pts <- sprintf("L%d,%d", seq_len(2000), (seq_len(2000) %% 7) * 3)
  d <- paste("M0,0", paste(pts, collapse = " "), "Z")
  elapsed <- system.time(p <- path_parse(d))[["elapsed"]]
  expect_equal(nrow(p), 2002L)
  expect_lt(elapsed, 0.5)
})

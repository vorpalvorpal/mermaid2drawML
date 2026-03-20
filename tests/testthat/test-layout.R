# tests/testthat/test-layout.R
# The R-based igraph layout engine (calculate_layout()) was removed in v0.2.0.
# Layout is now provided by mermaid-cli (dagre), read directly from SVG output.
# These tests are superseded by test-parse_svg.R and test-word_diagram.R.

test_that("calculate_layout no longer exported (v0.2.0 architectural change)", {
  skip("calculate_layout() removed in v0.2.0 — layout is now read from mermaid SVG output")
})

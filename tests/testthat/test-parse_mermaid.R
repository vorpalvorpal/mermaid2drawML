# tests/testthat/test-parse_mermaid.R
# The R-based regex Mermaid parser (parse_mermaid()) was removed in v0.2.0.
# Parsing is now delegated to @mermaid-js/parser via the Node.js bridge.
# These tests are superseded by test-parse_svg.R and test-word_diagram.R.

test_that("parse_mermaid no longer exported (v0.2.0 architectural change)", {
  skip("parse_mermaid() removed in v0.2.0 — use render_mermaid() + parse_mermaid_svg()")
})

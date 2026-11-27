# Shared test helpers: bridge fixtures and mocked rendering.

fixture_path <- function(name) {
  testthat::test_path("fixtures", paste0(name, ".json"))
}

# One parsed copy of each fixture per session.
.fixture_cache <- rlang::env()

fixture_render <- function(name) {
  if (is.null(.fixture_cache[[name]])) {
    .fixture_cache[[name]] <- read_render(fixture_path(name))
  }
  .fixture_cache[[name]]
}

skip_if_no_mermaid <- function() {
  testthat::skip_on_cran()
  testthat::skip_if(!nzchar(Sys.which("node")), "node is not on PATH")
  ok <- tryCatch({
    check_mermaid2drawml()
    TRUE
  }, error = function(e) FALSE)
  testthat::skip_if(!ok, "mermaid2drawml Node packages are not installed")
}

# Make render_mermaid() return `render` (a fixture name or a mermaid_render).
local_mock_render <- function(render, env = parent.frame()) {
  if (is.character(render)) render <- fixture_render(render)
  testthat::local_mocked_bindings(
    render_mermaid = function(...) render,
    .env = env
  )
}

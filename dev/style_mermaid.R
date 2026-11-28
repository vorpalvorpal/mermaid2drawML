# dev/style_mermaid.R
#
# Standalone helper: applies the standard LEMP/BMCC Mermaid styling to a
# diagram string and returns the modified text.
#
# Usage:
#   source("dev/style_mermaid.R")
#   bg_colour <- "Rosella"
#   styled <- style_mermaid(readLines("my_diagram.mermaid") |> paste(collapse = "\n"))
#   # or pass substitutions explicitly:
#   styled <- style_mermaid(raw_text, bg_colour = "Bush")
#
# The result can be passed to mermaid2drawml::body_add_mermaid().
# This file is NOT part of the mermaid2drawml package.

# ── BMCC colour palette ───────────────────────────────────────────────────
# The LEMP project's colours: each is a node class and a subgraph class
# below (stroke and, for nodes, text).

BMCC_COLOURS <- c(
  Rosella40    = "#FCA789",
  Rosella70    = "#FC5A3C",
  Rosella      = "#FF0000",
  Earth        = "#700017",
  Bush40       = "#A3C380",
  Bush80       = "#5E8933",
  Bush         = "#507811",
  Fern         = "#6EBD33",
  Canyon       = "#006A5C",
  Gum          = "#5A8263",
  Gum60        = "#8EAA91",
  Mist50       = "#B3C7E0",
  Mist         = "#6095C1",
  Mountains70  = "#5279A9",
  Mountains    = "#005490",
  Winter       = "#021D49",
  Rhododendron = "#9186BF",
  WildFlower   = "#FF671D",
  Wattle       = "#FFB71B",
  Escarpment50 = "#FEE6AA",
  Escarpment   = "#FED95E",
  Bark         = "#5E504E"
)

# ── Init block ────────────────────────────────────────────────────────────

.INIT_BLOCK <- "%%{init: {
  'theme': 'base',
  'flowchart': {
    'htmlLabels': false,
    'defaultRenderer': 'elk'
  },
  'themeVariables': {
    'background': '#FFFFFF',
    'lineColor': '#5E504E',
    'clusterBkg': 'transparent',
    'edgeLabelBackground': '#FFFFFF',
    'primaryTextColor': '#5E504E',
    'fontFamily': 'Myriad Pro',
    'fontSize': '20px'
  }
}}%%
"

# ── ClassDef blocks ───────────────────────────────────────────────────────
# One node class and one subgraph class ("sub" + name) per palette colour,
# then the two subgraph classes that are not colours.

.classdef_lines <- function(names, template) {
  paste0("classDef ", glue::glue(template, name = names, hex = BMCC_COLOURS[names]), "\n",
         collapse = "")
}

.NODE_CLASSDEFS <- paste0("\n", .classdef_lines(
  names(BMCC_COLOURS),
  "{name} fill:#FFFFFFAA,stroke:{hex},stroke-width:2px,color:{hex},font-size:20px;"
))

.SUBGRAPH_CLASSDEFS <- paste0("\n", .classdef_lines(
  names(BMCC_COLOURS),
  "sub{name} fill:transparent,stroke:{hex},stroke-width:2px,color:#5E504E,font-size:26px;"
),
"classDef subSub fill:transparent,stroke:#5E504E44,stroke-width:2px,color:#5E504E88,font-size:20px;\n",
"classDef subSpace fill:transparent,stroke:transparent,stroke-width:0px,color:transparent,font-size:20px,stroke-dasharray: 0 1;\n"
)

# ── style_mermaid() ───────────────────────────────────────────────────────

#' Apply standard BMCC styling to a Mermaid diagram string
#'
#' @param mermaid_code Character. Raw Mermaid text, optionally containing
#'   \code{@@\{var\}@@} placeholders.
#' @param ... Named values used to resolve placeholders (e.g.
#'   \code{bg_colour = "Rosella"}). Variables not supplied here are looked up
#'   in the calling environment, matching the original behaviour of
#'   \code{replace_mermaid_env_vars()}.
#' @return A single character string of styled Mermaid text ready for
#'   \code{mermaid2drawml::body_add_mermaid()}.
style_mermaid <- function(mermaid_code, ..., .envir = parent.frame()) {
  stopifnot(is.character(mermaid_code), length(mermaid_code) == 1L)
  # The values given here, in front of the caller's variables.
  scope <- rlang::env(.envir, ...)
  mermaid_code <- .resolve_mermaid_vars(mermaid_code, scope)
  # An init block of the diagram's own wins.
  if (!grepl("%%{init", mermaid_code, fixed = TRUE)) {
    mermaid_code <- paste0(.INIT_BLOCK, mermaid_code)
  }
  paste0(mermaid_code, .NODE_CLASSDEFS, .SUBGRAPH_CLASSDEFS)
}

.resolve_mermaid_vars <- function(text, envir) {
  hits <- regmatches(text, gregexpr("@@\\{([^}]+)\\}@@", text, perl = TRUE))[[1]]
  var_names <- unique(sub("^@@\\{(.*)\\}@@$", "\\1", hits))
  missing <- var_names[!vapply(var_names, exists, logical(1L), envir = envir, inherits = TRUE)]
  if (length(missing) > 0L) {
    cli::cli_abort(
      c("Unresolved {.code @@{{...}}@@} placeholder{?s}: {.val {missing}}.",
        i = "Supply them via {.arg ...} or define them in the calling environment."),
      class = "style_mermaid_unresolved"
    )
  }
  purrr::reduce(var_names, function(acc, nm) {
    gsub(paste0("@@\\{", nm, "\\}@@"), get(nm, envir = envir, inherits = TRUE), acc, perl = TRUE)
  }, .init = text)
}

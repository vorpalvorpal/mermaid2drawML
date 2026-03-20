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
# Hex values are the stroke colours from the LEMP project's classDef blocks.

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

.NODE_CLASSDEFS <- "
classDef Rosella40 fill:#FFFFFFAA,stroke:#FCA789,stroke-width:2px,color:#FCA789,font-size:20px;
classDef Rosella70 fill:#FFFFFFAA,stroke:#FC5A3C,stroke-width:2px,color:#FC5A3C,font-size:20px;
classDef Rosella fill:#FFFFFFAA,stroke:#FF0000,stroke-width:2px,color:#FF0000,font-size:20px;
classDef Earth fill:#FFFFFFAA,stroke:#700017,stroke-width:2px,color:#700017,font-size:20px;
classDef Bush40 fill:#FFFFFFAA,stroke:#A3C380,stroke-width:2px,color:#A3C380,font-size:20px;
classDef Bush80 fill:#FFFFFFAA,stroke:#5E8933,stroke-width:2px,color:#5E8933,font-size:20px;
classDef Bush fill:#FFFFFFAA,stroke:#507811,stroke-width:2px,color:#507811,font-size:20px;
classDef Fern fill:#FFFFFFAA,stroke:#6EBD33,stroke-width:2px,color:#6EBD33,font-size:20px;
classDef Canyon fill:#FFFFFFAA,stroke:#006A5C,stroke-width:2px,color:#006A5C,font-size:20px;
classDef Gum fill:#FFFFFFAA,stroke:#5A8263,stroke-width:2px,color:#5A8263,font-size:20px;
classDef Gum60 fill:#FFFFFFAA,stroke:#8EAA91,stroke-width:2px,color:#8EAA91,font-size:20px;
classDef Mist50 fill:#FFFFFFAA,stroke:#B3C7E0,stroke-width:2px,color:#B3C7E0,font-size:20px;
classDef Mist fill:#FFFFFFAA,stroke:#6095C1,stroke-width:2px,color:#6095C1,font-size:20px;
classDef Mountains70 fill:#FFFFFFAA,stroke:#5279A9,stroke-width:2px,color:#5279A9,font-size:20px;
classDef Mountains fill:#FFFFFFAA,stroke:#005490,stroke-width:2px,color:#005490,font-size:20px;
classDef Winter fill:#FFFFFFAA,stroke:#021D49,stroke-width:2px,color:#021D49,font-size:20px;
classDef Rhododendron fill:#FFFFFFAA,stroke:#9186BF,stroke-width:2px,color:#9186BF,font-size:20px;
classDef WildFlower fill:#FFFFFFAA,stroke:#FF671D,stroke-width:2px,color:#FF671D,font-size:20px;
classDef Wattle fill:#FFFFFFAA,stroke:#FFB71B,stroke-width:2px,color:#FFB71B,font-size:20px;
classDef Escarpment50 fill:#FFFFFFAA,stroke:#FEE6AA,stroke-width:2px,color:#FEE6AA,font-size:20px;
classDef Escarpment fill:#FFFFFFAA,stroke:#FED95E,stroke-width:2px,color:#FED95E,font-size:20px;
classDef Bark fill:#FFFFFFAA,stroke:#5E504E,stroke-width:2px,color:#5E504E,font-size:20px;
"

.SUBGRAPH_CLASSDEFS <- "
classDef subRosella40 fill:transparent,stroke:#FCA789,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subRosella70 fill:transparent,stroke:#FC5A3C,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subRosella fill:transparent,stroke:#FF0000,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subEarth fill:transparent,stroke:#700017,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subBush40 fill:transparent,stroke:#A3C380,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subBush80 fill:transparent,stroke:#5E8933,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subBush fill:transparent,stroke:#507811,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subFern fill:transparent,stroke:#6EBD33,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subCanyon fill:transparent,stroke:#006A5C,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subGum fill:transparent,stroke:#5A8263,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subGum60 fill:transparent,stroke:#8EAA91,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subMist50 fill:transparent,stroke:#B3C7E0,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subMist fill:transparent,stroke:#6095C1,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subMountains70 fill:transparent,stroke:#5279A9,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subMountains fill:transparent,stroke:#005490,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subWinter fill:transparent,stroke:#021D49,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subRhododendron fill:transparent,stroke:#9186BF,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subWildFlower fill:transparent,stroke:#FF671D,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subWattle fill:transparent,stroke:#FFB71B,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subEscarpment50 fill:transparent,stroke:#FEE6AA,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subEscarpment fill:transparent,stroke:#FED95E,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subBark fill:transparent,stroke:#5E504E,stroke-width:2px,color:#5E504E,font-size:26px;
classDef subSub fill:transparent,stroke:#5E504E44,stroke-width:2px,color:#5E504E88,font-size:20px;
classDef subSpace fill:transparent,stroke:transparent,stroke-width:0px,color:transparent,font-size:20px,stroke-dasharray: 0 1;
"

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

  # Merge explicit overrides into a child of the calling environment
  overrides <- list(...)
  if (length(overrides) > 0L) {
    env <- new.env(parent = .envir)
    for (nm in names(overrides)) assign(nm, overrides[[nm]], envir = env)
    .envir <- env
  }

  # 1. Resolve @@{var}@@ placeholders
  mermaid_code <- .resolve_mermaid_vars(mermaid_code, .envir)

  # 2. Prepend %%{init}%% block if not already present
  if (!grepl("%%{init", mermaid_code, fixed = TRUE)) {
    mermaid_code <- paste0(.INIT_BLOCK, mermaid_code)
  }

  # 3. Append node and subgraph classDef blocks
  paste0(mermaid_code, .NODE_CLASSDEFS, .SUBGRAPH_CLASSDEFS)
}

.resolve_mermaid_vars <- function(text, envir) {
  pattern <- "@@\\{([^}]+)\\}@@"
  hits     <- regmatches(text, gregexpr(pattern, text, perl = TRUE))[[1]]
  if (length(hits) == 0L) return(text)

  var_names <- unique(
    regmatches(hits, regexpr("(?<=@@\\{)[^}]+(?=\\}@@)", hits, perl = TRUE))
  )

  missing <- var_names[!vapply(var_names, exists, logical(1L),
                               envir = envir, inherits = TRUE)]
  if (length(missing) > 0L) {
    stop(
      "Unresolved @@{...}@@ placeholder(s): ",
      paste(missing, collapse = ", "),
      "\nSupply them via `...` or ensure they exist in the calling environment."
    )
  }

  for (nm in var_names) {
    val  <- get(nm, envir = envir, inherits = TRUE)
    text <- gsub(paste0("@@\\{", nm, "\\}@@"), val, text, perl = TRUE)
  }
  text
}

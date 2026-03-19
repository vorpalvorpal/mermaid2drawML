# Internal helpers -------------------------------------------------------------

default_node <- function(id) {
  list(id = id, label = id, shape = "rect",
       fill = NA_character_, stroke = NA_character_,
       color = NA_character_, class = NA_character_)
}

parse_style_string <- function(s) {
  parts  <- strsplit(s, ",")[[1]]
  result <- list(fill = NA_character_, stroke = NA_character_,
                 color = NA_character_, stroke_width = NA_real_)
  for (p in parts) {
    p   <- trimws(p)
    sep <- regexpr(":", p)
    if (sep < 1) next
    key <- trimws(substr(p, 1, sep - 1))
    val <- trimws(substr(p, sep + 1, nchar(p)))
    if (key == "fill")         result$fill   <- parse_css_colour(val)
    else if (key == "stroke")  result$stroke <- parse_css_colour(val)
    else if (key == "color" || key == "colour") result$color <- parse_css_colour(val)
    else if (key == "stroke-width") {
      num <- suppressWarnings(as.numeric(gsub("[^0-9.]", "", val)))
      if (!is.na(num)) result$stroke_width <- num / 1.333  # px -> pt approx
    }
  }
  result
}

# Node pattern table: (regex, shape, label_group)
# Applied in order; first match wins. Group 1 = id, group 2 = label (where applicable).
.node_patterns <- list(
  list(re = '^(\\w[\\w.-]*)@\\{',                          shape = "text",          label_grp = NA),
  list(re = '^(\\w[\\w.-]*)\\{\\{(.+)\\}\\}$',             shape = "hexagon",       label_grp = 2L),
  list(re = '^(\\w[\\w.-]*)\\{(.+)\\}$',                   shape = "diamond",       label_grp = 2L),
  list(re = '^(\\w[\\w.-]*)\\(\\((.+)\\)\\)$',             shape = "ellipse",       label_grp = 2L),
  list(re = '^(\\w[\\w.-]*)\\(\\[(.+)\\]\\)$',             shape = "roundRect",     label_grp = 2L),
  list(re = '^(\\w[\\w.-]*)\\[\\((.+)\\)\\]$',             shape = "rect",          label_grp = 2L),
  list(re = '^(\\w[\\w.-]*)\\[\\[(.+)\\]\\]$',             shape = "roundRect",     label_grp = 2L),
  list(re = '^(\\w[\\w.-]*)\\[\\\\(.+)/\\]$',              shape = "parallelogram", label_grp = 2L),
  list(re = '^(\\w[\\w.-]*)\\[/(.+)\\\\\\]$',              shape = "parallelogram", label_grp = 2L),
  list(re = '^(\\w[\\w.-]*)\\["([^"]+)"\\]$',              shape = "rect",          label_grp = 2L),
  list(re = '^(\\w[\\w.-]*)\\[(.+)\\]$',                   shape = "rect",          label_grp = 2L),
  list(re = '^(\\w[\\w.-]*)\\("([^"]+)"\\)$',              shape = "roundRect",     label_grp = 2L),
  list(re = '^(\\w[\\w.-]*)\\((.+)\\)$',                   shape = "roundRect",     label_grp = 2L)
)

parse_node_text <- function(text) {
  # text is already stripped of :::class
  text <- trimws(text)
  if (!nzchar(text)) return(NULL)

  # Handle @{shape: text, label: "..."} mermaid v10 shapes
  if (grepl("@\\{", text, perl = TRUE)) {
    id <- trimws(sub("@\\{.*", "", text, perl = TRUE))
    id <- trimws(id)
    if (!grepl("^\\w[\\w.-]*$", id, perl = TRUE)) id <- gsub("[^\\w.-]", "", id, perl = TRUE)
    lm <- regmatches(text, regexpr('label:\\s*"([^"]*)"', text, perl = TRUE))
    label <- if (length(lm) > 0 && nzchar(lm)) {
      sub('label:\\s*"([^"]*)"', "\\1", lm, perl = TRUE)
    } else id
    return(list(id = id, label = strip_html(label), shape = "text"))
  }

  for (pat in .node_patterns) {
    m <- regexec(pat$re, text, perl = TRUE)[[1]]
    if (m[1] > 0) {
      id    <- substr(text, m[2], m[2] + attr(m, "match.length")[2] - 1)
      label <- if (!is.na(pat$label_grp) && length(m) > pat$label_grp && m[pat$label_grp + 1] > 0) {
        substr(text, m[pat$label_grp + 1], m[pat$label_grp + 1] + attr(m, "match.length")[pat$label_grp + 1] - 1)
      } else id
      return(list(id = id, label = strip_html(label), shape = pat$shape))
    }
  }

  # Plain word identifier (allow dots, hyphens, underscores)
  if (grepl("^\\w[\\w._-]*$", text, perl = TRUE)) {
    return(list(id = text, label = text, shape = "rect"))
  }
  NULL
}

strip_class_annotation <- function(text) {
  # Returns list(text=..., class=...) after stripping :::ClassName
  class_name <- NA_character_
  if (grepl(":::", text, fixed = TRUE)) {
    class_name <- trimws(sub(".*:::(\\w+).*", "\\1", text, perl = TRUE))
    text       <- trimws(sub(":::.*$", "", text, perl = TRUE))
  }
  list(text = text, class = class_name)
}

# Edge detection ---------------------------------------------------------------

# Tokenise: split the line into chunks before/after quoted strings and
# pipe-labels, so that edge marker detection doesn't match inside node labels.
# Returns: NULL if no edge marker found
# Named list: marker, start, end (character positions in original line), line_type, arrow_start, arrow_end

detect_edge_marker <- function(line) {
  # Order matters: longest / most specific first
  patterns <- list(
    list(re = "<-->",        lt = "solid",  as = "arrow",  ae = "arrow"),
    list(re = "o--o",        lt = "solid",  as = "circle", ae = "circle"),
    list(re = "x--x",        lt = "solid",  as = "cross",  ae = "cross"),
    list(re = "={2,}>",      lt = "thick",  as = "none",   ae = "arrow"),
    list(re = "-\\.{1,}-?>", lt = "dashed", as = "none",   ae = "arrow"),
    list(re = "<-{2,}",      lt = "solid",  as = "arrow",  ae = "none"),
    list(re = "--x",         lt = "solid",  as = "none",   ae = "cross"),
    list(re = "--o",         lt = "solid",  as = "none",   ae = "circle"),
    list(re = "-{2,}>",      lt = "solid",  as = "none",   ae = "arrow"),
    list(re = "-{3,}",       lt = "solid",  as = "none",   ae = "none")
  )

  # We need to avoid matching inside square brackets or quoted strings
  # Build a mask of positions that are "inside" brackets/quotes
  n <- nchar(line)
  mask <- logical(n)  # TRUE = inside brackets/quotes, skip
  depth <- 0L
  in_quote <- FALSE
  quote_char <- ""
  i <- 1L
  while (i <= n) {
    ch <- substr(line, i, i)
    if (in_quote) {
      if (ch == quote_char) in_quote <- FALSE
      mask[i] <- TRUE
    } else if (ch == '"' || ch == "'") {
      in_quote  <- TRUE
      quote_char <- ch
      mask[i]  <- TRUE
    } else if (ch %in% c("[", "(", "{")) {
      depth <- depth + 1L
      if (depth > 0L) mask[i] <- TRUE
    } else if (ch %in% c("]", ")", "}")) {
      if (depth > 0L) mask[i] <- TRUE
      depth <- max(0L, depth - 1L)
    } else if (depth > 0L) {
      mask[i] <- TRUE
    }
    i <- i + 1L
  }

  for (p in patterns) {
    m <- gregexpr(p$re, line, perl = TRUE)[[1]]
    if (m[1] < 0) next
    for (j in seq_along(m)) {
      start <- m[j]
      end   <- m[j] + attr(m, "match.length")[j] - 1L
      # Check if any char of the match is masked
      if (any(mask[start:end])) next
      return(list(
        marker     = substr(line, start, end),
        start      = start,
        end        = end,
        line_type  = p$lt,
        arrow_start = p$as,
        arrow_end  = p$ae
      ))
    }
  }
  NULL
}

extract_edge_label <- function(left, right, marker_info) {
  label <- NA_character_

  # Check for pipe label: --> |label| at start of right side
  right_trim <- trimws(right)
  if (grepl("^\\|([^|]*)\\|", right_trim, perl = TRUE)) {
    m <- regexec("^\\|([^|]*)\\|", right_trim, perl = TRUE)[[1]]
    label <- trimws(substr(right_trim, m[2], m[2] + attr(m, "match.length")[2] - 1))
    right <- trimws(substr(right_trim, m[1] + attr(m, "match.length")[1], nchar(right_trim)))
  }

  # Check for inline label in marker left side: -- label -->
  # The marker itself may have been --label--> in the original
  # left_trim is what's before the edge marker
  if (is.na(label)) {
    left_trim <- trimws(left)
    ml <- regexpr("--([^->|]+)$", left_trim, perl = TRUE)
    if (ml > 0) {
      possible_label <- trimws(regmatches(left_trim, ml))
      possible_label <- trimws(sub("^--+", "", possible_label))
      possible_label <- trimws(sub("-+>?$", "", possible_label))
      if (nzchar(possible_label)) {
        label <- possible_label
        left  <- trimws(substr(left_trim, 1, ml - 1))
      }
    }
  }

  list(left = left, right = right, label = label)
}

parse_edge_line <- function(line) {
  # Returns a list of edge records (for chained edges), or NULL
  mk <- detect_edge_marker(line)
  if (is.null(mk)) return(NULL)

  left_raw  <- trimws(substr(line, 1, mk$start - 1))
  right_raw <- trimws(substr(line, mk$end + 1, nchar(line)))

  # Extract label from pipe syntax or inline
  extracted <- extract_edge_label(left_raw, right_raw, mk)
  left_raw  <- extracted$left
  right_raw <- extracted$right
  label     <- extracted$label

  # Parse left node — strip class annotation first
  left_ca  <- strip_class_annotation(left_raw)
  left_nd  <- parse_node_text(left_ca$text)
  if (is.null(left_nd)) return(NULL)
  left_nd$class <- left_ca$class

  # The right side might itself contain another edge (chaining)
  right_mk <- detect_edge_marker(right_raw)

  if (!is.null(right_mk)) {
    # Parse just the immediate right node (before the next marker)
    right_node_text <- trimws(substr(right_raw, 1, right_mk$start - 1))
    right_ca   <- strip_class_annotation(right_node_text)
    right_nd   <- parse_node_text(right_ca$text)
    if (is.null(right_nd)) return(NULL)
    right_nd$class <- right_ca$class

    # Recurse for the remainder
    rest <- parse_edge_line(right_raw)

    edge <- list(
      from        = left_nd$id,
      to          = right_nd$id,
      label       = label,
      line_type   = mk$line_type,
      arrow_start = mk$arrow_start,
      arrow_end   = mk$arrow_end
    )
    return(c(
      list(list(edge = edge, nodes = list(left_nd, right_nd))),
      rest
    ))
  } else {
    right_ca <- strip_class_annotation(right_raw)
    right_nd <- parse_node_text(right_ca$text)
    if (is.null(right_nd)) return(NULL)
    right_nd$class <- right_ca$class

    edge <- list(
      from        = left_nd$id,
      to          = right_nd$id,
      label       = label,
      line_type   = mk$line_type,
      arrow_start = mk$arrow_start,
      arrow_end   = mk$arrow_end
    )
    return(list(list(edge = edge, nodes = list(left_nd, right_nd))))
  }
}

# Main parser ------------------------------------------------------------------

#' Parse a Mermaid flowchart diagram
#'
#' @param mermaid_code A character string containing Mermaid flowchart/graph syntax.
#' @return A `mermaid_graph` S3 object containing direction, nodes, edges, subgraphs, and classes.
#' @export
parse_mermaid <- function(mermaid_code) {
  # Strip ```mermaid fences if present
  mermaid_code <- gsub("^```mermaid\\s*", "", mermaid_code, perl = TRUE)
  mermaid_code <- gsub("```\\s*$",        "", mermaid_code, perl = TRUE)

  lines <- strsplit(mermaid_code, "\n", fixed = TRUE)[[1]]

  direction   <- "TB"
  nodes       <- list()   # id -> node list
  edges_list  <- list()
  subgraphs   <- list()   # id -> subgraph list
  classes     <- list()   # name -> style list
  sg_stack    <- character(0)

  placeholder_warned <- FALSE

  for (raw_line in lines) {
    # Strip comments
    line <- gsub("%%[^\n]*", "", raw_line, perl = TRUE)
    # Replace style placeholders @@{...}@@
    if (grepl("@@\\{[^}]*\\}@@", line, perl = TRUE)) {
      if (!placeholder_warned) {
        rlang::warn(
          "Style placeholders (@@{...}@@) are not supported and have been replaced with defaults.",
          call. = FALSE
        )
        placeholder_warned <- TRUE
      }
      line <- gsub("@@\\{[^}]*\\}@@", "", line, perl = TRUE)
    }
    line <- trimws(line)
    if (!nzchar(line)) next

    # Header: flowchart / graph
    if (grepl("^(flowchart|graph)\\s+", line, perl = TRUE, ignore.case = TRUE)) {
      m <- regmatches(line, regexpr("^(?:flowchart|graph)\\s+(\\S+)", line, perl = TRUE, ignore.case = TRUE))
      if (length(m) > 0 && nzchar(m)) {
        dir_part <- trimws(sub("^(?:flowchart|graph)\\s+", "", m, perl = TRUE, ignore.case = TRUE))
        dir_part <- toupper(sub("^(TB|LR|BT|RL|TD).*", "\\1", dir_part, perl = TRUE))
        if (dir_part == "TD") dir_part <- "TB"
        if (nzchar(dir_part)) direction <- dir_part
      }
      next
    }

    # Subgraph open
    if (grepl("^subgraph\\b", line, perl = TRUE, ignore.case = TRUE)) {
      sg_def   <- trimws(sub("^subgraph\\s*", "", line, perl = TRUE, ignore.case = TRUE))
      sg_id    <- NA_character_
      sg_label <- NA_character_
      if (grepl("^\\[", sg_def)) {
        # No id: subgraph [label] or subgraph [ ]
        sg_id    <- paste0("__sg_", length(subgraphs) + 1L)
        inner    <- sub("^\\[(.*)\\]$", "\\1", sg_def)
        sg_label <- strip_html(trimws(inner))
      } else if (grepl('\\[', sg_def)) {
        sg_id    <- trimws(sub("\\s*\\[.*$", "", sg_def))
        inner    <- sub(".*\\[(.*)\\].*", "\\1", sg_def)
        sg_label <- strip_html(trimws(inner))
      } else if (grepl('^"', sg_def)) {
        sg_id    <- paste0("__sg_", length(subgraphs) + 1L)
        sg_label <- gsub('"', '', sg_def)
      } else {
        sg_id    <- sg_def
        sg_label <- sg_def
      }
      # Clean up sg_id - allow dots but keep it sane
      sg_id <- trimws(sg_id)
      if (!nzchar(sg_id) || sg_id == " " || sg_id == "[ ]") {
        sg_id <- paste0("__sg_", length(subgraphs) + 1L)
      }
      if (is.na(sg_label) || !nzchar(trimws(sg_label)) || sg_label == " ") {
        sg_label <- sg_id
      }
      subgraphs[[sg_id]] <- list(id = sg_id, label = sg_label, node_ids = character(0),
                                  fill = NA_character_, stroke = NA_character_)
      sg_stack <- c(sg_stack, sg_id)
      next
    }

    # Subgraph close
    if (grepl("^end\\s*$", line, perl = TRUE, ignore.case = TRUE)) {
      if (length(sg_stack) > 0) sg_stack <- sg_stack[-length(sg_stack)]
      next
    }

    # direction keyword inside subgraph
    if (grepl("^direction\\s+", line, perl = TRUE, ignore.case = TRUE)) next

    # classDef
    if (grepl("^classDef\\s+", line, perl = TRUE, ignore.case = TRUE)) {
      rest    <- trimws(sub("^classDef\\s+", "", line, perl = TRUE, ignore.case = TRUE))
      # class name is first token
      cd_name_m <- regexpr("^\\S+", rest, perl = TRUE)
      if (cd_name_m < 1) next
      cd_name <- substr(rest, cd_name_m, cd_name_m + attr(cd_name_m, "match.length") - 1)
      cd_body <- trimws(substr(rest, cd_name_m + attr(cd_name_m, "match.length"), nchar(rest)))
      # Remove trailing semicolon
      cd_body <- sub(";\\s*$", "", cd_body)
      classes[[cd_name]] <- parse_style_string(cd_body)
      next
    }

    # class assignment: class A,B,C className  OR  class A className;
    if (grepl("^class\\s+", line, perl = TRUE, ignore.case = TRUE)) {
      rest <- trimws(sub("^class\\s+", "", line, perl = TRUE, ignore.case = TRUE))
      rest <- sub(";\\s*$", "", rest)
      # Last whitespace-delimited token is the class name
      tokens    <- strsplit(trimws(rest), "\\s+", perl = TRUE)[[1]]
      if (length(tokens) < 2) next
      cn        <- tokens[length(tokens)]
      ids_part  <- paste(tokens[-length(tokens)], collapse = " ")
      ids       <- trimws(strsplit(ids_part, ",", fixed = TRUE)[[1]])
      for (nid in ids) {
        nid <- trimws(nid)
        if (!nzchar(nid)) next
        if (!nid %in% names(nodes)) nodes[[nid]] <- default_node(nid)
        nodes[[nid]]$class <- cn
      }
      next
    }

    # Try edge line
    edge_results <- parse_edge_line(line)
    if (!is.null(edge_results)) {
      for (er in edge_results) {
        # Register nodes
        for (nd in er$nodes) {
          nid <- nd$id
          if (!nid %in% names(nodes)) {
            nodes[[nid]]        <- default_node(nid)
            nodes[[nid]]$label  <- nd$label
            nodes[[nid]]$shape  <- nd$shape
          } else {
            if (!is.na(nd$shape) && nd$shape != "rect") nodes[[nid]]$shape <- nd$shape
            if (!is.na(nd$label) && nd$label != nid)    nodes[[nid]]$label <- nd$label
          }
          if (!is.null(nd$class) && !is.na(nd$class)) nodes[[nid]]$class <- nd$class
          # Add to all ancestor subgraphs
          for (sg_id in sg_stack) {
            if (!nid %in% subgraphs[[sg_id]]$node_ids)
              subgraphs[[sg_id]]$node_ids <- c(subgraphs[[sg_id]]$node_ids, nid)
          }
        }
        edges_list <- c(edges_list, list(er$edge))
      }
      next
    }

    # Standalone node declaration
    ca       <- strip_class_annotation(line)
    node_res <- parse_node_text(ca$text)
    if (!is.null(node_res)) {
      nid <- node_res$id
      if (!nid %in% names(nodes)) {
        nodes[[nid]]        <- default_node(nid)
        nodes[[nid]]$label  <- node_res$label
        nodes[[nid]]$shape  <- node_res$shape
      } else {
        if (!is.na(node_res$shape)) nodes[[nid]]$shape <- node_res$shape
        if (!is.na(node_res$label) && node_res$label != nid) nodes[[nid]]$label <- node_res$label
      }
      if (!is.na(ca$class)) nodes[[nid]]$class <- ca$class
      for (sg_id in sg_stack) {
        if (!nid %in% subgraphs[[sg_id]]$node_ids)
          subgraphs[[sg_id]]$node_ids <- c(subgraphs[[sg_id]]$node_ids, nid)
      }
    }
  }

  # Apply class styles to nodes
  for (nid in names(nodes)) {
    cl <- nodes[[nid]]$class
    if (!is.null(cl) && !is.na(cl) && cl %in% names(classes)) {
      st <- classes[[cl]]
      if (is.na(nodes[[nid]]$fill)   && !is.null(st$fill))   nodes[[nid]]$fill   <- st$fill
      if (is.na(nodes[[nid]]$stroke) && !is.null(st$stroke)) nodes[[nid]]$stroke <- st$stroke
      if (is.na(nodes[[nid]]$color)  && !is.null(st$color))  nodes[[nid]]$color  <- st$color
    }
  }

  # Build tibbles
  if (length(nodes) == 0) {
    nodes_df <- tibble::tibble(
      id = character(), label = character(), shape = character(),
      fill = character(), stroke = character(), color = character(), class = character()
    )
  } else {
    nodes_df <- tibble::tibble(
      id     = vapply(nodes, `[[`, character(1), "id"),
      label  = vapply(nodes, `[[`, character(1), "label"),
      shape  = vapply(nodes, `[[`, character(1), "shape"),
      fill   = vapply(nodes, function(n) if (is.null(n$fill)   || is.na(n$fill))   NA_character_ else n$fill,   character(1)),
      stroke = vapply(nodes, function(n) if (is.null(n$stroke) || is.na(n$stroke)) NA_character_ else n$stroke, character(1)),
      color  = vapply(nodes, function(n) if (is.null(n$color)  || is.na(n$color))  NA_character_ else n$color,  character(1)),
      class  = vapply(nodes, function(n) if (is.null(n$class)  || is.na(n$class))  NA_character_ else n$class,  character(1))
    )
  }

  if (length(edges_list) == 0) {
    edges_df <- tibble::tibble(
      from = character(), to = character(), label = character(),
      line_type = character(), arrow_start = character(), arrow_end = character()
    )
  } else {
    edges_df <- tibble::tibble(
      from        = vapply(edges_list, `[[`, character(1), "from"),
      to          = vapply(edges_list, `[[`, character(1), "to"),
      label       = vapply(edges_list, function(e) if (is.null(e$label) || is.na(e$label)) NA_character_ else e$label, character(1)),
      line_type   = vapply(edges_list, `[[`, character(1), "line_type"),
      arrow_start = vapply(edges_list, `[[`, character(1), "arrow_start"),
      arrow_end   = vapply(edges_list, `[[`, character(1), "arrow_end")
    )
  }

  if (length(subgraphs) == 0) {
    sgs_df <- tibble::tibble(
      id = character(), label = character(), node_ids = list(),
      fill = character(), stroke = character()
    )
  } else {
    sgs_df <- tibble::tibble(
      id       = vapply(subgraphs, `[[`, character(1), "id"),
      label    = vapply(subgraphs, `[[`, character(1), "label"),
      node_ids = lapply(subgraphs, `[[`, "node_ids"),
      fill     = vapply(subgraphs, function(sg) if (is.null(sg$fill)   || is.na(sg$fill))   NA_character_ else sg$fill,   character(1)),
      stroke   = vapply(subgraphs, function(sg) if (is.null(sg$stroke) || is.na(sg$stroke)) NA_character_ else sg$stroke, character(1))
    )
  }

  if (length(classes) == 0) {
    classes_df <- tibble::tibble(
      name = character(), fill = character(), stroke = character(),
      color = character(), stroke_width = numeric()
    )
  } else {
    classes_df <- tibble::tibble(
      name         = names(classes),
      fill         = vapply(classes, function(cl) if (is.null(cl$fill)         || is.na(cl$fill))         NA_character_ else cl$fill,         character(1)),
      stroke       = vapply(classes, function(cl) if (is.null(cl$stroke)       || is.na(cl$stroke))       NA_character_ else cl$stroke,       character(1)),
      color        = vapply(classes, function(cl) if (is.null(cl$color)        || is.na(cl$color))        NA_character_ else cl$color,        character(1)),
      stroke_width = vapply(classes, function(cl) if (is.null(cl$stroke_width) || is.na(cl$stroke_width)) NA_real_     else cl$stroke_width,  numeric(1))
    )
  }

  structure(
    list(direction = direction, nodes = nodes_df, edges = edges_df,
         subgraphs = sgs_df, classes = classes_df),
    class = "mermaid_graph"
  )
}

#' @export
print.mermaid_graph <- function(x, ...) {
  cat("Mermaid graph\n")
  cat("  Direction:", x$direction, "\n")
  cat("  Nodes    :", nrow(x$nodes), "\n")
  cat("  Edges    :", nrow(x$edges), "\n")
  cat("  Subgraphs:", nrow(x$subgraphs), "\n")
  cat("  Classes  :", nrow(x$classes), "\n")
  invisible(x)
}

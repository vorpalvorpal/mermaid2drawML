# R/parse_svg.R
#
# Parses the SVG produced by mermaid-cli into two data structures used by
# drawingml.R:
#
#   nodes  — tibble: id, label, shape, svg_cx, svg_cy, svg_w, svg_h,
#                    fill, stroke, class
#   edges  — tibble: id, from, to, label, path_d, arrow_start, arrow_end,
#                    line_type, label_x, label_y
#
# All svg_* coordinates are in raw SVG-user-unit pixels. Conversion to EMU
# is done in drawingml.R using the viewBox scale factor.

# ── Public entry point ─────────────────────────────────────────────────────

#' Parse a mermaid SVG into node and edge geometry tables
#'
#' @param svg Character string containing a complete SVG document as produced
#'   by `@mermaid-js/mermaid-cli`.
#' @param ast List. The AST from `@mermaid-js/parser` (used to enrich nodes
#'   with class/style information not always visible in the SVG).
#' @return A named list with:
#'   - `nodes`: tibble of node geometry
#'   - `edges`: tibble of edge geometry
#'   - `viewbox`: numeric(4) — minx, miny, width, height from the SVG viewBox
#' @keywords internal
parse_mermaid_svg <- function(svg, ast = NULL, source = NULL) {
  doc <- xml2::read_xml(svg)
  xml2::xml_ns_strip(doc)                  # drop SVG namespace for easier xpath

  vb        <- parse_viewbox(doc)
  style     <- extract_svg_style(doc)
  nodes     <- extract_nodes(doc, ast)

  # Override SVG-detected shapes with those declared in the mermaid source text
  # using the v11 @{ shape: X } syntax.
  if (!is.null(source) && nzchar(source)) {
    overrides <- parse_source_shapes(source)
    if (length(overrides) > 0L) {
      for (nid in names(overrides)) {
        rows <- nodes$id == nid
        if (any(rows)) nodes$shape[rows] <- overrides[[nid]]
      }
    }
  }

  subgraphs <- extract_subgraphs(doc, style)
  edges     <- extract_edges(doc, nodes)

  list(nodes = nodes, subgraphs = subgraphs, edges = edges, viewbox = vb, style = style)
}

#' Parse mermaid v11 @{ shape: X } declarations from source text
#'
#' @param code Character string of mermaid source code.
#' @return A named character vector mapping node ID -> internal shape name.
#' @keywords internal
parse_source_shapes <- function(code) {
  if (is.null(code) || !nzchar(code)) return(character(0))

  # Map mermaid shape key -> internal name
  key_to_internal <- c(
    rect="rect", rounded="roundRect", stadium="roundRect",
    diam="diamond", cyl="cylinder", circle="ellipse", "sm-circ"="ellipse",
    hex="hexagon", doc="doc", docs="docs",
    "win-pane"="winPane", "div-rect"="winPane",
    delay="delay", "sl-rect"="manualInput",
    "trap-t"="manualOp", "trap-b"="trapezoid",
    hourglass="collate", "curv-trap"="display",
    "fr-rect"="subprocess", flag="punchedTape",
    "bow-rect"="storedData", "cross-circ"="crossCirc",
    "h-cyl"="hCyl", "lin-cyl"="linCyl",
    tri="extract", "flip-tri"="mergeTri",
    bolt="bolt", "notch-rect"="notchRect",
    bang="bang", cloud="cloud",
    brace="brace", "brace-r"="braceR", braces="braces",
    "lean-r"="leanR", "lean-l"="leanL",
    "notch-pent"="notchPent", "f-circ"="junction",
    fork="fork", "dbl-circ"="dblCirc", "fr-circ"="frCirc",
    "lin-doc"="linDoc", "lin-rect"="linRect",
    odd="odd", "tag-doc"="tagDoc", "tag-rect"="tagRect",
    "st-rect"="stRect", text="text",
    parallelogram="parallelogram"
  )

  # Match NodeId@{ ... shape: key ... } — key may have hyphens
  pattern <- "([A-Za-z0-9_]+)@\\{[^}]*?\\bshape:\\s*([a-z][a-z0-9-]*)"
  m    <- gregexpr(pattern, code, perl = TRUE)
  hits <- regmatches(code, m)[[1]]
  result <- character(0)
  for (h in hits) {
    parts <- regmatches(h, regexec(pattern, h, perl = TRUE))[[1]]
    if (length(parts) == 3L) {
      nid   <- parts[2]
      key   <- parts[3]
      iname <- key_to_internal[key]
      if (!is.na(iname)) result[nid] <- iname
    }
  }
  result
}

# ── viewBox ────────────────────────────────────────────────────────────────

parse_viewbox <- function(doc) {
  vb_str <- xml2::xml_attr(doc, "viewBox")
  if (is.na(vb_str) || !nzchar(vb_str)) {
    # Fall back to width/height attributes
    w <- as.numeric(xml2::xml_attr(doc, "width")  %||% "800")
    h <- as.numeric(xml2::xml_attr(doc, "height") %||% "600")
    return(c(0, 0, w, h))
  }
  as.numeric(strsplit(trimws(vb_str), "[,\\s]+", perl = TRUE)[[1]])
}

# ── SVG stylesheet extraction ─────────────────────────────────────────────

# Reads the SVG <style> block for values that mermaid sets via CSS classes
# rather than inline attributes, so they would otherwise be invisible to the
# attribute-level parsers below.
#
# Returns:
#   font_size_px   — first font-size: Npx found (drives scaled Word font size)
#   edge_stroke    — 6-char hex from .flowchart-link { stroke: ... }
#   edge_sw_px     — stroke-width in px from .flowchart-link { stroke-width: ... }

extract_svg_style <- function(doc) {
  style_text <- ""
  style_el   <- xml2::xml_find_first(doc, ".//style")
  if (!inherits(style_el, "xml_missing")) style_text <- xml2::xml_text(style_el)

  # Font size ---------------------------------------------------------------
  font_size_px <- 14   # mermaid default when no %%{init}%% block is present
  m <- regmatches(style_text,
    regexpr("font-size:\\s*(\\d+(?:\\.\\d+)?)px", style_text, perl = TRUE))
  if (length(m) > 0L) {
    fs <- suppressWarnings(as.numeric(sub("px$", "", sub(".*font-size:\\s*", "", m[[1]]))))
    if (!is.na(fs) && fs > 0) font_size_px <- fs
  }

  # Font family -------------------------------------------------------------
  # Mermaid uses TWO font declarations that serve different purposes:
  #
  #   1. font-family: <value>  on #my-svg and .label — set by themeVariables
  #      fontFamily (e.g. "Myriad Pro").  Applies to <text> elements such as
  #      subgraph labels, but NOT to <foreignObject> node-label content.
  #
  #   2. --mermaid-font-family: "trebuchet ms",verdana,...  set via :root{} —
  #      a CSS custom property that mermaid's internal CSS uses for the actual
  #      <foreignObject> text (node labels).  This is always the mermaid
  #      built-in stack and is NOT affected by themeVariables.fontFamily.
  #
  # Node sizes (foreignObject dimensions) are therefore measured using the
  # --mermaid-font-family stack, NOT the declared font-family.  We extract
  # BOTH so callers can present the correct measurement font to Word (avoiding
  # metric mismatches) while still exposing the user's declared font if needed.
  .font_name_map <- c(
    "trebuchet ms"   = "Trebuchet MS",
    "verdana"        = "Verdana",
    "arial"          = "Arial",
    "helvetica"      = "Helvetica",
    "helvetica neue" = "Helvetica Neue",
    "times new roman"= "Times New Roman",
    "georgia"        = "Georgia",
    "courier new"    = "Courier New",
    "courier"        = "Courier",
    "tahoma"         = "Tahoma",
    "calibri"        = "Calibri",
    "segoe ui"       = "Segoe UI"
  )

  .first_font <- function(raw) {
    first <- trimws(strsplit(raw, ",")[[1]][1])
    first <- tolower(gsub('"', '', first, fixed = TRUE))
    first <- trimws(first)
    if (!nzchar(first) || first %in% c("sans-serif","serif","monospace","inherit"))
      return(NA_character_)
    if (first %in% names(.font_name_map)) return(.font_name_map[[first]])
    # Title-case unknown font names (e.g. "open sans" -> "Open Sans")
    gsub("(^|\\s)(\\S)", "\\1\\U\\2", first, perl = TRUE)
  }

  # 1. Declared font (themeVariables.fontFamily) — from font-family: <value>
  declared_font_family <- "Trebuchet MS"
  m <- regexec('font-family\\s*:\\s*([^;}\\n]+)', style_text, perl = TRUE)
  hit <- regmatches(style_text, m)[[1]]
  if (length(hit) == 2L) {
    f <- .first_font(hit[2])
    if (!is.na(f)) declared_font_family <- f
  }

  # 2. Measurement font (--mermaid-font-family CSS variable) — always Trebuchet MS
  #    unless mermaid changes its built-in default.  This is the font that
  #    Chromium actually used to measure node label widths.
  measurement_font_family <- "Trebuchet MS"
  m2 <- regexec('--mermaid-font-family\\s*:\\s*([^;}\\n]+)',
                style_text, perl = TRUE)
  hit2 <- regmatches(style_text, m2)[[1]]
  if (length(hit2) == 2L) {
    f <- .first_font(hit2[2])
    if (!is.na(f)) measurement_font_family <- f
  }

  # font_family is the measurement font: this is what should go in <w:rFonts>
  # to ensure Word text fits the SVG-measured node boxes without clipping.
  font_family <- measurement_font_family

  # Edge line colour --------------------------------------------------------
  # Mermaid emits:  .flowchart-link { stroke: #RRGGBB; fill: none; }
  edge_stroke <- "5E504E"   # Bark grey — mermaid's default lineColor
  m <- regmatches(style_text,
    regexpr("\\.flowchart-link[^{]*\\{[^}]*stroke\\s*:\\s*#([0-9A-Fa-f]{6})",
            style_text, perl = TRUE))
  if (length(m) > 0L) {
    col <- regmatches(m[[1]], regexpr("[0-9A-Fa-f]{6}$", m[[1]], perl = TRUE))
    if (length(col) > 0L) edge_stroke <- toupper(col)
  }

  # Edge stroke-width -------------------------------------------------------
  edge_sw_px <- 2
  m <- regmatches(style_text,
    regexpr("\\.flowchart-link[^{]*\\{[^}]*stroke-width\\s*:\\s*(\\d+(?:\\.\\d+)?)px",
            style_text, perl = TRUE))
  if (length(m) > 0L) {
    sw <- suppressWarnings(
      as.numeric(sub("px.*", "", sub(".*stroke-width\\s*:\\s*", "", m[[1]])))
    )
    if (!is.na(sw) && sw > 0) edge_sw_px <- sw
  }

  # Cluster default fill / stroke -----------------------------------------
  # Mermaid emits:  .cluster rect { fill: #ffffde; stroke: #aaaa33; ... }
  cluster_fill   <- "FFFFDE"
  cluster_stroke <- "AAAA33"
  m <- regmatches(style_text,
    regexpr("\\.cluster rect\\s*\\{([^}]*?)\\}", style_text, perl = TRUE))
  if (length(m) > 0L) {
    blk <- m[[1]]
    mf <- regmatches(blk,
      regexpr("fill:\\s*#([0-9A-Fa-f]{6})", blk, perl = TRUE))
    if (length(mf) > 0L) {
      col <- regmatches(mf[[1]], regexpr("[0-9A-Fa-f]{6}$", mf[[1]], perl = TRUE))
      if (length(col) > 0L) cluster_fill <- toupper(col)
    }
    ms <- regmatches(blk,
      regexpr("stroke:\\s*#([0-9A-Fa-f]{6})", blk, perl = TRUE))
    if (length(ms) > 0L) {
      col <- regmatches(ms[[1]], regexpr("[0-9A-Fa-f]{6}$", ms[[1]], perl = TRUE))
      if (length(col) > 0L) cluster_stroke <- toupper(col)
    }
  }

  list(font_size_px          = font_size_px,
       font_family            = font_family,           # measurement font (--mermaid-font-family)
       declared_font_family   = declared_font_family,  # themeVariables fontFamily
       edge_stroke            = edge_stroke,
       edge_sw_px             = edge_sw_px,
       cluster_fill           = cluster_fill,
       cluster_stroke         = cluster_stroke)
}

# ── Node extraction ────────────────────────────────────────────────────────

extract_nodes <- function(doc, ast) {
  # Mermaid renders each node as <g class="node ..."> with an id attribute.
  # The id is like "flowchart-A-12" or "mermaid-A" — we parse the mermaid id.
  node_gs <- xml2::xml_find_all(doc, ".//*[contains(@class,'node')][@id]")
  # Filter to direct node groups (not sub-elements that also happen to have class node)
  node_gs <- Filter(function(g) {
    cl <- xml2::xml_attr(g, "class") %||% ""
    grepl("\\bnode\\b", cl) && !grepl("\\bedgeLabel\\b", cl)
  }, node_gs)

  if (length(node_gs) == 0L) return(empty_nodes_tbl())

  rows <- lapply(node_gs, parse_one_node)
  rows <- Filter(Negate(is.null), rows)

  if (length(rows) == 0L) return(empty_nodes_tbl())

  # Enrich with AST class/style info
  tbl <- tibble::tibble(
    id       = vapply(rows, `[[`, character(1), "id"),
    label    = vapply(rows, `[[`, character(1), "label"),
    shape    = vapply(rows, `[[`, character(1), "shape"),
    svg_cx   = vapply(rows, `[[`, numeric(1),   "svg_cx"),
    svg_cy   = vapply(rows, `[[`, numeric(1),   "svg_cy"),
    svg_w    = vapply(rows, `[[`, numeric(1),   "svg_w"),
    svg_h    = vapply(rows, `[[`, numeric(1),   "svg_h"),
    txt_w    = vapply(rows, `[[`, numeric(1),   "txt_w"),
    txt_h    = vapply(rows, `[[`, numeric(1),   "txt_h"),
    fill     = vapply(rows, `[[`, character(1), "fill"),
    stroke   = vapply(rows, `[[`, character(1), "stroke"),
    color    = vapply(rows, `[[`, character(1), "color"),
    class    = vapply(rows, `[[`, character(1), "class")
  )

  tbl <- enrich_from_ast(tbl, ast)
  tbl
}

parse_one_node <- function(g) {
  svg_id <- xml2::xml_attr(g, "id") %||% ""
  # Extract the mermaid node id: "flowchart-A-12" → "A"
  mmd_id <- extract_mermaid_id(svg_id)
  if (!nzchar(mmd_id)) return(NULL)

  # CSS classes on the group — used for shape hints and class name
  classes <- strsplit(xml2::xml_attr(g, "class") %||% "", "\\s+")[[1]]

  # Position from transform="translate(cx, cy)" on the node group, plus any
  # cumulative translate from ancestor groups (e.g. the per-subgraph root groups
  # that mermaid emits when a diagram has multiple layout roots).
  tf  <- xml2::xml_attr(g, "transform") %||% ""
  pos <- parse_translate(tf) + ancestor_offset(g)

  # Find the primary shape element (first child that is a geometry element)
  shape_el <- find_shape_element(g)
  if (is.null(shape_el)) return(NULL)

  geom   <- element_geometry(shape_el, pos)  # list(shape, svg_cx, svg_cy, svg_w, svg_h)
  label  <- extract_node_label(g)
  fill   <- extract_fill(shape_el, g)
  stroke <- extract_stroke(shape_el, g)
  color  <- extract_text_color(shape_el, g)
  txtdim <- extract_label_dim(g)             # list(w, h) of the foreignObject text area

  # Class name assigned in mermaid (e.g. ":::plan") — found in the g's classes
  # Mermaid adds class names after "default" and known shape-markers
  known_shape_markers <- c("node", "default", "flowchart-label",
                            "rhombus", "circle", "hexagon", "rect",
                            "rounded", "stadium", "cylinder",
                            "subgraph", "cluster")
  user_classes <- setdiff(classes, known_shape_markers)
  user_class   <- if (length(user_classes) > 0L) user_classes[1L] else NA_character_

  list(
    id      = mmd_id,
    label   = label,
    shape   = geom$shape,
    svg_cx  = geom$svg_cx,
    svg_cy  = geom$svg_cy,
    svg_w   = geom$svg_w,
    svg_h   = geom$svg_h,
    txt_w   = txtdim$w,   # foreignObject width  (NA if not found)
    txt_h   = txtdim$h,   # foreignObject height (NA if not found)
    fill    = fill   %||% NA_character_,
    stroke  = stroke %||% NA_character_,
    color   = color  %||% NA_character_,
    class   = user_class
  )
}

extract_mermaid_id <- function(svg_id) {
  # Strip flowchart-/mermaid- prefix, then strip trailing -<digits> counter.
  # "flowchart-StartNode-3" → "StartNode-3" → "StartNode"
  # "flowchart-A-12"        → "A-12"        → "A"
  # "mermaid-B"             → "B"           → "B"  (no numeric suffix)
  s <- sub("^(?:flowchart|mermaid)-", "", svg_id, perl = TRUE)
  s <- sub("-\\d+$", "", s, perl = TRUE)
  s
}

parse_translate <- function(tf) {
  m <- regmatches(tf, regexpr("translate\\(([^)]+)\\)", tf, perl = TRUE))
  if (length(m) == 0L) return(c(0, 0))
  inner <- sub("translate\\(", "", sub("\\)", "", m))
  parts <- as.numeric(strsplit(trimws(inner), "[,\\s]+", perl = TRUE)[[1]])
  if (length(parts) < 2L) parts <- c(parts, 0)
  parts[1:2]
}

# Sum all translate() transforms on ancestors of `el` (excluding el itself).
ancestor_offset <- function(el) {
  off <- c(0, 0)
  for (anc in xml2::xml_parents(el)) {
    tf <- xml2::xml_attr(anc, "transform") %||% ""
    if (nzchar(tf)) off <- off + parse_translate(tf)
  }
  off
}

# Shift all coordinates in an SVG path string by (dx, dy).
# Assumes absolute (uppercase) commands using x,y pairs — M, L, C, S, Q, T —
# which is what mermaid-cli/ELK produces for flowchart edges.
offset_svg_path <- function(d, dx, dy) {
  if (dx == 0 && dy == 0) return(d)
  pattern <- "-?\\d+(?:\\.\\d+)?(?:[eE][+-]?\\d+)?"
  m <- gregexpr(pattern, d, perl = TRUE)
  nums <- as.numeric(regmatches(d, m)[[1]])
  offsets <- rep(c(dx, dy), length.out = length(nums))
  regmatches(d, m)[[1]] <- as.character(round(nums + offsets, 4))
  d
}

find_shape_element <- function(g) {
  # Mermaid marks the actual node shape with class="label-container" (or
  # "basic label-container"). Using this avoids accidentally picking up the
  # empty <rect/> placeholder that mermaid inserts inside <g class="label">,
  # which has no geometry attributes and would produce zero-size shapes.
  el <- xml2::xml_find_first(g, ".//*[contains(@class,'label-container')]")
  if (!inherits(el, "xml_missing")) return(el)

  # Fallback: search for geometry elements that are NOT inside the label group
  for (tag in c("rect", "polygon", "ellipse", "circle", "path")) {
    el <- xml2::xml_find_first(
      g, paste0(".//", tag, "[not(ancestor::*[contains(@class,'label')])]"))
    if (!inherits(el, "xml_missing")) return(el)
  }
  NULL
}

element_geometry <- function(el, pos) {
  tag <- xml2::xml_name(el)
  cx  <- pos[1]; cy <- pos[2]

  # Capture any transform on the element itself (e.g., path label-containers
  # have translate(-w/2, -h/2) to centre themselves on the node's translate point)
  el_tf  <- xml2::xml_attr(el, "transform") %||% ""
  el_off <- parse_translate(el_tf)   # c(dx, dy)

  if (tag == "rect") {
    w  <- as.numeric(xml2::xml_attr(el, "width")  %||% "0")
    h  <- as.numeric(xml2::xml_attr(el, "height") %||% "0")
    rx <- as.numeric(xml2::xml_attr(el, "rx")     %||% "0")
    # x/y are relative to translate point (which is the centre in mermaid)
    svg_cx <- cx          # already the centre
    svg_cy <- cy
    shape  <- if (!is.na(rx) && rx > 1) "roundRect" else "rect"
    list(shape = shape, svg_cx = svg_cx, svg_cy = svg_cy, svg_w = w, svg_h = h)

  } else if (tag == "polygon") {
    pts  <- parse_polygon_points(xml2::xml_attr(el, "points") %||% "")
    if (nrow(pts) == 0L) return(list(shape="rect", svg_cx=cx, svg_cy=cy, svg_w=60, svg_h=30))
    w    <- max(pts$x) - min(pts$x)
    h    <- max(pts$y) - min(pts$y)
    # Classify by number of distinct y-levels: 2→diamond, 3→hexagon-ish
    n_pts <- nrow(pts)
    shape <- if (n_pts == 4L) "diamond" else if (n_pts == 6L) "hexagon" else "rect"
    list(shape = shape, svg_cx = cx, svg_cy = cy, svg_w = w, svg_h = h)

  } else if (tag == "ellipse") {
    rx <- as.numeric(xml2::xml_attr(el, "rx") %||% "30")
    ry <- as.numeric(xml2::xml_attr(el, "ry") %||% "20")
    list(shape = "ellipse", svg_cx = cx, svg_cy = cy, svg_w = rx * 2, svg_h = ry * 2)

  } else if (tag == "circle") {
    r  <- as.numeric(xml2::xml_attr(el, "r") %||% "25")
    list(shape = "ellipse", svg_cx = cx, svg_cy = cy, svg_w = r * 2, svg_h = r * 2)

  } else if (tag == "g") {
    # Mermaid v11 renders all shapes as <g class="label-container ..."> groups
    # containing rough/sketch path elements.
    #   "outer-path" class  → roundRect / stadium shape
    #   1 child path, no outer-path → doc (folded-corner document page)
    #   2 child paths, no outer-path → docs (stacked document pages)
    el_cls     <- xml2::xml_attr(el, "class") %||% ""
    n_children <- length(xml2::xml_find_all(el, ".//path"))
    # win-pane: 1 child path + positive x/y translate (e.g. translate(2.5, 2.5))
    el_tf_str <- xml2::xml_attr(el, "transform") %||% ""
    el_off_g  <- parse_translate(el_tf_str)
    is_winpane <- n_children == 1L && el_off_g[1] > 0 && el_off_g[2] > 0
    shape <- if (grepl("outer-path", el_cls, fixed = TRUE)) {
      "roundRect"
    } else if (is_winpane) {
      "winPane"
    } else if (n_children == 1L) {
      "doc"
    } else if (n_children >= 2L) {
      "docs"
    } else {
      "rect"
    }
    bb     <- group_element_bbox(el)
    if (is.null(bb))
      return(list(shape = shape, svg_cx = cx + el_off[1], svg_cy = cy + el_off[2],
                  svg_w = 60, svg_h = 30))
    list(shape  = shape,
         svg_cx = cx + el_off[1] + bb$cx,
         svg_cy = cy + el_off[2] + bb$cy,
         svg_w  = bb$w,
         svg_h  = bb$h)

  } else {
    # path (or other element): bbox via path parser.
    # Account for the element's own transform (e.g., cylinder label-containers
    # use translate(-rx, -cy_offset) to centre the path on the node origin).
    d     <- xml2::xml_attr(el, "d") %||% ""
    bb    <- path_bbox(d)
    # Detect cylinder: path uses arc commands
    shape <- if (grepl("[Aa]", d, perl = TRUE)) "cylinder" else "rect"
    list(shape  = shape,
         svg_cx = cx + el_off[1] + bb$cx,
         svg_cy = cy + el_off[2] + bb$cy,
         svg_w  = bb$w,
         svg_h  = bb$h)
  }
}

# Aggregate bounding box from all descendant <path> elements inside a <g>.
# Returns list(cx, cy, w, h) in the g's local coordinate space,
# or NULL if no usable paths are found.
group_element_bbox <- function(g_el) {
  min_x <- Inf; max_x <- -Inf
  min_y <- Inf; max_y <- -Inf

  child_paths <- xml2::xml_find_all(g_el, ".//path")
  for (p in child_paths) {
    p_tf  <- xml2::xml_attr(p, "transform") %||% ""
    p_off <- parse_translate(p_tf)
    d     <- xml2::xml_attr(p, "d") %||% ""
    if (!nzchar(d)) next
    bb <- path_bbox(d)
    if (bb$w <= 1 && bb$h <= 1) next   # skip degenerate paths
    hw <- bb$w / 2; hh <- bb$h / 2
    lx <- p_off[1] + bb$cx - hw;  rx <- p_off[1] + bb$cx + hw
    ty <- p_off[2] + bb$cy - hh;  by <- p_off[2] + bb$cy + hh
    if (lx < min_x) min_x <- lx
    if (rx > max_x) max_x <- rx
    if (ty < min_y) min_y <- ty
    if (by > max_y) max_y <- by
  }

  if (is.infinite(min_x)) return(NULL)
  list(cx = (min_x + max_x) / 2,
       cy = (min_y + max_y) / 2,
       w  = max(max_x - min_x, 1),
       h  = max(max_y - min_y, 1))
}

parse_polygon_points <- function(pts_str) {
  if (!nzchar(pts_str)) return(data.frame(x = numeric(0), y = numeric(0)))
  pairs <- strsplit(trimws(pts_str), "[,\\s]+", perl = TRUE)[[1]]
  # pairs should be even length: x1 y1 x2 y2 ...
  nums  <- suppressWarnings(as.numeric(pairs))
  nums  <- nums[!is.na(nums)]
  if (length(nums) < 2L) return(data.frame(x = numeric(0), y = numeric(0)))
  n <- length(nums) %/% 2L
  data.frame(x = nums[seq(1, by = 2L, length.out = n)],
             y = nums[seq(2, by = 2L, length.out = n)])
}

extract_node_label <- function(g) {
  # Try <foreignObject> → inner text, then <text>
  fo <- xml2::xml_find_first(g, ".//foreignObject")
  if (!inherits(fo, "xml_missing")) {
    txt <- trimws(xml2::xml_text(fo))
    if (nzchar(txt)) return(txt)
  }
  txt_el <- xml2::xml_find_first(g, ".//text")
  if (!inherits(txt_el, "xml_missing")) {
    return(trimws(xml2::xml_text(txt_el)))
  }
  ""
}

extract_text_color <- function(shape_el, g) {
  # Mermaid classDefs set `color:#XXXXXX` which ends up as a CSS `color:`
  # property on the inner <g class="label"> child, NOT on the outer node <g>
  # or the shape element itself. Check elements in priority order:
  #   1. <g class="label"> style (most reliable for classDef colour)
  #   2. <foreignObject> / <span> descendant style (htmlLabels:true)
  #   3. <text> fill attribute (htmlLabels:false fallback)

  # 1. Inner label group
  label_g <- xml2::xml_find_first(g, ".//*[contains(@class,'label')][@style]")
  if (!inherits(label_g, "xml_missing")) {
    style <- xml2::xml_attr(label_g, "style") %||% ""
    m <- regmatches(style, regexpr("(?:^|;|\\s)color:\\s*([^;!]+)", style, perl = TRUE))
    if (length(m) > 0L) {
      col <- trimws(sub(".*color:\\s*", "", m))
      if (nzchar(col) && col != "none") return(parse_css_colour(col))
    }
  }

  # 2. <span> or <div> inside foreignObject
  for (tag in c("span", "div")) {
    el <- xml2::xml_find_first(g, paste0(".//", tag, "[@style]"))
    if (!inherits(el, "xml_missing")) {
      style <- xml2::xml_attr(el, "style") %||% ""
      m <- regmatches(style, regexpr("(?:^|;|\\s)color:\\s*([^;!]+)", style, perl = TRUE))
      if (length(m) > 0L) {
        col <- trimws(sub(".*color:\\s*", "", m))
        if (nzchar(col) && col != "none") return(parse_css_colour(col))
      }
    }
  }

  # 3. <text> fill (SVG text colour = fill, used when htmlLabels:false)
  txt_el <- xml2::xml_find_first(g, ".//text")
  if (!inherits(txt_el, "xml_missing")) {
    style <- xml2::xml_attr(txt_el, "style") %||% ""
    m <- regmatches(style, regexpr("(?:^|;)\\s*fill:\\s*([^;]+)", style, perl = TRUE))
    if (length(m) > 0L) {
      col <- trimws(sub(".*fill:\\s*", "", m))
      if (nzchar(col) && col != "none") return(parse_css_colour(col))
    }
    fa <- xml2::xml_attr(txt_el, "fill") %||% ""
    if (nzchar(fa) && fa != "none") return(parse_css_colour(fa))
  }

  NA_character_
}

# Strip the CSS `!important` annotation and surrounding whitespace so that
# values like "transparent !important" or "#700017 !important" parse cleanly.
.strip_important <- function(x) trimws(sub("\\s*!important.*$", "", x, perl = TRUE))

extract_fill <- function(shape_el, g) {
  # Mermaid v11 puts fill on child path elements inside the label-container <g>
  candidates <- if (xml2::xml_name(shape_el) == "g") {
    child_path <- xml2::xml_find_first(shape_el, ".//path[@style]")
    if (!inherits(child_path, "xml_missing"))
      list(child_path, shape_el, g)
    else
      list(shape_el, g)
  } else {
    list(shape_el, g)
  }
  for (el in candidates) {
    style <- xml2::xml_attr(el, "style") %||% ""
    m <- regmatches(style, regexpr("(?:^|;)\\s*fill:\\s*([^;]+)", style, perl = TRUE))
    if (length(m) > 0L) {
      col <- .strip_important(sub(".*fill:\\s*", "", m))
      if (nzchar(col)) return(parse_css_colour(col))   # "transparent" → NA via named_colours
    }
    fa <- .strip_important(xml2::xml_attr(el, "fill") %||% "")
    if (nzchar(fa)) return(parse_css_colour(fa))
  }
  NA_character_
}

extract_stroke <- function(shape_el, g) {
  # Mermaid v11 puts stroke on child path elements inside the label-container <g>
  candidates <- if (xml2::xml_name(shape_el) == "g") {
    child_path <- xml2::xml_find_first(shape_el, ".//path[@style]")
    if (!inherits(child_path, "xml_missing"))
      list(child_path, shape_el, g)
    else
      list(shape_el, g)
  } else {
    list(shape_el, g)
  }
  for (el in candidates) {
    style <- xml2::xml_attr(el, "style") %||% ""
    m <- regmatches(style, regexpr("(?:^|;)\\s*stroke:\\s*([^;]+)", style, perl = TRUE))
    if (length(m) > 0L) {
      col <- .strip_important(sub(".*stroke:\\s*", "", m))
      # 8-char hex (#RRGGBBAA) — hex_to_drawingml already strips alpha
      if (nzchar(col) && col != "none" && col != "transparent") return(parse_css_colour(col))
    }
    sa <- .strip_important(xml2::xml_attr(el, "stroke") %||% "")
    if (nzchar(sa) && sa != "none" && sa != "transparent") return(parse_css_colour(sa))
  }
  NA_character_
}

enrich_from_ast <- function(tbl, ast) {
  # NOT YET IMPLEMENTED.
  #
  # This stub exists to wire up the AST from @mermaid-js/parser for future use.
  # Currently, all node geometry, labels, colours, and shapes are read directly
  # from the mermaid SVG output, which is the sole source of truth for the
  # conversion. The AST is not consulted and the returned tibble is unchanged.
  #
  # What this function is intended to do eventually:
  #   - Recover the original `classDef` name for each node (the SVG only shows
  #     the computed fill/stroke colours, not the class name that produced them).
  #   - Supplement inline style information that mermaid-cli flattens into SVG
  #     attributes in ways that are hard to reverse.
  #   - Feed richer data into the round-trip metadata JSON stored in each shape's
  #     wp:docPr descr attribute, so a future Word→Mermaid converter can
  #     reconstruct classDef assignments rather than just raw hex colours.
  #
  # Practical note on @mermaid-js/parser v0.3.x:
  #   The parser only covers a subset of diagram types (architecture, gitGraph,
  #   info, packet, pie). For flowcharts — the most common diagram type — it
  #   returns a _parseError sentinel. Even for supported types, the AST has not
  #   been used here yet. Removing the `ast` parameter from the call chain would
  #   produce identical output today.
  if (is.null(ast) || !is.null(ast[["_parseError"]])) return(tbl)
  tbl
}

empty_nodes_tbl <- function() {
  tibble::tibble(
    id = character(), label = character(), shape = character(),
    svg_cx = numeric(), svg_cy = numeric(),
    svg_w = numeric(), svg_h = numeric(),
    txt_w = numeric(), txt_h = numeric(),
    fill = character(), stroke = character(), color = character(), class = character()
  )
}

# Returns the width and height of the <foreignObject> inside the node's label
# group. These are the exact pixel dimensions mermaid used for the text content,
# useful for sizing a DrawingML text overlay that avoids shape-geometry clipping.
extract_label_dim <- function(g) {
  fo <- xml2::xml_find_first(g, ".//foreignObject")
  if (inherits(fo, "xml_missing"))
    return(list(w = NA_real_, h = NA_real_))
  w <- suppressWarnings(as.numeric(xml2::xml_attr(fo, "width")  %||% NA_character_))
  h <- suppressWarnings(as.numeric(xml2::xml_attr(fo, "height") %||% NA_character_))
  list(w = w, h = h)
}

# ── Subgraph extraction ────────────────────────────────────────────────────

# Mermaid renders subgraphs as <g class="cluster [userClass]" id="SubgraphID">.
# The rect child gives the absolute SVG position (no transform on the <g>).
# "subSpace" clusters are invisible layout-spacers and are filtered out.

extract_subgraphs <- function(doc, style) {
  default_stroke <- style$cluster_stroke %||% "AAAA33"

  gs <- xml2::xml_find_all(doc, ".//g[contains(@class,'cluster')][@id]")
  gs <- Filter(function(g) {
    cl <- xml2::xml_attr(g, "class") %||% ""
    grepl("\\bcluster\\b", cl) && !grepl("\\bcluster-label\\b", cl)
  }, gs)

  rows <- lapply(gs, function(g) {
    svg_id  <- xml2::xml_attr(g, "id") %||% ""
    if (!nzchar(svg_id)) return(NULL)

    classes <- strsplit(xml2::xml_attr(g, "class") %||% "", "\\s+")[[1]]
    if ("subSpace" %in% classes) return(NULL)  # invisible spacer

    rect_el <- xml2::xml_find_first(g, "./rect")
    if (inherits(rect_el, "xml_missing")) return(NULL)

    x <- as.numeric(xml2::xml_attr(rect_el, "x")     %||% "0")
    y <- as.numeric(xml2::xml_attr(rect_el, "y")     %||% "0")
    w <- as.numeric(xml2::xml_attr(rect_el, "width") %||% "0")
    h <- as.numeric(xml2::xml_attr(rect_el, "height")%||% "0")
    if (w <= 0 || h <= 0) return(NULL)
    anc_off <- ancestor_offset(g)
    x <- x + anc_off[1]
    y <- y + anc_off[2]

    fill   <- extract_fill(rect_el, g)   # NA → <a:noFill/> (transparent border-only cluster)
    stroke <- extract_stroke(rect_el, g)

    # stroke-width and text color from the rect's inline style
    # (mermaid inlines `style="fill:... stroke:... stroke-width:Npx color:..."`)
    rect_style <- xml2::xml_attr(rect_el, "style") %||% ""

    sw_px <- NA_real_
    m_sw  <- regmatches(rect_style,
               regexpr("stroke-width:\\s*(\\d+(?:\\.\\d+)?)px", rect_style, perl = TRUE))
    if (length(m_sw) > 0L) {
      v <- suppressWarnings(as.numeric(sub("px.*", "", sub(".*stroke-width:\\s*", "", m_sw))))
      if (!is.na(v) && v > 0) sw_px <- v
    }

    sg_color <- NA_character_
    m_col    <- regmatches(rect_style,
                  regexpr("(?:^|;)\\s*color:\\s*([^;!]+)", rect_style, perl = TRUE))
    if (length(m_col) > 0L) {
      col <- .strip_important(sub(".*color:\\s*", "", m_col))
      if (nzchar(col)) sg_color <- parse_css_colour(col)
    }
    # Fallback: use stroke color for label text when no explicit color property
    if (is.na(sg_color)) sg_color <- stroke %||% NA_character_

    # Label text from cluster-label child
    label_g <- xml2::xml_find_first(g, ".//*[contains(@class,'cluster-label')]")
    label   <- if (!inherits(label_g, "xml_missing")) trimws(xml2::xml_text(label_g)) else ""

    known       <- c("cluster", "subSpace")
    user_class  <- setdiff(classes, known)
    user_class  <- if (length(user_class) > 0L) user_class[1L] else NA_character_

    list(
      id           = svg_id,
      label        = label,
      svg_x        = x,
      svg_y        = y,
      svg_w        = w,
      svg_h        = h,
      fill         = fill,                   # keep NA for transparent
      stroke       = stroke %||% default_stroke,
      stroke_width = sw_px,
      color        = sg_color,
      class        = user_class
    )
  })

  rows <- Filter(Negate(is.null), rows)
  if (length(rows) == 0L) return(empty_subgraphs_tbl())

  tibble::tibble(
    id           = vapply(rows, `[[`, character(1), "id"),
    label        = vapply(rows, `[[`, character(1), "label"),
    svg_x        = vapply(rows, `[[`, numeric(1),   "svg_x"),
    svg_y        = vapply(rows, `[[`, numeric(1),   "svg_y"),
    svg_w        = vapply(rows, `[[`, numeric(1),   "svg_w"),
    svg_h        = vapply(rows, `[[`, numeric(1),   "svg_h"),
    fill         = vapply(rows, `[[`, character(1), "fill"),
    stroke       = vapply(rows, `[[`, character(1), "stroke"),
    stroke_width = vapply(rows, function(r) r$stroke_width %||% NA_real_, numeric(1)),
    color        = vapply(rows, function(r) r$color %||% NA_character_,   character(1)),
    class        = vapply(rows, `[[`, character(1), "class")
  )
}

empty_subgraphs_tbl <- function() {
  tibble::tibble(
    id = character(), label = character(),
    svg_x = numeric(), svg_y = numeric(),
    svg_w = numeric(), svg_h = numeric(),
    fill = character(), stroke = character(),
    stroke_width = numeric(), color = character(), class = character()
  )
}

# ── Edge extraction ────────────────────────────────────────────────────────

extract_edges <- function(doc, nodes_tbl) {
  # Mermaid renders edge paths in <g class="edgePaths"> or similar.
  # Each path has id like "L-A-B-0".
  # Edge labels are in <g class="edgeLabels"> or "edgeLabel".

  paths <- xml2::xml_find_all(doc, ".//*[self::path or self::polyline][contains(@id,'L-') or contains(@id,'L_') or contains(@class,'edge')]")

  # Also look for paths inside edgePaths containers
  edge_containers <- xml2::xml_find_all(doc, ".//*[contains(@class,'edgePath') or contains(@class,'edgePaths')]")
  path_els <- c(paths, unlist(lapply(edge_containers, function(ec) {
    as.list(xml2::xml_find_all(ec, ".//path"))
  }), recursive = FALSE))

  # Deduplicate by pointer
  seen <- character(0)
  path_els <- Filter(function(p) {
    pid <- xml2::xml_attr(p, "id") %||% ""
    d   <- xml2::xml_attr(p, "d")  %||% ""
    key <- paste0(pid, "|", substr(d, 1, 20))
    if (key %in% seen) return(FALSE)
    seen <<- c(seen, key)
    nzchar(d)
  }, path_els)

  if (length(path_els) == 0L) return(empty_edges_tbl())

  rows <- lapply(path_els, function(p) parse_one_edge(p, doc, nodes_tbl))
  rows <- Filter(Negate(is.null), rows)
  if (length(rows) == 0L) return(empty_edges_tbl())

  tibble::tibble(
    id          = vapply(rows, `[[`, character(1), "id"),
    from        = vapply(rows, `[[`, character(1), "from"),
    to          = vapply(rows, `[[`, character(1), "to"),
    label       = vapply(rows, function(r) r$label %||% NA_character_, character(1)),
    path_d      = vapply(rows, `[[`, character(1), "path_d"),
    arrow_start = vapply(rows, `[[`, character(1), "arrow_start"),
    arrow_end   = vapply(rows, `[[`, character(1), "arrow_end"),
    line_type   = vapply(rows, `[[`, character(1), "line_type"),
    stroke      = vapply(rows, function(r) r$stroke %||% NA_character_, character(1)),
    label_x     = vapply(rows, function(r) r$label_x %||% NA_real_, numeric(1)),
    label_y     = vapply(rows, function(r) r$label_y %||% NA_real_, numeric(1))
  )
}

parse_one_edge <- function(path_el, doc, nodes_tbl) {
  edge_id <- xml2::xml_attr(path_el, "id") %||% ""
  d       <- xml2::xml_attr(path_el, "d")  %||% ""
  if (!nzchar(d)) return(NULL)
  anc_off <- ancestor_offset(path_el)
  d       <- offset_svg_path(d, anc_off[1], anc_off[2])

  # Parse from/to — support both mermaid v11 underscore format (L_A_B_0_0)
  # and older hyphen format (L-A-B-0). Node IDs are assumed not to contain _.
  from <- NA_character_; to <- NA_character_
  m <- regmatches(edge_id,
        regexec("^L_([^_]+)_([^_]+)_\\d+_\\d+$", edge_id, perl = TRUE))[[1]]
  if (length(m) == 3L) {
    from <- m[2]; to <- m[3]
  } else {
    m <- regmatches(edge_id, regexec("L-([^-]+)-([^-]+)", edge_id, perl = TRUE))[[1]]
    if (length(m) == 3L) { from <- m[2]; to <- m[3] }
  }

  # Arrow type from marker attributes
  me  <- xml2::xml_attr(path_el, "marker-end")   %||% ""
  ms  <- xml2::xml_attr(path_el, "marker-start") %||% ""
  arrow_end   <- if (nzchar(me)) classify_marker(me) else "none"
  arrow_start <- if (nzchar(ms)) classify_marker(ms) else "none"

  # Line type: use new mermaid v11 CSS class names (edge-pattern-*) first,
  # then fall back to legacy class/style heuristics.
  # "thick" must match edge-thickness-thick specifically (not edge-thickness-normal).
  cl    <- xml2::xml_attr(path_el, "class") %||% ""
  style <- xml2::xml_attr(path_el, "style") %||% ""
  line_type <- if (grepl("edge-pattern-dashed|edge-pattern-dotted", cl) ||
                   grepl("stroke-dasharray", style, ignore.case = TRUE)) {
    "dashed"
  } else if (grepl("edge-thickness-thick\\b", cl)) {
    "thick"
  } else {
    "solid"
  }

  # Stroke colour from inline style (linkStyle rules are inlined by mermaid-cli)
  stroke_col <- NA_character_
  m_sc <- regmatches(style,
            regexpr("(?:^|;)\\s*stroke:\\s*([^;!]+)", style, perl = TRUE))
  if (length(m_sc) > 0L) {
    col <- .strip_important(sub(".*stroke:\\s*", "", m_sc))
    if (nzchar(col) && col != "none") stroke_col <- parse_css_colour(col)
  }

  # Look for matching edge label
  lbl_info <- find_edge_label(edge_id, doc)

  list(
    id          = edge_id,
    from        = from %||% NA_character_,
    to          = to   %||% NA_character_,
    label       = lbl_info$label,
    path_d      = d,
    arrow_start = arrow_start,
    arrow_end   = arrow_end,
    line_type   = line_type,
    stroke      = stroke_col,
    label_x     = lbl_info$x,
    label_y     = lbl_info$y
  )
}

classify_marker <- function(marker_ref) {
  m <- tolower(marker_ref)
  if (grepl("arrow|arrow", m))  return("arrow")
  if (grepl("circle|dot",  m))  return("circle")
  if (grepl("cross|x",     m))  return("cross")
  "arrow"   # default — most markers in mermaid are arrowheads
}

find_edge_label <- function(edge_id, doc) {
  empty <- list(label = NA_character_, x = NA_real_, y = NA_real_)
  if (!nzchar(edge_id)) return(empty)

  # Look for <g id="{edge_id}-label"> or <g class="edgeLabel"> near the path
  lbl_id <- paste0(edge_id, "-label")
  lbl_g  <- xml2::xml_find_first(doc, paste0(".//*[@id='", lbl_id, "']"))
  if (inherits(lbl_g, "xml_missing")) {
    # Try without "-label" suffix matching
    lbl_g <- xml2::xml_find_first(doc,
      paste0(".//*[contains(@class,'edgeLabel') and contains(@id,'",
             sub("L-", "", edge_id), "')]"))
  }
  if (inherits(lbl_g, "xml_missing")) return(empty)

  txt <- trimws(xml2::xml_text(lbl_g))
  if (!nzchar(txt)) return(empty)

  tf  <- xml2::xml_attr(lbl_g, "transform") %||% ""
  pos <- parse_translate(tf)

  list(label = txt, x = pos[1], y = pos[2])
}

empty_edges_tbl <- function() {
  tibble::tibble(
    id = character(), from = character(), to = character(),
    label = character(), path_d = character(),
    arrow_start = character(), arrow_end = character(),
    line_type = character(), stroke = character(),
    label_x = numeric(), label_y = numeric()
  )
}

# ── SVG path → DrawingML custGeom helpers ─────────────────────────────────

#' Convert an SVG path d-string to a DrawingML custGeom XML snippet
#'
#' The path coordinates are normalised to the shape's bounding box so the
#' resulting custGeom can be placed with an `<a:xfrm>` at any position.
#'
#' @param d SVG path `d` attribute string.
#' @param scale Numeric. Multiply every SVG pixel by this to get EMU.
#' @return A named list:
#'   - `xml`: character, the `<a:custGeom>...</a:custGeom>` string
#'   - `x_emu`, `y_emu`: top-left of bounding box in EMU (absolute)
#'   - `w_emu`, `h_emu`: width / height in EMU
#' @keywords internal
svg_path_to_custgeom <- function(d, scale) {
  cmds <- parse_svg_path(d)
  if (length(cmds) == 0L) return(NULL)

  # Absolute coordinates for bounding box
  pts  <- all_path_points(cmds)
  if (nrow(pts) == 0L) return(NULL)

  min_x <- min(pts$x); min_y <- min(pts$y)
  max_x <- max(pts$x); max_y <- max(pts$y)
  w_svg <- max(max_x - min_x, 1);  h_svg <- max(max_y - min_y, 1)

  x_emu <- as.integer(round(min_x * scale))
  y_emu <- as.integer(round(min_y * scale))
  w_emu <- as.integer(round(w_svg * scale))
  h_emu <- as.integer(round(h_svg * scale))

  # Build path XML — coordinates normalised to bounding box origin, in EMU
  path_xml <- cmds_to_drawingml(cmds, min_x, min_y, scale)

  xml <- paste0(
    "<a:custGeom>",
      "<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>",
      "<a:rect l=\"0\" t=\"0\" r=\"r\" b=\"b\"/>",
      "<a:pathLst>",
        "<a:path w=\"", w_emu, "\" h=\"", h_emu, "\">",
          path_xml,
        "</a:path>",
      "</a:pathLst>",
    "</a:custGeom>"
  )

  list(xml = xml, x_emu = x_emu, y_emu = y_emu, w_emu = w_emu, h_emu = h_emu)
}

# ── SVG path parser ────────────────────────────────────────────────────────

parse_svg_path <- function(d) {
  # Tokenise into (command, args...) tuples.
  # Commands: M m L l H h V v C c Q q S s T t A a Z z
  d     <- trimws(d)
  tokens <- gregexpr("[MmLlHhVvCcQqSsTtAaZz]|[-+]?(?:\\d*\\.\\d+|\\d+)(?:[eE][-+]?\\d+)?",
                      d, perl = TRUE)
  parts  <- regmatches(d, tokens)[[1]]
  if (length(parts) == 0L) return(list())

  cmds   <- list()
  cur_cmd <- NULL
  cur_args <- numeric(0)
  cx <- 0; cy <- 0   # current point

  flush_cmd <- function() {
    if (!is.null(cur_cmd)) {
      cmds[[length(cmds) + 1L]] <<- list(cmd = cur_cmd, args = cur_args, cx = cx, cy = cy)
    }
  }

  for (tok in parts) {
    if (grepl("^[A-Za-z]$", tok)) {
      flush_cmd()
      cur_cmd  <- tok
      cur_args <- numeric(0)
    } else {
      cur_args <- c(cur_args, as.numeric(tok))
    }
  }
  flush_cmd()

  # Now resolve relative to absolute commands
  resolve_path(cmds)
}

resolve_path <- function(cmds) {
  out <- list()
  cx <- 0; cy <- 0   # current point
  sx <- 0; sy <- 0   # start of current subpath (for Z)

  for (entry in cmds) {
    cmd  <- entry$cmd
    args <- entry$args

    if (cmd == "M") {
      pts <- matrix(args, ncol = 2, byrow = TRUE)
      for (i in seq_len(nrow(pts))) {
        cx <- pts[i, 1]; cy <- pts[i, 2]
        if (i == 1L) { sx <- cx; sy <- cy }
        out[[length(out) + 1L]] <- list(cmd = "M", x = cx, y = cy)
      }
    } else if (cmd == "m") {
      pts <- matrix(args, ncol = 2, byrow = TRUE)
      for (i in seq_len(nrow(pts))) {
        cx <- cx + pts[i, 1]; cy <- cy + pts[i, 2]
        if (i == 1L) { sx <- cx; sy <- cy }
        out[[length(out) + 1L]] <- list(cmd = "M", x = cx, y = cy)
      }
    } else if (cmd == "L") {
      pts <- matrix(args, ncol = 2, byrow = TRUE)
      for (i in seq_len(nrow(pts))) {
        cx <- pts[i, 1]; cy <- pts[i, 2]
        out[[length(out) + 1L]] <- list(cmd = "L", x = cx, y = cy)
      }
    } else if (cmd == "l") {
      pts <- matrix(args, ncol = 2, byrow = TRUE)
      for (i in seq_len(nrow(pts))) {
        cx <- cx + pts[i, 1]; cy <- cy + pts[i, 2]
        out[[length(out) + 1L]] <- list(cmd = "L", x = cx, y = cy)
      }
    } else if (cmd == "H") {
      for (x in args) { cx <- x; out[[length(out) + 1L]] <- list(cmd = "L", x = cx, y = cy) }
    } else if (cmd == "h") {
      for (dx in args) { cx <- cx + dx; out[[length(out) + 1L]] <- list(cmd = "L", x = cx, y = cy) }
    } else if (cmd == "V") {
      for (y in args) { cy <- y; out[[length(out) + 1L]] <- list(cmd = "L", x = cx, y = cy) }
    } else if (cmd == "v") {
      for (dy in args) { cy <- cy + dy; out[[length(out) + 1L]] <- list(cmd = "L", x = cx, y = cy) }
    } else if (cmd == "C") {
      pts <- matrix(args, ncol = 6, byrow = TRUE)
      for (i in seq_len(nrow(pts))) {
        out[[length(out) + 1L]] <- list(cmd = "C",
          x1 = pts[i,1], y1 = pts[i,2], x2 = pts[i,3], y2 = pts[i,4],
          x  = pts[i,5], y  = pts[i,6])
        cx <- pts[i,5]; cy <- pts[i,6]
      }
    } else if (cmd == "c") {
      pts <- matrix(args, ncol = 6, byrow = TRUE)
      for (i in seq_len(nrow(pts))) {
        out[[length(out) + 1L]] <- list(cmd = "C",
          x1 = cx + pts[i,1], y1 = cy + pts[i,2],
          x2 = cx + pts[i,3], y2 = cy + pts[i,4],
          x  = cx + pts[i,5], y  = cy + pts[i,6])
        cx <- cx + pts[i,5]; cy <- cy + pts[i,6]
      }
    } else if (cmd == "Q") {
      pts <- matrix(args, ncol = 4, byrow = TRUE)
      for (i in seq_len(nrow(pts))) {
        out[[length(out) + 1L]] <- list(cmd = "Q",
          x1 = pts[i,1], y1 = pts[i,2], x = pts[i,3], y = pts[i,4])
        cx <- pts[i,3]; cy <- pts[i,4]
      }
    } else if (cmd == "q") {
      pts <- matrix(args, ncol = 4, byrow = TRUE)
      for (i in seq_len(nrow(pts))) {
        out[[length(out) + 1L]] <- list(cmd = "Q",
          x1 = cx + pts[i,1], y1 = cy + pts[i,2],
          x  = cx + pts[i,3], y  = cy + pts[i,4])
        cx <- cx + pts[i,3]; cy <- cy + pts[i,4]
      }
    } else if (cmd %in% c("Z", "z")) {
      out[[length(out) + 1L]] <- list(cmd = "Z")
      cx <- sx; cy <- sy
    } else if (cmd == "A") {
      # A rx ry x-rot large-arc sweep x y  (absolute, groups of 7)
      if (length(args) >= 7L) {
        pts <- matrix(args, ncol = 7, byrow = TRUE)
        for (i in seq_len(nrow(pts))) {
          cx <- pts[i, 6]; cy <- pts[i, 7]
          out[[length(out) + 1L]] <- list(cmd = "L", x = cx, y = cy)
        }
      }
    } else if (cmd == "a") {
      # a rx ry x-rot large-arc sweep dx dy  (relative, groups of 7)
      if (length(args) >= 7L) {
        pts <- matrix(args, ncol = 7, byrow = TRUE)
        for (i in seq_len(nrow(pts))) {
          cx <- cx + pts[i, 6]; cy <- cy + pts[i, 7]
          out[[length(out) + 1L]] <- list(cmd = "L", x = cx, y = cy)
        }
      }
    }
    # S, s, T, t — approximate with lineTo; not commonly used in mermaid flowcharts
  }
  out
}

all_path_points <- function(cmds) {
  xs <- numeric(0); ys <- numeric(0)
  for (cmd in cmds) {
    if (cmd$cmd %in% c("M", "L")) { xs <- c(xs, cmd$x); ys <- c(ys, cmd$y) }
    else if (cmd$cmd == "C") {
      xs <- c(xs, cmd$x1, cmd$x2, cmd$x); ys <- c(ys, cmd$y1, cmd$y2, cmd$y)
    } else if (cmd$cmd == "Q") {
      xs <- c(xs, cmd$x1, cmd$x); ys <- c(ys, cmd$y1, cmd$y)
    }
  }
  data.frame(x = xs, y = ys)
}

cmds_to_drawingml <- function(cmds, origin_x, origin_y, scale) {
  pt <- function(x, y) {
    xe <- as.integer(round((x - origin_x) * scale))
    ye <- as.integer(round((y - origin_y) * scale))
    paste0("<a:pt x=\"", xe, "\" y=\"", ye, "\"/>")
  }

  parts <- character(length(cmds))
  for (i in seq_along(cmds)) {
    cmd <- cmds[[i]]
    parts[i] <- switch(cmd$cmd,
      "M" = paste0("<a:moveTo>",   pt(cmd$x, cmd$y),                             "</a:moveTo>"),
      "L" = paste0("<a:lnTo>",     pt(cmd$x, cmd$y),                             "</a:lnTo>"),
      "C" = paste0("<a:cubicBezTo>", pt(cmd$x1, cmd$y1), pt(cmd$x2, cmd$y2), pt(cmd$x, cmd$y), "</a:cubicBezTo>"),
      "Q" = paste0("<a:quadBezTo>",  pt(cmd$x1, cmd$y1), pt(cmd$x, cmd$y),   "</a:quadBezTo>"),
      "Z" = "<a:close/>",
      ""
    )
  }
  paste0(parts, collapse = "")
}

# ── path_bbox (fallback for path-shaped nodes) ─────────────────────────────

path_bbox <- function(d) {
  cmds <- parse_svg_path(d)
  pts  <- all_path_points(cmds)
  if (nrow(pts) == 0L) return(list(cx = 0, cy = 0, w = 60, h = 30))
  min_x <- min(pts$x); max_x <- max(pts$x)
  min_y <- min(pts$y); max_y <- max(pts$y)
  list(cx = (min_x + max_x) / 2, cy = (min_y + max_y) / 2,
       w  = max(max_x - min_x, 1), h = max(max_y - min_y, 1))
}

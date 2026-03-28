# R/drawingml.R
#
# Converts parsed SVG geometry (from parse_svg.R) into DrawingML XML
# (<w:p>…</w:p>) ready for officer::body_add_xml().
#
# Architecture: single <wp:anchor> containing a <wpg:wgp> super-group.
#   - Subgraph groups are nested <wpg:grpSp> elements (emitted first = behind)
#   - Diamond/hexagon/parallelogram nodes are <wpg:grpSp> groups containing
#     a visual shape wsp + transparent text-overlay wsp
#   - Regular nodes are plain <wps:wsp> children
#   - Edges are <wps:wsp> children with custom geometry
#
# Key OOXML rules for nested groups (MS-ODRAWXML §2.16.1):
#   - Top-level wpg:wgp (directly in a:graphicData): NO wpg:cNvPr
#   - Nested wpg:grpSp (inside any group):           MUST have wpg:cNvPr
#   - Element order: cNvPr → cNvGrpSpPr → grpSpPr → children → extLst
#   - Use wpg:grpSp for nested groups, never wpg:wgp
#   - All a:xfrm elements must have non-zero a:ext and a:chExt

.dml_namespaces <- paste0(
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ',
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
)

.shape_map <- c(
  # Basic shapes
  rect          = "rect",
  roundRect     = "roundRect",
  diamond       = "diamond",
  ellipse       = "ellipse",
  hexagon       = "hexagon",
  parallelogram = "parallelogram",
  cylinder      = "can",
  # Document shapes (native DrawingML presets)
  doc           = "flowChartDocument",
  docs          = "flowChartMultidocument",
  # Flowchart shapes
  winPane       = "flowChartInternalStorage",
  delay         = "flowChartDelay",
  manualInput   = "flowChartManualInput",
  manualOp      = "flowChartManualOperation",
  trapezoid     = "trapezoid",
  collate       = "flowChartCollate",
  display       = "flowChartDisplay",
  subprocess    = "flowChartPredefinedProcess",
  punchedTape   = "flowChartPunchedTape",
  storedData    = "flowChartOnlineStorage",
  crossCirc     = "flowChartSummingJunction",
  hCyl          = "flowChartMagneticDisk",
  linCyl        = "flowChartMagneticDrum",
  extract       = "flowChartExtract",
  mergeTri      = "flowChartMerge",
  notchPent     = "flowChartPreparation",
  # Other presets
  bolt          = "lightningBolt",
  notchRect     = "snip1Rect",
  bang          = "irregularSeal1",
  cloud         = "cloud",
  brace         = "leftBrace",
  braceR        = "rightBrace",
  braces        = "bracePair",
  leanR         = "parallelogram",
  leanL         = "parallelogram",    # uses flipH="1" in xfrm
  # Special fill shapes (forced black; ellipse/rect geometry)
  junction      = "ellipse",          # f-circ — always solid black
  fork          = "rect",             # fork — always solid black
  # Double-outline / line-decorated shapes (approximations — TODO: custom geometry)
  dblCirc       = "ellipse",          # dbl-circ — TODO: double ellipse
  frCirc        = "ellipse",          # fr-circ — TODO: frame circle
  linDoc        = "flowChartDocument",# lin-doc — approximation
  linRect       = "rect",             # lin-rect — approximation
  odd           = "rect",             # odd — SVG custom TODO
  tagDoc        = "flowChartDocument",# tag-doc — approximation
  tagRect       = "rect",             # tag-rect — approximation
  stRect        = "rect",             # st-rect — used only by st_rect_page_wsp
  text          = "rect"              # transparent annotation label
)

# ── ID counter ─────────────────────────────────────────────────────────────

# Build the <w:rPr> XML for a labelled text run.
# Using a single helper ensures font, size, and colour are consistent
# everywhere and makes it easy to change them in one place.
text_rpr_xml <- function(text_col, ctx) {
  paste0(
    "<w:rPr>",
      "<w:rFonts",
        " w:ascii=\"",    ctx$font_family, "\"",
        " w:hAnsi=\"",    ctx$font_family, "\"",
        " w:cs=\"",       ctx$font_family, "\"",
      "/>",
      "<w:color w:val=\"", text_col, "\"/>",
      "<w:sz w:val=\"",    ctx$font_size_hp, "\"/>",
      "<w:szCs w:val=\"",  ctx$font_size_hp, "\"/>",
    "</w:rPr>"
  )
}

new_id_counter <- function(start = 1L) {
  n <- as.integer(start)
  list(
    next_id = function() { v <- n; n <<- n + 1L; v },
    peek    = function() n
  )
}

# ── Main exported function ─────────────────────────────────────────────────

#' Build DrawingML XML from parsed mermaid SVG geometry
#'
#' @param svg_data List from [parse_mermaid_svg()]: nodes, edges, viewbox.
#' @param ast Reserved for future use.
#' @param start_id Integer. First shape ID; increment between diagrams.
#' @param page_width_in Numeric. Usable content width in inches (default 6.0).
#' @param page_height_in Numeric. Max diagram height in inches (default 8.0).
#' @param margin_in Numeric. Left/top margin offset in inches (default 1.0).
#' @param default_fill Default node fill colour (6-char hex). Default "FFFFFF".
#' @param default_stroke Default stroke colour. Default "5E504E".
#' @param default_text_color Default text colour. Default "000000".
#' @param stroke_width_pt Stroke width in points. Default 1.5.
#' @return Named list: xml (character) and next_id (integer).
#' @export
build_diagram_xml <- function(svg_data,
                               ast                = NULL,
                               start_id           = 1L,
                               page_width_in      = 6.0,
                               page_height_in     = 8.0,
                               margin_in          = 1.0,
                               default_fill       = "FFFFFF",
                               default_stroke     = "5E504E",
                               default_text_color = "000000",
                               stroke_width_pt    = 1.5) {

  nodes  <- svg_data$nodes
  edges  <- svg_data$edges
  vb     <- svg_data$viewbox   # c(minx, miny, svg_w, svg_h)
  sty    <- svg_data$style %||%
    list(font_size_px = 14, edge_stroke = default_stroke, edge_sw_px = 2)

  svg_w  <- vb[3]; svg_h <- vb[4]
  scale  <- compute_scale(svg_w, svg_h, page_width_in, page_height_in)
  off_x  <- inches_to_emu(margin_in)
  off_y  <- inches_to_emu(margin_in)

  # Scale font size proportionally; floor at 8 half-pts (4pt).
  font_size_hp <- max(as.integer(round(sty$font_size_px * scale / 6350)), 8L)

  # Font family: use whatever mermaid/Chromium measured the text in, so that
  # Word's text renderer uses the same metrics and text fits the same boxes.
  font_family <- sty$font_family %||% "Trebuchet MS"

  # Edge stroke: colour from SVG stylesheet; width scaled from SVG px, min 0.75pt.
  # Word renders sub-0.5pt lines as hairlines; 0.75pt keeps edges clearly visible.
  edge_stroke <- sty$edge_stroke %||% default_stroke
  edge_sw_emu <- max(as.integer(round(sty$edge_sw_px * scale)), 9525L)

  subgraphs <- svg_data$subgraphs %||% empty_subgraphs_tbl()

  ctx <- list(
    vb             = vb,
    scale          = scale,
    font_size_hp   = font_size_hp,
    font_family    = font_family,
    edge_stroke    = edge_stroke,
    edge_sw_emu    = edge_sw_emu,
    default_fill   = default_fill,
    default_stroke = default_stroke,
    dtc            = default_text_color,
    sw_pt          = stroke_width_pt
  )

  ctr <- new_id_counter(as.integer(start_id) + 1L)  # +1: super-anchor uses start_id

  sg_w_emu <- max(as.integer(round(svg_w * scale)), 91440L)
  sg_h_emu <- max(as.integer(round(svg_h * scale)), 91440L)

  # Build subgraph containment tree
  sg_tree     <- build_subgraph_tree(subgraphs)
  node_sg_vec <- assign_nodes_to_subgraphs(nodes, sg_tree)
  if (nrow(nodes) > 0L) names(node_sg_vec) <- nodes$id

  parts <- character(0)

  # Root subgraphs — emitted first so they are drawn behind nodes/edges
  root_sg_ids <- names(sg_tree)[vapply(sg_tree, function(sg)
    is.na(sg$parent_id %||% NA_character_), logical(1))]

  for (sg_id in root_sg_ids) {
    xml <- emit_subgraph_grpSp(sg_id, sg_tree, vb[1], vb[2],
                                nodes, edges, node_sg_vec, ctx, ctr)
    if (nzchar(xml)) parts <- c(parts, xml)
  }

  # Nodes not inside any subgraph
  if (nrow(nodes) > 0L) {
    top_level_idx <- which(is.na(node_sg_vec))
    for (i in top_level_idx) {
      nd  <- as.list(nodes[i, ])
      xml <- emit_node(nd, vb[1], vb[2], ctx, ctr)
      if (nzchar(xml)) parts <- c(parts, xml)
    }
  }

  # Top-level edges: those not fully contained within a single root subgraph
  if (nrow(edges) > 0L) {
    for (i in seq_len(nrow(edges))) {
      e       <- as.list(edges[i, ])
      from_id <- e$from %||% NA_character_
      to_id   <- e$to   %||% NA_character_
      from_sg <- if (!is.na(from_id) && from_id %in% names(node_sg_vec))
                   node_sg_vec[[from_id]] else NA_character_
      to_sg   <- if (!is.na(to_id) && to_id %in% names(node_sg_vec))
                   node_sg_vec[[to_id]]   else NA_character_

      is_top_level <- is.na(from_sg) || is.na(to_sg) || !same_root(from_sg, to_sg, sg_tree)
      if (!is_top_level) next

      xml <- emit_edge(e, vb[1], vb[2], ctx, ctr)
      if (nzchar(xml)) parts <- c(parts, xml)

      lbl <- e$label %||% ""
      if (nzchar(lbl) &&
          !is.na(e$label_x %||% NA_real_) &&
          !is.na(e$label_y %||% NA_real_)) {
        lxml <- emit_edge_label(e, vb[1], vb[2], ctx, ctr)
        if (nzchar(lxml)) parts <- c(parts, lxml)
      }
    }
  }

  children_xml <- paste0(parts, collapse = "")
  anchor_xml   <- make_super_anchor(off_x, off_y, sg_w_emu, sg_h_emu,
                                     as.integer(start_id), children_xml)

  xml <- paste0('<w:p ', .dml_namespaces, '>', anchor_xml, '</w:p>')
  list(xml = xml, next_id = ctr$peek())
}

# ── Scale ─────────────────────────────────────────────────────────────────

compute_scale <- function(svg_w, svg_h, page_w_in, page_h_in) {
  if (is.na(svg_w) || svg_w <= 0 || is.na(svg_h) || svg_h <= 0)
    return(9525)  # 1 px = 9525 EMU at 96 dpi
  target_w <- inches_to_emu(page_w_in)
  target_h <- inches_to_emu(page_h_in)
  min(target_w / svg_w, target_h / svg_h)
}

# ── Group XML primitives ────────────────────────────────────────────────────
#
# make_root_wgp:    top-level wpg:wgp (NO cNvPr — spec requirement)
# make_nested_grpSp: nested wpg:grpSp (cNvPr REQUIRED — spec requirement)

make_root_wgp <- function(w_emu, h_emu, children_xml) {
  w_emu <- max(as.integer(w_emu), 91440L)
  h_emu <- max(as.integer(h_emu), 91440L)
  paste0(
    '<wpg:wgp>',
      '<wpg:cNvGrpSpPr/>',
      '<wpg:grpSpPr>',
        '<a:xfrm>',
          '<a:off x="0" y="0"/>',
          '<a:ext cx="', w_emu, '" cy="', h_emu, '"/>',
          '<a:chOff x="0" y="0"/>',
          '<a:chExt cx="', w_emu, '" cy="', h_emu, '"/>',
        '</a:xfrm>',
      '</wpg:grpSpPr>',
      children_xml,
    '</wpg:wgp>'
  )
}

make_nested_grpSp <- function(grp_id, name, x_rel, y_rel, w_emu, h_emu,
                               children_xml) {
  w_emu <- max(as.integer(w_emu), 91440L)
  h_emu <- max(as.integer(h_emu), 91440L)
  paste0(
    '<wpg:grpSp>',
      '<wpg:cNvPr id="', grp_id, '" name="', xml_escape(name), '"/>',
      '<wpg:cNvGrpSpPr/>',
      '<wpg:grpSpPr>',
        '<a:xfrm>',
          '<a:off x="', as.integer(x_rel), '" y="', as.integer(y_rel), '"/>',
          '<a:ext cx="', w_emu, '" cy="', h_emu, '"/>',
          '<a:chOff x="0" y="0"/>',
          '<a:chExt cx="', w_emu, '" cy="', h_emu, '"/>',
        '</a:xfrm>',
      '</wpg:grpSpPr>',
      children_xml,
    '</wpg:grpSp>'
  )
}

# ── Super-group anchor ──────────────────────────────────────────────────────

make_super_anchor <- function(x_page, y_page, w_emu, h_emu, shape_id,
                               children_xml) {
  w_emu <- max(as.integer(w_emu), 91440L)
  h_emu <- max(as.integer(h_emu), 91440L)
  # mc:AlternateContent with mc:Choice Requires="wpg" is required for Word to
  # activate its wpg namespace handler. Without this wrapper, nested wpg:grpSp
  # groups do not render even when the XML is otherwise spec-conformant.
  paste0(
    '<w:r><w:rPr><w:noProof/></w:rPr>',
    '<mc:AlternateContent>',
    '<mc:Choice Requires="wpg">',
    '<w:drawing>',
    '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" ',
      'simplePos="0" relativeHeight="251658240" behindDoc="0" ',
      'locked="0" layoutInCell="1" allowOverlap="1">',
    '<wp:simplePos x="0" y="0"/>',
    '<wp:positionH relativeFrom="page"><wp:posOffset>',
      as.integer(x_page), '</wp:posOffset></wp:positionH>',
    '<wp:positionV relativeFrom="page"><wp:posOffset>',
      as.integer(y_page), '</wp:posOffset></wp:positionV>',
    '<wp:extent cx="', w_emu, '" cy="', h_emu, '"/>',
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
    '<wp:wrapNone/>',
    '<wp:docPr id="', shape_id, '" name="mermaid:diagram" descr="{}"/>',
    '<wp:cNvGraphicFramePr/>',
    '<a:graphic>',
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">',
    make_root_wgp(w_emu, h_emu, children_xml),
    '</a:graphicData></a:graphic>',
    '</wp:anchor></w:drawing>',
    '</mc:Choice>',
    '<mc:Fallback/>',
    '</mc:AlternateContent>',
    '</w:r>'
  )
}

# ── Subgraph tree helpers ───────────────────────────────────────────────────

build_subgraph_tree <- function(subgraphs) {
  if (nrow(subgraphs) == 0L) return(list())
  n       <- nrow(subgraphs)
  parents <- rep(NA_character_, n)
  for (i in seq_len(n)) {
    sg_i      <- subgraphs[i, ]
    best_area <- Inf
    for (j in seq_len(n)) {
      if (i == j) next
      sg_j     <- subgraphs[j, ]
      contains <- sg_j$svg_x <= sg_i$svg_x &&
                  sg_j$svg_y <= sg_i$svg_y &&
                  (sg_j$svg_x + sg_j$svg_w) >= (sg_i$svg_x + sg_i$svg_w) &&
                  (sg_j$svg_y + sg_j$svg_h) >= (sg_i$svg_y + sg_i$svg_h)
      if (contains) {
        area_j <- sg_j$svg_w * sg_j$svg_h
        if (area_j < best_area) { best_area <- area_j; parents[i] <- sg_j$id }
      }
    }
  }
  result <- list()
  for (i in seq_len(n)) {
    sg           <- as.list(subgraphs[i, ])
    sg$parent_id <- parents[i]
    sg$children  <- subgraphs$id[!is.na(parents) & parents == sg$id]
    result[[sg$id]] <- sg
  }
  result
}

assign_nodes_to_subgraphs <- function(nodes, sg_tree) {
  if (length(sg_tree) == 0L || nrow(nodes) == 0L)
    return(rep(NA_character_, nrow(nodes)))
  vapply(seq_len(nrow(nodes)), function(i) {
    nd        <- nodes[i, ]
    best_area <- Inf
    best_id   <- NA_character_
    for (sg_id in names(sg_tree)) {
      sg <- sg_tree[[sg_id]]
      if (nd$svg_cx >= sg$svg_x && nd$svg_cx <= sg$svg_x + sg$svg_w &&
          nd$svg_cy >= sg$svg_y && nd$svg_cy <= sg$svg_y + sg$svg_h) {
        area <- sg$svg_w * sg$svg_h
        if (area < best_area) { best_area <- area; best_id <- sg_id }
      }
    }
    best_id
  }, character(1))
}

is_in_subtree <- function(sg_id, root_id, tree) {
  if (is.na(sg_id) || is.na(root_id)) return(FALSE)
  current <- sg_id
  while (!is.na(current)) {
    if (current == root_id) return(TRUE)
    parent <- tree[[current]]$parent_id %||% NA_character_
    if (is.na(parent)) break
    current <- parent
  }
  FALSE
}

root_of <- function(id, tree) {
  while (!is.na(tree[[id]]$parent_id %||% NA_character_))
    id <- tree[[id]]$parent_id
  id
}

same_root <- function(sg_a, sg_b, tree) {
  if (is.na(sg_a) || is.na(sg_b)) return(FALSE)
  root_of(sg_a, tree) == root_of(sg_b, tree)
}

# ── Colour helpers ─────────────────────────────────────────────────────────

is_dark <- function(hex) {
  if (is.na(hex) || nchar(hex) < 6L) return(FALSE)
  r   <- strtoi(substr(hex, 1L, 2L), 16L)
  g   <- strtoi(substr(hex, 3L, 4L), 16L)
  b   <- strtoi(substr(hex, 5L, 6L), 16L)
  (0.299 * r + 0.587 * g + 0.114 * b) < 128
}

node_text_colour <- function(nd, is_transparent, fill_hex, default_tc) {
  col <- nd$color %||% NA_character_
  if (!is.na(col)) return(col)
  if (!is_transparent && !is.na(fill_hex) && is_dark(fill_hex)) return("FFFFFF")
  default_tc
}

# ── Non-visual properties ───────────────────────────────────────────────────
# Element order for wps:wsp: cNvPr → (cNvSpPr|cNvCnPr) → spPr → ...
# cNvPr is optional per schema but MUST be present inside nested groups.

make_nvSpPr <- function(shape_id, name, descr_json) {
  paste0(
    '<wps:cNvPr id="', shape_id, '" name="', xml_escape(name), '"/>',
    '<wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>'
  )
}

make_nvCnPr <- function(shape_id, name, descr_json) {
  paste0(
    '<wps:cNvPr id="', shape_id, '" name="', xml_escape(name), '"/>',
    '<wps:cNvCnPr/>'
  )
}

# ── Subgraph group emission ─────────────────────────────────────────────────

emit_subgraph_grpSp <- function(sg_id, sg_tree, origin_x_px, origin_y_px,
                                 nodes, edges, node_sg_vec, ctx, ctr) {
  sg    <- sg_tree[[sg_id]]
  scale <- ctx$scale
  x_rel <- as.integer(round((sg$svg_x - origin_x_px) * scale))
  y_rel <- as.integer(round((sg$svg_y - origin_y_px) * scale))
  w_emu <- max(as.integer(round(sg$svg_w * scale)), 91440L)
  h_emu <- max(as.integer(round(sg$svg_h * scale)), 91440L)

  grp_id  <- ctr$next_id()
  bg_xml  <- subgraph_rect_wsp(sg, w_emu, h_emu, ctx, ctr)

  # Child subgraph groups (positions relative to this subgraph origin)
  child_sg_xml <- vapply(sg$children, function(child_id)
    emit_subgraph_grpSp(child_id, sg_tree, sg$svg_x, sg$svg_y,
                        nodes, edges, node_sg_vec, ctx, ctr),
    character(1))

  # Nodes directly in this subgraph (not in any child subgraph)
  direct_idx <- which(node_sg_vec == sg_id)
  node_xml <- vapply(direct_idx, function(i)
    emit_node(as.list(nodes[i, ]), sg$svg_x, sg$svg_y, ctx, ctr),
    character(1))

  # Edges whose both endpoints are in this subgraph subtree but not both
  # within the same single child subgraph
  edge_xml  <- character(nrow(edges))
  label_xml <- character(nrow(edges))
  for (i in seq_len(nrow(edges))) {
    e       <- as.list(edges[i, ])
    from_id <- e$from %||% NA_character_
    to_id   <- e$to   %||% NA_character_
    from_sg <- if (!is.na(from_id) && from_id %in% names(node_sg_vec))
                 node_sg_vec[[from_id]] else NA_character_
    to_sg   <- if (!is.na(to_id) && to_id %in% names(node_sg_vec))
                 node_sg_vec[[to_id]]   else NA_character_

    from_in <- !is.na(from_sg) && is_in_subtree(from_sg, sg_id, sg_tree)
    to_in   <- !is.na(to_sg)   && is_in_subtree(to_sg,   sg_id, sg_tree)
    if (!from_in || !to_in) next

    skip <- FALSE
    for (child_id in sg$children) {
      if (is_in_subtree(from_sg, child_id, sg_tree) &&
          is_in_subtree(to_sg,   child_id, sg_tree)) {
        skip <- TRUE; break
      }
    }
    if (skip) next

    edge_xml[i] <- emit_edge(e, sg$svg_x, sg$svg_y, ctx, ctr)
    lbl <- e$label %||% ""
    if (nzchar(lbl) && !is.na(e$label_x %||% NA_real_) && !is.na(e$label_y %||% NA_real_))
      label_xml[i] <- emit_edge_label(e, sg$svg_x, sg$svg_y, ctx, ctr)
  }

  all_pieces <- c(bg_xml, child_sg_xml, node_xml, edge_xml, label_xml)
  all_xml    <- paste0(all_pieces[nzchar(all_pieces)], collapse = "")

  make_nested_grpSp(grp_id,
                     paste0("mermaid:subgraph:", sg$id %||% ""),
                     x_rel, y_rel, w_emu, h_emu, all_xml)
}

# Background rect for a subgraph (position 0,0 within the subgraph group)
subgraph_rect_wsp <- function(sg, w_emu, h_emu, ctx, ctr) {
  shape_id   <- ctr$next_id()
  fill_hex   <- sg$fill %||% NA_character_
  fill_xml   <- if (is.na(fill_hex)) {
    "<a:noFill/>"
  } else {
    paste0("<a:solidFill><a:srgbClr val=\"", fill_hex, "\"/></a:solidFill>")
  }

  stroke_hex <- sg$stroke %||% "AAAA33"
  # stroke-width:Npx from SVG is in SVG-user-unit pixels; scale to EMU the same
  # way as every other dimension. Without scaling, 4px would become 4pt (≈50800 EMU)
  # which is far too thick. With scaling it becomes ~1pt.
  sg_sw_px <- sg$stroke_width %||% NA_real_
  if (!is.na(sg_sw_px) && sg_sw_px > 0) {
    sw_emu <- max(as.integer(round(sg_sw_px * ctx$scale)), 1588L)
  } else {
    sw_emu <- max(as.integer(ctx$sw_pt * 12700 * ctx$scale / 9525), 1588L)
  }
  stroke_xml <- paste0(
    "<a:ln w=\"", sw_emu, "\">",
    "<a:solidFill><a:srgbClr val=\"", stroke_hex, "\"/></a:solidFill>",
    "</a:ln>"
  )

  label    <- xml_escape(strip_html(sg$label %||% ""))
  # Use the subgraph's own colour for its label (from SVG `color:` property or stroke)
  text_col <- sg$color %||% NA_character_
  if (is.na(text_col)) text_col <- if (!is.na(fill_hex) && is_dark(fill_hex)) "FFFFFF" else ctx$dtc

  descr <- jsonlite::toJSON(list(
    v = "1", type = "subgraph",
    id = sg$id %||% "", label = sg$label %||% "",
    class = sg$class %||% NULL
  ), auto_unbox = TRUE, null = "null")

  nv_xml <- make_nvSpPr(shape_id,
                         paste0("mermaid:subgraph:", sg$id %||% ""),
                         descr)

  paste0(
    "<wps:wsp>",
      nv_xml,
      "<wps:spPr>",
        "<a:xfrm>",
          "<a:off x=\"0\" y=\"0\"/>",
          "<a:ext cx=\"", w_emu, "\" cy=\"", h_emu, "\"/>",
        "</a:xfrm>",
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
        fill_xml, stroke_xml,
      "</wps:spPr>",
      "<wps:txbx><w:txbxContent><w:p>",
        "<w:pPr><w:jc w:val=\"center\"/></w:pPr>",
        "<w:r>", text_rpr_xml(text_col, ctx),
          "<w:t xml:space=\"preserve\">", label, "</w:t></w:r>",
      "</w:p></w:txbxContent></wps:txbx>",
      "<wps:bodyPr anchor=\"t\" ",
        "lIns=\"0\" rIns=\"0\" ",
        "tIns=\"0\" bIns=\"0\">",
        "<a:normAutofit/></wps:bodyPr>",
    "</wps:wsp>"
  )
}

# ── Node emission ──────────────────────────────────────────────────────────

emit_node <- function(nd, origin_x_px, origin_y_px, ctx, ctr) {
  scale      <- ctx$scale
  w_emu      <- max(as.integer(round(nd$svg_w * scale)), 91440L)
  h_emu      <- max(as.integer(round(nd$svg_h * scale)), 45720L)
  x_rel      <- as.integer(round((nd$svg_cx - origin_x_px) * scale)) - w_emu %/% 2L
  y_rel      <- as.integer(round((nd$svg_cy - origin_y_px) * scale)) - h_emu %/% 2L
  shape_name <- nd$shape %||% "rect"

  # stRect: custom stacked-rect implementation
  if (shape_name == "stRect") {
    return(emit_st_rect_grpSp(nd, x_rel, y_rel, w_emu, h_emu, ctx, ctr))
  }

  # Shapes that need the visual-wsp + text-overlay-wsp group approach:
  # Non-rectangular geometries where Word's built-in text placement clips or
  # misaligns the label (typically the same shapes that use foreignObject in SVG).
  needs_group <- shape_name %in% c(
    "diamond", "hexagon",
    "parallelogram", "leanR", "leanL",
    "trapezoid",
    "extract",    # triangle (flowChartExtract)
    "mergeTri",   # inverted triangle (flowChartMerge)
    "notchPent",  # pentagon (flowChartPreparation)
    "collate"     # bowtie (flowChartCollate)
  )

  if (needs_group) {
    emit_node_grpSp(nd, x_rel, y_rel, w_emu, h_emu, ctx, ctr)
  } else {
    # All other shapes (rect, roundRect, ellipse, cylinder, flowChart*,
    # cloud, bolt, bang, brace*, text, etc.) use a single wps:wsp with
    # Word's built-in text body — no extra nesting required.
    node_wsp(nd, x_rel, y_rel, w_emu, h_emu, ctx, ctr)
  }
}

# wpg:grpSp for docs (stacked document pages):
# Two foldedCorner shapes offset from each other; front shape carries the text label.
emit_docs_grpSp <- function(nd, x_rel, y_rel, w_emu, h_emu, ctx, ctr) {
  grp_id <- ctr$next_id()

  # Stacking offset: ~18% of the smaller dimension, minimum 45720 EMU (0.05 in)
  off    <- max(as.integer(round(min(w_emu, h_emu) * 0.18)), 45720L)
  pg_w   <- w_emu - off
  pg_h   <- h_emu - off

  # Back page: offset to upper-right, same style, no text
  back_xml  <- docs_page_wsp(nd, off, 0L,  pg_w, pg_h, with_text = FALSE, ctx, ctr)
  # Front page: lower-left, with text label
  front_xml <- docs_page_wsp(nd, 0L, off,  pg_w, pg_h, with_text = TRUE,  ctx, ctr)

  make_nested_grpSp(grp_id,
                    paste0("mermaid:node:", nd$id %||% ""),
                    x_rel, y_rel, w_emu, h_emu,
                    paste0(back_xml, front_xml))
}

docs_page_wsp <- function(nd, x_pg, y_pg, pg_w, pg_h, with_text, ctx, ctr) {
  shape_id   <- ctr$next_id()
  fill_hex   <- nd$fill   %||% NA_character_
  stroke_hex <- nd$stroke %||% NA_character_

  # Both pages are always solid (never transparent) so that the front document
  # visually blocks the overlapping portion of the back document:
  #   back page  — always white (paper/substrate colour)
  #   front page — node fill colour, or white if the node has no fill
  fill_xml <- if (!with_text) {
    "<a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill>"
  } else {
    if (is.na(fill_hex))
      "<a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill>"
    else
      paste0("<a:solidFill><a:srgbClr val=\"", fill_hex, "\"/></a:solidFill>")
  }

  sw_emu     <- max(as.integer(ctx$sw_pt * 12700 * ctx$scale / 9525), 1588L)
  stroke_val <- if (!is.na(stroke_hex)) stroke_hex else ctx$default_stroke
  stroke_xml <- paste0(
    "<a:ln w=\"", sw_emu, "\">",
    "<a:solidFill><a:srgbClr val=\"", stroke_val, "\"/></a:solidFill>",
    "</a:ln>"
  )

  nv_xml <- make_nvSpPr(shape_id,
                        paste0("mermaid:", if (with_text) "node" else "back", ":",
                               nd$id %||% ""),
                        "{}")
  xfrm_xml <- paste0(
    "<a:xfrm>",
      "<a:off x=\"", x_pg, "\" y=\"", y_pg, "\"/>",
      "<a:ext cx=\"", pg_w, "\" cy=\"", pg_h, "\"/>",
    "</a:xfrm>"
  )
  geom_xml <- "<a:prstGeom prst=\"foldedCorner\"><a:avLst/></a:prstGeom>"

  if (!with_text) {
    paste0(
      "<wps:wsp>", nv_xml,
        "<wps:spPr>", xfrm_xml, geom_xml, fill_xml, stroke_xml, "</wps:spPr>",
        "<wps:bodyPr><a:noAutofit/></wps:bodyPr>",
      "</wps:wsp>"
    )
  } else {
    is_transparent <- is.na(fill_hex)
    label    <- xml_escape(strip_html(nd$label %||% nd$id %||% ""))
    text_col <- node_text_colour(nd, is_transparent, fill_hex, ctx$dtc)
        paste0(
      "<wps:wsp>", nv_xml,
        "<wps:spPr>", xfrm_xml, geom_xml, fill_xml, stroke_xml, "</wps:spPr>",
        "<wps:txbx><w:txbxContent><w:p>",
          "<w:pPr><w:jc w:val=\"center\"/></w:pPr>",
          "<w:r>", text_rpr_xml(text_col, ctx),
            "<w:t xml:space=\"preserve\">", label, "</w:t></w:r>",
        "</w:p></w:txbxContent></wps:txbx>",
        "<wps:bodyPr anchor=\"ctr\" ",
          "lIns=\"0\" rIns=\"0\" ",
          "tIns=\"0\" bIns=\"0\">",
          "<a:normAutofit/></wps:bodyPr>",
      "</wps:wsp>"
    )
  }
}

# wpg:grpSp for st-rect (Multi-Process): two stacked rect shapes in a group.
# Back page is white; front page carries the fill colour and text label.
# Modelled on emit_docs_grpSp / docs_page_wsp but uses rect preset geometry.
emit_st_rect_grpSp <- function(nd, x_rel, y_rel, w_emu, h_emu, ctx, ctr) {
  grp_id <- ctr$next_id()

  # Stacking offset: ~18% of the smaller dimension, minimum 45720 EMU (0.05 in)
  off    <- max(as.integer(round(min(w_emu, h_emu) * 0.18)), 45720L)
  pg_w   <- w_emu - off
  pg_h   <- h_emu - off

  # Back page: offset to upper-right, no text
  back_xml  <- st_rect_page_wsp(nd, off, 0L,  pg_w, pg_h, with_text = FALSE, ctx, ctr)
  # Front page: lower-left, with text label
  front_xml <- st_rect_page_wsp(nd, 0L, off,  pg_w, pg_h, with_text = TRUE,  ctx, ctr)

  make_nested_grpSp(grp_id,
                    paste0("mermaid:node:", nd$id %||% ""),
                    x_rel, y_rel, w_emu, h_emu,
                    paste0(back_xml, front_xml))
}

st_rect_page_wsp <- function(nd, x_pg, y_pg, pg_w, pg_h, with_text, ctx, ctr) {
  shape_id   <- ctr$next_id()
  fill_hex   <- nd$fill   %||% NA_character_
  stroke_hex <- nd$stroke %||% NA_character_

  # Back page is always white so it visually blocks the overlapping area.
  # Front page uses the node fill colour (or white if no fill).
  fill_xml <- if (!with_text) {
    "<a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill>"
  } else {
    if (is.na(fill_hex))
      "<a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill>"
    else
      paste0("<a:solidFill><a:srgbClr val=\"", fill_hex, "\"/></a:solidFill>")
  }

  sw_emu     <- max(as.integer(ctx$sw_pt * 12700 * ctx$scale / 9525), 1588L)
  stroke_val <- if (!is.na(stroke_hex)) stroke_hex else ctx$default_stroke
  stroke_xml <- paste0(
    "<a:ln w=\"", sw_emu, "\">",
    "<a:solidFill><a:srgbClr val=\"", stroke_val, "\"/></a:solidFill>",
    "</a:ln>"
  )

  nv_xml <- make_nvSpPr(shape_id,
                        paste0("mermaid:", if (with_text) "node" else "back", ":",
                               nd$id %||% ""),
                        "{}")
  xfrm_xml <- paste0(
    "<a:xfrm>",
      "<a:off x=\"", x_pg, "\" y=\"", y_pg, "\"/>",
      "<a:ext cx=\"", pg_w, "\" cy=\"", pg_h, "\"/>",
    "</a:xfrm>"
  )
  geom_xml <- "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"

  if (!with_text) {
    paste0(
      "<wps:wsp>", nv_xml,
        "<wps:spPr>", xfrm_xml, geom_xml, fill_xml, stroke_xml, "</wps:spPr>",
        "<wps:bodyPr><a:noAutofit/></wps:bodyPr>",
      "</wps:wsp>"
    )
  } else {
    is_transparent <- is.na(fill_hex)
    label    <- xml_escape(strip_html(nd$label %||% nd$id %||% ""))
    text_col <- node_text_colour(nd, is_transparent, fill_hex, ctx$dtc)
        paste0(
      "<wps:wsp>", nv_xml,
        "<wps:spPr>", xfrm_xml, geom_xml, fill_xml, stroke_xml, "</wps:spPr>",
        "<wps:txbx><w:txbxContent><w:p>",
          "<w:pPr><w:jc w:val=\"center\"/></w:pPr>",
          "<w:r>", text_rpr_xml(text_col, ctx),
            "<w:t xml:space=\"preserve\">", label, "</w:t></w:r>",
        "</w:p></w:txbxContent></wps:txbx>",
        "<wps:bodyPr anchor=\"ctr\" ",
          "lIns=\"0\" rIns=\"0\" ",
          "tIns=\"0\" bIns=\"0\">",
          "<a:normAutofit/></wps:bodyPr>",
      "</wps:wsp>"
    )
  }
}

# wpg:grpSp for diamond/hexagon/parallelogram:
# visual shape wsp at (0,0) + transparent text-overlay wsp centred in group.
emit_node_grpSp <- function(nd, x_rel, y_rel, w_emu, h_emu, ctx, ctr) {
  grp_id     <- ctr$next_id()
  visual_xml <- node_visual_wsp(nd, w_emu, h_emu, ctx, ctr)

  raw_tw <- nd$txt_w %||% NA_real_
  raw_th <- nd$txt_h %||% NA_real_
  tw_emu <- if (!is.na(raw_tw) && raw_tw > 0)
              as.integer(round(raw_tw * ctx$scale)) else w_emu
  th_emu <- if (!is.na(raw_th) && raw_th > 0)
              as.integer(round(raw_th * ctx$scale)) else h_emu
  tx_rel <- (w_emu - tw_emu) %/% 2L
  ty_rel <- (h_emu - th_emu) %/% 2L

  text_xml <- node_text_overlay_wsp(nd, tx_rel, ty_rel, tw_emu, th_emu, ctx, ctr)

  make_nested_grpSp(grp_id,
                     paste0("mermaid:node:", nd$id %||% ""),
                     x_rel, y_rel, w_emu, h_emu,
                     paste0(visual_xml, text_xml))
}

# Regular node wsp (shape + text, position within parent group)
node_wsp <- function(nd, x_rel, y_rel, w_emu, h_emu, ctx, ctr) {
  shape_id   <- ctr$next_id()
  scale      <- ctx$scale
  fill_hex   <- nd$fill   %||% NA_character_
  stroke_hex <- nd$stroke %||% NA_character_

  shape_name     <- nd$shape %||% "rect"
  is_transparent <- is.na(fill_hex)

  # junction (f-circ) and fork are always solid black regardless of SVG fill
  fill_xml <- if (shape_name %in% c("junction", "fork")) {
    "<a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>"
  } else if (identical(shape_name, "text")) {
    "<a:noFill/>"
  } else if (is_transparent) {
    "<a:noFill/>"
  } else {
    paste0("<a:solidFill><a:srgbClr val=\"", fill_hex, "\"/></a:solidFill>")
  }

  sw_emu     <- max(as.integer(ctx$sw_pt * 12700 * scale / 9525), 1588L)
  stroke_val <- if (!is.na(stroke_hex)) stroke_hex else ctx$default_stroke
  stroke_xml <- if (identical(shape_name, "text")) {
    "<a:ln w=\"0\"><a:noFill/></a:ln>"
  } else {
    paste0(
      "<a:ln w=\"", sw_emu, "\">",
      "<a:solidFill><a:srgbClr val=\"", stroke_val, "\"/></a:solidFill>",
      "</a:ln>"
    )
  }

  prst       <- .shape_map[shape_name]
  if (is.na(prst)) prst <- "rect"

  # leanL (lean-l) is the parallelogram preset mirrored horizontally
  flip_attr <- if (identical(shape_name, "leanL")) " flipH=\"1\"" else ""

  geom_xml <- paste0("<a:prstGeom prst=\"", prst, "\"><a:avLst/></a:prstGeom>")
  label    <- xml_escape(strip_html(nd$label %||% nd$id %||% ""))
  text_col <- node_text_colour(nd, is_transparent, fill_hex, ctx$dtc)


  descr <- jsonlite::toJSON(list(
    v = "1", type = "node",
    id = nd$id %||% "", label = nd$label %||% "",
    shape = shape_name, class = nd$class %||% NULL
  ), auto_unbox = TRUE, null = "null")

  nv_xml <- make_nvSpPr(shape_id,
                         paste0("mermaid:node:", nd$id %||% ""),
                         descr)

  paste0(
    "<wps:wsp>",
      nv_xml,
      "<wps:spPr>",
        "<a:xfrm", flip_attr, ">",
          "<a:off x=\"", x_rel, "\" y=\"", y_rel, "\"/>",
          "<a:ext cx=\"", w_emu, "\" cy=\"", h_emu, "\"/>",
        "</a:xfrm>",
        geom_xml, fill_xml, stroke_xml,
      "</wps:spPr>",
      "<wps:txbx><w:txbxContent><w:p>",
        "<w:pPr><w:jc w:val=\"center\"/></w:pPr>",
        "<w:r>", text_rpr_xml(text_col, ctx),
          "<w:t xml:space=\"preserve\">", label, "</w:t></w:r>",
      "</w:p></w:txbxContent></wps:txbx>",
      "<wps:bodyPr anchor=\"ctr\" ",
        "lIns=\"0\" rIns=\"0\" ",
        "tIns=\"0\" bIns=\"0\">",
        "<a:normAutofit/></wps:bodyPr>",
    "</wps:wsp>"
  )
}

# Visual-only wsp for diamond/hex group: shape fill/stroke, no text, at (0,0).
node_visual_wsp <- function(nd, w_emu, h_emu, ctx, ctr) {
  shape_id   <- ctr$next_id()
  scale      <- ctx$scale
  fill_hex   <- nd$fill   %||% NA_character_
  stroke_hex <- nd$stroke %||% NA_character_
  shape_name <- nd$shape  %||% "rect"

  # Forced solid-black fill for junction (f-circ) and fork shapes
  fill_xml <- if (shape_name %in% c("junction", "fork")) {
    "<a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>"
  } else if (is.na(fill_hex)) {
    "<a:noFill/>"
  } else {
    paste0("<a:solidFill><a:srgbClr val=\"", fill_hex, "\"/></a:solidFill>")
  }

  sw_emu     <- max(as.integer(ctx$sw_pt * 12700 * scale / 9525), 1588L)
  stroke_val <- if (!is.na(stroke_hex)) stroke_hex else ctx$default_stroke
  stroke_xml <- paste0(
    "<a:ln w=\"", sw_emu, "\">",
    "<a:solidFill><a:srgbClr val=\"", stroke_val, "\"/></a:solidFill>",
    "</a:ln>"
  )

  prst       <- .shape_map[shape_name]
  if (is.na(prst)) prst <- "rect"
  geom_xml   <- paste0("<a:prstGeom prst=\"", prst, "\"><a:avLst/></a:prstGeom>")

  # leanL (lean-l) uses the parallelogram preset mirrored horizontally
  flip_attr <- if (identical(shape_name, "leanL")) " flipH=\"1\"" else ""

  descr <- jsonlite::toJSON(list(
    v = "1", type = "node",
    id = nd$id %||% "", label = nd$label %||% "",
    shape = shape_name, class = nd$class %||% NULL
  ), auto_unbox = TRUE, null = "null")

  nv_xml <- make_nvSpPr(shape_id,
                         paste0("mermaid:node:", nd$id %||% ""),
                         descr)

  paste0(
    "<wps:wsp>",
      nv_xml,
      "<wps:spPr>",
        "<a:xfrm", flip_attr, ">",
          "<a:off x=\"0\" y=\"0\"/>",
          "<a:ext cx=\"", w_emu, "\" cy=\"", h_emu, "\"/>",
        "</a:xfrm>",
        geom_xml, fill_xml, stroke_xml,
      "</wps:spPr>",
      "<wps:bodyPr><a:noAutofit/></wps:bodyPr>",
    "</wps:wsp>"
  )
}

# Transparent rect text overlay wsp for diamond/hex group.
node_text_overlay_wsp <- function(nd, tx_rel, ty_rel, tw_emu, th_emu, ctx, ctr) {
  shape_id       <- ctr$next_id()
  fill_hex       <- nd$fill %||% NA_character_
  is_transparent <- is.na(fill_hex)
  label          <- xml_escape(strip_html(nd$label %||% nd$id %||% ""))
  text_col       <- node_text_colour(nd, is_transparent, fill_hex, ctx$dtc)

  nv_xml <- make_nvSpPr(shape_id,
                         paste0("mermaid:label:", nd$id %||% ""),
                         paste0('{"v":"1","type":"node-label","id":"',
                                nd$id %||% "", '"}'))

  paste0(
    "<wps:wsp>",
      nv_xml,
      "<wps:spPr>",
        "<a:xfrm>",
          "<a:off x=\"", tx_rel, "\" y=\"", ty_rel, "\"/>",
          "<a:ext cx=\"", tw_emu, "\" cy=\"", th_emu, "\"/>",
        "</a:xfrm>",
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
        "<a:noFill/><a:ln w=\"0\"><a:noFill/></a:ln>",
      "</wps:spPr>",
      "<wps:txbx><w:txbxContent><w:p>",
        "<w:pPr><w:jc w:val=\"center\"/></w:pPr>",
        "<w:r>", text_rpr_xml(text_col, ctx),
          "<w:t xml:space=\"preserve\">", label, "</w:t></w:r>",
      "</w:p></w:txbxContent></wps:txbx>",
      "<wps:bodyPr anchor=\"ctr\" lIns=\"0\" rIns=\"0\" tIns=\"0\" bIns=\"0\">",
        "<a:normAutofit/></wps:bodyPr>",
    "</wps:wsp>"
  )
}

# ── Edge emission ──────────────────────────────────────────────────────────

emit_edge <- function(e, origin_x_px, origin_y_px, ctx, ctr) {
  path_d <- e$path_d %||% ""
  if (!nzchar(path_d)) return("")

  cg <- svg_path_to_custgeom(path_d, ctx$scale)
  if (is.null(cg)) return("")

  shape_id <- ctr$next_id()
  x_rel    <- cg$x_emu - as.integer(round(origin_x_px * ctx$scale))
  y_rel    <- cg$y_emu - as.integer(round(origin_y_px * ctx$scale))
  w_emu    <- cg$w_emu
  h_emu    <- cg$h_emu

  lt       <- e$line_type %||% "solid"
  sw_emu   <- if (identical(lt, "thick")) as.integer(ctx$edge_sw_emu * 2L) else ctx$edge_sw_emu
  dash_xml <- if (identical(lt, "dashed")) "<a:prstDash val=\"dash\"/>" else ""

  make_end <- function(type, which) {
    type <- type %||% "none"
    if (type == "arrow")  return(paste0("<a:", which, "End type=\"arrow\" w=\"sm\" len=\"sm\"/>"))
    if (type == "circle") return(paste0("<a:", which, "End type=\"oval\" w=\"sm\" len=\"sm\"/>"))
    if (type == "cross")  return(paste0("<a:", which, "End type=\"diamond\" w=\"sm\" len=\"sm\"/>"))
    paste0("<a:", which, "End type=\"none\"/>")
  }

  descr <- jsonlite::toJSON(list(
    v = "1", type = "edge",
    from = e$from %||% NULL, to = e$to %||% NULL,
    label = e$label %||% NULL, line = lt,
    arrow_start = e$arrow_start %||% "none",
    arrow_end   = e$arrow_end   %||% "arrow"
  ), auto_unbox = TRUE, null = "null")

  nv_xml <- make_nvCnPr(shape_id,
                         paste0("mermaid:edge:",
                                e$from %||% "?", "-", e$to %||% "?"),
                         descr)

  paste0(
    "<wps:wsp>",
      nv_xml,
      "<wps:spPr>",
        "<a:xfrm>",
          "<a:off x=\"", x_rel, "\" y=\"", y_rel, "\"/>",
          "<a:ext cx=\"", w_emu, "\" cy=\"", h_emu, "\"/>",
        "</a:xfrm>",
        cg$xml,
        "<a:noFill/>",
        "<a:ln w=\"", sw_emu, "\">",
          "<a:solidFill><a:srgbClr val=\"",
            if (!is.na(e$stroke %||% NA_character_)) e$stroke else ctx$edge_stroke,
          "\"/></a:solidFill>",
          dash_xml,
          make_end(e$arrow_start, "head"),
          make_end(e$arrow_end,   "tail"),
        "</a:ln>",
      "</wps:spPr>",
      "<wps:bodyPr/>",
    "</wps:wsp>"
  )
}

emit_edge_label <- function(e, origin_x_px, origin_y_px, ctx, ctr) {
  lx <- as.numeric(e$label_x %||% NA_real_)
  ly <- as.numeric(e$label_y %||% NA_real_)
  if (is.na(lx) || is.na(ly)) return("")

  shape_id <- ctr$next_id()
  lbl      <- xml_escape(strip_html(e$label %||% ""))
  w_emu    <- inches_to_emu(1.2)
  h_emu    <- inches_to_emu(0.3)
  x_rel    <- as.integer(round((lx - origin_x_px) * ctx$scale)) - w_emu %/% 2L
  y_rel    <- as.integer(round((ly - origin_y_px) * ctx$scale)) - h_emu %/% 2L

  nv_xml <- make_nvSpPr(shape_id,
                         paste0("mermaid:edge_label:", e$from %||% "", ":", e$to %||% ""),
                         "{}")

  paste0(
    "<wps:wsp>",
      nv_xml,
      "<wps:spPr>",
        "<a:xfrm>",
          "<a:off x=\"", x_rel, "\" y=\"", y_rel, "\"/>",
          "<a:ext cx=\"", w_emu, "\" cy=\"", h_emu, "\"/>",
        "</a:xfrm>",
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
        "<a:noFill/><a:ln w=\"0\"><a:noFill/></a:ln>",
      "</wps:spPr>",
      "<wps:txbx><w:txbxContent><w:p>",
        "<w:pPr><w:jc w:val=\"center\"/></w:pPr>",
        "<w:r>", text_rpr_xml(ctx$dtc, ctx),
          "<w:t xml:space=\"preserve\">", lbl, "</w:t></w:r>",
      "</w:p></w:txbxContent></wps:txbx>",
      "<wps:bodyPr anchor=\"ctr\"><a:normAutofit/></wps:bodyPr>",
    "</wps:wsp>"
  )
}

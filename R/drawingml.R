# R/drawingml.R
#
# Converts parsed SVG geometry (from parse_svg.R) + the mermaid AST into a
# DrawingML XML string (<w:p>…</w:p>) ready for officer::body_add_xml().
#
# Nodes  → <wps:wsp> with <a:prstGeom> for standard shapes
# Edges  → <wps:wsp cNvCnPr/> with <a:custGeom> reproducing dagre bezier routing
# Labels → small transparent text-box at the label midpoint

.dml_namespaces <- paste0(
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ',
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
)

.shape_map <- c(
  rect          = "rect",
  roundRect     = "roundRect",
  diamond       = "diamond",
  ellipse       = "ellipse",
  hexagon       = "hexagon",
  parallelogram = "parallelogram",
  cylinder      = "can",
  text          = "rect"
)

# ── Main exported function ─────────────────────────────────────────────────

#' Build DrawingML XML from parsed mermaid SVG geometry
#'
#' @param svg_data List from [parse_mermaid_svg()]: nodes, edges, viewbox.
#' @param ast List. Reserved for future use. The AST from `@mermaid-js/parser`
#'   is accepted here for forward-compatibility but is not currently consulted.
#'   All shape data is derived from `svg_data`. See `enrich_from_ast()` in
#'   `parse_svg.R` for the intended use and current status.
#' @param start_id Integer. First shape ID; increment between diagrams.
#' @param page_width_in Numeric. Usable content width in inches (default 6.0).
#' @param page_height_in Numeric. Max diagram height in inches (default 8.0).
#' @param margin_in Numeric. Left/top margin offset in inches (default 1.0).
#' @param default_fill Default node fill colour (6-char hex). Default "FFFFFF".
#' @param default_stroke Default stroke colour. Default "4472C4".
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
  sty    <- svg_data$style %||% list(font_size_px = 14, edge_stroke = default_stroke, edge_sw_px = 2)

  svg_w  <- vb[3]; svg_h <- vb[4]
  scale  <- compute_scale(svg_w, svg_h, page_width_in, page_height_in)
  off_x  <- inches_to_emu(margin_in)
  off_y  <- inches_to_emu(margin_in)

  # Scale font size proportionally with the diagram; floor at 8 half-pts (4pt).
  # scale is in EMU/px; 6350 EMU = 0.5pt, so font_px * scale / 6350 = half-points.
  font_size_hp <- max(as.integer(round(sty$font_size_px * scale / 6350)), 8L)

  # Edge stroke: colour from SVG stylesheet; width scaled from SVG px, min 0.25pt.
  edge_stroke  <- sty$edge_stroke %||% default_stroke
  edge_sw_emu  <- max(as.integer(round(sty$edge_sw_px * scale)), 3175L)

  subgraphs <- svg_data$subgraphs %||% empty_subgraphs_tbl()

  shape_id <- as.integer(start_id)
  parts    <- character(0)

  # Subgraphs first — lowest z-order so they render behind everything else
  if (nrow(subgraphs) > 0L) {
    for (i in seq_len(nrow(subgraphs))) {
      sg  <- as.list(subgraphs[i, ])
      xml <- subgraph_shape_xml(sg, shape_id, vb, scale, off_x, off_y,
                                stroke_width_pt, font_size_hp, default_text_color)
      if (nzchar(xml)) { parts <- c(parts, xml); shape_id <- shape_id + 1L }
    }
  }

  if (nrow(nodes) > 0L) {
    for (i in seq_len(nrow(nodes))) {
      nd  <- as.list(nodes[i, ])
      xml <- node_shape_xml(nd, shape_id, vb, scale, off_x, off_y,
                            default_fill, default_stroke, default_text_color,
                            stroke_width_pt, font_size_hp)
      if (nzchar(xml)) { parts <- c(parts, xml); shape_id <- shape_id + 1L }
    }
  }

  if (nrow(edges) > 0L) {
    for (i in seq_len(nrow(edges))) {
      e   <- as.list(edges[i, ])
      xml <- edge_shape_xml(e, shape_id, vb, scale, off_x, off_y,
                            edge_stroke, edge_sw_emu)
      if (nzchar(xml)) { parts <- c(parts, xml); shape_id <- shape_id + 1L }

      if (!is.na(e$label %||% NA_character_) &&
          nzchar(e$label %||% "") &&
          !is.na(e$label_x %||% NA_real_) &&
          !is.na(e$label_y %||% NA_real_)) {
        lxml <- edge_label_xml(e, shape_id, vb, scale, off_x, off_y,
                               default_text_color)
        if (nzchar(lxml)) { parts <- c(parts, lxml); shape_id <- shape_id + 1L }
      }
    }
  }

  xml <- paste0('<w:p ', .dml_namespaces, '>',
                paste0(parts, collapse = ""),
                '</w:p>')

  list(xml = xml, next_id = shape_id)
}

# ── Scale ─────────────────────────────────────────────────────────────────

compute_scale <- function(svg_w, svg_h, page_w_in, page_h_in) {
  if (is.na(svg_w) || svg_w <= 0 || is.na(svg_h) || svg_h <= 0)
    return(9525)  # 1 px = 9525 EMU at 96 dpi
  target_w <- inches_to_emu(page_w_in)
  target_h <- inches_to_emu(page_h_in)
  min(target_w / svg_w, target_h / svg_h)
}

# ── Anchor wrapper ─────────────────────────────────────────────────────────

make_anchor <- function(x_emu, y_emu, w_emu, h_emu,
                         shape_id, name, descr_json,
                         z_order = 251658240L, behind_doc = FALSE, inner_xml) {
  w_emu <- max(as.integer(w_emu), 45720L)
  h_emu <- max(as.integer(h_emu), 45720L)
  paste0(
    '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>',
    '<wp:anchor distT="0" distB="0" distL="0" distR="0" ',
    'simplePos="0" relativeHeight="', z_order, '" ',
    'behindDoc="', if (behind_doc) "1" else "0", '" ',
    'locked="0" layoutInCell="1" allowOverlap="1">',
    '<wp:simplePos x="0" y="0"/>',
    '<wp:positionH relativeFrom="page">',
      '<wp:posOffset>', as.integer(x_emu), '</wp:posOffset></wp:positionH>',
    '<wp:positionV relativeFrom="page">',
      '<wp:posOffset>', as.integer(y_emu), '</wp:posOffset></wp:positionV>',
    '<wp:extent cx="', as.integer(w_emu), '" cy="', as.integer(h_emu), '"/>',
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
    '<wp:wrapNone/>',
    '<wp:docPr id="', shape_id,
      '" name="', xml_escape(name),
      '" descr="', xml_escape(descr_json), '"/>',
    '<wp:cNvGraphicFramePr/>',
    '<a:graphic>',
    '<a:graphicData ',
      'uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">',
    inner_xml,
    '</a:graphicData></a:graphic>',
    '</wp:anchor></w:drawing></w:r>'
  )
}

# ── Shape-specific text insets ────────────────────────────────────────────
#
# DrawingML preset geometries reduce the visible text area in ways that the
# generic scale-derived margins don't account for:
#
#   diamond   — pointed top/bottom/left/right; inscribed rect ≈ 50% w × 50% h
#   hexagon   — angled left/right ends; ≈ 13% of width lost per side
#
# For other shapes the standard scale-proportional margins are fine.

.node_l_ins <- function(shape, w_emu, scale) {
  base <- max(as.integer(91440 * scale / 9525), 9144L)
  if      (shape == "diamond") max(as.integer(w_emu * 0.27), base)
  else if (shape == "hexagon") max(as.integer(w_emu * 0.13), base)
  else                         base
}

.node_t_ins <- function(shape, h_emu, scale) {
  base <- max(as.integer(45720 * scale / 9525), 4572L)
  if (shape == "diamond") max(as.integer(h_emu * 0.27), base)
  else                    base
}

# ── Node shapes ────────────────────────────────────────────────────────────

node_shape_xml <- function(nd, shape_id, vb, scale, off_x, off_y,
                            default_fill, default_stroke, default_text_color,
                            stroke_width_pt, font_size_hp) {

  cx_emu <- as.integer(round((nd$svg_cx - vb[1]) * scale)) + off_x
  cy_emu <- as.integer(round((nd$svg_cy - vb[2]) * scale)) + off_y
  w_emu  <- max(as.integer(round(nd$svg_w * scale)), 91440L)  # min 0.1 in
  h_emu  <- max(as.integer(round(nd$svg_h * scale)), 45720L)  # min 0.05 in
  x_emu  <- cx_emu - w_emu %/% 2L
  y_emu  <- cy_emu - h_emu %/% 2L

  fill_hex   <- nd$fill   %||% NA_character_
  stroke_hex <- nd$stroke %||% NA_character_

  is_transparent <- is.na(fill_hex)
  fill_xml <- if (is_transparent) {
    "<a:noFill/>"
  } else {
    paste0("<a:solidFill><a:srgbClr val=\"", fill_hex, "\"/></a:solidFill>")
  }

  sw_emu     <- max(as.integer(stroke_width_pt * 12700 * scale / 9525), 1588L)
  stroke_val <- if (!is.na(stroke_hex)) stroke_hex else default_stroke
  stroke_xml <- paste0(
    "<a:ln w=\"", sw_emu, "\">",
    "<a:solidFill><a:srgbClr val=\"", stroke_val, "\"/></a:solidFill>",
    "</a:ln>"
  )

  shape_name <- nd$shape %||% "rect"
  prst       <- .shape_map[shape_name]
  if (is.na(prst)) prst <- "rect"

  if (identical(shape_name, "text")) {
    fill_xml   <- "<a:noFill/>"
    stroke_xml <- "<a:ln w=\"0\"><a:noFill/></a:ln>"
  }

  geom_xml <- paste0("<a:prstGeom prst=\"", prst, "\"><a:avLst/></a:prstGeom>")

  label    <- xml_escape(strip_html(nd$label %||% nd$id %||% ""))
  # nd$color may be NA_character_ (not NULL) when no classDef colour was found;
  # %||% only catches NULL, so check explicitly.
  text_col <- if (!is.na(nd$color %||% NA_character_)) {
    nd$color
  } else if (!is_transparent && !is.na(fill_hex) && is_dark(fill_hex)) {
    "FFFFFF"
  } else {
    default_text_color
  }

  inner <- paste0(
    "<wps:wsp>",
      "<wps:cNvSpPr><a:spLocks noChangeArrowheads=\"1\"/></wps:cNvSpPr>",
      "<wps:spPr>",
        "<a:xfrm><a:off x=\"0\" y=\"0\"/>",
          "<a:ext cx=\"", w_emu, "\" cy=\"", h_emu, "\"/></a:xfrm>",
        geom_xml, fill_xml, stroke_xml,
      "</wps:spPr>",
      "<wps:txbx><w:txbxContent><w:p>",
        "<w:pPr><w:jc w:val=\"center\"/></w:pPr>",
        "<w:r><w:rPr>",
          "<w:color w:val=\"", text_col, "\"/>",
          "<w:sz w:val=\"", font_size_hp, "\"/>",
          "<w:szCs w:val=\"", font_size_hp, "\"/>",
        "</w:rPr>",
          "<w:t xml:space=\"preserve\">", label, "</w:t></w:r>",
      "</w:p></w:txbxContent></wps:txbx>",
      "<wps:bodyPr anchor=\"ctr\" ",
        "lIns=\"", .node_l_ins(shape_name, w_emu, scale), "\" ",
        "rIns=\"", .node_l_ins(shape_name, w_emu, scale), "\" ",
        "tIns=\"", .node_t_ins(shape_name, h_emu, scale), "\" ",
        "bIns=\"", .node_t_ins(shape_name, h_emu, scale), "\">",
        "<a:normAutofit/></wps:bodyPr>",
    "</wps:wsp>"
  )

  descr <- jsonlite::toJSON(list(
    v = "1", type = "node",
    id = nd$id %||% "", label = nd$label %||% "",
    shape = shape_name, class = nd$class %||% NULL
  ), auto_unbox = TRUE, null = "null")

  make_anchor(x_emu, y_emu, w_emu, h_emu, shape_id,
              name       = paste0("mermaid:node:", nd$id %||% ""),
              descr_json = descr,
              z_order    = 251658240L,
              inner_xml  = inner)
}

# ── Subgraph (cluster) shapes ─────────────────────────────────────────────

subgraph_shape_xml <- function(sg, shape_id, vb, scale, off_x, off_y,
                                stroke_width_pt, font_size_hp,
                                default_text_color) {
  # Cluster rects use absolute SVG coordinates (no translate on the <g>)
  x_emu <- as.integer(round((sg$svg_x - vb[1]) * scale)) + off_x
  y_emu <- as.integer(round((sg$svg_y - vb[2]) * scale)) + off_y
  w_emu <- max(as.integer(round(sg$svg_w * scale)), 91440L)
  h_emu <- max(as.integer(round(sg$svg_h * scale)), 45720L)

  fill_hex <- sg$fill %||% NA_character_
  fill_xml <- if (is.na(fill_hex)) {
    "<a:noFill/>"
  } else {
    paste0("<a:solidFill><a:srgbClr val=\"", fill_hex, "\"/></a:solidFill>")
  }

  stroke_hex <- sg$stroke %||% "AAAA33"
  sw_emu     <- max(as.integer(stroke_width_pt * 12700 * scale / 9525), 1588L)
  stroke_xml <- paste0(
    "<a:ln w=\"", sw_emu, "\">",
    "<a:solidFill><a:srgbClr val=\"", stroke_hex, "\"/></a:solidFill>",
    "</a:ln>"
  )

  label    <- xml_escape(strip_html(sg$label %||% ""))
  text_col <- if (!is.na(fill_hex) && is_dark(fill_hex)) "FFFFFF" else default_text_color

  l_ins <- max(as.integer(91440 * scale / 9525), 9144L)
  t_ins <- max(as.integer(45720 * scale / 9525), 4572L)

  inner <- paste0(
    "<wps:wsp>",
      "<wps:cNvSpPr><a:spLocks noChangeArrowheads=\"1\"/></wps:cNvSpPr>",
      "<wps:spPr>",
        "<a:xfrm><a:off x=\"0\" y=\"0\"/>",
          "<a:ext cx=\"", w_emu, "\" cy=\"", h_emu, "\"/></a:xfrm>",
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
        fill_xml, stroke_xml,
      "</wps:spPr>",
      "<wps:txbx><w:txbxContent><w:p>",
        "<w:pPr><w:jc w:val=\"center\"/></w:pPr>",
        "<w:r><w:rPr>",
          "<w:color w:val=\"", text_col, "\"/>",
          "<w:sz w:val=\"", font_size_hp, "\"/>",
          "<w:szCs w:val=\"", font_size_hp, "\"/>",
        "</w:rPr>",
          "<w:t xml:space=\"preserve\">", label, "</w:t></w:r>",
      "</w:p></w:txbxContent></wps:txbx>",
      "<wps:bodyPr anchor=\"t\" ",
        "lIns=\"", l_ins, "\" rIns=\"", l_ins, "\" ",
        "tIns=\"", t_ins, "\" bIns=\"", t_ins, "\">",
        "<a:normAutofit/></wps:bodyPr>",
    "</wps:wsp>"
  )

  descr <- jsonlite::toJSON(list(
    v = "1", type = "subgraph",
    id = sg$id %||% "", label = sg$label %||% "",
    class = sg$class %||% NULL
  ), auto_unbox = TRUE, null = "null")

  # behindDoc=TRUE places subgraphs in the "behind text" layer, which is always
  # visually behind the "in front of text" layer used by nodes and edges.
  # This is more reliable than trying to win the relativeHeight z-order battle.
  make_anchor(x_emu, y_emu, w_emu, h_emu, shape_id,
              name       = paste0("mermaid:subgraph:", sg$id %||% ""),
              descr_json = descr,
              z_order    = 251658240L,
              behind_doc = TRUE,
              inner_xml  = inner)
}

# ── Edge shapes (custGeom from dagre bezier path) ──────────────────────────

edge_shape_xml <- function(e, shape_id, vb, scale, off_x, off_y,
                            edge_stroke, edge_sw_emu) {
  path_d <- e$path_d %||% ""
  if (!nzchar(path_d)) return("")

  cg <- svg_path_to_custgeom(path_d, scale)
  if (is.null(cg)) return("")

  # Shift from SVG coordinate space to page EMU
  x_emu <- cg$x_emu - as.integer(round(vb[1] * scale)) + off_x
  y_emu <- cg$y_emu - as.integer(round(vb[2] * scale)) + off_y
  w_emu <- cg$w_emu
  h_emu <- cg$h_emu

  lt       <- e$line_type %||% "solid"
  sw_emu   <- if (identical(lt, "thick")) as.integer(edge_sw_emu * 2L) else edge_sw_emu
  dash_xml <- if (identical(lt, "dashed")) "<a:prstDash val=\"dash\"/>" else ""

  make_end <- function(type, which) {
    type <- type %||% "none"
    if (type == "arrow")  return(paste0("<a:", which, "End type=\"arrow\" w=\"sm\" len=\"sm\"/>"))
    if (type == "circle") return(paste0("<a:", which, "End type=\"oval\" w=\"sm\" len=\"sm\"/>"))
    if (type == "cross")  return(paste0("<a:", which, "End type=\"diamond\" w=\"sm\" len=\"sm\"/>"))
    paste0("<a:", which, "End type=\"none\"/>")
  }

  inner <- paste0(
    "<wps:wsp>",
      "<wps:cNvCnPr/>",
      "<wps:spPr>",
        "<a:xfrm><a:off x=\"0\" y=\"0\"/>",
          "<a:ext cx=\"", w_emu, "\" cy=\"", h_emu, "\"/></a:xfrm>",
        cg$xml,
        "<a:noFill/>",
        "<a:ln w=\"", sw_emu, "\">",
          "<a:solidFill><a:srgbClr val=\"", edge_stroke, "\"/></a:solidFill>",
          dash_xml,
          make_end(e$arrow_start, "head"),
          make_end(e$arrow_end,   "tail"),
        "</a:ln>",
      "</wps:spPr>",
      "<wps:bodyPr/>",
    "</wps:wsp>"
  )

  descr <- jsonlite::toJSON(list(
    v = "1", type = "edge",
    from = e$from %||% NULL, to = e$to %||% NULL,
    label = e$label %||% NULL, line = lt,
    arrow_start = e$arrow_start %||% "none",
    arrow_end   = e$arrow_end   %||% "arrow"
  ), auto_unbox = TRUE, null = "null")

  make_anchor(x_emu, y_emu, w_emu, h_emu, shape_id,
              name       = paste0("mermaid:edge:",
                                  e$from %||% "?", "\u2192", e$to %||% "?"),
              descr_json = descr,
              z_order    = 251659000L,
              inner_xml  = inner)
}

# ── Edge label text boxes ──────────────────────────────────────────────────

edge_label_xml <- function(e, shape_id, vb, scale, off_x, off_y,
                            default_text_color) {
  lx <- as.numeric(e$label_x %||% NA_real_)
  ly <- as.numeric(e$label_y %||% NA_real_)
  if (is.na(lx) || is.na(ly)) return("")

  lbl   <- xml_escape(strip_html(e$label %||% ""))
  w_emu <- inches_to_emu(1.2)
  h_emu <- inches_to_emu(0.3)
  x_emu <- as.integer(round((lx - vb[1]) * scale)) + off_x - w_emu %/% 2L
  y_emu <- as.integer(round((ly - vb[2]) * scale)) + off_y - h_emu %/% 2L

  inner <- paste0(
    "<wps:wsp>",
      "<wps:cNvSpPr><a:spLocks noChangeArrowheads=\"1\"/></wps:cNvSpPr>",
      "<wps:spPr>",
        "<a:xfrm><a:off x=\"0\" y=\"0\"/>",
          "<a:ext cx=\"", w_emu, "\" cy=\"", h_emu, "\"/></a:xfrm>",
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
        "<a:noFill/><a:ln w=\"0\"><a:noFill/></a:ln>",
      "</wps:spPr>",
      "<wps:txbx><w:txbxContent><w:p>",
        "<w:pPr><w:jc w:val=\"center\"/></w:pPr>",
        "<w:r><w:rPr><w:color w:val=\"", default_text_color, "\"/>",
          "<w:sz w:val=\"16\"/></w:rPr>",
          "<w:t xml:space=\"preserve\">", lbl, "</w:t></w:r>",
      "</w:p></w:txbxContent></wps:txbx>",
      "<wps:bodyPr anchor=\"ctr\"><a:normAutofit/></wps:bodyPr>",
    "</wps:wsp>"
  )

  make_anchor(x_emu, y_emu, w_emu, h_emu, shape_id,
              name       = paste0("mermaid:edge_label:",
                                  e$from %||% "", ":", e$to %||% ""),
              descr_json = "{}",
              z_order    = 251660000L,
              inner_xml  = inner)
}

# ── Colour helpers ─────────────────────────────────────────────────────────

is_dark <- function(hex) {
  if (is.na(hex) || nchar(hex) < 6L) return(FALSE)
  r   <- strtoi(substr(hex, 1L, 2L), 16L)
  g   <- strtoi(substr(hex, 3L, 4L), 16L)
  b   <- strtoi(substr(hex, 5L, 6L), 16L)
  (0.299 * r + 0.587 * g + 0.114 * b) < 128
}

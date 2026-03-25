# R/drawingml.R
#
# Converts parsed SVG geometry (from parse_svg.R) into DrawingML XML
# (<w:p>…</w:p>) ready for officer::body_add_xml().
#
# Architecture: flat — every shape is its own <wp:anchor> element.
#   - Subgraph backgrounds: behindDoc="1" anchors (emitted first, drawn behind)
#   - Regular nodes:        anchor > graphicData[wps] > wps:wsp
#   - Diamond/hex/para nodes: anchor > graphicData[wpg] > wpg:wgp (depth-1)
#                              containing visual wsp + transparent text wsp
#   - Edges:                anchor > graphicData[wps] > wps:wsp (custom geom)
#
# NOTE: Nested wpg:wgp groups (depth >= 2) are spec-valid but do not render
# in Word's implementation. This flat approach avoids that limitation while
# still grouping diamond/hexagon shapes with their text overlays.

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
  rect          = "rect",
  roundRect     = "roundRect",
  diamond       = "diamond",
  ellipse       = "ellipse",
  hexagon       = "hexagon",
  parallelogram = "parallelogram",
  cylinder      = "can",
  text          = "rect"
)

# ── ID counter ─────────────────────────────────────────────────────────────

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

  # Edge stroke: colour from SVG stylesheet; width scaled from SVG px, min 0.25pt.
  edge_stroke <- sty$edge_stroke %||% default_stroke
  edge_sw_emu <- max(as.integer(round(sty$edge_sw_px * scale)), 3175L)

  subgraphs <- svg_data$subgraphs %||% empty_subgraphs_tbl()

  ctx <- list(
    vb             = vb,
    scale          = scale,
    font_size_hp   = font_size_hp,
    edge_stroke    = edge_stroke,
    edge_sw_emu    = edge_sw_emu,
    default_fill   = default_fill,
    default_stroke = default_stroke,
    dtc            = default_text_color,
    sw_pt          = stroke_width_pt
  )

  ctr   <- new_id_counter(as.integer(start_id))
  parts <- character(0)

  # 1. Subgraph backgrounds — behindDoc=1 so they appear behind nodes
  if (nrow(subgraphs) > 0L) {
    for (i in seq_len(nrow(subgraphs))) {
      sg  <- as.list(subgraphs[i, ])
      xml <- emit_subgraph_anchor(sg, off_x, off_y, vb, ctx, ctr)
      if (nzchar(xml)) parts <- c(parts, xml)
    }
  }

  # 2. Nodes
  if (nrow(nodes) > 0L) {
    for (i in seq_len(nrow(nodes))) {
      nd  <- as.list(nodes[i, ])
      xml <- emit_node(nd, off_x, off_y, vb, ctx, ctr)
      if (nzchar(xml)) parts <- c(parts, xml)
    }
  }

  # 3. Edges and their labels
  if (nrow(edges) > 0L) {
    for (i in seq_len(nrow(edges))) {
      e   <- as.list(edges[i, ])
      xml <- emit_edge(e, off_x, off_y, vb, ctx, ctr)
      if (nzchar(xml)) parts <- c(parts, xml)

      lbl <- e$label %||% ""
      if (nzchar(lbl) &&
          !is.na(e$label_x %||% NA_real_) &&
          !is.na(e$label_y %||% NA_real_)) {
        lxml <- emit_edge_label(e, off_x, off_y, vb, ctx, ctr)
        if (nzchar(lxml)) parts <- c(parts, lxml)
      }
    }
  }

  xml <- paste0('<w:p ', .dml_namespaces, '>',
                paste0(parts, collapse = ""),
                '</w:p>')

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

# ── Anchor primitives ──────────────────────────────────────────────────────

# Wrap a wps:wsp XML string in a wp:anchor for a single floating shape.
make_shape_anchor <- function(x_page, y_page, w_emu, h_emu, shape_id,
                               wsp_xml, behind_doc = FALSE) {
  w_emu   <- max(as.integer(w_emu), 91440L)
  h_emu   <- max(as.integer(h_emu), 91440L)
  behind  <- if (behind_doc) "1" else "0"
  rel_h   <- if (behind_doc) "1" else "251658240"
  paste0(
    '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>',
    '<wp:anchor distT="0" distB="0" distL="0" distR="0" ',
      'simplePos="0" relativeHeight="', rel_h, '" behindDoc="', behind, '" ',
      'locked="0" layoutInCell="1" allowOverlap="1">',
    '<wp:simplePos x="0" y="0"/>',
    '<wp:positionH relativeFrom="page"><wp:posOffset>',
      as.integer(x_page), '</wp:posOffset></wp:positionH>',
    '<wp:positionV relativeFrom="page"><wp:posOffset>',
      as.integer(y_page), '</wp:posOffset></wp:positionV>',
    '<wp:extent cx="', w_emu, '" cy="', h_emu, '"/>',
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
    '<wp:wrapNone/>',
    '<wp:docPr id="', shape_id, '" name="shape', shape_id, '"/>',
    '<wp:cNvGraphicFramePr/>',
    '<a:graphic>',
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">',
    wsp_xml,
    '</a:graphicData></a:graphic>',
    '</wp:anchor></w:drawing></w:r>'
  )
}

# Wrap children XML in a wp:anchor > wpg:wgp (depth-1 group).
make_group_anchor <- function(x_page, y_page, w_emu, h_emu, shape_id,
                               children_xml) {
  w_emu <- max(as.integer(w_emu), 91440L)
  h_emu <- max(as.integer(h_emu), 91440L)
  paste0(
    '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>',
    '<wp:anchor distT="0" distB="0" distL="0" distR="0" ',
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
    '<wp:docPr id="', shape_id, '" name="group', shape_id, '"/>',
    '<wp:cNvGraphicFramePr/>',
    '<a:graphic>',
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">',
    make_group_wgp(0L, 0L, w_emu, h_emu, children_xml),
    '</a:graphicData></a:graphic>',
    '</wp:anchor></w:drawing></w:r>'
  )
}

# Emit a depth-1 wpg:wgp group element (children use group-relative coords).
# chOff=(0,0) chExt=(w,h) gives 1:1 mapping between child coords and EMU.
make_group_wgp <- function(x_rel, y_rel, w_emu, h_emu, children_xml) {
  w_emu <- max(as.integer(w_emu), 91440L)
  h_emu <- max(as.integer(h_emu), 91440L)
  paste0(
    '<wpg:wgp>',
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
    '</wpg:wgp>'
  )
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

# ── Non-visual property stubs ──────────────────────────────────────────────
# wps:wsp only allows <wps:cNvSpPr> (shapes) or <wps:cNvCnPr/> (connectors).

make_nvSpPr <- function(shape_id, name, descr_json) {
  '<wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>'
}

make_nvCnPr <- function(shape_id, name, descr_json) {
  '<wps:cNvCnPr/>'
}

# ── Subgraph emission ──────────────────────────────────────────────────────

emit_subgraph_anchor <- function(sg, off_x, off_y, vb, ctx, ctr) {
  scale    <- ctx$scale
  x_abs    <- off_x + as.integer(round((sg$svg_x - vb[1]) * scale))
  y_abs    <- off_y + as.integer(round((sg$svg_y - vb[2]) * scale))
  w_emu    <- max(as.integer(round(sg$svg_w * scale)), 91440L)
  h_emu    <- max(as.integer(round(sg$svg_h * scale)), 91440L)
  shape_id <- ctr$next_id()

  fill_hex <- sg$fill %||% NA_character_
  fill_xml <- if (is.na(fill_hex)) {
    "<a:noFill/>"
  } else {
    paste0("<a:solidFill><a:srgbClr val=\"", fill_hex, "\"/></a:solidFill>")
  }

  stroke_hex <- sg$stroke %||% "AAAA33"
  sw_emu     <- max(as.integer(ctx$sw_pt * 12700 * scale / 9525), 1588L)
  stroke_xml <- paste0(
    "<a:ln w=\"", sw_emu, "\">",
    "<a:solidFill><a:srgbClr val=\"", stroke_hex, "\"/></a:solidFill>",
    "</a:ln>"
  )

  label    <- xml_escape(strip_html(sg$label %||% ""))
  text_col <- if (!is.na(fill_hex) && is_dark(fill_hex)) "FFFFFF" else ctx$dtc
  l_ins    <- max(as.integer(91440 * scale / 9525), 9144L)
  t_ins    <- max(as.integer(45720 * scale / 9525), 4572L)

  descr <- jsonlite::toJSON(list(
    v = "1", type = "subgraph",
    id = sg$id %||% "", label = sg$label %||% "",
    class = sg$class %||% NULL
  ), auto_unbox = TRUE, null = "null")

  nv_xml <- make_nvSpPr(shape_id,
                         paste0("mermaid:subgraph:", sg$id %||% ""),
                         descr)

  wsp_xml <- paste0(
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
        "<w:r><w:rPr>",
          "<w:color w:val=\"", text_col, "\"/>",
          "<w:sz w:val=\"", ctx$font_size_hp, "\"/>",
          "<w:szCs w:val=\"", ctx$font_size_hp, "\"/>",
        "</w:rPr>",
          "<w:t xml:space=\"preserve\">", label, "</w:t></w:r>",
      "</w:p></w:txbxContent></wps:txbx>",
      "<wps:bodyPr anchor=\"t\" ",
        "lIns=\"", l_ins, "\" rIns=\"", l_ins, "\" ",
        "tIns=\"", t_ins, "\" bIns=\"", t_ins, "\">",
        "<a:normAutofit/></wps:bodyPr>",
    "</wps:wsp>"
  )

  make_shape_anchor(x_abs, y_abs, w_emu, h_emu, shape_id, wsp_xml,
                    behind_doc = TRUE)
}

# ── Node emission ──────────────────────────────────────────────────────────

emit_node <- function(nd, off_x, off_y, vb, ctx, ctr) {
  scale      <- ctx$scale
  w_emu      <- max(as.integer(round(nd$svg_w * scale)), 91440L)
  h_emu      <- max(as.integer(round(nd$svg_h * scale)), 45720L)
  x_abs      <- off_x + as.integer(round((nd$svg_cx - vb[1]) * scale)) - w_emu %/% 2L
  y_abs      <- off_y + as.integer(round((nd$svg_cy - vb[2]) * scale)) - h_emu %/% 2L
  shape_name <- nd$shape %||% "rect"

  if (shape_name %in% c("diamond", "hexagon", "parallelogram")) {
    emit_node_group_anchor(nd, x_abs, y_abs, w_emu, h_emu, ctx, ctr)
  } else {
    shape_id <- ctr$next_id()
    wsp_xml  <- node_wsp_body(nd, w_emu, h_emu, ctx)
    make_shape_anchor(x_abs, y_abs, w_emu, h_emu, shape_id, wsp_xml)
  }
}

# Depth-1 group anchor for diamond/hexagon/parallelogram:
# visual shape wsp at (0,0) + transparent text-overlay wsp centred in group.
emit_node_group_anchor <- function(nd, x_abs, y_abs, w_emu, h_emu, ctx, ctr) {
  shape_id   <- ctr$next_id()
  visual_xml <- node_visual_wsp_body(nd, w_emu, h_emu, ctx)

  raw_tw <- nd$txt_w %||% NA_real_
  raw_th <- nd$txt_h %||% NA_real_
  tw_emu <- if (!is.na(raw_tw) && raw_tw > 0)
              as.integer(round(raw_tw * ctx$scale)) else w_emu
  th_emu <- if (!is.na(raw_th) && raw_th > 0)
              as.integer(round(raw_th * ctx$scale)) else h_emu
  tx_rel <- (w_emu - tw_emu) %/% 2L
  ty_rel <- (h_emu - th_emu) %/% 2L

  text_xml <- node_text_overlay_wsp_body(nd, tx_rel, ty_rel, tw_emu, th_emu, ctx)

  make_group_anchor(x_abs, y_abs, w_emu, h_emu, shape_id,
                    paste0(visual_xml, text_xml))
}

# Regular node: shape fill/stroke + text body in one wps:wsp.
# Position is (0,0) — the anchor handles page position.
node_wsp_body <- function(nd, w_emu, h_emu, ctx) {
  scale      <- ctx$scale
  fill_hex   <- nd$fill   %||% NA_character_
  stroke_hex <- nd$stroke %||% NA_character_

  is_transparent <- is.na(fill_hex)
  fill_xml <- if (is_transparent) {
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

  shape_name <- nd$shape %||% "rect"
  prst       <- .shape_map[shape_name]
  if (is.na(prst)) prst <- "rect"

  if (identical(shape_name, "text")) {
    fill_xml   <- "<a:noFill/>"
    stroke_xml <- "<a:ln w=\"0\"><a:noFill/></a:ln>"
  }

  geom_xml <- paste0("<a:prstGeom prst=\"", prst, "\"><a:avLst/></a:prstGeom>")
  label    <- xml_escape(strip_html(nd$label %||% nd$id %||% ""))
  text_col <- node_text_colour(nd, is_transparent, fill_hex, ctx$dtc)

  l_ins <- max(as.integer(91440 * scale / 9525), 9144L)
  t_ins <- max(as.integer(45720 * scale / 9525), 4572L)

  descr <- jsonlite::toJSON(list(
    v = "1", type = "node",
    id = nd$id %||% "", label = nd$label %||% "",
    shape = shape_name, class = nd$class %||% NULL
  ), auto_unbox = TRUE, null = "null")

  nv_xml <- make_nvSpPr(0L, paste0("mermaid:node:", nd$id %||% ""), descr)

  paste0(
    "<wps:wsp>",
      nv_xml,
      "<wps:spPr>",
        "<a:xfrm>",
          "<a:off x=\"0\" y=\"0\"/>",
          "<a:ext cx=\"", w_emu, "\" cy=\"", h_emu, "\"/>",
        "</a:xfrm>",
        geom_xml, fill_xml, stroke_xml,
      "</wps:spPr>",
      "<wps:txbx><w:txbxContent><w:p>",
        "<w:pPr><w:jc w:val=\"center\"/></w:pPr>",
        "<w:r><w:rPr>",
          "<w:color w:val=\"", text_col, "\"/>",
          "<w:sz w:val=\"", ctx$font_size_hp, "\"/>",
          "<w:szCs w:val=\"", ctx$font_size_hp, "\"/>",
        "</w:rPr>",
          "<w:t xml:space=\"preserve\">", label, "</w:t></w:r>",
      "</w:p></w:txbxContent></wps:txbx>",
      "<wps:bodyPr anchor=\"ctr\" ",
        "lIns=\"", l_ins, "\" rIns=\"", l_ins, "\" ",
        "tIns=\"", t_ins, "\" bIns=\"", t_ins, "\">",
        "<a:normAutofit/></wps:bodyPr>",
    "</wps:wsp>"
  )
}

# Visual-only wsp for diamond/hex group: shape fill/stroke, no text, at (0,0).
node_visual_wsp_body <- function(nd, w_emu, h_emu, ctx) {
  scale      <- ctx$scale
  fill_hex   <- nd$fill   %||% NA_character_
  stroke_hex <- nd$stroke %||% NA_character_

  fill_xml <- if (is.na(fill_hex)) {
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

  shape_name <- nd$shape %||% "rect"
  prst       <- .shape_map[shape_name]
  if (is.na(prst)) prst <- "rect"
  geom_xml   <- paste0("<a:prstGeom prst=\"", prst, "\"><a:avLst/></a:prstGeom>")

  descr <- jsonlite::toJSON(list(
    v = "1", type = "node",
    id = nd$id %||% "", label = nd$label %||% "",
    shape = shape_name, class = nd$class %||% NULL
  ), auto_unbox = TRUE, null = "null")

  nv_xml <- make_nvSpPr(0L, paste0("mermaid:node:", nd$id %||% ""), descr)

  paste0(
    "<wps:wsp>",
      nv_xml,
      "<wps:spPr>",
        "<a:xfrm>",
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
node_text_overlay_wsp_body <- function(nd, tx_rel, ty_rel, tw_emu, th_emu, ctx) {
  fill_hex       <- nd$fill %||% NA_character_
  is_transparent <- is.na(fill_hex)
  label          <- xml_escape(strip_html(nd$label %||% nd$id %||% ""))
  text_col       <- node_text_colour(nd, is_transparent, fill_hex, ctx$dtc)

  nv_xml <- make_nvSpPr(0L,
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
        "<w:r><w:rPr>",
          "<w:color w:val=\"", text_col, "\"/>",
          "<w:sz w:val=\"", ctx$font_size_hp, "\"/>",
          "<w:szCs w:val=\"", ctx$font_size_hp, "\"/>",
        "</w:rPr>",
          "<w:t xml:space=\"preserve\">", label, "</w:t></w:r>",
      "</w:p></w:txbxContent></wps:txbx>",
      "<wps:bodyPr anchor=\"ctr\" lIns=\"0\" rIns=\"0\" tIns=\"0\" bIns=\"0\">",
        "<a:normAutofit/></wps:bodyPr>",
    "</wps:wsp>"
  )
}

# ── Edge emission ──────────────────────────────────────────────────────────

emit_edge <- function(e, off_x, off_y, vb, ctx, ctr) {
  path_d <- e$path_d %||% ""
  if (!nzchar(path_d)) return("")

  cg <- svg_path_to_custgeom(path_d, ctx$scale)
  if (is.null(cg)) return("")

  shape_id <- ctr$next_id()
  x_abs    <- off_x + cg$x_emu - as.integer(round(vb[1] * ctx$scale))
  y_abs    <- off_y + cg$y_emu - as.integer(round(vb[2] * ctx$scale))
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
                                e$from %||% "?", "\u2192", e$to %||% "?"),
                         descr)

  wsp_xml <- paste0(
    "<wps:wsp>",
      nv_xml,
      "<wps:spPr>",
        "<a:xfrm>",
          "<a:off x=\"0\" y=\"0\"/>",
          "<a:ext cx=\"", w_emu, "\" cy=\"", h_emu, "\"/>",
        "</a:xfrm>",
        cg$xml,
        "<a:noFill/>",
        "<a:ln w=\"", sw_emu, "\">",
          "<a:solidFill><a:srgbClr val=\"", ctx$edge_stroke, "\"/></a:solidFill>",
          dash_xml,
          make_end(e$arrow_start, "head"),
          make_end(e$arrow_end,   "tail"),
        "</a:ln>",
      "</wps:spPr>",
      "<wps:bodyPr/>",
    "</wps:wsp>"
  )

  make_shape_anchor(x_abs, y_abs, w_emu, h_emu, shape_id, wsp_xml)
}

emit_edge_label <- function(e, off_x, off_y, vb, ctx, ctr) {
  lx <- as.numeric(e$label_x %||% NA_real_)
  ly <- as.numeric(e$label_y %||% NA_real_)
  if (is.na(lx) || is.na(ly)) return("")

  shape_id <- ctr$next_id()
  lbl      <- xml_escape(strip_html(e$label %||% ""))
  w_emu    <- inches_to_emu(1.2)
  h_emu    <- inches_to_emu(0.3)
  x_abs    <- off_x + as.integer(round((lx - vb[1]) * ctx$scale)) - w_emu %/% 2L
  y_abs    <- off_y + as.integer(round((ly - vb[2]) * ctx$scale)) - h_emu %/% 2L

  nv_xml <- make_nvSpPr(shape_id,
                         paste0("mermaid:edge_label:", e$from %||% "", ":", e$to %||% ""),
                         "{}")

  wsp_xml <- paste0(
    "<wps:wsp>",
      nv_xml,
      "<wps:spPr>",
        "<a:xfrm>",
          "<a:off x=\"0\" y=\"0\"/>",
          "<a:ext cx=\"", w_emu, "\" cy=\"", h_emu, "\"/>",
        "</a:xfrm>",
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
        "<a:noFill/><a:ln w=\"0\"><a:noFill/></a:ln>",
      "</wps:spPr>",
      "<wps:txbx><w:txbxContent><w:p>",
        "<w:pPr><w:jc w:val=\"center\"/></w:pPr>",
        "<w:r><w:rPr><w:color w:val=\"", ctx$dtc, "\"/>",
          "<w:sz w:val=\"16\"/></w:rPr>",
          "<w:t xml:space=\"preserve\">", lbl, "</w:t></w:r>",
      "</w:p></w:txbxContent></wps:txbx>",
      "<wps:bodyPr anchor=\"ctr\"><a:normAutofit/></wps:bodyPr>",
    "</wps:wsp>"
  )

  make_shape_anchor(x_abs, y_abs, w_emu, h_emu, shape_id, wsp_xml)
}

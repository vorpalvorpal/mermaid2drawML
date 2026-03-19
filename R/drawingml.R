# DrawingML shape name mapping -------------------------------------------------

.shape_map <- c(
  rect          = "rect",
  roundRect     = "roundRect",
  diamond       = "diamond",
  ellipse       = "ellipse",
  hexagon       = "hexagon",
  parallelogram = "parallelogram",
  text          = "rect",
  cylinder      = "can"
)

# Namespace declarations -------------------------------------------------------

.dml_namespaces <- paste0(
  'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ',
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
)

# Anchor wrapper ---------------------------------------------------------------

make_anchor <- function(x_in, y_in, w_in, h_in, shape_id, name, descr_json,
                         z_order = 251658240L, inner_xml) {
  x_emu <- inches_to_emu(x_in)
  y_emu <- inches_to_emu(y_in)
  w_emu <- inches_to_emu(max(w_in, 0.05))
  h_emu <- inches_to_emu(max(h_in, 0.05))
  paste0(
    '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>',
    '<wp:anchor distT="0" distB="0" distL="0" distR="0" ',
    'simplePos="0" relativeHeight="', z_order, '" ',
    'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">',
    '<wp:simplePos x="0" y="0"/>',
    '<wp:positionH relativeFrom="page"><wp:posOffset>', x_emu, '</wp:posOffset></wp:positionH>',
    '<wp:positionV relativeFrom="page"><wp:posOffset>', y_emu, '</wp:posOffset></wp:positionV>',
    '<wp:extent cx="', w_emu, '" cy="', h_emu, '"/>',
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
    '<wp:wrapNone/>',
    '<wp:docPr id="', shape_id, '" name="', xml_escape(name), '" descr="', xml_escape(descr_json), '"/>',
    '<wp:cNvGraphicFramePr/>',
    '<a:graphic>',
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">',
    inner_xml,
    '</a:graphicData></a:graphic>',
    '</wp:anchor></w:drawing></w:r>'
  )
}

# Node XML builder -------------------------------------------------------------

node_to_xml <- function(node, shape_id, default_fill, default_stroke,
                         default_text_color, stroke_width_pt) {
  # Resolve colours
  fill_col  <- if (!is.null(node$fill)   && !is.na(node$fill))   node$fill   else default_fill
  strk_col  <- if (!is.null(node$stroke) && !is.na(node$stroke)) node$stroke else default_stroke
  text_col  <- if (!is.null(node$color)  && !is.na(node$color))  node$color  else default_text_color

  # Fall back fill to white if still NA
  if (is.null(fill_col) || is.na(fill_col)) fill_col <- "FFFFFF"

  # Shape geometry
  shape_key  <- if (!is.null(node$shape) && !is.na(node$shape) && node$shape %in% names(.shape_map)) {
    node$shape
  } else "rect"
  prst       <- .shape_map[[shape_key]]

  is_text_shape <- (shape_key == "text")

  # Stroke width in EMU (1 pt = 12700 EMU)
  sw_emu <- as.integer(round(stroke_width_pt * 12700))

  # Fill XML
  fill_xml <- if (is_text_shape || is.na(fill_col)) {
    "<a:noFill/>"
  } else {
    paste0('<a:solidFill><a:srgbClr val="', fill_col, '"/></a:solidFill>')
  }

  # Stroke XML
  ln_xml <- if (is_text_shape) {
    '<a:ln><a:noFill/></a:ln>'
  } else if (is.null(strk_col) || is.na(strk_col)) {
    paste0('<a:ln w="', sw_emu, '"><a:noFill/></a:ln>')
  } else {
    paste0('<a:ln w="', sw_emu, '"><a:solidFill><a:srgbClr val="', strk_col, '"/></a:solidFill></a:ln>')
  }

  # Text colour XML
  text_col_xml <- if (!is.null(text_col) && !is.na(text_col)) {
    paste0('<a:solidFill><a:srgbClr val="', text_col, '"/></a:solidFill>')
  } else {
    '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
  }

  # Label — already stripped of HTML by parser
  label_safe <- xml_escape(if (!is.null(node$label) && !is.na(node$label)) node$label else node$id)

  # Metadata JSON for round-trip
  meta <- jsonlite::toJSON(
    list(
      v     = "1",
      type  = "node",
      id    = node$id,
      label = if (!is.null(node$label) && !is.na(node$label)) node$label else node$id,
      shape = shape_key,
      class = if (!is.null(node$class)  && !is.na(node$class))  node$class  else ""
    ),
    auto_unbox = TRUE
  )

  name_attr <- paste0("mermaid:node:", node$id)

  inner_xml <- paste0(
    '<wps:wsp>',
    '<wps:cNvPr id="', shape_id, '" name="', xml_escape(name_attr), '"/>',
    '<wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>',
    '<wps:spPr>',
    '<a:xfrm>',
    '<a:off x="', inches_to_emu(node$x), '" y="', inches_to_emu(node$y), '"/>',
    '<a:ext cx="', inches_to_emu(node$width), '" cy="', inches_to_emu(node$height), '"/>',
    '</a:xfrm>',
    '<a:prstGeom prst="', prst, '"><a:avLst/></a:prstGeom>',
    fill_xml,
    ln_xml,
    '</wps:spPr>',
    '<wps:txbx>',
    '<w:txbxContent>',
    '<w:p>',
    '<w:pPr><w:jc w:val="center"/></w:pPr>',
    '<w:r>',
    '<w:rPr>',
    '<w:color w:val="', if (!is.null(text_col) && !is.na(text_col)) text_col else "000000", '"/>',
    '<w:sz w:val="16"/><w:szCs w:val="16"/>',
    '</w:rPr>',
    '<w:t xml:space="preserve">', label_safe, '</w:t>',
    '</w:r>',
    '</w:p>',
    '</w:txbxContent>',
    '</wps:txbx>',
    '<wps:bodyPr anchor="ctr" anchorCtr="1" vert="horz">',
    '<a:noAutofit/>',
    '</wps:bodyPr>',
    '</wps:wsp>'
  )

  make_anchor(
    x_in      = node$x,
    y_in      = node$y,
    w_in      = node$width,
    h_in      = node$height,
    shape_id  = shape_id,
    name      = name_attr,
    descr_json = meta,
    z_order   = 251658240L,
    inner_xml = inner_xml
  )
}

# Edge XML builder -------------------------------------------------------------

edge_to_xml <- function(edge, shape_id, default_stroke, stroke_width_pt) {
  # If coords are missing, skip
  if (is.null(edge$x1) || is.na(edge$x1) ||
      is.null(edge$y1) || is.na(edge$y1) ||
      is.null(edge$x2) || is.na(edge$x2) ||
      is.null(edge$y2) || is.na(edge$y2)) {
    return("")
  }

  x1 <- edge$x1; y1 <- edge$y1
  x2 <- edge$x2; y2 <- edge$y2

  # Bounding box
  box_x <- min(x1, x2)
  box_y <- min(y1, y2)
  box_w <- abs(x2 - x1)
  box_h <- abs(y2 - y1)

  # Ensure minimum size
  box_w <- max(box_w, 0.01)
  box_h <- max(box_h, 0.01)

  flipH <- x1 > x2
  flipV <- y1 > y2

  # Stroke colour
  strk_col <- if (!is.null(default_stroke) && !is.na(default_stroke)) default_stroke else "4472C4"

  # Stroke width in EMU
  sw_emu <- as.integer(round(stroke_width_pt * 12700))

  # Dashed line
  dash_xml <- if (!is.null(edge$line_type) && !is.na(edge$line_type) && edge$line_type == "dashed") {
    '<a:prstDash val="dash"/>'
  } else {
    ""
  }

  # Arrow end helper
  arrow_end_xml <- function(end_type, tag) {
    if (is.null(end_type) || is.na(end_type)) end_type <- "none"
    type_attr <- switch(end_type,
      arrow  = "arrow",
      circle = "oval",
      cross  = "diamond",
      "none"
    )
    if (type_attr == "none") {
      paste0('<a:', tag, ' type="none"/>')
    } else {
      paste0('<a:', tag, ' type="', type_attr, '" w="med" len="med"/>')
    }
  }

  head_xml <- arrow_end_xml(edge$arrow_start, "headEnd")
  tail_xml <- arrow_end_xml(edge$arrow_end,   "tailEnd")

  # Metadata JSON
  meta <- jsonlite::toJSON(
    list(
      v          = "1",
      type       = "edge",
      from       = edge$from,
      to         = edge$to,
      label      = if (!is.null(edge$label) && !is.na(edge$label)) edge$label else "",
      line       = if (!is.null(edge$line_type)   && !is.na(edge$line_type))   edge$line_type   else "solid",
      arrow_start = if (!is.null(edge$arrow_start) && !is.na(edge$arrow_start)) edge$arrow_start else "none",
      arrow_end  = if (!is.null(edge$arrow_end)   && !is.na(edge$arrow_end))   edge$arrow_end   else "arrow"
    ),
    auto_unbox = TRUE
  )

  name_attr <- paste0("mermaid:edge:", edge$from, ":", edge$to)

  flip_attr <- ""
  if (flipH && flipV) flip_attr <- ' flipH="1" flipV="1"'
  else if (flipH)     flip_attr <- ' flipH="1"'
  else if (flipV)     flip_attr <- ' flipV="1"'

  inner_xml <- paste0(
    '<wps:wsp>',
    '<wps:cNvPr id="', shape_id, '" name="', xml_escape(name_attr), '"/>',
    '<wps:cNvCnPr/>',
    '<wps:spPr>',
    '<a:xfrm', flip_attr, '>',
    '<a:off x="', inches_to_emu(box_x), '" y="', inches_to_emu(box_y), '"/>',
    '<a:ext cx="', inches_to_emu(box_w), '" cy="', inches_to_emu(box_h), '"/>',
    '</a:xfrm>',
    '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>',
    '<a:noFill/>',
    '<a:ln w="', sw_emu, '">',
    '<a:solidFill><a:srgbClr val="', strk_col, '"/></a:solidFill>',
    dash_xml,
    head_xml,
    tail_xml,
    '</a:ln>',
    '</wps:spPr>',
    '<wps:bodyPr/>',
    '</wps:wsp>'
  )

  make_anchor(
    x_in      = box_x,
    y_in      = box_y,
    w_in      = box_w,
    h_in      = box_h,
    shape_id  = shape_id,
    name      = name_attr,
    descr_json = meta,
    z_order   = 251658241L,
    inner_xml = inner_xml
  )
}

# Edge label text box ----------------------------------------------------------

edge_label_to_xml <- function(edge, shape_id, default_text_color) {
  if (is.null(edge$label) || is.na(edge$label) || !nzchar(edge$label)) return("")
  if (is.null(edge$x1) || is.na(edge$x1) || is.null(edge$x2) || is.na(edge$x2)) return("")
  if (is.null(edge$y1) || is.na(edge$y1) || is.null(edge$y2) || is.na(edge$y2)) return("")

  mid_x  <- (edge$x1 + edge$x2) / 2
  mid_y  <- (edge$y1 + edge$y2) / 2
  lbl_w  <- 1.2
  lbl_h  <- 0.3
  box_x  <- mid_x - lbl_w / 2
  box_y  <- mid_y - lbl_h / 2

  text_col <- if (!is.null(default_text_color) && !is.na(default_text_color)) default_text_color else "000000"

  label_safe <- xml_escape(edge$label)
  name_attr  <- paste0("mermaid:edge_label:", edge$from, ":", edge$to)

  inner_xml <- paste0(
    '<wps:wsp>',
    '<wps:cNvPr id="', shape_id, '" name="', xml_escape(name_attr), '"/>',
    '<wps:cNvSpPr txBox="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>',
    '<wps:spPr>',
    '<a:xfrm>',
    '<a:off x="', inches_to_emu(box_x), '" y="', inches_to_emu(box_y), '"/>',
    '<a:ext cx="', inches_to_emu(lbl_w), '" cy="', inches_to_emu(lbl_h), '"/>',
    '</a:xfrm>',
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
    '<a:noFill/>',
    '<a:ln><a:noFill/></a:ln>',
    '</wps:spPr>',
    '<wps:txbx>',
    '<w:txbxContent>',
    '<w:p>',
    '<w:pPr><w:jc w:val="center"/></w:pPr>',
    '<w:r>',
    '<w:rPr>',
    '<w:color w:val="', text_col, '"/>',
    '<w:sz w:val="14"/><w:szCs w:val="14"/>',
    '</w:rPr>',
    '<w:t xml:space="preserve">', label_safe, '</w:t>',
    '</w:r>',
    '</w:p>',
    '</w:txbxContent>',
    '</wps:txbx>',
    '<wps:bodyPr anchor="ctr" anchorCtr="1" vert="horz">',
    '<a:noAutofit/>',
    '</wps:bodyPr>',
    '</wps:wsp>'
  )

  make_anchor(
    x_in      = box_x,
    y_in      = box_y,
    w_in      = lbl_w,
    h_in      = lbl_h,
    shape_id  = shape_id,
    name      = name_attr,
    descr_json = "{}",
    z_order   = 251658242L,
    inner_xml = inner_xml
  )
}

# Subgraph XML builder ---------------------------------------------------------

subgraph_to_xml <- function(sg, nodes_df, shape_id, default_stroke, padding_in = 0.2) {
  nids <- sg$node_ids
  if (length(nids) == 0) return("")

  sg_nodes <- nodes_df[nodes_df$id %in% nids, , drop = FALSE]
  if (nrow(sg_nodes) == 0) return("")

  # Check layout columns present
  if (!all(c("x", "y", "width", "height") %in% names(sg_nodes))) return("")

  min_x <- min(sg_nodes$x,                     na.rm = TRUE)
  min_y <- min(sg_nodes$y,                     na.rm = TRUE)
  max_x <- max(sg_nodes$x + sg_nodes$width,    na.rm = TRUE)
  max_y <- max(sg_nodes$y + sg_nodes$height,   na.rm = TRUE)

  box_x <- min_x - padding_in
  box_y <- min_y - padding_in
  box_w <- (max_x - min_x) + 2 * padding_in
  box_h <- (max_y - min_y) + 2 * padding_in

  box_w <- max(box_w, 0.1)
  box_h <- max(box_h, 0.1)

  strk_col <- if (!is.null(sg$stroke) && !is.na(sg$stroke)) sg$stroke else
    if (!is.null(default_stroke) && !is.na(default_stroke)) default_stroke else "4472C4"

  fill_col <- "F2F2F2"  # very light grey default

  label_safe <- xml_escape(if (!is.null(sg$label) && !is.na(sg$label)) sg$label else sg$id)

  # Metadata JSON
  meta <- jsonlite::toJSON(
    list(
      v     = "1",
      type  = "subgraph",
      id    = sg$id,
      label = if (!is.null(sg$label) && !is.na(sg$label)) sg$label else sg$id,
      nodes = as.list(nids)
    ),
    auto_unbox = TRUE
  )

  name_attr <- paste0("mermaid:subgraph:", sg$id)

  # Stroke width (1.5pt)
  sw_emu <- as.integer(round(1.5 * 12700))

  inner_xml <- paste0(
    '<wps:wsp>',
    '<wps:cNvPr id="', shape_id, '" name="', xml_escape(name_attr), '"/>',
    '<wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>',
    '<wps:spPr>',
    '<a:xfrm>',
    '<a:off x="', inches_to_emu(box_x), '" y="', inches_to_emu(box_y), '"/>',
    '<a:ext cx="', inches_to_emu(box_w), '" cy="', inches_to_emu(box_h), '"/>',
    '</a:xfrm>',
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
    '<a:solidFill><a:srgbClr val="', fill_col, '"/></a:solidFill>',
    '<a:ln w="', sw_emu, '">',
    '<a:solidFill><a:srgbClr val="', strk_col, '"/></a:solidFill>',
    '<a:prstDash val="dash"/>',
    '</a:ln>',
    '</wps:spPr>',
    '<wps:txbx>',
    '<w:txbxContent>',
    '<w:p>',
    '<w:pPr><w:jc w:val="left"/></w:pPr>',
    '<w:r>',
    '<w:rPr>',
    '<w:color w:val="', strk_col, '"/>',
    '<w:b/>',
    '<w:sz w:val="14"/><w:szCs w:val="14"/>',
    '</w:rPr>',
    '<w:t xml:space="preserve">', label_safe, '</w:t>',
    '</w:r>',
    '</w:p>',
    '</w:txbxContent>',
    '</wps:txbx>',
    '<wps:bodyPr anchor="t" anchorCtr="0" vert="horz">',
    '<a:noAutofit/>',
    '</wps:bodyPr>',
    '</wps:wsp>'
  )

  make_anchor(
    x_in      = box_x,
    y_in      = box_y,
    w_in      = box_w,
    h_in      = box_h,
    shape_id  = shape_id,
    name      = name_attr,
    descr_json = meta,
    z_order   = 2000000L,
    inner_xml = inner_xml
  )
}

# Main exported function -------------------------------------------------------

#' Build DrawingML XML for a laid-out mermaid_graph
#'
#' @param graph A laid-out `mermaid_graph` (from [calculate_layout()]).
#' @param start_id Integer starting shape ID. Increment between diagrams.
#' @param default_fill Default node fill colour (6-char hex). Default "FFFFFF".
#' @param default_stroke Default stroke colour. Default "4472C4".
#' @param default_text_color Default text colour. Default "000000".
#' @param stroke_width_pt Stroke width in points. Default 1.5.
#' @return Named list: `xml` (character) and `next_id` (integer).
#' @export
build_diagram_xml <- function(graph,
                               start_id           = 1L,
                               default_fill       = "FFFFFF",
                               default_stroke     = "4472C4",
                               default_text_color = "000000",
                               stroke_width_pt    = 1.5) {
  nodes     <- graph$nodes
  edges     <- graph$edges
  subgraphs <- graph$subgraphs
  shape_id  <- as.integer(start_id)
  parts     <- character(0)

  # 1. Subgraphs (z-order 2000000 - behind nodes)
  if (nrow(subgraphs) > 0) {
    for (i in seq_len(nrow(subgraphs))) {
      sg          <- as.list(subgraphs[i, ])
      sg$node_ids <- subgraphs$node_ids[[i]]
      xml_str     <- subgraph_to_xml(sg, nodes, shape_id, default_stroke)
      if (nzchar(xml_str)) {
        parts    <- c(parts, xml_str)
        shape_id <- shape_id + 1L
      }
    }
  }

  # 2. Nodes
  if (nrow(nodes) > 0) {
    for (i in seq_len(nrow(nodes))) {
      nd      <- as.list(nodes[i, ])
      xml_str <- node_to_xml(nd, shape_id, default_fill, default_stroke,
                              default_text_color, stroke_width_pt)
      parts    <- c(parts, xml_str)
      shape_id <- shape_id + 1L
    }
  }

  # 3. Edges + optional edge labels
  if (nrow(edges) > 0) {
    for (i in seq_len(nrow(edges))) {
      e       <- as.list(edges[i, ])
      xml_str <- edge_to_xml(e, shape_id, default_stroke, stroke_width_pt)
      if (nzchar(xml_str)) {
        parts    <- c(parts, xml_str)
        shape_id <- shape_id + 1L
      }
      if (!is.null(e$label) && !is.na(e$label) && nzchar(e$label)) {
        lxml <- edge_label_to_xml(e, shape_id, default_text_color)
        if (nzchar(lxml)) {
          parts    <- c(parts, lxml)
          shape_id <- shape_id + 1L
        }
      }
    }
  }

  xml <- paste0(
    '<w:p ', .dml_namespaces, '>',
    paste0(parts, collapse = ""),
    '</w:p>'
  )

  list(xml = xml, next_id = shape_id)
}

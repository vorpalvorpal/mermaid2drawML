# R/dml.R
# Pure DrawingML / WordprocessingML tag builders. Each XML construct is
# written in one place, and every attribute value goes through attr_value()
# and so through xml_escape(). Lengths given to these builders are already
# EMU; the conversion from scene px is in R/emit.R.
#
# Word is fussy about the wrapper (MS-ODRAWXML 2.16.1), and this markup is
# the one proven in Word by earlier versions of the package:
#   - the root wpg:wgp (directly in a:graphicData) has NO cNvPr;
#   - every nested wpg:grpSp MUST have one;
#   - element order: cNvPr, cNvGrpSpPr, grpSpPr, children;
#   - the drawing must sit in mc:AlternateContent / mc:Choice Requires="wpg",
#     or Word does not activate its wpg handler and nested groups vanish.

.dml_namespaces <- c(
  "xmlns:wpg" = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  "xmlns:wps" = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  "xmlns:wp"  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "xmlns:a"   = "http://schemas.openxmlformats.org/drawingml/2006/main",
  "xmlns:w"   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "xmlns:r"   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "xmlns:mc"  = "http://schemas.openxmlformats.org/markup-compatibility/2006"
)

# ---- generic element ---------------------------------------------------------

# Marks a shape id that R/emit.R numbers once the whole drawing exists. The
# character is one xml_escape() removes, so no escaped text can contain it.
.id_mark <- "\u0001"

#' An id attribute value to be numbered later (see number_ids())
#' @noRd
dml_id <- function() structure(.id_mark, class = "dml_id")

# Attribute values: integers for every number (EMU, half-points, angles; a
# double such as 1e5 would otherwise print in scientific notation), "1"/"0"
# for logicals, escaped text for the rest.
attr_value <- function(v) {
  if (inherits(v, "dml_id")) return(unclass(v))
  if (is.logical(v)) return(if (v) "1" else "0")
  if (is.numeric(v)) return(sprintf("%d", as.integer(round(v))))
  xml_escape(as.character(v))
}

#' One XML element
#'
#' @param .tag Tag name with its prefix (dotted, so that `name` can be an
#'   attribute).
#' @param ... Attributes as named arguments; a NULL or NA value leaves the
#'   attribute out.
#' @param .children Character vector of child markup, or NULL for an empty
#'   (self-closing) element.
#' @noRd
xml_tag <- function(.tag, ..., .children = NULL) {
  attrs <- purrr::discard(list(...), function(v) is.null(v) || (length(v) == 1L && is.na(v)))
  attr_str <- paste0(" ", names(attrs), "=\"", purrr::map_chr(attrs, attr_value), "\"",
                     collapse = "")
  if (length(attrs) == 0L) attr_str <- ""
  if (length(.children) == 0L) return(paste0("<", .tag, attr_str, "/>"))
  paste0("<", .tag, attr_str, ">", paste(.children, collapse = ""), "</", .tag, ">")
}

# ---- geometry ------------------------------------------------------------------

#' a:xfrm: offset and size in EMU, rotation in degrees clockwise
#'
#' `child = TRUE` adds a group's child offset and extent, equal to its own
#' size at origin 0: its children keep diagram coordinates relative to the
#' group's box.
#' @noRd
dml_xfrm <- function(x, y, w, h, rot = 0, flip_h = FALSE, flip_v = FALSE,
                     child = FALSE) {
  # DrawingML wants 60000ths of a degree in [0, 21600000).
  rot60k <- round(rot * 60000) %% 21600000
  xml_tag(
    "a:xfrm",
    rot = if (isTRUE(rot60k != 0)) rot60k,
    flipH = if (isTRUE(flip_h)) TRUE, flipV = if (isTRUE(flip_v)) TRUE,
    .children = c(
      xml_tag("a:off", x = x, y = y),
      xml_tag("a:ext", cx = w, cy = h),
      if (child) xml_tag("a:chOff", x = 0, y = 0),
      if (child) xml_tag("a:chExt", cx = w, cy = h)
    )
  )
}

#' a:prstGeom with its adjust values (a named numeric vector)
#' @noRd
dml_prst_geom <- function(prst, adj = numeric()) {
  gds <- purrr::imap_chr(as.list(adj), function(v, nm) {
    xml_tag("a:gd", name = nm, fmla = paste("val", sprintf("%d", as.integer(round(v)))))
  })
  xml_tag("a:prstGeom", prst = prst,
          .children = xml_tag("a:avLst", .children = unname(gds)))
}

# A point list, vectorised over its coordinates.
dml_pts <- function(x, y) sprintf("<a:pt x=\"%d\" y=\"%d\"/>", x, y)

#' a:custGeom with one a:path
#'
#' @param path Segments as in R/geometry.R (`cmd` M/L/C/Z, control points
#'   `x1`..`y2`, end point `x`, `y`), already in EMU relative to the box.
#' @param w,h The box in EMU (the path's coordinate space).
#' @noRd
dml_cust_geom <- function(path, w, h) {
  int <- function(v) as.integer(round(v))
  end <- dml_pts(int(path$x), int(path$y))
  # One column per command; each segment takes the column of its command.
  forms <- cbind(
    M = paste0("<a:moveTo>", end, "</a:moveTo>"),
    L = paste0("<a:lnTo>", end, "</a:lnTo>"),
    C = paste0("<a:cubicBezTo>", dml_pts(int(path$x1), int(path$y1)),
               dml_pts(int(path$x2), int(path$y2)), end, "</a:cubicBezTo>"),
    Z = rep("<a:close/>", nrow(path))
  )
  segs <- forms[cbind(seq_len(nrow(path)), match(path$cmd, colnames(forms)))]
  xml_tag(
    "a:custGeom",
    .children = c(
      "<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>",
      xml_tag("a:rect", l = "0", t = "0", r = "r", b = "b"),
      xml_tag("a:pathLst", .children = xml_tag("a:path", w = w, h = h, .children = segs))
    )
  )
}

# ---- paint ---------------------------------------------------------------------

#' a:solidFill (with a:alpha below 1) or a:noFill
#' @param hex "RRGGBB" or NA for none.
#' @noRd
dml_fill <- function(hex, alpha = 1) {
  if (is.na(hex)) return(xml_tag("a:noFill"))
  a <- if (!is.na(alpha) && alpha < 1) xml_tag("a:alpha", val = alpha * 100000)
  xml_tag("a:solidFill", .children = xml_tag("a:srgbClr", val = hex, .children = a))
}

#' a:ln: width in EMU, paint, preset dash, line ends
#'
#' @param hex NA draws no line.
#' @param dash A DrawingML preset dash name, or NA for solid.
#' @param head,tail Line-end types; "none" writes no end.
#' @param head_size,tail_size Each end's size classes c(w, len), "sm" |
#'   "med" | "lg".
#' @noRd
dml_ln <- function(width, hex, alpha = 1, dash = NA, head = "none", tail = "none",
                   head_size = c(w = "med", len = "med"),
                   tail_size = c(w = "med", len = "med")) {
  if (is.na(hex)) return(xml_tag("a:ln", .children = xml_tag("a:noFill")))
  line_end <- function(tag, type, size) {
    if (!identical(type, "none")) xml_tag(tag, type = type, w = size[["w"]], len = size[["len"]])
  }
  xml_tag("a:ln", w = width, .children = c(
    dml_fill(hex, alpha),
    if (!is.na(dash)) xml_tag("a:prstDash", val = dash),
    line_end("a:headEnd", head, head_size),
    line_end("a:tailEnd", tail, tail_size)
  ))
}

# ---- shapes and groups -----------------------------------------------------------

# Non-visual properties by role: closed shapes, connectors (lines and open
# paths), free text boxes.
.nv_props <- list(
  shape = function() xml_tag("wps:cNvSpPr", .children = xml_tag("a:spLocks", noChangeArrowheads = TRUE)),
  connector = function() xml_tag("wps:cNvCnPr"),
  text_box = function() xml_tag("wps:cNvSpPr", txBox = TRUE)
)

.body_anchors <- c(top = "t", center = "ctr", bottom = "b")

#' wps:bodyPr for a text body: no wrapping, insets in EMU c(l, t, r, b), and
#' the vertical anchor for the scene's `valign` ("top", "center", "bottom")
#' @noRd
dml_body_pr <- function(insets = c(0, 0, 0, 0), valign = "center") {
  xml_tag("wps:bodyPr", wrap = "none", lIns = insets[[1]], tIns = insets[[2]],
          rIns = insets[[3]], bIns = insets[[4]], anchor = .body_anchors[[valign]],
          .children = xml_tag("a:normAutofit"))
}

#' One wps:wsp
#'
#' @param name The shape's name (cNvPr).
#' @param role "shape", "connector" or "text_box".
#' @param sp_pr Children of wps:spPr: xfrm, geometry, fill, line.
#' @param txbx A wps:txbx element, or NULL.
#' @param body_pr A wps:bodyPr element (required by the schema, so a shape
#'   without text gets an empty one).
#' @noRd
dml_wsp <- function(name, role, sp_pr, txbx = NULL, body_pr = xml_tag("wps:bodyPr")) {
  xml_tag("wps:wsp", .children = c(
    xml_tag("wps:cNvPr", id = dml_id(), name = name),
    .nv_props[[role]](),
    xml_tag("wps:spPr", .children = sp_pr),
    txbx,
    body_pr
  ))
}

#' A nested group (wpg:grpSp, with its cNvPr) of `children` markup
#' @noRd
dml_grp <- function(name, x, y, w, h, children) {
  xml_tag("wpg:grpSp", .children = c(
    xml_tag("wpg:cNvPr", id = dml_id(), name = name),
    xml_tag("wpg:cNvGrpSpPr"),
    xml_tag("wpg:grpSpPr", .children = dml_xfrm(x, y, w, h, child = TRUE)),
    children
  ))
}

#' The root group (wpg:wgp, no cNvPr) of size w x h EMU
#' @noRd
dml_root_group <- function(w, h, children) {
  xml_tag("wpg:wgp", .children = c(
    xml_tag("wpg:cNvGrpSpPr"),
    xml_tag("wpg:grpSpPr", .children = dml_xfrm(0, 0, w, h, child = TRUE)),
    children
  ))
}

# ---- the in-line drawing and its paragraph -------------------------------------

#' The run holding an in-line drawing of `group`
#'
#' In line with the text (wp:inline), the drawing takes its own place in the
#' flow and can never overlap other paragraphs or another diagram.
#' @param descr,title The docPr alt text and title; NULL leaves one out.
#' @noRd
dml_inline <- function(w, h, group, descr = "", title = NULL) {
  inline <- xml_tag(
    "wp:inline", distT = 0, distB = 0, distL = 0, distR = 0,
    .children = c(
      xml_tag("wp:extent", cx = w, cy = h),
      xml_tag("wp:effectExtent", l = 0, t = 0, r = 0, b = 0),
      xml_tag("wp:docPr", id = dml_id(), name = "mermaid:diagram", descr = descr,
              title = title),
      xml_tag("wp:cNvGraphicFramePr"),
      xml_tag("a:graphic", .children = xml_tag(
        "a:graphicData",
        uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        .children = group
      ))
    )
  )
  xml_tag("w:r", .children = c(
    xml_tag("w:rPr", .children = xml_tag("w:noProof")),
    xml_tag("mc:AlternateContent", .children = c(
      xml_tag("mc:Choice", Requires = "wpg",
              .children = xml_tag("w:drawing", .children = inline)),
      xml_tag("mc:Fallback")
    ))
  ))
}

#' w:pPr with single line spacing and no space around, optional zero
#' indents, and the alignment
#'
#' Pinning these stops a template's Normal style (space after, exact or 1.5
#' line spacing) from clipping a drawing or overflowing a label.
#' @noRd
dml_ppr <- function(align, indents = TRUE) {
  xml_tag("w:pPr", .children = c(
    xml_tag("w:spacing", "w:before" = 0, "w:after" = 0, "w:line" = 240, "w:lineRule" = "auto"),
    if (indents) xml_tag("w:ind", "w:left" = 0, "w:right" = 0, "w:firstLine" = 0),
    xml_tag("w:jc", "w:val" = align)
  ))
}

#' One w:p of `content`; `namespaces = TRUE` declares the drawing namespaces
#' on it (the holder paragraph is inserted into the document on its own)
#' @noRd
dml_paragraph <- function(align, content, indents = TRUE, namespaces = FALSE) {
  ns <- if (namespaces) as.list(.dml_namespaces) else list()
  rlang::exec(xml_tag, "w:p", !!!ns, .children = c(dml_ppr(align, indents), content))
}

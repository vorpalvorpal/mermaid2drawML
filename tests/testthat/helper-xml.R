# Test helpers for the DrawingML: XPath with real namespace prefixes
# (xml2::xml_ns_strip() does not strip prefixed namespaces, and Word's XML is
# all prefixed), readers of shapes, text and the whole drawing, the schema's
# child order, and saved documents.

# -- Namespaces and XPath ------------------------------------------------------

dml_ns <- c(
  w   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  wp  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  wpg = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  a   = "http://schemas.openxmlformats.org/drawingml/2006/main",
  mc  = "http://schemas.openxmlformats.org/markup-compatibility/2006",
  r   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  pic = "http://schemas.openxmlformats.org/drawingml/2006/picture",
  w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
)

xfind <- function(x, xpath) xml2::xml_find_all(x, xpath, ns = dml_ns)
xfirst <- function(x, xpath) xml2::xml_find_first(x, xpath, ns = dml_ns)
xattr <- function(x, attr) xml2::xml_attr(x, attr, ns = dml_ns)
xnum <- function(x, attr) as.numeric(xattr(x, attr))
is_missing_node <- function(x) inherits(x, "xml_missing")
xml_missing_node <- function() xml2::xml_find_first(xml2::read_xml("<x/>"), "/nothing")

# A parsed document from an XML string, a build_diagram_xml() result or a
# mermaid_word_diagram (anything with $xml), or an already parsed document.
dml_doc <- function(x) {
  if (inherits(x, "xml_document") || inherits(x, "xml_node")) return(x)
  if (is.list(x)) x <- x$xml
  xml2::read_xml(x)
}

# -- Finding shapes by name ----------------------------------------------------

# All cNvPr elements (shapes and nested groups) with their names, in document
# order. Names are compared in R, so names holding quotes need no XPath escaping.
cnv_all <- function(doc) xfind(doc, "//wps:cNvPr | //wpg:cNvPr")
cnv_names <- function(doc) {
  n <- xml2::xml_attr(cnv_all(doc), "name")
  n[is.na(n)] <- ""
  n
}

# Every owner element (wps:wsp or wpg:grpSp) whose cNvPr name is exactly `name`.
owners_named <- function(doc, name) {
  cnv <- cnv_all(doc)
  xml2::xml_parent(cnv[cnv_names(doc) == name])
}

# The one element named `name` (the first if several), or an xml_missing.
shape_named <- function(doc, name) {
  hits <- owners_named(doc, name)
  if (length(hits) == 0L) xml_missing_node() else hits[[1]]
}

# Every element whose name starts with `prefix`.
shapes_named <- function(doc, prefix) {
  cnv <- cnv_all(doc)
  xml2::xml_parent(cnv[startsWith(cnv_names(doc), prefix)])
}

count_named <- function(doc, name) sum(cnv_names(doc) == name)

shape_name <- function(node) {
  xattr(xfirst(node, "./wps:cNvPr | ./wpg:cNvPr | ./pic:nvPicPr/pic:cNvPr"), "name")
}
is_group_node <- function(node) xml2::xml_name(node) %in% c("grpSp", "wgp")
is_wsp_node <- function(node) xml2::xml_name(node) == "wsp"

# Names of every named element below `node` (not including itself).
descendant_names <- function(node) {
  n <- xml2::xml_attr(xfind(node, ".//wps:cNvPr | .//wpg:cNvPr"), "name")
  setdiff(n, shape_name(node))
}

all_wsp <- function(doc) xfind(doc, "//wps:wsp")

# The one wps:wsp of a drawing; fails the test when there is not exactly one.
the_wsp <- function(doc) {
  s <- all_wsp(doc)
  testthat::expect_length(s, 1L)
  s[[1]]
}

# TRUE when every value is written as a whole number.
int_attr_ok <- function(v) all(grepl("^-?[0-9]+$", v))

# Pictures (pic:pic), in document order.
all_pics <- function(doc) xfind(doc, "//pic:pic")

# The r:embed of a picture's a:blip.
pic_embed <- function(node) xattr(xfirst(node, "./pic:blipFill/a:blip"), "r:embed")

# The a:hlinkClick r:id of a shape, group or picture, NA when it has none.
shape_link_id <- function(node) {
  xattr(xfirst(node, paste("./wps:cNvPr/a:hlinkClick | ./wpg:cNvPr/a:hlinkClick",
                           "| ./pic:nvPicPr/pic:cNvPr/a:hlinkClick")), "r:id")
}

# Every attribute in the relationships namespace (r:id, r:embed, ...) in a
# document: element, attribute and value, one row each.
r_attributes <- function(doc) {
  at <- xml2::xml_find_all(doc, sprintf("//@*[namespace-uri()='%s']", dml_ns[["r"]]))
  tibble::tibble(element = xml2::xml_name(xml2::xml_parent(at)),
                 attr = xml2::xml_name(at), value = xml2::xml_text(at))
}

# -- Reading one shape ---------------------------------------------------------

# Fill or line paint: the solidFill / noFill child of `el` (an spPr or a:ln).
paint_of <- function(el) {
  solid <- xfirst(el, "./a:solidFill/a:srgbClr")
  none <- xfirst(el, "./a:noFill")
  if (!is_missing_node(solid)) {
    list(kind = "solid", color = xattr(solid, "val"),
         alpha = xnum(xfirst(solid, "./a:alpha"), "val"))
  } else if (!is_missing_node(none)) {
    list(kind = "none", color = NA_character_, alpha = NA_real_)
  } else {
    list(kind = NA_character_, color = NA_character_, alpha = NA_real_)
  }
}

# The a:gradFill child of `el` (an spPr or a:ln), read out: `stops` (pos
# 0-100000, color, alpha with NA when not written), `ang` and `scaled` of
# a:lin, `path` of a:path, `fill_to_rect` c(l, t, r, b). NULL when `el` has
# no gradient fill.
grad_of <- function(el) {
  g <- xfirst(el, "./a:gradFill")
  if (is_missing_node(g)) return(NULL)
  gs <- xfind(g, "./a:gsLst/a:gs")
  clr <- purrr::map(as.list(gs), function(s) xfirst(s, "./a:srgbClr"))
  lin <- xfirst(g, "./a:lin")
  path <- xfirst(g, "./a:path")
  ftr <- xfirst(path, "./a:fillToRect")
  list(
    stops = tibble::tibble(
      pos = xnum(gs, "pos"),
      color = purrr::map_chr(clr, xattr, "val"),
      alpha = purrr::map_dbl(clr, function(c) xnum(xfirst(c, "./a:alpha"), "val"))
    ),
    ang = xnum(lin, "ang"), scaled = xattr(lin, "scaled"),
    path = xattr(path, "path"),
    fill_to_rect = c(l = xnum(ftr, "l"), t = xnum(ftr, "t"), r = xnum(ftr, "r"),
                     b = xnum(ftr, "b"))
  )
}

# The a:pattFill child of `el`: preset, foreground colour, and the
# background colour and alpha (NA when there is no a:bgClr). NULL when none.
patt_of <- function(el) {
  p <- xfirst(el, "./a:pattFill")
  if (is_missing_node(p)) return(NULL)
  fg <- xfirst(p, "./a:fgClr/a:srgbClr")
  bg <- xfirst(p, "./a:bgClr/a:srgbClr")
  list(prst = xattr(p, "prst"), fg = xattr(fg, "val"),
       fg_alpha = xnum(xfirst(fg, "./a:alpha"), "val"),
       has_bg = !is_missing_node(xfirst(p, "./a:bgClr")),
       bg = xattr(bg, "val"), bg_alpha = xnum(xfirst(bg, "./a:alpha"), "val"))
}

shape_sppr <- function(node) xfirst(node, "./wps:spPr")

shape_fill <- function(node) paint_of(shape_sppr(node))

end_type <- function(el) {
  t <- xattr(el, "type")
  if (is.na(t)) "none" else t
}

shape_ln <- function(node) {
  ln <- xfirst(node, "./wps:spPr/a:ln")
  head <- xfirst(ln, "./a:headEnd")
  tail <- xfirst(ln, "./a:tailEnd")
  c(list(present = !is_missing_node(ln), w = xnum(ln, "w")),
    paint_of(ln),
    list(dash = xattr(xfirst(ln, "./a:prstDash"), "val"),
         head = end_type(head), head_w = xattr(head, "w"), head_len = xattr(head, "len"),
         tail = end_type(tail), tail_w = xattr(tail, "w"), tail_len = xattr(tail, "len")))
}

shape_prst <- function(node) xattr(xfirst(node, "./wps:spPr/a:prstGeom"), "prst")

shape_is_custom <- function(node) !is_missing_node(xfirst(node, "./wps:spPr/a:custGeom"))

# Adjust values of a preset as a named numeric vector (name -> value).
shape_adj <- function(node) {
  gds <- xfind(node, "./wps:spPr/a:prstGeom/a:avLst/a:gd")
  stats::setNames(as.numeric(sub("^val\\s+", "", xattr(gds, "fmla"))), xattr(gds, "name"))
}

# The a:xfrm of a shape (wps:spPr) or group (wpg:grpSpPr).
shape_xfrm <- function(node) {
  xf <- xfirst(node, "./wps:spPr/a:xfrm | ./wpg:grpSpPr/a:xfrm | ./pic:spPr/a:xfrm")
  off <- xfirst(xf, "./a:off")
  ext <- xfirst(xf, "./a:ext")
  cho <- xfirst(xf, "./a:chOff")
  che <- xfirst(xf, "./a:chExt")
  rot <- xnum(xf, "rot")
  list(
    x = xnum(off, "x"), y = xnum(off, "y"),
    cx = xnum(ext, "cx"), cy = xnum(ext, "cy"),
    rot = if (is.na(rot)) 0 else rot,
    flipH = xattr(xf, "flipH") %in% c("1", "true"),
    flipV = xattr(xf, "flipV") %in% c("1", "true"),
    chx = xnum(cho, "x"), chy = xnum(cho, "y"),
    chcx = xnum(che, "cx"), chcy = xnum(che, "cy"),
    raw = list(x = xattr(off, "x"), y = xattr(off, "y"),
               cx = xattr(ext, "cx"), cy = xattr(ext, "cy"))
  )
}

shape_ext <- function(node) {
  xf <- shape_xfrm(node)
  c(cx = xf$cx, cy = xf$cy)
}

shape_off <- function(node) {
  xf <- shape_xfrm(node)
  c(x = xf$x, y = xf$y)
}

# The box of a shape or group in root-group EMU (unrotated), mapping through
# every enclosing group: parent = off + (child - chOff) * ext / chExt.
to_parent_space <- function(box, group) {
  g <- shape_xfrm(group)
  chx <- if (is.na(g$chx)) 0 else g$chx
  chy <- if (is.na(g$chy)) 0 else g$chy
  sx <- if (isTRUE(g$chcx > 0)) g$cx / g$chcx else 1
  sy <- if (isTRUE(g$chcy > 0)) g$cy / g$chcy else 1
  c(x = g$x + (box[["x"]] - chx) * sx, y = g$y + (box[["y"]] - chy) * sy,
    w = box[["w"]] * sx, h = box[["h"]] * sy)
}

abs_box <- function(node) {
  xf <- shape_xfrm(node)
  groups <- xfind(node, "ancestor::wpg:grpSp | ancestor::wpg:wgp")
  # Document order is outermost first; map from the innermost group outwards.
  purrr::reduce(rev(as.list(groups)), to_parent_space,
                .init = c(x = xf$x, y = xf$y, w = xf$cx, h = xf$cy))
}

# The axis-aligned footprint of a box rotated by `rot_60k` (60000ths of a
# degree) about its centre.
rotated_footprint <- function(box, rot_60k) {
  r <- rot_60k / 60000 * pi / 180
  w <- abs(box[["w"]] * cos(r)) + abs(box[["h"]] * sin(r))
  h <- abs(box[["w"]] * sin(r)) + abs(box[["h"]] * cos(r))
  cx <- box[["x"]] + box[["w"]] / 2
  cy <- box[["y"]] + box[["h"]] / 2
  c(x = cx - w / 2, y = cy - h / 2, w = w, h = h)
}

body_pr <- function(node) {
  bp <- xfirst(node, "./wps:bodyPr")
  if (is_missing_node(bp)) return(character())
  xml2::xml_attrs(bp)
}

# Text paragraphs of a shape's own text body (or of every shape below a group).
shape_paragraphs <- function(node) xfind(node, ".//w:txbxContent/w:p")

paragraph_text <- function(p) paste(xml2::xml_text(xfind(p, ".//w:t")), collapse = "")

shape_paragraph_texts <- function(node) {
  vapply(as.list(shape_paragraphs(node)), paragraph_text, character(1))
}

# Every txbxContent paragraph in the document, as text.
all_paragraph_texts <- function(doc) {
  vapply(as.list(xfind(doc, "//w:txbxContent/w:p")), paragraph_text, character(1))
}

paragraph_jc <- function(p) xattr(xfirst(p, "./w:pPr/w:jc"), "w:val")

# One row per run of a paragraph node set.
run_rows <- function(paras) {
  purrr::list_rbind(purrr::imap(as.list(paras), function(p, i) {
    purrr::list_rbind(purrr::map(as.list(xfind(p, "./w:r")), function(r) {
      rpr <- xfirst(r, "./w:rPr")
      has <- function(tag) !is_missing_node(xfirst(rpr, paste0("./w:", tag)))
      fonts <- xfirst(rpr, "./w:rFonts")
      tf <- xfirst(rpr, "./w14:textFill")
      tf_clr <- xfirst(tf, "./w14:solidFill/w14:srgbClr")
      tibble::tibble(
        para = i,
        text = paste(xml2::xml_text(xfind(r, "./w:t")), collapse = ""),
        bold = has("b"), italic = has("i"), underline = has("u"),
        color = xattr(xfirst(rpr, "./w:color"), "w:val"),
        sz = xnum(xfirst(rpr, "./w:sz"), "w:val"),
        szCs = xnum(xfirst(rpr, "./w:szCs"), "w:val"),
        ascii = xattr(fonts, "w:ascii"), hAnsi = xattr(fonts, "w:hAnsi"),
        eastAsia = xattr(fonts, "w:eastAsia"), cs = xattr(fonts, "w:cs"),
        # Word 2010's text fill: its colour and w14:alpha (a transparency).
        text_fill = !is_missing_node(tf),
        text_fill_color = xattr(tf_clr, "w14:val"),
        text_fill_alpha = xnum(xfirst(tf_clr, "./w14:alpha"), "w14:val")
      )
    }))
  }))
}

shape_runs <- function(node) run_rows(shape_paragraphs(node))

# Paragraph properties a text paragraph must carry: single spacing, no indents.
expect_reset_ppr <- function(p, info = NULL) {
  sp <- xfirst(p, "./w:pPr/w:spacing")
  ind <- xfirst(p, "./w:pPr/w:ind")
  testthat::expect_false(is_missing_node(sp), info = info)
  testthat::expect_false(is_missing_node(ind), info = info)
  testthat::expect_false(is_missing_node(xfirst(p, "./w:pPr/w:jc")), info = info)
  testthat::expect_equal(xattr(sp, "w:before"), "0", info = info)
  testthat::expect_equal(xattr(sp, "w:after"), "0", info = info)
  testthat::expect_equal(xattr(sp, "w:line"), "240", info = info)
  testthat::expect_equal(xattr(sp, "w:lineRule"), "auto", info = info)
  testthat::expect_equal(xattr(ind, "w:left"), "0", info = info)
  testthat::expect_equal(xattr(ind, "w:right"), "0", info = info)
  testthat::expect_equal(xattr(ind, "w:firstLine"), "0", info = info)
}

# -- Whole drawing -------------------------------------------------------------

# Every id in document order: wp:docPr first, then each cNvPr.
all_ids <- function(doc) {
  as.integer(xattr(xfind(doc, "//wp:docPr | //wps:cNvPr | //wpg:cNvPr | //pic:cNvPr"), "id"))
}

doc_pr <- function(doc) xfirst(doc, "//wp:inline/wp:docPr")

inline_extent <- function(doc) {
  e <- xfirst(doc, "//wp:inline/wp:extent")
  c(cx = xnum(e, "cx"), cy = xnum(e, "cy"))
}

root_group <- function(doc) xfirst(doc, "//a:graphicData/wpg:wgp")

root_ext <- function(doc) {
  e <- xfirst(doc, "//wpg:wgp/wpg:grpSpPr/a:xfrm/a:ext")
  c(cx = xnum(e, "cx"), cy = xnum(e, "cy"))
}

# -- Schema child order --------------------------------------------------------
#
# The order the ECMA-376 schemas (and Microsoft's wps/wpg extensions) require
# for the children of the elements this package writes: element -> its
# children's ranks in order. "a|b" is a choice: its members share a rank
# (fills, geometries, a group's members, a path's segments), so they may
# repeat or alternate among themselves. It covers gradient and pattern fills,
# hyperlinks, pictures and text fills as well as plain shapes. Optional
# children Word never needs are listed where they would go, to place what
# follows them.

.fill_choice <- "a:noFill|a:solidFill|a:gradFill|a:blipFill|a:pattFill|a:grpFill"
.line_fill_choice <- "a:noFill|a:solidFill|a:gradFill|a:pattFill"
.effect_choice <- "a:effectLst|a:effectDag"
.group_members <- "wps:wsp|wpg:grpSp|wpg:graphicFrame|pic:pic|w14:contentPart"

.rpr_base <- c(
  "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps",
  "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint",
  "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden", "w:color", "w:spacing",
  "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
  "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
  "w:eastAsianLayout", "w:specVanish", "w:oMath",
  # Word 2010's run effects follow the standard properties.
  "w14:glow", "w14:shadow", "w14:reflection", "w14:textOutline", "w14:textFill",
  "w14:scene3d", "w14:props3d", "w14:ligatures", "w14:numForm", "w14:numSpacing",
  "w14:stylisticSets", "w14:cntxtAlts",
  "w:rPrChange"
)

schema_child_order <- list(
  "wp:inline" = c("wp:extent", "wp:effectExtent", "wp:docPr", "wp:cNvGraphicFramePr",
                  "a:graphic"),
  "wpg:wgp" = c("wpg:cNvPr", "wpg:cNvGrpSpPr", "wpg:grpSpPr", .group_members,
                "wpg:extLst"),
  "wpg:grpSp" = c("wpg:cNvPr", "wpg:cNvGrpSpPr", "wpg:grpSpPr", .group_members,
                  "wpg:extLst"),
  "wps:wsp" = c("wps:cNvPr", "wps:cNvSpPr|wps:cNvCnPr", "wps:spPr", "wps:style",
                "wps:extLst", "wps:txbx|wps:linkedTxbx", "wps:bodyPr"),
  "wps:cNvPr" = c("a:hlinkClick", "a:hlinkHover", "a:extLst"),
  "wpg:cNvPr" = c("a:hlinkClick", "a:hlinkHover", "a:extLst"),
  "wps:spPr" = c("a:xfrm", "a:custGeom|a:prstGeom", .fill_choice, "a:ln", .effect_choice,
                 "a:scene3d", "a:sp3d", "a:extLst"),
  "wpg:grpSpPr" = c("a:xfrm", .fill_choice, .effect_choice, "a:scene3d", "a:extLst"),
  "a:xfrm" = c("a:off", "a:ext", "a:chOff", "a:chExt"),
  "a:ln" = c(.line_fill_choice, "a:prstDash|a:custDash", "a:round|a:bevel|a:miter",
             "a:headEnd", "a:tailEnd", "a:extLst"),
  "wps:bodyPr" = c("a:prstTxWarp", "a:noAutofit|a:normAutofit|a:spAutoFit", "a:scene3d",
                   "a:sp3d|a:flatTx", "a:extLst"),
  "w:pPr" = c(
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs",
    "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
    "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
    "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
    "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
    "w:pPrChange"
  ),
  # A paragraph mark's run properties may open with its revision marks.
  "w:rPr" = c("w:ins", "w:del", "w:moveFrom", "w:moveTo", .rpr_base),
  "a:custGeom" = c("a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst", "a:rect", "a:pathLst"),
  # Two elements share this name: a custom geometry's path (segments) and a
  # radial gradient's shade path (its focus rectangle). Neither ever holds
  # the other's children.
  "a:path" = "a:close|a:moveTo|a:lnTo|a:arcTo|a:quadBezTo|a:cubicBezTo|a:fillToRect",
  "a:gradFill" = c("a:gsLst", "a:lin|a:path", "a:tileRect"),
  "a:pattFill" = c("a:fgClr", "a:bgClr"),
  "pic:pic" = c("pic:nvPicPr", "pic:blipFill", "pic:spPr"),
  "pic:nvPicPr" = c("pic:cNvPr", "pic:cNvPicPr"),
  "pic:cNvPr" = c("a:hlinkClick", "a:hlinkHover", "a:extLst"),
  "pic:blipFill" = c("a:blip", "a:srcRect", "a:tile|a:stretch"),
  "pic:spPr" = c("a:xfrm", "a:custGeom|a:prstGeom", .fill_choice, "a:ln", .effect_choice,
                 "a:scene3d", "a:sp3d", "a:extLst")
)

# Namespace URIs by the prefix the table uses.
schema_ns <- dml_ns

# An element's name with the table's prefix for its namespace ("?:name" for
# a namespace the table does not know).
schema_name <- function(node) {
  uri <- xml2::xml_find_chr(node, "namespace-uri(.)")
  prefix <- names(schema_ns)[match(uri, schema_ns)]
  paste0(if (is.na(prefix)) "?" else prefix, ":", sub("^.*:", "", xml2::xml_name(node)))
}

# Problems with the order of `parent`'s element children, as text naming
# the parent and the child: a child the table does not allow there, or one
# that comes before a child it must follow.
child_order_problems <- function(parent, order) {
  ranks <- strsplit(order, "|", fixed = TRUE)
  rank_of <- stats::setNames(rep(seq_along(ranks), lengths(ranks)), unlist(ranks))
  pname <- schema_name(parent)
  kids <- purrr::map_chr(as.list(xml2::xml_children(parent)), schema_name)
  r <- unname(rank_of[kids])
  unknown <- kids[is.na(r)]
  out <- sprintf("%s: unexpected child %s", pname, unique(unknown))
  known <- r[!is.na(r)]
  back <- which(diff(known) < 0)
  c(out, sprintf("%s: %s comes after %s", pname, kids[!is.na(r)][back + 1L],
                 kids[!is.na(r)][back]))
}

# Every order problem in a document, for every element the table lists.
schema_order_problems <- function(doc) {
  unique(unlist(purrr::imap(schema_child_order, function(order, el) {
    parts <- strsplit(el, ":", fixed = TRUE)[[1]]
    xpath <- sprintf("//*[local-name()='%s' and namespace-uri()='%s']", parts[[2]],
                     schema_ns[[parts[[1]]]])
    purrr::map(as.list(xml2::xml_find_all(doc, xpath)), child_order_problems, order = order)
  })))
}

# -- Saved documents -------------------------------------------------------------

# Save an rdocx and read back its parts: `doc` (word/document.xml, parsed),
# `rels` (word/_rels/document.xml.rels as a tibble of id, type, target,
# target_mode) and `dir`, where the package was unzipped.
saved_docx <- function(rdocx, env = parent.frame()) {
  path <- withr::local_tempfile(fileext = ".docx", .local_envir = env)
  print(rdocx, target = path)
  dir <- withr::local_tempdir(.local_envir = env)
  utils::unzip(path, exdir = dir)
  rels_xml <- xml2::read_xml(file.path(dir, "word", "_rels", "document.xml.rels"))
  rel <- xml2::xml_children(rels_xml)
  list(
    doc = xml2::read_xml(file.path(dir, "word", "document.xml")),
    rels = tibble::tibble(id = xml2::xml_attr(rel, "Id"), type = xml2::xml_attr(rel, "Type"),
                          target = xml2::xml_attr(rel, "Target"),
                          target_mode = xml2::xml_attr(rel, "TargetMode")),
    dir = dir
  )
}

rel_type_image <- "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
rel_type_hyperlink <- "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

# The bytes of the part a relationship of word/document.xml targets.
rel_part_bytes <- function(pkg, target) {
  path <- file.path(pkg$dir, "word", target)
  if (!file.exists(path)) return(NULL)
  readBin(path, "raw", file.size(path))
}

# Problems with the r:id / r:embed attributes of a saved document: a value
# with no relationship, an image relationship whose part is missing or not a
# PNG, a hyperlink that is not external. Character, one line per problem.
relationship_problems <- function(pkg) {
  at <- r_attributes(pkg$doc)
  rel <- pkg$rels[match(at$value, pkg$rels$id), ]
  missing <- is.na(rel$id)
  out <- sprintf("%s r:%s=\"%s\" has no relationship", at$element[missing], at$attr[missing],
                 at$value[missing])
  img <- !missing & at$attr == "embed"
  for (k in which(img)) {
    if (!identical(rel$type[[k]], rel_type_image)) {
      out <- c(out, sprintf("r:embed=\"%s\" is a %s relationship", at$value[[k]], rel$type[[k]]))
    } else if (!is_png(rel_part_bytes(pkg, rel$target[[k]]))) {
      out <- c(out, sprintf("r:embed=\"%s\" targets %s, not a PNG part", at$value[[k]],
                            rel$target[[k]]))
    }
  }
  link <- !missing & at$element == "hlinkClick"
  for (k in which(link)) {
    if (!identical(rel$type[[k]], rel_type_hyperlink) ||
        !identical(rel$target_mode[[k]], "External")) {
      out <- c(out, sprintf("hlinkClick r:id=\"%s\" is not an external hyperlink", at$value[[k]]))
    }
  }
  out
}

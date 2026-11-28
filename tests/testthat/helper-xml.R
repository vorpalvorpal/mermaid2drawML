# Helpers for the scene and DrawingML tests (test-scene.R, test-semantics.R,
# test-dml_text.R, test-emit.R, test-word.R):
#   - XPath with real namespace prefixes (xml2::xml_ns_strip() does not strip
#     prefixed namespaces, and Word's XML is all prefixed);
#   - cached scenes and XML for the bridge fixtures;
#   - hand-built renders (mini_render() and prim_*()) for exact arithmetic.
# fixture_render() and local_mock_render() live in helper-render.R.

# -- Namespaces and XPath ------------------------------------------------------

dml_ns <- c(
  w   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  wp  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  wpg = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  a   = "http://schemas.openxmlformats.org/drawingml/2006/main",
  mc  = "http://schemas.openxmlformats.org/markup-compatibility/2006",
  r   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
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

shape_name <- function(node) xattr(xfirst(node, "./wps:cNvPr | ./wpg:cNvPr"), "name")
is_group_node <- function(node) xml2::xml_name(node) %in% c("grpSp", "wgp")
is_wsp_node <- function(node) xml2::xml_name(node) == "wsp"

# Names of every named element below `node` (not including itself).
descendant_names <- function(node) {
  n <- xml2::xml_attr(xfind(node, ".//wps:cNvPr | .//wpg:cNvPr"), "name")
  setdiff(n, shape_name(node))
}

all_wsp <- function(doc) xfind(doc, "//wps:wsp")

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
  xf <- xfirst(node, "./wps:spPr/a:xfrm | ./wpg:grpSpPr/a:xfrm")
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
      tibble::tibble(
        para = i,
        text = paste(xml2::xml_text(xfind(r, "./w:t")), collapse = ""),
        bold = has("b"), italic = has("i"), underline = has("u"),
        color = xattr(xfirst(rpr, "./w:color"), "w:val"),
        sz = xnum(xfirst(rpr, "./w:sz"), "w:val"),
        szCs = xnum(xfirst(rpr, "./w:szCs"), "w:val"),
        ascii = xattr(fonts, "w:ascii"), hAnsi = xattr(fonts, "w:hAnsi"),
        eastAsia = xattr(fonts, "w:eastAsia"), cs = xattr(fonts, "w:cs")
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
  as.integer(xattr(xfind(doc, "//wp:docPr | //wps:cNvPr | //wpg:cNvPr"), "id"))
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
# repeat or alternate among themselves. It lists children the package does
# not write yet (gradient and pattern fills, hyperlinks, pictures, text
# fills) so that writing them needs no change here. Optional children Word
# never needs are listed where they would go, to place what follows them.

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
  "a:path" = "a:close|a:moveTo|a:lnTo|a:arcTo|a:quadBezTo|a:cubicBezTo"
)

# Namespace URIs by the prefix the table uses.
schema_ns <- c(
  dml_ns,
  pic = "http://schemas.openxmlformats.org/drawingml/2006/picture",
  w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
)

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

# -- Expectations --------------------------------------------------------------

# Every element of `object` within `tol` of `expected` (absolute, unlike the
# relative tolerance of expect_equal()).
expect_near <- function(object, expected, tol, info = NULL) {
  act <- as.numeric(unlist(object))
  exp <- as.numeric(unlist(expected))
  testthat::expect_equal(length(act), length(exp), info = info)
  diff <- if (length(act) == length(exp) && length(act) > 0L) max(abs(act - exp)) else Inf
  testthat::expect_lte(
    diff, tol,
    label = sprintf("largest difference between c(%s) and c(%s)",
                    paste(signif(act, 9), collapse = ", "),
                    paste(signif(exp, 9), collapse = ", ")))
}

# -- Cached fixtures -----------------------------------------------------------

.dml_cache <- rlang::env()

all_fixture_names <- function() {
  sub("\\.json$", "", list.files(testthat::test_path("fixtures"), pattern = "\\.json$"))
}

# Evaluate `expr`, muffling only the warnings a real diagram is expected to
# raise: text below 7 pt on a large diagram, and unsupported prims (images).
quietly_expected <- function(expr) {
  withCallingHandlers(
    expr,
    mermaid2drawml_small_text = function(cnd) rlang::cnd_muffle(cnd),
    mermaid2drawml_unsupported = function(cnd) rlang::cnd_muffle(cnd)
  )
}

cached <- function(key, make) {
  if (is.null(.dml_cache[[key]])) .dml_cache[[key]] <- make()
  .dml_cache[[key]]
}

fixture_scene <- function(name) {
  cached(paste0("scene:", name), function() quietly_expected(build_scene(fixture_render(name))))
}

# build_diagram_xml() on a fixture (list(xml, next_id)), cached per arguments.
fixture_xml <- function(name, ...) {
  args <- list(...)
  cached(paste0("xml:", name, ":", rlang::hash(args)), function() {
    quietly_expected(rlang::exec(build_diagram_xml, fixture_render(name), !!!args))
  })
}

fixture_doc <- function(name, ...) {
  args <- list(...)
  cached(paste0("doc:", name, ":", rlang::hash(args)), function() {
    dml_doc(rlang::exec(fixture_xml, name, !!!args)$xml)
  })
}

# The bridge prims of one kind from a fixture.
fixture_prims <- function(name, kind = NULL) {
  p <- fixture_render(name)$prims
  if (is.null(kind)) p else purrr::keep(p, function(x) identical(x$kind, kind))
}

# The text of a text prim's rows (runs concatenated), one string per row.
prim_row_texts <- function(p) {
  vapply(p$rows, function(row) paste(vapply(row$runs, function(r) r$t, ""), collapse = ""),
         character(1))
}

# -- Scene items ---------------------------------------------------------------

item_named <- function(items, name) items[!is.na(items$name) & items$name == name, ]

# The names of group `name` and of every group nested in it.
subtree_groups <- function(items, name) {
  kids <- items$name[!is.na(items$parent) & items$parent == name &
                       items$kind == "group" & !is.na(items$name)]
  c(name, unlist(purrr::map(kids, function(k) subtree_groups(items, k))))
}

# The item called `name` and every item below it (following `parent`; the
# children need not have names of their own), in paint order.
subtree_rows <- function(items, name) {
  groups <- subtree_groups(items, name)
  items[(!is.na(items$name) & items$name == name) |
          (!is.na(items$parent) & items$parent %in% groups), ]
}

# -- Hand-built renders --------------------------------------------------------

`%|null|%` <- function(x, y) if (is.null(x)) y else x

# A JSON object (`{}`) even when empty.
json_obj <- function(x = list()) {
  if (length(x) == 0L) structure(list(), names = character()) else x
}

# Map points through an SVG matrix c(a, b, c, d, e, f).
apply_m <- function(m, x, y) {
  list(x = m[1] * x + m[3] * y + m[5], y = m[2] * x + m[4] * y + m[6])
}

bbox_of_points <- function(m, x, y) {
  p <- apply_m(m, x, y)
  c(min(p$x), min(p$y), max(p$x) - min(p$x), max(p$y) - min(p$y))
}

# Every field the bridge reports for a shape prim (bridge header), with
# defaults: white fill, black 1 px stroke, identity matrix.
prim_shape <- function(kind, tag, geom, bbox, m = c(1, 0, 0, 1, 0, 0),
                       fill = "#FFFFFF", fill_a = 1, stroke = "#000000",
                       stroke_a = 1, stroke_w = 1, dash = NULL,
                       marker_start = NULL, marker_end = NULL,
                       id = "", cls = "", data = list(), anc = list(),
                       html = NULL, n = NULL) {
  p <- list(
    kind = kind, tag = tag, id = id, cls = cls, data = json_obj(data),
    anc = anc, bbox = as.list(bbox), m = as.list(m), geom = geom,
    fill = fill, fill_a = if (is.null(fill)) 0 else fill_a,
    fill_approx = FALSE,
    stroke = stroke, stroke_a = if (is.null(stroke)) 0 else stroke_a,
    stroke_w = if (is.null(stroke)) 0 else stroke_w,
    dash = if (is.null(dash)) NULL else as.list(dash),
    marker_start = marker_start, marker_end = marker_end
  )
  if (!is.null(html)) p$html <- html
  if (!is.null(n)) p$n <- as.integer(n)
  p
}

prim_rect <- function(x, y, w, h, rx = 0, ry = rx, m = c(1, 0, 0, 1, 0, 0),
                      bbox = NULL, tag = "rect", ...) {
  bbox <- bbox %|null|% bbox_of_points(m, c(x, x + w, x, x + w), c(y, y, y + h, y + h))
  prim_shape("rect", tag, list(x = x, y = y, w = w, h = h, rx = rx, ry = ry),
             bbox = bbox, m = m, ...)
}

prim_ellipse <- function(cx, cy, rx, ry = rx, m = c(1, 0, 0, 1, 0, 0),
                         bbox = NULL, ...) {
  bbox <- bbox %|null|% bbox_of_points(m, c(cx - rx, cx + rx), c(cy - ry, cy + ry))
  prim_shape("ellipse", "ellipse", list(cx = cx, cy = cy, rx = rx, ry = ry),
             bbox = bbox, m = m, ...)
}

prim_line <- function(x1, y1, x2, y2, m = c(1, 0, 0, 1, 0, 0), bbox = NULL,
                      fill = NULL, ...) {
  bbox <- bbox %|null|% bbox_of_points(m, c(x1, x2), c(y1, y2))
  prim_shape("line", "line", list(x1 = x1, y1 = y1, x2 = x2, y2 = y2),
             bbox = bbox, m = m, fill = fill, ...)
}

# `points`: a list of c(x, y) pairs.
prim_poly <- function(points, closed = TRUE, m = c(1, 0, 0, 1, 0, 0),
                      bbox = NULL, ...) {
  xs <- vapply(points, `[[`, numeric(1), 1)
  ys <- vapply(points, `[[`, numeric(1), 2)
  bbox <- bbox %|null|% bbox_of_points(m, xs, ys)
  prim_shape("poly", if (closed) "polygon" else "polyline",
             list(points = purrr::map(points, as.list), closed = closed),
             bbox = bbox, m = m, ...)
}

# `bbox` defaults to the extent of the path's coordinate pairs, which is only
# right for absolute paths made of straight segments; give it for curves.
prim_path <- function(d, bbox = NULL, m = c(1, 0, 0, 1, 0, 0), ...) {
  if (is.null(bbox)) {
    nums <- as.numeric(regmatches(d, gregexpr("-?[0-9.]+(e-?[0-9]+)?", d))[[1]])
    xs <- nums[c(TRUE, FALSE)]
    ys <- nums[c(FALSE, TRUE)]
    bbox <- bbox_of_points(m, xs, ys)
  }
  prim_shape("path", "path", list(d = d), bbox = bbox, m = m, ...)
}

prim_image <- function(x, y, w, h, href = "https://example.com/x.png", anc = list()) {
  list(kind = "image", tag = "image", id = "", cls = "", data = json_obj(),
       anc = anc, bbox = list(x, y, w, h), href = href,
       m = list(1, 0, 0, 1, x, y))
}

prim_other <- function(tag, x, y, w, h, anc = list()) {
  list(kind = "other", tag = tag, id = "", cls = "", data = json_obj(),
       anc = anc, bbox = list(x, y, w, h))
}

# One text run as the bridge reports it.
text_run <- function(t, bold = FALSE, italic = FALSE, underline = FALSE,
                     color = "#000000", size = 16, font = "arial",
                     family = "arial, sans-serif") {
  list(t = t, bold = bold, italic = italic, underline = underline,
       color = color, size = size, family = family, font = font)
}

# A text prim. `text`: one string per row (one run each, styled by `...`, which
# goes to text_run()); or give `rows` as a list of lists of text_run()s. The
# rows split the box height evenly.
prim_text <- function(text = NULL, x, y, w, h, anchor = "middle", rot = 0,
                      rows = NULL, html = TRUE, anc = list(), id = "", cls = "",
                      n = NULL, ...) {
  rows <- rows %|null|% purrr::map(text, function(t) list(text_run(t, ...)))
  nr <- length(rows)
  rh <- h / nr
  p <- list(
    kind = "text", tag = if (html) "foreignobject" else "text",
    id = id, cls = cls, data = json_obj(), anc = anc,
    bbox = list(x, y, w, h), html = html, anchor = anchor, rot = rot,
    rows = purrr::imap(rows, function(runs, i) {
      list(bbox = list(x, y + (i - 1) * rh, w, rh), runs = runs)
    })
  )
  if (!is.null(n)) p$n <- as.integer(n)
  p
}

# A bridge success object holding `prims`, through the validating constructor.
# `markers`: a named list of marker_def()s, keyed by marker id (bridge 3).
mini_render <- function(prims = list(), viewBox = c(0, 0, 200, 100), db = NULL,
                        type = "flowchart-v2", title = NULL, desc = NULL,
                        markers = NULL) {
  prims <- purrr::imap(unname(prims), function(p, i) c(list(i = i - 1L), p))
  raw <- list(
    ok = TRUE, bridge = if (is.null(markers)) 2L else 3L, mermaid = "11.13.0",
    type = type, role = type,
    viewBox = as.list(viewBox), title = title, desc = desc,
    svg = "<svg xmlns=\"http://www.w3.org/2000/svg\"/>",
    prims = prims, db = db
  )
  if (!is.null(markers)) raw$markers <- markers
  json <- jsonlite::toJSON(raw, auto_unbox = TRUE, null = "null", digits = NA)
  as_mermaid_render(jsonlite::fromJSON(json, simplifyVector = FALSE))
}

# A render of a diagram type without node/edge units, so that only the
# generic conversion (spec 3a, 3b) applies.
generic_render <- function(prims = list(), viewBox = c(0, 0, 200, 100), ...) {
  mini_render(prims, viewBox = viewBox, type = "sequence", ...)
}

generic_items <- function(prims = list(), ...) build_scene(generic_render(prims, ...))$items

generic_doc <- function(prims = list(), ..., viewBox = c(0, 0, 200, 100)) {
  dml_doc(build_diagram_xml(generic_render(prims, viewBox = viewBox), ...)$xml)
}

# -- Flowchart-style units for hand-built renders ------------------------------
#
# The bridge says which element is which diagram unit: an `anc` entry (or,
# for an edge path, the prim itself) carries `unit = list(kind, id[, start,
# end])`, with `id` the unit's id in Mermaid's own data. These helpers set
# `unit` and nothing else that could mark a unit: the elements have neutral
# ids and classes (no `node`, `cluster`, `edgeLabel` or `edgePaths` class, no
# `flowchart-<id>-<n>` id, no `data-id` or `data-edge`), so a test built
# with them passes only when units are read from `unit`.

unit_node <- function(id) list(kind = "node", id = id)
unit_subgraph <- function(id) list(kind = "subgraph", id = id)
unit_edge <- function(id, start, end) list(kind = "edge", id = id, start = start, end = end)
unit_edge_label <- function(id, start, end) {
  list(kind = "edge_label", id = id, start = start, end = end)
}

# The db id of a node drawn as `flowchart-<id>-<n>`, the default unit id of
# the helpers that take a DOM id.
dom_node_id <- function(dom_id) sub("^flowchart-", "", sub("-\\d+$", "", dom_id))

# The two ends of a Mermaid edge id `L_<start>_<end>_<n>` (an end holding an
# underscore must be given explicitly).
edge_ends <- function(edge_id) {
  m <- regmatches(edge_id, regexec("^L_(.+)_(.+?)_\\d+$", edge_id))[[1]]
  if (length(m) == 3L) m[2:3] else c(NA_character_, NA_character_)
}

anc_root <- function() list(list(id = "", cls = "root"))
anc_node <- function(dom_id, cls = "default", id = dom_node_id(dom_id)) {
  c(anc_root(), list(list(id = "", cls = "nodes"),
                     list(id = "", cls = cls, unit = unit_node(id))))
}
anc_label <- function(anc) c(anc, list(list(id = "", cls = "label")))
anc_edge <- function() c(anc_root(), list(list(id = "", cls = "paths")))
anc_cluster <- function(id) {
  c(anc_root(), list(list(id = "", cls = "clusters"),
                     list(id = "", cls = "frame", data = list(look = "classic"),
                          unit = unit_subgraph(id))))
}
anc_cluster_label <- function(id) c(anc_cluster(id), list(list(id = "", cls = "cluster-label")))
anc_edge_label <- function(edge_id, start = edge_ends(edge_id)[[1]],
                           end = edge_ends(edge_id)[[2]]) {
  c(anc_root(), list(list(id = "", cls = "edgeLabels"),
                     list(id = "", cls = "label-box",
                          unit = unit_edge_label(edge_id, start, end)),
                     list(id = "", cls = "label")))
}

db_node <- function(id, shape = "squareRect", dom_id = paste0("flowchart-", id, "-0"),
                    label = id, parent = NULL, is_group = FALSE) {
  n <- list(id = id, label = label, shape = shape, isGroup = is_group)
  if (!is.null(dom_id)) n$domId <- dom_id
  if (!is.null(parent)) n$parentId <- parent
  n
}

db_edge <- function(start, end, id = paste0("L_", start, "_", end, "_0")) {
  list(id = id, start = start, end = end)
}

mini_db <- function(nodes = list(), edges = list()) list(nodes = nodes, edges = edges)

# An edge path prim as Mermaid draws it (open, stroked, arrow at the end); the
# bridge puts the edge's `unit` on the path itself.
prim_edge <- function(d, edge_id, marker_end = "flowchart-v2-pointEnd",
                      start = edge_ends(edge_id)[[1]], end = edge_ends(edge_id)[[2]],
                      ...) {
  p <- prim_path(d, anc = anc_edge(), fill = NULL, stroke = "#333333",
                 marker_end = marker_end, ...)
  p$unit <- unit_edge(edge_id, start, end)
  p
}

# -- Step 2: DOM serials (A2) and marker definitions (A1) ----------------------

# One `anc` entry with its DOM serial `n` (an unnamed <g> has id and cls "").
# `unit`: a unit_*() list when the element is a diagram unit.
anc_g <- function(n, cls = "", id = "", data = NULL, unit = NULL) {
  a <- list(id = id, cls = cls, n = as.integer(n))
  if (!is.null(data)) a$data <- data
  if (!is.null(unit)) a$unit <- unit
  a
}

# Wrap prims in one <g> with serial `n`: its entry goes first in each prim's
# `anc` (outermost first), so in_g(in_g(ps, 3), 2) puts ps in g 3 inside g 2.
in_g <- function(prims, n, cls = "", id = "", data = NULL, unit = NULL) {
  g <- anc_g(n, cls = cls, id = id, data = data, unit = unit)
  purrr::map(prims, function(p) {
    p$anc <- c(list(g), p$anc %|null|% list())
    p
  })
}

# Give each prim its own element serial `n`, from `from` in list order.
with_n <- function(prims, from = 1000L) {
  purrr::imap(unname(prims), function(p, i) {
    p$n <- as.integer(from + i - 1L)
    p
  })
}

# A marker's drawn child as the bridge reports it (A1): shape fields only, `m`
# and `bbox` in marker content coordinates, no i, id, anc, data or markers.
marker_prim <- function(p) {
  p[setdiff(names(p), c("i", "id", "anc", "data", "marker_start", "marker_end",
                        "html", "n"))]
}

# One entry of `render$markers` (A1). `prims` are prim_*() shapes in content
# coordinates; `viewBox` NULL for none.
marker_def <- function(viewBox = NULL, refX = 0, refY = 0, w = 3, h = 3,
                       units = "strokeWidth", orient = "auto", prims = list()) {
  list(viewBox = if (is.null(viewBox)) NULL else as.list(viewBox),
       refX = refX, refY = refY, w = w, h = h, units = units, orient = orient,
       prims = purrr::map(prims, marker_prim))
}

# Mermaid 11's own markers (from the survey SVGs), with their computed paint.

# flowchart-v2-pointEnd: a filled arrow point, tip at content x = 10.
marker_arrow_point <- function(fill = "#333333", stroke = "#333333",
                               w = 8, h = w, refX = 5, units = "userSpaceOnUse",
                               orient = "auto") {
  marker_def(c(0, 0, 10, 10), refX = refX, refY = 5, w = w, h = h, units = units,
             orient = orient,
             prims = list(prim_path("M 0 0 L 10 5 L 0 10 z", fill = fill,
                                    stroke = stroke, stroke_w = 1)))
}

# flowchart-v2-pointStart: the point turned round, orient auto.
marker_arrow_point_start <- function() {
  marker_def(c(0, 0, 10, 10), refX = 4.5, refY = 5, w = 8, h = 8,
             units = "userSpaceOnUse", orient = "auto",
             prims = list(prim_path("M 0 5 L 10 10 L 10 0 z", fill = "#333333",
                                    stroke = "#333333", stroke_w = 1)))
}

# flowchart-v2-circleEnd: a filled circle, refX beyond the content.
marker_circle_end <- function() {
  marker_def(c(0, 0, 10, 10), refX = 11, refY = 5, w = 11, h = 11,
             units = "userSpaceOnUse", orient = "auto",
             prims = list(prim_ellipse(5, 5, 5, fill = "#333333", stroke = "#333333")))
}

# flowchart-v2-crossEnd: two stroked strokes, no fill.
marker_cross_end <- function() {
  marker_def(c(0, 0, 11, 11), refX = 12, refY = 5.2, w = 11, h = 11,
             units = "userSpaceOnUse", orient = "auto",
             prims = list(prim_path("M 1,1 l 9,9 M 10,1 l -9,9", bbox = c(1, 1, 9, 9),
                                    fill = NULL, stroke = "#333333", stroke_w = 2)))
}

# class-extensionEnd: a hollow triangle, no viewBox, strokeWidth units;
# refX 1 (the base) sits on the path end and the tip points beyond it.
marker_extension_end <- function(orient = "auto", fill = NULL, stroke = "#333333") {
  marker_def(NULL, refX = 1, refY = 7, w = 20, h = 28, units = "strokeWidth",
             orient = orient,
             prims = list(prim_path("M 1,1 V 13 L18,7 Z", bbox = c(1, 1, 17, 12),
                                    fill = fill, stroke = stroke, stroke_w = 1)))
}

# class-extensionStart: the same triangle drawn turned round, refX 18, for a
# start with plain orient auto (190 x 240 is Mermaid's own clip box).
marker_extension_start <- function() {
  marker_def(NULL, refX = 18, refY = 7, w = 190, h = 240, units = "strokeWidth",
             orient = "auto",
             prims = list(prim_path("M 1,7 L18,13 V 1 Z", bbox = c(1, 1, 17, 12),
                                    fill = NULL, stroke = "#333333", stroke_w = 1)))
}

# class-compositionStart: a filled diamond, content x 1..18, refX 18.
marker_composition_start <- function() {
  marker_def(NULL, refX = 18, refY = 7, w = 190, h = 240, units = "strokeWidth",
             orient = "auto",
             prims = list(prim_path("M 18,7 L9,13 L1,7 L9,1 Z", bbox = c(1, 1, 17, 12),
                                    fill = "#333333", stroke = "#333333", stroke_w = 1)))
}

# sequencenumber: one filled circle, centred on (refX, refY).
marker_seq_number <- function() {
  marker_def(NULL, refX = 15, refY = 15, w = 60, h = 40, units = "strokeWidth",
             orient = "auto",
             prims = list(prim_ellipse(15, 15, 6, fill = "#333333", stroke = NULL)))
}

# er-oneOrMoreEnd: a crow's foot, one stroked open path (a bar and a lens).
marker_er_one_or_more_end <- function() {
  marker_def(NULL, refX = 27, refY = 18, w = 45, h = 36, units = "strokeWidth",
             orient = "auto",
             prims = list(prim_path("M3,9 L3,27 M9,18 Q27,0 45,18 Q27,36 9,18",
                                    bbox = c(3, 9, 42, 18), fill = NULL,
                                    stroke = "#333333", stroke_w = 1)))
}

# er-onlyOneStart: two bars, one stroked open path.
marker_er_only_one_start <- function() {
  marker_def(NULL, refX = 0, refY = 9, w = 18, h = 18, units = "strokeWidth",
             orient = "auto",
             prims = list(prim_path("M9,0 L9,18 M15,0 L15,18", fill = NULL,
                                    stroke = "#333333", stroke_w = 1)))
}

# A one-prim node unit (shape prim + centred label) for a db shape test.
# `id`: the node's db id (the unit id); by default the one in `dom_id`.
node_prims <- function(dom_id, shape_prim, label = "Label", box = NULL,
                       id = dom_node_id(dom_id)) {
  box <- box %|null|% unlist(shape_prim$bbox)
  tw <- min(40, box[3] / 2)
  th <- min(16, box[4] / 2)
  shape_prim$anc <- anc_node(dom_id, id = id)
  list(shape_prim,
       prim_text(label, x = box[1] + (box[3] - tw) / 2, y = box[2] + (box[4] - th) / 2,
                 w = tw, h = th, anc = anc_label(anc_node(dom_id, id = id))))
}

# A render with one flowchart node `id` of db shape `shape`, drawn by `prim`.
one_node_render <- function(shape, prim, id = "X", label = "Label",
                            viewBox = c(0, 0, 200, 100)) {
  dom <- paste0("flowchart-", id, "-0")
  mini_render(node_prims(dom, prim, label = label, id = id), viewBox = viewBox,
              db = mini_db(list(db_node(id, shape = shape, dom_id = dom, label = label))))
}

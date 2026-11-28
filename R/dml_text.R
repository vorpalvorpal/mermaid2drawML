# R/dml_text.R
# Text rows from the scene -> w:txbxContent. Mermaid's own rows become the
# paragraphs and the text is written exactly as the browser rendered it, so
# Word never re-wraps a label with different font metrics and no markdown or
# entity handling is needed here.

# Words of a font name that Word shows in capitals ("Trebuchet MS",
# "Segoe UI", "PT Sans", "Noto Sans SC").
.font_upper_words <- c("ms", "ui", "pt", "sc", "tc", "jp", "kr")

#' A font name as Word shows it
#'
#' The bridge reports names as the stylesheet wrote them, often lower case.
#' Words are title-cased, a few kept upper case; a name that already has an
#' upper-case letter is taken to be written as intended.
#' @param name Character vector.
#' @noRd
font_display_name <- function(name) {
  purrr::map_chr(name, function(n) {
    if (is.na(n) || grepl("[A-Z]", n)) return(n)
    words <- strsplit(n, " ", fixed = TRUE)[[1]]
    paste(ifelse(words %in% .font_upper_words, toupper(words), stringr::str_to_title(words)),
          collapse = " ")
  })
}

# The first family of a CSS font-family list, without its quotes.
first_family <- function(family) {
  gsub("^[\"']+|[\"']+$", "", trimws(sub(",.*$", "", family)))
}

#' The font of each run: the override, else the face the browser used, else
#' the first family the stylesheet asked for
#' @noRd
run_fonts <- function(runs, override = NULL) {
  if (!is.null(override)) return(rep(override, nrow(runs)))
  used <- !is.na(runs$font) & nzchar(runs$font)
  font_display_name(ifelse(used, runs$font, first_family(runs$family)))
}

#' w:sz for text of `size` px drawn at `scale` EMU per px
#'
#' Rounded down (with a tolerance so that an exact 24 stays 24), since text
#' larger than the browser measured would eat the slack in tight shapes; at
#' least 2, Word's smallest size.
#' @noRd
half_points <- function(size, scale, font_scale = 1) {
  pmax(2L, as.integer(floor(size * scale / .emu_per_half_pt * font_scale + 1e-6)))
}

#' One run. `alpha` below 1 adds Word 2010's text fill after the standard
#' properties, as the schema orders them; Word reads its w14:alpha as a
#' transparency, so it is (1 - alpha) * 100000. The plain w:color stays for
#' readers that know no w14.
#' @noRd
run_xml <- function(t, bold, italic, underline, color, alpha, hp, font) {
  rpr <- c(
    if (!is.na(font) && nzchar(font)) {
      xml_tag("w:rFonts", "w:ascii" = font, "w:hAnsi" = font, "w:eastAsia" = font,
              "w:cs" = font)
    },
    if (bold) xml_tag("w:b"),
    if (italic) xml_tag("w:i"),
    if (!is.na(color)) xml_tag("w:color", "w:val" = color),
    if (!is.na(hp)) c(xml_tag("w:sz", "w:val" = hp), xml_tag("w:szCs", "w:val" = hp)),
    # w:u comes after the sizes in the schema's sequence.
    if (underline) xml_tag("w:u", "w:val" = "single"),
    text_fill_xml(color, alpha)
  )
  xml_tag("w:r", .children = c(
    xml_tag("w:rPr", .children = rpr),
    xml_tag("w:t", "xml:space" = "preserve", .children = xml_escape(t))
  ))
}

text_fill_xml <- function(color, alpha) {
  if (is.na(color) || is.na(alpha) || alpha >= 1) return(NULL)
  xml_tag("w14:textFill", .children = xml_tag(
    "w14:solidFill", .children = xml_tag(
      "w14:srgbClr", "w14:val" = color,
      .children = xml_tag("w14:alpha", "w14:val" = (1 - alpha) * 100000))))
}

#' Runs wrapped in a Word HYPERLINK field: begin, the field code, separate,
#' the runs as its shown result, end
#'
#' A field needs no relationship in the document package, so XML with field
#' links is complete in itself. A quote in the URL is written \" as Word's
#' field syntax escapes it.
#' @noRd
field_link_runs <- function(runs_xml, url) {
  fld <- function(type) xml_tag("w:r", .children = xml_tag("w:fldChar", "w:fldCharType" = type))
  code <- sprintf(" HYPERLINK \"%s\" ", gsub("\"", "\\\\\"", url))
  instr <- xml_tag("w:r", .children = xml_tag("w:instrText", "xml:space" = "preserve",
                                               .children = xml_escape(code)))
  c(fld("begin"), instr, fld("separate"), runs_xml, fld("end"))
}

#' One text row (a tibble of runs) as a paragraph
#' @param style list(scale, font_scale, font_family): how text is sized and
#'   which font overrides the bridge's.
#' @param href A URL the row's runs link to as a field, or NULL.
#' @noRd
row_xml <- function(runs, align, style, href = NULL) {
  hp <- half_points(runs$size, style$scale, style$font_scale)
  runs_xml <- purrr::pmap_chr(
    list(runs$t, runs$bold, runs$italic, runs$underline, runs$color,
         runs$alpha %||% rep(1, nrow(runs)), hp, run_fonts(runs, style$font_family)),
    run_xml
  )
  dml_paragraph(align, if (is.null(href)) runs_xml else field_link_runs(runs_xml, href))
}

#' A scene text (see R/scene.R) as wps:txbx
#' @param href A URL every paragraph links to as a field, or NULL.
#' @noRd
dml_txbx <- function(text, style, href = NULL) {
  paras <- purrr::map_chr(text$rows, row_xml, align = text$align, style = style, href = href)
  xml_tag("wps:txbx", .children = xml_tag("w:txbxContent", .children = paras))
}

#' The size, in half-points, carried by the most characters of the diagram's
#' text; NA when it has none
#' @noRd
main_half_points <- function(items, scale, font_scale = 1) {
  rows <- purrr::list_flatten(purrr::map(purrr::compact(items$text), "rows"))
  runs <- purrr::list_rbind(rows)
  if (is.null(runs) || nrow(runs) == 0L) return(NA_integer_)
  chars <- tapply(nchar(runs$t), half_points(runs$size, scale, font_scale), sum)
  if (length(chars) == 0L) return(NA_integer_)
  as.integer(names(chars)[[which.max(chars)]])
}

#' Warn once when the main text comes out below 7 pt (14 half-points)
#' @noRd
warn_small_text <- function(hp) {
  if (is.na(hp) || hp >= 14L) return(invisible(NULL))
  points <- hp / 2
  pt <- format(points)
  cli::cli_warn(
    c("Most of the diagram's text is {pt} pt, below the 7 pt that stays readable.",
      i = "Text shrinks with a diagram scaled to fit the page; use a wider page, landscape orientation, or split the diagram into smaller ones."),
    class = "mermaid2drawml_small_text", points = points
  )
}

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

run_xml <- function(t, bold, italic, underline, color, hp, font) {
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
    if (underline) xml_tag("w:u", "w:val" = "single")
  )
  xml_tag("w:r", .children = c(
    xml_tag("w:rPr", .children = rpr),
    xml_tag("w:t", "xml:space" = "preserve", .children = xml_escape(t))
  ))
}

#' One text row (a tibble of runs) as a paragraph
#' @param style list(scale, font_scale, font_family): how text is sized and
#'   which font overrides the bridge's.
#' @noRd
row_xml <- function(runs, align, style) {
  hp <- half_points(runs$size, style$scale, style$font_scale)
  runs_xml <- purrr::pmap_chr(
    list(runs$t, runs$bold, runs$italic, runs$underline, runs$color, hp,
         run_fonts(runs, style$font_family)),
    run_xml
  )
  dml_paragraph(align, runs_xml)
}

#' A scene text (see R/scene.R) as wps:txbx
#' @noRd
dml_txbx <- function(text, style) {
  paras <- purrr::map_chr(text$rows, row_xml, align = text$align, style = style)
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
  pt <- format(hp / 2)
  cli::cli_warn(
    c("Most of the diagram's text is {pt} pt, below the 7 pt that stays readable.",
      i = "Text shrinks with a diagram scaled to fit the page; use a wider page, landscape orientation, or split the diagram into smaller ones."),
    class = "mermaid2drawml_small_text"
  )
}

#' @keywords internal
"_PACKAGE"

# EMU conversion ---------------------------------------------------------------

inches_to_emu <- function(x) as.integer(round(x * 914400))
emu_to_inches <- function(x) x / 914400

# XML helpers ------------------------------------------------------------------

xml_escape <- function(text) {
  if (is.null(text) || length(text) == 0) return(text)
  if (is.na(text)) return(text)
  text <- gsub("&",  "&amp;",  text, fixed = TRUE)
  text <- gsub("<",  "&lt;",   text, fixed = TRUE)
  text <- gsub(">",  "&gt;",   text, fixed = TRUE)
  text <- gsub('"',  "&quot;", text, fixed = TRUE)
  text <- gsub("'",  "&apos;", text, fixed = TRUE)
  text
}

hex_to_drawingml <- function(hex) {
  hex <- gsub("^#", "", trimws(hex))
  if (nchar(hex) == 3) {
    chars <- strsplit(hex, "")[[1]]
    hex   <- paste0(rep(chars, each = 2), collapse = "")
  }
  toupper(substr(hex, 1, 6))
}

.named_colours <- c(
  transparent = NA_character_, none = NA_character_,
  white = "FFFFFF", black = "000000", red = "FF0000",
  green = "008000", blue = "0000FF", orange = "FFA500",
  gold = "FFD700", yellow = "FFFF00", purple = "800080",
  pink = "FFC0CB", grey = "808080", gray = "808080",
  lightgrey = "D3D3D3", lightgray = "D3D3D3",
  darkgrey = "A9A9A9", darkgray = "A9A9A9",
  silver = "C0C0C0", navy = "000080", teal = "008080",
  lime = "00FF00", aqua = "00FFFF", cyan = "00FFFF",
  magenta = "FF00FF", fuchsia = "FF00FF", maroon = "800000",
  olive = "808000"
)

parse_css_colour <- function(css) {
  if (is.null(css) || is.na(css) || !nzchar(trimws(css))) return(NA_character_)
  css <- trimws(tolower(css))
  if (css %in% names(.named_colours)) return(.named_colours[[css]])
  if (grepl("^#", css)) return(hex_to_drawingml(css))
  if (grepl("^rgb", css)) {
    nums <- suppressWarnings(as.integer(
      regmatches(css, gregexpr("[0-9]+", css))[[1]]
    ))
    if (length(nums) >= 3) return(toupper(sprintf("%02X%02X%02X", nums[1], nums[2], nums[3])))
  }
  NA_character_
}

# Label cleaning ---------------------------------------------------------------

strip_html <- function(label) {
  if (is.null(label) || length(label) == 0 || is.na(label) || !nzchar(label)) return(label)
  label <- gsub('<[^>]+style=[^>]*>', '', label, perl = TRUE)
  label <- gsub('<[^>]+>', '', label, perl = TRUE)
  label <- gsub('&lt;',   '<', label, fixed = TRUE)
  label <- gsub('&gt;',   '>', label, fixed = TRUE)
  label <- gsub('&amp;',  '&', label, fixed = TRUE)
  label <- gsub('&nbsp;', ' ', label, fixed = TRUE)
  label <- gsub('&#160;', ' ', label, fixed = TRUE)
  label <- gsub('\\*\\*(.+?)\\*\\*', '\\1', label, perl = TRUE)
  label <- gsub('__(.+?)__',         '\\1', label, perl = TRUE)
  label <- gsub('\\*(.+?)\\*',       '\\1', label, perl = TRUE)
  trimws(gsub('\\s+', ' ', label))
}

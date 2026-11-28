# R/utils.R
# Units and XML escaping shared by the DrawingML builders.

# EMU (English Metric Units): 914400 per inch, so 9525 per CSS pixel at 96
# dpi, 12700 per point and 6350 per half-point (the unit of w:sz).
.emu_per_inch <- 914400
.emu_per_px <- 9525
.emu_per_pt <- 12700
.emu_per_half_pt <- 6350
.px_per_inch <- .emu_per_inch / .emu_per_px

# Characters XML 1.0 forbids. The browser can hand them over in text, and
# one of them marks the shape ids still to be numbered (R/emit.R), so no
# escaped value may keep them.
.xml_illegal_re <- "[\u0001-\u0008\u000B\u000C\u000E-\u001F]"

#' Escape text for an XML attribute value or text node
#' @param text Character vector; `NA` stays `NA`.
#' @noRd
xml_escape <- function(text) {
  text <- gsub(.xml_illegal_re, "", text, perl = TRUE)
  text <- gsub("&", "&amp;", text, fixed = TRUE)
  text <- gsub("<", "&lt;", text, fixed = TRUE)
  text <- gsub(">", "&gt;", text, fixed = TRUE)
  text <- gsub("\"", "&quot;", text, fixed = TRUE)
  gsub("'", "&apos;", text, fixed = TRUE)
}

#' Numbers as whole-number strings, at any size
#'
#' Going through R integers would turn anything past 2^31 into NA, and
#' `format()` would use scientific notation, so the digits are written
#' directly. Adding zero turns a negative zero into zero.
#' @param v Numeric vector; non-finite values are an error, never "NA".
#' @noRd
whole_number <- function(v) {
  if (!all(is.finite(v))) {
    cli::cli_abort("A non-finite number cannot be written as an XML attribute.",
                   class = "mermaid2drawml_render_error")
  }
  formatC(round(v) + 0, format = "f", digits = 0)
}

# ---- small data helpers -------------------------------------------------------

#' A tibble from columns, length-1 columns recycled
#'
#' For data frames made once per prim, unit or item: tibble()'s checks and
#' name repair cost more than the data. Columns must already have the right
#' types and lengths (1 or n).
#' @noRd
quick_tibble <- function(...) {
  cols <- list(...)
  n <- if (length(cols) == 0L) 0L else max(vapply(cols, vctrs::vec_size, integer(1)))
  if (any(vapply(cols, vctrs::vec_size, integer(1)) == 0L)) n <- 0L
  tibble::new_tibble(vctrs::vec_recycle_common(!!!cols, .size = n), nrow = n)
}

#' The row numbers of each key, split once
#'
#' @param keys A vector (NA is a key like any other).
#' @return A function of one key giving its rows, integer(0) for a key not
#'   in `keys`.
#' @noRd
rows_by_key <- function(keys) {
  levels <- unique(keys)
  rows <- split(seq_along(keys), factor(match(keys, levels), levels = seq_along(levels)))
  function(key) {
    j <- match(key, levels)
    if (is.na(j)) integer() else rows[[j]]
  }
}

# R/utils.R
# Units and XML escaping shared by the DrawingML builders.

# EMU (English Metric Units): 914400 per inch, so 9525 per CSS pixel at 96
# dpi, 12700 per point and 6350 per half-point (the unit of w:sz).
.emu_per_inch <- 914400
.emu_per_px <- 9525
.emu_per_pt <- 12700
.emu_per_half_pt <- 6350

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

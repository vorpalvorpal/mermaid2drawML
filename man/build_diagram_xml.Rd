\name{build_diagram_xml}
\alias{build_diagram_xml}
\title{Build DrawingML XML for a laid-out mermaid_graph}
\usage{
build_diagram_xml(
  graph,
  start_id = 1L,
  default_fill = "FFFFFF",
  default_stroke = "4472C4",
  default_text_color = "000000",
  stroke_width_pt = 1.5
)
}
\arguments{
\item{graph}{A laid-out \code{mermaid_graph} (from \code{\link{calculate_layout}}).}

\item{start_id}{Integer starting shape ID. Increment between diagrams.}

\item{default_fill}{Default node fill colour (6-char hex). Default \code{"FFFFFF"}.}

\item{default_stroke}{Default stroke colour. Default \code{"4472C4"}.}

\item{default_text_color}{Default text colour. Default \code{"000000"}.}

\item{stroke_width_pt}{Stroke width in points. Default 1.5.}
}
\value{
Named list with two elements:
\itemize{
\item \code{xml} — character string of DrawingML XML (a \verb{<w:p>} element)
\item \code{next_id} — integer, the next available shape ID
}
}
\description{
Converts a laid-out \code{mermaid_graph} into a DrawingML XML string suitable
for insertion into a Word document via \code{officer::body_add_xml()}.
}
\examples{
g <- parse_mermaid("flowchart TB\n  A[Start] --> B[End]") |>
  calculate_layout()
r <- build_diagram_xml(g)
nchar(r$xml)  # length of the XML string
r$next_id     # next available shape ID
}

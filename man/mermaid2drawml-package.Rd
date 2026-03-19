\name{mermaid2drawml-package}
\alias{mermaid2drawml}
\alias{mermaid2drawml-package}
\docType{package}
\title{mermaid2drawml: Convert Mermaid Diagrams to Editable Word Shapes}
\description{
Converts Mermaid flowchart diagrams to editable Microsoft Word shapes
using the DrawingML \code{wps:wsp} format. The resulting shapes are natively
editable in Word and can be inserted into documents via the officer package,
suitable for use in R Markdown workflows.
}
\details{
The main pipeline is:
\enumerate{
\item \code{\link{parse_mermaid}(code)} -- parse Mermaid syntax into a \code{mermaid_graph}
\item \code{\link{calculate_layout}(graph)} -- compute x/y positions using igraph Sugiyama layout
\item \code{\link{build_diagram_xml}(graph)} -- produce DrawingML XML
}

The user-facing convenience functions are \code{\link{mermaid_to_word}()} and
\code{\link{body_add_mermaid}()}.
}
\seealso{
\code{\link{parse_mermaid}}, \code{\link{calculate_layout}},
\code{\link{build_diagram_xml}}, \code{\link{mermaid_to_word}},
\code{\link{body_add_mermaid}}
}
\keyword{internal}

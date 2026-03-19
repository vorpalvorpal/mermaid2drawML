\name{body_add_mermaid}
\alias{body_add_mermaid}
\title{Add a Mermaid diagram to a Word document}
\usage{
body_add_mermaid(doc, mermaid_code, start_id = 1L, ...)
}
\arguments{
\item{doc}{An \code{rdocx} object from \code{officer::read_docx()}.}

\item{mermaid_code}{A character string of Mermaid flowchart syntax.}

\item{start_id}{Integer starting shape ID. Default 1L.}

\item{...}{Additional arguments passed to \code{\link{calculate_layout}}.}
}
\value{
The modified \code{rdocx} object (invisibly), for use in pipes.
}
\description{
A convenience wrapper around \code{\link{mermaid_to_word}} that inserts the
diagram directly into an \code{rdocx} document from the officer package.
}
\examples{
\dontrun{
library(officer)
doc <- read_docx() |>
  body_add_mermaid("flowchart TB\n  A[Start] --> B[End]")
print(doc, target = "output.docx")
}
}

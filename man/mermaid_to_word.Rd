\name{mermaid_to_word}
\alias{mermaid_to_word}
\title{Convert a Mermaid diagram to a Word-embeddable object}
\usage{
mermaid_to_word(mermaid_code, start_id = 1L, ...)
}
\arguments{
\item{mermaid_code}{A character string of Mermaid flowchart syntax.}

\item{start_id}{Integer starting shape ID. Increment between diagrams in the
same document to avoid ID collisions. Default 1L.}

\item{...}{Additional arguments passed to \code{\link{calculate_layout}}.}
}
\value{
A \code{mermaid_word_diagram} object with fields:
\itemize{
\item \code{xml} — DrawingML XML string
\item \code{graph} — the parsed and laid-out \code{mermaid_graph}
\item \code{next_id} — the next available shape ID
}
}
\description{
Parses a Mermaid flowchart, computes a hierarchical layout, and builds
DrawingML XML suitable for insertion into a Word document via
\code{\link{body_add_mermaid}} or \code{officer::body_add_xml()}.
}
\examples{
d <- mermaid_to_word("flowchart TB\n  A[Start] --> B[End]")
print(d)
}

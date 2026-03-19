\name{parse_mermaid}
\alias{parse_mermaid}
\title{Parse a Mermaid flowchart diagram}
\usage{
parse_mermaid(mermaid_code)
}
\arguments{
\item{mermaid_code}{A character string containing Mermaid flowchart/graph syntax.}
}
\value{
A \code{mermaid_graph} S3 object containing direction, nodes, edges,
subgraphs, and classes.
}
\description{
Parses Mermaid \code{flowchart} or \code{graph} syntax and returns a structured
\code{mermaid_graph} object containing the diagram's nodes, edges, subgraphs,
and class definitions.
}
\examples{
g <- parse_mermaid("flowchart TB\n  A[Start] --> B[End]")
print(g)
}

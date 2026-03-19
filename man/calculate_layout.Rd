\name{calculate_layout}
\alias{calculate_layout}
\title{Calculate layout positions for a mermaid_graph}
\usage{
calculate_layout(
  graph,
  page_width_in = 6,
  page_height_in = 8,
  node_w_in = 1.4,
  node_h_in = 0.55,
  h_gap_in = 0.3,
  v_gap_in = 0.4,
  margin_in = 1
)
}
\arguments{
\item{graph}{A \code{mermaid_graph} from \code{\link{parse_mermaid}}.}

\item{page_width_in}{Usable content width in inches (default 6.0).}

\item{page_height_in}{Maximum diagram height in inches (default 8.0).}

\item{node_w_in}{Default node width in inches (default 1.4).}

\item{node_h_in}{Default node height in inches (default 0.55).}

\item{h_gap_in}{Horizontal gap between nodes (default 0.3).}

\item{v_gap_in}{Vertical gap between ranks (default 0.4).}

\item{margin_in}{Page margin offset in inches (default 1.0).}
}
\value{
Modified \code{mermaid_graph} with layout columns (x, y, width, height) added
to nodes and (x1, y1, x2, y2) added to edges.
}
\description{
Uses igraph's Sugiyama hierarchical layout algorithm to compute x/y positions
for nodes in inches. The positions are relative to the page origin (top-left).
}
\examples{
g <- parse_mermaid("flowchart TB\n  A --> B --> C")
g <- calculate_layout(g)
g$nodes[, c("id", "x", "y")]
}

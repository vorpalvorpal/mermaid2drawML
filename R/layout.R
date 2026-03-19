#' Calculate layout positions for a mermaid_graph
#'
#' Uses igraph's Sugiyama hierarchical layout. Adds x, y, width, height to
#' nodes (in inches, origin at top-left of page) and x1, y1, x2, y2 to edges.
#'
#' @param graph A `mermaid_graph` from [parse_mermaid()].
#' @param page_width_in Usable content width in inches (default 6.0).
#' @param page_height_in Maximum diagram height in inches (default 8.0).
#' @param node_w_in Default node width in inches (default 1.4).
#' @param node_h_in Default node height in inches (default 0.55).
#' @param h_gap_in Horizontal gap between nodes (default 0.3).
#' @param v_gap_in Vertical gap between ranks (default 0.4).
#' @param margin_in Page margin offset in inches (default 1.0).
#' @return Modified `mermaid_graph` with layout columns added to nodes and edges.
#' @export
calculate_layout <- function(graph,
                              page_width_in  = 6.0,
                              page_height_in = 8.0,
                              node_w_in      = 1.4,
                              node_h_in      = 0.55,
                              h_gap_in       = 0.3,
                              v_gap_in       = 0.4,
                              margin_in      = 1.0) {

  nodes <- graph$nodes
  edges <- graph$edges
  n     <- nrow(nodes)

  # Initialise layout columns
  nodes$x      <- rep(margin_in, n)
  nodes$y      <- rep(margin_in, n)
  nodes$width  <- rep(node_w_in,  n)
  nodes$height <- rep(node_h_in,  n)

  add_edge_cols <- function(e) {
    e$x1 <- NA_real_; e$y1 <- NA_real_
    e$x2 <- NA_real_; e$y2 <- NA_real_
    e
  }

  if (n == 0) {
    graph$nodes <- nodes
    graph$edges <- add_edge_cols(edges)
    return(graph)
  }

  if (n == 1) {
    graph$nodes <- nodes
    graph$edges <- add_edge_cols(edges)
    return(graph)
  }

  direction <- toupper(graph$direction)

  # Valid edges only (both endpoints exist, no self-loops)
  valid_e <- edges[edges$from %in% nodes$id & edges$to %in% nodes$id &
                     edges$from != edges$to, , drop = FALSE]

  if (nrow(valid_e) == 0) {
    # Grid layout for disconnected nodes
    ncols   <- max(1L, ceiling(sqrt(n)))
    nodes$x <- margin_in + ((seq_len(n) - 1L) %% ncols) * (node_w_in + h_gap_in)
    nodes$y <- margin_in + (floor((seq_len(n) - 1L) / ncols)) * (node_h_in + v_gap_in)
    graph$nodes <- nodes
    graph$edges <- add_edge_cols(edges)
    return(graph)
  }

  g <- igraph::graph_from_data_frame(
    d        = valid_e[, c("from", "to"), drop = FALSE],
    directed = TRUE,
    vertices = data.frame(name = nodes$id, stringsAsFactors = FALSE)
  )

  lay <- igraph::layout_with_sugiyama(g)$layout
  # lay: rows = vertices, col1 = x (horizontal), col2 = y (vertical rank from bottom)

  # Orient for direction.
  # layout_with_sugiyama: col2 is the layer rank, descending from root.
  # Higher col2 value = closer to root (for TB that means top).
  # Negate col2 so that root is at the smaller coordinate value.
  if (direction %in% c("TB", "TD")) {
    # horizontal spread = col1, vertical = rank (negate so root is at top)
    coords <- data.frame(cx = lay[, 1], cy = -lay[, 2])
  } else if (direction == "BT") {
    # same horizontal, vertical not negated (root at bottom)
    coords <- data.frame(cx = lay[, 1], cy =  lay[, 2])
  } else if (direction == "LR") {
    # horizontal = rank (negate so root is at left), vertical = col1
    coords <- data.frame(cx = -lay[, 2], cy = lay[, 1])
  } else { # RL
    # horizontal = rank (not negated so root is at right), vertical = col1
    coords <- data.frame(cx =  lay[, 2], cy = lay[, 1])
  }

  # Normalise to [0,1]
  xr <- range(coords$cx); yr <- range(coords$cy)
  coords$cx <- if (diff(xr) > 0) (coords$cx - xr[1]) / diff(xr) else rep(0.5, n)
  coords$cy <- if (diff(yr) > 0) (coords$cy - yr[1]) / diff(yr) else rep(0.5, n)

  # Scale to page
  usable_w <- max(page_width_in  - node_w_in - 2 * h_gap_in, node_w_in)
  usable_h <- max(page_height_in - node_h_in - 2 * v_gap_in, node_h_in)
  coords$cx <- margin_in + coords$cx * usable_w
  coords$cy <- margin_in + coords$cy * usable_h

  # Match back by vertex name
  vnames <- igraph::V(g)$name
  for (i in seq_len(n)) {
    idx <- match(nodes$id[i], vnames)
    if (!is.na(idx)) {
      nodes$x[i] <- coords$cx[idx]
      nodes$y[i] <- coords$cy[idx]
    }
  }

  graph$nodes <- nodes

  # Edge endpoints
  edges <- add_edge_cols(edges)
  nd_idx <- setNames(seq_len(n), nodes$id)

  for (i in seq_len(nrow(edges))) {
    fi <- nd_idx[edges$from[i]]
    ti <- nd_idx[edges$to[i]]
    if (is.na(fi) || is.na(ti)) next
    fn <- nodes[fi, ]; tn <- nodes[ti, ]

    if (direction %in% c("TB", "TD")) {
      edges$x1[i] <- fn$x + fn$width  / 2; edges$y1[i] <- fn$y + fn$height
      edges$x2[i] <- tn$x + tn$width  / 2; edges$y2[i] <- tn$y
    } else if (direction == "BT") {
      edges$x1[i] <- fn$x + fn$width  / 2; edges$y1[i] <- fn$y
      edges$x2[i] <- tn$x + tn$width  / 2; edges$y2[i] <- tn$y + tn$height
    } else if (direction == "LR") {
      edges$x1[i] <- fn$x + fn$width;  edges$y1[i] <- fn$y + fn$height / 2
      edges$x2[i] <- tn$x;             edges$y2[i] <- tn$y + tn$height / 2
    } else { # RL
      edges$x1[i] <- fn$x;             edges$y1[i] <- fn$y + fn$height / 2
      edges$x2[i] <- tn$x + tn$width;  edges$y2[i] <- tn$y + tn$height / 2
    }
  }

  graph$edges <- edges
  graph
}

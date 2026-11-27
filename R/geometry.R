# R/geometry.R
# Pure geometry: SVG affine matrices and path data.
#
# A matrix is c(a, b, c, d, e, f) in SVG order:
#   x' = a x + c y + e,  y' = b x + d y + f.
# A path is a tibble of ABSOLUTE segments using only M, L, C and Z, with
# columns cmd, x1, y1, x2, y2, x, y (see path_parse()).

# ---- matrices ---------------------------------------------------------------

#' Apply an affine matrix to points
#' @param m Numeric length 6.
#' @param x,y Numeric vectors of the same length.
#' @return A tibble with columns `x` and `y`.
#' @noRd
mat_apply <- function(m, x, y) {
  # Both results are computed from the inputs before tibble() can shadow `x`.
  new_x <- m[[1]] * x + m[[3]] * y + m[[5]]
  new_y <- m[[2]] * x + m[[4]] * y + m[[6]]
  tibble::tibble(x = new_x, y = new_y)
}

#' Uniform scale factor of a matrix (square root of the absolute determinant)
#' @noRd
mat_scale <- function(m) sqrt(abs(m[[1]] * m[[4]] - m[[2]] * m[[3]]))

#' Rotation of a matrix in degrees clockwise, in (-180, 180]
#' @noRd
mat_rotation <- function(m) atan2(m[[2]], m[[1]]) * 180 / pi

#' TRUE when a matrix has no rotation or skew; flips count as upright
#' @noRd
mat_is_upright <- function(m, tol = 1e-6) {
  abs(m[[2]]) <= tol && abs(m[[3]]) <= tol
}

# ---- path construction ------------------------------------------------------

path_cols <- c("cmd", "x1", "y1", "x2", "y2", "x", "y")

# One constructor so every path tibble has the same columns and types.
path_tbl <- function(cmd = character(), x1 = NA_real_, y1 = NA_real_,
                     x2 = NA_real_, y2 = NA_real_,
                     x = NA_real_, y = NA_real_) {
  n <- length(cmd)
  rep_n <- function(v) if (n == 0L) numeric() else rep_len(as.numeric(v), n)
  tibble::tibble(cmd = cmd, x1 = rep_n(x1), y1 = rep_n(y1),
                 x2 = rep_n(x2), y2 = rep_n(y2),
                 x = rep_n(x), y = rep_n(y))
}

#' Build a path from a list of points
#'
#' @param points List of points, each `c(x, y)`, `list(x, y)` or a list with
#'   names `x` and `y` (as parsed from the bridge's JSON).
#' @param closed Append a `Z`.
#' @noRd
poly_path <- function(points, closed) {
  xy <- purrr::map(points, function(p) {
    p <- unlist(p)
    if (!is.null(names(p)) && all(c("x", "y") %in% names(p))) {
      as.numeric(p[c("x", "y")])
    } else {
      as.numeric(p[1:2])
    }
  })
  x <- purrr::map_dbl(xy, 1L)
  y <- purrr::map_dbl(xy, 2L)
  n <- length(x)
  if (n == 0L) return(path_tbl())
  path_tbl(
    cmd = c("M", rep("L", n - 1L), if (closed) "Z"),
    x = c(x, if (closed) NA_real_),
    y = c(y, if (closed) NA_real_)
  )
}

# ---- path parsing -----------------------------------------------------------

.num_re <- "[-+]?(?:\\d+\\.?\\d*|\\.\\d+)(?:[eE][-+]?\\d+)?"
.cmd_chars <- "MmLlHhVvCcSsQqTtAaZz"

# Arc arguments: flags are single characters and may be packed ("01-10").
.arc_re <- local({
  sep <- "[\\s,]*"
  n <- paste0("(", .num_re, ")")
  f <- "([01])"
  paste0(sep, n, sep, n, sep, n, sep, f, sep, f, sep, n, sep, n)
})

.arity <- c(M = 2L, L = 2L, H = 1L, V = 1L, C = 6L, S = 4L, Q = 4L, T = 2L,
            A = 7L, Z = 0L)

# Argument rows (one row per repeat of the command) as a numeric matrix.
path_args <- function(op, text) {
  k <- .arity[[toupper(op)]]
  if (k == 0L) return(matrix(numeric(), 0L, 0L))
  if (toupper(op) == "A") {
    m <- stringr::str_match_all(text, .arc_re)[[1]][, -1, drop = FALSE]
    storage.mode(m) <- "numeric"
    return(m)
  }
  v <- as.numeric(stringr::str_extract_all(text, .num_re)[[1]])
  n <- length(v) %/% k
  matrix(v[seq_len(n * k)], nrow = n, ncol = k, byrow = TRUE)
}

# One command letter with its arguments becomes one or more simple commands;
# repeated pairs after M are implicit linetos.
path_commands <- function(op, text) {
  args <- path_args(op, text)
  if (toupper(op) == "Z") return(list(list(op = op, a = numeric())))
  rows <- purrr::map(seq_len(nrow(args)), ~ args[.x, ])
  ops <- c(op, rep(if (op == "M") "L" else if (op == "m") "l" else op,
                   max(length(rows) - 1L, 0L)))
  purrr::map2(ops[seq_along(rows)], rows, ~ list(op = .x, a = .y))
}

path_tokenise <- function(d) {
  segs <- stringr::str_match_all(
    d, paste0("([", .cmd_chars, "])([^", .cmd_chars, "]*)"))[[1]]
  # A path must start with a moveto; anything else is not a path.
  if (nrow(segs) == 0L || !segs[1, 2] %in% c("M", "m")) return(list())
  purrr::list_flatten(purrr::map2(segs[, 2], segs[, 3], path_commands))
}

# ---- arcs -------------------------------------------------------------------

# Endpoint parameterisation to centre parameterisation (SVG 1.1 F.6.5), then
# one cubic per at most 90 degrees.
arc_to_cubics <- function(x0, y0, rx, ry, phi_deg, fa, fs, x, y) {
  if (x0 == x && y0 == y) return(path_tbl())
  rx <- abs(rx)
  ry <- abs(ry)
  if (rx == 0 || ry == 0) return(path_tbl("L", x = x, y = y))

  phi <- phi_deg * pi / 180
  cp <- cos(phi)
  sp <- sin(phi)
  dx <- (x0 - x) / 2
  dy <- (y0 - y) / 2
  x1p <- cp * dx + sp * dy
  y1p <- -sp * dx + cp * dy

  # Radii too small to span the chord are scaled up.
  lambda <- x1p^2 / rx^2 + y1p^2 / ry^2
  if (lambda > 1) {
    rx <- rx * sqrt(lambda)
    ry <- ry * sqrt(lambda)
  }

  num <- rx^2 * ry^2 - rx^2 * y1p^2 - ry^2 * x1p^2
  den <- rx^2 * y1p^2 + ry^2 * x1p^2
  coef <- (if (fa != fs) 1 else -1) * sqrt(max(0, num / den))
  cxp <- coef * rx * y1p / ry
  cyp <- -coef * ry * x1p / rx
  cx <- cp * cxp - sp * cyp + (x0 + x) / 2
  cy <- sp * cxp + cp * cyp + (y0 + y) / 2

  th1 <- atan2((y1p - cyp) / ry, (x1p - cxp) / rx)
  th2 <- atan2((-y1p - cyp) / ry, (-x1p - cxp) / rx)
  dth <- th2 - th1
  if (fs == 0 && dth > 0) dth <- dth - 2 * pi
  if (fs == 1 && dth < 0) dth <- dth + 2 * pi

  n <- max(1L, as.integer(ceiling(abs(dth) / (pi / 2) - 1e-9)))
  delta <- dth / n
  k <- 4 / 3 * tan(delta / 4)
  a1 <- th1 + (seq_len(n) - 1L) * delta
  a2 <- a1 + delta

  px <- function(a) cx + rx * cos(a) * cp - ry * sin(a) * sp
  py <- function(a) cy + rx * cos(a) * sp + ry * sin(a) * cp
  dpx <- function(a) -rx * sin(a) * cp - ry * cos(a) * sp
  dpy <- function(a) -rx * sin(a) * sp + ry * cos(a) * cp

  ex <- px(a2)
  ey <- py(a2)
  # Land exactly on the requested end point whatever the rounding.
  ex[n] <- x
  ey[n] <- y
  path_tbl(rep("C", n), x1 = px(a1) + k * dpx(a1), y1 = py(a1) + k * dpy(a1),
           x2 = ex - k * dpx(a2), y2 = ey - k * dpy(a2),
           x = ex, y = ey)
}

# ---- path folding -----------------------------------------------------------

# Parser state: current point, subpath start, and the last control point with
# its kind (for S and T reflection).
path_init <- function() {
  list(x = 0, y = 0, sx = 0, sy = 0, cx = 0, cy = 0, ck = "",
       rows = path_tbl())
}

# One simple command applied to the state: returns the new state, whose `rows`
# are the absolute segments that command produced.
path_step <- function(s, cmd) {
  op <- cmd$op
  a <- cmd$a
  ox <- if (op == tolower(op)) s$x else 0
  oy <- if (op == tolower(op)) s$y else 0
  moved <- function(rows, x, y, cx = x, cy = y, ck = "", sx = s$sx, sy = s$sy) {
    list(x = x, y = y, sx = sx, sy = sy, cx = cx, cy = cy, ck = ck, rows = rows)
  }
  cubic <- function(x1, y1, x2, y2, x, y, ck = "C", cx = x2, cy = y2) {
    moved(path_tbl("C", x1, y1, x2, y2, x, y), x, y, cx, cy, ck)
  }
  # A quadratic raised to a cubic with the same shape.
  quad <- function(qx, qy, x, y) {
    cubic(s$x + 2 / 3 * (qx - s$x), s$y + 2 / 3 * (qy - s$y),
          x + 2 / 3 * (qx - x), y + 2 / 3 * (qy - y), x, y,
          ck = "Q", cx = qx, cy = qy)
  }
  switch(
    toupper(op),
    M = moved(path_tbl("M", x = a[1] + ox, y = a[2] + oy),
              a[1] + ox, a[2] + oy, sx = a[1] + ox, sy = a[2] + oy),
    L = moved(path_tbl("L", x = a[1] + ox, y = a[2] + oy), a[1] + ox, a[2] + oy),
    H = moved(path_tbl("L", x = a[1] + ox, y = s$y), a[1] + ox, s$y),
    V = moved(path_tbl("L", x = s$x, y = a[1] + oy), s$x, a[1] + oy),
    C = cubic(a[1] + ox, a[2] + oy, a[3] + ox, a[4] + oy, a[5] + ox, a[6] + oy),
    S = {
      reflect <- s$ck == "C"
      cubic(if (reflect) 2 * s$x - s$cx else s$x,
            if (reflect) 2 * s$y - s$cy else s$y,
            a[1] + ox, a[2] + oy, a[3] + ox, a[4] + oy)
    },
    Q = quad(a[1] + ox, a[2] + oy, a[3] + ox, a[4] + oy),
    T = {
      reflect <- s$ck == "Q"
      quad(if (reflect) 2 * s$x - s$cx else s$x,
           if (reflect) 2 * s$y - s$cy else s$y,
           a[1] + ox, a[2] + oy)
    },
    A = moved(arc_to_cubics(s$x, s$y, a[1], a[2], a[3], a[4], a[5],
                            a[6] + ox, a[7] + oy),
              a[6] + ox, a[7] + oy),
    Z = moved(path_tbl("Z"), s$sx, s$sy)
  )
}

#' Parse SVG path data into absolute M, L, C and Z segments
#'
#' @param d Path data string.
#' @return A tibble with columns `cmd`, `x1`, `y1`, `x2`, `y2`, `x`, `y`;
#'   zero rows when `d` is empty or not a path.
#' @noRd
path_parse <- function(d) {
  if (length(d) != 1L || is.na(d)) return(path_tbl())
  cmds <- path_tokenise(d)
  if (length(cmds) == 0L) return(path_tbl())
  states <- purrr::accumulate(cmds, path_step, .init = path_init())[-1]
  purrr::list_rbind(purrr::map(states, "rows"))
}

# ---- path operations ---------------------------------------------------------

#' Apply a matrix to every point of a path, control points included
#' @noRd
path_transform <- function(path, m) {
  p <- mat_apply(m, path$x, path$y)
  c1 <- mat_apply(m, path$x1, path$y1)
  c2 <- mat_apply(m, path$x2, path$y2)
  path$x <- p$x
  path$y <- p$y
  path$x1 <- c1$x
  path$y1 <- c1$y
  path$x2 <- c2$x
  path$y2 <- c2$y
  path
}

# Values of a cubic Bezier (one coordinate) at the roots of its derivative
# inside (0, 1); these are the only interior candidates for its extremes.
cubic_extrema <- function(p0, p1, p2, p3) {
  a <- -p0 + 3 * p1 - 3 * p2 + p3
  b <- 2 * (p0 - 2 * p1 + p2)
  c <- p1 - p0
  roots <- if (abs(a) < 1e-12) {
    if (abs(b) < 1e-12) numeric() else -c / b
  } else {
    disc <- b^2 - 4 * a * c
    if (disc < 0) numeric() else (-b + c(-1, 1) * sqrt(disc)) / (2 * a)
  }
  t <- roots[roots > 0 & roots < 1]
  (1 - t)^3 * p0 + 3 * (1 - t)^2 * t * p1 + 3 * (1 - t) * t^2 * p2 + t^3 * p3
}

# Where each segment starts: the previous segment's end, with Z ending at the
# subpath start.
path_ends <- function(path) {
  is_m <- path$cmd == "M"
  start_row <- cummax(ifelse(is_m, seq_along(path$cmd), 0L))
  is_z <- path$cmd == "Z"
  tibble::tibble(
    x = ifelse(is_z, path$x[pmax(start_row, 1L)], path$x),
    y = ifelse(is_z, path$y[pmax(start_row, 1L)], path$y)
  )
}

#' True extent of a path as c(x, y, w, h), cubic extrema included
#' @noRd
path_bbox <- function(path) {
  if (nrow(path) == 0L) return(rep(NA_real_, 4L))
  ends <- path_ends(path)
  prev_x <- c(NA_real_, utils::head(ends$x, -1L))
  prev_y <- c(NA_real_, utils::head(ends$y, -1L))
  is_c <- path$cmd == "C"
  curve <- function(p0, p1, p2, p3) {
    unlist(purrr::pmap(list(p0[is_c], p1[is_c], p2[is_c], p3[is_c]),
                       cubic_extrema))
  }
  xs <- c(path$x[path$cmd != "Z"],
          curve(prev_x, path$x1, path$x2, path$x))
  ys <- c(path$y[path$cmd != "Z"],
          curve(prev_y, path$y1, path$y2, path$y))
  c(min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))
}

#' TRUE when every subpath ends with Z or returns to its start point
#' @noRd
path_closed <- function(path) {
  if (nrow(path) == 0L) return(FALSE)
  sub <- cumsum(path$cmd == "M")
  closes <- function(rows) {
    last <- rows[nrow(rows), ]
    first <- rows[1, ]
    last$cmd == "Z" ||
      (nrow(rows) > 1L && isTRUE(abs(last$x - first$x) < 1e-6) &&
         isTRUE(abs(last$y - first$y) < 1e-6))
  }
  all(purrr::map_lgl(split(path, sub), closes))
}

#' The box c(x, y, w, h) of a path that is one closed axis-aligned rectangle
#' of straight segments, else NULL
#' @noRd
path_as_rect <- function(path, tol = 0.01) {
  if (nrow(path) == 0L || sum(path$cmd == "M") != 1L || any(path$cmd == "C") ||
      !path_closed(path)) {
    return(NULL)
  }
  pts <- path[path$cmd != "Z", ]
  # A start point repeated before Z is not a fifth corner.
  if (nrow(pts) > 1L && abs(pts$x[nrow(pts)] - pts$x[1]) <= tol &&
      abs(pts$y[nrow(pts)] - pts$y[1]) <= tol) {
    pts <- pts[-nrow(pts), ]
  }
  if (nrow(pts) != 4L) return(NULL)

  dx <- abs(pts$x[c(2:4, 1L)] - pts$x)
  dy <- abs(pts$y[c(2:4, 1L)] - pts$y)
  horizontal <- dy <= tol & dx > tol
  vertical <- dx <= tol & dy > tol
  # Edges must alternate horizontal and vertical.
  ok <- (all(horizontal[c(1, 3)]) && all(vertical[c(2, 4)])) ||
    (all(vertical[c(1, 3)]) && all(horizontal[c(2, 4)]))
  if (!ok) return(NULL)
  c(min(pts$x), min(pts$y), diff(range(pts$x)), diff(range(pts$y)))
}

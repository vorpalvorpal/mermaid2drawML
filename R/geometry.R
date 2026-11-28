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

mat_translate <- function(dx, dy) c(1, 0, 0, 1, dx, dy)

mat_scaling <- function(s) c(s, 0, 0, s, 0, 0)

#' Rotation by `deg` degrees clockwise on screen (y points down)
#' @noRd
mat_rotate <- function(deg) {
  r <- deg * pi / 180
  c(cos(r), sin(r), -sin(r), cos(r), 0, 0)
}

#' The matrix that applies `first`, then `then`
#' @noRd
mat_then <- function(first, then) {
  a <- then
  b <- first
  c(a[[1]] * b[[1]] + a[[3]] * b[[2]], a[[2]] * b[[1]] + a[[4]] * b[[2]],
    a[[1]] * b[[3]] + a[[3]] * b[[4]], a[[2]] * b[[3]] + a[[4]] * b[[4]],
    a[[1]] * b[[5]] + a[[3]] * b[[6]] + a[[5]],
    a[[2]] * b[[5]] + a[[4]] * b[[6]] + a[[6]])
}

#' Matrices composed in the order they apply
#' @noRd
mat_chain <- function(...) purrr::reduce(list(...), mat_then)

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

#' Command letters and their numbers
#'
#' @return NULL when `d` is not a path; else list(ops, args): the command
#'   letters as written and, for each, the numbers that follow it (arc flags
#'   are single characters and may be packed, so arcs have their own pattern).
#' @noRd
path_tokens <- function(d) {
  segs <- stringr::str_match_all(
    d, paste0("([", .cmd_chars, "])([^", .cmd_chars, "]*)"))[[1]]
  # A path must start with a moveto; anything else is not a path.
  if (nrow(segs) == 0L || !segs[1, 2] %in% c("M", "m")) return(NULL)
  ops <- segs[, 2]
  args <- lapply(stringr::str_extract_all(segs[, 3], .num_re), as.numeric)
  is_arc <- toupper(ops) == "A"
  if (any(is_arc)) {
    arcs <- stringr::str_match_all(segs[is_arc, 3], .arc_re)
    args[is_arc] <- lapply(arcs, function(m) as.numeric(t(m[, -1L, drop = FALSE])))
  }
  list(ops = ops, args = args)
}

#' One simple command per repeat of each command letter
#'
#' Repeated argument groups after M are implicit linetos; numbers left over
#' after the last whole group are ignored.
#' @return list(ops, args, start): the simple commands, all their numbers in
#'   one vector, and where each command's numbers begin (0-based).
#' @noRd
path_simple_commands <- function(tokens) {
  k <- unname(.arity[toupper(tokens$ops)])
  reps <- ifelse(k == 0L, 1L, lengths(tokens$args) %/% pmax(k, 1L))
  used <- purrr::map2(tokens$args, k * reps, function(a, n) a[seq_len(n)])
  ops <- rep(tokens$ops, reps)
  later <- sequence(reps) > 1L
  ops[later & ops == "M"] <- "L"
  ops[later & ops == "m"] <- "l"
  width <- rep(k, reps)
  list(ops = ops, args = unlist(used, use.names = FALSE),
       start = cumsum(c(0L, width))[seq_along(ops)])
}

# ---- arcs -------------------------------------------------------------------

#' An elliptical arc as segments
#'
#' Endpoint parameterisation to centre parameterisation (SVG 1.1 F.6.5), then
#' one cubic per at most 90 degrees.
#' @return A numeric matrix, columns x1, y1, x2, y2, x, y, one row per
#'   segment: cubics, or one line (NA control points) when a radius is zero;
#'   no rows when the arc starts where it ends.
#' @noRd
arc_segments <- function(x0, y0, rx, ry, phi_deg, fa, fs, x, y) {
  if (x0 == x && y0 == y) return(matrix(numeric(), 0L, 6L))
  rx <- abs(rx)
  ry <- abs(ry)
  if (rx == 0 || ry == 0) return(matrix(c(NA, NA, NA, NA, x, y), 1L, 6L))

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
  cbind(px(a1) + k * dpx(a1), py(a1) + k * dpy(a1),
        ex - k * dpx(a2), ey - k * dpy(a2), ex, ey)
}

# ---- walking the commands -----------------------------------------------------

#' Absolute M, L, C and Z segments from simple commands
#'
#' One loop over the commands, carrying the current point, the subpath start
#' and the last control point with its kind (for S and T reflection) in local
#' variables and writing segments into a preallocated matrix: a tibble per
#' command made long paths (Mermaid's hand-drawn outlines) slow.
#' @noRd
path_walk <- function(ops, args, start) {
  ups <- toupper(ops)
  # An arc is at most four cubics.
  size <- sum(ifelse(ups == "A", 4L, 1L))
  seg <- matrix(NA_real_, size, 6L)
  cmd <- character(size)
  j <- 0L
  x <- 0
  y <- 0
  sx <- 0
  sy <- 0
  cx <- 0
  cy <- 0
  ck <- ""
  for (i in seq_along(ops)) {
    up <- ups[[i]]
    v <- args[start[[i]] + seq_len(.arity[[up]])]
    rel <- ops[[i]] != up
    ox <- if (rel) x else 0
    oy <- if (rel) y else 0
    if (up == "Z") {
      j <- j + 1L
      cmd[[j]] <- "Z"
      x <- sx
      y <- sy
      cx <- x
      cy <- y
      ck <- ""
      next
    }
    if (up %in% c("M", "L", "H", "V")) {
      nx <- switch(up, V = x, v[[1]] + ox)
      ny <- switch(up, H = y, V = v[[1]] + oy, v[[2]] + oy)
      j <- j + 1L
      cmd[[j]] <- if (up == "M") "M" else "L"
      seg[j, 5:6] <- c(nx, ny)
      if (up == "M") {
        sx <- nx
        sy <- ny
      }
      x <- nx
      y <- ny
      cx <- x
      cy <- y
      ck <- ""
      next
    }
    if (up == "A") {
      nx <- v[[6]] + ox
      ny <- v[[7]] + oy
      rows <- arc_segments(x, y, v[[1]], v[[2]], v[[3]], v[[4]], v[[5]], nx, ny)
      k <- nrow(rows)
      if (k > 0L) {
        seg[j + seq_len(k), ] <- rows
        cmd[j + seq_len(k)] <- ifelse(is.na(rows[, 1]), "L", "C")
        j <- j + k
      }
      x <- nx
      y <- ny
      cx <- x
      cy <- y
      ck <- ""
      next
    }
    # Curves: C and S are cubics; Q and T are quadratics raised to cubics
    # with the same shape.
    reflect_x <- 2 * x - cx
    reflect_y <- 2 * y - cy
    if (up %in% c("C", "S")) {
      if (up == "C") {
        c1 <- c(v[[1]] + ox, v[[2]] + oy)
        v <- v[3:6]
      } else {
        c1 <- if (ck == "C") c(reflect_x, reflect_y) else c(x, y)
      }
      c2 <- c(v[[1]] + ox, v[[2]] + oy)
      end <- c(v[[3]] + ox, v[[4]] + oy)
      ck <- "C"
      ctrl <- c2
    } else {
      if (up == "Q") {
        q <- c(v[[1]] + ox, v[[2]] + oy)
        v <- v[3:4]
      } else {
        q <- if (ck == "Q") c(reflect_x, reflect_y) else c(x, y)
      }
      end <- c(v[[1]] + ox, v[[2]] + oy)
      c1 <- c(x, y) + 2 / 3 * (q - c(x, y))
      c2 <- end + 2 / 3 * (q - end)
      ck <- "Q"
      ctrl <- q
    }
    j <- j + 1L
    cmd[[j]] <- "C"
    seg[j, ] <- c(c1, c2, end)
    x <- end[[1]]
    y <- end[[2]]
    cx <- ctrl[[1]]
    cy <- ctrl[[2]]
  }
  keep <- seq_len(j)
  tibble::tibble(cmd = cmd[keep], x1 = seg[keep, 1], y1 = seg[keep, 2],
                 x2 = seg[keep, 3], y2 = seg[keep, 4], x = seg[keep, 5],
                 y = seg[keep, 6])
}

#' Parse SVG path data into absolute M, L, C and Z segments
#'
#' @param d Path data string.
#' @return A tibble with columns `cmd`, `x1`, `y1`, `x2`, `y2`, `x`, `y`;
#'   zero rows when `d` is empty or not a path.
#' @noRd
path_parse <- function(d) {
  if (length(d) != 1L || is.na(d)) return(path_tbl())
  tokens <- path_tokens(d)
  if (is.null(tokens)) return(path_tbl())
  cmds <- path_simple_commands(tokens)
  if (length(cmds$ops) == 0L) return(path_tbl())
  path_walk(cmds$ops, cmds$args, cmds$start)
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

# Where each segment ends, with Z ending at the subpath start: list(x, y)
# (a plain list; this is called for every path end and box).
path_ends <- function(path) {
  is_m <- path$cmd == "M"
  start_row <- cummax(ifelse(is_m, seq_along(path$cmd), 0L))
  is_z <- path$cmd == "Z"
  list(
    x = ifelse(is_z, path$x[pmax(start_row, 1L)], path$x),
    y = ifelse(is_z, path$y[pmax(start_row, 1L)], path$y)
  )
}

#' Where a path starts or ends and the direction it runs there
#'
#' The direction is towards the first control or end point that differs from
#' the end (at the start) or from the last such point to the end (at the
#' end), as SVG orients markers.
#' @param at "start" or "end".
#' @return list(point = c(x, y), dir = unit c(dx, dy), NA when the path has
#'   no length at that end).
#' @noRd
path_end_frame <- function(path, at) {
  n <- nrow(path)
  ends <- path_ends(path)
  pt <- function(k) c(ends$x[[k]], ends$y[[k]])
  ctrl <- function(k, which) c(path[[paste0("x", which)]][[k]], path[[paste0("y", which)]][[k]])
  is_c <- function(k) identical(path$cmd[[k]], "C")
  if (at == "start") {
    point <- pt(1L)
    toward <- if (n < 2L) list() else if (is_c(2L)) list(ctrl(2L, 1), ctrl(2L, 2), pt(2L)) else list(pt(2L))
    away <- purrr::map(toward, function(q) q - point)
  } else {
    point <- pt(n)
    from <- if (n < 2L) list() else if (is_c(n)) list(ctrl(n, 2), ctrl(n, 1), pt(n - 1L)) else list(pt(n - 1L))
    away <- purrr::map(from, function(q) point - q)
  }
  d <- purrr::detect(away, function(v) isTRUE(sqrt(sum(v^2)) > 1e-9))
  list(point = point, dir = if (is.null(d)) c(NA_real_, NA_real_) else d / sqrt(sum(d^2)))
}

#' Move a path's start or end point by vector `v`; the control point next to
#' it moves with it, so a cubic keeps its direction there
#' @noRd
path_move_end <- function(path, at, v) {
  n <- nrow(path)
  if (at == "start") {
    path$x[[1]] <- path$x[[1]] + v[[1]]
    path$y[[1]] <- path$y[[1]] + v[[2]]
    if (n >= 2L && path$cmd[[2]] == "C") {
      path$x1[[2]] <- path$x1[[2]] + v[[1]]
      path$y1[[2]] <- path$y1[[2]] + v[[2]]
    }
    return(path)
  }
  # A closed path has no free end.
  if (path$cmd[[n]] == "Z") return(path)
  path$x[[n]] <- path$x[[n]] + v[[1]]
  path$y[[n]] <- path$y[[n]] + v[[2]]
  if (path$cmd[[n]] == "C") {
    path$x2[[n]] <- path$x2[[n]] + v[[1]]
    path$y2[[n]] <- path$y2[[n]] + v[[2]]
  }
  path
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

# ---- recognising figures -----------------------------------------------------

#' The drawn segments of a path: where each starts and ends and its midpoint
#' (a cubic evaluated at t = 0.5); a Z that closes no gap draws nothing
#' @noRd
path_segments <- function(path) {
  ends <- path_ends(path)
  x0 <- c(NA_real_, utils::head(ends$x, -1L))
  y0 <- c(NA_real_, utils::head(ends$y, -1L))
  is_c <- path$cmd == "C"
  mid <- function(p0, p1, p2, p3) ifelse(is_c, (p0 + 3 * p1 + 3 * p2 + p3) / 8, (p0 + p3) / 2)
  seg <- tibble::tibble(cmd = path$cmd, x0 = x0, y0 = y0, x = ends$x, y = ends$y,
                        mx = mid(x0, path$x1, path$x2, ends$x),
                        my = mid(y0, path$y1, path$y2, ends$y))
  gap <- abs(seg$x - seg$x0) + abs(seg$y - seg$y0)
  seg[seg$cmd != "M" & !(seg$cmd == "Z" & gap <= 1e-9), ]
}

# One closed subpath: the only kind of path a preset can stand for.
single_closed <- function(path) {
  nrow(path) > 0L && sum(path$cmd == "M") == 1L && path_closed(path)
}

#' TRUE when a closed path of at least four segments has every segment end
#' and midpoint within `tol` of the ellipse inscribed in `box`
#'
#' The distance is measured along the ray from the centre, which for the
#' near-circles this is meant for is the distance to the curve.
#' @noRd
path_on_ellipse <- function(path, box, tol) {
  if (!single_closed(path)) return(FALSE)
  seg <- path_segments(path)
  a <- box[[3]] / 2
  b <- box[[4]] / 2
  if (nrow(seg) < 4L || a <= 0 || b <= 0) return(FALSE)
  dx <- c(seg$x, seg$mx) - (box[[1]] + a)
  dy <- c(seg$y, seg$my) - (box[[2]] + b)
  r <- sqrt((dx / a)^2 + (dy / b)^2)
  dist <- ifelse(r > 0, abs(1 - 1 / r) * sqrt(dx^2 + dy^2), min(a, b))
  all(dist <= tol)
}

#' The common corner radius of a closed path that is a rectangle with four
#' equal round corners, or NA
#'
#' Every segment end and midpoint lies on the box boundary, except inside the
#' four corner squares of side r, where it stays outside the square inscribed
#' in the corner circle (an arc passes its corner, a straight snip cuts
#' through it). r is where the straight sides stop, the same at all four
#' corners along both sides, and some straight side remains.
#' @noRd
path_corner_radius <- function(path, box, tol) {
  if (!single_closed(path)) return(NA_real_)
  seg <- path_segments(path)
  px <- c(seg$x, seg$mx)
  py <- c(seg$y, seg$my)
  x0 <- box[[1]]
  y0 <- box[[2]]
  x1 <- x0 + box[[3]]
  y1 <- y0 + box[[4]]
  on <- list(left = abs(px - x0) <= tol, right = abs(px - x1) <= tol,
             top = abs(py - y0) <= tol, bottom = abs(py - y1) <= tol)
  ends_along <- function(v, keep, lo, hi) {
    if (!any(keep)) return(c(Inf, Inf))
    c(min(v[keep]) - lo, hi - max(v[keep]))
  }
  rs <- c(ends_along(px, on$top, x0, x1), ends_along(px, on$bottom, x0, x1),
          ends_along(py, on$left, y0, y1), ends_along(py, on$right, y0, y1))
  r <- mean(rs)
  # With no straight side at all the figure is a circle or worse, which is
  # the ellipse rule's to decide.
  if (!all(is.finite(rs)) || diff(range(rs)) > tol || r <= tol ||
      r > min(box[[3]], box[[4]]) / 2 + tol || r >= max(box[[3]], box[[4]]) / 2 - tol) {
    return(NA_real_)
  }
  boundary <- on$left | on$right | on$top | on$bottom
  # Corner circle centres, with the direction that points at their corner.
  corners <- tibble::tibble(cx = c(x0 + r, x1 - r, x1 - r, x0 + r),
                            cy = c(y0 + r, y0 + r, y1 - r, y1 - r),
                            sx = c(-1, 1, 1, -1), sy = c(-1, -1, 1, 1))
  half <- r / sqrt(2)
  in_corner <- purrr::pmap(corners, function(cx, cy, sx, sy) {
    ux <- sx * (px - cx)
    uy <- sy * (py - cy)
    ux >= -tol & ux <= r + tol & uy >= -tol & uy <= r + tol &
      (ux >= half - tol | uy >= half - tol)
  })
  if (all(boundary | Reduce(`|`, in_corner))) r else NA_real_
}

#' Every end point and segment midpoint of a path, as a tibble x, y
#' @noRd
path_points <- function(path) {
  seg <- path_segments(path)
  starts <- path[path$cmd == "M", c("x", "y")]
  tibble::tibble(x = c(starts$x, seg$x, seg$mx), y = c(starts$y, seg$y, seg$my))
}

# Distance of points from the boundary of box c(x, y, w, h).
box_boundary_dist <- function(x, y, box) {
  x0 <- box[[1]]
  y0 <- box[[2]]
  x1 <- x0 + box[[3]]
  y1 <- y0 + box[[4]]
  # Beyond the box on either axis, plus the way to the nearest side inside it.
  out <- sqrt(pmax(0, x0 - x, x - x1)^2 + pmax(0, y0 - y, y - y1)^2)
  inside <- pmin(abs(x - x0), abs(x - x1), abs(y - y0), abs(y - y1))
  ifelse(out > 0, out, inside)
}

# Distance of points from the ellipse inscribed in a box, along the ray from
# its centre (see path_on_ellipse()).
ellipse_dist <- function(x, y, box) {
  a <- box[[3]] / 2
  b <- box[[4]] / 2
  if (a <= 0 || b <= 0) return(rep(Inf, length(x)))
  dx <- x - (box[[1]] + a)
  dy <- y - (box[[2]] + b)
  r <- sqrt((dx / a)^2 + (dy / b)^2)
  ifelse(r > 0, abs(1 - 1 / r) * sqrt(dx^2 + dy^2), min(a, b))
}

# TRUE for points inside one of the four corner squares of side r of a box.
in_corner_squares <- function(x, y, box, r) {
  near_x <- x <= box[[1]] + r | x >= box[[1]] + box[[3]] - r
  near_y <- y <= box[[2]] + r | y >= box[[2]] + box[[4]] - r
  near_x & near_y
}

# How far points may be from each preset's figure; presets not listed are
# figures no path is compared with.
.figure_dist <- list(
  rect = function(x, y, box, r) box_boundary_dist(x, y, box),
  ellipse = function(x, y, box, r) ellipse_dist(x, y, box),
  roundRect = function(x, y, box, r) {
    r <- if (is.na(r)) 0 else r
    ifelse(in_corner_squares(x, y, box, r), 0, box_boundary_dist(x, y, box))
  }
)

#' TRUE when every end point and midpoint of `path` lies within `tol` of the
#' figure the preset `prst` draws in `box` (`r`: a roundRect's corner radius)
#' @noRd
path_on_figure <- function(path, prst, box, tol, r = NA_real_) {
  dist <- .figure_dist[[prst]]
  if (is.null(dist) || is.null(path) || nrow(path) == 0L) return(FALSE)
  pts <- path_points(path)
  all(dist(pts$x, pts$y, box, r) <= tol)
}

#' TRUE when box `inner` lies inside the ellipse inscribed in box `outer`
#' @noRd
box_in_ellipse <- function(inner, outer) {
  a <- outer[[3]] / 2
  b <- outer[[4]] / 2
  if (a <= 0 || b <= 0) return(FALSE)
  cx <- c(inner[[1]], inner[[1]] + inner[[3]]) - (outer[[1]] + a)
  cy <- c(inner[[2]], inner[[2]] + inner[[4]]) - (outer[[2]] + b)
  all(outer(cx^2 / a^2, cy^2 / b^2, `+`) <= 1)
}

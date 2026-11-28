# Page layout: margins, page size, orientation choice and section breaks.
# Offline: mermaid_to_word() is mocked, so no browser or Node is needed.

# ── Helpers ─────────────────────────────────────────────────────────────────

w_ns <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

fake_xml <- function(text = "DIAGRAM") {
  sprintf('<w:p xmlns:w="%s"><w:r><w:t>%s</w:t></w:r></w:p>', w_ns, text)
}

# Make mermaid_to_word() return a canned diagram; records the start_id it gets.
mock_diagram <- function(orientation = "landscape", margins = NULL,
                         seen = NULL, env = parent.frame()) {
  testthat::local_mocked_bindings(
    mermaid_to_word = function(mermaid_code, start_id = 1L, ...) {
      if (!is.null(seen)) seen$start_id <- start_id
      structure(
        list(xml = fake_xml(), next_id = 5L,
             orientation = orientation, margins = margins),
        class = "mermaid_word_diagram")
    },
    .env = env
  )
}

# Parse XML text with every namespace declaration and prefix removed, so the
# helpers below can use plain element names. (xml2::xml_ns_strip() removes
# only a default namespace, and Word's XML is all prefixed.)
strip_all <- function(s) {
  s <- gsub(' xmlns(:[A-Za-z0-9]+)?="[^"]*"', "", s)
  s <- gsub("<(/?)[A-Za-z0-9]+:", "<\\1", s)
  s <- gsub(" [A-Za-z0-9]+:([A-Za-z]+=)", " \\1", s)
  xml2::read_xml(s)
}

# A namespace-stripped copy of the body XML of an rdocx, or of a saved file.
body_of <- function(doc) {
  strip_all(as.character(officer::docx_body_xml(doc)))
}

saved_body_of <- function(doc) {
  path <- withr::local_tempfile(fileext = ".docx", .local_envir = parent.frame())
  print(doc, target = path)
  dir <- withr::local_tempdir(.local_envir = parent.frame())
  utils::unzip(path, files = "word/document.xml", exdir = dir)
  strip_all(paste(readLines(file.path(dir, "word", "document.xml"), warn = FALSE),
                  collapse = "\n"))
}

n_sect <- function(x) length(xml2::xml_find_all(x, "//sectPr"))

# Attribute value ignoring any namespace prefix on the attribute name.
attr_of <- function(node, name) {
  if (inherits(node, "xml_missing") || length(node) == 0L) return(NA_character_)
  a <- xml2::xml_attrs(node)
  names(a) <- sub("^.*:", "", names(a))
  if (name %in% names(a)) unname(a[[name]]) else NA_character_
}

# Direct children of w:body, each described as a list(kind, text, sect).
# kind: "sect" (paragraph carrying a section break), "diagram", "text",
# "final" (the body-level sectPr).
body_items <- function(x) {
  kids <- xml2::xml_children(xml2::xml_find_first(x, "//body"))
  lapply(kids, function(k) {
    nm <- xml2::xml_name(k)
    sect <- xml2::xml_find_first(k, "./pPr/sectPr")
    txt <- paste(xml2::xml_text(xml2::xml_find_all(k, ".//t")), collapse = "")
    if (nm == "sectPr") {
      list(kind = "final", text = "", sect = k)
    } else if (!inherits(sect, "xml_missing")) {
      list(kind = "sect", text = txt, sect = sect)
    } else if (grepl("DIAGRAM", txt, fixed = TRUE)) {
      list(kind = "diagram", text = txt, sect = NULL)
    } else {
      list(kind = "text", text = txt, sect = NULL)
    }
  })
}

kinds <- function(items) vapply(items, `[[`, "", "kind")
texts <- function(items) vapply(items, `[[`, "", "text")

pg_dims <- function(sect) {
  pg <- xml2::xml_find_first(sect, "./pgSz")
  c(w = as.numeric(attr_of(pg, "w")), h = as.numeric(attr_of(pg, "h")),
    landscape = identical(attr_of(pg, "orient"), "landscape"))
}

is_landscape_sect <- function(sect) {
  d <- pg_dims(sect)
  isTRUE(d[["landscape"]] == 1) ||
    (!is.na(d[["w"]]) && !is.na(d[["h"]]) && d[["w"]] > d[["h"]])
}

pg_mar <- function(sect) {
  m <- xml2::xml_find_first(sect, "./pgMar")
  vapply(c("top", "bottom", "left", "right"),
         function(s) as.numeric(attr_of(m, s)), numeric(1))
}

# The section (a `sect` node) that the item at index `i` belongs to: the first
# section break at or after it.
section_of <- function(items, i) {
  for (j in seq(i, length(items))) {
    if (items[[j]]$kind %in% c("sect", "final")) return(items[[j]]$sect)
  }
  NULL
}

doc_two_pars <- function() {
  officer::read_docx() |>
    officer::body_add_par("one") |>
    officer::body_add_par("two")
}

# ── .normalise_margins() ────────────────────────────────────────────────────

test_that(".normalise_margins() returns NULL for NULL", {
  expect_null(.normalise_margins(NULL))
})

test_that(".normalise_margins() spreads a single unnamed number to all sides", {
  expect_equal(.normalise_margins(0.5),
               list(top = 0.5, bottom = 0.5, left = 0.5, right = 0.5))
})

test_that(".normalise_margins() fills missing sides of a named vector with 1", {
  expect_equal(.normalise_margins(c(left = 0.25)),
               list(top = 1, bottom = 1, left = 0.25, right = 1))
})

test_that(".normalise_margins() always orders sides top, bottom, left, right", {
  res <- .normalise_margins(c(right = 0.1, left = 0.2, bottom = 0.3, top = 0.4))
  expect_named(res, c("top", "bottom", "left", "right"))
  expect_equal(unlist(res, use.names = FALSE), c(0.4, 0.3, 0.2, 0.1))
})

test_that(".normalise_margins() accepts a named list of numbers", {
  res <- .normalise_margins(list(bottom = 2, top = 0.5))
  expect_named(res, c("top", "bottom", "left", "right"))
  expect_equal(unlist(res, use.names = FALSE), c(0.5, 2, 1, 1))
})

test_that(".normalise_margins() accepts a single named value", {
  expect_equal(.normalise_margins(c(top = 0.7)),
               list(top = 0.7, bottom = 1, left = 1, right = 1))
})

test_that(".normalise_margins() accepts zero margins", {
  expect_equal(.normalise_margins(0),
               list(top = 0, bottom = 0, left = 0, right = 0))
})

test_that(".normalise_margins() rejects non-numeric input", {
  expect_error(.normalise_margins("wide"), class = "mermaid2drawml_bad_margins")
  expect_error(.normalise_margins(TRUE), class = "mermaid2drawml_bad_margins")
})

test_that(".normalise_margins() rejects an unnamed vector longer than one", {
  expect_error(.normalise_margins(c(1, 2)), class = "mermaid2drawml_bad_margins")
  expect_error(.normalise_margins(c(1, 2, 3, 4)),
               class = "mermaid2drawml_bad_margins")
})

test_that(".normalise_margins() rejects partly named input", {
  expect_error(.normalise_margins(c(top = 1, 2)),
               class = "mermaid2drawml_bad_margins")
})

test_that(".normalise_margins() rejects unknown names and names them", {
  expect_error(.normalise_margins(c(top = 1, sideways = 2, middle = 3)),
               class = "mermaid2drawml_bad_margins")
  expect_error(.normalise_margins(c(top = 1, sideways = 2)), "sideways")
})

test_that(".normalise_margins() rejects duplicated names", {
  expect_error(.normalise_margins(c(top = 1, top = 2)),
               class = "mermaid2drawml_bad_margins")
})

test_that(".normalise_margins() rejects NA values", {
  expect_error(.normalise_margins(NA_real_), class = "mermaid2drawml_bad_margins")
  expect_error(.normalise_margins(c(top = 1, left = NA)),
               class = "mermaid2drawml_bad_margins")
})

test_that(".normalise_margins() rejects negative values", {
  expect_error(.normalise_margins(-0.5), class = "mermaid2drawml_bad_margins")
  expect_error(.normalise_margins(c(top = 1, right = -1)),
               class = "mermaid2drawml_bad_margins")
})

test_that(".normalise_margins() raises cli (rlang) errors", {
  expect_error(.normalise_margins(-1), class = "rlang_error")
})

# ── .page_dims() ────────────────────────────────────────────────────────────

test_that(".page_dims() gives A4 portrait in inches", {
  expect_equal(.page_dims("portrait"), c(width = 8.27, height = 11.69))
})

test_that(".page_dims() gives A4 landscape in inches", {
  expect_equal(.page_dims("landscape"), c(width = 11.69, height = 8.27))
})

# ── .content_size() ─────────────────────────────────────────────────────────

test_that(".content_size() with NULL margins uses 1 inch on every side", {
  expect_equal(.content_size("portrait", NULL),
               c(width = 6.27, height = 9.69))
  expect_equal(.content_size("landscape", NULL),
               c(width = 9.69, height = 6.27))
})

test_that(".content_size() subtracts the supplied margins from the page", {
  m <- .normalise_margins(c(top = 0.5, bottom = 1.5, left = 0.25, right = 0.75))
  expect_equal(.content_size("portrait", m),
               c(width = 8.27 - 1, height = 11.69 - 2))
  expect_equal(.content_size("landscape", m),
               c(width = 11.69 - 1, height = 8.27 - 2))
})

test_that(".content_size() errors when margins leave no width", {
  m <- list(top = 1, bottom = 1, left = 5, right = 5)
  expect_error(.content_size("portrait", m),
               class = "mermaid2drawml_bad_margins")
})

test_that(".content_size() errors when margins leave no height", {
  m <- list(top = 6, bottom = 6, left = 1, right = 1)
  expect_error(.content_size("portrait", m),
               class = "mermaid2drawml_bad_margins")
  m2 <- list(top = 4, bottom = 4.27, left = 1, right = 1)  # exactly zero
  expect_error(.content_size("landscape", m2),
               class = "mermaid2drawml_bad_margins")
})

# ── .resolve_orientation() ──────────────────────────────────────────────────

test_that(".resolve_orientation() returns explicit choices as given", {
  expect_equal(.resolve_orientation("portrait", 1200, 400, portrait_width_in = 6.27),
               "portrait")
  expect_equal(.resolve_orientation("landscape", 100, 400, portrait_width_in = 6.27),
               "landscape")
})

test_that(".resolve_orientation() keeps a small wide diagram in portrait", {
  # 300 px = 3.1 in: aspect 3 but fits a portrait page easily
  expect_equal(.resolve_orientation("auto", 300, 100, max_scale = 1,
                                    portrait_width_in = 6.27),
               "portrait")
})

test_that(".resolve_orientation() turns a wide, large diagram landscape", {
  expect_equal(.resolve_orientation("auto", 1200, 400, max_scale = 1,
                                    portrait_width_in = 6.27),
               "landscape")
})

test_that(".resolve_orientation() keeps a large but squarish diagram portrait", {
  expect_equal(.resolve_orientation("auto", 1200, 1100, max_scale = 1,
                                    portrait_width_in = 6.27),
               "portrait")
})

test_that(".resolve_orientation() with max_scale = Inf only needs the aspect", {
  expect_equal(.resolve_orientation("auto", 300, 100, max_scale = Inf,
                                    portrait_width_in = 6.27),
               "landscape")
})

test_that(".resolve_orientation() honours a custom auto_aspect", {
  # 1200 x 960 px: aspect 1.25, and 12.5 in wide, so only the aspect decides
  expect_equal(.resolve_orientation("auto", 1200, 960, auto_aspect = 1.3,
                                    max_scale = 1, portrait_width_in = 6.27),
               "portrait")
  expect_equal(.resolve_orientation("auto", 1200, 960, auto_aspect = 1.2,
                                    max_scale = 1, portrait_width_in = 6.27),
               "landscape")
})

test_that(".resolve_orientation() uses max_scale in the width test", {
  # 480 px = 5 in: fits 6.27 in at scale 1, but at scale 2 is 10 in wide
  expect_equal(.resolve_orientation("auto", 480, 100, max_scale = 1,
                                    portrait_width_in = 6.27),
               "portrait")
  expect_equal(.resolve_orientation("auto", 480, 100, max_scale = 2,
                                    portrait_width_in = 6.27),
               "landscape")
})

test_that(".resolve_orientation() gives portrait for zero or missing height", {
  expect_equal(.resolve_orientation("auto", 1200, 0, max_scale = 1,
                                    portrait_width_in = 6.27),
               "portrait")
  expect_equal(.resolve_orientation("auto", 1200, NA_real_, max_scale = 1,
                                    portrait_width_in = 6.27),
               "portrait")
})

# ── body_add_mermaid(): sections ────────────────────────────────────────────

test_that("a portrait diagram without custom margins adds no section properties", {
  mock_diagram("portrait")
  doc <- doc_two_pars()
  before <- n_sect(body_of(doc))
  doc <- body_add_mermaid(doc, "flowchart TB\n A --> B")
  expect_equal(n_sect(body_of(doc)), before)
  expect_true("DIAGRAM" %in% texts(body_items(body_of(doc))))
})

test_that("a landscape diagram closes the earlier content in a portrait section", {
  mock_diagram("landscape")
  doc <- doc_two_pars() |> body_add_mermaid("flowchart LR\n A --> B")
  items <- body_items(body_of(doc))
  k <- kinds(items)

  i_diagram <- which(k == "diagram")
  sects <- which(k == "sect")
  expect_length(i_diagram, 1L)
  expect_length(sects, 2L)

  before <- sects[sects < i_diagram]
  after  <- sects[sects > i_diagram]
  expect_length(before, 1L)
  expect_length(after, 1L)

  # "one" and "two" precede the closing break
  expect_true(all(c("one", "two") %in% texts(items)[seq_len(before - 1L)]))
  # the closing break is not landscape
  expect_false(is_landscape_sect(items[[before]]$sect))
  # the break after the diagram carries the landscape layout
  expect_true(is_landscape_sect(items[[after]]$sect))
  d <- pg_dims(items[[after]]$sect)
  expect_equal(d[["landscape"]], 1)
  expect_gt(d[["w"]], d[["h"]])
})

test_that("in the saved file earlier paragraphs stay in a portrait section", {
  mock_diagram("landscape")
  doc <- doc_two_pars() |> body_add_mermaid("flowchart LR\n A --> B")
  items <- body_items(saved_body_of(doc))
  txt <- texts(items)

  for (p in c("one", "two")) {
    sect <- section_of(items, which(txt == p)[1])
    expect_false(is_landscape_sect(sect), info = p)
  }
  # and the diagram itself is in a landscape section
  expect_true(is_landscape_sect(section_of(items, which(kinds(items) == "diagram"))))
})

test_that("content added after a landscape diagram is outside its section", {
  mock_diagram("landscape")
  doc <- doc_two_pars() |>
    body_add_mermaid("flowchart LR\n A --> B") |>
    officer::body_add_par("three")
  items <- body_items(body_of(doc))
  k <- kinds(items)
  txt <- texts(items)

  i_three <- which(txt == "three")
  i_diagram <- which(k == "diagram")
  landscape_break <- which(k == "sect" &
    vapply(items, function(it) !is.null(it$sect) && is_landscape_sect(it$sect),
           logical(1)))
  expect_length(landscape_break, 1L)
  expect_gt(landscape_break, i_diagram)
  expect_gt(i_three, landscape_break)
})

test_that("a diagram first in an empty document adds exactly one section break", {
  mock_diagram("landscape")
  doc <- officer::read_docx()
  before <- length(xml2::xml_find_all(body_of(doc), "//body/p/pPr/sectPr"))
  doc <- body_add_mermaid(doc, "flowchart LR\n A --> B")
  x <- body_of(doc)
  after <- length(xml2::xml_find_all(x, "//body/p/pPr/sectPr"))
  expect_equal(after - before, 1L)

  # no section break comes before the diagram, so there is no blank first page
  k <- kinds(body_items(x))
  expect_lt(which(k == "diagram"), min(which(k == "sect")))
})

test_that("custom margins with portrait put the margins on the trailing section", {
  m <- list(top = 0.5, bottom = 0.5, left = 0.25, right = 0.25)
  mock_diagram("portrait", margins = m)
  doc <- doc_two_pars() |>
    body_add_mermaid("flowchart TB\n A --> B", page_margins_in = 0.25)
  items <- body_items(body_of(doc))
  k <- kinds(items)

  i_diagram <- which(k == "diagram")
  sects <- which(k == "sect")
  before <- sects[sects < i_diagram]
  after  <- sects[sects > i_diagram]
  expect_length(before, 1L)
  expect_length(after, 1L)

  # trailing section: custom margins in twips
  mar <- pg_mar(items[[after]]$sect)
  expect_equal(unname(mar[c("top", "bottom")]), c(720, 720))
  expect_equal(unname(mar[c("left", "right")]), c(360, 360))
  # and it is not landscape
  expect_false(is_landscape_sect(items[[after]]$sect))

  # preceding section keeps the document's own margins
  prev <- pg_mar(items[[before]]$sect)
  expect_false(isTRUE(all(prev == mar)))
  expect_false(isTRUE(unname(prev[["left"]]) == 360))
})

test_that("custom margin values reach pgMar in twips (inches x 1440)", {
  m <- list(top = 1.5, bottom = 0.75, left = 0.5, right = 2)
  mock_diagram("landscape", margins = m)
  doc <- doc_two_pars() |> body_add_mermaid("flowchart LR\n A --> B")
  items <- body_items(saved_body_of(doc))
  after <- which(kinds(items) == "sect")
  after <- after[after > which(kinds(items) == "diagram")]
  mar <- pg_mar(items[[after[1]]]$sect)
  expect_equal(unname(mar), c(2160, 1080, 720, 2880))
  expect_true(is_landscape_sect(items[[after[1]]]$sect))
})

# ── body_add_mermaid(): start_id ────────────────────────────────────────────

test_that("start_id = NULL is 1 for an empty document", {
  seen <- new.env()
  mock_diagram("portrait", seen = seen)
  body_add_mermaid(officer::read_docx(), "flowchart TB\n A --> B")
  expect_equal(as.integer(seen$start_id), 1L)
})

test_that("start_id = NULL is one more than the largest existing shape id", {
  seen <- new.env()
  mock_diagram("portrait", seen = seen)
  doc <- officer::read_docx()
  with_ids <- paste0(
    '<w:p xmlns:w="', w_ns, '" ',
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">',
    '<w:r><w:drawing><wp:inline><wp:docPr id="7" name="a"/></wp:inline>',
    '</w:drawing></w:r></w:p>')
  doc <- officer::body_add_xml(doc, with_ids)
  body_add_mermaid(doc, "flowchart TB\n A --> B")
  expect_equal(as.integer(seen$start_id), 8L)
})

test_that("an explicit start_id is passed through unchanged", {
  seen <- new.env()
  mock_diagram("portrait", seen = seen)
  body_add_mermaid(officer::read_docx(), "flowchart TB\n A --> B", start_id = 42L)
  expect_equal(as.integer(seen$start_id), 42L)
})

# ── Orientation "auto" from the size and scale actually used ────────────────

# A diagram 700 px wide and 200 high: 7.3 in at natural size, too wide for a
# portrait page's 6.27 in; 3.6 in at half size, which fits.
wide_700 <- function() {
  generic_render(list(prim_rect(0, 0, 700, 200, fill = "#ECECFF")),
                 viewBox = c(0, 0, 700, 200))
}

emu_per_in <- 914400

test_that("auto stays portrait when max_scale makes the diagram fit, else turns landscape", {
  local_mock_render(wide_700())
  half <- mermaid_to_word("graph LR", max_scale = 0.5)
  full <- mermaid_to_word("graph LR", max_scale = 1)
  expect_equal(half$orientation, "portrait")
  expect_equal(full$orientation, "landscape")
  # The drawing is the size that decision assumed, and fits its page.
  expect_near(inline_extent(dml_doc(half$xml))[["cx"]], 700 * 0.5 * 9525, tol = 1)
  expect_lte(inline_extent(dml_doc(half$xml))[["cx"]],
             .content_size("portrait", NULL)[["width"]] * emu_per_in)
  expect_near(inline_extent(dml_doc(full$xml))[["cx"]], 700 * 9525, tol = 1)
  expect_lte(inline_extent(dml_doc(full$xml))[["cx"]],
             .content_size("landscape", NULL)[["width"]] * emu_per_in)
})

test_that("body_add_mermaid() decides auto with its max_scale", {
  local_mock_render(wide_700())
  doc <- doc_two_pars()
  before <- n_sect(body_of(doc))
  half <- body_add_mermaid(doc_two_pars(), "graph LR", max_scale = 0.5)
  expect_equal(n_sect(body_of(half)), before)
  full <- body_add_mermaid(doc_two_pars(), "graph LR", max_scale = 1)
  sects <- xml2::xml_find_all(body_of(full), "//sectPr")
  expect_true(any(vapply(as.list(sects), is_landscape_sect, logical(1))))
})

test_that("with page_width_in and page_height_in given, auto adds no section", {
  # An explicit page size is the space the caller has, so "auto" has
  # nothing to decide: the drawing is fitted to that size and the page is
  # left as it is (portrait, no section of its own).
  local_mock_render(wide_700())
  d <- mermaid_to_word("graph LR", page_width_in = 6, page_height_in = 8)
  expect_equal(d$orientation, "portrait")
  expect_lte(inline_extent(dml_doc(d$xml))[["cx"]], 6 * emu_per_in + 1)

  doc <- doc_two_pars()
  before <- n_sect(body_of(doc))
  doc <- body_add_mermaid(doc, "graph LR", page_width_in = 6, page_height_in = 8)
  expect_equal(n_sect(body_of(doc)), before)
})

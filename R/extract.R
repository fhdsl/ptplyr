#' Extract Speaker Notes from a PowerPoint
#'
#' @param file Character. Path for `PPTX` file
#' @param ... additional arguments to pass to `extract_slide_note()`
#' particularly \code{xpath}
#'
#' @return Either a character vector or `NULL`
#' @export
#'
#' @importFrom utils unzip
#' @examples
#' ex_file = system.file("extdata", "example.pptx",
#' package = "ariExtra")
#' extract_notes(ex_file)
extract_notes = function(file, ...) {
  df = extract_slide_note(file, ...)
  if (is.null(df)) {
    return(NULL)
  }
  # need factor because they can be dumb with characters
  # and numerics and the file naming of PPTX files
  fac = basename(df$file)
  fac = factor(fac, levels = unique(fac))
  ss = split(df, fac)
  res = sapply(ss, function(x) {
    paste(x$text, collapse = " ")
  })
  if (any(trimws(res) %in% "")) {
    warning("Slides with no notes exists")
  }
  res[ res == ""] = ";"
  names(res) <- NULL
  return(res)
}

# Create a dataframe of slide xml files and the text on the slides
extract_slide_text = function(file, ...) {
  L = unzip_pptx(file)
  slides = L$slides

  if (length(slides) > 0) {
    # in case empty notes
    res = lapply(slides, function(x) {
      xx = extract_notes_xml(x, collapse_text = FALSE, ...)
      if (length(xx) == 0) {
        return(NULL)
      }
      snum = sub("[.]xml", "", sub("slide", "", basename(x)))
      snum = as.numeric(snum)
      data.frame(
        file = x,
        slide = snum,
        text = xx,
        index = 1:length(xx),
        stringsAsFactors = FALSE)
    })
    res = do.call(rbind, res)
    return(res)
  } else {
    return(NULL)
  }
}

# Create a dataframe of slide xml files and the respective notes under the slides
extract_slide_note = function(file, ...) {
  L = unzip_pptx(file)
  notes = L$notes
  slides = L$slides
  note_dir = L$note_dir

  if (length(notes) > 0) {
    # in case empty notes
    assoc_notes = sub("slide", "", basename(slides))
    assoc_notes = paste0("notesSlide", assoc_notes)
    assoc_notes = file.path(note_dir, assoc_notes)
    no_fe = !file.exists(assoc_notes)
    if (any(no_fe)) {
      file.create(assoc_notes[no_fe])
      notes = assoc_notes
    }
    res = lapply(notes, function(x) {
      if (file.size(x) == 0) {
        xx = ""
      } else {
        xx = extract_notes_xml(x, collapse_text = FALSE, ...)
      }
      if (length(xx) == 0) {
        xx = ""
      }
      snum = sub("[.]xml", "", sub("notesSlide", "", basename(x)))
      snum = as.numeric(snum)
      data.frame(
        file = x,
        slide = snum,
        text = xx,
        index = 1:length(xx),
        stringsAsFactors = FALSE)
    })
    res = do.call(rbind, res)
    return(res)
  } else {
    return(NULL)
  }
}

# Reorder XML files
reorder_xml = function(files) {
  if (length(files) == 0) {
    return(files)
  }
  nums = basename(files)
  nums = sub("[[:alpha:]]*(\\d.*)[.].*", "\\1", nums)
  nums = as.numeric(nums)
  if (any(is.na(nums))) {
    warning(paste0("Trying to parse set of files (example: ", files[1],
                   ") from PPTX, failed"))
    return(files)
  }
  files = files[order(nums)]
}

# Unzip PPTX and extract XML files of slides, notes
unzip_pptx = function(file) {
  tdir = tempfile()
  dir.create(tdir)
  # Extract folders from pptx
  res = unzip(file, exdir = tdir)
  rm(res)

  # Reorder xml slide files
  slide_dir = file.path(tdir, "ppt", "slides")
  slides = list.files(path = slide_dir, pattern = "[.]xml$", full.names = TRUE)
  slides = reorder_xml(slides)

  # Reorder xml note files
  note_dir = file.path(tdir, "ppt", "notesSlides")
  notes = list.files(path = note_dir, pattern = "[.]xml$", full.names = TRUE)
  notes = reorder_xml(notes)

  # Copy core.xml file from ariExtra into docProps/core.xml
  tdir = normalizePath(tdir)
  props_dir = file.path(tdir, "docProps")
  props_file = file.path(props_dir, "core.xml")
  ari_core_file = system.file("extdata", "docProps",
                              "core.xml", package = "ariExtra")
  if (!dir.exists(props_file)) {
    dir.create(props_dir, recursive = TRUE)
    file.copy(ari_core_file, props_file,
              overwrite = TRUE)
  }

  list(slides = slides, notes = notes,
       slide_dir = slide_dir, note_dir = note_dir,
       props_dir = props_dir, props_file = props_file,
       root_dir = tdir)
}

#' Extract Notes from XML
#'
#' @param file XML file from a PPTX
#' @param collapse_text should text be collapsed by spaces?
#' @param xpath \code{xpath} to pass to [xml2::xml_find_all()]
#'
#' @return A character vector
#' @export
#'
#' @importFrom xml2 read_xml xml_text xml_find_all
extract_notes_xml = function(file, collapse_text = TRUE, xpath = "//a:r//a:t") {
  xdoc = xml2::read_xml(file)
  # probably need to a:p//a:t and collapse all text within a a:p
  txt = xml2::xml_find_all(x = xdoc, xpath = xpath)
  txt = xml2::xml_text(txt)
  if (collapse_text) {
    txt = paste(txt, collapse = " ")
  }
  return(txt)
}

#' Convert a PPTX file to a PDF file
#'
#' Uses `docxtractr::convert_to_pdf()` for conversion.
#'
#' @param path Path to the PPTX file that needs to be converted to PDF.
#' @param verbose A logical value indicating whether to display progress
#'   messages during the conversion process. The default value is TRUE
#'
#' @importFrom docxtractr convert_to_pdf
#' @export
convert_pptx_pdf = function(path, verbose = TRUE) {
  pdf_file = tempfile(fileext = ".pdf")
  if (verbose) {
    message("Converting PPTX to PDF")
  }
  out = try({
    docxtractr::convert_to_pdf(path, pdf_file = pdf_file)
  })
  if (inherits(out, "try-error")) {
    docxtractr::convert_to_pdf(path, pdf_file = pdf_file)
  }
  if (verbose > 1) {
    message(paste0("PDF is at: ", pdf_file))
  }
  return(pdf_file)
}

#' Convert a PDF file to a series of PNG image files
#'
#' Uses `pdftools::pdf_convert()` for conversion.
#'
#' @param path Path to the PDF file that needs to be converted to PNGs.
#' @param verbose A logical value indicating whether to display progress
#'   messages during the conversion process. The default value is TRUE
#' @param dpi The resolution in dots per inch (dpi) to be used for the PNG
#'   images. The default value is 600.
#'
#' @importFrom pdftools poppler_config pdf_info pdf_convert
#' @export
convert_pdf_png = function(path, verbose = TRUE, dpi = 600) {
  fmts = pdftools::poppler_config()$supported_image_formats
  if ("png" %in% fmts) {
    format = "png"
  } else {
    format = fmts[1]
  }
  info = pdftools::pdf_info(pdf = path)
  filenames = vapply(seq.int(info$pages), function(x) {
    tempfile(fileext = paste0(".", format))
  }, FUN.VALUE = character(1))
  if (verbose) {
    message("Converting PDFs to PNGs")
  }
  pngs = pdftools::pdf_convert(
    pdf = path, dpi = dpi,
    format = format, filenames = filenames,
    verbose = as.logical(verbose))
  pngs
}

% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/extract.R
\name{extract_notes}
\alias{extract_notes}
\title{Extract Speaker Notes from a PowerPoint}
\usage{
extract_notes(file, ...)
}
\arguments{
\item{file}{Character. Path for \code{PPTX} file}

\item{...}{additional arguments to pass to \code{extract_slide_note()}
particularly \code{xpath}}
}
\value{
Either a character vector or \code{NULL}
}
\description{
Extract Speaker Notes from a PowerPoint
}
\examples{
ex_file = system.file("extdata", "example.pptx",
package = "ariExtra")
extract_notes(ex_file)
}

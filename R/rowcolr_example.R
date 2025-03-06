#' Get the path to an example Excel file
#'
#' Retrieves the path to the example `example.xlsx` file included in the `rowcolr` package.
#'
#' @param path Optional. A specific file name within the `extdata` directory. Defaults to `NULL`, which lists available files.
#' @return A character string with the file path or a directory listing.
#' @export
rowcolr_example <- function(path = NULL) {
  if (is.null(path)) {
    dir(system.file("extdata", package = "rowcolr"))
  } else {
    system.file("extdata", path, package = "rowcolr", mustWork = TRUE)
  }
}

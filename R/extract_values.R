#' Extract Values from an Excel File Based on Row and Column Identifiers
#'
#' This function reads an Excel file and extracts structured data by identifying row and column labels
#' using specific suffix patterns (e.g., "_row" for row identifiers and "_col" for column identifiers).
#'
#' @param file Path to the Excel file.
#' @param pattern_row A regex pattern to identify row labels (default: "_row$").
#' @param pattern_col A regex pattern to identify column labels (default: "_col$").
#' @return A tibble containing extracted values, including sheet name, row, column, and description.
#' @importFrom tidyxl xlsx_cells
#' @importFrom dplyr mutate select filter left_join everything
#' @importFrom tidyr crossing
#' @importFrom stringr str_detect str_extract
#' @export
#' @examples
#' # Extract values from an Excel file
#' dataset <- extract_values("data/example.xlsx")
extract_values <- function(file, pattern_row = "_row", pattern_col = "_col") {

  browser()
  # Check if file exists using base::file.exists
  if (!base::file.exists(file)) {
    base::stop("File not found: ", file)
  }

  # Read Excel cells using tidyxl::xlsx_cells
  raw_data <- tidyxl::xlsx_cells(file) |>
    dplyr::mutate(filename = base::basename(file))

  # Identify row and column labels based on patterns using stringr functions
  row_col_data <- raw_data |>
    dplyr::mutate(
      is_row_label = stringr::str_detect(character, pattern_row),
      is_col_label = stringr::str_detect(character, pattern_col)
    ) |>
    dplyr::select(is_row_label, is_col_label, dplyr::everything())

  # Extract row identifiers
  rows <- row_col_data |>
    dplyr::filter(is_row_label) |>
    dplyr::select(sheet, row, row_label = character)

  # Extract column identifiers
  cols <- row_col_data |>
    dplyr::filter(is_col_label) |>
    dplyr::select(sheet, col, col_label = character)

  # Generate all possible row-column combinations
  combinations <- tidyr::crossing(rows, cols) |>
    dplyr::filter(sheet.x == sheet.y) |>
    dplyr::select(sheet = sheet.x, row, col, row_label, col_label) |>
    dplyr::mutate(
      row_label = stringr::str_extract(row_label, ".*(?=_row)"),
      col_label = stringr::str_extract(col_label, ".*(?=_col)"),
      description = base::paste0(row_label, "_", col_label)
    )

  # Merge with original data to get values
  dataset <- combinations |>
    dplyr::left_join(
      row_col_data |>
        dplyr::select(sheet, row, col, value = numeric),
      by = c("sheet", "row", "col")
    ) |>
    dplyr::mutate(filename = base::basename(file)) |>
    dplyr::select(filename, sheet, row, col, row_label, col_label, description, value)

  return(dataset)
}

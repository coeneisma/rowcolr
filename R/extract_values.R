#' Extract Values from an Excel File Based on Row and Column Identifiers
#'
#' @description `r lifecycle::badge('experimental')`
#'
#' This function reads an Excel file and extracts structured data by identifying row and column labels
#' using specific suffix patterns or by matching against a predefined list of identifiers.
#'
#' @param file Path to the Excel file.
#' @param pattern_row A regex pattern to identify row labels (default: ".*_row$").
#' @param pattern_col A regex pattern to identify column labels (default: ".*_col$").
#' @param row_identifiers A character vector of row label identifiers (default: NULL).
#' @param col_identifiers A character vector of column label identifiers (default: NULL).
#' @param clean_description Logical. If TRUE, removes the identified suffix from the description (default: TRUE).
#' @return A tibble containing extracted values, including sheet name, row, column, and description.
#' @export
#' @examples
#' \dontrun{
#' # Extract values using regex patterns
#' dataset <- extract_values("data/example.xlsx")
#'
#' # Extract values using a predefined list of labels
#' dataset <- extract_values("data/example.xlsx",
#'                           row_identifiers = c("Total Assets", "Total Liabilities"),
#'                           col_identifiers = c("Year 2023", "Year 2024"))
#' }
extract_values <- function(file,
                           pattern_row = ".*_row$", pattern_col = ".*_col$",
                           row_identifiers = NULL, col_identifiers = NULL,
                           clean_description = TRUE) {

  # Check if the file exists
  if (!base::file.exists(file)) {
    base::stop("File not found: ", file)
  }

  # Read Excel cells
  raw_data <- tidyxl::xlsx_cells(file) |>
    dplyr::mutate(filename = base::basename(file))

  # Identify row and column labels based on regex patterns or identifier lists
  row_col_data <- raw_data |>
    dplyr::mutate(
      is_row_label = if (!is.null(row_identifiers)) character %in% row_identifiers else stringr::str_detect(character, pattern_row),
      is_col_label = if (!is.null(col_identifiers)) character %in% col_identifiers else stringr::str_detect(character, pattern_col)
    ) |>
    dplyr::select(sheet, row, col, character, is_row_label, is_col_label, dplyr::everything())

  # Extract row labels
  rows <- row_col_data |>
    dplyr::filter(is_row_label) |>
    dplyr::select(sheet_rows = sheet, row, row_label = character)

  # Extract column labels
  cols <- row_col_data |>
    dplyr::filter(is_col_label) |>
    dplyr::select(sheet_cols = sheet, col, col_label = character)

  # 🛠 Improved regex cleaning function
  clean_regex <- function(pattern) {
    # Remove ONLY leading ".*" or trailing ".*" or standalone "."
    pattern |>
      stringr::str_replace_all("^\\.*", "") |>  # Remove leading dots (.)
      stringr::str_replace_all("\\.*\\$$", "") |> # Remove trailing .* or .
      stringr::str_replace_all("^\\*", "") |>  # Remove leading *
      stringr::str_replace_all("\\$$", "") # Remove trailing $
  }

  clean_pattern_row <- clean_regex(pattern_row)
  clean_pattern_col <- clean_regex(pattern_col)

  # Generate all possible row-column combinations
  combinations <- tidyr::crossing(rows, cols) |>
    dplyr::filter(sheet_cols == sheet_rows) |>
    dplyr::select(sheet = sheet_cols, row, col, row_label, col_label) |>
    dplyr::mutate(
      # Remove suffix if clean_description is TRUE
      row_label_clean = if (clean_description) stringr::str_remove(row_label, clean_pattern_row) else row_label,
      col_label_clean = if (clean_description) stringr::str_remove(col_label, clean_pattern_col) else col_label,
      description = base::paste0(row_label_clean, "_", col_label_clean)
    )

  # Merge with the original dataset to get the corresponding values
  dataset <- combinations |>
    dplyr::left_join(
      row_col_data |>
        dplyr::select(sheet, row, col,
                      data_type, error, logical, numeric, date, character),
      by = c("sheet", "row", "col")
    ) |>
    dplyr::mutate(filename = base::basename(file)) |>
    dplyr::select(filename, sheet, row, col, row_label, col_label, description,
                  data_type, error, logical, numeric, date, character)

  return(dataset)
}

utils::globalVariables(c("sheet", "is_row_label", "is_col_label", "sheet_cols",
                         "sheet_rows", "row_label", "col_label", "row_label_clean",
                         "col_label_clean", "data_type", "error", "filename",
                         "description"))

#' Extract Values from an Excel File Based on Row and Column Identifiers
#'
#' @description This function reads an Excel file and extracts structured data
#' by identifying row and column labels. It ensures that each row label is only
#' matched with the closest column label to its right and only with column
#' labels from the nearest preceding row that contains column headers.
#'
#' @param file Path to the Excel file.
#' @param pattern_row A regex pattern to identify row labels (default:
#'   ".*_row$").
#' @param pattern_col A regex pattern to identify column labels (default:
#'   ".*_col$").
#' @param row_identifiers A character vector of row label identifiers (default:
#'   NULL).
#' @param col_identifiers A character vector of column label identifiers
#'   (default: NULL).
#' @param clean_description Logical. If TRUE, removes text matching
#'   `pattern_row` and `pattern_col` from row and column labels (default: TRUE).
#' @return A tibble containing extracted values, including sheet name, row,
#'   column, and description.
#' @export
#' @examples
#' # Extract values using regex patterns
#' dataset <- extract_values(rowcolr_example("example.xlsx"))
#'
#' # Extract values using a predefined list of labels
#' dataset <- extract_values(rowcolr_example("example.xlsx"),
#'                           row_identifiers = c("Total Assets", "Total Liabilities"),
#'                           col_identifiers = c("Year 2023", "Year 2024"))
extract_values <- function(file,
                           pattern_row = ".*_row$", pattern_col = ".*_col$",
                           row_identifiers = NULL, col_identifiers = NULL,
                           clean_description = TRUE) {

  # Check if the file exists
  if (!base::file.exists(file)) {
    base::stop("File not found: ", file)
  }

  # Read the Excel file and extract raw cell data
  raw_data <- tidyxl::xlsx_cells(file) |>
    dplyr::mutate(filename = base::basename(file))

  # Identify row and column labels based on regex patterns or predefined lists
  row_col_data <- raw_data |>
    dplyr::mutate(
      is_row_label = if (!is.null(row_identifiers)) character %in% row_identifiers else stringr::str_detect(character, pattern_row),
      is_col_label = if (!is.null(col_identifiers)) character %in% col_identifiers else stringr::str_detect(character, pattern_col)
    ) |>
    dplyr::select(sheet, row, col, character, is_row_label, is_col_label, dplyr::everything())

  # Extract row labels
  rows <- row_col_data |>
    dplyr::filter(is_row_label) |>
    dplyr::select(sheet_rows = sheet, row, row_col = col, row_label = character)  # Rename col to row_col

  # Extract column labels, including their row positions
  cols <- row_col_data |>
    dplyr::filter(is_col_label) |>
    dplyr::select(sheet_cols = sheet, col, row, col_label = character) |>
    dplyr::rename(col_row = row)  # Rename row to col_row (row of the column label)

  # Utility function to clean regex patterns
  clean_regex <- function(pattern) {
    pattern |>
      stringr::str_replace_all("^\\.*", "") |>
      stringr::str_replace_all("\\.*\\$$", "") |>
      stringr::str_replace_all("^\\*", "") |>
      stringr::str_replace_all("\\$$", "")
  }

  clean_pattern_row <- clean_regex(pattern_row)
  clean_pattern_col <- clean_regex(pattern_col)

  # Determine the closest preceding row with column labels for each row label
  rows <- rows |>
    dplyr::group_by(sheet_rows) |>
    dplyr::mutate(
      closest_col_row = purrr::map_dbl(row, ~ max(cols$col_row[cols$col_row < .x], na.rm = TRUE))
    ) |>
    dplyr::ungroup()

  # Find the closest column label to the right of each row label
  rows <- rows |>
    dplyr::rowwise() |>
    dplyr::mutate(
      # closest_col = ifelse(length(cols$col[cols$col > row$col]) > 0,
      #                      min(cols$col[cols$col > row$col]),
      #                      NA)
      closest_col = min(cols$col[cols$sheet_cols == sheet_rows &
                                   cols$col_row == closest_col_row &
                                   cols$col > row_col],
                        na.rm = TRUE)
    ) |>
    dplyr::ungroup()

  # Generate valid row-column intersections based on the closest column label row and position
  combinations <- tidyr::crossing(rows, cols) |>
    dplyr::filter(sheet_cols == sheet_rows,
                  col_row < row,
                  col_row == closest_col_row,
                  col == closest_col) |>
    dplyr::select(sheet = sheet_cols, row, col, row_label, col_label) |>
    dplyr::mutate(
      row_label_clean = if (clean_description) stringr::str_remove(row_label, clean_pattern_row) else row_label,
      col_label_clean = if (clean_description) stringr::str_remove(col_label, clean_pattern_col) else col_label,
      description = base::paste0(row_label_clean, "_", col_label_clean)
    )

  # Merge with the original dataset to retrieve corresponding values
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
                         "description", "col_row", "closest_col_row", "closest_col", "row_col"))

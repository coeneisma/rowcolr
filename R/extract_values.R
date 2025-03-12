#' Extract Values from an Excel File Based on Row and Column Identifiers
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#'   This function reads an Excel file and extracts structured data by
#'   identifying row and column labels. It ensures that each row label is only
#'   matched with the closest column label to its right and only with column
#'   labels from the nearest preceding row that contains column headers.
#'
#'   You can use **both regex patterns and explicit identifiers together**.
#'   - If `row_identifiers` and/or `col_identifiers` are provided, they take **priority** over `row_pattern` and `col_pattern`.
#'   - If **both** explicit identifiers and regex patterns are provided, the function will **first** match on identifiers. Any additional matches from regex patterns will be included.
#'
#' @param file Path to the Excel file.
#' @param row_identifiers A character vector of row label identifiers. If
#'   provided, takes precedence over `row_pattern`, but **does not disable it**.
#' @param col_identifiers A character vector of column label identifiers. If
#'   provided, takes precedence over `col_pattern`, but **does not disable it**.
#' @param row_pattern A regex pattern to identify row labels. **Used only if `row_identifiers` does not fully capture all row labels**. (Example: `".*_row$"`).
#' @param col_pattern A regex pattern to identify column labels. **Used only if `col_identifiers` does not fully capture all column labels**. (Example: `".*_col$"`).
#' @param clean_description Logical. If TRUE, removes text matching
#'   `row_pattern` and `col_pattern` from row and column labels (default: TRUE).
#' @return A tibble containing extracted values, including sheet name, row,
#'   column, and description.
#' @export
#' @examples
#' # Extract values using regex patterns
#' dataset <- extract_values(rowcolr_example("example.xlsx"), row_pattern = ".*_row$", col_pattern = ".*_col$")
#' dataset
#'
#' # Extract values using a predefined list of labels (ignoring regex)
#' dataset <- extract_values(rowcolr_example("example.xlsx"),
#'                           row_identifiers = c("Total assets (4+9)", "Total equity (10+11+12)"),
#'                           col_identifiers = c("2025_col"))
#' dataset
#'
#' # Extract values using BOTH identifiers and regex
#' dataset <- extract_values(rowcolr_example("example.xlsx"),
#'                           row_identifiers = c("Total assets (4+9)"),
#'                           row_pattern = ".*_row$")
#' dataset
extract_values <- function(file,
                           row_pattern = NULL, col_pattern = NULL,
                           row_identifiers = NULL, col_identifiers = NULL,
                           clean_description = TRUE) {

  if (missing(file)) {
    cli::cli_abort("The `file` argument is required. Please provide the path to an Excel file.")
  }

  if (!base::file.exists(file)) {
    cli::cli_abort("File {.path {file}} not found. Please check the file path and try again.")
  }

  if (is.null(row_pattern) && is.null(col_pattern) && is.null(row_identifiers) && is.null(col_identifiers)) {
    cli::cli_abort("No identifiers or regex patterns provided. Please specify `row_identifiers`, `col_identifiers`, `row_pattern`, or `col_pattern`.")
  }

  # Read the raw data from the Excel file using tidyxl::xlsx_cells
  raw_data <- tidyxl::xlsx_cells(file) |>
    dplyr::mutate(filename = base::basename(file))

  # Identify row and column labels based on exact identifiers and regex patterns
  row_col_data <- raw_data |>
    dplyr::mutate(
      is_row_label = if (!is.null(row_identifiers)) character %in% row_identifiers else FALSE,
      is_col_label = if (!is.null(col_identifiers)) character %in% col_identifiers else FALSE,

      # Add regex matches if the patterns exist
      is_row_label = if (!is.null(row_pattern)) {
        is_row_label | stringr::str_detect(character, row_pattern)
      } else {
        is_row_label
      },

      is_col_label = if (!is.null(col_pattern)) {
        is_col_label | stringr::str_detect(character, col_pattern)
      } else {
        is_col_label
      }
    )

  # Extract rows and columns that match the identifiers
  rows <- row_col_data |>
    dplyr::filter(is_row_label) |>
    dplyr::select(sheet_rows = sheet, row, row_col = col, row_label = character)

  cols <- row_col_data |>
    dplyr::filter(is_col_label) |>
    dplyr::select(sheet_cols = sheet, col, row, col_label = character) |>
    dplyr::rename(col_row = row)

  # For each row, identify the closest column label to its right from the preceding row
  rows <- rows |>
    dplyr::group_by(sheet_rows) |>
    dplyr::mutate(
      closest_col_row = purrr::map_dbl(row, ~ max(cols$col_row[cols$col_row < .x], na.rm = TRUE))
    ) |>
    dplyr::ungroup()

  # Identify the exact column corresponding to the closest column row
  rows <- rows |>
    dplyr::rowwise() |>
    dplyr::mutate(
      closest_col = suppressWarnings(min(cols$col[cols$sheet_cols == sheet_rows &
                                                    cols$col_row == closest_col_row &
                                                    cols$col > row_col],
                                         na.rm = TRUE))
    ) |>
    dplyr::ungroup()

  # Generate combinations of rows and columns based on the closest matching criteria
  combinations <- tidyr::crossing(rows, cols) |>
    dplyr::filter(sheet_cols == sheet_rows,
                  col_row < row,
                  col_row == closest_col_row,
                  col == closest_col) |>
    dplyr::select(sheet = sheet_cols, row, col, row_label, col_label) |>
    dplyr::mutate(
      row_label_clean = if (clean_description & !is.null(row_pattern)) stringr::str_remove(row_label, row_pattern) else row_label,
      col_label_clean = if (clean_description & !is.null(col_pattern)) stringr::str_remove(col_label, col_pattern) else col_label,
      description = base::paste0(row_label_clean, "_", col_label_clean)
    )

  # Join the extracted combinations with the original raw data and return the final dataset
  dataset <- combinations |>
    dplyr::left_join(
      row_col_data |>
        dplyr::select(sheet, row, col, data_type, error, logical, numeric, date, character),
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
                         "description", "col_row", "closest_col_row", "closest_col",
                         "row_col", "col_identifiers_example", "row_identifiers_example"))

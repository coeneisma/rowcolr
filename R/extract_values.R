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
#' @param fuzzy_threshold Numeric between 0 and 1. If set, enables fuzzy matching
#'   for identifiers. 0 = exact match required, 1 = any match accepted.
#'   Recommended: 0.8 (default: NULL, disabled).
#' @param fuzzy_method Character. Method for fuzzy matching from stringdist
#'   package: "osa", "lv", "dl", "jaccard", "jw" (default: "jw" for
#'   Jaro-Winkler).
#' @return A tibble containing extracted values with the following columns:
#'   - filename: Name of the source Excel file
#'   - sheet: Sheet name where the value was found
#'   - row, col: Row and column coordinates of the value
#'   - row_label, col_label: The matched row and column labels
#'   - description: Combined row and column label
#'   - data_type: Excel data type of the cell
#'   - error, logical, numeric, date, character: Cell values by type
#'   - fuzzy_threshold: The fuzzy matching threshold used (if any)
#'   - row_similarity: Similarity score for row label match (0-1)
#'   - col_similarity: Similarity score for column label match (0-1)
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
#'
#' # Extract values with fuzzy matching enabled
#' dataset <- extract_values(rowcolr_example("example.xlsx"),
#'                           row_identifiers = c("Total assets", "Total equity"),
#'                           col_identifiers = c("2025"),
#'                           fuzzy_threshold = 0.8)
#' dataset
extract_values <- function(file,
                           row_pattern = NULL, col_pattern = NULL,
                           row_identifiers = NULL, col_identifiers = NULL,
                           clean_description = TRUE,
                           fuzzy_threshold = NULL, fuzzy_method = "jw") {

  if (missing(file)) {
    cli::cli_abort("The `file` argument is required. Please provide the path to an Excel file.")
  }

  if (!base::file.exists(file)) {
    cli::cli_abort("File {.path {file}} not found. Please check the file path and try again.")
  }

  if (is.null(row_pattern) && is.null(col_pattern) && is.null(row_identifiers) && is.null(col_identifiers)) {
    cli::cli_abort("No identifiers or regex patterns provided. Please specify `row_identifiers`, `col_identifiers`, `row_pattern`, or `col_pattern`.")
  }

  # Validate fuzzy matching parameters
  if (!is.null(fuzzy_threshold)) {
    if (!is.numeric(fuzzy_threshold) || fuzzy_threshold < 0 || fuzzy_threshold > 1) {
      cli::cli_abort("`fuzzy_threshold` must be a numeric value between 0 and 1.")
    }
    if (!fuzzy_method %in% c("osa", "lv", "dl", "jaccard", "jw")) {
      cli::cli_abort("`fuzzy_method` must be one of: 'osa', 'lv', 'dl', 'jaccard', 'jw'.")
    }
  }

  # Read the raw data from the Excel file using tidyxl::xlsx_cells
  raw_data <- tidyxl::xlsx_cells(file) |>
    dplyr::mutate(filename = base::basename(file))

  # Identify row and column labels based on exact identifiers and regex patterns
  row_col_data <- raw_data |>
    dplyr::mutate(
      # Exact matches first
      is_row_label = if (!is.null(row_identifiers)) character %in% row_identifiers else FALSE,
      is_col_label = if (!is.null(col_identifiers)) character %in% col_identifiers else FALSE,
      
      # Initialize similarity scores
      row_similarity = 0,
      col_similarity = 0,
      
      # Add fuzzy matches if enabled and get similarity scores
      row_fuzzy_match = if (!is.null(row_identifiers) && !is.null(fuzzy_threshold)) {
        apply_fuzzy_matching(character, row_identifiers, fuzzy_threshold, fuzzy_method)
      } else {
        FALSE
      },
      
      col_fuzzy_match = if (!is.null(col_identifiers) && !is.null(fuzzy_threshold)) {
        apply_fuzzy_matching(character, col_identifiers, fuzzy_threshold, fuzzy_method)
      } else {
        FALSE
      },
      
      # Get similarity scores for fuzzy matches
      row_similarity = if (!is.null(row_identifiers) && !is.null(fuzzy_threshold)) {
        apply_fuzzy_matching(character, row_identifiers, fuzzy_threshold, fuzzy_method, return_scores = TRUE)
      } else {
        row_similarity
      },
      
      col_similarity = if (!is.null(col_identifiers) && !is.null(fuzzy_threshold)) {
        apply_fuzzy_matching(character, col_identifiers, fuzzy_threshold, fuzzy_method, return_scores = TRUE)
      } else {
        col_similarity
      },
      
      # Combine exact and fuzzy matches
      is_row_label = is_row_label | row_fuzzy_match,
      is_col_label = is_col_label | col_fuzzy_match,
      
      # Set similarity to 1 for exact matches
      row_similarity = ifelse(!is.null(row_identifiers) & character %in% row_identifiers, 1, row_similarity),
      col_similarity = ifelse(!is.null(col_identifiers) & character %in% col_identifiers, 1, col_similarity),

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
    dplyr::select(sheet_rows = sheet, row, row_col = col, row_label = character, row_similarity)

  cols <- row_col_data |>
    dplyr::filter(is_col_label) |>
    dplyr::select(sheet_cols = sheet, col, row, col_label = character, col_similarity) |>
    dplyr::rename(col_row = row)
  
  # Provide feedback about fuzzy matching if enabled
  if (!is.null(fuzzy_threshold) && (!is.null(row_identifiers) || !is.null(col_identifiers))) {
    n_row_matches <- nrow(rows)
    n_col_matches <- nrow(cols)
    
    if (n_row_matches > 0 || n_col_matches > 0) {
      cli::cli_inform(c(
        "i" = "Fuzzy matching enabled with threshold {fuzzy_threshold}",
        "i" = "Found {n_row_matches} row label{?s} and {n_col_matches} column label{?s}"
      ))
    }
  }

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
    dplyr::select(sheet = sheet_cols, row, col, row_label, col_label, row_similarity, col_similarity) |>
    dplyr::mutate(
      row_label_clean = if (clean_description & !is.null(row_pattern)) stringr::str_remove(row_label, row_pattern) else row_label,
      col_label_clean = if (clean_description & !is.null(col_pattern)) stringr::str_remove(col_label, col_pattern) else col_label,
      description = base::paste0(row_label_clean, "_", col_label_clean),
      fuzzy_threshold = fuzzy_threshold
    )

  # Join the extracted combinations with the original raw data and return the final dataset
  dataset <- combinations |>
    dplyr::left_join(
      row_col_data |>
        dplyr::select(sheet, row, col, data_type, error, logical, numeric, date, character),
      by = c("sheet", "row", "col")
    ) |>
    dplyr::mutate(filename = base::basename(file))
  
  # Select columns based on whether fuzzy matching was used
  if (!is.null(fuzzy_threshold)) {
    dataset <- dataset |>
      dplyr::select(filename, sheet, row, col, row_label, col_label, description,
                    data_type, error, logical, numeric, date, character,
                    fuzzy_threshold, row_similarity, col_similarity)
  } else {
    dataset <- dataset |>
      dplyr::select(filename, sheet, row, col, row_label, col_label, description,
                    data_type, error, logical, numeric, date, character)
  }

  # Warn if no data was extracted
  if (nrow(dataset) == 0) {
    cli::cli_warn("No data extracted. Check your identifiers/patterns.")
  }

  return(dataset)
}

utils::globalVariables(c("sheet", "is_row_label", "is_col_label", "sheet_cols",
                         "sheet_rows", "row_label", "col_label", "row_label_clean",
                         "col_label_clean", "data_type", "error", "filename",
                         "description", "col_row", "closest_col_row", "closest_col",
                         "row_col", "row_similarity", "col_similarity", "row_fuzzy_match",
                         "col_fuzzy_match", "fuzzy_threshold"))

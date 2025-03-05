
<!-- README.md is generated from README.Rmd. Please edit that file -->

# rowcolr

<!-- badges: start -->
<!-- badges: end -->

The goal of `rowcolr` is to extract structured data from Excel files by
identifying row and column labels using regex patterns or predefined
identifiers. The package is designed to support easy extraction of data
based on matching row and column labels, facilitating data wrangling and
cleaning tasks from spreadsheet data.

## Origin

The development of `rowcolr` was driven by a specific use case: existing
Excel files used as input forms by organizations in the arts and culture
sector to provide information to the Dutch Ministry of Education,
Culture, and Science (OCW). To minimize disruptions to the existing
workflow, the forms were kept as similar as possible to their original
format, while still allowing structured extraction of all relevant data
from the Excel files.

## Installation

You can install the development version of `rowcolr` from GitHub with:

``` r
devtools::install_github("coeneisma/rowcolr")
```

## Usage

The primary function of `rowcolr` is `extract_values()`, which allows
you to extract structured data from Excel files based on row and column
label patterns.

### Example: Extracting Data from an Excel File

``` r
library(rowcolr)

# Extract values from an Excel file using default row/column regex patterns
dataset <- extract_values("path/to/your/excel_file.xlsx")
head(dataset)
```

### Example: Using Specific Row and Column Identifiers

If you have predefined row or column labels, you can use them directly
by providing a character vector.

``` r
# Extract values using predefined row and column identifiers
dataset <- extract_values("path/to/your/excel_file.xlsx",
                           row_identifiers = c("Total Assets", "Net Profit"),
                           col_identifiers = c("2023", "2024"))
head(dataset)
```

### Example: Cleaning Descriptions by Removing Suffix Patterns

You can also clean the row and column labels before combining them into
the description.

``` r
# Extract values with cleaned descriptions
dataset <- extract_values("path/to/your/excel_file.xlsx", clean_description = TRUE)
head(dataset)
```

## Conclusion

The `rowcolr` package is a powerful tool for extracting structured data
from Excel files based on customizable row and column patterns. It is
designed to be flexible, allowing for both regex-based matching and
predefined lists of identifiers.

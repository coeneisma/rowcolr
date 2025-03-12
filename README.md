
<!-- README.md is generated from README.Rmd. Please edit that file -->

# rowcolr

<!-- badges: start -->

[![Lifecycle:
experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://lifecycle.r-lib.org/articles/stages.html#experimental)
[![CRAN
status](https://www.r-pkg.org/badges/version/rowcolr)](https://CRAN.R-project.org/package=rowcolr)
[![R-CMD-check](https://github.com/coeneisma/rowcolr/actions/workflows/R-CMD-check.yaml/badge.svg)](https://github.com/coeneisma/rowcolr/actions/workflows/R-CMD-check.yaml)

<!-- badges: end -->

The goal of `rowcolr` is to extract structured data from semi-structured
Excel files by identifying row and column labels using predefined
identifiers or regex patterns. Processing such data can be challenging
with existing packages. `rowcolr` simplifies this process by efficiently
locating values at the intersection of row and column labels, making
data wrangling and cleaning more seamless.

Built on the powerful `tidyxl` package, `rowcolr` ensures robust and
reliable extraction of cell contents from Excel files.

## Origin

The development of `rowcolr` was driven by a specific use case: existing
Excel files used as input forms by organizations in the arts and culture
sector to provide information to the Dutch Ministry of Education,
Culture, and Science (OCW). To minimize disruptions to the existing
workflow, the forms were kept as similar as possible to their original
format while still allowing structured extraction of all relevant data
from the Excel files.

## Installation

You can install the development version of `rowcolr` from GitHub with:

``` r
devtools::install_github("coeneisma/rowcolr")
```

## Usage

The primary function of `rowcolr` is `extract_values()`, which allows
you to extract structured data from Excel files by identifying row and
column labels. The function works by drawing a horizontal line from a
detected row label and a vertical line from a detected column label. The
value located at their intersection is extracted and stored.

The package can be loaded with:

``` r
library(rowcolr)
```

### Example: Extracting Data from an Excel File

Values can be extracted from an Excel file using row and column
identifiers or regex patterns:

``` r
# Extract values from an example Excel file with regex patterns
dataset <- extract_values(rowcolr_example("example.xlsx"), 
                          row_pattern = ".*_row$", 
                          col_pattern = ".*_col$")

head(dataset |> dplyr::select(-c(filename, sheet, row, col)))
#> # A tibble: 6 × 9
#>   row_label            col_label description data_type error logical  numeric
#>   <chr>                <chr>     <chr>       <chr>     <chr> <lgl>      <dbl>
#> 1 name_row             value_col _           character <NA>  NA            NA
#> 2 year_row             value_col _           numeric   <NA>  NA          2025
#> 3 kvk_row              value_col _           numeric   <NA>  NA      12345678
#> 4 test1_row            2026_col  _           character <NA>  NA            NA
#> 5 tes2_row             2026_col  _           character <NA>  NA            NA
#> 6 int_fixed_assets_row 2025_col  _           numeric   <NA>  NA        100000
#> # ℹ 2 more variables: date <dttm>, character <chr>
```

For more details on using the package, refer to the vignette:
`vignette("rowcolr")`.


<!-- README.md is generated from README.Rmd. Please edit that file -->

# rowcolr

<!-- badges: start -->

[![Lifecycle:
experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://lifecycle.r-lib.org/articles/stages.html#experimental)
<!-- badges: end -->

The goal of `rowcolr` is to extract structured data from Excel files by
identifying row and column labels using regex patterns or predefined
identifiers. The package is designed to facilitate easy extraction of
data by locating values at the intersection of row and column labels,
making data wrangling and cleaning more efficient.

`rowcolr` leverages the excellent `tidyxl` package to read data from
Excel files, ensuring robust and reliable extraction of cell contents.
The package is particularly useful for dealing with semi-structured data
in Excel files, which can be difficult to process using existing
packages.

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

Values can be extracted from an Excel file using the default row and
column regex patterns `.*_row$` and `.*_col$`:

``` r
# Extract values from an example Excel file
dataset <- extract_values(rowcolr_example("example.xlsx"))

head(dataset |> 
       dplyr::select(-c(filename, sheet, row, col)))
#> # A tibble: 6 × 9
#>   row_label             col_label description    data_type error logical numeric
#>   <chr>                 <chr>     <chr>          <chr>     <chr> <lgl>     <dbl>
#> 1 int_fixed_assets_row  2025_col  int_fixed_ass… numeric   <NA>  NA       100000
#> 2 tang_fixed_assets_row 2025_col  tang_fixed_as… numeric   <NA>  NA       200000
#> 3 fin_fixed_assets_row  2025_col  fin_fixed_ass… numeric   <NA>  NA       150000
#> 4 tot_fixed_assets_row  2025_col  tot_fixed_ass… numeric   <NA>  NA       450000
#> 5 stock_row             2025_col  stock_2025     numeric   <NA>  NA        25000
#> 6 receivables_row       2025_col  receivables_2… numeric   <NA>  NA        15000
#> # ℹ 2 more variables: date <dttm>, character <chr>
```

For more details on using the package, refer to the vignette:
`vignette("rowcolr")`.

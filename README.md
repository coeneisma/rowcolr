
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

The package can be loaded with:

``` r
library(rowcolr)
```

### Example: Extracting Data from an Excel File

Values can be extract from an Excel file using default row/column regex
patterns `.*_row$` and .\*\_col\$\`:

``` r
# 
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

More on the use of the package can be found in the vignette:
`vignette("rowcolr")`.

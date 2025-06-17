# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Package Overview

rowcolr is an R package for extracting structured data from Excel files by identifying row and column labels. The package is designed to handle Excel files where data is organized with descriptive labels in rows and columns, and extracts values at their intersections.

## Commands

### Development Commands
```bash
# Document the package (update NAMESPACE and man pages)
Rscript -e "devtools::document()"

# Check the package (full R CMD check)
Rscript -e "devtools::check()"

# Build the package
Rscript -e "devtools::build()"

# Install the package locally
Rscript -e "devtools::install()"

# Build pkgdown site locally
Rscript -e "pkgdown::build_site()"
```

### Testing
**Note**: This package currently has no test suite. When adding tests, create a `tests/testthat/` directory and use the testthat framework.

## Architecture

### Core Functionality
The package has two exported functions:
- `extract_values()`: Main function that extracts data from Excel files based on row/column identifiers or patterns
- `rowcolr_example()`: Helper to access example files in `inst/extdata/`

### Key Design Decisions
1. **Excel Parsing**: Uses `tidyxl::xlsx_cells()` for low-level cell access, preserving all Excel metadata
2. **Label Matching**: Supports both explicit identifiers and regex patterns for finding labels
3. **Intersection Algorithm**: Uses a "closest match" approach - for each row label, finds the nearest column label in the same row
4. **Data Structure**: Returns a tibble with original Excel coordinates and extracted values

### Dependencies
- **tidyxl**: Excel file parsing
- **dplyr/tidyr/purrr**: Data manipulation
- **stringr**: String operations and regex
- **cli**: User-friendly error messages

### File Organization
```
R/
├── extract_values.R    # Main extraction logic
└── rowcolr_example.R   # Example file helper

inst/
└── extdata/           # Example Excel files

vignettes/             # Package documentation
man/                   # Generated documentation
```

## Important Context
- Package follows tidyverse conventions (pipe operator, tibble output)
- All user-facing messages should use `cli` for consistency
- The package is marked as "experimental" in lifecycle
- Originally developed for Dutch government (OCW) use case
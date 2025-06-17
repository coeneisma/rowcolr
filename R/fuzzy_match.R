#' Find fuzzy matches for a target string in a vector of candidates
#'
#' @param target Character string to match
#' @param candidates Character vector of potential matches
#' @param threshold Numeric between 0 and 1. Minimum similarity score required
#' @param method Character. String distance method from stringdist package
#' @param return_scores Logical. If TRUE, returns similarity scores instead of logical vector
#' @return Logical vector indicating which candidates match, or numeric vector of scores
#' @keywords internal
find_fuzzy_matches <- function(target, candidates, threshold = 0.8, method = "jw", return_scores = FALSE) {
  if (length(candidates) == 0 || is.na(target) || target == "") {
    if (return_scores) {
      return(numeric(length(candidates)))
    } else {
      return(logical(length(candidates)))
    }
  }
  
  # Calculate string similarity (1 - normalized distance)
  similarities <- stringdist::stringsimmatrix(target, candidates, method = method)
  
  if (return_scores) {
    # Return similarity scores for matches above threshold, 0 for non-matches
    return(ifelse(similarities >= threshold, similarities, 0))
  } else {
    # Return logical vector of matches above threshold
    return(similarities >= threshold)
  }
}

#' Apply fuzzy matching to a vector of strings against identifiers
#'
#' @param strings Character vector to check for matches
#' @param identifiers Character vector of identifiers to match against
#' @param threshold Numeric between 0 and 1. Minimum similarity score
#' @param method Character. String distance method
#' @param return_scores Logical. If TRUE, returns best similarity scores
#' @return Logical vector indicating matches, or numeric vector of best scores
#' @keywords internal
apply_fuzzy_matching <- function(strings, identifiers, threshold = 0.8, method = "jw", return_scores = FALSE) {
  if (is.null(identifiers) || length(identifiers) == 0) {
    if (return_scores) {
      return(rep(0, length(strings)))
    } else {
      return(rep(FALSE, length(strings)))
    }
  }
  
  if (return_scores) {
    # For each string, return the best similarity score
    purrr::map_dbl(strings, ~ {
      if (is.na(.x) || .x == "") {
        return(0)
      }
      scores <- find_fuzzy_matches(.x, identifiers, threshold, method, return_scores = TRUE)
      max(scores)
    })
  } else {
    # For each string, check if it fuzzy matches any identifier
    purrr::map_lgl(strings, ~ {
      if (is.na(.x) || .x == "") {
        return(FALSE)
      }
      any(find_fuzzy_matches(.x, identifiers, threshold, method))
    })
  }
}
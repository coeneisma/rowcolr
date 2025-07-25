% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/fuzzy_match.R
\name{find_fuzzy_matches}
\alias{find_fuzzy_matches}
\title{Find fuzzy matches for a target string in a vector of candidates}
\usage{
find_fuzzy_matches(
  target,
  candidates,
  threshold = 0.8,
  method = "jw",
  return_scores = FALSE
)
}
\arguments{
\item{target}{Character string to match}

\item{candidates}{Character vector of potential matches}

\item{threshold}{Numeric between 0 and 1. Minimum similarity score required}

\item{method}{Character. String distance method from stringdist package}

\item{return_scores}{Logical. If TRUE, returns similarity scores instead of logical vector}
}
\value{
Logical vector indicating which candidates match, or numeric vector of scores
}
\description{
Find fuzzy matches for a target string in a vector of candidates
}
\keyword{internal}

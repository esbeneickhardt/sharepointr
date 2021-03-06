% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/sharepoint_functions.R
\name{get_sharepoint_file_names}
\alias{get_sharepoint_file_names}
\title{Get list of file names}
\usage{
get_sharepoint_file_names(
  sharepoint_token,
  sharepoint_url,
  sharepoint_digest_value,
  sharepoint_path
)
}
\arguments{
\item{sharepoint_token}{A SharePoint token from get_sharepoint_token()}

\item{sharepoint_url}{A SharePoint url, e.g. kksky.sharepoint.com}

\item{sharepoint_digest_value}{A SharePoint digest value from get_sharepoint_digest_value()}

\item{sharepoint_path}{Path to the file, e.g. Shared Documents/test}
}
\value{
character vector of file names
}
\description{
Get list of file names
}
\examples{
#no example yet
}

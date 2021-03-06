% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/Smartsheet-R-sdk.R
\name{colorize_sheet}
\alias{colorize_sheet}
\title{Colorize Smartsheet}
\usage{
colorize_sheet(sheet_name, clean_hex_col = TRUE, batch_size = 5000)
}
\arguments{
\item{sheet_name}{the target Smartsheet's exact sheet name}

\item{clean_hex_col}{defaults to TRUE and removes HEX_COLOR column when complete}
}
\value{
returns nothing
}
\description{
For a sheet with a column HEX_COLOR that sheet will be updated
(via a full replacement) so that those row have those colors and then the HEX_COLOR
column will be removed
}
\examples{
\dontrun{
colorize_sheet("sheet_name")
}
}

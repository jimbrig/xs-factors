library(RDCOMClient)

#' Batch Convert ".xls" Excel Workbooks to ".xlsx" - WINDOWS ONLY
#'
#' @param in_folder input folder path containing .xls files to convert
#' @param out_folder output folder where .xlsx files should be saves
#' @param delete_xls delete input .xls files ?
#'
#' @return invisible
#' @export
#'
#' @examples
#' library(RDCOMClient)
#' convert_xls_to_xlsx(in_folder = "data/raw/NCCI", out_folder = "data")
convert_xls_to_xlsx <- function(in_folder, out_folder, delete_xls = F) {
  if (missing(out_folder)) {
    out_folder <- in_folder
  }

  all_xls <- list.files(in_folder, pattern = ".xls$")

  if (length(all_xls) > 0) {
    all_xls_out <- gsub(".xls$", ".xlsx", all_xls)

    try({
      xls <- COMCreate("Excel.Application")

      lapply(1:length(all_xls), function(i) {
        cat(i, "\n")
        wb <- xls[["Workbooks"]]$Open(normalizePath(paste(in_folder, all_xls[i], sep = "\\")))
        wb$SaveAs(suppressWarnings(normalizePath(paste(out_folder, all_xls_out[i], sep = "\\"))), 51)
        wb$Close()
      })

      xls$Quit()
    }, silent = T)

    if (delete_xls) {
      all_xlsx_now <- list.files(in_folder, pattern = ".xlsx$")
      test <- setdiff(gsub(".xls$", "", all_xls), gsub(".xlsx$", "", all_xlsx_now))
      if (length(test) == 0) {
        try(unlink(paste(in_folder, all_xls, sep = "\\")), silent = T)
      }
    }
  }

  return(invisible(0))
}

#' List items in a SharePoint folder
#'
#' `sp_list_items()` lists both files and folders. `sp_list_files()` and
#' `sp_list_folders()` are convenience wrappers that filter to just one type.
#'
#' @param auth A `sharepointr_auth` object from [sharepoint_auth()].
#' @param path Library-relative path to a folder, e.g. `"Shared Documents"`
#'   or `"Shared Documents/Reports"`.
#'
#' @return A data frame with one row per item and columns `name`, `type`
#'   (`"file"` or `"folder"`), `size` (bytes, `NA` for folders), and
#'   `modified` (a `POSIXct` timestamp).
#' @export
#'
#' @examples
#' \dontrun{
#' sp_list_items(auth, "Shared Documents")
#' sp_list_files(auth, "Shared Documents/Reports")
#' sp_list_folders(auth, "Shared Documents")
#' }
sp_list_items <- function(auth, path) {
  sp_check_auth(auth)
  switch(auth$backend,
    graph = sp_graph_list_items(auth, path),
    rest = sp_rest_list_items(auth, path)
  )
}

#' @rdname sp_list_items
#' @export
sp_list_files <- function(auth, path) {
  items <- sp_list_items(auth, path)
  items[items$type == "file", , drop = FALSE]
}

#' @rdname sp_list_items
#' @export
sp_list_folders <- function(auth, path) {
  items <- sp_list_items(auth, path)
  items[items$type == "folder", , drop = FALSE]
}

#' Download a file from SharePoint
#'
#' @param auth A `sharepointr_auth` object from [sharepoint_auth()].
#' @param path Library-relative path to the file, e.g.
#'   `"Shared Documents/Reports/Sales.xlsx"`.
#' @param dest Local path to write the file to, e.g. `"Sales.xlsx"` or
#'   `"downloads/Sales.xlsx"`. To download under a different name than the
#'   one on SharePoint, just give a `dest` with that name.
#' @param overwrite Whether to overwrite `dest` if it already exists locally.
#'
#' @return `dest`, invisibly.
#' @export
#'
#' @examples
#' \dontrun{
#' sp_download_file(auth, "Shared Documents/Reports/Sales.xlsx", "Sales.xlsx")
#' }
sp_download_file <- function(auth, path, dest, overwrite = FALSE) {
  sp_check_auth(auth)
  switch(auth$backend,
    graph = sp_graph_download_file(auth, path, dest, overwrite),
    rest = sp_rest_download_file(auth, path, dest, overwrite)
  )
}

#' Upload a file to SharePoint
#'
#' Files larger than 4MB are uploaded in chunks via a resumable upload
#' session when `auth` uses the `"graph"` backend.
#'
#' @param auth A `sharepointr_auth` object from [sharepoint_auth()].
#' @param local_path Path to the local file to upload.
#' @param path Destination library-relative path on SharePoint, including
#'   the file name, e.g. `"Shared Documents/Reports/Sales.xlsx"`.
#' @param overwrite Whether to overwrite an existing file at `path`.
#'
#' @return `path`, invisibly.
#' @export
#'
#' @examples
#' \dontrun{
#' sp_upload_file(auth, "Sales.xlsx", "Shared Documents/Reports/Sales.xlsx")
#' }
sp_upload_file <- function(auth, local_path, path, overwrite = TRUE) {
  sp_check_auth(auth)
  switch(auth$backend,
    graph = sp_graph_upload_file(auth, local_path, path, overwrite),
    rest = sp_rest_upload_file(auth, local_path, path, overwrite)
  )
}

#' Delete a file or folder on SharePoint
#'
#' Deleted items are moved to the site recycle bin, not permanently deleted.
#'
#' @param auth A `sharepointr_auth` object from [sharepoint_auth()].
#' @param path Library-relative path to the file or folder to delete.
#'
#' @return `path`, invisibly.
#' @export
#'
#' @examples
#' \dontrun{
#' sp_delete_item(auth, "Shared Documents/Reports/Old.xlsx")
#' }
sp_delete_item <- function(auth, path) {
  sp_check_auth(auth)
  switch(auth$backend,
    graph = sp_graph_delete_item(auth, path),
    rest = sp_rest_delete_item(auth, path)
  )
}

#' Create a folder on SharePoint
#'
#' @param auth A `sharepointr_auth` object from [sharepoint_auth()].
#' @param path Library-relative path of the folder to create, e.g.
#'   `"Shared Documents/Reports/2026"`. Parent folders must already exist.
#'
#' @return `path`, invisibly.
#' @export
#'
#' @examples
#' \dontrun{
#' sp_create_folder(auth, "Shared Documents/Reports/2026")
#' }
sp_create_folder <- function(auth, path) {
  sp_check_auth(auth)
  switch(auth$backend,
    graph = sp_graph_create_folder(auth, path),
    rest = sp_rest_create_folder(auth, path)
  )
}

#' Check whether a file exists on SharePoint
#'
#' @param auth A `sharepointr_auth` object from [sharepoint_auth()].
#' @param path Library-relative path to the file.
#'
#' @return `TRUE` or `FALSE`.
#' @export
#'
#' @examples
#' \dontrun{
#' sp_file_exists(auth, "Shared Documents/Reports/Sales.xlsx")
#' }
sp_file_exists <- function(auth, path) {
  sp_check_auth(auth)
  switch(auth$backend,
    graph = sp_graph_file_exists(auth, path),
    rest = sp_rest_file_exists(auth, path)
  )
}

sp_check_auth <- function(auth) {
  if (!inherits(auth, "sharepointr_auth")) {
    stop("`auth` must be created with `sharepoint_auth()`.", call. = FALSE)
  }
}

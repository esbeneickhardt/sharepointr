sp_rest_headers_read <- function(auth) {
  c(list(Accept = "application/json;odata=verbose"), sp_rest_auth_headers(auth))
}

sp_rest_headers_write <- function(auth) {
  c(sp_rest_headers_read(auth), list(`X-RequestDigest` = sp_get_digest(auth)))
}

sp_rest_folder_url <- function(auth, path, suffix = "") {
  encoded <- sp_url_encode_segments(sp_server_relative_url(auth, path))
  paste0(auth$site_url, "/_api/web/GetFolderByServerRelativeUrl('", encoded, "')", suffix)
}

sp_rest_file_url <- function(auth, path, suffix = "") {
  encoded <- sp_url_encode_segments(sp_server_relative_url(auth, path))
  paste0(auth$site_url, "/_api/web/GetFileByServerRelativeUrl('", encoded, "')", suffix)
}

sp_rest_list_items <- function(auth, path) {
  files_resp <- httr2::request(sp_rest_folder_url(auth, path, "/Files")) |>
    httr2::req_headers(!!!sp_rest_headers_read(auth)) |>
    sp_with_error_handling() |>
    httr2::req_perform()
  files <- sp_rest_resp_body_json(files_resp, auth)

  folders_resp <- httr2::request(sp_rest_folder_url(auth, path, "/Folders")) |>
    httr2::req_headers(!!!sp_rest_headers_read(auth)) |>
    sp_with_error_handling() |>
    httr2::req_perform()
  folders <- sp_rest_resp_body_json(folders_resp, auth)

  file_items <- files$d$results
  folder_items <- Filter(function(f) f$Name != "Forms", folders$d$results)

  rbind(
    sp_rest_folders_to_df(folder_items),
    sp_rest_files_to_df(file_items)
  )
}

sp_rest_files_to_df <- function(items) {
  if (length(items) == 0) {
    return(sp_empty_items_df())
  }
  data.frame(
    name = vapply(items, function(x) x$Name, character(1)),
    type = "file",
    size = vapply(items, function(x) as.numeric(x$Length %||% NA), numeric(1)),
    modified = as.POSIXct(
      vapply(items, function(x) x$TimeLastModified %||% NA_character_, character(1)),
      format = "%Y-%m-%dT%H:%M:%OSZ", tz = "UTC"
    ),
    stringsAsFactors = FALSE
  )
}

sp_rest_folders_to_df <- function(items) {
  if (length(items) == 0) {
    return(sp_empty_items_df())
  }
  data.frame(
    name = vapply(items, function(x) x$Name, character(1)),
    type = "folder",
    size = NA_real_,
    modified = as.POSIXct(
      vapply(items, function(x) x$TimeLastModified %||% NA_character_, character(1)),
      format = "%Y-%m-%dT%H:%M:%OSZ", tz = "UTC"
    ),
    stringsAsFactors = FALSE
  )
}

sp_rest_download_file <- function(auth, path, dest, overwrite = FALSE) {
  if (file.exists(dest) && !overwrite) {
    stop("`dest` already exists and `overwrite = FALSE`: ", dest, call. = FALSE)
  }
  httr2::request(sp_rest_file_url(auth, path, "/$value")) |>
    httr2::req_headers(!!!sp_rest_headers_read(auth)) |>
    sp_with_error_handling() |>
    httr2::req_perform(path = dest)
  invisible(dest)
}

sp_rest_file_exists <- function(auth, path) {
  resp <- httr2::request(sp_rest_file_url(auth, path, "/Exists")) |>
    httr2::req_headers(!!!sp_rest_headers_read(auth)) |>
    sp_with_error_handling() |>
    httr2::req_perform()
  isTRUE(sp_rest_resp_body_json(resp, auth)$d$Exists)
}

sp_rest_upload_file <- function(auth, local_path, path, overwrite = TRUE) {
  if (!file.exists(local_path)) {
    stop("`local_path` does not exist: ", local_path, call. = FALSE)
  }
  parts <- sp_split_library_path(path)
  if (!nzchar(parts$item_path)) {
    stop("`path` must include a destination file name, e.g. \"Shared Documents/Test.xlsx\".", call. = FALSE)
  }
  segments <- strsplit(parts$item_path, "/", fixed = TRUE)[[1]]
  file_name <- segments[length(segments)]
  folder_item_path <- paste(segments[-length(segments)], collapse = "/")
  folder_path <- if (nzchar(folder_item_path)) {
    paste(parts$library, folder_item_path, sep = "/")
  } else {
    parts$library
  }

  suffix <- paste0(
    "/Files/Add(url='", utils::URLencode(file_name, reserved = TRUE),
    "',overwrite=", tolower(as.character(overwrite)), ")"
  )

  httr2::request(sp_rest_folder_url(auth, folder_path, suffix)) |>
    httr2::req_headers(!!!sp_rest_headers_write(auth)) |>
    httr2::req_method("POST") |>
    httr2::req_body_file(local_path) |>
    sp_with_error_handling() |>
    httr2::req_perform()
  invisible(path)
}

# Deleted items go to the site recycle bin. `path` may point at a file or a
# folder; a file recycle is attempted first, falling back to a folder recycle.
sp_rest_delete_item <- function(auth, path) {
  file_error <- tryCatch(
    {
      httr2::request(sp_rest_file_url(auth, path, "/recycle()")) |>
        httr2::req_headers(!!!sp_rest_headers_write(auth)) |>
        httr2::req_method("POST") |>
        sp_with_error_handling() |>
        httr2::req_perform()
      NULL
    },
    error = function(e) e
  )

  if (is.null(file_error)) {
    return(invisible(path))
  }

  folder_error <- tryCatch(
    {
      httr2::request(sp_rest_folder_url(auth, path, "/recycle()")) |>
        httr2::req_headers(!!!sp_rest_headers_write(auth)) |>
        httr2::req_method("POST") |>
        sp_with_error_handling() |>
        httr2::req_perform()
      NULL
    },
    error = function(e) e
  )

  if (!is.null(folder_error)) {
    stop(file_error)
  }
  invisible(path)
}

sp_rest_create_folder <- function(auth, path) {
  encoded <- sp_url_encode_segments(sp_server_relative_url(auth, path))
  url <- paste0(auth$site_url, "/_api/web/folders/add('", encoded, "')")
  httr2::request(url) |>
    httr2::req_headers(!!!sp_rest_headers_write(auth)) |>
    httr2::req_method("POST") |>
    sp_with_error_handling() |>
    httr2::req_perform()
  invisible(path)
}

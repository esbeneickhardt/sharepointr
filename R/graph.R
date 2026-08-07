sp_graph_base <- "https://graph.microsoft.com/v1.0"

sp_graph_request <- function(auth, url) {
  token <- sp_get_token(auth)
  httr2::request(url) |>
    httr2::req_headers(Authorization = paste0("Bearer ", token)) |>
    sp_with_error_handling()
}

sp_resolve_site_id <- function(auth) {
  cached <- auth$cache$site_id
  if (!is.null(cached)) {
    return(cached)
  }

  url <- paste0(
    sp_graph_base, "/sites/", sp_site_host(auth), ":", sp_site_server_relative_path(auth)
  )
  resp <- sp_graph_request(auth, url) |> httr2::req_perform()
  site_id <- httr2::resp_body_json(resp)$id

  auth$cache$site_id <- site_id
  site_id
}

sp_resolve_drive_id <- function(auth, library) {
  key <- paste0("drive_", library)
  cached <- auth$cache[[key]]
  if (!is.null(cached)) {
    return(cached)
  }

  site_id <- sp_resolve_site_id(auth)
  url <- paste0(sp_graph_base, "/sites/", site_id, "/drives")
  resp <- sp_graph_request(auth, url) |> httr2::req_perform()
  drives <- httr2::resp_body_json(resp)$value

  target <- tolower(utils::URLdecode(library))
  matches <- Filter(function(d) {
    web_url <- tolower(utils::URLdecode(d$webUrl %||% ""))
    endsWith(web_url, paste0("/", target))
  }, drives)

  if (length(matches) == 0) {
    available <- vapply(drives, function(d) d$name %||% "", character(1))
    stop(
      "Could not find a document library named '", library, "' on this site.\n",
      "Available libraries: ", paste(available, collapse = ", "),
      call. = FALSE
    )
  }

  drive_id <- matches[[1]]$id
  auth$cache[[key]] <- drive_id
  drive_id
}

# Builds the Graph drive-item URL for `path` (a library-relative path, e.g.
# "Shared Documents/Reports/file.xlsx"), with an optional suffix appended
# after the item reference, e.g. "/children" or "/content".
sp_graph_item_url <- function(auth, path, suffix = "") {
  parts <- sp_split_library_path(path)
  drive_id <- sp_resolve_drive_id(auth, parts$library)
  base <- paste0(sp_graph_base, "/drives/", drive_id, "/root")
  if (nzchar(parts$item_path)) {
    base <- paste0(base, ":/", sp_url_encode_segments(parts$item_path), ":")
  }
  paste0(base, suffix)
}

sp_graph_list_items <- function(auth, path) {
  url <- sp_graph_item_url(
    auth, path,
    "/children?$select=name,size,lastModifiedDateTime,file,folder&$top=999"
  )
  items <- list()
  repeat {
    resp <- sp_graph_request(auth, url) |> httr2::req_perform()
    body <- httr2::resp_body_json(resp)
    items <- c(items, body$value)
    next_link <- body[["@odata.nextLink"]]
    if (is.null(next_link)) break
    url <- next_link
  }
  sp_graph_items_to_df(items)
}

sp_graph_items_to_df <- function(items) {
  if (length(items) == 0) {
    return(sp_empty_items_df())
  }
  data.frame(
    name = vapply(items, function(x) x$name, character(1)),
    type = vapply(items, function(x) if (!is.null(x$folder)) "folder" else "file", character(1)),
    size = vapply(items, function(x) as.numeric(x$size %||% NA), numeric(1)),
    modified = as.POSIXct(
      vapply(items, function(x) x$lastModifiedDateTime %||% NA_character_, character(1)),
      format = "%Y-%m-%dT%H:%M:%OSZ", tz = "UTC"
    ),
    stringsAsFactors = FALSE
  )
}

sp_graph_download_file <- function(auth, path, dest, overwrite = FALSE) {
  if (file.exists(dest) && !overwrite) {
    stop("`dest` already exists and `overwrite = FALSE`: ", dest, call. = FALSE)
  }
  url <- sp_graph_item_url(auth, path, "/content")
  sp_graph_request(auth, url) |>
    httr2::req_perform(path = dest)
  invisible(dest)
}

sp_graph_file_exists <- function(auth, path) {
  url <- sp_graph_item_url(auth, path)
  resp <- tryCatch(
    sp_graph_request(auth, url) |> httr2::req_perform(),
    httr2_http_404 = function(e) NULL
  )
  !is.null(resp)
}

sp_graph_upload_file <- function(auth, local_path, path, overwrite = TRUE) {
  if (!file.exists(local_path)) {
    stop("`local_path` does not exist: ", local_path, call. = FALSE)
  }
  if (!overwrite && sp_graph_file_exists(auth, path)) {
    stop("A file already exists at '", path, "' and `overwrite = FALSE`.", call. = FALSE)
  }

  size <- file.size(local_path)
  if (size <= 4 * 1024 * 1024) {
    sp_graph_upload_small(auth, local_path, path)
  } else {
    sp_graph_upload_large(auth, local_path, path, size)
  }
  invisible(path)
}

sp_graph_upload_small <- function(auth, local_path, path) {
  url <- sp_graph_item_url(auth, path, "/content")
  sp_graph_request(auth, url) |>
    httr2::req_method("PUT") |>
    httr2::req_body_file(local_path) |>
    httr2::req_perform()
  invisible(path)
}

# Files >4MB must go through a resumable upload session:
# https://learn.microsoft.com/graph/api/driveitem-createuploadsession
sp_graph_upload_large <- function(auth, local_path, path, size) {
  session_url <- sp_graph_item_url(auth, path, "/createUploadSession")
  resp <- sp_graph_request(auth, session_url) |>
    httr2::req_method("POST") |>
    httr2::req_body_json(list(item = list(`@microsoft.graph.conflictBehavior` = "replace"))) |>
    httr2::req_perform()
  upload_url <- httr2::resp_body_json(resp)$uploadUrl

  chunk_size <- 10 * 1024 * 1024 # must be a multiple of 320 KiB
  con <- file(local_path, "rb")
  on.exit(close(con))

  start <- 0
  repeat {
    chunk <- readBin(con, "raw", n = chunk_size)
    if (length(chunk) == 0) break
    end <- start + length(chunk) - 1

    httr2::request(upload_url) |>
      httr2::req_method("PUT") |>
      httr2::req_headers(
        `Content-Length` = as.character(length(chunk)),
        `Content-Range` = paste0("bytes ", start, "-", end, "/", size)
      ) |>
      httr2::req_body_raw(chunk) |>
      sp_with_error_handling() |>
      httr2::req_perform()

    start <- end + 1
  }
  invisible(path)
}

sp_graph_delete_item <- function(auth, path) {
  url <- sp_graph_item_url(auth, path)
  sp_graph_request(auth, url) |>
    httr2::req_method("DELETE") |>
    httr2::req_perform()
  invisible(path)
}

sp_graph_create_folder <- function(auth, path) {
  parts <- sp_split_library_path(path)
  if (!nzchar(parts$item_path)) {
    stop("`path` must include a folder name to create, e.g. \"Shared Documents/NewFolder\".", call. = FALSE)
  }
  segments <- strsplit(parts$item_path, "/", fixed = TRUE)[[1]]
  name <- segments[length(segments)]
  parent_item_path <- paste(segments[-length(segments)], collapse = "/")
  parent_path <- if (nzchar(parent_item_path)) {
    paste(parts$library, parent_item_path, sep = "/")
  } else {
    parts$library
  }

  url <- sp_graph_item_url(auth, parent_path, "/children")
  sp_graph_request(auth, url) |>
    httr2::req_method("POST") |>
    httr2::req_body_json(list(
      name = name,
      folder = list(),
      `@microsoft.graph.conflictBehavior` = "fail"
    )) |>
    httr2::req_perform()
  invisible(path)
}

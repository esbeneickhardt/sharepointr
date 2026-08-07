`%||%` <- function(x, y) if (is.null(x)) y else x

sp_normalize_relative_path <- function(path) {
  if (is.null(path) || is.na(path)) {
    return("")
  }
  gsub("^/+|/+$", "", path)
}

sp_split_library_path <- function(path) {
  path <- sp_normalize_relative_path(path)
  if (!nzchar(path)) {
    stop("`path` must include at least a document library name, e.g. \"Shared Documents\".", call. = FALSE)
  }
  parts <- strsplit(path, "/", fixed = TRUE)[[1]]
  list(
    library = parts[1],
    item_path = paste(parts[-1], collapse = "/")
  )
}

sp_url_encode_segments <- function(path) {
  if (!nzchar(path)) {
    return(path)
  }
  segments <- strsplit(path, "/", fixed = TRUE)[[1]]
  paste(vapply(segments, utils::URLencode, character(1), reserved = TRUE), collapse = "/")
}

sp_site_host <- function(auth) {
  sub("^https?://([^/]+).*$", "\\1", auth$site_url)
}

sp_site_server_relative_path <- function(auth) {
  sub("^https?://[^/]+", "", auth$site_url)
}

sp_server_relative_url <- function(auth, path) {
  path <- sp_normalize_relative_path(path)
  site_path <- sp_site_server_relative_path(auth)
  if (nzchar(path)) {
    paste0(site_path, "/", path)
  } else {
    site_path
  }
}

# Extracts a readable message from a Graph or SharePoint REST error body,
# used as the `body` callback for httr2::req_error().
sp_error_body <- function(resp) {
  ctype <- httr2::resp_content_type(resp)
  if (!identical(ctype, "application/json")) {
    return(NULL)
  }
  body <- tryCatch(httr2::resp_body_json(resp), error = function(e) NULL)
  if (is.null(body) || is.null(body$error)) {
    return(NULL)
  }
  err <- body$error
  if (is.character(err$message)) {
    err$message
  } else if (is.list(err$message) && !is.null(err$message$value)) {
    err$message$value
  } else {
    NULL
  }
}

sp_with_error_handling <- function(req) {
  httr2::req_error(req, body = sp_error_body)
}

sp_empty_items_df <- function() {
  data.frame(
    name = character(),
    type = character(),
    size = numeric(),
    modified = as.POSIXct(character()),
    stringsAsFactors = FALSE
  )
}

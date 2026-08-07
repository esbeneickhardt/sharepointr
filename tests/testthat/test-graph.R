test_that("sp_resolve_site_id requests the host:path lookup and caches the result", {
  auth <- make_graph_auth()
  auth$cache$site_id <- NULL

  call_count <- 0
  httr2::local_mocked_responses(function(req) {
    call_count <<- call_count + 1
    expect_match(req$url, "sites/contoso.sharepoint.com:/sites/MySite", fixed = TRUE)
    httr2::response_json(body = list(id = "resolved-site-id"))
  })

  expect_equal(sp_resolve_site_id(auth), "resolved-site-id")
  expect_equal(sp_resolve_site_id(auth), "resolved-site-id")
  expect_equal(call_count, 1, label = "site id should be cached")
})

test_that("sp_resolve_drive_id matches a drive by its webUrl suffix", {
  auth <- make_graph_auth()
  auth$cache[["drive_Shared Documents"]] <- NULL

  httr2::local_mocked_responses(function(req) {
    expect_match(req$url, "/sites/site1/drives", fixed = TRUE)
    httr2::response_json(body = list(value = list(
      list(id = "drive-other", name = "Other", webUrl = "https://contoso.sharepoint.com/sites/MySite/Other"),
      list(id = "drive-docs", name = "Documents", webUrl = "https://contoso.sharepoint.com/sites/MySite/Shared Documents")
    )))
  })

  expect_equal(sp_resolve_drive_id(auth, "Shared Documents"), "drive-docs")
})

test_that("sp_resolve_drive_id errors with available libraries when no match", {
  auth <- make_graph_auth()

  httr2::local_mocked_responses(function(req) {
    httr2::response_json(body = list(value = list(
      list(id = "drive-docs", name = "Documents", webUrl = "https://contoso.sharepoint.com/sites/MySite/Shared Documents")
    )))
  })

  expect_error(sp_resolve_drive_id(auth, "Nonexistent"), "Could not find")
})

test_that("sp_graph_item_url builds colon-addressed item URLs", {
  auth <- make_graph_auth()
  expect_equal(
    sp_graph_item_url(auth, "Shared Documents/Reports/Sales.xlsx", "/content"),
    "https://graph.microsoft.com/v1.0/drives/drive1/root:/Reports/Sales.xlsx:/content"
  )
  expect_equal(
    sp_graph_item_url(auth, "Shared Documents", "/children"),
    "https://graph.microsoft.com/v1.0/drives/drive1/root/children"
  )
})

test_that("sp_graph_list_items parses files and folders into a data frame", {
  auth <- make_graph_auth()

  httr2::local_mocked_responses(list(
    httr2::response_json(body = list(value = list(
      list(name = "Report.xlsx", size = 1234, lastModifiedDateTime = "2026-01-01T10:00:00Z", file = list()),
      list(name = "Archive", size = 0, lastModifiedDateTime = "2026-01-02T10:00:00Z", folder = list())
    )))
  ))

  df <- sp_graph_list_items(auth, "Shared Documents")
  expect_equal(nrow(df), 2)
  expect_setequal(df$type, c("file", "folder"))
  expect_equal(df$name[df$type == "file"], "Report.xlsx")
})

test_that("sp_graph_list_items follows @odata.nextLink pagination", {
  auth <- make_graph_auth()

  call_count <- 0
  httr2::local_mocked_responses(function(req) {
    call_count <<- call_count + 1
    if (call_count == 1) {
      httr2::response_json(body = list(
        value = list(list(name = "A.txt", size = 1, lastModifiedDateTime = "2026-01-01T10:00:00Z", file = list())),
        `@odata.nextLink` = "https://graph.microsoft.com/v1.0/next-page"
      ))
    } else {
      httr2::response_json(body = list(
        value = list(list(name = "B.txt", size = 2, lastModifiedDateTime = "2026-01-01T10:00:00Z", file = list()))
      ))
    }
  })

  df <- sp_graph_list_items(auth, "Shared Documents")
  expect_equal(call_count, 2)
  expect_setequal(df$name, c("A.txt", "B.txt"))
})

test_that("sp_graph_download_file requests the item content endpoint with auth header", {
  auth <- make_graph_auth()
  dest <- withr::local_tempfile()

  captured <- NULL
  httr2::local_mocked_responses(function(req) {
    captured <<- req
    httr2::response(status_code = 200, body = charToRaw("filedata"))
  })

  sp_graph_download_file(auth, "Shared Documents/Sales.xlsx", dest, overwrite = TRUE)

  expect_match(captured$url, "drive1/root:/Sales.xlsx:/content", fixed = TRUE)
  expect_true("Authorization" %in% names(captured$headers))
})

test_that("sp_graph_download_file refuses to overwrite existing files by default", {
  auth <- make_graph_auth()
  dest <- withr::local_tempfile()
  writeLines("existing", dest)

  expect_error(
    sp_graph_download_file(auth, "Shared Documents/Sales.xlsx", dest, overwrite = FALSE),
    "already exists"
  )
})

test_that("sp_graph_file_exists returns TRUE/FALSE based on response status", {
  auth <- make_graph_auth()

  httr2::local_mocked_responses(list(httr2::response(status_code = 200)))
  expect_true(sp_graph_file_exists(auth, "Shared Documents/Sales.xlsx"))

  httr2::local_mocked_responses(list(httr2::response(status_code = 404)))
  expect_false(sp_graph_file_exists(auth, "Shared Documents/Missing.xlsx"))
})

test_that("sp_graph_upload_file PUTs small files directly to the content endpoint", {
  auth <- make_graph_auth()
  local_path <- withr::local_tempfile()
  writeBin(as.raw(1:10), local_path)

  captured <- NULL
  httr2::local_mocked_responses(function(req) {
    captured <<- req
    httr2::response_json(status_code = 201, body = list(name = "Sales.xlsx"))
  })

  sp_graph_upload_file(auth, local_path, "Shared Documents/Sales.xlsx")
  expect_match(captured$url, "root:/Sales.xlsx:/content", fixed = TRUE)
  expect_equal(captured$method, "PUT")
})

test_that("sp_graph_upload_file uses a chunked upload session for large files", {
  auth <- make_graph_auth()
  local_path <- withr::local_tempfile()
  writeBin(as.raw(rep(1, 5 * 1024 * 1024)), local_path) # 5 MiB, above the 4 MiB threshold

  requests <- list()
  httr2::local_mocked_responses(function(req) {
    requests[[length(requests) + 1]] <<- req
    if (grepl("createUploadSession", req$url, fixed = TRUE)) {
      httr2::response_json(body = list(uploadUrl = "https://upload.example.com/session1"))
    } else {
      httr2::response_json(status_code = 200, body = list(name = "Big.bin"))
    }
  })

  sp_graph_upload_file(auth, local_path, "Shared Documents/Big.bin")

  expect_length(requests, 2)
  expect_match(requests[[1]]$url, "createUploadSession", fixed = TRUE)
  expect_equal(requests[[2]]$url, "https://upload.example.com/session1")
  expect_equal(requests[[2]]$headers[["Content-Range"]], paste0("bytes 0-", 5 * 1024 * 1024 - 1, "/", 5 * 1024 * 1024))
})

test_that("sp_graph_delete_item sends a DELETE to the item endpoint", {
  auth <- make_graph_auth()
  captured <- NULL
  httr2::local_mocked_responses(function(req) {
    captured <<- req
    httr2::response(status_code = 204)
  })

  sp_graph_delete_item(auth, "Shared Documents/Old.xlsx")
  expect_equal(captured$method, "DELETE")
  expect_match(captured$url, "root:/Old.xlsx:", fixed = TRUE)
})

test_that("sp_graph_create_folder POSTs to the parent's children endpoint", {
  auth <- make_graph_auth()
  captured <- NULL
  httr2::local_mocked_responses(function(req) {
    captured <<- req
    httr2::response_json(status_code = 201, body = list(name = "2026"))
  })

  sp_graph_create_folder(auth, "Shared Documents/Reports/2026")
  expect_match(captured$url, "root:/Reports:/children", fixed = TRUE)
  expect_equal(captured$body$data$name, "2026")
})

test_that("sp_graph_create_folder errors when no folder name is given", {
  auth <- make_graph_auth()
  expect_error(sp_graph_create_folder(auth, "Shared Documents"), "must include a folder name")
})

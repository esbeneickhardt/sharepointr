test_that("sp_rest_folder_url and sp_rest_file_url build server-relative URLs", {
  auth <- make_rest_auth()
  expect_equal(
    sp_rest_folder_url(auth, "Shared Documents/test", "/Files"),
    "https://contoso.sharepoint.com/sites/MySite/_api/web/GetFolderByServerRelativeUrl('/sites/MySite/Shared%20Documents/test')/Files"
  )
  expect_equal(
    sp_rest_file_url(auth, "Shared Documents/test/a.txt", "/$value"),
    "https://contoso.sharepoint.com/sites/MySite/_api/web/GetFileByServerRelativeUrl('/sites/MySite/Shared%20Documents/test/a.txt')/$value"
  )
})

test_that("sp_rest_headers_read sends a Cookie header for cookie auth, not Authorization", {
  auth <- make_cookie_auth()
  headers <- sp_rest_headers_read(auth)
  expect_equal(headers$Cookie, "FedAuth=fedauth-value; rtFa=rtfa-value")
  expect_null(headers$Authorization)
})

test_that("sp_rest_list_items works over cookie auth and sends the cookie header", {
  auth <- make_cookie_auth()
  captured <- NULL
  httr2::local_mocked_responses(function(req) {
    captured <<- req
    if (grepl("/Files$", req$url)) {
      httr2::response_json(body = list(d = list(results = list(
        list(Name = "A.txt", Length = "1", TimeLastModified = "2026-01-01T10:00:00Z")
      ))))
    } else {
      httr2::response_json(body = list(d = list(results = list())))
    }
  })

  df <- sp_rest_list_items(auth, "Shared Documents")
  expect_equal(df$name, "A.txt")
  expect_true("Cookie" %in% names(captured$headers))
  expect_false("Authorization" %in% names(captured$headers))
})

test_that("an expired cookie session raises a clear error instead of a JSON parse error", {
  auth <- make_cookie_auth()
  httr2::local_mocked_responses(list(
    httr2::response(
      status_code = 200,
      headers = list(`Content-Type` = "text/html"),
      body = charToRaw("<html>Sign in to your account</html>")
    )
  ))

  expect_error(sp_rest_list_items(auth, "Shared Documents"), "expired")
})

test_that("sp_rest_list_items combines Files and Folders, dropping the 'Forms' folder", {
  auth <- make_rest_auth()

  httr2::local_mocked_responses(function(req) {
    if (grepl("/Files$", req$url)) {
      httr2::response_json(body = list(d = list(results = list(
        list(Name = "Report.xlsx", Length = "1234", TimeLastModified = "2026-01-01T10:00:00Z")
      ))))
    } else {
      httr2::response_json(body = list(d = list(results = list(
        list(Name = "Forms", TimeLastModified = "2026-01-01T10:00:00Z"),
        list(Name = "Archive", TimeLastModified = "2026-01-02T10:00:00Z")
      ))))
    }
  })

  df <- sp_rest_list_items(auth, "Shared Documents")
  expect_equal(nrow(df), 2)
  expect_false("Forms" %in% df$name)
  expect_setequal(df$type, c("file", "folder"))
})

test_that("sp_rest_download_file requests the file's $value endpoint", {
  auth <- make_rest_auth()
  dest <- withr::local_tempfile()

  captured <- NULL
  httr2::local_mocked_responses(function(req) {
    captured <<- req
    httr2::response(status_code = 200, body = charToRaw("filedata"))
  })

  sp_rest_download_file(auth, "Shared Documents/Sales.xlsx", dest, overwrite = TRUE)
  expect_match(captured$url, "GetFileByServerRelativeUrl", fixed = TRUE)
  expect_match(captured$url, "$value", fixed = TRUE)
  expect_true("Authorization" %in% names(captured$headers))
})

test_that("sp_rest_file_exists reads the Exists property", {
  auth <- make_rest_auth()

  httr2::local_mocked_responses(list(httr2::response_json(body = list(d = list(Exists = TRUE)))))
  expect_true(sp_rest_file_exists(auth, "Shared Documents/Sales.xlsx"))

  httr2::local_mocked_responses(list(httr2::response_json(body = list(d = list(Exists = FALSE)))))
  expect_false(sp_rest_file_exists(auth, "Shared Documents/Missing.xlsx"))
})

test_that("sp_rest_upload_file POSTs to Files/Add with the digest header", {
  auth <- make_rest_auth()
  local_path <- withr::local_tempfile()
  writeBin(as.raw(1:10), local_path)

  captured <- NULL
  httr2::local_mocked_responses(function(req) {
    captured <<- req
    httr2::response_json(status_code = 200, body = list(d = list(Name = "Sales.xlsx")))
  })

  sp_rest_upload_file(auth, local_path, "Shared Documents/Reports/Sales.xlsx")
  expect_match(captured$url, "Files/Add(url='Sales.xlsx',overwrite=true)", fixed = TRUE)
  expect_equal(captured$headers[["X-RequestDigest"]], "digest-val")
})

test_that("sp_rest_upload_file errors when no destination file name is given", {
  auth <- make_rest_auth()
  local_path <- withr::local_tempfile()
  writeBin(as.raw(1:10), local_path)
  expect_error(sp_rest_upload_file(auth, local_path, "Shared Documents"), "must include a destination")
})

test_that("sp_rest_delete_item recycles the file when the file endpoint succeeds", {
  auth <- make_rest_auth()

  call_count <- 0
  httr2::local_mocked_responses(function(req) {
    call_count <<- call_count + 1
    expect_match(req$url, "GetFileByServerRelativeUrl", fixed = TRUE)
    httr2::response(status_code = 200)
  })

  sp_rest_delete_item(auth, "Shared Documents/Old.xlsx")
  expect_equal(call_count, 1)
})

test_that("sp_rest_delete_item falls back to the folder endpoint when the file recycle fails", {
  auth <- make_rest_auth()

  httr2::local_mocked_responses(function(req) {
    if (grepl("GetFileByServerRelativeUrl", req$url, fixed = TRUE)) {
      httr2::response_json(status_code = 404, body = list(error = list(message = list(value = "File Not Found"))))
    } else {
      httr2::response(status_code = 200)
    }
  })

  expect_no_error(sp_rest_delete_item(auth, "Shared Documents/OldFolder"))
})

test_that("sp_rest_delete_item surfaces the original file error when both fail", {
  auth <- make_rest_auth()

  httr2::local_mocked_responses(function(req) {
    httr2::response_json(status_code = 404, body = list(error = list(message = list(value = "Not Found"))))
  })

  expect_error(sp_rest_delete_item(auth, "Shared Documents/Missing"), "Not Found")
})

test_that("sp_rest_create_folder POSTs to folders/add with the server-relative url", {
  auth <- make_rest_auth()
  captured <- NULL
  httr2::local_mocked_responses(function(req) {
    captured <<- req
    httr2::response_json(status_code = 200, body = list(d = list(Name = "2026")))
  })

  sp_rest_create_folder(auth, "Shared Documents/Reports/2026")
  expect_match(captured$url, "folders/add", fixed = TRUE)
  expect_equal(captured$headers[["X-RequestDigest"]], "digest-val")
})

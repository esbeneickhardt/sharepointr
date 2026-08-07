test_that("sp_normalize_relative_path strips leading/trailing slashes", {
  expect_equal(sp_normalize_relative_path("/a/b/"), "a/b")
  expect_equal(sp_normalize_relative_path("a/b"), "a/b")
  expect_equal(sp_normalize_relative_path(""), "")
  expect_equal(sp_normalize_relative_path(NULL), "")
})

test_that("sp_split_library_path splits library from item path", {
  parts <- sp_split_library_path("Shared Documents/Reports/Sales.xlsx")
  expect_equal(parts$library, "Shared Documents")
  expect_equal(parts$item_path, "Reports/Sales.xlsx")

  parts2 <- sp_split_library_path("Shared Documents")
  expect_equal(parts2$library, "Shared Documents")
  expect_equal(parts2$item_path, "")
})

test_that("sp_split_library_path errors on empty path", {
  expect_error(sp_split_library_path(""), "must include")
})

test_that("sp_url_encode_segments encodes spaces per segment, keeps slashes", {
  expect_equal(sp_url_encode_segments("Shared Documents/My File.xlsx"), "Shared%20Documents/My%20File.xlsx")
  expect_equal(sp_url_encode_segments(""), "")
})

test_that("sp_site_server_relative_path and sp_server_relative_url work", {
  auth <- sharepoint_auth(
    tenant_id = "tenant.onmicrosoft.com",
    client_id = "client",
    site_url = "https://contoso.sharepoint.com/sites/MySite",
    client_secret = "secret"
  )
  expect_equal(sp_site_host(auth), "contoso.sharepoint.com")
  expect_equal(sp_site_server_relative_path(auth), "/sites/MySite")
  expect_equal(
    sp_server_relative_url(auth, "Shared Documents/test"),
    "/sites/MySite/Shared Documents/test"
  )
  expect_equal(sp_server_relative_url(auth, ""), "/sites/MySite")
})

test_that("sp_error_body extracts Graph-style and REST-style error messages", {
  graph_resp <- httr2::response_json(
    status_code = 404,
    body = list(error = list(code = "itemNotFound", message = "The item was not found"))
  )
  expect_equal(sp_error_body(graph_resp), "The item was not found")

  rest_resp <- httr2::response_json(
    status_code = 400,
    body = list(error = list(message = list(lang = "en-US", value = "File not found")))
  )
  expect_equal(sp_error_body(rest_resp), "File not found")

  non_json_resp <- httr2::response(status_code = 500, headers = list(`Content-Type` = "text/plain"))
  expect_null(sp_error_body(non_json_resp))
})

test_that("sp_empty_items_df has the right shape", {
  df <- sp_empty_items_df()
  expect_equal(nrow(df), 0)
  expect_equal(names(df), c("name", "type", "size", "modified"))
})

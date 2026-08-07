test_that("public functions reject an auth object not made by sharepoint_auth()", {
  expect_error(sp_list_items("not-an-auth", "Shared Documents"), "sharepoint_auth")
  expect_error(sp_download_file(list(), "a", "b"), "sharepoint_auth")
})

test_that("sp_list_items dispatches to the graph backend", {
  auth <- make_graph_auth()
  httr2::local_mocked_responses(list(
    httr2::response_json(body = list(value = list(
      list(name = "A.txt", size = 1, lastModifiedDateTime = "2026-01-01T10:00:00Z", file = list())
    )))
  ))
  df <- sp_list_items(auth, "Shared Documents")
  expect_equal(df$name, "A.txt")
})

test_that("sp_list_items dispatches to the rest backend", {
  auth <- make_rest_auth()
  httr2::local_mocked_responses(function(req) {
    if (grepl("/Files$", req$url)) {
      httr2::response_json(body = list(d = list(results = list(
        list(Name = "A.txt", Length = "1", TimeLastModified = "2026-01-01T10:00:00Z")
      ))))
    } else {
      httr2::response_json(body = list(d = list(results = list())))
    }
  })
  df <- sp_list_items(auth, "Shared Documents")
  expect_equal(df$name, "A.txt")
})

test_that("sp_list_files and sp_list_folders filter sp_list_items", {
  auth <- make_graph_auth()
  httr2::local_mocked_responses(list(
    httr2::response_json(body = list(value = list(
      list(name = "A.txt", size = 1, lastModifiedDateTime = "2026-01-01T10:00:00Z", file = list()),
      list(name = "Sub", size = 0, lastModifiedDateTime = "2026-01-01T10:00:00Z", folder = list())
    )))
  ))
  expect_equal(sp_list_files(auth, "Shared Documents")$name, "A.txt")

  httr2::local_mocked_responses(list(
    httr2::response_json(body = list(value = list(
      list(name = "A.txt", size = 1, lastModifiedDateTime = "2026-01-01T10:00:00Z", file = list()),
      list(name = "Sub", size = 0, lastModifiedDateTime = "2026-01-01T10:00:00Z", folder = list())
    )))
  ))
  expect_equal(sp_list_folders(auth, "Shared Documents")$name, "Sub")
})

test_that("sp_file_exists dispatches per backend", {
  auth_graph <- make_graph_auth()
  httr2::local_mocked_responses(list(httr2::response(status_code = 200)))
  expect_true(sp_file_exists(auth_graph, "Shared Documents/A.txt"))

  auth_rest <- make_rest_auth()
  httr2::local_mocked_responses(list(httr2::response_json(body = list(d = list(Exists = TRUE)))))
  expect_true(sp_file_exists(auth_rest, "Shared Documents/A.txt"))
})

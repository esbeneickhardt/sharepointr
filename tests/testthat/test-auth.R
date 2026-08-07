test_that("sharepoint_auth validates required arguments", {
  expect_error(
    sharepoint_auth(client_id = "c", site_url = "https://x.sharepoint.com/sites/s", client_secret = "s"),
    "required"
  )
})

test_that("sharepoint_auth requires exactly one credential type", {
  expect_error(
    sharepoint_auth(
      tenant_id = "t", client_id = "c", site_url = "https://x.sharepoint.com/sites/s"
    ),
    "Provide exactly one"
  )
  expect_error(
    sharepoint_auth(
      tenant_id = "t", client_id = "c", site_url = "https://x.sharepoint.com/sites/s",
      client_secret = "s", certificate_path = "cert.pem"
    ),
    "not more than one"
  )
})

test_that("sharepoint_auth cookie authentication validates and forces the rest backend", {
  expect_error(
    sharepoint_auth(site_url = "https://x.sharepoint.com/sites/s", fedauth = "f"),
    "requires both"
  )
  expect_error(
    sharepoint_auth(
      site_url = "https://x.sharepoint.com/sites/s",
      fedauth = "f", rtfa = "r", backend = "graph"
    ),
    "only supports"
  )
  expect_error(
    sharepoint_auth(
      site_url = "https://x.sharepoint.com/sites/s",
      fedauth = "f", rtfa = "r", client_secret = "s"
    ),
    "not more than one"
  )

  auth <- sharepoint_auth(
    site_url = "https://x.sharepoint.com/sites/s",
    fedauth = "f", rtfa = "r"
  )
  expect_equal(auth$backend, "rest")
  expect_equal(auth$auth_method, "cookie")
  expect_null(auth$tenant_id)
  expect_null(auth$client_id)
})

test_that("sharepoint_auth defaults to the graph backend and strips trailing slashes", {
  auth <- sharepoint_auth(
    tenant_id = "t", client_id = "c",
    site_url = "https://x.sharepoint.com/sites/s/",
    client_secret = "secret"
  )
  expect_equal(auth$backend, "graph")
  expect_equal(auth$site_url, "https://x.sharepoint.com/sites/s")
  expect_s3_class(auth, "sharepointr_auth")
})

test_that("print.sharepointr_auth prints without erroring", {
  auth <- sharepoint_auth(
    tenant_id = "t", client_id = "c",
    site_url = "https://x.sharepoint.com/sites/s",
    client_secret = "secret"
  )
  expect_output(print(auth), "sharepointr_auth")
})

test_that("sp_token_scope differs between backends", {
  auth_graph <- sharepoint_auth(
    tenant_id = "t", client_id = "c",
    site_url = "https://contoso.sharepoint.com/sites/s",
    client_secret = "secret", backend = "graph"
  )
  auth_rest <- sharepoint_auth(
    tenant_id = "t", client_id = "c",
    site_url = "https://contoso.sharepoint.com/sites/s",
    client_secret = "secret", backend = "rest"
  )
  expect_equal(sp_token_scope(auth_graph), "https://graph.microsoft.com/.default")
  expect_equal(sp_token_scope(auth_rest), "https://contoso.sharepoint.com/.default")
})

test_that("sp_get_token fetches and caches a token (client secret flow)", {
  auth <- sharepoint_auth(
    tenant_id = "t", client_id = "c",
    site_url = "https://contoso.sharepoint.com/sites/s",
    client_secret = "secret"
  )

  call_count <- 0
  httr2::local_mocked_responses(function(req) {
    call_count <<- call_count + 1
    httr2::response_json(body = list(access_token = "tok-1", expires_in = 3600))
  })

  expect_equal(sp_get_token(auth), "tok-1")
  expect_equal(sp_get_token(auth), "tok-1")
  expect_equal(call_count, 1, label = "token should be cached, not refetched")
})

test_that("sp_get_token refreshes an expired token", {
  auth <- sharepoint_auth(
    tenant_id = "t", client_id = "c",
    site_url = "https://contoso.sharepoint.com/sites/s",
    client_secret = "secret"
  )
  auth$cache$token <- list(access_token = "stale", expires_at = Sys.time() - 10)

  httr2::local_mocked_responses(list(
    httr2::response_json(body = list(access_token = "fresh", expires_in = 3600))
  ))

  expect_equal(sp_get_token(auth), "fresh")
})

test_that("sp_get_digest fetches and caches a digest value (rest backend)", {
  auth <- sharepoint_auth(
    tenant_id = "t", client_id = "c",
    site_url = "https://contoso.sharepoint.com/sites/s",
    client_secret = "secret", backend = "rest"
  )
  auth$cache$token <- list(access_token = "tok", expires_at = Sys.time() + 3600)

  call_count <- 0
  httr2::local_mocked_responses(function(req) {
    call_count <<- call_count + 1
    httr2::response_json(body = list(d = list(GetContextWebInformation = list(
      FormDigestValue = "digest-value,1234",
      FormDigestTimeoutSeconds = 1800
    ))))
  })

  expect_equal(sp_get_digest(auth), "digest-value")
  expect_equal(sp_get_digest(auth), "digest-value")
  expect_equal(call_count, 1, label = "digest should be cached, not refetched")
})

test_that("sp_build_client_assertion produces a valid, verifiable JWT", {
  cert_path <- test_path("fixtures", "test-cert.pem")
  skip_if_not(file.exists(cert_path), "test certificate fixture missing")

  auth <- sharepoint_auth(
    tenant_id = "tenant.onmicrosoft.com",
    client_id = "11111111-2222-3333-4444-555555555555",
    site_url = "https://contoso.sharepoint.com/sites/MySite",
    certificate_path = cert_path,
    backend = "graph"
  )

  assertion <- sp_build_client_assertion(auth)
  parts <- strsplit(assertion, ".", fixed = TRUE)[[1]]
  expect_length(parts, 3)

  cert <- openssl::read_cert(cert_path)
  claim <- jose::jwt_decode_sig(assertion, cert)
  expect_equal(claim$iss, auth$client_id)
  expect_equal(claim$sub, auth$client_id)
  expect_equal(claim$aud, sp_token_url(auth))
})

test_that("sp_fetch_token_certificate posts a client assertion and parses the token", {
  cert_path <- test_path("fixtures", "test-cert.pem")
  skip_if_not(file.exists(cert_path), "test certificate fixture missing")

  auth <- sharepoint_auth(
    tenant_id = "t", client_id = "c",
    site_url = "https://contoso.sharepoint.com/sites/s",
    certificate_path = cert_path
  )

  seen_data <- NULL
  httr2::local_mocked_responses(function(req) {
    seen_data <<- req$body$data
    httr2::response_json(body = list(access_token = "cert-tok", expires_in = 3600))
  })

  expect_equal(sp_get_token(auth), "cert-tok")
  expect_equal(
    utils::URLdecode(as.character(seen_data$client_assertion_type)),
    "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
  )
  assertion <- utils::URLdecode(as.character(seen_data$client_assertion))
  expect_length(strsplit(assertion, ".", fixed = TRUE)[[1]], 3)
})

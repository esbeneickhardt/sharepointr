#' Create a SharePoint authentication object
#'
#' Builds an authentication object and passes it to every other `sp_*()`
#' function. Supports three authentication methods:
#'
#' - **Cookie** (`fedauth` + `rtfa`): reuses an already-signed-in browser
#'   session. No app registration or admin involvement needed - just your
#'   own SharePoint access. REST backend only. See "Cookie authentication"
#'   below for how to get the cookie values. Unofficial (not an API
#'   Microsoft documents or guarantees), and the cookies expire on your
#'   tenant's normal sign-in-frequency schedule, so this suits interactive
#'   or short-lived scripts better than long-running unattended jobs.
#' - **Client secret**: an Entra ID app registration authenticating via
#'   OAuth2 client credentials. Needs an admin to create the app
#'   registration and consent to its permissions once; after that it's
#'   fully unattended. Works with both backends.
#' - **Certificate**: same as client secret, but with a longer-lived,
#'   harder-to-leak credential. Works with both backends.
#'
#' Whichever method you use, the object lazily fetches and caches its
#' access token (and, on the REST backend, its request digest), refreshing
#' automatically as they expire - build it once per session.
#'
#' @section Cookie authentication:
#' 1. Sign in to your SharePoint site in a browser as you normally would.
#' 2. Open developer tools (F12) > **Application**/**Storage** > **Cookies**
#'    > your SharePoint domain.
#' 3. Copy the values of the `FedAuth` and `rtFa` cookies into `fedauth` and
#'    `rtfa` below.
#'
#' Treat these values like a password - anyone with them can act as you on
#' this site until they expire.
#'
#' @param site_url Full URL of the SharePoint site, e.g.
#'   `"https://contoso.sharepoint.com/sites/MySite"`.
#' @param tenant_id Entra ID tenant ID (GUID) or verified domain, e.g.
#'   `"contoso.onmicrosoft.com"`. Required for `client_secret` /
#'   `certificate_path`, not used for cookie authentication.
#' @param client_id Application (client) ID of the app registration.
#'   Required for `client_secret` / `certificate_path`, not used for cookie
#'   authentication.
#' @param client_secret Client secret value. Mutually exclusive with
#'   `certificate_path` and `fedauth`/`rtfa`.
#' @param certificate_path Path to a PEM file containing a certificate and
#'   its private key (e.g. `cat key.pem cert.pem > combined.pem`), or just
#'   the private key. Mutually exclusive with `client_secret` and
#'   `fedauth`/`rtfa`. PKCS12/PFX files are not supported.
#' @param certificate_password Password protecting `certificate_path`, if
#'   any.
#' @param fedauth,rtfa Values of the `FedAuth` and `rtFa` cookies from an
#'   authenticated browser session (see "Cookie authentication" below).
#'   Both required together; mutually exclusive with `client_secret` and
#'   `certificate_path`. Forces `backend = "rest"`.
#' @param backend Which API to use: `"graph"` (default, recommended for
#'   `client_secret`/`certificate_path`) or `"rest"` (the SharePoint REST
#'   API, `_api/web`; required for cookie authentication).
#'
#' @return An object of class `sharepointr_auth`.
#' @export
#'
#' @examples
#' \dontrun{
#' # Cookie (no app registration needed)
#' auth <- sharepoint_auth(
#'   site_url = "https://contoso.sharepoint.com/sites/MySite",
#'   fedauth = Sys.getenv("SHAREPOINT_FEDAUTH"),
#'   rtfa = Sys.getenv("SHAREPOINT_RTFA")
#' )
#'
#' # Client secret
#' auth <- sharepoint_auth(
#'   site_url = "https://contoso.sharepoint.com/sites/MySite",
#'   tenant_id = "contoso.onmicrosoft.com",
#'   client_id = "11111111-2222-3333-4444-555555555555",
#'   client_secret = Sys.getenv("SHAREPOINT_CLIENT_SECRET")
#' )
#'
#' # Certificate
#' auth <- sharepoint_auth(
#'   site_url = "https://contoso.sharepoint.com/sites/MySite",
#'   tenant_id = "contoso.onmicrosoft.com",
#'   client_id = "11111111-2222-3333-4444-555555555555",
#'   certificate_path = "combined.pem",
#'   certificate_password = Sys.getenv("SHAREPOINT_CERT_PASSWORD")
#' )
#'
#' sp_list_items(auth, "Shared Documents")
#' }
sharepoint_auth <- function(site_url,
                             tenant_id = NULL,
                             client_id = NULL,
                             client_secret = NULL,
                             certificate_path = NULL,
                             certificate_password = NULL,
                             fedauth = NULL,
                             rtfa = NULL,
                             backend = c("graph", "rest")) {
  backend_given <- !missing(backend)
  backend <- match.arg(backend)

  if (missing(site_url)) {
    stop("`site_url` is required.", call. = FALSE)
  }

  methods_given <- c(
    secret = !is.null(client_secret),
    certificate = !is.null(certificate_path),
    cookie = !is.null(fedauth) || !is.null(rtfa)
  )
  if (sum(methods_given) == 0) {
    stop(
      "Provide exactly one authentication method: `client_secret`, ",
      "`certificate_path`, or both `fedauth` and `rtfa`.",
      call. = FALSE
    )
  }
  if (sum(methods_given) > 1) {
    stop(
      "Provide exactly one authentication method, not more than one: ",
      paste(names(methods_given)[methods_given], collapse = ", "),
      " were all supplied.",
      call. = FALSE
    )
  }
  auth_method <- names(methods_given)[methods_given]

  if (auth_method == "cookie") {
    if (is.null(fedauth) || is.null(rtfa)) {
      stop("Cookie authentication requires both `fedauth` and `rtfa`.", call. = FALSE)
    }
    if (backend_given && backend != "rest") {
      stop(
        "Cookie authentication only supports `backend = \"rest\"` - ",
        "Microsoft Graph requires an OAuth token.",
        call. = FALSE
      )
    }
    backend <- "rest"
  } else if (is.null(tenant_id) || is.null(client_id)) {
    stop(
      "`tenant_id` and `client_id` are required for client_secret/certificate authentication.",
      call. = FALSE
    )
  }

  structure(
    list(
      tenant_id = tenant_id,
      client_id = client_id,
      site_url = sub("/+$", "", site_url),
      client_secret = client_secret,
      certificate_path = certificate_path,
      certificate_password = certificate_password,
      fedauth = fedauth,
      rtfa = rtfa,
      auth_method = unname(auth_method),
      backend = backend,
      cache = new.env(parent = emptyenv())
    ),
    class = "sharepointr_auth"
  )
}

#' @export
print.sharepointr_auth <- function(x, ...) {
  cat("<sharepointr_auth>\n")
  cat("  site:    ", x$site_url, "\n", sep = "")
  cat("  backend: ", x$backend, "\n", sep = "")
  cat("  method:  ", x$auth_method, "\n", sep = "")
  if (!is.null(x$client_id)) {
    cat("  client:  ", x$client_id, "\n", sep = "")
  }
  invisible(x)
}

sp_token_scope <- function(auth) {
  if (auth$backend == "graph") {
    "https://graph.microsoft.com/.default"
  } else {
    paste0("https://", sp_site_host(auth), "/.default")
  }
}

sp_token_url <- function(auth) {
  paste0("https://login.microsoftonline.com/", auth$tenant_id, "/oauth2/v2.0/token")
}

#' @return A bearer token string, fetching or refreshing it as needed.
#' @noRd
sp_get_token <- function(auth) {
  cached <- auth$cache$token
  if (!is.null(cached) && cached$expires_at > Sys.time() + 60) {
    return(cached$access_token)
  }

  token <- if (auth$auth_method == "secret") {
    sp_fetch_token_secret(auth)
  } else {
    sp_fetch_token_certificate(auth)
  }

  auth$cache$token <- token
  token$access_token
}

sp_fetch_token_secret <- function(auth) {
  resp <- httr2::request(sp_token_url(auth)) |>
    httr2::req_body_form(
      grant_type = "client_credentials",
      client_id = auth$client_id,
      client_secret = auth$client_secret,
      scope = sp_token_scope(auth)
    ) |>
    sp_with_error_handling() |>
    httr2::req_perform()

  sp_parse_token_response(resp)
}

sp_fetch_token_certificate <- function(auth) {
  assertion <- sp_build_client_assertion(auth)

  resp <- httr2::request(sp_token_url(auth)) |>
    httr2::req_body_form(
      grant_type = "client_credentials",
      client_id = auth$client_id,
      client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
      client_assertion = assertion,
      scope = sp_token_scope(auth)
    ) |>
    sp_with_error_handling() |>
    httr2::req_perform()

  sp_parse_token_response(resp)
}

sp_parse_token_response <- function(resp) {
  body <- httr2::resp_body_json(resp)
  list(
    access_token = body$access_token,
    expires_at = Sys.time() + as.numeric(body$expires_in)
  )
}

# Builds a signed JWT client assertion for certificate-based auth, per
# https://learn.microsoft.com/entra/identity-platform/certificate-credentials
sp_build_client_assertion <- function(auth) {
  key <- sp_read_private_key(auth$certificate_path, auth$certificate_password)
  cert <- openssl::read_cert(auth$certificate_path)

  der <- openssl::write_der(cert)
  x5t <- jose::base64url_encode(openssl::sha1(der))

  now <- as.numeric(Sys.time())
  claim <- jose::jwt_claim(
    iss = auth$client_id,
    sub = auth$client_id,
    aud = sp_token_url(auth),
    jti = openssl::base64_encode(openssl::rand_bytes(16)),
    nbf = now,
    exp = now + 300
  )

  jose::jwt_encode_sig(claim, key, header = list(alg = "RS256", typ = "JWT", x5t = x5t))
}

# `certificate_path` may be a PEM file containing only the private key, or a
# bundle with the certificate concatenated in as well (as produced by e.g.
# `cat key.pem cert.pem > combined.pem`). openssl::read_key() fails on the
# latter, so pull out just the private-key PEM block first.
sp_read_private_key <- function(path, password) {
  content <- readChar(path, file.info(path)$size, useBytes = TRUE)
  match <- regmatches(
    content,
    regexpr("(?s)-----BEGIN (?:RSA |EC )?PRIVATE KEY-----.*?-----END (?:RSA |EC )?PRIVATE KEY-----", content, perl = TRUE)
  )
  if (length(match) == 0) {
    return(openssl::read_key(path, password = password))
  }

  tmp <- tempfile(fileext = ".pem")
  on.exit(unlink(tmp))
  writeLines(match, tmp)
  openssl::read_key(tmp, password = password)
}

#' @return A valid request digest string for the SharePoint REST backend.
#' @noRd
sp_get_digest <- function(auth) {
  cached <- auth$cache$digest
  if (!is.null(cached) && cached$expires_at > Sys.time() + 30) {
    return(cached$value)
  }

  resp <- httr2::request(paste0(auth$site_url, "/_api/contextinfo")) |>
    httr2::req_headers(
      Accept = "application/json;odata=verbose",
      !!!sp_rest_auth_headers(auth)
    ) |>
    sp_with_error_handling() |>
    httr2::req_perform()

  info <- sp_rest_resp_body_json(resp, auth)$d$GetContextWebInformation
  parts <- strsplit(info$FormDigestValue, ",")[[1]]

  auth$cache$digest <- list(
    value = parts[1],
    expires_at = Sys.time() + as.numeric(info$FormDigestTimeoutSeconds)
  )
  auth$cache$digest$value
}

# Authorization header(s) for a SharePoint REST call: bearer token for
# secret/certificate auth, or the raw session cookies for cookie auth.
sp_rest_auth_headers <- function(auth) {
  if (auth$auth_method == "cookie") {
    list(Cookie = paste0("FedAuth=", auth$fedauth, "; rtFa=", auth$rtfa))
  } else {
    list(Authorization = paste0("Bearer ", sp_get_token(auth)))
  }
}

# Cookie sessions expire silently: SharePoint responds 200 with an HTML sign-in
# page rather than an HTTP error. Give a clear message instead of a JSON parse
# error in that case.
sp_rest_resp_body_json <- function(resp, auth) {
  tryCatch(
    httr2::resp_body_json(resp),
    error = function(e) {
      if (identical(auth$auth_method, "cookie")) {
        stop(
          "Failed to parse SharePoint's response as JSON. Your `fedauth`/`rtfa` ",
          "cookies have likely expired - sign in to SharePoint in your browser ",
          "again and copy fresh values.",
          call. = FALSE
        )
      }
      stop(e)
    }
  )
}

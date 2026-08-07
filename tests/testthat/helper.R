make_graph_auth <- function(library = "Shared Documents", drive_id = "drive1", site_id = "site1") {
  auth <- sharepoint_auth(
    tenant_id = "tenant.onmicrosoft.com",
    client_id = "client-id",
    site_url = "https://contoso.sharepoint.com/sites/MySite",
    client_secret = "secret",
    backend = "graph"
  )
  auth$cache$token <- list(access_token = "tok", expires_at = Sys.time() + 3600)
  auth$cache$site_id <- site_id
  auth$cache[[paste0("drive_", library)]] <- drive_id
  auth
}

make_rest_auth <- function() {
  auth <- sharepoint_auth(
    tenant_id = "tenant.onmicrosoft.com",
    client_id = "client-id",
    site_url = "https://contoso.sharepoint.com/sites/MySite",
    client_secret = "secret",
    backend = "rest"
  )
  auth$cache$token <- list(access_token = "tok", expires_at = Sys.time() + 3600)
  auth$cache$digest <- list(value = "digest-val", expires_at = Sys.time() + 1800)
  auth
}

make_cookie_auth <- function() {
  auth <- sharepoint_auth(
    site_url = "https://contoso.sharepoint.com/sites/MySite",
    fedauth = "fedauth-value",
    rtfa = "rtfa-value"
  )
  auth$cache$digest <- list(value = "digest-val", expires_at = Sys.time() + 1800)
  auth
}

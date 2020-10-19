#' Generate a SharePoint Token
#'
#' @param client_id The ID of the client you created using _layouts/15/appregnew.aspx
#' @param client_secret The secret for the client you created using _layouts/15/appregnew.aspx
#' @param tenant_id Bearer realm="tenant ID" when calling curl -X GET -v -H "Authorization: Bearer" https://url_of_sharepoint_site/_vti_bin/client.svc/
#' @param resource_id client_id="resource ID" when calling curl -X GET -v -H "Authorization: Bearer" https://url_of_sharepoint_site/_vti_bin/client.svc/
#' @param site_domain URL for SharePoint domain, e.g. kksky.sharepoint.com
#'
#' @importFrom httr content_type POST content
#'
#' @return A Token that can be used for calling the SharePoint API
#' @export
#'
#' @examples #no example yet
get_sharepoint_token <- function(client_id, client_secret, tenant_id, resource_id, site_domain){
  # Preparing call
  url <- paste0("https://accounts.accesscontrol.windows.net/", tenant_id, "/tokens/OAuth/2")
  headers <- httr::content_type("application/x-www-form-urlencoded")
  body <- paste0("grant_type=client_credentials", "&",
                 "client_id=", paste0(client_id, "@", tenant_id), "&",
                 "client_secret=", client_secret, "&",
                 "resource=", paste0(resource_id, "/", site_domain, "@", tenant_id))

  # Making call
  my_content <- httr::POST(url = url, headers, body = body)

  # Extracting token
  my_token <- httr::content(my_content)$access_token

  return(my_token)
}

#' Generate a Digest Value
#'
#' @param sharepoint_token A SharePoint token from get_sharepoint_token()
#' @param sharepoint_url A SharePoint url, e.g. kksky.sharepoint.com
#'
#' @importFrom utils URLencode
#'
#' @return
#' @export
#'
#' @examples #no example yet
get_sharepoint_digest_value <- function(sharepoint_token, sharepoint_url) {

  # Preparing call
  url <- utils::URLencode(paste0(sharepoint_url, "/_api/contextinfo"))
  headers <- httr::add_headers("Accept" = "application/json;odata=verbose",
                         "Authorization" = paste0("Bearer ", sharepoint_token))

  # Making call
  my_content <- httr::POST(url = url, headers)

  # Extracting digest value
  digest_value <- strsplit(httr::content(my_content)$d$GetContextWebInformation$FormDigestValue, ",")[[1]][1]

  return(digest_value)
}

#' Download a File from SharePoint
#'
#' @param sharepoint_token A SharePoint token from get_sharepoint_token()
#' @param sharepoint_url A SharePoint url, e.g. kksky.sharepoint.com
#' @param sharepoint_digest_value A SharePoint digest value from get_sharepoint_digest_value()
#' @param sharepoint_path Path to the file, e.g. Shared Documents/test
#' @param sharepoint_file_name Name of the file to download, e.g. Dokument.docx
#' @param out_path Local path to write file to, e.g. C:/Dokument.docx
#'
#' @importFrom httr add_headers GET content
#' @importFrom utils URLencode
#'
#' @return
#' @export
#'
#' @examples #no example yet
download_sharepoint_file <- function(sharepoint_token, sharepoint_url, sharepoint_digest_value, sharepoint_path, sharepoint_file_name, out_path) {

  # Preparing call
  url <- utils::URLencode(paste0(sharepoint_url, "/_api/web/GetFolderByServerRelativeUrl('", sharepoint_path, "')", "/Files('", sharepoint_file_name, "')/$value"))
  headers <- httr::add_headers("Accept" = "application/json;odata=verbose",
                         "Authorization" = paste0("Bearer ", sharepoint_token),
                         "X-RequestDigest" = sharepoint_digest_value)

  # Making call
  my_content <- httr::GET(url = url, headers)

  # Writing content to file
  writeBin(httr::content(my_content), paste0(out_path, "/", sharepoint_file_name))
}


#(ads)import httr would be fine too. Importing only function in use reduces namespace collisions.
#' Upload a File to SharePoint
#'
#' @param sharepoint_token A SharePoint token from get_sharepoint_token()
#' @param sharepoint_url A SharePoint url, e.g. kksky.sharepoint.com
#' @param sharepoint_digest_value A SharePoint digest value from get_sharepoint_digest_value()
#' @param sharepoint_path Path to the file, e.g. Shared Documents/test
#' @param sharepoint_file_name Name of the file in SharePoint, e.g. Dokument.docx
#' @param file_path Path to the file you want to upload, e.g. C:/Dokument.docx
#'
#' @importFrom httr upload_file POST add_headers
#' @importFrom utils URLencode
#' @return
#' @export
#'
#' @examples #no example yet
upload_file_to_sharepoint <- function(sharepoint_token, sharepoint_url, sharepoint_digest_value, sharepoint_path, sharepoint_file_name, file_path) {


  #Bonus info: if ever uploading to CRAN, they may as pedantic to require you to specify any dependency not base
  #utils, graphics, stats etc. are not base, but always provided in virtually any distribution and you don't have to load them
  # Prepare call
  url <- utils::URLencode(paste0(sharepoint_url, "/_api/web/GetFolderByServerRelativeUrl('", sharepoint_path, "')", "/Files/Add(url='", sharepoint_file_name, "',overwrite=true)"))
  headers <- httr::add_headers("Authorization" = paste0("Bearer ", sharepoint_token),
                         "X-RequestDigest" = sharepoint_digest_value)
  #it can be a good investment to literally mention any foreign package
  #it eases to read the source code for the first time
  #the risk of a namespace collisions, leading to weird bugs is reduced
  body <- httr::upload_file(file_path)

  # Making call
  my_content <- httr::POST(url = url, body = body, headers)
}

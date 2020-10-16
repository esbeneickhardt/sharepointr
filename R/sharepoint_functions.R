#' Generate a SharePoint Token
#'
#' @param client_id The ID of the client you created using _layouts/15/appregnew.aspx
#' @param client_secret The secret for the client you created using _layouts/15/appregnew.aspx
#' @param tenant_id Bearer realm="tenant ID" when calling curl -X GET -v -H "Authorization: Bearer" https://url_of_sharepoint_site/_vti_bin/client.svc/
#' @param resource_id client_id="resource ID" when calling curl -X GET -v -H "Authorization: Bearer" https://url_of_sharepoint_site/_vti_bin/client.svc/
#' @param site_domain URL for SharePoint domain, e.g. kksky.sharepoint.com
#'
#' @return A Token that can be used for calling the SharePoint API
#' @export
#'
#' @examples
get_sharepoint_token <- function(client_id, client_secret, tenant_id, resource_id, site_domain){
  # Preparing call
  url <- paste0("https://accounts.accesscontrol.windows.net/", tenant_id, "/tokens/OAuth/2")
  headers <- content_type("application/x-www-form-urlencoded")
  body <- paste0("grant_type=client_credentials", "&",
                 "client_id=", paste0(client_id, "@", tenant_id), "&",
                 "client_secret=", client_secret, "&",
                 "resource=", paste0(resource_id, "/", site_domain, "@", tenant_id))
  
  # Making call
  my_content <- POST(url = url, headers, body = body)
  
  # Extracting token
  my_token <- content(my_content)$access_token
  
  return(my_token)
}

#' Generate a Digest Value
#'
#' @param sharepoint_token A SharePoint token from get_sharepoint_token()
#' @param sharepoint_url A SharePoint url, e.g. kksky.sharepoint.com
#'
#' @return
#' @export
#'
#' @examples
get_sharepoint_digest_value <- function(sharepoint_token, sharepoint_url) {
  
  # Preparing call
  url <- URLencode(paste0(sharepoint_url, "/_api/contextinfo"))  
  headers <- add_headers("Accept" = "application/json;odata=verbose",
                         "Authorization" = paste0("Bearer ", sharepoint_token))  
  
  # Making call
  my_content <- POST(url = url, headers)
  
  # Extracting digest value
  digest_value <- strsplit(content(my_content)$d$GetContextWebInformation$FormDigestValue, ",")[[1]][1]
  
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
#' @return
#' @export
#'
#' @examples
download_sharepoint_file <- function(sharepoint_token, sharepoint_url, sharepoint_digest_value, sharepoint_path, sharepoint_file_name, out_path) {
  
  # Preparing call
  url <- URLencode(paste0(sharepoint_url, "/_api/web/GetFolderByServerRelativeUrl('", sharepoint_path, "')", "/Files('", sharepoint_file_name, "')/$value"))
  headers <- add_headers("Accept" = "application/json;odata=verbose",
                         "Authorization" = paste0("Bearer ", sharepoint_token),
                         "X-RequestDigest" = sharepoint_digest_value)
  
  # Making call
  my_content <- GET(url = url, headers)
  
  # Writing content to file
  writeBin(content(my_content), paste0(out_path, "/", sharepoint_file_name))
}

#' Upload a File to SharePoint
#'
#' @param sharepoint_token A SharePoint token from get_sharepoint_token()
#' @param sharepoint_url A SharePoint url, e.g. kksky.sharepoint.com
#' @param sharepoint_digest_value A SharePoint digest value from get_sharepoint_digest_value()
#' @param sharepoint_path Path to the file, e.g. Shared Documents/test
#' @param sharepoint_file_name Name of the file in SharePoint, e.g. Dokument.docx
#' @param file_path Path to the file you want to upload, e.g. C:/Dokument.docx
#'
#' @return
#' @export
#'
#' @examples
upload_file_to_sharepoint <- function(sharepoint_token, sharepoint_url, sharepoint_digest_value, sharepoint_path, sharepoint_file_name, file_path) {
  
  # Prepare call
  url <- URLencode(paste0(sharepoint_url, "/_api/web/GetFolderByServerRelativeUrl('", sharepoint_path, "')", "/Files/Add(url='", sharepoint_file_name, "',overwrite=true)"))
  headers <- add_headers("Authorization" = paste0("Bearer ", sharepoint_token),
                         "X-RequestDigest" = sharepoint_digest_value)
  body <- upload_file(file_path)
    
  # Making call
  my_content <- POST(url = url, body = body, headers)
}
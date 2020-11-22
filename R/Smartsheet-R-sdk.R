pkg.globals <- new.env()

pkg.globals$api_key <- "NONE"
pkg.globals$working_folder_id <- "NONE"






#' Set the api key for acessing Smartsheet
#'
#' @param key a string for the API key you want to use
#'
#' @return if this function works it returns nothing
#' @export
#'
#' @examples
#' \dontrun{
#' set_smartsheet_api_key("yourAPIkey")
#' }
set_smartsheet_api_key <- function(key) {
  r <- httr::GET("https://api.smartsheet.com/2.0/folders/411795358803844/folders",
                 httr::add_headers('Authorization' = paste('Bearer',key, sep = ' ')))
  if(grepl("errorCode",httr::content(r, "text"))){
    stop("Smartsheet Error: Your API key was invalid.")
  }
  pkg.globals$api_key <- key
}






#' Set a working folder in Smartsheet with a folder ID
#'
#' @param folder_id a string or numeric that exactly matches the target folder ID
#'
#' @return if this function works it returns nothing
#' @export
#'
#' @examples
#' \dontrun{
#' set_smart_working_folder("123456789")
#' set_smart_working_folder(123456789)
#' }
set_smart_working_folder <- function(folder_id) {
  if(pkg.globals$api_key == "NONE"){
    stop("rsmartsheet Error: Please set your api key with set_smartsheet_api_key() to use this function.")
  }
  r <- httr::GET(paste("https://api.smartsheet.com/2.0/folders/",folder_id,"/folders",sep=""),
                 httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
  if(grepl("errorCode",httr::content(r, "text"))){
    stop("Smartsheet Error: Your folder_id was invalid.")
  }
  pkg.globals$working_folder_id <- paste(folder_id,"/folders/",sep="")
}




#' Get Smartsheet Folder ID
#'
#' @param folder_name a string that exactly matches the target folder name
#'
#' @return the target folders ID as a string
#' @export
#'
#' @examples
#' \dontrun{
#' get_smart_folder_id("Folder_Name")
#' }
get_smart_folder_id<-function(folder_name){
  if(pkg.globals$api_key == "NONE"){
    stop("rsmartsheet Error: Please set your api key with set_smartsheet_api_key() to use this function.")
  }
  if(pkg.globals$working_folder_id == "NONE"){
    r <- httr::GET("https://api.smartsheet.com/2.0/home/folders/",
                   httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
  } else{
    r <- httr::GET(paste("https://api.smartsheet.com/2.0/folders/",pkg.globals$working_folder_id,sep=""),
                   httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
  }
  folders_listed <- jsonlite::fromJSON(httr::content(r, "text"))
  if(sum(folders_listed$data['name'] == folder_name)>1){
    stop("rsmartsheet Error: More than 1 folder found with that name")
  }
  if(sum(folders_listed$data['name'] == folder_name)==0){
    stop("rsmartsheet Error: No folder found with that name")
  }
  id <- toString(folders_listed$data$id[folders_listed$data$name == folder_name])
  return(id)
}






#' Create New Smartsheet in Folder with Specific ID
#'
#' @param file_path a path which locates the csv and provides the name
#' @param folder_id a string or numeric that exactly matches the target folder ID
#'
#' @return returns nothing
#' @export
#'
#' @examples
#' \dontrun{
#' csv_to_sheet_in_folder("a_folder/maybe_another_folder/sheet_name.csv", 123456789)
#' csv_to_sheet_in_folder("a_folder/maybe_another_folder/sheet_name.csv", "123456789")
#' }
csv_to_sheet_in_folder<-function(file_path, folder_id){
  if(pkg.globals$api_key == "NONE"){
    stop("rsmartsheet Error: Please set your api key with set_smartsheet_api_key() to use this function.")
  }
  return(httr::POST(url=paste("https://api.smartsheet.com/2.0/folders",folder_id,'sheets',paste('import?sheetName=',substr(basename(file_path),0,str_length(basename(file_path))-4),'&headerRowIndex=0&primaryColumnIndex=0',sep=''),sep='/'),
                    body=upload_file(file_path), httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'text/csv', 'Content-Disposition'='attachment')))
}






#' Delete an Existing Smartsheet in Folder with Specific ID
#'
#' @param sheet_name the target Smartsheets exact sheet name
#'
#' @return returns nothing
#' @export
#'
#' @examples
#' \dontrun{
#' delete_sheet_by_name("sheet_name")
#' }
delete_sheet_by_name<-function(sheet_name){
  if(pkg.globals$api_key == "NONE"){
    stop("rsmartsheet Error: Please set your api key with set_smartsheet_api_key() to use this function.")
  }
  r <- httr::GET("https://api.smartsheet.com/2.0/sheets?&includeAll=true",
                 httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
  sheets_listed <- jsonlite::fromJSON(httr::content(r, "text"))
  if(sum(sheets_listed$data['name'] == sheet_name)>1){
    stop("rsmartsheet Error: More than 1 sheet found with that name")
  }
  if(sum(sheets_listed$data['name'] == sheet_name)==0){
    stop("rsmartsheet Error: No sheet found with that name")
  }
  id <- toString(sheets_listed$data$id[sheets_listed$data$name == sheet_name])
  return(httr::DELETE(paste("https://api.smartsheet.com/2.0/sheets",id,sep='/'),
                      httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '))))
}






#' Get text CSV for an existing Smartsheet with a specific sheet name
#'
#' @param sheet_name the target Smartsheet's exact sheet name
#'
#' @return returns text of csv
#' @export
#'
#' @examples
#' \dontrun{
#' get_sheet_as_csv("sheet_name")
#' }
get_sheet_as_csv<-function(sheet_name){
  if(pkg.globals$api_key == "NONE"){
    stop("rsmartsheet Error: Please set your api key with set_smartsheet_api_key() to use this function.")
  }
  r <- httr::GET("https://api.smartsheet.com/2.0/sheets?&includeAll=true",
                 httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
  sheets_listed <- jsonlite::fromJSON(httr::content(r, "text"))
  if(sum(sheets_listed$data['name'] == sheet_name)>1){
    stop("rsmartsheet Error: More than 1 sheet found with that name")
  }
  if(sum(sheets_listed$data['name'] == sheet_name)==0){
    stop("rsmartsheet Error: No sheet found with that name")
  }
  id <- toString(sheets_listed$data$id[sheets_listed$data$name == sheet_name])
  return(httr::content(httr::GET(paste("https://api.smartsheet.com/2.0/sheets",id,sep='/'),
                                 httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Accept' = 'text/csv')), "text"))
}






#' Get text CSV for an existing attachment on a Smartsheet with a specific sheet name
#'
#' @param sheet_name the target Smartsheet's exact sheet name
#' @param attachments_name the target attachment exact file name
#'
#' @return returns text of csv
#' @export
#'
#' @examples
#' \dontrun{
#' get_sheet_as_csv("sheet_name", "atttachment.csv")
#' }
get_sheet_csv_attachment<-function(sheet_name, attachments_name){
  if(pkg.globals$api_key == "NONE"){
    stop("rsmartsheet Error: Please set your api key with set_smartsheet_api_key() to use this function.")
  }
  r <- httr::GET("https://api.smartsheet.com/2.0/sheets?&includeAll=true",
                 httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
  sheets_listed <- jsonlite::fromJSON(httr::content(r, "text"))
  if(sum(sheets_listed$data['name'] == sheet_name)>1){
    stop("rsmartsheet Error: More than 1 sheet found with that name")
  }
  if(sum(sheets_listed$data['name'] == sheet_name)==0){
    stop("rsmartsheet Error: No sheet found with that name")
  }
  id <- toString(sheets_listed$data$id[sheets_listed$data$name == sheet_name])
  attachments_listed <- jsonlite::fromJSON(httr::content(httr::GET(paste("https://api.smartsheet.com/2.0/sheets",id,'attachments',sep='/'),
                                                                   httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '))), "text"))
  if(sum(attachments_listed$data['name'] == attachments_name)>1){
    stop("rsmartsheet Error: More than 1 attachment found with that name")
  }
  if(sum(attachments_listed$data['name'] == attachments_name)==0){
    stop("rsmartsheet Error: No attahment found with that name")
  }
  aid <- toString(attachments_listed$data$id[attachments_listed$data$name == attachments_name])
  attached <- jsonlite::fromJSON(httr::content(httr::GET(paste("https://api.smartsheet.com/2.0/sheets",id,'attachments',aid,sep='/'),
                                                         httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '))), "text"))
  return(httr::content(httr::GET(attached$url),'text'))
}






#' Download a file of an existing attachment on a Smartsheet with a specific sheet name
#'
#' @param sheet_name the target Smartsheet's exact sheet name
#' @param file_path the download path
#' @param attachments_name if the downlad path file name is not the name of the target attached file
#'
#' @return returns nothing but downloads file to location
#' @export
#'
#' @examples
#' \dontrun{
#' download_sheet_attachment("sheet_name","a_folder/maybe_another_folder/sheet_name.csv")
#' download_sheet_attachment("sheet_name","a_folder/maybe_another_folder/sheet_name.csv", file_name="different_sheet_name")
#' }
download_sheet_attachment<-function(sheet_name, file_path, attachments_name=""){
  if(pkg.globals$api_key == "NONE"){
    stop("rsmartsheet Error: Please set your api key with set_smartsheet_api_key() to use this function.")
  }
  if(attachments_name==""){
    attachments_name <- basename(file_path)
  }
  r <- httr::GET("https://api.smartsheet.com/2.0/sheets?&includeAll=true",
                 httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
  sheets_listed <- jsonlite::fromJSON(httr::content(r, "text"))
  if(sum(sheets_listed$data['name'] == sheet_name)>1){
    stop("rsmartsheet Error: More than 1 sheet found with that name")
  }
  if(sum(sheets_listed$data['name'] == sheet_name)==0){
    stop("rsmartsheet Error: No sheet found with that name")
  }
  id <- toString(sheets_listed$data$id[sheets_listed$data$name == sheet_name])
  attachments_listed <- jsonlite::fromJSON(httr::content(GET(paste("https://api.smartsheet.com/2.0/sheets",id,'attachments',sep='/'),
                                                             httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '))), "text"))
  if(sum(attachments_listed$data['name'] == attachments_name)>1){
    stop("rsmartsheet Error: More than 1 attachment found with that name")
  }
  if(sum(attachments_listed$data['name'] == attachments_name)==0){
    stop("rsmartsheet Error: No attahment found with that name")
  }
  aid <- toString(attachments_listed$data$id[attachments_listed$data$name == attachments_name])
  attached <- jsonlite::fromJSON(httr::content(GET(paste("https://api.smartsheet.com/2.0/sheets",id,'attachments',aid,sep='/'),
                                                   httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '))), "text"))
  return(httr::GET(attached$url, write_disk(file_path, overwrite=TRUE)))
}






#' Download a file of an existing attachment on a Smartsheet with a specific sheet name
#'
#' @return returns table of smartsheets
#' @export
#'
#' @examples
#' \dontrun{
#' list_sheets()
#' }
list_sheets<-function(){
  if(pkg.globals$api_key == "NONE"){
    stop("rsmartsheet Error: Please set your api key with set_smartsheet_api_key() to use this function.")
  }
  r <- httr::GET("https://api.smartsheet.com/2.0/sheets?&includeAll=true",
                 httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
  sheets_listed <- jsonlite::fromJSON(httr::content(r, "text"))
  return(sheets_listed)
}






#' Replace an existing attachment file on a Smartsheet with a new file
#'
#' @param sheet_name the target Smartsheet's exact sheet name
#' @param file_path the path to the file that will overwrite the existing attachment
#' @param attachment_name the name of the attached file to replace
#'
#' @return returns nothing
#' @export
#'
#' @examples
#' \dontrun{
#' replace_sheet_attachment("sheet_name","a_folder/maybe_another_folder/attachment_name.csv", "other_attachment_name.csv")
#' }
replace_sheet_attachment<-function(sheet_name, file_path, attachment_name){
  if(pkg.globals$api_key == "NONE"){
    stop("rsmartsheet Error: Please set your api key with set_smartsheet_api_key() to use this function.")
  }
  if(basename(file_path)!=attachment_name){
    stop("rsmartsheet Error: upload file_path basename must equal attachment_name")
  }
  r <- httr::GET("https://api.smartsheet.com/2.0/sheets?&includeAll=true",
                 httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
  sheets_listed <- jsonlite::fromJSON(httr::content(r, "text"))
  if(sum(sheets_listed$data['name'] == sheet_name)>1){
    stop("rsmartsheet Error: More than 1 sheet found with that name")
  }
  if(sum(sheets_listed$data['name'] == sheet_name)==0){
    stop("rsmartsheet Error: No sheet found with that name")
  }
  id <- toString(sheets_listed$data$id[sheets_listed$data$name == sheet_name])
  attachments_listed <- jsonlite::fromJSON(httr::content(GET(paste("https://api.smartsheet.com/2.0/sheets",id,'attachments',sep='/'),
                                                             httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '))), "text"))
  if(sum(attachments_listed$data['name'] == attachment_name)>1){
    stop("rsmartsheet Error: More than 1 attachment found with that name")
  }
  if(sum(attachments_listed$data['name'] == attachment_name)==0){
    stop("rsmartsheet Error: No attahment found with that name")
  }
  aid <- toString(attachments_listed$data$id[attachments_listed$data$name == attachment_name])
  attached <- httr::DELETE(paste("https://api.smartsheet.com/2.0/sheets",id,'attachments',aid,sep='/'),
                           httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
  return(httr::POST(url=paste("https://api.smartsheet.com/2.0/sheets",id,'attachments',sep='/'),
                    body=upload_file(file_path), httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '),
                                                                   'Content-Type' = 'text/csv', 'Content-Disposition'=paste('attachment; filename="',attachment_name,'"',sep=''))))
}



#' Replace an existing Smartsheet with a new file
#'
#' @description Replaces the data in all existing rows and then adds new rows as appropriate.
#' Throws errors if the same columns are not present in both the the Sheet and csv to repalce it.
#'
#' @param sheet_name the target Smartsheet's exact sheet name
#' @param file_path the path to the file that will overwrite the existing attachment
#' @param never_delete prevents deleting extra rows if the replacement csv has fewer rows than target sheet
#'
#' @return returns nothing
#' @export
#'
#' @import magrittr
#'
#' @examples
#' \dontrun{
#' replace_sheet_with_csv("sheet_name","a_folder/maybe_another_folder/sheet.csv")
#' }
replace_sheet_with_csv<-function(sheet_name, file_path, never_delete=FALSE){
  if(pkg.globals$api_key == "NONE"){
    stop("rsmartsheet Error: Please set your api key with set_smartsheet_api_key() to use this function.")
  }
  # Find sheet by name
  r <- httr::GET("https://api.smartsheet.com/2.0/sheets?&includeAll=true",
                 httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
  sheets_listed <- jsonlite::fromJSON(httr::content(r, "text"))
  if(sum(sheets_listed$data['name'] == sheet_name)>1){
    stop("rsmartsheet Error: More than 1 sheet found with that name")
  }
  if(sum(sheets_listed$data['name'] == sheet_name)==0){
    stop("rsmartsheet Error: No sheet found with that name")
  }
  id <- toString(sheets_listed$data$id[sheets_listed$data$name == sheet_name])

  # Load in new data from csv to send
  data_to_send <- try(suppressWarnings(suppressMessages(readr::read_csv(file_path, col_types = readr::cols(.default = readr::col_character())))))
  if(inherits(data_to_send,"try-error")){
    stop("Could not read csv with readr")
  }

  # Download existing sheet data
  sheet_data <- jsonlite::fromJSON(str_replace_all(str_replace_all(httr::content(httr::GET(paste0("https://api.smartsheet.com/2.0/sheets/",id,"?level=2&include=objectValue"),
                                                                           httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '))), "text"),
                                                   '"id":([0-9]+)','"id":"\\1"')'"columnId":([0-9]+)','"columnId":"\\1"'), bigint_as_char=TRUE)

  # make column dictionary
  column_dict <- list()
  for (i in seq(1:nrow(sheet_data$columns))){
    column_dict[sheet_data$columns$title[i]] <- sheet_data$columns$id[i]
  }

  # in case their are no rows we want an empty vector for exisiting_rows not null
  exisiting_rows <- sheet_data$rows$id
  if(is.null(sheet_data$rows$id)){
    exisiting_rows <- vector()
  }

  # Check if columns match
  if(length(setdiff(colnames(data_to_send),sheet_data$columns$title))>0 |
     length(setdiff(sheet_data$columns$title,colnames(data_to_send)))>0){
    stop("Columns are not exactly the same between csv and target sheet")
  }

  # print mode
  if(length(exisiting_rows) == nrow(data_to_send)){
    print("replace_sheet_with_csv: perfect replacement")
    data_to_send <- suppressWarnings(suppressMessages(data_to_send %>%
      dplyr::mutate(id =dplyr::row_number()) %>%
      tidyr::pivot_longer(!id, names_to = "columnId",values_to="value") %>%
      dplyr::mutate(value=tidyr::replace_na(value,"")) %>%
      dplyr::mutate(columnId=as.character(unlist(unname(column_dict[columnId])))) %>%
      dplyr::group_split(id, keep=FALSE) %>%
      purrr::map_df(tidyr::nest) %>%
      dplyr::rename(cells=data) %>%
      dplyr::mutate(id=as.character(exisiting_rows))))
    return(httr::PUT(url=paste("https://api.smartsheet.com/2.0/sheets",id,'rows',sep='/'), body=jsonlite::toJSON(data_to_send),
                     httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'application/json')))
  }else if(length(exisiting_rows) < nrow(data_to_send)){
    print("replace_sheet_with_csv: adding new rows")
    data_to_send <- suppressWarnings(suppressMessages(data_to_send %>%
      dplyr::mutate(id =dplyr::row_number()) %>%
      tidyr::pivot_longer(!id, names_to = "columnId",values_to="value") %>%
      dplyr::mutate(value=tidyr::replace_na(value,"")) %>%
      dplyr::mutate(columnId=as.character(unlist(unname(column_dict[columnId])))) %>%
      dplyr::group_split(id, keep=FALSE) %>%
      purrr::map_df(tidyr::nest) %>%
      dplyr::rename(cells=data)))
    if(length(exisiting_rows)>0){
      data_to_update <- data_to_send %>%
        dplyr::slice(1:length(exisiting_rows)) %>%
        dplyr::mutate(id=as.character(exisiting_rows))
      r1 <- httr::PUT(url=paste("https://api.smartsheet.com/2.0/sheets",id,'rows',sep='/'), body=jsonlite::toJSON(data_to_update),
                      httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'application/json'))
      if(grepl("errorCode",httr::content(r1, "text"))){
        print(jsonlite::fromJSON(httr::content(r1, "text")))
        stop("Smartsheet Error: update row phase failed so skipping add row phase")
      }
    }
    data_to_add <- data_to_send %>%
      dplyr::slice(length(exisiting_rows):nrow(data_to_send)) %>%
      dplyr::mutate(toBottom=TRUE)
    r2 <- httr::POST(url=paste("https://api.smartsheet.com/2.0/sheets",id,'rows',sep='/'), body=jsonlite::toJSON(data_to_add),
                     httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'application/json'))
    return(list(response_one=r1,response_two=r2))
  } else if(length(exisiting_rows) > nrow(data_to_send) & never_delete==FALSE){
    print("replace_sheet_with_csv: deleteing some extra rows")
    data_to_send <- suppressWarnings(suppressMessages(data_to_send %>%
      dplyr::mutate(id =dplyr::row_number()) %>%
      tidyr::pivot_longer(!id, names_to = "columnId",values_to="value") %>%
      dplyr::mutate(value=tidyr::replace_na(value,"")) %>%
      dplyr::mutate(columnId=as.character(unlist(unname(column_dict[columnId])))) %>%
      dplyr::group_split(id, keep=FALSE) %>%
      purrr::map_df(tidyr::nest) %>%
      dplyr::rename(cells=data) %>%
      dplyr::mutate(id=as.character(head(exisiting_rows,length(cells))))))
    r1 <- httr::PUT(url=paste("https://api.smartsheet.com/2.0/sheets",id,'rows',sep='/'), body=jsonlite::toJSON(data_to_send),
                    httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'application/json'))
    if(grepl("errorCode",httr::content(r1, "text"))){
      print(jsonlite::fromJSON(httr::content(r1, "text")))
      stop("Smartsheet Error: update row phase failed so skipping delete row phase")
    }
    r2 <- httr::DELETE(url=paste0("https://api.smartsheet.com/2.0/sheets/",id,'/rows?ignoreRowsNotFound=true&ids=',
                                  paste(tail(exisiting_rows,length(exisiting_rows)-nrow(data_to_send)),collapse=",")),
                       httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
    return(list(response_one=r1,response_two=r2))
  } else if(length(exisiting_rows) > nrow(data_to_send) & never_delete){
    stop("ERROR replace_sheet_with_csv: csv has fewer rows than sheet requiring some rows to be deleted,\n set never_delete to FALSE and try again.")
  }
}

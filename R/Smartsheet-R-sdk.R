pkg.globals <- new.env()

pkg.globals$api_key <- "NONE"
pkg.globals$working_folder_id <- "NONE"
pkg.globals$allow_ids <- FALSE






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
  r <- httr::GET("https://api.smartsheet.com/2.0/sheets?&includeAll=false",
                 httr::add_headers('Authorization' = paste('Bearer',key, sep = ' ')))
  if(grepl("errorCode",httr::content(r, "text"))){
    stop("rmartsheet Error: Your API key was invalid.")
  }
  pkg.globals$api_key <- key
}



#' Check Smartsheet API Key
#'
#' @return if this function works it returns nothing
#' @export
#'
#' @examples
#' \dontrun{
#' check_smartsheet_api_key()
#' }
check_smartsheet_api_key <- function() {
  if(pkg.globals$api_key == "NONE"){
    stop("rsmartsheet Error: Please set your api key with set_smartsheet_api_key() to use this function.")
  }
}




#' Set whether you can use direct sheet id
#'
#' @description Set whether you can use direct sheet id references instead of sheet names
#'
#' @param inp logical TRUE or FALSE
#'
#' @return returns nothing
#' @export
#'
#' @examples
#' \dontrun{
#' use_direct_ids()
#' }
use_direct_ids <- function(inp) {
  pkg.globals$allow_ids <- as.logical(inp)
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

  check_smartsheet_api_key()

  r <- httr::GET(paste("https://api.smartsheet.com/2.0/folders/",folder_id,"/folders",sep=""),
                 httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
  if(grepl("errorCode",httr::content(r, "text"))){
    stop("rmartsheet Error: Your folder_id was invalid.")
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

  check_smartsheet_api_key()

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
csv_to_sheet_in_folder<-function(file_path, folder_id, all_text_number=FALSE){

  check_smartsheet_api_key()

  if(!stringr::str_detect(basename(file_path),"\\.csv$")){
    stop("rmartsheet Error: target file is not a .csv file")
  }

  sheet_name <- stringr::str_remove(basename(file_path),"\\.csv$")

  r <- httr::POST(url=paste("https://api.smartsheet.com/2.0/folders",folder_id,'sheets',paste('import?sheetName=',sheet_name,'&headerRowIndex=0&primaryColumnIndex=0',sep=''),sep='/'),
                    body=httr::upload_file(file_path), httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'text/csv', 'Content-Disposition'='attachment'))
  if(all_text_number){
    if(grepl("errorCode",httr::content(r, "text"))){
     print(jsonlite::fromJSON(httr::content(r, "text")))
     stop("rmartsheet Error: could not make sheet in the first place so aborting column type asssignment to TEXT_NUMBER")
    }
    set_sheet_columns_to_textnumber(sheet_name)
  }
  return(r)
}




#' Create New Smartsheet
#'
#' @param file_path a path which locates the csv and provides the name
#'
#' @return returns nothing
#' @export
#'
#' @examples
#' \dontrun{
#' csv_to_sheet("a_folder/maybe_another_folder/sheet_name.csv", 123456789)
#' csv_to_sheet("a_folder/maybe_another_folder/sheet_name.csv", "123456789")
#' }
csv_to_sheet<-function(file_path, all_text_number=FALSE){

  check_smartsheet_api_key()

  if(!stringr::str_detect(basename(file_path),"\\.csv$")){
    stop("rmartsheet Error: target file is not a .csv file")
  }

  sheet_name <- stringr::str_remove(basename(file_path),"\\.csv$")
  if(pkg.globals$working_folder_id == "NONE"){
    r <- httr::POST(url=paste("https://api.smartsheet.com/2.0/sheets",paste('import?sheetName=',sheet_name,'&headerRowIndex=0&primaryColumnIndex=0',sep=''),sep='/'),
                      body=httr::upload_file(file_path), httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'text/csv', 'Content-Disposition'='attachment'))
  } else{
    r <- httr::POST(url=paste("https://api.smartsheet.com/2.0/folders",pkg.globals$working_folder_id,'sheets',paste('import?sheetName=',sheet_name,'&headerRowIndex=0&primaryColumnIndex=0',sep=''),sep='/'),
                      body=httr::upload_file(file_path), httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'text/csv', 'Content-Disposition'='attachment'))
  }
  if(all_text_number){
    if(grepl("errorCode",httr::content(r, "text"))){
      print(jsonlite::fromJSON(httr::content(r, "text")))
      stop("rmartsheet Error: could not make sheet in the first place so aborting column type asssignment to TEXT_NUMBER")
    }
    set_sheet_columns_to_textnumber(sheet_name)
  }
  return(r)
}




#' Get Sheet Column Info
#'
#' @param sheet_name name of smartsheet
#'
#' @return returns table of column info
#' @export
#'
#' @examples
#' \dontrun{
#' get_sheet_column_info("sheet_name")
#' }
get_sheet_column_info<-function(sheet_name){

  check_smartsheet_api_key()

  id <- sheet_name_to_id(sheet_name)

  r <- httr::GET(paste0("https://api.smartsheet.com/2.0/sheets/",id,"/columns"),
                      httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))

  return(jsonlite::fromJSON(httr::content(r, "text"))$data)
}




#' Set Sheet Columns to Text Number Types
#'
#' @param sheet_name name of smartsheet
#'
#' @return returns nothing
#' @export
#'
#' @import magrittr
#'
#' @examples
#' \dontrun{
#' set_sheet_columns_to_textnumber("sheet_name")
#' }
set_sheet_columns_to_textnumber<-function(sheet_name){

  check_smartsheet_api_key()

  id <- sheet_name_to_id(sheet_name)

  col_data <- get_sheet_column_info(sheet_name)

  target_columns <- col_data %>% dplyr::filter(type!="TEXT_NUMBER") %>% dplyr::pull(id) %>% as.character()

  for (col in target_columns){
    r <- httr::PUT(paste0("https://api.smartsheet.com/2.0/sheets/",id,"/columns/",col),
                   httr::add_headers('Authorization' = paste('Bearer',smartkey, sep = ' ')), body='{"type": "TEXT_NUMBER"}')
    if(grepl("errorCode",httr::content(r, "text"))){
      print(paste("Column type update had error on column name:",target_columns[target_columns$id==col]$title))
      print(jsonlite::fromJSON(httr::content(r, "text")))
      stop("rmartsheet Error: column type update terminated early")
    }
  }
}




#' Get Smartsheet Name From ID
#'
#' @param sheet_name the name of a smartsheet
#'s
#' @return returns sheet ID
#' @export
#'
#' @examples
#' \dontrun{
#' sheet_name_to_id("sheet_name")
#' }
sheet_name_to_id<-function(sheet_name){

  if(!stringr::str_detect(sheet_name,"[A-Z]|[a-z]")){
    if(pkg.globals$allow_ids){
      return(sheet_name)
    } else{
      print("rsmartsheet Warning: It looks like you passed a sheet ID not a sheet name. Call use_direct_ids(TRUE) at the begining of your session to allow direct reference of sheets by ID")
      warning("rsmartsheet Warning: It looks like you passed a sheet ID not a sheet name.\nCall use_direct_ids(TRUE) at the begining of your session to allow direct reference of sheets by ID")
    }
  }

  check_smartsheet_api_key()

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
  return(id)
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

  check_smartsheet_api_key()

  id <- sheet_name_to_id(sheet_name)

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

  check_smartsheet_api_key()

  id <- sheet_name_to_id(sheet_name)
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

  check_smartsheet_api_key()

  id <- sheet_name_to_id(sheet_name)
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

  check_smartsheet_api_key()

  if(attachments_name==""){
    attachments_name <- basename(file_path)
  }
  id <- sheet_name_to_id(sheet_name)
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






#' Get a list of sheets that the user is able to access
#'
#' @return returns table of smartsheets
#' @export
#'
#' @examples
#' \dontrun{
#' list_sheets()
#' }
list_sheets<-function(){

  check_smartsheet_api_key()

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

  check_smartsheet_api_key()

  if(basename(file_path)!=attachment_name){
    stop("rsmartsheet Error: upload file_path basename must equal attachment_name")
  }
  id <- sheet_name_to_id(sheet_name)
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
                    body=httr::upload_file(file_path), httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '),
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
replace_sheet_with_csv<-function(sheet_name, file_path, never_delete=FALSE, batch_size=5000){

  check_smartsheet_api_key()

    # Find sheet by name
  id <- sheet_name_to_id(sheet_name)

  # Load in new data from csv to send
  data_to_send <- try(suppressWarnings(suppressMessages(readr::read_csv(file_path, col_types = readr::cols(.default = readr::col_character())))))
  if(inherits(data_to_send,"try-error")){
    stop("Could not read csv with readr")
  }

  # Download existing sheet data
  sheet_data <- jsonlite::fromJSON(stringr::str_replace_all(stringr::str_replace_all(httr::content(httr::GET(paste0("https://api.smartsheet.com/2.0/sheets/",id,"?level=2&include=objectValue"),
                                                                           httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '))), "text"),
                                                   '"id":([0-9]+)','"id":"\\1"'),'"columnId":([0-9]+)','"columnId":"\\1"'), bigint_as_char=TRUE)

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
    stop(paste("Columns are not exactly the same between csv and target sheet:",
               paste(setdiff(colnames(data_to_send),sheet_data$columns$title),collpase=", "),
               paste(setdiff(sheet_data$columns$title,colnames(data_to_send)),collpase=", ")))
  }

  make_indexes <- function(size){
    test <- seq(1:size)
    max <- batch_size
    y <- seq_along(test)
    chunks <- split(test, ceiling(y/max))
    return(list(low=lapply(chunks,min),high=lapply(chunks,max),size=length(chunks)))
  }

  # print mode
  if(length(exisiting_rows) == nrow(data_to_send)){
    if(nrow(data_to_send)==0){
      print("replace_sheet_with_csv: perfect replacement 0 rows in both")
      return()
    }
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
    responses <- list()
    indexes <- make_indexes(nrow(data_to_send))
    pb <- txtProgressBar(0, indexes$size, style = 3)
    for(i in seq(1:indexes$size)){
      r <- httr::PUT(url=paste("https://api.smartsheet.com/2.0/sheets",id,'rows',sep='/'), body=jsonlite::toJSON(data_to_send %>% dplyr::slice(indexes$low[[i]]:indexes$high[[i]])),
                     httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'application/json'))
      if(grepl("errorCode",httr::content(r, "text"))){
        print(paste("In chunk:",i))
        print(jsonlite::fromJSON(httr::content(r, "text")))
        stop("rmartsheet Error: add row phase failed")
      }
      responses[paste0("r",i)] <- list(r)
      setTxtProgressBar(pb, i)
    }
    return(responses)
  } else if(length(exisiting_rows) < nrow(data_to_send)){
    print("replace_sheet_with_csv: adding some new rows")
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
      responses_update <- list()
      indexes <- make_indexes(nrow(data_to_update))
      pb <- txtProgressBar(0, indexes$size, style = 3)
      for(i in seq(1:indexes$size)){
        r <- httr::PUT(url=paste("https://api.smartsheet.com/2.0/sheets",id,'rows',sep='/'), body=jsonlite::toJSON(data_to_update %>% dplyr::slice(indexes$low[[i]]:indexes$high[[i]])),
                       httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'application/json'))
        if(grepl("errorCode",httr::content(r, "text"))){
          print(paste("In chunk:",i))
          print(jsonlite::fromJSON(httr::content(r, "text")))
          stop("rmartsheet Error: update row phase failed so skipping add row phase")
        }
        responses_update[paste0("r",i)] <- list(r)
        setTxtProgressBar(pb, i)
      }
    }
    data_to_add <- data_to_send %>%
      dplyr::slice(length(exisiting_rows):nrow(data_to_send)) %>%
      dplyr::mutate(toBottom=TRUE)
     responses_add <- list()
    indexes <- make_indexes(nrow(data_to_add))
    pb <- txtProgressBar(0, indexes$size, style = 3)
    for(i in seq(1:indexes$size)){
      r <- httr::POST(url=paste("https://api.smartsheet.com/2.0/sheets",id,'rows',sep='/'), body=jsonlite::toJSON(data_to_add %>% dplyr::slice(indexes$low[[i]]:indexes$high[[i]])),
                      httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'application/json'))
      if(grepl("errorCode",httr::content(r, "text"))){
        print(paste("In chunk:",i))
        print(jsonlite::fromJSON(httr::content(r, "text")))
        stop("rmartsheet Error: add row phase failed")
      }
      responses_add[paste0("r",i)] <- list(r)
      setTxtProgressBar(pb, i)
    }
    if(length(exisiting_rows)>0){
      return(list(response_update=responses_update,response_add=responses_add))
    } else{
      return(list(response_add=responses_add))
    }
  } else if(length(exisiting_rows) > nrow(data_to_send) & never_delete==FALSE){
    if (nrow(data_to_send)>0){
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
      responses_add <- list()
      indexes <- make_indexes(nrow(data_to_send))
      pb <- txtProgressBar(0, indexes$size, style = 3)
      for(i in seq(1:indexes$size)){
        r <- httr::PUT(url=paste("https://api.smartsheet.com/2.0/sheets",id,'rows',sep='/'), body=jsonlite::toJSON(data_to_send %>% dplyr::slice(indexes$low[[i]]:indexes$high[[i]])),
                       httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'application/json'))
        if(grepl("errorCode",httr::content(r, "text"))){
          print(paste("In chunk:",i))
          print(jsonlite::fromJSON(httr::content(r, "text")))
          stop("rmartsheet Error: update row phase failed so skipping delete row phase")
        }
        responses_add[paste0("r",i)] <- list(r)
        setTxtProgressBar(pb, i)
      }
    } else{
      print("replace_sheet_with_csv: deleteing all extra rows")
      responses_add <- list()
    }
    delete_rows <- tail(exisiting_rows,length(exisiting_rows)-nrow(data_to_send))
    indexes <- make_indexes(length(delete_rows))
    responses_delete <- list()
    pb <- txtProgressBar(0, indexes$size, style = 3)
    for(i in seq(1:indexes$size)){
      r <- httr::DELETE(url=paste0("https://api.smartsheet.com/2.0/sheets/",id,'/rows?ignoreRowsNotFound=true&ids=', paste(delete_rows[c(indexes$low[[i]]:indexes$high[[i]])],collapse=",")),
                        httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' ')))
      if(grepl("errorCode",httr::content(r, "text"))){
        print(paste("In chunk:",i))
        print(jsonlite::fromJSON(httr::content(r, "text")))
        stop("rmartsheet Error: delete row phase failed")
      }
      responses_delete[paste0("r",i)] <- list(r)
      setTxtProgressBar(pb, i)
    }
    return(list(response_add=responses_add,response_delete=responses_delete))
  } else if(length(exisiting_rows) > nrow(data_to_send) & never_delete){
    stop("ERROR replace_sheet_with_csv: csv has fewer rows than sheet requiring some rows to be deleted,\n set never_delete to FALSE and try again.")
  }
}




#' Colorize Smartsheet
#'
#' @description For a sheet with a column HEX_COLOR that sheet will be updated
#' (via a full replacement) so that those row have those colors and then the HEX_COLOR
#' column will be removed
#'
#' @param sheet_name the target Smartsheet's exact sheet name
#' @param clean_hex_col defaults to TRUE and removes HEX_COLOR column when complete
#'
#' @return returns nothing
#' @export
#'
#' @import magrittr
#'
#' @examples
#' \dontrun{
#' colorize_sheet("sheet_name")
#' }
colorize_sheet<-function(sheet_name, clean_hex_col=TRUE, batch_size=5000){

  check_smartsheet_api_key()

  existing <- readr::read_csv(get_sheet_as_csv(sheet_name))
  if(!"HEX_COLOR" %in% colnames(existing)){
    stop("rsmartsheet Error: can't colorize without column 'HEX_COLOR'")
  }
  if(length(unique(stringr::str_length(existing$HEX_COLOR[is.na(existing$HEX_COLOR)==FALSE])))>1 &
     unique(stringr::str_length(existing$HEX_COLOR[is.na(existing$HEX_COLOR)==FALSE]))[1] == 9){
    stop("rsmartsheet Error: can't colorize with 'HEX_COLOR' column that has lengths other than 9 (starting with #) or NA")
  }
  palette <- c("#000000", "#FFFFFF", NA, "#FFEBEE", "#FFF3DF", "#FFFEE6", "#E7F5E9", "#E2F2FE", "#F4E4F5", "#F2E8DE", "#FFCCD2",
               "#FFE1AF", "#FEFF85", "#C6E7C8", "#B9DDFC", "#EBC7EF", "#EEDCCA", "#E5E5E5", "#F87E7D", "#FFCD7A", "#FEFF00", "#7ED085", "#5FB3F9",
               "#D190DA", "#D0AF8F", "#BDBDBD", "#EA352E", "#FF8D00", "#FFED00", "#40B14B", "#1061C3", "#9210AD", "#974C00", "#757575", "#991310",
               "#EA5000", "#EBC700", "#237F2E", "#0B347D", "#61058B", "#592C00")

  incolors <- unique(existing$HEX_COLOR[is.na(existing$HEX_COLOR)==FALSE])

  best_match <- list()
  dist_match <- list()
  for (color in incolors){
    best <- 10000000
    index <- 3
    for(pal in palette){
      target <- grDevices::col2rgb(color)
      test <- grDevices::col2rgb(pal)
      dist <- sqrt(((target[1]-test[1])^2)+((target[2]-test[2])^2)+((target[3]-test[3])^2))
      if(dist < best){
        index <- match(pal,palette)
        best <- dist
      }
    }
    best_match[color] <- index
    dist_match[color] <- best
  }

  id <- sheet_name_to_id(sheet_name)

  # Load in new data from csv to send
  data_to_send <- existing
  # Download existing sheet data
  sheet_data <- jsonlite::fromJSON(stringr::str_replace_all(stringr::str_replace_all(httr::content(httr::GET(paste0("https://api.smartsheet.com/2.0/sheets/",id,"?level=2&include=objectValue"),
                                                                                                             httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '))), "text"),
                                                                                     '"id":([0-9]+)','"id":"\\1"'),'"columnId":([0-9]+)','"columnId":"\\1"'), bigint_as_char=TRUE)

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
    stop(paste("Columns are not exactly the same between csv and target sheet:",
               paste(setdiff(colnames(data_to_send),sheet_data$columns$title),collpase=", "),
               paste(setdiff(sheet_data$columns$title,colnames(data_to_send)),collpase=", ")))
  }

  make_indexes <- function(size){
    test <- seq(1:size)
    max <- batch_size
    y <- seq_along(test)
    chunks <- split(test, ceiling(y/max))
    return(list(low=lapply(chunks,min),high=lapply(chunks,max),size=length(chunks)))
  }

  # print mode
  if(length(exisiting_rows) == nrow(data_to_send)){
    if(nrow(data_to_send)==0){
      print("replace_sheet_with_csv: perfect replacement 0 rows in both")
      return()
    }
    print("replace_sheet_with_csv: perfect replacement")
    data_to_send <- suppressWarnings(suppressMessages(data_to_send %>%
                                                        dplyr::mutate(id =dplyr::row_number()) %>%
                                                        tidyr::pivot_longer(-all_of(c("id","HEX_COLOR")), names_to = "columnId",values_to="value") %>%
                                                        dplyr::mutate(value=tidyr::replace_na(value,"")) %>%
                                                        dplyr::mutate(columnId=as.character(unlist(unname(column_dict[columnId])))) %>%
                                                        dplyr::mutate(format=paste0(",,,,,,,,,",
                                                                                    ifelse(is.na(HEX_COLOR),"0",as.character(best_match[HEX_COLOR]))
                                                                                    ,",,,,,,,")) %>%
                                                        dplyr::select(-HEX_COLOR) %>%
                                                        dplyr::group_split(id, keep=FALSE) %>%
                                                        purrr::map_df(tidyr::nest) %>%
                                                        dplyr::rename(cells=data) %>%
                                                        dplyr::mutate(id=as.character(exisiting_rows))))
    responses <- list()
    indexes <- make_indexes(nrow(data_to_send))
    pb <- txtProgressBar(0, indexes$size, style = 3)
    for(i in seq(1:indexes$size)){
      r <- httr::PUT(url=paste("https://api.smartsheet.com/2.0/sheets",id,'rows',sep='/'), body=jsonlite::toJSON(data_to_send %>% dplyr::slice(indexes$low[[i]]:indexes$high[[i]])),
                     httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'application/json'))
      if(grepl("errorCode",httr::content(r, "text"))){
        print(paste("In chunk:",i))
        print(jsonlite::fromJSON(httr::content(r, "text")))
        stop("rmartsheet Error: add row phase failed")
      }
      responses[paste0("r",i)] <- list(r)
      setTxtProgressBar(pb, i)
    }
    if(clean_hex_col){
      r <- httr::DELETE(url=paste("https://api.smartsheet.com/2.0/sheets",id,'columns',column_dict$HEX_COLOR,sep='/'),
                        httr::add_headers('Authorization' = paste('Bearer',pkg.globals$api_key, sep = ' '), 'Content-Type' = 'application/json'))
    }
    return(responses)
  } else {
    stop("ERROR replace_sheet_with_csv: was changed during the operation - aborting!")
  }
}

library(tidyverse)
library(readxl)


########## Customer Service ########################

build_worksheets_cs <- function(curr_df) {
  
  table_title_row_cs <- which(curr_df == "Customer Stuff" | curr_df == "Customer Things", arr.ind=TRUE)
  table_end_extant <- table_title_row_cs[,"col"] + 14
  table_selection_cs <- curr_df[table_title_row_cs[,"row"] + 1:28, table_title_row_cs[,"col"]:table_end_extant]
  table_end_row_cs <- which(table_selection_cs == "TOTAL", arr.ind=TRUE)[1,"row"] - 1
  table_end_col_cs <- which(is.na(table_selection_cs[1,]), arr.ind=TRUE)[1,"col"] - 1
  col_range <- (as.numeric(table_title_row_cs[,"col"])+as.numeric(table_end_col_cs)) - 1
  table_trimmed_cs = curr_df[table_title_row_cs[,"row"] + 1:table_end_row_cs, as.numeric(table_title_row_cs[,"col"]):col_range]
  second_row_count_na <- rowSums(is.na(table_trimmed_cs[2,])) / length(table_trimmed_cs[2,])
  second_row_numeric_bool <- sum(!is.na(as.numeric(table_trimmed_cs[2,]))) / length(table_trimmed_cs[2,]) 
  if (second_row_count_na == 1) {
    title_row <- table_trimmed_cs[1,1:table_end_col_cs]
    table_trimmed_cs <- table_trimmed_cs[-c(1,2),]
  } else if (second_row_numeric_bool > 0){ 
    title_row <- table_trimmed_cs[1,1:table_end_col_cs]
    table_trimmed_cs <- table_trimmed_cs[-c(1),]
  } else {
    table_trimmed_cs[is.na(table_trimmed_cs)] <- ""
    title_row <- trimws(paste(table_trimmed_cs[1,1:table_end_col_cs], table_trimmed_cs[2,1:table_end_col_cs]))
    table_trimmed_cs <- table_trimmed_cs[-c(1,2),]
  }
  colnames(table_trimmed_cs) <- title_row
  table_trimmed_cs$Date <- as.Date(as.numeric(table_trimmed_cs$Date),  origin = "1899-12-30")
  return(table_trimmed_cs)
}

sheet_names <- excel_sheets("C:/thepath/the_calls.xlsx") 
all_sheets_cs <- list()

for (s in sheet_names) {
  print(s)
  key <- s
  df <- read_excel("C:/thepath/the_calls.xlsx", sheet = s, col_names = FALSE, range = "A1:CO85")
  cleaned_df <- build_worksheets_cs(df)
  all_sheets_cs[[key]] <- cleaned_df
}
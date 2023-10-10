library(tidyverse)
library(janitor)
library(lubridate)
library(openxlsx)


clean <- function(student){
  
# select file  
  file_path <- paste0("student_", student, ".xlsx")
  
# select sheet  
  sheets <- readxl::excel_sheets(file_path)
  
  choice <- menu(sheets, title = "Select a sheet:")
  
  if (choice == 0) {
    cat("No sheet selected.")
    return(NULL)
  } 
  
# TYT
  else if (choice == 1) {
    cat("You selected sheet:", sheets[choice], "\n")
## set cols    
    data <- readxl::read_xlsx(paste0("student_", student, ".xlsx"))
    colnames(data) <- c("DATE",
                        "Tur_Dogru", "Tur_Yanlis", "Tur_Net",
                        "Tar_Dogru", "Tar_Yanlis", "Tar_Net",
                        "Cog_Dogru", "Cog_Yanlis", "Cog_Net",
                        "Fel_Dogru", "Fel_Yanlis", "Fel_Net",
                        "Din_Dogru", "Din_Yanlis", "Din_Net",
                        "Mat_Dogru", "Mat_Yanlis", "Mat_Net",
                        "Geo_Dogru", "Geo_Yanlis", "Geo_Net",
                        "Fiz_Dogru", "Fiz_Yanlis", "Fiz_Net",
                        "Kim_Dogru", "Kim_Yanlis", "Kim_Net",
                        "Bio_Dogru", "Bio_Yanlis", "Bio_Net",
                        "Tot_Dogru", "Tot_Yanlis", "Tot_Net")
    
    
    data <- data[-c(1:2, (nrow(data) - 1):nrow(data)), ]
## clean date
### trim notes
    data$DATE <- gsub("\\s*\\([^)]+\\)\\s*", "", data$DATE)
### standardize formats
    standardize_date <- function(date) {
      
      excel_date <- tryCatch(
        excel_numeric_to_date(as.numeric(as.character(date)), date_system = "modern"),
        error = function(e) NA  
      )
      
      if (!is.na(excel_date)) {
        return(excel_date)
      } else {
        
        european_date <- tryCatch(
          dmy(str_replace(date, "\\.", "-")),
          error = function(e) NA  
        )
        
        if (!is.na(european_date)) {
          return(european_date)
        } else {
          
          full_european_date <- tryCatch(
            dmy(date),
            error = function(e) NA  
          )
          
          if (!is.na(full_european_date)) {
            return(full_european_date)
          } else {
            return(NA)  
          }
        }
      }
    }
    
### Apply
    data$DATE <- sapply(data$DATE, standardize_date)

    data$DATE <- as.Date(data$DATE)
    
    data <- data %>% 
      mutate_if(function(x) !is.Date(x), as.numeric)
## NAs to 0s    
    data <- replace(data, is.na(data), 0)
    
## import    
    df_name_var <- paste0(student)
    
    assign(df_name_var, data, envir = .GlobalEnv)
  }
  
# AYT
  else if (choice == 2) {
    cat("Your selected data type", sheets[choice], "is not ready yet.")
    return(NULL)
    #data <- readxl::read_xlsx(paste0("student_", student, ".xlsx"), sheet = 2)
  }
# Brans
  else {
    cat("You selected sheet:", sheets[choice], "\n")
    data <- readxl::read_xlsx(paste0("student_", student, ".xlsx"), sheet = sheets[choice])
    colnames(data) <- c("DATE", "Dogru", "Yanlis", "Net")
    
    data <- data[-c(1:2, (nrow(data) - 1):nrow(data)), ]
    data <- data[complete.cases(data$Net), ]
    data[] <- lapply(data, function(x) ifelse(is.na(x), 0, x))
    
    data$DATE <- as.Date(substr(data$DATE, 1, 10), format = "%d.%m.%Y")
    
    df_name_var <- paste0(student,"_", sheets[choice])
    
    assign(df_name_var, data, envir = .GlobalEnv)
  }   
}
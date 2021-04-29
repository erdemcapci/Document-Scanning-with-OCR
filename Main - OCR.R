#install.packages("stringr")
#library(installr)
#updateR()

library(tesseract)
library(pdftools)
library(magick)
library(readxl)
library(writexl)
library(stringr)

library(dplyr)




v_neigb <- 50


result_df <- data.frame(Name_of_Document=character(),
                        Searched_Text=character(), 
                        Page=integer(),
                        Text_Neighbourhood=character(),
                        Text_Start=integer(),
                        Text_End=integer(),
                        stringsAsFactors=FALSE)

v_word_list <-read_excel("C:/Users/XXXXX/Desktop/RProjects/OCR/Input/Search_Words.xlsx", sheet = "Main")
v_word_list_vector<-unlist(v_word_list,use.names = FALSE)
eng <- tesseract("eng")
setwd("C:/Users/XXXXX/Desktop/RProjects/OCR")
getwd()
v_proj_dir <- getwd()
my.file.rename <- function(from, to) {
  todir <- dirname(to)
  if (!isTRUE(file.info(todir)$isdir)) dir.create(todir, recursive=TRUE)
  file.rename(from = from,  to = to)
}
v_pdf_path <- paste0(v_proj_dir,"/PDF_Files")
v_list_of_pdffiles <- list.files(path = v_pdf_path, pattern = "pdf",  full.names = TRUE)
v_pdf_count <-length(v_list_of_pdffiles)
setwd("./Png_Files")
getwd()
lapply(v_list_of_pdffiles, FUN = function(files) {
  pngfile<-pdf_convert(files, format = "png", dpi = 600)
})
v_list_of_pngfiles <- list.files(path = ".", pattern = "png",  full.names = TRUE)


lapply(v_list_of_pngfiles, FUN = function(files) {
  text <- tesseract::ocr(files, engine = eng)
  v_base_filename<-tools::file_path_sans_ext(basename(files))
  v_txt_path=paste0(v_proj_dir,"/Text_Files/",v_base_filename,".txt")
  write.table(text, file = v_txt_path, sep = "\t",
              row.names = TRUE, col.names = NA)
})

setwd("..")
getwd()

v_txt_path <- paste0(v_proj_dir,"/Text_Files")
v_list_of_txtfiles <- substr(list.files(path = v_txt_path, pattern = "txt",  full.names = FALSE), 1, nchar(list.files(path = v_txt_path, pattern = "txt",  full.names = FALSE))-4)
v_list_of_full_txtfiles <- list.files(path = v_txt_path, pattern = "txt",  full.names = TRUE)
v_txt_count <-length(v_list_of_txtfiles)
v_list_of_txtfiles
v_list_of_full_txtfiles

setwd("./Text_Files")
getwd()

v_data_frame <- data.frame(File_Page=character(), 
                 User=character(), 
                 stringsAsFactors=FALSE) 
#v_txt_count

rm(v_txt_text)
rm(v_txt_text_vector)
v_txt_text_vector=c()

for(i in 1:v_txt_count){
  v_list_of_full_txtfiles[i]
  v_txt_text<-readChar(v_list_of_full_txtfiles[i], file.info(v_list_of_full_txtfiles[i])$size)
  v_txt_text
  v_txt_text_temp<-gsub("\n", " ",v_txt_text[1], fixed=TRUE)
  v_txt_text_temp_1<-gsub("\r", "",v_txt_text_temp, fixed=TRUE)
  v_txt_text_temp_2<-gsub("\t", "",v_txt_text_temp_1, fixed=TRUE)
  v_txt_text_vector[i]<-gsub('\"',"",v_txt_text_temp_2, fixed=TRUE)
}


result_df <- data.frame(Name_of_Document=character(),
                        Searched_Text=character(), 
                        Page=character(),
                        Text_Neighbourhood=character(),
                        Text_Start=character(),
                        Text_End=character(),
                        stringsAsFactors=FALSE)

#result_df <- bind_rows(result_df, c(Name_of_Document="erdem",Searched_Text= "erdem_search", Page= "4", Text_Neighbourhood="asdfsdfsdfsf", Text_Start= "13",Text_End= "25"))

v_search_txt_count <-length(v_word_list_vector)

for(i in 1:v_txt_count){

  for(k in 1:v_search_txt_count){
  
v_search_results_l=str_locate_all(v_txt_text_vector[i], fixed(v_word_list_vector[k], ignore_case=TRUE))
v_search_results_t <- unlist(v_search_results_l[[1]], use.names=FALSE)
v_search_count <- length(v_search_results_t)/2
v_search_text_vector <- rep(v_word_list_vector[k], each=v_search_count)
v_name_of_doc <- rep(tools::file_path_sans_ext(basename(v_list_of_full_txtfiles[i])), each=v_search_count)
v_page_of_doc <- rep(tail(unlist(strsplit(tools::file_path_sans_ext(basename(v_list_of_full_txtfiles[i])),"_"), use.names=FALSE), n=1), each=v_search_count)


v_txt_start <- c()
v_txt_end <- c()
v_neigh_txt <- c()
v_search_not_found <- c()
rm(v_txt_start)
rm(v_txt_end)
rm(v_neigh_txt)
rm(v_search_not_found)
v_txt_start <- c()
v_txt_end <- c()
v_neigh_txt <- c()
v_search_not_found <- c()

if (v_search_count>0) {
  for(j in 1:v_search_count){
    
    v_txt_start[j] <- v_search_results_t[j,1]-v_neigb
    v_txt_end[j] <- v_search_results_t[j,2]+v_neigb
    v_neigh_txt[j] <- substr(v_txt_text_vector[i], v_txt_start[j], v_txt_end[j])
  }
}

df_sub <- v_search_results_t
colnames(df_sub)[1] <- "Text_Start"
colnames(df_sub)[2] <- "Text_End"
#v_search_results_v <- as.vector(v_search_results_l[1])

result_df <- rbind(result_df,cbind(df_sub, Searched_Text=v_search_text_vector,Name_of_Document=v_name_of_doc,Page=v_page_of_doc,Text_Neighbourhood=v_neigh_txt))

}
}

setwd("..")
setwd("./Output")
write_xlsx(result_df, 
           "Results.xlsx", 
           col_names = TRUE,
           format_headers = TRUE)



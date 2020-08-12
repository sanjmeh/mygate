# squad data analysis

library(wrapr)
library(tabulizer)
library(bit64)
library(readxl)
library(openxlsx)
library(data.table)
library(stringr)
library(splitstackshape)
library(magrittr)
library(lubridate)
library(ggplot2)
library(googlesheets)
library(googlesheets4)
library(googledrive)
library(htmlTable)
library(tidyverse)
library(cellranger)

bonus_18_19 <- "1jv0Bj0NokUv-OKCWBcKGoo3_iTy3zEuu5rD8HTukpWA"
leave_18_19 <- "13cbuAkD-Cf8QqZu_HZDM_0OPm0niOUsK_pJHSJTjMz0"

# uses pf.txt and esic.txt to get current guards UAN ids. Add to these when new guard joins.

# Extracts table from the PDF file of ESIC received from Squad India; uses tabulizer package as there is no wrapping of table cells.
read_esic_table <- function(path=".",esic_filepat="ECR APRIL"){
  fname <- list.files(path = path,pattern = esic_filepat)[1]
  message("Processing file:",fname)
  x1 <- tabulizer::extract_tables(fname)
  x2 <- x1 %>% map_dfr(~as.data.frame(as.matrix(.x),fill=T))
  setDT(x2)
  x2 <- x2[!(is.na(V8)|grepl("Page",V1)|V1=="")]
  x2 <- x2[,qc(V2,V9):=NULL]
  x2[,V1 := str_extract(V1,"\\d+")]
  x2 <- x2[!is.na(V1)]
  setnames(x2,qc(sn,ipnb,ipname,ndays,wage,ipcontr,reason))
}

# outputs a numeric matrix with first column as UAN and others deductions. 
# does not use tabulizer instead we use pdf_text as one table row consists of many data rows due to wrapping.
# Table extraction misses wages row grep UAN rows and hence sometimes we get only deductions.
read_pf_ecr <- function(path="EPF ECR COPIES", pf_filepat = "APRIL-2020"){
  fname <- list.files(path = path,pattern = pf_filepat,full.names = T)[1]
  message("Processing file:",fname)
  x1 <- pdf_text(fname)
  x3 <- x1 %>% map( ~str_split(.x,"\n") %>% map(str_trim)) %>% unlist
  p1 <- x3 %>% str_extract("(?=10\\d{10}).+[A-Z]") %>% na.omit
  data.table(x=p1) -> dt1
  dt1 %>% cSplit("x",sep=" ") -> dt2
  dt2 %>% lapply(as.character) %>% lapply(str_remove,",") %>% lapply(as.numeric) %>% as.data.table -> dt3
  as.matrix(dt3) -> m1
  if(mean(m1[,2],na.rm = T)<2000) 
    apply(m1,1,function(x) na.omit(x) %>% .[1:4]) -> t1
  else
    apply(m1,1,function(x) na.omit(x) %>% .[1:8]) -> t1
  t2 <- t(t1)
  if(dim(t2)[2]==4) t1[which(t1[,2] != t1[,3] + t1[,4]),4] <- 0 
  t2
}


#vislogy1 <- ovislog[crfact(entry)>="Jul_2018" & crfact(entry)<="Mar_2019"]

downsheet <- function(id=bonus_18_19,range="B2:BR28",saveas="bonus.csv"){
  x1 <- read_sheet(as_id(id),sheet =1,range = range,trim_ws = T)
  fwrite(x1,saveas)
}

load_bonus <- function(f="bonus.csv"){
  
  bonus <- fread(f,colClasses = list(character=c(1:5),numeric=c(6:69))) # assuming total 69 columns in the googlesheet for bonus
  setnames(bonus,grep("Total Duty",names(bonus)),paste0("totduty:",qc(Jul_2018,Aug_2018,Sep_2018,Oct_2018,Nov_2018,Dec_2018,Jan_2019,Feb_2019,Mar_2019)))
  setnames(bonus,grep("Normal Duty",names(bonus)),paste0("normduty:",qc(Jul_2018,Aug_2018,Sep_2018,Oct_2018,Nov_2018,Dec_2018,Jan_2019,Feb_2019,Mar_2019)))
  setnames(bonus,grep("Bonus",names(bonus)),paste0("bonus:",qc(Jul_2018,Aug_2018,Sep_2018,Oct_2018,Nov_2018,Dec_2018,Jan_2019,Feb_2019,Mar_2019)))
  setnames(bonus,grep("...69",names(bonus)),"Total_Bonus")
  setnames(bonus,grep("emp|mg",names(bonus),ig=T),qc(empname,myname))
  bon1 <- bonus[,.SD,.SDcols=grep("empname|myname|duty|bonus",names(bonus),ig=T)]
  melt(bon1,id.vars = qc(empname,myname))
}

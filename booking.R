library(data.table)
library(magrittr)
library(lubridate)
library(googlesheets)
library(googlesheets4)
library(stringr)
library(wrapr)
library(splitstackshape)
library(htmlTable)

if(search() %>% str_detect("key_all") %>% any() %>% not) attach("key_all")

fetch_booking_data <- function(link = mango_url){
  x1 <- googlesheets4::read_sheet(ss = link,sheet = 1,col_types = "tc__ccccc-----")
  setDT(x1)
  setnames(x1,qc(time,email,apt,flat,name,cont,order))
  x1 <- x1[!is.na(time)]
  return(x1)
}

process_bookings <- function(x1,start="20200526"){
  setnames(x1,qc(time,email,apt,flat,name,cont,order))
  #x1[,rasp:=NULL]
  #x1[,old:=NULL]
  x2 <- x1[time > ymd(20200525)]
  x3 <- cSplit(x2,"order",)
  x4 <- melt(x3,id.vars = qc(time,email,apt,flat,name,cont),na.rm = T)
  x4[,fruit:=str_extract(value,"(?<=kg ).{5}")]
  x4[,kg:=str_extract(value,"\\d+") %>% as.numeric()]
}


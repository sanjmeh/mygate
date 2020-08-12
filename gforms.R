library(purrr)
library(magrittr)
library(stringr)
library(splitstackshape)
library(data.table)
library(magrittr)
library(googlesheets4)
library(lubridate)
library(wrapr)
library(gmapsdistance)
library(htmlTable)

if(search() %>% str_detect("keys.data") %>% any() %>% not) attach("keys.data")
surveydates <- c("April 8th","April 9th","April 10th","April 11th","April 12th")

# pass pattern "kkd" or "hopcom"
download_survey <- function(pattern="kkd") {
  if(pattern=="hopcoms") {
    x1 <- sheets_read(as_sheets_id(hopcoms_link),sheet = 1,col_types = c("tcccccccc___cciccicccc"))
    setnames(x1,new = qc(ts,isbaf,bafn,apt,map,name,mob,pdates,email,delstatus,apt2,flats,addr,loc,pin,clustno,clustname,plan,seq))
  }
  else 
    if(pattern=="kkd") {
    x1 <- sheets_read(as_sheets_id(kkd_link),sheet = 1)
    setnames(x1,c(1:3),c("ts","email","apt"))
    setnames(x1,c("Your name", "Your apartment number" ),c("name","flat"))
    } else 
      if(pattern=="grapes") {
        x1 <- sheets_read(as_sheets_id(grapes_link),sheet = 1)
        setnames(x1,c("ts","email","apt","green","red","flat","name"))
      }
  setDT(x1)
  fwrite(x1, file = switch(pattern,kkd="kkd.csv",hopcoms= "hopcoms.csv", grapes= "grapes.csv"),dateTimeAs = "write.csv")
  return(x1)
}

proc_hop <- function(file="hopcoms.csv"){
        x1 <- fread(file)
        x1[,ts:=as_datetime(ts,tz = "Asia/Kolkata")]
        suppressWarnings(x1[,deldate:=parse_date_time(delstatus,orders=c("mdy","dmy")) %>% as.Date])
        x1[,delstatus:=ifelse(!is.na(deldate),"DELIVERED",delstatus)]
        x2 <- x1[delstatus==""] %>% unique(by=c("bafn"))
        x2[,addr_pin:=paste("Bengaluru+Karnataka",pin,sep="+") ]
        x2[,aptname_str:=apt2 %>% str_replace_all("[#,&]"," ") %>% str_replace_all("[\\s]+","+")]
        x2[,addr_part:=paste(aptname_str,addr_pin,sep="+")]
        x3 <- x2[,tmp:=str_replace_all(addr,"[,#]"," ") %>% strsplit("\\s+") %>% map(paste,collapse="+")][,addr_full:=paste(tmp,"Bangalore",sep = "+")][,tmp:=NULL]
}

proc_kkd <- function(file="kkd.csv"){
        x1 <- fread(file)
        setnames(x1,c(1:3),c("ts","email","apt")) # as first three columns will always be these: make sure while designing your form
        x1[,ts:=as_datetime(ts,tz = "Asia/Kolkata")]
        if(any(grepl("Your name",names(x1)))) setnames(x1,c("Your name"),c("name"))
        if(any(grepl("Your apartment number",names(x1)))) setnames(x1,c("Your apartment number"),c("apt"))
        if(any(grepl("any other item",names(x1)))) setnames(x1,c("Add any other item with full details and vendor will attempt to procure for you"),c("others"))
        x2 <- suppressWarnings(melt(x1,id.vars = qc(name,apt,flat,ts,email)))
        x2[,type:=str_extract(variable,"[A-Za-z ]*(?= \\[)")]
        x2[,stuff:=str_extract(variable,"(?= \\[).*") %>% str_remove_all("\\[|\\]")]
        x2[variable != "others",.(qty=str_extract_all(value,"\\d+") %>% unlist %>% as.integer %>% sum(na.rm = T)),by=.(apt,flat,stuff,type)]
        #x2[,qty:=ifelse(variable=="others",0,str_extract_all(value,"\\d+") %>% unlist %>% as.integer %>% sum(na.rm = T)),by=.(apt,flat,stuff)]
}

# aggregation can be "replace" or "sum"
proc_grapes <- function(file="grapes.csv",aggregation="replace"){
  x1 <- fread(file)
  x1[,qty_g:=str_extract(green,"\\d+") %>% as.double()]
  x1[,qty_r:=str_extract(red,"\\d+") %>% as.double()]
  if(grepl("repl",aggregation))
  x1[,.(ts=last(ts),green=last(qty_g), red=last(qty_r),ntime=.N), by = .(apt, flat, email, name)][order(apt,flat)]
  else
    x1[,.(ts=last(ts),green=sum(qty_g,na.rm = T), red=sum(qty_r,na.rm = T),ntime=.N), by = .(apt, flat, email, name)][order(apt,flat)]
}


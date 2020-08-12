library(magrittr)
library(wrapr)
library(readxl)
library(googlesheets)
library(googlesheets4)
library(data.table)
library(htmlTable)
library(lubridate)
library(stringr)
library(splitstackshape)
if(search() %>% str_detect("keys.data") %>% any() %>% not) attach("keys.data")

key_groc_form <- googlesheets4::as_sheets_id("1UIPzbLprgcC_hc9S4jqN9rRxJVFYhuMLhEbOPmrDiS8")
key_bigbask <- "1Lin_1UsKLEgQab0LUvMBJJL8r2AQM8M5fY9_odWCJqo"

download_orders <- function(){
 ords <- googlesheets4::read_sheet(key_groc_form,trim_ws = T)
 setDT(ords)
 fwrite(ords,"ords.csv",dateTimeAs = "write.csv")
 message("Written file ords.csv with ",nrow(ords)," rows")
}

download_bb_orders <- function(download=T){
  if(download==T) gs_download(gs_key(key_bigbask),overwrite = T)
  bbords <- seq_len(4) %>% map(read_excel,path="bigbasket-ordering.xlsx")
  bbords <- bbords %>% map(setDT)
  bb2 <- suppressWarnings(bbords %>% map(melt,id.vars = qc(SKU,sku_description,`Selling Price`),variable.name="flat",value.name="qty"))
  dtords <- bb2 %>% map(~.x[!is.na(qty)]) %>% rbindlist
  dtords <- dtords[grepl("[ABCDE][1-8]|[AB]-",flat)]
  dtords[,qty:=as.integer(qty)][,total:=qty * `Selling Price`]
}

bb_summary <- function(dt,summary=1) {
  if(summary==1) dtsum <- dt %>%  dcast(SKU + sku_description + `Selling Price` ~ flat,value.var = "qty", fun=sum)
  if(summary==2) dtsum <- dt %>%  dcast(SKU + sku_description + `Selling Price` ~ flat,value.var = "total", fun=sum)
  if(summary==3) dtsum <- dt[,.(total=sum(total)),by=flat] %>% janitor::adorn_totals(where = "row")
  if(summary==4) dtsum <- dt[,.(total_orders=sum(qty),total_amt=sum(qty *`Selling Price`)),by=.(SKU,sku_description)] %>%  janitor::adorn_totals(where = "row")

  return(dtsum)
}
 
proc_ord <- function(file="ords.csv"){
  ords <- fread(file)
  setnames(ords,grep("Time|Apartment|name|number|flat",names(ords),ig=T,val=T),c("time","apt","flat","name","number"))
  ords <- ords[,datetime:=parse_date_time(time,orders = c("ymdHMS"))]
  #ords[,flatu:=paste(Block,flat,sep = "_")] # add apat name prefix when we have more than one
  ords[,flatu:=flat] # add apat name prefix when we have more than one
  ords <- ords[,ord_by:=paste(name,flatu,sep = "/ ")][,Date:=as.Date(time)]
  ords[,time:=NULL]
}

tabulate_flatwise <- function(dt=ords,strtdate="mar29",strttim="00:00",earliest=20,uniq=T){
x1 <- melt(dt[datetime>=mdy_hm(paste(strtdate,2020,strttim))][seq_len(earliest)],
           id.vars=grep("flatu",names(dt),ig=T,val=T),measure.vars = patterns("Fru|Veg|Hyg|Groc")) %>% 
  cSplit(splitCols = "value") %>% 
  melt(measure.vars=patterns("value")) %>% 
  .[!is.na(value),-c("variable.1")] %>% 
    .[order(flatu,variable,value)]
if(uniq) unique(x1) else x1
}

htmlorddet <- function(ordtab){
ordtab[,.(category=variable,item=value)] %>% 
    htmlTable(rgroup = unique(ordtab$flatu),n.rgroup = ordtab[,.N,by=flatu][,N],align="l",css.rgroup = "font-size: 30px;")
}


tabulate_item_flat <- function(dt=ords, strtdate="mar29",strttim="00:00",earliest=20,uniq=T){
  x1 <- melt(dt[datetime>=mdy_hm(paste(strtdate,2020,strttim))][seq_len(earliest)],id.vars="flatu",measure.vars = patterns("Fru|Veg|Hyg|Groc")) %>% 
    cSplit(splitCols = "value") %>% 
    melt(measure.vars=patterns("value")) %>% 
    .[!is.na(value),-c("variable.1")] %>% 
    .[order(flatu,variable,value)]
  if(uniq) x1 <- unique(x1)
  x2 <- x1[,.N,by=.(category=variable,item=value,flatu)][order(flatu)]
    res1 <- dcast(x2,category + item ~ flatu,fill=0) %>% 
    janitor::adorn_totals(where = c("row","col"))
    majheaders <- names(res1)[2:ncol(res1)] %>% str_sub(1,3)
    res1
}

html_storelist  <- function(dt){
  dt[,-c("category")] %>% htmlTable(rgroup = unique(dt$category),n.rgroup=dt[,.N,by=category][,N],align = "l",css.rgroup = "font-size: 30px;")
}

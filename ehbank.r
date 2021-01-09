# ===== Library loading =======
library(lubridate)
library(assertthat)
library(purrr)
library(wrapr)
library(formattable)
library(tabulizer)
library(data.table)
library(magrittr)
library(dplyr)
library(cellranger)
library(splitstackshape)
library(anytime)
library(stringr)
library(readxl)
library(openxlsx)
library(googledrive)
library(googlesheets4)
library(htmlTable)
library(htmltools)
library(knitr)
library(kableExtra)
library(textclean)
library(janitor)
source("~/Dropbox/bank-statements/globalfns.r")
if(search() %>% str_detect("keys.data") %>% any() %>% not) attach("keys.data")

source("~/Dropbox/bank-statements/globalfns.r")
# ==== Description of steps =======
# New process after launchiing googlesheets4 on July 25: single line append command for new bank statements
# 1. pulltxt("SANJAUTH@EHAOA_1595689956077.txt")[cr_date>=ymd(20200708)][order(cr_date),.SD,.SDcols=names(x1)[1:7]] %>% sheet_append(key_embassy_master,.,sheet = "master")
# 2. pull_idbi_pdf()[cr_date>ymd(20200701) & day(cr_date)>=14] %>% sheet_append(ss = key_embassy_master,data = .,sheet = "IDBI")

# downmaster is now simplified skipping functionality of credit booking and supplementary credit tracking.

#source("park.r")
#source("mygate.R")
# ===== Loading global variables =================
drive_auth_configure(api_key = apikey_googlesheets)

if(!exists("rdu")) rdu <- fread("rdu.txt")
dr_hdfc <- readRDS("hdfcmaster_dribble.rds")
NAMES1 <- c( "cr_date",  "cr",       "db"  ,     "narr",     "balance" , "chq"   ,   "mode"  , "bank_tx_id",  "chq_date", "bank"  ,   "category", "flats"   , "month", "booked_cr",	"booked_db",	"rectnb" )
NAMES2 <- c( "cr_date",	"bank_tx_id",	"category",	"flats",	"booked_cr",	"rectnb")
total_due <- sum(rdu$sqft*5.05*1.18)
flatcats <- c("Sinking fund", "Maintenance", "Penalty", 
              "Interior work charges", "Shifting IN/OUT", "Common area")

# ====== FUNCTIONS ================
# to be modified to gs4: 
downcoll<- function(g2=gs_key(key_chqcoll)) as.data.table(gs_read(g2,ws = "new_format",range = cell_cols("A:G"),skip=2,col_names=F)) #this will work on google sheet Society cheque collection

# new function - using googlesheets4
downmaster <- function(keymaster=key_embassy_master){
  dt <- range_read(ss = key_embassy_master, sheet = "master",range = "A:R",col_types = c("TddcdccdTcc?cccccc"))
  setDT(dt)
  # confl_flats <- dt[!is.na(Flat_nbs) & !is.na(flats) & Flat_nbs != flats]
  # if(nrow(confl_flats)>0) {
  #   message("Conflicting flat taggings corrected:")
  #   print(confl_flats[,.(cr_date,Amount=cr,mode,narr,flats,Flat_nbs)])
  # }
  # dt[!is.na(Flat_nbs),flats:=Flat_nbs]
  # 
  # confl_cats <- dt[!is.na(Category) & !is.na(category) & category!=Category]
  # if(nrow(confl_cats)>0){
  #   message("Conflicting category taggings corrected:")
  #   print(confl_cats[,.(category,Category)])
  # }
  # dt[!is.na(Category),category:=Category]
  #dt[,category:=as.factor(category)]
  cat("Writing file..")
  fwrite(dt,file = "hdfc_master.csv",dateTimeAs = "write.csv")
  cat("written")
  dt
}

update_cells <- function(dt){
  dt[,bank_tx_id:=rownames(x1) %>% as.numeric()]
  dt[!is.na(cr) & !is.na(flats) & is.na(rectnb),rectnb:=bank_tx_id] # only update missing recpt nos.
  dt[,c("Category","Flat_nbs","rn") := NA]
  dt[,c(1:16)]
}

read_bank_file <- function(f="hdfc_master.csv"){
  fread(f,colClasses = list(character=c(4,6,7,9:10,13:18),numeric=c(2,3,5,8),POSIXct=c(1),factor=c(11,12)))
}


# NOt used now
proc_xls_hdfc <- function(dfile="bank_statement_dump2.xlsx"){
  n1 <- readxl::read_excel(dfile,sheet = "master",n_max = 0) %>% names()
  n2 <- readxl::read_excel(dfile,sheet = "supplementary",n_max = 0) %>% names()
  ctypes1 <- ifelse(grepl("date$",n1),"date",ifelse(grepl("cr$|db$|bal|amount|amt|id$",n1),"numeric","text")) # neat trick
  ctypes2 <- ifelse(grepl("date$",n2),"date",ifelse(grepl("cr$",n2),"numeric","text")) # assuming same heads as 1
  dt <- readxl::read_excel(path = dfile,sheet = "master",col_types = ctypes1) %>% setDT
  
  oldsupp <- readxl::read_excel(path = dfile,sheet = "supplementary",col_types = ctypes2) %>% setDT
  oldsupp[,bank_tx_id:=as.integer(bank_tx_id)]
  oldsupp[,flats:=as.integer(flats)]
  dt <- dt[!is.na(cr_date)] # filter out only valid rows
  oldsupp <- oldsupp[!is.na(cr_date)] # filter out only valid rows
  oldsupp[is.na(bank_tx_id),bank_tx_id:=as.integer(str_extract(rectnb,"\\d+"))]
  #new column names: Flat_nbs	month_of_maintenance	bank_chq_dt	bank_name	adjusted_amout	adjusted_category
  setnames(dt,old = grep("invoice",names(dt),ig=T,va=T),new="invoice_no")
  setnames(dt,old = grep("vendor",names(dt),ig=T,va=T),new="vendor")
  dt[,chq_date:=as.Date(chq_date)]
  dt[,bank_chq_date:=as.Date(bank_chq_date)]
  # in case we get numeric dates as often happens in excel
  # dt[grepl("^43",month_of_maintenance), month_of_maintenance:=as.integer(month_of_maintenance) %>%  as.Date(origin="1899-12-31") %>% crfact]
  # dt[grepl("^43",month), month:=as.integer(month) %>%  as.Date(origin="1899-12-31") %>% crfact]
  dt[!is.na(bank_chq_date), chq_date:= bank_chq_date]
  dt[!is.na(bank_name), bank := bank_name]
  # dt[!is.na(bank_chq_date), chq_date:=as.Date(x = as.numeric(bank_chq_dt),origin="1899-12-30")]
  # dt[!is.na(bank_name), bank := bank_name]
  dt[!is.na(Category),category:=Category] # use googlesheet updated categories from downloaded master
  #dt[!is.na(adjusted_category),category2:=adjusted_category] # use googlesheet updated categories from downloaded master
  if (nrow(dt[!is.na(Category)]) >0)
    base::message(sprintf("Updated %d categories",nrow(dt[!is.na(Category)])))
  if (nrow(dt[!is.na(adjusted_category)]) >0) # new row added
    base::message(sprintf("Updated %d ADJUSTED categories",nrow(dt[!is.na(adjusted_category)])))
  conflicting_flats <- dt[!is.na(Flat_nbs) & !is.na(flats) & Flat_nbs != flats]
  #conflicting_months <- dt[!is.na(month_of_maintenance) & !is.na(month) & month_of_maintenance != month]
  if(nrow(conflicting_flats)>0) {
    base::message("Clash of flat numbers:"); 
    print(conflicting_flats[,.(cr_date,cr,chq,narr,flats,new_flat_nbs=Flat_nbs)])
    }
  #if(nrow(conflicting_months)>0) {base::message("Clash of months numbers:"); print(conflicting_months)}
  stopifnot(readline(prompt = "Proceed?") %in% c("Y","y"))
 dt[!is.na(Flat_nbs), flats:=as.character(Flat_nbs) %>% str_replace_all("[/:,-]",";")]
 # dt[!is.na(Flat_nbs), flats:=str_split(Flat_nbs,pattern = "[-, :/;]") %>% lapply(drop_element,"^$") %>% lapply(paste,collapse=";")]
  
  #dt[!is.na(month_of_maintenance), month:=as.character(month_of_maintenance) %>% str_replace_all(",",";")]
  dt[chq==0,chq:=NA]
  if ( nrow( dt[!is.na(Flat_nbs) & mode %in% c("ONLINE")]) >0) 
    base::message(sprintf("Updated %d flats with online payments",nrow( dt[!is.na(Flat_nbs) & mode %in% c("ONLINE")])))
  if( nrow(dt[!is.na(adjusted_category)])) 
    base :: message(sprintf("Adjusted %d rows with multiple category of payments",nrow(dt[!is.na(adjusted_category)])))
  # Adjust the credit or debit for those transactioins that need to be split
  dt[!is.na(adjusted_category) & !is.na(cr), booked_cr:=cr - adjusted_amount]
  dt[!is.na(adjusted_category) & !is.na(db), booked_db:=db - adjusted_amount]
  # new column added for only bank transactions but this will need to be changed after we load the first round of bankids. So that by mistake we donot override.
  dt[(!is.na(cr) | !is.na(db)),bank_tx_id:=as.double(seq_len(.N))] 
  dt[!is.na(flats),rectnb:=ifelse(is.na(rectnb),as.character(bank_tx_id),rectnb)] # copy the bank tx id as recptnb for all non NA rect numbers; if bank_tx is also NA so be it.
  # the above is an important line of code that protects rect nbs that are already circulated. It ensures that even if bank_tx_id changes for the same record the rectnb remains the same for all old transactions.
  
  # commented out... work on pushing these to a new sheet - supplementary instead of adding rows in the same bank sheet.
  newsupp <- dt[!is.na(adjusted_category) & !is.na(cr),.(cr_date,booked_cr=adjusted_amount,bank_tx_id,category=adjusted_category,flats,rectnb=paste(bank_tx_id,"/2"))] # add a supplementary receipt number by suffixing to the original transaction id
  updated_supp <- rbind(oldsupp,newsupp,fill=T) %>% unique(by="bank_tx_id")
  updated_supp <- updated_supp[is.na(rectnb),rectnb:= paste(bank_tx_id,"/2")][booked_cr != 0][order(bank_tx_id)]
  setcolorder(dt, qc(cr_date,cr,db,narr,balance,chq,mode,bank_tx_id,chq_date,bank,category,flats,	month,booked_cr,booked_db,rectnb,vendor,invoice_no))
  dt <- dt[!is.na(cr_date),c(1:18)] # ignore data entry columms & NA rows
  dt <- unique(dt,by=qc(cr_date,narr,category,balance,bank_tx_id))
  dt <- mark_mode(dt)
  list(master=dt,supp=updated_supp)
}

# read the IDBI worksheet from the bank master downloaded xlsx file : convert this to direct googlesheet read now... pending
proc_xls_idbi <- function(dfile="bank_statement_dump2.xlsx"){
  n1 <- readxl::read_excel(dfile,sheet = "IDBI",n_max = 0) %>% names()
  ctypes1 <- ifelse(grepl("date$",n1),"date",ifelse(grepl("cr$|db$|bal|amount|amt|id$",n1),"numeric","text")) # neat trick
  dt <- readxl::read_excel(path = dfile,sheet = "IDBI",col_types = ctypes1) %>% setDT
  setnames(dt,old = grep("invoice",names(dt),ig=T,va=T),new="invoice_no")
  setnames(dt,old = grep("vendor",names(dt),ig=T,va=T),new="vendor")
  dt
}

# this still needs a downloaded excel file.. 
proc_xls_cheques <- function(dfile="bank_statement_dump2.xlsx"){
  n1 <- readxl::read_excel(dfile,sheet = "cheques_issued",n_max = 0) %>% names()
  ctypes1 <- ifelse(grepl("date$",n1,ig=T),"date",ifelse(grepl("cr$|db$|bal|amount|amt|id$",n1),"numeric","text")) # neat trick
  dt <- readxl::read_excel(path = dfile,sheet = "cheques_issued",col_types = ctypes1) %>% setDT
  dt
}

mark_mode <- function(st_master =NULL) {
  st_master[grepl(pattern = "RTGS|UPI|IMPS|NEFT|TPT|NET BANKING|FT - CR|PAYU|LHDF|HDFC[0-9]{8,}|NHDF|SI HD0|RAZPBESCOM",x = narr,ignore.case = T),mode:="ONLINE"]    
  st_master[grepl(pattern = "FD BOOKED",x = narr,ignore.case = T),mode:="FD BOOKED"]    
  st_master[grepl(pattern = "INST-ALERT|INTEREST|RTGS CHGS",x = narr,ignore.case = T),mode:="BANK"]    
  st_master[grepl(pattern = "cash",x = narr,ignore.case = T),mode:="CASH"]
  st_master[grepl("TAX RECOVERY FOR TD",narr),mode:="TDS"]
  st_master[is.na(mode),mode:="CHEQUE"]
  st_master
}

# not used
# gives 11 warnings for first 11 rows
pullsheets <- function(pat="EH",nine=T,csv=T){
  
  # FN1: read an excel file with HDFC bank dump into a data table
  pullone <- function(filename=NULL) {
    rcomma <- function(x) suppressWarnings(as.numeric(gsub(x=x,pattern = ",",replacement = "")) %>% round(digits = 0) )
    x1<- read_excel(filename,sheet = 1,skip = 22,col_names = F,col_types = c("text","text","text","skip","numeric","numeric","numeric"),trim_ws = T)
    names(x1) <- c("cr_date","narr","chq","db","cr","balance")
    x1$cr_date  %<>% dmy
    new_data <- colwise(rcomma,.(cr,db,balance))(x1) %>% cbind(x1[,setdiff(names(x1),c("cr","db","balance"))]) %>% as.data.table()
    setcolorder(new_data,c("cr_date","cr","db","narr","balance","chq"))
    new_data[!is.na(cr_date)][order(cr_date)]
  }  
  
  pull9col <- function(filename=NULL) {
    .x1<- fread(filename)
    setDT(.x1)
    .x1[,c(2,4,7):=NULL]
    setnames(.x1,c("cr_date","narr", "chq","dc","amt","balance"))
    .x1[,cr_date:=mdy_hms(cr_date)]
    #new_data <- colwise(rcomma,.(cr,db,balance))(x1) %>% cbind(x1[,setdiff(names(x1),c("cr","db","balance"))]) %>% as.data.table()
    mark_mode(.x1)
    .x1[dc=="D",db:=amt]
    .x1[dc=="C",cr:=amt]
    setcolorder(.x1,c("cr_date","cr","db","narr","balance","chq"))
    .x1[,dc:=NULL]
    .x1[,amt:=NULL]
    .x1[!is.na(cr_date)][order(cr_date)]
  } 
  # FN2: tag mode of payment column on the read data
  
  files <- list.files(pattern = pat)
  ldata <- lapply(files,ifelse(nine,pull9col,pullone))
  do.call(rbind,args = ldata) %>% as.data.table %>% unique %>% mark_mode %>%  .[order(cr_date)]
}

# new function - works on text downloads from ENET - one file at a time
# we are matching the pullsheet() output therefore removing some columns and renaming balance
pulltxt <- function(fname){
  .x1 <- fread(fname)
  .x1$`Transaction Date` %<>% dmy_hms
  .x1$`Value Date` %<>% dmy
  setnames(.x1,c("Transaction Date","Transaction Description","Transaction Amount","Debit / Credit", "Reference No.", "Running Balance" ),
           c("cr_date","narr","amt","dc", "chq", "balance"))
  mark_mode(.x1)
  .x1[dc=="D",db:=amt]
  .x1[dc=="C",cr:=amt]
  .x1[,chq:= as.character(chq)]
  .x1[,`:=`(c("Value Date","amt", "dc", "Transaction Branch" ),NULL)]
}

# not used after 1-2 time
# new function - Binod's csv file (the date is now without time)
# date needs to be manually finetuned everytime
pulltxtB <- function(fname,dttm=F){
  .x1 <- fread(fname,select = c(1,3:6,8:9))
  setnames(.x1,c("cr_date","narr","mode","chq","dc","amt","balance"))
  if(dttm==T)  .x1$cr_date %<>% dmy_hms() else 
  .x1[,cr_date:= cr_date %>% parse_date_time(orders = c("dmyHM","dmy"))]
  .x1$chq %<>% as.character()
  .x1[mode=="CHQ",mode:="CHEQUE"]
  .x1[mode=="SCW",mode:="CASH"]
  .x1[!grepl("CHEQ|CASH",mode),mode:="ONLINE"]
  .x1[,amt:=as.numeric(str_remove(amt,","))]
  .x1[,balance:=as.numeric(str_remove(balance,","))]
  .x1[dc=="D",db:=amt]
  .x1[dc=="C",cr:=amt]
  .x1[,dc:=NULL]
  .x1[,amt:=NULL]
  .x1[order(cr_date)]
}


getrate <- function(flat) {if(flat %in% rdu$flatn)  rdu[flatn==flat,sqft*5.05*1.18] else NA}
getrates <- function(flatrange) map_dbl(flatrange,getrate)

#--- These functions are used to summarise the general Accounting ------

# Petty cash processing - modify to gs4
dpetty <- function(outfile="petty2.xlsx") range_read(ss = as_id(key_pettytall),to = outfile,overwrite = T) 

# pettyc() will read from local excel file which is set at the default as used in dpetty()
pettyc <- function(dfile="/Users/sm/Dropbox/RWAdata/petty2.xlsx") {
    x <- read_excel(path = dfile,sheet = "pettycash",col_names = T,range = cell_cols("A:H"),trim_ws = T) %>% 
      as.data.table
    x[grepl("With",part),category:="Withdrawal"][!is.na(Date)]
}

replZero <- function(DT=NULL){ # matt dowle technique on SO here https://stackoverflow.com/a/7249454/1972786
  for (j in seq_len(ncol(DT)))
    set(DT,i = which(DT[[j]]==0),j,value = NA)
}

# modify to gs4
downlov <- function(k = key_rudra_master,ws="lov"){
  lov <- gs_download(gs_key(k),ws = "lov",to = ifelse(grepl("6NlOJJEitc$",k),"lov1.xlsx","lov2.xlsx"),overwrite = T)
  #read_excel(lov) %>% as.data.table()
}

loadlov <- function(){
  x1 <- read_excel("lov1.xlsx") %>% as.data.table()
  #x2 <- read_excel("lov2.xlsx") %>% as.data.table() removed lov2 as combined the two
  #rbind(x1[,c(1:2)],x2[,c(1,2)],fill = T) %>% unique
  x1 <- x1[,c(1:2)]
  setnames(x1,c("subcat","maincat"))
}


# expense and revenue trends - monthly
mtrnd <- function(statement=st_master,collapse = F, maincat=F,
                  cashdt= pettyc(),lovmer = loadlov(),
                  lastn=10, expenditure=T,tall=F,
                  rm="fixed|withdrawal|invested|sinking|transfer|petty") { # Note: this regex pattern is used on category to remove these transactions from the monthly trend
  
  # very dangerous function - try not to use 
  rmNA <- function(dt,n=1){
    dt[,1:(colnum <<- (tilln +n))] ->t1
    rowSums(t1[,c(4:(colnum-1)),with = F],na.rm = T) -> t2
    t1[which(t2>0)]
  }
  
  statement <- statement[,mnth:=crfact(cr_date)][ 
      ,type:="electronic"][
        !grepl(rm,category,ig=T) & as.integer(mnth)>as.integer(crfact(now()))-lastn ] # lastn months data
        
  cashdt <- cashdt[,mnth:=crfact(Date)][,type:="cash"][ # the column name Date could be confusing - target to change it soon
    !grepl(rm,category,ig=T) & as.integer(mnth)> as.integer(crfact(now()))-lastn] 
  
  assert_that(is.data.table(cashdt))
  assert_that(is.data.table(lovmer))
  if(maincat){
    if(expenditure) var <- "sumdb" else var <- "sumcr" # will be used in dtwide1 & dtwide2 to select variable
    varwide <- lev1(dig = statement,cash = cashdt,lov = lovmer,sumlev = 1) %>% 
      dcast(maincat ~ month,value.var = var,fun.aggregate = sum)
    varwide[rowSums(varwide[,2:length(varwide)])>0]  # this is used three times. To make a function.
  } else
  {
    if(expenditure) var <- "tot_deb" else var <- "tot_crd"
    cols<- character(0)
    sumtalld <- statement[,.(tot_deb=sum(db,na.rm = T),tot_crd = sum(cr,na.rm=T),count=.N),
                          by=.(type,category,mnth)]
    sumtallc <- cashdt[,.(tot_deb = sum(db,na.rm = T),tot_crd=sum(cr,na.rm=T),count= .N),
                       by=.(type,category,mnth)]
    dttall <- rbind(sumtalld,sumtallc)
    dttall <- dttall[loadlov(),on=.(category=subcat)][!is.na(mnth)]
  
    if(tall) return(dttall)
    dtwide1 <- dttall %>% dcast(formula = category + type + maincat ~ mnth,fun.aggregate = sum,value.var = var) # new line added
    dtwide2 <-  dttall %>% dcast(formula = category + maincat ~ mnth,value.var = var,fun.aggregate = sum) # cash and digital transactions added for each category
    dtwide1 <- dtwide1[rowSums(dtwide1[,4:length(dtwide1)])>0] # filter out all zero rows
    dtwide2 <- dtwide2[rowSums(dtwide2[,3:length(dtwide2)])>0]
    if(collapse) dtwide2 else dtwide1
  }
}

# ---- merging of statements ------
# use this directly after downmaster()
merge_st <- function(dfile="bank_statement_dump2.xlsx",
                     cashdt= pettyc(),lovmer = loadlov()) {
  # Combined download are merged and new records in IDBI are enriched with bank tx id, bank name, and type (electronic)
  st_hdfc <- proc_xls_hdfc(dfile)$master[,bank:="HDFC"]
  st_idbi <- proc_xls_idbi(dfile)[,bank:="IDBI"]
  st_chq <- proc_xls_cheques(dfile)
  st_chq[,chq_in_clearance:=T]
  st_el <- rbind(st_hdfc,st_idbi,st_chq,fill=T)
  st_el <- st_el[,mnth:=crfact(cr_date)][ ,type:="electronic"][,Date:=as.Date(cr_date)]
  st_el[is.na(mode) & bank=="IDBI", mode:="CHEQUE"]
  st_el[is.na(mode) & chq_in_clearance==T, mode:="CHEQUE"]
  # cash transaction are loaded and merged (particulars are merged with narration of bank statements)
  cashdt <- cashdt[,mnth:=crfact(Date)][,type:="cash"][,mode:="CASH"][,Date:=as.Date(Date)] %>% setnames(c("part"),c("narr"))
  x1 <- rbind(cashdt,st_el,fill=T)[order(Date)]
  x2 <- lovmer[x1,on=.(subcat=category)]
  setnames(x2,c("subcat"),"category")
  x2[,.(bank,Date,cr,db,narr,chq,bank_tx_id,type,mnth,mode,category,maincat,chq_in_clearance)]
}


# this will split only st_master as of now. Pass the source month, amount and pattern of narr. ANd also pass the new values and a ratio of distribution of the combined amt.
split_entry <- function(origin,m,amt,pat="",newcat="same",newdesc="same",shiftto=31,db=T,electronic=T,ratio=c(1,1),verbose=F){
  if(electronic){
    if(db) var <- "db" else var <- "cr"
    targetrow <- origin[month(cr_date)==m & round(get(var),0)==amt & grepl(pat,narr,ig=T)]
    if(!verbose) assert_that(nrow(targetrow)==1) else{
      print(targetrow)
      assert_that(nrow(targetrow)==1)
    }
  newdt <- rbind(origin,targetrow)
  lastrown=nrow(newdt)
  newdt <- newdt[lastrown,category:=ifelse(newcat=="same",category,newcat)][lastrown, # update the new row with newcat and new amount
                                       db:=(ratio[2]/sum(ratio))*amt][lastrown,
                                                                      cr_date:=cr_date+ddays(shiftto)]
  newdt[month(cr_date)==m &  round(get(var),0)==amt & grepl(pat,narr,ig=T), # reduce the amount from the  orginal entry
        (var):=(ratio[1]/sum(ratio))*amt][,mnth:=crfact(cr_date,many = F)]
  
  newdt[lastrown,
        narr := ifelse(newdesc=="same",narr,newdesc)]
  
  
  }
  newdt
}

# plots
plotbal <- function(afterdate=20180925,dt=st_master,thresh=100000){
  dt[cr_date>ymd(afterdate),.(DATE=cr_date,BAL=balance)] %>% ggplot(aes(DATE,BAL)) + geom_line() + geom_smooth(se = F) + geom_hline(yintercept = thresh,lty=2,color="red")
}

plotcum <- function(month_range=7:10,yearval=2018,dt=st_master){
  copy(dt) -> stmast
  stmast[is.na(cr),cr:=0]
  stmast[is.na(db),db:=0]
  stmast_molt <-
  stmast[as.integer(mnth) %in% month_range & 
           !grepl("Sink|Fixed",category,ig=T) & 
           year(cr_date)==yearval,
         .(day=day(cr_date),
           income=cumsum(cr),expense=cumsum(db)),
         by=mnth] %>% 
    melt(value.name = "amount",id.vars=c("mnth","day"))
stmast_molt %>% 
  ggplot(aes(day,amount), group=mnth) + 
  geom_line(aes(color=variable),size=2) + 
  facet_grid(~mnth)
}  

plotexp <- function(dt=st_master,main=T){
  x <- mtrnd(statement = dt,tall = T)[!grepl("INC:",maincat)] %>% 
    ggplot(aes(mnth,tot_deb)) + 
    geom_col(aes(fill=type))
  if(main)
    x + facet_wrap(~maincat,scales = "free_y")
  else
    x + facet_wrap(~category,scales = "free_y")
    
}

#st_master <- readRDS("stmast.RDS")

cashflow <- function(st=st_master,lastn=6,cum=F){
  # income <- mtrnd(lastn = lastn,st,cashdt = pettyc(),maincat = T,expenditure = F) %>% .[,-c(1)] %>% colSums() # income totals 
  # expense <- mtrnd(lastn = lastn, st,cashdt = pettyc(),maincat = T,expenditure = T) %>% .[,-c(1)] %>% colSums() # expense totals
  # income_expense <- rbind(income,expense) %>% as.data.table(keep.rownames = T) %>% setnames("rn","type") %>% melt(id.vars = c("type"),variable.name="month")
  income_expense_cum <-
    mtrnd(st,lastn = lastn, tall = T)[,.(exp=sum(tot_deb,na.rm = T),inc=sum(tot_crd,na.rm = T)),by=mnth][order(mnth)]
  cum_net_operating <- 
    income_expense_cum[,cumexp:=cumsum(exp)][,cuminc:=cumsum(inc)]
  cum_net_operating %>% melt(id.vars="mnth")
}

big_trend <- function(st=st_master,n=12,scenario=1){
  trexp <- mtrnd(st,collapse = T,maincat = F,lastn = n)
  trenxptall <- melt(trexp,variable.factor = T,id.vars = c("category","maincat"),variable.name = "month",value.name = "amt",na.rm = T)
  copy(trenxptall) -> tr_gl
  copy(trenxptall) -> tr_el
  # create two scenarios with Major expenses - can create more
  tr_gl[grepl("GST",category),Major_Expense:="GST"]
  tr_gl[grepl("Lift",category),Major_Expense:="LIFT_AMC"]
  tr_gl[is.na(Major_Expense),Major_Expense:="ROUTINE"]
  
  tr_el[grepl("Electricity",category),Major_Expense:="ELECTR_CONS"]
  tr_el[is.na(Major_Expense),Major_Expense:="ROUTINE"]
  
  switch(scenario,tr_gl,tr_el)
}

#for GBM
payments <- function(st=st_master,topn=15,main=F,totm=12,perp_mths=11){
  x2 <- mtrnd(statement = st,lastn=totm,lovmer = fread("catlookup.csv"),collapse = T,maincat = F,tall = T)
  x2 <- x2[mnth<="Feb_2019"]
 if(main) x2[,{
          strt = min(mnth)
          last = max(mnth)
          p2p_mths = as.numeric(last) - as.numeric(strt) + 1
          tot = sum(tot_deb,na.rm = T)
          umths = mnth %>% unique %>%  length # unique months paid in
          avginv =  tot/.N
          rlife = tot/p2p_mths # monthly expense for the life of the product
          runiq = tot/umths # months expense assuming sporadic use
          rrecur = tot/p2p_mths
          r12 = tot/12
          .(total_spent= tot, strt_mth=strt,end_mth=last,life= p2p_mths, uniq_mths =umths, tot_tx=.N, avg_inv=avginv,perp_mthly= rrecur,prop_month=r12)
  },
  by=maincat][order(-total_spent)][1:topn] else
  {
    x2 <- mtrnd(statement = st,lastn=totm,lovmer = fread("catlookup.csv"),collapse = T,maincat = F,tall = T)
    x2 <- x2[mnth<="Feb_2019"]
    x2[,{
      strt = min(mnth)
      last = max(mnth)
      p2p_mths = as.numeric(last) - as.numeric(strt) + 1
      tot = sum(tot_deb,na.rm = T)
      umths = mnth %>% unique %>%  length # unique months paid in
      avginv =  tot/.N
      rlife = tot/p2p_mths # monthly expense for the life of the product
      runiq = tot/umths # months expense assuming sporadic use
      rrecur = tot/p2p_mths
      r12 = tot/12
      .(total_spent= tot, strt_mth=strt,end_mth=last,life= p2p_mths, uniq_mths =umths, tot_tx=.N, avg_inv=avginv,perp_mthly= rrecur,prop_month=r12)
    },
    by=category][order(-total_spent)][1:topn]
  }
}

#name is lev1 but it takes sumlev from 1 to 4
lev1 <- function(dig=x2,cash=cashdt,lov = lovmer,sumlev = 1){ 
  # 1 = month & maincat
  # 2 = month, maincat & type of payment
  # 3 = maincat and type
  # 4 = month and type
  
  sumrz <- function(dt,lev=1){
    month.main <-   dt[,.(sumcr=sum(cr,na.rm=T),sumdb = sum(db,na.rm = T),count= .N),by=.(maincat,month)]
    month.main.type <-   dt[,.(sumcr=sum(cr,na.rm=T),sumdb = sum(db,na.rm = T),count= .N),by=.(maincat,month,type)]
    main.type <-   dt[,.(sumcr=sum(cr,na.rm=T),sumdb = sum(db,na.rm = T),count= .N),by=.(maincat,type)]
    month.type <-   dt[,.(sumcr=sum(cr,na.rm=T),sumdb = sum(db,na.rm = T),count= .N),by=.(month,type)]
    switch(lev,month.main,month.main.type,main.type,month.type)
  }
  unique(lov) ->lov
  a <- dig[,type:="electronic"][lov, on=.(category=subcat),nomatch=0]
  b <- cash[,type:="cash"][lov,on=.(category=subcat),nomatch=0]
  setnames(b,"date","cr_date")
  a[, month:=crfact(cr_date,many = T)]
  b[, month:=crfact(cr_date,many = T)]
  
  #if(verbose==2) {print(a);print(b) }
  rbind(sumrz(a,lev = sumlev),sumrz(b,lev = sumlev))
}


#Note: there is one more crfact function that does not have many parameter.
# many parameter if TRUE is the usual long format else it is just Jan to Dec (no year suffixed)
# we have replaced many=T in most functions. If you find any many=F existing change it.

# divide payment records of the flats in separate worksheets of excel: try it for prabhu
divsheets <- function(flats=c(101,102),starting="Jan_2018",st=st_master){
  wb <- openxlsx::createWorkbook()
  for(i in flats) addWorksheet(wb,sheetName = as.character(i))
  for(i in flats){
    j<- paste0("^",i,"$|;",i,"$|;",i,";|","^",i,";")
    tx <- st[category %in% flatcats & grepl(j,flats) & crfact(cr_date)>=starting,-c("db","balance")]
    tx[,cr_date := format(cr_date,"%b %d, %Y")]
    tx[,chq_date := format(chq_date,"%b %d, %Y")]
    if(any(!is.na(tx$chq_date))) bycheque=T else bycheque=F # no longer needed
    openxlsx::writeDataTable(wb,sheet = as.character(i),x = tx,startRow = 2,tableStyle = "none")
    openxlsx::mergeCells(wb,sheet = as.character(i),cols = 1:10,rows = 1)
    openxlsx::writeData(wb,sheet = as.character(i),startRow = 1,x = paste("PAYMENTS RECEIVED FROM FLAT:",i))
    openxlsx::addStyle(wb,sheet = as.character(i),rows = 1,cols=1,
                       style = createStyle(
                         fgFill = "cadetblue1",
                          valign="center",
                         indent=5,
                         fontSize = 24
                       ) )
    openxlsx::setRowHeights(wb,sheet = as.character(i),rows = 1,heights = 40)
    openxlsx::setRowHeights(wb,sheet = as.character(i),rows = 2:30,heights = 18)
    #if(bycheque) setColWidths(wb,sheet = as.character(i),cols = c(1,3,4,6,7,8),widths = c(12,"auto",11,12,19,15)) else
      setColWidths(wb,sheet = as.character(i),cols = c(1:10),widths = c("auto"),ignoreMergedCells = T)
      setColWidths(wb,sheet = as.character(i),cols = c(2),widths = c(8))
      #setColWidths(wb,sheet = as.character(i),cols = c(1,3,4,6,7,8),widths = c(12,"auto",8,2,2,15))
  }
  saveWorkbook(wb,file = "flatwise_sheets.xlsx",overwrite = T)
}



hang_paym <- function(flat=1103,r2=reco2){
  x1 <- r2[flatn==flat]
  x1[,rn:=rownames(x1)]
  x1[,.(rn1=first(rn),flatn=first(flatn),sqft=first(sqft), received_amt=first(credit),narration=first(narr),credit_date=first(cr_date)),by=txid][
    x1,on=.(rn1=rn)][,.(flatn,received_amt,txid,credit_date,invoice,mnthfor,due_date,narration)]
}

# new 6 char acct ids to be allotted to common names: DONOT regenerate acct numbers if minor changes happen in names spellings
create_accounts <- function(rdu,regenerate=F){
  if(!regenerate){
    accts <- fread("acctids.csv")
    rdu1 <- rdu[accts,on="flatn"]
  } else {
  names_mat <- rdu[,str_split(oname," ",simplify = T,n=3)][,c(2:3)] %>% str_remove_all("\\.|Mrs|Mr") %>% matrix(ncol = 2)
  names_mat[,2] %<>%  str_remove_all("\\s")
  len1 <- nchar(names_mat[,1])
  len1 <- ifelse(len1 <3,len1,3)
  len2 <- 6 - len1
  acct <- paste0(str_sub(names_mat[,1],1,len1),str_sub(names_mat[,2],1,len2)) %>% toupper() # joint the first three char of first name and second name
  rdu1 <- cbind(rdu,acct)
  #rdu1[flatn==1502,acct:="AISAMI"] # so that 1501 and 1502 get the same Acct id.
  rdu1[flatn==404,acct:="FARSHA"] # so that 404 and 1402 get the same Acct id.
  }
  rdu1[,invoice_val:=5.05*1.18*sqft]
  return(rdu1[,.(flatn,oname,acct,invoice_val)][,.(Flats=paste(flatn,collapse=";"),Names=paste(unique(oname),collapse=";"), tot_inv=sum(invoice_val) %>% round(2)),by=.(acct)][order(acct)])
}

# We donot overwrite these files. Just read from them or modify them by hand
# last changed 1501 and 1502 : split them into two accounts.
load_acctids <- function() fread("acctids.csv")
load_acctinv <- function() fread("acct_invoice.csv")


acct_ids <- load_acctids()
acct_inv <- load_acctinv()
acct_grp <- acct_ids[,.(flats=paste(flatn,collapse=";")),by=acct][acct_inv,on="acct"]

# load acct ids from local file instead of regenerating them from rdu
join_accts <- function(st){
  st_split <- cSplit(st,splitCols = "flats",direction = "wide",sep = "[,;:]",type.convert = F,drop = F,fixed = F )
  st_split[,flat_status:=fifelse(!(is.na(flats) | nchar(flats)==0) & is.na(as.integer(flats_1)),"Error","Fine")] # add a flat string status Error flag
  if(st_split[flat_status=="Error",.N] >0) message ("Found flat string error at: ",st_split[flat_status=="Error",flats])
  st_split <- st_split[,flatn:=ifelse(!is.na(flats),as.integer(flats_1),NA)][!is.na(flatn)] # this line will throw warnings if flats is non numeric. To remind Kiran
  st_split <- st_split[acct_ids,on=.(flatn),nomatch=0]
  st2 <- st_split[acct_inv,on=.(acct),nomatch=0]
  st2 %>%  select(setdiff(names(st2),grep("flats_",names(st2),value = T)))
}

# new reco code works on account ids and cumulative payments (Note: maintenance category payments are all cumulatively added)
# load this in google sheet worksheet named reco: remember to pass the list of two DTs as outputted by proc_xls_hdfc()
reco_cum <- function(sthdfc=st_hdfc,ason=Sys.Date()){
  st1 <- sthdfc
  st2 <- join_accts(st1) 
  setDT(st2)
  st2[,mtce_amt:=ifelse(category=="Maintenance",cr,0)] 
  st2[,mnths_actd_for:=mtce_amt/tot_inv]
  st3 <- st2[,.(mtcepaidfor=formattable::digits(round(sum(mnths_actd_for,na.rm = T),2),2),
                totmtcepaid = accounting(sum(mtce_amt,na.rm = T),0)
                ),
             by=acct][acct_grp,on="acct",nomatch=0]
  st3[!is.na(flats),paidupto:=last(marrf(round(mtcepaidfor - 0.3,0))),by=rownames(st3)] # round up only if > 0.8
  st3[,excess_paid:=round(totmtcepaid - tot_inv*round(mtcepaidfor,0),0)]
  setnames(st3,qc(tot_inv),qc(monthly_invoice))
  setcolorder(st3,qc(acct,flats,monthly_invoice,mtcepaidfor,paidupto,excess_paid))
  return(st3)
}

# scan difference of payments from invoices : incomplete
scan_diffs <- function(st,last=1){
  st2 <- join_accts(st)
  
}

dues_as_on <- function(ldt,ason=Sys.Date()){
  yr <- year(ason)
  mth <- month(ason)
  nmths <- (yr - 2018) *12 + mth
  st1 <- rbind(ldt[[1]],ldt[[2]][,cr:=booked_cr],fill=T)[as.Date(cr_date)<=ason]
  st2 <- join_accts(st1) 
  st2[,mtce_amt:=ifelse(category=="Maintenance",ifelse(is.na(booked_cr),cr,booked_cr),NA)] # booked_cr is more accurate when maintenance amount + penalty is paid in one transaction
 
  st2[,penal_amt:=ifelse(category=="Penalty",ifelse(is.na(booked_cr),cr,booked_cr),NA)]
  st2[,mnths_actd_for:=mtce_amt/tot_inv]
  st3 <- st2[,.(mtcepaidfor=round(sum(mnths_actd_for,na.rm = T),1),
                totmtcepaid = round(sum(mtce_amt,na.rm = T),0)
                #totpen_paid=sum(penal_amt,na.rm = T)
  ),
  by=acct][acct_grp,on="acct",nomatch=0] # monthly invoice is binded
  setnames(st3,c("tot_inv"),c("mnth_invoice"))
  st3[!is.na(flats),paidupto:=last(marrf(mtcepaidfor)),by=rownames(st3)]
  st3[,shortf := (mnth_invoice * nmths - totmtcepaid)]
  #st3[,pen := ifelse(nmths - mtcepaidfor >=1, round(nmths - mtcepaidfor,0)*500, 0)]
  st3[,mi := round(nmths - mtcepaidfor - 1,digits = 0)]
  st4 <- st3[mi >=1, .(intt= first(mnth_invoice)*mi*(mi + 1) * 0.025/2),by=acct][st3,on="acct"]
  st4[,ason:=ason]
  st4
}

# enter either an acctid or one of the many flats for the acct; output is all payments with rect nbs.
acct_paym_record <- function(acctid=NULL,ldt,flat=NULL,filt="Maint|Penal|Party"){
  stopifnot(any(!is.null(acctid),!is.null(flat)))
  if(is.null(acctid)) acctid <- acct_ids[flatn==flat,acct]
  st <- ldt$master %>% 
    rbind(ldt$supp,fill=T) %>% 
    join_accts() %>% subset(acct==acctid) %>% 
    mutate(date=as.Date(cr_date)) %>%  
    mutate_at("cr",~ifelse(is.na(booked_cr),cr,booked_cr)) %>% 
    arrange(date,desc(cr)) %>% 
    filter(grepl(filt,category)) %>% 
    select(acct,flats,date,cr,category,mode,chq,rectnb) %>% 
    setDT
  st
}

rep_paym_html <- function(ldt,mnthno=22){
  reco_cum(ldt) -> reco1
  reco1[,oname:=getoname(flats)[1],by=acct] -> reco2
  reco1[,ophn:=paste(get_ocont(flats),collapse = ", "),by=acct] -> reco2
  css1 <- ifelse(reco1$mtcepaidfor < mnthno,"background: orange",NA)
  reco1[,oname:=paste(getoname(flatstr = flats),collapse=", "),by=acct] %>% 
    htmlTable(align="r",css.cell = rep(css1,ncol(reco1)) %>% matrix(ncol = ncol(reco1)))
}

# Hard coded date changes for rolled over large payments, to avoid distortions in reports
change_date <- function(st){
  st2<- copy(st)
  st2[grepl("Electricity",category) & Date==ymd(20180312),Date:=Date + ddays(20)]
  st2[grepl("Squad",category) & db==250000,Date:=Date - ddays(20)]
  st2[grepl("Electricity",category) & chq==320,Date:=Date + ddays(20)]
  st2[grepl("Electricity",category) & chq==320,Date:=Date + ddays(20)]
}

# pass the DT that is output from the merge_st() function
enrich_st <- function(st){
  st2<- change_date(st)
  st2[,mnth:=crfact(Date)]
  st2[,cal_hyr := crhyr(Date)]
  st2[,f_hyr := crfhalf(Date)]
  st2[,qtr := crqtr(Date)]
  st2[,fy := crfy(Date)]
  st2 <- st2[order(Date)]
  day_incomes <- unique(st2$Date) %>% map_dbl(~dayinc(st2,.x))
  day_maint_coll <- unique(st2$Date) %>% map_dbl(~day_mcoll(st2,.x))
  day_coll <- unique(st2$Date) %>% map_dbl(~day_coll(st2,.x))
  day_expenses <-  unique(st2$Date) %>% map_dbl(~dayexp(st2,.x))
  netincr <- unique(st2$Date) %>% map_dbl(~incr(st2,.x))
  inc_dt <- data.table(udate=unique(st2$Date),netincr=netincr,today_exp = day_expenses,today_inc = day_incomes)
  inc_dt[,cbal:=cumsum(netincr)]
  inc_dt[,cexp:=cumsum(today_exp)]
  inc_dt[,cinc:=cumsum(today_inc)]
  inc_dt[,cmaint:=cumsum(day_maint_coll)]
  inc_dt[,cumcol:=cumsum(day_coll)]
  inc_dt[,cumsav:=cinc - cexp]
  st2[inc_dt,on=.(Date=udate)]
}

# calculates net increase in balance for a specific date 
incr <- function(st,dat1){
  st[Date==dat1, sum(cr,-db,na.rm = T)]
}

# net expense on a day
dayexp <- function(st,dat2){
  st[Date==dat2 & !grepl("Sinking|Transfer|Fixed|Invest|Withdrawal|Petty",category),sum(db,na.rm = T)]
}

# day income - all incomes
dayinc <- function(st,dat2){
  st[Date==dat2 & !grepl("Sinking|Transfer|Fixed|Redeem|Withdrawal|Petty",category),sum(cr,na.rm = T)]
}

# day maintenance collection
day_mcoll <- function(st,dat3){
  st[Date==dat3 & category=="Maintenance",sum(cr,na.rm = T)]
}

day_coll <- function(st,dat3){
  st[Date==dat3 & grepl("Maintenance|Collection|Interior|Shift|Penalt|Party",category),sum(cr,na.rm = T)]
}

# modify transfers to IDBI with salary amount : this is an approximation
adjust_idbi_transfers <- function(st){
  st[grepl("Transfer Out",category),category:="Salary_transfer"]
  st[grepl("Salary_transfer",category),db:=150000]
}

# filter only the expense or income transactions
rem_transfers <- function(st){
  st[!grepl("FD|Transfer|Sink",category)]
}

# donot use it for tab separated text downloads, use pull_idb_text()
pull_idbi_csv <- function(file= "~/Downloads/embassy.csv"){
  conv2n <- function(x) str_remove_all(x, "[,CDr]") %>% str_trim %>% as.double()
  x1 <- fread(file,encoding = "Latin-1")
  x2 <- x1[V1!=""]
  x3 <- x2[grepl("2019|2020",V1),cr_date:=str_trim(V1) %>% dmy()][!is.na(cr_date)]
  setnames(x3,c("V3","V4","V5","V6","V7"),c("chq","db","cr","balance","narr"))
  x3[,db := conv2n(db)]
  x3[,cr := conv2n(cr)]
  x3[,chq := textclean::replace_non_ascii(chq)]
  x3[,narr := textclean::replace_non_ascii(narr)]
  x3[,balance := conv2n(balance)]
  x4 <- x3[,rn:=as.numeric(rownames(x3))][order(-rn)]
  x4[,.(cr_date,db,cr,chq,narr,balance)]
}

# if idbi stmt is sent in doc format by Dasara it may be decoded by this function into a datatable
pull_idbi_doc <- function(textfile="~/Downloads/idbi_stmt.txt",gstempl=T){
  emb <- readr::read_file(textfile)
  idbraw <- emb %>% str_split("\r\n") %>% unlist %>% textclean::drop_element("^$")
  x1 <- idbraw %>% str_detect("Tran. Date") %>% dplyr::cumany() %>% idbraw[.]
  colnames <-  x1[1] %>% str_split("\t",simplify = T) %>% as.character()
  x2 <- split(x1[-1],c(1,2))
  dt1 <- paste(x2$`1` ,x2$`2`,sep = "\t") %>% data.table(x=.) %>% cSplit("x",sep = "\t",stripWhite = F)
  setnames(dt1,colnames %>% abbreviate() %>% tolower)
  dt2 <- dt1[-1,.(wthd,dpst,blnc)] %>% map(~str_replace_all(.x,",","") %>%  str_extract("\\d+")) %>% as.data.table() %>% cbind(dt1[-1],.)
  dt2 <- dt2[,-c(which(duplicated(colnames(dt2),fromLast = T))),with=F]
  dt2[,tr.d := as.Date(tr.d,"%d-%m-%Y")]
  dt2[,vldt := as.Date(vldt,"%d-%m-%Y")]
  dt2[,c.n. := str_trim(c.n.)]
  dt2[,nrrt := str_trim(nrrt)]
  dt2[,wthd := as.numeric(wthd)]
  dt2[,dpst := as.numeric(dpst)]
  dt2[,blnc := as.numeric(blnc)]
  dt3 <- dt2[,.(cr_date=tr.d,db=wthd,cr=dpst,chq=c.n.,narr=nrrt,balance=blnc)][order(cr_date)]
  if(gstempl) dt3 else dt2[order(tr.d)]
}

# if the statement is downloaded as tab separated, then this works (modified for new IDBI website)
pull_idbi_txt <- function(textfile="idbi/idbi_dec.txt",gstempl=T){
  emb <- readr::read_file(textfile)
  idbraw <- emb %>% str_split("\n") %>% unlist %>% textclean::drop_element("^$")
  x1 <- idbraw %>% str_detect("Txn Posted Date") %>% dplyr::cumany() %>% idbraw[.]
  dt1 <- data.table(x1[3:(length(x1) - 1)])
  dt2 <- cSplit(dt1,"V1",type.convert = "as.character",sep = "\t",stripWhite = F) # stripWhite=F is essential to retain blank columns
  dt2[,cr_date:=dmy(V1_01,tz = "Asia/Kolkata")]
  newcols <- dt2[,.(amt=V1_11,balance=V1_13)] %>% map(~str_replace_all(.x,",","") %>%  str_extract("\\d+")) %>% map(as.numeric) %>%  as.data.table()
  setnames(dt2,c( "V1_07", "V1_05","V1_09"),c("chq",  "narr","type"))
  dt2[,narr:=str_trim(narr)]
  dt2[,chq:=str_trim(chq)]
  dt2[,type:=str_trim(type)]
  dt3 <- cbind(dt2,newcols)[order(cr_date)]
  dt3[,db:=ifelse(type=="Dr.",amt,NA_real_)]
  dt3[,cr:=ifelse(type=="Cr.",amt,NA_real_)][,.(cr_date,cr,db,narr,balance,chq)]
} # usually we use htmlTable() to output the html table in viewer and copy paste into google sheet as simpler option

# assumes all .txt files containing IDBI monthly statements are lying in a directory; processes all files
pull_all_idbi <- function(dir="idbi"){
  x1 <- list.files(dir,full.names = T) %>% map(pull_idbi_txt) %>% rbindlist()
  x2 <- x1[order(cr_date)]
  x2[,balprev:=shift(balance)]
  x2[,cbal:=ifelse(.I>1,sum(balprev, cr, -db,na.rm = T),balance),by=.(cr_date,chq,cr,db)]
  if(x2[abs(cbal - balance) > 1,.N] > 0) {
    message("Possible missing or unordered transactions")
    x2[,bal_incorrect:=ifelse(abs(cbal - balance) > 1,TRUE,FALSE)]
    return(x2)
  }
  x2[,-c("balprev","cbal")][order(cr_date,chq)]
}
# Input is enriched & merged statement  - output of merge_st() and enrich_st()
# Output is a DT filtered with vendor payments only
vendor_payments <- function(ste){
  ste[is.na(cr) & !grepl("Transf|Sinking|Withdra|Petty|Salary",category),.(Date,mnth,db,narr,category,type,bank)]
}

#convert tally dump of CSV into a data.table of all transactions
proc_tally <- function(dumpfile="cash balance.csv",startrow=11){
  cbal <- fread(dumpfile)
  cbal[startrow:nrow(cbal)] -> cbal2
  cbal2[,V4:=NULL]
  names(cbal2) <- qc(Date,dir,prt,Date2,vchtype,vchno,db,cr)
  cbal2[,valid:=ifelse(nchar(Date)>0,1,0)]
  cbal2[,rownb:=cumsum(valid)]
  split(cbal2$prt,f = cbal2$rownb) -> prtfull
  prtfull %>% map(paste,collapse=";") %>% unlist -> prt2
  cbal2[valid==1,fullprt:=prt2]
  cbal3 <- cbal2[valid==1]
  cbal3[,cr:=as.numeric(cr)]
  cbal3[,db:=as.numeric(db)]
  cbal3[,Date:=dmy(Date)]
}

match_payments <- function(st){
  invdt <- rdu[,.(flatn,sqft,invoice=sqft*5.05*1.18)]
  st[category=="Maintenance"] %>% 
    cSplit("flats",sep = ";",direction = "tall") %>% 
    invdt[.,on=.(flatn=flats)] %>% 
    .[,.(InvoiceAMt=sum(invoice),flats=paste(flatn,collapse = ";")),by=.(cr_date,chq,narr,rectnb,Credited=cr)]
}

lastpaid <- function(st=st_master){
  lpaid <- st[category=="Maintenance"] %>% 
    cSplit("flats",sep = ";",direction = "tall") 
  lpaid[,.(lastpaid=last(cr_date),Amt=last(cr),mode=last(mode),ref=last(chq)),by=.(flatn=flats)]
}

pull_idbi_pdf <- function(datestr="31-12-2020",path="idbi",filename=NULL){
    ftonum <- function(x) x %>% as.character %>% str_remove_all(",") %>% as.numeric()
  if(datestr=="ALL"){
  files <- list.files(path = path,pattern = "OpTrans.*pdf$",full.names = T)
  stopifnot(length(files)>0)
  read_one_file <- function(fname){
    pdftools::pdf_text(fname) -> x1
    x1[[1]] %>% str_split("\n") %>% unlist %>% str_trim -> rawdata
    dt <- rawdata %>% 
      str_subset("/2020") %>% 
      str_subset("^Tran",negate = T) %>% 
      data.table(x=.) %>% 
      cSplit(splitCols = "x",sep = "   ")
    dt[!is.na(x_9)] %>% setnames(qc(sn,trdate,vdate,narr,chq,cred,ccy,amt,balance)) -> dta
    dt[is.na(x_9)] %>% setnames(qc(sn,trdate,vdate,narr,cred,ccy,amt,balance,dummy)) -> dtb
    dt <- rbind(dta,dtb,fill=T)
    dt[,dummy:=NULL]
    dt[,cr_date:=parse_date_time(trdate,orders = "dmyHMS",tz = "Asia/Kolkata")]
    dt[,cr:=ifelse(grepl("Cr",cred),ftonum(amt),NA)]
    dt[,db:=ifelse(grepl("Dr",cred),ftonum(amt),NA)]
    dt[,.(cr_date,cr,db,narr,balance=ftonum(balance),chq)]
  }
  
  merged_dt <- files %>% map(read_one_file) %>% rbindlist() %>% unique
  merged_dt[order(cr_date)]
  } else 
   {
     if(is.null(filename)) filename <-  sprintf("%s/OpTransactionHistoryUX5%s.pdf",path,datestr) else
       filename <- sprintf("%s/%s",path,filename)
     nstat1 <- pdftools::pdf_text(filename)[[1]] %>% str_split("\n") %>% .[[1]] %>% str_trim %>% str_detect("^Statement") %>% which()
     # if the line with Statement is not found on page 1 or goes to next page we will throw an error
     if(length(nstat1)==0)
       if(pdftools::pdf_info(filename)$pages>1){
         nstat2 <- 
           pdftools::pdf_text(filename)[[2]] %>% 
           str_split("\n") %>% .[[1]] %>% 
           str_trim %>% 
           str_detect("^Statement") %>% which() 
       }
     else 
       stop("Seems like there is a problem as there is only one page without the Statement Summary text found")
     #assertthat::assert_that((length(nstat1) + length(nstat2)) >0,msg = "Seems the IDBI first 2 pages donot contain the Statement summary... Try lesser date range")
     if(exists("nstat2")) {
       bottn1 <- 753 
      bottn2 <- 60 # picked for jul file - to be converted to linear formula
     } else
    bottn1 <- nstat1*15 + 255 # had to derive through linear algebra over several page sizes
    
    if(exists("bottn2")){ 
      bothpages <- 
        tabulizer::extract_tables(file = filename,pages = 1:2,area = list(c(294,38,bottn1,609),c(1.25,43,bottn2,612))) # this will return a list of pages
      x1 <- bothpages %>% map(~as.data.table(.x)) %>% rbindlist() %>% row_to_names(1)
    }
    else
     x1 <- tabulizer::extract_tables(file = filename,area = list(c(294,38,bottn1,609)))[[1]] %>% 
       as.data.table() %>% row_to_names(1)
    
      

if(grepl("srl",names(x1)[1],ig=T) & grepl("Balance",names(x1)[9],ig=T)){
  setnames(x1, qc(sn,trdate,vdate,narr,chq,cred,ccy,amt,balance))
  setDT(x1)
  x1[,cr_date:=parse_date_time(trdate,orders = "dmyHMS",tz = "Asia/Kolkata")]
  x1[,cr:=ifelse(grepl("Cr",cred),ftonum(amt),NA)]
  x1[,db:=ifelse(grepl("Dr",cred),ftonum(amt),NA)]
  x1[,.(cr_date,cr,db,narr,balance=ftonum(balance),chq)][order(cr_date)]
} else
      stop("The columns read in did not match to the 9 columns:",sprintf("idbi/OpTransactionHistoryUX5%s",datestr))
  } 
    
}

# use this on the csv file downloaded using googledrive::drive_download()
pull_hdfcmaster <- function(file="st_hdfc.csv"){
  x1 <- fread(file)
  x1[,cr_date:=parse_date_time(cr_date,orders = c("ymdHMS","mdyHMS"),tz = "Asia/Kolkata")]
  x2 <- x1[,c(1:12)]
  x2 %<>% map_at(.at = c(2,3,5),~ str_remove_all(.x,",") %>% str_extract("\\d+") %>% as.numeric)  
  x2 %<>% map_at(.at = c(7,10,11,12), as.factor)
  as.data.table(x2)
}

npv_diff <- function(flat,inv,daysbehind=0,dt=st_hdfc,dayrate=0.02/30){
  cutoffdate <- as.Date(Sys.time() - ddays(daysbehind))
  dt <- dt[cr_date>=ymd(20200101) & cr_date <= cutoffdate & grepl(flat,flats) & category=="Maintenance"]
  dt_dues <- data.table(dates=ymd(sprintf("2020%02d%s",1:9,"01")),amt=inv) # hard coded from Jan 2020
  dt_dues <-  dt_dues[dates <= cutoffdate]
  N <- as.numeric(cutoffdate - ymd(20200101)) # hard codes Jan 1 2020
  dt_dues[,days:=  -as.numeric(first(dates) - dates) + 1]
  cf_recd <- dt$cr
  day_int <- dt[, N - as.numeric(as.integer(cutoffdate - as.Date(cr_date))) + 1]
  npv_expected <- FinancialMath::NPV(cf0 = 0,times =dt_dues$days ,i = 0.02/30,cf = dt_dues$amt)
  npv_actual <- FinancialMath::NPV(cf0 = 0,i = dayrate,cf = cf_recd,times=day_int)
  totdue <- dt_dues[,sum(amt)]
  totrecd <- sum(cf_recd)
  list(day_credit=day_int,day_due=dt_dues$days,  cf=cf_recd, total_due=totdue, recd=totrecd, gap=  totdue - totrecd, gapnpv =  npv_expected - npv_actual)
}

# macro report - month on month starting from  bank statements rbind of HDFC and IDBI 
macro_report <- function(n=36)
st_ehbanks[year(cr_date)>=2018] %>% 
  mutate(month=crfact(cr_date)) %>% 
  group_by(month) %>% 
  dplyr::summarise(maint_coll=sum(ifelse(grepl("Mainten",category),cr,0)),
                   mthly_exp=sum(ifelse(!grepl("Trans|FD",category,ig=T),db,0),na.rm = T),
                   other_earnings=sum(ifelse(!grepl("Mainten|Trans|FD|Sink",category,ig=T) & !grepl("auto_redemption",narr),cr,0),na.rm = T)) %>% 
  mutate(cumcoll=cumsum(maint_coll),cum_invoice=cumsum(rep(total_due,n))) %>% 
  mutate(mtcesaving=  maint_coll - mthly_exp, prov_savings=cum_invoice - cumsum(mthly_exp)) %>% 
  mutate(savings_realised=cumsum(mtcesaving),cumdues=cum_invoice - cumcoll,roll_exp6=frollmean(mthly_exp,6), roll_coll6=frollmean(maint_coll,6))
  


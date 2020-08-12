library(splitstackshape)
library(stringr)
library(wrapr)
library(knitr)
library(kableExtra)
library(magrittr)
library(wrapr)
library(lubridate)
library(markdown)
library(rmarkdown)
library(data.table)
library(purrr)
library(googlesheets)



# ---- these functions return a vector of owner name or contact numbers; input can be either flat number or acctid.-----
getoname <- function(flatstr,acct_id=NA,res=rdu){
  if(is.na(acct_id)) {
    flts <- strsplit(as.character(flatstr),";") %>% unlist %>%  str_trim %>% unique
    res[flatn %in% flts, strsplit(oname,";") %>% unlist %>%  str_trim %>% unique]
  } else 
  {
    flts <- acct_grp[acct %in% acct_id,flats] %>% strsplit(.,";") %>% unlist %>%  str_trim %>% unique
    res[flatn %in% flts, .(oname=strsplit(oname,";")) %>% unlist %>%  str_trim %>% unique]
  }
}

# can return more than one element vector
get_ocont <- function(flatstr,acct_id=NA,res=rdu){
  if(is.na(acct_id)) {
  flts <- strsplit(as.character(flatstr),";") %>% unlist %>% str_trim %>% unique
    res[flatn %in% flts,strsplit(ophn,";")  %>% unlist %>% str_trim %>% unique]
  } else
  {
    flts <- acct_grp[acct %in% acct_id,flats] %>% strsplit(.,";")  %>% unlist %>% str_trim %>% unique
    res[flatn %in% flts,strsplit(ophn,";") %>% unlist %>% str_trim %>% unique]
  }
}

# Group the call always 'by account'. DONOT input a vector of accounts or flats, without grouping.
get_oemail <- function(flatstr,one_email=F,acct_id=NA,res=rdu){
  if(is.na(acct_id)) {
  flts <- strsplit(as.character(flatstr),";") %>% unlist %>% str_trim %>% unique
    emails <- res[flatn %in% flts,strsplit(oemail,";")  %>% unlist %>% str_trim %>% unique]
  } else
  {
    flts <- acct_grp[acct %in% acct_id,flats] %>% strsplit(.,";")  %>% unlist %>% str_trim %>% unique
    emails <- res[flatn %in% flts,strsplit(ophn,";") %>% unlist %>% str_trim %>% unique]
  }
  if(one_email) emails[1] else emails
}

# tenant contact
get_tcont <- function(flatstr,acct_id=NA,res=rdu){
  if(is.na(acct_id)) {
 flts <- strsplit(as.character(flatstr),";")  %>% unlist %>% str_trim %>% unique
    res[flatn %in% flts,.(tphn=strsplit(tphn,";") %>% unlist %>% str_trim %>% unique)]$tphn
} else
{
  flts <- acct_grp[acct %in% acct_id,flats] %>% strsplit(.,";") %>% unlist %>% str_trim %>% unique
  res[flatn %in% flts,.(tphn=strsplit(tphn,";")  %>% unlist %>% str_trim %>% unique )]$tphn
}
}

# tenant name (returns NA if no tenant)
get_tname <- function(flatstr,acct_id=NA,res=rdu){
  if(is.na(acct_id)) {
    flts <- strsplit(as.character(flatstr),";")  %>% unlist %>% str_trim %>% unique
    res[flatn %in% flts,strsplit(tname,";")  %>% unlist %>% str_trim %>% unique]
  } else 
  {
    flts <- acct_grp[acct %in% acct_id,flats] %>% strsplit(.,";")  %>% unlist %>% str_trim %>% unique
    res[flatn %in% flts,strsplit(tname,";")  %>% unlist %>% str_trim %>% unique ]
  }
}

identify_email <- function(email,res=rdu){
  stopifnot(is.character(email),grepl("@",email),length(email)==1)
  ox <- res[str_detect(oemail,email),.N]
  tx <- res[str_detect(temails,email),.N]
  if(tx>1) message("We have found the email of a tenant present in more than one flat, which is not fine.")
  if (ox > 0) res[grepl(email,oemail),paste(flatn,collapse = ",")] else
    if (tx > 0) res[grepl(email,temails),paste(flatn,collapse = ",")] else
      NA_integer_
    
}
#----- end of RDU functions ------



# these are creating pdf files -----
cr_pdf_greceipt <- function(dt,rown,rmd){ # rmd is rmd file name
  rmarkdown::render(input = rmd,intermediates_dir = "~/Downloads",
                    output_format = "pdf_document",
                    output_file = paste0("receipt_", dt[rown,paste(acct,rectnb,sep = "_")], ".pdf"), 
                    output_dir = "./receipts")
}

# NEW: email html body creation by looping render() on all accounts - preferred over latex version
cr_html_bodies <- function(ldt=stmast,rmdfile="cover2.Rmd",acct="ALL",curmnth=21,duedate="Oct 5, 2019",tilldt=Sys.Date(),penfile = "penalty.txt") {
  reco1 <- reco_cum(ldt)
  penalt <- load_penalty(infile = penfile,download = F)[acct_ids,on="flatn"]
  if(length(acct)==1 && acct=="ALL") accts_to_process <- acct_ids$acct else
    accts_to_process <- acct
    for(a in accts_to_process){
      rmarkdown::render(rmdfile,
                        params = list(acct=a,
                                      table=acct_paym_record(acctid = a,ldt = ldt),
                                      acctval = reco1[acct==a,monthly_invoice],
                                      curm = curmnth,
                                      duedate = duedate,
                                      pen = penalt[a==acct,latepen][1], # as some accts have multiple flats
                                      intt = penalt[a==acct,intt][1]
                                      ),
                        output_file = paste(a,"html",sep = "."),
                        output_dir = "html_body")
    }
}

cr_pdf_invoice<- function(dt,rown,rmd,mnth,yr){ # rmd is rmd file name
  rmarkdown::render(input = rmd,
                    output_format = "pdf_document",
                    output_file = paste0(dt[rown,flatn], ".pdf"),
                    output_dir = crfactm(mnth,yr) %>% paste0("./",.,"_invoices"))
}

cr_pdf_invoice_party<- function(dt,rown,rmd){ # rmd is rmd file name
  rmarkdown::render(input = rmd,
                    output_format = "pdf_document",
                    output_file = paste0(dt[rown,flatn],"_",dt[rown,date_of_use], ".pdf"),
                    output_dir = "./common_area/")
}

# ---- end of pdf file generators ----

# change the SOA sheet name and ensur it has the same headings as last time
# save this output in penalty.txt and then run create_invoices
down_soa <- function(tit="Invoice Statement For July 19",download=T){
  if(download) gs_download(gs_title(tit),ws = 1,to = "temp.csv",overwrite = T)
  x1 <- fread("temp.csv")
  x1 <- x1[8:63,c(3,8:10,12)]
  setnames(x1,c("flatn","latepen", "intt", "det","rem_priv"))
  x1[,latepen:= latepen %<>% str_remove(",") %>% as.numeric()]
  x1[,intt:= intt %<>% as.numeric()]
  x1[!is.na(latepen)]
}

load_penalty <- function(infile="penalty.txt",download=F){
  if(is.null(infile)) return(data.table(0))
  if(download==T) {
    x1 <- gs_read(gs_key(key_rudra_master),ws = "penalty",col_types = c("iddcc_")) %>% setDT
    setnames(x1,c("flatn","latepen", "intt", "det","rem_pub"))
    x1[,det:=str_replace_all(det,"&","and")]
    x1[,"rem_pub":=str_replace_all(rem_pub,"&","and")]
    fwrite(x1,infile,dateTimeAs = "write.csv")
    } else{
    x1 <-fread(infile,colClasses =c("integer","double","double","character","character"))
    setnames(x1,c("flatn","latepen", "intt", "det","rem_pub"))
    x1[,det:=str_replace_all(det,"&","and")] # to keep tex from crashing as & is a special char
    x1[,"rem_pub":=str_replace_all(rem_pub,"&","and")]
    }
  return(x1)
}

load_bounce <- function(infile="chq_bounce.txt",download=F){
  if(download==T){
    x1 <- gs_read(gs_key(key_rudra_master),ws = "bounce",col_types = c("icdccdc")) %>% setDT
    datecols <- names(x1)[4:5]
    x1[,(datecols) := map(.SD, ~lubridate::parse_date_time(.x, orders = c("ymd", "dmy", "mdy"))),.SDcols=datecols]
    setnames(x1,c("flatn","chqno",  "chq_amt","depon","bouncedon", "bcharges", "rem"))
    fwrite(x1,infile,dateTimeAs = "write.csv")
  } else
    x1 <-  fread(infile,colClasses =c("integer","character","double","character","character","double","character"))
  return(x1)
}

# Main function for creating monthly invoices
# Pls ensure st_master is updated with all payment tagging and penalty file is uptodate.
create_invoices <- function(for_mnth=5, goforflats="invoice.txt", penalty_data= "penalty.txt", rdu=fread("rdu.txt"),
                            chq_bounce_data="chq_bounce.txt", st=st_master,foryear=2020,rmd="invoice1.Rmd",savesummary=T,cutfiles=T,upload=F){
  cat(sprintf("Latest bank record that was tagged found for:%s",st[!is.na(category),last(cr_date)]))
  fact_mnth <- crfactm(for_mnth,foryear)
  #reco2 <- reconcile(sumlev = 2)[,.(flatn,cr_date,chq,credit,mode,mnthfor,due_date,delay_days,late_penalty,interest)]
  #rdu <- fread(flats_data) # ensure the latest googlesheet is downloaded
  penalty <- load_penalty(penalty_data,download = F)
  cbounced <- load_bounce()
  penalty[,intt:=intt/1.18]
  penalty[,latepen:=latepen/1.18]
  cbounced[,bcharges:=bcharges/1.18]
  #reco2[late_penalty>0,details:=paste0("payment for ",mnthfor," made on:", cr_date," (",delay_days," days later from due date ",due_date,")")]
  #penalt <- reco2[,.(sum_pen=sum(late_penalty),sum_int=sum(interest)),by=flatn]
  data <- penalty[rdu[,.(flatn,sqft,oname,oemail,ophn)],on="flatn"]
  data <- cbounced[data,on="flatn"]
  data[,formonth:=fact_mnth]
  data[,depon:=as.Date(depon)]
  data[,bouncedon:=as.Date(bouncedon)]
  data[,base_amount:=as.numeric(sqft*5.05)] # this is the base rate - hard coded
  data[,amount:=ifelse(is.na(bcharges),base_amount,base_amount+bcharges)]
  data[,amount:=ifelse(is.na(latepen),amount,amount+latepen)]
  data[,amount:=ifelse(is.na(intt),amount,amount+intt)]
  data[,cgst:=0.09*amount]
  data[,sgst:=0.09*amount]
  data[,total:=amount + cgst + sgst]
  data <- lastpaid()[data,on="flatn"]
  cols <- qc(intt,latepen,bcharges)
  data[,(cols):=map(.SD,~ifelse(is.na(.x),0,.x)),.SDcols=cols] # replacing NA with 0 neat trick in data.table
  if(is.numeric(goforflats)) sel_flats <- goforflats else
    if(goforflats=="ALL") sel_flats <- FLATS else 
      sel_flats <- fread(goforflats)$V1
  data[bcharges>0,bounceref:=paste("chq:",chqno,"amt:",chq_amt,"deposited:",depon,"bounced:", bouncedon),by=flatn]
  data <- data[flatn %in% sel_flats]
  cols<- qc(base_amount,amount,cgst,sgst,latepen,intt,bcharges,total)
  data[,(cols):=lapply(.SD,round,digits=2),.SDcols=cols] # can use this with map also now
  data[is.na(rem_pub),rem_pub:=""]
  summ<- data[,.(flatn,formonth,sqft,Owner=oname,base_amount,late_penalty=latepen,Interest=intt,Chq.bounce=bcharges,Bouceref=bounceref,Late_mnths=det,cgst,sgst,Total=total )]
  file_to_write <- sprintf("invoices_%s_%02d.csv",foryear,for_mnth)
  if(upload | savesummary) fwrite(summ,file = file_to_write)
  if(upload) gs_upload(file = file_to_write,overwrite = T)
    if(cutfiles) seq_len(nrow(data)) %>% walk(~cr_pdf_invoice(data,.x,rmd,for_mnth,foryear)) else # rmd is the rmd file name
    {
      cols <- qc(base_amount,Total)
      summ[,(cols):=lapply(.SD,format,big.mark=","),.SDcols=cols]
      summ
    }
}

create_party_invoices <- function(goforflats="party.txt",rmd="invoice_party.Rmd"){
  data <- rdu[,.(flatn,oname,tname,oemail,temails)]
  charges <- fread("party.txt",sep = ";",skip = 1)
  setnames(charges,c("flatn","date_of_use","type_facility","amount","flag")) # input file must be 5 columns
  charges[,date_of_use:=parse_date_time(date_of_use,orders = "mdy") %>% as.Date]
  data <- data[charges,on="flatn"]
  data[,cgst:=0.09*amount]
  data[,sgst:=0.09*amount]
  data[,total:=amount + cgst + sgst]
  data <- data[flag >0]
  seq_len(nrow(data)) %>% walk(~cr_pdf_invoice_party(data,.x,rmd)) # rmd is the rmd file name
}


#  working...
create_gen_receipts <- function(lstmnt,starting="2019-08-01",flatnbs="ALL", selaccts="ALL",ownernames="ALL", rect_range=1:10000,rmd="receipts_general.Rmd", sumfile="receipts_summary.csv",regenerate=F){
  strt <- as.Date(starting)
  st_acct <- join_accts(lstmnt[[1]])
  st_subset <- st_acct[as.Date(cr_date)>=as.Date(starting)]
  if(length(flatnbs)==1 && flatnbs !="ALL" || length(flatnbs) > 1) st_subset <- st_subset[grepl(paste(flatnbs,collapse = "|"),flats)] # remember this may overselect a few flats as this is a crude pattern match of flat number
  if(selaccts !="ALL") st_subset <- st_subset[acct %in% selaccts] 
  st_acct_det <- st_subset[acct_grp,on="acct",nomatch=0]
  st_enriched <- st_acct_det[,owner:=ifelse(ownernames=="ALL", getoname(flatn) %>% broman::vec2string(), getoname(flatn)[1]),by=flatn]
  cumacct <- reco_cum(lstmnt)
  dt <- st_enriched[cumacct,on="acct",nomatch=0]
  dt <- dt[str_extract(rectnb,"\\d+") %in% rect_range,.(cr_date,flats,acct,owner,category,paidupto,cr,booked_cr,rectnb,chq,mode,chq_date)][order(cr_date)]
  if(regenerate) seq_len(nrow(dt)) %>% walk(~cr_pdf_greceipt(dt,.x,rmd)) # rmd is receipt file name
  fwrite(dt,file=sumfile,dateTimeAs = "write.csv")
}


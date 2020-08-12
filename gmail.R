library(lubridate)
library(dplyr)
library(purrr)
library(splitstackshape)
library(htmlTable)
library(magrittr)
library(wrapr)
library(stringr)
library(data.table)
#library(readr)
library(formattable)
library(blastula)

# pick files from dir and post them as email bodies to the acctids
email_html_files <- function(dir="html_body",acctids="ALL"){
  allfiles <- list.files(dir,pattern = "html$",full.names = T)
  acct_grp[,sflat:=str_extract(flats,pattern = "\\d+")] # single flat
  #acct_grp[,oname:=getoname(sflat) %>% broman::vec2string(),by=acct]
  acct_grp[,To:=get_oemail(sflat,one_email = T),by=acct]
  dt1 <- data.table(files=allfiles)
  dt1[,acct:=str_extract(files,"[A-Z_]+(?=.html)")]
  dt1[,body:=readr::read_file(files),by=acct]
  dt2 <- acct_grp[dt1,on="acct"]
  dt2[,From:="Embassy Haven"]
  dt2[,Subject:=sprintf("Complete accounts and dues for %s [%s]",acct,flats)]
  dt3 <- dt2[!is.na(To)]
  dt4 <- if(acctids[1]!="ALL") dt3[acct %in% acctids] else dt3 # clever way to allow acctids vector in the same argument 
  emails <- dt4[,.(To,From,Subject,body)] %>% pmap(create_html_email)
  emails %>% map(gm_create_draft)
}

email_rmd_single <- function(file,emails=c("sanjmeh@gmail.com")){
    emime1 <- create_html_email(To=emails,From="Me",Subject = "Test",body = readr::read_file(file))
    gm_create_draft(emime1)
}
# new email code using blastula ready now
# blastula emails - no draft they are sent straight
# the only glitch is the column names of penalty.txt could be not matching - make them robust next
email_invoices <- function(flatlist="flats_to_email.txt",
                       cover="cover_invoice_blastula.Rmd",
                       dir_invoices="Apr_2020_invoices/",
                       this_mth = "April 2020", 
                       due_date = "Apr 15, 2020",
                       penl="penalty.txt",
                       exclude_flats=c(604))
{
    ALLFLATS<- setdiff(FLATS,exclude_flats)
    sendfunc <- function(x,y,z,...){
    smtp_send(email = x,to = y, from = "embassy.haven@gmail.com",subject = paste(this_mth," invoice for flat ",z),credentials = creds_file("blastula_gmail_creds"))
      cat(y,":GONE.")
  }
  
  if(is.numeric(flatlist)) sel_flats <- flatlist else 
    if(flatlist=="ALL") sel_flats <- ALLFLATS else sel_flats <- fread(flatlist)$V1
    
  my_dat <- 
    rdu[flatn %in% sel_flats,.(oname,oemail,flatn)] %>% 
    cSplit("oname",sep = ";",direction = "wide",type.convert = "as.character") %>% 
    cSplit("oemail",sep = ";",direction = "wide",type.convert = "as.character")
  my_dat[,fnames:=paste0(dir_invoices,flatn,".pdf")]
  #pen <- fread(penl,colClasses =c("integer","double","double","character","character"))
  pen <- load_penalty(infile = penl)
  if (nrow(pen)>0)  my_dat <- pen[my_dat,on=.(flatn)]
  
  # f1 <- plyr::catcolwise(as.character)
  # mydat2 <- f1(my_dat)
  #mydat3 <- cbind(my_dat[,setdiff(names(my_dat),names(mydat2)),with=F],mydat2)
  if(nrow(pen)>0)
  setcolorder(my_dat,qc(fnames,flatn,oname_1,rem_pub)) else # bringng fnames in beginning so that length of the DT does not matter
    setcolorder(my_dat,qc(fnames,flatn,oname_1)) 
  emails <- my_dat %>%  pmap(.f =  ~render_email(cover, render_options = list(params=list(mnth=this_mth,name=..3,flatn=..2,duedate=due_date,note=..4))) %>% add_attachment(file = ..1))
  my_dat[,rend_email:=list(emails)]
  setcolorder(my_dat,qc(rend_email,oemail_1,flatn))
  my_dat %>%  pwalk(.f = ~sendfunc(..1,..2,..3)) 
}

# save the output of this to file "acctwise.csv" and then manually delete / edit before running the next function 
rep_paid <- function(lstmast){
  flat_records <- acct_ids[,email1:=get_oemail(flatstr = flatn)[1],by=acct]
  acct_emails <- flat_records[,.(acct,email1)] %>% unique()
  reco1 <- reco_cum(lstmast)[acct_emails,on="acct"]
  reco1
}

###### 
# send emails informing advance paid and thanking.
# this email goes to all emails separated with |
# before running this func run reco_cum and filter on mtcepaidfor > N months
#####
email_advance_thanks <- function(file_to_read="acctwise.csv",
                       cover="cover_advance_blastula.Rmd",
                       include_only = NULL)
{
    # this sendfunc has just two parameters : email object and email address; subject is hard coded
    sendfunc <- function(x,y,...){
    smtp_send(email = x,to = y, from = "E H A O A",subject = paste("Acknowledgment of your advance payments"),credentials = creds_file("blastula_gmail_creds"))
      cat(y,":GONE.")
    }
    
    allrecs <- fread(file_to_read)
    if(is.null(include_only)) selrecs <- allrecs else selrecs <- allrecs[grepl(include_only,flats)]
 
  selrecs[,emailvect := list(str_split(emails,pattern = fixed("|"))),by=flats]  
  setcolorder(selrecs,qc(acct,flats,paidupto)) # bringing the 3 params inthe starting for passing them to pmap
  email_objects <- selrecs %>%  pmap(.f =  ~render_email(cover, render_options = list(params=list(acct=..1,flat = ..2, paidupto = ..3))))
  selrecs[,email_obj:=list(email_objects)]
  setcolorder(selrecs,c("email_obj","emailvect")) # this second setcolorder is essentially for sendfunc() pwalk sequence
  selrecs %>%  pwalk(.f = ~sendfunc(..1,..2)) 
}
  

  

# new blastula html emails attaching receipts
email_receipts <- function(rand = 0, indx=1:100,file_receipts="receipts_summary.csv",cover="cover_receipt_blastula.Rmd",dir_name="receipts",firstemail=T){
  
  rcpt <- fread(file_receipts)
  indx <- seq_len(nrow(rcpt)) %in% indx # restricting the processing to few rows of the receipt file
  rcpt <- rcpt[indx]
  rcpt[,cr_date:=parse_date_time(cr_date,orders = "ymdHMS")]
  rcpt[,chq_date:=parse_date_time(cr_date,orders = "ymdHMS")]
  acct_emails <-
    if(firstemail)
    acct_ids[,.(oemail=get_oemail(flatn,one_email = T)),by=acct][!is.na(oemail)] else
      acct_ids[,.(oemail=get_oemail(flatn,one_email = F)),by=acct][!is.na(oemail)]
  rcpt <- rcpt[acct_emails,on="acct",nomatch=0]  
  filenames <- list.files(dir_name,pattern = "pdf$",full.names = T)
  acct_in_fnames <- str_extract(filenames,"[A-Z]+(?=_)")
  id_receipts <- str_extract(filenames,"[A-Z]+_\\d+")
  dt_ids = data.table(f=filenames,ids=id_receipts)
  rcpt[,rids:=paste(acct,rectnb,sep = "_")]
  rcpt <- rcpt[dt_ids,on=.(rids = ids),nomatch=0][order(cr_date)]
  
  if(rand>0) usedt <- rcpt[sample(seq_len(nrow(rcpt)),rand)] else usedt <- rcpt
 
  edat <- 
   usedt %>% 
    mutate(
      name = owner,
      amt = cr,
      date_recd =  format(cr_date,"%b %d, %Y"),
      files=f
    ) %>%
    select(name, acct, date_recd, amt, files, oemail) %>% 
    setDT
  
  emails <-
    edat %>%  pmap(.f =  ~render_email(cover, render_options = list(params=list(name=..1,acct = ..2, Date=..3,amt = ..4))) %>% add_attachment(file = ..5))
  
  edat[,rend_email:=list(emails)]
  setcolorder(edat,qc(rend_email,oemail,date_recd))
  # the send function
  sendfunc <- function(x,y,z,...){
    smtp_send(email = x,to = y, 
              from = "EHAOA",
              subject = paste("Receipt for your payment made on",z),
              credentials = creds_file("blastula_gmail_creds"))
  }
  
  edat %>%  pwalk(.f = ~sendfunc(..1,..2,..3)) 
}

# Sends a fixed invite subject and fixed attachments in a personalised cover note to all
# USed for AGM invite
# acct_grp must exist as a variable to start with.
email_invite <- function(cover="cover_invite.Rmd",path="agm",flats_included="ALL"){
  f5 <- list.files(path = path,pattern = "pdf$",full.names = T)
  add_allattach <- function(e1,att=f5) {
    for (i in att) e1 <- add_attachment(e1,i)
    e1
  }
  edat <- acct_grp[,fflat:=str_split(flats,";")[[1]][1],by=acct
                   ][,emails:= list(list(get_oemail(fflat,one_email = F))),by=acct
                     ][rdu[,.(oname,flatn=as.character(flatn))],on=.(fflat=flatn)
                       ][!is.na(emails) & !is.na(acct)
                         ][,oname:=str_replace_all(oname,";",",")]
  if(flats_included[1]!="ALL") edat <- edat[fflat %in% flats_included]
  setcolorder(edat,qc(oname,flats))
  #edat <-   rdu[,invemails:=list(str_split(oemail,pattern = ";"))][,invnames:=str_replace_all(oname,";",",")][,.(flatn,invnames,invemails)][flatn == 504]
  email_objects <- 
  edat %>% 
     pmap(.f =  ~render_email(cover, render_options = list(params=list(name =..1,flatn = ..2))) %>% add_allattach) 
  
  edat[,rend_email:=list(email_objects)]
  setcolorder(edat,qc(rend_email,emails))
  
  # the send function
  sendfunc <- function(x,y,...){
    smtp_send(email = x,
              to = y, 
              from = "E H A O A",
              subject = paste("AGM scheduled for March 15, 2020"),
              credentials = creds_file("blastula_gmail_creds"))
  }
  edat %>%  pwalk(.f = ~sendfunc(..1,..2)) 
}

# returns a DT with emails
# to be rewritten with blastula
draft_email_contacts <- function(dt=rdu,allflats=F,send=F){
  if(allflats){
   flats_to_email <- FLATS 
  } else {
  file_flats="flats2.txt"
  flats_to_email <- fread(file_flats)$V1
  email_list <- dt$oemail %>% str_split(";") %>% unlist %>% str_trim
  }
  my_dat <- dt[flatn %in% flats_to_email,.(oname,oemail,ophn,flatn)] %>% cSplit("oemail",sep = ";",direction = "tall",type.convert = "character")
  my_dat[,oname:=str_remove(oname,",")]
  emailgrp <- my_dat[,.(phones=str_split(ophn,";") %>% unlist %>% str_trim %>% unique %>% paste(collapse = "; "),flats=paste(unique(flatn),collapse = "; ")),by=.(oemail,oname)]
  flatgrp <- my_dat[,.(emails=paste(unique(oemail),collapse = "; ")),by=.(flatn,oname)]
  emailgrp <- emailgrp[flatgrp[,.(oname,emails)],on="oname"] %>% unique
  emailgrp %<>% cSplit("emails",sep = ";",direction = "wide")
  emailgrp[,
        To := sprintf('%s <%s>', paste0("OWNER-", str_split(flats,";") %>% unlist %>% str_trim %>% paste(collapse = ".")), oemail),by=oemail
    ][, From := "EMBASSY HAVEN"
    ][, Subject := sprintf('Contact details verification for flat owner: %s', oname)
    ]
  emailgrp[phones=="NA",phones:="Not Available. Please send us your contact numbers urgently."]
  emailgrp[is.na(emails_2),emails_2:="Not Available. Consider providing us with a backup email id"]
  dt2 <- emailgrp[,{
    stylematrix <- matrix(rep("",12),nrow = 6)
    stylematrix[2:6,2] <- rep("font-weight: 900;",5)
    if(grepl("Not",emails_2)) stylematrix[4,2] <- "font-weight: 900; color:CornflowerBlue;"
    if(grepl("Not",phones)) stylematrix[5,2] <- "font-weight: 900; color:Crimson;"
    if(nchar(phones)<=11) {
      phones <- paste(phones,"(Consider providing us with a backup phone number)") 
      stylematrix[5,2] <- "font-weight: 900; color:CornflowerBlue;"
    }
    footnote<- "This is an auto generated message."
    mtable = list(Flat=paste(flats,collapse = ";"), Primary_email=emails_1,Backup_email=emails_2,Phone=paste(unique(phones),collapse = ";"),Legal_owner=oname)
    mtable2 <- mtable %>% as.data.table %>% t %>% as.data.table(keep.rownames = T)
    htable <- htmlTable(mtable2,
      header = c("Attribute","Record with EHAOA "),
      align = "l",align.header = "l",
      css.table = "border-collapse: collapse;",
      caption = "Kindly send in your edits/additions, by reply email.",
      tfoot = footnote,
      css.cell = stylematrix)
    .(To=To,From=From,Subject=Subject,body=htable)
  }, by=.(oemail,oname)]
 emails <-  dt2[!is.na(oemail)] %>% pmap(~create_html_email(To = ..3,From = ..4,Subject = ..5,body = ..6)) 
   if (send==T) gm_send_message(emails) else gm_create_draft(emails) 
}

# move this to blastula 
email_payment_hist <- function(start="Jan_2018", allcols=T, send=F,allflats=F,st=st_master,file_flats="flats_to_email.txt",dt=rdu){
  flats_to_email <- if(allflats) FLATS else fread(file_flats)$V1
  edat <- rdu[flatn %in% flats_to_email] %>% 
    cSplit("oemail",direction = "wide",sep = ";") %>% 
    .[,.(flatn,To=oemail_1,owner=oname,From="EHAOA")]
  
  edat[, Subject := sprintf('Payment record of flat no. %s, since %s',flatn,start)]
  
  dt2 <- if(allcols)
                edat[,{
    footnote<- sprintf("This is an auto generated email for flat number %d, sent on email id %s", flatn,To)
    caption = sprintf("These are all the payment transactions credited to EHAOA since %s from flat %d. Kindly let us know if there is any error/omission",start,flatn)
    stylemat <- matrix("font-weight: normal;",ncol = ncol(reco2),nrow = reco2[flatn==.BY[[2]],.N])
    rows_with_ONLINE <- {reco2[flatn==.BY[[2]],mode] == "ONLINE"} %>% which
    rows_with_delay <- {reco2[flatn==.BY[[2]],delay_days] > 7} %>% which
    stylemat[1,] <- rep("word-wrap: break-word;",ncol(reco2))
    stylemat[,c(2,9)] <- "font-weight: bold;"
    stylemat[rows_with_ONLINE,7] <- "color:green;"
    stylemat[rows_with_delay,9] <- "color:red; font-weight:bold;"
    #stylemat[,7] <- "color:Tomato; font-weight:900"
    htable <- htmlTable(reco2[flatn==.BY[[2]]], # BY[[2]] can be replaced by flatn
                        align = "lrrlrrrrrrr",align.header = "lrrlrrrrrrr",
                        header = qc(Flat_no,`Date of credit`,`Maintenance month`,`Due Date`, `Chq/UPI no.`, `Amount Paid (Rs.)`, Mode, `Receipt no.`, `No. of days delayed`, `Late payment penalty`, Interest),
                        css.table = "border: 1px solid black;",
                        col.columns = c("lightyellow1","seashell2"),
                        caption = caption,
                        css.cell = stylemat,
                        tfoot = footnote)
    .(From=From,Subject=Subject,body=htable)
  }, by=.(To,flatn)] else
    edat[,{
      footnote<- sprintf("This is an auto generated email for flat number %d, sent on email id %s", flatn,To)
      caption = sprintf("These are all the payment transactions credited to EHAOA since %s from flat %d. Kindly let us know if there is any error/omission",start,flatn)
      stylemat <- matrix("font-weight: normal;",ncol = 8,nrow = reco2[flatn==.BY[[2]],.N])
      rows_with_ONLINE <- {reco2[flatn==.BY[[2]],mode] == "ONLINE"} %>% which
      #rows_with_delay <- {reco2[flatn==.BY[[2]],delay_days] > 7} %>% which
      stylemat[1,] <- rep("word-wrap: break-word;",8)
      stylemat[,c(2)] <- "font-weight: bold;"
      stylemat[rows_with_ONLINE,7] <- "color:green;"
      #stylemat[rows_with_delay,9] <- "color:red; font-weight:bold;"
      #stylemat[,7] <- "color:Tomato; font-weight:900"
      htable <- htmlTable(reco2[flatn==.BY[[2]],c(1:8)], # BY[[2]] can be replaced by flatn
                          align = "lrrlrrrr",align.header = "lrrlrrrr",
                          header = qc(Flat_no,`Date of credit`,`Maintenance month`,`Due Date`, `Chq/UPI no.`, `Amount Paid (Rs.)`, Mode, `Receipt no.`),
                          css.table = "border: 1px solid black;",
                          col.columns = c("lightyellow","seashell"),
                          caption = caption,
                          css.cell = stylemat,
                          tfoot = footnote)
      .(From=From,Subject=Subject,body=htable)
    }, by=.(To,flatn)]
  
  dt2 <- dt2[!is.na(To)]
  
  emails <- dt2[,.(To,From,Subject,body)] %>%
    pmap(create_html_email) # will not work with new gmailR package.. yet to change
  
  if(send==T) emails %>% walk(gm_send_message) else
    emails %>% walk(gm_create_draft)
}

# will send / draft emails of a certan year, and months given to only those flats who have not paid in the month.
# move this to blastula
email_reminder <- function(lstmast,currnt=crfact(Sys.Date)){
  reco <- reco_cum(ldt = lstmast)
  edat <- rdu[penddt,on="flatn"] %>% cSplit("oemail",direction = "wide",sep = ";") %>% .[,.(flatn,To=oemail_1,owner=oname,From="EHAOA",mths)]
  edat[,Subject:= sprintf('reminder for paying overdue flat maintenance charges (flat no: %s) for the month(s) of %s', flatn,mths)]
  edat[,body:= sprintf( readr::read_file("reminder.txt"), owner,flatn,mths)]
  emails <- edat[,.(To,Subject,From,body)] %>%
    pmap(create_text_email)
  
  if(send==T) emails %>% walk(gm_send_message) else
    emails %>% walk(gm_create_draft)
}



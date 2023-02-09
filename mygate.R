library(wrapr)
library(dplyr)
#library(tabulizer)
library(DescTools)
library(bit64)
library(knitr)
library(readxl)
library(openxlsx)
library(data.table)
library(stringr)
library(splitstackshape)
library(magrittr)
library(lubridate)
library(ggplot2)
library(googledrive)
library(googlesheets4)
library(htmlTable)
library(assertthat)
library(purrr)
library(cellranger)
#pacman::p_load(kableExtra)
library(RColorBrewer)
library(janitor)

#if(dir.exists("~/Dropbox/bank-statements")) source("~/Dropbox/bank-statements/globalfns.r") else
  if(dir.exists("~/R/bank-statements")) source("~/R/bank-statements/globalfns.r")
#if(search() %>% str_detect("keys.data") %>% any() %>% not) attach("keys.data")
#if(search() %>% str_detect("mygate.keydata") %>% any() %>% not) attach("mygate.keydata")
embassy_str <- "Tech|Plum|Electr|House|Gard|STP"
squad_str <- "Superv|Lady|Guard|Security"

#fpath <- normalizePath(dirname(path = "."))
# if(grepl("Dropbox",fpath)) okpath <- grepl("RWAdata$",fpath) else
#   if(grepl("rstudio",fpath)) okpath <- grepl("embassy$",fpath)
# assert_that(okpath,msg = "You are not in the correct directory")
curyr <- 2020
curmth <- month(now())
# if(!exists("duty_matrix")) duty_matrix <- readRDS("dutyhrs.RDS")
# if(!exists("vislog")) vislog <- readRDS("disk_vislog.RDS")

#ptrn <- namewords %>% paste0(collapse = "|")

options(scipen = 99,digits = 2)

# download a fresh copy of resident_data_upload in rdu.txt. Read it with fread("rdu.txt") when needed
down_rdu <- function(key=key_rdu){ #google sheet resident_data_upload_mygate
    rdu <- as.data.table( read_sheet(ss = key,sheet = "master"))
    names(rdu) <- c( "blk" ,    "flatn",   "oname",   "tname" ,  "oemail",  "temails", "ophn" ,   "tphn",    "status" ,"sqft","soldby","app" )
    rdu[,oemail:=str_replace(oemail,"[^a-z]$",replacement = "")]
    rdu[,oemail:=str_replace_all(oemail,"^\\W| ","")]
    rdu[,ophn:=as.character(ophn)]
    rdu[,tphn:=as.character(tphn)]
    fwrite(rdu,"rdu.txt")
    base::message("New value of rdu saved in `rdu`")
    rdu <<- rdu[1:56]
}

simp_colnames <- function(dt){
  setDT(dt)
  setnames(dt, c( "blk" ,    "flatn",   "oname",   "tname" ,  "oemail",  "temails", "ophn" ,   "tphn",    "status" ,"sqft","soldby","app" ))
}

# upload cleaned up rdu back to googlesheet
# review this and try using write_sheet()
up_rdu <- function(dt,upl=F){
  dt2 <- copy(dt)
  dt2 <- tidyr::replace_na(dt2,(list(oemail="",tname="",temails="",tphn="",ophn="",app="")))
  drdu <- read_excel("resident-data-upload-mygate.xlsx")
  setnames(dt2,names(drdu))
  if(upl==T) gs_edit_cells(gs_key(key_rdu),ws = 1,anchor = "A1",input = dt2) else
  dt2
}

# if(!exists("rdu")) rdu <- fread("rdu.txt")
# if(!exists("flats")){
#   flats <- fread("flats.csv")
#   setnames(flats, c( "Society", "Flat" ,   "flat_type", "occ_type",  "Status",    "p_intc",    "s_intc",    "members"   ) )
# }
# FLATS = rdu$flatn

# new function to clash emails from three list : rdu, google group and mygate
# try with tenant=F to get full email clash with mygate
comp_emails <- function(havenfile="haven-residents.csv",mygatefile="resident_details (1).csv", tenant=F){
  .x1 <- fread(mygatefile)
  setnames(.x1,qc(society,flatn,name,utype,status,mob,email,pass,created))
  .x1$created %<>% dmy
  .x1$name %<>% str_trim
  .x1[,flatn:=str_extract(flatn,"\\d+") %>% as.integer()]
  mgate<- .x1[,-c("society")]
  ggrp <- fread(havenfile)[,.(email=`Email address`)]


  rdu2 <- rdu %>% cSplit(splitCols = "temails",sep = ";",direction = "tall",type.convert = "character") %>% .[!is.na(temails)] %>% setDT
  tmissed_in_google <- setdiff(rdu2$temails,ggrp$email)
  cat("\nFollowing emails of tenants are missed in haven google groups, present in google sheet rdu:\n")
  print(rdu2[temails %in% tmissed_in_google,.(flatn,tname,temails,tphn)])

  rdu3 <- rdu %>% cSplit(splitCols = "oemail",sep = ";",direction = "tall",type.convert = "character") %>% .[!is.na(oemail)] %>% setDT
  omissed_in_google <- setdiff(rdu3$oemail,ggrp$email)
  cat("\nOwner email present in rdu but but missing in google group:\n")
  print(rdu3[oemail %in% omissed_in_google,.(flatn,oname,oemail,ophn)])

  missed_in_rdu <- setdiff(ggrp$email,c(rdu2$temails,rdu3$oemails))

  cat("Present in google group but but missing in rdu:\n")
  print(missed_in_rdu)
}


# replace NA in place of 0 across all columns
replNA <- function(DT){ # matt dowle technique on SO here https://stackoverflow.com/a/7249454/1972786
  for (j in seq_len(ncol(DT)))
    set(DT,i = which(is.na(DT[[j]])),j,value = 0)
}

# replace 0 in place of NA across all columns
replZero <- function(DT){ # matt dowle technique on SO here https://stackoverflow.com/a/7249454/1972786
  for (j in seq_len(ncol(DT)))
    set(DT,i = which(DT[[j]]==0),j,value = NA)
}

# reads one csv file
read_vislog <- function(mygatefile="visaug.csv",tall=T,cols=8) {
  if(cols==8) {vislog = fread(mygatefile,select = 1:8)
  setnames(vislog, old = c("name","mob","type","entry","exit","flats","from","vehicle"))
  }
  else {
    vislog = fread(mygatefile)
    setnames(vislog,  c("soc","name","mob","type","subtype","flats", "status","vehicle", "entry","exit",c("X1","X2")) )
  }

  vislog[name!="Name"] -> vislog
  # if(mygatefile=="visjan.csv") # visjan.csv has different date format
  #   vislog[,entry:=ymd_hms(entry,tz = "Asia/Kolkata")][,exit:=ymd_hms(exit,tz = "Asia/Kolkata")][,stay:=as.duration(exit-entry)][,stayhrs:=as.numeric(stay/3600)] else
      vislog[,entry:=dmy_hms(entry,tz = "Asia/Kolkata")][,exit:=dmy_hms(exit,tz = "Asia/Kolkata")][,stay:=as.duration(exit-entry)][,stayhrs:=as.numeric(stay/3600)]

  vislog[,mob:=as.numeric(mob)]
  n_names_trailsp <- vislog[grepl(".+ $",name),name] %>% uniqueN()
  if(n_names_trailsp>0) {
    cat(paste( "..corrected",n_names_trailsp,"names with trailing spaces while importing",mygatefile,"\n"))
    vislog[grepl(".+ $",name),name:=str_remove(name," $")] # remove trailing spaces from name
  }
  vislog[,flats:= str_extract_all("\\d+",string = flats)]
  #vislog <- vislog[,flats:=.(str_extract_all("\\d+",string = flats))]
  compress <- function(x) paste(x,collapse = ";")
  vislog$flats <- sapply(X = vislog$flats,FUN = compress)
  vislog[flats=="",flats:=NA]
  vislog[,stayhrs:=round(stayhrs,3)]
  vislog[,weekday:=weekdays(entry)][,weekno :=week(entry)]
  vislog[,shift:=ifelse(hour(entry) %in% c(18:23,0) & stayhrs<=18, "NIGHT",
                        ifelse(hour(entry) %in% c(18:23,0) & stayhrs>18, "NIGHT+DAY",
                               ifelse(hour(entry) %in% c(7:11) & stayhrs>18, "DAY+NIGHT",
                                      ifelse(hour(entry) %in% c(7:11) & stayhrs<=18, "DAY",
                                             "ERR")
                               )
                        )
  )
  ]

  # test to be removed
  #vislog[grepl("RAME",name) & crfact(entry)=="Oct_2019" & grepl("Secu",type),hour(entry)]

  vislong <- cSplit(vislog,splitCols = "flats",sep = ";",direction = "long")
  if(tall) vislong else vislog
}

# read all csv files in one datatable & correct few names
read_allcsv <- function(names="^vis...\\.csv",path=".",save=T){
  files <- list.files(pattern = names,path = path,full.names = T)
  x1 <- files %>%  map(.f = ~read_vislog(mygatefile=.x)) %>% rbindlist()
  # Donot use purrr::map_dfr() as datatable copy is created instead of in place change.
  # map_dfr() created DT will generate warnings whenever you create a new column... as a shallow copy of the DT is created
  vislog <- correct_names(dt = x1,corr = "name_corr.csv") %>% .[order(entry)]# this will not only correct a few names but output an unique DT
  saveRDS(vislog,"disk_vislog.RDS")
  vislog
}

# modified to incorporate attendance.csv DT output as well. Just remember to change the namecol and typecol.
correct_names <- function(dt=vislog,corr="name_corr.csv",namecol="name",typecol="type"){
  dtchanges <- fread(corr)
  x1 <- dtchanges[dt,on=c(oldname=namecol,type=typecol),nomatch=0][,.(oldname,newname,type)] %>% unique
  x2 <- x1[dt,on=c("oldname"=namecol,"type"=typecol)]
  x3 <- x2[,name:=ifelse(is.na(newname),oldname,newname)][,c("oldname","newname"):=NULL] %>% unique
  # setcolorder(x3,neworder = names(vislog))
  x3
}

# guard attrition calculations
# have to use vislog with old data since inception
# now these 2 functions just return new / exited guards. Also the start date is Aug 2018.
marknew <- function(data=vislog,vis="Security|Guard"){
  data <- data[grepl(vis,type,ig=T)][,ndays := (year(entry)-2018)*365 + yday(entry)]
  data2 <- data[,
       {
         allnames <- data[ndays < .BY,unique(name) ] # cool trick
         NEW= ifelse(name %in% allnames,F,T)
         .(name=name,NEW=NEW,First_entry=entry,type=type)
       },
       by=ndays]
  unique(data2,by="name")[ndays>first(ndays)+10]
}

markexit <- function(data=vislog,vis="Security|Guard"){
  data <- data[grepl(vis,type,ig=T)][,ndays := (year(entry)-2018)*365 + yday(entry)]
  data2 <- data[,
       {
         allnames <- data[ndays > .BY,unique(name) ] # same trick, direction of inequality reveresed
         Attrited= ifelse(name %in% allnames,F,T)
         .(name=name,Attrited=Attrited,Last_entry=entry,type=type)
       },
       by=ndays]
   unique(data2[Attrited==T],by="name")[ndays<last(ndays) ]
}

# May want to update these codes as more guards leave us
#exit_codes <- fread("guards_exit_code.txt")



# a report that prints out stay period of all guards since inception
# have to use vislog with old data since inception
pr_html_guards <- function(){
  base::message("Following Guards attrited:")
  attr_guards <- marknew()[,-c("type")][markexit(),on="name"][is.na(ndays),ndays:=244]
  attr_guards[,days_stayed:=i.ndays - ndays]
  attr_guards2 <- attr_guards[exit_codes,on="name"]
  attr_guards <- exit_codes[attr_guards,on="name"]
  attr_guards[code==0,status:=NA_character_]
  attr_guards[code==1,status:="LEFT ON HIS OWN"]
  attr_guards[code==2,status:="SACKED"]
  all_guards <- vislog[grepl(squad_string,type,ig=T),.(join_date=first(entry)),by=.(name)]
  old_guards <- attr_guards[all_guards,on=.(name)][,age:=now() - join_date][is.na(Attrited)][age>ddays(365)]
  print(attr_guards)
  base::message("Following Guards have crossed one year. Send them a bouquet")
  print(old_guards)
  all_guards <- attr_guards[all_guards,on="name"][is.na(status),age:=now() - join_date][,.(name,Attrited,days_stayed,type,join_date,status,age)]
  all_guards %>%
    mutate(joining_date=format(join_date,"%b %d, %Y"), age = round(age)) %>%
    dplyr::select(name,Attrited,days_stayed,type,joining_date,status,age) %>%
    htmlTable(align=c("lllllrr"))
}

# --- SECURITY -----
#is used to show daily attendance - but should not be added for monthly. For monthly use calcattendnce()
# this is the rule based attendance


# reduce the usage of analyze except for SD and mean/median analysis; prefer to use show_gaps() as more versatile
analyse <- function(mnth=10,DT=vislog,emp_flat=F,emp_rwa=F,emp_security=F,flat_interaction=T,curday=day(now())){
  cat("\nStarted with ",nrow(DT)," records")
  assert_that(nrow(DT[month(entry)==mnth])>0,msg = "Zero rows found")
  if(!month(now()) == mnth) curday <- days_in_month(mnth)
  flag <- which(c(emp_flat,emp_rwa,emp_security))
  assert_that(length(flag)==1,msg = "Select exactly one, and only employee category.")
  dt1 <-  # take the right subset string between the three : flat emps, society emp, or security
    switch (flag,
            DT[!is.na(flats)],
            DT[grepl(embassy_string,type,ig=T)],
            DT[grepl(squad_string,type,ig=T)]
    )
  cat("\nSubsetted to ",nrow(dt1)," records containing relevant type of visitors")
  cat("\nGrouping on date basis to compute total stay times...")
  dt2 <- dt1[month(entry) %in% mnth,
     .(tstay=sum(stay) %>% as.duration,fentry=hour(min((entry))),lexit=hour(max((exit))),
       trips=.N,nightstay=ifelse(max(day(exit)>min(day(entry)) & max(hour(exit)>=3)),T,F)),
     by=.(name=str_trim(name),type,mob,day=day(entry),weekday=weekdays(entry),mnth=month(entry),flats)] # dt1 is summarised with grouping on day of month.
  cat("..DONE\n")


  cat(paste("\nStarted calculating unbiased means for ",ifelse(emp_flat,"Flat Employee data",ifelse(emp_rwa,"RWA employee", "Visitor data")),".. "))
  base::message(nrow(dt2)," records\n")

  output3 <- dt2[,{
                  stay <- .SD$tstay %>% as.duration()
                  trips <- .SD$trips
                  fentry <- .SD$fentry
                  lexit <- .SD$lexit
                  tyexit <- dt2[day!=.BY[[3]] & name == .BY[[1]] & mob==.BY[[2]],mean(lexit,na.rm = T)]
                  meanstay2 <- dt2[day!=.BY[[3]] & name == .BY[[1]] & mob==.BY[[2]],mean(tstay,na.rm = T) %>% as.duration() %>% round(0)]
                  Ndays <- dt2[name==.BY[[1]] & mob== .BY[[2]],.N]
    .(flats=flats,type=type,stay=stay,mean=meanstay2,trips=trips,
      first_entry=fentry,last_exit=lexit, typical_exit_hour=tyexit,ndays=Ndays)
  },
  by=.(name,mob,date=day)
  ]
  output3
}

find_multiple_cat <- function(dt=vislog,mnth=8,count=F){
  dt[mob!=9999999999 & month(entry) %in% mnth, {
      mobs <- .SD[,.(name,type,mob)] %>% unique %>% .[,.N,by=mob] %>% .[N>1]
      ss <- .SD[mob %in% mobs$mob,.(name=paste(unique(name),collapse = ";"),
                                    type=paste(unique(type),collapse = ";"),
                                    flats=paste(unique(flats),collapse = ";")),
                by=mob] %>% unique
      if(count) x=nrow(ss) else ss[order(mob)]
  }
  ]
}

# pass latest month; output names of visitors with same mobile no.
mult_cat_latest <- function(m=5){
  x1 <- seq_len(m) %>% map_dfr(~find_multiple_cat(mnth = .x,count = F),.id = "mnth")
  x1[,.(count=.N,name=paste(unique(name),collapse =";"),
        type=paste(unique(type),collapse=";"),
        mnth=paste(unique(mnth),collapse=";"),
        flats=paste(unique(flats),collapse=";")
        ),
     by=mob][order(-count)][grepl(as.character(m),mnth)]

}

absence <- function(mnth=9,categ="driver|maid",flatwise=F,tillday=day(now()),dt = vislog){
  if(!month(now())==mnth) tillday <- days_in_month(mnth)
  dt1 <- dt[month(entry)==mnth & grepl(categ,type,ig=T)]
  assert_that(nrow(dt1)>0,msg = "ZERO rows for the month found")
  dt2 <- dt1[,.(days_present=list(unique(day(entry))),alldays=list(1:tillday)),by=.(name,flats,type)]
  absent_days <- dt2$days_present %>% map2(.x = dt2$alldays,.y=.,.f = setdiff)
  #absent_counts <- absent_days %>% map_int(.f = length) %>% as.data.table(x = .) %>% setnames("A")
  absent_counts <- lapply(absent_days, length) %>% unlist
  absent_days2 <- setDT(list(absent_days))
  dt3 <- cbind(dt2,data.table(absent_days),A=absent_counts)
  dt3[,alldays:=NULL][,P:=uniqueN(unlist(days_present)),by=name]
  if(flatwise) dt3[!is.na(flats)][order(flats,type,name)] else dt3[,flats:=NULL][order(type,name)]
}

# make plots
mnthf <- factor(month.abb,levels = month.abb,ordered = T)

# visitor counts and stay
visplot <- function(m="May_2019",categ=c("Driver"),stay=T){
  dailycnt <- vislog[!is.na(flats) & crfact(entry)>=m & type %in% categ,.(cnt=.N,avgstay=mean(stayhrs,na.rm = T)),by=.(type,month=crfact(entry),day(entry))]
  ndlycnt <- vislog[!is.na(flats) & crfact(entry)>=m & type %in% categ,.(cnt=.N,avgstay=mean(stayhrs,na.rm = T),tstay=sum(stayhrs,na.rm = T)),by=.(name,month=crfact(entry),day(entry))]
  #dailycnt2 <- vislog[!is.na(flats) & month(entry)>=10 & type %in% categ,.(cnt=.N,avgstay=mean(stayhrs,na.rm = T)),by=.(name,type,month(entry))]
  #pstay2 <- dailycnt2 %>% ggplot(aes(name,avgstay)) + geom_col() + facet_wrap(~ type + mnthf[month])

  if(stay==T)  dailycnt %>% ggplot(aes(day,avgstay)) + geom_col() + facet_wrap(~ type + month) else
    dailycnt %>% ggplot(aes(day,cnt)) + geom_col() + facet_wrap(~ type + month)
}

# plot for attrition
# focus on 2019
plotattr <- function(data=vislog,filter="Security|Guard",TITLE="ATTRITION",start_mnth="Dec_2018"){
  new1<- marknew(data,vis=filter)[,.(mnth=crfact(entry),NEW,name)][order(mnth,NEW)] %>% unique
  new1[mnth>=start_mnth,.N,by=.(mnth,NEW)] %>% ggplot(aes(mnth,N)) + geom_col(aes(fill=NEW)) + labs(x="MONTH",y="UNIQUE EMPLOYEES",title=TITLE)
}

#plot for checkins
plotcheckins <- function(data=vislog,filter="Security",startmnth="May_2019",endmnth= NA,TITLE="CHECK INS"){
  data <- data[grepl(filter,type,ig=T)]
  data[crfact(entry)>=startmnth,.N,by=.(mth=crfact(entry),type)] %>% ggplot(aes(mth,N)) +
    geom_line(aes(group=type,color=type)) +
    geom_point() +
    labs(x="Month",y="TOTAL CHECK IN COUNTS",title=TITLE)
}

#plot 5 : does not work delete it.
plotbub <- function(dt=vislog,m=5,vistype="^Driver",flatwise=T,flatno=NULL){
  dt2 <- copy(dt[month(entry)==m & grepl(vistype,type)])
  if(!is.null(flatno))
    dt2 <- dt2[grepl(flatno,flats)]
  #anal[,day2:=day(exit)]
  if(flatwise) {
    dt2$flats %<>%  factor(levels=FLATS,ordered=T)
    dt2[order(flats)] %>%
    ggplot(aes(day,y=interaction(flats,name,lex.order = T,drop = T))) +
    geom_point(aes(size=tstay,color=exception,shape=nightstay)) +
    labs(x=paste("DAY OF MONTH:",month.abb[m]),y=paste("FLAT /",vistype),title="CHECKIN EFFICIENCY",subtitle= "BUBBLE CHART") +
    scale_x_continuous(breaks = 1:30)
  } else
    dt2 %>%
    ggplot(aes(day,name)) +
    geom_point(aes(size=tstay,color=nightstay)) +
    #geom_point(aes(day2,y= name),color="lightgreen",size=0.9) +
    labs(x=paste("DAY OF MONTH:",month.abb[m]),y=vistype,title="CHECKIN EFFICIENCY",subtitle= "BUBBLE CHART") +
    scale_x_continuous(breaks = 1:30)
}
#p5 <- plotbub()

# second plot this works on new processed vislog
plotbub2 <- function(dt=vislog,m=5,vistype="^Driver",upperhrs = 16, vdline=15, flat.name=T,flat.only=F,flatno=NULL,text2=paste(vistype,":",month.name[m])){
  dt2 <- copy(dt[month(entry)==m & grepl(vistype,type,ignore.case = T)])
  dt2 <- dt2[,.(tstay=sum(stayhrs,na.rm = T)),by=.(name,mob,type,flats,weekno,weekday,day(entry))]
  dt2[tstay>0.5,mhalf:=T]
  dt2[tstay>8,m8:=T]
  dt2[tstay>16,m16:=T]
  dt2[tstay>24,m24:=T]
  dt2[tstay>upperhrs,exception:=T]

  if(!is.null(flatno))
    dt2 <- dt2[grepl(flatno,flats)]
  if(flat.name) {
    dt2$flats %<>%  factor(levels=FLATS,ordered=T)
    assert_that(dt2$exception %>% is.na %>% not %>% any,msg = "There is not a single exception.")
    dt2[!is.na(flats)][order(flats)] %>%
      ggplot(aes(day,y=interaction(flats,name,lex.order = T,drop = T))) +
      geom_point(aes(size=tstay,color=exception)) +
     # geom_point(aes(day2,y=interaction(flats,name,lex.order = T,drop = T)),color="lightgreen",size=0.9) +
      labs(x=paste("DAY OF MONTH:",month.abb[m]),y=paste("FLAT /",vistype),title="CHECKIN EFFICIENCY",subtitle= text2) +
      scale_x_continuous(breaks = 1:30) +
      geom_vline(xintercept = vdline,lty=2)
  } else
    if(flat.only){
      dt2$flats %<>%  factor(levels=FLATS,ordered=T)
      dt2[order(flats)] %>%
        ggplot(aes(day,y=flats,lex.order = T,drop = T)) +
        geom_point(aes(size=tstay,color=exception)) +
        # geom_point(aes(day2,y=interaction(flats,name,lex.order = T,drop = T)),color="lightgreen",size=0.9) +
        labs(x=paste("DAY OF MONTH:",month.abb[m]),y=paste("FLAT /",vistype),title="CHECKIN EFFICIENCY",subtitle= text2) +
        scale_x_continuous(breaks = 1:30) +
        geom_vline(xintercept = vdline,lty=2)
    } else
    dt2 %>%
    ggplot(aes(day,name)) +
    geom_point(aes(size=tstay,color=exception)) +
    #geom_point(aes(day2,y= name),color="lightgreen",size=0.9) +
    labs(x=paste("DAY OF MONTH:",month.abb[m]),y=vistype,title="CHECKIN EFFICIENCY",subtitle= text2) +
    scale_x_continuous(breaks = 1:30)+
    geom_vline(xintercept = vdline,lty=2)
}

# input= the downloaded roster now a csv file; output employee wise EWD for the month
ewd_rost <- function(month=3,uptoline=21,holidays=F){
  rostfile=paste0("roster_",tolower(month.abb[month]),".csv")
  xros <- fread(rostfile,nrows = 22)
  setDT(xros)
  ewdcol <- grep("EWD|Expected", names(xros),ig=T)
  holrow <- grep("DATE",xros[[1]],ig=T)
  assert_that(length(holrow)>0,msg = "Date row not found for Holidays")
  table1 <- xros[1:(holrow-1),c(1,ewdcol),with=F] %>% textclean::drop_empty_row()
  names(table1) <- qc(name,ewd)
  table1$name %<>% str_trim()
  table1 <- table1[!is.na(ewd)]
  table2 <- xros[(holrow+1):uptoline,c(1,2)] %>% textclean::drop_empty_row()
  #table2 <- table2[2:nrow(table2)]
  setnames(table2,c("date","occasion"))
  table2[,date:=parse_date_time(date,orders = c("bdy","dby","ymd"),tz = "Asia/Kolkata")]
  if(holidays) table2 else table1
}

holidaydt <- function(passed_mnths=8:10){
  x1 <- passed_mnths %>% purrr::map_dfr(.f = ewd_rost,holidays=T,.id = "month")
  x1$month <- as.numeric(x1$month) + 7
  x1$mnth <- month.abb[x1$month]
  x1[,month:=NULL]
}

pasteSh <- function(x) paste(unique(x),collapse = "+")


# Not used bur use to recollect excel processing
upd_staff_attendance <-
  function(
    mnth= "Jul_2019",
    staff_string=squad_string, # this is now hardcoded. No other option
    removenames = "ZZ",
    file_squad="squad.xlsx",file_emb = "emb.xlsx", file_flat = "flat_emp.xlsx",
    upload=F, plot = F,
    data=vislog,
    google_sheet_file = "dummy"
  ){
    #mnth<- mnum(fmnth)
    if(grepl("guard",staff_string,ig=T)) tmpfile=file_squad
    wb <- createWorkbook(title="staff_wise_chckin")

    # remove undesired person
    data <- data[!grepl(removenames,name,ig=T)]

    # initialise with the shift database   - not to be used for attendance now.

    dtall <- show_gap(mth = mnth,dt = dtshifts,type_str = staff_string)
    # prepare xsec - aggregated on stay time and trips per day per person
    # xsec <- copy(dtshifts)
    # xsec[,tstay:=sum(stayhrs),by=.(name,mob,day(entry))]
    #xsec[,percstay:=stayhrs/tstay*100]
    #xsec[,trips:=.N,by=.(name,day(entry))]

    monthname <- month.name[mnth]

    addWorksheet(wb,"summary",tabColour = "red")
    addWorksheet(wb,"wide",tabColour = "blue")
    addWorksheet(wb,"daily",tabColour = "grey")
    #addWorksheet(wb,"swipes",tabColour = "grey")
    sty.cen <- createStyle(halign = "center")
    sty.rt <- createStyle(halign = "right")
    sty.tall <- createStyle(valign = "center")
    sty.title1 <- createStyle(wrapText = F,fontName = "Al Tarikh", textDecoration ="bold",
                              fontSize = 18,fontColour = "black",
                              halign = "center",valign = "center",fgFill = colours()[633])
    sty.title2 <- createStyle(wrapText = T,fontName = "Al Bayan", textDecoration ="bold",
                              fontSize = 14,fontColour = "black",
                              halign = "right",valign = "center")

    sty.shade_pink <- createStyle(fontName = "Al Bayan", textDecoration ="bold",bgFill = "deeppink2")
    sty.shade_green <- createStyle(fontName = "Al Bayan", textDecoration ="bold",bgFill = "chartreuse2")
    sty.shade_night <- createStyle(fontColour = "darkviolet",textDecoration ="bold")
    sty.shade_N2 <- createStyle(fontName = "Al Bayan", textDecoration ="bold",bgFill = "dodgerblue3")
    sty.shade_D2 <- createStyle(fontName = "Al Bayan", textDecoration ="bold",bgFill = "orange1")

    sty.date.number <- createStyle(wrapText = F,fontName = "Al Tarikh", textDecoration ="bold",
                                   fontSize = 14,fontColour = "black",
                                   halign = "right",valign = "center",fgFill = "grey93" )

    sty.hours <- createStyle(numFmt = "0.0",textDecoration = "italic",halign = "right")
    snum <- createStyle(numFmt = "#,##0.00")


    # main function to write a table with style for title
    append_short_tables <- function(dtreport,sheetnumber,title="DUMMY",sty,cols=40, offset=0){
      cat("\nAppending short table in sheet:",sheetnumber, "rowno:",runner[sheetnumber])
      rownumber <- runner[sheetnumber]
      writeData(wb,sheetnumber,x = title,startRow = rownumber + 1,startCol = 1)
      writeDataTable(wb, sheet = sheetnumber, x = get(dtreport), startCol = 1 + offset, startRow = rownumber + 2,
                     colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleLight1",
                     bandedRows = TRUE, withFilter = F,tableName = dtreport)
      mergeCells(wb,sheetnumber,cols = seq_len(ncol(get(dtreport))+ offset),rows= rownumber + 1) # table heading
      addStyle(wb,sheetnumber,style = sty.title1,rows = rownumber + 1,cols = 1,stack=T,gridExpand = T) # table heading
      addStyle(wb,sheetnumber,style = sty.date.number,rows = rownumber + 2,cols = offset + seq_len(cols),gridExpand = T) # date headings
      addStyle(wb,sheetnumber,style = createStyle(halign = "left"),rows = rownumber + 2,cols =ifelse(offset, 2, 1:2),stack=T,gridExpand = T) # date headings
      setRowHeights(wb,sheetnumber,rows = rownumber + 1,heights = 36) # table heading
      #addStyle(wb,sheetnumber,style = sty.title2,rows = rownumber + 2,cols = cols,stack=T) # column heading - just below table heading
      runner[sheetnumber] <<- runner[sheetnumber] + 17 # this is the global row counter to track utilised rows
    } # end of function

    runner<- integer(10L)
    runner[1:10]<- 1  # running counter for consumed rows in the workbooks indexed in runner
    # tab_sum1 <- calcattendance(m=mnth,visdt = data)
    # tab_sum2 <- calcattendance(m=mnth,visdt = data) # have to modify this.
    #tab_hc<- dtshifts[,.(tothrs=sum(stayhrs),totheads=sum(stayhrs)/ifelse(grepl("Lady",type),9,12)),by=.(type,day(entry))] %>% dcast(type ~ day,value.var="totheads",fun.aggregate=sum,na.rm=T)
    tab_shift <- dtshifts %>% dcast(name + type ~ day(entry),value.var="shift",fun.aggregate=pasteSh)
    tab_hrs<- calc_att_daily(m = mnth,embassy = F,hrs = T,withtotals = T)
    tab_att1<- calc_att_daily(m = mnth,embassy = F,hrs = F,withtotals = T)
    #tab_count <- dtshifts[,.(count=.N),by=.(name,type,day(entry))] %>% dcast(name + type ~ day,value.var = "count",fun.aggregate=sum)
    tab_doub<- dtshifts %>% dcast(name + type ~ day(entry),value.var="double",fun.aggregate=sum,na.rm=T)
    #tab_trip<- dtshifts %>% dcast(name + type ~ day(entry),value.var="triple",fun.aggregate=sum,na.rm=T)
    tab_tall<- dtshifts
    # removed the extra columns from from func calls. Still to decide if these are a good idea or hard coding the styling is easier to manage.
    append_short_tables("tab_sum2",1,title = paste("SUMMARY: ATTENDANCE-",toupper(monthname),"[Uptil: ",max(dtshifts$exit,dtshifts$entry,na.rm = T)," ]"))
    append_short_tables("tab_shift",2,title = "SHIFTS - DAY AND NIGHT")
    #append_short_tables("tab_count",2,title = "NUMBER OF TIMES CHECKED OUT")
    append_short_tables("tab_hrs",2,"HOURS ON DUTY")
    append_short_tables("tab_att1",2,"ATTENDANCE BASED ON DAILY DUTY HOURS")
    append_short_tables("tab_doub",2,"DOUBLE SHIFTS OR SWIPE MISSED")
    #append_short_tables("tab_trip",2,"TRIPLE SHIFTS OR SWIPE MISSED")
    append_short_tables("tab_hc",2, "HEADCOUNT SUMMARY",offset = 1)
    append_short_tables("tab_tall",3,"RAW DATA : DAILY LOG",sty = sty.tall,cols=ncol(tab_tall))
    cat("\ncols in wide reports:",ncol(tab_shift))
    conditionalFormatting(wb,sheet = 2,cols = 3:(ncol(tab_shift)+2),rows = 4:17,type = "contains", rule = "NIGHT",style =sty.shade_night )
    conditionalFormatting(wb,sheet = 2,cols = 3:(ncol(tab_shift)+2),rows = 4:17,type = "contains", rule = "NIGHT+DAY",style =sty.shade_N2 )
    conditionalFormatting(wb,sheet = 2,cols = 3:(ncol(tab_shift)+2),rows = 4:17,type = "contains", rule = "DAY+NIGHT",style =sty.shade_D2 )
    conditionalFormatting(wb,sheet = 2,cols = 3:(ncol(tab_shift)+2),rows = 21:34,rule = ">1",style =sty.shade_green )
    conditionalFormatting(wb,sheet = 2,cols = 3:(ncol(tab_shift)+2),rows = 38:51,rule = ">15",style =sty.shade_D2 )
    conditionalFormatting(wb,sheet = 2,cols = 3:(ncol(tab_shift)+2),rows = 55:68,rule = "=1",style =sty.shade_pink )
    conditionalFormatting(wb,sheet = 2,cols = 3:(ncol(tab_shift)+2),rows = 72:85,rule = "=1",style =sty.shade_pink )
    conditionalFormatting(wb,sheet = 2,cols = 3:(ncol(tab_shift)+2),rows = 89:102,rule = ">1",style =sty.shade_N2 )

    setColWidths(wb,1,cols = c(1,2,6),widths = c(25,15,25),ignoreMergedCells = T)
    setColWidths(wb,2,cols = 1:2,widths = c(25,15),ignoreMergedCells = T)
    setColWidths(wb,3,cols = 1:3,widths = c(25,15,10),ignoreMergedCells = T)
    #setColWidths(wb,4,cols = c(1:3,5:7),widths = c(25,15,10,15,15,15),ignoreMergedCells = T)

    addStyle(wb,1,style = snum,rows = 1:100,cols = 9:10,stack = T,gridExpand = T) # last two columns are decimal formatted
    addStyle(wb,2,style = snum,rows = 106:110,cols = 3: (ncol(tab_hc)+2) ,stack = T,gridExpand = T) # decimal formatted HC
    addStyle(wb,3,style = snum,rows = 1:1000,cols = c(8:10,15), stack = T,gridExpand = T) # decimal formatted HC
    #addStyle(wb,4,style = snum,rows = 1:1000,cols = c(8:10),stack = T,gridExpand = T) # decimal formatted HC

    openXL(wb)
    saveWorkbook(wb,file = tmpfile,overwrite = T)
    if(upload) gs_upload(file = tmpfile,sheet_title = google_sheet_file,overwrite = T)
  }

################## STAFF ATTENDANCE AND PAYMENTS ################

# Process the downloaded attendance.csv from new dashboard
# new columns in mygate raw data added "In Date" and "Out Date" since Oct 2022
# made changes to incorporate this; simplified outdate computations.
proc_attend <- function(fname="attendance_detail.csv",mth=10){
  x1 <- fread(fname) |> clean_names()
 # use standard R way of decoding time (thanks Ronak Shah in stackoverflow)
  x1[,inat:=paste(in_date,in_time) |> parse_date_time(orders = "dmyIM %p",tz = "Asia/Kolkata")]
  x1[,outat:=paste(out_date,out_time) |> parse_date_time(orders = "dmyIM %p",tz = "Asia/Kolkata")]
  x1[,in_date:=as_date(inat)]
  x1[,out_date:=as_date(outat)]
  x1[!is.na(outat),hours:=as.numeric(outat - inat,unit = "hours")]
  x1[,fullday_hours:=sum(hours),by = .(in_date,name)] # new line added, now the fullhrs cannot be sum aggregated.
  x2 <- x1[,.(name,type,in_date,fullday_hours)] |> unique()
  x2[,att:=ifelse(grepl("Secur",type),round_to_fraction(fullday_hours/12,denominator = 2),NA)]
  x2[,att:=ifelse(grepl("Lady",type),round_to_fraction(fullday_hours/8,denominator = 2),att)]
  x2[,att:=ifelse(grepl("Ele",type),round_to_fraction(fullday_hours/9.5,denominator = 2),att)]
  x2[,att:=ifelse(grepl("Plu",type),round_to_fraction(fullday_hours/9.5,denominator = 2),att)]
  x2[,att:=ifelse(grepl("House",type),round_to_fraction(fullday_hours/8,denominator = 2),att)]
  x2[,att:=ifelse(grepl("STP",type),round_to_fraction(fullday_hours/12,denominator = 2),att)]
  x2[,att:=ifelse(grepl("Gardener",type),round_to_fraction(fullday_hours/8,denominator = 2),att)]
  x2[,ot:=case_when(
      grepl("Plu|Ele",type) & att == 1 ~  pmax(0,fullday_hours - 9.5),
      T ~  0,
  )]
}

markattendance <- function(dt = tall_stats,dm=duty_matrix,rule=c(half=0.75,zero=0.25)){
  dt <- dt[dm,on="type",nomatch=0]
  dt[,att:=1]
  dt[tstay < rule["half"]*dutyhrs,att:=0.5]
  dt[tstay < rule["zero"]*dutyhrs,att:=0]
}

# simplified by not loading employee list from google sheet and no strict matching names.
# this is fantastic. Reliable and fast.
calc_att_daily <- function(m="Jul_2020",embassy=T,hrs=F,withtotals=T,dt=vislog){
  #line added only for Aug 2020 for one gardener who was missed in type
  dt[grepl("rafik",name,ig=T) & m=="Aug_2020",type:="Society Gardener-Male"]
  staffdt <- dt[(grepl(ifelse(embassy,embassy_str,squad_str),type,ig=T) | mob == "9008005492") & crfact(entry)==m][duty_matrix,on="type",nomatch=0]
  wide1 <- staffdt %>%
    .[,.(tstay=sum(stayhrs),trips=.N),by=.(day(entry),name,type)] %>%
    markattendance() %>% # this marks half day for less than 75% hours of duty and 0 for < 25%
    dcast(name + type ~ day,fun.aggregate = sum,value.var="att")
  wide2 <- staffdt[,.(tstay=sum(stayhrs) %>% round(digits = 2),trips=.N),by=.(day(entry),name,type,dutyhrs)] %>%
    dcast(name + type ~ day,fun.aggregate = sum,value.var="tstay")
  if(withtotals){
    wide1 <- janitor::adorn_totals(wide1,where = c("row","col"))
    wide2 <- janitor::adorn_totals(wide2,where = c("row","col"))
  }
  if(hrs) wide2 else wide1
}

# a very short code without bells and whistles to check attendance of any staff quickly
quick_hrs <- function(pat="Housekee",m=3,data=vislog){
  data[grepl(pat,type,ig=T) & month(entry)==m][,.(stay=sum(stayhrs)),by=.(name,day(entry),type,weekdays(entry))] %>% dcast(name + type ~ day)
}

# this is the  manhours based attendance so not on a daily basis. Donot use it for internal staff.
calcattendance <- function(m="Oct_2019",visdt=vislog,dm=duty_matrix,type=squad_string){
  dt<- show_gap(m,dt = visdt,type_str = type)
  maxdate <- max(day(dt$entry))
  x1 <- dt[double==T | consec==T,.(dble_shifts=.N),by=name]
  #x11 <- dt[triple==T,.(trple_shifts=.N),by=.(name)]
  x2 <- dt[,{
    da<-unique(day(entry))
    ab<- setdiff(seq_len(maxdate),da)
    .(abdays=list(ab),abn=length(ab),prdays=maxdate-length(ab))},
    by=.(name,type)]
  x3 <- dt[dm,on="type",nomatch=0][,
                                   .(tothrs=sum(stayhrs,na.rm = T),
                                     duties=sum(stayhrs,na.rm = T)/first(dutyhrs),duty_hrs=first(dutyhrs)),
                                   by=.(name,type)]
  x11[x1,on="name"][x2,on="name"][x3,on=.(name,type)] %>% setcolorder(neworder = c("name","type","duty_hrs"))
}

emp_subset <- function(mnth="May_2019",data=vislog){
  loademp()-> empids
  data[,name:=str_trim(name)] # remove this line once the input data is clean
  data <- copy(data[name %in% empids$name_mygate & grepl(embassy_string,type) & crfact(entry) %in% mnth])
  data <- data[,`:=`(from=NULL,vehicle=NULL,flats=NULL)]
  duty_matrix[data,on="type"]
}
# this is the main function for attendance and double shifts
show_gap <- function(mth="Oct_2019",type_str="Security|Guard",dt=vislog,flats=F){
  dtshifts <- sec_shift(mnth = mth,dtvis = dt, pattern = type_str)
  #dt2 <- copy(dt[grepl(type_str,type,ig=T) & crfact(entry)==mth][order(entry)])
  # dtshifts[,prevexit:=shift(exit,type = "lag"),by=.(name,mob)]
  # dtshifts[,prevstay:=shift(stayhrs,type = "lag"),by=.(name,mob)]
  # #vislog2$prevexit %<>% as.POSIXct(tz = "UTC",origin="1970-01-01")
  # dtshifts[,gap:=entry - prevexit]
  #dt2[,gaphrs:=as.numeric(gap)/3600 %>% round(2)]
  dtshifts[away_before < dminutes(60) & (stayhrs + prevstay) > 22,consec:=T]
  dtshifts[stayhrs > 22,double:=T]

  if(flats) dtshifts[order(name,entry),.(name,type,flats,prevexit,entry,exit,stayhrs=tot_shift,stay_out,prevstay,gap,weekday=weekdays(entry),shift,double,consec)] else
    dtshifts[order(name,entry),.(name,type,prevexit,entry,exit,stayhrs=tot_shift,stay_out,prevstay,gap,weekday=weekdays(entry),shift,double,consec)]
}

# new function to see monthly trend on stay period or number of times 2 or more entries made for same person
trnd_checkin <- function(typeofvis="^Driver",mean=F,namestr=""){
  if(!mean){
  x1 <- vislog[entry>=ymd(20180801) & grepl(typeofvis,type) & grepl(namestr,name,ignore.case = T),
               .(meantrips = mean(.N)),by=.(day(entry),month(entry),name)][
           meantrips>1][
             ,.N,by="month"]

  x1[,mnth:=crfactm(mnth = month,2018)]
  output <- x1[month<=curmth,mnth:=crfactm(month,2019)][,.(mnth,Nmulti=N)][order(mnth)]
  } else{
    x1 <- vislog[entry>=ymd(20180801) & grepl(typeofvis,type) & grepl(namestr,name,ignore.case = T),
           .(staypv = mean(stayhrs,na.rm = T),.N),by=month(entry)]
  x1[,mnth:=crfactm(mnth = month,2018)]
  output <- x1[month<=curmth,mnth:=crfactm(month,2019)][,.(mnth,staypv,N)][order(mnth)]
  }
  output
}

vis_sum <- function(tpat="^Technical",npat="",starting="Aug_2018"){
  vislog[grepl(tpat,type,ig=T) & grepl(npat,name,ig=T),
         .(avgstay=round(mean(stayhrs,na.rm = T),2),
           tottrips=.N,
           totdays=NROW(unique(day(entry)))),
         by=.(name,type,crfact(entry))][
           ,totstay:=avgstay*tottrips][order(crfact,name)]
}


longstay <- function(stdate=NA,enddate=now(),mnth="Jan_2019"){
  if(is.na(stdate))
    vislog[crfact(entry)==mnth][order(-stay)][1:25][,-c("stay","shift","weekno","from","vehicle")] else
      vislog[entry>=ymd(stdate)][order(-stay)][1:25][,-c("stay","shift","weekno","from","vehicle")]
}

# new function mainly for technicians OT - now simplified
# still to use duty_matrix start and end; right now hard coded at 8 am and 19 hrs
otcalc <- function(m="Apr_2019",pattern=embassy_string,logdt = vislog){
  dt <- logdt[grepl(pattern,type,ig=T) & crfact(entry)==m]
  dt[,totday:=sum(stayhrs,na.rm = T),by=.(name,day(entry))]
  dt[,flats:=NULL]; dt[,from:=NULL]; dt[,vehicle:=NULL]
  dt[,next_entry:=shift(entry,type = "lead"),by=name]
  dt[,prev_exit:=shift(exit,type = "lag"),by=name]
  dt[,stay_out:=as.numeric(next_entry - exit,units="hours")] # retaining this column for ease of recollection
  dt[,away_next:=stay_out] # copy it for compatible names
  dt[,away_before:=as.numeric(entry - prev_exit,units="hours")]
  dt[,kind:=ifelse(hour(entry) %in% c(6:9)  & away_before > 5,"FIRST",NA_character_)]
  dt[,kind:=ifelse(hour(entry) %in% c(10:14)  & away_before >10,"FIRST",kind)]
  dt[,kind:=ifelse(away_next > 3  & hour(exit) %in% c(16:18,20),"LAST",kind)]
  dt[,kind:=ifelse(away_next > 12,"LAST",kind)]
  dt[,kind:=ifelse(away_next > 1  & hour(exit) == 19 ,"LAST",kind)]
  dt[,kind:=ifelse(away_next < 3  & away_before < 3 & hour(exit) <19 ,"INTERMEDIATE",kind)]
  dt[,kind:=ifelse(is.na(next_entry)  & hour(exit) %in% c(16:21),"LAST",kind)]
  dt[,kind:=ifelse(hour(exit) %in% 18:19 & (away_before>10 | hour(entry) %in% c(8)),"SINGLE",kind)]
  dt[,kind:=ifelse(away_next > 10 & away_before>10,"SINGLE",kind)]
  dt[,kind:=ifelse(hour(entry) >=19 & stayhrs <= 10,"CALL",kind)]
  dt[,kind:=ifelse(hour(entry) <= 6 & hour(exit) <=7,"CALL",kind)]
  dt[,kind:=ifelse((hour(entry)>19) & stayhrs >= 8,"NIGHT",kind)]
  dt[kind=="CALL",otnight :=round(stayhrs+0.5)]
  dt[is.na(otnight),otnight:=0]
  dt <- dt[kind %in% c("FIRST","INTERMEDIATE"),tawayn:=sum(away_next,na.rm = T),by=.(name,day(entry))] # summarisaton but returning full DT
  dt[is.na(tawayn),tawayn:=0]
  dt2 <- dt[,.(mnth=first(crfact(entry)),tothrs = sum(stayhrs) + first(tawayn),away=first(tawayn), otnight=first(otnight)),by=.(name,day(entry),weekday)]
  dt2[,ot_tot:=otnight + max(tothrs-10,0) %>% round(2),by=.(name,day)]
  dt2[,att:=ifelse(tothrs>=10,1,tothrs/10)]
  dt2[,att:=ifelse(att>0.75,1,ifelse(att<.25,0,0.5))]
  dt2[att==0 & tothrs >0,ot_tot:=ifelse(ot_tot>0,ot_tot,round(tothrs+0.5,0)) %>% round(digits = 2)]
  dt2[,tothrs:= round(tothrs,2)][,away:=round(away,2)]
}

# new datatable with shift details
sec_shift <- function(mnth="Oct_2019",pattern=squad_string,dtvis=vislog){
  yr <- str_sub(mnth,-4)
  mno <- which(month.abb==str_sub(mnth,1,3))
  daysmnth <- days_in_month(mno)
  # fix leap year
  if(mno==2 & dmy(paste("01",mnth)) %>% leap_year()) daysmnth <- 29
  dt <- copy(dtvis[grepl(pattern,type,ig=T) & crfact(entry)==mnth & !is.na(exit)])
  dt[,totday:=sum(stayhrs,na.rm = T),by=.(name,day(entry))]
  dt[,flats:=NULL]; dt[,from:=NULL]; dt[,vehicle:=NULL]
  day_shift_start <-  formatC(seq_len(daysmnth),digits = 1,flag = "0") %>% paste(yr,mno,.,sep = "-") %>% paste("08:00") %>% as.POSIXct()
  night_shift_start <-  formatC(seq_len(daysmnth),digits = 1,flag = "0") %>% paste(yr,mno,.,sep = "-") %>% paste("20:00") %>% as.POSIXct()
  mat_day_shifts <- matrix(c(day_shift_start,night_shift_start),ncol = 2)
  mat_night_shifts <- matrix(c(night_shift_start,shift(day_shift_start,type = "lead",fill = last(day_shift_start) + ddays(1))),ncol = 2)
  dt[,
     nightsh := seq_len(nrow(dt)) %>% map(~Overlap(mat_night_shifts,c(dt$entry[.x],dt$exit[.x]))/3600) %>% map_int(~which(.x > 1) %>% ifelse(length(.)==0,NA,.) )][,
       daysh := seq_len(nrow(dt)) %>% map(~Overlap(mat_day_shifts,c(dt$entry[.x],dt$exit[.x]))/3600) %>% map_int(~which(.x > 1) %>% ifelse(length(.)==0,NA,.) )][,
          hrs_night := seq_len(nrow(dt)) %>% map(~Overlap(mat_night_shifts,c(dt$entry[.x],dt$exit[.x]))/3600) %>% map_dbl(sum)][,
            hrs_day := seq_len(nrow(dt)) %>% map(~Overlap(mat_day_shifts,c(dt$entry[.x],dt$exit[.x]))/3600) %>% map_dbl(sum)
       ]
}


# change shift to "day", "night" or "all"; collapse=F AND shift="all" will give day and night in alternate rows
# pipe in the output of prev function to this
report_shifts <- function(dt,shift="day",collapse=F){
  dt_day <- dt[!is.na(daysh),.(tot_shift=sum(hrs_day)),by=.(name,type,daysh)] %>%
    dcast(name + type ~ daysh,value.var="tot_shift",fun.aggregate=function(x) if(length(x)==0) "-" else format(x,digits = 1))
  dt_night <- dt[!is.na(nightsh),.(tot_shift=sum(hrs_night)),by=.(name,type,nightsh)] %>%
    dcast(name + type  ~ nightsh,value.var="tot_shift",fun.aggregate=function(x) if(length(x)==0) "-" else format(x,digits = 1))

  if(grepl("^[Dd]",shift)) return(dt_day)  else
    if(grepl("^[Nn]",shift)) return(dt_night) else
      if(grepl("^[Aa]",shift) & collapse==F){
        dt_day[,shift:="DAY"]
        dt_night[,shift:="NIGHT"]
        rbind(dt_day,dt_night) %>% setcolorder(qc(name,shift)) %>% .[order(name)]
      } else
        dt[,.(tot_shift=sum(hrs_day,hrs_night)),by=.(name,type,day(entry))] %>%
    dcast(name + type ~ day,value.var="tot_shift",fun.aggregate=function(x) if(length(x)==0) "-" else format(x,digits = 1))

}

create_shift_rep <- function(mnth="Oct_2019",pattern=squad_string,dtvis=vislog){
  sec_shift(mnth = mnth,pattern = pattern,dtvis = dtvis) %>%
    report_shifts(shift="a") %>%
    kable() %>%
    kable_styling() %>%
    collapse_rows(columns = 1)
}

# use this to directly append the current month's OT to the same excel sheet: not working now
create_ot_excel <- function(names=c("plum","electr"),mnths=c("Jul_2019"),exfile="ot.xlsx",overwrite=F){
  if(file.exists(exfile) & !overwrite) wb <- openxlsx::loadWorkbook(exfile) else wb <- createWorkbook() # open the same excel sheet if it exists
  namestr <- rep(names,each=length(mnths))
  mnthstr <- rep(mnths,times=length(names))
  tabs <- namestr %>% map2(mnthstr,~otcalc(mnth = .y,pattern = .x))
  sheetnames <- tabs %>% map_chr(~.x[,paste(first(name),first(mnth))]) %>% walk(~addWorksheet(wb,.x))
  sheetnames %>% walk2(tabs,~writeDataTable(wb,sheet = .x,x = .y))
  saveWorkbook(wb,exfile,overwrite = T)
}

cleaned_up_stay <- function(m=3,pat=embassy_string){
  dtlong <- otcalc(pattern = pat)
  dtlong[kind %in% c("SINGLE","LAST"),dutyend:=exit]
  dtlong[kind %in% c("SINGLE","FIRST"),dutystart:=entry]
  dt2 <- dtlong[month(entry)==m][!all(is.na(dutyend),is.na(dutystart))]
  dt2[,.(dutyst=max(dutystart,na.rm = T),dutyfin = max(dutyend,na.rm = T)),by=.(name,day(entry))]
}

vehshape <- function(txtfile="vehicle_logs/vehlog_feb15_mar15.csv"){
  x2 <- fread(txtfile)
  x2[,dattim:=dmy_hm(Time)]
  x2[,Date:=as.Date(dattim)]
  x2[,flatn:=as.numeric(str_extract(Flat,"\\d+"))]
  setnames(x2,qc(flat,vehnb,vehtype,entry,dummy,dattim,Date,flatn))
  x2[,dummy:=NULL]
}

################## COVID TRACKING ################

# pass the vislog and a days range (pass 1:31 if full month needed)
plot_viscount <- function(dt = vislog,days=1:31,tit="EMBASSY HAVEN VISITOR COUNT",subtit="Month on month"){
dt <- dt[day(entry) %in% days & !is.na(flats) & !grepl("Deliver",type,ig=T) & !grepl("stay",from,ig=T)]
dt[,count:=.N,by=crfact(entry)]
dt[,countmax:=max(count,na.rm = T)]
data <- dt[,.(month=crfact(entry),countmax,count)] %>% unique
data[,perct:=1 - count/countmax]
data[,perct:=formattable::percent(perct)]
data[,labpos:=mean(c(countmax,count)),by=month]
data %>% ggplot(aes(month,count)) + geom_col() +
  labs(title = tit,subtitle = subtit,x="Month",y="VISITOR COUNT") +
  geom_errorbar(aes(ymin=count,ymax=countmax),width=0.2) +
  geom_label(aes(y=labpos,label=perct))
}

# Gets confusing to decode, as improvements are not obvious
plot_flatcounts <- function(dt = vislog,days=1:31,tit="FLAT COUNTS",subtit="Day on day changes"){
dt <- dt[day(entry) %in% days & !is.na(flats)]
dt <- exclude_stayGate(dt)
dt1 <- dt[,{uvis = uniqueN(mob); .(count=as.integer(uvis))},by=.(flats,Date=as.Date(entry))]
dt1[,.N,by=.(Date,count)]

}

plot_flatwise <- function(dt = vislog,boxtext="July 15-16",days=1:31,tit="EMBASSY HAVEN",
                          subtit="Unique Visitor count since Jun 01, 2020",table=F){
  dt <- dt[day(entry) %in% days & !grepl("Delivery",type)]
  dt[,flatn:=as.factor(flats)]
  dt[,mth:=crfact(entry)]
  dt <- exclude_stayGate(dt) # remove absent maids and stay at gate visitors
  dt1 <- dt[,{uvis = uniqueN(mob); .(count=as.integer(uvis))},by=.(flatn,day(entry))]
  dt2 <- dt[,.(flatn,name,mob,type,from)] %>% unique

  #dt[,count:=.N,by=.(flatn)]
  #data <- dt[,.(flatn,count)] %>% unique
  data <- dt1[!is.na(flatn)]
  data[,Density:= case_when(count>=3 ~ 1,
                         count==2 ~ 2,
                         count<=1 ~ 3,
                         TRUE ~ 4
                         )]
  posy <- 0.75*max(data$count,na.rm = T)
  posx <- 0.5*length(data$flatn)
  data[,Density:=factor(Density)]
  if(table==T) dt2 else
    data %>%
    ggplot(aes(flatn,count)) +
    geom_col(aes(fill=Density),position = "dodge") +
    scale_fill_manual(values = brewer.pal(4,"PiYG"),labels=c("High","Medium","Low","Excellent")) +
    ylim(0,NA) +
    coord_flip() +
    labs(title = tit,subtitle = subtit,x="FLAT",y="UNIQUE VISITOR COUNT")+
    theme(axis.text.x = element_text(size = 18,face = "bold")) + theme(text = element_text(size = 18)) +
    geom_label(aes(x=17,y=posy),data = data.table(),label=boxtext, size=12,label.padding = unit(1.5,"lines"),label.size = 1)

}

# comparative visitors
compare_flvis <- function(dt = vjul,days=16:31){
  active_flats <- dt[!is.na(flats),unique(flats)]
  dt[,dayofm:=day(entry)]
  dt[,mth:=crfact(entry)]
  dt <- dt[dayofm %in% days & !grepl("Delivery|White Lin",type,ig=T) & !is.na(flats)]
  dt[,flatn:=as.factor(flats)]
  dt <- exclude_stayGate(dt) # remove absent maids and stay at gate visitors
  dt1 <- dt[day(entry) %in% days,{uvis1 = uniqueN(mob); .(count=as.integer(uvis1))},by=.(flatn,Date=as.Date(entry))]
  dt1[,Date := format(Date,format="%B-%d")]
  zflats <- dt1[,as.character(flatn)] %>% setdiff(as.character(active_flats),.) %>% as.integer()
  return(list(dt1=dt1,zero_vis=zflats))

}

exclude_stayGate <- function(dt,exclusionfile="lockdown_exclusions.csv"){
  x1 <- fread(exclusionfile) # assume the file has two columns : flatn and name
  x1[,flatn:=as.factor(flatn)]
  dt <- dt[!grepl("stay\\s*at\\s*gate",x = from,ig=T)] # exclude stay at gate guests
  anti_join(dt,x1)
}

# excludes weekends but includes every visitor
plot_daily_visitors <- function(dt=vislog,removecommon=T){
  filter <- dt[,!weekday %in% c("Saturday","Sunday") ]
  if(removecommon==T) filter <- filter & dt[,!is.na(flats) & !grepl("stay",from,ig=T)]

  dt[filter,.N,.(date=as.Date(entry,tz="Asia/Kolkata"))][order(date)] %>% ggplot() + geom_line(aes(date,N)) + scale_x_date(date_breaks = "months",date_labels = "%b")
}

# load file named `entry_exit(%s).csv` sent by mygate for missing days
load_missed_data <- function(serials=3:4){
  serials %>% map(~sprintf("entry_exit (%s).csv",.x)) %>% map(fread) %>% rbindlist -> x1
}

# pass the output of `load_missed_data` and output will be appended data to master DT
append_missed_data <- function(x1,master=vislog_dec){
  x1 %>%
    mutate(mob=Mobile,entry=dmy_hms(Intime,tz = "Asia/Kolkata"),exit=dmy_hms(Outtime,tz = "Asia/Kolkata")) %>%
    mutate(stayhrs=as.numeric(exit - entry)/60) %>%
    select(name=Name,flats=Flat,type=`Sub Type`,entry,exit,stayhrs) %>% # if you want more columns add them
    rbind(master,.,fill=T) %>% # the other cols would be filled by NA
    arrange(entry)
}

# Outputs the enriched attendance along with overtime column for a single employee at time.
proc_overtime <- function(file = "~/Downloads/attendance_detail.csv",name="nag",month = 10){
  proc_attend(file,mth = month) %>%
    mutate(day = day(Date_IN),Mobile = as.character(Mobile),ot = ifelse(att==0,hours,hours -10),dayow = weekdays(Date_IN)) %>%
    mutate(ot = ifelse(ot<0,0,ot)) %>%
    filter(grepl(name,Name,ig=T))
}

load_mnthly_visitor_file <- function(filepat = "Entry_Exit"){
    fnames <- list.files(".",pattern = filepat)
    x1 <-
        fnames |>
        map(fread) |>
        rbindlist()
    x2 <- x1 |> clean_names()
    x2[,entry_time:=parse_date_time(entry_time,orders = "dmyHMS",tz = "Asia/Kolkata")]
    x2[,mobile:=as.character(mobile)]
    x2[,mnth:=month(entry_time)]
    x3 <-
        x2 |> cSplit("flatlist",sep = ",",type.convert = "as.is") |>
        tidyr::pivot_longer(cols = starts_with("flat"),values_drop_na = T,names_to = "flatn")
    return(x3)
}

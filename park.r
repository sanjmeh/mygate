require(data.table)
require(dplyr)
require(magrittr)
require(stringr)
require(purrr)
require(tidyr)
require(googlesheets)
require(splitstackshape)
source("mygate.r")
if(search() %>% str_detect("keys.data") %>% any() %>% not) attach("keys.data")

rawimp1 <- fread("carpmystic.csv",drop = c(3,4,5,6)) #tanveer sheet
rawimp2 <- fread("ecpl_parking.csv") #nalini sheet

names(rawimp1) <- c("flatn","name","parkstr")
cSplit(rawimp1,splitCols = "parkstr",sep = "&",direction = "wide") -> dtpm
dtpm[,soldby:="MYSTIC"]


rawimp2[V1!=""][,.(name=V3,flatn=V5,parkstr=V6)] -> dtpe
dtpe[2:nrow(dtpe)] -> dtpe
cSplit(dtpe,splitCols = "parkstr",sep = ",",direction = "wide") -> dtpe
dtpe[,soldby:="ECPL"]


dtp = rbind(dtpe,dtpm,fill=T)
dtp[,flatn:=as.integer(flatn)]

dtp %>% melt.data.table(id.vars = c("name","flatn","soldby"),value.name = "slot", na.rm = T) ->dtp2
allotted_parking <- dtp2[,flatn:=as.integer(flatn)]
setcolorder(allotted_parking,neworder = c("slot","variable", "flatn","name","soldby"))
allotted_parking[order(slot)]->allotted_parking


read_veh <- function(ghandle = gs_key(key_parkingdetails),survey=T) { # Parking Details of Rudra - manual noting of car numbers
    manual = as.data.table(gs_read(ghandle,ws = 1,range = cell_cols("A:D")))
    names(manual)<-c("labelled_flatn","slot","vehicle","type")
    manual$labelled_flatn %<>% as.integer()
    manual$slot %<>% as.integer()
    #manual[labelled_flatn==204 & vehicle == "KA01MG4432" ,slot:=1010] # removing duplicate slot 10; line to be deleted after data cleaned for slot 10 
    unique(manual[!is.na(vehicle)]) -> manual
    manual$vehicle %>% str_extract_all("\\d+") -> lnumbers
    manual[,numb:={sapply(X = lnumbers,FUN = function(x) x[2]) %>% as.integer()}]    # take the second number
    signed = as.data.table(gs_read(ghandle,ws = "signed",range = cell_cols("A:B")))
    names(signed)<- c("number","flatn")
    signed[!is.na(flatn)] -> signed
    signed[,flatn:=as.integer(flatn)]
    if(survey) manual else signed
}


free_slots_mystic = c(25,30,  31,  34,  35,  44,  51,  59,  60,  63,  79,  80, 107, 111, 123)
dfree1 <- data.table(slot=free_slots_mystic)
dfree1[,variable:="UNALLOC"]
dfree1[,c("flatn","name"):=NA]
dfree1[,soldby:="MYSTIC"]

free_slots_ecpl = c(47,69,70,118)
dfree2 <- data.table(slot=free_slots_ecpl)
dfree2[,variable:="UNALLOC"]
dfree2[,c("flatn","name"):=NA]
dfree2[,soldby:="ECPL"]

rbind(allotted_parking,dfree1,dfree2)[order(slot)] -> all_parking

fread("haven-residents.csv",skip = 1) ->hres #dump from google group
names(hres)<- c("email","nick","grpst","emailst","emailpr","postperm","yr","mth","day","hr","min","sec","tz")
hres[,joined_on:=as.Date(paste0(yr,"-",mth,"-",day))]



#dtp2 %>% spread(key = variable,value = slot) ->dtpw #this is just reversing to dtp shape, so a repeat in some sense. Just to practice
invalid_plates <- function(dt=manual,park=all_parking) dt[all_parking,on="slot"][labelled_flatn!=flatn,.(slot,labelled_flatn,'Actual allocation'=flatn,vehicle)][order(slot)]

if(!exists("rdu")) rdu <- down_rdu()

#print the slot wise cars as parked on ground for two consecutive days.
print_cars <- function() allotted_parking[
    manual,on="slot"][
        rdu,.(slot,flatn,name,soldby,vehicle,type),on="flatn"][
        !is.na(slot) & !is.na(vehicle)][
            ,.(vehnos=list(unique(vehicle)),flats=list(unique(flatn))),by=slot][
                order(slot)]


rdu_exp<- function(x=rdu){
    names(x) <- c("blk" ,    "flatn" ,  "oname" ,  "tname" ,  "oemail"  ,"temails", "ophn"    ,"tphn" ,   "status",  "isres" ,  "sqft",    "soldby",  "app")
    rdu_ph <- cSplit(rdu,splitCols = c("ophn","tphn"),sep = ";")[
      ,ophn1:=as.double(str_replace_all(ophn_1,"[-+ ]",""))][
        ,ophn2:=as.double(str_replace_all(ophn_2,"[-+ ]",""))][
          ,tphn1:=as.double(str_replace_all(tphn_1,"[-+ ]",""))][
            ,tphn2:=as.double(str_replace_all(tphn_2,"[-+ ]",""))]
    rdu_exp <- cSplit(rdu_ph,splitCols = c("oemail","temails"),sep = ";",stripWhite = T)
    #rdu_exp[,]
    # rdu_exp<- cSplit(indt = rdu,splitCols = c("Owner_emails","Owner_phones","Tenant_emails","Tenant_phones") ,sep = ";",direction = "wide",drop = T,fixed = T,stripWhite = T)
    # df<- as.data.frame(rdu); df_exp<- as.data.frame(rdu_exp)
    # df[is.na(df)]<-"" ; df_exp[is.na(df_exp)]<-""
    names(x)<-c( "Block" ,    "Flat number",   "Owner_name",   "Tenant_name" ,  "Owner_emails",  "Tenant_emails", "Owner_phones" ,   "Tenant_phones",    "Status" , "Is Residing?","Flat sqft","Sold by","App"  )
    #gs_edit_cells(ss = gs_key(key_rdu),ws = "resid_expanded",input = rdu_exp,anchor = "A1",trim = T)
    x
}


# uncomment to upload new parking updates
# message("Uploading parking worksheet")
# gs_edit_cells(ss = gs_key("1xSd47bmtLjuv2Kb9AdUD0F8eNl5Y5Arc0mwiv_0FXWo"),ws = "parked",input = cars_in_parking[order(flatn)])




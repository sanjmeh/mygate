library(magrittr)
library(officer)
library(data.table)
library(ggplot2)
source("mygate.R")
source("ehbank.r")

read_vislog("visjun.csv") -> vislog_june
read_vislog("visjul.csv") -> vislog_july

d1 <- vislog_june[!is.na(flats) & stayhrs<24 & grepl("^Driver",type,ig=T)]
d2 <- vislog_july[!is.na(flats) & stayhrs<24 & grepl("^Driver",type,ig=T)]

m1 <- vislog_june[!is.na(flats) & stayhrs<24 & grepl("^Maid",type,ig=T)]
m2 <- vislog_july[!is.na(flats) & stayhrs<24 & grepl("^Maid",type,ig=T)]


acc <- split_entry(st_master,m = 3,amt = 85472,pat = "BESCOM",shiftto = 31,ratio = c(1,100)) %>%  # split Bescom double payment
  split_entry(m=7,pat = "LOK INTER",amt = 35293,shiftto = -31,ratio = c(1,4)) %>%  # swimming pool payment to be divided in 3 months
  split_entry(m=6,pat = "LOK INTER",amt = 28234,shiftto = -31,ratio = c(1,1)) %>% # swimming pool payment to be divided in 3 months
  mtrnd(tall = T,tilln = 7)

#acc[,sum(tot_deb),by=maincat][order(-V1)][1:3][,V1]
top3 <- acc[,sum(tot_deb),by=maincat][order(-V1)][1:3][,maincat]
top5 <- acc[,sum(tot_deb),by=maincat][order(-V1)][1:5][,maincat]
top8 <- acc[,sum(tot_deb),by=maincat][order(-V1)][1:8][,maincat]
copy(acc) -> acc_copy
acc3 <- acc[!maincat %in% top3,maincat:="OTHERS"]
copy(acc_copy) -> acc
acc5 <- acc[!maincat %in% top5,maincat:="OTHERS"]

plot_cumul <- bstall %>% 
  .[grepl("cum",variable)] %>% 
  ggplot(aes(mnth,value)) + geom_col(aes(fill=variable),position="dodge") +
  labs(title="CUMULATIVE EXPENSES VS TO CUM MAINT. COLLECTIONS",subtitle ="MONTH ON MONTH",
       caption = "Expenses logged are cash flow not necessarily allocated for work in the same month")

salcomp <- data.table(type=c("security","hk","tech","supr","gard"), # assuming 11 security, 9 HK, 4 tech and 1 supr.
                      ads=c(190000,143480,85347,0,27425),
                      newexp=c(266370,90000,60000,20000,45000))
salcomp2 <- melt(salcomp,id.var="type") 

plot_salcomp <- salcomp2 %>% ggplot(aes(factor(type),value)) + 
  geom_col(aes(fill=variable),position="dodge") + 
  xlab(label = "Category of employee") + ylab("Amount(Rs.)") +
  labs(title="COST COMPARISON",subtitle ="ADS and post ADS era",
       caption = "newexp is ")

plot_majorexp3 <- ggplot(acc3,aes(mnth,tot_deb,fill=maincat)) + 
  geom_col() + geom_hline(yintercept = 880000,lty=2,size=1.5) + 
  geom_text(aes(x="Jan",y=900000),label="8.8 lakhs",size=5) + 
  labs(title="EXPENSES",subtitle ="MONTH ON MONTH",
       caption = "Expenses logged are cash flow not necessarily for work in the same month")

plot_majorexp5 <- ggplot(acc5,aes(mnth,tot_deb,fill=maincat)) + 
  geom_col() + geom_hline(yintercept = 880000,lty=2,size=1.5) + 
  geom_text(aes(x="Jan",y=900000),label="8.8 lakhs",size=5) + 
  labs(title="EXPENSES",subtitle ="MONTH ON MONTH",
       caption = "Expenses logged are cash flow not necessarily for work in the same month")

plot_garden <- acc[grepl("landsc|gard",category,ig=T)] %>%  
  ggplot(aes(month,amt,fill=category)) + 
  geom_col() + geom_hline(yintercept = 50000,lty=2) + 
  geom_text(aes(x="Jan",y=52000),label = "50K - expected steady state expense")+
  labs(title="LANDSCAPING : COMPONENTS",subtitle ="Cash Flow",
       caption = "Expenses logged are cash flow not necessarily for work in the same month")

plot_expincome <- melt(income_expense,id.vars = "mnth") %>% 
  ggplot(aes(mnth,value)) + 
  geom_col(aes(fill=variable),position = "dodge") +
  geom_hline(yintercept = 880000,lty=2,size=2) +
  geom_text(aes(x="Jan",y=900000),label="8.8 lakhs",size=5) + 
  labs(title="EXPENSES vs. INCOME",subtitle ="MONTH ON MONTH - CASH FLOW",
       caption = "Expenses logged are cash flow not necessarily for work in the same month")

# plot_electricity

g1_dr <- ggplot(d1) + 
  geom_point(aes(factor(flats),hour(entry),color = "Entry Time"),position = "jitter",alpha=0.9) +  
  geom_point(aes(factor(flats),hour(exit),color = "Exit Time"),position="jitter",alpha=0.9) + 
  scale_color_manual(values = c("limegreen","steelblue3")) + 
  labs(subtitle = "DRIVER REPORTING AND EXIT TIMES",title= "JUNE-2018",caption="Data taken from MyGate databse as logged by the security")

g2_dr <- ggplot(d2) + 
  geom_point(aes(factor(flats),hour(entry),color = "Entry Time"),position = "jitter",alpha=0.9) +  
  geom_point(aes(factor(flats),hour(exit),color = "Exit Time"),position="jitter",alpha=0.9) + 
  scale_color_manual(values = c("limegreen","steelblue3")) +
  labs(subtitle = "DRIVER REPORTING AND EXIT TIMES",title= "JULY-2018",caption="Data taken from MyGate databse as logged by the security")

g1_maid <- ggplot(m1) + 
  geom_point(aes(factor(flats),hour(entry),color = "Entry Time"),position = "jitter",alpha=0.9) +  
  geom_point(aes(factor(flats),hour(exit),color = "Exit Time"),position="jitter",alpha=0.9) + 
  scale_color_manual(values = c("limegreen","steelblue3")) + 
  labs(subtitle = "MAID REPORTING AND EXIT TIMES",title= "JUNE-2018",caption="Data taken from MyGate databse as logged by the security")

g2_maid <- ggplot(m2) + 
  geom_point(aes(factor(flats),hour(entry),color = "Entry Time"),position = "jitter",alpha=0.9) +  
  geom_point(aes(factor(flats),hour(exit),color = "Exit Time"),position="jitter",alpha=0.9) + 
  scale_color_manual(values = c("limegreen","steelblue3")) + 
  labs(subtitle = "MAID REPORTING AND EXIT TIMES",title= "JULY-2018",caption="Data taken from MyGate databse as logged by the security")



slides <- read_pptx(path = "agm.pptx") %>% 
  add_slide(layout = "Title and Content",master = "Office Theme") %>% 
  ph_with_gg(value = plot_cumul) %>% 
  add_slide(layout = "Title and Content",master = "Office Theme") %>% 
  ph_with_gg(value = plot_majorexp3) %>% 
  add_slide(layout = "Title and Content",master = "Office Theme") %>% 
  ph_with_gg(value = plot_majorexp5) %>% 
  add_slide(layout = "Title and Content",master = "Office Theme") %>% 
  ph_with_gg(value = plot_salcomp) %>% 
  add_slide(layout = "Title and Content",master = "Office Theme") %>% 
  ph_with_table(value = trexp_col) %>% 
  add_slide(layout = "Title and Content",master = "Office Theme") %>% 
  ph_with_table(value = trexp_main) %>% 
  add_slide(layout = "Title and Content",master = "Office Theme") %>% 
  ph_with_table(value = trexp)
  
print(slides, target = "agm.pptx")
  
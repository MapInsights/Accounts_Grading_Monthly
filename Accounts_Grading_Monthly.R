#Account grading
library(RDCOMClient)
library(stringr)
library(dplyr)
library(ggplot2)
library(xlsx)
library(DBI)
library(RODBC)
library(lubridate)

start <- as.Date(Sys.time()) -1
setwd('C:\\Programs\\gtc_tasks\\Accounts_Grading_Monthly')
wd<-getwd()
filename<-paste("Accounts_Grading_",start,".xlsx",sep="")


# Set the directory path for later
spreadsheets_dir <- paste(home_dir,"spreadsheets",sep = "/")

# Make sure you delete the folders and files from last week 
unlink("spreadsheets", recursive = TRUE, force = FALSE)

# Create a directory for spreadsheets otherwise R having a heart attack 
dir.create("spreadsheets",showWarnings = F)



odbcChannel <- odbcConnect('echo_core')
data  <- sqlQuery( odbcChannel, "select isnull(ca.parent_id,ca.id) 'id',isnull(cap.name,ca.name) 'name',isnull(cap.number,ca.number)'number',sum(j.totalnetprice) 'totalnetprice',sum(j.totalCharge)'totalCharge', count(j.id) 'JobCount'
                   from echo_core_prod..jobs j
                   left join echo_core_prod..customer_accounts ca on ca.id = j.customer_account_id
                   left join echo_core_prod..customer_accounts cap on cap.id = ca.parent_id
                   where j.jobdate > DATEADD(month,-6,getdate())-- '2015-12-21'
                   and j.jobStatus in (7,10)
                   and ca.number not in ('G51','G10','LHR','G50','G56','G9','1','G6','GTC888','LONGTC1387','G9.5','G60','LHR cash','G53','G4','G57','GTC888','G9.1','G4.1','G4')
                   
                   group by  isnull(ca.parent_id,ca.id),isnull(cap.name,ca.name) ,isnull(cap.number,ca.number)
                   ")

grade <- sqlQuery( odbcChannel, "

select distinct isnull(ca.parent_id,ca.id) 'id',isnull(parent.grade_id,ca.grade_id) 'gradeId', cg.name
                   from Echo_core_prod..customer_accounts ca
                   left join echo_core_prod..customer_accounts parent on parent.id = ca.parent_id
                   left join Echo_core_prod..customer_grades cg on ca.grade_id = cg.id
                   --left join Echo_core_prod..customer_grades cgp on parent.grade_id = cg.id
                   ")


opendate <- sqlQuery( odbcChannel, 
"select ca.id, ca.name, ca.number, ca.parent_id,parent.name, parent.number,ca.grade_id,parent.grade_id 'ParentGrade',ca.dateOpened, parent.dateOpened 'ParentOpened'
from Echo_core_prod..customer_accounts ca
                      left join Echo_core_prod..customer_accounts parent on parent.id = ca.parent_id
                     ")

odbcClose(odbcChannel)

opendate[is.na(opendate$ParentOpened),"ParentOpened"]<-opendate[is.na(opendate$ParentOpened),"dateOpened"]
opendate[is.na(opendate$parent_id),"parent_id"]<-opendate[is.na(opendate$parent_id),"id"]


opendateslim<-opendate[,c(4,10)]

#use group by to get first date opened, and do same for grade.s
opendateslim2 <- group_by(opendate,parent_id) %>% summarise(open = min(ParentOpened))
grade2<-group_by(grade,id) %>% summarise(grade= first(name))


datamerge<-merge(data,opendateslim2,by.x = "id",by.y="parent_id" ,all.x = TRUE )

datamerge<-merge(datamerge,grade2, by ="id",all.x=TRUE)


datamerge$DaysOpen<-as.numeric(round(difftime(start,datamerge$open,units = "days"),0))

datamerge$DailyJobs<-datamerge$JobCount/pmin(as.numeric(datamerge$DaysOpen),365/2)
datamerge$AveFare<-datamerge$totalnetprice / datamerge$JobCount

datamerge$totalMargin<-datamerge$totalnetprice - datamerge$totalCharge
datamerge$MarginDay <- datamerge$totalMargin / pmin(as.numeric(datamerge$DaysOpen),365/2)
datamerge$MarginPct <- datamerge$totalMargin / datamerge$totalnetprice



#Calc points
datamerge$points<- datamerge$MarginDay * datamerge$AveFare / 1000 * 365

AllTrips<- sum(datamerge$DailyJobs)
datamerge$PctTrips <- datamerge$DailyJobs / AllTrips

final<-datamerge[order( datamerge$points, decreasing=TRUE),]
final$AcumTripsPct<-cumsum(final$PctTrips)

final$NewGrade<-1
final[final$AcumTripsPct<=0.8,"NewGrade"]<- 2
final[final$AcumTripsPct<=0.6,"NewGrade"]<- 3
final[final$AcumTripsPct<=0.4,"NewGrade"]<- 4
final[final$AcumTripsPct<=0.2,"NewGrade"]<- 5


write.xlsx(final,file = paste("spreadsheets",filename,sep="/"))


#Now send to distribution list

string<-paste(wd,"spreadsheets",filename,sep="/")
string
#Now send it to Lee

# Send mail for 3D

library(RDCOMClient)
OutApp <- COMCreate("Outlook.Application")
outMail = OutApp$CreateItem(0)

#OutApp <- COMCreate("Outlook.Application")
#outMail = OutApp$CreateItem(0)
outMail[["subject"]] = 'Monthly Account Grading'
outMail[["To"]] = "haider.variava@greentomatocars.com;daria.alekseeva@greentomatocars.com;antony.carolan@greentomatocars.com;ian.bates@greentomatocars.com;james.rowe@greentomatocars.com;sean.sauter@greentomatocars.com"
#outMail[["To"]] = "antony.carolan@greentomatocars.com"
outMail[["body"]] ="Hi, Monthly account grading report is attached. Antony"
outMail[["Attachments"]]$Add(string)

outMail$Send()
rm(list = c("OutApp","outMail"))

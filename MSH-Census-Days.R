#--------------------Census Days
##Make sure to change dates in save file within Export Function and rerun the export function
#-----Reads and manipulates raw file to be ready to append master file
rawfile <- function(){
  #Read raw file
  raw <- readxl::read_excel(file.choose())
  #Only take necessary columns
  raw1 <- data.frame(Unit= raw$UNIT,UnitDescription=raw$`UNIT DESCRIPTION`,
                     MonthYear=raw$`MONTH YEAR`,Date=raw$DATE,Census=raw$`Census for the Day`)
  #Filter out blank or zero rows
  raw2 <- subset(raw1,raw1$Census!=  | !is.na(raw1$Census))
  #Check how many days are in the file. Stop if not equal to 14
  filedays <- length(unique(raw2$Date))
  raw2<<-raw2
  stopifnot(filedays==14)
}

#-----Append raw file to a master file after the rawfile function
master <- function(){
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/BETA Test - Premier Tracking/Productivity/Volume - Data/MSH Data/Inpatient Census Days/Calculation Worksheets")
  #upload current master file
  Masterold <<- read.csv("Master.csv",check.names = F)
  #turn raw2 date in factor
  raw2$Date <- as.factor(raw2$Date)
  #append the current master with the new raw file to make the new master
  Masternew <<- rbind(Masterold,raw2)
  #overwrite the new master over the old master
  write.csv(Masternew,"Master.csv", row.names = F)
}

#-----Trend by Month
Monthly <- function(){
  library(lubridate)
  #adjust the date format of the new master table
  Masternew$MonthYear <- dmy(paste0("01/",Masternew$MonthYear))
  #find all unique month-year pairs in the new master table
  months <- unique(Masternew$MonthYear)
  #aggregate total patient days by month-year pair
  trend <<- aggregate(Census~MonthYear,Masternew,sum)
  #print the trend
  trend
}

#-----Upload Cost Center Crosswalk and assign cost center to each department
VolumeID <- function(){
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/BETA Test - Premier Tracking/Productivity/Volume - Data/MSH Data/Inpatient Census Days/Calculation Worksheets")
  #upload volume ID crosswalk with every department and their respective volumeID
  Crosswalk <- read.csv("VolumeID.csv", check.names = F, colClasses = c(rep("factor",4)))
  #merge volumeID and raw file. Assigns CC and volume ID to each unit in raw file
  PP <- merge(Crosswalk,raw2, by.x="UNIT", by.y="Unit")
  #aggregate merged table by cost center to get biweekly sum
  #PP2 <<- aggregate(Census~`Cost Center`,PP,sum)
  PP2 <<- aggregate(Census~VolumeID,PP,sum)
}

#-----Create census upload for PP
#MAKE SURE TO CHANGE DATE IN SAVE TITLE
Export <- function(date1,date2,file){
  #Create upload format
  upload <<- data.frame(partner="729805", hospital="NY0014", CC=substr(PP2$VolumeID,start=1,stop=8),
                        start=date1, end=date2, volumeID=PP2$VolumeID, Census=PP2$Census, budget="0")
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/BETA Test - Premier Tracking/Productivity/Volume - Data/MSH Data/Inpatient Census Days/Calculation Worksheets")
  #save export
  write.table(upload,file=file,sep=",",col.names=F,row.names=F)
}

#-----------------------------------------------------------------------------------------------

#Execute Functions
rawfile()
master()
Monthly()
VolumeID()
##date1 should equal the first date of the PP 
##date2 should equal the last date of the PP
Export(date1="10/27/2019",date2="11/09/2019",file="MSH_Census Days_27OCT2019 to 09NOV2019.csv")

#In case you are doing multiple uploads
rm(Masternew,Masterold,PP2,raw,raw1,raw2,trend,upload,filedays,PP,Crosswalk)
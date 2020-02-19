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
  raw2 <- subset(raw1,!is.na(raw1$Census))
  #Check how many days are in the file. Stop if not equal to 14
  filedays <- length(unique(raw2$Date))
  raw2<<-raw2
  stopifnot(filedays==14)
  print(min(raw2$Date))
  print(max(raw2$Date))
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
  MonthTrend <<- aggregate(Census~MonthYear,Masternew,sum)
  #print the trend
  MonthTrend
}

#-----Upload Cost Center Crosswalk and assign cost center to each department
VolumeID <- function(date){
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/BETA Test - Premier Tracking/Productivity/Volume - Data/MSH Data/Inpatient Census Days/Calculation Worksheets")
  #upload volume ID crosswalk with every department and their respective volumeID
  Crosswalk <- read.csv("VolumeID.csv", check.names = F, colClasses = c(rep("factor",4)))
  #merge volumeID and raw file. Assigns CC and volume ID to each unit in raw file
  PP <- merge(Crosswalk,raw2, by.x="UNIT", by.y="Unit")
  #aggregate merged table by cost center to get biweekly sum
  #PP2 <<- aggregate(Census~`Cost Center`,PP,sum)
  PP2 <<- aggregate(Census~VolumeID,PP,sum)
  #Set directory
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/BETA Test - Premier Tracking/Productivity/Volume - Data/MSH Data/Inpatient Census Days/Calculation Worksheets")
  #Read current PP trend file
  Current <- read.csv(file="PP_Trend.csv",check.names = F)
  #Read current while maintaining CC and VolumeID
  PPTrend <- read.csv(file="PP_Trend.csv", check.names = F, colClasses = c(rep("factor",length(Current))))
  #Append current PP census to PP Trend
  PPTrend_new <- cbind(PPTrend,PP2$Census)
  #adjust column name to the end date of this payperiod
  colnames(PPTrend_new)[length(PPTrend_new)] <- date
  #Display PP Trend
  PPTrend_new <<- PPTrend_new
  #Prompt user to make overwrite decision
  Decision <- readline(prompt="Overwrite PP Trend? (yes/no): ")
  #If user answers yes then overwrite the trend file
  if(Decision == "yes"){
    setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/BETA Test - Premier Tracking/Productivity/Volume - Data/MSH Data/Inpatient Census Days/Calculation Worksheets")
    write.csv(PPTrend_new,file="PP_Trend.csv", row.names = F)
  }
}

#-----Create census upload for PP
#MAKE SURE TO CHANGE DATE IN SAVE TITLE
Export <- function(date1,date2,file){
  #Create upload format
  upload <<- data.frame(partner="729805", hospital="NY0014", CC=substr(PP2$VolumeID,start=1,stop=8),
                        start=date1, end=date2, volumeID=PP2$VolumeID, Census=PP2$Census, budget="0",
                        Unit = PPTrend_new$Department, Volume = "Patient Days")
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/BETA Test - Premier Tracking/Productivity/Volume - Data/MSH Data/Inpatient Census Days/Calculation Worksheets")
  #save export
  write.table(upload,file=file,sep=",",col.names=F,row.names=F)
}

#-----------------------------------------------------------------------------------------------

#Execute Functions
rawfile()
master()
Monthly()
#date is the end date of the pp file
VolumeID("11/23/2019")
##date1 should equal the first date of the PP 
##date2 should equal the last date of the PP
Export(date1="11/10/2019",date2="11/23/2019",file="MSH_Census Days_10NOV2019 to 23NOV2019.csv")

#In case you are doing multiple uploads
rm(Masternew,Masterold,PP2,raw,raw1,raw2,trend,upload,filedays,PP,Crosswalk)
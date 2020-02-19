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
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/MSH Data/Inpatient Census Days/Calculation Worksheets")
  #upload current master file
  Masterold <- readxl::read_excel("Master.xlsx", col_types = c("text", "text", "date", "date", "numeric"))
  library(lubridate)
  #adjust the monthyear format of the raw table to prepare for append
  raw2$MonthYear <- dmy(paste0("19/",raw2$MonthYear))
  #put monthyear in same format as master table
  raw2$MonthYear <- as.POSIXct(raw2$MonthYear)
  #append the current master with the new raw file to make the new master
  Masternew <<- rbind(Masterold,raw2)
}

#-----Upload Cost Center Crosswalk and assign cost center to each department
VolumeID <- function(date){
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/MSH Data/Inpatient Census Days/Calculation Worksheets")
  #upload volume ID crosswalk with every department and their respective volumeID
  Crosswalk <- read.csv("VolumeID.csv", check.names = F, colClasses = c(rep("factor",4)))
  #merge volumeID and raw file. Assigns CC and volume ID to each unit in raw file
  PP <- merge(Crosswalk,raw2, by.x="UNIT", by.y="Unit")
  #aggregate merged table by cost center to get biweekly sum
  PP2 <<- aggregate(Census~`Cost Center`,PP,sum)
  #PP2 <<- aggregate(Census~UNIT + `UNIT DESCRIPTION` + `Cost Center` + VolumeID,PP,sum)
  #Set directory
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/MSH Data/Inpatient Census Days/Calculation Worksheets")
  #Read current while maintaining CC and VolumeID
  PPTrend <- read.csv(file="PP_Trend.csv", check.names = F, colClasses = c(rep("factor",length(PP2))))
  #Append current PP census to PP Trend
  PPTrend_new <- cbind(PPTrend,PP2$Census)
  #adjust column name to the end date of this payperiod
  colnames(PPTrend_new)[length(PPTrend_new)] <- date
  #Display PP Trend
  PPTrend_new <<- PPTrend_new
}

#-----Create census upload for PP
#MAKE SURE TO CHANGE DATE IN SAVE TITLE
Export <- function(date1,date2){
  #Create upload format
  upload <<- data.frame(partner="729805", hospital="NY0014", CC=PP2$`Cost Center`,start=date1, 
                        end=date2, volumeID=paste0(PP2$`Cost Center`,"1"), Census=PP2$Census, budget="0",
                        Unit = PPTrend_new$Department, Volume = "Patient Days")
}

Save <- function(file){
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/MSH Data/Inpatient Census Days/Calculation Worksheets")
  #overwrite the new master over the old master
  write.csv(Masternew,"Master.csv", row.names = F)
  #save export
  write.table(upload,file=file,sep=",",col.names=F,row.names=F)
  #save PP Trend
  write.csv(PPTrend_new,file="PP_Trend.csv", row.names = F)
}

#-----------------------------------------------------------------------------------------------

#Execute Functions
rawfile()
master()
#date is the end date of the pp file
VolumeID("11/23/2019")
##date1 should equal the first date of the PP 
##date2 should equal the last date of the PP
Export(date1="11/10/2019",date2="11/23/2019")
#Check Masternew, upload and PPTrend_new tables
Save(file="MSH_Census Days_10NOV2019 to 23NOV2019.csv")


#In case you are doing multiple uploads
rm(Masternew,Masterold,PP2,raw,raw1,raw2,trend,upload,filedays,PP,Crosswalk)
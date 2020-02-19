#--------------------Census Days For PP
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
  MinDate <<- print(min(raw2$Date))
  MaxDate <<- print(max(raw2$Date))
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
VolumeID <- function(){
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
  colnames(PPTrend_new)[length(PPTrend_new)] <- as.character(paste0(substr(MaxDate,start=6,stop=7),"/",substr(MaxDate,start=9,stop=10),"/",substr(MaxDate,start=1,stop=4)))
  #Display PP Trend
  PPTrend_new <<- PPTrend_new
}

#-----Create census upload for PP
Export <- function(){
  #Create upload format
  ExportMax <- paste0(substr(MaxDate,start=6,stop=7),"/",substr(MaxDate,start=9,stop=10),"/",substr(MaxDate,start=1,stop=4))
  ExportMin <- paste0(substr(MinDate,start=6,stop=7),"/",substr(MinDate,start=9,stop=10),"/",substr(MinDate,start=1,stop=4))
  upload <<- data.frame(partner="729805", hospital="NY0014", CC=PP2$`Cost Center`,start=ExportMin, 
                        end=ExportMax, volumeID=paste0(PP2$`Cost Center`,"1"), Census=PP2$Census, budget="0",
                        Unit = PPTrend_new$Department, Volume = "Patient Days")
}

Save <- function(){
  setwd("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/MSH Data/Inpatient Census Days/Calculation Worksheets")
  library(lubridate)
  smonth <- toupper(month.abb[month(MinDate)])
  emonth <- toupper(month.abb[month(MaxDate)])
  sday <- format(as.Date(MinDate, format="%Y-%m-%d"), format="%d")
  eday <- format(as.Date(MaxDate, format="%Y-%m-%d"), format="%d")
  syear <- substr(MinDate, start=1, stop=4)
  eyear <- substr(MaxDate, start=1, stop=4)
  name <- paste0("MSH_Census Days_",sday,smonth,syear," to ",eday,emonth,eyear,".csv")
  #save export
  write.table(upload,file=name,sep=",",col.names=F,row.names=F)
  #save PP Trend
  write.csv(PPTrend_new,file="PP_Trend.csv", row.names = F)
  #overwrite the new master over the old master
  MasterName <- paste0("Master - ",eday,emonth,eyear)
  write.xlsx(as.data.frame(Masternew), "Master.xlsx", col.names = T, row.names = F, sheetName = MasterName )
}

#-----------------------------------------------------------------------------------------------

#Execute Functions
rawfile()
master()
#Upload Cost Center Crosswalk and assign cost center to each department
VolumeID()
#Create census upload for PP
Export()
#Check Masternew, upload and PPTrend_new tables
Save()


#In case you are doing multiple uploads
rm(Masternew,Masterold,PP2,raw,raw1,raw2,trend,upload,filedays,PP,Crosswalk)
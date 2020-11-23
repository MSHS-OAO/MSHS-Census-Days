#--------------------Census Days For PP
library(dplyr)
library(tidyr)
library(readxl)
library(zoo)

#--------------raw
raw <- read_xlsx(file.choose(),skip = 4)
census <- raw %>%
  select(-KP7...40) %>%
  rename(KP7 = KP7...50) %>%
  filter(!is.na(Date),
         Date != "Total") %>%
  mutate(Date = as.numeric(Date),
         Weekday = as.numeric(Weekday)) %>%
  mutate(Date = as.Date(Date, origin = "1899-12-30"),
         Weekday = as.Date(Weekday, origin = "1899-12-30")) 
census <- census %>%
  pivot_longer(names_to = "DepID", cols = 3:ncol(census), values_to = "Census")
#add check for dates to make sure whole month


#------------volumeID
volID <- read.csv("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/MSH Data/Inpatient Census Days/New Calculation Worksheets/volID.csv",
                  stringsAsFactors = F,colClasses = c(rep("character",6)))
paycycle <- read_xlsx("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/MSH Data/Inpatient Census Days/New Calculation Worksheets/Pay Cycle Calendar.xlsx") %>%
  select(Date,Start.Date,End.Date) %>%
  mutate(Date = as.Date(Date),
         Start.Date = as.Date(Start.Date),
         End.Date = as.Date(End.Date))
census_vol <- left_join(census,volID,by=c("DepID"="UNIT")) %>%
  left_join(paycycle,by=("Date"="Date"))

#------------master
#make sure to check census_vol does not overlap with Master.RDS
master_old <- readRDS("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/MSH Data/Inpatient Census Days/New Calculation Worksheets/Master.RDS")
if(max(master_old$Date) < min(census_vol$Date)){
  master_new <- rbind(master_old,census_vol)
  saveRDS(master_new,"J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/MSH Data/Inpatient Census Days/New Calculation Worksheets/Master.RDS")
}
trend <- master_new %>% 
  ungroup()%>%
  group_by(Oracle,End.Date) %>%
  summarise(Census = sum(Census, na.rm = T)) %>%
  pivot_wider(id_cols = c(Oracle),names_from = End.Date,values_from = Census)


#------------upload
#user input to filter master on neccessary pay period
upload <- function(start,end){
  upload <- master_new %>%
    filter(Start.Date >= as.Date(start,format = "%m/%d/%Y"),
           End.Date <= as.Date(end,format = "%m/%d/%Y"),
           !is.na(Oracle)) %>%
    group_by(Oracle,OracleVol,Start.Date,End.Date) %>%
    summarise(Census = sum(Census, na.rm = T)) %>%
    mutate(Partner = "729805",
           Hosp = "NY0014",
           Budget = "0") %>%
    select(Partner,Hosp,Oracle,Start.Date,End.Date,OracleVol,Census,Budget) %>%
    ungroup() %>%
    mutate(Start.Date = format(Start.Date,format = "%m/%d/%Y"),
           End.Date = format(End.Date,format = "%m/%d/%Y"))
  return(upload)
}

census_export <- upload(start = "09/27/2020",end = "10/24/2020")
write.table(census_export,"J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/MSH Data/Inpatient Census Days/New Calculation Worksheets/Uploads/MSH_Census_27SEP2020 to 24OCT2020.csv",col.names = F,row.names = F,sep = ",")
trend <- census_export %>% pivot_wider(id_cols = Oracle,names_from = End.Date,values_from = Census)

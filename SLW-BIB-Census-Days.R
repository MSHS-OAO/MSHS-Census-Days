dir <- 'J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/Census Days'
setwd(dir)

# Load Libraries ----------------------------------------------------------
library(tidyverse)
library(xlsx)

# User Input --------------------------------------------------------------
warning("Update Pay Period Start and End Dates Needed:")
Sys.sleep(2)
pp.start <- as.Date('2020-06-21') # start date of first pay period needed
pp.end <- as.Date('2020-08-01') # end date of the last pay period needed
cat(paste("Pay period starting on",format(pp.start, "%m/%d/%Y"), 'and ending on',format(pp.end, "%m/%d/%Y") ),fill = T)

# Import Dictionaries -------------------------------------------------------
map_CC_Vol <-  read.xlsx('BIBSLW_Volume ID_Cost Center_ Mapping.xlsx', sheetIndex = 1)
dict_PC <- read.xlsx('Pay Cycle Dictionaries.xlsx', sheetIndex = 1)
colnames(dict_PC)[1] <- 'Census.Date'
dict_PC <- dict_PC %>% drop_na()

# Import Data -------------------------------------------------------------
data_census <- read.xlsx(choose.files(caption = "Select Census File", multi = F, default= 'J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/Census Days/Source Data'), sheetIndex = 1)

# QC ----------------------------------------------------------------------
if(pp.end > range(data_census$Census.Date)[2]){stop('Data Missing from Census for Pay Periods Needed')}
if(range(data_census$Census.Date)[2] > range(dict_PC$End.Date)[2]){stop("Update Pay Cycle Dictionary")}

# Pre Processing ----------------------------------------------------------
data_census <- data_census %>%
  mutate(Site = as.character(Site),
        Nursing.Station.Code = as.character(Nursing.Station.Code))
map_CC_Vol <- map_CC_Vol %>%
  mutate(Site = as.character(Site),
         Nursing.Station.Code = as.character(Nursing.Station.Code))
data_upload <- left_join(data_census, drop_na(subset(map_CC_Vol, select = c('Site', 'Nursing.Station.Code', 'CostCenter', 'VolumeID')))) #issue this is adding rows
data_upload <- left_join(data_upload, dict_PC)

# MSW Upload File ---------------------------------------------------------
data_upload_MSW <- data_upload %>%
  as.data.frame() %>%
  filter(Census.Date >= pp.start,
         Census.Date <= pp.end,
         Site == 'RVT') %>%
  mutate(Corp = 729805,
         Site = 'NY2162',
         Start.Date = format(Start.Date, "%m/%d/%Y"),
         End.Date = format(End.Date, "%m/%d/%Y")) %>%
  select(Corp, Site, CostCenter, Start.Date, End.Date, VolumeID, Census.Day) %>%
  group_by(Corp, Site, CostCenter, Start.Date, End.Date, VolumeID) %>%
  summarise(Volume = sum(Census.Day, na.rm=T)) %>%
  mutate(Budget = 0)
data_upload_MSW <- na.omit(data_upload_MSW)

# MSM Upload File ---------------------------------------------------------
data_upload_MSM <- data_upload %>%
  as.data.frame() %>%
  filter(Census.Date >= pp.start,
         Census.Date <= pp.end,
         Site == 'STL') %>%
  mutate(Corp = 729805,
         Site = 'NY2163',
         Start.Date = format(Start.Date, "%m/%d/%Y"),
         End.Date = format(End.Date, "%m/%d/%Y")) %>%
  select(Corp, Site, CostCenter, Start.Date, End.Date, VolumeID, Census.Day) %>%
  group_by(Corp, Site, CostCenter, Start.Date, End.Date, VolumeID) %>%
  summarise(Volume = sum(Census.Day, na.rm=T)) %>%
  mutate(Budget = 0)
data_upload_MSM <- na.omit(data_upload_MSM)

# MSBIB Upload File ---------------------------------------------------------
data_upload_MSBIB <- data_upload %>%
  as.data.frame() %>%
  filter(Census.Date >= pp.start,
         Census.Date <= pp.end,
         Site == 'BIB' | Site == 'BIPTR') %>%
  mutate(Corp = 729805,
         Site = '630571',
         Start.Date = format(Start.Date, "%m/%d/%Y"),
         End.Date = format(End.Date, "%m/%d/%Y")) %>%
  select(Corp, Site, CostCenter, Start.Date, End.Date, VolumeID, Census.Day) %>%
  group_by(Corp, Site, CostCenter, Start.Date, End.Date, VolumeID) %>%
  summarise(Volume = sum(Census.Day, na.rm=T)) %>%
  mutate(Budget = 0)
data_upload_MSBIB <- na.omit(data_upload_MSBIB)

# Export Files ------------------------------------------------------------
setwd(paste0(dir, '/Upload Files'))
write.table(data_upload_MSW, file = paste0("MSW_Census Days_", format(pp.start,"%d%b%y"), " to ", format(pp.end, "%d%b%y"), ".csv"), sep = ',' , row.names = F,col.names = F)
write.table(data_upload_MSM, file = paste0("MSM_Census Days_", format(pp.start,"%d%b%y"), " to ", format(pp.end, "%d%b%y"), ".csv"), sep = ',', row.names = F, col.names = F)
write.table(data_upload_MSBIB, file = paste0("MSBIB_Census Days_", format(pp.start,"%d%b%y"), " to ", format(pp.end, "%d%b%y"), ".csv"), sep = ',', row.names = F, col.names = F)

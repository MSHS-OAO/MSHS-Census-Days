dir <- 'J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/Census Days'
#setwd(dir)

# Load Libraries ----------------------------------------------------------
library(tidyverse)
library(xlsx)

# User Input --------------------------------------------------------------
pp.start <- as.Date('2020-08-02') # start date of first pay period needed
pp.end <- as.Date('2020-08-29') # end date of the last pay period needed
warning("Update Pay Periods Start and End Dates Needed:")
cat(paste("Pay period starting on",format(pp.start, "%m/%d/%Y"), 'and ending on',format(pp.end, "%m/%d/%Y") ),fill = T)
Sys.sleep(2)

# Import Dictionaries -------------------------------------------------------
map_CC_Vol <-  read.xlsx(paste0(dir, '/BIBSLW_Volume ID_Cost Center_ Mapping.xlsx'), sheetIndex = 1)
dict_PC <- read.xlsx(paste0(dir,'/Pay Cycle Dictionaries.xlsx'), sheetIndex = 1)
colnames(dict_PC)[1] <- 'Census.Date'
dict_PC <- dict_PC %>% drop_na()

# Import Data -------------------------------------------------------------
import_recent_file <- function(folder.path) {
  File.Name <- list.files(path = folder.path,pattern = 'xlsx$',full.names = F)
  File.Path <- list.files(path = folder.path,pattern = 'xlsx$',full.names = T)
  File.Date <- as.Date(sapply(File.Name, function(x) substr(x,nchar(x)-12, nchar(x)-5)),format = '%m.%d.%y')
  File.Table <- data.table::data.table(File.Name, File.Date, File.Path)  %>%
    arrange(desc(File.Date))
  data_recent <- read.xlsx(file = File.Table$File.Path[1], sheetIndex = 1)
  return(data_recent)
}
data_census <- import_recent_file(paste0(dir, '/Source Data'))
#data_census <- read.xlsx(choose.files(caption = "Select Census File", multi = F, default= 'J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/Census Days/Source Data'), sheetIndex = 1)

# QC ----------------------------------------------------------------------
if(pp.end > range(data_census$Census.Date)[2]){stop('Data Missing from Census for Pay Periods Needed')}
if(range(data_census$Census.Date)[2] > range(dict_PC$End.Date)[2]){stop("Update Pay Cycle Dictionary")}

# Pre Processing ----------------------------------------------------------
data_census <- data_census %>%
  mutate(Site = as.character(Site),
         Nursing.Station.Code = as.character(Nursing.Station.Code))
map_CC_Vol <- map_CC_Vol %>%
  select (Site, Nursing.Station.Code, CostCenter, VolumeID) %>%
  mutate(Nursing.Station.Code = as.character(Nursing.Station.Code)) %>%
  drop_na()
data_upload <- left_join(data_census, map_CC_Vol)
data_upload <- left_join(data_upload, dict_PC)

# Upload File Creation ----------------------------------------------------
upload_file <- function(site.census, site.premier){
  upload <- data_upload %>%
    as.data.frame() %>%
    filter(Census.Date >= pp.start,
           Census.Date <= pp.end,
           Site == site.census) %>%
    mutate(Corp = 729805,
           Site = site.premier,
           Start.Date = format(Start.Date, "%m/%d/%Y"),
           End.Date = format(End.Date, "%m/%d/%Y")) %>%
    select(Corp, Site, CostCenter, Start.Date, End.Date, VolumeID, Census.Day) %>%
    group_by(Corp, Site, CostCenter, Start.Date, End.Date, VolumeID) %>%
    summarise(Volume = sum(Census.Day, na.rm=T)) %>%
    mutate(Budget = 0)
  upload <- na.omit(upload)
  return(upload)
}
data_upload_MSW <- upload_file('RVT', 'NY2162')
data_upload_MSM <- upload_file('STL', 'NY2163')
data_upload_MSBIB <- rbind(upload_file('BIB','630571'), upload_file('BIPTR','630571'))

# Export Files ------------------------------------------------------------
setwd(paste0(dir, '/Upload Files'))
write.table(data_upload_MSW, file = paste0("MSW_Census Days_", format(pp.start,"%d%b%y"), " to ", format(pp.end, "%d%b%y"), ".csv"), sep = ',' , row.names = F,col.names = F)
write.table(data_upload_MSM, file = paste0("MSM_Census Days_", format(pp.start,"%d%b%y"), " to ", format(pp.end, "%d%b%y"), ".csv"), sep = ',', row.names = F, col.names = F)
write.table(data_upload_MSBIB, file = paste0("MSBIB_Census Days_", format(pp.start,"%d%b%y"), " to ", format(pp.end, "%d%b%y"), ".csv"), sep = ',', row.names = F, col.names = F)

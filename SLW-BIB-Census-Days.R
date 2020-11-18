dir <- 'J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/Census Days'

# Load Libraries ----------------------------------------------------------
library(tidyverse)
library(xlsx)

# User Input --------------------------------------------------------------
pp.start <- as.Date('2020-09-27') # start date of first pay period needed
pp.end <- as.Date('2020-10-24') # end date of the last pay period needed
warning("Update Pay Periods Start and End Dates Needed:")
cat(paste("Pay period starting on",format(pp.start, "%m/%d/%Y"), 'and ending on',format(pp.end, "%m/%d/%Y") ),fill = T)
Sys.sleep(2)

# Constants ---------------------------------------------------------------
#Names of sites in census files (old, new)
site_MSW <- c('RVT', 'MSW')
site_MSM <- c('STL', 'MSM')
site_MSBI <- c('BIPTR', 'MSBITR')
site_MSB <- c('BIB', 'MSB')

# Import Dictionaries -------------------------------------------------------
map_CC_Vol <-  read.xlsx(paste0(dir, '/BIBSLW_Volume ID_Cost Center_ Mapping.xlsx'), sheetIndex = 1)
dict_PC <- read.xlsx(paste0(dir,'/Pay Cycle Dictionaries.xlsx'), sheetIndex = 1)
colnames(dict_PC)[1] <- 'Census.Date'
dict_PC <- dict_PC %>% drop_na()

# Import Data -------------------------------------------------------------
import_recent_file <- function(folder.path, place) {
  #Importing File information from Folder
  File.Name <- list.files(path = folder.path,pattern = 'xlsx$',full.names = F)
  File.Path <- list.files(path = folder.path,pattern = 'xlsx$',full.names = T)
  File.Date <- as.Date(sapply(File.Name, function(x) substr(x,nchar(x)-12, nchar(x)-5)),format = '%m.%d.%y')
  File.Table <<- data.table::data.table(File.Name, File.Date, File.Path) %>%
    arrange(desc(File.Date))
  #Importing Data 
  data_recent <- read.xlsx(file = File.Table$File.Path[place], sheetIndex = 1)
  data_recent <- data_recent %>% mutate(Source = File.Table$File.Path[place]) #File Source Column for Reference
  return(data_recent)
}
data_census <- import_recent_file(paste0(dir, '/Source Data'), 1)
#data_census <- read.xlsx(choose.files(caption = "Select Census File", multi = F, default= 'J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/Census Days/Source Data'), sheetIndex = 1)
#if(format(pp.start, "%Y") != format(pp.end, "%Y")){
#  File.Table2 <- File.Table %>% mutate(File.Year = format(File.Date,'%Y'))
#  File.Table2 <- File.Table2[nrow(File.Table2)-sum(File.Table2$File.Year == format(pp.start,"%Y")),]
 # data_census2  <- read.xlsx(file = File.Table2$File.Path, sheetIndex = 1)
 # data_census2 <- data_census2 %>% mutate(Soruce = File.Table2$File.Path)
 # data_census <- rbind(data_census, data_census2) %>% distinct()
 # }

# QC ----------------------------------------------------------------------
if(pp.end > range(data_census$Census.Date)[2]){stop('Data Missing from Census for Pay Periods Needed')}
if(pp.start < range(data_census$Census.Date)[1]){
  data_census2 <- import_recent_file(paste0(dir, '/Source Data'), 2)
  if(any(!unique(data_census$Site) %in% unique(data_census2$Site))){
    #If the site names are not the same update the 2nd census file to the new names
    site.Table <- as.data.frame(rbind(site_MSM, site_MSW, site_MSBI, site_MSB))
    colnames(site.Table) <- c('Site', 'Site.New')
    site.Table <- site.Table %>% mutate(Site = as.character(Site))
    data_census2 <- left_join(data_census2, site.Table)
    data_census2 <- data_census2 %>% mutate(Site = NULL) %>% rename(Site = Site.New)
    #Combine the 1st and 2nd census file
    data_census <- rbind(data_census, data_census2) %>% mutate(Source = NULL) %>% distinct()
  } else{
    #If the sites match combine the 1st and 2nd census file
    data_census <- rbind(data_census, data_census2) %>% mutate(Source = NULL) %>% distinct()
  }
}#If payperiod crosses a year import 2nd census file from previous year
if(range(data_census$Census.Date)[2] > range(dict_PC$End.Date)[2]){stop("Update Pay Cycle Dictionary")}

# Pre Processing ----------------------------------------------------------
data_census <- data_census %>%
  mutate(Site = as.character(Site),
         Nursing.Station.Code = as.character(Nursing.Station.Code))
map_CC_Vol <- map_CC_Vol %>%
  select (Site, Nursing.Station.Code, CostCenter, VolumeID) %>%
  mutate(Nursing.Station.Code = as.character(Nursing.Station.Code)) %>%
  drop_na()
if(any(!unique(map_CC_Vol$Site) %in% unique(data_census$Site))){
  site.Table <- as.data.frame(rbind(site_MSM, site_MSW, site_MSBI, site_MSB))
  colnames(site.Table) <- c('Site', 'Site.New')
  site.Table <- site.Table %>% mutate(Site = as.character(Site))
  map_CC_Vol <- left_join(map_CC_Vol, site.Table)
  map_CC_Vol <- map_CC_Vol %>% mutate(Site = NULL) %>% rename(Site = Site.New)
} #depending on old or new names used for sites update dictionary
data_upload <- left_join(data_census, map_CC_Vol)
data_upload <- left_join(data_upload, dict_PC)

# QC --------------------------------------------------------------------
if(nrow(data_census) != nrow(data_upload)){stop('Check Dictionaries for Duplicates')} #checking to see if vlookups are duplicating rows

# Upload File Creation ----------------------------------------------------
upload_file <- function(site.census, site.premier, map_cc){
  upload <- data_upload %>%
    as.data.frame() %>%
    filter(Census.Date >= pp.start,
           Census.Date <= pp.end,
           Site %in% site.census) %>%
    mutate(Corp = 729805,
           Site = site.premier,
           Start.Date = format(Start.Date, "%m/%d/%Y"),
           End.Date = format(End.Date, "%m/%d/%Y")) %>%
    select(Corp, Site, CostCenter, Start.Date, End.Date, VolumeID, Census.Day) %>%
    group_by(Corp, Site, CostCenter, Start.Date, End.Date, VolumeID) %>%
    summarise(Volume = sum(Census.Day, na.rm=T)) %>%
    mutate(Budget = 0)
  upload <- na.omit(upload)
  #Adding Zeros
  payperiods <- upload %>% ungroup() %>% select(Start.Date, End.Date) %>% distinct()
  map_cc <- map_cc %>% filter(Site %in% site.census) %>% select(CostCenter,VolumeID) %>% distinct()
  map_cc <- merge(map_cc, payperiods)
  map_cc <- map_cc %>% mutate(Concat = paste0(VolumeID, Start.Date))
  upload <- upload %>% mutate(Concat = paste0(VolumeID, Start.Date))
  zero_rows <- map_cc[!(map_cc$Concat %in% upload$Concat),]
  zero_rows <- zero_rows %>% 
    mutate(Corp = 729805,Site = site.premier,Volume = 0,Budget = 0)
  upload <- plyr::rbind.fill(upload, zero_rows)
  upload$Concat <- NULL
  return(upload)
}
data_upload_MSW <- upload_file(site_MSW, 'NY2162', map_CC_Vol)
data_upload_MSM <- upload_file(site_MSM, 'NY2163', map_CC_Vol)
data_upload_MSBIB <- rbind(upload_file(site_MSBI,'630571', map_CC_Vol), upload_file(site_MSB,'630571', map_CC_Vol))

# Export Files ------------------------------------------------------------
#setwd(paste0(dir, '/Upload Files'))
write.table(data_upload_MSW, file = paste0(dir,'/Upload Files', "/MSW_Census Days_", format(pp.start,"%d%b%y"), " to ", format(pp.end, "%d%b%y"), ".csv"), sep = ',' , row.names = F,col.names = F)
write.table(data_upload_MSM, file = paste0(dir,'/Upload Files',"/MSM_Census Days_", format(pp.start,"%d%b%y"), " to ", format(pp.end, "%d%b%y"), ".csv"), sep = ',', row.names = F, col.names = F)
write.table(data_upload_MSBIB, file = paste0(dir,'/Upload Files',"/MSBIB_Census Days_", format(pp.start,"%d%b%y"), " to ", format(pp.end, "%d%b%y"), ".csv"), sep = ',', row.names = F, col.names = F)

# Generating Quality Chart ------------------------------------------------
quality_chart <- function(data, site.census) {
  data_chart <- data_upload %>% ungroup()  %>% filter(Site == site.census) %>%
    select(Nursing.Station.Code, CostCenter, End.Date,Census.Day) %>%
    mutate(End.Date = format(End.Date, "%m.%d.%y")) %>%
    pivot_wider(names_from = End.Date, values_from = Census.Day, values_fn = list(Census.Day = sum))
}
chart_master <- lapply(as.list(unique(data_upload$Site)), function(x) quality_chart(data_upload, x))

# Export Quality Charts ---------------------------------------------------
write.xlsx2(chart_master[1],file = paste0(dir, '/Quality Chart_',format(Sys.time(), '%d%b%y'),'.xlsx'), row.names = F, sheetName = unique(data_upload$Site)[1])
sapply(2:length(chart_master),function(x) write.xlsx2(chart_master[x], file =paste0(dir, '/Quality Chart_',format(Sys.time(), '%d%b%y'),'.xlsx'), row.names = F, sheetName = unique(data_upload$Site)[x],append = T))


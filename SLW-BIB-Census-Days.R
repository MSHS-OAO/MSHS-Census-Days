dir <- 'J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/Census Days'

# Load Libraries ----------------------------------------------------------
library(tidyverse)
library(xlsx)

# User Input --------------------------------------------------------------
pp.start <- as.Date('2021-02-28') # start date of first pay period needed
pp.end <- as.Date('2021-03-27') # end date of the last pay period needed
if(pp.end < pp.start){stop("End date before Start date")} # inital QC check on date range
warning("Update Pay Periods Start and End Dates Needed:") #reminder to update dates
cat(paste("Pay period starting on",format(pp.start, "%m/%d/%Y"), 'and ending on',format(pp.end, "%m/%d/%Y") ),fill = T)
Sys.sleep(2)

# Constants ---------------------------------------------------------------
#Current names of sites - one for each site (order matters)
site_names <- c("MSB", "MSBI", "MSM", "MSW")
# Old site names - all old names and current name (order matters, must match above)
site_old_names <- list(c("BIB","MSB"), c("BIPTR","MSBITR","MSBI"),c("STL","MSM"), c("RVT","MSW"))

# Import Dictionaries -------------------------------------------------------
map_CC_Vol <-  read.xlsx(paste0(dir, '/BIBSLW_Volume ID_Cost Center_ Mapping.xlsx'), sheetIndex = 1)
dict_PC <- read.xlsx(paste0(dir,'/Pay Cycle Dictionaries.xlsx'), sheetIndex = 1)
colnames(dict_PC)[1] <- 'Census.Date'
dict_PC <- dict_PC %>% drop_na()
#Checking dates requested are valid payperiods
if(!pp.start %in% dict_PC$Start.Date){
  stop("Start date entered is not the start of a payperiod, please enter another start date")
}else if(!pp.end %in% dict_PC$End.Date){stop("End date entered is not the end of a pay period, please enter another end date")}


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
#select which file you want instead of most recent file
#data_census <- read.xlsx(choose.files(caption = "Select Census File", multi = F, default= 'J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Volume - Data/Multisite Volumes/Census Days/Source Data'), sheetIndex = 1)

# QC ----------------------------------------------------------------------
#Checking range of dates requested exists in the census file
if(pp.end > range(data_census$Census.Date)[2] | pp.start <range(data_census$Census.Date)[1]){
  if(format(pp.end, "%Y") > format(pp.start, "%Y") & pp.end > range(data_census$Census.Date)[2]){
    #If the end date requested goes beyond the calendar year of the census file
    warning(paste0("File only has calendar year to date data, end date will be changed to "),format(pp.start, "%Y"),"-12-31")
    pp.end.new <- as.Date(paste0(format(pp.start, "%Y"),"-12-31"))
  }else if(format(pp.start, "%Y") < format(pp.end, "%Y") & pp.start < range(data_census$Census.Date)[1]){
    #If the start date requested is before the calendar year of the census file
    warning(paste0("File only has calendar year to date data, start date will be changed to "),format(pp.end, "%Y"),"-01-01")
    pp.start.new <- as.Date(paste0(format(pp.end, "%Y"),'-01-01'))
  }else{#If date range requested is outside the range of the census file
    stop('Data Missing from Census file for Pay Periods Needed. Please add most recent file to the source data folder.')
  }}

#Checking if any sites have been renamed in Census file
if(any(!c(as.vector(unique(data_census$Site), mode = 'any')) %in% c(site_names,unlist(site_old_names)))){
  new_site_names <- which(!c(as.vector(unique(data_census$Site),mode = 'any')) %in% c(site_names,unlist(site_old_names)))
  new_site_names <- c(as.vector(unique(data_census$Site),mode = 'any'))[new_site_names]
  warning("New site name(s) found: ",paste(new_site_names,collapse = ", "))
  stop("Please update the new site names in the constants section and rerun code")
}

#Checking the pay cycle dictionary is up to date
if(range(data_census$Census.Date)[2] > range(dict_PC$End.Date)[2]){stop("Update Pay Cycle Dictionary")}

# Pre Processing ----------------------------------------------------------
data_census <- data_census %>%
  mutate(Site = as.character(Site),
         Nursing.Station.Code = as.character(Nursing.Station.Code))
map_CC_Vol <- map_CC_Vol %>%
  select (Site, Nursing.Station.Code, CostCenter, VolumeID) %>%
  mutate(Nursing.Station.Code = as.character(Nursing.Station.Code)) %>%
  drop_na() %>% distinct()
# if any of the files have old site names update them to new site names
if(any(!unique(map_CC_Vol$Site) %in% site_names) | any(!unique(data_census$Site) %in% site_names)){
  for(i in 1:length(site_names)){
  data_census$Site <- gsub(paste(unlist(site_old_names[i]),collapse = "|"), site_names[i],data_census$Site)
  map_CC_Vol$Site <- gsub(paste(unlist(site_old_names[i]),collapse = "|"), site_names[i],map_CC_Vol$Site)
  }
}
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
  #Adding Zeros for nursing stations with no census
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
#If there is a payperiod that goes into a different calendar year
new_start_end <- function(upload_file){
  if(exists('pp.end.new')){
    upload_file$End.Date <- gsub(format(pp.end,"%m/%d/%Y"),format(pp.end.new,"%m/%d/%Y"), upload_file$End.Date)
    return(upload_file)
  }else if(exists('pp.start.new')){
    upload_file$Start.Date <- gsub(format(pp.start,"%m/%d/%Y"),format(pp.start.new,"%m/%d/%Y"), upload_file$Start.Date)
    return(upload_file)
  }else{return(upload_file)}
}
#Creating upload file for each site
data_upload_MSW <- new_start_end(upload_file(site_names[4], 'NY2162', map_CC_Vol))
data_upload_MSM <- new_start_end(upload_file(site_names[3], 'NY2163', map_CC_Vol))
data_upload_MSBIB <- new_start_end(rbind(upload_file(site_names[1],'630571', map_CC_Vol), upload_file(site_names[2],'630571', map_CC_Vol)))

# Export Files ------------------------------------------------------------
write.table(data_upload_MSW, file = paste0(dir,'/Upload Files', "/MSW_Census Days_", if(exists('pp.start.new')){format(pp.start.new,"%d%b%y")}else{format(pp.start,"%d%b%y")}, " to ", if(exists('pp.end.new')){format(pp.end.new,"%d%b%y")}else{format(pp.end, "%d%b%y")}, ".csv"), sep = ',' , row.names = F,col.names = F)
write.table(data_upload_MSM, file = paste0(dir,'/Upload Files',"/MSM_Census Days_", if(exists('pp.start.new')){format(pp.start.new,"%d%b%y")}else{format(pp.start,"%d%b%y")}, " to ", if(exists('pp.end.new')){format(pp.end.new,"%d%b%y")}else{format(pp.end, "%d%b%y")}, ".csv"), sep = ',', row.names = F, col.names = F)
write.table(data_upload_MSBIB, file = paste0(dir,'/Upload Files',"/MSBIB_Census Days_", if(exists('pp.start.new')){format(pp.start.new,"%d%b%y")}else{format(pp.start,"%d%b%y")}, " to ", if(exists('pp.end.new')){format(pp.end.new,"%d%b%y")}else{format(pp.end, "%d%b%y")}, ".csv"), sep = ',', row.names = F, col.names = F)

# Generating Quality Chart ------------------------------------------------
quality_chart <- function(data, site.census) {
  data_chart <- data_upload %>% ungroup()  %>% filter(Site == site.census) %>%
    select(Nursing.Station.Code, CostCenter, End.Date,Census.Day) %>%
    arrange(End.Date) %>%
    mutate(End.Date = format(End.Date, "%m.%d.%y")) %>%
    pivot_wider(names_from = End.Date, values_from = Census.Day, values_fn = list(Census.Day = sum))%>%
    arrange(Nursing.Station.Code)
}
chart_master <- lapply(as.list(unique(data_upload$Site)), function(x) quality_chart(data_upload, x))

# Export Quality Charts ---------------------------------------------------
write.xlsx2(chart_master[1],file = paste0(dir, '/Quality Chart_',format(Sys.time(), '%d%b%y'),'.xlsx'), row.names = F, sheetName = unique(data_upload$Site)[1])
sapply(2:length(chart_master),function(x) write.xlsx2(chart_master[x], file =paste0(dir, '/Quality Chart_',format(Sys.time(), '%d%b%y'),'.xlsx'), row.names = F, sheetName = unique(data_upload$Site)[x],append = T))


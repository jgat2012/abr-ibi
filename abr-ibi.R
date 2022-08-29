## Load packages
pacman::p_load(
  rio,        # importing data  
  here,       # relative file pathways
  readxl,     # read excel
  writexl,    # write excel
  strex,      #String manipulation
  janitor,    # data cleaning and tables
  lubridate,  # working with dates
  tidyverse,  # data management and visualization,
  scales,     # easily convert proportions to percents  
  labelled,   # Variable and values labelling
  flextable,  # Format tables
  openxlsx,   # Write to Excel
)

# use here from the here package
here <- here::here
# use clean_names from the janitor package
clean_names <- janitor::clean_names

## Set data folder
data_folder  <- "data"
output_folder<- "output"

################################################################################
#################   I. Importing & Exploring Data         ######################
###################                                     ########################
################################################################################


## Get list of files in directory
list_of_files <- list.files("data/",pattern ="^2021IBIABRHIVKinshasa_DataQualityDiscrepancies",full.names = TRUE )


################################################################################
#################        II. Cleaning Data        ##############################
######################                        ##################################
################################################################################



################################################################################
#################        III. Analysing Data      ##############################
######################                        ##################################
################################################################################

## Loop through each file and append it in new data frame
  ## Structure of data rules
struc_data <-data.frame(matrix(nrow = 0,ncol=4))
final_data <- createWorkbook("ABR-IBIDataQualityWorkbook")

openxlsx::addWorksheet(final_data, "Missing-5", gridLines = TRUE)
openxlsx::addWorksheet(final_data, "Coherence-6", gridLines = TRUE)

data_five  <- data.frame(matrix(nrow = 0,ncol=6))
data_coh_6 <- data.frame(matrix(nrow = 0,ncol = 9))


x <-  1
for (myfile in list_of_files) {
  cur_data   <- read.csv(file = myfile)
  nocol      <- ncol(cur_data)
  rule_name  <- strsplit(myfile, "_")[[1]][3]
  full_name  <- myfile
  
  ## Populate final_data file
  
  ## Rules with more than 5 variables but that can be converted to 5 variables rules
  to_fiveMiss_rule_1 <- c("MissingGroupeDgeDuPatientFE","MissingHmoculturePrleveAuCHKFE",
                            "MissingNumroAdmissionCHKFE","MissingTempratureAxillaireFE")
  
  to_fiveMiss_rule_2 <- c("Missing19SoinsInfDomCRF","Missing21HospCHK30jCRF",
                          "Missing25CRF","Missing27CRF","MissingDiagSortie319CRF",
                          "MissingHosp30DerniersJoursCRF","MissingInjectableCRF",
                          "MissingModeSortie323CRF","MissingTttARV321CRF","MissingTttTB320CRF"
                          )
  
    ## if data has 5 columns and rule is missing
  if(nocol == 5 &&  grepl("Missing",myfile)){
    ## append 5th col ==missing column, 6th col = rule_name
    cur_data$variablemanquante <- colnames(cur_data)[5]
    cur_data$regle          <- rule_name
    cur_data$id_etude_crf   <- ""
    
    ## Remove missing variable from data to keep structure consistent
    cur_data <- cur_data[-c(2,5)]
    
    data_five <- data_five %>%
      rbind(
        cur_data
      )
   
    
  } 
  ## Missing data
  else if(nocol == 6 && grepl(paste(to_fiveMiss_rule_1,collapse="|"),myfile)){
    ## append 6th col ==missing column, th col = rule_name
    cur_data$variablemanquante <- colnames(cur_data)[6]
    cur_data$regle          <- rule_name
    cur_data$id_etude_crf   <- ""
    
    ## Remove variables not needed from data to keep structure consistent
    cur_data <- cur_data[-c(2,5,6)]
    
    data_five <- data_five %>%
      rbind(
        cur_data
      )
  } 
  ## Missing data
  else if(nocol == 6 && grepl(paste(to_fiveMiss_rule_2,collapse="|"),myfile)){
    ## append 6th col ==missing column, th col = rule_name
    cur_data$variablemanquante <- colnames(cur_data)[6]
    cur_data$regle          <- rule_name
    
    ## Remove variables not needed from data to keep structure consistent
    cur_data <- cur_data[-c(2,6)]
    
    data_five <- data_five %>%
      rbind(
        cur_data
      )
  }
  
  ##Discrepencies between data in different forms with 6 variables
  else if (nocol == 6 && grepl("^data/2021IBIABRHIVKinshasa_DataQualityDiscrepancies_Coh",myfile)){
    
    cur_data$nomvar1    <- colnames(cur_data)[5]
    cur_data$nomvar2    <- colnames(cur_data)[6]
    cur_data$valeurvar1 <- cur_data[,5]
    cur_data$valeurvar2 <- cur_data[,6]
    cur_data$regle      <- rule_name
    
    
    ## Remove variables not needed from data to keep structure consistent
    cur_data <- cur_data[-c(5,6)]
    
    data_coh_6 <- data_coh_6 %>%
      rbind(
        cur_data
      )
    
   
    
  }
  
  
  struc_data <- struc_data %>%
    
   rbind(
     c(x,rule_name,nocol,full_name)
   )
  x <- x+1
}
## Change heading of data frame
colnames(struc_data) <- c("No","Rule","NoColumns","Full Name")


## Change data types of struc_data
struc_data  <- struc_data %>%
  mutate(
    No        = as.numeric(No),
    NoColumns = as.numeric(NoColumns)
  )

### ------ Write outputs to Excel file ----- ####
openxlsx::writeData(final_data, sheet = 1, data_five,borderStyle = "thin",withFilter = TRUE)
openxlsx::writeData(final_data, sheet = 2, data_coh_6,borderStyle = "thin",withFilter = TRUE)

## Save final file
file_name<-paste("output/ABR-IBIDataQuality_",Sys.Date(),".xlsx",sep = "")

openxlsx::saveWorkbook(final_data, file_name, overwrite = TRUE)



cur_data

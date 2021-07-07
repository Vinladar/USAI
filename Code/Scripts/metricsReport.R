# This script is designed to take the RMA Detail report in Intuitive and output an Excel file with stats regarding RMAs.
# This version is an updated version for the metrics (created 6.2.2020)
# Import the required packages.
library(dplyr)
library(data.table)
library(openxlsx)
library(ggplot2)

# Read in the Master YTD RMA file.
YTD_master <- read.xlsx("Code/Failure_rate/2021/YTD/2021_YTD.xlsx") %>%
  select(c("RMA_ID", "Item_ID", "Item_Name", "Return_Qty", "Reason_Code", "Month"))
YTD_2020 <- read.xlsx("Code/Failure_rate/2020/2020_YTD.xlsx") %>%
  select(c("RMA.ID", "Item.ID", "Item.Name", "Return.Qty", "Reason.Code", "Month"))
names(YTD_2020) <- c("RMA_ID", "Item_ID", "Item_Name", "Return_Qty", "Reason_Code", "Month")
YTD_drivers <- filter(YTD_master, grepl("SP-[0-9]{3}-[0-9]{4}", Item_ID))
YTD_engines <- filter(YTD_master, grepl("LEM-", Item_ID))
drivers_2020 <- filter(YTD_2020, grepl("SP-[0-9]{3}-[0-9]{4}", Item_ID))
engines_2020 <- filter(YTD_2020, grepl("LEM-", Item_ID))

# These are all of the support files that are used to merge data from various locations.
drivers <- read.csv("Code/supportFiles/Drivers.csv")
driver_to_E2 <- read.csv("Code/supportFiles/driver_to_E2.csv")
reasonCodes <- read.csv("Code/supportFiles/Reasoncode.csv")
driver_to_partFam <- read.csv("Code/supportFiles/driverToPartFam.csv")
light_engines <- read.csv("Code/supportFiles/LightEngines.csv")
light_engine_to_partFam <- read.csv("Code/supportFiles/LightEngineToPartFam.csv")
LEM_to_LED <- read.csv("Code/supportFiles/LEM_to_LED.csv")

getMonthlyData <- function(month, driver_out, engine_out) {
  monthlyDrivers <- filter(YTD_drivers, Month == month)
  monthlyEngines <- filter(YTD_engines, Month == month)
  getDrivers(monthlyDrivers, driver_out)
  getLightEngines(monthlyEngines, engine_out)
}

getDrivers <- function(monthlyDrivers, driver_out) {
  driverIDs <- data.frame(do.call("rbind", strsplit(as.character(monthlyDrivers$Item_ID), "-", fixed = TRUE)))
  driversFinal <- cbind(monthlyDrivers, driverIDs) %>%
    select(c("RMA_ID", "Item_ID", "Item_Name", "Return_Qty", "Reason_Code", "X4"))
  names(driversFinal) <- c("RMA_ID", "Item_ID", "Item_Name", "Return_Qty", "Reason_Code", "DIM_Type")
  groupedDrivers <- merge(driversFinal, driver_to_E2, by.x = c("Item_ID"), by.y = c("SP_Kit"), all.x = TRUE) %>%
    select(c("Item_ID", "RMA_ID", "Item_Name", "Return_Qty", "Reason_Code", "DIM_Type", "E2", "SP_Acct_Val", "E2_Acct_Val", "SP_Price")) %>%
    group_by(E2, DIM_Type, E2_Acct_Val, SP_Acct_Val, SP_Price)
  groupedDrivers$SP_Price <- as.numeric(groupedDrivers$SP_Price)
  groupedDrivers$SP_Acct_Val <- as.numeric(groupedDrivers$SP_Acct_Val)
  groupedDrivers$E2_Acct_Val <- as.numeric(groupedDrivers$E2_Acct_Val)
  groupedDrivers1 <- summarise(groupedDrivers, "# of RMAs" = length(unique(RMA_ID)),  Qty = sum(Return_Qty))
  groupedDrivers1 <- arrange(groupedDrivers1, desc(Qty))
  write.csv(groupedDrivers1, driver_out)
}

getLightEngines <- function(monthlyEngines, engine_out) {
  rmaEngines <- merge(monthlyEngines, LEM_to_LED, by.x = c("Item_ID"), by.y = c("LEM_Kit"), all.x = TRUE)
  rmaEngines$LEM <- substring(rmaEngines$Item_ID, 1, 7)
  rmaEngines <- merge(rmaEngines, light_engine_to_partFam, by.x = c("LEM"), by.y = c("Item.ID"), all.x = TRUE)
  names(rmaEngines) <- c("LEM", "Item_ID", "RMA_ID", "Item_Name" , "Return_Qty", 
                         "Reason_Code", "Month", "LED", "LEM_Acct_Val", 
                         "LED_Acct_Val", "LEM_Price", "Product_Family")
  groupedEngines <- group_by(rmaEngines, LED, Product_Family, LED_Acct_Val, LEM_Acct_Val, LEM_Price)
  groupedEngines1 <- summarize(groupedEngines, "# of RMAs" = length(unique(RMA_ID)),  Qty = sum(Return_Qty))
  groupedEngines1 <- arrange(groupedEngines1, desc(Qty)) %>%
    select(c("# of RMAs", "Qty", "LED", "Product_Family", "LED_Acct_Val", "LEM_Acct_Val", "LEM_Price"))
  write.csv(groupedEngines1, engine_out)
}

getYTDData <- function(driver_output, engine_output) {
  getYTDDrivers(driver_output)
  getYTDEngines(engine_output)
}

getYTDDrivers <- function(driver_output) {
  YTD_drivers <- merge(YTD_drivers, driver_to_E2, by.x = c("Item_ID"), by.y = c("SP_Kit"), all.x = TRUE) %>%
    select(c(E2, Return_Qty, SP_Acct_Val, E2_Acct_Val, SP_Price, Month)) %>%
    group_by(E2, Month, SP_Acct_Val)
  Grouped_YTD_drivers <- summarize(YTD_drivers, Return_Qty = sum(Return_Qty)) %>%
    arrange(desc(E2))
  write.csv(Grouped_YTD_drivers, driver_output)
}

getYTDEngines <- function(engine_output) {
  YTD_engines <- merge(YTD_engines, LEM_to_LED, by.x = c("Item_ID"), by.y = c("LEM_Kit"), all.x = TRUE) %>%
    select(c(LED, Return_Qty, LEM_Acct_Val, LED_Acct_Val, LEM_Price, Month)) %>%
    group_by(LED, Month, LEM_Acct_Val)
  Grouped_YTD_engines <- summarize(YTD_engines, Return_Qty = sum(Return_Qty)) %>%
    arrange(desc(LED))
  write.csv(Grouped_YTD_engines, engine_output)
}

get2020YTD <- function(driver_output, engine_output) {
  
}
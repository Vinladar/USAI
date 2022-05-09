# This script is designed to take the RMA Detail report in Intuitive and output an Excel file with stats regarding RMAs.
# This version is an updated version for the metrics (created 6.2.2020)
# Import the required packages.
library(dplyr)
library(tidyverse)
library(data.table)
library(openxlsx)
library(ggplot2)

# Read in the Master YTD RMA file.
YTD_2020 <- read.xlsx("Code/Failure_rate/2020/2020_YTD.xlsx") %>%
  select(c("RMA.ID", "Item.ID", "Item.Name", "Return.Qty", "Reason.Code", "Month"))
names(YTD_2020) <- c("RMA_ID", "Item_ID", "Item_Name", "Return_Qty", "Reason_Code", "Month")
YTD_master <- read.xlsx("Code/Failure_rate/2022/YTD_2022.xlsx")
names(YTD_master) <- c("RMA_ID", "Item_ID", "Return_Qty", "Item_Name", "Reason_Code", "Month")
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
LEM_to_LED <- LEM_to_LED[, -6]
LEM_to_LED$LEM_Acct_Val <- as.numeric(LEM_to_LED$LEM_Acct_Val)
LEM_to_LED$LED_Acct_Val <- as.numeric(LEM_to_LED$LED_Acct_Val)
LEM_to_LED$LEM_Price <- as.numeric(LEM_to_LED$LEM_Price)
names(LEM_to_LED) <- c("LEM_Kit", "LED", "LEM_Acct_Val", "LED_Acct_Val", "LEM_Price")

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
  driversFinal1 <- driversFinal %>%
    left_join(driver_to_E2, by = c("Item_ID" = "SP_Kit")) %>%
    select("RMA_ID", "Item_ID", "Item_Name", "Return_Qty", "Reason_Code", 
           "DIM_Type", "E2", "SP_Acct_Val", "E2_Acct_Val", "SP_Price")
  
  driversFinal1$SP_Acct_Val <- as.numeric(driversFinal1$SP_Acct_Val)
  driversFinal1$E2_Acct_Val <- as.numeric(driversFinal1$E2_Acct_Val)
  driversFinal1$Return_Qty <- as.numeric(driversFinal1$Return_Qty)
  nullDrivers <- filter(driversFinal1, is.na(E2))
  if (nrow(nullDrivers) > 0) 
    return(write.csv(nullDrivers, str_c(driver_out, "NULLdrivers.csv")))
  driversFinal1$Total_cost <- driversFinal1$Return_Qty * driversFinal1$SP_Acct_Val
  groupedDrivers <- driversFinal1 %>%
    group_by(E2, DIM_Type, E2_Acct_Val, SP_Acct_Val)
  
  groupedDrivers1 <- summarise(groupedDrivers, "# of RMAs" = length(unique(RMA_ID)),  Qty = sum(Return_Qty), Total_cost = sum(Total_cost))
  groupedDrivers1 <- select(groupedDrivers1, "# of RMAs", "Qty", "E2", "DIM_Type", "E2_Acct_Val", "SP_Acct_Val", "Total_cost")
  groupedDrivers1 <- arrange(groupedDrivers1, desc(Qty))
  groupedDriversNoOutliers <- filter(groupedDrivers1, E2 != "E2-526" & E2 != "E2-527")
  write.csv(groupedDrivers1, str_c(driver_out, "Drivers.csv"))
  write.csv(head(groupedDrivers1, 15), str_c(driver_out, "DriversTop15.csv"))
  write.csv(groupedDriversNoOutliers, str_c(driver_out, "DriversNoOutliers.csv"))
  write.csv(head(groupedDriversNoOutliers, 15), str_c(driver_out, "DriversTop15NoOutliers.csv"))
}

getLightEngines <- function(monthlyEngines, engine_out) {
  rmaEngines <- monthlyEngines %>%
    left_join(LEM_to_LED, by = c("Item_ID" = "LEM_Kit"))
  nullEngines <- filter(rmaEngines, is.na(LED))
  if (nrow(nullEngines) > 0) 
    return(write.csv(nullEngines, str_c(engine_out, "NULLengines.csv")))
  rmaEngines$LEM <- substring(rmaEngines$Item_ID, 1, 7)
  rmaEngines1 <- rmaEngines %>%
    left_join(light_engine_to_partFam, by = c("LEM" = "Item.ID"))
  names(rmaEngines1) <- c("RMA_ID", "Item_ID", "Return_Qty", "Item_Name", 
                          "Reason_Code", "Month", "LED", "LEM_Acct_Val", 
                          "LED_Acct_Val", "LEM_Price", "LEM", "Product_Family")
  rmaEngines1$Return_Qty <- as.numeric(rmaEngines1$Return_Qty)
  rmaEngines1$LEM_Acct_Val <- as.numeric(rmaEngines1$LEM_Acct_Val)
  rmaEngines1$LED_Acct_Val <- as.numeric(rmaEngines1$LED_Acct_Val)
  rmaEngines1$LEM_Price <- as.numeric(rmaEngines1$LEM_Price)
  groupedEngines <- rmaEngines1 %>%
    group_by(LED, Product_Family, LED_Acct_Val, LEM_Acct_Val, LEM_Price)
  groupedEngines1 <- summarize(groupedEngines, "# of RMAs" = length(unique(RMA_ID)),  Qty = sum(Return_Qty))
  groupedEngines1 <- arrange(groupedEngines1, desc(Qty)) %>%
    select(c("# of RMAs", "Qty", "LED", "Product_Family", "LED_Acct_Val", "LEM_Acct_Val"))
  groupedEngines1$Total_cost <- groupedEngines1$LEM_Acct_Val * groupedEngines1$Qty
  groupedEnginesNoOutliers <- filter(groupedEngines1, !str_detect(LED, "LED-213"))
  groupedEnginesNoOutliers <- filter(groupedEnginesNoOutliers, !str_detect(LED, "LED-203"))
  write.csv(groupedEngines1, str_c(engine_out, "Engines.csv"))
  write.csv(head(groupedEngines1, 15), str_c(engine_out, "EnginesTop15.csv"))
  write.csv(head(groupedEnginesNoOutliers, 15), str_c(engine_out, "EnginesTop15NoOutliers.csv"))
  write.csv(groupedEnginesNoOutliers, str_c(engine_out, "EnginesNoOutliers.csv"))
}

getYTDData <- function(driver_output, engine_output) {
  getYTDDrivers(driver_output)
  getYTDEngines(engine_output)
}

getYTDDrivers <- function(driver_output) {
  YTD_drivers1 <- merge(YTD_drivers, driver_to_E2, by.x = c("Item_ID"), by.y = c("SP_Kit"), all.x = TRUE) %>%
    select(c(E2, Return_Qty, SP_Acct_Val, Month))
  YTD_drivers1$Total_Cost <- YTD_drivers1$Return_Qty * YTD_drivers1$SP_Acct_Val
  Grouped_YTD_drivers <- group_by(YTD_drivers1, E2, Month)
  Grouped_YTD_drivers <- summarize(Grouped_YTD_drivers, Return_Qty = sum(Return_Qty), Total_Cost = sum(Total_Cost)) %>%
    arrange(desc(E2))
  write.csv(Grouped_YTD_drivers, driver_output)
}

getYTDEngines <- function(engine_output) {
  YTD_engines <- merge(YTD_engines, LEM_to_LED, by.x = c("Item_ID"), by.y = c("LEM_Kit"), all.x = TRUE) %>%
    select(c(LED, Return_Qty, LEM_Acct_Val, Month))
  YTD_engines$Total_Cost <- YTD_engines$Return_Qty * YTD_engines$LEM_Acct_Val
  Grouped_YTD_engines <- group_by(YTD_engines, LED, Month)
  Grouped_YTD_engines <- summarize(Grouped_YTD_engines, Return_Qty = sum(Return_Qty), Total_Cost = sum(Total_Cost)) %>%
    arrange(desc(LED))
  write.csv(Grouped_YTD_engines, engine_output)
}

get2020YTD <- function(driver_output, engine_output) {
  
}
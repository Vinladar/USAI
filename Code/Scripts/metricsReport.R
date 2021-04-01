# This script is designed to take the RMA Detail report in Intuitive and output an Excel file with stats regarding RMAs.
# This version is an updated version for the metrics (created 6.2.2020)
# Import the required packages.

library(dplyr)
library(data.table)
library(openxlsx)
library(ggplot2)

file_location <- "Code/Failure_rate/2021/March_2021.xlsx"
driver_out <- "Code/Output/March_2021/Drivers.csv"
engines_out <- "Code/Output/March_2021/Engines.csv"
codes_out <- "Code/Output/March_2021/Codes.csv"

loadData <- function(fileDir) {
  rma <- read.xlsx(fileDir)
}

drivers <- read.csv("Code/supportFiles/Drivers.csv")
driver_to_E2 <- read.csv("Code/supportFiles/driver_to_E2.csv")
reasonCodes <- read.csv("Code/supportFiles/Reasoncode.csv")
driver_to_partFam <- read.csv("Code/supportFiles/driverToPartFam.csv")
light_engines <- read.csv("Code/supportFiles/LightEngines.csv")
light_engine_to_partFam <- read.csv("Code/supportFiles/LightEngineToPartFam.csv")
LEM_to_LED <- read.xlsx("Code/supportFiles/LEM_to_LED.xlsx")

rma <- loadData(file_location)
rma2 <- select(rma, RMA.ID, Item.ID, Item.Name, Return.Qty, Reason.Code, Month)
names(rma2) <- c("RMA_ID", "Item_ID", "Item_Name", "Return_Qty", "Reason_Code", "Month")

getDrivers <- function(rma) {
  rmaDrivers <- rma[grep("SP-[0-9]{3}-[0-9]{4}", rma$Item_ID),]
  driverIDs <- data.frame(do.call("rbind", strsplit(as.character(rmaDrivers$Item_ID), "-", fixed = TRUE)))
  driversFinal <- cbind(rmaDrivers, driverIDs)[, c(1, 2, 4, 6, 10)]
  names(driversFinal) <- c("RMA_ID", "Item_ID", "Return_Qty", "Reason_Code", "DIM_Type")
  groupedDrivers <- merge(driversFinal, driver_to_E2, by.x = c("Item_ID"), by.y = c("SP_Kit"), all.x = TRUE) %>%
    group_by(E2, DIM_Type)
  groupedDrivers1 <- groupedDrivers[, c(6, 2, 3, 4, 5, 7, 8, 9)]
  groupedDrivers1$SP_Price <- as.numeric(groupedDrivers1$SP_Price)
  groupedDrivers1$SP_Acct_Val <- groupedDrivers1$SP_Acct_Val * groupedDrivers1$Return_Qty
  groupedDrivers1$E2_Acct_Val <- groupedDrivers1$E2_Acct_Val * groupedDrivers1$Return_Qty
  groupedDrivers1$SP_Price <- groupedDrivers1$SP_Price * groupedDrivers1$Return_Qty
  groupedDrivers1 <- summarise(groupedDrivers1, "# of RMAs" = length(unique(RMA_ID)),  Qty = sum(Return_Qty), E2_Acct_Val = sum(E2_Acct_Val), SP_Acct_Val = sum(SP_Acct_Val), SP_Price = sum(SP_Price))
  groupedDrivers1 <- groupedDrivers1[, c(3, 4, 1, 2, 5, 6, 7)]
  groupedDrivers1$CostToReplace <- groupedDrivers1$SP_Price - groupedDrivers1$SP_Acct_Val
  groupedDrivers1 <- arrange(groupedDrivers1, desc(Qty))
  write.csv(groupedDrivers1, driver_out)
}

getLightEngines <- function(rma) {
  names(light_engine_to_partFam) <- c("Item_ID", "Product_Family")
  rmaEngines <- rma[grep("LEM-", rma$Item_ID),]
  rmaEngines <- merge(rmaEngines, LEM_to_LED, by.x = c("Item_ID"), by.y = c("LEM_Kit"), all.x = TRUE)
  rmaEngines$LEM <- substring(rmaEngines$Item_ID, 1, 7)
  rmaEngines <- merge(rmaEngines, light_engine_to_partFam, by.x = c("LEM"), by.y = c("Item_ID"), all.x = TRUE)
  groupedEngines <- group_by(rmaEngines, LED, Product_Family)[, c(3, 5, 6, 8, 9, 10, 11, 12)]
  groupedEngines$LEM_Acct_Val <- groupedEngines$LEM_Acct_Val * groupedEngines$Return_Qty
  groupedEngines$LED_Acct_Val <- groupedEngines$LED_Acct_Val * groupedEngines$Return_Qty
  groupedEngines$LEM_Price <- groupedEngines$LEM_Price * groupedEngines$Return_Qty
  groupedEngines1 <- summarize(groupedEngines, "# of RMAs" = length(unique(RMA_ID)),  Qty = sum(Return_Qty), LEM_Acct_Val = sum(LEM_Acct_Val), LED_Acct_Val = sum(LED_Acct_Val), LEM_Price = sum(LEM_Price))
  groupedEngines1 <- arrange(groupedEngines1, desc(Qty))[, c(3, 4, 1, 2, 6, 5, 7)]
  groupedEngines1$CostToReturn <- groupedEngines1$LEM_Price - groupedEngines1$LEM_Acct_Val
  write.csv(groupedEngines1, engines_out)
}

getReasonCodes <- function(rma) {
  codes <- rma[, c(6, 4, 1)]
  groupedCodes <- group_by(codes, reason) %>%
    summarize(Number_of_RMAs = length(unique(RMA_ID)), Qty = sum(Return_Qty)) %>%
    arrange(desc(Number_of_RMAs))
  write.csv(groupedCodes, codes_out)
}

generateReports <- function(rma) {
  rma <- merge(rma, reasonCodes, by.x = c("Reason_Code"), by.y = c("code"), all.x = TRUE)[,c(2:7)]
  getDrivers(rma)
  getLightEngines(rma)
  getReasonCodes(rma)
}

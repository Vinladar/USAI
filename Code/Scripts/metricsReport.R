# This script is designed to take the RMA Detail report in Intuitive and output an Excel file with stats regarding RMAs.
# This version is an updated version for the metrics (created 6.2.2020)
# Import the required packages.

library(dplyr)
library(data.table)
library(openxlsx)
library(ggplot2)

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

rma <- loadData("Code/Failure_rate/2021/January_2021.xlsx")
rma2 <- select(rma, RMA.ID, Item.ID, Item.Name, Return.Qty, Reason.Code, Month)
rma3 <- merge(rma2, reasonCodes, by.x = c("Reason.Code"), by.y = c("reason"), all.x = TRUE)[, c(2, 3, 5, 6, 7)]

getDrivers <- function(rma) {
  rmaDrivers <- rma[grep("SP-[0-9]{3}-[0-9]{4}", rma$Item.ID),]
  driverIDs <- data.frame(do.call("rbind", strsplit(as.character(rmaDrivers$Item.ID), "-", fixed = TRUE)))
  driversFinal <- cbind(rmaDrivers, driverIDs)[, c(1, 2, 3, 4, 5, 9)]
  names(driversFinal) <- c("RMA_ID", "Item_ID", "Return_Qty", "Month", "Reason_Code", "DIM_Type")
  groupedDrivers <- merge(driversFinal, driver_to_E2, by.x = c("Item_ID"), by.y = c("SP_Kit"), all.x = TRUE) %>%
    group_by(E2, DIM_Type)
  groupedDrivers1 <- groupedDrivers[, c(2, 3, 5, 6, 7)]
  groupedDrivers1 <- summarise(groupedDrivers1, "# of RMAs" = length(unique(RMA_ID)),  Qty = sum(Return_Qty))
  groupedDrivers1 <- groupedDrivers1[, c(3, 4, 1, 2)]
  groupedDrivers1 <- arrange(groupedDrivers1, desc(Qty))
  write.csv(groupedDrivers1, "Code/Output/January_2021/Drivers.csv")
}

getLightEngines <- function(rma) {
  names(light_engine_to_partFam) <- c("Item_ID", "Product_Family")
  rmaEngines <- rma[grep("LEM-", rma$Item.ID),]
  rmaEngines <- merge(rmaEngines, LEM_to_LED, by.x = c("Item.ID"), by.y = c("LEM_Kit"), all.x = TRUE)
  rmaEngines$LEM <- substring(rmaEngines$Item.ID, 1, 7)
  rmaEngines <- merge(rmaEngines, light_engine_to_partFam, by.x = c("LEM"), by.y = c("Item_ID"), all.x = TRUE)
  groupedEngines <- group_by(rmaEngines, LED, Product_Family)[, c(3, 4, 6, 7, 8)]
  groupedEngines <- summarize(groupedEngines, "# of RMAs" = length(unique(RMA.ID)),  Qty = sum(Return.Qty))[, c(3, 4, 1, 2)]
  groupedEngines <- arrange(groupedEngines, desc(Qty))
  write.csv(groupedEngines, "Code/Output/January_2021/Engines.csv")
}

getReasonCodes <- function(rma) {
  codes <- rma[, c(5, 3, 1)]
  groupedCodes <- group_by(codes, code) %>%
    summarize(Number_of_RMAs = length(unique(RMA.ID)), Qty = sum(Return.Qty)) %>%
    arrange(desc(Qty))
  write.csv(groupedCodes, "Code/Output/January_2021/Codes.csv")
}

generateReports <- function(rma) {
  getDrivers(rma)
  getLightEngines(rma)
  getReasonCodes(rma)
}

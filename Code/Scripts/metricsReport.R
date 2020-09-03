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
LEM_to_LED <- read_excel("Code/supportFiles/LEM_to_LED.xlsx")

rma <- loadData("Code/Failure_rate/August_2020.xlsx")
rma2 <- select(rma, RMA.ID, Item.ID, Item.Name, Return.Qty, Reason.Code)
rma3 <- merge(rma2, reasonCodes, by.x = c("Reason.Code"), by.y = c("reason"), all.x = TRUE)[c(2:6)]
rma3 <- rename(rma3, Reason.Code = code)

getDrivers <- function(rma) {
  rmaDrivers <- rma[grep("SP-[0-9]{3}-[0-9]{4}", rma$Item.ID),]
  groupedDrivers <- merge(rmaDrivers, driver_to_E2, by.x = c("Item.ID"), by.y = c("SP_Kit"), all.x = TRUE) %>%
    group_by(E2)
  groupedDrivers <- groupedDrivers[, c(2, 4, 6)]
  groupedDrivers1 <- summarise(groupedDrivers, "# of RMAs" = length(unique(RMA.ID)),  Qty = sum(Return.Qty))
  groupedDrivers1 <- groupedDrivers1[, c(2, 3, 1)]
  groupedDrivers1 <- arrange(groupedDrivers1, desc(Qty))
  write.csv(groupedDrivers1, "Code/Output/August_2020/Drivers.csv")
}

getLightEngines <- function(rma) {
  rmaEngines <- rma[grep("LEM-", rma$Item.ID),]
  rmaEngines <- merge(rmaEngines, LEM_to_LED, by.x = c("Item.ID"), by.y = c("LEM_Kit"), all.x = TRUE)
  groupedEngines <- group_by(rmaEngines, LED)
  groupedEngines <- summarize(groupedEngines, "# of RMAs" = length(unique(RMA.ID)),  Qty = sum(Return.Qty))[, c(2, 3, 1)]
  groupedEngines <- arrange(groupedEngines, desc(Qty))
  write.csv(groupedEngines, "Code/Output/August_2020/Engines.csv")
}

getReasonCodes <- function(rma) {
  codes <- rma[, c(5, 4, 1)]
  groupedCodes <- group_by(codes, Reason.Code) %>%
    summarize(Number_of_RMAs = length(unique(RMA.ID)), Qty = sum(Return.Qty)) %>%
    arrange(desc(Qty))
  write.csv(groupedCodes, "Code/Output/August_2020/Codes.csv")
}

generateReports <- function(rma) {
  getDrivers(rma)
  getLightEngines(rma)
  getReasonCodes(rma)
}

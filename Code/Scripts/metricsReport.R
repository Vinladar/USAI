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

rma <- loadData("Code/Failure_rate/May_2020.xlsx")
rma2 <- select(rma, RMA.ID, Item.ID, Item.Name, Return.Qty, Reason.Code)
rma3 <- merge(rma2, reasonCodes, by.x = c("Reason.Code"), by.y = c("reason"))[c(2:6)]
rma3 <- rename(rma3, Reason.Code = code)



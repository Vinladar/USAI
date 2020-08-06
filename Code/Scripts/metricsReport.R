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

rma <- loadData("Code/Failure_rate/July_2020.xlsx")
rma2 <- select(rma, RMA.ID, Item.ID, Item.Name, Return.Qty, Reason.Code)
rma3 <- merge(rma2, reasonCodes, by.x = c("Reason.Code"), by.y = c("reason"), all.x = TRUE)[c(2:6)]
rma3 <- rename(rma3, Reason.Code = code)

getDrivers <- function(rma) {
  rmaDrivers <- rma[grep("SP-[0-9]{3}-[0-9]{4}", rma$Item.ID),]
  groupedDrivers <- merge(rmaDrivers, driver_to_E2, by.x = c("Item.ID"), by.y = c("SP_Kit")) %>%
    group_by(E2)
# groupedDrivers <- summarize(groupedDrivers, Qty = sum(Return.Qty))
  splitDrivers <- group_split(groupedDrivers)
  splitDrivers <- lapply(splitDrivers, function(x) {
    cbind(x, Group = factor(1:nrow(x)))
  })
  groupedDrivers <- rbindlist(splitDrivers)[,c(2, 3, 4)]
  g <- ggplot(data = groupedDrivers, 
              aes(x = reorder(E2, -Return.Qty), y = Return.Qty, position = "stacked", fill = Group))
  g + geom_bar(stat = "identity") + 
    labs(x = "Driver") +
    theme(axis.text.x = element_text(angle = 90), legend.position = "none") +
    scale_fill_brewer(palette="Set1")
}

getLightEngines <- function(rma) {
  rmaEngines <- rma[grep("LEM-", rma$Item.ID),]
  rmaEngines <- cbind(rmaEngines, Item.ID = substring(rmaEngines$Item.ID, 1, 7))[,c(6, 4, 1)]
  groupedEngines <- group_by(rmaEngines, Item.ID)
  groupedEngines <- summarize(groupedEngines, "# of RMAs" = length(unique(RMA.ID)),  Qty = sum(Return.Qty))
  groupedEngines <- arrange(groupedEngines, desc(Qty))
  g <- ggplot(data = groupedEngines, aes(x = reorder(Item.ID, -Qty), y = Qty))
  g + geom_bar(stat = "identity") + labs(x = "Light Engine") + geom_text(aes(label = Qty, y = Qty + 2.0), position = position_dodge(0.9), vjust = 0)
}

getReasonCodes <- function(rma) {
  codes <- rma[, c(5, 4, 1)]
  groupedCodes <- group_by(codes, Reason.Code) %>%
    summarize("Number of RMAs" = length(unique(RMA.ID)), Qty = sum(Return.Qty)) %>%
    arrange(desc(Qty))
  g <- ggplot(data = groupedCodes, aes(x = reorder(Reason.Code, -Qty), y = Qty))
  g + geom_bar(stat = "identity") + labs(x = "Reason Code") + geom_text(aes(label = Qty, y = Qty + 2.0), position = position_dodge(0.9), vjust = 0)
}

generateGraphs <- function(rma) {
  getDrivers(rma)
  getLightEngines(rma)
  getReasonCodes(rma)
}

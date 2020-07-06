# This script is designed to take the RMA Detail report in Intuitive and output an Excel file with stats regarding RMAs.
# Import the required packages.
library(dplyr)
library(data.table)
library(openxlsx)
library(ggplot2)

# Generate the light engine report. 
loadLightEngines <- function(X, output = "Code/Output/Output.xlsx") {
# Remove the columns from the dataframe that are not needed.
  X <- select(X, RMA.ID, Item.ID, Return.Qty)
  lightEngines <- X[grep("LEM-", X$Item.ID),]
  lightEngines$Item.ID <- substring(lightEngines$Item.ID, 1, 7)
  grpLightEngines <- group_by(lightEngines, Item.ID)
  grpLightEngines <- arrange(grpLightEngines, Return.Qty)
  grpLightEngines <- summarise(grpLightEngines, length(unique(RMA.ID)), sum(Return.Qty))
  names(grpLightEngines) <- c("Item_ID", "RMA_ID", "Qty")
  grpLightEngines <- arrange(grpLightEngines, desc(Qty))
  qplot(y = Qty, data = grpLightEngines[-grpLightEngines$RMA_ID,], geom = "bar")
}

# Generates the Reason code report.
reasonCodes <- function(X, output) {
# Remove the columns from the dataframe that are not needed.
	X2 <- as.data.table(X[, c("Reason.Code", "Return.Qty")])
# Cast the return quantity as a numeric value so that we can aggregate the results.
	X2$Return.Qty <- as.numeric(X2$Return.Qty)
# Aggregate the results
	X3 <- aggregate(X2$Return.Qty, by = list(X2$Reason.Code), sum)
# Change the column names.
	names(X3) <- c("Reason_Code", "Quantity")
# Remove blank entries.
	X4 <- X3[!is.na(X3$Quantity),]
# Sort the results by quantity in descending order.
	X4 <- X4[order(-X4$Quantity),]
# Output the results to a file.
#	write.xlsx(X4, output, asTable = FALSE)
}

# Generate the driver report.
loadDrivers <- function(X, output) {
# Remove the columns from the dataframe that are not needed.
	X <- X[,c("Item.ID", "Return.Qty")]
# Collect only the driver part numbers and quantities.
	drivers <- X[grep("SP-[0-9]{3}-[0-9]{4}", X$Item.ID),]
# Remove uneeded characters from EML drivers
	drivers[grep("EML$", drivers$Item.ID),1] <- substring(drivers[grep("EML$", drivers$Item.ID),1], 1, 13)
# Replaces the current values with XXXX.
	substring(drivers[,1], 8, 11) <- "XXXX"
# Cast the quanitities as numeric values.
	drivers$Return.Qty <- as.numeric(drivers$Return.Qty)
# Aggregate the results.
	drivers <- as.data.table(aggregate(drivers$Return.Qty, by = list(drivers$Item.ID), sum))
# Rename the columns
	names(drivers) <- c("Driver", "Quantity")
# Sort the results by quantities in descending order.
	drivers <- drivers[order(-drivers$Quantity),]
# Remove any with a quantity of 0.
	drivers <- filter(drivers, drivers$Quantity != 0)
# Output the results.
	write.xlsx(drivers, output, asTable = FALSE)
}

# Runs all three of the report functions.
generateReports <- function(fileName, lightEngines = "engines.xlsx", drivers = "drivers.xlsx", codes = "codes.xlsx") {
	X <- read.xlsx(fileName)
	rma <- read.csv("Code/supportFiles/driver_to_E2.csv")
	loadLightEngines(X, lightEngines)
	loadDrivers(X, drivers)
	reasonCodes(X, codes)
}






















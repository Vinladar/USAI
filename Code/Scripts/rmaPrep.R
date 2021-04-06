library(data.table)
library(dplyr)
library(openxlsx)

# Test line of code.

so_master <- read.xlsx("Code/supportFiles/Sales_Order_Master/SO_Master.xlsx")

# Process the RMA information
rma_edit <- function(rma, rmaDetails, outputString = "Default.xlsx") {
	rma <- buildTable(rma, rmaDetails)
	rmaOutput <- addProductFamilies(rma)
	rmaOutput <- addUserDefs(rmaOutput)
	write.xlsx(rmaOutput, outputString, asTable = FALSE)
}

# Build the table. 
# There are a few steps in this part of the process. First, the columns are cleaned.
# Then the Dates are converted and the sales order numbers are added.
# Column names are then redefined, error codes are added, then the summaries are formatted
buildTable <- function(rma, rmaDetails) {
	names(rma) <- c("RMA_DateCreated", "CBA_Addr", "CBA_City", "CBA_Country", "Customer Name", 
	                "CBA_Address", "RMA_Comments", "RMAL_Problem Description", "IMA_Prod Fam", 
	                "RMAL_Return Quantity", "CBA_Currency Type", "IMA_Item ID", "CBA_EMail", 
	                "CBA_Name", "RMA_ID", "CBA_Doc Delivery Method", "CBA_Fax")
  rma <- arrange(rma, desc(RMA_ID))
  names(rmaDetails) <- c("RMA_ID", "RMA Line", "SO Req'd Date", "SO_ID", "SO Line", "Invoice ID", 
                         "Customer ID", "Item ID", "Rev", "Item Name", "Return Qty", "Unit Price", 
                         "RMA Status", "Credit/Replace Option", "Reason_Code", "Disc. %", "Sales Tax", 
                         "Extra Charges",  "Extra Charge Desc", "Extended Amt", "Project ID", 
                         "Misc Line Description", "Problem Desc.", "Replacement Complete") 
  rmaDetails <- arrange(rmaDetails, desc(RMA_ID))
	rma[, c(2, 3, 4, 6, 11, 13, 14, 16, 17)] = ""
	## rma[,1] = convertToDate(rma[, 1], origin = "1900-01-01")
	rma[,2] <- rmaDetails$SO_ID
	columns <- c("Date", "orig_SO", "ShipDate", "Specifier", "Customer", 
		"Job_Name", "RMA_Comments", "Summary", "Product_Family", "Qty", "Qty_On_Order", 
		"Prob_Part", "Replaced", "Repl_SO", "RMA", "Comments", "Code")
	colnames(rma) <- columns
	rma$Code <- rmaDetails$Reason_Code
	rma <- rma[,-7]
	rma$Summary <- as.data.table(strsplit(rma$Summary, " // "))[2] %>%
		t()
	rma$orig_SO <- rmaDetails$SO_ID
	rma
}

# Read the reason code CSV and change them in the rma table.
addReasonCodes <- function(rma) {
	codes <- read.csv("Code/supportFiles/Reasoncode.csv", header = TRUE)
	rma <- merge(rma, codes, by.x = c("Code"), by.y = c("code"))[c(2:17)]
	rma
}

# Split into groups for drivers, light engines, and other parts, then 
#	change the product families.
addProductFamilies <- function(rma) {
	engines <- rma[grep("LEM-", rma$Prob_Part),]
	drivers <- rma[grep("SP-[0-9]{3}-[0-9]{4}-", rma$Prob_Part),]
	others <- rma
	others <- others[-grep("LEM-|SP-[0-9]{3}-[0-9]{4}-", others$Prob_Part),]
	engineToPartFam <- read.csv("Code/supportFiles/LightEngineToPartFam.csv")
	names(engineToPartFam) <- c("Item.ID", "Product.Family")
	driverToPartFam <- read.csv("Code/supportFiles/driverToPartFam.csv")
	names(driverToPartFam) <- c("Item.ID", "Product.Family")
	engines <- cbind(engines, PartFam = substring(engines$Prob_Part, 1, 7))
	drivers <- cbind(drivers, PartFam = substring(drivers$Prob_Part, 1, 6))
	engines <- merge(engines, engineToPartFam, by.x = c("PartFam"), by.y = c("Item.ID"))
	drivers <- merge(drivers, driverToPartFam, by.x = c("PartFam"), by.y = c("Item.ID"))
	engines$Product_Family <- engines$Product.Family
	drivers$Product_Family <- drivers$Product.Family
	engines <- engines[c(2:17)]
	drivers <- drivers[c(2:17)]
	rmaOutput <- rbind(engines, drivers) %>%
		rbind(others) %>%
		as.data.table()
	rmaOutput <- rmaOutput[order(-rmaOutput$RMA),]
	rmaOutput
}

addUserDefs <- function(rma) {
	so_master <- read.csv("Code/supportFiles/so_User_Defs.csv")
	rma <- left_join(rma, so_master, by = c("orig_SO" = "Sales.Order.ID"))
	rma$Specifier <- rma$Specifier.User.Def.3
	rma$Job_Name <- rma$Job.Name.User.Def.5
	rma <- rma[1:16]
	rma
}

















# Read the reason code CSV and change them in the rma table.
addReasonCodes <- function(rma) {
	codes <- read.csv("H:/Code/supportFiles/Reasoncode.csv", header = TRUE)
	rma <- merge(rma, codes, by.x = c("Code"), by.y = c("reason"))[c(2:17)]
	rma
}

# Split into groups for drivers, light engines, and other parts, then 
#	change the product families.
addProductFamilies <- function(rma) {
	engines <- rma[grep("LEM-", rma$Prob_Part),]
	drivers <- rma[grep("SP-[0-9]{3}-[0-9]{4}-", rma$Prob_Part),]
	others <- rma
	others <- others[-grep("LEM-|SP-[0-9]{3}-[0-9]{4}-", others$Prob_Part),]
	engineToPartFam <- read.csv("H:/Code/supportFiles/LightEngineToPartFam.csv")
	names(engineToPartFam) <- c("Item.ID", "Product.Family")
	driverToPartFam <- read.csv("H:/Code/supportFiles/driverToPartFam.csv")
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
	so_master <- read.csv("H:/Code/supportFiles/so_User_Defs.csv")
	rma <- left_join(rma, so_master, by = c("orig_SO" = "Sales.Order.ID"))
	rma$Specifier <- rma$Specifier.User.Def.3
	rma$Job_Name <- rma$Job.Name.User.Def.5
	rma <- rma[1:16]
	rma
}

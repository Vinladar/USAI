library(openxlsx)
library(data.table)

# Read the RMA info
X <- as.data.table(read.csv("Failure_rate/Jan_Dec_2019.csv"))

# New table with the desired information
Items <- X[,c("Item.ID", "Reason.Code", "Return.Qty")]

# Break the datasets into light engines, drivers, and others
# Light engines
lightEngines <- Items[grep("LEM-", Items$Item.ID),]
lightEngineToPartFam <- as.data.table(read.csv("supportFiles/LightEngineToPartFam.csv"))
names(lightEngineToPartFam) <- c("Item.ID", "Product.Family")


# Drivers
drivers <- Items[grep("SP-[0-9]{3}-[0-9]{4}-", Items$Item.ID),]
driverToPartFam <- read.csv("supportFiles/driverToPartFam.csv")
names(driverToPartFam) <- c("Item.ID", "Product.Family")

# Others
others <- Items
others <- others[-grep("LEM-", others$Item.ID),]
others <- others[-grep("SP-[0-9]{3}-[0-9]{4}-", others$Item.ID),]


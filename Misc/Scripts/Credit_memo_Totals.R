# This script is designed to take the result from the RMA Listing report in Intuitive and determine
#   exactly how much we have spent on RMAs.
# It starts with the unedited report and it removes unneeded columns, then removes everything with a restock
#   fee. 
# After that, it removes all rows that don't have any Credit Memo information, then finds the sum of the
#   extended prices in the remaining rows.
library(openxlsx)

calculate_totals <- function(Data, Output) {
  clean_data <- Data[, c(1, 3, 5, 7, 9, 11, 12, 13, 21, 22, 23, 24, 25, 30, 32, 33, 37)]
  names(clean_data) <- c("Customer Corp Name", "Sales Order ID", "Customer ID", "Return Qty", 
                         "Customer PO Line Nbr",  "Discount Percentage", "Item Name", "Product Family", 
                         "Problem Desc", "Credit Memo Date / ID", "Item ID", "Credit Memo ID", "Unit Price", 
                         "RMA ID", "Extended Price", "Credit Memo Complete", "Restocking Percent")
  no_restock <- clean_data[clean_data$`Restocking Percent` == "", ]
  no_restock_credit_memo <- no_restock[grep("..", no_restock$`Credit Memo Date / ID`), ]
  no_restock_credit_memo$`Extended Price` <- as.numeric(no_restock_credit_memo$`Extended Price`)
  final <- no_restock_credit_memo[grep("DEFECTIVE", no_restock_credit_memo$`Problem Desc`), ]
  total <- sum(final$`Extended Price`, na.rm = TRUE)
  write.xlsx(final, Output)
}

# This script is designed to take the result from the RMA Listing report in Intuitive and determine
#   exactly how much we have spent on RMAs.
# It starts with the unedited report and it removes unneeded columns, then removes everything with a restock
#   fee. 
# After that, it removes all rows that don't have any Credit Memo information, then finds the sum of the
#   extended prices in the remaining rows.
library(openxlsx)

raw_data <- read.xlsx("Misc/Credit_Memo_RMAs_2019_to_10_2020.xlsx")
clean_data <- raw_data[, c(1, 3, 5, 7, 9, 11, 12, 13, 21, 22, 23, 24, 25, 30, 32, 33, 37)]
no_restock <- clean_data[clean_data$Restocking.Percent == "",]
no_restock_credit_memo <- no_restock[grep("..", no_restock$`Credit.Memo.Date./.ID`), ]
no_restock_credit_memo$Extended.Price <- as.numeric(no_restock_credit_memo$Extended.Price)
total <- sum(no_restock_credit_memo$Extended.Price, na.rm = TRUE)

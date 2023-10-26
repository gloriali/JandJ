# barcode for Clover
library(dplyr)
library("readxl")
mastersku <- read.delim("../AlphaPay/1-MasterSKU-All-Product-2023-10-25.txt", skip = 3, header = T, as.is = T, colClasses = c(UPC.CA = "character"))
rownames(mastersku) <- mastersku$MSKU
inventory <- read.delim("../VancouverBaby/inventory-vancouverbaby-items.txt")
inventory$Product.Code <- mastersku[inventory$SKU, "UPC.CA"]
write.csv(inventory, file = "../VancouverBaby/inventory-vancouverbaby-barcode.csv", row.names = F, na = "")

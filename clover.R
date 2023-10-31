# barcode for Clover
library(dplyr)
library("readxl")
mastersku <- read.delim("../Clover/1-MasterSKU-All-Product-2023-10-25.txt", skip = 3, header = T, as.is = T, colClasses = c(UPC.CA = "character"))
rownames(mastersku) <- mastersku$MSKU
inventory <- read.delim("../VancouverBaby/inventory-vancouverbaby-items.txt")
inventory$Product.Code <- mastersku[inventory$SKU, "UPC.CA"]
write.csv(inventory, file = "../VancouverBaby/inventory-vancouverbaby-barcode.csv", row.names = F, na = "")

# Combine Vancouver Baby left over with Richmond stock
library(dplyr)
vbaby <- read.csv("../Clover/inventory20231030-items.csv", as.is = T)
rownames(vbaby) <- vbaby$Name
vbaby[is.na(vbaby$Quantity), "Quantity"] <- 0
richmond <- read.csv("../Square/catalog-2023-10-28-1448.csv", as.is = T)
rownames(richmond) <- richmond$SKU
richmond[is.na(richmond$Current.Quantity.Jan...Jul.Richmond), "Current.Quantity.Jan...Jul.Richmond"] <- 0
vbaby$Price <- as.numeric(richmond[vbaby$Name, "Price"])
vbaby$Quantity <- vbaby$Quantity + richmond[vbaby$Name, "Current.Quantity.Jan...Jul.Richmond"]
vbaby[vbaby$Quantity < 0, "Quantity"] <- 0
vbaby[grep("NA", vbaby$Product.Code), "Product.Code"] <- ""
write.csv(vbaby, file = "../Clover/inventory20231030-items-upload.csv", row.names = F, na = "")

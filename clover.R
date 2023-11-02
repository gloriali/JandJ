# -------- woo price for clover -------- 
library(dplyr)
woo <- read.csv("../woo/wc-product-export-31-10-2023-1698777245011.csv", as.is = T)
woo1 <- woo %>% filter(!is.na(woo$Regular.price) & !duplicated(woo$SKU) & !(woo$SKU == "")) 
rownames(woo1) <- woo1$SKU
sales <- woo1 %>% filter(!is.na(woo1$Sale.price)) 
rownames(sales) <- sales$SKU

clover <- read.csv("../Clover/inventory20231031-items.csv", as.is = T)
rownames(clover) <- clover$Name
clover$Price <- woo1[clover$Name, "Regular.price"]
clover[sales$SKU, "Price"] <- sales[sales$SKU, "Sale.price"]
write.csv(clover, file = "../Clover/inventory20231031-items-upload.csv", row.names = F, na = "")

# -------- master file barcode for Clover -------------
library(dplyr)
mastersku <- read.csv("../woo/1-MasterSKU-All-Product-2023-10-25.csv", skip = 3, header = T, as.is = T, colClasses = c(UPC.Active = "character"))
rownames(mastersku) <- mastersku$MSKU
clover <- read.csv("../Clover/inventory20231031-items.csv", as.is = T)
clover$Product.Code <- mastersku[clover$SKU, "UPC.Active"]
write.csv(clover, file = "../Clover/inventory20231031-items-upload.csv", row.names = F, na = "")

# -------- Combine Vancouver Baby left over with Richmond stock -------- 
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

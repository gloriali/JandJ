# Clover system management
# -------- woo price for clover -------- 
# download current woo price: wc > Products > All Products > Export.
# download clover inventory: clover_item > Inventory > Items > Export.
library(dplyr)
library(openxlsx)
woo <- read.csv(paste0("../woo/", list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))))), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") 
rownames(woo) <- woo$SKU
sales <- woo %>% filter(!is.na(Sale.price)) 
rownames(sales) <- sales$SKU

clover <- loadWorkbook(paste0("../Clover/", list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"))))
clover_item <- readWorkbook(clover, "Items") %>%
  mutate(Price = ifelse(Name %in% sales$SKU, sales[Name, "Sale.price"], woo[Name, "Regular.price"]))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
writeData(clover, "Items", clover_item)
saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"))
# upload to Clover > Inventory

# ---------------- request stock from Surrey warehouse --------------------
# download current Surrey stock: xoro > Item Inventory Snapshot > Export all - csv; 
# download current Richmond stock: clover_item > Inventory > Items > Export
library(dplyr)
library(openxlsx)
xoro <- read.csv(paste0("../xoro/", list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"))), as.is = T) %>% filter(Store == "Warehouse - JJ")
rownames(xoro) <- xoro$Item.

clover <- loadWorkbook(paste0("../Clover/", list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"))))
clover_item <- readWorkbook(clover, "Items") %>% mutate(cat = gsub("-.*", "", Name))
clover_item[is.na(clover_item$Quantity) | clover_item$Quantity < 0, "Quantity"] <- 0
rownames(clover_item) <- clover_item$Name

n <- 1 # Qty per SKU to stock at Richmond
n_xoro <- 10 # min Qty in stock at Surrey to request
request <- c("BCV", "WPF", "WJA") # categories to restock
order <- data.frame(StoreCode = "WH-JJ", ItemNumber=(clover_item %>% filter(Quantity < n & cat %in% request))$Name, Qty = n - clover_item[(clover_item %>% filter(Quantity < n & cat %in% request))$Name, "Quantity"], LocationName = "BIN", UnitCost = "", ReasonCode = "RWT", Memo = "Richmond Transfer to Miranda", UploadRule = "D", AdjAccntName = "", TxnDate = "", ItemIdentifierCode = "", ImportError = "")
order <- order %>% filter(xoro[order$ItemNumber, "ATS"] > n_xoro) %>% filter(Qty > 0) %>% mutate(cat = gsub("-.*", "", ItemNumber), size = gsub("\\w+-\\w+-", "", ItemNumber)) %>% arrange(cat, size) %>% select(-c("cat", "size"))
write.csv(order, file = paste0("../Clover/order", format(Sys.Date(), "%m%d%Y"), ".csv"), row.names = F, na = "")
# email Shikshit

# -------- master file barcode for clover_item -------------
# master SKU file: OneDrive > TWK 2020 share
# download clover inventory: clover_item > Inventory > Items > Export.
library(dplyr)
library(openxlsx)
mastersku <- read.xlsx(paste0("../../TWK 2020 share/", list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-")), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU

clover <- loadWorkbook(paste0("../Clover/", list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"))))
clover_item <- readWorkbook(clover, "Items") %>% mutate(Product.Code = mastersku[SKU, "UPC.Active"])
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
writeData(clover, "Items", clover_item)
saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"))
# upload to Clover > Inventory

# -------- Black Friday sales price -------- 
# download current woo price: wc > Products > All Products > Export.
# download clover inventory: clover_item > Inventory > Items > Export.
library(dplyr)
library(openxlsx)
woo <- read.csv(paste0("../woo/", list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))))), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & !(SKU == "")) 
rownames(woo) <- woo$SKU

clover <- loadWorkbook(paste0("../Clover/", list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"))))
clover_item <- readWorkbook(clover, "Items") %>% mutate(Price = woo[Name, "Regular.price"]*0.8)
clover_item[grepl("WSS-CNL", clover_item$Name), "Price"] <- 39
clover_item[grepl("WSS-DNL", clover_item$Name), "Price"] <- 39
clover_item[grepl("WSS-UNC", clover_item$Name), "Price"] <- 39
clover_item[grepl("WSS-TRZ", clover_item$Name), "Price"] <- 39
clover_item[grepl("BRC", clover_item$Name), "Price"] <- 34
clover_item[grepl("LBS", clover_item$Name), "Price"] <- 27
clover_item[grepl("AWWJ", clover_item$Name), "Price"] <- 59
clover_item[grepl("ICP", clover_item$Name), "Price"] <- 69
clover_item[grepl("SWS", clover_item$Name), "Price"] <- 25
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
writeData(clover, "Items", clover_item)
saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"))

# -------- Combine Vancouver Baby left over with Richmond stock -------- 
library(dplyr)
vbaby <- read.csv("../clover_item/inventory20231030-items.csv", as.is = T)
rownames(vbaby) <- vbaby$Name
vbaby[is.na(vbaby$Quantity), "Quantity"] <- 0
richmond <- read.csv("../Square/catalog-2023-10-28-1448.csv", as.is = T)
rownames(richmond) <- richmond$SKU
richmond[is.na(richmond$Current.Quantity.Jan...Jul.Richmond), "Current.Quantity.Jan...Jul.Richmond"] <- 0
vbaby$Price <- as.numeric(richmond[vbaby$Name, "Price"])
vbaby$Quantity <- vbaby$Quantity + richmond[vbaby$Name, "Current.Quantity.Jan...Jul.Richmond"]
vbaby[vbaby$Quantity < 0, "Quantity"] <- 0
vbaby[grep("NA", vbaby$Product.Code), "Product.Code"] <- ""
write.csv(vbaby, file = "../clover_item/inventory20231030-items-upload.csv", row.names = F, na = "")

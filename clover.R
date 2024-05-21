# Clover system management
# -------- update clover with woo -------- 
# download current woo: wc > Products > All Products > Export.
# download clover inventory: clover_item > Inventory > Items > Export.
library(dplyr)
library(openxlsx)
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price), cat = toupper(gsub("-.*", "", SKU)))
rownames(woo) <- woo$SKU
price <- woo %>% mutate(cat = gsub("-.*", "", SKU)) %>% group_by(cat) %>% summarise(Price = max(Sale.price)) %>% as.data.frame()
price <- rbind(price, data.frame(cat = c("MSWS"), Price = c(25))) %>% `row.names<-`(toupper(.[, "cat"])) 
#clover <- openxlsx::loadWorkbook(rownames(file.info(list.files(path = "../Clover/", pattern = "inventory[0-9]+.xlsx", full.names = TRUE)) %>% filter(mtime == max(mtime))))
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% 
  mutate(cat = toupper(gsub("-.*", "", Name)), Price = ifelse(Name %in% woo$SKU, woo[Name, "Sale.price"], price[cat, "Price"]), Price.Type = ifelse(is.na(Price), "Variable", "Fixed"), Alternate.Name = woo[Name, "Name"], Tax.Rates = ifelse(woo[Name, "Tax.class"] == "full", "GST+PST", "GST"), Tax.Rates = ifelse(is.na(Tax.Rates), "GST", Tax.Rates)) %>% select(-cat)
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item), gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
# upload to Clover > Inventory

# ---------------- request stock from Surrey warehouse --------------------
# download current Surrey stock: xoro > Item Inventory Snapshot > Export all - csv; 
# download current Richmond stock: clover_item > Inventory > Items > Export
library(dplyr)
library(openxlsx)
xoro <- read.csv(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"), full.names = T), as.is = T) %>% filter(Store == "Warehouse - JJ")
rownames(xoro) <- xoro$Item.
clover <- loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
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

# ---------------- adjust inventory --------------------
# download current Richmond stock: clover_item > Inventory > Items > Export
library(dplyr)
library(openxlsx)
adjust_inventory <- read.csv("../Clover/order04302024.csv", as.is = T) %>% `row.names<-`(toupper(.[, "ItemNumber"])) 
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- readWorkbook(clover, "Items") %>% mutate(SKU = toupper(SKU), Quantity = ifelse(SKU %in% rownames(adjust_inventory), Quantity + adjust_inventory[SKU, "Qty"], Quantity))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item), gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)

# -------- sync master file with clover_item -------------
# master SKU file: OneDrive > TWK 2020 share
# download clover inventory: clover_item > Inventory > Items > Export.
library(dplyr)
library(openxlsx)
library(tidyr)
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price), cat = toupper(gsub("-.*", "", SKU))) %>% `row.names<-`(toupper(.[, "SKU"])) 
price <- woo %>% group_by(cat) %>% summarise(Price = max(Sale.price)) %>% as.data.frame()
price <- rbind(price, data.frame(cat = c("MSWS"), Price = c(25))) %>% `row.names<-`(toupper(.[, "cat"])) 
mastersku <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../FBArefill/Raw Data File/", pattern = "1-MasterSKU-All-Product-", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% filter(!is.na(Name), !duplicated(toupper(Name))) %>% `row.names<-`(toupper(.[, "Name"])) 
category_exclude <- c("FVMU", "MFJM", "MSKB", "MSXP", "MWBF", "MWBS", "MWJA", "MWMT", "MWPF", "MWPS", "MWSF", "MWSS", "PJWU","PLAU", "PTAU", "STORE", "TPLU", "TPSU", "TSWU", "TTLU", "TTSU", "WJAU", "WPSU", "XBMU")
clover_item_upload <- data.frame(Clover.ID = "", Name = toupper(mastersku[mastersku$MSKU.Status == "Active", "MSKU"])) %>% filter(!is.na(Name)) %>% 
  mutate(cat = mastersku[Name, "Category.SKU"], Alternate.Name = woo[Name, "Name"], Price = ifelse(Name %in% woo$SKU, woo[Name, "Sale.price"], price[cat, "Price"]), Price.Type = ifelse(is.na(Price), "Variable", "Fixed"), Price.Unit = NA, Tax.Rates = ifelse(woo[Name, "Tax.class"] == "full", "GST+PST", "GST"), Tax.Rates = ifelse(is.na(Tax.Rates), "GST", Tax.Rates), Cost = 0, Product.Code = gsub("/.*", "", mastersku[Name, "UPC.Active"]), SKU = Name, Modifier.Groups = NA, Quantity = ifelse(Name %in% clover_item$Name, clover_item[Name, "Quantity"], 0), Printer.Labels = NA, Hidden = "No", Non.revenue.item = "No") %>%
  filter(!(cat %in% category_exclude)) %>% arrange(cat, Name)
clover_cat_upload <- clover_item_upload %>% group_by(cat) %>% mutate(id = row_number(cat)) %>% select(cat, id, Name) %>% spread(id, Name, fill = "") %>% t() 
clover_cat_upload <- cbind(data.frame(V0 = c("Category Name", "Items in Category", rep("", times = nrow(clover_cat_upload) - 2))), clover_cat_upload)
clover_item_upload <- clover_item_upload %>% select(-cat)
clover_item_upload <- clover_item_upload %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item)+1000, gridExpand = T)
writeData(clover, sheet = "Items", clover_item_upload)
deleteData(clover, sheet = "Categories", cols = 1:1000, rows = 1:1000, gridExpand = T)
writeData(clover, sheet = "Categories", clover_cat_upload, colNames = F)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
# upload to Clover > Inventory
# upload to Clover > Inventory

# -------- Inventory recount -------- 
library(dplyr)
library(openxlsx)
clover <- loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_qty <- readWorkbook(clover, "Items") %>% select(SKU, Quantity) %>% mutate(New_qty = "", cat = gsub("-.*", "", SKU), size = gsub("^\\w+-\\w+-", "", SKU)) %>% arrange(cat, size) %>% select(SKU, New_qty, Quantity)
write.xlsx(clover_qty, file = paste0("../Clover/recount_inventory-", format(Sys.Date(), "%Y%m%d"), ".xlsx"))
# recount inventory
clover_qty1 <- read.xlsx("../Clover/recount_inventory-20231219.xlsx", sheet = 1, skipEmptyCols = T) %>% mutate(qty_1 = ifelse(is.na(New_qty) | New_qty == "", 0, as.numeric(New_qty)))
clover_qty_g <- read.xlsx("../Clover/recount_inventory-20240103gloria.xlsx", sheet = 1, skipEmptyCols = T) %>% mutate(qty_g = ifelse(is.na(New_qty) | New_qty == "", 0, as.numeric(New_qty)))
clover_qty_m <- read.xlsx("../Clover/recount_inventory-20240103minranda.xlsx", sheet = 1, skipEmptyCols = T) %>% mutate(qty_m = ifelse(is.na(New_qty) | New_qty == "", 0, as.numeric(New_qty)))
clover_qty_s <- read.xlsx("../Clover/recount_inventory-20240102shirley.xlsx", sheet = 1, skipEmptyCols = T) %>% mutate(qty_s = ifelse(is.na(New_qty) | New_qty == "", 0, as.numeric(New_qty)))
clover_qty <- merge(merge(clover_qty1, clover_qty_g, by = "SKU"), merge(clover_qty_m, clover_qty_s, by = "SKU"), by = "SKU") %>%
  mutate(Quantity = Quantity.x.x, New_qty = qty_g + qty_m + qty_s - qty_1 * 2) %>% select(SKU, Quantity, New_qty)
diff <- clover_qty[clover_qty$New_qty != clover_qty$Quantity, ] %>% arrange(desc(abs(Quantity-New_qty)))
write.csv(clover_qty, file = "../Clover/recount_inventory-20240103_final.csv", row.names = F, quote = F)
clover_qty <- read.xlsx("../Clover/recount_inventory-20240103_final.xlsx", sheet = 1)
rownames(clover_qty) <- clover_qty$SKU
clover <- loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- readWorkbook(clover, "Items") %>% mutate(Quantity = clover_qty[SKU, "qty"])
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
writeData(clover, "Items", clover_item)
saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"))

# -------- Black Friday sales price -------- 
# download current woo price: wc > Products > All Products > Export.
# download clover inventory: clover_item > Inventory > Items > Export.
library(dplyr)
library(openxlsx)
woo <- read.csv(list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & !(SKU == "")) 
rownames(woo) <- woo$SKU

clover <- loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
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

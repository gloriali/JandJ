# Routine tasks
library(dplyr)
library(stringr)
library(openxlsx)

# input: woo > Products > All Products > Export > all
# input: xoro > Item Inventory Snapshot > Export all
woo <- read.csv(list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price))
rownames(woo) <- woo$SKU
xoro <- read.csv(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"), full.names = T), as.is = T) %>% filter(Store == "Warehouse - JJ") %>% mutate(Item. = toupper(Item.))
rownames(xoro) <- xoro$Item.

# -------- Sync XHS price to woo; inventory to xoro: Mon Wed Fri ----------------
# input: AllValue > Products > Export > All products
products_XHS <- read.xlsx(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), sheet = 1)
products_upload <- products_XHS %>% mutate(Inventory = ifelse(xoro[SKU, "ATS"] < 20, 0, xoro[SKU, "ATS"]), Price = woo[SKU, "Sale.price"], Compare.At.Price = woo[SKU, "Regular.price"], Tags = ifelse(Price == Compare.At.Price, "正价", "特价"))
colnames(products_upload) <- gsub("\\.", " ", colnames(products_upload))
write.xlsx(products_upload, file = paste0("../XHS/products_upload-", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"))
# upload to AllValue > Products

# -------- Sync Square price to woo; inventory to xoro: weekly -------------
square <- data.frame(Token = "", ItemName = woo$SKU, VariationName = "Regular", SKU = woo$SKU, Description = "", Category = xoro[woo$SKU, "Item.Category"], Price = woo$Sale.price, PriceRichmond = "", PriceSurrey = "", NewQuantityRichmond = xoro[woo$SKU, "ATS"], NewQuantitySurrey = xoro[woo$SKU, "ATS"], EnabledRichmond = "Y", CurrentQuantityRichmond = "", CurrentQuantitySurrey = "", StockAlertEnabledRichmond = "N", StockAlertCountRichmond = 0, EnabledSurrey = "Y", StockAlertEnabledSurrey = "N", StockAlertCountSurrey = 0, SEOTitle = "",	SEODescription = "",	Permalink = "", SquareOnlineItemVisibility = "UNAVAILABLE",	ShippingEnabled = "",	SelfserveOrderingEnabled = "",	DeliveryEnabled = "",	PickupEnabled = "", Sellable = "Y",	Stockable = "Y",	SkipDetailScreeninPOS = "",	OptionName1 = "",	OptionValue1 = "", TaxPST = ifelse(woo$Tax.class == "full", "Y", "N")) %>%
  mutate(NewQuantityRichmond = ifelse(is.na(NewQuantityRichmond), 0, NewQuantityRichmond), NewQuantitySurrey = ifelse(is.na(NewQuantitySurrey), 0, NewQuantitySurrey))
write.table(square, file = paste0("../Square/square-upload-", Sys.Date(), ".csv"), sep = ",", row.names = F, col.names = c(colnames(square)[1:ncol(square)-1], "Tax - PST (7%)"), na = "")
# upload to Square > Items

# -------- Sync Clover to woo: weekly -------------
# input Clover > Inventory > Items > Export.
clover <- loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- readWorkbook(clover, "Items") %>% mutate(Price = woo[Name, "Sale.price"], Alternate.Name = woo[Name, "Name"], Tax.Rates = ifelse(woo$Tax.class == "full", "GST+PST", "PST"))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
writeData(clover, "Items", clover_item)
saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
# upload to Clover > Inventory

# -------- Add POS sales to yotpo rewards program: weekly -----------
# customer info: Clover > Customers > Download
# sales: Clover > Transactions > Payments > select dates > Export
customer <- read.csv(list.files(path = "../Clover/", pattern = paste0("Customers-", format(Sys.Date(), "%Y%m%d")), full.names = T), as.is = T) %>%
  mutate(Name = paste0(First.Name, " ", Last.Name)) %>% distinct(Email.Address, .keep_all = T) %>% filter(Name != " ", Email.Address != "")
rownames(customer) <- customer$Name
payments <- read.csv(list.files(path = "../Clover/", pattern = paste0("Payments-", format(Sys.Date(), "%Y%m%d")), full.names = T), as.is = T) %>% 
  mutate(email = customer[Customer.Name, "Email.Address"], Refund.Amount = ifelse(is.na(Refund.Amount), 0, as.numeric(Refund.Amount))) %>% filter(!is.na(email))
point <- data.frame(Email = payments$email, Points = as.integer(payments$Amount - payments$Refund.Amount))
# customer info: Square > Customers > Export customers
# sales: Square > Transactions > select dates > Export Transactions CSV
square_customer <- read.csv("../Square/customers.csv", as.is = T)
rownames(square_customer) <- square_customer$Square.Customer.ID
transactions <- read.csv("../Square/transactions-2023-11-09-2023-11-17.csv", as.is = T) %>% filter(Customer.ID != "")
point <- bind_rows(point, data.frame(Email = square_customer[transactions$Customer.ID, "Email.Address"], Points = as.integer(gsub("\\$", "", transactions$Total.Collected)))) %>% filter(Points != 0, !is.na(Email))
write.csv(point, file = paste0("../yotpo/", format(Sys.Date(), "%m%d%Y"), "-yotpo.csv"), row.names = F)
# upload to Yotpo
non_included <- read.csv("../yotpo/non_included.csv", as.is = T)
error <- read.csv("../yotpo/error_report.csv", as.is = T)
non_included <- bind_rows(non_included, error %>% select(Email, Points)) %>% count(Email, wt = Points)
write.table(non_included, file = "../yotpo/non_included.csv", sep = ",", row.names = F, col.names = c("Email", "Points"))

# -------- Request stock from Surrey: at request --------------------
# input Clover > Inventory > Items > Export.
request <- c("BCV", "WPF", "WJA") # categories to restock
n <- 1       # Qty per SKU to stock at Richmond
n_xoro <- 10 # min Qty in stock at Surrey to request
clover <- loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- readWorkbook(clover, "Items") %>% mutate(cat = gsub("-.*", "", Name))
clover_item[is.na(clover_item$Quantity) | clover_item$Quantity < 0, "Quantity"] <- 0
rownames(clover_item) <- clover_item$Name
order <- data.frame(StoreCode = "WH-JJ", ItemNumber=(clover_item %>% filter(Quantity < n & cat %in% request))$Name, Qty = n - clover_item[(clover_item %>% filter(Quantity < n & cat %in% request))$Name, "Quantity"], LocationName = "BIN", UnitCost = "", ReasonCode = "RWT", Memo = "Richmond Transfer to Miranda", UploadRule = "D", AdjAccntName = "", TxnDate = "", ItemIdentifierCode = "", ImportError = "")
order <- order %>% filter(xoro[order$ItemNumber, "ATS"] > n_xoro) %>% filter(Qty > 0) %>% mutate(cat = gsub("-.*", "", ItemNumber), size = gsub("\\w+-\\w+-", "", ItemNumber)) %>% arrange(cat, size) %>% select(-c("cat", "size"))
write.csv(order, file = paste0("../Clover/order", format(Sys.Date(), "%m%d%Y"), ".csv"), row.names = F, na = "")
# email Shikshit

# -------- Discontinued SKUs management: monthly --------------------
library(dplyr)
library(openxlsx)
library(scales)
new_season <- "24"   # New season contains
in_season <- "F"     # "F": Sept - Feb; "S": Mar - Aug
qty_offline <- 3     # Qty to move to offline sales
qty_sold <- 50       # Sales rate low: base Qty sold on website last month
size_percent <- 0.5  # % of sizes not available 
mastersku <- read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
xoro <- read.csv(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"), full.names = T), as.is = T) %>% filter(Store == "Warehouse - JJ") %>% 
  mutate(seasons = mastersku[Item., "Seasons"], cat_print = gsub("(\\w+-\\w+)-.*", "\\1", Item.))
rownames(xoro) <- xoro$Item.
## move offline 
offline <- xoro %>% filter(!grepl(new_season, seasons) & ATS <= qty_offline & ATS > 0) %>% arrange(Item.)
write.csv(offline, file = paste0("../SKUmanagement/offline_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\
## suggest to website deals page 
woo <- read.csv(list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price), discount = (Regular.price - Sale.price)/Regular.price)
rownames(woo) <- woo$SKU
xoro_lastmonth <- read.csv("../xoro/Item Inventory Snapshot_11102023.csv", as.is = T) %>% filter(Store == "Warehouse - JJ") 
rownames(xoro_lastmonth) <- xoro_lastmonth$Item.
JJ_month_ratio <- read.xlsx("../FBArefill/Historic sales and inv. data for all cats v29 (20231205).xlsx", sheet = "MonSaleR", startRow = 2, cols = c(1, 26:37)) %>% filter(!is.na(Category)) 
JJ_month_ratio <- JJ_month_ratio %>% mutate(monthR = as.integer(format(Sys.Date(), "%d"))/31*JJ_month_ratio[, gsub("0", "", format(Sys.Date(), "%m"))] + (31 - as.integer(format(Sys.Date(), "%d")))/31*JJ_month_ratio[, as.integer(format(Sys.Date(), "%m"))])
rownames(JJ_month_ratio) <- JJ_month_ratio$Category
sizes_all <- xoro %>% mutate(cat_print_size = gsub("(\\w+-\\w+-\\w+).*", "\\1", Item.)) %>% filter(!duplicated(cat_print_size)) %>% count(cat_print)
rownames(sizes_all) <- sizes_all$cat_print
size_limited <- xoro %>% filter(!grepl(new_season, seasons) & ATS > qty_offline) %>% count(cat_print) %>% mutate(all = sizes_all[cat_print, "n"], percent = (all - n)/all) %>% filter(percent >= size_percent & percent < 1)
rownames(size_limited) <- size_limited$cat_print
deals <- xoro %>% mutate(lastmonth = xoro_lastmonth[Item., "ATS"], cat = gsub("-.*", "", Item.), threshold = ifelse(cat %in% JJ_month_ratio$Category, qty_sold*JJ_month_ratio[cat, "monthR"], ifelse(grepl(in_season, seasons), 8, 2))) %>% filter(cat_print %in% size_limited$cat_print & !grepl(new_season, seasons) & ATS > qty_offline & lastmonth - ATS < threshold & lastmonth - ATS > 0) %>% 
  mutate(size_percent = percent(size_limited[cat_print, "percent"]), discount = percent(woo[Item., "discount"], accuracy = 0.1)) %>% arrange(Item.)
write.csv(deals, file = paste0("../SKUmanagement/deals_", Sys.Date(), ".csv"), row.names = F, na = "")
# email Joren and Kamer
## to inactivate 
clover <- read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items")
rownames(clover) <- clover$SKU
names <- grep("-WK", getSheetNames("../FBArefill/Historic sales and inv. data for all cats v29 (20231205).xlsx"), value = T)
inventory <- data.frame()
for(name in names){
  cat_inventory <- read.xlsx("../FBArefill/Historic sales and inv. data for all cats v29 (20231205).xlsx", sheet = name, startRow = 2)
  inventory <- rbind(inventory, cat_inventory)
}
inventory_s <- inventory %>% count(Adjust_MSKU, wt = Inv_Total_End.WK)
rownames(inventory_s) <- inventory_s$Adjust_MSKU
qty0 <- mastersku %>% select(MSKU.Status, Seasons.SKU, MSKU) %>% mutate(Inv_clover = clover[MSKU, "Quantity"], Inv_Total_EndWK = inventory_s[MSKU, "n"]) %>% 
  filter(!grepl(new_season, Seasons.SKU) & MSKU.Status == "Active" & Inv_clover == 0 & Inv_Total_EndWK == 0)
write.csv(qty0, file = paste0("../SKUmanagement/discontinued_qty0_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\

# -------- Prepare to generate barcode image: at request ------------
library(dplyr)
library(openxlsx)
POn <- "P257"
startRow <- 9
mastersku <- read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
barcode <- read.xlsx(list.files(path = "../../TWK 2020 share/twk general/1-orders (formerly upcoming shipments)/", pattern = paste0(POn, ".*.xlsx"), full.names = T, recursive = T), sheet = 1, startRow = startRow, cols = readcols) %>% select(SKU, Design.Version) %>% 
  mutate(SKU = str_trim(SKU), Print.English = mastersku[SKU, "Print.Name"], Print.Chinese = mastersku[SKU, "Print.Chinese"], Size = mastersku[SKU, "Size"], UPC.Active = mastersku[SKU, "UPC.Active"], Image = "") %>% filter(SKU != "", !is.na(SKU))
write.xlsx(barcode, file = paste0("../Clover/barcode", format(Sys.Date(), "%Y%m%d"), ".xlsx"))
# Email Shirley

# -------- Sync master file barcode with clover: at request -------------
# master SKU file: OneDrive > TWK 2020 share
# download clover inventory: clover_item > Inventory > Items > Export.
library(dplyr)
library(openxlsx)
new_season <- "24"
mastersku <- read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
clover <- loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- readWorkbook(clover, "Items") %>% mutate(Product.Code = mastersku[SKU, "UPC.Active"])
new_items <- data.frame(Clover.ID = "", Name = mastersku[mastersku$MSKU.Status == "Active" & grepl(new_season, mastersku$Seasons) & !(mastersku$MSKU %in% clover_item$Name), "MSKU"]) %>%
  mutate(Alternate.Name = "", Price = 0, Price.Type = "Fixed", Price.Unit = NA, Tax.Rates = "GST", Cost = 0, Product.Code = mastersku[Name, "UPC.Active"], SKU = Name, Modifier.Groups = NA, Quantity = 0, Printer.Labels = NA, Hidden = "No", Non.revenue.item = "No")
new_items <- new_items %>% rename_with(~ gsub("Non.revenue.item", "Non-revenue.item", colnames(new_items)))
clover_item <- rbind(clover_item, new_items)
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
writeData(clover, "Items", clover_item)
saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
# upload to Clover > Inventory

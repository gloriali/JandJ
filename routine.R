# Routine tasks
library(dplyr)
library(stringr)
library(openxlsx)
library(xlsx)
library(scales)
library(ggplot2)

# input: woo > Products > All Products > Export > all
# input: xoro > Item Inventory Snapshot > Export all
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price)) %>% `row.names<-`(toupper(.[, "SKU"])) 
xoro <- read.xlsx2(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".xlsx"), full.names = T), sheetIndex = 1) %>% filter(Store == "Warehouse - JJ") %>% mutate(Item. = toupper(Item.), ATS = as.numeric(ATS)) %>% `row.names<-`(toupper(.[, "Item."])) 

# -------- Sync XHS price to woo; inventory to xoro: weekly ----------------
# input: AllValue > Products > Export > All products
products_description <- read.xlsx2("../XHS/products_description.xlsx", sheetIndex = 1) %>% `row.names<-`(toupper(.[, "SPU"])) 
wb <- openxlsx::loadWorkbook(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T))
openxlsx::protectWorkbook(wb, protect = F)
openxlsx::saveWorkbook(wb, list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), overwrite = T)
products_XHS <- read.xlsx2(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), sheetIndex = 1)
products_upload <- products_XHS %>% mutate(Inventory = ifelse(xoro[SKU, "ATS"] < 10, 0, xoro[SKU, "ATS"]), Price = woo[SKU, "Sale.price"], Compare.At.Price = woo[SKU, "Regular.price"], Tags = ifelse(Price == Compare.At.Price, "正价", "特价"))
products_upload <- products_upload %>% mutate(Product.Name = products_description[toupper(SPU), "Product.Name"], SEO.Product.Name = Product.Name, Description = products_description[toupper(SPU), "Description"], Mobile.Description = products_description[toupper(SPU), "Description"], SEO.Description = products_description[toupper(SPU), "Description"], Describe = products_description[toupper(SPU), "Description"])
colnames(products_upload) <- gsub("\\.", " ", colnames(products_upload))
openxlsx::write.xlsx(products_upload, file = paste0("../XHS/products_upload-", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"))
# upload to AllValue > Products

# -------- Sync Square price to woo; inventory to xoro: weekly -------------
square <- data.frame(Token = "", ItemName = woo$SKU, ItemType = "", VariationName = "Regular", SKU = woo$SKU, Description = "", Category = xoro[woo$SKU, "Item.Category"], Price = woo$Sale.price, PriceRichmond = "", PriceSurrey = "", NewQuantityRichmond = xoro[woo$SKU, "ATS"], NewQuantitySurrey = xoro[woo$SKU, "ATS"], EnabledRichmond = "Y", CurrentQuantityRichmond = "", CurrentQuantitySurrey = "", StockAlertEnabledRichmond = "N", StockAlertCountRichmond = 0, EnabledSurrey = "Y", StockAlertEnabledSurrey = "N", StockAlertCountSurrey = 0, SEOTitle = "",	SEODescription = "",	Permalink = "", SquareOnlineItemVisibility = "UNAVAILABLE",	ShippingEnabled = "",	SelfserveOrderingEnabled = "",	DeliveryEnabled = "",	PickupEnabled = "", Sellable = "Y",	Stockable = "Y",	SkipDetailScreeninPOS = "",	OptionName1 = "",	OptionValue1 = "", TaxPST = ifelse(woo$Tax.class == "full", "Y", "N")) %>%
  mutate(NewQuantityRichmond = ifelse(is.na(NewQuantityRichmond), 0, NewQuantityRichmond), NewQuantitySurrey = ifelse(is.na(NewQuantitySurrey), 0, NewQuantitySurrey))
write.table(square, file = paste0("../Square/square-upload-", Sys.Date(), ".csv"), sep = ",", row.names = F, col.names = c(colnames(square)[1:ncol(square)-1], "Tax - PST (7%)"), na = "")
# upload to Square > Items

# -------- Sync Clover to woo: weekly -------------
# input Clover > Inventory > Items > Export.
price <- woo %>% mutate(cat = gsub("-.*", "", SKU)) %>% group_by(cat) %>% summarise(Price = max(Sale.price)) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "cat"])) 
#clover <- openxlsx::loadWorkbook(rownames(file.info(list.files(path = "../Clover/", pattern = "inventory[0-9]+.xlsx", full.names = TRUE)) %>% filter(mtime == max(mtime))))
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% 
  mutate(cat = toupper(gsub("-.*", "", SKU)), Price = ifelse(Name %in% woo$SKU, woo[Name, "Sale.price"], price[cat, "Price"]), Price.Type = ifelse(is.na(Price), "Variable", "Fixed"), Alternate.Name = woo[Name, "Name"], Tax.Rates = ifelse(woo[Name, "Tax.class"] == "full", "GST+PST", "GST"), Tax.Rates = ifelse(is.na(Tax.Rates), "GST", Tax.Rates)) %>% select(-cat)
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item), gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
# upload to Clover > Inventory

# -------- Add POS sales to yotpo rewards program: weekly -----------
# customer info: Clover > Customers > Download
# sales: Clover > Transactions > Payments > select dates > Export
customer <- read.csv(list.files(path = "../Clover/", pattern = paste0("Customers-", format(Sys.Date(), "%Y%m%d")), full.names = T), as.is = T) %>%
  mutate(Name = paste0(First.Name, " ", Last.Name)) %>% distinct(Email.Address, .keep_all = T) %>% filter(Name != " ", Email.Address != "") %>% `row.names<-`(.[, "Name"])
payments <- read.csv(list.files(path = "../Clover/", pattern = paste0("Payments-", format(Sys.Date(), "%Y%m%d")), full.names = T), as.is = T) %>% 
  mutate(email = customer[Customer.Name, "Email.Address"], Refund.Amount = ifelse(is.na(Refund.Amount), 0, as.numeric(Refund.Amount))) %>% filter(!is.na(email))
point <- data.frame(Email = payments$email, Points = as.integer(payments$Amount - payments$Refund.Amount))
# customer info: Square > Customers > Export customers
# sales: Square > Transactions > select dates > Export Transactions CSV
square_customer <- read.csv("../Square/customers.csv", as.is = T) %>% `row.names<-`(.[, "Square.Customer.ID"]) 
transactions <- read.csv("../Square/transactions-2024-03-20-2024-04-11.csv", as.is = T) %>% filter(Customer.ID != "")
point <- bind_rows(point, data.frame(Email = square_customer[transactions$Customer.ID, "Email.Address"], Points = as.integer(gsub("\\$", "", transactions$Total.Collected)))) %>% filter(Points != 0, !is.na(Email))
write.csv(point, file = paste0("../yotpo/", format(Sys.Date(), "%m%d%Y"), "-yotpo.csv"), row.names = F)
# upload to Yotpo
non_included <- read.csv("../yotpo/non_included.csv", as.is = T)
error <- read.csv("../yotpo/error_report.csv", as.is = T)
non_included <- bind_rows(non_included, error %>% select(Email, Points)) %>% count(Email, wt = Points)
write.table(non_included, file = "../yotpo/non_included.csv", sep = ",", row.names = F, col.names = c("Email", "Points"))

# -------- Request stock from Surrey: at request --------------------
# input Clover > Inventory > Items > Export.
request <- c("GHA", "GBX", "SPW", "HAV0", "HAD0", "UG1") # categories to restock
n <- 3       # Qty per SKU to stock at Richmond
n_xoro <- 10 # min Qty in stock at Surrey to request
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- readWorkbook(clover, "Items") %>% mutate(cat = gsub("-.*", "", Name), Quantity = ifelse(is.na(Quantity) | Quantity < 0, 0, Quantity)) %>% filter(!duplicated(Name), !is.na(Name)) %>% `row.names<-`(toupper(.[, "Name"])) 
order <- data.frame(StoreCode = "WH-JJ", ItemNumber=(clover_item %>% filter(Quantity < n & cat %in% request))$Name, Qty = n - clover_item[(clover_item %>% filter(Quantity < n & cat %in% request))$Name, "Quantity"], LocationName = "BIN", UnitCost = "", ReasonCode = "RWT", Memo = "Richmond Transfer to Miranda", UploadRule = "D", AdjAccntName = "", TxnDate = "", ItemIdentifierCode = "", ImportError = "")
order <- order %>% filter(xoro[order$ItemNumber, "ATS"] > n_xoro) %>% filter(Qty > 0) %>% mutate(cat = gsub("-.*", "", ItemNumber), size = gsub("\\w+-\\w+-", "", ItemNumber)) %>% arrange(cat, size) %>% select(-c("cat", "size"))
write.csv(order, file = paste0("../Clover/order", format(Sys.Date(), "%m%d%Y"), ".csv"), row.names = F, na = "")
# email Shikshit

# ---------------- Adjust Clover Inventory: at request --------------------
# download current Richmond stock: clover_item > Inventory > Items > Export
library(dplyr)
library(openxlsx)
adjust_inventory <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "order.*_receive.csv", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% `row.names<-`(toupper(.[, "ItemNumber"])) 
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- readWorkbook(clover, "Items") %>% mutate(SKU = toupper(SKU), Quantity = ifelse(SKU %in% rownames(adjust_inventory), Quantity + adjust_inventory[SKU, "Qty"], Quantity))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item), gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
# upload to Clover > Inventory

# -------- Batch update Clover inventory: at request -------------
clover_update <- read.csv("../Clover/Order04182024.csv", as.is = T) %>% mutate(Qty = 0 - Qty) %>% `row.names<-`(toupper(.[, "ItemNumber"])) 
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% mutate(Quantity = ifelse(Name %in% clover_update$ItemNumber, Quantity + clover_update[Name, "Qty"], Quantity))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item), gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)

# -------- Analysis: monthly --------------------
new_season <- "24"   # New season contains (yr)
qty_offline <- 3     # Qty to move to offline sales
month <- format(Sys.Date(), "%m")
if(month %in% c("09", "10", "11", "12", "01", "02")){in_season <- "F"}else(in_season <- "S")
RawData <- list.files(path = "../FBArefill/Raw Data File/", pattern = "All Marketplace All SKU Categories", full.names = T)
mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T)[1], sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
mastersku_adjust <- mastersku %>% filter(!duplicated(Adjust.SKU)) %>% `row.names<-`(toupper(.[, "Adjust.SKU"])) 
xoro <- read.xlsx2(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".xlsx"), full.names = T), sheetIndex = 1) %>% filter(Store == "Warehouse - JJ") %>% 
  mutate(Item. = toupper(Item.), ATS = as.numeric(ATS), Seasons = ifelse(Item. %in% mastersku$MSKU, mastersku[Item., "Seasons.SKU"], mastersku_adjust[Item., "Seasons.SKU"]), SPU = gsub("(\\w+-\\w+)-.*", "\\1", Item.)) %>% `row.names<-`(.[, "Item."])
## Raw data summary
monR_last_yr <- read.csv(list.files(path = "../Analysis/", pattern = "MonthlyRatio2023_category_adjust_2024-01-25.csv", full.names = T), as.is = T) %>% `row.names<-`(.[, "Category"])
sheets <- grep("-WK", getSheetNames(RawData), value = T)
inventory <- data.frame()
for(sheet in sheets){
  inventory_i <- openxlsx::read.xlsx(RawData, sheet = sheet, startRow = 2)
  inventory <- rbind(inventory, inventory_i)
}
inventory[is.na(inventory)] <- 0
write.csv(inventory, file = paste0("../Analysis/Sales_Inventory_SKU_", Sys.Date(), ".csv"), row.names = F, na = "")
names <- c(Category = "Cat_SKU", Month01="Jan", Month02="Feb", Month03="Mar", Month04="Apr", Month05="May", Month06="June", Month07="July", Month08="Aug", Month09="Sep", Month10="Oct", Month11="Nov", Month12="Dec")
sheets <- grep("-SKU$", getSheetNames(RawData), value = T)
if(month == "01"){columns <- c(1:11, 26:37)}else{columns <- c(1:11, 13:(as.numeric(month)-1+12), (25+as.numeric(month)):37)}
sales_cat_last12month <- data.frame()
sales_SKU_last12month <- data.frame()
for(sheet in sheets){
  sales_last12month_i <- openxlsx::read.xlsx(RawData, sheet = sheet, cols = columns, startRow = 2) 
  if(sum(grepl("janandjul.com", sales_last12month_i[[1]]))){
    sales_cat_last12month_i <- sales_last12month_i %>% filter(.[[1]] == "All Marketplaces") %>% mutate_all(~ replace(., is.na(.), 0)) %>% rename_with(~ gsub("20", "Total.20", gsub("_x000D_.Total", "", .x)))
    sales_cat_last12month <- rbind(sales_cat_last12month, sales_cat_last12month_i)
    sales_SKU_last12month_i <- openxlsx::read.xlsx(RawData, sheet = sheet, cols = columns, startRow = 2) %>% 
      filter(!(Sales.Channel %in% c("Summary", "Sales Channel", "All Marketplaces"))) %>% mutate_all(~ replace(., is.na(.), 0)) %>% rename_with(~ gsub("20", "Total.20", gsub("_x000D_.Total", "", .x)))
    sales_SKU_last12month <- rbind(sales_SKU_last12month, sales_SKU_last12month_i)
  }
}
sales_SKU_last12month <- sales_SKU_last12month %>% rename(all_of(names)) %>% select(Sales.Channel, Adjust_MSKU, Month01:Month12) %>% mutate(across(Month01:Month12, ~as.numeric(.)))
write.csv(sales_SKU_last12month, file = paste0("../Analysis/Sales_SKU_last12month_", Sys.Date(), ".csv"), row.names = F)
if(month %in% c("01", "02")){
  monR <- monR_last_yr %>% select(Category, Month01:Month12) %>% `row.names<-`(toupper(.[, "Category"]))
}else{
  monR_last_yr <- monR_last_yr %>% rowwise() %>% mutate(T.new = sum(c_across(Month01:as.name(sprintf("Month%02d", (as.numeric(month)-1)))))) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "Category"]))
  sales_cat_last12month <- sales_cat_last12month %>% rename(all_of(names)) %>% mutate(across(Month01:Month12, ~as.numeric(.))) %>% rowwise() %>% mutate(T.new = sum(c_across(Month01:as.name(sprintf("Month%02d", (as.numeric(month)-1))))), MonR.new = monR_last_yr[Category, "T.new"]) %>% 
    filter(!is.na(MonR.new)) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "Category"]))
  monR_new <- sales_cat_last12month %>% mutate(across(Month01:as.name(sprintf("Month%02d", (as.numeric(month)-1))), ~ ./T.new*MonR.new))
  monR_new[monR_new$T.new == 0, ] <- monR_new[monR_new$T.new == 0, ] %>% rows_update(., monR_last_yr %>% select(Category, Month01:as.name(sprintf("Month%02d", (as.numeric(month)-1)))), by = "Category", unmatched = "ignore")
  monR <- merge(monR_new %>% select(Category, Month01:as.name(sprintf("Month%02d", (as.numeric(month)-1)))), monR_last_yr %>% select(Category, as.name(sprintf("Month%02d", as.numeric(month))):Month12), by = "Category") %>% `row.names<-`(toupper(.[, "Category"]))
}
write.csv(monR, file = paste0("../Analysis/MonthlyRatio_", Sys.Date(), ".csv"), row.names = F)
## SKUs to move offline
offline <- xoro %>% filter(!grepl(new_season, Seasons) & ATS <= qty_offline & ATS > 0) %>% arrange(Item.) %>% 
  mutate(StoreCode = "WH-JJ", ItemNumber = Item., Qty = ATS, LocationName = "BIN", UnitCost = "", ReasonCode = "RWT", Memo = "Richmond Transfer to Miranda", UploadRule = "D", AdjAccntName = "", TxnDate = "", ItemIdentifierCode = "", ImportError = "") %>% select(StoreCode, ItemNumber, Qty, LocationName, UnitCost, ReasonCode, Memo, UploadRule, AdjAccntName, TxnDate, ItemIdentifierCode, ImportError)
write.csv(offline, file = paste0("../Analysis/offline_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\
## SKUs to inactive
clover <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items") %>% `row.names<-`(toupper(.[, "Name"]))
inventory_s <- inventory %>% count(Adjust_MSKU, wt = Inv_Total_End.WK) %>% `row.names<-`(toupper(.[, "Adjust_MSKU"]))
qty0 <- mastersku %>% select(MSKU.Status, Seasons.SKU, MSKU) %>% mutate(MSKU = toupper(MSKU), Inv_clover = clover[MSKU, "Quantity"], Inv_Total_EndWK = inventory_s[MSKU, "n"]) %>% 
  filter(!grepl(new_season, Seasons.SKU) & MSKU.Status == "Active" & Inv_clover == 0 & Inv_Total_EndWK == 0)
write.csv(qty0, file = paste0("../Analysis/discontinued_qty0_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\
## Suggested deals for the website
discount_method <- openxlsx::read.xlsx("../Analysis/JJ_discount_method.xlsx", sheet = 1, startRow = 2, rowNames = T)
if(month == "01"){last3m <- c("Month11", "Month12", "Month01")}else if(month %in% c("02", "03")){last3m <- sprintf("Month%02d", c(1:(as.numeric(month)-1), (as.numeric(month)-3+12):12))}else{last3m <- sprintf("Month%02d", c((as.numeric(month)-3):(as.numeric(month)-1)))}
monR <- monR %>% mutate(T.last3m = rowSums(across(all_of(last3m)))) %>% `row.names<-`(.[, "Category"])
inventory_SPU <- inventory %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU)) %>% group_by(SPU) %>% summarise(qty_SPU = sum(Inv_Total_End.WK, na.rm = T), qty_WH = sum(Inv_WH.JJ, na.rm = T), qty_AMZ = qty_SPU - qty_WH) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SKU_last12month <- sales_SKU_last12month %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU), T.last3m = rowSums(across(all_of(last3m))), T.yr = rowSums(across(Month01:Month12)))
sales_SPU_last3month <- sales_SKU_last12month %>% group_by(SPU) %>% summarise(T.last3m = sum(T.last3m)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SPU_last3month_JJ <- sales_SKU_last12month %>% filter(Sales.Channel == "janandjul.com") %>% group_by(SPU) %>% summarise(T.last3m = sum(T.last3m)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SPU_last12month_AMZ <- sales_SKU_last12month %>% filter(grepl("Amazon", Sales.Channel)) %>% group_by(SPU) %>% summarise(T.last12m = sum(T.yr)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sizes_all <- xoro %>% mutate(SPU_size = gsub("(\\w+-\\w+-\\w+).*", "\\1", Item.)) %>% filter(!duplicated(SPU_size)) %>% count(SPU) %>% `row.names<-`(.[, "SPU"])
size_limited <- xoro %>% filter(ATS > qty_offline) %>% count(SPU) %>% mutate(all = sizes_all[SPU, "n"], percent = ifelse(all < n, 0, (all - n)/all)) %>% `row.names<-`(.[, "SPU"])
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% filter(Published == 1, !is.na(Regular.price)) %>% filter(!duplicated(SKU), SKU != "") %>% 
  mutate(SKU = toupper(SKU), Qty = xoro[SKU, "ATS"], Seasons = ifelse(SKU %in% mastersku$MSKU, mastersku[SKU, "Seasons.SKU"], mastersku_adjust[SKU, "Seasons.SKU"]), discount = ifelse(is.na(Sale.price), 0, (Regular.price - Sale.price)/Regular.price), cat = mastersku[SKU, "Category.SKU"], SPU = gsub("(\\w+-\\w+)-.*", "\\1", SKU)) %>% 
  select(ID, SKU, Name, Seasons, cat, SPU, Regular.price, discount, Qty) %>% `row.names<-`(.[, "SKU"])
woo_deals <- woo %>% filter(!grepl(new_season, Seasons), Qty > qty_offline, Regular.price > 0) %>% mutate(Qty_SPU = inventory_SPU[SPU, "qty_SPU"], Qty_WH = inventory_SPU[SPU, "qty_WH"], Qty_AMZ = inventory_SPU[SPU, "qty_AMZ"], MonR_last3m = monR[cat, "T.last3m"], Sales_SPU_last3m = sales_SPU_last3month[SPU, "T.last3m"], Sales_SPU_last3m_JJ = sales_SPU_last3month_JJ[SPU, "T.last3m"], Sales_SPU_1yr = Sales_SPU_last3m/MonR_last3m, Sales_SPU_1yr_JJ = Sales_SPU_last3m_JJ/MonR_last3m, Sales_SPU_1yr_AMZ = sales_SPU_last12month_AMZ[SPU, "T.last12m"]) %>% 
  filter(!is.na(Qty_SPU)) %>% mutate(time_yrs = ifelse(Qty_AMZ >= Sales_SPU_1yr_AMZ, Qty_WH/Sales_SPU_1yr_JJ, Qty_SPU/Sales_SPU_1yr), time_yrs = paste0("Sold out ", as.character(as.integer(ifelse(time_yrs > 3, 3, time_yrs)) + 1), " yr"), size_percent_missing = gsub("Sizes.missing.0%", "Sizes.missing.<.50%", paste0("Sizes.missing.", as.character(as.integer(ifelse(size_limited[SPU, "percent"] < 0.5, 0, size_limited[SPU, "percent"])*10)*10), "%"))) %>%
  rowwise() %>% mutate(Suggest_discount = as.numeric(gsub("% Off", "", discount_method[time_yrs, size_percent_missing]))/100, Suggest_price = round((1 - Suggest_discount) * Regular.price, digits = 2)) %>% arrange(-Suggest_discount, SKU) %>% filter(Suggest_discount > 0, Suggest_discount > discount) 
write.csv(woo_deals, file = paste0("../Analysis/Deals_all_", Sys.Date(), ".csv"), row.names = F)
# Email Will
if(month %in% c("02", "03", "04", "05")){woo_deals <- woo_deals %>% filter(!grepl("S", Seasons))}
if(month %in% c("08", "09", "10", "11")){woo_deals <- woo_deals %>% filter(!grepl("F", Seasons))}
write.csv(woo_deals, file = paste0("../Analysis/Deals_", Sys.Date(), ".csv"), row.names = F)
# Email Joren, Kamer
## Low inventory to adjust ads
qty_refill <- 12
inventory <- inventory %>% mutate(Inv_WH_AMZ.CA.US = Inv_WH.JJ + Inv_AMZ.USA + Inv_AMZ.CA) %>% `row.names<-`(.[, "Adjust_MSKU"])
Sales_Inv_ThisMonth <- sales_SKU_last12month %>% filter(Sales.Channel %in% c("Amazon.com", "Amazon.ca", "janandjul.com")) %>% group_by(Adjust_MSKU) %>% summarise(Sales.last3m = sum(T.last3m)) %>% filter(Adjust_MSKU %in% inventory$Adjust_MSKU) %>%
  mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU), cat = mastersku[Adjust_MSKU, "Category.SKU"], Seasons = mastersku[Adjust_MSKU, "Seasons.SKU"], Status = mastersku[Adjust_MSKU, "MSKU.Status"], AMZ_flag = mastersku[Adjust_MSKU, "Pause.Plan.FBA_x000D_.CA/US/MX/AU/UK/DE"], monR_last3m = monR[cat, "T.last3m"], monR_this = monR[cat, paste0("Month", month)], Sales.this = Sales.last3m/monR_last3m*monR_this, Inv_WH_AMZ.CA.US = inventory[Adjust_MSKU, "Inv_WH_AMZ.CA.US"], Qty_left = ifelse(Inv_WH_AMZ.CA.US < Sales.this, 0, as.integer(Inv_WH_AMZ.CA.US - Sales.this)), Enough_Inv = (Qty_left >= qty_refill))
Sales_Inv_ThisMonth_SPU <- Sales_Inv_ThisMonth %>% group_by(SPU) %>% summarise(Enough_Inv = mean(Enough_Inv)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
Sales_Inv_ThisMonth <- Sales_Inv_ThisMonth %>% mutate(Percent_Sizes_SPU_Enough_Inv = paste0(as.character(as.integer(Sales_Inv_ThisMonth_SPU[SPU, "Enough_Inv"]*100)), "%")) %>% filter(grepl(in_season, Seasons), Status == "Active", !Enough_Inv) %>% arrange(Qty_left, Adjust_MSKU)
write.csv(Sales_Inv_ThisMonth, file = paste0("../Analysis/Sales_Inv_Prediction_", Sys.Date(), ".csv"), row.names = F)
# Email Adrienne, Joren and Kamer 
## Low inventory to add PO
if(month == "12"){next4m <- c("Month01", "Month02", "Month03", "Month04")}else if(month %in% c("10", "11")){next4m <- sprintf("Month%02d", c(1:(as.numeric(month)-12+3), as.numeric(month):12))}else{next4m <- sprintf("Month%02d", c(as.numeric(month):(as.numeric(month)+3)))}
PO <- rbind(openxlsx::read.xlsx("../PO/2024FW China to Global Shipment Tracking.xlsx", sheet = "2024FW", startRow = 4, fillMergedCells = T, cols = c(7, 15)), openxlsx::read.xlsx("../PO/2024SS China to Global Shipment Tracking.xlsx", sheet = "2024SS", startRow = 4, fillMergedCells = T, cols = c(7, 15))) %>% 
  filter(!is.na(SKU)) %>% group_by(SKU) %>% summarise(Remain = sum(as.numeric(Remain..in.Pcs))) %>% as.data.frame() %>% `row.names<-`(.[, "SKU"])
monR <- monR %>% mutate(T.next4m = rowSums(across(all_of(next4m)))) %>% `row.names<-`(.[, "Category"])
Sales_Inv_Next4m <- sales_SKU_last12month %>% filter(Sales.Channel %in% c("Amazon.com", "Amazon.ca", "janandjul.com")) %>% group_by(Adjust_MSKU) %>% summarise(Sales.last3m = sum(T.last3m)) %>% filter(Adjust_MSKU %in% inventory$Adjust_MSKU) %>%
  mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU), cat = mastersku[Adjust_MSKU, "Category.SKU"], Seasons = mastersku[Adjust_MSKU, "Seasons.SKU"], Status = mastersku[Adjust_MSKU, "MSKU.Status"], AMZ_flag = mastersku[Adjust_MSKU, "Pause.Plan.FBA_x000D_.CA/US/MX/AU/UK/DE"], monR_last3m = monR[cat, "T.last3m"], monR_next4m = monR[cat, "T.next4m"], Sales.next4m = as.integer(Sales.last3m/monR_last3m*monR_next4m), Inv_WH_AMZ.CA.US = inventory[Adjust_MSKU, "Inv_WH_AMZ.CA.US"], PO_remain = ifelse(Adjust_MSKU %in% PO$SKU, PO[Adjust_MSKU, "Remain"], 0), diff = Sales.next4m - Inv_WH_AMZ.CA.US - PO_remain, Enough_Inv = (Inv_WH_AMZ.CA.US + PO_remain >= Sales.next4m)) %>%
  filter(!Enough_Inv, grepl("24", Seasons))
write.csv(Sales_Inv_Next4m, file = paste0("../Analysis/Sales_Inv_Prediction_Next4m_", Sys.Date(), ".csv"), row.names = F)
# Discuss with Mei, Florence, Cindy and Matt

# -------- Prepare to generate barcode image: at request ------------
library(dplyr)
library(openxlsx)
library(stringr)
season <- "2024FW"
startRow <- 9
mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"])) 
qty0 <- data.frame()
skus <- c()
for(i in 292:342){
  POn <- paste0("P", i)
  if(!sum(grepl(POn, list.dirs(paste0("../PO/order/", season, "/"), recursive = F)))) next
  print(POn)
  file <- list.files(path = paste0("../PO/order/", season, "/"), pattern = paste0(POn, ".*.xlsx"), full.names = T, recursive = T)
  for(f in 1:length(file)){
    barcode <- read.xlsx(file[f], sheet = 1, startRow = startRow) %>% select(SKU, Design.Version, TOTAL.ORDER) %>% 
      mutate(SKU = str_trim(SKU), Category = mastersku[SKU, "Category.SKU"], Print.English = mastersku[SKU, "Print.Name"], Print.Chinese = mastersku[SKU, "Print.Chinese"], Size = mastersku[SKU, "Size"], UPC.Active = gsub("/.*", "", mastersku[SKU, "UPC.Active"]), Image = "") %>%
      filter(SKU != "", !is.na(SKU))
    skus <- c(skus, barcode[barcode$TOTAL.ORDER != 0, "SKU"])
    qty0 <- rbind(qty0, barcode %>% filter(grepl("NEW", Design.Version), TOTAL.ORDER == 0, !is.na(UPC.Active)))
    barcode <- barcode %>% filter(TOTAL.ORDER != 0) %>% select(SKU, Print.English, Print.Chinese,	Category, Size, UPC.Active, Design.Version, Image)
    barcode_split <- split(barcode, f = barcode$Category )
    for(j in 1:length(barcode_split)){
      barcode_j <- barcode_split[[j]]
      cat <- barcode_j[1, "Category"]
      dir <- grep(cat, list.dirs("../../TWK Product Labels/", recursive = F), value = T)
      if(length(dir) == 0){
        dir.create(paste0("../../TWK Product Labels/", cat, "/", season, "/"), recursive = T)
        write.xlsx(barcode, file = paste0(paste0("../../TWK Product Labels/", cat, "/", season, "/"), POn, "_", cat, "-Barcode Labels.xlsx"))
      }
      else{
        if(!sum(grepl("24F", list.dirs(dir, recursive = F)))){
          dir.create(paste0(dir, "/", season, "/"))
          write.xlsx(barcode, file = paste0(dir, "/", season, "/", POn, "_", cat, "-Barcode Labels.xlsx"))
        }
        else{
          write.xlsx(barcode, file = paste0(grep("24F", list.dirs(dir, recursive = F), value = T), "/", POn, "_", cat, "-Barcode Labels.xlsx"))
        }
      }
    }
  }
}
qty0 <- qty0 %>% filter(!(SKU %in% skus))
write.xlsx(qty0, file = "../../TWK Product Labels/SKU_new_qty0_UPC.xlsx", rowNames = F)

# -------- Sync master file barcode with clover: at request -------------
# master SKU file: OneDrive > TWK 2020 share
# download clover inventory: clover_item > Inventory > Items > Export.
library(dplyr)
library(openxlsx)
library(tidyr)
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price), cat = toupper(gsub("-.*", "", SKU))) %>% `row.names<-`(toupper(.[, "SKU"])) 
price <- woo %>% group_by(cat) %>% summarise(Price = max(Sale.price)) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "cat"])) 
mastersku <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../FBArefill/Raw Data File/", pattern = "1-MasterSKU-All-Product-", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% filter(!is.na(Name), !duplicated(Name)) %>% `row.names<-`(toupper(.[, "Name"])) 
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

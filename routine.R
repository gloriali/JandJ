# Routine tasks
library(dplyr)
library(stringr)
library(openxlsx)
library(xlsx)

# input: woo > Products > All Products > Export > all
# input: xoro > Item Inventory Snapshot > Export all
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price))
rownames(woo) <- woo$SKU
xoro <- read.xlsx2(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".xlsx"), full.names = T), sheetIndex = 1) %>% filter(Store == "Warehouse - JJ") %>% mutate(Item. = toupper(Item.), ATS = as.numeric(ATS))
rownames(xoro) <- xoro$Item.

# -------- Sync XHS price to woo; inventory to xoro: Mon Wed Fri ----------------
# input: AllValue > Products > Export > All products
products_description <- read.xlsx2("../XHS/products_description.xlsx", sheetIndex = 1)
rownames(products_description) <- products_description$SPU
products_XHS <- read.xlsx2(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), sheetIndex = 1)
products_upload <- products_XHS %>% mutate(Inventory = ifelse(xoro[SKU, "ATS"] < 20, 0, xoro[SKU, "ATS"]), Price = woo[SKU, "Sale.price"], Compare.At.Price = woo[SKU, "Regular.price"], Tags = ifelse(Price == Compare.At.Price, "正价", "特价"))
products_upload <- products_upload %>% mutate(Product.Name = products_description[toupper(SPU), "Product.Name"], SEO.Product.Name = Product.Name, Description = products_description[toupper(SPU), "Description"], Mobile.Description = products_description[toupper(SPU), "Description"], SEO.Description = products_description[toupper(SPU), "Description"], Describe = products_description[toupper(SPU), "Description"])
colnames(products_upload) <- gsub("\\.", " ", colnames(products_upload))
openxlsx::write.xlsx(products_upload, file = paste0("../XHS/products_upload-", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"))
# upload to AllValue > Products

# -------- Sync Square price to woo; inventory to xoro: weekly -------------
square <- data.frame(Token = "", ItemName = woo$SKU, VariationName = "Regular", SKU = woo$SKU, Description = "", Category = xoro[woo$SKU, "Item.Category"], Price = woo$Sale.price, PriceRichmond = "", PriceSurrey = "", NewQuantityRichmond = xoro[woo$SKU, "ATS"], NewQuantitySurrey = xoro[woo$SKU, "ATS"], EnabledRichmond = "Y", CurrentQuantityRichmond = "", CurrentQuantitySurrey = "", StockAlertEnabledRichmond = "N", StockAlertCountRichmond = 0, EnabledSurrey = "Y", StockAlertEnabledSurrey = "N", StockAlertCountSurrey = 0, SEOTitle = "",	SEODescription = "",	Permalink = "", SquareOnlineItemVisibility = "UNAVAILABLE",	ShippingEnabled = "",	SelfserveOrderingEnabled = "",	DeliveryEnabled = "",	PickupEnabled = "", Sellable = "Y",	Stockable = "Y",	SkipDetailScreeninPOS = "",	OptionName1 = "",	OptionValue1 = "", TaxPST = ifelse(woo$Tax.class == "full", "Y", "N")) %>%
  mutate(NewQuantityRichmond = ifelse(is.na(NewQuantityRichmond), 0, NewQuantityRichmond), NewQuantitySurrey = ifelse(is.na(NewQuantitySurrey), 0, NewQuantitySurrey))
write.table(square, file = paste0("../Square/square-upload-", Sys.Date(), ".csv"), sep = ",", row.names = F, col.names = c(colnames(square)[1:ncol(square)-1], "Tax - PST (7%)"), na = "")
# upload to Square > Items

# -------- Sync Clover to woo: weekly -------------
# input Clover > Inventory > Items > Export.
price <- woo %>% mutate(cat = gsub("-.*", "", SKU)) %>% group_by(cat) %>% summarise(Price = max(Sale.price)) %>% as.data.frame()
rownames(price) <- price$cat
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
request <- c("IHT", "SKG", "WJA", "WGS") # categories to restock
n <- 1       # Qty per SKU to stock at Richmond
n_xoro <- 10 # min Qty in stock at Surrey to request
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- readWorkbook(clover, "Items") %>% mutate(cat = gsub("-.*", "", Name)) %>% filter(!duplicated(Name), !is.na(Name))
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
f <- "../FBArefill/Historic sales and inv. data for all cats v31 (20240104).xlsx"
f1 <- "../xoro/Item Inventory Snapshot_12012023.csv"
mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
xoro <- read.csv(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"), full.names = T), as.is = T) %>% filter(Store == "Warehouse - JJ") %>% 
  mutate(seasons = mastersku[Item., "Seasons"], cat_print = gsub("(\\w+-\\w+)-.*", "\\1", Item.))
rownames(xoro) <- xoro$Item.
## move offline 
offline <- xoro %>% filter(!grepl(new_season, seasons) & ATS <= qty_offline & ATS > 0) %>% arrange(Item.)
write.csv(offline, file = paste0("../SKUmanagement/offline_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\
## suggest to website deals page 
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price), discount = (Regular.price - Sale.price)/Regular.price)
rownames(woo) <- woo$SKU
xoro_lastmonth <- read.csv(f1, as.is = T) %>% filter(Store == "Warehouse - JJ") 
rownames(xoro_lastmonth) <- xoro_lastmonth$Item.
JJ_month_ratio <- read.xlsx(f, sheet = "MonSaleR", startRow = 2, cols = c(1, 26:37)) %>% filter(!is.na(Category)) 
JJ_month_ratio <- JJ_month_ratio %>% mutate(monthR = as.integer(format(Sys.Date(), "%d"))/31*JJ_month_ratio[, gsub("0", "", format(Sys.Date(), "%m"))] + (31 - as.integer(format(Sys.Date(), "%d")))/31*JJ_month_ratio[, as.character(as.integer(format(Sys.Date(), "%m"))-1)])
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
clover <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items") %>% filter(!duplicated(Name), !is.na(Name))
rownames(clover) <- clover$SKU
names <- grep("-WK", getSheetNames(f), value = T)
inventory <- data.frame()
for(name in names){
  cat_inventory <- openxlsx::read.xlsx(f, sheet = name, startRow = 2)
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
library(stringr)
season <- "2024FW"
startRow <- 9
mastersku <- read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
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
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price), cat = toupper(gsub("-.*", "", SKU)))
rownames(woo) <- woo$SKU
price <- woo %>% group_by(cat) %>% summarise(Price = max(Sale.price)) %>% as.data.frame()
rownames(price) <- price$cat
mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% filter(!is.na(Name), !duplicated(Name))
rownames(clover_item) <- clover_item$Name
category_exclude <- c("FVMU", "MFJM", "MSKB", "MSXP", "MWBF", "MWBS", "MWJA", "MWMT", "MWPF", "MWPS", "MWSF", "MWSS", "PJWU","PLAU", "PTAU", "STORE", "TPLU", "TPSU", "TSWU", "TTLU", "TTSU", "WJAU", "WPSU", "XBMU")
clover_item_upload <- data.frame(Clover.ID = "", Name = mastersku[mastersku$MSKU.Status == "Active", "MSKU"]) %>% filter(!is.na(Name)) %>% 
  mutate(cat = mastersku[Name, "Category.SKU"], Alternate.Name = woo[Name, "Name"], Price = ifelse(Name %in% woo$SKU, woo[Name, "Sale.price"], price[cat, "Price"]), Price.Type = ifelse(is.na(Price), "Variable", "Fixed"), Price.Unit = NA, Tax.Rates = ifelse(woo[Name, "Tax.class"] == "full", "GST+PST", "GST"), Tax.Rates = ifelse(is.na(Tax.Rates), "GST", Tax.Rates), Cost = 0, Product.Code = gsub("/.*", "", mastersku[Name, "UPC.Active"]), SKU = Name, Modifier.Groups = NA, Quantity = ifelse(Name %in% clover_item$Name, clover_item[Name, "Quantity"], 0), Printer.Labels = NA, Hidden = "No", Non.revenue.item = "No") %>%
  filter(!(cat %in% category_exclude)) %>% arrange(cat, Name)
clover_cat_upload <- clover_item_upload %>% group_by(cat) %>% mutate(id = row_number(cat)) %>%  select(cat, id, Name) %>% spread(id, Name, fill = "") %>% t() 
clover_cat_upload <- cbind(data.frame(V0 = c("Category Name", "Items in Category", rep("", times = nrow(clover_cat_upload) - 2))), clover_cat_upload)
clover_item_upload <- clover_item_upload %>% select(-cat)
clover_item_upload <- clover_item_upload %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item)+1000, gridExpand = T)
writeData(clover, sheet = "Items", clover_item_upload)
deleteData(clover, sheet = "Categories", cols = 1:1000, rows = 1:1000, gridExpand = T)
writeData(clover, sheet = "Categories", clover_cat_upload, colNames = F)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
# upload to Clover > Inventory

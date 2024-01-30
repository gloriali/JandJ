# SKUs analysis and management
library(dplyr)
library(openxlsx)
library(scales)
library(ggplot2)
library(xlsx)
new_season <- "24"   # New season contains
in_season <- "F"     # "F": Sept - Feb; "S": Mar - Aug
qty_offline <- 3     # Qty to move to offline sales
size_percent <- 0.5  # % of sizes not available 
RawData <- list.files(path = "../FBArefill/Raw Data File/", pattern = "Historic sales and inv. data for all cats", full.names = T)

mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T)[1], sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
mastersku_adjust <- mastersku %>% filter(!duplicated(Adjust.SKU)) %>% `row.names<-`(toupper(.[, "Adjust.SKU"])) 
xoro <- read.xlsx2(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".xlsx"), full.names = T), sheetIndex = 1) %>% filter(Store == "Warehouse - JJ") %>% 
  mutate(Item. = toupper(Item.), ATS = as.numeric(ATS), Seasons = ifelse(Item. %in% mastersku$MSKU, mastersku[Item., "Seasons.SKU"], mastersku_adjust[Item., "Seasons.SKU"]), SPU = gsub("(\\w+-\\w+)-.*", "\\1", Item.)) %>% `row.names<-`(.[, "Item."])

# ------------- move offline ---------------------------
offline <- xoro %>% filter(!grepl(new_season, Seasons) & ATS <= qty_offline & ATS > 0) %>% arrange(Item.) %>% 
  mutate(StoreCode = "WH-JJ", ItemNumber = Item., Qty = ATS, LocationName = "BIN", UnitCost = "", ReasonCode = "RWT", Memo = "Richmond Transfer to Miranda", UploadRule = "D", AdjAccntName = "", TxnDate = "", ItemIdentifierCode = "", ImportError = "") %>% 
  select(StoreCode, ItemNumber, Qty, LocationName, UnitCost, ReasonCode, Memo, UploadRule, AdjAccntName, TxnDate, ItemIdentifierCode, ImportError)
write.csv(offline, file = paste0("../Analysis/offline_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\

# ------------- to inactivate --------------------------
clover <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items") %>% `row.names<-`(toupper(.[, "Name"]))
sheets <- grep("-WK", getSheetNames(RawData), value = T)
inventory <- data.frame()
for(sheet in sheets){
  inventory_i <- openxlsx::read.xlsx(RawData, sheet = sheet, startRow = 2)
  inventory <- rbind(inventory, inventory_i)
}
write.csv(inventory, file = paste0("../Analysis/Sales_Inventory_SKU_", Sys.Date(), ".csv"), row.names = F, na = "")
inventory_s <- inventory %>% count(Adjust_MSKU, wt = Inv_Total_End.WK) %>% `row.names<-`(toupper(.[, "Adjust_MSKU"]))
qty0 <- mastersku %>% select(MSKU.Status, Seasons.SKU, MSKU) %>% mutate(MSKU = toupper(MSKU), Inv_clover = clover[MSKU, "Quantity"], Inv_Total_EndWK = inventory_s[MSKU, "n"]) %>% 
  filter(!grepl(new_season, Seasons.SKU) & MSKU.Status == "Active" & Inv_clover == 0 & Inv_Total_EndWK == 0)
write.csv(qty0, file = paste0("../Analysis/discontinued_qty0_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\

# ------------- move to website deals page -------------
monR_last_yr <- read.csv(list.files(path = "../Analysis/", pattern = "MonthlyRatio2023_category_adjust_2024-01-25.csv", full.names = T), as.is = T) %>% `row.names<-`(.[, "Category"])
month <- format(Sys.Date(), "%m")
sheets <- grep("-SKU$", getSheetNames(RawData), value = T)
if(month == "01"){columns <- c(1:11, 26:37)}else{columns <- c(1:11, 13:(as.numeric(month)-1+12), (25+as.numeric(month)):37)}
sales_cat_last12month <- data.frame()
for(sheet in sheets){
  sales_cat_last12month_i <- openxlsx::read.xlsx(RawData, sheet = sheet, cols = columns, startRow = 2) 
  if(sum(grepl("janandjul.com", sales_cat_last12month_i[[1]]))){
    sales_cat_last12month_i <- sales_cat_last12month_i %>% filter(.[[1]] == "All Marketplaces") %>% mutate_all(~ replace(., is.na(.), 0)) %>% rename_with(~ gsub("20", "Total.20", gsub("_x000D_.Total", "", .x)))
    sales_cat_last12month <- rbind(sales_cat_last12month, sales_cat_last12month_i)
  }
}
sales_cat_last12month <- sales_cat_last12month %>% filter(Total.2023 > 0)
if(month == "01"){
  monR <- monR_last_yr
}else{
  monR_last_yr <- monR_last_yr %>% rowwise() %>% mutate(T.new = sum(.[, sprintf("Month%02d", c(1:(as.numeric(month)-1)))]))
}
# email Joren and Kamer

# ------------- summarize raw sales data ------------- 
sheets <- getSheetNames(RawData)
sheets_cat <- sheets[!grepl("-", sheets)&!grepl("Sale", sheets)]
sales_cat <- data.frame()
monR_cat <- data.frame()
for(sheet in sheets_cat){
  sales_cat_i <- openxlsx::read.xlsx(RawData, sheet = sheet, cols = c(1:19), startRow = 2) %>% 
    filter(Sales.Channel == "All Channels", Year == "2023")
  if(nrow(sales_cat_i) == 0){
    sales_cat_i <- openxlsx::read.xlsx(RawData, sheet = sheet, cols = c(1:19), startRow = 2) %>% 
      filter(Sales.Channel == "janandjul.com", Year == "2023")
  }
  sales_cat <- rbind(sales_cat, sales_cat_i)
  monR_cat_i <- openxlsx::read.xlsx(RawData, sheet = sheet, cols = c(1:6, 20:32), startRow = 2) %>% 
    filter(Sales.Channel == "All Channels", Year == "2023")
  if(nrow(monR_cat_i) == 0){
    monR_cat_i <- openxlsx::read.xlsx(RawData, sheet = sheet, cols = c(1:6, 20:32), startRow = 2) %>% 
      filter(Sales.Channel == "janandjul.com", Year == "2023")
  }
  monR_cat <- rbind(monR_cat, monR_cat_i)
}
write.csv(sales_cat, file = paste0("../Analysis/SalesAll2023_category_", Sys.Date(), ".csv"), row.names = F, na = "")
write.csv(monR_cat, file = paste0("../Analysis/MonthlyRatio2023_category_", Sys.Date(), ".csv"), row.names = F, na = "")
sheets_SKU <- sheets[grepl("-SKU$", sheets)]
sales_SKU <- data.frame()
for(sheet in sheets_SKU){
  sales_SKU_i <- openxlsx::read.xlsx(RawData, sheet = sheet, cols = c(1:11, 26:37), startRow = 2) %>% 
    filter(!(Sales.Channel %in% c("Summary", "Sales Channel", "All Marketplaces"))) %>% mutate_all(~ replace(., is.na(.), 0)) %>% rename_with(~ gsub("20", "Total.20", gsub("_x000D_.Total", "", .x)))
  sales_SKU <- rbind(sales_SKU, sales_SKU_i)
}
sales_SKU <- sales_SKU %>% filter(Total.2023 > 0) %>% filter(!duplicated(Adjust_MSKU))
write.csv(sales_SKU, file = paste0("../Analysis/Sales2023_SKU_", Sys.Date(), ".csv"), row.names = F, na = "")

# ------------- FW clearance -------------
monR_cat <- read.csv(list.files(path = "../Analysis/", pattern = "MonthlyRatio2023_category_adjust_2024-01-25.csv", full.names = T), as.is = T) %>% mutate(FW_peak = Sep + Oct + Nov + Dec, Jan2Jun = Jan + Feb + Mar + Apr + May + June) %>% `row.names<-`(.[, "Category"])
inventory <- read.csv(paste0("../Analysis/Sales_Inventory_SKU_", Sys.Date(), ".csv"), as.is = T) %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU)) %>% `row.names<-`(.[, "Adjust_MSKU"])
inventory_SPU <- inventory %>% group_by(SPU) %>% summarise(qty_SPU = sum(Inv_Total_End.WK)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SKU <- read.csv(list.files(path = "../Analysis/", pattern = "SalesJJ2023_SKU_.*.csv", full.names = T), as.is = T) %>% mutate(FW_peak = Sep + Oct + Nov + Dec, Jan2Jun = Jan + Feb + Mar + Apr + May + June, SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU)) %>% `row.names<-`(.[, "Adjust_MSKU"])
sales_peak_SPU <- sales_SKU %>% group_by(SPU) %>% summarise(sales_peak_SPU = sum(FW_peak)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_Jan2Jun_SPU <- sales_SKU %>% group_by(SPU) %>% summarise(sales_Jan2Jun_SPU = sum(Jan2Jun)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_yr_SPU <- sales_SKU %>% group_by(SPU) %>% summarise(sales_yr_SPU = sum(Total.2023)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T)[1], sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
mastersku_adjust <- mastersku %>% filter(!duplicated(Adjust.SKU)) %>% `row.names<-`(toupper(.[, "Adjust.SKU"])) 
xoro <- read.xlsx2(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".xlsx"), full.names = T), sheetIndex = 1) %>% filter(Store == "Warehouse - JJ") %>% 
  mutate(Item. = toupper(Item.), ATS = as.numeric(ATS), Seasons = ifelse(Item. %in% mastersku$MSKU, mastersku[Item., "Seasons.SKU"], mastersku_adjust[Item., "Seasons.SKU"]), SPU = gsub("(\\w+-\\w+)-.*", "\\1", Item.)) %>% `row.names<-`(.[, "Item."])
sizes_all <- xoro %>% mutate(SPU_size = gsub("(\\w+-\\w+-\\w+).*", "\\1", Item.)) %>% filter(!duplicated(SPU_size)) %>% count(SPU) %>% `row.names<-`(.[, "SPU"])
size_limited <- xoro %>% filter(ATS > qty_offline) %>% count(SPU) %>% mutate(all = sizes_all[SPU, "n"], percent = (all - n)/all) %>% `row.names<-`(.[, "SPU"])

woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% filter(!is.na(Regular.price), !duplicated(SKU), SKU != "") %>% 
  mutate(SKU = toupper(SKU), Qty = ifelse(SKU %in% xoro$Item., xoro[SKU, "ATS"], as.numeric(Stock)), Seasons = ifelse(SKU %in% mastersku$MSKU, mastersku[SKU, "Seasons.SKU"], mastersku_adjust[SKU, "Seasons.SKU"]), discount = ifelse(is.na(Sale.price), 0, (Regular.price - Sale.price)/Regular.price), cat = mastersku[SKU, "Category.SKU"], SPU = gsub("(\\w+-\\w+)-.*", "\\1", SKU)) %>% 
  select(ID, SKU, Name, Seasons, cat, SPU, Regular.price, discount, Qty) %>% `row.names<-`(.[, "SKU"])
woo_SPU_Qty <- woo %>% group_by(SPU) %>% summarise(SPU_Qty = sum(Qty)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
woo_FW <- woo %>% filter(grepl("F", Seasons), Qty > qty_offline, Regular.price > 0) %>% mutate(Qty_SPU = inventory_SPU[SPU, "qty_SPU"], Growth_cat = monR_cat[cat, "Year.Year.Growth"], MonR_peak = ifelse(!is.na(Growth_cat), monR_cat[cat, "FW_peak"], NA), MonR_Jan2Jun = ifelse(!is.na(Growth_cat), monR_cat[cat, "Jan2Jun"], NA), Sales_peak_SPU = sales_peak_SPU[SPU, "sales_peak_SPU"]) %>% 
  mutate(MonR_peak = ifelse(is.na(MonR_peak), median(MonR_peak, na.rm = T), MonR_peak), MonR_Jan2Jun = ifelse(is.na(MonR_Jan2Jun), median(MonR_Jan2Jun, na.rm = T), MonR_Jan2Jun), Sales_yr_SPU = pmax(Sales_peak_SPU/MonR_peak, sales_yr_SPU[SPU, "sales_yr_SPU"]), Sales_Jan2Jun_SPU = Sales_yr_SPU * MonR_Jan2Jun)
write.csv(woo_FW %>% filter(is.na(Sales_yr_SPU)), file = paste0("../Analysis/Clearance_2023FW_NoSalesData_", Sys.Date(), ".csv"), row.names = F, na = "")
woo_FW_24F <- woo_FW %>% filter(!is.na(Sales_yr_SPU), grepl(new_season, Seasons)) %>% mutate(Qty_predict_Jun = ifelse(Qty_SPU > Sales_Jan2Jun_SPU, as.integer(Qty_SPU - Sales_Jan2Jun_SPU), 0), times_QtyJun_Sales = round(Qty_predict_Jun/Sales_Jan2Jun_SPU, digits = 1)) %>% 
  mutate(Suggest_discount = ifelse(times_QtyJun_Sales <= 0.5 | Qty_predict_Jun <= 20, 0.05, NA), Suggest_discount = pmax(Suggest_discount, discount), Suggest_sale_price = round((1 - Suggest_discount) * Regular.price, digits = 2))
woo_FW_24F[is.na(woo_FW_24F$Suggest_discount), ] <- woo_FW_24F[is.na(woo_FW_24F$Suggest_discount), ] %>% mutate(Suggest_discount = as.numeric(cut_number(times_QtyJun_Sales, n = 4)) * 0.05 + 0.05, Suggest_discount = pmax(Suggest_discount, discount), Suggest_sale_price = round((1 - Suggest_discount) * Regular.price, digits = 2))
#plot(density(woo_FW_24F$Qty_predict_Jun), xlim = c(0,200))
#plot(density(woo_FW_24F$times_QtyJun_Sales), xlim = c(0,20))
write.csv(woo_FW_24F, file = paste0("../Analysis/Clearance_2023FW_CarryOver_", Sys.Date(), ".csv"), row.names = F, na = "")
woo_FW_discontinued <- woo_FW %>% filter(!is.na(Sales_yr_SPU), !grepl(new_season, Seasons)) %>% mutate(times_yr = round(Qty_SPU/Sales_yr_SPU, digits = 1), size_percent_missing = size_limited[SPU, "percent"]) %>% 
  mutate(Suggest_discount = ifelse(times_yr <= 1 & size_percent_missing < size_percent, 0.05, NA), Suggest_discount = pmax(Suggest_discount, discount), Suggest_sale_price = round((1 - Suggest_discount) * Regular.price, digits = 2))
woo_FW_discontinued[is.na(woo_FW_discontinued$Suggest_discount) & grepl("23", woo_FW_discontinued$Seasons) & woo_FW_discontinued$size_percent_missing < size_percent, ] <- woo_FW_discontinued[is.na(woo_FW_discontinued$Suggest_discount) & grepl("23", woo_FW_discontinued$Seasons) & woo_FW_discontinued$size_percent_missing < size_percent, ] %>% mutate(Suggest_discount = as.numeric(cut_number(times_yr, n = 4)) * 0.05 + 0.05, Suggest_discount = pmax(Suggest_discount, discount), Suggest_sale_price = round((1 - Suggest_discount) * Regular.price, digits = 2))
woo_FW_discontinued[is.na(woo_FW_discontinued$Suggest_discount), ] <- woo_FW_discontinued[is.na(woo_FW_discontinued$Suggest_discount), ] %>% mutate(Suggest_discount = as.numeric(cut_number(times_yr, n = 5)) * 0.05 + 0.25, Suggest_discount = pmax(Suggest_discount, discount), Suggest_sale_price = round((1 - Suggest_discount) * Regular.price, digits = 2))
write.csv(woo_FW_discontinued, file = paste0("../Analysis/Clearance_2023FW_discontinued_", Sys.Date(), ".csv"), row.names = F, na = "")


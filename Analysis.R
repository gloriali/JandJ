# SKUs analysis and management
library(dplyr)
library(openxlsx)
library(scales)
library(ggplot2)
library(xlsx)
new_season <- "24"   # New season contains (yr)
qty_offline <- 3     # Qty to move to offline sales
month <- format(Sys.Date(), "%m")
if(month %in% c("09", "10", "11", "12", "01", "02")){in_season <- "F"}else(in_season <- "S")
RawData <- list.files(path = "../FBArefill/Raw Data File/", pattern = "All Marketplace All SKU Categories", full.names = T)
mastersku <- openxlsx::read.xlsx(list.files(path = "../FBArefill/Raw Data File/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
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
inventory[is.na(inventory)] <- 0
write.csv(inventory, file = paste0("../Analysis/Sales_Inventory_SKU_", Sys.Date(), ".csv"), row.names = F, na = "")
inventory_s <- inventory %>% count(Adjust_MSKU, wt = Inv_Total_End.WK) %>% `row.names<-`(toupper(.[, "Adjust_MSKU"]))
qty0 <- mastersku %>% select(MSKU.Status, Seasons.SKU, MSKU) %>% mutate(MSKU = toupper(MSKU), Inv_clover = clover[MSKU, "Quantity"], Inv_Total_EndWK = inventory_s[MSKU, "n"]) %>% 
  filter(!grepl(new_season, Seasons.SKU) & MSKU.Status == "Active" & Inv_clover == 0 & Inv_Total_EndWK == 0)
write.csv(qty0, file = paste0("../Analysis/discontinued_qty0_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\

# ------------- update monthly ratio and sales -------------
monR_last_yr <- read.csv(list.files(path = "../Analysis/", pattern = "MonthlyRatio2023_category_adjust_2024-01-25.csv", full.names = T), as.is = T) %>% `row.names<-`(.[, "Category"])
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

# ------------- move to website deals page -------------
#inventory <- read.csv("../Analysis/Sales_Inventory_SKU_2024-02-12.csv", as.is = T)
#sales_SKU_last12month <- read.csv("../Analysis/Sales_SKU_last12month_2024-01-31.csv", as.is = T)
#monR <- read.csv("../Analysis/MonthlyRatio_2024-01-31.csv", as.is = T) %>% `row.names<-`(.[, "Category"])
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

# ------------- low inventory to adjust ads -------------
#inventory <- read.csv("../Analysis/Sales_Inventory_SKU_2024-02-12.csv", as.is = T)
#sales_SKU_last12month <- read.csv("../Analysis/Sales_SKU_last12month_2024-01-31.csv", as.is = T)
#monR <- read.csv("../Analysis/MonthlyRatio_2024-01-31.csv", as.is = T) %>% `row.names<-`(.[, "Category"])
#if(month == "01"){last3m <- c("Month11", "Month12", "Month01")}else if(month %in% c("02", "03")){last3m <- sprintf("Month%02d", c(1:(as.numeric(month)-1), (as.numeric(month)-3+12):12))}else{last3m <- sprintf("Month%02d", c((as.numeric(month)-3):(as.numeric(month)-1)))}
#monR <- monR %>% mutate(T.last3m = rowSums(across(all_of(last3m)))) %>% `row.names<-`(.[, "Category"])
qty_refill <- 12
inventory <- inventory %>% mutate(Inv_WH_AMZ.CA.US = Inv_WH.JJ + Inv_AMZ.USA + Inv_AMZ.CA) %>% `row.names<-`(.[, "Adjust_MSKU"])
Sales_Inv_ThisMonth <- sales_SKU_last12month %>% filter(Sales.Channel %in% c("Amazon.com", "Amazon.ca", "janandjul.com")) %>% group_by(Adjust_MSKU) %>% summarise(Sales.last3m = sum(T.last3m)) %>% filter(Adjust_MSKU %in% inventory$Adjust_MSKU) %>%
  mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU), cat = mastersku[Adjust_MSKU, "Category.SKU"], Seasons = mastersku[Adjust_MSKU, "Seasons.SKU"], Status = mastersku[Adjust_MSKU, "MSKU.Status"], AMZ_flag = mastersku[Adjust_MSKU, "Pause.Plan.FBA_x000D_.CA/US/MX/AU/UK/DE"], monR_last3m = monR[cat, "T.last3m"], monR_this = monR[cat, paste0("Month", month)], Sales.this = Sales.last3m/monR_last3m*monR_this, Inv_WH_AMZ.CA.US = inventory[Adjust_MSKU, "Inv_WH_AMZ.CA.US"], Qty_left = ifelse(Inv_WH_AMZ.CA.US < Sales.this, 0, as.integer(Inv_WH_AMZ.CA.US - Sales.this)), Enough_Inv = (Qty_left >= qty_refill))
Sales_Inv_ThisMonth_SPU <- Sales_Inv_ThisMonth %>% group_by(SPU) %>% summarise(Enough_Inv = mean(Enough_Inv)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
Sales_Inv_ThisMonth <- Sales_Inv_ThisMonth %>% mutate(Percent_Sizes_SPU_Enough_Inv = paste0(as.character(as.integer(Sales_Inv_ThisMonth_SPU[SPU, "Enough_Inv"]*100)), "%")) %>% filter(grepl(in_season, Seasons), Status == "Active", !Enough_Inv) %>% arrange(Qty_left, Adjust_MSKU)
write.csv(Sales_Inv_ThisMonth, file = paste0("../Analysis/OOS_ThisMonth", Sys.Date(), ".csv"), row.names = F)
# Email Adrienne, Joren and Kamer 

# ------------- low inventory to add PO -------------
#inventory <- read.csv("../Analysis/Sales_Inventory_SKU_2024-01-31.csv", as.is = T) %>% mutate_all(~ replace(., is.na(.), 0))
#sales_SKU_last12month <- read.csv("../Analysis/Sales_SKU_last12month_2024-01-31.csv", as.is = T)
#monR <- read.csv("../Analysis/MonthlyRatio_2024-01-31.csv", as.is = T) %>% `row.names<-`(.[, "Category"])
#inventory <- inventory %>% mutate(Inv_WH_AMZ.CA.US = Inv_WH.JJ + Inv_AMZ.USA + Inv_AMZ.CA) %>% `row.names<-`(.[, "Adjust_MSKU"])
#if(month == "01"){last3m <- c("Month11", "Month12", "Month01")}else if(month %in% c("02", "03")){last3m <- sprintf("Month%02d", c(1:(as.numeric(month)-1), (as.numeric(month)-3+12):12))}else{last3m <- sprintf("Month%02d", c((as.numeric(month)-3):(as.numeric(month)-1)))}
if(month == "12"){next4m <- c("Month01", "Month02", "Month03", "Month04")}else if(month %in% c("10", "11")){next4m <- sprintf("Month%02d", c(1:(as.numeric(month)-12+3), as.numeric(month):12))}else{next4m <- sprintf("Month%02d", c(as.numeric(month):(as.numeric(month)+3)))}
PO <- rbind(openxlsx::read.xlsx("../PO/2024FW China to Global Shipment Tracking.xlsx", sheet = "2024FW", startRow = 4, fillMergedCells = T, cols = c(7, 15)), openxlsx::read.xlsx("../PO/2024SS China to Global Shipment Tracking.xlsx", sheet = "2024SS", startRow = 4, fillMergedCells = T, cols = c(7, 15))) %>% 
  filter(!is.na(SKU)) %>% group_by(SKU) %>% summarise(Remain = sum(as.numeric(Remain..in.Pcs))) %>% as.data.frame() %>% `row.names<-`(.[, "SKU"])
monR <- monR %>% mutate(T.next4m = rowSums(across(all_of(next4m)))) %>% `row.names<-`(.[, "Category"])
Sales_Inv_Next4m <- sales_SKU_last12month %>% filter(Sales.Channel %in% c("Amazon.com", "Amazon.ca", "janandjul.com")) %>% group_by(Adjust_MSKU) %>% summarise(Sales.last3m = sum(T.last3m)) %>% filter(Adjust_MSKU %in% inventory$Adjust_MSKU) %>%
  mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU), cat = mastersku[Adjust_MSKU, "Category.SKU"], Seasons = mastersku[Adjust_MSKU, "Seasons.SKU"], Status = mastersku[Adjust_MSKU, "MSKU.Status"], AMZ_flag = mastersku[Adjust_MSKU, "Pause.Plan.FBA_x000D_.CA/US/MX/AU/UK/DE"], monR_last3m = monR[cat, "T.last3m"], monR_next4m = monR[cat, "T.next4m"], Sales.next4m = Sales.last3m/monR_last3m*monR_next4m, Inv_WH_AMZ.CA.US = inventory[Adjust_MSKU, "Inv_WH_AMZ.CA.US"], PO_remain = ifelse(Adjust_MSKU %in% PO$SKU, PO[Adjust_MSKU, "Remain"], 0), Enough_Inv = (Inv_WH_AMZ.CA.US + PO_remain >= Sales.next4m))
Sales_Inv_Next4m_SPU <- Sales_Inv_Next4m %>% group_by(SPU) %>% summarise(Enough_Inv = mean(Enough_Inv)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
Sales_Inv_Next4m <- Sales_Inv_Next4m %>% mutate(Enough_Inv_SPU = paste0(as.character(as.integer(Sales_Inv_Next4m_SPU[SPU, "Enough_Inv"]*100)), "%")) %>% filter(grepl(new_season, Seasons), Status == "Active", !Enough_Inv)
write.csv(Sales_Inv_Next4m, file = paste0("../Analysis/OOS_Next4m_", Sys.Date(), ".csv"), row.names = F)
# Discuss with Mei, Florence, Cindy and Matt

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

# ------------- Warehouse sale ------------
defects <- xoro %>% filter(grepl("^M[A-Z][A-Z][A-Z]-", Item.), ATS > 0) %>% arrange(Item.)
write.csv(defects, file = paste0("../Analysis/defects_", Sys.Date(), ".csv"), row.names = F, na = "")

clover <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items") %>% `row.names<-`(toupper(.[, "Name"]))
offline <- xoro %>% mutate(qty = clover[Item., "Quantity"], Price = clover[Item., "Price"]) %>% filter(ATS <= 0, qty > 0, !grepl("24", Seasons)) %>% arrange(Item.)
write.csv(offline, file = paste0("../Analysis/offline_all_", Sys.Date(), ".csv"), row.names = F, na = "")

inventory <- read.csv("../Analysis/Sales_Inventory_SKU_2024-05-01.csv", as.is = T)
sales_SKU_last12month <- read.csv("../Analysis/Sales_SKU_last12month_2024-05-01.csv", as.is = T)
monR <- read.csv("../Analysis/MonthlyRatio_2024-05-01.csv", as.is = T) %>% `row.names<-`(.[, "Category"])
if(month == "01"){last3m <- c("Month11", "Month12", "Month01")}else if(month %in% c("02", "03")){last3m <- sprintf("Month%02d", c(1:(as.numeric(month)-1), (as.numeric(month)-3+12):12))}else{last3m <- sprintf("Month%02d", c((as.numeric(month)-3):(as.numeric(month)-1)))}
monR <- monR %>% mutate(T.last3m = rowSums(across(all_of(last3m)))) %>% `row.names<-`(.[, "Category"])
inventory_SPU <- inventory %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU)) %>% group_by(SPU) %>% summarise(qty_SPU = sum(Inv_Total_End.WK, na.rm = T), qty_WH = sum(Inv_WH.JJ, na.rm = T), qty_AMZ = qty_SPU - qty_WH) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SKU_last12month <- sales_SKU_last12month %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU), T.last3m = rowSums(across(all_of(last3m))), T.yr = rowSums(across(Month01:Month12)))
sales_SPU_last3month <- sales_SKU_last12month %>% group_by(SPU) %>% summarise(T.last3m = sum(T.last3m)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SPU_last3month_JJ <- sales_SKU_last12month %>% filter(Sales.Channel == "janandjul.com") %>% group_by(SPU) %>% summarise(T.last3m = sum(T.last3m)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SPU_last12month_AMZ <- sales_SKU_last12month %>% filter(grepl("Amazon", Sales.Channel)) %>% group_by(SPU) %>% summarise(T.last12m = sum(T.yr)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% filter(!is.na(Regular.price)) %>% filter(!duplicated(SKU), SKU != "") %>% 
  mutate(SKU = toupper(SKU), Qty = xoro[SKU, "ATS"], Seasons = ifelse(SKU %in% mastersku$MSKU, mastersku[SKU, "Seasons.SKU"], mastersku_adjust[SKU, "Seasons.SKU"]), discount = ifelse(is.na(Sale.price), 0, (Regular.price - Sale.price)/Regular.price), cat = mastersku[SKU, "Category.SKU"], SPU = gsub("(\\w+-\\w+)-.*", "\\1", SKU)) %>% 
  select(ID, SKU, Name, Seasons, cat, SPU, Regular.price, discount, Qty) %>% `row.names<-`(.[, "SKU"])
warehouse_sales <- woo %>% filter(!grepl(new_season, Seasons), Qty > qty_offline, Regular.price > 0) %>% mutate(Qty_SPU = inventory_SPU[SPU, "qty_SPU"], Qty_WH = inventory_SPU[SPU, "qty_WH"], Qty_AMZ = inventory_SPU[SPU, "qty_AMZ"], MonR_last3m = monR[cat, "T.last3m"], Sales_SPU_last3m = sales_SPU_last3month[SPU, "T.last3m"], Sales_SPU_last3m_JJ = sales_SPU_last3month_JJ[SPU, "T.last3m"], Sales_SPU_1yr = Sales_SPU_last3m/MonR_last3m, Sales_SPU_1yr_JJ = Sales_SPU_last3m_JJ/MonR_last3m, Sales_SPU_1yr_AMZ = sales_SPU_last12month_AMZ[SPU, "T.last12m"]) %>% 
  filter(!is.na(Qty_SPU)) %>% mutate(time_yrs = round(ifelse(Qty_AMZ >= Sales_SPU_1yr_AMZ, Qty_WH/Sales_SPU_1yr_JJ, Qty_SPU/Sales_SPU_1yr), digits = 1)) %>%
  filter(!grepl("23", Seasons), !grepl("24", Seasons), time_yrs > 1.5) %>% select(SKU, Name, Seasons, Regular.price, discount, Qty, time_yrs) %>% arrange(SKU)
write.csv(warehouse_sales, file = paste0("../Analysis/warehouse_sales_", Sys.Date(), ".csv"), row.names = F, na = "")

# ---------- ABC show 2024 -----------
xoro <- read.xlsx2(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".xlsx"), full.names = T), sheetIndex = 1) %>% filter(Store == "Warehouse - JJ") %>% 
  mutate(Item. = toupper(Item.), ATS = as.numeric(ATS)) %>% `row.names<-`(.[, "Item."])
clover <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items") %>% `row.names<-`(toupper(.[, "Name"]))
request_SKU <- read.csv("../TradeShows/ABC2024/request_SKU20240423.csv", as.is = T) %>% mutate(SKU = toupper(SKU), qty = xoro[SKU, "ATS"], On.PO = xoro[SKU, "On.PO"], clover = clover[SKU, "Quantity"], Warehouse = ifelse(qty > 0 | On.PO > 0, "Surrey", ifelse(clover > 0, "Clover", "Sample")))
write.csv(request_SKU, file = paste0("../TradeShows/ABC2024/Stock_Surrey", Sys.Date(), ".csv"), row.names = F) 
ABC_surrey <- request_SKU %>% filter(Warehouse == "Surrey") %>% select(SKU) %>% arrange(SKU) %>%
  mutate(StoreCode = "WH-JJ", ItemNumber = SKU, Qty = -1, LocationName = "BIN", UnitCost = "", ReasonCode = "SPC", Memo = "ABC trade show 2024", UploadRule = "D", AdjAccntName = "", TxnDate = "", ItemIdentifierCode = "", ImportError = "") %>% select(-SKU)
write.csv(ABC_surrey, file = paste0("../TradeShows/ABC2024/ABC_Surrey", Sys.Date(), ".csv"), row.names = F)
ABC_richmond <- request_SKU %>% filter(!(SKU %in% ABC_surrey$ItemNumber)) 
write.csv(ABC_richmond, file = paste0("../TradeShows/ABC2024/ABC_Richmond", Sys.Date(), ".csv"), row.names = F)
giveaway <- read.csv("../TradeShows/ABC2024/giveaway.csv", as.is = T) %>% mutate(qty = xoro[SKU, "ATS"], On.PO = xoro[SKU, "On.PO"])
write.csv(giveaway, file = paste0("../TradeShows/ABC2024/giveaway", Sys.Date(), ".csv"), row.names = F)

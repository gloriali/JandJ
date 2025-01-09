# SKUs analysis and management
library(dplyr)
library(openxlsx)
library(scales)
library(ggplot2)
library(tidyr)
library(xlsx)
library(gridExtra)
new_season <- "24"   # New season contains (yr)
qty_offline <- 3     # Qty to move to offline sales
month <- format(Sys.Date(), "%m")
date <- as.numeric(format(Sys.Date(), "%d"))/31
monthSS <- c("Month03", "Month04", "Month05", "Month06", "Month07", "Month08")
monthFW <- c("Month09", "Month10", "Month11", "Month12", "Month01", "Month02")
if(month %in% c("09", "10", "11", "12", "01", "02")){in_season <- "F"}else(in_season <- "S")
RawData <- list.files(path = "../FBArefill/Raw Data File/", pattern = "All Marketplace All SKU Categories", full.names = T)
mastersku <- openxlsx::read.xlsx(list.files(path = "../FBArefill/Raw Data File/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
mastersku_adjust <- mastersku %>% filter(!duplicated(Adjust.SKU)) %>% `row.names<-`(toupper(.[, "Adjust.SKU"])) 
netsuite_item_S <- read.csv(list.files(path = "../FBArefill/Raw Data File/", pattern = "Items_S_.*", full.names = T), as.is = T) %>%
  mutate(Name = toupper(Name), Seasons = ifelse(Name %in% mastersku$MSKU, mastersku[Name, "Seasons.SKU"], mastersku_adjust[Name, "Seasons.SKU"]), SPU = gsub("(\\w+-\\w+)-.*", "\\1", Name)) %>% `row.names<-`(.[, "Name"])

# ------------- move offline ---------------------------
offline <- netsuite_item_S %>% filter(!grepl(new_season, Seasons) & Warehouse.Available <= qty_offline & Warehouse.Available > 0) %>% arrange(Name) %>% 
  mutate(Date = format(Sys.Date(), "%m/%d/%Y"), TO.TYPE = "Surrey-Richmond", SEASON = "24F", FROM.WAREHOUSE = "WH-SURREY", TO.WAREHOUSE = "WH-RICHMOND", REF.NO = paste0("TO-S2R", format(Sys.Date(), "%y%m%d")), Memo = "Richmond Refill", ORDER.PLACED.BY = "Gloria Li", ITEM = Name, Quantity = Warehouse.Available) %>% select(Date:Quantity)
write.csv(offline, file = paste0("../Analysis/offline_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\

# ------------- to inactivate --------------------------
clover <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items") %>% filter(!is.na(Name)) %>% `row.names<-`(toupper(.[, "Name"]))
sheets <- grep("-WK", getSheetNames(RawData), value = T)
inventory <- data.frame()
for(sheet in sheets){
  inventory_i <- openxlsx::read.xlsx(RawData, sheet = sheet, startRow = 2)
  inventory <- rbind(inventory, inventory_i)
}
inventory[is.na(inventory)] <- 0
inventory <- inventory %>% rename("Inv_Total_End.WK" = "Inv_Total_End.WK.(Not.Inc..CN)")
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
sales_SKU_last12month <- sales_SKU_last12month %>% rename(all_of(names)) %>% select(Sales.Channel, Adjust_MSKU, Month01:Month12) %>% mutate(across(Month01:Month12, ~as.numeric(.))) %>% mutate(across(as.name(paste0("Month", month)), ~round(./date, 0)))
write.csv(sales_SKU_last12month, file = paste0("../Analysis/Sales_SKU_last12month_", Sys.Date(), ".csv"), row.names = F)
sales_cat_last12month <- sales_cat_last12month %>% select(-(1:5))
write.csv(sales_cat_last12month, file = paste0("../Analysis/Sales_categories_last12month_", Sys.Date(), ".csv"), row.names = F)
if(month %in% c("01")){
  monR <- monR_last_yr %>% select(Category, Month01:Month12) %>% `row.names<-`(toupper(.[, "Category"]))
}else{
  monR_last_yr <- monR_last_yr %>% rowwise() %>% mutate(T.new = sum(c_across(Month01:as.name(sprintf("Month%02d", (as.numeric(month))))))) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "Category"]))
  sales_cat_last12month <- sales_cat_last12month %>% rename(all_of(names)) %>% mutate(across(Month01:Month12, ~as.numeric(.))) %>% mutate(across(as.name(paste0("Month", month)), ~round(./date, 0))) %>% 
    rowwise() %>% mutate(T.new = sum(c_across(Month01:as.name(sprintf("Month%02d", (as.numeric(month)))))), MonR.new = monR_last_yr[Category, "T.new"]) %>% 
    filter(!is.na(MonR.new)) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "Category"]))
  monR_new <- sales_cat_last12month %>% mutate(across(Month01:as.name(sprintf("Month%02d", (as.numeric(month)))), ~ ./T.new*MonR.new))
  monR_new[monR_new$T.new == 0, ] <- monR_new[monR_new$T.new == 0, ] %>% rows_update(., monR_last_yr %>% select(Category, Month01:as.name(sprintf("Month%02d", (as.numeric(month))))), by = "Category", unmatched = "ignore")
  monR <- merge(monR_new %>% select(Category, Month01:as.name(sprintf("Month%02d", (as.numeric(month))))), monR_last_yr %>% select(Category, as.name(sprintf("Month%02d", as.numeric(month)+1)):Month12), by = "Category") %>% `row.names<-`(toupper(.[, "Category"]))
}
write.csv(monR, file = paste0("../Analysis/MonthlyRatio_", Sys.Date(), ".csv"), row.names = F)

# ------------- move to website deals page -------------
#inventory <- read.csv("../Analysis/Sales_Inventory_SKU_2024-06-03.csv", as.is = T)
#sales_SKU_last12month <- read.csv("../Analysis/Sales_SKU_last12month_2024-06-03.csv", as.is = T)
#monR <- read.csv("../Analysis/MonthlyRatio_2024-06-03.csv", as.is = T) %>% `row.names<-`(.[, "Category"])
discount_method <- openxlsx::read.xlsx("../Analysis/JJ_discount_method.xlsx", sheet = 1, startRow = 2, rowNames = T)
if(month == "01"){last3m <- c("Month11", "Month12", "Month01")}else if(month %in% c("02", "03")){last3m <- sprintf("Month%02d", c(1:(as.numeric(month)-1), (as.numeric(month)-3+12):12))}else{last3m <- sprintf("Month%02d", c((as.numeric(month)-3):(as.numeric(month)-1)))}
monR <- monR %>% mutate(T.last3m = rowSums(across(all_of(last3m)))) %>% `row.names<-`(.[, "Category"])
inventory_SPU <- inventory %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU)) %>% group_by(SPU) %>% summarise(qty_SPU = sum(Inv_Total_End.WK, na.rm = T), qty_WH = sum(Inv_WH.JJ, na.rm = T), qty_AMZ = qty_SPU - qty_WH) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SKU_last12month <- sales_SKU_last12month %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU), T.last3m = rowSums(across(all_of(last3m))), T.yr = rowSums(across(Month01:Month12)))
sales_SPU_last3month <- sales_SKU_last12month %>% group_by(SPU) %>% summarise(T.last3m = sum(T.last3m)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SPU_last3month_JJ <- sales_SKU_last12month %>% filter(Sales.Channel == "janandjul.com") %>% group_by(SPU) %>% summarise(T.last3m = sum(T.last3m)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SPU_last12month_AMZ <- sales_SKU_last12month %>% filter(grepl("Amazon", Sales.Channel)) %>% group_by(SPU) %>% summarise(T.last12m = sum(T.yr)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sizes_all <- netsuite_item_S %>% mutate(SPU_size = gsub("(\\w+-\\w+-\\w+).*", "\\1", Name)) %>% filter(!duplicated(SPU_size)) %>% count(SPU) %>% `row.names<-`(.[, "SPU"])
size_limited <- netsuite_item_S %>% filter(Warehouse.Available > qty_offline) %>% count(SPU) %>% mutate(all = sizes_all[SPU, "n"], percent = ifelse(all < n, 0, (all - n)/all)) %>% `row.names<-`(.[, "SPU"])
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% filter(Published == 1, !is.na(Regular.price)) %>% filter(!duplicated(SKU), SKU != "") %>% 
  mutate(SKU = toupper(SKU), Qty = netsuite_item_S[SKU, "Warehouse.Available"], Seasons = ifelse(SKU %in% mastersku$MSKU, mastersku[SKU, "Seasons.SKU"], mastersku_adjust[SKU, "Seasons.SKU"]), discount = ifelse(is.na(Sale.price), 0, (Regular.price - Sale.price)/Regular.price), cat = mastersku[SKU, "Category.SKU"], SPU = gsub("(\\w+-\\w+)-.*", "\\1", SKU)) %>% 
  select(ID, SKU, Name, Seasons, cat, SPU, Regular.price, discount, Qty) %>% `row.names<-`(.[, "SKU"])
overstock <- woo %>% filter(!grepl(new_season, Seasons), Qty > qty_offline, Regular.price > 0) %>% mutate(Qty_SPU = inventory_SPU[SPU, "qty_SPU"], Qty_WH = inventory_SPU[SPU, "qty_WH"], Qty_AMZ = inventory_SPU[SPU, "qty_AMZ"], MonR_last3m = monR[cat, "T.last3m"], Sales_SPU_last3m = sales_SPU_last3month[SPU, "T.last3m"], Sales_SPU_last3m_JJ = sales_SPU_last3month_JJ[SPU, "T.last3m"], Sales_SPU_1yr = Sales_SPU_last3m/MonR_last3m, Sales_SPU_1yr_JJ = Sales_SPU_last3m_JJ/MonR_last3m, Sales_SPU_1yr_AMZ = sales_SPU_last12month_AMZ[SPU, "T.last12m"]) %>% 
  filter(!is.na(Qty_SPU)) %>% mutate(time_yrs = round(ifelse(Qty_AMZ >= Sales_SPU_1yr_AMZ, Qty_WH/Sales_SPU_1yr_JJ, Qty_SPU/Sales_SPU_1yr), 1), size_percent_missing = gsub("Sizes.missing.0%", "Sizes.missing.<.50%", paste0("Sizes.missing.", as.character(as.integer(ifelse(size_limited[SPU, "percent"] < 0.5, 0, size_limited[SPU, "percent"])*10)*10), "%"))) %>% 
  filter(time_yrs >= 1) %>% select(SKU:Qty, time_yrs, size_percent_missing) %>% arrange(SKU)
write.csv(overstock, file = paste0("../Analysis/Overstock_", Sys.Date(), ".csv"), row.names = F)
# Email Will
woo_deals <- woo %>% filter(!grepl(new_season, Seasons), Qty > qty_offline, Regular.price > 0) %>% mutate(Qty_SPU = inventory_SPU[SPU, "qty_SPU"], Qty_WH = inventory_SPU[SPU, "qty_WH"], Qty_AMZ = inventory_SPU[SPU, "qty_AMZ"], MonR_last3m = monR[cat, "T.last3m"], Sales_SPU_last3m = sales_SPU_last3month[SPU, "T.last3m"], Sales_SPU_last3m_JJ = sales_SPU_last3month_JJ[SPU, "T.last3m"], Sales_SPU_1yr = Sales_SPU_last3m/MonR_last3m, Sales_SPU_1yr_JJ = Sales_SPU_last3m_JJ/MonR_last3m, Sales_SPU_1yr_AMZ = sales_SPU_last12month_AMZ[SPU, "T.last12m"]) %>% 
  filter(!is.na(Qty_SPU)) %>% mutate(time_yrs = ifelse(Qty_AMZ >= Sales_SPU_1yr_AMZ, Qty_WH/Sales_SPU_1yr_JJ, Qty_SPU/Sales_SPU_1yr), time_yrs = paste0("Sold out ", as.character(as.integer(ifelse(time_yrs > 3, 3, time_yrs)) + 1), " yr"), size_percent_missing = gsub("Sizes.missing.0%", "Sizes.missing.<.50%", paste0("Sizes.missing.", as.character(as.integer(ifelse(size_limited[SPU, "percent"] < 0.5, 0, size_limited[SPU, "percent"])*10)*10), "%"))) %>%
  rowwise() %>% mutate(Suggest_discount = as.numeric(gsub("% Off", "", discount_method[time_yrs, size_percent_missing]))/100, Suggest_price = round((1 - Suggest_discount) * Regular.price, digits = 2)) %>% arrange(-Suggest_discount, SKU) %>% filter(Suggest_discount > 0, Suggest_discount > discount) 
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

# ------------- Bulk dead inventory ------------- 
#sales_SKU_last12month <- read.csv("../Analysis/Sales_SKU_last12month_2024-01-31.csv", as.is = T)
sales_SPU_last12month <- sales_SKU_last12month %>% mutate(SPU = paste0(mastersku[Adjust_MSKU, "Category.SKU"], "-", mastersku[Adjust_MSKU, "Print.SKU"]), T.SS = rowSums(across(all_of(monthSS))), T.FW = rowSums(across(all_of(monthFW)))) %>% group_by(SPU) %>% summarise(T.SS = sum(T.SS), T.FW = sum(T.FW)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% filter(Published == 1, !is.na(Regular.price)) %>% filter(!duplicated(SKU), SKU != "") %>% 
  mutate(SKU = toupper(SKU), Qty = xoro[SKU, "ATS"], Seasons = ifelse(SKU %in% mastersku$MSKU, mastersku[SKU, "Seasons.SKU"], mastersku_adjust[SKU, "Seasons.SKU"]), discount = ifelse(is.na(Sale.price), 0, (Regular.price - Sale.price)/Regular.price), cat = mastersku[SKU, "Category.SKU"], SPU = gsub("(\\w+-\\w+)-.*", "\\1", SKU)) %>% 
  select(ID, SKU, Name, Seasons, cat, SPU, Regular.price, discount, Qty) %>% `row.names<-`(.[, "SKU"])
discontinued <- woo %>% filter(!grepl(new_season, Seasons), Qty > qty_offline, Regular.price > 0) %>% group_by(SPU) %>% summarise(Seasons = Seasons[1], Qty = sum(Qty)) %>% mutate(sales.SS = ifelse(SPU %in% sales_SPU_last12month$SPU, sales_SPU_last12month[SPU, "T.SS"], 0), sales.FW = ifelse(SPU %in% sales_SPU_last12month$SPU, sales_SPU_last12month[SPU, "T.FW"], 0), time_yrs = Qty/(sales.SS + sales.FW))
bulk_dead_SS <- discontinued %>% filter(!(grepl("F", Seasons)), time_yrs > 1, Qty > 500)
write.csv(bulk_dead_SS, file = paste0("../Analysis/bulk_dead_SS_", Sys.Date(), ".csv"), row.names = F)
bulk_dead_FW <- discontinued %>% filter(!(grepl("S", Seasons)), time_yrs > 1, Qty > 500)
write.csv(bulk_dead_FW, file = paste0("../Analysis/bulk_dead_FW_", Sys.Date(), ".csv"), row.names = F)

# ------------- Discontinued limit sizes ------------- 
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% filter(Published == 1, !is.na(Regular.price)) %>% filter(!duplicated(SKU), SKU != "") %>% 
  mutate(SKU = toupper(SKU), Qty = xoro[SKU, "ATS"], Seasons = ifelse(SKU %in% mastersku$MSKU, mastersku[SKU, "Seasons.SKU"], mastersku_adjust[SKU, "Seasons.SKU"]), discount = ifelse(is.na(Sale.price), 0, (Regular.price - Sale.price)/Regular.price), cat = mastersku[SKU, "Category.SKU"], SPU = gsub("(\\w+-\\w+)-.*", "\\1", SKU)) %>% 
  select(ID, SKU, Name, Seasons, cat, SPU, Regular.price, discount, Qty) %>% `row.names<-`(.[, "SKU"])
sizes_all <- woo %>% count(SPU) %>% mutate(cat = gsub("-.*", "", SPU)) %>% group_by(cat) %>% summarize(n_all = max(n)) %>% as.data.frame() %>% `row.names<-`(.[, "cat"])
size_limited <- woo %>% filter(!grepl(new_season, Seasons), Qty > qty_offline, Regular.price > 0) %>% count(SPU) %>% mutate(cat = gsub("-.*", "", SPU), n_all = sizes_all[cat, "n_all"]) %>% filter(n != n_all)
write.csv(size_limited, file = paste0("../Analysis/size_limited_", Sys.Date(), ".csv"), row.names = F)

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

# ----------- Sales trend -------------
categories <- read.csv("../Analysis/Categories.csv", as.is = T) %>% `row.names<-`(.[, "Category.SKU"])
mastersku <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../FBArefill/Raw Data File/", pattern = "1-MasterSKU-All-Product-", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
sales_cat_last12month <- read.csv(rownames(file.info(list.files(path = "../Analysis/", pattern = "Sales_categories_last12month_", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% `row.names<-`(.[, "Cat_SKU"]) %>% 
  mutate(Category.Group = categories[Cat_SKU, "Category.Group"]) %>% filter(!is.na(Category.Group), Cat_SKU %in% (mastersku %>% filter(grepl("2[34]", Seasons)))$Category.SKU, Total.2023 != 0)
pdf(paste0("../Analysis/Sales_trend_", Sys.Date(), ".pdf"), width = 9, height = 4)
for(group in unique(sales_cat_last12month$Category.Group)){
  cat_sales_annual <- sales_cat_last12month %>% filter(Category.Group == group) %>% select(Cat_SKU, contains("Total")) %>% 
    pivot_longer(!Cat_SKU, names_to = "Year", values_to = "Sales") %>% mutate(Year = gsub("Total\\.", "", Year))
  cat_sales_annual_plot <- ggplot(cat_sales_annual, aes(Year, Sales, color = Cat_SKU, group = Cat_SKU)) + geom_point() + geom_line() + 
    labs(title = paste0(group, "\nSales for the last 4 years"), color = "") + xlab("") + ylab("") + theme_bw() + theme(plot.title = element_text(hjust = 0.5))
  cat_sales_monthly <- sales_cat_last12month %>% filter(Category.Group == group) %>% select(-contains("Total"), -Category.Group) %>% 
    pivot_longer(!Cat_SKU, names_to = "Month", values_to = "Sales") %>% mutate(Month = factor(Month, levels = Month[1:12]))
  cat_sales_monthly_plot <- ggplot(cat_sales_monthly, aes(Month, Sales, color = Cat_SKU, group = Cat_SKU)) + geom_point() + geom_line() + 
    labs(title = paste0(group, "\nSales for the last 12 months"), color = "") + xlab("") + theme_bw() + theme(plot.title = element_text(hjust = 0.5)) + guides(color = "none")
  grid.arrange(cat_sales_monthly_plot, cat_sales_annual_plot, ncol = 2, widths = c(3, 2))
}
dev.off()

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

# ------------- SS clearance -------------
inventory <- read.csv("../Analysis/Sales_Inventory_SKU_2024-07-22.csv", as.is = T)
sales_SKU_last12month <- read.csv("../Analysis/Sales_SKU_last12month_2024-07-22.csv", as.is = T)
monR <- read.csv("../Analysis/MonthlyRatio_2024-07-22.csv", as.is = T) %>% `row.names<-`(.[, "Category"])
discount_method <- openxlsx::read.xlsx("../Analysis/JJ_discount_method.xlsx", sheet = 1, startRow = 2, rowNames = T)
if(month == "01"){last3m <- c("Month11", "Month12", "Month01")}else if(month %in% c("02", "03")){last3m <- sprintf("Month%02d", c(1:(as.numeric(month)), (as.numeric(month)-2+12):12))}else{last3m <- sprintf("Month%02d", c((as.numeric(month)-2):(as.numeric(month))))}
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
SS_clearance <- woo %>% filter(grepl("S", Seasons), Qty > qty_offline, Regular.price > 0) %>% mutate(Qty_SPU = inventory_SPU[SPU, "qty_SPU"], Qty_WH = inventory_SPU[SPU, "qty_WH"], Qty_AMZ = inventory_SPU[SPU, "qty_AMZ"], MonR_last3m = monR[cat, "T.last3m"], Sales_SPU_last3m = sales_SPU_last3month[SPU, "T.last3m"], Sales_SPU_last3m_JJ = sales_SPU_last3month_JJ[SPU, "T.last3m"], Sales_SPU_1yr = Sales_SPU_last3m/MonR_last3m, Sales_SPU_1yr_JJ = Sales_SPU_last3m_JJ/MonR_last3m, Sales_SPU_1yr_AMZ = sales_SPU_last12month_AMZ[SPU, "T.last12m"]) %>% 
  mutate(time = Qty_SPU/Sales_SPU_1yr, time_yrs = paste0("Sold out ", as.character(as.integer(ifelse(time > 3, 3, time)) + 1), " yr"), size_percent_missing = gsub("Sizes.missing.0%", "Sizes.missing.<.50%", paste0("Sizes.missing.", as.character(as.integer(ifelse(size_limited[SPU, "percent"] < 0.5, 0, size_limited[SPU, "percent"])*10)*10), "%"))) %>%
  rowwise() %>% mutate(Suggest_discount = as.numeric(gsub("% Off", "", discount_method[time_yrs, size_percent_missing]))/100, Suggest_price = round((1 - Suggest_discount) * Regular.price, digits = 2)) %>% arrange(-Suggest_discount, SKU) %>% filter(time > 1 | size_percent_missing != "Sizes.missing.<.50%" | !grepl("24", Seasons))
write.csv(SS_clearance, file = paste0("../Analysis/Clearance_2024SS_", Sys.Date(), ".csv"), row.names = F, na = "")

# ------------- Warehouse sale ------------
defects <- xoro %>% filter(grepl("^M[A-Z][A-Z][A-Z]-", Item.), ATS > 5) %>% arrange(Item.)
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
  filter(!grepl("24", Seasons), time_yrs > 1.5) %>% select(SKU, Name, Seasons, Regular.price, discount, Qty, time_yrs) %>% arrange(SKU)
write.csv(warehouse_sales, file = paste0("../Analysis/hat_sales_", Sys.Date(), ".csv"), row.names = F, na = "")
### list of items to sell for TJX
TJX <- woo %>% filter(!grepl(new_season, Seasons), Qty > 5, Regular.price > 0) %>% mutate(Wholesale.price = round(Regular.price/2, 2), Qty_SPU = inventory_SPU[SPU, "qty_SPU"], Qty_WH = inventory_SPU[SPU, "qty_WH"], Qty_AMZ = inventory_SPU[SPU, "qty_AMZ"], MonR_last3m = monR[cat, "T.last3m"], Sales_SPU_last3m = sales_SPU_last3month[SPU, "T.last3m"], Sales_SPU_last3m_JJ = sales_SPU_last3month_JJ[SPU, "T.last3m"], Sales_SPU_1yr = Sales_SPU_last3m/MonR_last3m, Sales_SPU_1yr_JJ = Sales_SPU_last3m_JJ/MonR_last3m, Sales_SPU_1yr_AMZ = sales_SPU_last12month_AMZ[SPU, "T.last12m"]) %>% 
  filter(!is.na(Qty_SPU)) %>% mutate(time_yrs = round(ifelse(Qty_AMZ >= Sales_SPU_1yr_AMZ, Qty_WH/Sales_SPU_1yr_JJ, Qty_SPU/Sales_SPU_1yr), digits = 1)) %>%
  filter(!grepl("24", Seasons), time_yrs >= 4) %>% select(SKU, Name, Seasons, Regular.price, Wholesale.price, Qty, time_yrs) %>% arrange(SKU)
write.csv(TJX, file = paste0("../Analysis/TJX_", Sys.Date(), ".csv"), row.names = F, na = "")

### restock WJA, WPS, WJT, WMT
warehouse_sales <- read.csv("../Analysis/warehouse_sales_2024-06-04.csv", as.is = T) %>% mutate(Sales.price = gsub("\\$", "", Sales.price)) %>% `row.names<-`(.[, "SKU"])
warehouse_sales_new <- data.frame(Type = "Discontinued", SKU = c((xoro %>% filter(!grepl("24", Seasons), grepl("^WMT-", Item.), ATS > 0))$Item., (xoro %>% filter(!grepl("24", Seasons), grepl("^WJA-", Item.), ATS > 0))$Item., (xoro %>% filter(!grepl("24", Seasons), grepl("^WJT-", Item.), ATS > 0))$Item., (xoro %>% filter(!grepl("24", Seasons), grepl("^WPS-", Item.), ATS > 0))$Item., (xoro %>% filter(grepl("^BRC-DNL", Item.), ATS > 0))$Item., (xoro %>% filter(grepl("^BRC-TRZ", Item.), ATS > 0))$Item., (xoro %>% filter(grepl("^KEH-", Item.), ATS > 0))$Item., (xoro %>% filter(grepl("^KMT-", Item.), ATS > 0))$Item.)) %>% 
  mutate(Sales.price = ifelse(grepl("WJA", SKU), 50, 25), Sales.price = ifelse(grepl("^WMT", SKU), 15, Sales.price), Sales.price = ifelse(grepl("^K", SKU), 10, Sales.price), Sales.price = ifelse(grepl("^BRC", SKU), 30, Sales.price), Sales.price = ifelse(grepl("^WJT", SKU), 45, Sales.price), Surrey_Qty = xoro[SKU, "ATS"], Seasons = mastersku[SKU, "Seasons"], Qty_transfer = 0)
warehouse_sales <- rbind(warehouse_sales, warehouse_sales_new) %>% mutate(Seasons = mastersku[SKU, "Seasons"]) %>% mutate(Surrey_Qty = xoro[SKU, "ATS"], Clover_Qty = clover_item[SKU, "Quantity"]) %>% filter(Surrey_Qty > 0)
write.csv(warehouse_sales, file = "../Analysis/warehouse_sales_2024-06-12.csv", row.names = F, na = "")
### setup clover and square
price_deal <- read.csv("../Analysis/price_deal.csv", as.is = T) %>% `row.names<-`(.[, "SKU"])
warehouse_sales <- read.csv("../Analysis/warehouse_sales_2024-06-12.csv", as.is = T) %>% mutate(Sales.price = gsub("\\$", "", Sales.price)) %>% `row.names<-`(.[, "SKU"])
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% 
  mutate(Price = ifelse(Name %in% warehouse_sales$SKU, warehouse_sales[Name, "Sales.price"], Price), Price = ifelse(Name %in% price_deal$Name, price_deal[Name, "Price"], Price), Price = ifelse(grepl("^BRC", Name), "30.00", Price), Price = ifelse(grepl("^KMT", Name), "8.00", Price), Price = ifelse(grepl("^WJT", Name), "45.00", Price), Price = ifelse(grepl("^WPS", Name), "25.00", Price), Price = ifelse(grepl("^WJA", Name), "50.00", Price), Price = ifelse(grepl("^WMT", Name), "15.00", Price), Price = ifelse(grepl("SSS", Name), "10.00", Price), Price = ifelse(grepl("SSW", Name), "10.00", Price), Price = ifelse(grepl("BTP", Name), "40.00", Price), Price = ifelse(grepl("WSS-CNL", Name), "40.00", Price), Price = ifelse(grepl("WSS-DNL", Name), "40.00", Price), Price = ifelse(grepl("WSS-TRZ", Name), "40.00", Price), Price = ifelse(grepl("WSS-UNC", Name), "40.00", Price), Price = ifelse(grepl("ICP", Name), "60.00", Price), Price.Type = ifelse(is.na(Price), "Variable", "Fixed")) %>% `row.names<-`(.[, "Name"])
#defective <- warehouse_sales %>% filter(Type == "Defective") %>% mutate(Original = gsub("^M", "", SKU), Product.Code = clover_item[Original, "Product.Code"])
#clover_new <- data.frame(Clover.ID = "", Name = defective$SKU, Alternate.Name = "", Price = defective$Sales.price, Price.Type = "Fixed", Price.Unit = "", Tax.Rates = "GST", Cost = 0, Product.Code = defective$Product.Code, SKU = defective$SKU, Modifier.Groups = "", Quantity = 0, Printer.Labels = "", Hidden = "No", Non.revenue.item = "No") 
#clover_new <- clover_new %>% rename_with(~ gsub("Non.revenue.item", "Non-revenue.item", colnames(clover_new)))
#clover_item <- rbind(clover_item, clover_new) %>% mutate(Product.Code = ifelse(Name %in% defective$Original, "", Product.Code))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item), gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price)) %>% `row.names<-`(toupper(.[, "SKU"])) 
square_upload <- data.frame(Token = "", ItemName = clover_item$Name, ItemType = "", VariationName = "Regular", SKU = clover_item$Name, Description = "", Category = "", Price = clover_item$Price, PriceRichmond = "", PriceSurrey = "", NewQuantityRichmond = 0, NewQuantitySurrey = 0, EnabledRichmond = "Y", CurrentQuantityRichmond = "", CurrentQuantitySurrey = "", StockAlertEnabledRichmond = "N", StockAlertCountRichmond = 0, EnabledSurrey = "Y", StockAlertEnabledSurrey = "N", StockAlertCountSurrey = 0, SEOTitle = "",	SEODescription = "",	Permalink = "", SquareOnlineItemVisibility = "UNAVAILABLE",	ShippingEnabled = "",	SelfserveOrderingEnabled = "",	DeliveryEnabled = "",	PickupEnabled = "", Sellable = "Y",	Stockable = "Y",	SkipDetailScreeninPOS = "",	OptionName1 = "",	OptionValue1 = "") %>% 
  mutate(TaxPST = ifelse(woo[SKU, "Tax.class"] == "full", "Y", "N"), TaxPST = ifelse(is.na(TaxPST), "N", TaxPST))
write.table(square_upload, file = paste0("../Square/square-upload-", Sys.Date(), ".csv"), sep = ",", row.names = F, col.names = c(colnames(square_upload)[1:ncol(square_upload)-1], "Tax - PST (7%)"), na = "")
### update clover with square sales and refill
xoro <- read.xlsx2(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".xlsx"), full.names = T), sheetIndex = 1) %>% filter(Store == "Warehouse - JJ") %>% mutate(Item. = toupper(Item.), ATS = as.numeric(ATS)) %>% `row.names<-`(toupper(.[, "Item."])) 
warehouse_sales <- read.csv("../Analysis/warehouse_sales_2024-06-12.csv", as.is = T) %>% mutate(Sales.price = gsub("\\$", "", Sales.price)) %>% `row.names<-`(.[, "SKU"])
#square <- read.csv(list.files(path = "../Square/", pattern = paste0("catalog-", format(Sys.Date(), "%Y-%m-%d")), full.names = T), as.is = T) %>% `row.names<-`(toupper(.[, "SKU"])) 
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
#clover_item <- openxlsx::readWorkbook(clover, "Items") %>% mutate(Quantity = Quantity + square[Name, "Current.Quantity.Richmond"] + square[Name, "Current.Quantity.Surrey"]) %>% `row.names<-`(.[, "Name"])
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% `row.names<-`(.[, "Name"])
order <- data.frame(Item = (clover_item %>% filter(Name %in% warehouse_sales$SKU, Quantity < 10))$Name) %>% 
  mutate(Qty = clover_item[Item, "Quantity"], ATS = xoro[Item, "ATS"], cat = gsub("-.*", "", Item), size = gsub("\\w+-\\w+-", "", Item), Qty_transfer = ifelse(ATS < 10, ATS, 20 - Qty)) %>% filter(ATS > 0) %>% arrange(cat, size) %>% select(-c("cat", "size"))
#order <- clover_item %>% filter(grepl("^WJ", Name)) %>% select(Name, Quantity) %>% mutate(Seasons = mastersku[Name, "Seasons"], cat = gsub("-.*", "", Name), size = gsub("\\w+-\\w+-", "", Name), cat_size = paste0(cat, "-", size), Qty_xoro = xoro[Name, "ATS"]) %>% filter(Qty_xoro + Quantity > 0, !grepl("24", Seasons)) %>% arrange(cat_size)
write.csv(order, file = paste0("../Clover/order", format(Sys.Date(), "%m%d%Y"), ".csv"), row.names = F, na = "")
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item), gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
square_upload <- data.frame(Token = "", ItemName = clover_item$Name, ItemType = "", VariationName = "Regular", SKU = clover_item$Name, Description = "", Category = "", Price = clover_item$Price, PriceRichmond = "", PriceSurrey = "", NewQuantityRichmond = 0, NewQuantitySurrey = 0, EnabledRichmond = "Y", CurrentQuantityRichmond = "", CurrentQuantitySurrey = "", StockAlertEnabledRichmond = "N", StockAlertCountRichmond = 0, EnabledSurrey = "Y", StockAlertEnabledSurrey = "N", StockAlertCountSurrey = 0, SEOTitle = "",	SEODescription = "",	Permalink = "", SquareOnlineItemVisibility = "UNAVAILABLE",	ShippingEnabled = "",	SelfserveOrderingEnabled = "",	DeliveryEnabled = "",	PickupEnabled = "", Sellable = "Y",	Stockable = "Y",	SkipDetailScreeninPOS = "",	OptionName1 = "",	OptionValue1 = "") %>% 
  mutate(TaxPST = ifelse(woo[SKU, "Tax.class"] == "full", "Y", "N"), TaxPST = ifelse(is.na(TaxPST), "N", TaxPST))
write.table(square_upload, file = paste0("../Square/square-upload-", Sys.Date(), ".csv"), sep = ",", row.names = F, col.names = c(colnames(square_upload)[1:ncol(square_upload)-1], "Tax - PST (7%)"), na = "")
### clover receive stock
order <- read.xlsx2("../Clover/order_2024-06-12_packed.xlsx", sheetIndex = 1) %>% `row.names<-`(.[, "SKU"])
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% mutate(Quantity = ifelse(Name %in% order$SKU, Quantity + as.numeric(order[Name, "QTY_picked"]), Quantity), Price = ifelse(Name %in% order$SKU, order[Name, "Sales.price"], Price))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item), gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)

## hats sales for online
sales_SKU_thisyear <- data.frame()
sheets <- getSheetNames(RawData)
sheets_SKU <- sheets[grepl("-SKU$", sheets)&(!grepl("PW2", sheets))&(!grepl("WJO", sheets))&(!grepl("WPO", sheets))&(!grepl("ISJ", sheets))&(!grepl("WJS", sheets))]
for(sheet in sheets_SKU){
  sales_SKU_i <- openxlsx::read.xlsx(RawData, sheet = sheet, cols = c(1:8), startRow = 2) %>% 
    filter(!(Sales.Channel %in% c("Summary", "Sales Channel", "All Marketplaces"))) %>% rename_with(~ gsub("20", "Total.20", gsub("_x000D_.Total", "", .x)))
  sales_SKU_thisyear <- rbind(sales_SKU_thisyear, sales_SKU_i)
}
sales_SKU_thisyearT <- sales_SKU_thisyear %>% group_by(Adjust_MSKU) %>% summarise(Total.2024 = sum(Total.2024)) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "Adjust_MSKU"]))
sales_SKU_thisyearAMZ <- sales_SKU_thisyear %>% filter(grepl("Amazon", Sales.Channel)) %>% group_by(Adjust_MSKU) %>% summarise(Total.2024 = sum(Total.2024)) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "Adjust_MSKU"]))
sales_SKU_thisyearJJ <- sales_SKU_thisyear %>% filter(grepl("janandjul", Sales.Channel)) %>% group_by(Adjust_MSKU) %>% summarise(Total.2024 = sum(Total.2024)) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "Adjust_MSKU"]))

categories <- c("AAA", "ACA", "ACB", "AHJ", "AJP", "AJS", "HAD0", "HAV0", "HBS", "HCA0", "HCB0", "HCF0", "HJP", "HJS", "HXC", "HXP", "HXU")
sizes_all <- data.frame(cat = categories, n_all = c(2, 1, 1, 1, 3, 2, 4, 4, 3, 4, 4, 4, 1, 1, 2, 2, 2)) %>% `row.names<-`(.[, "cat"])
hats_sales <- read.csv("../Analysis/hats_sales_2024-05-27.csv", as.is = T) %>% mutate(Total.2024 = sales_SKU_thisyearT[toupper(SKU), "Total.2024"], Total.2024.JJ = sales_SKU_thisyearJJ[toupper(SKU), "Total.2024"], Total.2024.AMZ = sales_SKU_thisyearAMZ[toupper(SKU), "Total.2024"])
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% filter(!is.na(Regular.price)) %>% filter(!duplicated(SKU), SKU != "") %>% 
  mutate(SKU = toupper(SKU), Qty = xoro[SKU, "ATS"], Seasons = ifelse(SKU %in% mastersku$MSKU, mastersku[SKU, "Seasons.SKU"], mastersku_adjust[SKU, "Seasons.SKU"]), discount = ifelse(is.na(Sale.price), 0, (Regular.price - Sale.price)/Regular.price), cat = mastersku[SKU, "Category.SKU"], SPU = gsub("(\\w+-\\w+)-.*", "\\1", SKU)) %>% 
  select(SKU, Name, Seasons, cat, SPU, Regular.price, discount, Qty) %>% `row.names<-`(.[, "SKU"])
hats_sales_old <- woo %>% filter(cat %in% categories, Qty > qty_offline, !grepl("24", Seasons), !grepl("23", Seasons), !grepl("22", Seasons), !(SKU %in% hats_sales$SKU)) 
hats_limit <- woo %>% filter(cat %in% categories, Qty > qty_offline) %>% count(SPU) %>% mutate(cat = gsub("-.*", "", SPU), n_all = sizes_all[cat, "n_all"]) %>% filter(n != n_all)
hats_sales_limit <- woo %>% filter(cat %in% categories, Qty > qty_offline, !grepl("24", Seasons), SPU %in% hats_limit$SPU, !(SKU %in% hats_sales$SKU), !(SKU %in% hats_sales_old$SKU)) 
hats_sales_other <- rbind(hats_sales_old, hats_sales_limit) %>% mutate(Total.2024 = sales_SKU_thisyearT[toupper(SKU), "Total.2024"], Total.2024.JJ = sales_SKU_thisyearJJ[toupper(SKU), "Total.2024"], Total.2024.AMZ = sales_SKU_thisyearAMZ[toupper(SKU), "Total.2024"], time_yrs = Qty/Total.2024) %>% 
  select(SKU, Name, Seasons, Regular.price, Qty, time_yrs, Total.2024, Total.2024.JJ, Total.2024.AMZ)
write.csv(hats_sales_other, file = "../Analysis/hats_sales_other_2024-05-29.csv", row.names = F)

# ----------------- B2S / Grand opening --------------
## update Clover and Square price
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price)) %>% `row.names<-`(toupper(.[, "SKU"])) 
xoro <- read.xlsx2(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".xlsx"), full.names = T), sheetIndex = 1) %>% filter(Store == "Warehouse - JJ") %>% mutate(Item. = toupper(Item.), ATS = as.numeric(ATS)) %>% `row.names<-`(toupper(.[, "Item."])) 
# price <- woo %>% mutate(cat = gsub("-.*", "", SKU)) %>% group_by(cat) %>% summarise(Price = round(max(Regular.price) * 0.85, 2)) %>% as.data.frame()
# write.csv(price, file = "../Clover/B2S_opening_price.csv", row.names = F)
# manual update price
price <- read.csv("../Clover/B2S_opening_price.csv", as.is = T)
price <- rbind(price, data.frame(cat = c("MISC5", "MISC10", "MISC15", "MISC20"), Price = c(5, 10, 15, 20))) %>% `row.names<-`(toupper(.[, "cat"])) 
mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T)[1], sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
cat_sep <- c("WJA", "WJT", "WPF", "WPS")
size_s <- c("1T", "2T", "3T", "4T", "5T", "6Y")
special_price <- c("WSS-CNL", "WSS-DNL", "WSS-TRZ", "WSS-UNC")
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% filter(Name != "") %>% mutate(cat = gsub("-.*", "", Name), size = toupper(mastersku[Name, "Size"]), SPU = paste0(toupper(mastersku[Name, "Category.SKU"]), "-", toupper(mastersku[Name, "Print.SKU"])), cat = ifelse(cat %in% cat_sep, ifelse(size %in% size_s, paste0(cat, "_S"), paste0(cat, "_L")), cat)) %>% `row.names<-`(toupper(.[, "Name"])) %>% 
  mutate(Price = ifelse(cat %in% rownames(price), price[cat, "Price"], ifelse(woo[Name, "Sale.price"] < woo[Name, "Regular.price"] * 0.85, woo[Name, "Sale.price"], round(woo[Name, "Regular.price"] * 0.85, 2))), Price = ifelse(SPU %in% special_price, 50, Price), Price.Type = ifelse(is.na(Price), "Variable", "Fixed"), Alternate.Name = woo[Name, "Name"], Tax.Rates = ifelse(woo[Name, "Tax.class"] == "full", "GST+PST", "GST"), Tax.Rates = ifelse(is.na(Tax.Rates), "GST", Tax.Rates)) %>% select(-cat, -size, -SPU)
square <- data.frame(Token = "", ItemName = clover_item$Name, ItemType = "", VariationName = "Regular", SKU = clover_item$Name, Description = "", Category = "", Price = clover_item$Price, PriceRichmond = "", PriceSurrey = "", NewQuantityRichmond = xoro[toupper(clover_item$Name), "ATS"], NewQuantitySurrey = xoro[toupper(clover_item$Name), "ATS"], EnabledRichmond = "Y", CurrentQuantityRichmond = "", CurrentQuantitySurrey = "", StockAlertEnabledRichmond = "N", StockAlertCountRichmond = 0, EnabledSurrey = "Y", StockAlertEnabledSurrey = "N", StockAlertCountSurrey = 0, SEOTitle = "",	SEODescription = "",	Permalink = "", SquareOnlineItemVisibility = "UNAVAILABLE",	ShippingEnabled = "",	SelfserveOrderingEnabled = "",	DeliveryEnabled = "",	PickupEnabled = "", Sellable = "Y",	Stockable = "Y",	SkipDetailScreeninPOS = "",	OptionName1 = "",	OptionValue1 = "", TaxPST = ifelse(clover_item$Tax.Rates == "GST+PST", "Y", "N")) %>%
  mutate(NewQuantityRichmond = ifelse(is.na(NewQuantityRichmond), 0, NewQuantityRichmond), NewQuantitySurrey = ifelse(is.na(NewQuantitySurrey), 0, NewQuantitySurrey)) %>% filter(!is.na(Price))
write.table(square, file = paste0("../Square/square-upload-", Sys.Date(), ".csv"), sep = ",", row.names = F, col.names = c(colnames(square)[1:ncol(square)-1], "Tax - PST (7%)"), na = "")
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item) + 100, gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
## Receive stock
order <- read.csv("../Clover/order08142024.csv", as.is = T) %>% `row.names<-`(.[, "ItemNumber"])
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% mutate(Quantity = ifelse(Name %in% order$ItemNumber, Quantity + order[Name, "Qty"], Quantity))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item) + 100, gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
## Set barcode for MWPF MWPS
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% filter(!is.na(Name)) %>% `row.names<-`(.[, "Name"]) 
clover_item_WP <- clover_item %>% filter(grepl("^WP", Name)) 
clover_item <- clover_item %>% mutate(Product.Code = ifelse(grepl("^MWP", Name), clover_item_WP[gsub("^MWP", "WP", Name), "Product.Code"], Product.Code))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item) + 100, gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
## deduct Square sales stock
square_so <- read.csv(rownames(file.info(list.files(path = "../Square/", pattern = "items", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% 
  group_by(Item) %>% summarize(Qty = sum(Count)) %>% as.data.frame() %>% `row.names<-`(.[, "Item"]) 
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% mutate(Quantity = ifelse(Name %in% square_so$Item, Quantity - square_so[Name, "Qty"], Quantity))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item) + 100, gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)

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

# ---------- best sellers ----------
xoro <- read.xlsx2(list.files(path = "../xoro/", pattern = paste0("^Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".xlsx"), full.names = T), sheetIndex = 1) %>% filter(Store == "Warehouse - JJ") %>% mutate(Item. = toupper(Item.), ATS = as.numeric(ATS)) %>% `row.names<-`(.[, "Item."])
mastersku <- openxlsx::read.xlsx(list.files(path = "../FBArefill/Raw Data File/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
sales_SKU_thisyear <- data.frame()
sheets <- getSheetNames(RawData)
sheets_SKU <- sheets[grepl("-SKU$", sheets)&(!grepl("PW2", sheets))&(!grepl("WJO", sheets))&(!grepl("WPO", sheets))&(!grepl("ISJ", sheets))&(!grepl("WJS", sheets))]
for(sheet in sheets_SKU){
  sales_SKU_i <- openxlsx::read.xlsx(RawData, sheet = sheet, cols = c(1:8), startRow = 2) %>% 
    filter(!(Sales.Channel %in% c("Summary", "Sales Channel", "All Marketplaces"))) %>% rename_with(~ gsub("20", "Total.20", gsub("_x000D_.Total", "", .x)))
  sales_SKU_thisyear <- rbind(sales_SKU_thisyear, sales_SKU_i)
}
sales_SPU_thisyearT <- sales_SKU_thisyear %>% mutate(SPU = paste0(mastersku[Adjust_MSKU, "Category.SKU"], "-", mastersku[Adjust_MSKU, "Print.SKU"])) %>% group_by(SPU) %>% summarise(Total.2024 = sum(Total.2024)) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "SPU"]))
sales_SKU_thisyearT <- sales_SKU_thisyear %>% group_by(Adjust_MSKU) %>% summarise(Total.2024 = sum(Total.2024)) %>% as.data.frame() %>% mutate(Category = mastersku[Adjust_MSKU, "Category.SKU"], SPU = paste0(mastersku[Adjust_MSKU, "Category.SKU"], "-", mastersku[Adjust_MSKU, "Print.SKU"]), Total.2024.SPU = sales_SPU_thisyearT[SPU, "Total.2024"], ATS = xoro[Adjust_MSKU, "ATS"]) %>% `row.names<-`(toupper(.[, "Adjust_MSKU"]))
categories <- c("GBX", "GHA", "GUA", "GUX")
TopSeller_sunglasses <- sales_SKU_thisyearT %>% filter(Category %in% categories) %>% arrange(-Total.2024.SPU)
write.csv(TopSeller_sunglasses, file = paste0("../Analysis/TopSeller_sunglasses_", Sys.Date(), ".csv"), row.names = F)
categories <- c("SPW", "SKB", "SKG", "SWS")
TopSeller_shoes <- sales_SKU_thisyearT %>% filter(Category %in% categories) %>% arrange(Category, -Total.2024.SPU)
write.csv(TopSeller_shoes, file = paste0("../Analysis/TopSeller_shoes_", Sys.Date(), ".csv"), row.names = F)
categories <- c("HAD0", "HAV0", "HBS", "HCA0", "HCB0", "HCF0", "HJP", "HJS", "HXC", "HXP", "HXU")
TopSeller_hats <- sales_SKU_thisyearT %>% filter(Category %in% categories) %>% arrange(-Total.2024.SPU)
write.csv(TopSeller_hats, file = paste0("../Analysis/TopSeller_hats_", Sys.Date(), ".csv"), row.names = F)

# ----------- Always on sale -------------
xoro <- read.xlsx2(list.files(path = "../xoro/", pattern = paste0("^Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".xlsx"), full.names = T), sheetIndex = 1) %>% filter(Store == "Warehouse - JJ") %>% mutate(Item. = toupper(Item.), ATS = as.numeric(ATS)) %>% `row.names<-`(.[, "Item."])
woo_old <- read.csv("../woo/wc-product-export-17-6-2024-1718640775225.csv", as.is = T) %>% filter(SKU != "") %>% `row.names<-`(.[, "SKU"])
clearance <- read.csv("../Analysis/Retail JJ Deals set to end on 2030.csv", as.is = T) %>% 
  mutate(old.price = woo_old[SKU, "Sale.price"], qty = ifelse(SKU %in% xoro$Item., xoro[SKU, "ATS"], 0), Sale.price = ifelse(is.na(Sale.price), old.price, Sale.price), discount = ifelse(is.na(Sale.price), 0, round((Regular.price - Sale.price)/Regular.price, 2)))
write.csv(clearance, file = "../Analysis/Retail JJ Deals set to end on 2030.csv", row.names = F, na = "")

# ----------- Boxing Day sale 2024 -------------
# IHT arrived late and will use last year's (Dec, Jan, Feb) sales for prediction
categories <- c("KEH", "KMT", "IHT", "ISS", "IPS", "ISB", "ISJ", "ICP", "IPC", "WMT", "WGS", "BST", "XBM", "XBK", "XLB", "XPC", "SWS", "SKX", "SKG", "UG1", "UJ1", "UV2", "UT1", "USA", "HXU", "HXP", "HXC", "HAD0", "HAV0", "HBS", "HCA0", "HCB0", "HCF0", "HJP", "HJS")
netsuite_item_S <- read.csv(list.files(path = "../FBArefill/Raw Data File/", pattern = "Items_S_.*", full.names = T), as.is = T) %>%
  mutate(Name = toupper(Name), Seasons = ifelse(Name %in% mastersku$MSKU, mastersku[Name, "Seasons.SKU"], mastersku_adjust[Name, "Seasons.SKU"]), SPU = gsub("(\\w+-\\w+)-.*", "\\1", Name)) %>% `row.names<-`(.[, "Name"])
inventory <- read.csv(rownames(file.info(list.files(path = "../Analysis/", pattern = "Sales_Inventory_SKU_", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T)
sales_SKU_last12month <- read.csv(rownames(file.info(list.files(path = "../Analysis/", pattern = "Sales_SKU_last12month_", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T)
monR <- read.csv(rownames(file.info(list.files(path = "../Analysis/", pattern = "MonthlyRatio_", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% `row.names<-`(.[, "Category"])
if(month == "01"){last3m <- c("Month11", "Month12", "Month01")}else if(month %in% c("02", "03")){last3m <- sprintf("Month%02d", c(1:(as.numeric(month)-1), (as.numeric(month)-3+12):12))}else{last3m <- sprintf("Month%02d", c((as.numeric(month)-3):(as.numeric(month)-1)))}
monR <- monR %>% mutate(T.last3m = ifelse(Category == "IHT", rowSums(across(all_of(c("Month12", "Month01", "Month02")))), rowSums(across(all_of(last3m))))) %>% `row.names<-`(.[, "Category"])
inventory_SPU <- inventory %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU)) %>% group_by(SPU) %>% summarise(qty_SPU = sum(Inv_Total_End.WK, na.rm = T), qty_WH = sum(Inv_WH.JJ, na.rm = T), qty_AMZ = qty_SPU - qty_WH) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
netsuite_item_S_SPU <- netsuite_item_S %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Name)) %>% group_by(SPU) %>% summarise(qty_WH = sum(Warehouse.Available, na.rm = T)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SKU_last12month <- sales_SKU_last12month %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU), T.last3m = ifelse(grepl("IHT", SPU), rowSums(across(all_of(c("Month12", "Month01", "Month02")))), rowSums(across(all_of(last3m)))), T.yr = rowSums(across(Month01:Month12)))
sales_SPU_last3month <- sales_SKU_last12month %>% group_by(SPU) %>% summarise(T.last3m = sum(T.last3m)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SPU_last3month_JJ <- sales_SKU_last12month %>% filter(Sales.Channel == "janandjul.com") %>% group_by(SPU) %>% summarise(T.last3m = sum(T.last3m)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SPU_last12month_AMZ <- sales_SKU_last12month %>% filter(grepl("Amazon", Sales.Channel)) %>% group_by(SPU) %>% summarise(T.last12m = sum(T.yr)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sizes_all <- netsuite_item_S %>% mutate(SPU_size = gsub("(\\w+-\\w+-\\w+).*", "\\1", Name)) %>% filter(!duplicated(SPU_size)) %>% count(SPU) %>% `row.names<-`(.[, "SPU"])
size_limited <- netsuite_item_S %>% filter(Warehouse.Available > qty_offline) %>% count(SPU) %>% mutate(all = sizes_all[SPU, "n"], percent = ifelse(all < n, 0, (all - n)/all)) %>% `row.names<-`(.[, "SPU"])
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% filter(Published == 1, !is.na(Regular.price)) %>% filter(!duplicated(SKU), SKU != "") %>% 
  mutate(SKU = toupper(SKU), Qty = netsuite_item_S[SKU, "Warehouse.Available"], Qty = ifelse(is.na(Qty), 0, Qty), Seasons = ifelse(SKU %in% mastersku$MSKU, mastersku[SKU, "Seasons.SKU"], mastersku_adjust[SKU, "Seasons.SKU"]), discount = ifelse(is.na(Sale.price), 0, (Regular.price - Sale.price)/Regular.price), cat = mastersku[SKU, "Category.SKU"], SPU = gsub("(\\w+-\\w+)-.*", "\\1", SKU)) %>% 
  select(ID, SKU, Name, Seasons, cat, SPU, Regular.price, discount, Qty) %>% `row.names<-`(.[, "SKU"])
boxing_day <- woo %>% filter(Qty > qty_offline, Regular.price > 0, cat %in% categories) %>% mutate(Qty_SPU = inventory_SPU[SPU, "qty_SPU"], Qty_WH = netsuite_item_S_SPU[SPU, "qty_WH"], Qty_AMZ = inventory_SPU[SPU, "qty_AMZ"], MonR_last3m = monR[cat, "T.last3m"], Sales_SPU_last3m = sales_SPU_last3month[SPU, "T.last3m"], Sales_SPU_last3m_JJ = sales_SPU_last3month_JJ[SPU, "T.last3m"], Sales_SPU_1yr = Sales_SPU_last3m/MonR_last3m, Sales_SPU_1yr_JJ = Sales_SPU_last3m_JJ/MonR_last3m, Sales_SPU_1yr_AMZ = sales_SPU_last12month_AMZ[SPU, "T.last12m"]) %>% 
  filter(!is.na(Qty_SPU)) %>% mutate(time_yrs = round(ifelse(Qty_AMZ >= Sales_SPU_1yr_AMZ, Qty_WH/Sales_SPU_1yr_JJ, Qty_SPU/Sales_SPU_1yr), 2), time_yrs = ifelse(is.na(time_yrs), 0, time_yrs), size_percent_missing = gsub("Sizes.missing.0%", "Sizes.missing.<.50%", paste0("Sizes.missing.", as.character(as.integer(ifelse(size_limited[SPU, "percent"] < 0.5, 0, size_limited[SPU, "percent"])*10)*10), "%"))) %>% 
  select(SKU:Qty, time_yrs, size_percent_missing) %>% mutate(Suggest_discount = 0, Suggest_discount = ifelse(grepl("25F", Seasons), ifelse(time_yrs > 4, 0.2, 0.1), Suggest_discount), Suggest_discount = ifelse(cat == "IHT", 0.2, Suggest_discount)) %>% arrange(Suggest_discount, -time_yrs)
boxing_day[1:54, "Suggest_discount"] <- 0.4
boxing_day[55:101, "Suggest_discount"] <- 0.3
boxing_day[102:250, "Suggest_discount"] <- 0.2
boxing_day[251:500, "Suggest_discount"] <- 0.1
boxing_day <- boxing_day %>% mutate(Suggest_discount = ifelse(cat == "HXC", Suggest_discount + 0.1, Suggest_discount), Suggest_discount = ifelse(Suggest_discount > discount, Suggest_discount, discount), Sale_price = round(Regular.price * (1 - Suggest_discount), 2)) %>% arrange(cat, -time_yrs)
write.csv(boxing_day, file = paste0("../Analysis/Boxing_day_deals1_", Sys.Date(), ".csv"), row.names = F)

# ----------- Wholesale performance analysis 2023 2024 -------------
library(reshape2)
library(ggh4x)
wholesale <- openxlsx::loadWorkbook(list.files(path = "../Wholesale/", pattern = "Wholesale_Invoices", full.names = T))
wholesale_Xoro <- openxlsx::readWorkbook(wholesale, 2) %>% mutate(REF.NO = SO.Ref.Number, Date = as.Date(Date, origin = "1899-12-30"), Year = format(Date, "%Y"), Month = format(Date, "%m"), Currency = Currency.Code, Amount.Discount = Total.Discount, Tax = Total.Tax.Amount, Country = Ship.to.Country, State = Ship.To.State, Email = Main.Email) %>%
  select(REF.NO, Date, Year, Month, Sales.Rep, Currency, Total.Amount, Shipping.Cost, Amount.Discount, Tax, Total.Qty, Country, State, Customer.Name, Email)
Customers_Xoro <- wholesale_Xoro %>% distinct(Email, .keep_all = T) %>% `row.names<-`(.[, "Email"])
wholesale_NS <- openxlsx::readWorkbook(wholesale, 1) %>% mutate(Date = as.Date(Date, origin = "1899-12-30"), Year = format(Date, "%Y"), Month = format(Date, "%m"), Total.Amount = Amount.Foreign.Currency, Amount.Discount = ifelse(is.na(Amount.Discount), 0, Amount.Discount), Tax = Amount.Transaction.Tax.Total, Total.Qty = TOTAL_QTY_R, Country = Shipping.Country.Code, State = Shipping.State.Province, Customer.Name = ifelse(Email %in% Customers_Xoro$Email, Customers_Xoro[Email, "Customer.Name"], Company.Name)) %>% 
  select(REF.NO, Date, Year, Month, Sales.Rep, Currency, Total.Amount, Shipping.Cost, Amount.Discount, Tax, Total.Qty, Country, State, Customer.Name, Email)
wholesale_Faire <- openxlsx::readWorkbook(wholesale, 3) %>% group_by(Order.Number) %>% mutate(REF.NO = Order.Number, Date = as.Date(Order.Date, origin = "1899-12-30"), Year = format(Date, "%Y"), Month = format(Date, "%m"), Sales.Rep = "Faire", Currency = "USD", Total.Amount = sum(Wholesale.Price), Shipping.Cost = 0, Amount.Discount = 0, Tax = 0, Total.Qty = sum(Quantity), Country = gsub("United States", "US-Faire", gsub("Canada", "CA", Country)), State = State, Customer.Name = Retailer.Name, Email = "") %>% ungroup() %>%
  select(REF.NO, Date, Year, Month, Sales.Rep, Currency, Total.Amount, Shipping.Cost, Amount.Discount, Tax, Total.Qty, Country, State, Customer.Name, Email) %>% distinct(REF.NO, .keep_all = T) %>% filter(Year %in% c(2023, 2024))
wholesale_2023_2024 <- rbind(wholesale_Faire, wholesale_Xoro, wholesale_NS, data.frame(REF.NO = "", Date = "", Year = c(2023, 2024), Month = "", Sales.Rep = "Taiwan", Currency = "USD", Total.Amount = c(580179.66, 366010.45), Shipping.Cost = c(2691, 1350), Amount.Discount = c(144259.5, 66224.11), Tax = c(0, 0), Total.Qty = c(37798, 21676), Country = "TW", State = "TW", Customer.Name = "Taobaby", Email = "taobabyjp@gmail.com"))
wholesale_2023_2024_Country <- wholesale_2023_2024 %>% group_by(Country, Year) %>% summarise(Average.Amount = ifelse(Country == "TW", 0, mean(Total.Amount, na.rm = T)), Total.Amount = sum(Total.Amount), Total.Qty = sum(Total.Qty)) %>% ungroup() %>%
  melt(id.vars = c("Year","Country"), variable.name = "Total") %>% filter(Country %in% c("AU", "CA", "CZ", "FR", "TW", "US", "US-Faire")) 
(wholesale_2023_2024_Country_plot <- ggplot(wholesale_2023_2024_Country, aes(Country, value)) + 
    geom_col(aes(fill = Year), position = position_dodge()) + 
    scale_y_continuous(labels = scales::comma) + 
    facet_grid(rows = vars(Total), scales = "free_y") + 
    ggtitle("Wholesale annual amount and quantity in different countries") + xlab("") + ylab("") + theme_bw() + theme(plot.title = element_text(hjust = 0.5)))
wholesale_2023_2024_SalesRep <- wholesale_2023_2024 %>% group_by(Sales.Rep, Year) %>% summarise(Average.Amount = ifelse(Sales.Rep %in% c("Taiwan", "Uniform"), 0, mean(Total.Amount, na.rm = T)), Total.Amount = sum(Total.Amount), Total.Qty = sum(Total.Qty)) %>% ungroup() %>%
  melt(id.vars = c("Year","Sales.Rep"), variable.name = "Total")
(wholesale_2023_2024_SalesRep_plot <- ggplot(wholesale_2023_2024_SalesRep, aes(Sales.Rep, value)) + 
    geom_col(aes(fill = Year), position = position_dodge()) + 
    scale_y_continuous(labels = scales::comma) + 
    facet_grid(rows = vars(Total), scales = "free_y") + 
    ggtitle("Wholesale annual amount and quantity for different Sales Reps") + xlab("") + ylab("") + 
    theme_bw() + theme(plot.title = element_text(hjust = 0.5), axis.text.x = element_text(angle = 15, vjust = 1, hjust=1)))
wholesale_2023_2024_State <- wholesale_2023_2024 %>% group_by(Country, State, Year) %>% summarise(Average.Amount = ifelse(Sales.Rep %in% c("Taiwan", "Uniform"), 0, mean(Total.Amount, na.rm = T)), Total.Amount = sum(Total.Amount), Total.Qty = sum(Total.Qty)) %>% ungroup() %>%
  melt(id.vars = c("Year","Country", "State"), variable.name = "Total") %>% filter(Country %in% c("CA", "US"), !is.na(State), State != "-")
(wholesale_2023_2024_State_plot <- ggplot(wholesale_2023_2024_State, aes(State, value)) + 
    geom_col(aes(fill = Year), position = position_dodge()) + 
    scale_y_continuous(labels = scales::comma) + 
    facet_grid2(rows = vars(Total), cols = vars(Country), scales = "free", space = "free_x", independent = "y") + 
    ggtitle("Wholesale annual amount and quantity in different States") + xlab("") + ylab("") + theme_bw() + 
    theme(plot.title = element_text(hjust = 0.5), axis.text.x = element_text(angle = 90, hjust = 0.5)))
pdf(file = "../Wholesale/Wholesale_summary_2023_2024.pdf", width = 13, height = 8)
grid.arrange(wholesale_2023_2024_SalesRep_plot, wholesale_2023_2024_Country_plot, ncol=2)
(wholesale_2023_2024_State_plot)
dev.off()
wholesale_2023_2024_Store <- wholesale_2023_2024 %>% group_by(Customer.Name) %>% mutate(Start.Date = min(Date)) %>%
  group_by(Customer.Name, Year) %>% summarise(Email = Email[1], Sales.Rep = Sales.Rep[1], Start.Date = Start.Date[1], Country = Country[1], State = State[1], Average.Amount = round(mean(Total.Amount, na.rm = T), 2), Total.Amount = sum(Total.Amount), Total.Qty = sum(Total.Qty)) %>%
  tidyr::pivot_wider(names_from  = c(Year), values_from = c('Total.Amount', 'Average.Amount', 'Total.Qty')) %>% arrange(Sales.Rep, desc(Total.Amount_2024)) %>% mutate(Growth.Rate = ifelse(is.na(Total.Amount_2023), 1, ifelse(is.na(Total.Amount_2024), -1, round((Total.Amount_2024 - Total.Amount_2023)/Total.Amount_2023, 3))))
write.csv(wholesale_2023_2024_Store, file = "../Wholesale/wholesale_2023_2024_Store.csv", row.names = F, na = "", quote = F)

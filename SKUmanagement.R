# discontinued SKUs management
library(dplyr)
library(openxlsx)
library(scales)
new_season <- "24"   # New season contains
in_season <- "F"     # "F": Sept - Feb; "S": Mar - Aug
qty_offline <- 3     # Qty to move to offline sales
qty_sold <- 50       # Sales rate low: Qty sold on website last month
size_percent <- 0.5  # % of sizes not available 

mastersku <- read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
xoro <- read.csv(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"), full.names = T), as.is = T) %>% filter(Store == "Warehouse - JJ") %>% 
  mutate(seasons = mastersku[Item., "Seasons"], cat_print = gsub("(\\w+-\\w+)-.*", "\\1", Item.))
rownames(xoro) <- xoro$Item.

# ------------- move offline ---------------------------
offline <- xoro %>% filter(!grepl(new_season, seasons) & ATS <= qty_offline & ATS > 0) %>% arrange(Item.)
write.csv(offline, file = paste0("../SKUmanagement/offline_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\

# ------------- to inactivate --------------------------
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

# ------------- move to website deals page -------------
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
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

# ------------- FW clearance -------------
sheets <- getSheetNames("../FBArefill/Raw Data File/Archive/Historic sales and inv. data for all cats v31 (20240104).xlsx")
sheets_cat <- sheets[!grepl("-", sheets)&!grepl("Sale", sheets)]
sales_cat <- data.frame()
monR_cat <- data.frame()
for(sheet in sheets_cat){
  sales_cat_i <- openxlsx::read.xlsx("../FBArefill/Raw Data File/Archive/Historic sales and inv. data for all cats v31 (20240104).xlsx", sheet = sheet, cols = c(1:19), startRow = 2) %>% 
    filter(Sales.Channel == "All Channels", Year == "2023")
  if(nrow(sales_cat_i) == 0){
    sales_cat_i <- openxlsx::read.xlsx("../FBArefill/Raw Data File/Archive/Historic sales and inv. data for all cats v31 (20240104).xlsx", sheet = sheet, cols = c(1:19), startRow = 2) %>% 
      filter(Sales.Channel == "janandjul.com", Year == "2023")
  }
  sales_cat <- rbind(sales_cat, sales_cat_i)
  monR_cat_i <- openxlsx::read.xlsx("../FBArefill/Raw Data File/Archive/Historic sales and inv. data for all cats v31 (20240104).xlsx", sheet = sheet, cols = c(1:6, 20:32), startRow = 2) %>% 
    filter(Sales.Channel == "All Channels", Year == "2023")
  if(nrow(monR_cat_i) == 0){
    monR_cat_i <- openxlsx::read.xlsx("../FBArefill/Raw Data File/Archive/Historic sales and inv. data for all cats v31 (20240104).xlsx", sheet = sheet, cols = c(1:6, 20:32), startRow = 2) %>% 
      filter(Sales.Channel == "janandjul.com", Year == "2023")
  }
  monR_cat <- rbind(monR_cat, monR_cat_i)
}
write.csv(sales_cat, file = paste0("../SKUmanagement/SalesAll2023_category_", Sys.Date(), ".csv"), row.names = F, na = "")
write.csv(monR_cat, file = paste0("../SKUmanagement/MonthlyRatio2023_category_", Sys.Date(), ".csv"), row.names = F, na = "")

sheets_SKU <- sheets[grepl("-SKU$", sheets)]
sales_SKU <- data.frame()
monR_SKU <- data.frame()
for(sheet in sheets_SKU){
  sales_SKU_i <- openxlsx::read.xlsx("../FBArefill/Raw Data File/Archive/Historic sales and inv. data for all cats v31 (20240104).xlsx", sheet = sheet, cols = c(1:11, 26:37), startRow = 2) %>% 
    filter(Sales.Channel == "janandjul.com") %>% mutate_all(~ replace(., is.na(.), 0)) %>% rename_with(~ gsub("_x000D_", "", .x))
  sales_SKU <- rbind(sales_SKU, sales_SKU_i)
  monR_SKU_i <- openxlsx::read.xlsx("../FBArefill/Raw Data File/Archive/Historic sales and inv. data for all SKUs v31 (20240104).xlsx", sheet = sheet, cols = c(1:6, 20:32), startRow = 2) %>% 
    filter(Sales.Channel == "All Channels", Year == "2023")
  if(nrow(monR_SKU_i) == 0){
    monR_SKU_i <- openxlsx::read.xlsx("../FBArefill/Raw Data File/Archive/Historic sales and inv. data for all SKUs v31 (20240104).xlsx", sheet = sheet, cols = c(1:6, 20:32), startRow = 2) %>% 
      filter(Sales.Channel == "janandjul.com", Year == "2023")
  }
  monR_SKU <- rbind(monR_SKU, monR_SKU_i)
}
write.csv(sales_SKU, file = paste0("../SKUmanagement/SalesAll2023_SKUegory_", Sys.Date(), ".csv"), row.names = F, na = "")
write.csv(monR_SKU, file = paste0("../SKUmanagement/MonthlyRatio2023_SKUegory_", Sys.Date(), ".csv"), row.names = F, na = "")

mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T)[1], sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
mastersku_adjust <- mastersku %>% filter(!duplicated(Adjust.SKU))
rownames(mastersku_adjust) <- mastersku_adjust$Adjust.SKU
xoro <- read.xlsx2(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".xlsx"), full.names = T), sheetIndex = 1) %>% filter(Store == "Warehouse - JJ") %>% mutate(Item. = toupper(Item.), ATS = as.numeric(ATS), cat_print = gsub("(\\w+-\\w+)-.*", "\\1", Item.))
rownames(xoro) <- xoro$Item.
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% 
  mutate(Qty = ifelse(SKU %in% xoro$Item., xoro[SKU, "ATS"], as.numeric(Stock)), Seasons = ifelse(SKU %in% mastersku$MSKU, mastersku[SKU, "Seasons.SKU"], mastersku_adjust[SKU, "Seasons.SKU"]), Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price), discount = (Regular.price - Sale.price)/Regular.price) %>% 
  select(ID, SKU, Name, Sale.price, Regular.price, Qty, Seasons, discount)
rownames(woo) <- woo$SKU
woo_FW <- woo %>% filter(grepl("F", Seasons), Qty > 0)

sizes_all <- xoro %>% mutate(cat_print_size = gsub("(\\w+-\\w+-\\w+).*", "\\1", Item.)) %>% filter(!duplicated(cat_print_size)) %>% count(cat_print)
rownames(sizes_all) <- sizes_all$cat_print
size_limited <- xoro %>% filter(!grepl(new_season, seasons) & ATS > qty_offline) %>% count(cat_print) %>% mutate(all = sizes_all[cat_print, "n"], percent = (all - n)/all) %>% filter(percent >= size_percent & percent < 1)
rownames(size_limited) <- size_limited$cat_print


# discontinued SKUs management
new_season <- "24"   # New season contains
qty_offline <- 3     # Qty to move to offline sales
# qty_jj <- 12         # Qty to stop sending to Amazon, i.e. JJ website sales only
qty_sold <- 5        # Sales rate low: Qty sold on website last month
size_percent <- 0.5  # % of sizes not available 

library(dplyr)
library(openxlsx)
library(scales)
mastersku <- read.xlsx("../SKUmanagement/1-MasterSKU-All-Product-2023-12-01.xlsx", sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
woo <- read.csv(paste0("../woo/", list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))))), as.is = T)
woo <- woo %>% filter(!is.na(woo$Regular.price) & !duplicated(woo$SKU) & !(woo$SKU == "")) 
sales <- woo %>% filter(!is.na(woo$Sale.price)) %>% mutate(discount = (Regular.price-Sale.price)/Regular.price)
rownames(sales) <- sales$SKU
clover <- read.xlsx(paste0("../Clover/", list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"))), sheet = "Items", )
rownames(clover) <- clover$SKU
xoro_lastmonth <- read.csv("../xoro/Item Inventory Snapshot_11032023.csv", as.is = T) %>% filter(Store == "Warehouse - JJ") 
rownames(xoro_lastmonth) <- xoro_lastmonth$Item.
xoro <- read.csv(paste0("../xoro/", list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"))), as.is = T) %>% filter(Store == "Warehouse - JJ") %>% 
  mutate(cat_print = gsub("(\\w+-\\w+)-.*", "\\1", Item.), status = mastersku[Item., "MSKU.Status"], seasons = mastersku[Item., "Seasons"], lastmonth = xoro_lastmonth[Item., "ATS"], clover = clover[Item., "Quantity"])
rownames(xoro) <- xoro$Item.

# ------------- move offline -------------
offline0 <- xoro %>% filter(!grepl(new_season, seasons) & status == "Active" & ATS == 0 & clover == 0) 
offline <- xoro %>% filter(!grepl(new_season, seasons) & ATS <= qty_offline & ATS > 0) %>% arrange(Item.)
write.csv(offline0, file = paste0("../SKUmanagement/offline_qty0_", Sys.Date(), ".csv"), row.names = F, na = "")
write.csv(offline, file = paste0("../SKUmanagement/offline_", Sys.Date(), ".csv"), row.names = F, na = "")

# ------------- move to website deals page -------------
sizes_all <- xoro %>% filter(status == "Active") %>% count(cat_print)
rownames(sizes_all) <- sizes_all$cat_print
size0 <- xoro %>% filter(!grepl(new_season, seasons) & status == "Active" & ATS == 0) %>% count(cat_print) %>% mutate(all = sizes_all[cat_print, "n"], percent = n/all) %>% filter(percent >= size_percent & percent < 1)
rownames(size0) <- size0$cat_print
deals <- xoro %>% filter(!grepl(new_season, seasons) & status == "Active" & ATS > qty_offline & lastmonth - ATS < qty_sold & lastmonth - ATS > 0 & cat_print %in% size0$cat_print) %>% mutate(size_percent = size0[cat_print, "percent"]) %>% arrange(Item.)
inactive <- xoro %>% filter(status != "Active" & ATS > qty_offline)
inactive_sizes_all <- xoro %>% filter(cat_print %in% inactive$cat_print) %>% count(cat_print)
rownames(inactive_sizes_all) <- inactive_sizes_all$cat_print
inactive_size0 <- xoro %>% filter(cat_print %in% inactive$cat_print & ATS == 0) %>% count(cat_print) %>% mutate(all = inactive_sizes_all[cat_print, "n"], percent = n/all) %>% filter(percent >= size_percent & percent < 1)
rownames(inactive_size0) <- inactive_size0$cat_print
inactive <- inactive %>% filter(cat_print %in% inactive_size0$cat_print) %>% mutate(size_percent = inactive_size0[cat_print, "percent"])
deals <- rbind(deals, inactive) %>% mutate(size_percent = percent(size_percent), discount = percent(ifelse(Item. %in% sales$SKU, sales[Item., "discount"], 0), accuracy = 0.1)) %>% arrange(Item.)
write.csv(deals, file = paste0("../SKUmanagement/deals_", Sys.Date(), ".csv"), row.names = F, na = "")

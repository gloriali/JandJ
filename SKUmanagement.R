# discontinued SKUs management
library(dplyr)
library(openxlsx)
library(scales)
new_season <- "24"   # New season contains
in_season <- "F"     # "F": Sept - Feb; "S": Mar - Aug
qty_offline <- 3     # Qty to move to offline sales
qty_sold <- 50       # Sales rate low: Qty sold on website last month
size_percent <- 0.5  # % of sizes not available 

mastersku <- read.xlsx(paste0("../../TWK 2020 share/", list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-")), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
xoro <- read.csv(paste0("../xoro/", list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"))), as.is = T) %>% filter(Store == "Warehouse - JJ") %>% 
  mutate(seasons = mastersku[Item., "Seasons"], cat_print = gsub("(\\w+-\\w+)-.*", "\\1", Item.))
rownames(xoro) <- xoro$Item.

# ------------- move offline ---------------------------
offline <- xoro %>% filter(!grepl(new_season, seasons) & ATS <= qty_offline & ATS > 0) %>% arrange(Item.)
write.csv(offline, file = paste0("../SKUmanagement/offline_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\

# ------------- move to website deals page -------------
woo <- read.csv(paste0("../woo/", list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))))), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") 
sales <- woo %>% filter(!is.na(Sale.price)) %>% mutate(discount = (Regular.price-Sale.price)/Regular.price)
rownames(sales) <- sales$SKU
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
  mutate(size_percent = percent(size_limited[cat_print, "percent"]), discount = percent(ifelse(Item. %in% sales$SKU, sales[Item., "discount"], 0), accuracy = 0.1)) %>% arrange(Item.)
write.csv(deals, file = paste0("../SKUmanagement/deals_", Sys.Date(), ".csv"), row.names = F, na = "")
# email Joren and Kamer

# ------------- to inactivate --------------------------
clover <- read.xlsx(paste0("../Clover/", list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"))), sheet = "Items")
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

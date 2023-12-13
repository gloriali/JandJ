# FBA refill
library(dplyr)
library(openxlsx)

mastersku <- read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
xoro <- read.csv(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"), full.names = T), as.is = T) %>% filter(Store == "Warehouse - JJ")
rownames(xoro) <- xoro$Item.
config_US <- read.xlsx("../../TWK Analysis/0 - Analysis to Share - Sales and Inventory/0 - performance analysis/Category focus/2023 category focus.xlsx", sheet = "CONFIG", startRow = 2, cols = c(1:6))
config_CA <- read.xlsx("../../TWK Analysis/0 - Analysis to Share - Sales and Inventory/0 - performance analysis/Category focus/2023 category focus.xlsx", sheet = "CONFIG", startRow = 2, cols = c(1:2, 7:10))
month_ratio_US <- read.xlsx("../FBArefill/Historic sales and inv. data for all cats v29 (20231205).xlsx", sheet = "MonSaleR", startRow = 2, cols = c(1:13)) %>% filter(!is.na(Category)) 
month_ratio_CA <- read.xlsx("../FBArefill/Historic sales and inv. data for all cats v29 (20231205).xlsx", sheet = "MonSaleR", startRow = 2, cols = c(1, 14:25)) %>% filter(!is.na(Category)) 

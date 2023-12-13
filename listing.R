# Listing - create and manage listing files
library(dplyr)
library(openxlsx)
library(zoo)
library(stringr)
library(tidyr)

## Little Red Book Listing
mastersku <- read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
xoro <- read.csv(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"), full.names = T), as.is = T) %>% filter(Store == "Warehouse - JJ")
rownames(xoro) <- xoro$Item.
woo <- read.csv(list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price))
rownames(woo) <- woo$SKU
dimensions <- read.xlsx("../../TWK Listing/Product&DimsLog.xlsx", sheet = "DimensionLog", fillMergedCells = T) %>% mutate(Category = gsub("\\(.*", "", gsub("（.*", "", Category)), Category = ifelse(Category == "", NA, Category), Category = na.locf(Category), Category = strsplit(Category, "/"), Size = strsplit(Size, "/")) %>% 
  unnest(Category) %>% unnest(Size) %>% as.data.frame() %>% mutate(Category = str_trim(Category), Size = toupper(str_trim(Size)))
rownames(dimensions) <- paste(dimensions$Category, dimensions$Size, sep = "_")
products_XHS <- read.xlsx(list.files(path = "../Listing/LittleRedBook/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), sheet = 1)

### -------- update inventory with xoro ----------------
products_XHS$Inventory <- ifelse(xoro[products_XHS$SKU, "ATS"] < 20, 0, xoro[products_XHS$SKU, "ATS"])
colnames(products_XHS) <- gsub("\\.", " ", colnames(products_XHS))
write.xlsx(products_XHS, file = paste0("../Listing/LittleRedBook/products_upload-", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"))

### -------- XHS Existing descriptions -----------------
products_description <- products_XHS %>% filter(!duplicated(SPU)) %>% mutate(SPU = toupper(SPU), Description = str_trim(paste(Description, SEO.Description, Describe))) %>% select(SPU, Product.Name, Description, Vendor, Categories, Tags, Option1.Name, Image.Src)
write.xlsx(products_description, file = "../Listing/LittleRedBook/products_description.xlsx")

### -------- XHS Create listing ------------------------
dimensions <- read.xlsx("../Listing/LittleRedBook/products_description.xlsx", sheet = 1, fillMergedCells = T) 


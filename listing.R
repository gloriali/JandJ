# Listing - create and manage listing files
library(dplyr)
library(openxlsx)
library(zoo)
library(stringr)
library(tidyr)

## Little Red book Listing
mastersku <- read.xlsx(paste0("../../TWK 2020 share/", list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-")), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
xoro <- read.csv(paste0("../xoro/", list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"))), as.is = T) %>% filter(Store == "Warehouse - JJ")
rownames(xoro) <- xoro$Item.
woo <- read.csv(paste0("../woo/", list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))))), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price))
rownames(woo) <- woo$SKU
dimensions <- read.xlsx("../../TWK Listing/Product&DimsLog.xlsx", sheet = "DimensionLog", fillMergedCells = T) %>% mutate(Category = gsub("\\(.*", "", gsub("（.*", "", Category)), Category = ifelse(Category == "", NA, Category), Category = na.locf(Category), Category = strsplit(Category, "/"), Size = strsplit(Size, "/")) %>% 
  unnest(Category) %>% unnest(Size) %>% as.data.frame() %>% mutate(Category = str_trim(Category), Size = toupper(str_trim(Size)))
rownames(dimensions) <- paste(dimensions$Category, dimensions$Size, sep = "_")

## -------- Existing descriptions -----------------
products_XHS <- read.xlsx(paste0("../Listing/LittleRedBook/", list.files(path = "../Listing/LittleRedBook/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"))), sheet = 1)
products_description <- products_XHS %>% filter(!duplicated(SPU)) %>% mutate(SPU = toupper(SPU), Description = str_trim(paste(Description, SEO.Description, Describe))) %>% select(SPU, Product.Name, Description, Vendor, Categories, Tags, Option1.Name, Image.Src)
write.xlsx(products_description, file = paste0("../Listing/LittleRedBook/products_description_", Sys.Date(), ".xlsx"))

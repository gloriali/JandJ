# Listing - create and manage listing files
library(dplyr)
library(openxlsx)
library(zoo)
library(stringr)
library(tidyr)

## Little Red Book Listing
xoro <- read.csv(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"), full.names = T), as.is = T) %>% filter(Store == "Warehouse - JJ") %>% mutate(Item. = toupper(Item.))
rownames(xoro) <- xoro$Item.
woo <- read.csv(list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(SKU = toupper(SKU), Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price))
rownames(woo) <- woo$SKU
# !AllValue > Products> Export all product - !!!open in excel > Review > Unprotect sheet > save!!!
products_XHS <- read.xlsx(list.files(path = "../Listing/XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), sheet = 1)

### -------- update inventory with xoro ----------------
products_upload <- products_XHS %>% mutate(Inventory = ifelse(xoro[SKU, "ATS"] < 20, 0, xoro[SKU, "ATS"]), Price = woo[SKU, "Sale.price"], Compare.At.Price = woo[SKU, "Regular.price"], Tags = ifelse(Price == Compare.At.Price, "正价", "特价"))
colnames(products_upload) <- gsub("\\.", " ", colnames(products_upload))
write.xlsx(products_upload, file = paste0("../Listing/XHS/products_upload-", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"))

### -------- XHS Existing descriptions -----------------
products_description <- products_XHS %>% filter(!duplicated(SPU)) %>% mutate(SPU = toupper(SPU), cat = gsub("-.*", "", SKU), Description = paste(Description, SEO.Description, Describe), Description = str_trim(gsub("NA", "", Description))) %>% select(SPU, Product.Name, cat, Description, Vendor, Categories, Tags, Option1.Name, Image.Src)
products_description_cat <- products_description %>% select(cat, Product.Name, Description, Vendor, Categories, Option1.Name) %>% filter(!duplicated(cat)) %>% mutate(Product.Name = gsub("\\s?-\\s?\\w+$", "", Product.Name))
write.xlsx(products_description, file = "../Listing/XHS/products_description.xlsx")
write.xlsx(products_description_cat, file = "../Listing/XHS/products_description_categories.xlsx")

### -------- XHS Create listing ------------------------
mastersku <- read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) 
rownames(mastersku) <- mastersku$MSKU
dimensions <- read.xlsx("../../TWK Listing/Product&DimsLog.xlsx", sheet = "DimensionLog", fillMergedCells = T) %>% mutate(Category = gsub("\\(.*", "", gsub("（.*", "", Category)), Category = ifelse(Category == "", NA, Category), Category = na.locf(Category), Category = strsplit(Category, "/"), Size = strsplit(Size, "/")) %>% 
  unnest(Category) %>% unnest(Size) %>% as.data.frame() %>% mutate(Category = str_trim(Category), Size = gsub("[TY]", "", toupper(str_trim(Size)))) %>% filter(!duplicated(paste(Category, Size, sep = "_")))
rownames(dimensions) <- paste(dimensions$Category, dimensions$Size, sep = "_")
products_description <- read.xlsx("../Listing/XHS/products_description.xlsx", sheet = 1)
rownames(products_description) <- products_description$SPU
products_description_cat <- read.xlsx("../Listing/XHS/products_description_categories.xlsx", sheet = 1)
rownames(products_description_cat) <- products_description_cat$cat

categories <- c("UG1", "UJ1", "USA", "USS", "UT1", "UV2", "UVS", "UVT")
woo$cat <- gsub("-.*", "", woo$SKU)
new_listing <- data.frame(SKU = woo[woo$cat %in% categories, "SKU"]) %>% 
  mutate(status = mastersku[SKU, "MSKU.Status"], SPU = woo[SKU, "Parent"], cat = gsub("-.*", "", SKU), s = paste(cat, gsub("[tyTY]", "", str_split_i(SKU, "-", i = 3)), sep = "_"), Product.Name = ifelse(cat %in% products_description_cat$cat, products_description_cat[cat, "Product.Name"], ""), Description = ifelse(cat %in% products_description_cat$cat, products_description_cat[cat, "Description"], ""), Vendor = "Jan & Jul", Categories = ifelse(cat %in% products_description_cat$cat, products_description_cat[cat, "Categories"], ""), Option1.Name = woo[SKU, "Attribute.1.name"], Option1.Value = woo[SKU, "Attribute.1.value.s."], Weight = dimensions[s, "Weight.to.Use.(g)"], Price = woo[SKU, "Sale.price"], Compare.At.Price = woo[SKU, "Regular.price"], Tags = ifelse(Price == Compare.At.Price, "正价", "特价"), Requires.Shipping = "TRUE", Taxable = "TRUE", Barcode = mastersku[SKU, "UPC.Active"], Image.Src = ifelse(SPU %in% products_description$SPU, products_description$Image.Src, ""), SEO.Product.Name = Product.Name, SEO.Description = Description, Weight.Unit = "g", Inventory = ifelse(xoro[SKU, "ATS"] < 20, 0, xoro[SKU, "ATS"])) %>%
  filter(status == "Active", !is.na(Inventory))  %>% select(-cat, -s, -status)
colnames(new_listing) <- gsub("\\.", " ", colnames(new_listing))
write.xlsx(new_listing, file = paste0("../Listing/XHS/new_listing-", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"))


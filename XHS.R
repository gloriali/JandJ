# Little Red Book Listing and syncing
library(dplyr)
library(xlsx)
library(openxlsx)
library(zoo)
library(stringr)
library(tidyr)

## Little Red Book Listing
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price)) %>% `row.names<-`(toupper(.[, "SKU"])) 
netsuite_item <- read.csv(list.files(path = "../NetSuite/", pattern = paste0("Items_All_", format(Sys.Date(), "%Y%m%d"), ".csv"), full.names = T), as.is = T)
netsuite_item[netsuite_item == "" | is.na(netsuite_item)] <- 0
netsuite_item_S <- netsuite_item %>% filter(Inventory.Warehouse == "WH-SURREY") %>% `row.names<-`(toupper(.[, "Name"])) 
products_description <- read.xlsx2("../XHS/products_description.xlsx", sheetIndex = 1) %>% `row.names<-`(toupper(.[, "SPU"])) 
wb <- openxlsx::loadWorkbook(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T))
openxlsx::protectWorkbook(wb, protect = F)
openxlsx::saveWorkbook(wb, list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), overwrite = T)
products_XHS <- read.xlsx2(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), sheetIndex = 1)

### -------------- Sync XHS price to woo; inventory to NS S ---------------
# input: AllValue > Products > Export > All products
products_description <- read.xlsx2("../XHS/products_description.xlsx", sheetIndex = 1) %>% `row.names<-`(toupper(.[, "SPU"])) 
wb <- openxlsx::loadWorkbook(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T))
openxlsx::protectWorkbook(wb, protect = F)
openxlsx::saveWorkbook(wb, list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), overwrite = T)
products_XHS <- read.xlsx2(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), sheetIndex = 1)
products_XHS[products_XHS=="NA"] <- ""
products_upload <- products_XHS %>% mutate(Inventory = ifelse(netsuite_item_S[SKU, "Warehouse.Available"] < 10, 0, netsuite_item_S[SKU, "Warehouse.Available"]), Price = woo[SKU, "Sale.price"], Compare.At.Price = woo[SKU, "Regular.price"], Tags = ifelse(Price == Compare.At.Price, "正价", "特价"))
products_upload <- products_upload %>% mutate(Product.Name = products_description[toupper(SPU), "Product.Name"], SEO.Product.Name = Product.Name, Description = products_description[toupper(SPU), "Description"], Mobile.Description = products_description[toupper(SPU), "Description"], SEO.Description = products_description[toupper(SPU), "Description"], Describe = products_description[toupper(SPU), "Description"])
colnames(products_upload) <- gsub("\\.", " ", colnames(products_upload))
openxlsx::write.xlsx(products_upload, file = paste0("../XHS/products_upload-", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"), na.string = "")
# upload to AllValue > Products

### -------- extract and update existing descriptions -----------------
products_description <- products_XHS %>% filter(!duplicated(SPU)) %>% mutate(SPU = toupper(SPU), cat = gsub("-.*", "", SKU), Description = paste(Description, SEO.Description, Describe), Description = str_trim(gsub("NA", "", Description))) %>% select(SPU, Product.Name, cat, Description, Categories, Option1.Name, Image.Src)
products_description_cat <- products_description %>% select(cat, Product.Name, Description, Categories, Option1.Name) %>% filter(!duplicated(cat)) %>% mutate(Product.Name = gsub("\\s?-\\s?\\w+$", "", Product.Name))
write.xlsx(products_description, file = "../XHS/products_description.xlsx", row.names = F)
write.xlsx(products_description_cat, file = "../XHS/products_description_categories.xlsx", row.names = F)
products_description_cat <- openxlsx::read.xlsx("../XHS/products_description_categories.xlsx", sheet = 1) %>% `row.names<-`(toupper(.[, "cat"])) 
products_description <- read.xlsx2("../XHS/products_description.xlsx", sheetIndex = 1) %>% `row.names<-`(toupper(.[, "SPU"])) 
products_description <- products_description %>% mutate(Product.Name = products_description_cat[cat, "Product.Name"], Description = products_description_cat[cat, "Description"], Categories = products_description_cat[cat, "Categories"], Option1.Name = products_description_cat[cat, "Option1.Name"])
write.xlsx(products_description, file = "../XHS/products_description.xlsx", row.names = F)

### -------- create new listing ------------------------
new_season <- "25"
mastersku <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../FBArefill/Raw Data File/", pattern = "1-MasterSKU-All-Product-", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% mutate(SPU = paste(Category.SKU, Print.SKU, sep = "-")) %>% `row.names<-`(toupper(.[, "MSKU"]))
catprint <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU_CatPrintsFactory.xlsx", full.names = T), sheet = "SKU- product category", fillMergedCells = T) %>% filter(!duplicated(SKU))
rownames(catprint) <- catprint$SKU
image <- read.csv("../woo/ImageSrc.csv", as.is = T, header = T)
rownames(image) <- image$SKU
woo$cat <- ifelse(woo$SKU %in% mastersku$MSKU, mastersku[woo$SKU, "Category.SKU"], gsub("-.*", "", woo$SKU))

#### description and details for the new listing
##### Add a new category
categories <- c("AJC", "FVM", "HBU", "HLH", "HLC", "SMF")
products_description <- read.xlsx2("../XHS/products_description.xlsx", sheetIndex = 1)
products_description_cat <- read.xlsx2("../XHS/products_description_categories.xlsx", sheetIndex = 1)
rownames(products_description_cat) <- products_description_cat$cat
products_description_cat <- rbind(products_description_cat, data.frame(cat = categories) %>% mutate(Product.Name = catprint[cat, "Style.Name"], Description = "", Categories = catprint[cat, "Product.Category"], Option1.Name = "Size"))
write.xlsx(products_description_cat, file = "../XHS/products_description_categories.xlsx", row.names = F)
# !open products_description_cat in excel and fill in info
products_description_cat <- read.xlsx2("../XHS/products_description_categories.xlsx", sheetIndex = 1)
rownames(products_description_cat) <- products_description_cat$cat
mastersku_SPU <- mastersku %>% filter(MSKU.Status == "Active") %>% filter(!duplicated(SPU))
rownames(mastersku_SPU) <- mastersku_SPU$SPU
#new_description <- data.frame(SPU = mastersku_SPU[mastersku_SPU$Category.SKU %in% categories, "SPU"]) %>% mutate(cat = mastersku_SPU[SPU, "Category.SKU"], Seasons = mastersku_SPU[SPU, "Seasons.SKU"], Product.Name = paste0(products_description_cat[cat, "Product.Name"], " | ", mastersku_SPU[SPU, "Print.Chinese"]), Description = products_description_cat[cat, "Description"], Categories = products_description_cat[cat, "Categories"], Option1.Name = products_description_cat[cat, "Option1.Name"], Image.Src = image[SPU, "Images"]) %>% filter(!(SPU %in% toupper(products_XHS$SPU)))
new_description <- data.frame(SPU = mastersku_SPU$SPU) %>% mutate(cat = mastersku_SPU[SPU, "Category.SKU"], Seasons = mastersku_SPU[SPU, "Seasons.SKU"], Product.Name = paste0(products_description_cat[cat, "Product.Name"], " | ", mastersku_SPU[SPU, "Print.Chinese"]), Description = products_description_cat[cat, "Description"], Categories = products_description_cat[cat, "Categories"], Option1.Name = products_description_cat[cat, "Option1.Name"], Image.Src = image[SPU, "Images"]) %>%
  filter(!(SPU %in% toupper(products_XHS$SPU)), cat %in% products_description_cat$cat, !(SPU %in% products_description$SPU), SPU %in% toupper(gsub("^(\\w+-\\w+)-.*", "\\1", woo$SKU)))
rownames(new_description) <- new_description$SPU
products_description <- rbind(products_description, new_description %>% select(-Seasons))
write.xlsx(products_description, file = "../XHS/products_description.xlsx", row.names = F)
# !open products_description in excel and fill in missing info
products_description <- read.xlsx2("../XHS/products_description.xlsx", sheetIndex = 1) 
rownames(products_description) <- products_description$SPU
new_listing <- data.frame(SKU = woo$SKU) %>% 
  mutate(status = mastersku[SKU, "MSKU.Status"], seasons = mastersku[SKU, "Seasons"], SPU = mastersku[SKU, "SPU"], cat = mastersku[SKU, "Category.SKU"], s = paste(cat, gsub("[tyTY]", "", mastersku[SKU, "Size"]), sep = "_"), Product.Name = products_description[SPU, "Product.Name"], Description = products_description[SPU, "Description"], Mobile.Description = Description, Vendor = "Jan & Jul", Categories = products_description[SPU, "Categories"], Option1.Name = products_description[SPU, "Option1.Name"], Option1.Value = mastersku[SKU, "Size"], Option2.Name = "", Option2.Value = "", Option3.Name = "", Option3.Value = "", Weight = woo[SKU, "Weight..g."], Price = woo[SKU, "Sale.price"], Compare.At.Price = woo[SKU, "Regular.price"], Tags = ifelse(Price == Compare.At.Price, "正价", "特价"), Requires.Shipping = "TRUE", Taxable = "TRUE", Barcode = mastersku[SKU, "UPC.Active"], Image.Src = products_description[SPU, "Image.Src"], Image.Position = "", SEO.Product.Name = Product.Name, SEO.Description = Description, Variant.Image = "", Weight.Unit = "g", Cost.Price = "", Describe = "", Inventory = ifelse(netsuite_item_S[SKU, "Warehouse.Available"] < 20, 0, netsuite_item_S[SKU, "Warehouse.Available"])) %>%
  filter(status == "Active", SKU %in% netsuite_item_S$Name, Inventory > 0, !(SPU %in% toupper(products_XHS$SPU)), cat %in% products_description_cat$cat, SPU %in% products_description$SPU)  %>% 
  select(SPU, Product.Name, Description, Mobile.Description, Vendor, Categories, Tags, Option1.Name, Option1.Value, Option2.Name, Option2.Value, Option3.Name, Option3.Value, SKU, Weight, Price, Compare.At.Price, Requires.Shipping, Taxable, Barcode, Image.Src, Image.Position, SEO.Product.Name, SEO.Description, Variant.Image, Weight.Unit, Cost.Price, Inventory, Describe)
colnames(new_listing) <- gsub("\\.", " ", colnames(new_listing))
openxlsx::write.xlsx(new_listing, file = paste0("../XHS/new_listing-", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"), sheetName = "Sample")
# upload to AllValue > Products > change status to "Listed". All Value > Little Red Book Applet > Products management > Add promotional product > Select categories

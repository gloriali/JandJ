# Little Red Book Listing and syncing
library(dplyr)
library(xlsx)
library(openxlsx)
library(zoo)
library(stringr)
library(tidyr)

## Little Red Book Listing
xoro <- read.csv(list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"), full.names = T), as.is = T) %>% filter(Store == "Warehouse - JJ") %>% mutate(Item. = toupper(Item.))
rownames(xoro) <- xoro$Item.
woo <- read.csv(list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y"))), ".*.csv"), full.names = T)) %>% 
  mutate(SKU = toupper(SKU), Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price)) %>% filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "")
rownames(woo) <- woo$SKU
products_description <- read.xlsx2("../XHS/products_description.xlsx", sheetIndex = 1)
rownames(products_description) <- products_description$SPU
products_XHS <- read.xlsx2(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), sheetIndex = 1)

### -------- update inventory with xoro, price with woo, product description ----------------
products_upload <- products_XHS %>% mutate(Inventory = ifelse(xoro[SKU, "ATS"] < 20, 0, xoro[SKU, "ATS"]), Price = woo[SKU, "Sale.price"], Compare.At.Price = woo[SKU, "Regular.price"], Tags = ifelse(Price == Compare.At.Price, "正价", "特价"))
products_upload <- products_upload %>% mutate(Description = products_description[toupper(SPU), "Description"], Mobile.Description = products_description[toupper(SPU), "Description"], SEO.Description = products_description[toupper(SPU), "Description"], Describe = products_description[toupper(SPU), "Description"])
colnames(products_upload) <- gsub("\\.", " ", colnames(products_upload))
openxlsx::write.xlsx(products_upload, file = paste0("../XHS/products_upload-", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"))
# upload to AllValue > Products

### -------- extract and update existing descriptions -----------------
products_description <- products_XHS %>% filter(!duplicated(SPU)) %>% mutate(SPU = toupper(SPU), cat = gsub("-.*", "", SKU), Description = paste(Description, SEO.Description, Describe), Description = str_trim(gsub("NA", "", Description))) %>% select(SPU, Product.Name, cat, Description, Categories, Option1.Name, Image.Src)
products_description_cat <- products_description %>% select(cat, Product.Name, Description, Categories, Option1.Name) %>% filter(!duplicated(cat)) %>% mutate(Product.Name = gsub("\\s?-\\s?\\w+$", "", Product.Name))
write.xlsx(products_description, file = "../XHS/products_description.xlsx")
write.xlsx(products_description_cat, file = "../XHS/products_description_categories.xlsx")
products_description_cat <- openxlsx::read.xlsx("../XHS/products_description_categories.xlsx", sheet = 1)
rownames(products_description_cat) <- products_description_cat$cat
products_description <- products_description %>% mutate(Description = products_description_cat[cat, "Description"])
write.xlsx(products_description, file = "../XHS/products_description.xlsx")

### -------- create new listing ------------------------
new_season <- "24"
categories <- products_description_cat$cat
mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% mutate(SPU = paste(Category.SKU, Print.SKU, sep = "-"))
rownames(mastersku) <- mastersku$MSKU
woo$cat <- ifelse(woo$SKU %in% mastersku$MSKU, mastersku[woo$SKU, "Category.SKU"], gsub("-.*", "", woo$SKU))
woo_cat <- woo %>% filter(!duplicated(cat)) 
rownames(woo_cat) <- woo_cat$cat
image <- read.table("../woo/ImageSrc.tsv", sep = "\t", as.is = T, header = T)
rownames(image) <- image$SKU

#### description and details for the new listing
products_description_cat <- read.xlsx2("../XHS/products_description_categories.xlsx", sheetIndex = 1)
rownames(products_description_cat) <- products_description_cat$cat
products_description_cat <- rbind(products_description_cat, data.frame(cat = categories) %>% mutate(Product.Name = woo_cat[cat, "Name"], Description = "", Categories = "", Option1.Name = woo_cat[cat, "Attribute.1.name"]))
write.xlsx(products_description_cat, file = "../XHS/products_description_categories.xlsx")
# !open products_description_cat in excel and fill in info
products_description_cat <- read.xlsx2("../XHS/products_description_categories.xlsx", sheetIndex = 1)
rownames(products_description_cat) <- products_description_cat$cat
mastersku_SPU <- mastersku %>% filter(MSKU.Status == "Active", !duplicated(SPU))
rownames(mastersku_SPU) <- mastersku_SPU$SPU
new_description <- data.frame(SPU = mastersku_SPU[mastersku_SPU$Category.SKU %in% categories, "SPU"]) %>% mutate(cat = mastersku_SPU[SPU, "Category.SKU"], Product.Name = paste0(products_description_cat[cat, "Product.Name"], " - ", mastersku_SPU[SPU, "Print.Chinese"]), Description = products_description_cat[cat, "Description"], Categories = products_description_cat[cat, "Categories"], Option1.Name = products_description_cat[cat, "Option1.Name"], Image.Src = image[SPU, "Images"]) %>%
  filter(!(SPU %in% toupper(products_XHS$SPU)))
rownames(new_description) <- new_description$SPU

new_listing <- data.frame(SKU = woo[woo$cat %in% categories, "SKU"]) %>% 
  mutate(status = mastersku[SKU, "MSKU.Status"], seasons = mastersku[SKU, "Seasons"], SPU = mastersku[SKU, "SPU"], cat = mastersku[SKU, "Category.SKU"], s = paste(cat, gsub("[tyTY]", "", mastersku[SKU, "Size"]), sep = "_"), Product.Name = new_description[SPU, "Product.Name"], Description = new_description[SPU, "Description"], Mobile.Description = Description, Vendor = "Jan & Jul", Categories = new_description[SPU, "Categories"], Option1.Name = new_description[SPU, "Option1.Name"], Option1.Value = mastersku[SKU, "Size"], Weight = woo[SKU, "Weight..g."], Price = woo[SKU, "Sale.price"], Compare.At.Price = woo[SKU, "Regular.price"], Tags = ifelse(Price == Compare.At.Price, "正价", "特价"), Requires.Shipping = "TRUE", Taxable = "TRUE", Barcode = mastersku[SKU, "UPC.Active"], Image.Src = new_description[SPU, "Image.Src"], Image.Position = "", SEO.Product.Name = Product.Name, SEO.Description = Description, Variant.Image = "", Weight.Unit = "g", Cost.Price = "", Describe = "", Inventory = ifelse(xoro[SKU, "ATS"] < 20, 0, xoro[SKU, "ATS"])) %>%
  filter(status == "Active", SKU %in% xoro$Item., grepl(new_season, seasons) | Inventory > 0, SPU %in% new_description$SPU)  %>% 
  select(SPU, Product.Name, Description, Mobile.Description, Vendor, Categories, Tags, Option1.Name, Option1.Value, SKU, Weight, Price, Compare.At.Price, Requires.Shipping, Taxable, Barcode, Image.Src, Image.Position, SEO.Product.Name, SEO.Description, Variant.Image, Weight.Unit, Cost.Price, Inventory, Describe)
colnames(new_listing) <- gsub("\\.", " ", colnames(new_listing))
openxlsx::write.xlsx(new_listing, file = paste0("../XHS/new_listing-", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"))
# upload to AllValue > Products > change status to "Listed". All Value > Little Red Book Applet > Products management > Add promotional product > Select categories

# Generate Square items library for uploading
setwd("~/OneDrive - Jan and Jul/TWK - Gloria/JandJ")
library(dplyr)

# ------- match Square price to woo; stock to xoro ------
# input: woo > Products > All Products > Export > all
# input: xoro > Item Inventory Snapshot > Export all
woo <- read.csv(paste0("../woo/", list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))))), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price))
rownames(woo) <- woo$SKU
xoro <- read.csv(paste0("../xoro/", list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"))), as.is = T) %>% filter(Store == "Warehouse - JJ")
rownames(xoro) <- xoro$Item.

square <- data.frame(ItemName = woo$SKU, SKU = woo$SKU, Description = "", Category = xoro[woo$SKU, "Item.Category"], Price = woo$Sale.price, PriceRichmond = woo$Sale.price, PriceSurrey = woo$Sale.price, NewQuantityRichmond = xoro[woo$SKU, "ATS"], NewQuantitySurrey = xoro[woo$SKU, "ATS"], Token = "", VariationName = "Regular", EnabledRichmond = "Y", CurrentQuantityRichmond = "", CurrentQuantitySurrey = "", StockAlertEnabledRichmond = "N", StockAlertCountRichmond = 0, EnabledSurrey = "Y", StockAlertEnabledSurrey = "N", StockAlertCountSurrey = 0, TaxPST = woo$Tax.class)
rownames(square) <- square$SKU
square$TaxPST <- gsub("parent", "N", gsub("full", "Y", square$TaxPST))
square$PriceRichmond <- square$Price
square$PriceSurrey <- square$Price
write.table(square, file = paste0("../Square/square-upload-", Sys.Date(), ".csv"), sep = ",", row.names = F, col.names = c(colnames(square)[1:ncol(square)-1], "Tax-PST(7%)"), na = "")
# upload to Square > Items

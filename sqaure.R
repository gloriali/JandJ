# Generate Square items library for uploading
setwd("~/OneDrive - Jan and Jul/TWK - Gloria/JandJ")
library(dplyr)
# ------- match Square price to woo ------
# input: woo > Products > All Products > Export > all
woo <- read.csv(paste0("../woo/", list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))))), as.is = T)
woo1 <- woo %>% filter(!is.na(woo$Regular.price) & !duplicated(woo$SKU) & !(woo$SKU == "")) 
sales <- woo1 %>% filter(!is.na(woo1$Sale.price)) 
rownames(sales) <- sales$SKU

# ------- match Square stock to xoro ------
# input: xoro > Item Inventory Snapshot > Export all
xoro <- read.csv(paste0("../xoro/", list.files(path = "../xoro/", pattern = paste0("Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".csv"))), as.is = T) %>% filter(Store == "Warehouse - JJ")
rownames(xoro) <- xoro$Item.

square <- data.frame(ItemName = woo1$SKU, SKU = woo1$SKU, Description = "", Category = xoro[woo1$SKU, "Item.Category"], Price = woo1$Regular.price, PriceRichmond = woo1$Regular.price, PriceSurrey = woo1$Regular.price, NewQuantityRichmond = xoro[woo1$SKU, "ATS"], NewQuantitySurrey = xoro[woo1$SKU, "ATS"], Token = "", VariationName = "Regular", EnabledRichmond = "Y", CurrentQuantityRichmond = "", CurrentQuantitySurrey = "", StockAlertEnabledRichmond = "N", StockAlertCountRichmond = 0, EnabledSurrey = "Y", StockAlertEnabledSurrey = "N", StockAlertCountSurrey = 0, TaxPST = woo1$Tax.class)
rownames(square) <- square$SKU
square$TaxPST <- gsub("parent", "N", gsub("full", "Y", square$TaxPST))
square[sales$SKU, "Price"] <- sales[sales$SKU, "Sale.price"]
write.table(square, file = paste0("../Square/square-upload-", Sys.Date(), ".csv"), sep = ",", row.names = F, col.names = c(colnames(square)[1:ncol(square)-1], "Tax-PST(7%)"), na = "")

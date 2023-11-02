# update price for Square
# input: woo > Products > All Products > Export > all
library(dplyr)
woo <- read.csv(paste0("../woo/", list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))))), as.is = T)
woo1 <- woo %>% filter(!is.na(woo$Regular.price) & !duplicated(woo$SKU) & !(woo$SKU == "")) 
woo1[grep("-", woo1$Stock), "Stock"] <- 0
woo1$Stock <- as.numeric(woo1$Stock)
sales <- woo1 %>% filter(!is.na(woo1$Sale.price)) 
rownames(sales) <- sales$SKU

square <- data.frame(ItemName = woo1$SKU, SKU = woo1$SKU, Description = "", Category = gsub("\\ *\\|.*", "", woo1$Name), Price = woo1$Regular.price, PriceRichmond = woo1$Regular.price, PriceSurrey = woo1$Regular.price, NewQuantityRichmond = woo1$Stock, NewQuantitySurrey = woo1$Stock, Token = "", VariationName = "Regular", EnabledRichmond = "Y", CurrentQuantityRichmond = "", CurrentQuantitySurrey = "", StockAlertEnabledRichmond = "N", StockAlertCountRichmond = 0, EnabledSurrey = "Y", StockAlertEnabledSurrey = "N", StockAlertCountSurrey = 0, TaxPST = woo1$Tax.class)
rownames(square) <- square$SKU
square$TaxPST <- gsub("parent", "N", gsub("full", "Y", square$TaxPST))
square[sales$SKU, "Price"] <- sales[sales$SKU, "Sale.price"]
write.table(square, file = paste0("../Square/square-upload-", Sys.Date(), ".csv"), sep = ",", row.names = F, col.names = c(colnames(square)[1:ncol(square)-1], "Tax-PST(7%)"), na = "")

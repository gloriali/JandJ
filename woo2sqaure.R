# update price for Square
# input: woo > Products > All Products > Export > all
library(dplyr)
woo <- read.csv("../Square/wc-product-export-30-10-2023-1698685138568.csv", as.is = T)
woo1 <- woo %>% filter(!is.na(woo$Regular.price) & !duplicated(woo$SKU) & !(woo$SKU == "")) 
woo1[grep("-", woo1$Stock), "Stock"] <- 0
woo1$Stock <- as.numeric(woo1$Stock)
sales <- woo1 %>% filter(!is.na(woo1$Sale.price)) 
rownames(sales) <- sales$SKU

square <- data.frame(ItemName = woo1$SKU, SKU = woo1$SKU, Description = woo1$Name, Category = gsub("\\ *\\|.*", "", woo1$Name), Price = woo1$Regular.price, PriceRichmond = woo1$Regular.price, PriceSurrey = woo1$Regular.price, NewQuantityRichmond = woo1$Stock, NewQuantitySurrey = woo1$Stock, Token = NA, VariationName = "Regular", EnabledRichmond = "Y", CurrentQuantityRichmond = NA, CurrentQuantitySurrey = NA, StockAlertEnabledRichmond = "N", StockAlertCountRichmond = 0, EnabledSurrey = "Y", StockAlertEnabledSurrey = "N", StockAlertCountSurrey = 0)
rownames(square) <- square$SKU
square[sales$SKU, "Price"] <- sales[sales$SKU, "Sale.price"]
write.csv(square, file = "../Square/square-upload-20231030.csv", row.names = F, na = "")

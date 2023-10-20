# update price for Sqaure
library(dplyr)
woo <- read.csv("../Sqaure/wc-product-export-19-10-2023-1697732118756.csv")
woo1 <- woo %>% filter(!is.na(woo$Regular.price)) 
woo1 <- woo1 %>% filter(!duplicated(woo1$SKU))
sales <- woo1 %>% filter(!is.na(woo1$Sale.price)) 
rownames(sales) <- sales$SKU

square <- data.frame(ItemName = woo1$SKU, SKU = woo1$SKU, Description = woo1$Name, Category = gsub("\\ *\\|.*", "", woo1$Name), Price = woo1$Regular.price, Stock = woo1$Stock)
rownames(square) <- square$SKU
square[sales$SKU, "Price"] <- sales[sales$SKU, "Sale.price"]
write.csv(square, file = "../Sqaure/square.csv", row.names = F)

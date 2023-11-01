# update price for clover
library(dplyr)
woo <- read.csv("../woo/wc-product-export-31-10-2023-1698777245011.csv", as.is = T)
woo1 <- woo %>% filter(!is.na(woo$Regular.price) & !duplicated(woo$SKU) & !(woo$SKU == "")) 
rownames(woo1) <- woo1$SKU
sales <- woo1 %>% filter(!is.na(woo1$Sale.price)) 
rownames(sales) <- sales$SKU

clover <- read.csv("../Clover/inventory20231031-items.csv", as.is = T)
rownames(clover) <- clover$Name
clover$Price <- woo1[clover$Name, "Regular.price"]
clover[sales$SKU, "Price"] <- sales[sales$SKU, "Sale.price"]
write.csv(clover, file = "../Clover/inventory20231031-items-upload.csv", row.names = F, na = "")


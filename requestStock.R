# request stock from Surrey warehouse
# download current Surrey stock: wc > Products > All Products > Export; download current Richmond stock: clover > Inventory > Items > Export
library(dplyr)
request <- c("ICP", "IPC", "IPS", "ISB", "WJA", "WMR", "WMT", "WPF")
woo <- read.csv(paste0("../woo/", list.files(path = "../woo/", pattern = paste0("wc-product-export-", sub("-0", "", sub("^0", "", format(Sys.Date(), "%d-%m-%Y")))))), as.is = T)
woo1 <- woo %>% filter(!is.na(woo$Regular.price) & !duplicated(woo$SKU) & !(woo$SKU == "")) 
woo1[grep("-", woo1$Stock), "Stock"] <- 0
woo1$Stock <- as.numeric(woo1$Stock)
rownames(woo1) <- woo1$SKU

clover <- read.csv(paste0("../Clover/", list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), "-items.csv"))), as.is = T)
clover$cat <- gsub("-.*", "", clover$Name)
clover[is.na(clover$Quantity), "Quantity"] <- 0

order <- data.frame(StoreCode = "WH-JJ", ItemNumber=(clover %>% filter(Quantity<1 & cat %in% request))$Name, Qty = 1, LocationName = "BIN", UnitCost = "", ReasonCode = "RWT", Memo = "Richmond Transfer to Miranda", UploadRule = "D", AdjAccntName = "", TxnDate = "", ItemIdentifierCode = "", ImportError = "")
order <- order %>% filter(woo1[order$ItemNumber, "Stock"] > 10)
write.csv(order, file = paste0("../Clover/order", format(Sys.Date(), "%m%d%Y"), ".csv"), row.names = F, na = "")

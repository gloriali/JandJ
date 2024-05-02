# SKUs analysis and management
library(dplyr)
library(openxlsx)
library(scales)
library(ggplot2)
library(xlsx)

# ------------- upload Richmond inventory ---------------------------
netsuite_items <- read.csv("../NetSuite/Items531.csv", as.is = T)
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
inventory_update <- openxlsx::readWorkbook(clover, "Items") %>% filter(SKU %in% netsuite_items$MSKU, Quantity > 0) %>% mutate(Item = SKU, Adjust.Qty.By = Quantity, BIN = "BIN", Reason.Code = "Found Inventory", MEMO = "Inventory Upload", Location = "WH-RICHMOND", External.ID = paste0("IAR", format(Sys.Date(), "%y%m%d"))) %>%
  select(Item, Adjust.Qty.By, Quantity, BIN, Reason.Code, MEMO, Location, External.ID)
colnames(inventory_update) <- gsub("\\.", " ", colnames(inventory_update))
write.csv(inventory_update, file = paste0("../NetSuite/inventory_update-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F)

# ------------- upload Clover SO ---------------------------
clover_so <- read.csv("../Clover/LineItemsExport-20240502_0000_PDT-20240502_2359_PDT.csv", as.is = T) 
netsuite_so <- clover_so %>% filter(Item.SKU != "", Order.Payment.State == "Paid") %>% mutate(Order.date = format(as.Date(gsub(" .*", "", Line.Item.Date), "%d-%b-%Y"), "%m/%d/%Y")) %>% group_by(Order.date) %>% 
  mutate(ID = data.table::rleid(Order.ID), REF.ID = paste0("POP", format(as.Date(Order.date, "%m/%d/%Y"), "%Y%m%d"), "-", sprintf("%02d", ID)), Order.Type = "JJR OFFLINE", Department = "Retail : Store Richmond", Warehouse = "WH-Richmond", Quantity = 1, Price.level = "Custom", Rate = Item.Revenue, Amount = Item.Total, Coupon.Discount = Total.Discount + Order.Discount.Proportion, Coupon.Code = gsub(" -.*","", paste0(Order.Discounts, Discounts)), Tax.Code = ifelse(Item.Tax.Rate == 0.05, "CA-BC-GST", ifelse(Item.Tax.Rate == 0.12, "CA-BC-TAX", ""))) %>% 
  select(Order.date, REF.ID, Order.Type, Department, Warehouse, Order.ID, Item.SKU, Quantity, Price.level, Rate, Amount, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount)
write.csv(netsuite_so, file = paste0("../NetSuite/SO-clover-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F)

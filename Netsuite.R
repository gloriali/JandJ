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
clover_so <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "LineItemsExport-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T)
netsuite_so <- clover_so %>% filter(Item.SKU != "", Order.Payment.State == "Paid", !grepl("XHS", Order.Discounts)) %>% mutate(Order.date = format(as.Date(gsub(" .*", "", Line.Item.Date), "%d-%b-%Y"), "%m/%d/%Y")) %>% group_by(Order.date) %>% 
  mutate(Payment.Option = "Clover", Class = "FBM : CA", MEMO = "Clover sales", ID = data.table::rleid(Order.ID), REF.ID = paste0("CL", format(as.Date(Order.date, "%m/%d/%Y"), "%Y%m%d"), "-", sprintf("%02d", ID)), Order.Type = "JJR SHOPR", Department = "Retail : Store Richmond", Warehouse = "WH-Richmond", Quantity = 1, Price.level = "Custom", Rate = Item.Revenue, Amount = Item.Total, Coupon.Discount = Total.Discount + Order.Discount.Proportion, Coupon.Code = gsub("NA", "", gsub(" -.*","", paste0(Order.Discounts, Discounts))), Tax.Code = ifelse(Item.Tax.Rate == 0.05, "CA-BC-GST", ifelse(Item.Tax.Rate == 0.12, "CA-BC-TAX", ""))) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, Price.level, Rate, Amount, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount)
write.csv(netsuite_so, file = paste0("../NetSuite/SO-clover-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F)

# ------------- upload Square SO ---------------------------
square_so <- read.csv(rownames(file.info(list.files(path = "../Square/", pattern = "items-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T)
netsuite_so <- square_so %>% filter(Item != "", Location == "Surrey") %>% mutate(Order.date = format(as.Date(Date, "%Y-%m-%d"), "%m/%d/%Y")) %>% group_by(Order.date) %>% 
  mutate(Payment.Option = "Square", Class = "FBM : CA", MEMO = "Square sales", Order.ID = Transaction.ID, Item.SKU = Item, ID = data.table::rleid(Order.ID), REF.ID = paste0("SQ", format(as.Date(Order.date, "%m/%d/%Y"), "%Y%m%d"), "-", sprintf("%02d", ID)), Order.Type = "JJR SHOPS", Department = "Retail : Store Surrey", Warehouse = "WH-SURREY", Quantity = Qty, Price.level = "Custom", Gross.Sales = as.numeric(gsub("\\$", "", Gross.Sales)), Net.Sales = as.numeric(gsub("\\$", "", Net.Sales)), Discounts = as.numeric(gsub("\\$", "", Discounts)), Tax = as.numeric(gsub("\\$", "", Tax)), Rate = round(Gross.Sales/Qty, 2), Amount = Net.Sales, Coupon.Discount = Discounts, Coupon.Code = "", Tax.Code = ifelse(round(Tax/Net.Sales, 2) == 0.05, "CA-BC-GST", ifelse(round(Tax/Net.Sales, 2) == 0.12, "CA-BC-TAX", "")), Tax.Amount = Tax) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, Price.level, Rate, Amount, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount)
write.csv(netsuite_so, file = paste0("../NetSuite/SO-square-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F)

# SKUs analysis and management
library(dplyr)
library(openxlsx)
library(scales)
library(ggplot2)
library(xlsx)
library(data.table)
# netsuite_R <- fread(list.files(path = "../NetSuite/", pattern = paste0("CurrentInventorySnapshot", format(Sys.Date(), "%Y%m%d"), ".csv"), full.names = T), data.table = F, skip = 7, select = c(1, 11:17)) %>% `row.names<-`(.[, "V1"])
# netsuite_R[netsuite_R == "" | is.na(netsuite_R)] <- 0
# netsuite_R[, 2:8] <- lapply(netsuite_R[, 2:8], function(x) as.numeric(gsub("\\,", "", x)))
# netsuite_S <- fread(list.files(path = "../NetSuite/", pattern = paste0("CurrentInventorySnapshot", format(Sys.Date(), "%Y%m%d"), ".csv"), full.names = T), data.table = F, skip = 7, select = c(1, 18:24)) %>% `row.names<-`(.[, "V1"])
# netsuite_S[netsuite_S == "" | is.na(netsuite_S)] <- 0
# netsuite_S[, 2:8] <- lapply(netsuite_S[, 2:8], function(x) as.numeric(gsub("\\,", "", x)))
# netsuite_C <- fread(list.files(path = "../NetSuite/", pattern = paste0("CurrentInventorySnapshot", format(Sys.Date(), "%Y%m%d"), ".csv"), full.names = T), data.table = F, skip = 7, select = c(1, 4:10)) %>% `row.names<-`(.[, "V1"])
# netsuite_C[netsuite_C == "" | is.na(netsuite_C)] <- 0
# netsuite_C[, 2:8] <- lapply(netsuite_C[, 2:8], function(x) as.numeric(gsub("\\,", "", x)))

# ------------- update Richmond inventory ---------------------------
netsuite_item <- read.csv(list.files(path = "../NetSuite/", pattern = paste0("Items", format(Sys.Date(), "%Y%m%d"), ".csv"), full.names = T), as.is = T) %>% mutate(Name = toupper(Name))
netsuite_item[netsuite_item == "" | is.na(netsuite_item)] <- 0
clover_item <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items") %>% filter(Name %in% netsuite_item$Name) %>% mutate(Quantity = ifelse(Quantity < 0, 0, Quantity))
netsuite_item <- netsuite_item %>% filter(Inventory.Warehouse == "WH-RICHMOND") %>% `row.names<-`(.[, "Name"])
inventory_update <- clover_item %>% mutate(Date = format(Sys.Date(), "%m/%d/%Y"), ITEM = Name, ADJUST.QTY.BY = ifelse(Name %in% netsuite_item$Name, Quantity - netsuite_item[Name, "Warehouse.On.Hand"], Quantity), Reason.Code = "CC", MEMO = "Clover Inventory Update", WAREHOUSE = "WH-RICHMOND", External.ID = paste0("IA", format(Sys.Date(), "%y%m%d"))) %>%
  select(Date, ITEM, ADJUST.QTY.BY, Reason.Code, MEMO, WAREHOUSE, External.ID) %>% filter(ADJUST.QTY.BY != 0)
colnames(inventory_update) <- gsub("QTY BY", "QTY. BY", gsub("\\.", " ", colnames(inventory_update)))
write.csv(inventory_update, file = paste0("../NetSuite/IA-Richmond-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F)

# ------------- upload Clover SO ---------------------------
customer <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "Customers-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>%
  mutate(Name = paste0(First.Name, " ", Last.Name)) %>% distinct(Email.Address, .keep_all = T) %>% filter(Name != " ", Email.Address != "") %>% `row.names<-`(.[, "Name"])
payments <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "Payments-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% mutate(Phone = customer[Customer.Name, "Phone.Number"], Email = customer[Customer.Name, "Email.Address"]) %>% `row.names<-`(.[, "Order.ID"]) 
clover_so <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "LineItemsExport-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% mutate(Recipient = payments[Order.ID, "Customer.Name"], Recipient.Phone = payments[Order.ID, "Phone"], Recipient.Email = payments[Order.ID, "Email"], Tender = payments[Order.ID, "Tender"])
netsuite_so <- clover_so %>% filter(Item.SKU != "", Order.Payment.State == "Paid", !grepl("XHS", Order.Discounts)) %>% mutate(Order.date = format(as.Date(gsub(" .*", "", Line.Item.Date), "%d-%b-%Y"), "%m/%d/%Y")) %>% group_by(Order.date) %>% 
  mutate(Payment.Option = ifelse(Tender == "Cash", "Cash", "Clover"), Class = "FBM : CA", MEMO = "Clover sales", BIN = "BIN", Customer = "15139 JJR SHOPR", ID = data.table::rleid(Order.ID), REF.ID = paste0("CL", format(as.Date(Order.date, "%m/%d/%Y"), "%Y%m%d"), "-", sprintf("%02d", ID)), Order.Type = "JJR SHOPR", Department = "Retail : Store Richmond", Warehouse = "WH-RICHMOND", Quantity = 1, Price.level = "Custom", Rate = Item.Total, Coupon.Discount = Total.Discount + Order.Discount.Proportion, Coupon.Discount = ifelse(Coupon.Discount == 0, "", Coupon.Discount), Coupon.Code = gsub("NA", "", gsub(" -.*","", paste0(Order.Discounts, Discounts))), Tax.Code = ifelse(Item.Tax.Rate == 0.05, "CA-BC-GST", ifelse(Item.Tax.Rate == 0.12, "CA-BC-TAX", ""))) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, BIN, Price.level, Rate, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount, Customer, Recipient, Recipient.Phone, Recipient.Email)
write.csv(netsuite_so, file = paste0("../NetSuite/SO-clover-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F)

# ------------- upload Square SO ---------------------------
customer <- read.csv("../Square/customers.csv", as.is = T) %>% `row.names<-`(.[, "Square.Customer.ID"])
payments <- read.csv(rownames(file.info(list.files(path = "../Square/", pattern = "transactions-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% `row.names<-`(.[, "Payment.ID"])
square_so <- read.csv(rownames(file.info(list.files(path = "../Square/", pattern = "items-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% 
  mutate(Cash = payments[Payment.ID, "Cash"], Recipient.Email = ifelse(is.na(Customer.ID), "", customer[Customer.ID, "Email.Address"]), Recipient.Phone = ifelse(is.na(Customer.ID), "", customer[Customer.ID, "Phone.Number"]))
netsuite_so <- square_so %>% filter(Item != "") %>% mutate(Order.date = format(as.Date(Date, "%Y-%m-%d"), "%m/%d/%Y")) %>% group_by(Order.date) %>% 
  mutate(Payment.Option = ifelse(Cash == "$0.00" | is.na(Cash), "Square", "Cash"), Class = "FBM : CA", MEMO = "Square sales", BIN = "BIN", Customer = ifelse(Location == "Surrey", "15145 JJR SHOPS", "15139 JJR SHOPR"), Recipient = ifelse(is.na(Customer.Name), "", Customer.Name), Order.ID = Transaction.ID, Item.SKU = Item, ID = data.table::rleid(Order.ID), REF.ID = paste0("SQ", format(as.Date(Order.date, "%m/%d/%Y"), "%Y%m%d"), "-", sprintf("%02d", ID)), Order.Type = ifelse(Location == "Surrey", "JJR SHOPS", "JJR SHOPR"), Department = ifelse(Location == "Surrey", "Retail : Store Surrey", "Retail : Store Richmond"), Warehouse = ifelse(Location == "Surrey", "WH-SURREY", "WH-RICHMOND"), Quantity = Qty, Price.level = "Custom", Coupon.Discount = as.numeric(gsub("\\$", "", Discounts)), Coupon.Discount = ifelse(Coupon.Discount == 0, "", Coupon.Discount), Coupon.Code = "", Tax = as.numeric(gsub("\\$", "", Tax)), Rate = round(as.numeric(gsub("\\$", "", Net.Sales))/Qty, 2), Tax.Code = ifelse(round(Tax/as.numeric(gsub("\\$", "", Net.Sales)), 2) == 0.05, "CA-BC-GST", ifelse(round(Tax/as.numeric(gsub("\\$", "", Net.Sales)), 2) == 0.12, "CA-BC-TAX", "")), Tax.Amount = Tax) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, BIN, Price.level, Rate, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount, Customer, Recipient, Recipient.Phone, Recipient.Email)
write.csv(netsuite_so, file = paste0("../NetSuite/SO-square-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F)

# ------------- upload XHS SO ---------------------------
XHS_so <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../XHS/", pattern = "order_export", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = 1) 
netsuite_so <- XHS_so %>% filter(Financial.Status == "paid") %>% mutate(Order.date = format(as.Date(gsub(" .*", "", Created.At), "%Y-%m-%d"), "%m/%d/%Y")) %>% 
  mutate(Payment.Option = "AlphaPay", Class = "FBM : CN", MEMO = "XHS sales", BIN = "BIN", Customer = "15143 XHS CN", Order.ID = Order.No, REF.ID = paste0("XHS", Order.No), Order.Type = "XHS CN", Department = "Retail : Marketplace : Redbook", Warehouse = "WH-RICHMOND", Item.SKU = LineItem.SKU, Quantity = as.numeric(LineItem.Quantity), Price.level = "Custom", Rate = round((as.numeric(LineItem.Total) - as.numeric(LineItem.Total.Discount))/Quantity, 2), Coupon.Discount = as.numeric(LineItem.Total.Discount), Coupon.Discount = ifelse(Coupon.Discount == 0, "", Coupon.Discount), Coupon.Code = "", Tax.Amount = "", Tax.Code = "", Recipient = paste0(Shipping.Last.Name, Shipping.First.Name), Recipient.Phone = Phone, Recipient.Email = Email) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, BIN, Price.level, Rate, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount, Customer, Recipient, Recipient.Phone, Recipient.Email)
Encoding(netsuite_so$Customer) = "UTF-8"
write_excel_csv(netsuite_so, file = paste0("../NetSuite/SO-XHS-", format(Sys.Date(), "%Y%m%d"), ".csv"))
write.csv(netsuite_so, file = paste0("../NetSuite/SO-XHS-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, fileEncoding = "UTF-8")

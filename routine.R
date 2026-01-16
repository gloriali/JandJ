# Routine tasks
library(dplyr)
library(stringr)
library(openxlsx)
library(xlsx)
library(scales)
library(ggplot2)
library(tidyr)
library(readr)

# input: NS > Items > Export all warehouses
netsuite_item <- read.csv(rownames(file.info(list.files(path = "../NetSuite/", pattern = "Items_All_", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) 
netsuite_item[netsuite_item == "" | is.na(netsuite_item)] <- 0
netsuite_item_S <- netsuite_item %>% filter(Inventory.Warehouse == "WH-SURREY") %>% `row.names<-`(.[, "Name"]) 
netsuite_item_R <- netsuite_item %>% filter(Inventory.Warehouse == "WH-RICHMOND") %>% `row.names<-`(.[, "Name"]) 
# input: woo > Products > All Products > Export > all
woo <- read.csv(rownames(file.info(list.files(path = "../woo/", pattern = "wc-product-export-", full.names = T)) %>% filter(mtime == max(mtime))), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price) | (Sys.time() < strptime(Date.sale.price.starts, format = "%Y-%m-%d %H:%M:%S") | Sys.time() > strptime(Date.sale.price.ends, format = "%Y-%m-%d %H:%M:%S")), Regular.price, Sale.price)) %>% `row.names<-`(.[, "SKU"])
#woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
#  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price)) %>% `row.names<-`(.[, "SKU"])
#woo <- woo %>% mutate(Sale.price = ifelse(grepl("^[HGUS]", SKU), round(0.8*Regular.price, 2), Sale.price))
#Cat_sale <- data.frame(CAT = c("LBT","LBP", "LAN", "LAB", "LCT", "LCP", "FJM", "FPM", "FSM", "FJC", "FVM", "FHA", "FMR", "FAN", "KHB", "KHP", "KMN", "KMT", "KEH"), discount = c(rep(0.8, 17), rep(0.6, 2))) %>% `row.names<-`(.[, "CAT"])
#woo <- woo %>% mutate(cat = gsub("-.*", "", SKU), Sale.price = ifelse(cat %in% Cat_sale$CAT, round(Regular.price * Cat_sale[cat, "discount"], 2), Sale.price))

# ------------- upload Clover SO to NS, correct Clover inventory and update price: daily ---------------------------
customer <- read.csv("../Clover/Customers.csv", as.is = T) %>% mutate(Name = paste0(First.Name, " ", Last.Name)) %>% distinct(Email.Address, .keep_all = T) %>% filter(Name != " ", Email.Address != "") %>% `row.names<-`(.[, "Name"])
payments <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "Payments-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% mutate(Phone = customer[Customer.Name, "Phone.Number"], Email = customer[Customer.Name, "Email.Address"]) %>% filter(Result == "SUCCESS") %>% `row.names<-`(.[, "Order.ID"]) 
clover_so <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "LineItemsExport-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% mutate(Item.SKU = ifelse(Item.SKU %in% netsuite_item$Name, Item.SKU, "MISC-ITEM"), Recipient = payments[Order.ID, "Customer.Name"], Recipient.Phone = payments[Order.ID, "Phone"], Recipient.Email = payments[Order.ID, "Email"], Tender = payments[Order.ID, "Tender"])
netsuite_so <- clover_so %>% filter(Item.SKU != "", Order.Payment.State == "Paid", Refunded == "false", !grepl("XHS", Order.Discounts), !grepl("Surrey", Order.Discounts)) %>% mutate(Order.date = format(as.Date(gsub(" .*", "", Line.Item.Date), "%d-%b-%Y"), "%m/%d/%Y")) %>% group_by(Order.date) %>% 
  mutate(Payment.Option = ifelse(Tender == "Cash", "Cash", ifelse(grepl("WeChat", Tender), "AlphaPay", "Clover")), Class = "FBM : CA", MEMO = "Clover sales", Customer = ifelse(Order.Employee.Name == "Garman", "55 JJR SHOPS", "54 JJR SHOPR"), ID = data.table::rleid(Order.ID), REF.ID = paste0("CL", format(as.Date(Order.date, "%m/%d/%Y"), "%y%m%d"), "-", sprintf("%02d", ID)), Order.Type = ifelse(Order.Employee.Name == "Garman", "JJR SHOPS", "JJR SHOPR"), Department = ifelse(Order.Employee.Name == "Garman", "Retail : Store Surrey", "Retail : Store Richmond"), Warehouse = ifelse(Order.Employee.Name == "Garman", "WH-SURREY", "WH-RICHMOND"), Quantity = 1, Price.level = "Custom", Rate = Item.Total, Coupon.Discount = Total.Discount + Order.Discount.Proportion, Coupon.Discount = ifelse(Coupon.Discount == 0, NA, Coupon.Discount), Coupon.Code = gsub("NA", "", gsub(" -.*","", paste0(Order.Discounts, Discounts))), Tax.Code = ifelse(Item.Tax.Rate == 0.05, "CA-BC-GST", ifelse(Item.Tax.Rate == 0.12, "CA-BC-TAX", "")), SHIPPING.CARRIER = "", SHIPPING.METHOD = "", SHIPPING.COST = 0) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, Price.level, Rate, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount, Customer, Recipient, Recipient.Phone, Recipient.Email, SHIPPING.CARRIER, SHIPPING.METHOD, SHIPPING.COST) %>% filter(!is.na(Payment.Option))
write.csv(netsuite_so, file = paste0("../Clover/SO-clover-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")
# upload to NS
update_INV <- T
adjust_inventory <- clover_so %>% filter(Order.Employee.Name == "Garman") %>% group_by(Item.SKU) %>% summarise(Quantity = n()) %>% as.data.frame() %>% `row.names<-`(.[, "Item.SKU"])
price <- woo %>% mutate(cat = gsub("-.*", "", SKU)) %>% group_by(cat) %>% summarise(Price = max(Sale.price)) %>% as.data.frame()
price <- rbind(price, data.frame(cat = c("MISC5", "MISC10", "MISC15", "MISC20", "MISC25", "MISC30", "MISC35", "MISC45", "DBRC", "DBTB", "DBTL", "DBTP", "DLBS", "DWJA", "DWJT", "DWPF", "DWPS", "DWSF", "DWSS", "DXBK", "MAJC"), Price = c(5, 10, 15, 20, 25, 30, 35, 45, 30, 35, 40, 40, 25, 50, 40, 30, 25, 60, 55, 30, 99.99))) %>% `row.names<-`(toupper(.[, "cat"])) 
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
if(update_INV){
  print("Update INV")
  clover_item <- readWorkbook(clover, "Items") %>% mutate(SKU = SKU, Quantity = ifelse(SKU %in% rownames(adjust_inventory), Quantity + adjust_inventory[SKU, "Quantity"], Quantity)) %>% filter(Name != "") %>% 
    mutate(cat = gsub("-.*", "", Name), Price = ifelse(Name %in% woo$SKU, woo[Name, "Sale.price"], price[cat, "Price"]), Price.Type = ifelse(is.na(Price), "Variable", "Fixed"), Alternate.Name = woo[Name, "Name"], Tax.Rates = ifelse(woo[Name, "Tax.class"] == "full", "GST+PST", "GST"), Tax.Rates = ifelse(is.na(Tax.Rates), "GST", Tax.Rates)) %>% select(-cat)
  }else{
  clover_item <- readWorkbook(clover, "Items") %>% filter(Name != "") %>% mutate(cat = gsub("-.*", "", Name), Price = ifelse(Name %in% woo$SKU, woo[Name, "Sale.price"], price[cat, "Price"]), Price.Type = ifelse(is.na(Price), "Variable", "Fixed"), Alternate.Name = woo[Name, "Name"], Tax.Rates = ifelse(woo[Name, "Tax.class"] == "full", "GST+PST", "GST"), Tax.Rates = ifelse(is.na(Tax.Rates), "GST", Tax.Rates)) %>% select(-cat)
}
clover_item <- clover_item %>% mutate(Price = ifelse(grepl("^SS", Name), 10, Price))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item) + 100, gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
# upload to Clover > Inventory
# -------- Add POS sales to yotpo rewards program: daily -----------
# customer info: Clover > Customers > Download
# sales: Clover > Transactions > Payments > select dates > Export
customer <- read.csv("../Clover/Customers.csv", as.is = T) %>% mutate(Name = paste0(First.Name, " ", Last.Name)) %>% distinct(Email.Address, .keep_all = T) %>% filter(Name != " ", Email.Address != "") %>% `row.names<-`(.[, "Name"])
payments <- read.csv(list.files(path = "../Clover/", pattern = paste0("Payments-", format(Sys.Date(), "%Y%m%d")), full.names = T), as.is = T) %>% 
  mutate(email = customer[Customer.Name, "Email.Address"], Refund.Amount = ifelse(is.na(Refund.Amount), 0, as.numeric(Refund.Amount))) %>% filter(!is.na(email), Result == "SUCCESS")
point <- data.frame(Email = payments$email, Points = as.integer(payments$Amount - payments$Refund.Amount)) %>% filter(Points != 0, !is.na(Email)) %>% mutate(Email = gsub(",.*", "", Email), Email = gsub('\\"', "", Email))
write.csv(point, file = paste0("../yotpo/", format(Sys.Date(), "%m%d%Y"), "-yotpo.csv"), row.names = F)
# upload to Yotpo
non_included <- read.csv("../yotpo/non_included.csv", as.is = T)
error <- read.csv("../yotpo/error_report.csv", as.is = T)
non_included <- bind_rows(non_included, error %>% select(Email, Points)) %>% count(Email, wt = Points)
write.table(non_included, file = "../yotpo/non_included.csv", sep = ",", row.names = F, col.names = c("Email", "Points"))

# ------------- upload XHS SO to NS, sync XHS price and inventory: weekly ---------------------------
XHS_so <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../XHS/", pattern = "order_export", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = 1) %>%
  mutate(Shipping.Address2 = replace_na(Shipping.Address2, ""), Shipping.Address1 = paste(Shipping.Address1, Shipping.Address2), Shipping.Country = gsub("加拿大", "CA", gsub("美国", "US", Shipping.Country)), Shipping.Province = gsub("不列颠哥伦比亚", "BC", Shipping.Province))
netsuite_so <- XHS_so %>% filter(Financial.Status == "paid", Shipping.Country == "中国") %>% mutate(Order.date = format(as.Date(gsub(" .*", "", Created.At), "%Y-%m-%d"), "%m/%d/%Y")) %>% 
  mutate(Payment.Option = "AlphaPay", Class = "FBM : CN", MEMO = "XHS sales", Customer = "56 XHS CN", Order.ID = Order.No, REF.ID = paste0("XHSCN-", Order.No), Order.Type = "XHS CN", Department = "Retail : Marketplace : XiaoHongShu", Warehouse = "WH-RICHMOND", Item.SKU = LineItem.SKU, Quantity = as.numeric(LineItem.Quantity), Price.level = "Custom", Rate = round((as.numeric(LineItem.Total) - as.numeric(LineItem.Total.Discount))/Quantity, 2), Coupon.Discount = as.numeric(LineItem.Total.Discount), Coupon.Discount = ifelse(Coupon.Discount == 0, "", Coupon.Discount), Coupon.Code = "", Tax.Amount = "", Tax.Code = "", Recipient = paste0(Shipping.Last.Name, Shipping.First.Name), Recipient.Phone = Phone, Recipient.Email = Email, SHIPPING.CARRIER = "Longxing", SHIPPING.METHOD = "", SHIPPING.COST = Shipping) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, Price.level, Rate, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount, Customer, Recipient, Recipient.Phone, Recipient.Email, SHIPPING.CARRIER, SHIPPING.METHOD, SHIPPING.COST)
Encoding(netsuite_so$Recipient) = "UTF-8"
write_excel_csv(netsuite_so, file = paste0("../XHS/CashSale-XHS-", format(Sys.Date(), "%Y%m%d"), ".csv"), na = "")
HSTonly_states <- c("NB", "NL")
HST_states <- c("NS", "ON", "PE")
GSTonly_states <- c("AB", "NT", "NU", "YT")
PST_cats <- c("AAA", "ACA", "ACB", "AHJ", "AJA", "AJC", "AJP", "AJS", "AWWJ", "DRC", "GBX", "GHA", "GUA", "GUX", "XBK", "XBM", "XLB", "XPC")
netsuite_so <- XHS_so %>% filter(Financial.Status == "paid", Shipping.Country != "中国") %>% 
  mutate(Order.Type.Code = paste("XHS", Shipping.Country), Order.ID = paste0("XHS", Shipping.Country, "-", Order.No), Order.Date = format(as.Date(gsub(" .*", "", Created.At), "%Y-%m-%d"), "%m/%d/%Y"), Shipping.Name = ifelse(Delivery.Method == "pickup", Pickup.Person.Name, paste(Shipping.First.Name, Shipping.Last.Name)), Cat = gsub("-.*", "", LineItem.SKU)) %>% 
  mutate(Customer.ID = ifelse(grepl("CA", Order.Type.Code), "57 XHS CA", "28876 XHS US"), Ref.Number = Order.ID,	Date = format(Sys.Date(), "%m/%d/%Y"), Sales.Effective.Date = Order.Date, Ship.By.Date = format(as.Date(Date, "%m/%d/%Y") + 1, "%m/%d/%Y"),	Price.Level = "Custom",	SKU = LineItem.SKU,	Unit.Price = LineItem.UnitPrice, QTY = LineItem.Quantity, Amount = as.numeric(LineItem.Total) - as.numeric(LineItem.Total.Discount), Tax.Code = ifelse(Shipping.Country == "US", ".", ifelse(Shipping.Province %in% GSTonly_states, paste0("CA-", Shipping.Province, "-GST"), ifelse(Shipping.Province %in% HSTonly_states, paste0("CA-", Shipping.Province, "-HST"), ifelse(Shipping.Province %in% HST_states, ifelse(Cat %in% PST_cats, paste0("CA-", Shipping.Province, "-HST"), paste0("CA-", Shipping.Province, "-GST")), ifelse(Cat %in% PST_cats, paste0("CA-", Shipping.Province, "-TAX"), paste0("CA-", Shipping.Province, "-GST")))))), Coupon_Discount = LineItem.Total.Discount, Order.Coupon.Code = LineItem.Promotion.Detail, Shipping.cost = Shipping, Shipping.Tax.Code = ".", Woocommerce.Order.Total = Total, Commit = "Available QTY",	Marketplace.PAYMENT.METHOD = Payment.Method, Payment.Option = ifelse(Payment.Method == "Stripe", "Stripe CAD Allvalue", "AlphaPay"), Terms = "CIA", To.Be.Emailed = "No", Shipping.Postal.Code = Shipping.Zip, Billing.Name = Shipping.Name, Billing.Address1 = Shipping.Address1, Billing.City = Shipping.City, Billing.Province = Shipping.Province, Billing.Country = Shipping.Country, Billing.Postal.Code = Shipping.Zip, Department = "Retail : Marketplace : XiaoHongShu", Class = paste("FBM :", Shipping.Country),	Warehouse = "WH-SURREY", Recipient = Shipping.Name, Recipient.Email = Email, Recipient.Phone = Phone, Marketplace.Shipping.Method = ifelse(Delivery.Method != "pickup", ifelse(Shipping == 0, "Free shipping over $69", "Parcel with Tracking"), ifelse(Shipping.City == "Vancouver", "Pickup at UBC Vancouver Campus", ifelse(Shipping.City == "Richmond", "Pickup at Office in Richmond, BC", "Pickup at Warehouse in Surrey, BC"))), Shipstation.Carrier = ifelse(Shipping.Country == "CA", "Canada Post", "USPS"),	Shipstation.Shipping.Service = ifelse(Shipping.Country == "CA", "Expedited Parcel", "USPS Priority Mail"), SHIPSTATION.SERVICE = ifelse(Delivery.Method != "pickup", "Standard", ifelse(Shipping.City == "Vancouver", "UBC Pickup", paste(Shipping.City, "Pickup"))), Order.Issue = "", Memo = "XHS order") %>%
  select(Customer.ID, Order.Type.Code, Order.ID, Ref.Number, Date, Order.Date, Sales.Effective.Date, Ship.By.Date, Price.Level, SKU, Unit.Price, QTY, Amount, Tax.Code, Coupon_Discount, Order.Coupon.Code, Shipping.cost, Shipping.Tax.Code, Woocommerce.Order.Total, Marketplace.Shipping.Method, Commit, Marketplace.PAYMENT.METHOD, Payment.Option, Terms, To.Be.Emailed, Shipping.Name, Shipping.Address1, Shipping.City, Shipping.Province, Shipping.Country, Shipping.Postal.Code, Billing.Name, Billing.Address1, Billing.City, Billing.Province, Billing.Country, Billing.Postal.Code, Department, Class, Warehouse, Recipient, Recipient.Email, Recipient.Phone, Shipstation.Carrier, Shipstation.Shipping.Service, SHIPSTATION.SERVICE, Order.Issue, Memo)
netsuite_so <- netsuite_so %>% rename_with(~ gsub("\\.", " ", colnames(netsuite_so)))
write_excel_csv(netsuite_so, file = paste0("../XHS/SO-XHS-", format(Sys.Date(), "%Y%m%d"), ".csv"), na = "")
# upload to NS
# input: AllValue > Products > Export > All products
products_description <- openxlsx::read.xlsx("../XHS/products_description.xlsx", sheet = 1) %>% `row.names<-`(toupper(.[, "SPU"])) 
wb <- openxlsx::loadWorkbook(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T))
openxlsx::protectWorkbook(wb, protect = F)
openxlsx::saveWorkbook(wb, list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), overwrite = T)
products_XHS <- read.xlsx2(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), sheetIndex = 1)
products_XHS[products_XHS=="NA"] <- ""
products_upload <- products_XHS %>% mutate(Inventory = ifelse(netsuite_item_S[SKU, "Warehouse.Available"] < 5, 0, netsuite_item_S[SKU, "Warehouse.Available"]), Price = ifelse(grepl("^L", SKU), woo[SKU, "Sale.price"] + 10, woo[SKU, "Sale.price"]), Compare.At.Price = ifelse(grepl("^L", SKU), woo[SKU, "Regular.price"] + 10, woo[SKU, "Regular.price"]), Tags = ifelse(Price == Compare.At.Price, "正价", "特价"))
products_upload <- products_upload %>% mutate(Inventory = ifelse(grepl("^MWPF", SKU), 0, Inventory))
products_upload <- products_upload %>% mutate(Product.Name = products_description[toupper(SPU), "Product.Name"], SEO.Product.Name = Product.Name, Description = products_description[toupper(SPU), "Description"], Mobile.Description = products_description[toupper(SPU), "Description"], SEO.Description = products_description[toupper(SPU), "Description"], Describe = products_description[toupper(SPU), "Description"])
colnames(products_upload) <- gsub("\\.", " ", colnames(products_upload))
openxlsx::write.xlsx(products_upload, file = paste0("../XHS/products_upload-", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"), na.string = "")
# upload to AllValue > Products

# ------------- upload Richmond returns/exchanges to NS: weekly ---------------------------
return <- openxlsx::read.xlsx("../../TWK 2020 share/twk general/2 - show room operations/0-Showroom records.xlsx", sheet = 1, detectDates = T, fillMergedCells = T) %>% filter(is.na(NS.Action)) %>% 
  mutate(Date = format(Sys.Date(), "%m/%d/%Y"), ITEM = `Return.Item.SKU.(Inv.in)`, ADJUST.QTY.BY = Return.Qty, Reason.Code = ifelse(is.na(Exchange.Qty), "RT", "EXC"), MEMO = paste0(Original.Order.ID, " | ", Order.Channel, " | ", Return.Reason, " | ", Notes), WAREHOUSE = "WH-RICHMOND", External.ID = paste0("IAR-", format(Sys.Date(), "%y%m%d"), "-1")) %>%
  select(Date, ITEM, ADJUST.QTY.BY, Reason.Code, MEMO, WAREHOUSE, External.ID) %>% filter(ADJUST.QTY.BY != 0)
colnames(return) <- gsub("QTY BY", "QTY. BY", gsub("\\.", " ", colnames(return)))
write.csv(return, file = paste0("../Clover/RT-Richmond-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")

# ------------- Sync NS Richmond inventory with Clover: weekly ---------------------------
clover_item <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items") %>% filter(Name %in% netsuite_item_R$Name) %>% mutate(Quantity = ifelse(Quantity < 0, 0, Quantity))
inventory_update <- clover_item %>% mutate(Status = "Good", Date = format(Sys.Date(), "%m/%d/%Y"), ITEM = Name, ADJUST.QTY.BY = ifelse(Name %in% netsuite_item_R$Name, Quantity - netsuite_item_R[Name, "Warehouse.On.Hand"], Quantity), Reason.Code = "CC", MEMO = "Clover Inventory Update", WAREHOUSE = "WH-RICHMOND", External.ID = paste0("IAR-", format(Sys.Date(), "%y%m%d"), "-2")) %>%
  select(Date, ITEM, ADJUST.QTY.BY, Reason.Code, Status, MEMO, WAREHOUSE, External.ID) %>% filter(ADJUST.QTY.BY != 0)
colnames(inventory_update) <- gsub("QTY BY", "QTY. BY", gsub("\\.", " ", colnames(inventory_update)))
write.csv(inventory_update, file = paste0("../NetSuite/IA-Richmond-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")
## if there are unsuccessful Clover SO upload due to NS qty=0 
SO_error <- read.csv("../Clover/results.csv", as.is = T) %>% group_by(Item.SKU) %>% summarise(Qty = sum(Quantity)) %>% as.data.frame() %>% `row.names<-`(.[, "Item.SKU"])
clover_item <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items") %>% filter(Name %in% netsuite_item$Name) %>% mutate(Quantity = ifelse(Quantity < 0, 0, Quantity))
inventory_update <- clover_item %>% mutate(Status = "Good", Date = format(Sys.Date(), "%m/%d/%Y"), ITEM = Name, ADJUST.QTY.BY = ifelse(Name %in% netsuite_item_R$Name, Quantity - netsuite_item_R[Name, "Warehouse.On.Hand"], Quantity), ADJUST.QTY.BY = ifelse((ITEM %in% rownames(SO_error)) & (netsuite_item_R[Name, "Warehouse.On.Hand"] < SO_error[ITEM, "Qty"]), ADJUST.QTY.BY + SO_error[ITEM, "Qty"] - netsuite_item_R[Name, "Warehouse.On.Hand"], ADJUST.QTY.BY), Reason.Code = "CC", MEMO = "Clover Inventory Update", WAREHOUSE = "WH-RICHMOND", External.ID = paste0("IAR-", format(Sys.Date(), "%y%m%d"), "-2")) %>%
  select(Date, ITEM, ADJUST.QTY.BY, Reason.Code, Status, MEMO, WAREHOUSE, External.ID) %>% filter(ADJUST.QTY.BY != 0)
colnames(inventory_update) <- gsub("QTY BY", "QTY. BY", gsub("\\.", " ", colnames(inventory_update)))
write.csv(inventory_update, file = paste0("../NetSuite/IAR-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")

# -------- Request stock from Surrey: at request --------------------
# input Clover > Inventory > Items > Export.
netsuite_item <- read.csv(list.files(path = "../NetSuite/", pattern = paste0("Items_All_", format(Sys.Date(), "%Y%m%d"), ".csv"), full.names = T), as.is = T)
netsuite_item[netsuite_item == "" | is.na(netsuite_item)] <- 0
netsuite_item_S <- netsuite_item %>% filter(Inventory.Warehouse == "WH-SURREY") %>% `row.names<-`(.[, "Name"])
season <- "25F"
#request <- c("WSS", "WSF", "WJT", "WJA", "WPS", "WPF", "WBS", "WBF", "BTL", "BRC", "SWS")
request <- c("AJA", "AJC", "AJM", "BCV", "BSW", "BSA", "BRC", "BST", "BSL", "BTB", "BTL", "BTT", "BST", "FAN", "FHA", "FJC", "FJM", "FPM", "FMR", "FSM", "FVM", "IHT", "IPC", "ICP", "IPS", "ISJ", "ISS", "ISB", "KEH", "KHB", "KHM", "KHP", "KMN", "KMT", "LAB", "LAN", "LBT", "LBP", "LCP", "LCT", "SKG", "SKB", "SKX", "SMC", "SMF", "SWS", "WJA", "WJT", "WPF", "WPS", "WBF", "WJO", "WPO", "WHO", "WHR", "WBS", "WGF", "WGS", "WMT", "WRM", "WSF", "WSS", "XBK", "XBM", "XLB", "XPC") # categories to restock for FW
#request <- c("SWS", "SMC", "SBS", "SMF", "XBM", "XBK", "BRC", "SKG", "SKB", "SKX", "SJD", "SJF", "SPW", "LBT", "LBP", "HAV0", "HCA0", "HCB0", "HAD0", "HCF0", "HXP", "HXU", "HXC", "HBS", "HBU", "HLC", "HLH", "GUA", "GUX", "GHA", "GBX", "UG1", "UJ1", "USA", "UT1", "UV2", "USS", "UST") # categories to restock for SS
cat <- c("WJA", "WPF", "WPS", "XBK", "XBM", "XLB", "SWS", "BRC", "BTB", "BTL", "FHA", "IHT", "KHB", "KHP", "WHR")
size <- c("2T", "3T", "4T", "5T", "6Y", "OS", "24", "25", "26", "27", "28", "29", "8", "9", "10", "11", "12", "13", "L", "XL")
n <- 3       # Qty per SKU to stock at Richmond
n1 <- 3   # Qty for SKUs in cat & size to refill
n_S <- 8 # min Qty in stock at Surrey to request
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- readWorkbook(clover, "Items") %>% mutate(cat = gsub("-.*", "", Name), Quantity = ifelse(is.na(Quantity) | Quantity < 0, 0, Quantity)) %>% filter(!duplicated(Name), !is.na(Name)) %>% `row.names<-`(.[, "Name"])
# order <- data.frame(StoreCode = "WH-JJ", ItemNumber=(clover_item %>% filter(Quantity < n & cat %in% request))$Name, Qty = n - clover_item[(clover_item %>% filter(Quantity < n & cat %in% request))$Name, "Quantity"], LocationName = "BIN", UnitCost = "", ReasonCode = "RWT", Memo = "Richmond Transfer to Miranda", UploadRule = "D", AdjAccntName = "", TxnDate = "", ItemIdentifierCode = "", ImportError = "")
# order <- order %>% filter(xoro[order$ItemNumber, "ATS"] > n_xoro) %>% filter(Qty > 0) %>% mutate(cat = gsub("-.*", "", ItemNumber), size = gsub("\\w+-\\w+-", "", ItemNumber)) %>% arrange(cat, size) %>% select(-c("cat", "size"))
order <- data.frame(Date = format(Sys.Date(), "%m/%d/%Y"), TO.TYPE = "Surrey-Richmond", SEASON = season, FROM.WAREHOUSE = "WH-SURREY", TO.WAREHOUSE = "WH-RICHMOND", REF.NO = paste0("TO-S2R", format(Sys.Date(), "%y%m%d")), Memo = "Richmond Refill", ORDER.PLACED.BY = "Gloria Li", ITEM = (clover_item %>% filter(Quantity < n))$Name) %>% 
  mutate(CAT = netsuite_item_S[ITEM, "Item.Category.SKU"], SIZE = netsuite_item_S[ITEM, "Item.Size"], Seasons = netsuite_item_S[ITEM, "Item.SKU.Seasons"], Quantity = ifelse(CAT %in% cat & SIZE %in% size & grepl(season, Seasons), n1 - clover_item[ITEM, "Quantity"], n - clover_item[ITEM, "Quantity"])) %>% filter(ITEM %in% netsuite_item_S$Name, netsuite_item_S[ITEM, "Warehouse.Available"] > n_S) %>% filter(Quantity > 1) %>% mutate(cat = gsub("-.*", "", ITEM), size = gsub("\\w+-\\w+-", "", ITEM))
order <- order %>% filter(cat %in% request) %>% select(-CAT, -SIZE, -Seasons)
order <- order %>% rename_with(~ gsub("\\.", " ", colnames(order))) %>% select(-c("cat", "size")) %>% arrange(ITEM) %>% mutate(stock = clover_item[ITEM, "Quantity"])
write.csv(order, file = paste0("../Clover/order", format(Sys.Date(), "%m%d%Y"), ".csv"), row.names = F, na = "")
# upload to NS & email Shikshit

# ---------------- Check and update AMZPrep shipment orders ---------------
synced <- openxlsx::read.xlsx("../AMZPrep/SyncedOrder.xlsx", sheet = 1)
shipped <- openxlsx::read.xlsx("../AMZPrep/ShippedOrder.xlsx", sheet = 1)
records <- openxlsx::loadWorkbook("../../TWK Order Process/01 - shipping tracking/3-AMZPrep Order Tracking - 20251223.xlsx")
records_AMZPrep <- readWorkbook(records, "Shipments", detectDates = T) %>% mutate(eShipper.Status = ifelse(Order.ID %in% synced$Extra.Note.1, "Synced", ifelse(Order.ID %in% shipped$Extra.Note.1, "Shipped", eShipper.Status)))
new_synced <- synced %>% filter(!(Extra.Note.1 %in% records_AMZPrep$Order.ID)) %>% 
  mutate(Date = gsub(" .*", "", Shipment.Order.Date), Order.ID = Extra.Note.1, Shipped.from.eShipper = "Whole", Province = `Billing.State.Code/Region`, Total.Qty = Total.Quantity, eShipper.Qty = Total.Quantity, Auto.FR = "Y", Shipping.Required = ifelse(Shipping.Service == "Cheapest Service", "", "Express"), FR.Created = "Y", eShipper.Status = "Synced", Items = ifelse(Total.Quantity == 1, SKU, "")) %>%
  select(Date:Items) %>% distinct(Order.ID, .keep_all = T)
new_shipped <- shipped %>% filter(!(Extra.Note.1 %in% records_AMZPrep$Order.ID)) %>% 
  mutate(Date = gsub(" .*", "", Shipment.Order.Date), Order.ID = Extra.Note.1, Shipped.from.eShipper = "Whole", Province = `Billing.State.Code/Region`, Total.Qty = Total.Quantity, eShipper.Qty = Total.Quantity, Auto.FR = "Y", Shipping.Required = ifelse(Shipping.Service == "Cheapest Service", "", "Express"), FR.Created = "Y", eShipper.Status = "Shipped", Items = ifelse(Total.Quantity == 1, SKU, "")) %>%
  select(Date:Items) %>% distinct(Order.ID, .keep_all = T)
records_AMZPrep <- rbind(records_AMZPrep, new_synced, new_shipped)
records_AMZPrep <- records_AMZPrep %>% rename_with(~ gsub("\\.", " ", colnames(records_AMZPrep)))
deleteData(records, sheet = "Shipments", cols = 1:ncol(records_AMZPrep), rows = 1:nrow(records_AMZPrep), gridExpand = T)
writeData(records, sheet = "Shipments", records_AMZPrep)
openxlsx::saveWorkbook(records, file = "../../TWK Order Process/01 - shipping tracking/3-AMZPrep Order Tracking - 20251223.xlsx", overwrite = T)
duplicated <- rbind(synced, shipped) %>% mutate(SO = gsub("-.*", "", Order.Code), FR = gsub(".*-", "", Order.Code)) %>% group_by(SO) %>% filter(n_distinct(FR) > 1) %>% ungroup() %>% 
  select(Order.Code:Status) %>% distinct(Order.Code, .keep_all = T) %>% arrange(Order.Code)

# ---------------- Adjust Clover Inventory: at request --------------------
# download current Richmond stock: clover_item > Inventory > Items > Export
library(dplyr)
library(openxlsx)
adjust_inventory <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "order08282025.csv", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% `row.names<-`(.[, "ITEM"])
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- readWorkbook(clover, "Items") %>% mutate(SKU = SKU, Quantity = ifelse(SKU %in% rownames(adjust_inventory), Quantity + adjust_inventory[SKU, "Quantity"], Quantity))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item), gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
# upload to Clover > Inventory

# -------- Analysis: monthly --------------------
new_season <- "25"   # New season contains (yr)
qty_offline <- 1     # Qty to move to offline sales
month <- format(Sys.Date(), "%m")
if(month %in% c("09", "10", "11", "12", "01", "02")){in_season <- "F"}else(in_season <- "S")
RawData <- list.files(path = "../FBArefill/Raw Data File/", pattern = "All Marketplace by", full.names = T)
mastersku <- openxlsx::read.xlsx(list.files(path = "../FBArefill/Raw Data File/", pattern = "1-MasterSKU-All-Product-", full.names = T)[1], sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(.[, "MSKU"])
mastersku_adjust <- mastersku %>% filter(!duplicated(Adjust.SKU)) %>% `row.names<-`(.[, "Adjust.SKU"])
netsuite_item_S <- read.csv(list.files(path = "../NetSuite/", pattern = paste0("Items_All_", format(Sys.Date(), "%Y%m%d"), ".csv"), full.names = T), as.is = T) %>% filter(Inventory.Warehouse == "WH-SURREY") %>% 
  mutate(Name = Name, Seasons = ifelse(Name %in% mastersku$MSKU, mastersku[Name, "Seasons.SKU"], mastersku_adjust[Name, "Seasons.SKU"]), SPU = gsub("(\\w+-\\w+)-.*", "\\1", Name)) %>% `row.names<-`(.[, "Name"])
## Raw data summary
monR_last_yr <- read.csv(list.files(path = "../Analysis/", pattern = "MonthlyRatio2023_category_adjust_2024-01-25.csv", full.names = T), as.is = T) %>% `row.names<-`(.[, "Category"])
sheets <- grep("-WK", getSheetNames(RawData), value = T)
inventory <- data.frame()
for(sheet in sheets){
  inventory_i <- openxlsx::read.xlsx(RawData, sheet = sheet, startRow = 2)
  inventory <- rbind(inventory, inventory_i)
}
inventory[is.na(inventory)] <- 0
#inventory <- inventory %>% mutate(Inv_Total_End.WK = `Inv_Total_End.WK.(Not.Inc..CN.RM)` + Inv_WH.China + Inv_WH.Richmond)
write.csv(inventory, file = paste0("../Analysis/Sales_Inventory_SKU_", Sys.Date(), ".csv"), row.names = F, na = "")
names <- c(Category = "Cat_SKU", Month01="Jan", Month02="Feb", Month03="Mar", Month04="Apr", Month05="May", Month06="June", Month07="July", Month08="Aug", Month09="Sep", Month10="Oct", Month11="Nov", Month12="Dec")
sheets <- grep("-SKU$", getSheetNames(RawData), value = T)
if(month == "01"){columns <- c(1:11, 26:37)}else{columns <- c(1:11, 13:(as.numeric(month)-1+12), (25+as.numeric(month)):37)}
sales_cat_last12month <- data.frame()
sales_SKU_last12month <- data.frame()
for(sheet in sheets){
  sales_last12month_i <- openxlsx::read.xlsx(RawData, sheet = sheet, cols = columns, startRow = 2) 
  if(sum(grepl("janandjul.com", sales_last12month_i[[1]]))){
    sales_cat_last12month_i <- sales_last12month_i %>% filter(.[[1]] == "All Marketplaces") %>% mutate_all(~ replace(., is.na(.), 0)) %>% rename_with(~ gsub("20", "Total.20", gsub("_x000D_.Total", "", .x)))
    sales_cat_last12month <- rbind(sales_cat_last12month, sales_cat_last12month_i)
    sales_SKU_last12month_i <- openxlsx::read.xlsx(RawData, sheet = sheet, cols = columns, startRow = 2) %>% 
      filter(!(Sales.Channel %in% c("Summary", "Sales Channel", "All Marketplaces"))) %>% mutate_all(~ replace(., is.na(.), 0)) %>% rename_with(~ gsub("20", "Total.20", gsub("_x000D_.Total", "", .x)))
    sales_SKU_last12month <- rbind(sales_SKU_last12month, sales_SKU_last12month_i)
  }
}
sales_SKU_last12month <- sales_SKU_last12month %>% rename(all_of(names)) %>% select(Sales.Channel, Adjust_MSKU, Month01:Month12) %>% mutate(across(Month01:Month12, ~as.numeric(.)))
write.csv(sales_SKU_last12month, file = paste0("../Analysis/Sales_SKU_last12month_", Sys.Date(), ".csv"), row.names = F)
if(month %in% c("01", "02")){
  monR <- monR_last_yr %>% select(Category, Month01:Month12) %>% `row.names<-`(toupper(.[, "Category"]))
}else{
  monR_last_yr <- monR_last_yr %>% rowwise() %>% mutate(T.new = sum(c_across(Month01:as.name(sprintf("Month%02d", (as.numeric(month)-1)))))) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "Category"]))
  sales_cat_last12month <- sales_cat_last12month %>% rename(all_of(names)) %>% mutate(across(Month01:Month12, ~as.numeric(.))) %>% rowwise() %>% mutate(T.new = sum(c_across(Month01:as.name(sprintf("Month%02d", (as.numeric(month)-1))))), MonR.new = monR_last_yr[Category, "T.new"]) %>% 
    filter(!is.na(MonR.new)) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "Category"]))
  monR_new <- sales_cat_last12month %>% mutate(across(Month01:as.name(sprintf("Month%02d", (as.numeric(month)-1))), ~ ./T.new*MonR.new))
  monR_new[monR_new$T.new == 0, ] <- monR_new[monR_new$T.new == 0, ] %>% rows_update(., monR_last_yr %>% select(Category, Month01:as.name(sprintf("Month%02d", (as.numeric(month)-1)))), by = "Category", unmatched = "ignore")
  monR <- merge(monR_new %>% select(Category, Month01:as.name(sprintf("Month%02d", (as.numeric(month)-1)))), monR_last_yr %>% select(Category, as.name(sprintf("Month%02d", as.numeric(month))):Month12), by = "Category") %>% `row.names<-`(toupper(.[, "Category"]))
}
write.csv(monR, file = paste0("../Analysis/MonthlyRatio_", Sys.Date(), ".csv"), row.names = F)
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\
## SKUs to inactive
clover <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items")  %>% filter(!is.na(Name)) %>% `row.names<-`(.[, "Name"])
inventory_s <- inventory %>% count(Adjust_MSKU, wt = Inv_Total_End.WK) %>% `row.names<-`(.[, "Adjust_MSKU"])
qty0 <- mastersku %>% select(MSKU.Status, Seasons.SKU, MSKU) %>% mutate(MSKU = MSKU, Inv_clover = clover[MSKU, "Quantity"], Inv_Total_EndWK = inventory_s[MSKU, "n"]) %>% 
  filter(!grepl(new_season, Seasons.SKU) & MSKU.Status == "Active" & Inv_clover == 0 & Inv_Total_EndWK == 0)
write.csv(qty0, file = paste0("../Analysis/discontinued_qty0_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\
## Suggested deals for the website
discount_method <- openxlsx::read.xlsx("../Analysis/JJ_discount_method.xlsx", sheet = 1, startRow = 2, rowNames = T)
if(month == "01"){last3m <- c("Month11", "Month12", "Month01")}else if(month %in% c("02", "03")){last3m <- sprintf("Month%02d", c(1:(as.numeric(month)-1), (as.numeric(month)-3+12):12))}else{last3m <- sprintf("Month%02d", c((as.numeric(month)-3):(as.numeric(month)-1)))}
monR <- monR %>% mutate(T.last3m = rowSums(across(all_of(last3m)))) %>% `row.names<-`(.[, "Category"])
inventory_SPU <- inventory %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU)) %>% group_by(SPU) %>% summarise(qty_SPU = sum(Inv_Total_End.WK, na.rm = T), qty_WH = sum(Inv_WH.Surrey, na.rm = T), qty_AMZ = qty_SPU - qty_WH) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SKU_last12month <- sales_SKU_last12month %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU), T.last3m = rowSums(across(all_of(last3m))), T.yr = rowSums(across(Month01:Month12)))
sales_SPU_last3month <- sales_SKU_last12month %>% group_by(SPU) %>% summarise(T.last3m = sum(T.last3m)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SPU_last3month_JJ <- sales_SKU_last12month %>% filter(Sales.Channel == "janandjul.com") %>% group_by(SPU) %>% summarise(T.last3m = sum(T.last3m)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SPU_last12month_AMZ <- sales_SKU_last12month %>% filter(grepl("Amazon", Sales.Channel)) %>% group_by(SPU) %>% summarise(T.last12m = sum(T.yr)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sizes_all <- netsuite_item_S %>% mutate(SPU_size = gsub("(\\w+-\\w+-\\w+).*", "\\1", Name)) %>% filter(!duplicated(SPU_size)) %>% count(SPU) %>% `row.names<-`(.[, "SPU"])
size_limited <- netsuite_item_S %>% filter(Warehouse.Available > qty_offline) %>% count(SPU) %>% mutate(all = sizes_all[SPU, "n"], percent = ifelse(all < n, 0, (all - n)/all)) %>% `row.names<-`(.[, "SPU"])
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% filter(Published == 1, In.stock. == 1, !is.na(Regular.price)) %>% filter(!duplicated(SKU), SKU != "") %>% 
  mutate(SKU = SKU, Qty = netsuite_item_S[SKU, "Warehouse.Available"], Seasons = ifelse(SKU %in% mastersku$MSKU, mastersku[SKU, "Seasons.SKU"], mastersku_adjust[SKU, "Seasons.SKU"]), discount = ifelse(is.na(Sale.price), 0, (Regular.price - Sale.price)/Regular.price), cat = mastersku[SKU, "Category.SKU"], SPU = gsub("(\\w+-\\w+)-.*", "\\1", SKU)) %>% 
  select(ID, SKU, Name, Seasons, cat, SPU, Regular.price, discount, Qty) %>% `row.names<-`(.[, "SKU"])
overstock <- woo %>% filter(!grepl(new_season, Seasons), Qty > qty_offline, Regular.price > 0) %>% mutate(Qty_SPU = inventory_SPU[SPU, "qty_SPU"], Qty_WH = inventory_SPU[SPU, "qty_WH"], Qty_AMZ = inventory_SPU[SPU, "qty_AMZ"], MonR_last3m = monR[cat, "T.last3m"], Sales_SPU_last3m = sales_SPU_last3month[SPU, "T.last3m"], Sales_SPU_last3m_JJ = sales_SPU_last3month_JJ[SPU, "T.last3m"], Sales_SPU_1yr = Sales_SPU_last3m/MonR_last3m, Sales_SPU_1yr_JJ = Sales_SPU_last3m_JJ/MonR_last3m, Sales_SPU_1yr_AMZ = sales_SPU_last12month_AMZ[SPU, "T.last12m"]) %>% 
  filter(!is.na(Qty_SPU)) %>% mutate(time_yrs = round(ifelse(Qty_AMZ >= Sales_SPU_1yr_AMZ, Qty_WH/Sales_SPU_1yr_JJ, Qty_SPU/Sales_SPU_1yr), 1), size_percent_missing = gsub("Sizes.missing.0%", "Sizes.missing.<.50%", paste0("Sizes.missing.", as.character(as.integer(ifelse(size_limited[SPU, "percent"] < 0.5, 0, size_limited[SPU, "percent"])*10)*10), "%"))) %>% 
  select(SKU:Qty, time_yrs, size_percent_missing) %>% arrange(SKU)
write.csv(overstock, file = paste0("../Analysis/Overstock_", Sys.Date(), ".csv"), row.names = F)
# Email Will
woo_deals <- woo %>% filter(!grepl(new_season, Seasons), Qty > qty_offline, Regular.price > 0) %>% mutate(Qty_SPU = inventory_SPU[SPU, "qty_SPU"], Qty_WH = inventory_SPU[SPU, "qty_WH"], Qty_AMZ = inventory_SPU[SPU, "qty_AMZ"], MonR_last3m = monR[cat, "T.last3m"], Sales_SPU_last3m = sales_SPU_last3month[SPU, "T.last3m"], Sales_SPU_last3m_JJ = sales_SPU_last3month_JJ[SPU, "T.last3m"], Sales_SPU_1yr = Sales_SPU_last3m/MonR_last3m, Sales_SPU_1yr_JJ = Sales_SPU_last3m_JJ/MonR_last3m, Sales_SPU_1yr_AMZ = sales_SPU_last12month_AMZ[SPU, "T.last12m"]) %>% 
  filter(!is.na(Qty_SPU)) %>% mutate(time_yr = ifelse(Qty_AMZ >= Sales_SPU_1yr_AMZ, Qty_WH/Sales_SPU_1yr_JJ, Qty_SPU/Sales_SPU_1yr), time_yrs = paste0("Sold out ", as.character(as.integer(ifelse(time_yr > 3, 3, time_yr)) + 1), " yr"), size_percent_missing = gsub("Sizes.missing.0%", "Sizes.missing.<.50%", paste0("Sizes.missing.", as.character(as.integer(ifelse(size_limited[SPU, "percent"] < 0.5, 0, size_limited[SPU, "percent"])*10)*10), "%"))) %>%
  rowwise() %>% mutate(Suggest_discount = as.numeric(gsub("% Off", "", discount_method[time_yrs, size_percent_missing]))/100, Suggest_price = round((1 - Suggest_discount) * Regular.price, digits = 2)) %>% arrange(-Suggest_discount, SKU) %>% filter(Suggest_discount > 0) 
if(month %in% c("02", "03", "04", "05")){woo_deals <- woo_deals %>% filter(!grepl("S", Seasons))}
if(month %in% c("08", "09", "10", "11")){woo_deals <- woo_deals %>% filter(!grepl("F", Seasons))}
write.csv(woo_deals, file = paste0("../Analysis/Deals_", Sys.Date(), ".csv"), row.names = F)
# Email Joren, Kamer
## Low inventory to adjust ads
qty_refill <- 12
inventory <- inventory %>% mutate(Inv_WH_AMZ.CA.US = Inv_WH.JJ + Inv_AMZ.USA + Inv_AMZ.CA) %>% `row.names<-`(.[, "Adjust_MSKU"])
Sales_Inv_ThisMonth <- sales_SKU_last12month %>% filter(Sales.Channel %in% c("Amazon.com", "Amazon.ca", "janandjul.com")) %>% group_by(Adjust_MSKU) %>% summarise(Sales.last3m = sum(T.last3m)) %>% filter(Adjust_MSKU %in% inventory$Adjust_MSKU) %>%
  mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU), cat = mastersku[Adjust_MSKU, "Category.SKU"], Seasons = mastersku[Adjust_MSKU, "Seasons.SKU"], Status = mastersku[Adjust_MSKU, "MSKU.Status"], AMZ_flag = mastersku[Adjust_MSKU, "Pause.Plan.FBA_x000D_.CA/US/MX/AU/UK/DE"], monR_last3m = monR[cat, "T.last3m"], monR_this = monR[cat, paste0("Month", month)], Sales.this = Sales.last3m/monR_last3m*monR_this, Inv_WH_AMZ.CA.US = inventory[Adjust_MSKU, "Inv_WH_AMZ.CA.US"], Qty_left = ifelse(Inv_WH_AMZ.CA.US < Sales.this, 0, as.integer(Inv_WH_AMZ.CA.US - Sales.this)), Enough_Inv = (Qty_left >= qty_refill | Sales.this == 0))
Sales_Inv_ThisMonth_SPU <- Sales_Inv_ThisMonth %>% group_by(SPU) %>% summarise(Enough_Inv = mean(Enough_Inv)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
Sales_Inv_ThisMonth <- Sales_Inv_ThisMonth %>% mutate(Percent_Sizes_SPU_Enough_Inv = paste0(as.character(as.integer(Sales_Inv_ThisMonth_SPU[SPU, "Enough_Inv"]*100)), "%")) %>% filter(grepl(in_season, Seasons), Status == "Active", !Enough_Inv) %>% arrange(Qty_left, Adjust_MSKU)
write.csv(Sales_Inv_ThisMonth, file = paste0("../Analysis/Sales_Inv_Prediction_", Sys.Date(), ".csv"), row.names = F)
# Email Adrienne, Joren and Kamer 
## Low inventory to add PO
if(month == "12"){next4m <- c("Month01", "Month02", "Month03", "Month04")}else if(month %in% c("10", "11")){next4m <- sprintf("Month%02d", c(1:(as.numeric(month)-12+3), as.numeric(month):12))}else{next4m <- sprintf("Month%02d", c(as.numeric(month):(as.numeric(month)+3)))}
PO <- read.csv(rownames(file.info(list.files(path = "../PO/", pattern = "PurchaseOrders.*.csv", full.names = TRUE)) %>% filter(mtime == max(mtime)))) %>% filter(Warehouse == "WH-SURREY") %>% mutate(Quantity.on.Shipments = ifelse(Quantity.on.Shipments == "", 0, as.numeric(Quantity.on.Shipments)), Qty.Remain = as.numeric(Quantity) - Quantity.on.Shipments) %>% 
  group_by(Item) %>% summarise(Qty.Remain = sum(Qty.Remain)) %>% filter(Qty.Remain > 0) %>% as.data.frame() %>% `row.names<-`(.[, "Item"])
monR <- monR %>% mutate(T.next4m = rowSums(across(all_of(next4m)))) %>% `row.names<-`(.[, "Category"])
Sales_Inv_Next4m <- sales_SKU_last12month %>% filter(Sales.Channel %in% c("Amazon.com", "Amazon.ca", "janandjul.com")) %>% group_by(Adjust_MSKU) %>% summarise(Sales.last3m = sum(T.last3m)) %>% filter(Adjust_MSKU %in% inventory$Adjust_MSKU) %>%
  mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU), cat = mastersku[Adjust_MSKU, "Category.SKU"], Seasons = mastersku[Adjust_MSKU, "Seasons.SKU"], Status = mastersku[Adjust_MSKU, "MSKU.Status"], AMZ_flag = mastersku[Adjust_MSKU, "Pause.Plan.FBA_x000D_.CA/US/MX/AU/UK/DE"], monR_last3m = monR[cat, "T.last3m"], monR_next4m = monR[cat, "T.next4m"], Sales.next4m = as.integer(Sales.last3m/monR_last3m*monR_next4m), Inv_WH_AMZ.CA.US = inventory[Adjust_MSKU, "Inv_WH_AMZ.CA.US"], PO_remain = ifelse(Adjust_MSKU %in% PO$SKU, PO[Adjust_MSKU, "Qty.Remain"], 0), diff = Sales.next4m - Inv_WH_AMZ.CA.US - PO_remain, Enough_Inv = (Inv_WH_AMZ.CA.US + PO_remain >= Sales.next4m)) %>%
  filter(!Enough_Inv, grepl("24", Seasons))
write.csv(Sales_Inv_Next4m, file = paste0("../Analysis/Sales_Inv_Prediction_Next4m_", Sys.Date(), ".csv"), row.names = F)
# Discuss with Mei, Florence, Cindy and Matt

# -------- Sync master file barcode with clover: at request -------------
# master SKU file: OneDrive > TWK 2020 share
# download clover inventory: clover_item > Inventory > Items > Export.
library(dplyr)
library(openxlsx)
library(tidyr)
woo <- read.csv(rownames(file.info(list.files(path = "../woo/", pattern = "wc-product-export-", full.names = T)) %>% filter(mtime == max(mtime))), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price) | (Sys.time() < strptime(Date.sale.price.starts, format = "%Y-%m-%d %H:%M:%S") | Sys.time() > strptime(Date.sale.price.ends, format = "%Y-%m-%d %H:%M:%S")), Regular.price, Sale.price)) %>% `row.names<-`(.[, "SKU"])
price <- woo %>% mutate(cat = gsub("-.*", "", SKU)) %>% group_by(cat) %>% summarise(Price = max(Sale.price)) %>% as.data.frame()
price <- rbind(price, data.frame(cat = c("MISC5", "MISC10", "MISC15", "MISC20", "DBRC", "DBTB", "DBTL", "DBTP", "DLBS", "DWJA", "DWJT", "DWPF", "DWPS", "DWSF", "DWSS", "DXBK"), Price = c(5, 10, 15, 20, 30, 35, 40, 40, 25, 50, 40, 30, 25, 60, 55, 30))) %>% `row.names<-`(toupper(.[, "cat"])) 
mastersku <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../FBArefill/Raw Data File/", pattern = "1-MasterSKU-All-Product-", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(.[, "MSKU"])
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% filter(!is.na(Name), !duplicated(Name)) %>% `row.names<-`(.[, "Name"]) 
#category_exclude <- c("FVMU", "PJWU","PLAU", "PTAU", "STORE", "TPLU", "TPSU", "TSWU", "TTLU", "TTSU", "WJAU", "WPSU", "XBMU")
category_exclude <- c("")
clover_item_upload <- data.frame(Clover.ID = "", Name = mastersku[mastersku$MSKU.Status == "Active", "MSKU"]) %>% filter(!is.na(Name)) %>% 
  mutate(cat = mastersku[Name, "Category.SKU"], Alternate.Name = woo[Name, "Name"], Price = ifelse(Name %in% woo$SKU, woo[Name, "Sale.price"], price[cat, "Price"]), Price.Type = ifelse(is.na(Price), "Variable", "Fixed"), Price.Unit = NA, Tax.Rates = ifelse(woo[Name, "Tax.class"] == "full", "GST+PST", "GST"), Tax.Rates = ifelse(is.na(Tax.Rates), "GST", Tax.Rates), Cost = 0, Product.Code = gsub("/.*", "", mastersku[Name, "UPC.Active"]), SKU = Name, Modifier.Groups = NA, Quantity = ifelse(Name %in% clover_item$Name, clover_item[Name, "Quantity"], 0), Printer.Labels = NA, Hidden = "No", Non.revenue.item = "No") %>%
  filter(!(cat %in% category_exclude)) %>% arrange(cat, Name)
clover_item_upload <- clover_item_upload %>% mutate(Price = ifelse(grepl("^SS", Name), 10, Price))
clover_item_upload <- clover_item_upload %>% mutate(Product.Code = ifelse(grepl("^M", Name), gsub("/.*", "", mastersku[gsub("^M", "", Name), "UPC.Active"]), Product.Code))
clover_cat_upload <- clover_item_upload %>% group_by(cat) %>% mutate(id = row_number(cat)) %>% select(cat, id, Name) %>% spread(id, Name, fill = "") %>% t() 
clover_cat_upload <- cbind(data.frame(V0 = c("Category Name", "Items in Category", rep("", times = nrow(clover_cat_upload) - 2))), clover_cat_upload)
clover_item_upload <- clover_item_upload %>% select(-cat)
clover_item_upload <- clover_item_upload %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item)+10, rows = 1:nrow(clover_item)+1000, gridExpand = T)
writeData(clover, sheet = "Items", clover_item_upload)
deleteData(clover, sheet = "Categories", cols = 1:1000, rows = 1:1000, gridExpand = T)
writeData(clover, sheet = "Categories", clover_cat_upload, colNames = F)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
# upload to Clover > Inventory

# -------- Prepare to generate barcode image: at request ------------
library(dplyr)
library(openxlsx)
library(stringr)
season <- "2025SS"
startRow <- 9
mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(.[, "MSKU"]) 
qty0 <- data.frame()
skus <- c()
for(i in c(346, 368, 387:403)){
  POn <- paste0("P", i)
  if(!sum(grepl(POn, list.dirs(paste0("../PO/order/", season, "/"), recursive = F)))) next
  print(POn)
  file <- list.files(path = paste0("../PO/order/", season, "/"), pattern = paste0(POn, ".*.xlsx"), full.names = T, recursive = T)
  for(f in 1:length(file)){
    barcode <- openxlsx::read.xlsx(file[f], sheet = 1, startRow = startRow) %>% select(SKU, Design.Version, TOTAL.ORDER) %>% 
      mutate(SKU = str_trim(SKU), Category = mastersku[SKU, "Category.SKU"], Print.English = mastersku[SKU, "Print.Name"], Print.Chinese = mastersku[SKU, "Print.Chinese"], Size = mastersku[SKU, "Size"], UPC.Active = gsub("/.*", "", mastersku[SKU, "UPC.Active"]), Image = "") %>%
      filter(SKU != "", !is.na(SKU))
    skus <- c(skus, barcode[barcode$TOTAL.ORDER != 0, "SKU"])
    qty0 <- rbind(qty0, barcode %>% filter(grepl("NEW", Design.Version), TOTAL.ORDER == 0, !is.na(UPC.Active)))
    barcode <- barcode %>% filter(TOTAL.ORDER != 0) %>% select(SKU, Print.English, Print.Chinese,	Category, Size, UPC.Active, Design.Version, Image)
    barcode_split <- split(barcode, f = barcode$Category)
    for(j in 1:length(barcode_split)){
      barcode_j <- barcode_split[[j]]
      cat <- barcode_j[1, "Category"]
      dir <- grep(cat, list.dirs("../../TWK Product Labels/", recursive = F), value = T)
      if(length(dir) == 0){
        dir.create(paste0("../../TWK Product Labels/", cat, "/", season, "/"), recursive = T)
        write.xlsx(barcode, file = paste0(paste0("../../TWK Product Labels/", cat, "/", season, "/"), POn, "_", cat, "-Barcode Labels.xlsx"))
      }
      else{
        if(!sum(grepl("25S", list.dirs(dir, recursive = F)))){
          dir.create(paste0(dir, "/", season, "/"))
          write.xlsx(barcode, file = paste0(dir, "/", season, "/", POn, "_", cat, "-Barcode Labels.xlsx"))
        }
        else{
          write.xlsx(barcode, file = paste0(grep("25S", list.dirs(dir, recursive = F), value = T), "/", POn, "_", cat, "-Barcode Labels.xlsx"))
        }
      }
    }
  }
}
qty0 <- qty0 %>% filter(!(SKU %in% skus))
write.xlsx(qty0, file = "../../TWK Product Labels/SKU_new_qty0_UPC.xlsx", rowNames = F)

# -------- Prepare PreProduction_Checklist: at request ------------
library(stringi)
season <- "26S"
folder <- stri_remove_empty(gsub(" *- *sent", "", gsub(paste0("\\.\\.\\/PO\\/order\\/", season, "\\/"), "", list.files(path = paste0("../PO/order/", season, "/"), pattern = "^P"))))
checklist <- stri_split_regex(folder, pattern = " *- *", n = 3, simplify = T) %>% as.data.frame() %>% mutate(V3 = gsub(" *-.*", "", V3))
write.csv(checklist, "../PO/order/checklist.csv", quote = F, row.names = F)

# ------------ Prep for 3PL CA East -------------------
East <- c("MB", "ON", "QC", "NB", "NS", "NL", "PE")
#cat <- c('AJP', 'AWP', 'BSL', 'BST', 'FHA', 'FMG', 'HBU', 'HLH', 'ICP', 'IPC', 'IPS', 'ISJ', 'KHM', 'MBH', 'MKMT', 'MWBF', 'MWBS', 'MWPS', 'PWJ', 'SKB-INSOL', 'UST')
mastersku <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../FBArefill/Raw Data File/", pattern = "1-MasterSKU-All-Product-", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(.[, "MSKU"])
JJ_orders <- read.csv("../woo/orders-20241001-20250131.csv", as.is = T) %>% filter(Order.Status == "Completed", Country.Code..Shipping. != "", SKU != "") %>% mutate(Category = ifelse(SKU %in% row.names(mastersku), mastersku[SKU, "Category.SKU"], gsub("-.*", "", SKU)), Country = Country.Code..Shipping., State = State.Code..Shipping., Qty = Quantity....Refund.) 
JJ_orders_CA <- JJ_orders %>% group_by(Category, Country) %>% summarise(Quantity = sum(Qty)) %>% group_by(Category) %>% mutate(Total_qty = sum(Quantity), Proportion = round(Quantity/Total_qty, 2)) %>% filter(Country == "CA")
JJ_orders_East <- JJ_orders %>% filter(Country == "CA") %>% mutate(Region = ifelse(State %in% East, "East", "West"))%>% group_by(Category, Region) %>% summarise(Quantity = sum(Qty)) %>% group_by(Category) %>% mutate(Total_qty = sum(Quantity), Proportion = round(Quantity/Total_qty, 2)) %>% filter(Region == "East") %>% as.data.frame() %>% `row.names<-`(toupper(.[, "Category"])) 
JJ_orders_propotion <- data.frame(Category = JJ_orders_CA$Category, CA_propotion = JJ_orders_CA$Proportion) %>% mutate(East_proportion = ifelse(Category %in% JJ_orders_East$Category, JJ_orders_East[Category, "Proportion"], 0))
#JJ_orders_propotion <- JJ_orders_propotion %>% filter(Category %in% cat)
write.csv(JJ_orders_propotion, file = paste0("../woo/JJ_orders_propotion_", Sys.Date(), ".csv"), row.names = F)

# ------------ upload regular POs -------------------
library(stringi)
season <- "25F"
folder <- stri_remove_empty(gsub(" *- *sent", "", gsub(paste0("\\.\\.\\/PO\\/order\\/", season, "\\/"), "", list.files(path = paste0("../PO/order/", season, "/"), pattern = "^P"))))
SeasonStart <- read.csv("../PO/SeasonStart.csv", as.is = T) %>% mutate(receive_date = format(as.Date(paste0(arrival_date, "-2025"), format = "%d-%B-%Y"), "%m/%d/%Y")) %>% `row.names<-`(.[, "category"]) 
PO_NS <- data.frame(); total <- 0; items <- 0
for(f in folder){
  memo <- f; ID <- trimws(gsub("-.*", "", f)); category <- trimws(str_split_1(f, "-")[2]); vendor <- trimws(str_split_1(f, "-")[3])
  print(paste(ID, category, vendor, memo))
  file <- list.files(path = paste0("../PO/order/", season, "/"), pattern = paste0(ID, ".*.xlsx"), full.names = TRUE, recursive = T)
  if(length(file) != 1){print(paste(ID, "INPUT FILE ERROR:", length(file), "files found")); next}
  PO <- openxlsx::read.xlsx(file, sheet = 1, startRow = 8) 
  if(sum(grepl("加拿大总数量", colnames(PO))) == 1){
    PO_CA <- PO %>% select("产品编号", "加拿大总数量") %>% filter(grepl("-.*", 产品编号)) %>% mutate(cat = gsub("-.*", "", 产品编号))
    CAT <- paste((PO_CA %>% distinct(cat))$cat, collapse = " ")
    if(CAT != category){print(paste0("Category ERROR: ", category, CAT))}
    if(category %in% SeasonStart$category){ReceiveDate <- SeasonStart[category, "receive_date"]}else{ReceiveDate <- format(as.Date(gsub("出货日期.ShipDate.", "", colnames(PO)[which(grepl("加拿大总数量", colnames(PO)))+1]), "%Y.%m.%d") + 40, "%m/%d/%Y")}
    PO_CA_NS <- PO_CA %>% mutate(PO.TYPE = "Regular", CATEGORY = category, SEASON = season, REF.NO = paste0(ID, "-CA"), WAREHOUSE = "WH-SURREY", VENDOR = vendor, CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = memo, ITEM = 产品编号,  QUANTITY = round(as.numeric(ifelse(is.na(加拿大总数量), 0, 加拿大总数量))), Tax.Code = "CA-Zero", External.ID = paste0(ID, "-CA")) %>% 
      filter(QUANTITY > 0) %>% select(PO.TYPE:External.ID) %>% rename_with(~ gsub("\\.", " ", .))
    print(paste("CA", sum(PO_CA_NS$QUANTITY), PO[2, "加拿大总数量"]))
    total <- total + as.numeric(PO[2, "加拿大总数量"])
    items <- items + nrow(PO_CA_NS)
    if(nrow(PO_CA_NS)){
      Encoding(PO_CA_NS$MEMO) = "UTF-8"
      PO_NS <- rbind(PO_NS, PO_CA_NS)
      write_excel_csv(PO_CA_NS, file = paste0(gsub("(.*\\/).*", "\\1", file), "NS_PO_CA_", ID, ".csv"), na = "")
    }
  }
  if(sum(grepl("英国总数量", colnames(PO))) == 1){
    PO_UK <- PO %>% select("产品编号", "英国总数量") %>% filter(grepl("-.*", 产品编号)) %>% mutate(cat = gsub("-.*", "", 产品编号))
    CAT <- paste((PO_UK %>% distinct(cat))$cat, collapse = " ")
    if(CAT != category){print(paste0("Category ERROR: ", category, CAT))}
    if(category %in% SeasonStart$category){ReceiveDate <- SeasonStart[category, "receive_date"]}else{ReceiveDate <- format(as.Date(gsub("出货日期.ShipDate.", "", colnames(PO)[which(grepl("加拿大总数量", colnames(PO)))+1]), "%Y.%m.%d") + 40, "%m/%d/%Y")}
    PO_UK_NS <- PO_UK %>% mutate(PO.TYPE = "Regular", CATEGORY = category, SEASON = season, REF.NO = paste0(ID, "-UK"), WAREHOUSE = "WH-AMZ : FBA-UK", VENDOR = vendor, CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = memo, ITEM = 产品编号,  QUANTITY = round(as.numeric(ifelse(is.na(英国总数量), 0, 英国总数量))), Tax.Code = "CA-Zero", External.ID = paste0(ID, "-UK")) %>% 
      filter(QUANTITY > 0) %>% select(PO.TYPE:External.ID) %>% rename_with(~ gsub("\\.", " ", .))
    print(paste("UK", sum(PO_UK_NS$QUANTITY), PO[2, "英国总数量"]))
    total <- total + as.numeric(PO[2, "英国总数量"])
    items <- items + nrow(PO_UK_NS)
    if(nrow(PO_UK_NS)){
      Encoding(PO_UK_NS$MEMO) = "UTF-8"
      PO_NS <- rbind(PO_NS, PO_UK_NS)
      write_excel_csv(PO_UK_NS, file = paste0(gsub("(.*\\/).*", "\\1", file), "NS_PO_UK_", ID, ".csv"), na = "")
    }
  }
  if(sum(grepl("德国总数量", colnames(PO))) == 1){
    PO_DE <- PO %>% select("产品编号", "德国总数量") %>% filter(grepl("-.*", 产品编号)) %>% mutate(cat = gsub("-.*", "", 产品编号))
    CAT <- paste((PO_DE %>% distinct(cat))$cat, collapse = " ")
    if(CAT != category){print(paste0("Category ERROR: ", category, CAT))}
    if(category %in% SeasonStart$category){ReceiveDate <- SeasonStart[category, "receive_date"]}else{ReceiveDate <- format(as.Date(gsub("出货日期.ShipDate.", "", colnames(PO)[which(grepl("加拿大总数量", colnames(PO)))+1]), "%Y.%m.%d") + 40, "%m/%d/%Y")}
    PO_DE_NS <- PO_DE %>% mutate(PO.TYPE = "Regular", CATEGORY = category, SEASON = season, REF.NO = paste0(ID, "-DE"), WAREHOUSE = "WH-AMZ : FBA-DE", VENDOR = vendor, CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = memo, ITEM = 产品编号,  QUANTITY = round(as.numeric(ifelse(is.na(德国总数量), 0, 德国总数量))), Tax.Code = "CA-Zero", External.ID = paste0(ID, "-DE")) %>% 
      filter(QUANTITY > 0) %>% select(PO.TYPE:External.ID) %>% rename_with(~ gsub("\\.", " ", .))
    print(paste("DE", sum(PO_DE_NS$QUANTITY), PO[2, "德国总数量"]))
    total <- total + as.numeric(PO[2, "德国总数量"])
    items <- items + nrow(PO_DE_NS)
    if(nrow(PO_DE_NS)){
      Encoding(PO_DE_NS$MEMO) = "UTF-8"
      PO_NS <- rbind(PO_NS, PO_DE_NS)
      write_excel_csv(PO_DE_NS, file = paste0(gsub("(.*\\/).*", "\\1", file), "NS_PO_DE_", ID, ".csv"), na = "")
    }
  }
  if(sum(grepl("澳大利亚总数量", colnames(PO))) == 1){
    PO_AU <- PO %>% select("产品编号", "澳大利亚总数量") %>% filter(grepl("-.*", 产品编号)) %>% mutate(cat = gsub("-.*", "", 产品编号))
    CAT <- paste((PO_AU %>% distinct(cat))$cat, collapse = " ")
    if(CAT != category){print(paste0("Category ERROR: ", category, CAT))}
    if(category %in% SeasonStart$category){ReceiveDate <- SeasonStart[category, "receive_date"]}else{ReceiveDate <- format(as.Date(gsub("出货日期.ShipDate.", "", colnames(PO)[which(grepl("加拿大总数量", colnames(PO)))+1]), "%Y.%m.%d") + 40, "%m/%d/%Y")}
    PO_AU_NS <- PO_AU %>% mutate(PO.TYPE = "Regular", CATEGORY = category, SEASON = season, REF.NO = paste0(ID, "-AU"), WAREHOUSE = "WH-AMZ : FBA-AU", VENDOR = vendor, CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = memo, ITEM = 产品编号,  QUANTITY = round(as.numeric(ifelse(is.na(澳大利亚总数量), 0, 澳大利亚总数量))), Tax.Code = "CA-Zero", External.ID = paste0(ID, "-AU")) %>% 
      filter(QUANTITY > 0) %>% select(PO.TYPE:External.ID) %>% rename_with(~ gsub("\\.", " ", .))
    print(paste("AU", sum(PO_AU_NS$QUANTITY), PO[2, "澳大利亚总数量"]))
    total <- total + as.numeric(PO[2, "澳大利亚总数量"])
    items <- items + nrow(PO_AU_NS)
    if(nrow(PO_AU_NS)){
      Encoding(PO_AU_NS$MEMO) = "UTF-8"
      PO_NS <- rbind(PO_NS, PO_AU_NS)
      write_excel_csv(PO_AU_NS, file = paste0(gsub("(.*\\/).*", "\\1", file), "NS_PO_AU_", ID, ".csv"), na = "")
    }
  }
  if(sum(grepl("日本总数量", colnames(PO))) == 1){
    PO_JP <- PO %>% select("产品编号", "日本总数量") %>% filter(grepl("-.*", 产品编号)) %>% mutate(cat = gsub("-.*", "", 产品编号))
    CAT <- paste((PO_JP %>% distinct(cat))$cat, collapse = " ")
    if(CAT != category){print(paste0("Category ERROR: ", category, CAT))}
    if(category %in% SeasonStart$category){ReceiveDate <- SeasonStart[category, "receive_date"]}else{ReceiveDate <- format(as.Date(gsub("出货日期.ShipDate.", "", colnames(PO)[which(grepl("加拿大总数量", colnames(PO)))+1]), "%Y.%m.%d") + 40, "%m/%d/%Y")}
    PO_JP_NS <- PO_JP %>% mutate(PO.TYPE = "Regular", CATEGORY = category, SEASON = season, REF.NO = paste0(ID, "-JP"), WAREHOUSE = "WH-AMZ : FBA-JP", VENDOR = vendor, CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = memo, ITEM = 产品编号,  QUANTITY = round(as.numeric(ifelse(is.na(日本总数量), 0, 日本总数量))), Tax.Code = "CA-Zero", External.ID = paste0(ID, "-JP")) %>% 
      filter(QUANTITY > 0) %>% select(PO.TYPE:External.ID) %>% rename_with(~ gsub("\\.", " ", .))
    print(paste("JP", sum(PO_JP_NS$QUANTITY), PO[2, "日本总数量"]))
    total <- total + as.numeric(PO[2, "日本总数量"])
    items <- items + nrow(PO_JP_NS)
    if(nrow(PO_JP_NS)){
      Encoding(PO_JP_NS$MEMO) = "UTF-8"
      PO_NS <- rbind(PO_NS, PO_JP_NS)
      write_excel_csv(PO_JP_NS, file = paste0(gsub("(.*\\/).*", "\\1", file), "NS_PO_JP_", ID, ".csv"), na = "")
    }
  }
  if(sum(grepl("中国总数量", colnames(PO))) == 1){
    PO_CN <- PO %>% select("产品编号", "中国总数量") %>% filter(grepl("-.*", 产品编号)) %>% mutate(cat = gsub("-.*", "", 产品编号))
    CAT <- paste((PO_CN %>% distinct(cat))$cat, collapse = " ")
    if(CAT != category){print(paste0("Category ERROR: ", category, CAT))}
    if(category %in% SeasonStart$category){ReceiveDate <- SeasonStart[category, "receive_date"]}else{ReceiveDate <- format(as.Date(gsub("出货日期.ShipDate.", "", colnames(PO)[which(grepl("加拿大总数量", colnames(PO)))+1]), "%Y.%m.%d") + 40, "%m/%d/%Y")}
    PO_CN_NS <- PO_CN %>% mutate(PO.TYPE = "Regular", CATEGORY = category, SEASON = season, REF.NO = paste0(ID, "-CN"), WAREHOUSE = "WH-CHINA", VENDOR = vendor, CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = memo, ITEM = 产品编号,  QUANTITY = round(as.numeric(ifelse(is.na(中国总数量), 0, 中国总数量))), Tax.Code = "CA-Zero", External.ID = paste0(ID, "-CN")) %>% 
      filter(QUANTITY > 0) %>% select(PO.TYPE:External.ID) %>% rename_with(~ gsub("\\.", " ", .))
    print(paste("CN", sum(PO_CN_NS$QUANTITY), PO[2, "中国总数量"]))
    total <- total + as.numeric(PO[2, "中国总数量"])
    items <- items + nrow(PO_CN_NS)
    if(nrow(PO_CN_NS)){
      Encoding(PO_CN_NS$MEMO) = "UTF-8"
      PO_NS <- rbind(PO_NS, PO_CN_NS)
      write_excel_csv(PO_CN_NS, file = paste0(gsub("(.*\\/).*", "\\1", file), "NS_PO_CN_", ID, ".csv"), na = "")
    }
  }
  print(ReceiveDate)
}
print(paste("Total No. of items: ", nrow(PO_NS), items, "; Total Qty: ", sum(PO_NS$QUANTITY), total))
write_excel_csv(PO_NS, file = paste0("../PO/order/", season, "/NS_PO_regular_", format(Sys.Date(), "%Y%m%d"), ".csv"), na = "")

# ------------ upload wholesaler POs: CZ-Mylerie, CA-Clement, FR-Petits, US-BabiesR -------------------
ID <- "P553"; season <- "26S"; type <- "CZ-Mylerie"; warehouse <- "WH-CHINA"
file <- list.files(path = paste0("../PO/order/", season, "/"), pattern = paste0(ID, ".*.xlsx"), full.names = TRUE, recursive = T)
PO_NS <- data.frame(); total <- 0; items <- 0; CAT <- c()
for(f in file){
  sheets <- getSheetNames(f)[getSheetNames(f) != "Export Summary"]
  memo <- gsub(paste0(".*", ID, " *\\- *"), "", gsub(paste0("\\/", ID, "-.*\\.xlsx"), "", f))
  for(s in sheets){
    print(c(f, s))
    PO <- openxlsx::read.xlsx(f, sheet = s, startRow = 8)
    ReceiveDate <- format(as.Date(gsub("出货日期.ShipDate.", "", colnames(PO)[which(grepl("中国总数量", colnames(PO)))+1]), "%Y.%m.%d") + 40, "%m/%d/%Y")
    PO_WS <- PO %>% select("产品编号", "订单.总数量") %>% filter(grepl("-.*-", 产品编号)) %>% mutate(cat = gsub("-.*", "", 产品编号))
    CAT <- c(CAT, paste((PO_WS %>% distinct(cat))$cat, collapse = " "))
    PO_WS_NS <- PO_WS %>% mutate(PO.TYPE = type, CATEGORY = "", SEASON = season, REF.NO = ID, WAREHOUSE = warehouse, VENDOR = "China", CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = memo, ITEM = 产品编号,  QUANTITY = as.numeric(ifelse(is.na(订单.总数量), 0, 订单.总数量)), Tax.Code = "CA-Zero", External.ID = ID) %>% 
      filter(QUANTITY > 0) %>% select(PO.TYPE:External.ID) %>% rename_with(~ gsub("\\.", " ", .))
    total <- total + sum(PO_WS_NS$QUANTITY)
    items <- items + nrow(PO_WS_NS)
    print(paste0("No. of items: ", nrow(PO_WS_NS), "; Qty: ", sum(PO_WS_NS$QUANTITY)))
    PO_NS <- rbind(PO_NS, PO_WS_NS)
  }
}
PO_NS$CATEGORY <- paste(CAT, collapse = " ")
print(paste("Total No. of items: ", nrow(PO_NS), items, "; Total Qty: ", sum(PO_NS$QUANTITY), total))
write_excel_csv(PO_NS, file = paste0(gsub("(.*\\/).*", "\\1", f), "NS_PO_", ID, ".csv"), na = "")

# ------------ upload CEFA POs -------------------
ID <- "CEFA7"; season <- "25F"; ReceiveDate <- "10/20/2025"
netsuite_item <- read.csv(rownames(file.info(list.files(path = "../NetSuite/", pattern = "Items_All_", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T)
PO <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../PO/order/CEFA/", pattern = "Order.*.xlsx", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = 1, startRow = 2) %>% 
  filter(grepl("-.*-", SKU.Number), !is.na(Quantity), Quantity > 0) %>% mutate(cat = gsub("-.*", "", SKU.Number))
CAT <- paste((PO %>% distinct(cat))$cat, collapse = " ")
new_SKU <- PO %>% filter(!(SKU.Number %in% netsuite_item$Name)) %>% select(SKU.Number, Quantity)
if(nrow(new_SKU)){write.csv(new_SKU, file = paste0("../PO/order/CEFA/", "new_SKU_", ID, ".csv"), row.names = F, quote = F, na = "")}
PO_NS <- PO %>% mutate(PO.TYPE = "Daycare Uniforms", CATEGORY = CAT, SEASON = season, REF.NO = ID, WAREHOUSE = "WH-SURREY", VENDOR = "China", CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = paste0("CEFA order for ", season), ITEM = SKU.Number,  QUANTITY = Quantity, Tax.Code = "CA-Zero", External.ID = REF.NO) %>% 
  filter(Quantity > 0) %>% select(PO.TYPE:External.ID) %>% rename_with(~ gsub("\\.", " ", .))
write.csv(PO_NS, file = paste0("../PO/order/CEFA/", "NS_PO_", ID, ".csv"), row.names = F, quote = F, na = "")

# ------------ upload inbound shipment for POs -------------------
library(tidyr)
season <- "26S"; warehouse <- "WH-SURREY"; PO_suffix <- "-CA"
RefNo <- "26SSCA3"; ShippingDate <- "11/29/2025"; ReceiveDate <- "12/29/2025"; AMZ.Shipment.ID <- ""
PO_detail <- read.csv(rownames(file.info(list.files(path = "../PO/", pattern = "PurchaseOrders", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% filter(Item != "") %>% 
  mutate(Quantity = as.numeric(Quantity), Quantity.Fulfilled.Received = as.numeric(Quantity.Fulfilled.Received), Quantity.on.Shipments = ifelse(is.na(as.numeric(Quantity.on.Shipments)), 0, as.numeric(Quantity.on.Shipments)), Quantity.Remain = Quantity - Quantity.on.Shipments, Quantity.Remain = ifelse(Quantity.Remain < 0, 0, Quantity.Remain)) %>% `row.names<-`(paste0(.[, "REF.NO"], "_", .[, "Item"]))
POn <- PO_detail %>% filter(!duplicated(REF.NO)) %>% `row.names<-`(.[, "REF.NO"])
Weight_manual <- openxlsx::read.xlsx("../PO/shipment/UnitWeight_ManualCheck.xlsx", sheet = 1) %>% `row.names<-`(paste0(.[, "Category"], "_", .[, "Size"]))
Weight_master <- openxlsx::read.xlsx(list.files(path = "../PO/shipment/", pattern = "Product&DimsLog", full.names = T), sheet = "DimensionLog", fillMergedCells = T) %>% filter(!is.na(Size)) %>% `row.names<-`(paste0(.[, "Category"], "_", .[, "Size"])) %>% mutate(Unit.Weight = round(`Sample.Weight(g)`, 0))
Unit_weight <- rbind(Weight_manual %>% select(Category, Size, Unit.Weight), Weight_master %>% select(Category, Size, Unit.Weight) %>% filter(!(rownames(Weight_master) %in% rownames(Weight_manual))))
shipment_in <- list.files(path = "../PO/shipment/", pattern = paste0(".*", RefNo, ".*.xlsx"), recursive = T, full.names = T)
memo <- gsub(",", " ", gsub(paste0(RefNo, " *"), "", gsub(" *ETA.*\\/.*", "", gsub("../PO/shipment/+", "", shipment_in))))
sheets <- getSheetNames(shipment_in)[]
summary <- openxlsx::read.xlsx(shipment_in, sheet = 1, fillMergedCells = T)
shipment <- data.frame(); attachment <- data.frame(); Nbox <- 0
for(s in 2:length(sheets)){
  sheet <- sheets[s]
  #PO_suffix <- ifelse(s%in% c(7, 8), "-CN", "-CA")
  shipment_i <- openxlsx::read.xlsx(shipment_in, sheet = sheet, startRow = 5, fillMergedCells = T) %>% 
    select(1:PER.CTN) %>% filter(grepl(".*\\-.*", SKU)) %>% tidyr::fill(1:Season, .direction = "down") %>% mutate(`PO.#` = str_trim(`PO.#`), Unit.weight = ifelse(as.numeric(`Unit.Weight(g)`) < 10, round(as.numeric(`Unit.Weight(g)`)*1000, 0), round(as.numeric(`Unit.Weight(g)`), 0))) 
  if(sum(grepl("CTN.NO", colnames(shipment_i))) == 1){
    print(paste(1, sheet))
    shipment_o <- data.frame(REF.NO = RefNo, EXPECTED.SHIPPING.DATE = ShippingDate, EXPECTED.DELIVERY.DATE = ReceiveDate, MEMO = memo, PO.REF.NO = paste0(shipment_i$`PO.#`, PO_suffix), BOX.NO = paste0(s, "-", shipment_i$`CTN.NO`), ITEM = shipment_i$SKU, QUANTITY = as.numeric(shipment_i$QTY.SHIPPED), LOCATION = warehouse) %>% mutate(PO = paste0("PO#", POn[PO.REF.NO, "Document.Number"]), Unit.Weight.Shipment = as.numeric(shipment_i$Unit.weight), Unit.Weight.woo = Unit_weight[gsub("-[A-Z]+-", "_", ITEM), "Unit.Weight"], GR.Weight = round(as.numeric(shipment_i$PER.CTN), 2)) %>% group_by(BOX.NO) %>% mutate(EST.Weight.Shipment = round(sum(QUANTITY*Unit.Weight.Shipment/1000) + 1, 2), EST.Weight.woo = round(sum(QUANTITY*Unit.Weight.woo/1000) + 1, 2))
  }else{
    print(paste(2, sheet))
    colnames(shipment_i) <- c("Start", "End", colnames(shipment_i)[3:length(colnames(shipment_i))])
    shipment_i <- shipment_i %>% mutate(Start = as.numeric(Start), End = as.numeric(End), QTY.Per.Box = as.numeric(QTY.Per.Box))
    shipment_o <- data.frame()
    for(i in 1:nrow(shipment_i)){
      for(n in shipment_i[i, "Start"]:shipment_i[i, "End"]){
        shipment_o <- rbind(shipment_o, data.frame(REF.NO = RefNo, EXPECTED.SHIPPING.DATE = ShippingDate, EXPECTED.DELIVERY.DATE = ReceiveDate, MEMO = memo, PO.REF.NO = paste0(shipment_i[i, "PO.#"], PO_suffix), BOX.NO = paste0(s, "-", n), ITEM = shipment_i[i, "SKU"], QUANTITY = as.numeric(shipment_i[i, "QTY.Per.Box"]), LOCATION = warehouse) %>% mutate(PO = paste0("PO#", POn[PO.REF.NO, "Document.Number"]), Unit.Weight.Shipment = as.numeric(shipment_i[i, "Unit.weight"]), Unit.Weight.woo = Unit_weight[gsub("-[A-Z]+-", "_", ITEM), "Unit.Weight"], GR.Weight = round(as.numeric(shipment_i[i, "PER.CTN"]), 2)))
      }
    }
    shipment_o <- shipment_o %>% group_by(BOX.NO) %>% mutate(EST.Weight.Shipment = round(sum(QUANTITY*Unit.Weight.Shipment/1000) + 1, 2), EST.Weight.woo = round(sum(QUANTITY*Unit.Weight.woo/1000) + 1, 2))
  }
  Nbox <- Nbox + length(unique(shipment_o$BOX.NO))
  print(paste0("Total Qty ", sum(shipment_o$QUANTITY), "; Total Box# ", length(unique(shipment_o$BOX.NO))))
  shipment <- rbind(shipment, shipment_o %>% select(REF.NO:PO))
  attachment <- rbind(attachment, data.frame(Box.Number = c(gsub(".*-", "", shipment_o$BOX.NO), ""), SKU = c(shipment_o$ITEM, ""), Packing.Slip.QTY = c(shipment_o$QUANTITY, ""), GR.Weight = c(shipment_o$GR.Weight, ""), Unit.Weight.woo = c(shipment_o$Unit.Weight.woo, ""), Unit.Weight.Shipment = c(shipment_o$Unit.Weight.Shipment, ""), EST.Weight.woo = c(shipment_o$EST.Weight.woo, ""), EST.Weight.Shipment = c(shipment_o$EST.Weight.Shipment, ""), Notes = ""))
}
print(paste0("Qty: ", sum(shipment$QUANTITY), "; #Boxes: ", Nbox))
if(sum(shipment$QUANTITY) == summary[nrow(summary), "QTY"]){print(paste0("Total QTY matches: ", sum(shipment$QUANTITY)))}else{print(paste("Total QTY NOT matching:", sum(shipment$QUANTITY), summary[nrow(summary), "QTY"]))}
if(Nbox == summary[nrow(summary), "Number.of.Box"]){print(paste0("Total Box# matches: ", Nbox))}else{print(paste("Total Box# NOT matching:", Nbox, summary[nrow(summary), "Number.of.Box"]))}
attachment <- attachment %>% rename_with(~ gsub("\\.", " ", .))
write.csv(attachment, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, "_receiving_by_box.csv"), row.names = F, quote = F, na = "")
# check for over-receiving
shipment <- shipment %>% mutate(PO.REF.NO = ifelse(PO == "PO#NA", gsub("-.*", "", PO.REF.NO), PO.REF.NO), PO = paste0("PO#", POn[PO.REF.NO, "Document.Number"]))
shipment <- shipment %>% group_by(PO.REF.NO, ITEM) %>% mutate(QUANTITY = sum(QUANTITY), BOX.NO = paste0(BOX.NO, collapse = " | ")) %>% distinct(PO.REF.NO, ITEM, .keep_all = T) %>% ungroup()
OverReceive <- shipment %>% mutate(Qty.Remain = PO_detail[paste0(PO.REF.NO, "_", ITEM), "Quantity.Remain"], Qty.Remain = ifelse(is.na(Qty.Remain), 0, Qty.Remain), Qty = QUANTITY - Qty.Remain) %>% filter(QUANTITY > Qty.Remain)
if(nrow(OverReceive)){
  View(OverReceive)
  PO_OverReceive <- OverReceive %>% mutate(PO.TYPE = "Over Received", CATEGORY = "MIX", SEASON = season, REF.NO = paste0("Over.", RefNo), WAREHOUSE = gsub("FBA", "WH-AMZ : FBA", warehouse), VENDOR = "China", CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = paste0("Over receiving in shipment ", RefNo), ITEM. = ITEM, QUANTITY. = Qty, Tax.Code = "CA-Zero", External.ID = REF.NO) %>% 
    select(PO.TYPE:External.ID, REF.NO, MEMO) %>% group_by(ITEM.) %>% mutate(QUANTITY. = sum(QUANTITY.)) %>% distinct(ITEM., .keep_all = T) %>% ungroup() %>% rename_with(~ gsub("\\.", " ", .))
  write.csv(PO_OverReceive, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_PO_", RefNo, "_OverReceive.csv"), row.names = F, na = "")
}else{
  print("No over-receiving. ")
}
## upload over-receiving PO 
# output for NS: upload shipment to Inbound Shipment and attach attachment in Communication tab
if(nrow(OverReceive)){
  p <- "PO#PO000434" # over-receiving PO 
  shipment <- shipment %>% mutate(Qty.Remain = PO_detail[paste0(PO.REF.NO, "_", ITEM), "Quantity.Remain"], Qty.Remain = ifelse(is.na(Qty.Remain), 0, Qty.Remain), QUANTITY = ifelse(QUANTITY > Qty.Remain, Qty.Remain, QUANTITY)) %>% filter(PO != "PO#NA") %>% select(-Qty.Remain)
  shipment <- rbind(shipment, data.frame(REF.NO = RefNo, EXPECTED.SHIPPING.DATE = ShippingDate, EXPECTED.DELIVERY.DATE = ReceiveDate, MEMO = memo, PO.REF.NO = paste0("Over.", RefNo), BOX.NO = OverReceive$BOX.NO, ITEM = OverReceive$ITEM, QUANTITY = OverReceive$Qty, LOCATION = warehouse, PO = p)) %>% filter(QUANTITY != 0) %>% 
    mutate(BOX.NO = ifelse(nchar(BOX.NO) > 300, substring(BOX.NO, 1, 299), BOX.NO)) %>% arrange(ITEM) 
}
shipment <- shipment %>% group_by(PO.REF.NO, ITEM) %>% mutate(QUANTITY = sum(QUANTITY), BOX.NO = paste0(BOX.NO, collapse = " | "), AMZ.Shipment.ID = AMZ.Shipment.ID) %>% distinct(PO.REF.NO, ITEM, .keep_all = T) %>% ungroup()
write.csv(shipment, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, ".csv"), row.names = F, quote = F, na = "", fileEncoding = "UTF-8")
## generate receiving csv file
#id <- "INBSHIP14"
#receiving <- shipment %>% mutate(ID = id, PO = gsub("PO#PO", "PO-", PO), Item = ITEM, Qty = QUANTITY) %>% select(ID, PO, Item, Qty)
#write.csv(receiving, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, "_receiving.csv"), row.names = F, quote = F, na = "")
## check weights
weights <- read.csv(list.files(path = "../PO/", pattern = paste0("NS_", RefNo, "_receiving_by_box.csv"), recursive = T, full.names = T), as.is = T) %>% mutate(Diff.woo = Total.Weight..g./1000 - EST.Weight.woo, Diff.shipment = Total.Weight..g./1000 - EST.Weight.Shipment)
woo_update <- weights %>% filter(abs(Diff.woo) > abs(Diff.shipment)) %>% distinct(SKU, .keep_all = T) %>% mutate(Category = gsub("-.*", "", SKU), Size = gsub(".*-", "", SKU))
write.csv(woo_update, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "Weight_diff_woo_update_", RefNo, ".csv"), row.names = F, quote = F, na = "")
weight_check <- weights %>% filter(abs(Diff.woo) > 0.5, abs(Diff.shipment) > 0.5)
write.csv(weight_check, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "Weight_check_", RefNo, ".csv"), row.names = F, quote = F, na = "")

# ------------ inbound shipment for PO TO combined -------------------
library(tidyr)
season <- "25S"; warehouse <- "WH-SURREY"; PO_suffix <- "-CA"
RefNo <- "25SSCA4"; ShippingDate <- "1/26/2025"; ReceiveDate <- "3/10/2025"
PO_detail <- read.csv(rownames(file.info(list.files(path = "../PO/", pattern = "PurchaseOrders", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% filter(Item != "") %>% 
  mutate(Quantity = as.numeric(Quantity), Quantity.on.Shipments = ifelse(is.na(as.numeric(Quantity.on.Shipments)), 0, as.numeric(Quantity.on.Shipments)), Quantity.Remain = Quantity - Quantity.on.Shipments, Quantity.Remain = ifelse(Quantity.Remain < 0, 0, Quantity.Remain)) %>% `row.names<-`(paste0(.[, "REF.NO"], "_", .[, "Item"]))
POn <- PO_detail %>% filter(!duplicated(REF.NO)) %>% `row.names<-`(.[, "REF.NO"])
shipment_in <- list.files(path = "../PO/shipment/", pattern = paste0(".*", RefNo, ".*.xlsx"), recursive = T, full.names = T)
memo <- gsub(",", " ", gsub(paste0(RefNo, " *"), "", gsub(" *ETA.*\\/.*", "", gsub("../PO/shipment/+", "", shipment_in))))
sheets <- getSheetNames(shipment_in)[]
summary <- openxlsx::read.xlsx(shipment_in, sheet = 1, fillMergedCells = T)
shipment <- data.frame(); TO <- data.frame(); attachment <- data.frame(); Nbox <- 0
for(s in 2:length(sheets)){
  sheet <- sheets[s]
  shipment_i <- openxlsx::read.xlsx(shipment_in, sheet = sheet, startRow = 5, fillMergedCells = T) %>% 
    select(1:Season) %>% filter(grepl(".*\\-.*", SKU)) %>% tidyr::fill(1:Season, .direction = "down")
  if(sum(grepl("CTN.NO", colnames(shipment_i))) == 1){
    print(paste(1, sheet))
    shipment_o <- data.frame(REF.NO = RefNo, EXPECTED.SHIPPING.DATE = ShippingDate, EXPECTED.DELIVERY.DATE = ReceiveDate, MEMO = memo, PO.REF.NO = paste0(shipment_i$`PO.#`, PO_suffix), BOX.NO = paste0(s, "-", shipment_i$`CTN.NO`), ITEM = shipment_i$SKU, QUANTITY = as.numeric(shipment_i$QTY.SHIPPED), LOCATION = warehouse) %>% mutate(PO = paste0("PO#", POn[PO.REF.NO, "Document.Number"]))
  }else{
    print(paste(2, sheet))
    colnames(shipment_i) <- c("Start", "End", colnames(shipment_i)[3:length(colnames(shipment_i))])
    shipment_i <- shipment_i %>% mutate(Start = as.numeric(Start), End = as.numeric(End), QTY.Per.Box = as.numeric(QTY.Per.Box))
    shipment_o <- data.frame()
    for(i in 1:nrow(shipment_i)){
      for(n in shipment_i[i, "Start"]:shipment_i[i, "End"]){
        shipment_o <- rbind(shipment_o, data.frame(REF.NO = RefNo, EXPECTED.SHIPPING.DATE = ShippingDate, EXPECTED.DELIVERY.DATE = ReceiveDate, MEMO = memo, PO.REF.NO = paste0(shipment_i[i, "PO.#"], PO_suffix), BOX.NO = paste0(s, "-", n), ITEM = shipment_i[i, "SKU"], QUANTITY = as.numeric(shipment_i[i, "QTY.Per.Box"]), LOCATION = warehouse) %>% mutate(PO = paste0("PO#", POn[PO.REF.NO, "Document.Number"])))
      }
    }
  }
  Nbox <- Nbox + length(unique(shipment_o$BOX.NO))
  print(paste0("Total Qty ", sum(shipment_o$QUANTITY), "; Total Box# ", length(unique(shipment_o$BOX.NO))))
  if(grepl("CR", sheets[s])){
    TO <- rbind(TO, data.frame(Date = format(Sys.Date(), "%m/%d/%Y"), TO.TYPE = "China-Surrey", SEASON = "24F", FROM.WAREHOUSE = "WH-CHINA", TO.WAREHOUSE = "WH-SURREY", REF.NO = paste0("TO-", RefNo), Memo = memo, ORDER.PLACED.BY = "Gloria Li", ITEM = shipment_o$ITEM, Quantity = shipment_o$QUANTITY))
  }else{
    shipment <- rbind(shipment, shipment_o)
  }
  attachment <- rbind(attachment, data.frame(Box.Number = c(gsub(".*-", "", shipment_o$BOX.NO), ""), SKU = c(shipment_o$ITEM, ""), Packing.Slip.QTY = c(shipment_o$QUANTITY, ""), Notes = ""))
}
if(nrow(TO)){
  if(sum(shipment$QUANTITY) + sum(TO$Quantity) == summary[nrow(summary), "QTY"]){print(paste0("Total QTY matches: ", summary[nrow(summary), "QTY"]))}else{print(paste("Total QTY NOT matching:", sum(shipment$QUANTITY) + sum(TO$Quantity), summary[nrow(summary), "QTY"]))}
}else{
  if(sum(shipment$QUANTITY) == summary[nrow(summary), "QTY"]){print(paste0("Total QTY matches: ", sum(shipment$QUANTITY)))}else{print(paste("Total QTY NOT matching:", sum(shipment$QUANTITY), summary[nrow(summary), "QTY"]))}
}
if(Nbox == summary[nrow(summary), "Number.of.Box"]){print(paste0("Total Box# matches: ", Nbox))}else{print(paste("Total Box# NOT matching:", Nbox, summary[nrow(summary), "Number.of.Box"]))}
# check for over-receiving
if(nrow(TO)){
  netsuite_item <- read.csv(list.files(path = "../NetSuite/", pattern = paste0("Items_All_", format(Sys.Date(), "%Y%m%d"), ".csv"), full.names = T), as.is = T) %>% filter(Inventory.Warehouse == unique(TO$FROM.WAREHOUSE)) %>% `row.names<-`(.[, "Name"]) 
  netsuite_item[netsuite_item == "" | is.na(netsuite_item)] <- 0
  TO <- TO %>% group_by(ITEM) %>% mutate(Quantity = sum(Quantity)) %>% distinct(ITEM, .keep_all = T) %>% ungroup() %>% mutate(Qty.availeble = ifelse(ITEM %in% netsuite_item$Name, netsuite_item[ITEM, "Warehouse.Available"], 0), Qty.add = Quantity - Qty.availeble)
  if(nrow(TO %>% filter(Qty.add > 0))){
    inventory_update <- data.frame(Date = format(Sys.Date(), "%m/%d/%Y"), ITEM = (TO %>% filter(Qty.add > 0))$ITEM, ADJUST.QTY.BY = (TO %>% filter(Qty.add > 0))$Qty.add, Reason.Code = "CC", MEMO = paste0(RefNo, " China Reserve TO"), WAREHOUSE = "WH-CHINA", External.ID = paste0("IAC-", format(Sys.Date(), "%y%m%d"), "-1"))
    colnames(inventory_update) <- gsub("QTY BY", "QTY. BY", gsub("\\.", " ", colnames(inventory_update)))
    View(inventory_update)
    write.csv(inventory_update, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "IAC-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")
  }
  TO <- TO %>% select(-Qty.availeble, -Qty.add)
}else{
  print("No TO in this shipment. ")
}
shipment <- shipment %>% mutate(PO.REF.NO = ifelse(PO == "PO#NA", gsub("-.*", "", PO.REF.NO), PO.REF.NO), PO = paste0("PO#", POn[PO.REF.NO, "Document.Number"]))
shipment <- shipment %>% group_by(PO.REF.NO, ITEM) %>% mutate(QUANTITY = sum(QUANTITY), BOX.NO = paste0(BOX.NO, collapse = " | ")) %>% distinct(PO.REF.NO, ITEM, .keep_all = T) %>% ungroup()
OverReceive <- shipment %>% mutate(Qty.Remain = PO_detail[paste0(PO.REF.NO, "_", ITEM), "Quantity.Remain"], Qty.Remain = ifelse(is.na(Qty.Remain), 0, Qty.Remain), Qty = QUANTITY - Qty.Remain) %>% filter(QUANTITY > Qty.Remain)
if(nrow(OverReceive)){
  PO_OverReceive <- OverReceive %>% mutate(PO.TYPE = "Over Received", CATEGORY = "MIX", SEASON = season, REF.NO = paste0("Over.", RefNo), WAREHOUSE = warehouse, VENDOR = "China", CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = paste0("Over receiving in shipment ", RefNo), ITEM. = ITEM, QUANTITY. = Qty, Tax.Code = "CA-Zero", External.ID = REF.NO) %>% 
    select(PO.TYPE:External.ID) %>% rename_with(~ gsub("\\.", " ", .))
  PO_OverReceive <- PO_OverReceive %>% group_by(ITEM) %>% mutate(QUANTITY = sum(QUANTITY)) %>% ungroup %>% distinct(ITEM, .keep_all = T)
  View(PO_OverReceive)
  write.csv(PO_OverReceive, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_PO_", RefNo, "_OverReceive.csv"), row.names = F, na = "")
}else{
  print("No over-receiving. ")
}
TO <- TO %>% filter(Quantity > 0) %>% rename_with(~ gsub("\\.", " ", .))
write.csv(TO, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, "_TO.csv"), row.names = F, quote = F, na = "")
attachment <- attachment %>% rename_with(~ gsub("\\.", " ", .))
write.csv(attachment, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, "_receiving_by_box.csv"), row.names = F, quote = F, na = "")
## upload over-receiving PO 
# output for NS: upload shipment to Inbound Shipment and attach attachment in Communication tab
if(nrow(OverReceive)){
  p <- "PO#PO000234" # over-receiving PO 
  shipment <- shipment %>% mutate(Qty.Remain = PO_detail[paste0(PO.REF.NO, "_", ITEM), "Quantity.Remain"], Qty.Remain = ifelse(is.na(Qty.Remain), 0, Qty.Remain), QUANTITY = ifelse(QUANTITY > Qty.Remain, Qty.Remain, QUANTITY)) %>% filter(PO != "PO#NA") %>% select(-Qty.Remain)
  shipment <- rbind(shipment, data.frame(REF.NO = RefNo, EXPECTED.SHIPPING.DATE = ShippingDate, EXPECTED.DELIVERY.DATE = ReceiveDate, MEMO = memo, PO.REF.NO = paste0("Over.", RefNo), BOX.NO = OverReceive$BOX.NO, ITEM = OverReceive$ITEM, QUANTITY = OverReceive$Qty, LOCATION = warehouse, PO = p)) %>% filter(QUANTITY != 0) %>% 
    mutate(BOX.NO = ifelse(nchar(BOX.NO) > 300, substring(BOX.NO, 1, 299), BOX.NO)) %>% arrange(ITEM) 
}
shipment <- shipment %>% group_by(ITEM, PO) %>% mutate(QUANTITY = sum(QUANTITY)) %>% ungroup %>% distinct(ITEM, PO, .keep_all = T)
write.csv(shipment, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, ".csv"), row.names = F, quote = F, na = "")
## generate receiving csv file
id <- "INBSHIP16"
receiving <- shipment %>% mutate(ID = id, PO = gsub("PO#PO", "PO-", PO), Item = ITEM, Qty = QUANTITY) %>% select(ID, PO, Item, Qty)
write.csv(receiving, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, "_receiving.csv"), row.names = F, quote = F, na = "")
TO_receiving <- data.frame(TO = "TO000117", Internal.ID = "243637", Item = TO$ITEM, Qty = TO$Quantity)
TO_receiving <- TO_receiving %>% rename_with(~ gsub("\\.", " ", .))
write.csv(TO_receiving, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, "_TO_receiving.csv"), row.names = F, quote = F, na = "")

# ------------ check shipment Qty -------------------
RefNo <- "25FWCA7"
attachment <- read.csv(list.files(path = "../PO/shipment/", pattern = paste0(RefNo, "_receiving_by_box.csv"), recursive = T, full.names = T), as.is = T) %>% filter(SKU!="") %>% `row.names<-`(paste0(.[, "SKU"], "_", .[, "Box.Number"]))
shipment_weight <- openxlsx::read.xlsx(list.files(path = "../PO/shipment/", pattern = paste0("BoxWeight_", RefNo, ".xlsx"), recursive = T, full.names = T)) %>% 
  mutate(Total.Weight = `Total.Weight.(g)`/1000, Unit.Weight.woo = attachment[paste0(SKU, "_", Box.Number), "Unit.Weight.woo"], Unit.Weight.Shipment = attachment[paste0(SKU, "_", Box.Number), "Unit.Weight.Shipment"], EST.Weight.woo = attachment[paste0(SKU, "_", Box.Number), "EST.Weight.woo"], EST.Weight.Shipment = attachment[paste0(SKU, "_", Box.Number), "EST.Weight.Shipment"], Diff.woo = Total.Weight - EST.Weight.woo, Diff.Shipment = Total.Weight - EST.Weight.Shipment)
discrepancy <- shipment_weight %>% filter(abs(Diff.woo) > Unit.Weight.woo/1000 & abs(Diff.Shipment) > Unit.Weight.Shipment/1000)
View(discrepancy)
discrepancy_size <- discrepancy %>% mutate(Category = gsub("-.*", "", SKU), Size = gsub(".*-", "", SKU)) %>% group_by(Category, Size) %>% summarise(N = n(), Diff.woo = mean(Diff.woo), Diff.Shipment = mean(Diff.Shipment))
View(discrepancy_size)
write.csv(discrepancy, file = paste0("../PO/shipment/Weight_discrepancy_", RefNo, ".csv"), row.names = F, quote = F)
write.csv(discrepancy_size, file = paste0("../PO/shipment/Weight_discrepancy_size_", RefNo, ".csv"), row.names = F, quote = F)

# ------------- upload Square SO to NS ---------------------------
customer <- read.csv("../Square/customers.csv", as.is = T) %>% `row.names<-`(.[, "Square.Customer.ID"])
payments <- read.csv(rownames(file.info(list.files(path = "../Square/", pattern = "transactions-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% filter(Event.Type != "Refund") %>% `row.names<-`(.[, "Payment.ID"])
square_so <- read.csv(rownames(file.info(list.files(path = "../Square/", pattern = "items-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% 
  mutate(Item = ifelse(Item %in% netsuite_item$Name, Item, "MISC-ITEM"), Cash = payments[Payment.ID, "Cash"], Recipient.Email = ifelse(is.na(Customer.ID), "", customer[Customer.ID, "Email.Address"]), Recipient.Phone = ifelse(is.na(Customer.ID), "", customer[Customer.ID, "Phone.Number"]))
netsuite_so <- square_so %>% filter(Item != "") %>% mutate(Order.date = format(as.Date(Date, "%Y-%m-%d"), "%m/%d/%Y")) %>% group_by(Order.date) %>% 
  mutate(Payment.Option = ifelse(Cash == "$0.00" | is.na(Cash), "Square", "Cash"), Class = "FBM : CA", MEMO = "Square sales", Customer = ifelse(Location == "Surrey", "55 JJR SHOPS", "54 JJR SHOPR"), Recipient = ifelse(is.na(Customer.Name), "", Customer.Name), Order.ID = Transaction.ID, Item.SKU = Item, ID = data.table::rleid(Order.ID), REF.ID = paste0("SQ", format(as.Date(Order.date, "%m/%d/%Y"), "%y%m%d"), "-", sprintf("%02d", ID)), Order.Type = ifelse(Location == "Surrey", "JJR SHOPS", "JJR SHOPR"), Department = ifelse(Location == "Surrey", "Retail : Store Surrey", "Retail : Store Richmond"), Warehouse = ifelse(Location == "Surrey", "WH-SURREY", "WH-RICHMOND"), Quantity = Qty, Price.level = "Custom", Coupon.Discount = as.numeric(gsub("\\$", "", Discounts)), Coupon.Discount = ifelse(Coupon.Discount == 0, "", as.character(Coupon.Discount)), Coupon.Code = "", Tax = as.numeric(gsub("\\$", "", Tax)), Rate = round(as.numeric(gsub("\\$", "", Net.Sales))/Qty, 2), Tax.Code = ifelse(round(Tax/as.numeric(gsub("\\$", "", Net.Sales)), 2) == 0.05, "CA-BC-GST", ifelse(round(Tax/as.numeric(gsub("\\$", "", Net.Sales)), 2) == 0.12, "CA-BC-TAX", "")), Tax.Amount = Tax, SHIPPING.CARRIER = "", SHIPPING.METHOD = "", SHIPPING.COST = 0) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, Price.level, Rate, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount, Customer, Recipient, Recipient.Phone, Recipient.Email, SHIPPING.CARRIER, SHIPPING.METHOD, SHIPPING.COST) %>% filter(Quantity > 0)
write.csv(netsuite_so, file = paste0("../Square/SO-square-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")
# upload to NS

# -------- Sync Square price to woo; inventory to NS S -------------
square <- data.frame(Token = "", ItemName = woo$SKU, ItemType = "", VariationName = "Regular", SKU = woo$SKU, Description = "", Category = netsuite_item_S[woo$SKU, "Item.Category.SKU"], Price = woo$Sale.price, PriceRichmond = "", PriceSurrey = "", NewQuantityRichmond = netsuite_item_S[woo$SKU, "Warehouse.Available"], NewQuantitySurrey = netsuite_item_S[woo$SKU, "Warehouse.Available"], EnabledRichmond = "Y", CurrentQuantityRichmond = "", CurrentQuantitySurrey = "", StockAlertEnabledRichmond = "N", StockAlertCountRichmond = 0, EnabledSurrey = "Y", StockAlertEnabledSurrey = "N", StockAlertCountSurrey = 0, SEOTitle = "",	SEODescription = "",	Permalink = "", SquareOnlineItemVisibility = "UNAVAILABLE",	ShippingEnabled = "",	SelfserveOrderingEnabled = "",	DeliveryEnabled = "",	PickupEnabled = "", Sellable = "Y",	Stockable = "Y",	SkipDetailScreeninPOS = "",	OptionName1 = "",	OptionValue1 = "", TaxPST = ifelse(woo$Tax.class == "full", "Y", "N")) %>%
  mutate(NewQuantityRichmond = ifelse(is.na(NewQuantityRichmond), 0, NewQuantityRichmond), NewQuantitySurrey = ifelse(is.na(NewQuantitySurrey), 0, NewQuantitySurrey))
write.table(square, file = paste0("../Square/square-upload-", Sys.Date(), ".csv"), sep = ",", row.names = F, col.names = c(colnames(square)[1:ncol(square)-1], "Tax - PST (7%)"), na = "")
# upload to Square > Items

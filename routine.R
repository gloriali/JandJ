# Routine tasks
library(dplyr)
library(stringr)
library(openxlsx)
library(xlsx)
library(scales)
library(ggplot2)
library(tidyr)
library(readr)

netsuite_item <- read.csv(rownames(file.info(list.files(path = "../NetSuite/", pattern = "Items_All_", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T)

# ------------- upload Square SO to NS: daily ---------------------------
customer <- read.csv("../Square/customers.csv", as.is = T) %>% `row.names<-`(.[, "Square.Customer.ID"])
payments <- read.csv(rownames(file.info(list.files(path = "../Square/", pattern = "transactions-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% `row.names<-`(.[, "Payment.ID"])
square_so <- read.csv(rownames(file.info(list.files(path = "../Square/", pattern = "items-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% 
  mutate(Item = ifelse(Item %in% netsuite_item$Name, Item, "MISC-ITEM"), Cash = payments[Payment.ID, "Cash"], Recipient.Email = ifelse(is.na(Customer.ID), "", customer[Customer.ID, "Email.Address"]), Recipient.Phone = ifelse(is.na(Customer.ID), "", customer[Customer.ID, "Phone.Number"]))
netsuite_so <- square_so %>% filter(Item != "") %>% mutate(Order.date = format(as.Date(Date, "%Y-%m-%d"), "%m/%d/%Y")) %>% group_by(Order.date) %>% 
  mutate(Payment.Option = ifelse(Cash == "$0.00" | is.na(Cash), "Square", "Cash"), Class = "FBM : CA", MEMO = "Square sales", Customer = ifelse(Location == "Surrey", "55 JJR SHOPS", "54 JJR SHOPR"), Recipient = ifelse(is.na(Customer.Name), "", Customer.Name), Order.ID = Transaction.ID, Item.SKU = Item, ID = data.table::rleid(Order.ID), REF.ID = paste0("SQ", format(as.Date(Order.date, "%m/%d/%Y"), "%y%m%d"), "-", sprintf("%02d", ID)), Order.Type = ifelse(Location == "Surrey", "JJR SHOPS", "JJR SHOPR"), Department = ifelse(Location == "Surrey", "Retail : Store Surrey", "Retail : Store Richmond"), Warehouse = ifelse(Location == "Surrey", "WH-SURREY", "WH-RICHMOND"), Quantity = Qty, Price.level = "Custom", Coupon.Discount = as.numeric(gsub("\\$", "", Discounts)), Coupon.Discount = ifelse(Coupon.Discount == 0, "", Coupon.Discount), Coupon.Code = "", Tax = as.numeric(gsub("\\$", "", Tax)), Rate = round(as.numeric(gsub("\\$", "", Net.Sales))/Qty, 2), Tax.Code = ifelse(round(Tax/as.numeric(gsub("\\$", "", Net.Sales)), 2) == 0.05, "CA-BC-GST", ifelse(round(Tax/as.numeric(gsub("\\$", "", Net.Sales)), 2) == 0.12, "CA-BC-TAX", "")), Tax.Amount = Tax, SHIPPING.CARRIER = "", SHIPPING.METHOD = "", SHIPPING.COST = 0) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, Price.level, Rate, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount, Customer, Recipient, Recipient.Phone, Recipient.Email, SHIPPING.CARRIER, SHIPPING.METHOD, SHIPPING.COST) %>% filter(Quantity > 0)
write.csv(netsuite_so, file = paste0("../Square/SO-square-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")
# upload to NS

# ------------- upload XHS SO to NS: weekly ---------------------------
XHS_so <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../XHS/", pattern = "order_export", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = 1) 
netsuite_so <- XHS_so %>% filter(Financial.Status == "paid") %>% mutate(Order.date = format(as.Date(gsub(" .*", "", Created.At), "%Y-%m-%d"), "%m/%d/%Y")) %>% 
  mutate(Payment.Option = "AlphaPay", Class = "FBM : CN", MEMO = "XHS sales", Customer = "56 XHS CN", Order.ID = Order.No, REF.ID = paste0("XHSCN-", Order.No), Order.Type = "XHS CN", Department = "Retail : Marketplace : XiaoHongShu", Warehouse = "WH-RICHMOND", Item.SKU = LineItem.SKU, Quantity = as.numeric(LineItem.Quantity), Price.level = "Custom", Rate = round((as.numeric(LineItem.Total) - as.numeric(LineItem.Total.Discount))/Quantity, 2), Coupon.Discount = as.numeric(LineItem.Total.Discount), Coupon.Discount = ifelse(Coupon.Discount == 0, "", Coupon.Discount), Coupon.Code = "", Tax.Amount = "", Tax.Code = "", Recipient = paste0(Shipping.Last.Name, Shipping.First.Name), Recipient.Phone = Phone, Recipient.Email = Email, SHIPPING.CARRIER = "Longxing", SHIPPING.METHOD = "", SHIPPING.COST = Shipping) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, Price.level, Rate, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount, Customer, Recipient, Recipient.Phone, Recipient.Email, SHIPPING.CARRIER, SHIPPING.METHOD, SHIPPING.COST)
Encoding(netsuite_so$Recipient) = "UTF-8"
write_excel_csv(netsuite_so, file = paste0("../XHS/SO-XHS-", format(Sys.Date(), "%Y%m%d"), ".csv"), na = "")
# upload to NS

# ------------- upload Clover SO to NS: weekly ---------------------------
customer <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "Customers-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>%
  mutate(Name = paste0(First.Name, " ", Last.Name)) %>% distinct(Email.Address, .keep_all = T) %>% filter(Name != " ", Email.Address != "") %>% `row.names<-`(.[, "Name"])
payments <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "Payments-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% mutate(Phone = customer[Customer.Name, "Phone.Number"], Email = customer[Customer.Name, "Email.Address"]) %>% filter(Result == "SUCCESS") %>% `row.names<-`(.[, "Order.ID"]) 
clover_so <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "LineItemsExport-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% mutate(Item.SKU = ifelse(Item.SKU %in% netsuite_item$Name, Item.SKU, "MISC-ITEM"), Recipient = payments[Order.ID, "Customer.Name"], Recipient.Phone = payments[Order.ID, "Phone"], Recipient.Email = payments[Order.ID, "Email"], Tender = payments[Order.ID, "Tender"])
netsuite_so <- clover_so %>% filter(Item.SKU != "", Order.Payment.State == "Paid", !grepl("XHS", Order.Discounts), !grepl("Prepaid", Order.Discounts), !grepl("Surrey", Order.Discounts)) %>% mutate(Order.date = format(as.Date(gsub(" .*", "", Line.Item.Date), "%d-%b-%Y"), "%m/%d/%Y")) %>% group_by(Order.date) %>% 
  mutate(Payment.Option = ifelse(Tender == "Cash", "Cash", ifelse(grepl("WeChat", Tender), "AlphaPay", "Clover")), Class = "FBM : CA", MEMO = "Clover sales", Customer = "54 JJR SHOPR", ID = data.table::rleid(Order.ID), REF.ID = paste0("CL", format(as.Date(Order.date, "%m/%d/%Y"), "%y%m%d"), "-", sprintf("%02d", ID)), Order.Type = "JJR SHOPR", Department = "Retail : Store Richmond", Warehouse = "WH-RICHMOND", Quantity = 1, Price.level = "Custom", Rate = Item.Total, Coupon.Discount = Total.Discount + Order.Discount.Proportion, Coupon.Discount = ifelse(Coupon.Discount == 0, NA, Coupon.Discount), Coupon.Code = gsub("NA", "", gsub(" -.*","", paste0(Order.Discounts, Discounts))), Tax.Code = ifelse(Item.Tax.Rate == 0.05, "CA-BC-GST", ifelse(Item.Tax.Rate == 0.12, "CA-BC-TAX", "")), SHIPPING.CARRIER = "", SHIPPING.METHOD = "", SHIPPING.COST = 0) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, Price.level, Rate, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount, Customer, Recipient, Recipient.Phone, Recipient.Email, SHIPPING.CARRIER, SHIPPING.METHOD, SHIPPING.COST) %>% filter(!is.na(Payment.Option))
write.csv(netsuite_so, file = paste0("../Clover/SO-clover-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")
# upload to NS

# ====================================================================

# input: woo > Products > All Products > Export > all
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price)) %>% `row.names<-`(toupper(.[, "SKU"])) 
# input: NS > Items > Export all warehouses
netsuite_item <- read.csv(list.files(path = "../NetSuite/", pattern = paste0("Items_All_", format(Sys.Date(), "%Y%m%d"), ".csv"), full.names = T), as.is = T)
netsuite_item[netsuite_item == "" | is.na(netsuite_item)] <- 0
netsuite_item_S <- netsuite_item %>% filter(Inventory.Warehouse == "WH-SURREY") %>% `row.names<-`(toupper(.[, "Name"])) 
netsuite_item_R <- netsuite_item %>% filter(Inventory.Warehouse == "WH-RICHMOND") %>% `row.names<-`(toupper(.[, "Name"])) 
write.csv(netsuite_item_S, file = paste0("../FBArefill/Raw Data File/Items_S_", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")

# -------- Sync XHS price to woo; inventory to NS S: weekly ----------------
# input: AllValue > Products > Export > All products
products_description <- read.xlsx2("../XHS/products_description.xlsx", sheetIndex = 1) %>% `row.names<-`(toupper(.[, "SPU"])) 
wb <- openxlsx::loadWorkbook(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T))
openxlsx::protectWorkbook(wb, protect = F)
openxlsx::saveWorkbook(wb, list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), overwrite = T)
products_XHS <- read.xlsx2(list.files(path = "../XHS/", pattern = paste0("products_export\\(", format(Sys.Date(), "%Y-%m-%d"), ".*.xlsx"), full.names = T), sheetIndex = 1)
products_XHS[products_XHS=="NA"] <- ""
products_upload <- products_XHS %>% mutate(Inventory = ifelse(netsuite_item_S[SKU, "Warehouse.Available"] < 10, 0, netsuite_item_S[SKU, "Warehouse.Available"]), Price = woo[SKU, "Sale.price"], Compare.At.Price = woo[SKU, "Regular.price"], Tags = ifelse(Price == Compare.At.Price, "正价", "特价"))
#products_upload <- products_XHS %>% mutate(Inventory = ifelse(xoro[SKU, "ATS"] < 10, 0, xoro[SKU, "ATS"]), Price = woo[SKU, "Regular.price"], Compare.At.Price = woo[SKU, "Regular.price"], Tags = ifelse(Price == Compare.At.Price, "正价", "特价"))
products_upload <- products_upload %>% mutate(Product.Name = products_description[toupper(SPU), "Product.Name"], SEO.Product.Name = Product.Name, Description = products_description[toupper(SPU), "Description"], Mobile.Description = products_description[toupper(SPU), "Description"], SEO.Description = products_description[toupper(SPU), "Description"], Describe = products_description[toupper(SPU), "Description"])
colnames(products_upload) <- gsub("\\.", " ", colnames(products_upload))
openxlsx::write.xlsx(products_upload, file = paste0("../XHS/products_upload-", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"), na.string = "")
# upload to AllValue > Products

# -------- Sync Square price to woo; inventory to NS S: weekly -------------
square <- data.frame(Token = "", ItemName = woo$SKU, ItemType = "", VariationName = "Regular", SKU = woo$SKU, Description = "", Category = netsuite_item_S[woo$SKU, "Item.Category.SKU"], Price = woo$Sale.price, PriceRichmond = "", PriceSurrey = "", NewQuantityRichmond = netsuite_item_S[woo$SKU, "Warehouse.Available"], NewQuantitySurrey = netsuite_item_S[woo$SKU, "Warehouse.Available"], EnabledRichmond = "Y", CurrentQuantityRichmond = "", CurrentQuantitySurrey = "", StockAlertEnabledRichmond = "N", StockAlertCountRichmond = 0, EnabledSurrey = "Y", StockAlertEnabledSurrey = "N", StockAlertCountSurrey = 0, SEOTitle = "",	SEODescription = "",	Permalink = "", SquareOnlineItemVisibility = "UNAVAILABLE",	ShippingEnabled = "",	SelfserveOrderingEnabled = "",	DeliveryEnabled = "",	PickupEnabled = "", Sellable = "Y",	Stockable = "Y",	SkipDetailScreeninPOS = "",	OptionName1 = "",	OptionValue1 = "", TaxPST = ifelse(woo$Tax.class == "full", "Y", "N")) %>%
  mutate(NewQuantityRichmond = ifelse(is.na(NewQuantityRichmond), 0, NewQuantityRichmond), NewQuantitySurrey = ifelse(is.na(NewQuantitySurrey), 0, NewQuantitySurrey))
write.table(square, file = paste0("../Square/square-upload-", Sys.Date(), ".csv"), sep = ",", row.names = F, col.names = c(colnames(square)[1:ncol(square)-1], "Tax - PST (7%)"), na = "")
# upload to Square > Items

# -------- Sync Clover price to woo: weekly -------------
# input Clover > Inventory > Items > Export.
price <- woo %>% mutate(cat = gsub("-.*", "", SKU)) %>% group_by(cat) %>% summarise(Price = max(Sale.price)) %>% as.data.frame()
price <- rbind(price, data.frame(cat = c("MISC5", "MISC10", "MISC15", "MISC20", "DBRC", "DBTB", "DBTL", "DBTP", "DLBS", "DWJA", "DWJT", "DWPF", "DWPS", "DWSF", "DWSS", "DXBK"), Price = c(5, 10, 15, 20, 30, 35, 40, 40, 25, 50, 40, 30, 25, 60, 55, 30))) %>% `row.names<-`(toupper(.[, "cat"])) 
#clover <- openxlsx::loadWorkbook(rownames(file.info(list.files(path = "../Clover/", pattern = "inventory[0-9]+.xlsx", full.names = TRUE)) %>% filter(mtime == max(mtime))))
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% filter(Name != "") %>% 
  mutate(cat = toupper(gsub("-.*", "", Name)), Price = ifelse(toupper(Name) %in% toupper(woo$SKU), woo[Name, "Sale.price"], price[cat, "Price"]), Price.Type = ifelse(is.na(Price), "Variable", "Fixed"), Alternate.Name = woo[Name, "Name"], Tax.Rates = ifelse(woo[Name, "Tax.class"] == "full", "GST+PST", "GST"), Tax.Rates = ifelse(is.na(Tax.Rates), "GST", Tax.Rates)) %>% select(-cat)
clover_item <- clover_item %>% mutate(Price = ifelse(grepl("^SS", Name), 10, Price))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item) + 100, gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
# upload to Clover > Inventory

# ------------- Sync NS Richmond inventory with Clover: weekly ---------------------------
clover_item <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items") %>% filter(Name %in% netsuite_item_R$Name) %>% mutate(Quantity = ifelse(Quantity < 0, 0, Quantity))
inventory_update <- clover_item %>% mutate(Date = format(Sys.Date(), "%m/%d/%Y"), ITEM = Name, ADJUST.QTY.BY = ifelse(toupper(Name) %in% rownames(netsuite_item_R), Quantity - netsuite_item_R[toupper(Name), "Warehouse.On.Hand"], Quantity), Reason.Code = "CC", MEMO = "Clover Inventory Update", WAREHOUSE = "WH-RICHMOND", External.ID = paste0("IAR-", format(Sys.Date(), "%y%m%d"), "-1")) %>%
  select(Date, ITEM, ADJUST.QTY.BY, Reason.Code, MEMO, WAREHOUSE, External.ID) %>% filter(ADJUST.QTY.BY != 0)
colnames(inventory_update) <- gsub("QTY BY", "QTY. BY", gsub("\\.", " ", colnames(inventory_update)))
write.csv(inventory_update, file = paste0("../NetSuite/IA-Richmond-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")
## if there are unsuccessful Clover SO upload due to NS qty=0 
SO_error <- read.csv("../NetSuite/results.csv", as.is = T) %>% group_by(Item.SKU) %>% summarise(Qty = sum(Quantity)) %>% as.data.frame() %>% `row.names<-`(toupper(.[, "Item.SKU"]))
clover_item <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items") %>% filter(Name %in% netsuite_item$Name) %>% mutate(Quantity = ifelse(Quantity < 0, 0, Quantity))
inventory_update <- clover_item %>% mutate(STATUS = "Good", Date = format(Sys.Date(), "%m/%d/%Y"), ITEM = Name, ADJUST.QTY.BY = ifelse(toupper(Name) %in% rownames(netsuite_item_R), Quantity - netsuite_item_R[toupper(Name), "Warehouse.On.Hand"], Quantity), ADJUST.QTY.BY = ifelse(ITEM %in% rownames(SO_error), ADJUST.QTY.BY + SO_error[ITEM, "Qty"], ADJUST.QTY.BY), Reason.Code = "CC", MEMO = "Clover Inventory Update", WAREHOUSE = "WH-RICHMOND", External.ID = paste0("IAR-", format(Sys.Date(), "%y%m%d"), "-1")) %>%
  select(Date, ITEM, ADJUST.QTY.BY, Reason.Code, STATUS, MEMO, WAREHOUSE, External.ID) %>% filter(ADJUST.QTY.BY != 0)
colnames(inventory_update) <- gsub("QTY BY", "QTY. BY", gsub("\\.", " ", colnames(inventory_update)))
write.csv(inventory_update, file = paste0("../NetSuite/IA-Richmond-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")

# -------- Add POS sales to yotpo rewards program: weekly -----------
# customer info: Clover > Customers > Download
# sales: Clover > Transactions > Payments > select dates > Export
customer <- read.csv(list.files(path = "../Clover/", pattern = paste0("Customers-", format(Sys.Date(), "%Y%m%d")), full.names = T), as.is = T) %>%
  mutate(Name = paste0(First.Name, " ", Last.Name)) %>% distinct(Email.Address, .keep_all = T) %>% filter(Name != " ", Email.Address != "") %>% `row.names<-`(.[, "Name"])
payments <- read.csv(list.files(path = "../Clover/", pattern = paste0("Payments-", format(Sys.Date(), "%Y%m%d")), full.names = T), as.is = T) %>% 
  mutate(email = customer[Customer.Name, "Email.Address"], Refund.Amount = ifelse(is.na(Refund.Amount), 0, as.numeric(Refund.Amount))) %>% filter(!is.na(email))
point <- data.frame(Email = payments$email, Points = as.integer(payments$Amount - payments$Refund.Amount)) %>% filter(Points != 0, !is.na(Email)) %>% mutate(Email = gsub(",.*", "", Email), Email = gsub('\\"', "", Email))
write.csv(point, file = paste0("../yotpo/", format(Sys.Date(), "%m%d%Y"), "-yotpo.csv"), row.names = F)
# customer info: Square > Customers > Export customers
# sales: Square > Transactions > select dates > Export Transactions CSV
#square_customer <- read.csv("../Square/customers.csv", as.is = T) %>% `row.names<-`(.[, "Square.Customer.ID"]) 
#transactions <- read.csv("../Square/transactions-2024-03-20-2024-04-11.csv", as.is = T) %>% filter(Customer.ID != "")
#point <- bind_rows(point, data.frame(Email = square_customer[transactions$Customer.ID, "Email.Address"], Points = as.integer(gsub("\\$", "", transactions$Total.Collected)))) %>% filter(Points != 0, !is.na(Email))
# upload to Yotpo
non_included <- read.csv("../yotpo/non_included.csv", as.is = T)
error <- read.csv("../yotpo/error_report.csv", as.is = T)
non_included <- bind_rows(non_included, error %>% select(Email, Points)) %>% count(Email, wt = Points)
write.table(non_included, file = "../yotpo/non_included.csv", sep = ",", row.names = F, col.names = c("Email", "Points"))

# -------- Request stock from Surrey: at request --------------------
# input Clover > Inventory > Items > Export.
netsuite_item <- read.csv(list.files(path = "../NetSuite/", pattern = paste0("Items_All_", format(Sys.Date(), "%Y%m%d"), ".csv"), full.names = T), as.is = T)
netsuite_item[netsuite_item == "" | is.na(netsuite_item)] <- 0
netsuite_item_S <- netsuite_item %>% filter(Inventory.Warehouse == "WH-SURREY") %>% `row.names<-`(toupper(.[, "Name"])) 
request <- c("SWS", "BRC", "BSL", "BTB", "BTL", "BTT", "BST", "WJA", "WJT", "WPF", "WPS", "WBF", "WBS", "WGS", "WMT", "WSF", "WSS", "XBK", "XBM", "XLB", "XPC", "SKG", "SKB", "SKX", "IHT", "FHA", "IPC", "ISJ", "FAN", "FJM", "FPM", "DRC", "KEH", "KMT", "LBT", "LBP") # categories to restock
n <- 3       # Qty per SKU to stock at Richmond
n_S <- 8 # min Qty in stock at Surrey to request
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- readWorkbook(clover, "Items") %>% mutate(cat = gsub("-.*", "", Name), Quantity = ifelse(is.na(Quantity) | Quantity < 0, 0, Quantity)) %>% filter(!duplicated(Name), !is.na(Name)) %>% `row.names<-`(toupper(.[, "Name"])) 
# order <- data.frame(StoreCode = "WH-JJ", ItemNumber=(clover_item %>% filter(Quantity < n & cat %in% request))$Name, Qty = n - clover_item[(clover_item %>% filter(Quantity < n & cat %in% request))$Name, "Quantity"], LocationName = "BIN", UnitCost = "", ReasonCode = "RWT", Memo = "Richmond Transfer to Miranda", UploadRule = "D", AdjAccntName = "", TxnDate = "", ItemIdentifierCode = "", ImportError = "")
# order <- order %>% filter(xoro[order$ItemNumber, "ATS"] > n_xoro) %>% filter(Qty > 0) %>% mutate(cat = gsub("-.*", "", ItemNumber), size = gsub("\\w+-\\w+-", "", ItemNumber)) %>% arrange(cat, size) %>% select(-c("cat", "size"))
order <- data.frame(Date = format(Sys.Date(), "%m/%d/%Y"), TO.TYPE = "Surrey-Richmond", SEASON = "24F", FROM.WAREHOUSE = "WH-SURREY", TO.WAREHOUSE = "WH-RICHMOND", REF.NO = paste0("TO-S2R", format(Sys.Date(), "%y%m%d")), Memo = "Richmond Refill", ORDER.PLACED.BY = "Gloria Li", ITEM = (clover_item %>% filter(Quantity < n))$Name) %>% 
  mutate(Quantity = n - clover_item[ITEM, "Quantity"]) %>% filter(ITEM %in% netsuite_item_S$Name, netsuite_item_S[ITEM, "Warehouse.Available"] > n_S) %>% filter(Quantity > 1) %>% mutate(cat = gsub("-.*", "", ITEM), size = gsub("\\w+-\\w+-", "", ITEM))
order <- order %>% filter(cat %in% request)
order <- order %>% rename_with(~ gsub("\\.", " ", colnames(order))) %>% select(-c("cat", "size")) 
write.csv(order, file = paste0("../Clover/order", format(Sys.Date(), "%m%d%Y"), ".csv"), row.names = F, na = "")
# upload to NS & email Shikshit

# ---------------- Adjust Clover Inventory: at request --------------------
# download current Richmond stock: clover_item > Inventory > Items > Export
library(dplyr)
library(openxlsx)
adjust_inventory <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "order", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% `row.names<-`(toupper(.[, "ITEM"])) 
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- readWorkbook(clover, "Items") %>% mutate(SKU = toupper(SKU), Quantity = ifelse(SKU %in% rownames(adjust_inventory), Quantity + adjust_inventory[SKU, "Quantity"], Quantity))
clover_item <- clover_item %>% rename_with(~ gsub("\\.", " ", colnames(clover_item)))
deleteData(clover, sheet = "Items", cols = 1:ncol(clover_item), rows = 1:nrow(clover_item), gridExpand = T)
writeData(clover, sheet = "Items", clover_item)
openxlsx::saveWorkbook(clover, file = paste0("../Clover/inventory", format(Sys.Date(), "%Y%m%d"), "-upload.xlsx"), overwrite = T)
# upload to Clover > Inventory

# -------- Analysis: monthly --------------------
new_season <- "24"   # New season contains (yr)
qty_offline <- 3     # Qty to move to offline sales
month <- format(Sys.Date(), "%m")
if(month %in% c("09", "10", "11", "12", "01", "02")){in_season <- "F"}else(in_season <- "S")
RawData <- list.files(path = "../FBArefill/Raw Data File/", pattern = "All Marketplace by", full.names = T)
mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T)[1], sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
mastersku_adjust <- mastersku %>% filter(!duplicated(Adjust.SKU)) %>% `row.names<-`(toupper(.[, "Adjust.SKU"])) 
netsuite_item_S <- read.csv(list.files(path = "../FBArefill/Raw Data File/", pattern = "Items_S_.*", full.names = T), as.is = T) %>%
  mutate(Name = toupper(Name), Seasons = ifelse(Name %in% mastersku$MSKU, mastersku[Name, "Seasons.SKU"], mastersku_adjust[Name, "Seasons.SKU"]), SPU = gsub("(\\w+-\\w+)-.*", "\\1", Name)) %>% `row.names<-`(.[, "Name"])
## Raw data summary
monR_last_yr <- read.csv(list.files(path = "../Analysis/", pattern = "MonthlyRatio2023_category_adjust_2024-01-25.csv", full.names = T), as.is = T) %>% `row.names<-`(.[, "Category"])
sheets <- grep("-WK", getSheetNames(RawData), value = T)
inventory <- data.frame()
for(sheet in sheets){
  inventory_i <- openxlsx::read.xlsx(RawData, sheet = sheet, startRow = 2)
  inventory <- rbind(inventory, inventory_i)
}
inventory[is.na(inventory)] <- 0
inventory <- inventory %>% mutate(Inv_Total_End.WK = `Inv_Total_End.WK.(Not.Inc..CN.RM)` + Inv_WH.China + Inv_WH.Richmond)
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
## SKUs to move offline
offline <- netsuite_item_S %>% filter(!grepl(new_season, Seasons) & Warehouse.Available <= qty_offline & Warehouse.Available > 0) %>% arrange(Name) %>% 
  mutate(Date = format(Sys.Date(), "%m/%d/%Y"), TO.TYPE = "Surrey-Richmond", SEASON = "24F", FROM.WAREHOUSE = "WH-SURREY", TO.WAREHOUSE = "WH-RICHMOND", REF.NO = paste0("TO-S2R", format(Sys.Date(), "%y%m%d")), Memo = "Qty <= 3 items to transfer to Richmond ", ORDER.PLACED.BY = "Gloria Li", ITEM = Name, Quantity = Warehouse.Available) %>% select(Date:Quantity)
offline <- offline %>% rename_with(~ gsub("\\.", " ", colnames(offline)))
write.csv(offline, file = paste0("../Analysis/offline_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\
## SKUs to inactive
clover <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items")  %>% filter(!is.na(Name)) %>% `row.names<-`(toupper(.[, "Name"]))
inventory_s <- inventory %>% count(Adjust_MSKU, wt = Inv_Total_End.WK) %>% `row.names<-`(toupper(.[, "Adjust_MSKU"]))
qty0 <- mastersku %>% select(MSKU.Status, Seasons.SKU, MSKU) %>% mutate(MSKU = toupper(MSKU), Inv_clover = clover[MSKU, "Quantity"], Inv_Total_EndWK = inventory_s[MSKU, "n"]) %>% 
  filter(!grepl(new_season, Seasons.SKU) & MSKU.Status == "Active" & Inv_clover == 0 & Inv_Total_EndWK == 0)
write.csv(qty0, file = paste0("../Analysis/discontinued_qty0_", Sys.Date(), ".csv"), row.names = F, na = "")
# copy to TWK Analysis\0 - Analysis to Share - Sales and Inventory\
## Suggested deals for the website
discount_method <- openxlsx::read.xlsx("../Analysis/JJ_discount_method.xlsx", sheet = 1, startRow = 2, rowNames = T)
if(month == "01"){last3m <- c("Month11", "Month12", "Month01")}else if(month %in% c("02", "03")){last3m <- sprintf("Month%02d", c(1:(as.numeric(month)-1), (as.numeric(month)-3+12):12))}else{last3m <- sprintf("Month%02d", c((as.numeric(month)-3):(as.numeric(month)-1)))}
monR <- monR %>% mutate(T.last3m = rowSums(across(all_of(last3m)))) %>% `row.names<-`(.[, "Category"])
inventory_SPU <- inventory %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU)) %>% group_by(SPU) %>% summarise(qty_SPU = sum(Inv_Total_End.WK, na.rm = T), qty_WH = sum(Inv_WH.JJ, na.rm = T), qty_AMZ = qty_SPU - qty_WH) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SKU_last12month <- sales_SKU_last12month %>% mutate(SPU = gsub("(\\w+-\\w+)-.*", "\\1", Adjust_MSKU), T.last3m = rowSums(across(all_of(last3m))), T.yr = rowSums(across(Month01:Month12)))
sales_SPU_last3month <- sales_SKU_last12month %>% group_by(SPU) %>% summarise(T.last3m = sum(T.last3m)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SPU_last3month_JJ <- sales_SKU_last12month %>% filter(Sales.Channel == "janandjul.com") %>% group_by(SPU) %>% summarise(T.last3m = sum(T.last3m)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sales_SPU_last12month_AMZ <- sales_SKU_last12month %>% filter(grepl("Amazon", Sales.Channel)) %>% group_by(SPU) %>% summarise(T.last12m = sum(T.yr)) %>% as.data.frame() %>% `row.names<-`(.[, "SPU"])
sizes_all <- netsuite_item_S %>% mutate(SPU_size = gsub("(\\w+-\\w+-\\w+).*", "\\1", Name)) %>% filter(!duplicated(SPU_size)) %>% count(SPU) %>% `row.names<-`(.[, "SPU"])
size_limited <- netsuite_item_S %>% filter(Warehouse.Available > qty_offline) %>% count(SPU) %>% mutate(all = sizes_all[SPU, "n"], percent = ifelse(all < n, 0, (all - n)/all)) %>% `row.names<-`(.[, "SPU"])
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% filter(Published == 1, !is.na(Regular.price)) %>% filter(!duplicated(SKU), SKU != "") %>% 
  mutate(SKU = toupper(SKU), Qty = netsuite_item_S[SKU, "Warehouse.Available"], Seasons = ifelse(SKU %in% mastersku$MSKU, mastersku[SKU, "Seasons.SKU"], mastersku_adjust[SKU, "Seasons.SKU"]), discount = ifelse(is.na(Sale.price), 0, (Regular.price - Sale.price)/Regular.price), cat = mastersku[SKU, "Category.SKU"], SPU = gsub("(\\w+-\\w+)-.*", "\\1", SKU)) %>% 
  select(ID, SKU, Name, Seasons, cat, SPU, Regular.price, discount, Qty) %>% `row.names<-`(.[, "SKU"])
overstock <- woo %>% filter(!grepl(new_season, Seasons), Qty > qty_offline, Regular.price > 0) %>% mutate(Qty_SPU = inventory_SPU[SPU, "qty_SPU"], Qty_WH = inventory_SPU[SPU, "qty_WH"], Qty_AMZ = inventory_SPU[SPU, "qty_AMZ"], MonR_last3m = monR[cat, "T.last3m"], Sales_SPU_last3m = sales_SPU_last3month[SPU, "T.last3m"], Sales_SPU_last3m_JJ = sales_SPU_last3month_JJ[SPU, "T.last3m"], Sales_SPU_1yr = Sales_SPU_last3m/MonR_last3m, Sales_SPU_1yr_JJ = Sales_SPU_last3m_JJ/MonR_last3m, Sales_SPU_1yr_AMZ = sales_SPU_last12month_AMZ[SPU, "T.last12m"]) %>% 
  filter(!is.na(Qty_SPU)) %>% mutate(time_yrs = round(ifelse(Qty_AMZ >= Sales_SPU_1yr_AMZ, Qty_WH/Sales_SPU_1yr_JJ, Qty_SPU/Sales_SPU_1yr), 1), size_percent_missing = gsub("Sizes.missing.0%", "Sizes.missing.<.50%", paste0("Sizes.missing.", as.character(as.integer(ifelse(size_limited[SPU, "percent"] < 0.5, 0, size_limited[SPU, "percent"])*10)*10), "%"))) %>% 
  filter(time_yrs >= 1) %>% select(SKU:Qty, time_yrs, size_percent_missing) %>% arrange(SKU)
write.csv(overstock, file = paste0("../Analysis/Overstock_", Sys.Date(), ".csv"), row.names = F)
# Email Will
woo_deals <- woo %>% filter(!grepl(new_season, Seasons), Qty > qty_offline, Regular.price > 0) %>% mutate(Qty_SPU = inventory_SPU[SPU, "qty_SPU"], Qty_WH = inventory_SPU[SPU, "qty_WH"], Qty_AMZ = inventory_SPU[SPU, "qty_AMZ"], MonR_last3m = monR[cat, "T.last3m"], Sales_SPU_last3m = sales_SPU_last3month[SPU, "T.last3m"], Sales_SPU_last3m_JJ = sales_SPU_last3month_JJ[SPU, "T.last3m"], Sales_SPU_1yr = Sales_SPU_last3m/MonR_last3m, Sales_SPU_1yr_JJ = Sales_SPU_last3m_JJ/MonR_last3m, Sales_SPU_1yr_AMZ = sales_SPU_last12month_AMZ[SPU, "T.last12m"]) %>% 
  filter(!is.na(Qty_SPU)) %>% mutate(time_yrs = ifelse(Qty_AMZ >= Sales_SPU_1yr_AMZ, Qty_WH/Sales_SPU_1yr_JJ, Qty_SPU/Sales_SPU_1yr), time_yrs = paste0("Sold out ", as.character(as.integer(ifelse(time_yrs > 3, 3, time_yrs)) + 1), " yr"), size_percent_missing = gsub("Sizes.missing.0%", "Sizes.missing.<.50%", paste0("Sizes.missing.", as.character(as.integer(ifelse(size_limited[SPU, "percent"] < 0.5, 0, size_limited[SPU, "percent"])*10)*10), "%"))) %>%
  rowwise() %>% mutate(Suggest_discount = as.numeric(gsub("% Off", "", discount_method[time_yrs, size_percent_missing]))/100, Suggest_price = round((1 - Suggest_discount) * Regular.price, digits = 2)) %>% arrange(-Suggest_discount, SKU) %>% filter(Suggest_discount > 0, Suggest_discount > discount) 
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

# -------- Prepare to generate barcode image: at request ------------
library(dplyr)
library(openxlsx)
library(stringr)
season <- "2025SS"
startRow <- 9
mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"])) 
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

# -------- Sync master file barcode with clover: at request -------------
# master SKU file: OneDrive > TWK 2020 share
# download clover inventory: clover_item > Inventory > Items > Export.
library(dplyr)
library(openxlsx)
library(tidyr)
woo <- read.csv(list.files(path = "../woo/", pattern = gsub("-0", "-", paste0("wc-product-export-", format(Sys.Date(), "%d-%m-%Y"))), full.names = T), as.is = T) %>% 
  filter(!is.na(Regular.price) & !duplicated(SKU) & SKU != "") %>% mutate(Sale.price = ifelse(is.na(Sale.price), Regular.price, Sale.price), cat = toupper(gsub("-.*", "", SKU))) %>% `row.names<-`(toupper(.[, "SKU"])) 
price <- woo %>% mutate(cat = gsub("-.*", "", SKU)) %>% group_by(cat) %>% summarise(Price = max(Sale.price)) %>% as.data.frame()
price <- rbind(price, data.frame(cat = c("MISC5", "MISC10", "MISC15", "MISC20", "DBRC", "DBTB", "DBTL", "DBTP", "DLBS", "DWJA", "DWJT", "DWPF", "DWPS", "DWSF", "DWSS", "DXBK"), Price = c(5, 10, 15, 20, 30, 35, 40, 40, 25, 50, 40, 30, 25, 60, 55, 30))) %>% `row.names<-`(toupper(.[, "cat"])) 
mastersku <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../FBArefill/Raw Data File/", pattern = "1-MasterSKU-All-Product-", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
clover <- openxlsx::loadWorkbook(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T))
clover_item <- openxlsx::readWorkbook(clover, "Items") %>% filter(!is.na(Name), !duplicated(toupper(Name))) %>% `row.names<-`(toupper(.[, "Name"])) 
category_exclude <- c("FVMU", "PJWU","PLAU", "PTAU", "STORE", "TPLU", "TPSU", "TSWU", "TTLU", "TTSU", "WJAU", "WPSU", "XBMU")
clover_item_upload <- data.frame(Clover.ID = "", Name = toupper(mastersku[mastersku$MSKU.Status == "Active", "MSKU"])) %>% filter(!is.na(Name)) %>% 
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

# ------------ upload regular POs -------------------
season <- "25S"
PO_NS <- data.frame(); total <- 0; items <- 0
for(p in 370:403){
  ID <- paste0("P", p)
  file <- list.files(path = paste0("../PO/order/", season, "/"), pattern = paste0(ID, ".*.xlsx"), full.names = TRUE, recursive = T)
  if(length(file) != 1){print(paste(ID, "INPUT FILE ERROR:", length(file), "files found")); next}
  memo <- gsub("\\.xlsx", "",gsub(paste0(".*", ID, " *\\- *"), "", file))
  print(paste(ID, memo))
  PO <- openxlsx::read.xlsx(file, sheet = 1, startRow = 8) 
  if(sum(grepl("加拿大总数量", colnames(PO))) == 1){
    ReceiveDate <- format(as.Date(gsub("出货日期.ShipDate.", "", colnames(PO)[which(grepl("加拿大总数量", colnames(PO)))+1]), "%Y.%m.%d") + 40, "%m/%d/%Y")
    PO_CA <- PO %>% select("产品编号", "加拿大总数量") %>% filter(grepl("-.*-", 产品编号)) %>% mutate(cat = gsub("-.*", "", 产品编号))
    CAT <- paste((PO_CA %>% distinct(cat))$cat, collapse = " ")
    PO_CA_NS <- PO_CA %>% mutate(PO.TYPE = "Regular", CATEGORY = CAT, SEASON = season, REF.NO = paste0(ID, "-CA"), WAREHOUSE = "WH-SURREY", VENDOR = "China", CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = memo, ITEM = 产品编号,  QUANTITY = as.numeric(ifelse(is.na(加拿大总数量), 0, 加拿大总数量)), Tax.Code = "CA-Zero", External.ID = paste0(ID, "-CA")) %>% 
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
    ReceiveDate <- format(as.Date(gsub("出货日期.ShipDate.", "", colnames(PO)[which(grepl("英国总数量", colnames(PO)))+1]), "%Y.%m.%d") + 40, "%m/%d/%Y")
    PO_UK <- PO %>% select("产品编号", "英国总数量") %>% filter(grepl("-.*-", 产品编号)) %>% mutate(cat = gsub("-.*", "", 产品编号))
    CAT <- paste((PO_UK %>% distinct(cat))$cat, collapse = " ")
    PO_UK_NS <- PO_UK %>% mutate(PO.TYPE = "Regular", CATEGORY = CAT, SEASON = season, REF.NO = paste0(ID, "-UK"), WAREHOUSE = "WH-AMZ : FBA-UK", VENDOR = "China", CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = memo, ITEM = 产品编号,  QUANTITY = as.numeric(ifelse(is.na(英国总数量), 0, 英国总数量)), Tax.Code = "CA-Zero", External.ID = paste0(ID, "-UK")) %>% 
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
    ReceiveDate <- format(as.Date(gsub("出货日期.ShipDate.", "", colnames(PO)[which(grepl("德国总数量", colnames(PO)))+1]), "%Y.%m.%d") + 40, "%m/%d/%Y")
    PO_DE <- PO %>% select("产品编号", "德国总数量") %>% filter(grepl("-.*-", 产品编号)) %>% mutate(cat = gsub("-.*", "", 产品编号))
    CAT <- paste((PO_DE %>% distinct(cat))$cat, collapse = " ")
    PO_DE_NS <- PO_DE %>% mutate(PO.TYPE = "Regular", CATEGORY = CAT, SEASON = season, REF.NO = paste0(ID, "-DE"), WAREHOUSE = "WH-AMZ : FBA-DE", VENDOR = "China", CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = memo, ITEM = 产品编号,  QUANTITY = as.numeric(ifelse(is.na(德国总数量), 0, 德国总数量)), Tax.Code = "CA-Zero", External.ID = paste0(ID, "-DE")) %>% 
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
  if(sum(grepl("中国总数量", colnames(PO))) == 1){
    ReceiveDate <- format(as.Date(gsub("出货日期.ShipDate.", "", colnames(PO)[which(grepl("中国总数量", colnames(PO)))+1]), "%Y.%m.%d") + 40, "%m/%d/%Y")
    PO_CN <- PO %>% select("产品编号", "中国总数量") %>% filter(grepl("-.*-", 产品编号)) %>% mutate(cat = gsub("-.*", "", 产品编号))
    CAT <- paste((PO_CN %>% distinct(cat))$cat, collapse = " ")
    PO_CN_NS <- PO_CN %>% mutate(PO.TYPE = "Regular", CATEGORY = CAT, SEASON = season, REF.NO = paste0(ID, "-CN"), WAREHOUSE = "WH-CHINA", VENDOR = "China", CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = memo, ITEM = 产品编号,  QUANTITY = as.numeric(ifelse(is.na(中国总数量), 0, 中国总数量)), Tax.Code = "CA-Zero", External.ID = paste0(ID, "-CN")) %>% 
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
}
print(paste("Total No. of items: ", nrow(PO_NS), items, "; Total Qty: ", sum(PO_NS$QUANTITY), total))
write_excel_csv(PO_NS, file = paste0("../PO/order/", season, "/NS_PO_regular_", format(Sys.Date(), "%Y%m%d"), ".csv"), na = "")

# ------------ upload wholesaler POs: CZ-Mylerie, CA-Clement, FR-Petits -------------------
ID <- "P406"; season <- "25S"; type <- "CA-Clement"; warehouse <- "WH-SURREY"
file <- list.files(path = paste0("../PO/order/", season, "/"), pattern = paste0(ID, ".*.xlsx"), full.names = TRUE, recursive = T)
PO_NS <- data.frame(); total <- 0; items <- 0; CAT <- c()
for(f in file){
  sheets <- getSheetNames(f)[getSheetNames(f) != "Export Summary"]
  memo <- gsub(paste0(".*", ID, " *\\- *"), "", gsub(paste0("\\/", ID, "-.*\\.xlsx"), "", f))
  for(s in sheets){
    print(c(memo, s))
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
ID <- "CEFA5"; season <- "24F"; ReceiveDate <- "11/20/2024"
PO <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../PO/order/CEFA/", pattern = "Order form.*.xlsx", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = 1, startRow = 3) %>% 
  filter(grepl("-.*-", SKU.Number)) %>% mutate(Quantity = ifelse(is.na(Quantity), 0, Quantity), cat = gsub("-.*", "", SKU.Number))
CAT <- paste((PO %>% distinct(cat))$cat, collapse = " ")
PO_NS <- PO %>% mutate(PO.TYPE = "Daycare Uniforms", CATEGORY = CAT, SEASON = season, REF.NO = ID, WAREHOUSE = "WH-SURREY", VENDOR = "China", CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = paste0("CEFA order for ", season), ITEM = SKU.Number,  QUANTITY = Quantity, Tax.Code = "CA-Zero", External.ID = REF.NO) %>% 
  filter(Quantity > 0) %>% select(PO.TYPE:External.ID) %>% rename_with(~ gsub("\\.", " ", .))
write.csv(PO_NS, file = paste0("../PO/order/CEFA/", "NS_PO_", ID, ".csv"), row.names = F, na = "")

# ------------ upload inbound shipment for POs -------------------
library(tidyr)
season <- "24F"; warehouse <- "WH-SURREY"
RefNo <- "24FWCA14"; ShippingDate <- "10/25/2024"; ReceiveDate <- "11/20/2024"
PO_detail <- read.csv(rownames(file.info(list.files(path = "../PO/", pattern = "PurchaseOrders", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% filter(Item != "", Warehouse == warehouse) %>% 
  mutate(Quantity = as.numeric(Quantity), Quantity.Fulfilled.Received = as.numeric(Quantity.Fulfilled.Received), Quantity.on.Shipments = ifelse(is.na(as.numeric(Quantity.on.Shipments)), 0, as.numeric(Quantity.on.Shipments)), Quantity.Remain = Quantity - Quantity.on.Shipments - Quantity.Fulfilled.Received, Quantity.Remain = ifelse(Quantity.Remain < 0, 0, Quantity.Remain)) %>% `row.names<-`(paste0(.[, "REF.NO"], "_", .[, "Item"]))
POn <- PO_detail %>% filter(!duplicated(REF.NO)) %>% `row.names<-`(.[, "REF.NO"])
shipment_in <- list.files(path = "../PO/shipment/", pattern = paste0(".*", RefNo, ".*.xlsx"), recursive = T, full.names = T)
memo <- gsub("ETA.*\\/.*", "", gsub("../PO/shipment/", "", shipment_in))
sheets <- getSheetNames(shipment_in)[]
summary <- openxlsx::read.xlsx(shipment_in, sheet = 1, fillMergedCells = T)
shipment <- data.frame(); attachment <- data.frame(); Nbox <- 0
for(s in 2:length(sheets)){
  sheet <- sheets[s]
  shipment_i <- openxlsx::read.xlsx(shipment_in, sheet = sheet, startRow = 5, fillMergedCells = T) %>% 
    select(1:Season) %>% filter(grepl(".*\\-.*", SKU)) %>% tidyr::fill(1:Season, .direction = "down")
  #shipment_i$`PO.#` <- paste0("P", RefNo)
  if(sum(grepl("CTN.NO", colnames(shipment_i))) == 1){
    print(paste(1, sheet))
    shipment_o <- data.frame(REF.NO = RefNo, EXPECTED.SHIPPING.DATE = ShippingDate, EXPECTED.DELIVERY.DATE = ReceiveDate, MEMO = memo, PO.REF.NO = shipment_i$`PO.#`, BOX.NO = paste0(s, "-", shipment_i$`CTN.NO`), ITEM = shipment_i$SKU, QUANTITY = shipment_i$QTY.SHIPPED, LOCATION = warehouse) %>% mutate(PO = paste0("PO#", POn[PO.REF.NO, "Document.Number"]))
  }else{
    print(paste(2, sheet))
    colnames(shipment_i) <- c("Start", "End", colnames(shipment_i)[3:length(colnames(shipment_i))])
    shipment_i <- shipment_i %>% mutate(Start = as.numeric(Start), End = as.numeric(End), QTY.Per.Box = as.numeric(QTY.Per.Box))
    shipment_o <- data.frame()
    for(i in 1:nrow(shipment_i)){
      for(n in shipment_i[i, "Start"]:shipment_i[i, "End"]){
        shipment_o <- rbind(shipment_o, data.frame(REF.NO = RefNo, EXPECTED.SHIPPING.DATE = ShippingDate, EXPECTED.DELIVERY.DATE = ReceiveDate, MEMO = memo, PO.REF.NO = shipment_i[i, "PO.#"], BOX.NO = paste0(s, "-", n), ITEM = shipment_i[i, "SKU"], QUANTITY = shipment_i[i, "QTY.Per.Box"], LOCATION = warehouse) %>% mutate(PO = paste0("PO#", POn[PO.REF.NO, "Document.Number"])))
      }
    }
  }
  Nbox <- Nbox + length(unique(shipment_o$BOX.NO))
  print(paste0("Total Qty ", sum(shipment_o$QUANTITY), "; Total Box# ", length(unique(shipment_o$BOX.NO))))
  shipment <- rbind(shipment, shipment_o)
  attachment <- rbind(attachment, data.frame(Box.Number = c(gsub(".*-", "", shipment_o$BOX.NO), ""), SKU = c(shipment_o$ITEM, ""), Packing.Slip.QTY = c(shipment_o$QUANTITY, ""), Notes = ""))
}
if(sum(shipment$QUANTITY) == summary[nrow(summary), "QTY"]){print(paste0("Total QTY matches: ", sum(shipment$QUANTITY)))}else{print(paste("Total QTY NOT matching:", sum(shipment$QUANTITY), summary[nrow(summary), "QTY"]))}
if(Nbox == summary[nrow(summary), "Number.of.Box"]){print(paste0("Total Box# matches: ", Nbox))}else{print(paste("Total Box# NOT matching:", Nbox, summary[nrow(summary), "Number.of.Box"]))}
# check for over-receiving
shipment <- shipment %>% group_by(PO.REF.NO, ITEM) %>% mutate(QUANTITY = sum(QUANTITY), BOX.NO = paste0(BOX.NO, collapse = " | ")) %>% distinct(PO.REF.NO, ITEM, .keep_all = T) %>% ungroup()
OverReceive <- shipment %>% mutate(Qty.Remain = PO_detail[paste0(PO.REF.NO, "_", ITEM), "Quantity.Remain"], Qty.Remain = ifelse(is.na(Qty.Remain), 0, Qty.Remain), Qty = QUANTITY - Qty.Remain) %>% filter(QUANTITY > Qty.Remain)
if(nrow(OverReceive)){
  PO_OverReceive <- OverReceive %>% mutate(PO.TYPE = "Over Received", CATEGORY = "MIX", SEASON = season, REF.NO = paste0("Over.", RefNo), WAREHOUSE = warehouse, VENDOR = "China", CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = paste0("Over receiving in shipment ", RefNo), ITEM. = ITEM, QUANTITY. = Qty, Tax.Code = "CA-Zero", External.ID = REF.NO) %>% 
    select(PO.TYPE:External.ID) %>% rename_with(~ gsub("\\.", " ", .))
  write.csv(PO_OverReceive, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_PO_", RefNo, "_OverReceive.csv"), row.names = F, na = "")
}else{
  print("No over-receiving. ")
}
## upload over-receiving PO 
# output for NS: upload shipment to Inbound Shipment and attach attachment in Communication tab
if(nrow(OverReceive)){
  p <- "PO#PO000095" # over-receiving PO 
  shipment <- shipment %>% mutate(Qty.Remain = PO_detail[paste0(PO.REF.NO, "_", ITEM), "Quantity.Remain"], Qty.Remain = ifelse(is.na(Qty.Remain), 0, Qty.Remain), QUANTITY = ifelse(QUANTITY > Qty.Remain, Qty.Remain, QUANTITY)) %>% filter(PO != "PO#NA") %>% select(-Qty.Remain)
  shipment <- rbind(shipment, data.frame(REF.NO = RefNo, EXPECTED.SHIPPING.DATE = ShippingDate, EXPECTED.DELIVERY.DATE = ReceiveDate, MEMO = memo, PO.REF.NO = "Over.24FWCA14", BOX.NO = OverReceive$BOX.NO, ITEM = OverReceive$ITEM, QUANTITY = OverReceive$Qty, LOCATION = warehouse, PO = p)) %>% arrange(ITEM)
}
write.csv(shipment, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, ".csv"), row.names = F, na = "")
attachment <- attachment %>% rename_with(~ gsub("\\.", " ", .))
write.csv(attachment, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, "_attachment.csv"), row.names = F, na = "")


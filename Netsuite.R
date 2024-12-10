# SKUs analysis and management
library(dplyr)
library(openxlsx)
library(scales)
library(ggplot2)
library(xlsx)
library(data.table)
library(tidyr)
library(readr)
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
inventory_update <- clover_item %>% mutate(Date = format(Sys.Date(), "%m/%d/%Y"), ITEM = Name, ADJUST.QTY.BY = ifelse(Name %in% netsuite_item$Name, Quantity - netsuite_item[Name, "Warehouse.On.Hand"], Quantity), Reason.Code = "CC", MEMO = "Clover Inventory Update", WAREHOUSE = "WH-RICHMOND", External.ID = paste0("IAR-", format(Sys.Date(), "%y%m%d"), "-1")) %>%
  select(Date, ITEM, ADJUST.QTY.BY, Reason.Code, MEMO, WAREHOUSE, External.ID) %>% filter(ADJUST.QTY.BY != 0)
colnames(inventory_update) <- gsub("QTY BY", "QTY. BY", gsub("\\.", " ", colnames(inventory_update)))
write.csv(inventory_update, file = paste0("../NetSuite/IA-Richmond-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")

# ------------- upload Clover SO ---------------------------
customer <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "Customers-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>%
  mutate(Name = paste0(First.Name, " ", Last.Name)) %>% distinct(Email.Address, .keep_all = T) %>% filter(Name != " ", Email.Address != "") %>% `row.names<-`(.[, "Name"])
payments <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "Payments-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% mutate(Phone = customer[Customer.Name, "Phone.Number"], Email = customer[Customer.Name, "Email.Address"]) %>% `row.names<-`(.[, "Order.ID"]) 
clover_so <- read.csv(rownames(file.info(list.files(path = "../Clover/", pattern = "LineItemsExport-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% mutate(Recipient = payments[Order.ID, "Customer.Name"], Recipient.Phone = payments[Order.ID, "Phone"], Recipient.Email = payments[Order.ID, "Email"], Tender = payments[Order.ID, "Tender"])
netsuite_so <- clover_so %>% filter(Item.SKU != "", Order.Payment.State == "Paid", !grepl("XHS", Order.Discounts), !grepl("Surrey", Order.Discounts)) %>% mutate(Order.date = format(as.Date(gsub(" .*", "", Line.Item.Date), "%d-%b-%Y"), "%m/%d/%Y")) %>% group_by(Order.date) %>% 
  mutate(Payment.Option = ifelse(Tender == "Cash", "Cash", ifelse(grepl("WeChat", Tender), "AlphaPay", "Clover")), Class = "FBM : CA", MEMO = "Clover sales", Customer = "54 JJR SHOPR", ID = data.table::rleid(Order.ID), REF.ID = paste0("CL", format(as.Date(Order.date, "%m/%d/%Y"), "%y%m%d"), "-", sprintf("%02d", ID)), Order.Type = "JJR SHOPR", Department = "Retail : Store Richmond", Warehouse = "WH-RICHMOND", Quantity = 1, Price.level = "Custom", Rate = Item.Total, Coupon.Discount = Total.Discount + Order.Discount.Proportion, Coupon.Discount = ifelse(Coupon.Discount == 0, NA, Coupon.Discount), Coupon.Code = gsub("NA", "", gsub(" -.*","", paste0(Order.Discounts, Discounts))), Tax.Code = ifelse(Item.Tax.Rate == 0.05, "CA-BC-GST", ifelse(Item.Tax.Rate == 0.12, "CA-BC-TAX", "")), SHIPPING.CARRIER = "", SHIPPING.METHOD = "", SHIPPING.COST = 0) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, Price.level, Rate, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount, Customer, Recipient, Recipient.Phone, Recipient.Email, SHIPPING.CARRIER, SHIPPING.METHOD, SHIPPING.COST)
write.csv(netsuite_so, file = paste0("../Clover/SO-clover-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")

# ------------- upload Square SO ---------------------------
customer <- read.csv("../Square/customers.csv", as.is = T) %>% `row.names<-`(.[, "Square.Customer.ID"])
payments <- read.csv(rownames(file.info(list.files(path = "../Square/", pattern = "transactions-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% `row.names<-`(.[, "Payment.ID"])
square_so <- read.csv(rownames(file.info(list.files(path = "../Square/", pattern = "items-", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% 
  mutate(Cash = payments[Payment.ID, "Cash"], Recipient.Email = ifelse(is.na(Customer.ID), "", customer[Customer.ID, "Email.Address"]), Recipient.Phone = ifelse(is.na(Customer.ID), "", customer[Customer.ID, "Phone.Number"]))
netsuite_so <- square_so %>% filter(Item != "") %>% mutate(Order.date = format(as.Date(Date, "%Y-%m-%d"), "%m/%d/%Y")) %>% group_by(Order.date) %>% 
  mutate(Payment.Option = ifelse(Cash == "$0.00" | is.na(Cash), "Square", "Cash"), Class = "FBM : CA", MEMO = "Square sales", Customer = ifelse(Location == "Surrey", "55 JJR SHOPS", "54 JJR SHOPR"), Recipient = ifelse(is.na(Customer.Name), "", Customer.Name), Order.ID = Transaction.ID, Item.SKU = Item, ID = data.table::rleid(Order.ID), REF.ID = paste0("SQ", format(as.Date(Order.date, "%m/%d/%Y"), "%y%m%d"), "-", sprintf("%02d", ID)), Order.Type = ifelse(Location == "Surrey", "JJR SHOPS", "JJR SHOPR"), Department = ifelse(Location == "Surrey", "Retail : Store Surrey", "Retail : Store Richmond"), Warehouse = ifelse(Location == "Surrey", "WH-SURREY", "WH-RICHMOND"), Quantity = Qty, Price.level = "Custom", Coupon.Discount = as.numeric(gsub("\\$", "", Discounts)), Coupon.Discount = ifelse(Coupon.Discount == 0, "", Coupon.Discount), Coupon.Code = "", Tax = as.numeric(gsub("\\$", "", Tax)), Rate = round(as.numeric(gsub("\\$", "", Net.Sales))/Qty, 2), Tax.Code = ifelse(round(Tax/as.numeric(gsub("\\$", "", Net.Sales)), 2) == 0.05, "CA-BC-GST", ifelse(round(Tax/as.numeric(gsub("\\$", "", Net.Sales)), 2) == 0.12, "CA-BC-TAX", "")), Tax.Amount = Tax, SHIPPING.CARRIER = "", SHIPPING.METHOD = "", SHIPPING.COST = 0) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, Price.level, Rate, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount, Customer, Recipient, Recipient.Phone, Recipient.Email, SHIPPING.CARRIER, SHIPPING.METHOD, SHIPPING.COST)
write.csv(netsuite_so, file = paste0("../Square/SO-square-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")

# ------------- upload XHS SO ---------------------------
XHS_so <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../XHS/", pattern = "order_export", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = 1) 
netsuite_so <- XHS_so %>% filter(Financial.Status == "paid") %>% mutate(Order.date = format(as.Date(gsub(" .*", "", Created.At), "%Y-%m-%d"), "%m/%d/%Y")) %>% 
  mutate(Payment.Option = "AlphaPay", Class = "FBM : CN", MEMO = "XHS sales", Customer = "56 XHS CN", Order.ID = Order.No, REF.ID = paste0("XHSCN-", Order.No), Order.Type = "XHS CN", Department = "Retail : Marketplace : XiaoHongShu", Warehouse = "WH-RICHMOND", Item.SKU = LineItem.SKU, Quantity = as.numeric(LineItem.Quantity), Price.level = "Custom", Rate = round((as.numeric(LineItem.Total) - as.numeric(LineItem.Total.Discount))/Quantity, 2), Coupon.Discount = as.numeric(LineItem.Total.Discount), Coupon.Discount = ifelse(Coupon.Discount == 0, "", Coupon.Discount), Coupon.Code = "", Tax.Amount = "", Tax.Code = "", Recipient = paste0(Shipping.Last.Name, Shipping.First.Name), Recipient.Phone = Phone, Recipient.Email = Email, SHIPPING.CARRIER = "Longxing", SHIPPING.METHOD = "", SHIPPING.COST = Shipping) %>% 
  select(Payment.Option, Class, Order.date, REF.ID, Order.Type, Department, Warehouse, MEMO, Order.ID, Item.SKU, Quantity, Price.level, Rate, Coupon.Discount, Coupon.Code, Tax.Code, Tax.Amount, Customer, Recipient, Recipient.Phone, Recipient.Email, SHIPPING.CARRIER, SHIPPING.METHOD, SHIPPING.COST)
Encoding(netsuite_so$Recipient) = "UTF-8"
write_excel_csv(netsuite_so, file = paste0("../XHS/SO-XHS-", format(Sys.Date(), "%Y%m%d"), ".csv"), na = "")

# ------------- upload CN inventory ---------------------------
xoro_CN <- read.xlsx2(list.files(path = "../xoro/", pattern = paste0("CN_Item Inventory Snapshot_", format(Sys.Date(), "%m%d%Y"), ".xlsx"), full.names = T), sheetIndex = 1) %>% filter(ATS > 0) %>% mutate(Item. = toupper(Item.), ATS = as.numeric(ATS)) %>% `row.names<-`(toupper(.[, "Item."])) 
inventory_update <- xoro_CN %>% mutate(Date = format(Sys.Date(), "%m/%d/%Y"), ITEM = Item., ADJUST.QTY.BY = ATS, Reason.Code = "CC", MEMO = "China Inventory Upload", WAREHOUSE = "WH-CHINA", External.ID = paste0("IAC-", format(Sys.Date(), "%y%m%d"), "-1")) %>%
  select(Date, ITEM, ADJUST.QTY.BY, Reason.Code, MEMO, WAREHOUSE, External.ID) %>% filter(ADJUST.QTY.BY != 0)
colnames(inventory_update) <- gsub("QTY BY", "QTY. BY", gsub("\\.", " ", colnames(inventory_update)))
write.csv(inventory_update, file = paste0("../NetSuite/IA-China-", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")

# ------------- Clover items not in NS ---------------------------
mastersku <- openxlsx::read.xlsx(list.files(path = "../../TWK 2020 share/", pattern = "1-MasterSKU-All-Product-", full.names = T)[1], sheet = "MasterFile", startRow = 4, fillMergedCells = T) %>% `row.names<-`(toupper(.[, "MSKU"]))
netsuite_item <- read.csv(list.files(path = "../NetSuite/", pattern = paste0("Items_", format(Sys.Date(), "%Y%m%d"), ".csv"), full.names = T), as.is = T) %>% mutate(Name = toupper(Name))
clover_item <- openxlsx::read.xlsx(list.files(path = "../Clover/", pattern = paste0("inventory", format(Sys.Date(), "%Y%m%d"), ".xlsx"), full.names = T), sheet = "Items") 
items <- clover_item %>% filter(!(Name %in% netsuite_item$Name)) %>% mutate(Seasons = mastersku[Name, "Seasons.SKU"], Status = mastersku[Name, "MSKU.Status"]) %>% 
  select(Name, Alternate.Name, Status, Seasons, Quantity)
write.csv(items, file = paste0("../NetSuite/items_Clover_notNS", format(Sys.Date(), "%Y%m%d"), ".csv"), row.names = F, na = "")

# ------------ upload inbound shipment for POs -------------------
library(tidyr)
season <- "25S"; warehouse <- "FBA-US"; PO_suffix <- "-CA"
RefNo <- "25SSUS1"; ShippingDate <- "12/3/2024"; ReceiveDate <- "1/10/2025"
PO_detail <- read.csv(rownames(file.info(list.files(path = "../PO/", pattern = "PurchaseOrders", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% filter(Item != "") %>% 
  mutate(Quantity = as.numeric(Quantity), Quantity.Fulfilled.Received = as.numeric(Quantity.Fulfilled.Received), Quantity.on.Shipments = ifelse(is.na(as.numeric(Quantity.on.Shipments)), 0, as.numeric(Quantity.on.Shipments)), Quantity.Remain = Quantity - Quantity.on.Shipments - Quantity.Fulfilled.Received, Quantity.Remain = ifelse(Quantity.Remain < 0, 0, Quantity.Remain)) %>% `row.names<-`(paste0(.[, "REF.NO"], "_", .[, "Item"]))
POn <- PO_detail %>% filter(!duplicated(REF.NO)) %>% `row.names<-`(.[, "REF.NO"])
shipment_in <- list.files(path = "../PO/shipment/", pattern = paste0(".*", RefNo, ".*.xlsx"), recursive = T, full.names = T)
memo <- gsub(paste0(RefNo, " *"), "", gsub(" *ETA.*\\/.*", "", gsub("../PO/shipment/", "", shipment_in)))
sheets <- getSheetNames(shipment_in)[]
summary <- openxlsx::read.xlsx(shipment_in, sheet = 1, fillMergedCells = T)
shipment <- data.frame(); attachment <- data.frame(); Nbox <- 0
for(s in 2:length(sheets)){
  sheet <- sheets[s]
  shipment_i <- openxlsx::read.xlsx(shipment_in, sheet = sheet, startRow = 5, fillMergedCells = T) %>% 
    select(1:Season) %>% filter(grepl(".*\\-.*", SKU)) %>% tidyr::fill(1:Season, .direction = "down")
  if(sum(grepl("CTN.NO", colnames(shipment_i))) == 1){
    print(paste(1, sheet))
    shipment_o <- data.frame(REF.NO = RefNo, EXPECTED.SHIPPING.DATE = ShippingDate, EXPECTED.DELIVERY.DATE = ReceiveDate, MEMO = memo, PO.REF.NO = paste0(shipment_i$`PO.#`, PO_suffix), BOX.NO = paste0(s, "-", shipment_i$`CTN.NO`), ITEM = shipment_i$SKU, QUANTITY = shipment_i$QTY.SHIPPED, LOCATION = warehouse) %>% mutate(PO = paste0("PO#", POn[PO.REF.NO, "Document.Number"]))
  }else{
    print(paste(2, sheet))
    colnames(shipment_i) <- c("Start", "End", colnames(shipment_i)[3:length(colnames(shipment_i))])
    shipment_i <- shipment_i %>% mutate(Start = as.numeric(Start), End = as.numeric(End), QTY.Per.Box = as.numeric(QTY.Per.Box))
    shipment_o <- data.frame()
    for(i in 1:nrow(shipment_i)){
      for(n in shipment_i[i, "Start"]:shipment_i[i, "End"]){
        shipment_o <- rbind(shipment_o, data.frame(REF.NO = RefNo, EXPECTED.SHIPPING.DATE = ShippingDate, EXPECTED.DELIVERY.DATE = ReceiveDate, MEMO = memo, PO.REF.NO = paste0(shipment_i[i, "PO.#"], PO_suffix), BOX.NO = paste0(s, "-", n), ITEM = shipment_i[i, "SKU"], QUANTITY = shipment_i[i, "QTY.Per.Box"], LOCATION = warehouse) %>% mutate(PO = paste0("PO#", POn[PO.REF.NO, "Document.Number"])))
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
write.csv(shipment, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, ".csv"), row.names = F, quote = F, na = "")
attachment <- attachment %>% rename_with(~ gsub("\\.", " ", .))
write.csv(attachment, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, "_receiving_by_box.csv"), row.names = F, quote = F, na = "")
## generate receiving csv file
id <- "INBSHIP6"
receiving <- shipment %>% mutate(ID = id, PO = gsub("PO#PO", "PO-", PO), Item = ITEM, Qty = QUANTITY) %>% select(ID, PO, Item, Qty)
write.csv(receiving, file = paste0(gsub("(.*\\/).*", "\\1", shipment_in), "NS_", RefNo, "_receiving.csv"), row.names = F, quote = F, na = "")

# ------------ upload CEFA POs -------------------
ID <- "CEFA5"; season <- "24F"; ReceiveDate <- "11/20/2024"
PO <- openxlsx::read.xlsx(rownames(file.info(list.files(path = "../PO/order/CEFA/", pattern = "Order form.*.xlsx", full.names = TRUE)) %>% filter(mtime == max(mtime))), sheet = 1, startRow = 3) %>% 
  filter(grepl("-.*-", SKU.Number)) %>% mutate(Quantity = ifelse(is.na(Quantity), 0, Quantity), cat = gsub("-.*", "", SKU.Number))
CAT <- paste((PO %>% distinct(cat))$cat, collapse = " ")
PO_NS <- PO %>% mutate(PO.TYPE = "Daycare Uniforms", CATEGORY = CAT, SEASON = season, REF.NO = ID, WAREHOUSE = "WH-SURREY", VENDOR = "China", CURRENCY = "CAD", ORDER.DATE = format(Sys.Date(), "%m/%d/%Y"), ORDER.PLACED.BY = "Gloria Li", APPROVAL.STATUS = "APPROVED", DUE.DATE = ReceiveDate, MEMO = paste0("CEFA order for ", season), ITEM = SKU.Number,  QUANTITY = Quantity, Tax.Code = "CA-Zero", External.ID = REF.NO) %>% 
  filter(Quantity > 0) %>% select(PO.TYPE:External.ID) %>% rename_with(~ gsub("\\.", " ", .))
write.csv(PO_NS, file = paste0("../PO/order/CEFA/", "NS_PO_", ID, ".csv"), row.names = F, na = "")

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

# ------------ Adjust CN inventory for AU customer order -------------------
PO_detail <- read.csv(rownames(file.info(list.files(path = "../PO/", pattern = "PurchaseOrders", full.names = TRUE)) %>% filter(mtime == max(mtime))), as.is = T) %>% filter(Item != "") %>% 
  mutate(Quantity = as.numeric(Quantity), Quantity.Fulfilled.Received = as.numeric(Quantity.Fulfilled.Received), Quantity.on.Shipments = ifelse(is.na(as.numeric(Quantity.on.Shipments)), 0, as.numeric(Quantity.on.Shipments)), Quantity.Remain = Quantity - Quantity.on.Shipments - Quantity.Fulfilled.Received, Quantity.Remain = ifelse(Quantity.Remain < 0, 0, Quantity.Remain)) %>% `row.names<-`(paste0(.[, "REF.NO"], "_", .[, "Item"]))
POn <- PO_detail %>% filter(!duplicated(REF.NO)) %>% `row.names<-`(.[, "REF.NO"])
netsuite_item <- read.csv(list.files(path = "../NetSuite/", pattern = paste0("Items_All_", format(Sys.Date(), "%Y%m%d"), ".csv"), full.names = T), as.is = T)
netsuite_item[netsuite_item == "" | is.na(netsuite_item)] <- 0
netsuite_item_CN <- netsuite_item %>% filter(Inventory.Warehouse == "WH-CHINA") %>% `row.names<-`(toupper(.[, "Name"])) 
AU_SO <- openxlsx::loadWorkbook(list.files(path = "../NetSuite/", pattern = paste0("AU coustomer order-packing list.xlsx"), full.names = T))
AU_SO1 <- openxlsx::readWorkbook(AU_SO, "batch 1") %>% mutate(Qty_CN = netsuite_item_CN[SKU, "Warehouse.On.Hand"], Qty_missing = ifelse(QTY < Qty_CN, 0, QTY - Qty_CN))
AU_SO2 <- openxlsx::readWorkbook(AU_SO, "batch 2", startRow = 5, cols = c(1:7)) %>% filter(!is.na(SKU)) %>% `row.names<-`(toupper(.[, "SKU"])) 
AU_SO2_Inbound <- data.frame(REF.NO = "IN_TWMU241127-PARA", EXPECTED.SHIPPING.DATE = "12/04/2024", EXPECTED.DELIVERY.DATE = "12/05/2024", MEMO = "Will leave in WH-CHINA for TWMU241127-PARA", PO.REF.NO = paste0(AU_SO2$`PO.#`, "-CA"), BOX.NO = 1, ITEM = AU_SO2$SKU, QUANTITY = AU_SO2$QTY.SHIPPED, LOCATION = "WH-CHINA") %>% mutate(PO = paste0("PO#", POn[PO.REF.NO, "Document.Number"]))
write.csv(AU_SO2_Inbound, file = "../PO/shipment/IN_TWMU241127-PARA.csv", row.names = F)
AU_SO2_Inbound_receive <- data.frame(ID = "INBSHIP5", PO = gsub("PO#PO", "PO-", AU_SO2_Inbound$PO), Item = AU_SO2_Inbound$ITEM, Qty = AU_SO2_Inbound$QUANTITY)
write.csv(AU_SO2_Inbound_receive, file = "../PO/shipment/IN_TWMU241127-PARA_receiving.csv", row.names = F, quote = F)

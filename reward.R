# add POS sales to yotpo rewards program system
## --------------- Clover sales to yotpo -----------
# customer info: Clover > Customers > Download
# sales: Clover > Transactions > Payments > select dates > Export
library(dplyr)
customer <- read.csv(paste0("../Clover/", list.files(path = "../Clover/", pattern = paste0("Customers-", format(Sys.Date(), "%Y%m%d")))), as.is = T) %>%
  mutate(Name = paste0(First.Name, " ", Last.Name)) %>% 
  filter(Name != " ", Email.Address != "")
rownames(customer) <- customer$Name

payments <- read.csv(paste0("../Clover/", list.files(path = "../Clover/", pattern = paste0("Payments-", format(Sys.Date(), "%Y%m%d")))), as.is = T) 
payments$email <- customer[payments$Customer.Name, "Email.Address"]
payments[is.na(payments$Refund.Amount), "Refund.Amount"] <- 0
payments <- payments %>% filter(!is.na(email))

point <- data.frame(Email = payments$email, Points = as.integer(payments$Amount - payments$Refund.Amount))

## --------------- Square sales to yotpo --------------
# customer info: Square > Customers > Export customers
# sales: Square > Transactions > select dates > Export Transactions CSV
square_customer <- read.csv(paste0("../Square/", list.files(path = "../Square/", pattern = paste0("customers-", format(Sys.Date(), "%Y%m%d")))), as.is = T)
rownames(square_customer) <- square_customer$Square.Customer.ID
transactions <- read.csv(paste0("../Square/", list.files(path = "../Square/", pattern = paste0("transactions-.*", format(Sys.Date(), "%Y-%m-%d")))), as.is = T) %>%
  filter(Customer.ID != "")
point <- bind_rows(point, data.frame(Email = square_customer[transactions$Customer.ID, "Email.Address"], Points = as.integer(gsub("\\$", "", transactions$Total.Collected))))
write.csv(point, file = paste0("../yotpo/", format(Sys.Date(), "%m%d%Y"), "-yotpo.csv"), row.names = F)

# combine PO files and shipping info into PO tracking file
library(dplyr)
library(openxlsx)

# ------ write header for new Shipment Tracking.xlsx file ------
season <- "2023FW"
POtrack <- createWorkbook()
addWorksheet(POtrack, paste(season, "Shipment Tracking"))
mergeCells(POtrack, 1, cols = 1:10, rows = 1:3)
writeData(POtrack, 1, paste(season, "China to Global Shipment Tracking"), startRow = 1, startCol = 1)
addStyle(POtrack, 1, createStyle(fontSize = 22, halign = "center", valign = "center", textDecoration = "bold", border = "TopBottomLeftRight"), rows = 1, cols = 1)
mergeCells(POtrack, 1, cols = 11:12, rows = 1)
writeData(POtrack, 1, "Country/Region", startRow = 1, startCol = 11)
mergeCells(POtrack, 1, cols = 11:12, rows = 2)
writeData(POtrack, 1, "Container #", startRow = 2, startCol = 11)
mergeCells(POtrack, 1, cols = 11:12, rows = 3)
writeData(POtrack, 1, "Shipment Details", startRow = 3, startCol = 11)
addStyle(POtrack, 1, createStyle(halign = "center", valign = "center", textDecoration = "bold", border = "TopBottomLeftRight"), rows = 1:3, cols = 11)
mergeCells(POtrack, 1, cols = 13:15, rows = 1)
writeData(POtrack, 1, "CA/US", startRow = 1, startCol = 13)
addStyle(POtrack, 1, createStyle(fontSize = 16, halign = "center", valign = "center", textDecoration = "bold", fgFill = "#47b38f", border = "TopBottomLeftRight"), rows = 1, cols = 13)
mergeCells(POtrack, 1, cols = 16:18, rows = 1)
writeData(POtrack, 1, "UK", startRow = 1, startCol = 16)
addStyle(POtrack, 1, createStyle(fontSize = 16, halign = "center", valign = "center", textDecoration = "bold", fgFill = "#2db0e3", border = "TopBottomLeftRight"), rows = 1, cols = 16)
mergeCells(POtrack, 1, cols = 19:21, rows = 1)
writeData(POtrack, 1, "DE", startRow = 1, startCol = 19)
addStyle(POtrack, 1, createStyle(fontSize = 16, halign = "center", valign = "center", textDecoration = "bold", fgFill = "#db8625", border = "TopBottomLeftRight"), rows = 1, cols = 19)
mergeCells(POtrack, 1, cols = 22:24, rows = 1)
writeData(POtrack, 1, "Wholesaler", startRow = 1, startCol = 22)
addStyle(POtrack, 1, createStyle(fontSize = 16, halign = "center", valign = "center", textDecoration = "bold", fgFill = "#7ee86d", border = "TopBottomLeftRight"), rows = 1, cols = 22)
mergeCells(POtrack, 1, cols = 25:27, rows = 1)
writeData(POtrack, 1, "CN Initial Reserve", startRow = 1, startCol = 25)
addStyle(POtrack, 1, createStyle(fontSize = 16, halign = "center", valign = "center", textDecoration = "bold", fgFill = "#f7261b", border = "TopBottomLeftRight"), rows = 1, cols = 25)
cnames <- c("PO#", "CAT", "Season",	"Factory", "Total PO in Pcs", "Name", "SKU", "Seasons", "Colour/Graphic(EN)", "Colour/Graphic(CN)", "Remain. in Pcs", "Remain. in %", "PO Plan", "Remain. in Pcs", "Remain. in %", "PO Plan", "Remain. in Pcs", "Remain. in %", "PO Plan", "Remain. in Pcs", "Remain. in %", "PO Plan", "Remain. in Pcs", "Remain. in %", "PO Plan", "Remain. in Pcs", "Remain. in %")
writeData(POtrack, 1, t(as.matrix(cnames)), startRow = 4, startCol = 1, colNames = F)
addStyle(POtrack, 1, createStyle(halign = "center", valign = "center", textDecoration = "bold", fgFill = "#fcbc3d", wrapText = T, border = "TopBottomLeftRight"), rows = 4, cols = 1:27)
freezePane(POtrack, 1, firstActiveRow = 5, firstActiveCol = 13)
saveWorkbook(POtrack, paste0("../PO/", season, " China to Global Shipment Tracking.xlsx"), overwrite = TRUE)

# ----------- input PO -------------
masterSKU <- read.xlsx("../PO/1-MasterSKU-All-Product-2023-11-07.xlsx", sheet = 1, startRow = 4)
season <- "2023FW"
POn <- paste0("P", c(203))
POtrack <- loadWorkbook(paste0("../PO/", season, " China to Global Shipment Tracking.xlsx"))
PO <- read.xlsx(paste0("../PO/order/", list.files(path = "../PO/order/", pattern = paste0(POn, ".*.xlsx"), recursive = T)), sheet = 1)



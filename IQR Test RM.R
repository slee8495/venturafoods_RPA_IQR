library(tidyverse)
library(readxl)
library(writexl)
library(reshape2)
library(officer)
library(openxlsx)
library(lubridate)
library(magrittr)
library(visdat)
library(simputation)
library(skimr)


##################################################################################################################################################################
##################################################################################################################################################################
##################################################################################################################################################################

# Planner Address Book (If updated, correct this link) ----
Planner_adress <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/Address Book - 04.26.22.xlsx", 
                             sheet = "Sheet1", col_types = c("text", 
                                                             "text", "text", "text", "text"))

names(Planner_adress) <- str_replace_all(names(Planner_adress), c(" " = "_"))

colnames(Planner_adress)[1] <- "Planner"


# Exception Report ----

exception_report <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_2/exception report.xlsx", 
                               col_types = c("text", "text", "text", 
                                             "text", "numeric", "text", "text", "text", 
                                             "text", "text", "text", "text", "text", 
                                             "text", "numeric", "numeric", "numeric", 
                                             "numeric", "numeric", "numeric", 
                                             "numeric", "text", "text", "text", 
                                             "text", "text", "text", "text", "numeric", 
                                             "text", "text"))

exception_report[-1:-2,] -> exception_report

colnames(exception_report) <- exception_report[1, ]
exception_report[-1, ] -> exception_report

colnames(exception_report)[1] <- "B_P"
colnames(exception_report)[2] <- "ItemNo"
colnames(exception_report)[3] <- "Buyer"
colnames(exception_report)[4] <- "Planner"
colnames(exception_report)[5] <- "Supplier_No"
colnames(exception_report)[6] <- "Payee Number"
colnames(exception_report)[7] <- "MPF or Line"
colnames(exception_report)[8] <- "Order Policy Code"
colnames(exception_report)[9] <- "Order Policy Value"
colnames(exception_report)[10] <- "Plan Code"
colnames(exception_report)[11] <- "Fence Rule"
colnames(exception_report)[12] <- "Plan Fence Days"
colnames(exception_report)[13] <- "Msg Display Fence"
colnames(exception_report)[14] <- "Freeze Fence"
colnames(exception_report)[15] <- "Leadtime Days"
colnames(exception_report)[16] <- "Reorder MIN"
colnames(exception_report)[17] <- "Reorder MAX"
colnames(exception_report)[18] <- "Reorder Multiple"
colnames(exception_report)[19] <- "Safety Stock"
colnames(exception_report)[20] <- "Reorder Point"
colnames(exception_report)[21] <- "Reorder Qty"
colnames(exception_report)[22] <- "Avg Demand Weeks"
colnames(exception_report)[23] <- "Formula Type"
colnames(exception_report)[24] <- "Sort Code"
colnames(exception_report)[25] <- "Schedule Group"
colnames(exception_report)[26] <- "Model"
colnames(exception_report)[27] <- "Description"
colnames(exception_report)[28] <- "UOM"
colnames(exception_report)[29] <- "PL QTY"
colnames(exception_report)[30] <- "Planning Formula"
colnames(exception_report)[31] <- "Costing Formula"

names(exception_report) <- str_replace_all(names(exception_report), c(" " = "_"))


exception_report %<>% 
  dplyr::mutate(ref = paste0(B_P, "_", ItemNo)) %>% 
  dplyr::relocate(ref) 



# Campus_ref pulling ----

Campus_ref <- read_excel("S:/Supply Chain Projects/RStudio/BoM/Master formats/RM_on_Hand/Campus_ref.xlsx", 
                         col_types = c("numeric", "text", "text", 
                                       "numeric"))



colnames(Campus_ref)[1] <- "B_P"
colnames(Campus_ref)[2] <- "Description"
colnames(Campus_ref)[3] <- "Campus_Name"
colnames(Campus_ref)[4] <- "Campus"

Campus_ref %<>% 
  dplyr::mutate(Location = B_P) 

# Vlookup for Campus_ref

merge(exception_report, Campus_ref[, c("B_P", "Campus")], by = "B_P", all.x = TRUE) %>% 
  dplyr::mutate(campus_ref = paste0(Campus, "_", ItemNo)) %>% 
  dplyr::relocate(ref, campus_ref, Campus) %>% 
  dplyr::rename(Loc_SKU = campus_ref,
                campus = Campus) -> exception_report_original


# get the RM Item only. 

exception_report_original %<>% 
  dplyr::mutate(ItemNo = as.numeric(ItemNo)) %>% 
  dplyr::mutate(item_na = is.na(ItemNo)) %>% 
  dplyr::filter(item_na == FALSE) %>% 
  dplyr::mutate(campus_na = is.na(campus)) %>% 
  dplyr::filter(campus_na == FALSE) -> exception_report

exception_report$ItemNo <- as.character(exception_report$ItemNo)



# remove duplicated value - prioritize bigger Loc Number (RM only)  - exception_report_1 (supplier #)

exception_report %>% 
  dplyr::mutate(B_P = as.integer(B_P)) %>% 
  dplyr::arrange(Loc_SKU, desc(B_P)) -> exception_report_1

exception_report_1[-which(duplicated(exception_report_1$Loc_SKU)),] -> exception_report_1


# remove duplicated value - prioritize larger number of Lead time - exception_report_2 (Lead time)

exception_report$Leadtime_Days[is.na(exception_report$Leadtime_Days)] <- 0

exception_report %>% 
  dplyr::arrange(Loc_SKU, desc(Leadtime_Days)) -> exception_report_2

exception_report_2[-which(duplicated(exception_report_2$Loc_SKU)),] -> exception_report_2

# remove duplicated value - prioritize larger number of MOQ - exception_report_3 (MOQ)

exception_report$Reorder_MIN [is.na(exception_report$Reorder_MIN )] <- 0

exception_report %>% 
  dplyr::arrange(Loc_SKU, desc(Reorder_MIN)) -> exception_report_3

exception_report_3[-which(duplicated(exception_report_3$Loc_SKU)),] -> exception_report_3


# remove duplicated value - prioritize smaller number of B_P - exception_report_4 (Planner, UoM)

exception_report %>% 
  dplyr::mutate(B_P = as.integer(B_P)) %>% 
  dplyr::arrange(Loc_SKU, B_P) -> exception_report_4

exception_report_4[-which(duplicated(exception_report_4$Loc_SKU)),] -> exception_report_4

# Pivoting exception_report
exception_report_original$Safety_Stock[is.na(exception_report_original$Safety_Stock)] <- 0

reshape2::dcast(exception_report_original, Loc_SKU ~ ., value.var = "Safety_Stock", sum) %>% 
  dplyr::rename(Safety_Stock = ".") -> exception_report_pivot


# Read IQR Report ----

RM_data <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_2/Raw Material Inventory Health (IQR) - 05.18.22.xlsx", 
                      sheet = "RM data", col_names = FALSE, 
                      col_types = c("text", "text", "text", 
                                    "text", "text", "text", "text", "text", 
                                    "text", "numeric", "date", "text", "numeric", 
                                    "text", "text", "numeric", "numeric", "numeric", 
                                    "numeric", "numeric", "numeric", "numeric", "numeric", 
                                    "numeric", "numeric", "numeric", "numeric", "numeric", 
                                    "numeric", "numeric", "numeric", "numeric", "numeric", 
                                    "numeric", "numeric", "text", "numeric", "numeric", 
                                    "numeric", "numeric", "numeric", "numeric", "numeric", 
                                    "numeric", "numeric", "numeric", "numeric", "numeric", 
                                    "numeric", "numeric", "text", "text"))

RM_data[-1:-3,] -> RM_data
colnames(RM_data) <- RM_data[1, ]
RM_data[-1, ] -> RM_data


colnames(RM_data)[1] <- "Mfg Loc"
colnames(RM_data)[2] <- "Loc Name"
colnames(RM_data)[3] <- "Item"
colnames(RM_data)[4] <- "Loc SKU"
colnames(RM_data)[5] <- "Supplier No"
colnames(RM_data)[6] <- "Description"
colnames(RM_data)[7] <- "Used in Priority SKU?"
colnames(RM_data)[8] <- "Type"
colnames(RM_data)[9] <- "Item Type"
colnames(RM_data)[10] <- "Shelf Life day"
colnames(RM_data)[11] <- "Birthday"
colnames(RM_data)[12] <- "UoM"
colnames(RM_data)[13] <- "Lead time"
colnames(RM_data)[14] <- "Planner"
colnames(RM_data)[15] <- "Planner Name"
colnames(RM_data)[16] <- "Standard Cost"
colnames(RM_data)[17] <- "MOQ"
colnames(RM_data)[18] <- "EOQ"
colnames(RM_data)[19] <- "Safety Stock"
colnames(RM_data)[20] <- "Max Cycle Stock"
colnames(RM_data)[21] <- "Usable"
colnames(RM_data)[22] <- "Quality hold"
colnames(RM_data)[23] <- "Quality hold in cost"
colnames(RM_data)[24] <- "Soft Hold"
colnames(RM_data)[25] <- "On_Hand_usable_and_soft_hold"
colnames(RM_data)[26] <- "On Hand in cost"
colnames(RM_data)[27] <- "Target Inv"
colnames(RM_data)[28] <- "Target Inv in cost"
colnames(RM_data)[29] <- "Max inv"
colnames(RM_data)[30] <- "Max inv cost"
colnames(RM_data)[31] <- "OPV"
colnames(RM_data)[32] <- "PO in next 28 days"
colnames(RM_data)[33] <- "Receipt in the next 28 days"
colnames(RM_data)[34] <- "DOS"
colnames(RM_data)[35] <- "At Risk in cost"
colnames(RM_data)[36] <- "Inv Health"
colnames(RM_data)[37] <- "Current month dep demand"
colnames(RM_data)[38] <- "Next month dep demand"
colnames(RM_data)[39] <- "Total dep. demand Next 6 Months"
colnames(RM_data)[40] <- "Total Last 6 mos Sales"
colnames(RM_data)[41] <- "Total Last 12 mos Sales "
colnames(RM_data)[42] <- "has Max?"
colnames(RM_data)[43] <- "on_hand_Inv_greaterthan_max"
colnames(RM_data)[44] <- "on hand Inv<=max"
colnames(RM_data)[45] <- "on hand Inv>target"
colnames(RM_data)[46] <- "on hand Inv<=target"
colnames(RM_data)[47] <- "IQR_cost"
colnames(RM_data)[48] <- "UPI_cost"
colnames(RM_data)[49] <- "IQR_cost_and_Hold_cost"
colnames(RM_data)[50] <- "UPI_cost_and_Hold_cost"


names(RM_data) <- stringr::str_replace_all(names(RM_data), c(" " = "_"))

RM_data %<>% 
  dplyr::mutate(Loc_SKU = gsub("-", "_", Loc_SKU)) %>%
  dplyr::relocate(Loc_SKU, .before = Supplier_No) %>% 
  dplyr::mutate(ref = Loc_SKU)


# Inventory Analysis Read RM ----

Inventory_analysis <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_2/Inventory Analysis 05.25.22.xlsx", 
                                    col_types = c("text", "text", "text", 
                                                  "text", "text", "text", "text", "numeric", 
                                                  "numeric", "numeric"),
                                    sheet = "RM")



Inventory_analysis[-1,] -> Inventory_analysis
colnames(Inventory_analysis) <- Inventory_analysis[1, ]
Inventory_analysis[-1, ] -> Inventory_analysis


colnames(Inventory_analysis)[1] <- "Location"
colnames(Inventory_analysis)[2] <- "Location_Nm"
colnames(Inventory_analysis)[3] <- "SKU"
colnames(Inventory_analysis)[4] <- "Label"
colnames(Inventory_analysis)[5] <- "Description"
colnames(Inventory_analysis)[6] <- "Inventory_Status"
colnames(Inventory_analysis)[7] <- "Inventory_Hold_Status"
colnames(Inventory_analysis)[8] <- "Last_Purchase_Price"
colnames(Inventory_analysis)[9] <- "Total_Cost"
colnames(Inventory_analysis)[10] <- "Inventory_Qty_Cases"


merge(Inventory_analysis, Campus_ref[, c("Location", "Campus")], by = "Location", all.x = TRUE) -> Inventory_analysis

Inventory_analysis %<>%  
  dplyr::mutate(campus_ref = paste0(Campus, "_", SKU), campus_ref = gsub("-", "", campus_ref)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", SKU), ref = gsub("-", "", ref)) %>% 
  dplyr::relocate(ref, campus_ref, Campus) %>% 
  dplyr::rename(campus = Campus)


# Inventory_analysis_pivot_ref

reshape2::dcast(Inventory_analysis, ref ~ Inventory_Hold_Status, value.var = "Inventory_Qty_Cases", sum) -> pivot_ref_Inventory_analysis
reshape2::dcast(Inventory_analysis, campus_ref ~ Inventory_Hold_Status, value.var = "Inventory_Qty_Cases", sum) -> pivot_campus_ref_Inventory_analysis

pivot_campus_ref_Inventory_analysis %<>% 
  dplyr::rename(Usable = Useable, Loc_SKU = campus_ref, Hard_Hold = "Hard Hold", Soft_Hold = "Soft Hold")

# BoM_dep_demand ----
BoM_dep_demand <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_2/BOM_Detail Constrained RM - 05.25.22.xlsx",
                             sheet = "dep demand")

colnames(BoM_dep_demand)[1] <- "Loc_SKU"
colnames(BoM_dep_demand)[2] <- "current_month"
colnames(BoM_dep_demand)[3] <- "next_month"
colnames(BoM_dep_demand)[ncol(BoM_dep_demand)] <- "sum_of_months"


BoM_dep_demand %<>% 
  dplyr::mutate(Loc_SKU = gsub("-", "_", Loc_SKU))



# Standard Cost # From MicroStrategy ----
Standard_Cost <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Standard Cost.xlsx", 
                            col_types = c("text", "text", "text", 
                                          "text", "numeric"))


Standard_Cost[-1,] -> Standard_Cost
colnames(Standard_Cost) <- Standard_Cost[1, ]
Standard_Cost[-1, ] -> Standard_Cost

colnames(Standard_Cost)[1] <- "Item"
colnames(Standard_Cost)[4] <- "Location_Nm"
colnames(Standard_Cost)[5] <- "Standard_Cost"

Standard_Cost %<>% 
  dplyr::mutate(Loc_SKU = paste0(Location, "_", Item))

# Consumption data component # Updated once a month ----
consumption_data <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/consumption data component - 05.04.22.xlsx")

consumption_data[-1:-2,] -> consumption_data
colnames(consumption_data) <- consumption_data[1, ]
consumption_data[-1, ] -> consumption_data

colnames(consumption_data)[1] <- "Loc_SKU"
colnames(consumption_data)[ncol(consumption_data)-1] <- "sum_12mos"
colnames(consumption_data)[ncol(consumption_data)] <- "sum_6mos"

consumption_data %<>% 
  dplyr::mutate(Loc_SKU = gsub("-", "_", Loc_SKU))


# SS Optimization RM for EOQ ----
SS_optimization <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/SS Optimization for EOQ - RM.xlsx",
                              sheet = "Sheet1")

SS_optimization[-1:-5,] -> SS_optimization
colnames(SS_optimization) <- SS_optimization[1, ]
SS_optimization[-1, ] -> SS_optimization

colnames(SS_optimization)[3] <- "Loc_SKU"
colnames(SS_optimization)[29] <- "Standard_Cost"
colnames(SS_optimization)[48] <- "EOQ_adjusted"


data.frame(SS_optimization$Loc_SKU) -> ss_opt_Loc_SKU

ss_opt_Loc_SKU %>% 
  dplyr::mutate(SS_optimization.Loc_SKU = gsub("-", "_", SS_optimization.Loc_SKU)) %>% 
  dplyr::rename(Loc_SKU = SS_optimization.Loc_SKU) %>% 
  dplyr::bind_cols(SS_optimization) %>% 
  dplyr::select(-"Loc_SKU...4") %>% 
  dplyr::rename(Loc_SKU = "Loc_SKU...1") -> SS_optimization

SS_optimization[-which(duplicated(SS_optimization$Loc_SKU)),] -> SS_optimization


# Custord PO ----
po <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_2/wo receipt custord po - 05.25.22.xlsx", 
                 sheet = "po", col_names = FALSE)

po %>% 
  dplyr::rename(aa = "...1") %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8"), sep = "~") %>% 
  dplyr::select(-"3") -> po


substr(po$`1`, nchar(po$`1`)-2, nchar(po$`1`)) -> label_po


po %>% 
  dplyr::bind_cols(label_po) %>% 
  dplyr::rename(label = "...8",
                aa = "1") %>% 
  dplyr::mutate(label_na = as.integer(label),
                label_na = !is.na(label_na)) %>% 
  dplyr::filter(label_na == TRUE) %>% 
  tidyr::separate(aa, c("global", "rp", "Item")) %>% 
  dplyr::select(-global, -rp, -"4", -"8", -label, -label_na) %>% 
  dplyr::rename(Location = "2",
                Qty = "5",
                po_no = "6",
                date = "7") %>% 
  dplyr::mutate(Qty = as.double(Qty),
                date = as.Date(date)) %>% 
  dplyr::mutate(Location = sub("^0+", "", Location),
                Item = sub("^0+", "", Item)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item)) %>% 
  dplyr::relocate(ref, Item, Location) -> po


# PO_Pivot 
po %>% 
  dplyr::mutate(date = as.Date(date)) %>% 
  dplyr::mutate(next_28_days = ifelse(date <= Sys.Date() + 28, "Y", "N")) %>% 
  reshape2::dcast(ref ~ next_28_days, value.var = "Qty", sum) %>% 
  dplyr::rename(Loc_SKU = ref) -> PO_Pivot



# Custord Receipt ----
receipt <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_2/wo receipt custord po - 05.25.22.xlsx", 
                      sheet = "receipt", col_names = FALSE)


receipt %>% 
  dplyr::rename(aa = "...1") %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8"), sep = "~") %>% 
  dplyr::select(-"3", -"8") -> receipt

substr(receipt$`1`, nchar(receipt$`1`)-2, nchar(receipt$`1`)) -> label_receipt


receipt %>% 
  dplyr::bind_cols(label_receipt) %>% 
  dplyr::rename(label = "...7",
                aa = "1") %>% 
  dplyr::mutate(label_na = as.integer(label),
                label_na = !is.na(label_na)) %>% 
  dplyr::filter(label_na == TRUE) %>% 
  tidyr::separate(aa, c("global", "rp", "Item")) %>% 
  dplyr::select(-global, -rp, -"4", -label, -label_na) %>% 
  dplyr::rename(Location = "2",
                Qty = "5",
                po_no = "6",
                date = "7") %>% 
  dplyr::mutate(Qty = as.double(Qty),
                date = as.Date(date)) %>% 
  dplyr::mutate(Location = sub("^0+", "", Location),
                Item = sub("^0+", "", Item)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item)) %>% 
  dplyr::relocate(ref, Item, Location) -> receipt



# Receipt Pivot
receipt %<>% 
  dplyr::mutate(date = as.Date(date)) %>% 
  dplyr::mutate(next_28_days = ifelse(date <= Sys.Date() + 28, "Y", "N")) %>% 
  reshape2::dcast(ref ~ next_28_days, value.var = "Qty", sum) %>% 
  dplyr::rename(Loc_SKU = ref) -> Receipt_Pivot  



################################################################################################################
################################################## Canada Read #################################################
################################################################################################################

JD_OH <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Canada/JD_OH_SS_20220525 - for R.xlsx", 
                                      col_types = c("numeric", "text", "text", 
                                                    "text", "numeric", "numeric", "text", 
                                                    "numeric", "numeric", "text", "text", 
                                                    "text"))

exception_report$Planner
str(RM_data)
str(exception_report)
JD_OH
colnames(JD_OH) <- JD_OH[1, ]
JD_OH[-1, ] -> JD_OH

colnames(JD_OH)[1]  <- "B_P"
colnames(JD_OH)[2]  <- "ItemNo"
colnames(JD_OH)[3]  <- "Stock_Type"
colnames(JD_OH)[4]  <- "Description"
colnames(JD_OH)[5]  <- "Balance_Usable"
colnames(JD_OH)[6]  <- "Balance_Hold"
colnames(JD_OH)[7]  <- "Lot_Status"
colnames(JD_OH)[8]  <- "On_Hand"
colnames(JD_OH)[9]  <- "Safety_Stock"
colnames(JD_OH)[10] <- "GL_Class"
colnames(JD_OH)[11] <- "Planner"
colnames(JD_OH)[12] <- "Planner_Name"

JD_OH %>% 
  dplyr::filter(B_P == 624) %>% 
  dplyr::mutate(ItemNo = as.integer(ItemNo)) %>% 
  dplyr::filter(ItemNo != is.na(ItemNo)) -> JD_OH


# add hold status to JD_OH

Lot_Status_Code <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/Lot Status Code.xlsx")

Lot_Status_Code[-1, ] -> Lot_Status_Code

colnames(Lot_Status_Code)[1] <- "Lot_Status"
colnames(Lot_Status_Code)[2] <- "Description"
colnames(Lot_Status_Code)[3] <- "Hard_Soft_Hold"

merge(JD_OH, Lot_Status_Code[, c("Lot_Status", "Hard_Soft_Hold")], by = "Lot_Status", all.x = TRUE) %>% 
  dplyr::rename(hold_status = Hard_Soft_Hold) -> JD_OH


reshape2::dcast(JD_OH, B_P + ItemNo + Description + Balance_Usable ~ . , value.var = "Balance_Hold", sum) %>% 
  dplyr::select(-".") -> JD_OH_Pivot_1

# Useable
merge(JD_OH_Pivot_1, Campus_ref[, c("B_P", "Campus")], by = "B_P", all.x = TRUE) %>% 
  dplyr::mutate(campus_ref = paste0(Campus, "_", ItemNo)) %>% 
  dplyr::mutate(ref = paste0(B_P, "_", ItemNo)) %>% 
  dplyr::rename(campus = Campus, Location = B_P, SKU = ItemNo, Inventory_Qty_Cases = Balance_Usable) %>% 
  dplyr::mutate(Location_Nm = "Edmonton", Label = "n/a", Inventory_Status = "n/a", Inventory_Hold_Status = "Useable",
                Last_Purchase_Price = "n/a", Total_Cost = "n/a") %>% 
  dplyr::relocate(ref, campus_ref, campus, Location, Location_Nm, SKU, Label, Description, Inventory_Status,
                  Inventory_Hold_Status, Last_Purchase_Price, Total_Cost, Inventory_Qty_Cases) %>% 
  dplyr::mutate(Location = as.character(Location),
                SKU = as.character(SKU),
                Total_Cost = as.numeric(Total_Cost)) -> JD_OH_Useable

# Hard Hold
JD_OH %>% 
  dplyr::filter(hold_status == "Hard") %>% 
  reshape2::dcast(B_P + ItemNo + Description ~ . , value.var = "Balance_Hold", sum) %>% 
  dplyr::rename(Location = B_P, 
                SKU = ItemNo,
                Inventory_Qty_Cases = ".") %>% 
  merge(Campus_ref[, c("Location", "Campus")], by = "Location", all.x = TRUE) %>% 
  dplyr::rename(campus = Campus) %>% 
  dplyr::mutate(ref = paste0(Location, "_", SKU),
                campus_ref = paste0(campus, "-", SKU),
                Location_Nm = "Edmonton",
                Label = "n/a",
                Inventory_Status = "n/a", 
                Inventory_Hold_Status = "Hard Hold",
                Last_Purchase_Price = "n/a",
                Total_Cost = "n/a") %>% 
  dplyr::relocate(ref, campus_ref, campus, Location, Location_Nm, SKU, Label, Description, Inventory_Status,
                  Inventory_Hold_Status, Last_Purchase_Price, Total_Cost, Inventory_Qty_Cases) %>% 
  dplyr::mutate(Location = as.character(Location),
                SKU = as.character(SKU),
                Total_Cost = as.numeric(Total_Cost)) -> JD_OH_Hard_Hold



# Soft Hold
JD_OH %>% 
  dplyr::filter(hold_status == "Soft") %>% 
  reshape2::dcast(B_P + ItemNo + Description ~ . , value.var = "Balance_Hold", sum) %>% 
  dplyr::rename(Location = B_P, 
                SKU = ItemNo,
                Inventory_Qty_Cases = ".") %>% 
  merge(Campus_ref[, c("Location", "Campus")], by = "Location", all.x = TRUE) %>% 
  dplyr::rename(campus = Campus) %>% 
  dplyr::mutate(ref = paste0(Location, "_", SKU),
                campus_ref = paste0(campus, "-", SKU),
                Location_Nm = "Edmonton",
                Label = "n/a",
                Inventory_Status = "n/a", 
                Inventory_Hold_Status = "Soft Hold",
                Last_Purchase_Price = "n/a",
                Total_Cost = "n/a") %>% 
  dplyr::relocate(ref, campus_ref, campus, Location, Location_Nm, SKU, Label, Description, Inventory_Status,
                  Inventory_Hold_Status, Last_Purchase_Price, Total_Cost, Inventory_Qty_Cases) %>% 
  dplyr::mutate(Location = as.character(Location),
                SKU = as.character(SKU),
                Total_Cost = as.numeric(Total_Cost)) -> JD_OH_Soft_Hold


rbind(Inventory_analysis, JD_OH_Useable, JD_OH_Hard_Hold, JD_OH_Soft_Hold) -> Inventory_analysis

#####################################################################################################################
######################################################## ETL ########################################################
#####################################################################################################################

# vlookup - Supplier No 
merge(RM_data, exception_report_1[, c("Loc_SKU", "Supplier_No")], by = "Loc_SKU", all.x = TRUE) %>%
  dplyr::mutate(supplier_no_na = is.na(Supplier_No.y), Supplier_No = ifelse(supplier_no_na == FALSE, Supplier_No.y, "DNRR")) %>% 
  dplyr::select(-Supplier_No.x, -Supplier_No.y, -supplier_no_na) %>% 
  dplyr::relocate(Loc_SKU, Supplier_No, .after = "Item") -> RM_data

# vlookup - UoM
merge(RM_data, exception_report_4[, c("ref", "UOM")], by = "ref", all.x = TRUE) -> RM_data

rm_na_1 <- as.matrix(RM_data[, "UOM"])
rm_na_1[is.na(rm_na_1)] <- "DNRR"
RM_data[, "UOM"] <- rm_na_1

RM_data %<>% 
  dplyr::relocate(UOM, .after = UoM) %>% 
  dplyr::select(-UoM) 

# vlookup - Lead Time
merge(RM_data, exception_report_2[, c("Loc_SKU", "Leadtime_Days")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::relocate(Leadtime_Days, .after = Lead_time) %>% 
  dplyr::select(-Lead_time) -> RM_data

rm_na_2 <- as.matrix(RM_data[, "Leadtime_Days"])
rm_na_2[is.na(rm_na_2)] <- "DNRR"
RM_data[, "Leadtime_Days"] <- rm_na_2

RM_data %<>% 
  dplyr::rename(Lead_time = Leadtime_Days)


# vlookup - Planner
merge(RM_data, exception_report_4[, c("ref", "Planner")], by = "ref", all.x = TRUE) %>% 
  dplyr::relocate(Planner.y, .after = "Planner.x") %>% 
  dplyr::select(-Planner.x) %>% 
  dplyr::rename(Planner = Planner.y) -> RM_data

rm_na_3 <- as.matrix(RM_data[, "Planner"])
rm_na_3[is.na(rm_na_3)] <- "DNRR"
RM_data[, "Planner"] <- rm_na_3


# vlookup - Planner Name
merge(RM_data, Planner_adress[, c("Planner", "Alpha_Name")], by = "Planner", all.x = TRUE) %>% 
  dplyr::relocate(Alpha_Name, .after = Planner_Name) %>% 
  dplyr::select(-Planner_Name) %>% 
  dplyr::rename(Planner_Name = Alpha_Name) %>% 
  dplyr::relocate(Planner, .before = Planner_Name)-> RM_data


# vlookup - Standard Cost
merge(RM_data, SS_optimization[, c("Loc_SKU", "Standard_Cost")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::relocate(Standard_Cost.y, .after = Standard_Cost.x) %>% 
  dplyr::select(-Standard_Cost.x) %>% 
  dplyr::rename(Standard_Cost = Standard_Cost.y) %>%
  dplyr::relocate(Item, .before = Loc_SKU) -> RM_data

RM_data$Standard_Cost <- as.double(RM_data$Standard_Cost)

RM_data %<>% 
  dplyr::mutate(Standard_Cost = sprintf("%.2f", Standard_Cost)) %>% 
  dplyr::mutate(Standard_Cost = as.double(Standard_Cost))


# vlookup - MOQ
merge(RM_data, exception_report_3[, c("Loc_SKU", "Reorder_MIN")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::relocate(Reorder_MIN, .after = MOQ) %>% 
  dplyr::select(-MOQ) %>% 
  dplyr::rename(MOQ = Reorder_MIN) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# vlookup - Safety Stock
merge(RM_data, exception_report_pivot[, c("Loc_SKU", "Safety_Stock")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(Safety_Stock.y = sprintf("%.2f", Safety_Stock.y)) %>% 
  dplyr::mutate(Safety_Stock.y = gsub("NA", "0", Safety_Stock.y)) %>% 
  dplyr::mutate(Safety_Stock.y = as.integer(Safety_Stock.y)) %>% 
  dplyr::relocate(Safety_Stock.y, .after = Safety_Stock.x) %>% 
  dplyr::select(-Safety_Stock.x) %>% 
  dplyr::rename(Safety_Stock = Safety_Stock.y) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# vlookup - Usable
merge(RM_data, pivot_campus_ref_Inventory_analysis[, c("Loc_SKU", "Usable")], by = "Loc_SKU", all.x = TRUE) %>%
  dplyr::mutate(Usable.y = sprintf("%.2f", Usable.y)) %>% 
  dplyr::mutate(Usable.y = gsub("NA", "0", Usable.y)) %>% 
  dplyr::mutate(Usable.y = as.integer(Usable.y)) %>% 
  dplyr::relocate(Usable.y, .after = Usable.x) %>% 
  dplyr::select(-Usable.x) %>% 
  dplyr::rename(Usable = Usable.y) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# vlookup - Quantity hold
merge(RM_data, pivot_campus_ref_Inventory_analysis[, c("Loc_SKU", "Hard_Hold")], by = "Loc_SKU", all.x = TRUE) %>%
  dplyr::mutate(Hard_Hold = sprintf("%.2f", Hard_Hold)) %>% 
  dplyr::mutate(Hard_Hold = gsub("NA", "0", Hard_Hold)) %>% 
  dplyr::mutate(Hard_Hold = as.integer(Hard_Hold)) %>% 
  dplyr::relocate(Hard_Hold, .after = Quality_hold) %>% 
  dplyr::select(-Quality_hold) %>% 
  dplyr::rename(Quality_hold = Hard_Hold) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data

# Calculation - Quality Hold in $$
RM_data %<>% 
  dplyr::mutate(Quality_hold_in_cost = Quality_hold * Standard_Cost) %>% 
  dplyr::mutate(Quality_hold_in_cost = sprintf("%.2f", Quality_hold_in_cost)) %>% 
  dplyr::mutate(Quality_hold_in_cost = gsub("NA", "0", Quality_hold_in_cost)) %>% 
  dplyr::mutate(Quality_hold_in_cost = as.double(Quality_hold_in_cost))

RM_data$Quality_hold_in_cost <- as.double(RM_data$Quality_hold_in_cost)

RM_data %<>% 
  dplyr::mutate(Quality_hold_in_cost = sprintf("%.2f", Quality_hold_in_cost)) %>% 
  dplyr::mutate(Quality_hold_in_cost = as.double(Quality_hold_in_cost))


# vlookup - Soft Hold
merge(RM_data, pivot_campus_ref_Inventory_analysis[, c("Loc_SKU", "Soft_Hold")], by = "Loc_SKU", all.x = TRUE) %>%
  dplyr::mutate(Soft_Hold.y = sprintf("%.2f", Soft_Hold.y)) %>% 
  dplyr::mutate(Soft_Hold.y = gsub("NA", "0", Soft_Hold.y)) %>% 
  dplyr::mutate(Soft_Hold.y = as.integer(Soft_Hold.y)) %>% 
  dplyr::relocate(Soft_Hold.y, .after = Soft_Hold.x) %>% 
  dplyr::select(-Soft_Hold.x) %>% 
  dplyr::rename(Soft_Hold = Soft_Hold.y) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data

# Calculation - On Hand (usable + soft hold)
RM_data %>% 
  dplyr::mutate(On_Hand_usable_and_soft_hold = Usable + Soft_Hold) -> RM_data

# Calculation - On Hand in $$
RM_data %<>% 
  dplyr::mutate(On_Hand_in_cost = On_Hand_usable_and_soft_hold * Standard_Cost) %>% 
  dplyr::mutate(On_Hand_in_cost = sprintf("%.2f", On_Hand_in_cost)) %>% 
  dplyr::mutate(On_Hand_in_cost = gsub("NA", "0", On_Hand_in_cost)) %>% 
  dplyr::mutate(On_Hand_in_cost = as.double(On_Hand_in_cost))



# vlookup - OPV
merge(RM_data, exception_report_4[, c("Loc_SKU", "Order_Policy_Value")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(Order_Policy_Value = as.integer(Order_Policy_Value)) %>% 
  dplyr::mutate(Order_Policy_Value = sprintf("%.2f", Order_Policy_Value)) %>% 
  dplyr::mutate(Order_Policy_Value = gsub("NA", "0", Order_Policy_Value)) %>% 
  dplyr::mutate(Order_Policy_Value = as.integer(Order_Policy_Value)) %>% 
  dplyr::relocate(Order_Policy_Value, .after = OPV) %>% 
  dplyr::select(-OPV) %>% 
  dplyr::rename(OPV = Order_Policy_Value) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# vlookup - PO in next 28 days
merge(RM_data, PO_Pivot[, c("Loc_SKU", "Y")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(Y = sprintf("%.2f", Y)) %>% 
  dplyr::mutate(Y = gsub("NA", "0", Y)) %>% 
  dplyr::mutate(Y = as.integer(Y)) %>% 
  dplyr::relocate(Y, .after = PO_in_next_28_days) %>% 
  dplyr::select(-PO_in_next_28_days) %>% 
  dplyr::rename(PO_in_next_28_days = Y) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data



# vlookup - Receipt in next 28 days
merge(RM_data, Receipt_Pivot[, c("Loc_SKU", "Y")], by = "Loc_SKU", all.x = TRUE) %>%
  dplyr::mutate(Y = sprintf("%.2f", Y)) %>%
  dplyr::mutate(Y = gsub("NA", "0", Y)) %>%
  dplyr::mutate(Y = as.integer(Y)) %>% 
  dplyr::relocate(Y, .after = Receipt_in_the_next_28_days) %>% 
  dplyr::select(-Receipt_in_the_next_28_days) %>% 
  dplyr::rename(Receipt_in_the_next_28_days = Y) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# vlookup - Current month dep demand
merge(RM_data, BoM_dep_demand[, c("Loc_SKU", "current_month")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(current_month = sprintf("%.2f", current_month)) %>% 
  dplyr::mutate(current_month = gsub("NA", "0", current_month)) %>% 
  dplyr::mutate(current_month = as.double(current_month)) %>% 
  dplyr::relocate(current_month, .after = Current_month_dep_demand) %>% 
  dplyr::select(-Current_month_dep_demand) %>% 
  dplyr::rename(Current_month_dep_demand = current_month) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data



# vlookup - Next month dep demand
merge(RM_data, BoM_dep_demand[, c("Loc_SKU", "next_month")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(next_month = sprintf("%.2f", next_month)) %>% 
  dplyr::mutate(next_month = gsub("NA", "0", next_month)) %>% 
  dplyr::mutate(next_month = as.double(next_month)) %>% 
  dplyr::relocate(next_month, .after = Next_month_dep_demand) %>% 
  dplyr::select(-Next_month_dep_demand) %>% 
  dplyr::rename(Next_month_dep_demand = next_month) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# vlookup - Total dep. demand Next 6 Months
merge(RM_data, BoM_dep_demand[, c("Loc_SKU", "sum_of_months")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(sum_of_months = sprintf("%.2f", sum_of_months)) %>% 
  dplyr::mutate(sum_of_months = gsub("NA", "0", sum_of_months)) %>% 
  dplyr::mutate(sum_of_months = as.double(sum_of_months)) %>% 
  dplyr::relocate(sum_of_months, .after = Total_dep._demand_Next_6_Months) %>% 
  dplyr::select(-Total_dep._demand_Next_6_Months) %>% 
  dplyr::rename(Total_dep_demand_Next_6_Months = sum_of_months) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# Calculation - DOS
RM_data %>% 
  dplyr::mutate(DOS = On_Hand_usable_and_soft_hold / (pmax(Current_month_dep_demand, Next_month_dep_demand)/30)) %>% 
  dplyr::mutate(DOS = sprintf("%.2f", DOS)) %>% 
  dplyr::mutate(DOS = gsub("NA", "0", DOS)) %>%
  dplyr::mutate(DOS = gsub("NaN", "0", DOS)) %>% 
  dplyr::mutate(DOS = as.double(DOS)) -> RM_data


# vlookup - Total Last 6 mos Sales
merge(RM_data, consumption_data[, c("Loc_SKU", "sum_6mos")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(sum_6mos = as.double(sum_6mos)) %>% 
  dplyr::mutate(sum_6mos = sprintf("%.2f", sum_6mos)) %>% 
  dplyr::mutate(sum_6mos = gsub("NA", "0", sum_6mos)) %>% 
  dplyr::mutate(sum_6mos = as.double(sum_6mos)) %>% 
  dplyr::relocate(sum_6mos, .after = Total_Last_6_mos_Sales) %>% 
  dplyr::select(-Total_Last_6_mos_Sales) %>% 
  dplyr::rename(Total_Last_6_mos_Sales = sum_6mos) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data

# vlookup - Total Last 12 mos Sales
merge(RM_data, consumption_data[, c("Loc_SKU", "sum_12mos")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(sum_12mos = as.double(sum_12mos)) %>% 
  dplyr::mutate(sum_12mos = sprintf("%.2f", sum_12mos)) %>% 
  dplyr::mutate(sum_12mos = gsub("NA", "0", sum_12mos)) %>% 
  dplyr::mutate(sum_12mos = as.double(sum_12mos)) %>% 
  dplyr::relocate(sum_12mos, .after = Total_Last_12_mos_Sales_) %>% 
  dplyr::select(-Total_Last_12_mos_Sales_) %>% 
  dplyr::rename(Total_Last_12_mos_Sales = sum_12mos) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data

# vlookup - EOQ
merge(RM_data, SS_optimization[, c("Loc_SKU", "EOQ_adjusted")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(EOQ_adjusted = as.double(EOQ_adjusted)) %>% 
  dplyr::mutate(EOQ_adjusted = sprintf("%.2f", EOQ_adjusted)) %>% 
  dplyr::mutate(EOQ_adjusted = gsub("NA", "0", EOQ_adjusted)) %>% 
  dplyr::mutate(EOQ_adjusted = as.double(EOQ_adjusted)) %>% 
  dplyr::relocate(EOQ_adjusted, .after = EOQ) %>% 
  dplyr::select(-EOQ) %>% 
  dplyr::rename(EOQ = EOQ_adjusted) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# Calculation - Max Cycle Stock
RM_data %>% 
  dplyr::mutate(Max_Cycle_Stock =
                  pmax(EOQ, MOQ, OPV*(Next_month_dep_demand/20.83),Total_Last_12_mos_Sales/250)) %>% 
  dplyr::mutate(Max_Cycle_Stock = sprintf("%.2f", Max_Cycle_Stock)) %>% 
  dplyr::mutate(Max_Cycle_Stock = gsub("NA", "0", Max_Cycle_Stock)) %>% 
  dplyr::mutate(Max_Cycle_Stock = as.integer(Max_Cycle_Stock)) -> RM_data


# Calculation - Target Inv
RM_data %>% 
  dplyr::mutate(Target_Inv = Safety_Stock + Max_Cycle_Stock / 2) -> RM_data

# Calculation - Target Inv in $$
RM_data %<>% 
  dplyr::mutate(Target_Inv_in_cost = Target_Inv * Standard_Cost) %>% 
  dplyr::mutate(Target_Inv_in_cost = as.double(Target_Inv_in_cost)) %>% 
  dplyr::mutate(Target_Inv_in_cost = sprintf("%.2f", Target_Inv_in_cost)) %>% 
  dplyr::mutate(Target_Inv_in_cost = gsub("NA", "0", Target_Inv_in_cost)) %>% 
  dplyr::mutate(Target_Inv_in_cost = as.double(Target_Inv_in_cost)) 


# Calculation - Max inv
RM_data %>% 
  dplyr::mutate(Max_inv = Safety_Stock + Max_Cycle_Stock) -> RM_data

# Calculation - Max inv $$
RM_data %<>% 
  dplyr::mutate(Max_inv_cost = Max_inv * Standard_Cost) %>% 
  dplyr::mutate(Max_inv_cost = as.double(Max_inv_cost)) %>% 
  dplyr::mutate(Max_inv_cost = sprintf("%.2f", Max_inv_cost)) %>% 
  dplyr::mutate(Max_inv_cost = gsub("NA", "0", Max_inv_cost)) %>% 
  dplyr::mutate(Max_inv_cost = as.double(Max_inv_cost)) 


# Calculation - has Max?
RM_data %>% 
  dplyr::mutate("has_Max?" = ifelse(Max_inv > 0, 1, 0)) -> RM_data

# Calculation - on hand Inv > max
RM_data %<>% 
  dplyr::mutate(on_hand_Inv_greaterthan_max = ifelse(On_Hand_usable_and_soft_hold > Max_inv, 1, 0))

# Calculation - on hand Inv <= max
RM_data %>% 
  dplyr::mutate("on_hand_Inv<=max" = ifelse(On_Hand_usable_and_soft_hold <= Max_inv, 1, 0)) -> RM_data

# Calculation - on hand Inv > target
RM_data %>% 
  dplyr::mutate("on_hand_Inv>target" = ifelse(On_Hand_usable_and_soft_hold > Target_Inv, 1, 0)) -> RM_data

# Calculation - on hand Inv <= target
RM_data %>% 
  dplyr::mutate("on_hand_Inv<=target" = ifelse(On_Hand_usable_and_soft_hold <= Target_Inv, 1, 0)) -> RM_data


# Calculation - Inv Health
# add today's date col
RM_data %>% 
  dplyr::mutate(today = Sys.Date()-6,
                today = as.Date(today, format = "%Y-%m-%d"),
                Birthday = as.Date(Birthday, format = "%Y-%m-%d"),
                diff_days = today - Birthday,
                diff_days = as.numeric(diff_days),
                Inv_Health = ifelse(On_Hand_usable_and_soft_hold < Safety_Stock, "BELOW SS", 
                                    ifelse(Item_Type == "Non-Commodity" & DOS > 0.6*Shelf_Life_day, "AT RISK", 
                                           ifelse(On_Hand_usable_and_soft_hold > 0 & Lead_time == "DNRR", "DEAD", 
                                                  ifelse(On_Hand_usable_and_soft_hold > 0 & Current_month_dep_demand == 0 & 
                                                           Next_month_dep_demand == 0 & Total_dep_demand_Next_6_Months == 0 & 
                                                           diff_days > 90, "DEAD", 
                                                         ifelse(on_hand_Inv_greaterthan_max == 0 | diff_days < 91, "HEALTHY", "EXECESS")))))) %>% 
  dplyr::select(-today, -diff_days) -> RM_data



# Calculation - At Risk in $$
RM_data %<>% 
  dplyr::mutate(At_Risk_in_cost  = ifelse(Inv_Health=="At Risk",
                                         (On_Hand_usable_and_soft_hold -((pmax(Current_month_dep_demand,Next_month_dep_demand)/30) 
                                                                         *(Shelf_Life_day*0.6)))*Standard_Cost,0)) 

# Calculation - IQR $$
RM_data %<>% 
  dplyr::mutate(IQR_cost = ifelse(Inv_Health == "DEAD" | Inv_Health == "HEALTHY" | Inv_Health == "BELOW SS", On_Hand_in_cost, 
                                  ifelse(Inv_Health == "AT RISK", At_Risk_in_cost, On_Hand_in_cost - Max_inv_cost))) 


RM_data$IQR_cost <- as.double(RM_data$IQR_cost)

RM_data %<>% 
  dplyr::mutate(IQR_cost = sprintf("%.2f", IQR_cost)) %>% 
  dplyr::mutate(IQR_cost = as.double(IQR_cost))

# Calculation - UPI $$
RM_data %<>% 
  dplyr::mutate(UPI_cost = ifelse(Inv_Health == "AT RISK", At_Risk_in_cost,
                                 ifelse(Inv_Health == "EXCESS", On_Hand_in_cost - Max_inv_cost,
                                        ifelse(Inv_Health == "DEAD", On_Hand_in_cost, 0))))



RM_data$UPI_cost <- as.double(RM_data$UPI_cost)

RM_data %<>% 
  dplyr::mutate(UPI_cost = sprintf("%.2f", UPI_cost)) %>% 
  dplyr::mutate(UPI_cost = as.double(UPI_cost))

# Calculation - IQR $$ + Hold $$
RM_data %<>% 
  dplyr::mutate(IQR_cost_and_Hold_cost = IQR_cost + Quality_hold_in_cost) 
  
RM_data$IQR_cost_and_Hold_cost <- as.double(RM_data$IQR_cost_and_Hold_cost)
  
RM_data %<>% 
    dplyr::mutate(IQR_cost_and_Hold_cost = sprintf("%.2f", IQR_cost_and_Hold_cost)) %>% 
    dplyr::mutate(IQR_cost_and_Hold_cost = as.double(IQR_cost_and_Hold_cost))

# Calculation - UPI $$ + Hold $$
RM_data %<>% 
  dplyr::mutate(UPI_cost_and_Hold_cost = UPI_cost + Quality_hold_in_cost)

RM_data$UPI_cost_and_Hold_cost <- as.double(RM_data$UPI_cost_and_Hold_cost)

RM_data %<>% 
  dplyr::mutate(UPI_cost_and_Hold_cost = sprintf("%.2f", UPI_cost_and_Hold_cost)) %>% 
  dplyr::mutate(UPI_cost_and_Hold_cost = as.double(UPI_cost_and_Hold_cost))





#####################################################################################################################
########################################## Change Col names to original #############################################
#####################################################################################################################

RM_data %<>% 
  dplyr::mutate(Loc_SKU = gsub("_", "-", Loc_SKU)) %>% 
  dplyr::relocate(Mfg_Loc, Loc_Name)



colnames(RM_data)[1]<-"Mfg Loc"
colnames(RM_data)[2]<-"Loc Name"
colnames(RM_data)[3]<-"Item"
colnames(RM_data)[4]<-"Loc-SKU"
colnames(RM_data)[5]<-"Supplier#"
colnames(RM_data)[6]<-"Description"
colnames(RM_data)[7]<-"Used in Priority SKU?"
colnames(RM_data)[8]<-"Type"
colnames(RM_data)[9]<-"Item Type"
colnames(RM_data)[10]<-"Shelf Life (day)"
colnames(RM_data)[11]<-"Birthday"
colnames(RM_data)[12]<-"UoM"
colnames(RM_data)[13]<-"Lead time"
colnames(RM_data)[14]<-"Planner"
colnames(RM_data)[15]<-"Planner Name"
colnames(RM_data)[16]<-"Standard Cost"
colnames(RM_data)[17]<-"MOQ"
colnames(RM_data)[18]<-"EOQ"
colnames(RM_data)[19]<-"Safety Stock"
colnames(RM_data)[20]<-"Max Cycle Stock"
colnames(RM_data)[21]<-"Usable"
colnames(RM_data)[22]<-"Quality hold"
colnames(RM_data)[23]<-"Quality hold in $$"
colnames(RM_data)[24]<-"Soft Hold"
colnames(RM_data)[25]<-"On Hand(usable + soft hold)"
colnames(RM_data)[26]<-"On Hand in $$"
colnames(RM_data)[27]<-"Target Inv"
colnames(RM_data)[28]<-"Target Inv in $$"
colnames(RM_data)[29]<-"Max inv"
colnames(RM_data)[30]<-"Max inv $$"
colnames(RM_data)[31]<-"OPV"
colnames(RM_data)[32]<-"PO in next 30 days"
colnames(RM_data)[33]<-"Receipt in the next 30 days"
colnames(RM_data)[34]<-"DOS"
colnames(RM_data)[35]<-"At Risk in $$"
colnames(RM_data)[36]<-"Inv Health"
colnames(RM_data)[37]<-"Current month dep demand"
colnames(RM_data)[38]<-"Next month dep demand"
colnames(RM_data)[39]<-"Total dep. demand Next 6 Months"
colnames(RM_data)[40]<-"Total Last 6 mos Sales"
colnames(RM_data)[41]<-"Total Last 12 mos Sales"
colnames(RM_data)[42]<-"has Max?"
colnames(RM_data)[43]<-"on hand Inv >max"
colnames(RM_data)[44]<-"on hand Inv <= max"
colnames(RM_data)[45]<-"on hand Inv > target"
colnames(RM_data)[46]<-"on hand Inv <= target"
colnames(RM_data)[47]<-"IQR $$"
colnames(RM_data)[48]<-"UPI$$"
colnames(RM_data)[49]<-"IQR $$ + Hold $$"
colnames(RM_data)[50]<-"UPI$$ + Hold $$"



writexl::write_xlsx(RM_data, "IQR_Report_5.25.2022.xlsx")

#### Look for the opportunity for save in RData format #### so that data can be easily pulled next following week. 



# when you run, make sure, Sys.Date() -6




# OPV (exception report) is from ref. not campus ref. Need to revise it. 


# Inventory analysis combine completed. move to Margaret's file
# 9:22 Canada video



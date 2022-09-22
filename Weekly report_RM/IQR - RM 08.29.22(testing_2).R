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
Planner_adress <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/Address Book - 08.23.22.xlsx", 
                             sheet = "Sheet1", col_types = c("text", 
                                                             "text", "text", "text", "text"))

names(Planner_adress) <- str_replace_all(names(Planner_adress), c(" " = "_"))

colnames(Planner_adress)[1] <- "Planner"


# Exception Report ----

exception_report <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_7 9.21.22/exception report 09.21.22.xlsx")

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

exception_report[, -32] -> exception_report
names(exception_report) <- str_replace_all(names(exception_report), c(" " = "_"))


exception_report %<>% 
  dplyr::mutate(ref = paste0(B_P, "_", ItemNo)) %>% 
  dplyr::relocate(ref) 

readr::type_convert(exception_report) -> exception_report

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
                campus = Campus) -> exception_report


# get the RM Item only. 
exception_report %>% 
  dplyr::mutate(ItemNo = as.numeric(ItemNo)) %>% 
  dplyr::mutate(item_na = is.na(ItemNo)) %>% 
  dplyr::filter(item_na == FALSE) %>% 
  dplyr::mutate(campus_na = is.na(campus)) %>% 
  dplyr::filter(campus_na == FALSE) -> exception_report

exception_report$ItemNo <- as.character(exception_report$ItemNo)


# exception report for safety_stock
exception_report %>% 
  dplyr::select(Loc_SKU, Safety_Stock) -> exception_report_ss

exception_report_ss %>% 
  dplyr::group_by(Loc_SKU) %>% 
  dplyr::summarise(Safety_Stock = sum(Safety_Stock, na.rm = TRUE)) -> exception_report_ss


# exception report for lead time
exception_report %>% 
  dplyr::arrange(Loc_SKU, desc(Leadtime_Days)) -> exception_report_lead

exception_report_lead[!duplicated(exception_report_lead[,c("Loc_SKU")]),] -> exception_report_lead

# exception report for MOQ
exception_report %>% 
  dplyr::arrange(Loc_SKU, desc(Reorder_MIN)) -> exception_report_moq

exception_report_moq[!duplicated(exception_report_moq[,c("Loc_SKU")]),] -> exception_report_moq


# exception report for Supplier No
exception_report %>% 
  dplyr::mutate(Supplier_No = replace(Supplier_No, is.na(Supplier_No), 0)) %>% 
  dplyr::rename(Supplier = Supplier_No) -> exception_report_supplier_no


# remove duplicated value - prioritize bigger Loc Number (RM only)

exception_report %>% 
  dplyr::mutate(B_P = as.integer(B_P)) %>% 
  dplyr::arrange(Loc_SKU, desc(B_P)) -> exception_report


# exception report Planner NA to 0
exception_report %>% 
  dplyr::mutate(Planner = replace(Planner, is.na(Planner), 0)) -> exception_report


# Pivoting exception_report
reshape2::dcast(exception_report, Loc_SKU ~ ., value.var = "Safety_Stock", sum) %>% 
  dplyr::rename(Safety_Stock = ".") -> exception_report_pivot




# Read IQR Report ----

RM_data <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Raw Material Inventory Health (IQR) - 09.14.22.xlsx", 
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
colnames(RM_data)[35] <- "At Risk in $$"
colnames(RM_data)[36] <- "Inv Health"
colnames(RM_data)[37] <- "Current month dep demand"
colnames(RM_data)[38] <- "Next month dep demand"
colnames(RM_data)[39] <- "Total dep. demand Next 6 Months"
colnames(RM_data)[40] <- "Total Last 6 mos Sales"
colnames(RM_data)[41] <- "Total Last 12 mos Sales "
colnames(RM_data)[42] <- "has Max?"
colnames(RM_data)[43] <- "on hand Inv>max"
colnames(RM_data)[44] <- "on hand Inv<=max"
colnames(RM_data)[45] <- "on hand Inv>target"
colnames(RM_data)[46] <- "on hand Inv<=target"
colnames(RM_data)[47] <- "IQR $$"
colnames(RM_data)[48] <- "UPI$$"
colnames(RM_data)[49] <- "IQR $$+Hold $$"
colnames(RM_data)[50] <- "UPI$$+Hold $$"


names(RM_data) <- stringr::str_replace_all(names(RM_data), c(" " = "_"))

RM_data %>% 
  dplyr::mutate(Loc_SKU = gsub("-", "_", Loc_SKU)) %>% 
  dplyr::relocate(Loc_SKU, .before = Supplier_No) -> RM_data



# Inventory Analysis Read RM ----

Inventory_analysis_RM <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_7 9.21.22/Inventory Report for all locations - 09.21.22.xlsx", 
                                    sheet = "RM")



Inventory_analysis_RM[-1,] -> Inventory_analysis_RM
colnames(Inventory_analysis_RM) <- Inventory_analysis_RM[1, ]
Inventory_analysis_RM[-1, ] -> Inventory_analysis_RM


colnames(Inventory_analysis_RM)[1] <- "Location"
colnames(Inventory_analysis_RM)[2] <- "Location_Nm"
colnames(Inventory_analysis_RM)[3] <- "Campus"
colnames(Inventory_analysis_RM)[4] <- "SKU"
colnames(Inventory_analysis_RM)[5] <- "Description"
colnames(Inventory_analysis_RM)[6] <- "Inventory_Status"
colnames(Inventory_analysis_RM)[7] <- "Inventory_Hold_Status"
colnames(Inventory_analysis_RM)[8] <- "Inventory_Qty_Cases"


Inventory_analysis <- Inventory_analysis_RM
readr::type_convert(Inventory_analysis) -> Inventory_analysis

# Vlookup - campus
# merge(Inventory_analysis, Campus_ref[, c("Location", "Campus")], by = "Location", all.x = TRUE) -> Inventory_analysis

Inventory_analysis %>%  
  dplyr::mutate(SKU = sub("^0+", "", SKU)) %>% 
  dplyr::mutate(campus_ref = paste0(Campus, "_", SKU), campus_ref = gsub("-", "", campus_ref)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", SKU), ref = gsub("-", "", ref)) %>% 
  dplyr::relocate(ref, campus_ref, Campus) %>% 
  dplyr::rename(campus = Campus) -> Inventory_analysis


# Inventory_analysis_pivot_ref

reshape2::dcast(Inventory_analysis, ref ~ Inventory_Hold_Status, value.var = "Inventory_Qty_Cases", sum) -> pivot_ref_Inventory_analysis
reshape2::dcast(Inventory_analysis, campus_ref ~ Inventory_Hold_Status, value.var = "Inventory_Qty_Cases", sum) -> pivot_campus_ref_Inventory_analysis

pivot_campus_ref_Inventory_analysis %<>% 
  dplyr::rename(Usable = Useable, Loc_SKU = campus_ref, Hard_Hold = "Hard Hold", Soft_Hold = "Soft Hold")

# BoM_dep_demand ----
BoM_dep_demand <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_7 9.21.22/Bill of Material.xlsx",
                             sheet = "Sheet1")

BoM_dep_demand %>% 
  janitor::clean_names() %>% 
  dplyr::rename(Loc_SKU = comp_ref) %>% 
  dplyr::mutate(Loc_SKU = gsub("-", "_", Loc_SKU)) %>% 
  data.frame() -> BoM_dep_demand

BoM_dep_demand[is.na(BoM_dep_demand)] <- 0

BoM_dep_demand %>% 
  dplyr::group_by(Loc_SKU) %>% 
  dplyr::summarise(mon_a_dep_demand = sum(mon_a_dep_demand),
                   mon_b_dep_demand = sum(mon_b_dep_demand),
                   mon_c_dep_demand = sum(mon_c_dep_demand),
                   mon_d_dep_demand = sum(mon_d_dep_demand),
                   mon_e_dep_demand = sum(mon_e_dep_demand),
                   mon_f_dep_demand = sum(mon_f_dep_demand)) %>% 
  dplyr::rename(current_month = mon_a_dep_demand,
                next_month = mon_b_dep_demand) %>% 
  dplyr::mutate(sum_of_months = current_month + next_month + mon_c_dep_demand + mon_d_dep_demand + mon_e_dep_demand + mon_f_dep_demand) -> BoM_dep_demand
  


BoM_dep_demand %>% filter(Loc_SKU == "622_240202061")


# Consumption data component # Updated once a month ----
consumption_data <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/consumption data component - 09.14.22.xlsx")

consumption_data[-1:-2,] -> consumption_data
colnames(consumption_data) <- consumption_data[1, ]
consumption_data[-1, ] -> consumption_data


colnames(consumption_data)[1] <- "Loc_SKU"
colnames(consumption_data)[ncol(consumption_data)-1] <- "sum_12mos"
colnames(consumption_data)[ncol(consumption_data)] <- "sum_6mos"

consumption_data %>% 
  dplyr::mutate(Loc_SKU = gsub("-", "_", Loc_SKU)) -> consumption_data

consumption_data %>% 
  data.frame() %>% 
  readr::type_convert() -> consumption_data

consumption_data[is.na(consumption_data)] <- 0


# SS Optimization RM for EOQ ----
SS_optimization <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_6 9.14.22/SS Optimization by Location - Raw Material August 2022.xlsx",
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
po <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_7 9.21.22/wo receipt custord po - 09.21.22.xlsx", 
                 sheet = "po", col_names = FALSE)



po %>% 
  dplyr::rename(aa = "...1") %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8"), sep = "~") %>% 
  dplyr::rename(a = "1") %>% 
  tidyr::separate(a, c("global", "rp", "Item")) %>% 
  dplyr::rename(Loc = "2",
                Qty = "5",
                PO_No = "6",
                date = "7") %>% 
  dplyr::select(-global, -rp, -"3", -"4", -"8") %>% 
  dplyr::mutate(date = as.Date(date)) %>% 
  dplyr::mutate(year = year(date),
                month = month(date),
                day = day(date))%>% 
  readr::type_convert() %>% 
  dplyr::mutate(month_year = paste0(month, "_", year)) %>% 
  dplyr::mutate(Loc = sub("^0+", "", Loc),
                Item = sub("^0+", "", Item)) %>% 
  dplyr::mutate(ref = paste0(Loc, "_", Item)) %>% 
  dplyr::relocate(ref) -> PO



# PO_Pivot 
PO %>% 
  dplyr::mutate(next_28_days = ifelse(date >= Sys.Date() & date <= Sys.Date() + 28, "Y", "N")) -> PO


reshape2::dcast(PO, ref ~ next_28_days, value.var = "Qty", sum) %>% 
  dplyr::rename(Loc_SKU = ref) -> PO_Pivot

rm(po)

# Custord Receipt ----
receipt <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_7 9.21.22/wo receipt custord po - 09.21.22.xlsx", 
                      sheet = "receipt", col_names = FALSE)


# Base receipt variable
receipt %>% 
  dplyr::rename(aa = "...1") %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8"), sep = "~") %>% 
  dplyr::rename(a = "1") %>% 
  tidyr::separate(a, c("global", "rp", "Item")) %>% 
  dplyr::rename(Loc = "2",
                Qty = "5",
                date = "7") %>% 
  dplyr::select(-global, -rp, -"3", -"4", -"6", -"8") %>% 
  dplyr::mutate(Item = gsub("^0+", "", Item),
                Loc = gsub("^0+", "", Loc)) %>% 
  dplyr::mutate(date = as.Date(date),
                year = year(date),
                month = month(date),
                day = day(date)) %>% 
  readr::type_convert() %>% 
  dplyr::mutate(ref = paste0(Loc, "_", Item),
                next_28_days = ifelse(date >= Sys.Date() & date <= Sys.Date() + 28, "Y", "N")) %>% 
  dplyr::relocate(ref) -> Receipt


rm(receipt)

# Receipt_Pivot 
reshape2::dcast(Receipt, ref ~ next_28_days, value.var = "Qty", sum) %>% 
  dplyr::rename(Loc_SKU = ref) -> Receipt_Pivot  

#####################################################################################################################
######################################################## ETL ########################################################
#####################################################################################################################

# vlookup - UoM
merge(RM_data, exception_report[, c("Loc_SKU", "UOM")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(UOM = replace(UOM, is.na(UOM), "DNRR")) %>% 
  dplyr::relocate(UOM, .after = UoM) %>% 
  dplyr::select(-UoM) -> RM_data

RM_data[!duplicated(RM_data[,c("Loc_SKU")]),] -> RM_data

# vlookup - Supplier No
merge(RM_data, exception_report_supplier_no[, c("Loc_SKU", "Supplier")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::arrange(Loc_SKU, desc(Supplier)) %>% 
  dplyr::select(-Supplier_No) %>% 
  dplyr::rename(Supplier_No = Supplier) %>% 
  dplyr::mutate(Supplier_No = replace(Supplier_No, is.na(Supplier_No), "DNRR")) -> RM_data

RM_data[!duplicated(RM_data[,c("Loc_SKU")]),] -> RM_data

# vlookup - Lead Time
exception_report_lead %>% 
  dplyr::mutate(Leadtime_Days = replace(Leadtime_Days, is.na(Leadtime_Days), 0)) -> exception_report_lead

merge(RM_data, exception_report_lead[, c("Loc_SKU", "Leadtime_Days")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::relocate(Leadtime_Days, .after = Lead_time) %>% 
  dplyr::mutate(Leadtime_Days = replace(Leadtime_Days, is.na(Leadtime_Days), "DNRR")) %>% 
  dplyr::select(-Lead_time) %>% 
  dplyr::rename(Lead_time = Leadtime_Days) -> RM_data


# vlookup - Planner
merge(RM_data, exception_report[, c("Loc_SKU", "Planner")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::relocate(Planner.y, .after = "Planner.x") %>% 
  dplyr::select(-Planner.x) %>% 
  dplyr::rename(Planner = Planner.y) %>% 
  dplyr::mutate(Planner = replace(Planner, is.na(Planner), "DNRR")) -> RM_data

RM_data[!duplicated(RM_data[,c("Loc_SKU")]),] -> RM_data

# vlookup - Planner Name
merge(RM_data, Planner_adress[, c("Planner", "Alpha_Name")], by = "Planner", all.x = TRUE) %>% 
  dplyr::relocate(Alpha_Name, .after = Planner_Name) %>% 
  dplyr::select(-Planner_Name) %>% 
  dplyr::rename(Planner_Name = Alpha_Name) %>% 
  dplyr::relocate(Planner, .before = Planner_Name) %>% 
  dplyr::mutate(Planner_Name = ifelse(Planner == "DNRR", "DNRR", Planner_Name)) -> RM_data



# vlookup - MOQ
exception_report_moq %>% 
  dplyr::mutate(Reorder_MIN = replace(Reorder_MIN, is.na(Reorder_MIN), 0)) -> exception_report_moq

merge(RM_data, exception_report_moq[, c("Loc_SKU", "Reorder_MIN")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::relocate(Reorder_MIN, .after = MOQ) %>% 
  dplyr::select(-MOQ) %>% 
  dplyr::rename(MOQ = Reorder_MIN) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# vlookup - Safety Stock
merge(RM_data, exception_report_ss[, c("Loc_SKU", "Safety_Stock")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(Safety_Stock.y = round(Safety_Stock.y, 0)) %>% 
  dplyr::mutate(Safety_Stock.y = replace(Safety_Stock.y, is.na(Safety_Stock.y), 0)) %>% 
  dplyr::relocate(Safety_Stock.y, .after = Safety_Stock.x) %>% 
  dplyr::select(-Safety_Stock.x) %>% 
  dplyr::rename(Safety_Stock = Safety_Stock.y) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# vlookup - Usable
merge(RM_data, pivot_campus_ref_Inventory_analysis[, c("Loc_SKU", "Usable")], by = "Loc_SKU", all.x = TRUE) %>%
  dplyr::mutate(Usable.y = round(Usable.y, 2)) %>% 
  dplyr::mutate(Usable.y = replace(Usable.y, is.na(Usable.y), 0)) %>% 
  dplyr::mutate(Usable.y = as.integer(Usable.y)) %>% 
  dplyr::relocate(Usable.y, .after = Usable.x) %>% 
  dplyr::select(-Usable.x) %>% 
  dplyr::rename(Usable = Usable.y) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# vlookup - Quality Hold
merge(RM_data, pivot_campus_ref_Inventory_analysis[, c("Loc_SKU", "Hard_Hold")], by = "Loc_SKU", all.x = TRUE) %>%
  dplyr::mutate(Hard_Hold = round(Hard_Hold, 2)) %>% 
  dplyr::mutate(Hard_Hold = replace(Hard_Hold, is.na(Hard_Hold), 0)) %>% 
  dplyr::relocate(Hard_Hold, .after = Quality_hold) %>% 
  dplyr::select(-Quality_hold) %>% 
  dplyr::rename(Quality_hold = Hard_Hold) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data

# Calculation - Quality Hold in $$
RM_data %>% 
  dplyr::mutate(Quality_hold_in_cost = Quality_hold * Standard_Cost) %>% 
  dplyr::mutate(Quality_hold_in_cost = round(Quality_hold_in_cost, 2)) %>% 
  dplyr::mutate(Quality_hold_in_cost = replace(Quality_hold_in_cost, is.na(Quality_hold_in_cost), 0)) %>% 
  dplyr::rename("Quality_hold_in_$$" = Quality_hold_in_cost) -> RM_data


# vlookup - Soft Hold
merge(RM_data, pivot_campus_ref_Inventory_analysis[, c("Loc_SKU", "Soft_Hold")], by = "Loc_SKU", all.x = TRUE) %>%
  dplyr::mutate(Soft_Hold.y = round(Soft_Hold.y, 2)) %>% 
  dplyr::mutate(Soft_Hold.y = replace(Soft_Hold.y, is.na(Soft_Hold.y), 0)) %>% 
  dplyr::relocate(Soft_Hold.y, .after = Soft_Hold.x) %>% 
  dplyr::select(-Soft_Hold.x) %>% 
  dplyr::rename(Soft_Hold = Soft_Hold.y) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data

# Calculation - On Hand (usable + soft hold)
RM_data %>% 
  dplyr::mutate(On_Hand_usable_and_soft_hold = Usable + Soft_Hold) -> RM_data

# Calculation - On Hand in $$
RM_data %>% 
  dplyr::mutate(On_Hand_in_cost = On_Hand_usable_and_soft_hold * Standard_Cost) %>% 
  dplyr::mutate(On_Hand_in_cost = round(On_Hand_in_cost, 2)) %>% 
  dplyr::mutate(On_Hand_in_cost = replace(On_Hand_in_cost, is.na(On_Hand_in_cost), 0)) %>% 
  dplyr::rename("On_Hand_in_$$" = On_Hand_in_cost) -> RM_data



# vlookup - OPV
exception_report %>% 
  dplyr::arrange(Loc_SKU, desc(Order_Policy_Value)) -> exception_report_opv

exception_report_opv[!duplicated(exception_report_opv[,c("Loc_SKU")]),] -> exception_report_opv

merge(RM_data, exception_report_opv[, c("Loc_SKU", "Order_Policy_Value")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(Order_Policy_Value = replace(Order_Policy_Value, is.na(Order_Policy_Value), 0)) %>% 
  dplyr::relocate(Order_Policy_Value, .after = OPV) %>% 
  dplyr::select(-OPV) %>% 
  dplyr::rename(OPV = Order_Policy_Value) -> RM_data

RM_data[!duplicated(RM_data[,c("Loc_SKU")]),] -> RM_data



# vlookup - PO in next 28 days
merge(RM_data, PO_Pivot[, c("Loc_SKU", "Y")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(Y = round(Y, 2)) %>% 
  dplyr::mutate(Y = replace(Y, is.na(Y), 0)) %>% 
  dplyr::relocate(Y, .after = PO_in_next_28_days) %>% 
  dplyr::select(-PO_in_next_28_days) %>% 
  dplyr::rename(PO_in_next_28_days = Y) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data



# vlookup - Receipt in next 28 days
merge(RM_data, Receipt_Pivot[, c("Loc_SKU", "Y")], by = "Loc_SKU", all.x = TRUE) %>%
  dplyr::mutate(Y = round(Y, 2)) %>% 
  dplyr::mutate(Y = replace(Y, is.na(Y), 0)) %>% 
  dplyr::relocate(Y, .after = Receipt_in_the_next_28_days) %>% 
  dplyr::select(-Receipt_in_the_next_28_days) %>% 
  dplyr::rename(Receipt_in_the_next_28_days = Y) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# vlookup - Current month dep demand
merge(RM_data, BoM_dep_demand[, c("Loc_SKU", "current_month")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(current_month = replace(current_month, is.na(current_month), 0)) %>% 
  dplyr::relocate(current_month, .after = Current_month_dep_demand) %>% 
  dplyr::select(-Current_month_dep_demand) %>% 
  dplyr::rename(Current_month_dep_demand = current_month) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data



# vlookup - Next month dep demand
merge(RM_data, BoM_dep_demand[, c("Loc_SKU", "next_month")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(next_month = replace(next_month, is.na(next_month), 0)) %>% 
  dplyr::relocate(next_month, .after = Next_month_dep_demand) %>% 
  dplyr::select(-Next_month_dep_demand) %>% 
  dplyr::rename(Next_month_dep_demand = next_month) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# vlookup - Total dep. demand Next 6 Months
merge(RM_data, BoM_dep_demand[, c("Loc_SKU", "sum_of_months")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(sum_of_months = replace(sum_of_months, is.na(sum_of_months), 0)) %>% 
  dplyr::relocate(sum_of_months, .after = Total_dep._demand_Next_6_Months) %>% 
  dplyr::select(-Total_dep._demand_Next_6_Months) %>% 
  dplyr::rename(Total_dep_demand_Next_6_Months = sum_of_months) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# Calculation - DOS
RM_data %>% 
  dplyr::mutate(DOS = On_Hand_usable_and_soft_hold / (pmax(Current_month_dep_demand, Next_month_dep_demand)/30)) %>% 
  dplyr::mutate(DOS = round(DOS, 0)) %>% 
  dplyr::mutate(DOS = replace(DOS, is.na(DOS), 0)) %>% 
  dplyr::mutate(DOS = replace(DOS, is.nan(DOS), 0)) %>% 
  dplyr::mutate(DOS = replace(DOS, is.infinite(DOS), 0)) -> RM_data


# vlookup - Total Last 6 mos Sales
merge(RM_data, consumption_data[, c("Loc_SKU", "sum_6mos")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(sum_6mos = as.double(sum_6mos)) %>% 
  dplyr::mutate(sum_6mos = round(sum_6mos, 2)) %>% 
  dplyr::mutate(sum_6mos = replace(sum_6mos, is.na(sum_6mos), 0)) %>% 
  dplyr::relocate(sum_6mos, .after = Total_Last_6_mos_Sales) %>% 
  dplyr::select(-Total_Last_6_mos_Sales) %>% 
  dplyr::rename(Total_Last_6_mos_Sales = sum_6mos) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data

# vlookup - Total Last 12 mos Sales
merge(RM_data, consumption_data[, c("Loc_SKU", "sum_12mos")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(sum_12mos = as.double(sum_12mos)) %>% 
  dplyr::mutate(sum_12mos = round(sum_12mos, 2)) %>% 
  dplyr::mutate(sum_12mos = replace(sum_12mos, is.na(sum_12mos), 0)) %>% 
  dplyr::relocate(sum_12mos, .after = Total_Last_12_mos_Sales_) %>% 
  dplyr::select(-Total_Last_12_mos_Sales_) %>% 
  dplyr::rename(Total_Last_12_mos_Sales = sum_12mos) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data

# vlookup - EOQ
merge(RM_data, SS_optimization[, c("Loc_SKU", "EOQ_adjusted")], by = "Loc_SKU", all.x = TRUE) %>% 
  dplyr::mutate(EOQ_adjusted = as.double(EOQ_adjusted)) %>% 
  dplyr::mutate(EOQ_adjusted = round(EOQ_adjusted, 0)) %>% 
  dplyr::mutate(EOQ_adjusted = replace(EOQ_adjusted, is.na(EOQ_adjusted), 0)) %>% 
  dplyr::relocate(EOQ_adjusted, .after = EOQ) %>% 
  dplyr::select(-EOQ) %>% 
  dplyr::rename(EOQ = EOQ_adjusted) %>% 
  dplyr::relocate(Loc_SKU, .after = Item) -> RM_data


# Calculation - Max Cycle Stock
RM_data %>% 
  dplyr::mutate(Max_Cycle_Stock =
                  pmax(EOQ, MOQ, OPV*(Next_month_dep_demand/20.83), OPV*(Total_Last_12_mos_Sales/250))) %>% 
  dplyr::mutate(Max_Cycle_Stock = round(Max_Cycle_Stock, 2)) %>% 
  dplyr::mutate(Max_Cycle_Stock = replace(Max_Cycle_Stock, is.na(Max_Cycle_Stock), 0)) -> RM_data


# Calculation - Target Inv
RM_data %>% 
  dplyr::mutate(Target_Inv = Safety_Stock + Max_Cycle_Stock / 2) -> RM_data

# Calculation - Target Inv in $$
RM_data %>% 
  dplyr::mutate(Target_Inv_in_cost = Target_Inv * Standard_Cost) %>% 
  dplyr::mutate(Target_Inv_in_cost = as.double(Target_Inv_in_cost)) %>% 
  dplyr::mutate(Target_Inv_in_cost = round(Target_Inv_in_cost, 2)) %>% 
  dplyr::mutate(Target_Inv_in_cost = replace(Target_Inv_in_cost, is.na(Target_Inv_in_cost), 0)) %>% 
  dplyr::rename("Target_Inv_in_$$" = Target_Inv_in_cost) -> RM_data


# Calculation - Max inv
RM_data %>% 
  dplyr::mutate(Max_inv = Safety_Stock + Max_Cycle_Stock) -> RM_data

# Calculation - Max inv $$
RM_data %>% 
  dplyr::mutate(Max_inv_cost = Max_inv * Standard_Cost) %>% 
  dplyr::mutate(Max_inv_cost = as.double(Max_inv_cost)) %>% 
  dplyr::mutate(Max_inv_cost = round(Max_inv_cost, 2)) %>% 
  dplyr::mutate(Max_inv_cost = replace(Max_inv_cost, is.na(Max_inv_cost), 0)) %>% 
  dplyr::rename("Max_inv_$$" = Max_inv_cost) -> RM_data


# Calculation - has Max?
RM_data %>% 
  dplyr::mutate("has_Max?" = ifelse(Max_inv > 0, 1, 0)) -> RM_data

# Calculation - on hand Inv > max
RM_data %>% 
  dplyr::mutate("on_hand_Inv>max" = ifelse(On_Hand_usable_and_soft_hold > Max_inv, 1, 0)) %>% 
  dplyr::rename(on_hand_Inv_greaterthan_max = "on_hand_Inv>max") -> RM_data

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
  dplyr::mutate(today = Sys.Date(),
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
  dplyr::select(-today, -diff_days) %>% 
  dplyr::rename("on_hand_Inv>max" = on_hand_Inv_greaterthan_max) -> RM_data



# Calculation - At Risk in $$
RM_data %<>% 
  dplyr::mutate("At_Risk_in_$$" = ifelse(Inv_Health=="At Risk",
                                         (On_Hand_usable_and_soft_hold -((pmax(Current_month_dep_demand,Next_month_dep_demand)/30) 
                                                                         *(Shelf_Life_day*0.6)))*Standard_Cost,0)) 

# Calculation - IQR $$
RM_data %>% 
  dplyr::rename(On_Hand_in_cost = "On_Hand_in_$$",
                At_Risk_in_cost = "At_Risk_in_$$",
                Max_inv_cost = "Max_inv_$$") %>% 
  dplyr::mutate("IQR_$$" = ifelse(Inv_Health == "DEAD" | Inv_Health == "HEALTHY" | Inv_Health == "BELOW SS", On_Hand_in_cost, 
                                  ifelse(Inv_Health == "AT RISK", At_Risk_in_cost, On_Hand_in_cost - Max_inv_cost))) -> RM_data


# Calculation - UPI $$
RM_data %<>% 
  dplyr::mutate("UPI$$" = ifelse(Inv_Health == "AT RISK", At_Risk_in_cost,
                                 ifelse(Inv_Health == "EXCESS", On_Hand_in_cost - Max_inv_cost,
                                        ifelse(Inv_Health == "DEAD", On_Hand_in_cost, 0)))) %>% 
  dplyr::rename("On_Hand_in_$$" = On_Hand_in_cost,
                "At_Risk_in_$$" = At_Risk_in_cost,
                "Max_inv_$$" = Max_inv_cost) 

# Calculation - IQR $$ + Hold $$
RM_data %<>% 
  dplyr::rename(IQR_cost = "IQR_$$",
                Quality_hold_in_Cost = "Quality_hold_in_$$") %>% 
  dplyr::mutate("IQR_$$+Hold_$$" = IQR_cost + Quality_hold_in_Cost) 

# Calculation - UPI $$ + Hold $$
RM_data %<>% 
  dplyr::rename(UPI_cost = "UPI$$") %>% 
  dplyr::mutate("UPI$$+Hold_$$" = UPI_cost + Quality_hold_in_Cost) %>% 
  dplyr::rename("IQR_$$" = IQR_cost,
                "Quality_hold_in_$$" = Quality_hold_in_Cost,
                "UPI$$" = UPI_cost)




######## Deleting items that we don't need ###########
RM_data %>% dplyr::filter(Loc_SKU != "60_8883") -> RM_data




#####################################################################################################################
########################################## Change Col names to original #############################################
#####################################################################################################################

# sum supposed to be
# OPV: 273956  (Need an explanation again the sorting logic)
# Supplier NO: 1283530430   (This one, I got the similar number)

# Total row number, I have 9689 instead of 9691
# check with "36_44391", "36_45854"

# test
RM_data %>% 
  dplyr::mutate(Lead_time = as.numeric(Lead_time),
                Supplier_No = as.numeric(Supplier_No)) -> test_data

sum(test_data$Lead_time, na.rm = TRUE)  
sum(test_data$MOQ, na.rm = TRUE)
sum(test_data$Supplier_No, na.rm = TRUE)
sum(test_data$Safety_Stock)
sum(test_data$Usable)
sum(test_data$Quality_hold)
sum(test_data$OPV)

test_data %>% filter(Item == 97491)

exception_report %>% 
  filter(Loc_SKU == "75_42556")

exception_report_ss %>% 
  dplyr::filter(Loc_SKU == "75_42556")

RM_data %>% 
  dplyr::mutate(Loc_SKU = gsub("_", "-", Loc_SKU)) %>% 
  writexl::write_xlsx("test.xlsx")

RM_data %>% filter(Loc_SKU == "60_5198")


#


########### Don't forget to rearrange!! #################

RM_data %>% 
  dplyr::mutate(Loc_SKU = gsub("_", "-", Loc_SKU)) %>% 
  dplyr::relocate(Mfg_Loc, Loc_Name) -> RM_data


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



writexl::write_xlsx(RM_data, "IQR_Report_8.17.2022.xlsx")





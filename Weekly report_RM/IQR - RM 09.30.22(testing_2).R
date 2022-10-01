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
library(janitor)



##################################################################################################################################################################
##################################################################################################################################################################
##################################################################################################################################################################

# Planner Address Book (If updated, correct this link) ----
planner_adress <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/Address Book - 08.23.22.xlsx", 
                             sheet = "Sheet1", col_types = c("text", 
                                                             "text", "text", "text", "text"))

planner_adress %>% 
  janitor::clean_names() -> planner_adress

colnames(planner_adress)[1] <- "planner"


# Exception Report ----

exception_report <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_8 9.28.22/exception report 09.28.22.xlsx")

exception_report[-1:-2,] -> exception_report

colnames(exception_report) <- exception_report[1, ]
exception_report[-1, ] -> exception_report

exception_report %>% 
  janitor::clean_names() -> exception_report


exception_report[, -32] -> exception_report

exception_report %>% 
  dplyr::mutate(ref = paste0(b_p, "_", item_number)) %>% 
  dplyr::relocate(ref) -> exception_report

readr::type_convert(exception_report) -> exception_report

# Campus_ref pulling ----

campus_ref <- read_excel("S:/Supply Chain Projects/RStudio/BoM/Master formats/RM_on_Hand/Campus_ref.xlsx", 
                         col_types = c("numeric", "text", "text", 
                                       "numeric"))

campus_ref %>% 
  janitor::clean_names() -> campus_ref

str(campus_ref)

colnames(campus_ref)[1] <- "b_p"

campus_ref %>% 
  dplyr::mutate(location = b_p) -> campus_ref

# Vlookup for Campus_ref

merge(exception_report, campus_ref[, c("b_p", "campus")], by = "b_p", all.x = TRUE) %>% 
  dplyr::mutate(campus_ref = paste0(campus, "_", item_number)) %>% 
  dplyr::relocate(ref, campus_ref, campus) %>% 
  dplyr::rename(loc_sku = campus_ref) -> exception_report


# get the RM Item only. 
exception_report %>% 
  dplyr::mutate(item_number = as.numeric(item_number)) %>% 
  dplyr::mutate(item_na = is.na(item_number)) %>% 
  dplyr::filter(item_na == FALSE) %>% 
  dplyr::mutate(campus_na = is.na(campus)) %>% 
  dplyr::filter(campus_na == FALSE) -> exception_report

exception_report$item_number <- as.character(exception_report$item_number)


# exception report for safety_stock
exception_report %>% 
  dplyr::select(loc_sku, safety_stock) -> exception_report_ss

exception_report_ss %>% 
  dplyr::group_by(loc_sku) %>% 
  dplyr::summarise(safety_stock = sum(safety_stock, na.rm = TRUE)) -> exception_report_ss


# exception report for lead time
exception_report %>% 
  dplyr::arrange(loc_sku, desc(leadtime_days)) -> exception_report_lead

exception_report_lead[!duplicated(exception_report_lead[,c("loc_sku")]),] -> exception_report_lead

# exception report for MOQ
exception_report %>% 
  dplyr::arrange(loc_sku, desc(reorder_min)) -> exception_report_moq

exception_report_moq[!duplicated(exception_report_moq[,c("loc_sku")]),] -> exception_report_moq


# exception report for Supplier No
exception_report %>% 
  dplyr::mutate(supplier = replace(supplier, is.na(supplier), 0)) -> exception_report_supplier_no


# remove duplicated value - prioritize bigger Loc Number (RM only)

exception_report %>% 
  dplyr::mutate(b_p = as.integer(b_p)) %>% 
  dplyr::arrange(loc_sku, desc(b_p)) -> exception_report


# exception report Planner NA to 0
exception_report %>% 
  dplyr::mutate(planner = replace(planner, is.na(planner), 0)) -> exception_report


# Pivoting exception_report
reshape2::dcast(exception_report, loc_sku ~ ., value.var = "safety_stock", sum) %>% 
  dplyr::rename(safety_stock = ".") -> exception_report_pivot




# Read IQR Report ----

rm_data <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Raw Material Inventory Health (IQR) - 09.21.22 rev2.xlsx", 
                      sheet = "RM data", col_names = FALSE)

rm_data[-1:-3,] -> rm_data
colnames(rm_data) <- rm_data[1, ]
rm_data[-1, ] -> rm_data

rm_data %>% 
  janitor::clean_names() %>% 
  readr::type_convert()-> rm_data

str(rm_data)

colnames(rm_data)[23] <- "quality_hold_in_cost"
colnames(rm_data)[26] <- "on_hand_in_cost"
colnames(rm_data)[28] <- "target_inv_in_cost"
colnames(rm_data)[30] <- "max_inv_cost"
colnames(rm_data)[35] <- "at_risk_in_cost"
colnames(rm_data)[43] <- "on_hand_inv_greater_than_max"
colnames(rm_data)[44] <- "on_hand_inv_less_or_equal_than_max"
colnames(rm_data)[45] <- "on_hand_inv_greater_than_target"
colnames(rm_data)[46] <- "on_hand_inv_less_or_equal_than_target"
colnames(rm_data)[47] <- "iqr_cost"
colnames(rm_data)[48] <- "upi_cost"
colnames(rm_data)[49] <- "iqr_cost_plus_hold_cost"
colnames(rm_data)[50] <- "upi_cost_plus_hold_cost"



rm_data %>% 
  dplyr::mutate(loc_sku = gsub("-", "_", loc_sku)) %>% 
  dplyr::relocate(loc_sku, .before = supplier_number) -> rm_data



# Inventory Analysis Read RM ----

inventory_analysis_rm <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_8 9.28.22/Inventory Report for all locations - 09.28.22.xlsx", 
                                    sheet = "RM")



inventory_analysis_rm[-1,] -> inventory_analysis_rm
colnames(inventory_analysis_rm) <- inventory_analysis_rm[1, ]
inventory_analysis_rm[-1, ] -> inventory_analysis_rm

inventory_analysis_rm %>% 
  janitor::clean_names() %>% 
  readr::type_convert() -> inventory_analysis_rm

colnames(inventory_analysis_rm)[2] <- "location_name"
colnames(inventory_analysis_rm)[5] <- "description"
colnames(inventory_analysis_rm)[7] <- "inventory_hold_status"
colnames(inventory_analysis_rm)[8] <- "inventory_qty_cases"


inventory_analysis <- inventory_analysis_rm
readr::type_convert(inventory_analysis) -> inventory_analysis

# Vlookup - campus
# merge(Inventory_analysis, Campus_ref[, c("Location", "Campus")], by = "Location", all.x = TRUE) -> Inventory_analysis

inventory_analysis %>%  
  dplyr::mutate(item = sub("^0+", "", item)) %>% 
  dplyr::mutate(campus_ref = paste0(campus, "_", item), campus_ref = gsub("-", "", campus_ref)) %>% 
  dplyr::mutate(ref = paste0(location, "_", item), ref = gsub("-", "", ref)) %>% 
  dplyr::relocate(ref, campus_ref, campus) -> inventory_analysis


# Inventory_analysis_pivot_ref

reshape2::dcast(inventory_analysis, ref ~ inventory_hold_status, value.var = "inventory_qty_cases", sum) -> pivot_ref_inventory_analysis
reshape2::dcast(inventory_analysis, campus_ref ~ inventory_hold_status, value.var = "inventory_qty_cases", sum) -> pivot_campus_ref_inventory_analysis

pivot_campus_ref_inventory_analysis %>% 
  dplyr::rename(usable = Useable, loc_sku = campus_ref, hard_hold = "Hard Hold", soft_hold = "Soft Hold") -> pivot_campus_ref_inventory_analysis

# BoM_dep_demand ----
bom_dep_demand <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_8 9.28.22/Bill of Material.xlsx",
                             sheet = "Sheet1")

bom_dep_demand %>% 
  janitor::clean_names() %>% 
  dplyr::rename(loc_sku = comp_ref) %>% 
  dplyr::mutate(loc_sku = gsub("-", "_", loc_sku)) %>% 
  data.frame() -> bom_dep_demand

bom_dep_demand[is.na(bom_dep_demand)] <- 0

bom_dep_demand %>% 
  dplyr::group_by(loc_sku) %>% 
  dplyr::summarise(mon_a_dep_demand = sum(mon_a_dep_demand),
                   mon_b_dep_demand = sum(mon_b_dep_demand),
                   mon_c_dep_demand = sum(mon_c_dep_demand),
                   mon_d_dep_demand = sum(mon_d_dep_demand),
                   mon_e_dep_demand = sum(mon_e_dep_demand),
                   mon_f_dep_demand = sum(mon_f_dep_demand)) %>% 
  dplyr::rename(current_month = mon_a_dep_demand,
                next_month = mon_b_dep_demand) %>% 
  dplyr::mutate(sum_of_months = current_month + next_month + mon_c_dep_demand + mon_d_dep_demand + mon_e_dep_demand + mon_f_dep_demand) -> bom_dep_demand



# Consumption data component # Updated once a month ----
consumption_data <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/consumption data component - 09.14.22.xlsx")

consumption_data[-1:-2,] -> consumption_data
colnames(consumption_data) <- consumption_data[1, ]
consumption_data[-1, ] -> consumption_data


colnames(consumption_data)[1] <- "loc_sku"
colnames(consumption_data)[ncol(consumption_data)-1] <- "sum_12mos"
colnames(consumption_data)[ncol(consumption_data)] <- "sum_6mos"

consumption_data %>% 
  dplyr::mutate(loc_sku = gsub("-", "_", loc_sku)) -> consumption_data

consumption_data %>% 
  data.frame() %>% 
  readr::type_convert() -> consumption_data

consumption_data[is.na(consumption_data)] <- 0


# SS Optimization RM for EOQ ----
ss_optimization <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_6 9.14.22/SS Optimization by Location - Raw Material August 2022.xlsx",
                              sheet = "Sheet1")

ss_optimization[-1:-5,] -> ss_optimization
colnames(ss_optimization) <- ss_optimization[1, ]
ss_optimization[-1, ] -> ss_optimization

ss_optimization %>% 
  janitor::clean_names() %>% 
  readr::type_convert() -> ss_optimization


colnames(ss_optimization)[3] <- "loc_sku"
colnames(ss_optimization)[29] <- "standard_cost"
colnames(ss_optimization)[48] <- "eoq_adjusted"


data.frame(ss_optimization$loc_sku) -> ss_opt_loc_sku

ss_opt_loc_sku %>% 
  dplyr::mutate(ss_optimization.loc_sku = gsub("-", "_", ss_optimization.loc_sku)) %>% 
  dplyr::rename(loc_sku = ss_optimization.loc_sku) %>% 
  dplyr::bind_cols(ss_optimization) %>% 
  dplyr::select(-loc_sku...4) %>% 
  dplyr::rename(loc_sku = loc_sku...1) -> ss_optimization

ss_optimization[-which(duplicated(ss_optimization$loc_sku)),] -> ss_optimization

# Custord PO ----
po <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_8 9.28.22/wo receipt custord po - 09.28.22.xlsx", 
                 sheet = "po", col_names = FALSE)



po %>% 
  dplyr::rename(aa = "...1") %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8"), sep = "~") %>% 
  dplyr::rename(a = "1") %>% 
  tidyr::separate(a, c("global", "rp", "Item")) %>% 
  dplyr::rename(loc = "2",
                qty = "5",
                po_number = "6",
                date = "7") %>% 
  dplyr::select(-global, -rp, -"3", -"4", -"8") %>% 
  dplyr::mutate(date = as.Date(date)) %>% 
  dplyr::mutate(year = year(date),
                month = month(date),
                day = day(date))%>% 
  readr::type_convert() %>% 
  dplyr::mutate(month_year = paste0(month, "_", year)) %>% 
  dplyr::mutate(loc = sub("^0+", "", loc),
                Item = sub("^0+", "", Item)) %>% 
  dplyr::mutate(ref = paste0(loc, "_", Item)) %>% 
  dplyr::rename(item = Item) %>% 
  dplyr::relocate(ref) -> po



# PO_Pivot 
po %>% 
  dplyr::mutate(next_28_days = ifelse(date >= Sys.Date() & date <= Sys.Date() + 28, "Y", "N")) -> po


reshape2::dcast(po, ref ~ next_28_days, value.var = "qty", sum) %>% 
  dplyr::rename(loc_sku = ref) -> PO_Pivot



# Custord Receipt ----
receipt <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Test_8 9.28.22/wo receipt custord po - 09.28.22.xlsx", 
                      sheet = "receipt", col_names = FALSE)


# Base receipt variable
receipt %>% 
  dplyr::rename(aa = "...1") %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8"), sep = "~") %>% 
  dplyr::rename(a = "1") %>% 
  tidyr::separate(a, c("global", "rp", "Item")) %>% 
  dplyr::rename(loc = "2",
                qty = "5",
                date = "7") %>% 
  dplyr::select(-global, -rp, -"3", -"4", -"6", -"8") %>% 
  dplyr::mutate(Item = gsub("^0+", "", Item),
                loc = gsub("^0+", "", loc)) %>% 
  dplyr::mutate(date = as.Date(date),
                year = year(date),
                month = month(date),
                day = day(date)) %>% 
  readr::type_convert() %>% 
  dplyr::mutate(ref = paste0(loc, "_", Item),
                next_28_days = ifelse(date >= Sys.Date() & date <= Sys.Date() + 28, "Y", "N")) %>% 
  dplyr::relocate(ref) %>% 
  dplyr::rename(item = Item) -> receipt




# receipt_pivot 
reshape2::dcast(receipt, ref ~ next_28_days, value.var = "qty", sum) %>% 
  dplyr::rename(loc_sku = ref) -> receipt_pivot  

#####################################################################################################################
######################################################## ETL ########################################################
#####################################################################################################################

# vlookup - UoM
merge(rm_data, exception_report[, c("loc_sku", "uom")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::mutate(uom = replace(uom, is.na(uom), "DNRR")) %>% 
  dplyr::relocate(uom, .after = uo_m) %>% 
  dplyr::select(-uo_m) -> rm_data

rm_data[!duplicated(rm_data[,c("loc_sku")]),] -> rm_data


# vlookup - Supplier No
merge(rm_data, exception_report_supplier_no[, c("loc_sku", "supplier")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::arrange(loc_sku, desc(supplier)) %>% 
  dplyr::select(-supplier_number) %>% 
  dplyr::rename(supplier_number = supplier) %>% 
  dplyr::mutate(supplier_number = replace(supplier_number, is.na(supplier_number), "DNRR")) -> rm_data

rm_data[!duplicated(rm_data[,c("loc_sku")]),] -> rm_data

# vlookup - Lead Time

exception_report_lead %>% 
  dplyr::mutate(leadtime_days = replace(leadtime_days, is.na(leadtime_days), 0)) -> exception_report_lead

merge(rm_data, exception_report_lead[, c("loc_sku", "leadtime_days")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::relocate(leadtime_days, .after = lead_time) %>% 
  dplyr::mutate(leadtime_days = replace(leadtime_days, is.na(leadtime_days), "DNRR")) %>% 
  dplyr::select(-lead_time) %>% 
  dplyr::rename(lead_time = leadtime_days) -> rm_data


# vlookup - Planner
merge(rm_data, exception_report[, c("loc_sku", "planner")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::relocate(planner.y, .after = "planner.x") %>% 
  dplyr::select(-planner.x) %>% 
  dplyr::rename(planner = planner.y) %>% 
  dplyr::mutate(planner = replace(planner, is.na(planner), "DNRR")) -> rm_data

rm_data[!duplicated(rm_data[,c("loc_sku")]),] -> rm_data

# vlookup - Planner Name
merge(rm_data, planner_adress[, c("planner", "alpha_name")], by = "planner", all.x = TRUE) %>% 
  dplyr::relocate(alpha_name, .after = planner_name) %>% 
  dplyr::select(-planner_name) %>% 
  dplyr::rename(planner_name = alpha_name) %>% 
  dplyr::relocate(planner, .before = planner_name) %>% 
  dplyr::mutate(planner_name = ifelse(planner == "DNRR", "DNRR", planner_name)) -> rm_data



# vlookup - MOQ
exception_report_moq %>% 
  dplyr::mutate(reorder_min = replace(reorder_min, is.na(reorder_min), 0)) -> exception_report_moq

merge(rm_data, exception_report_moq[, c("loc_sku", "reorder_min")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::relocate(reorder_min, .after = moq) %>% 
  dplyr::select(-moq) %>% 
  dplyr::rename(moq = reorder_min) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data


# vlookup - Safety Stock
merge(rm_data, exception_report_ss[, c("loc_sku", "safety_stock")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::mutate(safety_stock.y = round(safety_stock.y, 0)) %>% 
  dplyr::mutate(safety_Stock.y = replace(safety_stock.y, is.na(safety_stock.y), 0)) %>% 
  dplyr::relocate(safety_stock.y, .after = safety_stock.x) %>% 
  dplyr::select(-safety_stock.x) %>% 
  dplyr::rename(safety_stock = safety_stock.y) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data


# vlookup - Usable
merge(rm_data, pivot_campus_ref_inventory_analysis[, c("loc_sku", "usable")], by = "loc_sku", all.x = TRUE) %>%
  dplyr::mutate(usable.y = round(usable.y, 2)) %>% 
  dplyr::mutate(usable.y = replace(usable.y, is.na(usable.y), 0)) %>% 
  dplyr::mutate(usable.y = as.integer(usable.y)) %>% 
  dplyr::relocate(usable.y, .after = usable.x) %>% 
  dplyr::select(-usable.x) %>% 
  dplyr::rename(usable = usable.y) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data


# vlookup - Quality Hold
merge(rm_data, pivot_campus_ref_inventory_analysis[, c("loc_sku", "hard_hold")], by = "loc_sku", all.x = TRUE) %>%
  dplyr::mutate(hard_hold = round(hard_hold, 2)) %>% 
  dplyr::mutate(hard_Hold = replace(hard_hold, is.na(hard_hold), 0)) %>% 
  dplyr::relocate(hard_hold, .after = quality_hold) %>% 
  dplyr::select(-quality_hold) %>% 
  dplyr::rename(quality_hold = hard_hold) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data

# Calculation - Quality Hold in $$
rm_data %>% 
  dplyr::mutate(standard_cost = as.numeric(standard_cost)) -> rm_data

rm_data %>% 
  dplyr::mutate(quality_hold_in_cost = quality_hold * standard_cost) %>% 
  dplyr::mutate(quality_hold_in_cost = round(quality_hold_in_cost, 2)) %>% 
  dplyr::mutate(quality_hold_in_cost = replace(quality_hold_in_cost, is.na(quality_hold_in_cost), 0)) -> rm_data


# vlookup - Soft Hold
merge(rm_data, pivot_campus_ref_inventory_analysis[, c("loc_sku", "soft_hold")], by = "loc_sku", all.x = TRUE) %>%
  dplyr::mutate(soft_hold.y = round(soft_hold.y, 2)) %>% 
  dplyr::mutate(soft_hold.y = replace(soft_hold.y, is.na(soft_hold.y), 0)) %>% 
  dplyr::relocate(soft_hold.y, .after = soft_hold.x) %>% 
  dplyr::select(-soft_hold.x) %>% 
  dplyr::rename(soft_hold = soft_hold.y) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data

# Calculation - On Hand (usable + soft hold)
rm_data %>% 
  dplyr::mutate(on_hand_usable_and_soft_hold = usable + soft_hold) -> rm_data

# Calculation - On Hand in $$
rm_data %>% 
  dplyr::mutate(on_hand_in_cost = on_hand_usable_and_soft_hold * standard_cost) %>% 
  dplyr::mutate(on_hand_in_cost = round(on_hand_in_cost, 2)) %>% 
  dplyr::mutate(on_hand_in_cost = replace(on_hand_in_cost, is.na(on_hand_in_cost), 0)) -> rm_data



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
  dplyr::mutate(Max_Cycle_Stock = round(Max_Cycle_Stock, 1)) %>% 
  dplyr::mutate(Max_Cycle_Stock = replace(Max_Cycle_Stock, is.na(Max_Cycle_Stock), 0)) %>% 
  dplyr::mutate(Max_Cycle_Stock = ifelse(Lead_time == "DNRR", EOQ, Max_Cycle_Stock)) -> RM_data


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
# RM_data %>% 
#   dplyr::mutate(today = Sys.Date(),
#                 today = as.Date(today, format = "%Y-%m-%d"),
#                 Birthday = as.Date(Birthday, format = "%Y-%m-%d"),
#                 diff_days = today - Birthday,
#                 diff_days = as.numeric(diff_days),
#                 Inv_Health = ifelse(On_Hand_usable_and_soft_hold < Safety_Stock, "BELOW SS", 
#                                     ifelse(Item_Type == "Non-Commodity" & DOS > 0.6*Shelf_Life_day, "AT RISK", 
#                                            ifelse(On_Hand_usable_and_soft_hold > 0 & Lead_time == "DNRR", "DEAD", 
#                                                   ifelse(On_Hand_usable_and_soft_hold > 0 & Current_month_dep_demand == 0 & 
#                                                            Next_month_dep_demand == 0 & Total_dep_demand_Next_6_Months == 0 & 
#                                                            diff_days > 90, "DEAD", 
#                                                          ifelse(on_hand_Inv_greaterthan_max == 0 | diff_days < 91, "HEALTHY", "EXCESS")))))) %>% 
#   dplyr::select(-today, -diff_days) %>% 
#   dplyr::rename("on_hand_Inv>max" = on_hand_Inv_greaterthan_max) -> RM_data


RM_data %>% 
  dplyr::mutate(Shelf_Life_day = as.numeric(Shelf_Life_day),
                Birthday = as.integer(Birthday)) -> RM_data


RM_data %>% 
  dplyr::mutate(today = Sys.Date(),
                today = as.Date(today, format = "%Y-%m-%d"),
                Birthday = as.Date(Birthday, origin = "1899-12-30"),
                diff_days = today - Birthday,
                diff_days = as.numeric(diff_days),
                Inv_Health = ifelse(On_Hand_usable_and_soft_hold < Safety_Stock, "BELOW SS", (ifelse(Item_Type == "Non-Commodity" & DOS > 0.6 * Shelf_Life_day, "AT RISK",
                                                                                                     ifelse((On_Hand_usable_and_soft_hold > 0 & Lead_time == "DNRR") | (On_Hand_usable_and_soft_hold > 0 & Current_month_dep_demand == 0 & Next_month_dep_demand == 0 & Total_dep_demand_Next_6_Months == 0 & diff_days > 90), "DEAD",
                                                                                                            ifelse(on_hand_Inv_greaterthan_max == 0, "HEALTHY", "EXCESS")))))) %>% 
  dplyr::rename("on_hand_Inv>max" = on_hand_Inv_greaterthan_max) -> RM_data





# Calculation - At Risk in $$
RM_data %>% 
  dplyr::mutate("At_Risk_in_$$" = ifelse(Inv_Health == "At Risk",
                                         (On_Hand_usable_and_soft_hold -((pmax(Current_month_dep_demand,Next_month_dep_demand)/30) 
                                                                         *(Shelf_Life_day*0.6)))*Standard_Cost,0)) -> RM_data

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


# test excel
writexl::write_xlsx(RM_data, "test.xlsx")

#####################################################################################################################
########################################## Change Col names to original #############################################
#####################################################################################################################

RM_data %>% 
  dplyr::filter(Loc_SKU == "622_310000611")

#


########### Don't forget to rearrange and bring cols only what you need! #################
RM_data %>% 
  dplyr::mutate(Loc_SKU = gsub("_", "-", Loc_SKU)) %>% 
  janitor::clean_names() %>% 
  dplyr::select() -> RM_data


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





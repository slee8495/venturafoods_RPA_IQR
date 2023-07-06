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
library(rio)


##################################################################################################################################################################
##################################################################################################################################################################
##################################################################################################################################################################

# Planner Address Book (If updated, correct this link) ----
planner_adress <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/Address Book - 07.03.23.xlsx", 
                             sheet = "Sheet1", col_types = c("text", 
                                                             "text", "text", "text", "text"))

planner_adress %>% 
  janitor::clean_names() -> planner_adress

colnames(planner_adress)[1] <- "planner"


# Exception Report ----

exception_report <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/BoM version 2/Weekly Run/7.5.2023/exception report.xlsx")

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

campus_ref <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/campus reference.xlsx", 
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

rm_data <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Raw Material Inventory Health (IQR) - 06.28.23.xlsx", 
                      sheet = "RM data", col_names = FALSE)

rm_data[-1:-3,] -> rm_data
colnames(rm_data) <- rm_data[1, ]
rm_data[-1, ] -> rm_data

rm_data %>% 
  janitor::clean_names() %>% 
  readr::type_convert()-> rm_data

str(rm_data)

colnames(rm_data)[7]  <- "class"
colnames(rm_data)[23] <- "quality_hold_in_cost"
colnames(rm_data)[26] <- "on_hand_in_cost"
colnames(rm_data)[28] <- "target_inv_in_cost"
colnames(rm_data)[30] <- "max_inv_cost"
colnames(rm_data)[35] <- "at_risk_in_cost"
colnames(rm_data)[44] <- "on_hand_inv_greater_than_max"
colnames(rm_data)[45] <- "on_hand_inv_less_or_equal_than_max"
colnames(rm_data)[46] <- "on_hand_inv_greater_than_target"
colnames(rm_data)[47] <- "on_hand_inv_less_or_equal_than_target"
colnames(rm_data)[48] <- "iqr_cost"
colnames(rm_data)[49] <- "upi_cost"
colnames(rm_data)[50] <- "iqr_cost_plus_hold_cost"
colnames(rm_data)[51] <- "upi_cost_plus_hold_cost"
colnames(rm_data)[54] <- "current_month_dep_demand_in_cost"
colnames(rm_data)[55] <- "next_month_dep_demand_in_cost"



rm_data %>% 
  dplyr::mutate(loc_sku = gsub("-", "_", loc_sku)) %>% 
  dplyr::relocate(loc_sku, .before = supplier_number) -> rm_data



# Inventory Analysis Read RM ----

inventory_analysis_rm <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/BoM version 2/Weekly Run/7.5.2023/Inventory Report for all locations (29).xlsx")



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

inventory_analysis %>% 
  dplyr::mutate(inventory_qty_cases = ifelse(is.na(inventory_qty_cases), 0, inventory_qty_cases)) -> inventory_analysis


# Inventory_analysis_pivot_ref

reshape2::dcast(inventory_analysis, ref ~ inventory_hold_status, value.var = "inventory_qty_cases", sum) -> pivot_ref_inventory_analysis
reshape2::dcast(inventory_analysis, campus_ref ~ inventory_hold_status, value.var = "inventory_qty_cases", sum) -> pivot_campus_ref_inventory_analysis

pivot_campus_ref_inventory_analysis %>% 
  dplyr::rename(usable = Useable, loc_sku = campus_ref, hard_hold = "Hard Hold", soft_hold = "Soft Hold") -> pivot_campus_ref_inventory_analysis

# BoM_dep_demand ----
bom_dep_demand <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/BoM version 2/Weekly Run/7.5.2023/Bill of Material_070523.xlsx",
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
# consumption_data <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/consumption data component - 05.04.23.xlsx")
# 
# consumption_data[-1:-2,] -> consumption_data
# colnames(consumption_data) <- consumption_data[1, ]
# consumption_data[-1, ] -> consumption_data
# 
# 
# colnames(consumption_data)[1] <- "loc_sku"
# colnames(consumption_data)[ncol(consumption_data)-1] <- "sum_12mos"
# colnames(consumption_data)[ncol(consumption_data)] <- "sum_6mos"
# 
# consumption_data %>% 
#   dplyr::mutate(loc_sku = gsub("-", "_", loc_sku)) -> consumption_data
# 
# consumption_data %>% 
#   data.frame() %>% 
#   readr::type_convert() -> consumption_data
# 
# consumption_data[is.na(consumption_data)] <- 0



# Consumption data component # Updated once a month ----
consumption_data <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/Raw Material Monthly Consumption - 06.09.23.xlsx")
consumption_data[-1:-3, ] -> consumption_data
colnames(consumption_data) <- consumption_data[1, ]
consumption_data[-1, ] -> consumption_data

consumption_data -> consumption_data

consumption_data %>% 
  janitor::clean_names() %>% 
  data.frame() %>%
  readr::type_convert() %>% 
  dplyr::select(-na_6) %>% 
  dplyr::rename(loc_sku = ref,
                unit.cost = unit_cost,
                sum_12mos = na_7) %>% 
  dplyr::mutate(sum_6mos = monthly_usage_8 + monthly_usage_9 + monthly_usage_10 + monthly_usage_11 + monthly_usage_12 + monthly_usage_13) %>% 
  dplyr::relocate(sum_6mos, .before = sum_12mos) %>% 
  dplyr::mutate(loc_sku = gsub("-", "_", loc_sku)) -> consumption_data


consumption_data[is.na(consumption_data)] <- 0



# SS Optimization RM for EOQ ----
ss_optimization <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/SS Optimization by Location - Raw Material June 2023.xlsx",
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

ss_optimization[!duplicated(ss_optimization[,c("loc_sku")]),] -> ss_optimization

# Custord PO ----
po <- read.csv("Z:/IMPORT_JDE_OPENPO.csv",
               header = FALSE)



po %>% 
  dplyr::rename(aa = V1) %>% 
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
  dplyr::rename(loc_sku = ref) -> po_pivot



# Custord Receipt ----
receipt <- read.csv("Z:/IMPORT_RECEIPTS.csv",
                    header = FALSE)


# Base receipt variable
receipt %>% 
  dplyr::rename(aa = V1) %>% 
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
  dplyr::mutate(supplier_number = replace(supplier_number, is.na(supplier_number), "DNRR")) %>% 
  dplyr::relocate(supplier_number, .after = loc_sku) -> rm_data

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
  dplyr::mutate(safety_stock.y = replace(safety_stock.y, is.na(safety_stock.y), 0)) %>% 
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
  dplyr::mutate(hard_hold = replace(hard_hold, is.na(hard_hold), 0)) %>% 
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
  dplyr::mutate(on_hand_usable_soft_hold = usable + soft_hold) -> rm_data

# Calculation - On Hand in $$
rm_data %>% 
  dplyr::mutate(on_hand_in_cost = on_hand_usable_soft_hold * standard_cost) %>% 
  dplyr::mutate(on_hand_in_cost = round(on_hand_in_cost, 2)) %>% 
  dplyr::mutate(on_hand_in_cost = replace(on_hand_in_cost, is.na(on_hand_in_cost), 0)) -> rm_data



# vlookup - OPV
exception_report %>% 
  dplyr::arrange(loc_sku, desc(order_policy_value)) -> exception_report_opv

exception_report_opv[!duplicated(exception_report_opv[,c("loc_sku")]),] -> exception_report_opv

merge(rm_data, exception_report_opv[, c("loc_sku", "order_policy_value")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::mutate(order_policy_value = replace(order_policy_value, is.na(order_policy_value), 0)) %>% 
  dplyr::relocate(order_policy_value, .after = opv) %>% 
  dplyr::select(-opv) %>% 
  dplyr::rename(opv = order_policy_value) -> rm_data

rm_data[!duplicated(rm_data[,c("loc_sku")]),] -> rm_data



# vlookup - PO in next 28 days
merge(rm_data, po_pivot[, c("loc_sku", "Y")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::mutate(Y = round(Y, 2)) %>% 
  dplyr::mutate(Y = replace(Y, is.na(Y), 0)) %>% 
  dplyr::relocate(Y, .after = po_in_next_30_days) %>% 
  dplyr::select(-po_in_next_30_days) %>% 
  dplyr::rename(po_in_next_30_days = Y) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data



# vlookup - Receipt in next 28 days
merge(rm_data, receipt_pivot[, c("loc_sku", "Y")], by = "loc_sku", all.x = TRUE) %>%
  dplyr::mutate(Y = round(Y, 2)) %>% 
  dplyr::mutate(Y = replace(Y, is.na(Y), 0)) %>% 
  dplyr::relocate(Y, .after = receipt_in_the_next_30_days) %>% 
  dplyr::select(-receipt_in_the_next_30_days) %>% 
  dplyr::rename(receipt_in_the_next_30_days = Y) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data


# vlookup - Current month dep demand
merge(rm_data, bom_dep_demand[, c("loc_sku", "current_month")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::mutate(current_month = replace(current_month, is.na(current_month), 0)) %>% 
  dplyr::relocate(current_month, .after = current_month_dep_demand) %>% 
  dplyr::select(-current_month_dep_demand) %>% 
  dplyr::rename(current_month_dep_demand = current_month) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data



# vlookup - Next month dep demand
merge(rm_data, bom_dep_demand[, c("loc_sku", "next_month")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::mutate(next_month = replace(next_month, is.na(next_month), 0)) %>% 
  dplyr::relocate(next_month, .after = next_month_dep_demand) %>% 
  dplyr::select(-next_month_dep_demand) %>% 
  dplyr::rename(next_month_dep_demand = next_month) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data


# vlookup - Total dep. demand Next 6 Months
merge(rm_data, bom_dep_demand[, c("loc_sku", "sum_of_months")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::mutate(sum_of_months = replace(sum_of_months, is.na(sum_of_months), 0)) %>% 
  dplyr::relocate(sum_of_months, .after = total_dep_demand_next_6_months) %>% 
  dplyr::select(-total_dep_demand_next_6_months) %>% 
  dplyr::rename(total_dep_demand_next_6_months = sum_of_months) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data


# Calculation - DOS
rm_data %>% 
  dplyr::mutate(dos = on_hand_usable_soft_hold / (pmax(current_month_dep_demand, next_month_dep_demand)/30)) %>% 
  dplyr::mutate(dos = round(dos, 0)) %>% 
  dplyr::mutate(dos = replace(dos, is.na(dos), 0)) %>% 
  dplyr::mutate(dos = replace(dos, is.nan(dos), 0)) %>% 
  dplyr::mutate(dos = replace(dos, is.infinite(dos), 0)) -> rm_data


# vlookup - Total Last 6 mos Sales
merge(rm_data, consumption_data[, c("loc_sku", "sum_6mos")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::mutate(sum_6mos = as.double(sum_6mos)) %>% 
  dplyr::mutate(sum_6mos = round(sum_6mos, 2)) %>% 
  dplyr::mutate(sum_6mos = replace(sum_6mos, is.na(sum_6mos), 0)) %>% 
  dplyr::relocate(sum_6mos, .after = total_last_6_mos_sales) %>% 
  dplyr::select(-total_last_6_mos_sales) %>% 
  dplyr::rename(total_last_6_mos_sales = sum_6mos) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data

# vlookup - Total Last 12 mos Sales
merge(rm_data, consumption_data[, c("loc_sku", "sum_12mos")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::mutate(sum_12mos = as.double(sum_12mos)) %>% 
  dplyr::mutate(sum_12mos = round(sum_12mos, 2)) %>% 
  dplyr::mutate(sum_12mos = replace(sum_12mos, is.na(sum_12mos), 0)) %>% 
  dplyr::relocate(sum_12mos, .after = total_last_12_mos_sales) %>% 
  dplyr::select(-total_last_12_mos_sales) %>% 
  dplyr::rename(total_last_12_mos_sales = sum_12mos) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data

# vlookup - EOQ
merge(rm_data, ss_optimization[, c("loc_sku", "eoq_adjusted")], by = "loc_sku", all.x = TRUE) %>% 
  dplyr::mutate(eoq_adjusted = as.double(eoq_adjusted)) %>% 
  dplyr::mutate(eoq_adjusted = round(eoq_adjusted, 0)) %>% 
  dplyr::mutate(eoq_adjusted = replace(eoq_adjusted, is.na(eoq_adjusted), 0)) %>% 
  dplyr::relocate(eoq_adjusted, .after = eoq) %>% 
  dplyr::select(-eoq) %>% 
  dplyr::rename(eoq = eoq_adjusted) %>% 
  dplyr::relocate(loc_sku, .after = item) -> rm_data




# Calculation - Moq in days
rm_data %>% 
  dplyr::mutate(moq_in_days = ifelse(lead_time == "DNRR", "DNRR", moq/(total_dep_demand_next_6_months/180)),
                moq_in_days = replace(moq_in_days, is.na(moq_in_days), 999),
                moq_in_days = as.numeric(moq_in_days),
                moq_in_days = round(moq_in_days, 1),
                moq_in_days = replace(moq_in_days, is.na(moq_in_days), "DNRR")) -> rm_data


# Calculation - Max Cycle Stock
rm_data %>% 
  dplyr::mutate(max_cycle_stock =
                  pmax(eoq, moq, opv*(next_month_dep_demand/20.83), opv*(total_last_12_mos_sales/250))) %>% 
  dplyr::mutate(max_cycle_stock = round(max_cycle_stock, 1)) %>% 
  dplyr::mutate(max_cycle_stock = replace(max_cycle_stock, is.na(max_cycle_stock), 0)) %>% 
  dplyr::mutate(max_cycle_stock = ifelse(lead_time == "DNRR", eoq, max_cycle_stock)) -> rm_data


# Calculation - Target Inv
rm_data %>% 
  dplyr::mutate(target_inv = safety_stock + max_cycle_stock / 2) -> rm_data

# Calculation - Target Inv in $$
rm_data %>% 
  dplyr::mutate(target_inv_in_cost = target_inv * standard_cost) %>% 
  dplyr::mutate(target_inv_in_cost = as.double(target_inv_in_cost)) %>% 
  dplyr::mutate(target_inv_in_cost = round(target_inv_in_cost, 2)) %>% 
  dplyr::mutate(target_inv_in_cost = replace(target_inv_in_cost, is.na(target_inv_in_cost), 0)) -> rm_data


# Calculation - Max inv
rm_data %>% 
  dplyr::mutate(max_inv = safety_stock + max_cycle_stock) -> rm_data

# Calculation - Max inv $$
rm_data %>% 
  dplyr::mutate(max_inv_cost = max_inv * standard_cost) %>% 
  dplyr::mutate(max_inv_cost = as.double(max_inv_cost)) %>% 
  dplyr::mutate(max_inv_cost = round(max_inv_cost, 2)) %>% 
  dplyr::mutate(max_inv_cost = replace(max_inv_cost, is.na(max_inv_cost), 0)) -> rm_data


# Calculation - MOQ Flag
rm_data %>% 
  dplyr::mutate(moq_flag = ifelse(lead_time == "DNRR", "DNRR",
                                  ifelse(total_dep_demand_next_6_months == 0, "No demand", 
                                         ifelse(moq / (total_dep_demand_next_6_months / 180) >= (shelf_life_day * 0.6), "High MOQ", 
                                                "OK")))) -> rm_data



# Calculation - has Max?
rm_data %>% 
  dplyr::mutate(has_max = ifelse(max_inv > 0, 1, 0)) -> rm_data

# Calculation - on hand Inv > max
rm_data %>% 
  dplyr::mutate(on_hand_inv_greater_than_max = ifelse(on_hand_usable_soft_hold > max_inv, 1, 0)) -> rm_data

# Calculation - on hand Inv <= max
rm_data %>% 
  dplyr::mutate(on_hand_inv_less_or_equal_than_max = ifelse(on_hand_usable_soft_hold <= max_inv, 1, 0)) -> rm_data

# Calculation - on hand Inv > target
rm_data %>% 
  dplyr::mutate(on_hand_inv_greater_than_target = ifelse(on_hand_usable_soft_hold > target_inv, 1, 0)) -> rm_data

# Calculation - on hand Inv <= target
rm_data %>% 
  dplyr::mutate(on_hand_inv_less_or_equal_than_target = ifelse(on_hand_usable_soft_hold <= target_inv, 1, 0)) -> rm_data


# Calculation - Inv Health
rm_data %>% 
  dplyr::mutate(shelf_life_day = as.numeric(shelf_life_day),
                birthday = as.integer(birthday)) -> rm_data


rm_data %>% 
  dplyr::mutate(today = Sys.Date(),
                today = as.Date(today, format = "%Y-%m-%d"),
                birthday = as.Date(birthday, origin = "1899-12-30"),
                diff_days = today - birthday,
                diff_days = as.numeric(diff_days),
                inv_health = ifelse(on_hand_usable_soft_hold < safety_stock, "BELOW SS", (ifelse(item_type == "Non-Commodity" & dos > 0.6 * shelf_life_day, "AT RISK",
                                                                                                 ifelse((on_hand_usable_soft_hold > 0 & lead_time == "DNRR") | (on_hand_usable_soft_hold > 0 & current_month_dep_demand == 0 & next_month_dep_demand == 0 & total_dep_demand_next_6_months == 0 & diff_days > 90), "DEAD",
                                                                                                        ifelse(on_hand_inv_greater_than_max == 0, "HEALTHY", "EXCESS")))))) -> rm_data





# Calculation - At Risk in $$
rm_data %>% 
  dplyr::mutate(at_risk_in_cost = ifelse(inv_health == "At Risk",
                                         (on_hand_usable_soft_hold -((pmax(current_month_dep_demand, next_month_dep_demand)/30) 
                                                                     *(shelf_life_day * 0.6))) * standard_cost,0)) -> rm_data

# Calculation - IQR $$
rm_data %>% 
  dplyr::mutate(iqr_cost = ifelse(inv_health == "DEAD" | inv_health == "HEALTHY" | inv_health == "BELOW SS", on_hand_in_cost, 
                                  ifelse(inv_health == "AT RISK", at_risk_in_cost, on_hand_in_cost - max_inv_cost))) -> rm_data


# Calculation - UPI $$
rm_data %>% 
  dplyr::mutate(upi_cost = ifelse(inv_health == "AT RISK", at_risk_in_cost,
                                  ifelse(inv_health == "EXCESS", on_hand_in_cost - max_inv_cost,
                                         ifelse(inv_health == "DEAD", on_hand_in_cost, 0)))) -> rm_data
# Calculation - IQR $$ + Hold $$
rm_data %>% 
  dplyr::mutate(iqr_cost_plus_hold_cost = iqr_cost + quality_hold_in_cost) -> rm_data 

# Calculation - UPI $$ + Hold $$
rm_data %>% 
  dplyr::mutate(upi_cost_plus_hold_cost = upi_cost + quality_hold_in_cost) -> rm_data


# Calculation - current month dep demand in $$
rm_data %>% 
  dplyr::mutate(current_month_dep_demand_in_cost = current_month_dep_demand * standard_cost,
                current_month_dep_demand_in_cost = round(current_month_dep_demand_in_cost, 2)) -> rm_data



# Calculation - next month dep demand in $$
rm_data %>% 
  dplyr::mutate(next_month_dep_demand_in_cost = next_month_dep_demand * standard_cost,
                next_month_dep_demand_in_cost = round(next_month_dep_demand_in_cost, 2)) -> rm_data



######## Deleting items that we don't need ###########
rm_data %>% 
  dplyr::filter(loc_sku != "60_8883") %>% 
  dplyr::filter(loc_sku != "75_16795") %>% 
  dplyr::filter(loc_sku != "75_21645") -> rm_data


##### update 5/31/2023 #####

# Planner Name N/A to 0
rm_data %>% 
  dplyr::mutate(planner_name = ifelse(is.na(planner_name) & planner == 0, 0, planner_name)) -> rm_data

# MOQ N/A to 0
rm_data %>% 
  dplyr::mutate(moq  = ifelse(is.na(moq), 0, moq)) -> rm_data

# Usable, hold, label work ----
inv_bal <- read_csv("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Weekly Report run/7.5.23/ATT86890.csv")

inv_bal[-1:-2, ] -> inv_bal
colnames(inv_bal) <- inv_bal[1, ]
inv_bal[-1, ] -> inv_bal

inv_bal %>% 
  janitor::clean_names() %>% 
  dplyr::rename(b_p = bp,
                usable = na,
                hard_hold = na_2) %>% 
  tidyr::separate(b_p, c("a", "b", "c", "d", "e")) %>% 
  dplyr::select(b, item, usable, hard_hold, soft_hold) %>% 
  dplyr::rename(b_p = b) %>% 
  dplyr::mutate(b_p = as.double(b_p),
                soft_hold = as.double(soft_hold),
                hard_hold = as.double(hard_hold),
                usable = as.double(usable),
                usable = replace(usable, is.na(usable), 0),
                hard_hold = replace(hard_hold, is.na(hard_hold), 0),
                soft_hold = replace(soft_hold, is.na(soft_hold), 0)) %>% 
  dplyr::left_join(campus_ref) %>% 
  dplyr::mutate(loc_sku = paste0(campus, "_", item)) %>% 
  dplyr::rename(usable_2 = usable,
                quality_hold_2 = hard_hold,
                soft_hold_2 = soft_hold) %>% 
  dplyr::select(loc_sku, usable_2, quality_hold_2, soft_hold_2) -> inv_bal

inv_bal[!duplicated(inv_bal[,c("loc_sku")]),] -> inv_bal


rm_data %>% 
  dplyr::mutate(label_ref = paste0(mfg_loc, "_", item_type)) %>% 
  dplyr::left_join(inv_bal) %>% 
  dplyr::mutate(usable = ifelse(label_ref == "25_Label" | label_ref == "55_Label" | label_ref == "86_Label", usable_2, usable),
                quality_hold = ifelse(label_ref == "25_Label" | label_ref == "55_Label" | label_ref == "86_Label", quality_hold_2, quality_hold),
                soft_hold = ifelse(label_ref == "25_Label" | label_ref == "55_Label" | label_ref == "86_Label", soft_hold_2, soft_hold)) %>% 
  dplyr::select(-usable_2, -quality_hold_2, -soft_hold_2, -label_ref) %>% 
  dplyr::mutate(usable = replace(usable, is.na(usable), 0),
                hard_hold = replace(quality_hold, is.na(quality_hold), 0),
                soft_hold = replace(soft_hold, is.na(soft_hold), 0)) -> rm_data


# Quality/soft Hold blank to 0
rm_data %>% 
  dplyr::mutate(quality_hold = ifelse(is.na(quality_hold), 0, quality_hold),
                quality_hold = ifelse(quality_hold < 0, 0, quality_hold),
                soft_hold = ifelse(is.na(soft_hold), 0, soft_hold),
                soft_hold = ifelse(soft_hold < 0, 0, soft_hold)) -> rm_data


# Arrange ----
rm_data_for_arrange <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Weekly Report run/7.5.23/Raw Material Inventory Health (IQR) - 07.05.23.xlsx",
                                  sheet = "RM data")

rm_data_for_arrange[-1:-2, ] -> rm_data_for_arrange
colnames(rm_data_for_arrange) <- rm_data_for_arrange[1, ]
rm_data_for_arrange[-1, ] -> rm_data_for_arrange

rm_data_for_arrange %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::select(loc_sku) %>% 
  dplyr::mutate(arrange = row_number(),
                loc_sku = gsub("-", "_", loc_sku)) -> rm_data_for_arrange

rm_data %>% 
  dplyr::left_join(rm_data_for_arrange) %>% 
  dplyr::arrange(arrange) %>% 
  dplyr::select(-arrange)-> rm_data





#####################################################################################################################
########################################## Change Col names to original #############################################
#####################################################################################################################
rm_data %>% select(next_month_dep_demand_in_cost)

########### Don't forget to rearrange and bring cols only what you need! #################
rm_data %>% 
  dplyr::mutate(loc_sku = gsub("_", "-", loc_sku)) %>% 
  dplyr::select(mfg_loc, loc_name, item, loc_sku, supplier_number, description, class, item_type, shelf_life_day,
                birthday, uom, lead_time, planner, planner_name, standard_cost, moq, moq_in_days, eoq, safety_stock, max_cycle_stock, usable, quality_hold,
                quality_hold_in_cost, soft_hold, on_hand_usable_soft_hold, on_hand_in_cost, target_inv, target_inv_in_cost, max_inv, max_inv_cost,
                opv, po_in_next_30_days, receipt_in_the_next_30_days, dos, at_risk_in_cost, moq_flag, inv_health,  
                current_month_dep_demand, next_month_dep_demand,
                total_dep_demand_next_6_months, total_last_6_mos_sales, total_last_12_mos_sales, has_max, on_hand_inv_greater_than_max,
                on_hand_inv_less_or_equal_than_max, on_hand_inv_greater_than_target, on_hand_inv_less_or_equal_than_target,
                iqr_cost, upi_cost, iqr_cost_plus_hold_cost, upi_cost_plus_hold_cost, current_month_dep_demand_in_cost, 
                next_month_dep_demand_in_cost) -> rm_data


colnames(rm_data)[1]<-"Mfg Loc"
colnames(rm_data)[2]<-"Loc Name"
colnames(rm_data)[3]<-"Item"
colnames(rm_data)[4]<-"Loc-SKU"
colnames(rm_data)[5]<-"Supplier#"
colnames(rm_data)[6]<-"Description"
colnames(rm_data)[7]<-"Class"
colnames(rm_data)[8]<-"Item Type"
colnames(rm_data)[9]<-"Shelf Life (day)"
colnames(rm_data)[10]<-"Birthday"
colnames(rm_data)[11]<-"UoM"
colnames(rm_data)[12]<-"Lead time"
colnames(rm_data)[13]<-"Planner"
colnames(rm_data)[14]<-"Planner Name"
colnames(rm_data)[15]<-"Standard Cost"
colnames(rm_data)[16]<-"MOQ"
colnames(rm_data)[17]<-"MOQ in days"
colnames(rm_data)[18]<-"EOQ"
colnames(rm_data)[19]<-"Safety Stock"
colnames(rm_data)[20]<-"Max Cycle Stock"
colnames(rm_data)[21]<-"Usable"
colnames(rm_data)[22]<-"Quality hold"
colnames(rm_data)[23]<-"Quality hold in $$"
colnames(rm_data)[24]<-"Soft Hold"
colnames(rm_data)[25]<-"On Hand(usable + soft hold)"
colnames(rm_data)[26]<-"On Hand in $$"
colnames(rm_data)[27]<-"Target Inv"
colnames(rm_data)[28]<-"Target Inv in $$"
colnames(rm_data)[29]<-"Max inv"
colnames(rm_data)[30]<-"Max inv $$"
colnames(rm_data)[31]<-"OPV"
colnames(rm_data)[32]<-"PO in next 30 days"
colnames(rm_data)[33]<-"Receipt in the next 30 days"
colnames(rm_data)[34]<-"DOS"
colnames(rm_data)[35]<-"At Risk in $$"
colnames(rm_data)[36]<-"MOQ Flag"
colnames(rm_data)[37]<-"Inv Health"
colnames(rm_data)[38]<-"Current month dep demand"
colnames(rm_data)[39]<-"Next month dep demand"
colnames(rm_data)[40]<-"Total dep. demand Next 6 Months"
colnames(rm_data)[41]<-"Total Last 6 mos Sales"
colnames(rm_data)[42]<-"Total Last 12 mos Sales"
colnames(rm_data)[43]<-"has Max?"
colnames(rm_data)[44]<-"on hand Inv >max"
colnames(rm_data)[45]<-"on hand Inv <= max"
colnames(rm_data)[46]<-"on hand Inv > target"
colnames(rm_data)[47]<-"on hand Inv <= target"
colnames(rm_data)[48]<-"IQR $$"
colnames(rm_data)[49]<-"UPI$$"
colnames(rm_data)[50]<-"IQR $$ + Hold $$"
colnames(rm_data)[51]<-"UPI$$ + Hold $$"
colnames(rm_data)[52]<-"current month dep demand in $$"
colnames(rm_data)[53]<-"next month dep demand in $$"

writexl::write_xlsx(rm_data, "IQR_RM_Report_070523.xlsx")


file.rename(from="C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/IQR/venturafoods_RPA_IQR/IQR_RM_Report_070523.xlsx",
            to="C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Weekly Report run/7.5.23/IQR_RM_Report_070523.xlsx")






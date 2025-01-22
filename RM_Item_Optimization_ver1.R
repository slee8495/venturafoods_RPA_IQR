library(tidyverse)
library(readxl)
library(writexl)
library(reshape2)
library(officer)
library(openxlsx)
library(lubridate)
library(magrittr)
library(skimr)
library(bizdays)
library(janitor) 

##################################################################################################################################################################
##################################################################################################################################################################
##################################################################################################################################################################

exception_report <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/JDE Exception report extract/2025/exception report 2025.01.21.xlsx")
exception_report_dnrr <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/JDE DNRR Exception report extract/2025/exception report DOU 2025.01.21.xlsx")
inventory_fg <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Inventory/Inventory with Lot Report v.2 - 2025.01.21.xlsx",
                           sheet = "FG")
inventory_rm <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Inventory/Inventory with Lot Report v.2 - 2025.01.21.xlsx",
                           sheet = "RM")
oo_bt_fg <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/01212025/US and CAN OO BT where status _ J.xlsx")
dsx <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2025/DSX Forecast Backup - 2025.01.16.xlsx")
jde_25_55_label <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Inventory/JDE Inventory Lot Detail - 2025.01.21.xlsx")
lot_status_code <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Lot Status Code.xlsx")
bom <- read.xlsx("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/01212025/Bill of Material_01212025.xlsx")
campus_ref <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Campus reference.xlsx")
iom_live <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/SS Optimization by Location - Finished Goods LIVE.xlsx",
                       sheet = "CVM & Focus label & Contract")
iom_live_1st_sheet <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/SS Optimization by Location - Finished Goods LIVE.xlsx")
iom_live_1st_sheet_rm <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/SS Optimization by Location - Raw Material LIVE.xlsx")


complete_sku_list <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/01212025/Complete SKU list - Linda.xlsx")
unit_cost <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/01212025/Unit_Cost.xlsx")
class_ref <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Class reference (JDE).xlsx")
iqr_rm <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Raw Material Inventory Health (IQR) NEW TEMPLATE - 01.14.2025.xlsx",
                     sheet = "RM data") ### Use Pre week ###

supplier_address_book <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Address Book/Address Book - 2025.01.07.xlsx",
                                    sheet = "supplier")

###################################################################

exception_report[-1:-2, ] -> exception_report
colnames(exception_report) <- exception_report[1, ]
exception_report[-1, -32] -> exception_report

###################################################################

exception_report_dnrr[-1:-2, ] -> exception_report_dnrr
colnames(exception_report_dnrr) <- exception_report_dnrr[1, ]
exception_report_dnrr[-1, -32] -> exception_report_dnrr


###################################################################

complete_sku_list[-1, ] -> complete_sku_list
colnames(complete_sku_list) <- complete_sku_list[1, ]
complete_sku_list[-1, ] -> complete_sku_list

#################################################### Finished Goods ############################################################################################    

colnames(iom_live) <- iom_live[1, ]
iom_live[-1, ] -> iom_live

###################################################################

inventory_fg[-1, ] -> inventory_fg
colnames(inventory_fg) <- inventory_fg[1, ]
inventory_fg[-1, ] -> inventory_fg


###################################################################

dsx[-1,] -> dsx
colnames(dsx) <- dsx[1, ]
dsx[-1, ] -> dsx

###################################################################

oo_bt_fg %>% 
  dplyr::slice(c(-1, -3)) -> oo_bt_fg_2

colnames(oo_bt_fg_2) <- oo_bt_fg_2[1, ]
oo_bt_fg_2[-1, ] -> oo_bt_fg_2

###################################################################

unit_cost[-1, ] -> unit_cost
colnames(unit_cost) <- unit_cost[1, ]
unit_cost[-1, ] -> unit_cost

###################################################################

iom_live_1st_sheet[-1:-6, ] -> iom_live_1st_sheet
colnames(iom_live_1st_sheet) <- iom_live_1st_sheet[1, ]
iom_live_1st_sheet[-1, ] -> iom_live_1st_sheet

###################################################################

iom_live_1st_sheet_rm[-1:-5, ] -> iom_live_1st_sheet_rm
colnames(iom_live_1st_sheet_rm) <- iom_live_1st_sheet_rm[1, ]
iom_live_1st_sheet_rm[-1, ] -> iom_live_1st_sheet_rm

###################################################################

iqr_rm[-1:-2, ] -> iqr_rm
colnames(iqr_rm) <- iqr_rm[1, ]
iqr_rm[-1, ] -> iqr_rm

###################################################################


inventory_rm[-1, ] -> inventory_rm
colnames(inventory_rm) <- inventory_rm[1, ]
inventory_rm[-1, ] -> inventory_rm

###################################################################


jde_25_55_label[-1:-5, ] -> jde_25_55_label
colnames(jde_25_55_label) <- jde_25_55_label[1, ]
jde_25_55_label[-1, ] -> jde_25_55_label


###################################################################


# 1. Has inventory (useable, soft hold, hard hold all included)


inventory_rm %>% 
  janitor::clean_names() %>% 
  dplyr::select(campus_no, item, current_inventory_balance) %>% 
  dplyr::rename(inventory = current_inventory_balance) %>% 
  dplyr::mutate(inventory = as.double(inventory)) %>% 
  dplyr::mutate(ref = paste0(campus_no, "_", item)) %>% 
  dplyr::group_by(ref) %>% 
  dplyr::summarise(inventory = sum(inventory)) %>% 
  tidyr::separate(ref, into = c("campus", "item"), sep = "_") %>%
  dplyr::mutate(campus = as.double(campus),
                item = as.double(item)) %>%
  dplyr::mutate(ref = paste0(campus, "_", item)) %>%
  dplyr::filter(inventory > 0) %>% 
  dplyr::select(-inventory) -> has_on_hand_inventory_rm_1



jde_25_55_label %>% 
  janitor::clean_names() %>% 
  dplyr::filter(mpf == "LBL") %>% 
  dplyr::mutate(on_hand = as.double(on_hand)) %>% 
  dplyr::filter(on_hand > 0) %>%
  dplyr::select(bp, item_number) %>% 
  dplyr::left_join(campus_ref %>% janitor::clean_names() %>% select(location, campus) %>% rename(bp = location)) %>% 
  dplyr::mutate(ref = paste0(campus, "_", item_number)) %>% 
  dplyr::distinct(ref) %>% 
  tidyr::separate(ref, into = c("campus", "item"), sep = "_") %>%
  dplyr::mutate(campus = as.double(campus),
                item = as.double(item)) %>% 
  dplyr::mutate(ref = paste0(campus, "_", item)) -> has_on_hand_inventory_rm_2


bind_rows(has_on_hand_inventory_rm_1, has_on_hand_inventory_rm_2) -> has_on_hand_inventory_rm

has_on_hand_inventory_rm %>% 
  dplyr::mutate(campus = as.character(campus),
                item = as.character(item)) -> has_on_hand_inventory_rm


# 2. Zero inventory but has dependent demand for next 6 months

bom %>% 
  janitor::clean_names() %>% 
  dplyr::select(comp_ref, mon_a_dep_demand, mon_b_dep_demand, mon_c_dep_demand, mon_d_dep_demand, mon_e_dep_demand, mon_f_dep_demand) %>% 
  dplyr::mutate(dep_demand = mon_a_dep_demand + mon_b_dep_demand + mon_c_dep_demand + mon_d_dep_demand + mon_e_dep_demand + mon_f_dep_demand) %>% 
  dplyr::filter(dep_demand > 0) %>% 
  dplyr::select(comp_ref) %>% 
  dplyr::distinct(comp_ref) %>% 
  tidyr::separate(comp_ref, into = c("campus", "item"), sep = "-") %>% 
  plyr::mutate(ref = paste0(campus, "_", item)) -> zero_on_hand_has_dependent_demand_rm




#  3. Zero inventory, no dependent demand for next 6 months but show ACTIVE in JDE
exception_report %>% 
  janitor::clean_names() %>%
  dplyr::filter(mpf_or_line == "LBL" | mpf_or_line == "PKG" | mpf_or_line == "ING") %>% 
  dplyr::select(b_p, item_number) %>% 
  dplyr::left_join(campus_ref %>% janitor::clean_names() %>% select(location, campus) %>% rename(b_p = location)) %>% 
  dplyr::mutate(ref = paste0(campus, "_", item_number)) %>% 
  dplyr::select(campus, item_number, ref) %>% 
  dplyr::rename(item = item_number) %>% 
  dplyr::distinct(ref, .keep_all = TRUE) %>%
  dplyr::filter(stringr::str_detect(item, "^[0-9]+$")) %>% 
  dplyr::distinct(ref, .keep_all = TRUE) -> active_items_rm


active_items_rm %>% 
  dplyr::mutate(campus = as.character(campus),
                item = as.character(item)) -> active_items_rm


dplyr::bind_rows(has_on_hand_inventory_rm, 
                 zero_on_hand_has_dependent_demand_rm, 
                 active_items_rm) %>%
  dplyr::distinct(ref) %>% 
  tidyr::separate(ref, c("location", "item")) %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) -> final_data_rm



## Conclusion
dplyr::bind_rows(has_on_hand_inventory_rm, 
                 zero_on_hand_has_dependent_demand_rm, 
                 active_items_rm) %>%
  dplyr::distinct(ref) %>% 
  tidyr::separate(ref, c("location", "item")) %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) -> final_data_rm


# Final touch
final_data_rm %>% 
  dplyr::filter(!(location %in% c("16", "22", "502", "503", "690", "691", "214", "331", "601", "602", "608", "621", "636", "660", "675"))) %>% 
  dplyr::filter(!(ref %in% c("60_8883", "75_16795", "75_21645"))) %>% 
  dplyr::filter(!(item %in% c("1", "34688"))) %>% 
  dplyr::filter(!(stringr::str_detect(item, "^[0-9]{3}$"))) %>% 
  dplyr::filter(!(location %in% c("622", "624") & stringr::str_detect(item, "^[0-9]{5}$"))) -> final_data_rm




##################################################################################################################################################################
##################################################################################################################################################################
##################################################################################################################################################################


# mfg_loc
final_data_rm %>% 
  dplyr::left_join(campus_ref %>% 
                     janitor::clean_names() %>% 
                     dplyr::rename(mfg_loc = campus), by = "location") %>% 
  dplyr::select(-location, -ref) %>% 
  dplyr::mutate(loc_sku = paste0(mfg_loc, "_", item)) %>% 
  dplyr::filter(mfg_loc %in% c(10, 25, 30, 33, 34, 36, 43, 55, 60, 75, 86, 622, 624)) %>% 
  dplyr::relocate(mfg_loc, location_name, item, loc_sku) -> final_data_rm


# Supplier#, Supplier Name
final_data_rm %>% 
  dplyr::left_join(exception_report %>% 
                     janitor::clean_names() %>% 
                     dplyr::mutate(loc_sku = paste0(b_p, "_", item_number)) %>% 
                     dplyr::select(loc_sku, supplier), by = "loc_sku") %>%
  
  dplyr::left_join(exception_report_dnrr %>%
                     janitor::clean_names() %>% 
                     dplyr::mutate(loc_sku = paste0(b_p, "_", item_number)) %>%
                     dplyr::select(loc_sku, supplier), by = "loc_sku") %>% 
  
  dplyr::mutate(supplier = dplyr::coalesce(supplier.x, supplier.y)) %>% 
  dplyr::select(-supplier.x, -supplier.y) %>% 
  
  dplyr::mutate(supplier = ifelse(is.na(supplier), 0, supplier)) %>% 
  
  dplyr::left_join(supplier_address_book %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(address_number, alpha_name) %>% 
                     dplyr::rename(supplier = address_number,
                                   supplier_name = alpha_name) %>% 
                     dplyr::mutate(supplier = as.character(supplier)), by = "supplier") %>% 
  dplyr::mutate(supplier_name = ifelse(is.na(supplier_name), 0, supplier_name)) -> final_data_rm
  


# Description
final_data_rm %>% 
  dplyr::left_join(rbind(exception_report, exception_report_dnrr) %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(item_number, description) %>% 
                     dplyr::distinct(item_number, .keep_all = TRUE) %>% 
                     dplyr::select(item_number, description) %>% 
                     dplyr::rename(item = item_number), by = "item") %>% 
  dplyr::filter(!is.na(description) & description != "") %>% 
  dplyr::filter(!stringr::str_detect(description, "(?i)Condensate|Water|TEMPERATURE RECORDERS|Pallet")) -> final_data_rm


# Remove RPS from exception report
final_data_rm %>% 
  dplyr::left_join(exception_report %>% 
                     janitor::clean_names() %>% 
                     dplyr::mutate(loc_sku = paste0(b_p, "_", item_number)) %>%
                     dplyr::select(mpf_or_line, loc_sku) %>%
                     dplyr::distinct(loc_sku, .keep_all = TRUE), by = "loc_sku") %>% 
  dplyr::filter(mpf_or_line != "RPS") %>% 
  dplyr::select(-mpf_or_line) -> final_data_rm


# Class

final_data_rm %>% 
  dplyr::left_join(bom %>% janitor::clean_names() %>% 
                     dplyr::rename(item = component) %>% 
                     dplyr::select(item, commodity_class) %>% 
                     dplyr::mutate(item = as.character(item)) %>% 
                     dplyr::distinct(item, .keep_all = TRUE), by = "item") %>% 
  dplyr::mutate(commodity_class = as.character(commodity_class)) %>% 

  dplyr::left_join(class_ref %>% 
                     janitor::clean_names() %>% 
                     dplyr::rename(commodity_class = code, class = description) %>%
                     dplyr::mutate(commodity_class = as.character(commodity_class)) %>% 
                     dplyr::distinct(class, .keep_all = TRUE) %>% 
                     dplyr::filter(!is.na(commodity_class)), by = "commodity_class") %>% 
  
  dplyr::left_join(iom_live_1st_sheet_rm %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(loc_sku, class) %>% 
                     dplyr::mutate(loc_sku = gsub("-", "_", loc_sku)) %>% 
                     dplyr::distinct(loc_sku, .keep_all = TRUE), by = "loc_sku") %>% 
  dplyr::mutate(class = dplyr::coalesce(class.x, class.y)) %>% 
  dplyr::select(-class.x, -class.y) %>% 
  
  dplyr::left_join(iqr_rm %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(item, class) %>% 
                     dplyr::distinct(item, .keep_all = TRUE), by = "item") %>% 
  dplyr::mutate(class = dplyr::coalesce(class.x, class.y)) %>% 
  dplyr::select(-class.x, -class.y) %>% 
  dplyr::filter(!stringr::str_detect(class, "(?i)REFINERY PROCESS SUPPLIES")) %>% 
  dplyr::mutate(class = ifelse(is.na(class), 0, class)) -> final_data_rm


# Item Type
final_data_rm %>% 
  dplyr::left_join(exception_report %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(item_number, mpf_or_line) %>% 
                     dplyr::rename(item = item_number) %>% 
                     dplyr::distinct(item, .keep_all = TRUE), by = "item") %>% 
  dplyr::mutate(item_type = dplyr::case_when(
    mpf_or_line == "PKG" ~ "Packaging",
    mpf_or_line == "LBL" ~ "Label",
    mpf_or_line == "ING" ~ "Non-Commodity",
    commodity_class < 500 ~ "Non-Commodity",
    commodity_class == 570 ~ "Label",
    commodity_class >= 500 & commodity_class < 900 & commodity_class != 570 ~ "Packaging",
    commodity_class > 900 ~ "Commodity Oil",
    item %in% c("BCH", "BLD", "FGT", "RPS", "SFM", "SSA", "WIP") ~ "WIP",
    item == "OHD" ~ "Overhead",
    TRUE ~ "Other" 
  )) %>% 
  dplyr::select(-mpf_or_line, -commodity_class)  -> final_data_rm





# Shelf Life (day), Birthday
final_data_rm %>% 
  dplyr::left_join(iom_live_1st_sheet_rm %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(loc_sku, shelf_life_day, birthday) %>% 
                     dplyr::mutate(
                       birthday = suppressWarnings(as.numeric(birthday)), 
                       birthday = as.Date(birthday, origin = "1899-12-30"), 
                       loc_sku = gsub("-", "_", loc_sku)
                     ), 
                   by = "loc_sku") %>% 
  dplyr::mutate(
    shelf_life_day = ifelse(is.na(shelf_life_day), 0, shelf_life_day)
  ) -> final_data_rm






# UOM, Lead Time, Planner, Planner Name
final_data_rm %>% 
  dplyr::mutate(uom = "IQR Report",
                lead_time = "IQR Report",
                planner = "IQR Report",
                planner_name = "IQR Report") -> final_data_rm


# Standard Cost
final_data_rm %>% 
  dplyr::left_join(unit_cost %>% 
                     janitor::clean_names() %>% 
                     dplyr::mutate(location = as.double(location)) %>% 
                     dplyr::mutate(ref = paste0(location, "_", item)) %>% 
                     dplyr::rename(unit_cost = simulated_cost,
                                   loc_sku = ref) %>% 
                     dplyr::select(loc_sku, unit_cost), by = "loc_sku") %>% 
  dplyr::mutate(unit_cost = ifelse(is.na(unit_cost), NA, as.numeric(unit_cost))) %>% 
  dplyr::left_join(iom_live_1st_sheet %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(ship_ref, unit_cost) %>% 
                     dplyr::mutate(ship_ref = gsub("-", "_", ship_ref)) %>% 
                     dplyr::mutate(unit_cost = as.double(unit_cost)) %>% 
                     dplyr::rename(loc_sku = ship_ref,
                                   unit_cost_2 = unit_cost), by = "loc_sku") %>% 
  dplyr::mutate(unit_cost = ifelse(is.na(unit_cost), unit_cost_2, unit_cost),
                unit_cost = ifelse(is.na(unit_cost), 0, unit_cost)) %>% 
  dplyr::select(-unit_cost_2) %>% 
  dplyr::rename(standard_cost = unit_cost) -> final_data_rm


# MOQ, MOQ in days
final_data_rm %>% 
  dplyr::mutate(moq = "IQR Report",
                moq_in_days = "IQR Report") -> final_data_rm


# EOQ            ######################### Question ###################
final_data_rm %>% 
  dplyr::mutate(eoq = "IQR Report") -> final_data_rm


# Safety Stock
final_data_rm %>% 
  dplyr::mutate(safety_stock = "IQR Report") -> final_data_rm


# Safety Stock $
final_data_rm %>% 
  dplyr::mutate(safety_stock_dollar = "Formula") -> final_data_rm

# Max Cycle Stock
final_data_rm %>% 
  dplyr::mutate(max_cycle_stock = "Formula") -> final_data_rm


# Useable, Quality Hold
final_data_rm %>% 
  dplyr::mutate(useable = "IQR Report",
                quality_hold = "IQR Report") -> final_data_rm


# Quality Hold $
final_data_rm %>% 
  dplyr::mutate(quality_hold_dollar = "Formula") -> final_data_rm


# Soft Hold
final_data_rm %>% 
  dplyr::mutate(soft_hold = "IQR Report") -> final_data_rm

# On Hand (usable + soft hold),	On Hand $,	Total Inventory $,	Target Inv,	Target Inv in $$,	Max inv,	Max inv $$
final_data_rm %>% 
  dplyr::mutate(on_hand = "Formula",
                on_hand_dollar = "Formula",
                total_inventory_dollar = "Formula",
                target_inv = "Formula",
                target_inv_dollar = "Formula",
                max_inv = "Formula",
                max_inv_dollar = "Formula") -> final_data_rm


# OPV,	PO in next 30 days,	Receipt in the next 30 days
final_data_rm %>% 
  dplyr::mutate(opv = "IQR Report",
                po_in_next_30_days = "IQR Report",
                receipt_in_next_30_days = "IQR Report") -> final_data_rm


# DOS,	Dead $,	Lot At Risk $,	Excess $,	Healthy Cycle Stock $,	SS OH $,	UPI $,	UPI $ (w/ Hold),	MOQ Flag,	Inv Health Flag
final_data_rm %>% 
  dplyr::mutate(dos = "Formula",
                dead_dollar = "Formula",
                lot_at_risk_dollar = "Formula",
                excess_dollar = "Formula",
                healthy_cycle_stock_dollar = "Formula",
                ss_oh_dollar = "Formula",
                upi_dollar = "Formula",
                upi_dollar_w_hold = "Formula",
                moq_flag = "Formula",
                inv_health_flag = "Formula") -> final_data_rm


# Current month dep demand,	Next month dep demand,	Total dep. demand Next 6 Months,	Total Last 6 mos Sales,	Total Last 12 mos Sales 
final_data_rm %>% 
  dplyr::mutate(current_month_dep_demand = "IQR Report",
                next_month_dep_demand = "IQR Report",
                total_dep_demand_next_6_months = "IQR Report",
                total_last_6_mos_sales = "IQR Report",
                total_last_12_mos_sales = "IQR Report") -> final_data_rm



# has Max?,	on hand Inv >max,	on hand Inv <= max,	on hand Inv > target,	on hand Inv <= target,	current month dep demand in $$,	next month dep demand in $$,	OH - Max in $,	IQR$,	IQR + hold $
final_data_rm %>% 
  dplyr::mutate(has_max = "Formula",
                on_hand_inv_gt_max = "Formula",
                on_hand_inv_lte_max = "Formula",
                on_hand_inv_gt_target = "Formula",
                on_hand_inv_lte_target = "Formula",
                current_month_dep_demand_in_dollar = "Formula",
                next_month_dep_demand_in_dollar = "Formula",
                oh_max_in_dollar = "Formula",
                iqr = "Formula",
                iqr_hold = "Formula") -> final_data_rm




###################################################################################################################################################


writexl::write_xlsx(final_data_rm, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/weekly Report run/2025/01.21.2025/rm_optimization.xlsx")
















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

exception_report <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/JDE Exception report extract/2024/exception report 2024.10.08.xlsx")
exception_report_dnrr <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/JDE DNRR Exception report extract/2024/exception report DOU 2024.10.08.xlsx")
inventory_fg <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Inventory/Inventory with Lot Report v.2 - 2024.10.08.xlsx",
                           sheet = "FG")
inventory_rm <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Inventory/Inventory with Lot Report v.2 - 2024.10.08.xlsx",
                           sheet = "RM")
oo_bt_fg <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/10012024/US and CAN OO BT where status _ J.xlsx")
dsx <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2024/DSX Forecast Backup - 2024.10.08.xlsx")
jde_25_55_label <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Inventory/JDE Inventory Lot Detail - 2024.10.08.xlsx")
lot_status_code <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Lot Status Code.xlsx")
bom <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/10012024/Bill of Material_100124.xlsx")
campus_ref <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Campus reference.xlsx")
complete_sku_list <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/10012024/Complete SKU list - Linda.xlsx")
iom_live <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/SS Optimization by Location - Finished Goods LIVE.xlsx")
unit_cost <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/10082024/Unit_Cost.xlsx")

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

iom_live[-1:-6, ] -> iom_live
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

### 1. Has on hand inventory (useable, soft hold, hard hold all included)
inventory_fg %>% 
  janitor::clean_names() %>% 
  dplyr::select(location, item, current_inventory_balance) %>% 
  dplyr::rename(inventory = current_inventory_balance) %>% 
  dplyr::mutate(inventory = as.double(inventory)) %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) %>% 
  dplyr::group_by(ref) %>% 
  dplyr::summarise(inventory = sum(inventory)) %>% 
  tidyr::separate(ref, into = c("location", "item"), sep = "_") %>%
  dplyr::mutate(ref = paste0(location, "_", item)) %>% 
  dplyr::filter(inventory > 0) %>%
  dplyr::select(-inventory) -> has_on_hand_inventory_fg


### 2. Zero on hand, but has open customer orders & branch transfer for next 30 days
oo_bt_fg_2 %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(oo_cases = as.double(oo_cases),
                oo_cases = ifelse(is.na(oo_cases), 0, oo_cases),
                b_t_open_order_cases = as.double(b_t_open_order_cases),
                b_t_open_order_cases = ifelse(is.na(b_t_open_order_cases), 0, b_t_open_order_cases)) %>% 
  dplyr::rename(item = product_label_sku) %>% 
  dplyr::mutate(item = gsub("-", "", item)) %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) %>% 
  dplyr::mutate(oo_bt_cases = oo_cases + b_t_open_order_cases) %>% 
  dplyr::group_by(ref) %>%
  dplyr::summarise(oo_bt_cases = sum(oo_bt_cases)) %>% 
  tidyr::separate(ref, into = c("location", "item"), sep = "_") %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) %>%
  dplyr::filter(oo_bt_cases > 0) %>%
  dplyr::select(-oo_bt_cases) -> zero_on_hand_has_open_orders_fg


### 3. Zero on hand, zero open customer orders, but has forecast for next 12 months 
dsx %>% 
  janitor::clean_names() %>% 
  dplyr::select(forecast_month_year_id, location_no, product_label_sku_code, adjusted_forecast_cases) %>% 
  dplyr::mutate(forecast_month_year_id = as.double(forecast_month_year_id),
                location_no = as.double(location_no),
                adjusted_forecast_cases = as.double(adjusted_forecast_cases)) %>% 
  dplyr::rename(forecast_month = forecast_month_year_id,
                location = location_no,
                item = product_label_sku_code,
                forecast = adjusted_forecast_cases) %>% 
  dplyr::mutate(item = gsub("-", "", item)) %>%
  dplyr::mutate(ref = paste0(location, "_", item)) %>% 
  dplyr::filter(forecast_month >= as.numeric(format(floor_date(today(), "month"), "%Y%m")) &
                  forecast_month <= as.numeric(format(floor_date(today() + months(11), "month"), "%Y%m"))) %>% 
  dplyr::mutate(forecast = ifelse(is.na(forecast), 0, forecast)) %>% 
  dplyr::group_by(ref) %>%
  dplyr::summarise(forecast = sum(forecast)) %>% 
  tidyr::separate(ref, into = c("location", "item"), sep = "_") %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) %>% 
  dplyr::filter(forecast > 0) %>%
  dplyr::select(-forecast) -> zero_on_hand_zero_open_orders_has_forecast_fg






### 4. None of the above but show ACTIVE in JDE
exception_report %>% 
  janitor::clean_names() %>% 
  dplyr::select(b_p, item_number) %>% 
  dplyr::mutate(ref = paste0(b_p, "_", item_number)) %>% 
  dplyr::rename(location = b_p, item = item_number) %>% 
  dplyr::distinct(ref, .keep_all = TRUE) %>% 
  dplyr::filter(stringr::str_detect(item, "^[0-9]+[a-zA-Z]+$")) %>%
  dplyr::distinct(ref, .keep_all = TRUE) -> active_items_fg





## Conclusion
dplyr::bind_rows(has_on_hand_inventory_fg, 
                 zero_on_hand_has_open_orders_fg, 
                 zero_on_hand_zero_open_orders_has_forecast_fg, 
                 active_items_fg) %>%
  dplyr::distinct(ref) %>% 
  tidyr::separate(ref, c("location", "item")) %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) -> final_data_fg




##################################################################################################################################################################
##################################################################################################################################################################
##################################################################################################################################################################


# mfg_loc
final_data_fg %>% 
  dplyr::left_join(complete_sku_list %>% 
                     janitor::clean_names() %>% 
                     data.frame() %>% 
                     dplyr::select(item_location_no_v2, product_label_sku, product_manufacturing_location) %>% 
                     dplyr::rename(item = product_label_sku,
                                   location = item_location_no_v2,
                                   mfg_loc = product_manufacturing_location) %>% 
                     dplyr::mutate(item = gsub("-", "", item)) %>% 
                     dplyr::mutate(ref = paste0(location, "_", item)) %>% 
                     dplyr::select(ref, mfg_loc), by = "ref") %>% 
  dplyr::relocate(location, mfg_loc, item, ref) -> final_data_fg


# Campus
final_data_fg %>% 
  dplyr::left_join(campus_ref %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(location, campus), by = "location") %>% 
  dplyr::relocate(campus, .after = mfg_loc) -> final_data_fg



# Category, Platform, Macro-Platform, Sub Type
final_data_fg %>% 
  dplyr::mutate(category = "IQR Report",
                platform = "IQR Report",
                macro_platform = "IQR Report",
                sub_type = "Formula") %>% 
  dplyr::relocate(category, platform, macro_platform, sub_type, .after = item) -> final_data_fg


# CVM
final_data_fg %>% 
  dplyr::left_join(iom_live %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(ship_ref, cvm_channel, focus_label) %>% 
                     dplyr::rename(ref = ship_ref,
                                   cvm = cvm_channel) %>% 
                     dplyr::mutate(ref = gsub("-", "_", ref)), by = "ref") %>% 
  dplyr::relocate(cvm, focus_label, .after = sub_type) -> final_data_fg



# ref, mfg_ref, campus_ref, base, label
final_data_fg %>% 
  dplyr::mutate(mfg_ref = "Formula",
                campus_ref = "Formula",
                base = "Formula",
                label = "Formula") %>% 
  dplyr::relocate(ref, mfg_ref, campus_ref, base, label, .after = focus_label) -> final_data_fg



# Description
final_data_fg %>% 
  dplyr::left_join(rbind(exception_report, exception_report_dnrr) %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(item_number, description) %>% 
                     dplyr::distinct(item_number, .keep_all = TRUE) %>% 
                     dplyr::rename(item = item_number), by = "item") -> final_data_fg


# MTO/MTS, MPF, Planner, Planner Name
final_data_fg %>% 
  dplyr::mutate(mto_mts = "IQR Report",
                mpf = "IQR Report",
                planner = "IQR Report",
                planner_name = "IQR Report") %>% 
  dplyr::relocate(mto_mts, mpf, planner, planner_name, .after = description) %>% data.frame() -> final_data_fg



# Qty per pallet
final_data_fg %>% 
  dplyr::left_join(complete_sku_list %>% 
                     janitor::clean_names() %>% 
                     data.frame() %>% 
                     dplyr::select(fg_cases_per_pallet, item_location_no_v2, product_label_sku) %>% 
                     dplyr::mutate(product_label_sku = gsub("-", "", product_label_sku)) %>% 
                     dplyr::mutate(ref = paste0(item_location_no_v2, "_", product_label_sku)) %>% 
                     dplyr::select(ref, fg_cases_per_pallet), by = "ref") %>% 
  dplyr::rename(qty_per_pallet = fg_cases_per_pallet) -> final_data_fg


# Storage condition
final_data_fg %>% 
  dplyr::left_join(iom_live %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(ship_ref, ref_code) %>% 
                     dplyr::rename(ref = ship_ref) %>% 
                     dplyr::mutate(ref = gsub("-", "_", ref)), by = "ref") %>% 
  dplyr::rename(storage_condition = ref_code) -> final_data_fg


# Pack Size
final_data_fg %>% 
  dplyr::left_join(complete_sku_list %>% 
                     janitor::clean_names() %>% 
                     data.frame() %>% 
                     dplyr::select(na_6, item_location_no_v2, product_label_sku) %>% 
                     dplyr::mutate(product_label_sku = gsub("-", "", product_label_sku)) %>% 
                     dplyr::mutate(ref = paste0(item_location_no_v2, "_", product_label_sku)) %>% 
                     dplyr::select(ref, na_6) %>% 
                     dplyr::rename(pack_size = na_6), by = "ref") -> final_data_fg


# Formula
final_data_fg %>% 
  dplyr::left_join(iom_live %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(ship_ref, formula) %>% 
                     dplyr::rename(ref = ship_ref) %>% 
                     dplyr::mutate(ref = gsub("-", "_", ref)), by = "ref") -> final_data_fg


# Net Wt LBS
final_data_fg %>% 
  dplyr::left_join(complete_sku_list %>% 
                     janitor::clean_names() %>% 
                     data.frame() %>% 
                     dplyr::select(fg_net_weight, item_location_no_v2, product_label_sku) %>% 
                     dplyr::mutate(product_label_sku = gsub("-", "", product_label_sku)) %>% 
                     dplyr::mutate(ref = paste0(item_location_no_v2, "_", product_label_sku)) %>% 
                     dplyr::select(ref, fg_net_weight ) %>% 
                     dplyr::rename(net_wt_lbs = fg_net_weight ), by = "ref") -> final_data_fg


# Unit Cost
final_data_fg %>% 
  dplyr::left_join(unit_cost %>% 
                     janitor::clean_names() %>% 
                     dplyr::mutate(location = as.double(location)) %>% 
                     dplyr::mutate(ref = paste0(location, "_", item)) %>% 
                     dplyr::rename(unit_cost = simulated_cost) %>% 
                     dplyr::select(ref, unit_cost), by = "ref") %>% 
  dplyr::mutate(unit_cost = ifelse(is.na(unit_cost), NA, round(as.numeric(unit_cost), 2))) -> final_data_fg



# JDE MOQ
final_data_fg %>% 
  dplyr::mutate(jde_moq = "IQR Report") -> final_data_fg




# Shelf life, Ship Shelf Life %, Shippable Shelf Life
final_data_fg %>% 
  dplyr::left_join(complete_sku_list %>% 
                     janitor::clean_names() %>% 
                     data.frame() %>% 
                     dplyr::mutate(product_label_sku = gsub("-", "", product_label_sku)) %>% 
                     dplyr::mutate(ref = paste0(item_location_no_v2, "_", product_label_sku)) %>% 
                     dplyr::select(ref, product_ship_shelf_life_percent, product_shelf_life_days) %>% 
                     dplyr::rename(shelf_life = product_shelf_life_days,
                                   shippable_shelf_life_percent = product_ship_shelf_life_percent) %>% 
                     dplyr::mutate(shippable_shelf_life_percent = as.numeric(shippable_shelf_life_percent),
                                   shelf_life = as.numeric(shelf_life),
                                   shippable_shelf_life = shelf_life * shippable_shelf_life_percent / 100, 
                                   shippable_shelf_life = round(shippable_shelf_life, 0))) %>% 
  dplyr::select(-shelf_life, -shippable_shelf_life_percent) -> final_data_fg


# Hold Days
final_data_fg %>% 
  dplyr::left_join(complete_sku_list %>% 
                     janitor::clean_names() %>% 
                     data.frame() %>% 
                     dplyr::mutate(product_label_sku = gsub("-", "", product_label_sku)) %>% 
                     dplyr::mutate(ref = paste0(item_location_no_v2, "_", product_label_sku)) %>% 
                     dplyr::select(ref, qc_hold_days) %>% 
                     dplyr::rename(hold_days = qc_hold_days), by = "ref") -> final_data_fg



# Current SS
final_data_fg %>% 
  dplyr::mutate(current_ss = "IQR Report") -> final_data_fg



# current_ss_plt, max_cycle_stock_lag1, max_cycle_stock, Max Cycle Stock Mfg Adjusted Forward, Max Cycle Stock $, 
# Max Cycle Stock (plt), Avg Cycle Stock (plt) 
final_data_fg %>% 
  dplyr::mutate(current_ss_plt = "Formula",
                max_cycle_stock_lag1 = "Formula",
                max_cycle_stock = "Formula",
                max_cycle_stock_mfg_adjusted_forward = "Formula",
                max_cycle_stock_dollars = "Formula",
                max_cycle_stock_plt = "Formula",
                avg_cycle_stock_plt = "Formula") -> final_data_fg

# Usable, Quality Hold
final_data_fg %>% 
  dplyr::mutate(usable = "IQR Report",
                quality_hold = "IQR Report") -> final_data_fg

# Quality hold $, Quality hold (plt)
final_data_fg %>% 
  dplyr::mutate(quality_hold_dollars = "Formula",
                quality_hold_plt = "Formula") -> final_data_fg

# Soft Hold
final_data_fg %>% 
  dplyr::mutate(soft_hold = "IQR Report") -> final_data_fg


# On Hand (usable + soft hold), On Hand in pounds,	On Hand $,	On Hand (plt),	Total Inventory (plt),	
# On Hand - Max $,	On Hand - Mfg Adjusted Forward Max $,	Inventory Target,	Inventory Target (lbs.),	Inventory Target $
# Inventory Target (plt),	Max Inventory Target,	Max Inventory Target (lbs.),	Max Inventory Target $,	Max Inventory Target (plt)
# Mfg Adjusted Forward Inv Target,	Mfg Adjusted Forward Inv Target (lbs.),	Mfg Adjusted Forward Inv Target $
# Mfg Adjusted Forward Inv Max,	Mfg Adjusted Forward Inv Max (lbs.),	Mfg Adjusted Forward Inv Max $,	
# Forward Inv Target lag 1 Current Month Fcst, Forward Inv Target lag 1 Current Month Fcst (lbs.)	
# Forward Inv Target lag 1 Current Month Fcst $

final_data_fg %>% 
  dplyr::mutate(on_hand = "Formula",
                on_hand_lbs = "Formula",
                on_hand_dollars = "Formula",
                on_hand_plt = "Formula",
                total_inventory_plt = "Formula",
                on_hand_max_dollars = "Formula",
                on_hand_mfg_adjusted_forward_max_dollars = "Formula",
                inventory_target = "Formula",
                inventory_target_lbs = "Formula",
                inventory_target_dollars = "Formula",
                inventory_target_plt = "Formula",
                max_inventory_target = "Formula",
                max_inventory_target_lbs = "Formula",
                max_inventory_target_dollars = "Formula",
                max_inventory_target_plt = "Formula",
                mfg_adjusted_forward_inv_target = "Formula",
                mfg_adjusted_forward_inv_target_lbs = "Formula",
                mfg_adjusted_forward_inv_target_dollars = "Formula",
                mfg_adjusted_forward_inv_max = "Formula",
                mfg_adjusted_forward_inv_max_lbs = "Formula",
                mfg_adjusted_forward_inv_max_dollars = "Formula",
                forward_inv_target_lag1_current_month_fcst = "Formula",
                forward_inv_target_lag1_current_month_fcst_lbs = "Formula",
                forward_inv_target_lag1_current_month_fcst_dollars = "Formula") -> final_data_fg

# OPV,	CustOrd in next 7 days,	CustOrd in next 14 days,	CustOrd in next 21 days,	CustOrd in next 28 days
final_data_fg %>% 
  dplyr::mutate(opv = "IQR Report",
                custord_in_next_7_days = "IQR Report",
                custord_in_next_14_days = "IQR Report",
                custord_in_next_21_days = "IQR Report",
                custord_in_next_28_days = "IQR Report") -> final_data_fg

# CustOrd in next 28 days $
final_data_fg %>% 
  dplyr::mutate(custord_in_next_28_days_dollars = "Formula") -> final_data_fg


# Mfg CustOrd in next 7 days,	Mfg CustOrd in next 14 days,	Mfg CustOrd in next 21 days,	Mfg CustOrd in next 28 days

final_data_fg %>% 
  dplyr::mutate(mfg_custord_in_next_7_days = "IQR Report",
                mfg_custord_in_next_14_days = "IQR Report",
                mfg_custord_in_next_21_days = "IQR Report",
                mfg_custord_in_next_28_days = "IQR Report") -> final_data_fg

# Mfg CustOrd in next 28 days $,	Current XS Pallets
final_data_fg %>% 
  dplyr::mutate(mfg_custord_in_next_28_days_dollars = "Formula",
                current_xs_pallets = "IQR Report") -> final_data_fg

# Firm WO in next 28 days,	Receipt in the next 28 days
final_data_fg %>% 
  dplyr::mutate(firm_wo_in_next_28_days = "IQR Report",
                receipt_in_next_28_days = "IQR Report") -> final_data_fg


# DOS,	DOS after CustOrd,	Max Inv DOS,	Target Inv DOS (includes Orders),	
# Mfg DOS,	Mfg DOS after CustOrd,	Mfg Adjusted Forward Max Inv DOS,	Mfg Forward Target Inv DOS (fcst only),	
# Mfg Adjusted Forward Target Inv DOS (includes Orders),	Inv Health,	Mfg Inv Health

final_data_fg %>% 
  dplyr::mutate(dos = "Formula",
                dos_after_custord = "Formula",
                max_inv_dos = "Formula",
                target_inv_dos = "Formula",
                mfg_dos = "Formula",
                mfg_dos_after_custord = "Formula",
                mfg_adjusted_forward_max_inv_dos = "Formula",
                mfg_forward_target_inv_dos = "Formula",
                mfg_adjusted_forward_target_inv_dos = "Formula",
                inv_health = "Formula",
                mfg_inv_health = "Formula") -> final_data_fg


# Lag 1 Current Month Fcst,	Lag 1 Current Month Fcst $,	Current Month Fcst,	Next Month Fcst,	Mfg Current Month Fcst,
# Mfg Next Month Fcst,	Total Last 6 mos Sales,	Total Last 12 mos Sales, Total Forecast Next 12 Months,	Total mfg Forecast Next 12 Months
final_data_fg %>% 
  dplyr::mutate(lag1_current_month_fcst = "IQR Report",
                lag1_current_month_fcst_dollars = "Formula",
                current_month_fcst = "IQR Report",
                next_month_fcst = "IQR Report",
                mfg_current_month_fcst = "IQR Report",
                mfg_next_month_fcst = "IQR Report",
                total_last_6_mos_sales = "IQR Report",
                total_last_12_mos_sales = "IQR Report",
                total_forecast_next_12_months = "IQR Report",
                total_mfg_forecast_next_12_months = "IQR Report") -> final_data_fg

# has adjusted forward looking Max?, on hand Inv > AF max,	on hand Inv <= AF max,
# on hand Inv > Adjusted Forward looking target,	on hand Inv <= AF target,	on hand Inv after CustOrd > AF max,
# on hand Inv after CustOrd <= AF max,	on hand Inv after CustOrd > AF target,	on hand Inv after CustOrd <= AF target,
# on hand inv after 28 days CustOrd > 0,	has mfg adjusted forward looking Max?,	on hand Inv > mfg AF max,
# on hand Inv <= mfg AF max,	on hand Inv > mfg Adjusted Forward looking target,	on hand Inv <= mfg AF target,
# on hand Inv after CustOrd > mfg AF max,	on hand Inv after CustOrd <= mfg AF max,	on hand Inv after CustOrd > mfg AF target
# on hand Inv after CustOrd <= mfg AF target,	on hand inv after mfg 28 days CustOrd > 0,	Current Month Fcst $,	Next Month Fcst $

final_data_fg %>% 
  dplyr::mutate(has_adjusted_forward_looking_max = "Formula",
                on_hand_inv_gt_af_max = "Formula",
                on_hand_inv_lte_af_max = "Formula",
                on_hand_inv_gt_af_target = "Formula",
                on_hand_inv_lte_af_target = "Formula",
                on_hand_inv_after_custord_gt_af_max = "Formula",
                on_hand_inv_after_custord_lte_af_max = "Formula",
                on_hand_inv_after_custord_gt_af_target = "Formula",
                on_hand_inv_after_custord_lte_af_target = "Formula",
                on_hand_inv_after_28_days_custord_gt_0 = "Formula",
                has_mfg_adjusted_forward_looking_max = "Formula",
                on_hand_inv_gt_mfg_af_max = "Formula",
                on_hand_inv_lte_mfg_af_max = "Formula",
                on_hand_inv_gt_mfg_af_target = "Formula",
                on_hand_inv_lte_mfg_af_target = "Formula",
                on_hand_inv_after_custord_gt_mfg_af_max = "Formula",
                on_hand_inv_after_custord_lte_mfg_af_max = "Formula",
                on_hand_inv_after_custord_gt_mfg_af_target = "Formula",
                on_hand_inv_after_custord_lte_mfg_af_target = "Formula",
                on_hand_inv_after_mfg_28_days_custord_gt_0 = "Formula",
                current_month_fcst_dollars = "Formula",
                next_month_fcst_dollars = "Formula") -> final_data_fg


# Open Orders (All)
final_data_fg %>% 
  dplyr::mutate(open_orders_all = "IQR Report") -> final_data_fg


# On Hand - Open Order (All),	On Hand - Open Orders (all) $,	Campus weighted cost,	Campus OH - OO
final_data_fg %>% 
  dplyr::mutate(on_hand_open_order_all = "Formula",
                on_hand_open_order_all_dollars = "Formula",
                campus_weighted_cost = "IQR Report",
                campus_oh_oo = "Formula") -> final_data_fg


###################################################################################################################################################

writexl::write_xlsx(final_data_fg, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/10.08.2024/fg_optimization.xlsx")




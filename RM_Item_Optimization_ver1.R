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

exception_report <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/JDE Exception report extract/2024/exception report 2024.12.03.xlsx")
exception_report_dnrr <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/JDE DNRR Exception report extract/2024/exception report DOU 2024.12.03.xlsx")
inventory_fg <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Inventory/Inventory with Lot Report v.2 - 2024.12.03.xlsx",
                           sheet = "FG")
inventory_rm <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Inventory/Inventory with Lot Report v.2 - 2024.12.03.xlsx",
                           sheet = "RM")
oo_bt_fg <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/12032024/US and CAN OO BT where status _ J.xlsx")
dsx <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2024/DSX Forecast Backup - 2024.12.02.xlsx")
jde_25_55_label <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Inventory/JDE Inventory Lot Detail - 2024.12.03.xlsx")
lot_status_code <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Lot Status Code.xlsx")
bom <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/12032024/Bill of Material_12032024.xlsx")
campus_ref <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Campus reference.xlsx")
iom_live <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/SS Optimization by Location - Finished Goods LIVE.xlsx",
                       sheet = "CVM & Focus label & Contract")
iom_live_1st_sheet <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/SS Optimization by Location - Finished Goods LIVE.xlsx")



complete_sku_list <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/12032024/Complete SKU list - Linda.xlsx")
unit_cost <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Report ingredients/Stan/12032024/Unit_Cost.xlsx")

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


# 1. Has inventory (useable, soft hold, hard hold all included)

inventory_rm[-1, ] -> inventory_rm
colnames(inventory_rm) <- inventory_rm[1, ]
inventory_rm[-1, ] -> inventory_rm

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


jde_25_55_label[-1:-5, ] -> jde_25_55_label
colnames(jde_25_55_label) <- jde_25_55_label[1, ]
jde_25_55_label[-1, ] -> jde_25_55_label

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
  dplyr::filter(!(ref %in% c("60_8883", "75_16975", "75_21645"))) -> final_data_rm


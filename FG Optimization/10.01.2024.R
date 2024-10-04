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

exception_report <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/09.03.2024/exception report.xlsx")
inventory_fg <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/09.03.2024/inventory.xlsx",
                           sheet = "FG")
inventory_rm <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/09.03.2024/inventory.xlsx",
                           sheet = "RM")
oo_bt_fg <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/09.03.2024/US and CAN OO BT where status _ J.xlsx")
dsx <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2024/DSX Forecast Backup - 2024.08.30.xlsx")
jde_25_55_label <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/09.03.2024/JDE 25,55.xlsx")
lot_status_code <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Lot Status Code.xlsx")
bom <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/BoM version 2/Weekly Run/2024/09.03.2024/Bill of Material_090324.xlsx")
campus_ref <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Campus reference.xlsx")

###################################################################

exception_report[-1:-2, ] -> exception_report
colnames(exception_report) <- exception_report[1, ]
exception_report[-1, -32] -> exception_report

###################################################################




#################################################### Finished Goods ############################################################################################    


### 1. Has on hand inventory (useable, soft hold, hard hold all included)


inventory_fg[-1, ] -> inventory_fg
colnames(inventory_fg) <- inventory_fg[1, ]
inventory_fg[-1, ] -> inventory_fg


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
oo_bt_fg %>% 
  dplyr::slice(c(-1, -3)) -> oo_bt_fg_2

colnames(oo_bt_fg_2) <- oo_bt_fg_2[1, ]
oo_bt_fg_2[-1, ] -> oo_bt_fg_2


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
dsx[-1,] -> dsx
colnames(dsx) <- dsx[1, ]
dsx[-1, ] -> dsx

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
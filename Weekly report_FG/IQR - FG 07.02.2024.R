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



# dir.create("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/06.18.2024")

file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/06.25.2024/Finished Goods Inventory Health Adjusted Forward (IQR) NEW TEMPLATE - 06.25.2024.xlsx",
          "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/07.02.2024/Finished Goods Inventory Health Adjusted Forward (IQR) NEW TEMPLATE - 07.02.2024.xlsx",
          overwrite = TRUE)


# for exposure analysis tab
# https://venturafoods.sharepoint.com/sites/ExpiredProductReporting/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FExpiredProductReporting%2FShared%20Documents%2FExpiration%20Risk%20Management%2FFinished%20Good%20Risk%2FFG%20%2D%20Raw%20Data%20Weekly%20Risk%20Original%20Files%2D%20For%20Downloading%20Only&p=true&ga=1



##################################################################################################################################################################

specific_date <- as.Date("2024-07-02")

# (Path Revision Needed) Planner Address Book (If updated, correct this link) ----
# sdrive: S:/Supply Chain Projects/Linda Liang/reference files/Address Book - 04.26.22.xlsx
Planner_address <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Address Book/Address Book - 2024.07.02.xlsx", 
                              sheet = "employee", col_types = c("text", 
                                                                "text", "text", "text", "text"))

names(Planner_address) <- str_replace_all(names(Planner_address), c(" " = "_"))

colnames(Planner_address)[1] <- "Planner"

Planner_address %>% 
  dplyr::select(1:2) -> Planner_address


# macro_platform ----
macro_platform <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/07.02.2024/Finished Goods Inventory Health Adjusted Forward (IQR) NEW TEMPLATE - 07.02.2024.xlsx",
                             sheet = "Macro-Platform",
                             col_names = FALSE)

colnames(macro_platform) <- macro_platform[1, ]
macro_platform[-1, ] -> macro_platform

macro_platform %>% 
  dplyr::rename(Macro_Platform = `Macro-Platform`) -> macro_platform

## FG_ref_to_mpg_ref 

FG_ref_to_mfg_ref <- read_excel("S:/Supply Chain Projects/RStudio/BoM/Master formats/FG_On_Hand/FG_ref_to_mfg_ref.xlsx")

FG_ref_to_mfg_ref %<>% 
  dplyr::mutate(ref = gsub("-", "_", ref),
                Campus_Ref = gsub("-", "_", Campus_Ref),
                Mfg_Ref = gsub("-", "_", Mfg_Ref)) %>% 
  dplyr::rename(campus_ref = Campus_Ref,
                mfg_ref = Mfg_Ref,
                mfg_loc = "mfg loc") %>% 
  dplyr::mutate(mfg_loc = gsub("-$", "", mfg_loc))

FG_ref_to_mfg_ref[!duplicated(FG_ref_to_mfg_ref[,c("mfg_loc", "ref")]),] -> FG_ref_to_mfg_ref


# (Path Revision Needed) Exception Report ----

exception_report <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/07.02.2024/exception report.xlsx", 
                               sheet = "Sheet1",
                               col_types = c("text", "text", "text", 
                                             "text", "numeric", "text", "text", "text", 
                                             "text", "text", "text", "text", "text", 
                                             "text", "numeric", "numeric", "numeric", 
                                             "numeric", "numeric", "numeric", 
                                             "numeric", "text", "text", "text", 
                                             "text", "text", "text", "text", "numeric", 
                                             "text", "text", "text"))

exception_report[-1:-2, -32] -> exception_report

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


exception_report %>% 
  dplyr::mutate(ref = paste0(B_P, "_", ItemNo)) %>% 
  dplyr::relocate(ref) ->  exception_report

exception_report[!duplicated(exception_report[,c("ref")]),] -> exception_report

# (Path Revision Needed) Campus_ref pulling ----
# S drive: "S:/Supply Chain Projects/RStudio/BoM/Master formats/RM_on_Hand/Campus_ref.xlsx"

Campus_ref <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Campus reference.xlsx")

Campus_ref %>% 
  janitor::clean_names() %>% 
  dplyr::rename(B_P = location,
                Campus = campus) %>% 
  dplyr::mutate(B_P = as.numeric(B_P)) -> Campus_ref


# Vlookup for Campus_ref

merge(exception_report, Campus_ref[, c("B_P", "Campus")], by = "B_P", all.x = TRUE) %>% 
  dplyr::mutate(campus_ref = paste0(Campus, "_", ItemNo)) %>% 
  dplyr::relocate(ref, campus_ref, Campus) %>% 
  dplyr::rename(Loc_SKU = campus_ref,
                campus = Campus) -> exception_report

# get the FG Item only. 

# right formula
exception_report$ItemNo -> label

data.frame(substr(label, nchar(label)-2, nchar(label))) -> label


exception_report %<>% 
  dplyr::bind_cols(label)

colnames(exception_report)[ncol(exception_report)] <- "label"

exception_report %>% 
  dplyr::mutate(label_test = as.integer(label)) %>% 
  dplyr::mutate(label_test = is.na(label_test)) %>% 
  dplyr::filter(label_test == TRUE) -> exception_report

# Planner NA to 0 in exception_report before vlookup
exception_report %>% 
  dplyr::mutate(Planner = replace(Planner, is.na(Planner), 0)) -> exception_report


# MPF NA to 0 in exception_report before vlookup
exception_report %>% 
  dplyr::mutate(MPF_or_Line = replace(MPF_or_Line, is.na(MPF_or_Line), 0)) -> exception_report

# Pivoting exception_report
reshape2::dcast(exception_report, Loc_SKU ~ ., value.var = "Safety_Stock", sum) %>% 
  dplyr::rename(Safety_Stock = ".") -> exception_report_pivot



# (Path Revision Needed) Custord PO ----
po <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/07.02.2024/Copy of PO Reporting Tool - 07.02.24.xlsx",
                 sheet = "Daily Open PO")


po %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(loc_item = gsub("-", "_", loc_item)) %>% 
  dplyr::select(loc_item, x2nd_item_number, location, quantity_to_receive, promised_delivery_date) %>% 
  dplyr::rename(ref = loc_item,
                item = x2nd_item_number,
                loc = location,
                qty = quantity_to_receive,
                date = promised_delivery_date) %>% 
  dplyr::mutate(date = as.Date(date),
                year = year(date),
                month = month(date),
                day = day(date),
                month_year = paste0(month, "_", year)) -> po


# PO_Pivot 
po %>% 
  dplyr::mutate(next_28_days = ifelse(date >= specific_date & date <= specific_date + 28, "Y", "N")) -> po


reshape2::dcast(po, ref ~ next_28_days, value.var = "qty", sum) %>% 
  dplyr::rename(loc_sku = ref) -> po_pivot


# (Path Revision Needed) Custord Receipt ----
receipt <- read.csv("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/DSXIE/2024/07.02/receipt.csv",
                    header = FALSE)


# Base receipt variable
receipt %>% 
  dplyr::select(-1) %>% 
  dplyr::slice(-1) %>% 
  dplyr::rename(aa = V2) %>% 
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
  
  dplyr::mutate(loc = as.numeric(str_replace_all(loc, "[A-Za-z]", ""))) %>% 
  
  dplyr::mutate(ref = paste0(loc, "_", Item),
                next_28_days = ifelse(date >= specific_date & date <= specific_date + 28, "Y", "N")) %>% 
  dplyr::relocate(ref) %>% 
  dplyr::rename(item = Item) -> receipt




# receipt_pivot 
reshape2::dcast(receipt, ref ~ next_28_days, value.var = "qty", sum) -> Receipt_Pivot  




# (Path Revision Needed) Custord wo ----
wo <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/07.02.2024/Wo.xlsx")


wo %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) %>% 
  dplyr::mutate(in_next_7_days = ifelse(date >= specific_date & date < specific_date+7, "Y", "N")) %>% 
  dplyr::rename(Item = item,
                Location = location,
                Qty = production_scheduled_cases) %>% 
  reshape2::dcast(ref ~ in_next_7_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(N = as.integer(N)) -> wo_pivot



# (Path Revision Needed) custord custord ----
# Open Customer Order File pulling ----  Change Directory ----
custord <- read.xlsx("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/07.02.2024/US and CAN OO BT where status _ J.xlsx",
                     colNames = FALSE)

custord %>% 
  dplyr::slice(c(-1, -3)) -> custord

colnames(custord) <- custord[1, ]
custord[-1, ] -> custord


custord %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(product_label_sku = gsub("-", "", product_label_sku)) %>% 
  dplyr::mutate(ref = paste0(location, "_", product_label_sku)) %>% 
  dplyr::mutate(oo_cases = as.double(oo_cases),
                oo_cases = ifelse(is.na(oo_cases), 0, oo_cases),
                b_t_open_order_cases = as.double(b_t_open_order_cases),
                b_t_open_order_cases = ifelse(is.na(b_t_open_order_cases), 0, b_t_open_order_cases)) %>%
  dplyr::mutate(Qty = oo_cases + b_t_open_order_cases) %>% 
  dplyr::mutate(sales_order_requested_ship_date = as_date(as.integer(sales_order_requested_ship_date), origin = "1899-12-30")) %>% 
  dplyr::select(ref, product_label_sku, location, Qty, sales_order_requested_ship_date) %>% 
  dplyr::rename(Item = product_label_sku,
                Location = location,
                date = sales_order_requested_ship_date) %>% 
  dplyr::group_by(ref, Item, Location, date) %>% 
  dplyr::summarise(Qty = sum(Qty)) %>% 
  dplyr::relocate(Qty, .after = Location) %>% 
  dplyr::mutate(in_next_7_days = ifelse(date >= specific_date & date < specific_date +7, "Y", "N"),
                in_next_14_days = ifelse(date < specific_date + 14, "Y", "N"),
                in_next_21_days = ifelse(date < specific_date + 21, "Y", "N"),
                in_next_28_days = ifelse(date < specific_date + 28, "Y", "N")) -> custord






########################################################################################################################

# custord_pivot_1, custord_pivot_2, custord_pivot_3, custord_pivot_4

reshape2::dcast(custord, ref ~ in_next_7_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(Total = N + Y) -> custord_pivot_1


reshape2::dcast(custord, ref ~ in_next_14_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(Total = N + Y) -> custord_pivot_2


reshape2::dcast(custord, ref ~ in_next_21_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(Total = N + Y) -> custord_pivot_3


reshape2::dcast(custord, ref ~ in_next_28_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(Total = N + Y) -> custord_pivot_4



# mfg custord 
merge(custord, FG_ref_to_mfg_ref[, c("ref", "mfg_ref")], by = "ref", all.x = TRUE) -> custord_mfg


# custord_mfg - pivot
reshape2::dcast(custord_mfg, mfg_ref ~ in_next_7_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(Total = N + Y) -> custord_mfg_pivot_1


reshape2::dcast(custord_mfg, mfg_ref ~ in_next_14_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(Total = N + Y) -> custord_mfg_pivot_2


reshape2::dcast(custord_mfg, mfg_ref ~ in_next_21_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(Total = N + Y) -> custord_mfg_pivot_3


reshape2::dcast(custord_mfg, mfg_ref ~ in_next_28_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(Total = N + Y) -> custord_mfg_pivot_4


# (Path Revision Needed) Lag 1 DSX Forecast pulling (Previous month file)---- Change Directory ----

DSX_Forecast_Backup_pre <- read_excel(
  "S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2024/DSX Forecast Backup - 2024.06.03.xlsx")

DSX_Forecast_Backup_pre[-1,] -> DSX_Forecast_Backup_pre
colnames(DSX_Forecast_Backup_pre) <- DSX_Forecast_Backup_pre[1, ]
DSX_Forecast_Backup_pre[-1, ] -> DSX_Forecast_Backup_pre

colnames(DSX_Forecast_Backup_pre)[1]  <- "Primary_Channel_ID"
colnames(DSX_Forecast_Backup_pre)[2]  <- "Segmentation_ID"
colnames(DSX_Forecast_Backup_pre)[3]  <- "Sub_Segment_ID"
colnames(DSX_Forecast_Backup_pre)[4]  <- "Forecast_Month_Year_Code_Segment_ID"
colnames(DSX_Forecast_Backup_pre)[5]  <- "Product_Manufacturing_Location_Code"
colnames(DSX_Forecast_Backup_pre)[6]  <- "Product_Manufacturing_Location_Name"
colnames(DSX_Forecast_Backup_pre)[7]  <- "Location_No"
colnames(DSX_Forecast_Backup_pre)[8]  <- "Location_Name"
colnames(DSX_Forecast_Backup_pre)[9]  <- "Product_Label_SKU_Code"
colnames(DSX_Forecast_Backup_pre)[10] <- "Product_Label_SKU_Name"
colnames(DSX_Forecast_Backup_pre)[11] <- "Product_Category_Name"
colnames(DSX_Forecast_Backup_pre)[12] <- "Product_Platform_Name"
colnames(DSX_Forecast_Backup_pre)[13] <- "Product_Group_Code"
colnames(DSX_Forecast_Backup_pre)[14] <- "Product_Group_Short_Name"
colnames(DSX_Forecast_Backup_pre)[15] <- "Product_Manufacturing_Line_Area_No_Code"
colnames(DSX_Forecast_Backup_pre)[16] <- "ABC_4_ID"
colnames(DSX_Forecast_Backup_pre)[17] <- "Safety_Stock_ID"
colnames(DSX_Forecast_Backup_pre)[18] <- "MTO_MTS_Gross_Requirements_Calc_Method_ID"
colnames(DSX_Forecast_Backup_pre)[19] <- "Adjusted_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup_pre)[20] <- "Adjusted_Forecast_Cases"
colnames(DSX_Forecast_Backup_pre)[21] <- "Stat_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup_pre)[22] <- "Stat_Forecast_Cases"
colnames(DSX_Forecast_Backup_pre)[23] <- "Cust_Ref_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup_pre)[24] <- "Cust_Ref_Forecast_Cases"




readr::type_convert(DSX_Forecast_Backup_pre) -> DSX_Forecast_Backup_pre

# Data wrangling - DSX_Forecast_Backup_pre 

DSX_Forecast_Backup_pre %>% 
  dplyr::mutate(Product_Label_SKU_Code = gsub("-", "", Product_Label_SKU_Code)) %>% 
  dplyr::mutate(ref = paste0(Location_No, "_", Product_Label_SKU_Code)) %>% 
  dplyr::mutate(mfg_ref = paste0(Product_Manufacturing_Location_Code, "_", Product_Label_SKU_Code)) %>% 
  dplyr::relocate(ref, mfg_ref) %>% 
  dplyr::mutate(Forecast_Month_Year_Code_Segment_ID = as.character(Forecast_Month_Year_Code_Segment_ID)) -> DSX_Forecast_Backup_pre


DSX_Forecast_Backup_pre %>% 
  dplyr::mutate(Safety_Stock_ID = replace(Safety_Stock_ID, is.na(Safety_Stock_ID), 0),
                Adjusted_Forecast_Pounds_lbs = replace(Adjusted_Forecast_Pounds_lbs, is.na(Adjusted_Forecast_Pounds_lbs), 0),
                Adjusted_Forecast_Cases = replace(Adjusted_Forecast_Cases, is.na(Adjusted_Forecast_Cases), 0),
                Stat_Forecast_Pounds_lbs = replace(Stat_Forecast_Pounds_lbs, is.na(Stat_Forecast_Pounds_lbs), 0),
                Stat_Forecast_Cases = replace(Stat_Forecast_Cases, is.na(Stat_Forecast_Cases), 0),
                Cust_Ref_Forecast_Pounds_lbs = replace(Cust_Ref_Forecast_Pounds_lbs, is.na(Cust_Ref_Forecast_Pounds_lbs), 0),
                Cust_Ref_Forecast_Cases = replace(Cust_Ref_Forecast_Cases, is.na(Cust_Ref_Forecast_Cases), 0)) -> DSX_Forecast_Backup_pre



# value n/a to 0
DSX_Forecast_Backup_pre$Adjusted_Forecast_Pounds_lbs -> dsx_na_1_pre
DSX_Forecast_Backup_pre$Adjusted_Forecast_Cases      -> dsx_na_2_pre
DSX_Forecast_Backup_pre$Stat_Forecast_Pounds_lbs     -> dsx_na_3_pre
DSX_Forecast_Backup_pre$Stat_Forecast_Cases          -> dsx_na_4_pre
DSX_Forecast_Backup_pre$Cust_Ref_Forecast_Pounds_lbs -> dsx_na_5_pre
DSX_Forecast_Backup_pre$Cust_Ref_Forecast_Cases      -> dsx_na_6_pre


dsx_na_1_pre[is.na(dsx_na_1_pre)] <- 0
dsx_na_2_pre[is.na(dsx_na_2_pre)] <- 0
dsx_na_3_pre[is.na(dsx_na_3_pre)] <- 0
dsx_na_4_pre[is.na(dsx_na_4_pre)] <- 0
dsx_na_5_pre[is.na(dsx_na_5_pre)] <- 0
dsx_na_6_pre[is.na(dsx_na_6_pre)] <- 0

cbind(DSX_Forecast_Backup_pre, dsx_na_1_pre, dsx_na_2_pre, dsx_na_3_pre, 
      dsx_na_4_pre, dsx_na_5_pre, dsx_na_6_pre) -> DSX_Forecast_Backup_pre

DSX_Forecast_Backup_pre[, -21:-26] -> DSX_Forecast_Backup_pre

colnames(DSX_Forecast_Backup_pre)[21] <- "Adjusted_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup_pre)[22] <- "Adjusted_Forecast_Cases"
colnames(DSX_Forecast_Backup_pre)[23] <- "Stat_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup_pre)[24] <- "Stat_Forecast_Cases"
colnames(DSX_Forecast_Backup_pre)[25] <- "Cust_Ref_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup_pre)[26] <- "Cust_Ref_Forecast_Cases"


reshape2::dcast(DSX_Forecast_Backup_pre, ref ~ Forecast_Month_Year_Code_Segment_ID , 
                value.var = "Adjusted_Forecast_Cases", sum) -> DSX_pivot_1_pre



colnames(DSX_pivot_1_pre)[1]  <- "ref"

colnames(DSX_pivot_1_pre)[2]  <- "last_mon_fcst"
colnames(DSX_pivot_1_pre)[3]  <- "Mon_a_fcst"
colnames(DSX_pivot_1_pre)[4]  <- "Mon_b_fcst"
colnames(DSX_pivot_1_pre)[5]  <- "Mon_c_fcst"
colnames(DSX_pivot_1_pre)[6]  <- "Mon_d_fcst"
colnames(DSX_pivot_1_pre)[7]  <- "Mon_e_fcst"
colnames(DSX_pivot_1_pre)[8]  <- "Mon_f_fcst"
colnames(DSX_pivot_1_pre)[9]  <- "Mon_g_fcst"


# (Path Revision Needed) DSX Forecast pulling (Current Month file) ---- Change Directory ----


DSX_Forecast_Backup <- read_excel(
  "S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2024/DSX Forecast Backup - 2024.07.02.xlsx") ### Match with BoM date! ####

DSX_Forecast_Backup[-1,] -> DSX_Forecast_Backup
colnames(DSX_Forecast_Backup) <- DSX_Forecast_Backup[1, ]
DSX_Forecast_Backup[-1, ] -> DSX_Forecast_Backup

colnames(DSX_Forecast_Backup)[1]  <- "Primary_Channel_ID"
colnames(DSX_Forecast_Backup)[2]  <- "Segmentation_ID"
colnames(DSX_Forecast_Backup)[3]  <- "Sub_Segment_ID"
colnames(DSX_Forecast_Backup)[4]  <- "Forecast_Month_Year_Code_Segment_ID"
colnames(DSX_Forecast_Backup)[5]  <- "Product_Manufacturing_Location_Code"
colnames(DSX_Forecast_Backup)[6]  <- "Product_Manufacturing_Location_Name"
colnames(DSX_Forecast_Backup)[7]  <- "Location_No"
colnames(DSX_Forecast_Backup)[8]  <- "Location_Name"
colnames(DSX_Forecast_Backup)[9]  <- "Product_Label_SKU_Code"
colnames(DSX_Forecast_Backup)[10] <- "Product_Label_SKU_Name"
colnames(DSX_Forecast_Backup)[11] <- "Product_Category_Name"
colnames(DSX_Forecast_Backup)[12] <- "Product_Platform_Name"
colnames(DSX_Forecast_Backup)[13] <- "Product_Group_Code"
colnames(DSX_Forecast_Backup)[14] <- "Product_Group_Short_Name"
colnames(DSX_Forecast_Backup)[15] <- "Product_Manufacturing_Line_Area_No_Code"
colnames(DSX_Forecast_Backup)[16] <- "ABC_4_ID"
colnames(DSX_Forecast_Backup)[17] <- "Safety_Stock_ID"
colnames(DSX_Forecast_Backup)[18] <- "MTO_MTS_Gross_Requirements_Calc_Method_ID"
colnames(DSX_Forecast_Backup)[19] <- "Adjusted_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup)[20] <- "Adjusted_Forecast_Cases"
colnames(DSX_Forecast_Backup)[21] <- "Stat_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup)[22] <- "Stat_Forecast_Cases"
colnames(DSX_Forecast_Backup)[23] <- "Cust_Ref_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup)[24] <- "Cust_Ref_Forecast_Cases"



readr::type_convert(DSX_Forecast_Backup) -> DSX_Forecast_Backup


# Data wrangling - DSX_Forecast_Backup 

DSX_Forecast_Backup %>% 
  dplyr::mutate(Product_Label_SKU_Code = gsub("-", "", Product_Label_SKU_Code)) %>% 
  dplyr::mutate(ref = paste0(Location_No, "_", Product_Label_SKU_Code)) %>% 
  dplyr::mutate(mfg_ref = paste0(Product_Manufacturing_Location_Code, "_", Product_Label_SKU_Code)) %>% 
  dplyr::relocate(ref, mfg_ref) %>% 
  dplyr::mutate(Forecast_Month_Year_Code_Segment_ID = as.character(Forecast_Month_Year_Code_Segment_ID)) -> DSX_Forecast_Backup


DSX_Forecast_Backup %>% 
  dplyr::mutate(Safety_Stock_ID = replace(Safety_Stock_ID, is.na(Safety_Stock_ID), 0),
                Adjusted_Forecast_Pounds_lbs = replace(Adjusted_Forecast_Pounds_lbs, is.na(Adjusted_Forecast_Pounds_lbs), 0),
                Adjusted_Forecast_Cases = replace(Adjusted_Forecast_Cases, is.na(Adjusted_Forecast_Cases), 0),
                Stat_Forecast_Pounds_lbs = replace(Stat_Forecast_Pounds_lbs, is.na(Stat_Forecast_Pounds_lbs), 0),
                Stat_Forecast_Cases = replace(Stat_Forecast_Cases, is.na(Stat_Forecast_Cases), 0),
                Cust_Ref_Forecast_Pounds_lbs = replace(Cust_Ref_Forecast_Pounds_lbs, is.na(Cust_Ref_Forecast_Pounds_lbs), 0),
                Cust_Ref_Forecast_Cases = replace(Cust_Ref_Forecast_Cases, is.na(Cust_Ref_Forecast_Cases), 0)) -> DSX_Forecast_Backup




# value n/a to 0
DSX_Forecast_Backup$Adjusted_Forecast_Pounds_lbs -> dsx_na_1
DSX_Forecast_Backup$Adjusted_Forecast_Cases      -> dsx_na_2
DSX_Forecast_Backup$Stat_Forecast_Pounds_lbs     -> dsx_na_3
DSX_Forecast_Backup$Stat_Forecast_Cases          -> dsx_na_4
DSX_Forecast_Backup$Cust_Ref_Forecast_Pounds_lbs -> dsx_na_5
DSX_Forecast_Backup$Cust_Ref_Forecast_Cases      -> dsx_na_6


dsx_na_1[is.na(dsx_na_1)] <- 0
dsx_na_2[is.na(dsx_na_2)] <- 0
dsx_na_3[is.na(dsx_na_3)] <- 0
dsx_na_4[is.na(dsx_na_4)] <- 0
dsx_na_5[is.na(dsx_na_5)] <- 0
dsx_na_6[is.na(dsx_na_6)] <- 0

cbind(DSX_Forecast_Backup, dsx_na_1, dsx_na_2, dsx_na_3, dsx_na_4, dsx_na_5, dsx_na_6) -> DSX_Forecast_Backup

DSX_Forecast_Backup[, -21:-26] -> DSX_Forecast_Backup

colnames(DSX_Forecast_Backup)[21] <- "Adjusted_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup)[22] <- "Adjusted_Forecast_Cases"
colnames(DSX_Forecast_Backup)[23] <- "Stat_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup)[24] <- "Stat_Forecast_Cases"
colnames(DSX_Forecast_Backup)[25] <- "Cust_Ref_Forecast_Pounds_lbs"
colnames(DSX_Forecast_Backup)[26] <- "Cust_Ref_Forecast_Cases"

# DSX forecast backup pivot (ref)
reshape2::dcast(DSX_Forecast_Backup, ref ~ Forecast_Month_Year_Code_Segment_ID , 
                value.var = "Adjusted_Forecast_Cases", sum) -> DSX_pivot_1



colnames(DSX_pivot_1)[1]  <- "ref"

colnames(DSX_pivot_1)[2]  <- "last_mon_fcst"
colnames(DSX_pivot_1)[3]  <- "Mon_a_fcst"
colnames(DSX_pivot_1)[4]  <- "Mon_b_fcst"
colnames(DSX_pivot_1)[5]  <- "Mon_c_fcst"
colnames(DSX_pivot_1)[6]  <- "Mon_d_fcst"
colnames(DSX_pivot_1)[7]  <- "Mon_e_fcst"
colnames(DSX_pivot_1)[8]  <- "Mon_f_fcst"
colnames(DSX_pivot_1)[9]  <- "Mon_g_fcst"

DSX_pivot_1 %>% 
  dplyr::select(1:14) %>% 
  janitor::adorn_totals(where = "col", na.rm = TRUE, name = "total_12_month") %>% 
  dplyr::mutate(total_12_month = total_12_month - last_mon_fcst) -> DSX_pivot_1




# DSX forecast backup pivot (mfg_ref)

reshape2::dcast(DSX_Forecast_Backup, mfg_ref ~ Forecast_Month_Year_Code_Segment_ID , 
                value.var = "Adjusted_Forecast_Cases", sum) -> DSX_mfg_pivot_1



colnames(DSX_mfg_pivot_1)[1]  <- "mfg_ref"

colnames(DSX_mfg_pivot_1)[2]  <- "last_mon_fcst"
colnames(DSX_mfg_pivot_1)[3]  <- "Mon_a_fcst"
colnames(DSX_mfg_pivot_1)[4]  <- "Mon_b_fcst"
colnames(DSX_mfg_pivot_1)[5]  <- "Mon_c_fcst"
colnames(DSX_mfg_pivot_1)[6]  <- "Mon_d_fcst"
colnames(DSX_mfg_pivot_1)[7]  <- "Mon_e_fcst"
colnames(DSX_mfg_pivot_1)[8]  <- "Mon_f_fcst"
colnames(DSX_mfg_pivot_1)[9]  <- "Mon_g_fcst"


DSX_mfg_pivot_1 %>% 
  dplyr::select(1:14) %>% 
  janitor::adorn_totals(where = "col", na.rm = TRUE, name = "total_12_month") %>% 
  dplyr::mutate(total_12_month = total_12_month - last_mon_fcst) -> DSX_mfg_pivot_1


# (Path Revision Needed) Inventory ----
inventory <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/07.02.2024/Inventory.xlsx",
                        sheet = "FG")

inventory[-1, ] -> inventory
colnames(inventory) <- inventory[1, ]
inventory[-1, ] -> inventory


### ref inventory

inventory %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(item = gsub("-", "", item)) %>%
  filter(!str_starts(description, "PWS ") & 
           !str_starts(description, "SUB ") & 
           !str_starts(description, "THW ") & 
           !str_starts(description, "PALLET")) %>% 
  dplyr::mutate(Loc_SKU = paste0(location, "_", item)) %>% 
  dplyr::select(Loc_SKU, inventory_hold_status, current_inventory_balance) %>% 
  dplyr::mutate(current_inventory_balance = as.numeric(current_inventory_balance)) %>% 
  tidyr::pivot_wider(names_from = inventory_hold_status, 
                     values_from = current_inventory_balance, 
                     values_fn = list(current_inventory_balance = sum)) %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(soft_hold = replace(soft_hold, is.na(soft_hold), 0),
                hard_hold = replace(hard_hold, is.na(hard_hold), 0),
                useable = replace(useable, is.na(useable), 0)) %>% 
  dplyr::rename(Useable = useable,
                Hard_Hold = hard_hold,
                Soft_Hold = soft_hold,
                ref = loc_sku) %>%
  dplyr::relocate(ref, Hard_Hold, Soft_Hold, Useable) -> pivot_ref_Inventory_analysis




### campus inventory

inventory %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(item = gsub("-", "", item)) %>%
  filter(!str_starts(description, "PWS ") & 
           !str_starts(description, "SUB ") & 
           !str_starts(description, "THW ") & 
           !str_starts(description, "PALLET")) %>% 
  dplyr::mutate(Loc_SKU = paste0(campus_no, "_", item)) %>% 
  dplyr::select(Loc_SKU, inventory_hold_status, current_inventory_balance) %>% 
  dplyr::mutate(current_inventory_balance = as.numeric(current_inventory_balance)) %>% 
  tidyr::pivot_wider(names_from = inventory_hold_status, 
                     values_from = current_inventory_balance, 
                     values_fn = list(current_inventory_balance = sum)) %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(soft_hold = replace(soft_hold, is.na(soft_hold), 0),
                hard_hold = replace(hard_hold, is.na(hard_hold), 0),
                useable = replace(useable, is.na(useable), 0)) %>% 
  dplyr::rename(Usable = useable,
                Hard_Hold = hard_hold,
                Soft_Hold = soft_hold,
                Loc_SKU = loc_sku) %>%
  dplyr::relocate(Loc_SKU, Hard_Hold, Soft_Hold, Usable) -> pivot_campus_ref_Inventory_analysis






# (Path Revision Needed) Main Dataset Board ----  

IQR_FG_sample <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/07.02.2024/Finished Goods Inventory Health Adjusted Forward (IQR) NEW TEMPLATE - 07.02.2024.xlsx",
                            sheet = "Location FG")

IQR_FG_sample[-1:-2,] -> IQR_FG_sample

colnames(IQR_FG_sample) <- IQR_FG_sample[1, ]
IQR_FG_sample[-1, ] -> IQR_FG_sample

IQR_FG_sample %>% 
  janitor::clean_names() -> IQR_FG_sample


IQR_FG_sample %>% 
  dplyr::mutate(ref = gsub("-", "_", ref),
                campus_ref = gsub("-", "_", campus_ref),
                mfg_ref = gsub("-", "_", mfg_ref)) %>% 
  dplyr::rename(loc_sku = campus_ref) -> IQR_FG_sample


# (Path Revision Needed) read SD & CV file ----
sdcv <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Standard Deviation & CV/Standard Deviation, CV,  June 2024.xlsx")

sdcv[-1:-3,] -> sdcv
colnames(sdcv) <- sdcv[1,]
sdcv[-1, ] -> sdcv

sdcv %>% 
  janitor::clean_names() %>% 
  dplyr::rename(last_6_month_sales = last_6_months_sales,
                last_12_month_sales = last_12_months_sales,
                total_forecast_next_12_months = x3_months_fcst) -> sdcv


sdcv %>%
  dplyr::select(ref, last_6_month_sales, last_12_month_sales, total_forecast_next_12_months) %>%
  dplyr::mutate(ref = gsub("-", "_", ref),
                last_6_month_sales = as.double(last_6_month_sales),
                last_12_month_sales = as.double(last_12_month_sales),
                total_forecast_next_12_months = as.double(total_forecast_next_12_months)) -> sdcv

sdcv %>% 
  dplyr::group_by(ref) %>% 
  dplyr::summarise(sum(last_6_month_sales), sum(last_12_month_sales), sum(total_forecast_next_12_months)) %>% 
  dplyr::rename(last_6_month_sales = "sum(last_6_month_sales)",
                last_12_month_sales = "sum(last_12_month_sales)",
                total_forecast_next_12_months = "sum(total_forecast_next_12_months)") -> sdcv

##########################################################################################################
##################################################  ETL  #################################################
##########################################################################################################

readr::type_convert(IQR_FG_sample) -> IQR_FG_sample


##################################### vlookups #########################################

# vlookup - MTO_MTS (with if, iferror)
merge(IQR_FG_sample, exception_report[, c("ref", "Order_Policy_Code")], by = "ref", all.x = TRUE) %>% 
  dplyr::select(-opv) %>% 
  dplyr::rename(opv = Order_Policy_Code) %>% 
  dplyr::mutate(opv = as.integer(opv)) -> IQR_FG_sample

IQR_FG_sample$opv[is.na(IQR_FG_sample$opv)] <- 1


IQR_FG_sample %>% 
  dplyr::mutate(mto_mts = ifelse(opv == 1, "MTO", "MTS")) -> IQR_FG_sample



# vlookup - MPF
merge(IQR_FG_sample, exception_report[, c("ref", "MPF_or_Line")], by = "ref", all.x = TRUE) %>% 
  dplyr::select(-mpf) %>% 
  dplyr::rename(mpf = MPF_or_Line) %>% 
  dplyr::mutate(mpf = replace(mpf, is.na(mpf), "DNRR")) -> IQR_FG_sample


# vlookup - Planner
merge(IQR_FG_sample, exception_report[, c("ref", "Planner")], by = "ref", all.x = TRUE) %>% 
  dplyr::select(-planner) %>% 
  dplyr::rename(planner = Planner) %>% 
  dplyr::mutate(planner = replace(planner, is.na(planner), "DNRR")) -> IQR_FG_sample



# vlookup - Planner Name 
Planner_address %>% 
  dplyr::rename(planner = Planner) -> Planner_address

merge(IQR_FG_sample, Planner_address[, c("planner", "Alpha_Name")], by = "planner", all.x = TRUE) %>% 
  dplyr::mutate(planner_name = ifelse(planner == 0, "NA",
                                      ifelse(planner == "DNRR", "DNRR",
                                             Alpha_Name))) %>% 
  dplyr::select(-Alpha_Name) -> IQR_FG_sample



# vlookup - JDE MOQ
merge(IQR_FG_sample, exception_report[, c("ref", "Reorder_MIN")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(reorder_min_na = !is.na(Reorder_MIN)) %>% 
  dplyr::mutate(jde_moq = ifelse(reorder_min_na == TRUE, Reorder_MIN, 0)) %>% 
  dplyr::select(-reorder_min_na, -Reorder_MIN) -> IQR_FG_sample



# vlookup - Current SS
merge(IQR_FG_sample, exception_report[, c("ref", "Safety_Stock")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(safety_stock_na = !is.na(Safety_Stock)) %>% 
  dplyr::mutate(current_ss = ifelse(safety_stock_na == TRUE, Safety_Stock, 0)) %>% 
  dplyr::select(-safety_stock_na, -Safety_Stock) -> IQR_FG_sample



# vlookup - Useable
merge(IQR_FG_sample, pivot_ref_Inventory_analysis[, c("ref", "Useable")], by = "ref", all.x = TRUE) %>%
  dplyr::mutate(useable_na = !is.na(Useable)) %>% 
  dplyr::mutate(usable = ifelse(useable_na == TRUE, Useable, 0)) %>%
  dplyr::select(-useable_na, -Useable)  -> IQR_FG_sample



# vlookup - Quality hold
merge(IQR_FG_sample, pivot_ref_Inventory_analysis[, c("ref", "Hard_Hold")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(hard_hold_na = !is.na(Hard_Hold)) %>% 
  dplyr::mutate(quality_hold = ifelse(hard_hold_na == TRUE, Hard_Hold, 0)) %>% 
  dplyr::select(-hard_hold_na, -Hard_Hold) -> IQR_FG_sample



# vlookup - Soft Hold
merge(IQR_FG_sample, pivot_ref_Inventory_analysis[, c("ref", "Soft_Hold")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(soft_hold_na = !is.na(Soft_Hold)) %>% 
  dplyr::mutate(soft_hold = ifelse(soft_hold_na == TRUE, Soft_Hold, 0)) %>% 
  dplyr::select(-soft_hold_na, -Soft_Hold) -> IQR_FG_sample



# vlookup - opv
merge(IQR_FG_sample, exception_report[, c("ref", "Order_Policy_Value")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(opv_na = !is.na(Order_Policy_Value)) %>% 
  dplyr::mutate(opv = ifelse(opv_na == TRUE, Order_Policy_Value, 0)) %>% 
  dplyr::select(-opv_na, -Order_Policy_Value) -> IQR_FG_sample



# vlookup - CustOrd in next 7 days
merge(IQR_FG_sample, custord_pivot_1[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(cust_ord_in_next_7_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample



# vlookup - CustOrd in next 14 days
merge(IQR_FG_sample, custord_pivot_2[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(cust_ord_in_next_14_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample



# vlookup - CustOrd in next 21 days
merge(IQR_FG_sample, custord_pivot_3[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(cust_ord_in_next_21_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample



# vlookup - CustOrd in next 28 days
merge(IQR_FG_sample, custord_pivot_4[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(cust_ord_in_next_28_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample



# vlookup - Firm WO in next 28 days
merge(IQR_FG_sample, wo_pivot[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(firm_wo_in_next_28_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample



# vlookup - Receipt in the next 28 days
merge(IQR_FG_sample, Receipt_Pivot[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(receipt_in_the_next_28_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - Lag_1_current_month_fcst  
merge(IQR_FG_sample, DSX_pivot_1_pre[, c("ref", "Mon_b_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(mon_b_na = !is.na(Mon_b_fcst)) %>% 
  dplyr::mutate(lag_1_current_month_fcst = ifelse(mon_b_na == TRUE, Mon_b_fcst, 0)) %>% 
  dplyr::select(-Mon_b_fcst, -mon_b_na) %>% 
  dplyr::mutate(lag_1_current_month_fcst = round(lag_1_current_month_fcst , 0)) -> IQR_FG_sample



# vlookup - Current Month Fcst
merge(IQR_FG_sample, DSX_pivot_1[, c("ref", "Mon_a_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(mon_a_na = !is.na(Mon_a_fcst)) %>% 
  dplyr::mutate(current_month_fcst = ifelse(mon_a_na == TRUE, Mon_a_fcst, 0)) %>% 
  dplyr::select(-Mon_a_fcst, - mon_a_na) %>% 
  dplyr::mutate(current_month_fcst = round(current_month_fcst , 0)) -> IQR_FG_sample



# vlookup - Next Month Fcst
merge(IQR_FG_sample, DSX_pivot_1[, c("ref", "Mon_b_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(mon_b_na = !is.na(Mon_b_fcst)) %>% 
  dplyr::mutate(next_month_fcst = ifelse(mon_b_na == TRUE, Mon_b_fcst, 0)) %>% 
  dplyr::select(-Mon_b_fcst, - mon_b_na) %>% 
  dplyr::mutate(next_month_fcst  = round(next_month_fcst, 0)) -> IQR_FG_sample



# vlookup - Total Last 6 mos Sales
merge(IQR_FG_sample, sdcv[, c("ref", "last_6_month_sales")], by = "ref", all.x = TRUE) %>% 
  dplyr::select(-total_last_6_mos_sales) %>% 
  dplyr::rename(total_last_6_mos_sales = last_6_month_sales) %>% 
  dplyr::mutate(total_last_6_mos_sales = ifelse(is.na(total_last_6_mos_sales), 0, total_last_6_mos_sales)) -> IQR_FG_sample



# vlookup - Total Last 12 mos Sales 
merge(IQR_FG_sample, sdcv[, c("ref", "last_12_month_sales")], by = "ref", all.x = TRUE) %>% 
  dplyr::select(-total_last_12_mos_sales) %>% 
  dplyr::rename(total_last_12_mos_sales = last_12_month_sales) %>% 
  dplyr::mutate(total_last_12_mos_sales = ifelse(is.na(total_last_12_mos_sales), 0, total_last_12_mos_sales)) -> IQR_FG_sample



# vlookup - Total Forecast Next 12 Months
merge(IQR_FG_sample, DSX_pivot_1[, c("ref", "total_12_month")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(total_12_month = replace(total_12_month, is.na(total_12_month), 0)) %>% 
  dplyr::select(-total_forecast_next_12_months) %>% 
  dplyr::rename(total_forecast_next_12_months = total_12_month) -> IQR_FG_sample




# vlookup - Mfg CustOrd in next 7 days
merge(IQR_FG_sample, custord_mfg_pivot_1[, c("mfg_ref", "Y")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(mfg_cust_ord_in_next_7_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - Mfg CustOrd in next 14 days
merge(IQR_FG_sample, custord_mfg_pivot_2[, c("mfg_ref", "Y")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(mfg_cust_ord_in_next_14_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - Mfg CustOrd in next 21 days
merge(IQR_FG_sample, custord_mfg_pivot_3[, c("mfg_ref", "Y")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(mfg_cust_ord_in_next_21_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - Mfg CustOrd in next 28 days
merge(IQR_FG_sample, custord_mfg_pivot_4[, c("mfg_ref", "Y")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(mfg_cust_ord_in_next_28_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - Mfg Current Month Fcst
merge(IQR_FG_sample, DSX_mfg_pivot_1[, c("mfg_ref", "Mon_a_fcst")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(mon_a_na = !is.na(Mon_a_fcst)) %>% 
  dplyr::mutate(mfg_current_month_fcst = ifelse(mon_a_na == TRUE, Mon_a_fcst, 0)) %>% 
  dplyr::select(-Mon_a_fcst, - mon_a_na) %>% 
  dplyr::mutate(mfg_current_month_fcst = round(mfg_current_month_fcst , 0)) -> IQR_FG_sample


# vlookup - Mfg Next Month Fcst
merge(IQR_FG_sample, DSX_mfg_pivot_1[, c("mfg_ref", "Mon_b_fcst")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(mon_b_na = !is.na(Mon_b_fcst)) %>% 
  dplyr::mutate(mfg_next_month_fcst = ifelse(mon_b_na == TRUE, Mon_b_fcst, 0)) %>% 
  dplyr::select(-Mon_b_fcst, - mon_b_na) %>% 
  dplyr::mutate(mfg_next_month_fcst  = round(mfg_next_month_fcst, 0)) -> IQR_FG_sample



# vlookup - Total Forecast Next 12 Months
merge(IQR_FG_sample, DSX_mfg_pivot_1[, c("mfg_ref", "total_12_month")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(total_12_month = replace(total_12_month, is.na(total_12_month), 0)) %>% 
  dplyr::select(-total_mfg_forecast_next_12_months) %>% 
  dplyr::rename(total_mfg_forecast_next_12_months = total_12_month) -> IQR_FG_sample



##################################### calculates #########################################
readr::type_convert(IQR_FG_sample) -> IQR_FG_sample


# calculate - Max Cycle Stock
IQR_FG_sample %>% 
  dplyr::mutate(max_cycle_stock = ifelse(opv == 0, pmax(cust_ord_in_next_28_days, jde_moq), pmax(jde_moq, 
                                                                                                 ifelse(opv >= 20, cust_ord_in_next_28_days,
                                                                                                        ifelse(opv >= 12 & opv < 20, cust_ord_in_next_21_days,
                                                                                                               ifelse(opv >= 8 & opv < 12, cust_ord_in_next_14_days, cust_ord_in_next_7_days))) +
                                                                                                   current_month_fcst / 20.83 * hold_days,
                                                                                                 current_month_fcst / 20.83 * (opv + hold_days),
                                                                                                 total_last_12_mos_sales / 250 * (opv + hold_days) ))) %>% 
  dplyr::mutate(max_cycle_stock = round(max_cycle_stock, 0)) -> IQR_FG_sample




# calculate - max_cycle_stock_lag_1
IQR_FG_sample %>% 
  dplyr::mutate(max_cycle_stock_lag_1 = ifelse(opv == 0, pmax(cust_ord_in_next_28_days, jde_moq), pmax(jde_moq, 
                                                                                                       ifelse(opv >= 20, cust_ord_in_next_28_days,
                                                                                                              ifelse(opv >= 12 & opv < 20, cust_ord_in_next_21_days,
                                                                                                                     ifelse(opv >= 8 & opv < 12, cust_ord_in_next_14_days, cust_ord_in_next_7_days))) +
                                                                                                         lag_1_current_month_fcst / 20.83 * hold_days,
                                                                                                       lag_1_current_month_fcst / 20.83 * (opv + hold_days) ))) %>% 
  dplyr::mutate(max_cycle_stock_lag_1 = round(max_cycle_stock_lag_1, 0)) -> IQR_FG_sample




# calculate - Max Cycle Stock Adjusted Forward 
IQR_FG_sample %>% 
  dplyr::mutate(max_cycle_stock = ifelse(opv == 0, pmax(cust_ord_in_next_28_days, jde_moq), pmax(jde_moq, ifelse(opv >= 20, cust_ord_in_next_28_days,
                                                                                                                 ifelse(opv >= 12 & opv < 20, cust_ord_in_next_21_days,
                                                                                                                        ifelse(opv >= 8 & opv < 12, cust_ord_in_next_14_days, cust_ord_in_next_7_days))) +
                                                                                                   pmax(current_month_fcst, next_month_fcst) / 20.83 * hold_days,
                                                                                                 pmax(current_month_fcst, next_month_fcst) / 20.83 * (opv + hold_days) ))) %>% 
  dplyr::mutate(max_cycle_stock = round(max_cycle_stock, 0)) -> IQR_FG_sample




# calculate - Quality hold in $$
IQR_FG_sample %>% 
  dplyr::mutate(quality_hold_2 = quality_hold * unit_cost) -> IQR_FG_sample


# calculate - On Hand (usable + soft hold)
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_usable_soft_hold = usable + soft_hold) -> IQR_FG_sample


# calculate - On Hand in pounds
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_usable_soft_hold = as.numeric(on_hand_usable_soft_hold),
                net_wt_lbs = as.numeric(net_wt_lbs)) %>% 
  dplyr::mutate(on_hand_in_pounds = on_hand_usable_soft_hold * net_wt_lbs) -> IQR_FG_sample



# calculate - On Hand in $$
IQR_FG_sample %>% 
  dplyr::mutate(on_hand = on_hand_usable_soft_hold * unit_cost) -> IQR_FG_sample



# (Business Days) calculate - Forward Inv Target Current Month Fcst (Business days input - ## Current Month ##) -----------------------------------
IQR_FG_sample %>% 
  dplyr::mutate(inventory_target = (pmax((current_month_fcst / 20) * opv, jde_moq)) / 2 + current_ss,
                inventory_target = round(inventory_target, 0)) -> IQR_FG_sample


# calculate - inventory_target_in_lbs
IQR_FG_sample %>% 
  dplyr::mutate(inventory_target_lbs = inventory_target * net_wt_lbs,
                inventory_target_lbs = round(inventory_target_lbs, 0)) -> IQR_FG_sample


# calculate - Forward Inv Target Current Month Fcst in $$
IQR_FG_sample %>% 
  dplyr::mutate(inventory_target_2 = inventory_target * unit_cost,
                inventory_target_2 = round(inventory_target_2, 2)) -> IQR_FG_sample



#calculate - forward_inv_target_lag_1_current_month_fcst
IQR_FG_sample %>% 
  dplyr::mutate(forward_inv_target_lag_1_current_month_fcst = max_cycle_stock_lag_1 / 2 + current_ss,
                forward_inv_target_lag_1_current_month_fcst = round(forward_inv_target_lag_1_current_month_fcst , 0)) -> IQR_FG_sample


# calculate - Forward Inv Target lag 1 Current Month Fcst in lbs.
IQR_FG_sample %>% 
  dplyr::mutate(forward_inv_target_lag_1_current_month_fcst_lbs = forward_inv_target_lag_1_current_month_fcst * net_wt_lbs,
                forward_inv_target_lag_1_current_month_fcst_lbs = round(forward_inv_target_lag_1_current_month_fcst_lbs, 0)) -> IQR_FG_sample


# calculate - Forward Inv Target lag 1 Current Month Fcst in $$
IQR_FG_sample %>% 
  dplyr::mutate(forward_inv_target_lag_1_current_month_fcst_2 = forward_inv_target_lag_1_current_month_fcst * unit_cost,
                forward_inv_target_lag_1_current_month_fcst_2 = round(forward_inv_target_lag_1_current_month_fcst_2, 0)) -> IQR_FG_sample


# calculate - CustOrd in next 28 days in $$
IQR_FG_sample %>% 
  dplyr::mutate(cust_ord_in_next_28_days_2 = cust_ord_in_next_28_days * unit_cost) -> IQR_FG_sample


# calculate - dos
IQR_FG_sample %>% 
  dplyr::mutate(dos = on_hand_usable_soft_hold / pmax((ifelse(opv == 0 | opv >= 20, 
                                                              cust_ord_in_next_28_days, 
                                                              ifelse(opv < 20  & opv >= 12, cust_ord_in_next_21_days,
                                                                     ifelse(opv < 12  & opv >= 8, cust_ord_in_next_14_days, cust_ord_in_next_7_days))) / opv), 
                                                      (pmax(current_month_fcst, next_month_fcst) / 20.83) ),
                dos_na = !is.na(dos),
                dos = ifelse(dos_na == TRUE, dos, 0),
                dos = round(dos, 1),
                dos = replace(dos, is.infinite(dos), 0)) %>% 
  dplyr::select(-dos_na) -> IQR_FG_sample





# calculate - dos after CustOrd
IQR_FG_sample %>% 
  dplyr::mutate(dos_after_cust_ord = (on_hand_usable_soft_hold - ifelse(opv == 0 | opv >= 20, cust_ord_in_next_28_days, 
                                                                        ifelse(opv < 20 & opv >= 12, cust_ord_in_next_21_days, 
                                                                               ifelse(opv < 12 & opv >= 8, cust_ord_in_next_14_days, cust_ord_in_next_7_days))))/
                  pmax((ifelse(opv == 0 | opv >= 20, cust_ord_in_next_28_days, 
                               ifelse(opv < 20 & opv >= 12, cust_ord_in_next_21_days,
                                      ifelse(opv < 12 & opv >= 8, cust_ord_in_next_14_days, cust_ord_in_next_7_days)))/opv),
                       (pmax(current_month_fcst, next_month_fcst)/20.83))) %>% 
  dplyr::mutate(dos_na = !is.na(dos_after_cust_ord),
                dos_after_cust_ord = ifelse(dos_na == TRUE, dos_after_cust_ord, 0),
                dos_after_cust_ord = round(dos_after_cust_ord, 1),
                dos_after_cust_ord = replace(dos_after_cust_ord, is.infinite(dos_after_cust_ord), 0)) %>% 
  dplyr::select(-dos_na) %>% 
  dplyr::mutate(dos_after_cust_ord = round(dos_after_cust_ord, 1)) -> IQR_FG_sample


# Max Inv Dos
IQR_FG_sample %>% 
  mutate(
    max_inv_dos = if_else(
      is.na(max_inventory_target) | is.na(opv) | is.na(cust_ord_in_next_7_days) | is.na(cust_ord_in_next_14_days) | is.na(cust_ord_in_next_21_days) | is.na(cust_ord_in_next_28_days) | is.na(current_month_fcst) | is.na(next_month_fcst), 
      0,
      max_inventory_target / max(
        case_when(
          opv == 0 | opv >= 20 ~ cust_ord_in_next_28_days / opv,
          opv < 20 & opv >= 12 ~ cust_ord_in_next_21_days / opv,
          opv < 12 & opv >= 8 ~ cust_ord_in_next_14_days / opv,
          TRUE ~ cust_ord_in_next_7_days / opv
        ),
        max(current_month_fcst, next_month_fcst) / 20.83
      )
    )
  ) -> IQR_FG_sample




# calculate - target_inv_dos_includes_orders
IQR_FG_sample %>% 
  dplyr::mutate(aa = current_ss / pmax((ifelse(opv == 0 | opv >= 20, cust_ord_in_next_28_days,
                                               ifelse(opv < 20 & opv >= 12, cust_ord_in_next_21_days,
                                                      ifelse(opv < 12 & opv >= 8, cust_ord_in_next_14_days, cust_ord_in_next_7_days))) / opv),
                                       (pmax(current_month_fcst, next_month_fcst) / 20.83)),
                aa = replace(aa, is.na(aa), 0),
                aa = replace(aa, is.nan(aa), 0),
                aa = replace(aa, is.infinite(aa), 0)) %>% 
  dplyr::mutate(target_inv_dos_includes_orders = ifelse(total_forecast_next_12_months == 0, 0, 
                                                        opv + aa)) %>% 
  dplyr::mutate(target_inv_dos_includes_orders = round(target_inv_dos_includes_orders, 1)) %>% 
  dplyr::select(-aa) -> IQR_FG_sample



# calculate - on hand Inv after CustOrd > AF max
IQR_FG_sample %>%
  mutate(
    on_hand_inv_after_cust_ord_af_max = if_else(
      is.na(on_hand_usable_soft_hold) | is.na(max_inventory_target) | is.na(opv) | is.na(cust_ord_in_next_7_days) | is.na(cust_ord_in_next_14_days) | is.na(cust_ord_in_next_21_days) | is.na(cust_ord_in_next_28_days),
      NA_real_,
      if_else(
        on_hand_usable_soft_hold - case_when(
          opv == 0 | opv >= 20 ~ cust_ord_in_next_28_days,
          opv < 20 & opv >= 12 ~ cust_ord_in_next_21_days,
          opv < 12 & opv >= 8 ~ cust_ord_in_next_14_days,
          TRUE ~ cust_ord_in_next_7_days) > max_inventory_target,1,0))) -> IQR_FG_sample




# calculate - Inv Health
IQR_FG_sample %>% 
  dplyr::mutate(inv_health = ifelse(on_hand_usable_soft_hold < current_ss, "BELOW SS",
                                    ifelse(dos_after_cust_ord > shippable_shelf_life, "AT RISK" ,
                                           ifelse(total_forecast_next_12_months <= 0 & cust_ord_in_next_28_days <= 0,
                                                  ifelse(on_hand_usable_soft_hold > 0, "DEAD",
                                                         ifelse(on_hand_inv_after_cust_ord_af_max == 0, "HEALTHY", "EXCESS")),
                                                  ifelse(on_hand_inv_after_cust_ord_af_max == 1, "EXCESS", "HEALTHY"))))) -> IQR_FG_sample


# calculate - Lag 1 Current Month Fcst in cost
IQR_FG_sample %>% 
  dplyr::mutate(lag_1_current_month_fcst_2 = lag_1_current_month_fcst * unit_cost) -> IQR_FG_sample


# calculate - has adjusted forward looking Max?
IQR_FG_sample %>% 
  dplyr::mutate(has_adjusted_forward_looking_max = ifelse(max_inventory_target > 0, 1, 0)) -> IQR_FG_sample


# calculate - on hand Inv > AF max
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_af_max = ifelse(on_hand_usable_soft_hold > max_inventory_target, 1, 0)) -> IQR_FG_sample


# calculate - on hand Inv <= AF max
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_af_max_2 = ifelse(on_hand_usable_soft_hold <= max_inventory_target, 1, 0)) -> IQR_FG_sample


# calculate - on hand Inv > Adjusted Forward looking target
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_adjusted_forward_looking_target = ifelse(on_hand_usable_soft_hold > max_inventory_target, 1, 0)) -> IQR_FG_sample


# calculate - on hand Inv <= AF target
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_af_target = ifelse(on_hand_usable_soft_hold <= max_inventory_target, 1, 0)) -> IQR_FG_sample


# calculate - on hand Inv after CustOrd <= AF max
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_after_cust_ord_af_max_2 = ifelse(on_hand_usable_soft_hold - (ifelse(opv == 0 | opv >= 20, cust_ord_in_next_28_days,
                                                                                                ifelse(opv < 20 & opv >= 12, cust_ord_in_next_21_days,
                                                                                                       ifelse(opv < 12 & opv >= 8, cust_ord_in_next_14_days, cust_ord_in_next_7_days))))
                                                             <= max_inventory_target , 1, 0)) -> IQR_FG_sample


# calculate - on hand Inv after CustOrd > AF target
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_after_cust_ord_af_target = ifelse(on_hand_usable_soft_hold - (ifelse(opv == 0 | opv >= 20, cust_ord_in_next_28_days,
                                                                                                 ifelse(opv < 20 & opv >= 12, cust_ord_in_next_21_days,
                                                                                                        ifelse(opv < 12 & opv >= 8, cust_ord_in_next_14_days, cust_ord_in_next_7_days))))
                                                              > max_inventory_target, 1, 0)) -> IQR_FG_sample



# calculate - on hand Inv after CustOrd <= AF target
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_after_cust_ord_af_target_2 = ifelse(on_hand_usable_soft_hold - (ifelse(opv == 0 | opv >= 20, cust_ord_in_next_28_days,
                                                                                                   ifelse(opv < 20 & opv >= 12, cust_ord_in_next_21_days,
                                                                                                          ifelse(opv < 12 & opv >= 8, cust_ord_in_next_14_days, cust_ord_in_next_7_days))))
                                                                <= max_inventory_target, 1, 0)) -> IQR_FG_sample


# calculate - on hand inv after 28 days CustOrd > 0
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_after_28_days_cust_ord_0 = ifelse(on_hand_usable_soft_hold - cust_ord_in_next_28_days > 0, 1,0)) -> IQR_FG_sample



# calculate - Max Cycle Stock Mfg Adjusted Forward
IQR_FG_sample %>%
  dplyr::mutate(max_cycle_stock_mfg_adjusted_forward = ifelse(opv == 0,pmax(mfg_cust_ord_in_next_28_days, jde_moq),
                                                              pmax(jde_moq, ifelse(opv >= 20, mfg_cust_ord_in_next_28_days,
                                                                                   ifelse(opv >= 12 & opv < 20, mfg_cust_ord_in_next_21_days,
                                                                                          ifelse(opv >= 8 & opv < 12, mfg_cust_ord_in_next_14_days, mfg_cust_ord_in_next_7_days)))
                                                                   + pmax(mfg_current_month_fcst, mfg_next_month_fcst) / 20.83 * hold_days, pmax(mfg_current_month_fcst, mfg_next_month_fcst) / 20.83 *
                                                                     (opv+hold_days)))) %>%
  dplyr::mutate(max_cycle_stock_mfg_adjusted_forward = round(max_cycle_stock_mfg_adjusted_forward, 0)) -> IQR_FG_sample




# calculate - Mfg Adjusted Forward Inv Max
IQR_FG_sample %>% 
  dplyr::mutate(mfg_adjusted_forward_inv_max = max_cycle_stock_mfg_adjusted_forward + current_ss) %>% 
  dplyr::mutate(mfg_adjusted_forward_inv_max = round(mfg_adjusted_forward_inv_max, 0)) -> IQR_FG_sample




# calculate - On Hand - Mfg Adjusted Forward Max in $$
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_mfg_adjusted_forward_max = ifelse((on_hand_usable_soft_hold - mfg_adjusted_forward_inv_max) * 
                                                            unit_cost < 0, 0, (on_hand_usable_soft_hold - 
                                                                                 mfg_adjusted_forward_inv_max) * unit_cost) )  -> IQR_FG_sample



# calculate - Mfg Adjusted Forward Inv Target
IQR_FG_sample %>% 
  dplyr::mutate(mfg_adjusted_forward_inv_target = max_cycle_stock_mfg_adjusted_forward / 2 + current_ss) %>% 
  dplyr::mutate(mfg_adjusted_forward_inv_target = round(mfg_adjusted_forward_inv_target, 0)) -> IQR_FG_sample



# calculate - Mfg Adjusted Forward Inv Target in lbs.
IQR_FG_sample %>% 
  dplyr::mutate(mfg_adjusted_forward_inv_target_lbs = mfg_adjusted_forward_inv_target * net_wt_lbs) %>% 
  dplyr::mutate(mfg_adjusted_forward_inv_target_lbs = round(mfg_adjusted_forward_inv_target_lbs, 2))-> IQR_FG_sample


# calculate - Mfg Adjusted Forward Inv Target in $$
IQR_FG_sample %>% 
  dplyr::mutate(mfg_adjusted_forward_inv_target_2 = mfg_adjusted_forward_inv_target * unit_cost) %>% 
  dplyr::mutate(mfg_adjusted_forward_inv_target_2 = round(mfg_adjusted_forward_inv_target_2, 2)) -> IQR_FG_sample


# calculate - Mfg Adjusted Forward Inv Max in lbs.
IQR_FG_sample %>% 
  dplyr::mutate(mfg_adjusted_forward_inv_max_lbs = mfg_adjusted_forward_inv_max * net_wt_lbs) %>% 
  dplyr::mutate(mfg_adjusted_forward_inv_max_lbs = round(mfg_adjusted_forward_inv_max_lbs, 2)) -> IQR_FG_sample


# calculate - Mfg Adjusted Forward Inv Max in $$
IQR_FG_sample %>% 
  dplyr::mutate(mfg_adjusted_forward_inv_max_2 = mfg_adjusted_forward_inv_max * unit_cost) %>% 
  dplyr::mutate(mfg_adjusted_forward_inv_max_2 = round(mfg_adjusted_forward_inv_max_2, 2)) -> IQR_FG_sample


# calculate - Mfg CustOrd in next 28 days in $$
IQR_FG_sample %>% 
  dplyr::mutate(mfg_cust_ord_in_next_28_days_2 = mfg_cust_ord_in_next_28_days * unit_cost) %>% 
  dplyr::mutate(mfg_cust_ord_in_next_28_days_2 = round(mfg_cust_ord_in_next_28_days_2, 2)) -> IQR_FG_sample


# calculate - Mfg dos 

# test
IQR_FG_sample %>% 
  dplyr::mutate(mfg_dos = on_hand_usable_soft_hold / pmax((ifelse(opv == 0 | opv >= 20, 
                                                                  mfg_cust_ord_in_next_28_days, 
                                                                  ifelse(opv < 20  & opv >= 12, mfg_cust_ord_in_next_21_days,
                                                                         ifelse(opv < 12  & opv >= 8, mfg_cust_ord_in_next_14_days, 
                                                                                mfg_cust_ord_in_next_7_days))) / opv), 
                                                          (pmax(mfg_current_month_fcst, mfg_next_month_fcst) / 20.83) ),
                dos_mfg_na = !is.na(mfg_dos),
                mfg_dos = ifelse(dos_mfg_na == TRUE, mfg_dos, 0),
                mfg_dos = round(mfg_dos, 1),
                mfg_dos = replace(mfg_dos, is.infinite(mfg_dos), 0)) %>% 
  dplyr::select(-dos_mfg_na) -> IQR_FG_sample


# calculate - Mfg dos after CustOrd 
IQR_FG_sample %>% 
  dplyr::mutate(mfg_dos_after_cust_ord = (on_hand_usable_soft_hold - ifelse(opv == 0 | opv >= 20, mfg_cust_ord_in_next_28_days, 
                                                                            ifelse(opv < 20 & opv >= 12, mfg_cust_ord_in_next_21_days, 
                                                                                   ifelse(opv < 12 & opv >= 8, mfg_cust_ord_in_next_14_days, 
                                                                                          mfg_cust_ord_in_next_7_days))))/
                  pmax((ifelse(opv == 0 | opv >= 20, mfg_cust_ord_in_next_28_days, 
                               ifelse(opv < 20 & opv >= 12, mfg_cust_ord_in_next_21_days,
                                      ifelse(opv < 12 & opv >= 8, mfg_cust_ord_in_next_14_days, mfg_cust_ord_in_next_7_days)))/opv),
                       (pmax(mfg_current_month_fcst, mfg_next_month_fcst)/20.83))) %>% 
  dplyr::mutate(dos_mfg_na = !is.na(mfg_dos_after_cust_ord),
                mfg_dos_after_cust_ord = ifelse(dos_mfg_na == TRUE, mfg_dos_after_cust_ord, 0),
                mfg_dos_after_cust_ord = round(mfg_dos_after_cust_ord, 1),
                mfg_dos_after_cust_ord = replace(mfg_dos_after_cust_ord, is.infinite(mfg_dos_after_cust_ord), 0)) %>% 
  dplyr::select(-dos_mfg_na) %>% 
  dplyr::mutate(mfg_dos_after_cust_ord = round(mfg_dos_after_cust_ord, 1)) -> IQR_FG_sample


# calculate - mfg_adjusted_forward_max_inv_dos
IQR_FG_sample %>% 
  dplyr::mutate(mfg_adjusted_forward_max_inv_dos =  (mfg_adjusted_forward_inv_target / 
                                                       pmax((ifelse(opv == 0 | opv >= 20, mfg_cust_ord_in_next_28_days,
                                                                    ifelse(opv < 20 & opv >= 12, mfg_cust_ord_in_next_21_days,
                                                                           ifelse(opv < 12 & opv >= 8, mfg_cust_ord_in_next_14_days, 
                                                                                  mfg_cust_ord_in_next_7_days)))
                                                             / opv),
                                                            (pmax(mfg_current_month_fcst, mfg_next_month_fcst) / 20.83)))   ) %>% 
  dplyr::mutate(dos_mfg_na = !is.na(mfg_adjusted_forward_max_inv_dos),
                mfg_adjusted_forward_max_inv_dos = ifelse(dos_mfg_na == TRUE, mfg_adjusted_forward_max_inv_dos, 0),
                mfg_adjusted_forward_max_inv_dos = round(mfg_adjusted_forward_max_inv_dos, 1),
                mfg_adjusted_forward_max_inv_dos = replace(mfg_adjusted_forward_max_inv_dos, is.infinite(mfg_adjusted_forward_max_inv_dos), 0)) %>% 
  dplyr::select(-dos_mfg_na) %>% 
  dplyr::mutate(mfg_adjusted_forward_max_inv_dos = round(mfg_adjusted_forward_max_inv_dos, 1)) -> IQR_FG_sample



# calculate - Mfg Forward Target Inv dos (fcst only)
IQR_FG_sample %>% 
  dplyr::mutate(aa = current_ss / (pmax(mfg_current_month_fcst, mfg_next_month_fcst) / 20.83),
                aa = replace(aa, is.na(aa), 0),
                aa = replace(aa, is.nan(aa), 0),
                aa = replace(aa, is.infinite(aa), 0)) %>% 
  dplyr::mutate(mfg_forward_target_inv_dos_fcst_only = ifelse(total_mfg_forecast_next_12_months == 0, 0,
                                                              opv + aa)) %>% 
  dplyr::mutate(mfg_forward_target_inv_dos_fcst_only = round(mfg_forward_target_inv_dos_fcst_only, 0)) %>% 
  dplyr::select(-aa) -> IQR_FG_sample



# calculate - Mfg Adjusted Forward Target Inv dos (includes Orders)
IQR_FG_sample %>% 
  dplyr::mutate(aa = current_ss / pmax((ifelse(opv == 0 | opv >= 20, mfg_cust_ord_in_next_28_days,
                                               ifelse(opv < 20 & opv >= 12, mfg_cust_ord_in_next_21_days,
                                                      ifelse(opv < 12 & opv >= 8, mfg_cust_ord_in_next_14_days, mfg_cust_ord_in_next_7_days))) / opv),
                                       (pmax(current_month_fcst, next_month_fcst) / 20.83)),
                aa = replace(aa, is.na(aa), 0),
                aa = replace(aa, is.nan(aa), 0),
                aa = replace(aa, is.infinite(aa), 0)) %>% 
  dplyr::mutate(mfg_adjusted_forward_target_inv_dos_includes_orders = ifelse(total_mfg_forecast_next_12_months == 0, 0, 
                                                                             opv + aa)) %>% 
  dplyr::mutate(mfg_adjusted_forward_target_inv_dos_includes_orders = round(mfg_adjusted_forward_target_inv_dos_includes_orders, 0)) %>% 
  dplyr::select(-aa) -> IQR_FG_sample


# on hand Inv after CustOrd > mfg AF max
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_after_cust_ord_mfg_af_max =  ifelse(on_hand_usable_soft_hold-
                                                                  (ifelse(opv == 0 | opv >= 20, mfg_cust_ord_in_next_28_days,
                                                                          ifelse(opv < 20 & opv >= 12, mfg_cust_ord_in_next_21_days,
                                                                                 ifelse(opv < 12 & opv >= 8, mfg_cust_ord_in_next_14_days, mfg_cust_ord_in_next_7_days))))
                                                                > mfg_adjusted_forward_inv_max, 1, 0)) -> IQR_FG_sample


# calculate - Mfg Inv Health
IQR_FG_sample %>% 
  dplyr::mutate(mfg_inv_health = ifelse(on_hand_usable_soft_hold < current_ss, "BELOW SS",
                                        ifelse(mfg_dos_after_cust_ord > shippable_shelf_life, "AT RISK",
                                               ifelse(total_mfg_forecast_next_12_months <= 0 & mfg_cust_ord_in_next_28_days <= 0,
                                                      ifelse(on_hand_usable_soft_hold > 0, "DEAD",
                                                             ifelse(on_hand_inv_after_cust_ord_mfg_af_max == 0, "HEALTHY", "EXCESS")),
                                                      ifelse(on_hand_inv_after_cust_ord_mfg_af_max == 1, "EXCESS", "HEALTHY"))))) -> IQR_FG_sample




# calculate - has mfg adjusted forward looking Max?
IQR_FG_sample %>% 
  dplyr::mutate(has_mfg_adjusted_forward_looking_max = ifelse(mfg_adjusted_forward_inv_max > 0, 1, 0)) -> IQR_FG_sample




# calculate - on hand Inv > mfg AF max
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_mfg_af_max = ifelse(on_hand_usable_soft_hold > mfg_adjusted_forward_inv_max, 1, 0)) -> IQR_FG_sample


# calculate - on hand Inv <= mfg AF max
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_mfg_af_max_2 = ifelse(on_hand_usable_soft_hold <= mfg_adjusted_forward_inv_max, 1, 0)) -> IQR_FG_sample



# calculate - on hand Inv > mfg Adjusted Forward looking target
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_mfg_adjusted_forward_looking_target = ifelse(on_hand_usable_soft_hold > 
                                                                           mfg_adjusted_forward_inv_target, 1, 0)) -> IQR_FG_sample


# calculate - on hand Inv <= mfg AF target
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_mfg_af_target = ifelse(on_hand_usable_soft_hold <= 
                                                     mfg_adjusted_forward_inv_target, 1, 0)) -> IQR_FG_sample



# calculate - on hand Inv after CustOrd <= mfg AF max
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_after_cust_ord_mfg_af_max_2 =  
                  ifelse(on_hand_usable_soft_hold - (ifelse(opv == 0 | opv >= 20, mfg_cust_ord_in_next_28_days,
                                                            ifelse(opv < 20 & opv >= 12, mfg_cust_ord_in_next_21_days,
                                                                   ifelse(opv < 12 & opv >= 8, mfg_cust_ord_in_next_14_days, mfg_cust_ord_in_next_7_days))))
                         <= mfg_adjusted_forward_inv_max, 1,0)) -> IQR_FG_sample




# calculate - on hand Inv after CustOrd > mfg AF target
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_after_cust_ord_mfg_af_target = ifelse(on_hand_usable_soft_hold-(
    ifelse(opv == 0 | opv >= 20, mfg_cust_ord_in_next_28_days,
           ifelse(opv < 20 & opv >= 12, mfg_cust_ord_in_next_21_days,
                  ifelse(opv < 12 & opv >= 8, mfg_cust_ord_in_next_14_days, mfg_cust_ord_in_next_7_days)))) > mfg_adjusted_forward_inv_target,1,0)) -> IQR_FG_sample



# calculate - on hand Inv after CustOrd <= mfg AF target
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_inv_after_cust_ord_mfg_af_target_2 = 
                  ifelse(on_hand_usable_soft_hold-(
                    ifelse(opv == 0 | opv >= 20, mfg_cust_ord_in_next_28_days,
                           ifelse(opv < 20 & opv >= 12, mfg_cust_ord_in_next_21_days,
                                  ifelse(opv < 12 & opv >= 8, mfg_cust_ord_in_next_14_days, mfg_cust_ord_in_next_7_days))))
                    <= mfg_adjusted_forward_inv_target,1,0)) -> IQR_FG_sample



# calculate - on hand inv after mfg 28 days CustOrd > 0
IQR_FG_sample %>%
  dplyr::mutate(on_hand_inv_after_mfg_28_days_cust_ord_0 =
                  ifelse(on_hand_usable_soft_hold - mfg_cust_ord_in_next_28_days > 0, 1, 0)) -> IQR_FG_sample




# Planner Name N/A
IQR_FG_sample %>% 
  dplyr::mutate(planner_name = ifelse(is.na(planner_name) & planner == 0, 0, planner_name)) -> IQR_FG_sample


# pre_data
pre_data <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/06.25.2024/iqr_fg_rstudio_06252024.xlsx")

pre_data %>% 
  janitor::clean_names() %>% 
  dplyr::select(item_2, category, platform, macro_platform) %>% 
  dplyr::rename(category_2 = category,
                platform_2 = platform,
                macro_platform_2 = macro_platform) -> pre_data

pre_data[!duplicated(pre_data[,c("item_2")]),] -> pre_data


IQR_FG_sample %>% 
  dplyr::left_join(pre_data) %>% 
  dplyr::mutate(category = ifelse(is.na(category), category_2, category),
                platform = ifelse(is.na(platform), platform_2, platform),
                macro_platform = ifelse(is.na(macro_platform), macro_platform_2, macro_platform)) %>% 
  dplyr::select(-category_2, -platform_2, -macro_platform_2) -> IQR_FG_sample


# Category & Platform
completed_sku_list <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/07.02.2024/Completed SKU list - Linda.xlsx")
completed_sku_list[-1:-2, ]  %>% 
  janitor::clean_names() %>% 
  dplyr::select(x6, x7, x9, x11) %>% 
  dplyr::rename(Parent_Item_Number = x6,
                Category = x9,
                Platform = x11) %>% 
  dplyr::mutate(Parent_Item_Number = gsub("-", "", Parent_Item_Number)) -> completed_sku_list

completed_sku_list[!duplicated(completed_sku_list[,c("Parent_Item_Number")]),] -> completed_sku_list

completed_sku_list %>% 
  dplyr::select(Parent_Item_Number, Category) %>% 
  dplyr::rename(item_2 = Parent_Item_Number,
                category = Category)-> completed_sku_list_category


completed_sku_list %>% 
  dplyr::select(Parent_Item_Number, Platform) %>% 
  dplyr::rename(item_2 = Parent_Item_Number,
                platform = Platform)-> completed_sku_list_platform


IQR_FG_sample %>% 
  dplyr::select(-category, -platform) %>% 
  dplyr::left_join(completed_sku_list_category) %>% 
  dplyr::left_join(completed_sku_list_platform) -> IQR_FG_sample


macro_platform[!duplicated(macro_platform[,c("Platform")]),] -> macro_platform

macro_platform %>% 
  dplyr::rename(platform = Platform,
                macro_platform = Macro_Platform) -> macro_platform

IQR_FG_sample %>% 
  dplyr::select(-macro_platform) %>% 
  dplyr::left_join(macro_platform) -> IQR_FG_sample

# On Priority List
priority_sku <- read_excel("S:/Supply Chain Projects/RStudio/Priority_Sku_and_uniques.xlsx",
                           col_names = FALSE)

colnames(priority_sku) <- priority_sku[1, ]
priority_sku[-1, ] -> priority_sku

colnames(priority_sku)[1] <- "priority_sku"

priority_sku %>% 
  dplyr::mutate(Item_2 = priority_sku) %>% 
  dplyr::rename(on_priority_list = priority_sku,
                item_2 = Item_2)-> priority_sku


IQR_FG_sample %>% 
  dplyr::left_join(priority_sku) %>% 
  dplyr::mutate(on_priority_list = ifelse(is.na(on_priority_list), "N", "Y")) -> IQR_FG_sample


############################## Added on 9/12/2023 ################################## NEW TEMPLATE ##

# Current SS $
IQR_FG_sample %>% 
  dplyr::mutate(current_ss_2 = current_ss * unit_cost) -> IQR_FG_sample


# Max Cycle Stock $
IQR_FG_sample %>% 
  dplyr::mutate(max_cycle_stock_mfg_adjusted_forward = max_cycle_stock * unit_cost) -> IQR_FG_sample


# Max Inventory Target
IQR_FG_sample %>% 
  dplyr::mutate(max_inventory_target = current_ss + max_cycle_stock) -> IQR_FG_sample


# Max Inventory Target (lbs.)
IQR_FG_sample %>% 
  dplyr::mutate(max_inventory_target_lbs = max_inventory_target * net_wt_lbs) -> IQR_FG_sample

# Max Inventory Target $
IQR_FG_sample %>% 
  dplyr::mutate(max_inventory_target_2 = max_inventory_target * unit_cost) -> IQR_FG_sample

# Current Month Fcst $
IQR_FG_sample %>% 
  dplyr::mutate(current_month_fcst_2 = current_month_fcst * unit_cost) -> IQR_FG_sample

# Next Month Fcst $
IQR_FG_sample %>% 
  dplyr::mutate(next_month_fcst_2 = next_month_fcst * unit_cost) -> IQR_FG_sample



############################################ Custord ###################################################
custord %>% 
  dplyr::select(1:5) %>% 
  dplyr::mutate(ref = gsub("_", "-", ref))-> custord_open_order

########################################################################################################

############################################ Added 9/21/23 #############################################
custord_open_order %>% 
  dplyr::mutate(ref = gsub("-", "_", ref)) %>% 
  dplyr::select(ref, Qty) %>% 
  dplyr::rename(open_orders_all = Qty) -> custord_for_merge

custord_for_merge %>% 
  dplyr::group_by(ref) %>% 
  dplyr::summarise(open_orders_all = sum(open_orders_all)) %>% 
  dplyr::ungroup()-> custord_for_merge



IQR_FG_sample %>% 
  dplyr::select(-open_orders_all) %>% 
  dplyr::left_join(custord_for_merge) %>% 
  dplyr::mutate(open_orders_all = ifelse(is.na(open_orders_all), 0, open_orders_all)) -> IQR_FG_sample


############################################ Added 12/12/23 #############################################
# Description 
IQR_FG_sample %>% 
  dplyr::select(-description) %>% 
  dplyr::left_join(completed_sku_list %>% select(Parent_Item_Number, x7) %>% rename(item_2 = Parent_Item_Number, description = x7)) %>%
  dplyr::relocate(description, .after = label) -> IQR_FG_sample


# relocation cvm, focus_label
IQR_FG_sample %>% 
  dplyr::select(-focus_label) %>% 
  dplyr::rename(focus_label = on_priority_list) %>%
  dplyr::relocate(cvm, .after = sub_type) %>% 
  dplyr::relocate(focus_label, .after = cvm) -> IQR_FG_sample


############################################ Added 05/09/23 #############################################
IQR_FG_sample %>% 
  dplyr::mutate(category = ifelse(is.na(category), "NA", category),
                platform= ifelse(is.na(platform), "NA", platform),
                macro_platform = ifelse(is.na(macro_platform), "NA", macro_platform)) %>% 
  dplyr::select(-Category) -> IQR_FG_sample


#### Added 05/21/2024 #### DOU Exception Report ####

exception_report_dou <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/JDE DNRR Exception report extract/2024/exception report DOU 2024.07.02.xlsx")
exception_report_dou %>% 
  janitor::clean_names() %>% 
  dplyr::slice(-1:-2) -> exception_report_dou

colnames(exception_report_dou) <- exception_report_dou[1, ]
exception_report_dou[-1, ] -> exception_report_dou

exception_report_dou %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(ref = paste0(b_p, "_", item_number)) %>% 
  dplyr::select(ref, planner) %>% 
  dplyr::distinct() -> exception_report_dou



IQR_FG_sample %>%
  left_join(exception_report_dou %>%
              filter(ref %in% IQR_FG_sample$ref[which(IQR_FG_sample$planner == "DNRR")]), by = "ref") %>%
  dplyr::mutate(planner = ifelse(planner.x == "DNRR", planner.y, planner.x)) %>% 
  dplyr::select(-planner.x, -planner.y, -planner_name) %>%
  dplyr::left_join(Planner_address %>% rename(planner_name = Alpha_Name)) %>% 
  dplyr::mutate(planner_name = ifelse(planner == "DNRR", "DNRR", planner_name)) %>% 
  dplyr::mutate(planner = ifelse(is.na(planner) | planner == 0, "DNRR", planner),
                planner_name = ifelse(is.na(planner_name) | planner_name == 0, "DNRR", planner_name)) -> IQR_FG_sample









########################################################################################################

# Arrange ----
fg_data_for_arrange <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/07.02.2024/Finished Goods Inventory Health Adjusted Forward (IQR) NEW TEMPLATE - 07.02.2024.xlsx",
                                  sheet = "Location FG")

fg_data_for_arrange[-1:-2, ] -> fg_data_for_arrange
colnames(fg_data_for_arrange) <- fg_data_for_arrange[1, ]
fg_data_for_arrange[-1, ] -> fg_data_for_arrange

fg_data_for_arrange %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::select(ref) %>%
  dplyr::mutate(ref = gsub("-", "_", ref)) %>% 
  dplyr::mutate(arrange = row_number()) -> fg_data_for_arrange

IQR_FG_sample %>% 
  dplyr::left_join(fg_data_for_arrange) %>% 
  dplyr::arrange(arrange) %>% 
  dplyr::select(-arrange)-> IQR_FG_sample





####################################### transform to original format ####################################


IQR_FG_sample %>% 
  dplyr::mutate(ref = gsub("_", "-", ref),
                loc_sku = gsub("_", "-", loc_sku),
                mfg_ref = gsub("_", "-", mfg_ref)) %>%
  dplyr::rename(campus_ref = loc_sku) %>% 
  dplyr::relocate(loc, mfg_loc, campus, item_2, category, platform, macro_platform, sub_type, cvm, focus_label, ref, 
                  mfg_ref, campus_ref, base, label, description, mto_mts, mpf, planner, planner_name, qty_per_pallet, storage_condition, pack_size, formula, 
                  net_wt_lbs, unit_cost, jde_moq,
                  shippable_shelf_life, hold_days, current_ss, current_ss_2, max_cycle_stock_lag_1, max_cycle_stock, 
                  max_cycle_stock_mfg_adjusted_forward, 
                  max_cycle_stock_2,
                  usable, quality_hold, quality_hold_2, soft_hold, 
                  on_hand_usable_soft_hold, on_hand_in_pounds, on_hand, on_hand_max, on_hand_mfg_adjusted_forward_max, inventory_target, 
                  inventory_target_lbs, inventory_target_2,
                  max_inventory_target, max_inventory_target_lbs, max_inventory_target_2,
                  mfg_adjusted_forward_inv_target, mfg_adjusted_forward_inv_target_lbs,
                  mfg_adjusted_forward_inv_target_2, 
                  mfg_adjusted_forward_inv_max, mfg_adjusted_forward_inv_max_lbs,
                  mfg_adjusted_forward_inv_max_2, 
                  forward_inv_target_lag_1_current_month_fcst,
                  forward_inv_target_lag_1_current_month_fcst_lbs, forward_inv_target_lag_1_current_month_fcst_2,
                  opv, cust_ord_in_next_7_days, cust_ord_in_next_14_days, cust_ord_in_next_21_days, cust_ord_in_next_28_days,
                  cust_ord_in_next_28_days_2, mfg_cust_ord_in_next_7_days, mfg_cust_ord_in_next_14_days,
                  mfg_cust_ord_in_next_21_days, mfg_cust_ord_in_next_28_days, mfg_cust_ord_in_next_28_days_2,
                  firm_wo_in_next_28_days, receipt_in_the_next_28_days, dos, 
                  dos_after_cust_ord, max_inv_dos,
                  target_inv_dos_includes_orders, mfg_dos, mfg_dos_after_cust_ord,
                  mfg_adjusted_forward_max_inv_dos, mfg_forward_target_inv_dos_fcst_only, 
                  mfg_adjusted_forward_target_inv_dos_includes_orders, inv_health, mfg_inv_health, 
                  lag_1_current_month_fcst, lag_1_current_month_fcst_2, current_month_fcst, next_month_fcst, mfg_current_month_fcst,
                  mfg_next_month_fcst, total_last_6_mos_sales, total_last_12_mos_sales, total_forecast_next_12_months,
                  total_mfg_forecast_next_12_months, 
                  has_adjusted_forward_looking_max, 
                  on_hand_inv_af_max, on_hand_inv_af_max_2, on_hand_inv_adjusted_forward_looking_target, on_hand_inv_af_target, on_hand_inv_after_cust_ord_af_max,
                  on_hand_inv_after_cust_ord_af_max_2, on_hand_inv_after_cust_ord_af_target, on_hand_inv_after_cust_ord_af_target_2, 
                  on_hand_inv_after_28_days_cust_ord_0, has_mfg_adjusted_forward_looking_max, on_hand_inv_mfg_af_max, on_hand_inv_mfg_af_max_2,
                  on_hand_inv_mfg_adjusted_forward_looking_target, on_hand_inv_mfg_af_target, on_hand_inv_after_cust_ord_mfg_af_max, 
                  on_hand_inv_after_cust_ord_mfg_af_max_2, on_hand_inv_after_cust_ord_mfg_af_target, on_hand_inv_after_cust_ord_mfg_af_target_2,
                  on_hand_inv_after_mfg_28_days_cust_ord_0, current_month_fcst_2, next_month_fcst_2, open_orders_all) -> IQR_FG_sample



# (Path Revision Needed)
writexl::write_xlsx(IQR_FG_sample, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/07.02.2024/iqr_fg_rstudio_07022024.xlsx")
writexl::write_xlsx(custord_open_order, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/07.02.2024/Open Order.xlsx")





######################################################################################################################################################

#### DOS File Moving from RM. 
file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Weekly Report run/2024/07.02.2024/Inventory Health (IQR) Tracker - DOS.xlsx",
          "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/07.02.2024/Inventory Health (IQR) Tracker - DOS.xlsx")



#### IQR main file Moving to S Drive. 
file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/07.02.2024/Finished Goods Inventory Health Adjusted Forward (IQR) NEW TEMPLATE - 07.02.2024.xlsx",
          "S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Finished Goods Inventory Health Adjusted Forward (IQR) NEW TEMPLATE - 07.02.2024.xlsx",
          overwrite = TRUE)



### DOS Tracker moving to S Drive
file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/07.02.2024/Inventory Health (IQR) Tracker - DOS.xlsx",
          "S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Inventory Health (IQR) Tracker - DOS.xlsx",
          overwrite = TRUE)

### IQR pre week moving in S Drive
file.copy("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Finished Goods Inventory Health Adjusted Forward (IQR) NEW TEMPLATE - 06.25.2024.xlsx",
          "S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2024/Finished Goods Inventory Health Adjusted Forward (IQR) NEW TEMPLATE - 06.25.2024.xlsx",
          overwrite = TRUE)


######################################################### Do this once a month for a pre actual sales ########################################################
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/F9E860846E4878AEF29B5E949CF675DC/K53--K46
# Make sure to have only one Month (Last Month) from the Dossier

pre_month_actual <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/07.02.2024/Actual Shipped.xlsx")

pre_month_actual %>% 
  janitor::clean_names() -> pre_month_actual

pre_month_actual[-1, ] -> pre_month_actual
colnames(pre_month_actual)[6] <- "qty"

pre_month_actual %>% 
  dplyr::mutate(qty = as.double(qty),
                qty = ifelse(is.na(qty), 0, qty)) %>% 
  dplyr::rename(loc = location,
                item_2 = calendar_month_year) %>% 
  dplyr::select(loc, item_2, qty) %>% 
  dplyr::mutate(item_2 = gsub("-", "", item_2),
                ref = paste0(loc, "_", item_2)) %>% 
  dplyr::left_join(IQR_FG_sample %>% select(ref, unit_cost) %>% mutate(ref = gsub("-", "_", ref))) %>% 
  dplyr::mutate(total_qty = unit_cost * qty,
                total_qty = ifelse(is.na(total_qty), 0, total_qty)) -> pre_month_actual

pre_month_actual$total_qty %>% sum()


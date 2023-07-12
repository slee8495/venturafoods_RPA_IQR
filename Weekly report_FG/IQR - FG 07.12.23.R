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

# (Path Revision Needed) Planner Address Book (If updated, correct this link) ----
# sdrive: S:/Supply Chain Projects/Linda Liang/reference files/Address Book - 04.26.22.xlsx

Planner_address <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/Address Book - 07.03.23.xlsx", 
                              sheet = "Sheet1", col_types = c("text", 
                                                              "text", "text", "text", "text"))

names(Planner_address) <- str_replace_all(names(Planner_address), c(" " = "_"))

colnames(Planner_address)[1] <- "Planner"

Planner_address %>% 
  dplyr::select(1:2) -> Planner_address

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

exception_report <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/BoM version 2/Weekly Run/7.12.2023/exception report.xlsx", 
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


exception_report %<>% 
  dplyr::mutate(ref = paste0(B_P, "_", ItemNo)) %>% 
  dplyr::relocate(ref) 



# (Path Revision Needed) Campus_ref pulling ----
# S drive: "S:/Supply Chain Projects/RStudio/BoM/Master formats/RM_on_Hand/Campus_ref.xlsx"
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


# (Path Revision Needed) Custord Receipt ----
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
reshape2::dcast(receipt, ref ~ next_28_days, value.var = "qty", sum) -> Receipt_Pivot  




# (Path Revision Needed) Custord wo ----
wo <- read.csv("Z:/IMPORT_JDE_OPENWO.csv",
               header = FALSE)


wo %>% 
  dplyr::rename(aa = V1) %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8"), sep = "~") %>% 
  dplyr::select(-"3") %>% 
  dplyr::rename(aa = "1") %>%  
  tidyr::separate(aa, c("global", "rp", "Item")) %>% 
  dplyr::select(-global, -rp, -"4", -"8") %>% 
  dplyr::rename(Location = "2",
                Qty = "5",
                wo_no = "6",
                date = "7") %>% 
  dplyr::mutate(Qty = as.double(Qty),
                date = as.Date(date)) %>% 
  dplyr::mutate(Location = sub("^0+", "", Location)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item)) %>% 
  dplyr::relocate(ref, Item, Location) %>% 
  dplyr::mutate(in_next_7_days = ifelse(date >= Sys.Date() & date < Sys.Date()+7, "Y", "N") )-> wo

# wo pivot
wo %>% 
  reshape2::dcast(ref ~ in_next_7_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(N = as.integer(N)) -> wo_pivot




# (Path Revision Needed) custord custord ----
# Open Customer Order File pulling ----  Change Directory ----
custord <- read.csv("Z:/IMPORT_CUSTORDS.csv",
                    header = FALSE)



custord %>% 
  dplyr::rename(aa = V1) %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8", "9"), sep = "~") %>% 
  dplyr::select(-"3", -"6", -"7", -"8") -> custord


substr(custord$`1`, nchar(custord$`1`)-2, nchar(custord$`1`)) -> label_custord

custord %>% 
  dplyr::bind_cols(label_custord) %>% 
  dplyr::rename(label = "...6",
                aa = "1") %>% 
  dplyr::mutate(label_na = as.integer(label),
                label_na = is.na(label_na)) %>% 
  dplyr::filter(label_na == TRUE) %>% 
  tidyr::separate(aa, c("global", "rp", "Item")) %>% 
  dplyr::select(-global, -rp, -"4", -label, -label_na) %>% 
  dplyr::rename(Location = "2",
                Qty = "5",
                date = "9") %>% 
  dplyr::mutate(Qty = as.double(Qty),
                date = as.Date(date)) %>% 
  dplyr::mutate(Location = sub("^0+", "", Location)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item),
                in_next_7_days = ifelse(date < Sys.Date() + 7, "Y", "N"),
                in_next_14_days = ifelse(date < Sys.Date() + 14, "Y", "N"),
                in_next_21_days = ifelse(date < Sys.Date() + 21, "Y", "N"),
                in_next_28_days = ifelse(date < Sys.Date() + 28, "Y", "N")) %>% 
  dplyr::relocate(ref, Item, Location) -> custord


# (Path Revision Needed) Loc 624 for custord ----
loc_624_custord <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/BoM version 2/Weekly Run/7.12.2023/Canada Open Orders (14).xlsx", 
                              col_names = FALSE)

loc_624_custord[-1:-2, ] -> loc_624_custord
colnames(loc_624_custord) <- loc_624_custord[1, ]
loc_624_custord[-1, ] -> loc_624_custord

colnames(loc_624_custord)[1] <- "Location"
colnames(loc_624_custord)[2] <- "Location_name"
colnames(loc_624_custord)[3] <- "Item"
colnames(loc_624_custord)[4] <- "description"
colnames(loc_624_custord)[5] <- "date"
colnames(loc_624_custord)[6] <- "Qty"


loc_624_custord %>% 
  dplyr::mutate(Item = gsub("-", "", Item)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item)) %>% 
  dplyr::relocate(ref, Item, Location, Qty, date) %>% 
  dplyr::select(-Location_name, -description) %>% 
  dplyr::mutate(date = as.integer(date),
                date = as.Date(date, origin = "1899-12-30"),
                Qty = as.double(Qty)) %>% 
  dplyr::mutate(in_next_7_days = ifelse(date < Sys.Date() + 7, "Y", "N"),
                in_next_14_days = ifelse(date < Sys.Date() + 14, "Y", "N"),
                in_next_21_days = ifelse(date < Sys.Date() + 21, "Y", "N"),
                in_next_28_days = ifelse(date < Sys.Date() + 28, "Y", "N")) %>% 
  dplyr::mutate(Qty = replace(Qty, is.na(Qty), 0)) -> loc_624_custord



rbind(custord, loc_624_custord) -> custord

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


# (Path Revision Needed) DSX Forecast pulling (Previous month file)---- Change Directory ----

DSX_Forecast_Backup_pre <- read_excel(
  "S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/2023/DSX Forecast Backup - 2023.06.01.xlsx")

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

# DSX_Forecast_Backup_pre$Forecast_Month_Year_Code_Segment_ID <- as.double(DSX_Forecast_Backup_pre$Forecast_Month_Year_Code_Segment_ID)
# DSX_Forecast_Backup_pre$Product_Manufacturing_Location_Code <- as.double(DSX_Forecast_Backup_pre$Product_Manufacturing_Location_Code)
# DSX_Forecast_Backup_pre$Location_No <- as.double(DSX_Forecast_Backup_pre$Location_No)
# DSX_Forecast_Backup_pre$Product_Manufacturing_Line_Area_No_Code <- as.double(DSX_Forecast_Backup_pre$Product_Manufacturing_Line_Area_No_Code)
# DSX_Forecast_Backup_pre$Safety_Stock_ID <- as.double(DSX_Forecast_Backup_pre$Safety_Stock_ID)
# DSX_Forecast_Backup_pre$Adjusted_Forecast_Pounds_lbs <- as.double(DSX_Forecast_Backup_pre$Adjusted_Forecast_Pounds_lbs)
# DSX_Forecast_Backup_pre$Adjusted_Forecast_Cases <- as.double(DSX_Forecast_Backup_pre$Adjusted_Forecast_Cases)
# DSX_Forecast_Backup_pre$Stat_Forecast_Pounds_lbs <- as.double(DSX_Forecast_Backup_pre$Stat_Forecast_Pounds_lbs)
# DSX_Forecast_Backup_pre$Stat_Forecast_Cases <- as.double(DSX_Forecast_Backup_pre$Stat_Forecast_Cases)
# DSX_Forecast_Backup_pre$Cust_Ref_Forecast_Pounds_lbs <- as.double(DSX_Forecast_Backup_pre$Cust_Ref_Forecast_Pounds_lbs)
# DSX_Forecast_Backup_pre$Cust_Ref_Forecast_Cases <- as.double(DSX_Forecast_Backup_pre$Cust_Ref_Forecast_Cases)


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
  "S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/DSX Forecast Backup - 2023.07.11.xlsx")

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

# DSX_Forecast_Backup$Forecast_Month_Year_Code_Segment_ID <- as.double(DSX_Forecast_Backup$Forecast_Month_Year_Code_Segment_ID)
# DSX_Forecast_Backup$Product_Manufacturing_Location_Code <- as.double(DSX_Forecast_Backup$Product_Manufacturing_Location_Code)
# DSX_Forecast_Backup$Location_No <- as.double(DSX_Forecast_Backup$Location_No)
# DSX_Forecast_Backup$Product_Manufacturing_Line_Area_No_Code <- as.double(DSX_Forecast_Backup$Product_Manufacturing_Line_Area_No_Code)
# DSX_Forecast_Backup$Safety_Stock_ID <- as.double(DSX_Forecast_Backup$Safety_Stock_ID)
# DSX_Forecast_Backup$Adjusted_Forecast_Pounds_lbs <- as.double(DSX_Forecast_Backup$Adjusted_Forecast_Pounds_lbs)
# DSX_Forecast_Backup$Adjusted_Forecast_Cases <- as.double(DSX_Forecast_Backup$Adjusted_Forecast_Cases)
# DSX_Forecast_Backup$Stat_Forecast_Pounds_lbs <- as.double(DSX_Forecast_Backup$Stat_Forecast_Pounds_lbs)
# DSX_Forecast_Backup$Stat_Forecast_Cases <- as.double(DSX_Forecast_Backup$Stat_Forecast_Cases)
# DSX_Forecast_Backup$Cust_Ref_Forecast_Pounds_lbs <- as.double(DSX_Forecast_Backup$Cust_Ref_Forecast_Pounds_lbs)
# DSX_Forecast_Backup$Cust_Ref_Forecast_Cases <- as.double(DSX_Forecast_Backup$Cust_Ref_Forecast_Cases)

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

# Read FG ----
Inventory_analysis_FG <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/BoM version 2/Weekly Run/7.12.2023/Inventory Report for all locations (30).xlsx")


Inventory_analysis_FG[-1,] -> Inventory_analysis_FG
colnames(Inventory_analysis_FG) <- Inventory_analysis_FG[1, ]
Inventory_analysis_FG[-1, ] -> Inventory_analysis_FG

colnames(Inventory_analysis_FG)[1] <- "Location"
colnames(Inventory_analysis_FG)[2] <- "Location_Nm"
colnames(Inventory_analysis_FG)[3] <- "campus"
colnames(Inventory_analysis_FG)[4] <- "SKU"
colnames(Inventory_analysis_FG)[5] <- "Description"
colnames(Inventory_analysis_FG)[6] <- "Inventory_Status"
colnames(Inventory_analysis_FG)[7] <- "Inventory_Hold_Status"
colnames(Inventory_analysis_FG)[8] <- "Inventory_Qty_Cases"

Inventory_analysis_FG %>% 
  dplyr::mutate(campus_ref = paste0(campus, "_", SKU), campus_ref = gsub("-", "", campus_ref)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", SKU), ref = gsub("-", "", ref)) %>% 
  dplyr::relocate(ref, campus_ref) -> Inventory_analysis_FG

readr::type_convert(Inventory_analysis_FG) -> Inventory_analysis_FG


# Inventory_analysis_pivot_ref

reshape2::dcast(Inventory_analysis_FG, ref ~ Inventory_Hold_Status, value.var = "Inventory_Qty_Cases", sum) -> pivot_ref_Inventory_analysis

names(pivot_ref_Inventory_analysis) <- str_replace_all(names(pivot_ref_Inventory_analysis), c(" " = "_"))

reshape2::dcast(Inventory_analysis_FG, campus_ref ~ Inventory_Hold_Status, value.var = "Inventory_Qty_Cases", sum) %>% 
  dplyr::rename(Usable = Useable, Loc_SKU = campus_ref, Hard_Hold = "Hard Hold", Soft_Hold = "Soft Hold") -> pivot_campus_ref_Inventory_analysis

names(pivot_campus_ref_Inventory_analysis) <- str_replace_all(names(pivot_campus_ref_Inventory_analysis), c(" " = "_"))


# (Path Revision Needed) Main Dataset Board ----

IQR_FG_sample <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Finished Goods Inventory Health Adjusted Forward (IQR) - 07.05.23.xlsx",
                            sheet = "FG without BKO BKM TST")

IQR_FG_sample[-1:-2,] -> IQR_FG_sample

colnames(IQR_FG_sample) <- IQR_FG_sample[1, ]
IQR_FG_sample[-1, ] -> IQR_FG_sample

names(IQR_FG_sample) <- str_replace_all(names(IQR_FG_sample), c(" " = "_"))
names(IQR_FG_sample) <- str_replace_all(names(IQR_FG_sample), c("-" = "_"))
names(IQR_FG_sample) <- str_replace_all(names(IQR_FG_sample), c("/" = "_"))
names(IQR_FG_sample) <- str_replace_all(names(IQR_FG_sample), c(">" = "greater"))
names(IQR_FG_sample) <- str_replace_all(names(IQR_FG_sample), c("<=" = "less_or_equal"))


colnames(IQR_FG_sample)[9] <- "On_Priority_list"
colnames(IQR_FG_sample)[34] <- "Useable"
colnames(IQR_FG_sample)[36] <- "Quality_hold_in_cost"
colnames(IQR_FG_sample)[38] <- "On_Hand_usable_and_soft_hold"
colnames(IQR_FG_sample)[40] <- "On_Hand_in_cost"
colnames(IQR_FG_sample)[41] <- "On_Hand_Adjusted_Forward_Max_in_cost"
colnames(IQR_FG_sample)[42] <- "On_Hand_Mfg_Adjusted_Forward_Max_in_cost"
colnames(IQR_FG_sample)[43] <- "Forward_Inv_Target_Current_Month_Fcst"
colnames(IQR_FG_sample)[45] <- "Forward_Inv_Target_Current_Month_Fcst_in_cost"
colnames(IQR_FG_sample)[48] <- "Adjusted_Forward_Inv_Target_in_cost"
colnames(IQR_FG_sample)[51] <- "Adjusted_Forward_Inv_Max_in_cost"
colnames(IQR_FG_sample)[54] <- "Mfg_Adjusted_Forward_Inv_Target_in_cost"
colnames(IQR_FG_sample)[57] <- "Mfg_Adjusted_Forward_Inv_Max_in_cost"
colnames(IQR_FG_sample)[58] <- "Forward_Inv_Target_lag_1_Current_Month_Fcst"
colnames(IQR_FG_sample)[60] <- "Forward_Inv_Target_lag_1_Current_Month_Fcst_in_cost"
colnames(IQR_FG_sample)[66] <- "CustOrd_in_next_28_days_in_cost"
colnames(IQR_FG_sample)[71] <- "Mfg_CustOrd_in_next_28_days_in_cost"
colnames(IQR_FG_sample)[77] <- "Forward_Target_Inv_DOS_fcst_only"
colnames(IQR_FG_sample)[78] <- "Adjusted_Forward_Target_Inv_DOS_includes_Orders"
colnames(IQR_FG_sample)[82] <- "Mfg_Forward_Target_Inv_DOS_fcst_only"
colnames(IQR_FG_sample)[83] <- "Mfg_Adjusted_Forward_Target_Inv_DOS_includes_Orders"
colnames(IQR_FG_sample)[87] <- "Lag_1_Current_Month_Fcst_in_cost"
colnames(IQR_FG_sample)[96] <- "has_adjusted_forward_looking_Max"
colnames(IQR_FG_sample)[106] <- "has_mfg_adjusted_forward_looking_Max"


IQR_FG_sample %>% 
  dplyr::mutate(Ref = gsub("-", "_", Ref),
                Campus_Ref = gsub("-", "_", Campus_Ref),
                Mfg_Ref = gsub("-", "_", Mfg_Ref)) %>% 
  dplyr::rename(ref = Ref,
                Loc_SKU = Campus_Ref,
                mfg_ref = Mfg_Ref) -> IQR_FG_sample


# (Path Revision Needed) read SD & CV file ----
sdcv <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Standard Deviation & CV/Standard Deviation, CV,  June 2023.xlsx")

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

# ## mfg_loc & Mfg_Ref
# merge(IQR_FG_sample, FG_ref_to_mfg_ref[, c("ref", "mfg_loc")], by = "ref", all.x = TRUE) %>%
#   dplyr::relocate(mfg_loc.y, .after = mfg_loc.x) %>%
#   dplyr::select(-mfg_loc.x) %>%
#   dplyr::rename(mfg_loc = mfg_loc.y) -> IQR_FG_sample
# 
# 
# ## Mfg Ref
# IQR_FG_sample %<>%
#   dplyr::mutate(mfg_ref = paste0(mfg_loc, "_", Item_2)) %>%
#   dplyr::select(-Mfg_Ref)


readr::type_convert(IQR_FG_sample) -> IQR_FG_sample

##################################### vlookups #########################################

# vlookup - MTO_MTS (with if, iferror)
merge(IQR_FG_sample, exception_report[, c("ref", "Order_Policy_Code")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Order_Policy_Code = as.integer(Order_Policy_Code)) -> IQR_FG_sample

IQR_FG_sample$Order_Policy_Code[is.na(IQR_FG_sample$Order_Policy_Code)] <- 1

IQR_FG_sample %<>% 
  dplyr::mutate(MTO_MTS = ifelse(Order_Policy_Code == 1, "MTO", "MTS")) %>% 
  dplyr::select(-Order_Policy_Code)

# vlookup - MPF
merge(IQR_FG_sample, exception_report[, c("ref", "MPF_or_Line")], by = "ref", all.x = TRUE) %>% 
  dplyr::relocate(MPF_or_Line, .after = MPF) %>% 
  dplyr::select(-MPF) %>% 
  dplyr::rename(MPF = MPF_or_Line) %>% 
  dplyr::mutate(MPF = replace(MPF, is.na(MPF), "DNRR")) -> IQR_FG_sample


# vlookup - Planner
merge(IQR_FG_sample, exception_report[, c("ref", "Planner")], by = "ref", all.x = TRUE) %>% 
  dplyr::relocate(Planner.y, .after = Planner.x) %>% 
  dplyr::select(-Planner.x) %>% 
  dplyr::rename(Planner = Planner.y) %>% 
  dplyr::mutate(Planner = replace(Planner, is.na(Planner), "DNRR")) -> IQR_FG_sample

# vlookup - Planner Name 
merge(IQR_FG_sample, Planner_address[, c("Planner", "Alpha_Name")], by = "Planner", all.x = TRUE) %>% 
  dplyr::mutate(Planner_Name = ifelse(Planner == 0, "NA",
                                      ifelse(Planner == "DNRR", "DNRR",
                                             Alpha_Name))) %>% 
  dplyr::select(-Alpha_Name) -> IQR_FG_sample


# vlookup - JDE MOQ
merge(IQR_FG_sample, exception_report[, c("ref", "Reorder_MIN")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(reorder_min_na = !is.na(Reorder_MIN)) %>% 
  dplyr::mutate(JDE_MOQ = ifelse(reorder_min_na == TRUE, Reorder_MIN, 0)) %>% 
  dplyr::select(-reorder_min_na, -Reorder_MIN) -> IQR_FG_sample

# vlookup - Current SS
merge(IQR_FG_sample, exception_report[, c("ref", "Safety_Stock")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(safety_stock_na = !is.na(Safety_Stock)) %>% 
  dplyr::mutate(Current_SS = ifelse(safety_stock_na == TRUE, Safety_Stock, 0)) %>% 
  dplyr::select(-safety_stock_na, -Safety_Stock) -> IQR_FG_sample


# vlookup - Useable
merge(IQR_FG_sample, pivot_ref_Inventory_analysis[, c("ref", "Useable")], by = "ref", all.x = TRUE) %>%
  dplyr::mutate(useable_na = !is.na(Useable.y)) %>% 
  dplyr::mutate(Useable = ifelse(useable_na == TRUE, Useable.y, 0)) %>%
  dplyr::relocate(Useable, .after = Useable.x) %>% 
  dplyr::select(-useable_na, -Useable.y, -Useable.x) -> IQR_FG_sample


# vlookup - Quality hold
merge(IQR_FG_sample, pivot_ref_Inventory_analysis[, c("ref", "Hard_Hold")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(hard_hold_na = !is.na(Hard_Hold)) %>% 
  dplyr::mutate(Quality_hold = ifelse(hard_hold_na == TRUE, Hard_Hold, 0)) %>% 
  dplyr::select(-hard_hold_na, -Hard_Hold) -> IQR_FG_sample

# vlookup - Soft Hold
merge(IQR_FG_sample, pivot_ref_Inventory_analysis[, c("ref", "Soft_Hold")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(soft_hold_na = !is.na(Soft_Hold.y)) %>% 
  dplyr::mutate(Soft_Hold = ifelse(soft_hold_na == TRUE, Soft_Hold.y, 0)) %>% 
  dplyr::select(-soft_hold_na, -Soft_Hold.y, -Soft_Hold.x) %>% 
  dplyr::relocate(Soft_Hold, .after = Quality_hold_in_cost)-> IQR_FG_sample

# vlookup - OPV
merge(IQR_FG_sample, exception_report[, c("ref", "Order_Policy_Value")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(opv_na = !is.na(Order_Policy_Value)) %>% 
  dplyr::mutate(OPV = ifelse(opv_na == TRUE, Order_Policy_Value, 0)) %>% 
  dplyr::select(-opv_na, -Order_Policy_Value) -> IQR_FG_sample


# vlookup - CustOrd in next 7 days
merge(IQR_FG_sample, custord_pivot_1[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(CustOrd_in_next_7_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - CustOrd in next 14 days
merge(IQR_FG_sample, custord_pivot_2[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(CustOrd_in_next_14_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - CustOrd in next 21 days
merge(IQR_FG_sample, custord_pivot_3[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(CustOrd_in_next_21_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - CustOrd in next 28 days
merge(IQR_FG_sample, custord_pivot_4[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(CustOrd_in_next_28_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - Firm WO in next 28 days
merge(IQR_FG_sample, wo_pivot[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(Firm_WO_in_next_28_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - Receipt in the next 28 days
merge(IQR_FG_sample, Receipt_Pivot[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(Receipt_in_the_next_28_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample

# vlookup - Lag_1_Current_Month_Fcst  
merge(IQR_FG_sample, DSX_pivot_1_pre[, c("ref", "Mon_b_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(mon_b_na = !is.na(Mon_b_fcst)) %>% 
  dplyr::mutate(Lag_1_Current_Month_Fcst = ifelse(mon_b_na == TRUE, Mon_b_fcst, 0)) %>% 
  dplyr::select(-Mon_b_fcst, -mon_b_na) %>% 
  dplyr::mutate(Lag_1_Current_Month_Fcst = round(Lag_1_Current_Month_Fcst , 0)) -> IQR_FG_sample


# vlookup - Current Month Fcst
merge(IQR_FG_sample, DSX_pivot_1[, c("ref", "Mon_a_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(mon_a_na = !is.na(Mon_a_fcst)) %>% 
  dplyr::mutate(Current_Month_Fcst = ifelse(mon_a_na == TRUE, Mon_a_fcst, 0)) %>% 
  dplyr::select(-Mon_a_fcst, - mon_a_na) %>% 
  dplyr::mutate(Current_Month_Fcst = round(Current_Month_Fcst , 0)) -> IQR_FG_sample

# vlookup - Next Month Fcst
merge(IQR_FG_sample, DSX_pivot_1[, c("ref", "Mon_b_fcst")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(mon_b_na = !is.na(Mon_b_fcst)) %>% 
  dplyr::mutate(Next_Month_Fcst = ifelse(mon_b_na == TRUE, Mon_b_fcst, 0)) %>% 
  dplyr::select(-Mon_b_fcst, - mon_b_na) %>% 
  dplyr::mutate(Next_Month_Fcst  = round(Next_Month_Fcst, 0)) -> IQR_FG_sample


# vlookup - Total Last 6 mos Sales
merge(IQR_FG_sample, sdcv[, c("ref", "last_6_month_sales")], by = "ref", all.x = TRUE) %>% 
  dplyr::relocate(last_6_month_sales, .after = Total_Last_6_mos_Sales) %>% 
  dplyr::select(-Total_Last_6_mos_Sales) %>% 
  dplyr::rename(Total_Last_6_mos_Sales = last_6_month_sales) -> IQR_FG_sample


# vlookup - Total Last 12 mos Sales 
merge(IQR_FG_sample, sdcv[, c("ref", "last_12_month_sales")], by = "ref", all.x = TRUE) %>% 
  dplyr::relocate(last_12_month_sales, .after = Total_Last_12_mos_Sales) %>% 
  dplyr::select(-Total_Last_12_mos_Sales) %>% 
  dplyr::rename(Total_Last_12_mos_Sales = last_12_month_sales) -> IQR_FG_sample


# vlookup - Total Forecast Next 12 Months
merge(IQR_FG_sample, DSX_pivot_1[, c("ref", "total_12_month")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(total_12_month = replace(total_12_month, is.na(total_12_month), 0)) %>% 
  dplyr::relocate(total_12_month, .after = Total_Forecast_Next_12_Months) %>% 
  dplyr::select(-Total_Forecast_Next_12_Months) %>% 
  dplyr::rename(Total_Forecast_Next_12_Months = total_12_month) -> IQR_FG_sample




# vlookup - Mfg CustOrd in next 7 days
merge(IQR_FG_sample, custord_mfg_pivot_1[, c("mfg_ref", "Y")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(Mfg_CustOrd_in_next_7_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - Mfg CustOrd in next 14 days
merge(IQR_FG_sample, custord_mfg_pivot_2[, c("mfg_ref", "Y")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(Mfg_CustOrd_in_next_14_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - Mfg CustOrd in next 21 days
merge(IQR_FG_sample, custord_mfg_pivot_3[, c("mfg_ref", "Y")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(Mfg_CustOrd_in_next_21_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - Mfg CustOrd in next 28 days
merge(IQR_FG_sample, custord_mfg_pivot_4[, c("mfg_ref", "Y")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(y_na = !is.na(Y)) %>% 
  dplyr::mutate(Mfg_CustOrd_in_next_28_days = ifelse(y_na == TRUE, Y, 0)) %>% 
  dplyr::select(-Y, -y_na) -> IQR_FG_sample


# vlookup - Mfg Current Month Fcst
merge(IQR_FG_sample, DSX_mfg_pivot_1[, c("mfg_ref", "Mon_a_fcst")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(mon_a_na = !is.na(Mon_a_fcst)) %>% 
  dplyr::mutate(Mfg_Current_Month_Fcst = ifelse(mon_a_na == TRUE, Mon_a_fcst, 0)) %>% 
  dplyr::select(-Mon_a_fcst, - mon_a_na) %>% 
  dplyr::mutate(Mfg_Current_Month_Fcst = round(Mfg_Current_Month_Fcst , 0)) -> IQR_FG_sample

# vlookup - Mfg Next Month Fcst
merge(IQR_FG_sample, DSX_mfg_pivot_1[, c("mfg_ref", "Mon_b_fcst")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(mon_b_na = !is.na(Mon_b_fcst)) %>% 
  dplyr::mutate(Mfg_Next_Month_Fcst = ifelse(mon_b_na == TRUE, Mon_b_fcst, 0)) %>% 
  dplyr::select(-Mon_b_fcst, - mon_b_na) %>% 
  dplyr::mutate(Mfg_Next_Month_Fcst  = round(Mfg_Next_Month_Fcst, 0)) -> IQR_FG_sample



# vlookup - Total Forecast Next 12 Months
merge(IQR_FG_sample, DSX_mfg_pivot_1[, c("mfg_ref", "total_12_month")], by = "mfg_ref", all.x = TRUE) %>% 
  dplyr::mutate(total_12_month = replace(total_12_month, is.na(total_12_month), 0)) %>% 
  dplyr::relocate(total_12_month, .after = Total_mfg_Forecast_Next_12_Months) %>% 
  dplyr::select(-Total_mfg_Forecast_Next_12_Months) %>% 
  dplyr::rename(Total_mfg_Forecast_Next_12_Months = total_12_month) -> IQR_FG_sample



##################################### calculates #########################################
readr::type_convert(IQR_FG_sample) -> IQR_FG_sample

# calculate - Max Cycle Stock
IQR_FG_sample %<>% 
  dplyr::mutate(Max_Cycle_Stock = ifelse(OPV == 0, pmax(CustOrd_in_next_28_days, JDE_MOQ), pmax(JDE_MOQ, 
                                                                                                ifelse(OPV >= 20, CustOrd_in_next_28_days,
                                                                                                       ifelse(OPV >= 12 & OPV < 20, CustOrd_in_next_21_days,
                                                                                                              ifelse(OPV >= 8 & OPV < 12, CustOrd_in_next_14_days, CustOrd_in_next_7_days))) +
                                                                                                  Current_Month_Fcst / 20.83 * Hold_days,
                                                                                                Current_Month_Fcst / 20.83 * (OPV + Hold_days),
                                                                                                Total_Last_12_mos_Sales / 250 * (OPV + Hold_days) ))) %>% 
  dplyr::mutate(Max_Cycle_Stock = round(Max_Cycle_Stock, 0)) 




# calculate - Max_Cycle_Stock_lag_1
IQR_FG_sample %<>% 
  dplyr::mutate(Max_Cycle_Stock_lag_1 = ifelse(OPV == 0, pmax(CustOrd_in_next_28_days, JDE_MOQ), pmax(JDE_MOQ, 
                                                                                                      ifelse(OPV >= 20, CustOrd_in_next_28_days,
                                                                                                             ifelse(OPV >= 12 & OPV < 20, CustOrd_in_next_21_days,
                                                                                                                    ifelse(OPV >= 8 & OPV < 12, CustOrd_in_next_14_days, CustOrd_in_next_7_days))) +
                                                                                                        Lag_1_Current_Month_Fcst / 20.83 * Hold_days,
                                                                                                      Lag_1_Current_Month_Fcst / 20.83 * (OPV + Hold_days) ))) %>% 
  dplyr::mutate(Max_Cycle_Stock_lag_1 = round(Max_Cycle_Stock_lag_1, 0)) 




# calculate - Max Cycle Stock Adjusted Forward 
IQR_FG_sample %<>% 
  dplyr::mutate(Max_Cycle_Stock_Adjusted_Forward = ifelse(OPV == 0, pmax(CustOrd_in_next_28_days, JDE_MOQ), pmax(JDE_MOQ, 
                                                                                                                 ifelse(OPV >= 20, CustOrd_in_next_28_days,
                                                                                                                        ifelse(OPV >= 12 & OPV < 20, CustOrd_in_next_21_days,
                                                                                                                               ifelse(OPV >= 8 & OPV < 12, CustOrd_in_next_14_days, CustOrd_in_next_7_days))) +
                                                                                                                   pmax(Current_Month_Fcst, Next_Month_Fcst) / 20.83 * Hold_days,
                                                                                                                 pmax(Current_Month_Fcst, Next_Month_Fcst) / 20.83 * (OPV + Hold_days) ))) %>% 
  dplyr::mutate(Max_Cycle_Stock_Adjusted_Forward = round(Max_Cycle_Stock_Adjusted_Forward, 0)) 




# calculate - Quality hold in $$
IQR_FG_sample %<>% 
  dplyr::mutate(Quality_hold_in_cost = Quality_hold * Unit_Cost) 


# calculate - On Hand (usable + soft hold)
IQR_FG_sample %<>% 
  dplyr::mutate(On_Hand_usable_and_soft_hold = Useable + Soft_Hold)


# calculate - On Hand in pounds
IQR_FG_sample %>% 
  dplyr::mutate(On_Hand_usable_and_soft_hold = as.numeric(On_Hand_usable_and_soft_hold),
                Net_Wt_Lbs = as.numeric(Net_Wt_Lbs)) %>% 
  dplyr::mutate(On_Hand_in_pounds = On_Hand_usable_and_soft_hold * Net_Wt_Lbs) -> IQR_FG_sample


# calculate - On Hand in $$
IQR_FG_sample %<>% 
  dplyr::mutate(On_Hand_in_cost = On_Hand_usable_and_soft_hold * Unit_Cost) 


# calculate - Adjusted Forward Inv Max
IQR_FG_sample %<>% 
  dplyr::mutate(Adjusted_Forward_Inv_Max = Max_Cycle_Stock_Adjusted_Forward + Current_SS)


# calculate - On Hand - Adjusted Forward Max in $$
IQR_FG_sample %<>% 
  dplyr::mutate(On_Hand_Adjusted_Forward_Max_in_cost = ifelse( (On_Hand_usable_and_soft_hold - Adjusted_Forward_Inv_Max) * Unit_Cost  < 0,
                                                               0, (On_Hand_usable_and_soft_hold - Adjusted_Forward_Inv_Max) * Unit_Cost   )) 



# (Business Days) calculate - Forward Inv Target Current Month Fcst (Business days input - ## Current Month ##) -----------------------------------
IQR_FG_sample %<>% 
  dplyr::mutate(Forward_Inv_Target_Current_Month_Fcst = (pmax((Current_Month_Fcst / 20) * OPV, JDE_MOQ)) / 2 + Current_SS,
                Forward_Inv_Target_Current_Month_Fcst = round(Forward_Inv_Target_Current_Month_Fcst, 0))


# calculate - Forward_Inv_Target_Current_Month_Fcst_in_lbs
IQR_FG_sample %<>% 
  dplyr::mutate(Forward_Inv_Target_Current_Month_Fcst_in_lbs. = Forward_Inv_Target_Current_Month_Fcst * Net_Wt_Lbs,
                Forward_Inv_Target_Current_Month_Fcst_in_lbs. = round(Forward_Inv_Target_Current_Month_Fcst_in_lbs., 0))


# calculate - Forward Inv Target Current Month Fcst in $$
IQR_FG_sample %<>% 
  dplyr::mutate(Forward_Inv_Target_Current_Month_Fcst_in_cost = Forward_Inv_Target_Current_Month_Fcst * Unit_Cost,
                Forward_Inv_Target_Current_Month_Fcst_in_cost = round(Forward_Inv_Target_Current_Month_Fcst_in_cost, 2))


# calculate - Adjusted Forward Inv Target
IQR_FG_sample %<>% 
  dplyr::mutate(Adjusted_Forward_Inv_Target = Max_Cycle_Stock_Adjusted_Forward / 2 + Current_SS,
                Adjusted_Forward_Inv_Target = round(Adjusted_Forward_Inv_Target, 0))


# calculate - Adjusted Forward Inv Target in lbs.
IQR_FG_sample %<>% 
  dplyr::mutate(Adjusted_Forward_Inv_Target_in_lbs. = Adjusted_Forward_Inv_Target * Net_Wt_Lbs,
                Adjusted_Forward_Inv_Target_in_lbs. = round(Adjusted_Forward_Inv_Target_in_lbs., 0))


# calculate - Adjusted Forward Inv Target in $$
IQR_FG_sample %<>% 
  dplyr::mutate(Adjusted_Forward_Inv_Target_in_cost = Adjusted_Forward_Inv_Target * Unit_Cost)


# calculate - Adjusted Forward Inv Max in lbs.
IQR_FG_sample %<>% 
  dplyr::mutate(Adjusted_Forward_Inv_Max_in_lbs. = Adjusted_Forward_Inv_Max * Net_Wt_Lbs,
                Adjusted_Forward_Inv_Max_in_lbs. = round(Adjusted_Forward_Inv_Max_in_lbs., 0))

# calculate - Adjusted Forward Inv Max in $$
IQR_FG_sample %<>% 
  dplyr::mutate(Adjusted_Forward_Inv_Max_in_cost = Adjusted_Forward_Inv_Max * Unit_Cost)


#calculate - Forward_Inv_Target_lag_1_Current_Month_Fcst
IQR_FG_sample %<>% 
  dplyr::mutate(Forward_Inv_Target_lag_1_Current_Month_Fcst = Max_Cycle_Stock_lag_1 / 2 + Current_SS,
                Forward_Inv_Target_lag_1_Current_Month_Fcst = round(Forward_Inv_Target_lag_1_Current_Month_Fcst , 0))


# calculate - Forward Inv Target lag 1 Current Month Fcst in lbs.
IQR_FG_sample %<>% 
  dplyr::mutate(Forward_Inv_Target_lag_1_Current_Month_Fcst_in_lbs. = Forward_Inv_Target_lag_1_Current_Month_Fcst * Net_Wt_Lbs,
                Forward_Inv_Target_lag_1_Current_Month_Fcst_in_lbs. = round(Forward_Inv_Target_lag_1_Current_Month_Fcst_in_lbs., 0))


# calculate - Forward Inv Target lag 1 Current Month Fcst in $$
IQR_FG_sample %<>% 
  dplyr::mutate(Forward_Inv_Target_lag_1_Current_Month_Fcst_in_cost = Forward_Inv_Target_lag_1_Current_Month_Fcst * Unit_Cost,
                Forward_Inv_Target_lag_1_Current_Month_Fcst_in_cost = round(Forward_Inv_Target_lag_1_Current_Month_Fcst_in_cost, 0))


# calculate - CustOrd in next 28 days in $$
IQR_FG_sample %<>% 
  dplyr::mutate(CustOrd_in_next_28_days_in_cost = CustOrd_in_next_28_days * Unit_Cost)


# calculate - DOS
IQR_FG_sample %<>% 
  dplyr::mutate(DOS = On_Hand_usable_and_soft_hold / pmax((ifelse(OPV == 0 | OPV >= 20, 
                                                                  CustOrd_in_next_28_days, 
                                                                  ifelse(OPV < 20  & OPV >= 12, CustOrd_in_next_21_days,
                                                                         ifelse(OPV < 12  & OPV >= 8, CustOrd_in_next_14_days, CustOrd_in_next_7_days))) / OPV), 
                                                          (pmax(Current_Month_Fcst, Next_Month_Fcst) / 20.83) ),
                dos_na = !is.na(DOS),
                DOS = ifelse(dos_na == TRUE, DOS, 0),
                DOS = round(DOS, 1),
                DOS = replace(DOS, is.infinite(DOS), 0)) %>% 
  dplyr::select(-dos_na)





# calculate - DOS after CustOrd
IQR_FG_sample %<>% 
  dplyr::mutate(DOS_after_CustOrd = (On_Hand_usable_and_soft_hold - ifelse(OPV == 0 | OPV >= 20, CustOrd_in_next_28_days, 
                                                                           ifelse(OPV < 20 & OPV >= 12, CustOrd_in_next_21_days, 
                                                                                  ifelse(OPV < 12 & OPV >= 8, CustOrd_in_next_14_days, CustOrd_in_next_7_days))))/
                  pmax((ifelse(OPV == 0 | OPV >= 20, CustOrd_in_next_28_days, 
                               ifelse(OPV < 20 & OPV >= 12, CustOrd_in_next_21_days,
                                      ifelse(OPV < 12 & OPV >= 8, CustOrd_in_next_14_days, CustOrd_in_next_7_days)))/OPV),
                       (pmax(Current_Month_Fcst, Next_Month_Fcst)/20.83))) %>% 
  dplyr::mutate(dos_na = !is.na(DOS_after_CustOrd),
                DOS_after_CustOrd = ifelse(dos_na == TRUE, DOS_after_CustOrd, 0),
                DOS_after_CustOrd = round(DOS_after_CustOrd, 1),
                DOS_after_CustOrd = replace(DOS_after_CustOrd, is.infinite(DOS_after_CustOrd), 0)) %>% 
  dplyr::select(-dos_na) %>% 
  dplyr::mutate(DOS_after_CustOrd = round(DOS_after_CustOrd, 1))


# calculate - Adjusted Forward Max Inv DOS
IQR_FG_sample %<>% 
  dplyr::mutate(Adjusted_Forward_Max_Inv_DOS = (Adjusted_Forward_Inv_Max / pmax((ifelse(OPV == 0 | OPV >= 20, CustOrd_in_next_28_days,
                                                                                        ifelse(OPV < 20 & OPV >= 12, CustOrd_in_next_21_days,
                                                                                               ifelse(OPV < 12 & OPV >= 8, CustOrd_in_next_14_days, CustOrd_in_next_7_days))) / OPV),
                                                                                (pmax(Current_Month_Fcst, Next_Month_Fcst) / 20.83)))) %>% 
  dplyr::mutate(dos_na = !is.na(Adjusted_Forward_Max_Inv_DOS),
                Adjusted_Forward_Max_Inv_DOS = ifelse(dos_na == TRUE, Adjusted_Forward_Max_Inv_DOS, 0),
                Adjusted_Forward_Max_Inv_DOS = round(Adjusted_Forward_Max_Inv_DOS, 1),
                Adjusted_Forward_Max_Inv_DOS = replace(Adjusted_Forward_Max_Inv_DOS, is.infinite(Adjusted_Forward_Max_Inv_DOS), 0)) %>% 
  dplyr::select(-dos_na) %>% 
  dplyr::mutate(Adjusted_Forward_Max_Inv_DOS = round(Adjusted_Forward_Max_Inv_DOS, 1))





# calculate - Forward_Target_Inv_DOS_fcst_only
IQR_FG_sample %<>% 
  dplyr::mutate(aa = Current_SS / (pmax(Current_Month_Fcst, Next_Month_Fcst) / 20.83),
                aa = replace(aa, is.na(aa), 0),
                aa = replace(aa, is.nan(aa), 0),
                aa = replace(aa, is.infinite(aa), 0)) %>% 
  dplyr::mutate(Forward_Target_Inv_DOS_fcst_only = ifelse(Total_Forecast_Next_12_Months == 0, 0,
                                                          OPV + aa)) %>% 
  dplyr::mutate(Forward_Target_Inv_DOS_fcst_only = round(Forward_Target_Inv_DOS_fcst_only, 1)) %>% 
  dplyr::select(-aa)



# calculate - Adjusted_Forward_Target_Inv_DOS_includes_Orders
IQR_FG_sample %<>% 
  dplyr::mutate(aa = Current_SS / pmax((ifelse(OPV == 0 | OPV >= 20, CustOrd_in_next_28_days,
                                               ifelse(OPV < 20 & OPV >= 12, CustOrd_in_next_21_days,
                                                      ifelse(OPV < 12 & OPV >= 8, CustOrd_in_next_14_days, CustOrd_in_next_7_days))) / OPV),
                                       (pmax(Current_Month_Fcst, Next_Month_Fcst) / 20.83)),
                aa = replace(aa, is.na(aa), 0),
                aa = replace(aa, is.nan(aa), 0),
                aa = replace(aa, is.infinite(aa), 0)) %>% 
  dplyr::mutate(Adjusted_Forward_Target_Inv_DOS_includes_Orders = ifelse(Total_Forecast_Next_12_Months == 0, 0, 
                                                                         OPV + aa)) %>% 
  dplyr::mutate(Adjusted_Forward_Target_Inv_DOS_includes_Orders = round(Adjusted_Forward_Target_Inv_DOS_includes_Orders, 1)) %>% 
  dplyr::select(-aa)



# calculate - on hand Inv after CustOrd > AF max
IQR_FG_sample %<>% 
  dplyr::mutate(on_hand_Inv_after_CustOrd_greater_AF_max = ifelse(On_Hand_usable_and_soft_hold - (ifelse(OPV == 0 | OPV >= 20, CustOrd_in_next_28_days, 
                                                                                                         ifelse(OPV < 20 & OPV >= 12, CustOrd_in_next_21_days,
                                                                                                                ifelse(OPV < 12 & OPV >= 8, CustOrd_in_next_14_days, CustOrd_in_next_7_days)))) 
                                                                  > Adjusted_Forward_Inv_Max, 1,0))



# calculate - Inv Health
IQR_FG_sample %<>% 
  dplyr::mutate(Inv_Health = ifelse(On_Hand_usable_and_soft_hold < Current_SS, "BELOW SS",
                                    ifelse(DOS_after_CustOrd > Shippable_Shelf_Life, "AT RISK" ,
                                           ifelse(Total_Forecast_Next_12_Months <= 0 & CustOrd_in_next_28_days <= 0,
                                                  ifelse(On_Hand_usable_and_soft_hold > 0, "DEAD",
                                                         ifelse(on_hand_Inv_after_CustOrd_greater_AF_max == 0, "HEALTHY", "EXCESS")),
                                                  ifelse(on_hand_Inv_after_CustOrd_greater_AF_max == 1, "EXCESS", "HEALTHY"))))) 


# calculate - Lag 1 Current Month Fcst in cost
IQR_FG_sample %<>% 
  dplyr::mutate(Lag_1_Current_Month_Fcst_in_cost = Lag_1_Current_Month_Fcst * Unit_Cost)


# calculate - has adjusted forward looking Max?
IQR_FG_sample %<>% 
  dplyr::mutate(has_adjusted_forward_looking_Max = ifelse(Adjusted_Forward_Inv_Max > 0, 1, 0))


# calculate - on hand Inv > AF max
IQR_FG_sample %<>% 
  dplyr::mutate(on_hand_Inv_greater_AF_max = ifelse(On_Hand_usable_and_soft_hold > Adjusted_Forward_Inv_Max, 1, 0))


# calculate - on hand Inv <= AF max
IQR_FG_sample %<>% 
  dplyr::mutate(on_hand_Inv_less_or_equal_AF_max = ifelse(On_Hand_usable_and_soft_hold <= Adjusted_Forward_Inv_Max, 1, 0))


# calculate - on hand Inv > Adjusted Forward looking target
IQR_FG_sample %<>% 
  dplyr::mutate(on_hand_Inv_greater_Adjusted_Forward_looking_target = ifelse(On_Hand_usable_and_soft_hold > Adjusted_Forward_Inv_Target, 1, 0))


# calculate - on hand Inv <= AF target
IQR_FG_sample %<>% 
  dplyr::mutate(on_hand_Inv_less_or_equal_AF_target = ifelse(On_Hand_usable_and_soft_hold <= Adjusted_Forward_Inv_Target, 1, 0))


# calculate - on hand Inv after CustOrd <= AF max
IQR_FG_sample %<>% 
  dplyr::mutate(on_hand_Inv_after_CustOrd_less_or_equal_AF_max = ifelse(On_Hand_usable_and_soft_hold - (ifelse(OPV == 0 | OPV >= 20, CustOrd_in_next_28_days,
                                                                                                               ifelse(OPV < 20 & OPV >= 12, CustOrd_in_next_21_days,
                                                                                                                      ifelse(OPV < 12 & OPV >= 8, CustOrd_in_next_14_days, CustOrd_in_next_7_days))))
                                                                        <= Adjusted_Forward_Inv_Max , 1, 0))


# calculate - on hand Inv after CustOrd > AF target
IQR_FG_sample %<>% 
  dplyr::mutate(on_hand_Inv_after_CustOrd_greater_AF_target = ifelse(On_Hand_usable_and_soft_hold - (ifelse(OPV == 0 | OPV >= 20, CustOrd_in_next_28_days,
                                                                                                            ifelse(OPV < 20 & OPV >= 12, CustOrd_in_next_21_days,
                                                                                                                   ifelse(OPV < 12 & OPV >= 8, CustOrd_in_next_14_days, CustOrd_in_next_7_days))))
                                                                     > Adjusted_Forward_Inv_Target, 1, 0))



# calculate - on hand Inv after CustOrd <= AF target
IQR_FG_sample %<>% 
  dplyr::mutate(on_hand_Inv_after_CustOrd_less_or_equal_AF_target = ifelse(On_Hand_usable_and_soft_hold - (ifelse(OPV == 0 | OPV >= 20, CustOrd_in_next_28_days,
                                                                                                                  ifelse(OPV < 20 & OPV >= 12, CustOrd_in_next_21_days,
                                                                                                                         ifelse(OPV < 12 & OPV >= 8, CustOrd_in_next_14_days, CustOrd_in_next_7_days))))
                                                                           <= Adjusted_Forward_Inv_Target, 1, 0))


# calculate - on hand inv after 28 days CustOrd > 0
IQR_FG_sample %<>% 
  dplyr::mutate(on_hand_inv_after_28_days_CustOrd_greater_0 = ifelse(On_Hand_usable_and_soft_hold - CustOrd_in_next_28_days > 0, 1,0))



# calculate - Max Cycle Stock Mfg Adjusted Forward
IQR_FG_sample %<>%
  dplyr::mutate(Max_Cycle_Stock_Mfg_Adjusted_Forward = ifelse(OPV == 0,pmax(Mfg_CustOrd_in_next_28_days, JDE_MOQ),
                                                              pmax(JDE_MOQ, ifelse(OPV >= 20, Mfg_CustOrd_in_next_28_days,
                                                                                   ifelse(OPV >= 12 & OPV < 20, Mfg_CustOrd_in_next_21_days,
                                                                                          ifelse(OPV >= 8 & OPV < 12, Mfg_CustOrd_in_next_14_days, Mfg_CustOrd_in_next_7_days)))
                                                                   + pmax(Mfg_Current_Month_Fcst, Mfg_Next_Month_Fcst) / 20.83 * Hold_days, pmax(Mfg_Current_Month_Fcst, Mfg_Next_Month_Fcst) / 20.83 *
                                                                     (OPV+Hold_days)))) %>%
  dplyr::mutate(Max_Cycle_Stock_Mfg_Adjusted_Forward = round(Max_Cycle_Stock_Mfg_Adjusted_Forward, 0))




# calculate - Mfg Adjusted Forward Inv Max
IQR_FG_sample %<>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Inv_Max = Max_Cycle_Stock_Mfg_Adjusted_Forward + Current_SS) %>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Inv_Max = round(Mfg_Adjusted_Forward_Inv_Max, 0))




# calculate - On Hand - Mfg Adjusted Forward Max in $$
IQR_FG_sample %<>% 
  dplyr::mutate(On_Hand_Mfg_Adjusted_Forward_Max_in_cost = ifelse((On_Hand_usable_and_soft_hold - Mfg_Adjusted_Forward_Inv_Max) * 
                                                                    Unit_Cost < 0, 0, (On_Hand_usable_and_soft_hold - 
                                                                                         Mfg_Adjusted_Forward_Inv_Max) * Unit_Cost) )



# calculate - Mfg Adjusted Forward Inv Target
IQR_FG_sample %<>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Inv_Target = Max_Cycle_Stock_Mfg_Adjusted_Forward / 2 + Current_SS) %>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Inv_Target = round(Mfg_Adjusted_Forward_Inv_Target, 0))



# calculate - Mfg Adjusted Forward Inv Target in lbs.
IQR_FG_sample %<>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Inv_Target_in_lbs. = Mfg_Adjusted_Forward_Inv_Target * Net_Wt_Lbs) %>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Inv_Target_in_lbs. = round(Mfg_Adjusted_Forward_Inv_Target_in_lbs., 2))


# calculate - Mfg Adjusted Forward Inv Target in $$
IQR_FG_sample %<>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Inv_Target_in_cost = Mfg_Adjusted_Forward_Inv_Target * Unit_Cost) %>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Inv_Target_in_cost = round(Mfg_Adjusted_Forward_Inv_Target_in_cost, 2))


# calculate - Mfg Adjusted Forward Inv Max in lbs.
IQR_FG_sample %<>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Inv_Max_in_lbs. = Mfg_Adjusted_Forward_Inv_Max * Net_Wt_Lbs) %>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Inv_Max_in_lbs. = round(Mfg_Adjusted_Forward_Inv_Max_in_lbs., 2))


# calculate - Mfg Adjusted Forward Inv Max in $$
IQR_FG_sample %<>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Inv_Max_in_cost = Mfg_Adjusted_Forward_Inv_Max * Unit_Cost) %>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Inv_Max_in_cost = round(Mfg_Adjusted_Forward_Inv_Max_in_cost, 2))

# calculate - Mfg CustOrd in next 28 days in $$
IQR_FG_sample %<>% 
  dplyr::mutate(Mfg_CustOrd_in_next_28_days_in_cost = Mfg_CustOrd_in_next_28_days * Unit_Cost) %>% 
  dplyr::mutate(Mfg_CustOrd_in_next_28_days_in_cost = round(Mfg_CustOrd_in_next_28_days_in_cost, 2))


# calculate - Mfg Dos 

# test
IQR_FG_sample %<>% 
  dplyr::mutate(Mfg_DOS = On_Hand_usable_and_soft_hold / pmax((ifelse(OPV == 0 | OPV >= 20, 
                                                                      Mfg_CustOrd_in_next_28_days, 
                                                                      ifelse(OPV < 20  & OPV >= 12, Mfg_CustOrd_in_next_21_days,
                                                                             ifelse(OPV < 12  & OPV >= 8, Mfg_CustOrd_in_next_14_days, 
                                                                                    Mfg_CustOrd_in_next_7_days))) / OPV), 
                                                              (pmax(Mfg_Current_Month_Fcst, Mfg_Next_Month_Fcst) / 20.83) ),
                dos_mfg_na = !is.na(Mfg_DOS),
                Mfg_DOS = ifelse(dos_mfg_na == TRUE, Mfg_DOS, 0),
                Mfg_DOS = round(Mfg_DOS, 1),
                Mfg_DOS = replace(Mfg_DOS, is.infinite(Mfg_DOS), 0)) %>% 
  dplyr::select(-dos_mfg_na)



# calculate - Mfg DOS after CustOrd 
IQR_FG_sample %<>% 
  dplyr::mutate(Mfg_DOS_after_CustOrd = (On_Hand_usable_and_soft_hold - ifelse(OPV == 0 | OPV >= 20, Mfg_CustOrd_in_next_28_days, 
                                                                               ifelse(OPV < 20 & OPV >= 12, Mfg_CustOrd_in_next_21_days, 
                                                                                      ifelse(OPV < 12 & OPV >= 8, Mfg_CustOrd_in_next_14_days, 
                                                                                             Mfg_CustOrd_in_next_7_days))))/
                  pmax((ifelse(OPV == 0 | OPV >= 20, Mfg_CustOrd_in_next_28_days, 
                               ifelse(OPV < 20 & OPV >= 12, Mfg_CustOrd_in_next_21_days,
                                      ifelse(OPV < 12 & OPV >= 8, Mfg_CustOrd_in_next_14_days, Mfg_CustOrd_in_next_7_days)))/OPV),
                       (pmax(Mfg_Current_Month_Fcst, Mfg_Next_Month_Fcst)/20.83))) %>% 
  dplyr::mutate(dos_mfg_na = !is.na(Mfg_DOS_after_CustOrd),
                Mfg_DOS_after_CustOrd = ifelse(dos_mfg_na == TRUE, Mfg_DOS_after_CustOrd, 0),
                Mfg_DOS_after_CustOrd = round(Mfg_DOS_after_CustOrd, 1),
                Mfg_DOS_after_CustOrd = replace(Mfg_DOS_after_CustOrd, is.infinite(Mfg_DOS_after_CustOrd), 0)) %>% 
  dplyr::select(-dos_mfg_na) %>% 
  dplyr::mutate(Mfg_DOS_after_CustOrd = round(Mfg_DOS_after_CustOrd, 1))


# calculate - Mfg_Adjusted_Forward_Max_Inv_DOS
IQR_FG_sample %<>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Max_Inv_DOS =  (Mfg_Adjusted_Forward_Inv_Target / 
                                                       pmax((ifelse(OPV == 0 | OPV >= 20, Mfg_CustOrd_in_next_28_days,
                                                                    ifelse(OPV < 20 & OPV >= 12, Mfg_CustOrd_in_next_21_days,
                                                                           ifelse(OPV < 12 & OPV >= 8, Mfg_CustOrd_in_next_14_days, 
                                                                                  Mfg_CustOrd_in_next_7_days)))
                                                             / OPV),
                                                            (pmax(Mfg_Current_Month_Fcst, Mfg_Next_Month_Fcst) / 20.83)))   ) %>% 
  dplyr::mutate(dos_mfg_na = !is.na(Mfg_Adjusted_Forward_Max_Inv_DOS),
                Mfg_Adjusted_Forward_Max_Inv_DOS = ifelse(dos_mfg_na == TRUE, Mfg_Adjusted_Forward_Max_Inv_DOS, 0),
                Mfg_Adjusted_Forward_Max_Inv_DOS = round(Mfg_Adjusted_Forward_Max_Inv_DOS, 1),
                Mfg_Adjusted_Forward_Max_Inv_DOS = replace(Mfg_Adjusted_Forward_Max_Inv_DOS, is.infinite(Mfg_Adjusted_Forward_Max_Inv_DOS), 0)) %>% 
  dplyr::select(-dos_mfg_na) %>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Max_Inv_DOS = round(Mfg_Adjusted_Forward_Max_Inv_DOS, 1))



# calculate - Mfg Forward Target Inv DOS (fcst only)
IQR_FG_sample %<>% 
  dplyr::mutate(aa = Current_SS / (pmax(Mfg_Current_Month_Fcst, Mfg_Next_Month_Fcst) / 20.83),
                aa = replace(aa, is.na(aa), 0),
                aa = replace(aa, is.nan(aa), 0),
                aa = replace(aa, is.infinite(aa), 0)) %>% 
  dplyr::mutate(Mfg_Forward_Target_Inv_DOS_fcst_only = ifelse(Total_mfg_Forecast_Next_12_Months == 0, 0,
                                                              OPV + aa)) %>% 
  dplyr::mutate(Mfg_Forward_Target_Inv_DOS_fcst_only = round(Mfg_Forward_Target_Inv_DOS_fcst_only, 0)) %>% 
  dplyr::select(-aa)


# calculate - Mfg Adjusted Forward Target Inv DOS (includes Orders)
IQR_FG_sample %<>% 
  dplyr::mutate(aa = Current_SS / pmax((ifelse(OPV == 0 | OPV >= 20, Mfg_CustOrd_in_next_28_days,
                                               ifelse(OPV < 20 & OPV >= 12, Mfg_CustOrd_in_next_21_days,
                                                      ifelse(OPV < 12 & OPV >= 8, Mfg_CustOrd_in_next_14_days, Mfg_CustOrd_in_next_7_days))) / OPV),
                                       (pmax(Current_Month_Fcst, Next_Month_Fcst) / 20.83)),
                aa = replace(aa, is.na(aa), 0),
                aa = replace(aa, is.nan(aa), 0),
                aa = replace(aa, is.infinite(aa), 0)) %>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Target_Inv_DOS_includes_Orders = ifelse(Total_mfg_Forecast_Next_12_Months == 0, 0, 
                                                                             OPV + aa)) %>% 
  dplyr::mutate(Mfg_Adjusted_Forward_Target_Inv_DOS_includes_Orders = round(Mfg_Adjusted_Forward_Target_Inv_DOS_includes_Orders, 0)) %>% 
  dplyr::select(-aa)


# on hand Inv after CustOrd > mfg AF max
IQR_FG_sample %<>% 
  dplyr::mutate(on_hand_Inv_after_CustOrd_greater_mfg_AF_max =  ifelse(On_Hand_usable_and_soft_hold-
                                                                         (ifelse(OPV == 0 | OPV >= 20, Mfg_CustOrd_in_next_28_days,
                                                                                 ifelse(OPV < 20 & OPV >= 12, Mfg_CustOrd_in_next_21_days,
                                                                                        ifelse(OPV < 12 & OPV >= 8, Mfg_CustOrd_in_next_14_days, Mfg_CustOrd_in_next_7_days))))
                                                                       > Mfg_Adjusted_Forward_Inv_Max, 1, 0))




# calculate - Mfg Inv Health
IQR_FG_sample %<>% 
  dplyr::mutate(Mfg_Inv_Health = ifelse(On_Hand_usable_and_soft_hold < Current_SS, "BELOW SS",
                                        ifelse(Mfg_DOS_after_CustOrd > Shippable_Shelf_Life, "AT RISK",
                                               ifelse(Total_mfg_Forecast_Next_12_Months <= 0 & Mfg_CustOrd_in_next_28_days <= 0,
                                                      ifelse(On_Hand_usable_and_soft_hold > 0, "DEAD",
                                                             ifelse(on_hand_Inv_after_CustOrd_greater_mfg_AF_max == 0, "HEALTHY", "EXCESS")),
                                                      ifelse(on_hand_Inv_after_CustOrd_greater_mfg_AF_max == 1, "EXCESS", "HEALTHY")))))




# calculate - has mfg adjusted forward looking Max?
IQR_FG_sample %>% 
  dplyr::mutate(has_mfg_adjusted_forward_looking_Max = ifelse(Mfg_Adjusted_Forward_Inv_Max > 0, 1, 0)) -> IQR_FG_sample




# calculate - on hand Inv > mfg AF max
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_Inv_greater_mfg_AF_max = ifelse(On_Hand_usable_and_soft_hold > Mfg_Adjusted_Forward_Inv_Max, 1, 0)) -> IQR_FG_sample


# calculate - on hand Inv <= mfg AF max
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_Inv_less_or_equal_mfg_AF_max = ifelse(On_Hand_usable_and_soft_hold <= Mfg_Adjusted_Forward_Inv_Max, 1, 0)) -> IQR_FG_sample



# calculate - on hand Inv > mfg Adjusted Forward looking target
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_Inv_greater_mfg_Adjusted_Forward_looking_target = ifelse(On_Hand_usable_and_soft_hold > 
                                                                                   Mfg_Adjusted_Forward_Inv_Target, 1, 0)) -> IQR_FG_sample


# calculate - on hand Inv <= mfg AF target
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_Inv_less_or_equal_mfg_AF_target = ifelse(On_Hand_usable_and_soft_hold <= 
                                                                   Mfg_Adjusted_Forward_Inv_Target, 1, 0)) -> IQR_FG_sample



# calculate - on hand Inv after CustOrd <= mfg AF max
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_Inv_after_CustOrd_less_or_equal_mfg_AF_max =  
                  ifelse(On_Hand_usable_and_soft_hold - (ifelse(OPV == 0 | OPV >= 20, Mfg_CustOrd_in_next_28_days,
                                                                ifelse(OPV < 20 & OPV >= 12, Mfg_CustOrd_in_next_21_days,
                                                                       ifelse(OPV < 12 & OPV >= 8, Mfg_CustOrd_in_next_14_days, Mfg_CustOrd_in_next_7_days))))
                         <= Mfg_Adjusted_Forward_Inv_Max, 1,0)) -> IQR_FG_sample




# calculate - on hand Inv after CustOrd > mfg AF target
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_Inv_after_CustOrd_greater_mfg_AF_target = ifelse(On_Hand_usable_and_soft_hold-(
    ifelse(OPV == 0 | OPV >= 20, Mfg_CustOrd_in_next_28_days,
           ifelse(OPV < 20 & OPV >= 12, Mfg_CustOrd_in_next_21_days,
                  ifelse(OPV < 12 & OPV >= 8, Mfg_CustOrd_in_next_14_days, Mfg_CustOrd_in_next_7_days)))) > Mfg_Adjusted_Forward_Inv_Target,1,0)) -> IQR_FG_sample



# calculate - on hand Inv after CustOrd <= mfg AF target
IQR_FG_sample %>% 
  dplyr::mutate(on_hand_Inv_after_CustOrd_less_or_equal_mfg_AF_target = 
                  ifelse(On_Hand_usable_and_soft_hold-(
                    ifelse(OPV == 0 | OPV >= 20, Mfg_CustOrd_in_next_28_days,
                           ifelse(OPV < 20 & OPV >= 12, Mfg_CustOrd_in_next_21_days,
                                  ifelse(OPV < 12 & OPV >= 8, Mfg_CustOrd_in_next_14_days, Mfg_CustOrd_in_next_7_days))))
                    <= Mfg_Adjusted_Forward_Inv_Target,1,0)) -> IQR_FG_sample



# calculate - on hand inv after mfg 28 days CustOrd > 0
IQR_FG_sample %>%
  dplyr::mutate(on_hand_inv_after_mfg_28_days_CustOrd_greater_0 =
                  ifelse(On_Hand_usable_and_soft_hold - Mfg_CustOrd_in_next_28_days > 0, 1, 0)) -> IQR_FG_sample




######## added 7/5/2023 #########
# wo_2
wo_2 <- read.csv("Z:/IMPORT_JDE_OPENWO.csv",
                 header = FALSE)


wo_2 %>% 
  dplyr::rename(aa = V1) %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8"), sep = "~") %>% 
  dplyr::select(-"3") %>% 
  dplyr::rename(aa = "1") %>%  
  tidyr::separate(aa, c("global", "rp", "Item")) %>% 
  dplyr::select(-global, -rp) %>% 
  dplyr::rename(Location = "2",
                Workorder = "4",
                qty = "5",
                wo_no = "6",
                date = "7",
                wo_no_2 = "8") %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item)) %>% 
  dplyr::mutate(qty = as.double(qty),
                date = as.Date(date)) %>% 
  dplyr::mutate(Location = sub("^0+", "", Location)) %>% 
  
  dplyr::relocate(ref, Item, Location, Workorder) %>% 
  dplyr::mutate(in_next_7_days = ifelse(date >= Sys.Date() & date < Sys.Date()+7, "Y", "N") ) %>% 
  dplyr::mutate(ref = gsub("_", "-", ref)) -> wo_2


# receipt_2
receipt_2 <- read.csv("Z:/IMPORT_RECEIPTS.csv",
                      header = FALSE)


# Base receipt variable
receipt_2 %>% 
  dplyr::rename(aa = V1) %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8"), sep = "~") %>% 
  dplyr::rename(a = "1") %>% 
  tidyr::separate(a, c("global", "rp", "Item")) %>% 
  dplyr::rename(Location = "2",
                scheduled = "4",
                qty = "5",
                receipt_no = "6",
                date = "7",
                receipt_no_2 = "8") %>% 
  dplyr::select(-global, -rp, -"3") %>% 
  dplyr::mutate(Item = gsub("^0+", "", Item),
                Location = gsub("^0+", "", Location)) %>% 
  dplyr::mutate(date = as.Date(date)) %>% 
  readr::type_convert() %>% 
  dplyr::mutate(ref = paste0(Location, "-", Item),
                next_7_days = ifelse(date >= Sys.Date() & date <= Sys.Date() + 7, "Y", "N")) %>% 
  dplyr::rename(item = Item) %>% 
  dplyr::left_join(Campus_ref %>% mutate(Campus = as.character(Campus),
                                         Location = as.character(Location)) %>% select(Location, Campus)) %>% 
  dplyr::mutate(campus_ref = paste0(Campus, "-", item)) %>% 
  dplyr::relocate(ref, campus_ref, Campus, item, Location, scheduled, qty, receipt_no, date, receipt_no_2, next_7_days) -> receipt_2



# Planner Name N/A
IQR_FG_sample %>% 
  dplyr::mutate(Planner_Name = ifelse(is.na(Planner_Name) & Planner == 0, 0, Planner_Name)) -> IQR_FG_sample



# Arrange ----
fg_data_for_arrange <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/7.12.2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 07.12.23.xlsx",
                                  sheet = "FG without BKO BKM TST")

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
                Loc_SKU = gsub("_", "-", Loc_SKU),
                mfg_ref = gsub("_", "-", mfg_ref)) %>%
  dplyr::rename(Campus_Ref = Loc_SKU) %>% 
  dplyr::relocate(Loc, mfg_loc, Campus, Item_2, category, Platform, Macro_Platform, Sub_Type, On_Priority_list, ref, 
                  mfg_ref, Campus_Ref, Base, Label,Description, MTO_MTS, MPF, Planner, Planner_Name, Pack_Size, Formula, 
                  Net_Wt_Lbs, Unit_Cost, JDE_MOQ,
                  Shippable_Shelf_Life, Hold_days, Current_SS, Max_Cycle_Stock, Max_Cycle_Stock_lag_1, 
                  Max_Cycle_Stock_Adjusted_Forward, Max_Cycle_Stock_Mfg_Adjusted_Forward, 
                  Useable, Quality_hold, Quality_hold_in_cost, Soft_Hold, 
                  On_Hand_usable_and_soft_hold, On_Hand_in_pounds, On_Hand_in_cost, On_Hand_Adjusted_Forward_Max_in_cost,
                  On_Hand_Mfg_Adjusted_Forward_Max_in_cost, Forward_Inv_Target_Current_Month_Fcst, 
                  Forward_Inv_Target_Current_Month_Fcst_in_lbs., Forward_Inv_Target_Current_Month_Fcst_in_cost,
                  Adjusted_Forward_Inv_Target, Adjusted_Forward_Inv_Target_in_lbs., Adjusted_Forward_Inv_Target_in_cost,
                  Adjusted_Forward_Inv_Max, Adjusted_Forward_Inv_Max_in_lbs., Adjusted_Forward_Inv_Max_in_cost,
                  Mfg_Adjusted_Forward_Inv_Target, Mfg_Adjusted_Forward_Inv_Target_in_lbs.,
                  Mfg_Adjusted_Forward_Inv_Target_in_cost, Mfg_Adjusted_Forward_Inv_Max, Mfg_Adjusted_Forward_Inv_Max_in_lbs.,
                  Mfg_Adjusted_Forward_Inv_Max_in_cost, Forward_Inv_Target_lag_1_Current_Month_Fcst,
                  Forward_Inv_Target_lag_1_Current_Month_Fcst_in_lbs., Forward_Inv_Target_lag_1_Current_Month_Fcst_in_cost,
                  OPV, CustOrd_in_next_7_days, CustOrd_in_next_14_days, CustOrd_in_next_21_days, CustOrd_in_next_28_days,
                  CustOrd_in_next_28_days_in_cost, Mfg_CustOrd_in_next_7_days, Mfg_CustOrd_in_next_14_days,
                  Mfg_CustOrd_in_next_21_days, Mfg_CustOrd_in_next_28_days, Mfg_CustOrd_in_next_28_days_in_cost,
                  Firm_WO_in_next_28_days, Receipt_in_the_next_28_days, DOS, 
                  DOS_after_CustOrd, Adjusted_Forward_Max_Inv_DOS, Forward_Target_Inv_DOS_fcst_only, 
                  Adjusted_Forward_Target_Inv_DOS_includes_Orders, Mfg_DOS, Mfg_DOS_after_CustOrd,
                  Mfg_Adjusted_Forward_Max_Inv_DOS, Mfg_Forward_Target_Inv_DOS_fcst_only, 
                  Mfg_Adjusted_Forward_Target_Inv_DOS_includes_Orders, Inv_Health, Mfg_Inv_Health, Lag_1_Current_Month_Fcst,
                  Lag_1_Current_Month_Fcst_in_cost, Current_Month_Fcst, Next_Month_Fcst, Mfg_Current_Month_Fcst,
                  Mfg_Next_Month_Fcst, Total_Last_6_mos_Sales, Total_Last_12_mos_Sales, Total_Forecast_Next_12_Months,
                  Total_mfg_Forecast_Next_12_Months, has_adjusted_forward_looking_Max, 
                  on_hand_Inv_greater_AF_max, on_hand_Inv_less_or_equal_AF_max, 
                  on_hand_Inv_greater_Adjusted_Forward_looking_target,
                  on_hand_Inv_less_or_equal_AF_target, on_hand_Inv_after_CustOrd_greater_AF_max,
                  on_hand_Inv_after_CustOrd_less_or_equal_AF_max, on_hand_Inv_after_CustOrd_greater_AF_target, 
                  on_hand_Inv_after_CustOrd_less_or_equal_AF_target, on_hand_inv_after_28_days_CustOrd_greater_0,
                  has_mfg_adjusted_forward_looking_Max, on_hand_Inv_greater_mfg_AF_max, on_hand_Inv_less_or_equal_mfg_AF_max,
                  on_hand_Inv_greater_mfg_Adjusted_Forward_looking_target, on_hand_Inv_less_or_equal_mfg_AF_target,
                  on_hand_Inv_after_CustOrd_greater_mfg_AF_max, on_hand_Inv_after_CustOrd_less_or_equal_mfg_AF_max,
                  on_hand_Inv_after_CustOrd_greater_mfg_AF_target, on_hand_Inv_after_CustOrd_less_or_equal_mfg_AF_target,
                  on_hand_inv_after_mfg_28_days_CustOrd_greater_0) -> IQR_FG_sample


colnames(IQR_FG_sample)[1]<-"Loc"
colnames(IQR_FG_sample)[2]<-"mfg loc"
colnames(IQR_FG_sample)[3]<-"Campus"
colnames(IQR_FG_sample)[4]<-"Item 2"
colnames(IQR_FG_sample)[5]<-"category"
colnames(IQR_FG_sample)[6]<-"Platform"
colnames(IQR_FG_sample)[7]<-"Macro-Platform"
colnames(IQR_FG_sample)[8]<-"Sub Type"
colnames(IQR_FG_sample)[9]<-"On Priority list?"
colnames(IQR_FG_sample)[10]<-"Ref"
colnames(IQR_FG_sample)[11]<-"Mfg Ref"
colnames(IQR_FG_sample)[12]<-"Campus Ref"
colnames(IQR_FG_sample)[13]<-"Base"
colnames(IQR_FG_sample)[14]<-"Label"
colnames(IQR_FG_sample)[15]<-"Description"
colnames(IQR_FG_sample)[16]<-"MTO/MTS"
colnames(IQR_FG_sample)[17]<-"MPF"
colnames(IQR_FG_sample)[18]<-"Planner"
colnames(IQR_FG_sample)[19]<-"Planner Name"
colnames(IQR_FG_sample)[20]<-"Pack Size"
colnames(IQR_FG_sample)[21]<-"Formula"
colnames(IQR_FG_sample)[22]<-"Net Wt Lbs"
colnames(IQR_FG_sample)[23]<-"Unit Cost"
colnames(IQR_FG_sample)[24]<-"JDE MOQ"
colnames(IQR_FG_sample)[25]<-"Shippable Shelf Life"
colnames(IQR_FG_sample)[26]<-"Hold days"
colnames(IQR_FG_sample)[27]<-"Current SS"
colnames(IQR_FG_sample)[28]<-"Max Cycle Stock"
colnames(IQR_FG_sample)[29]<-"Max Cycle Stock lag 1"
colnames(IQR_FG_sample)[30]<-"Max Cycle Stock Adjusted Forward "
colnames(IQR_FG_sample)[31]<-"Max Cycle Stock Mfg Adjusted Forward "
colnames(IQR_FG_sample)[32]<-"Usable"
colnames(IQR_FG_sample)[33]<-"Quality hold"
colnames(IQR_FG_sample)[34]<-"Quality hold in $$"
colnames(IQR_FG_sample)[35]<-"Soft Hold"
colnames(IQR_FG_sample)[36]<-"On Hand (usable + soft hold)"
colnames(IQR_FG_sample)[37]<-"On Hand in pounds"
colnames(IQR_FG_sample)[38]<-"On Hand in $$"
colnames(IQR_FG_sample)[39]<-"On Hand - Adjusted Forward Max in $$"
colnames(IQR_FG_sample)[40]<-"On Hand - Mfg Adjusted Forward Max in $$"
colnames(IQR_FG_sample)[41]<-"Forward Inv Target Current Month Fcst"
colnames(IQR_FG_sample)[42]<-"Forward Inv Target Current Month Fcst in lbs."
colnames(IQR_FG_sample)[43]<-"Forward Inv Target Current Month Fcst in $$"
colnames(IQR_FG_sample)[44]<-"Adjusted Forward Inv Target"
colnames(IQR_FG_sample)[45]<-"Adjusted Forward Inv Target in lbs."
colnames(IQR_FG_sample)[46]<-"Adjusted Forward Inv Target in $$"
colnames(IQR_FG_sample)[47]<-"Adjusted Forward Inv Max"
colnames(IQR_FG_sample)[48]<-"Adjusted Forward Inv Max in lbs."
colnames(IQR_FG_sample)[49]<-"Adjusted Forward Inv Max in $$"
colnames(IQR_FG_sample)[50]<-"Mfg Adjusted Forward Inv Target"
colnames(IQR_FG_sample)[51]<-"Mfg Adjusted Forward Inv Target in lbs."
colnames(IQR_FG_sample)[52]<-"Mfg Adjusted Forward Inv Target in $$"
colnames(IQR_FG_sample)[53]<-"Mfg Adjusted Forward Inv Max"
colnames(IQR_FG_sample)[54]<-"Mfg Adjusted Forward Inv Max in lbs."
colnames(IQR_FG_sample)[55]<-"Mfg Adjusted Forward Inv Max in $$"
colnames(IQR_FG_sample)[56]<-"Forward Inv Target lag 1 Current Month Fcst"
colnames(IQR_FG_sample)[57]<-"Forward Inv Target lag 1 Current Month Fcst in lbs."
colnames(IQR_FG_sample)[58]<-"Forward Inv Target lag 1 Current Month Fcst in $$"
colnames(IQR_FG_sample)[59]<-"OPV"
colnames(IQR_FG_sample)[60]<-"CustOrd in next 7 days"
colnames(IQR_FG_sample)[61]<-"CustOrd in next 14 days"
colnames(IQR_FG_sample)[62]<-"CustOrd in next 21 days"
colnames(IQR_FG_sample)[63]<-"CustOrd in next 28 days"
colnames(IQR_FG_sample)[64]<-"CustOrd in next 28 days in $$"
colnames(IQR_FG_sample)[65]<-"Mfg CustOrd in next 7 days"
colnames(IQR_FG_sample)[66]<-"Mfg CustOrd in next 14 days"
colnames(IQR_FG_sample)[67]<-"Mfg CustOrd in next 21 days"
colnames(IQR_FG_sample)[68]<-"Mfg CustOrd in next 28 days"
colnames(IQR_FG_sample)[69]<-"Mfg CustOrd in next 28 days in $$"
colnames(IQR_FG_sample)[70]<-"Firm WO in next 28 days"
colnames(IQR_FG_sample)[71]<-"Receipt in the next 28 days"
colnames(IQR_FG_sample)[72]<-"DOS"
colnames(IQR_FG_sample)[73]<-"DOS after CustOrd"
colnames(IQR_FG_sample)[74]<-"Adjusted Forward Max Inv DOS"
colnames(IQR_FG_sample)[75]<-"Forward Target Inv DOS (fcst only)"
colnames(IQR_FG_sample)[76]<-"Adjusted Forward Target Inv DOS (includes Orders)"
colnames(IQR_FG_sample)[77]<-"Mfg DOS"
colnames(IQR_FG_sample)[78]<-"Mfg DOS after CustOrd"
colnames(IQR_FG_sample)[79]<-"Mfg Adjusted Forward Max Inv DOS"
colnames(IQR_FG_sample)[80]<-"Mfg Forward Target Inv DOS (fcst only)"
colnames(IQR_FG_sample)[81]<-"Mfg Adjusted Forward Target Inv DOS (includes Orders)"
colnames(IQR_FG_sample)[82]<-"Inv Health"
colnames(IQR_FG_sample)[83]<-"Mfg Inv Health"
colnames(IQR_FG_sample)[84]<-"Lag 1 Current Month Fcst"
colnames(IQR_FG_sample)[85]<-"Lag 1 Current Month Fcst in $$"
colnames(IQR_FG_sample)[86]<-"Current Month Fcst"
colnames(IQR_FG_sample)[87]<-"Next Month Fcst"
colnames(IQR_FG_sample)[88]<-"Mfg Current Month Fcst"
colnames(IQR_FG_sample)[89]<-"Mfg Next Month Fcst"
colnames(IQR_FG_sample)[90]<-"Total Last 6 mos Sales"
colnames(IQR_FG_sample)[91]<-"Total Last 12 mos Sales "
colnames(IQR_FG_sample)[92]<-"Total Forecast Next 12 Months"
colnames(IQR_FG_sample)[93]<-"Total mfg Forecast Next 12 Months"
colnames(IQR_FG_sample)[94]<-"has adjusted forward looking Max?"
colnames(IQR_FG_sample)[95]<-"on hand Inv > AF max"
colnames(IQR_FG_sample)[96]<-"on hand Inv <= AF max"
colnames(IQR_FG_sample)[97]<-"on hand Inv > Adjusted Forward looking target"
colnames(IQR_FG_sample)[98]<-"on hand Inv <= AF target"
colnames(IQR_FG_sample)[99]<-"on hand Inv after CustOrd > AF max"
colnames(IQR_FG_sample)[100]<-"on hand Inv after CustOrd <= AF max"
colnames(IQR_FG_sample)[101]<-"on hand Inv after CustOrd > AF target"
colnames(IQR_FG_sample)[102]<-"on hand Inv after CustOrd <= AF target"
colnames(IQR_FG_sample)[103]<-"on hand inv after 28 days CustOrd > 0"
colnames(IQR_FG_sample)[104]<-"has mfg adjusted forward looking Max?"
colnames(IQR_FG_sample)[105]<-"on hand Inv > mfg AF max"
colnames(IQR_FG_sample)[106]<-"on hand Inv <= mfg AF max"
colnames(IQR_FG_sample)[107]<-"on hand Inv > mfg Adjusted Forward looking target"
colnames(IQR_FG_sample)[108]<-"on hand Inv <= mfg AF target"
colnames(IQR_FG_sample)[109]<-"on hand Inv after CustOrd > mfg AF max"
colnames(IQR_FG_sample)[110]<-"on hand Inv after CustOrd <= mfg AF max"
colnames(IQR_FG_sample)[111]<-"on hand Inv after CustOrd > mfg AF target"
colnames(IQR_FG_sample)[112]<-"on hand Inv after CustOrd <= mfg AF target"
colnames(IQR_FG_sample)[113]<-"on hand inv after mfg 28 days CustOrd > 0"


# (Path Revision Needed)
writexl::write_xlsx(wo_2, "wo.xlsx")
writexl::write_xlsx(receipt_2, "receipt.xlsx")
writexl::write_xlsx(IQR_FG_sample, "IQR_FG_report_071223.xlsx")


file.rename(from="C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/IQR/venturafoods_RPA_IQR/IQR_FG_report_071223.xlsx",
            to="C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/7.12.2023/IQR_FG_Report_071223.xlsx")


file.rename(from="C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/IQR/venturafoods_RPA_IQR/wo.xlsx",
            to="C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/7.12.2023/wo.xlsx")

file.rename(from="C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/IQR/venturafoods_RPA_IQR/receipt.xlsx",
            to="C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/7.12.2023/receipt.xlsx")


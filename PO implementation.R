# (Path Revision Needed) Custord PO ----
po <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/05.14.2024/po.xlsx",
                 sheet = "Daily Open PO")


po %>% 
  janitor::clean_names() %>% 
  dplyr::rename(loc_sku = loc_item) %>% 
  dplyr::mutate(loc_sku = gsub("-", "_", loc_sku)) 
  dplyr::select(loc_sku, )


  
## hmm.. I want to figure below things out. 
  ## - date column: which one is the one to use? 
  ## - Qty, and transfer? what column would be the matched one. 
  
  

po %>% 
  dplyr::select(-1) %>% 
  dplyr::slice(-1) %>% 
  dplyr::rename(aa = V2) %>% 
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
  dplyr::mutate(next_28_days = ifelse(date >= specific_date & date <= specific_date + 28, "Y", "N")) -> po


reshape2::dcast(po, ref ~ next_28_days, value.var = "qty", sum) %>% 
  dplyr::rename(loc_sku = ref) -> po_pivot

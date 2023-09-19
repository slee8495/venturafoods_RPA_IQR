DSX_pivot_1 %>% 
  dplyr::select(1, ncol(DSX_pivot_1)) %>% 
  dplyr::mutate(total_12_month = round(total_12_month, 0)) %>% 
  dplyr::filter(ref == "208_17974CGS")

DSX_mfg_pivot_1 %>% 
  dplyr::select(1, ncol(DSX_mfg_pivot_1)) %>% 
  dplyr::mutate(total_12_month = round(total_12_month, 0)) %>% 
  dplyr::filter(mfg_ref == "-1_17974CGS")


650
932
986

IQR_FG_sample %>% 
  dplyr::filter(item_2 == "17974CGS")

I %>% 
  dplyr::filter(item_2 == "17974CGS") %>% 
  dplyr::select(total_mfg_12_month, total_12_month)

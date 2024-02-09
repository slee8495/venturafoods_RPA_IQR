custord

custord %>%
  gather(key = "timeframe", value = "value", starts_with("in_next_")) %>%
  group_by(ref, timeframe) %>%
  summarise(Count_of_Y = sum(value == "Y", na.rm = TRUE)) %>%
  spread(key = timeframe, value = Count_of_Y) %>% 
  dplyr::select(ref, in_next_7_days, in_next_14_days, in_next_21_days, in_next_28_days) -> data

data.frame(
  ref = "Total",
  in_next_7_days = sum(data$in_next_7_days),
  in_next_14_days = sum(data$in_next_14_days),
  in_next_21_days = sum(data$in_next_21_days),
  in_next_28_days = sum(data$in_next_28_days)
)


m
ref in_next_7_days in_next_14_days in_next_21_days in_next_28_days
1 Total          12778           18799           20380           20652

ref in_next_7_days in_next_14_days in_next_21_days in_next_28_days
1 Total          13147           19302           20906           21202

ref in_next_7_days in_next_14_days in_next_21_days in_next_28_days
1 Total          13187           19342           20946           21242



custord_pivot_1$Y %>% sum()
custord_pivot_2$Y %>% sum()
custord_pivot_3$Y %>% sum()
custord_pivot_4$Y %>% sum()



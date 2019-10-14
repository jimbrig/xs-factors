ncci_data <- bind_rows(ncci_loss_data, ncci_lossalae_data)

if (!dir.exists("cache")) dir.create("cache")
fst::write_fst(ncci_data, "cache/ncci_data_10-19")

out <- ncci_data %>%
  pivot_wider(id_cols = c(hg, type, limit), names_from = state, values_from = elf)

if (!dir.exists("output")) dir.create("output")
write.xlsx(out, "output/ncci_excess_factors_10-19.xlsx")

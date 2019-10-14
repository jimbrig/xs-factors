alae_files_meta <- files_meta %>%
  filter(!is.na(alae_tab))

alae_files_in <- alae_files_meta %>%
  pull(local_path_xlsx) %>%
  set_names(alae_files_meta$state)

ncci_lossalae_data <- pmap_dfr(list(xlsxFile = alae_files_in,
                                    sheet = alae_files_meta$alae_tab,
                                    startRow = alae_files_meta$start_row,
                                    rows = alae_files_meta$rows,
                                    cols = alae_files_meta$cols),
                               openxlsx::read.xlsx,
                               colNames = FALSE,
                               na.strings = c("NA", "n/a"),
                               .id = "state") %>%
  set_names(c("state", "limit", "elf")) %>%
  tibble::add_column(hg = rep(rep(LETTERS[1:7], each = 40), length(alae_files_in))) %>%
  left_join(select(alae_files_meta, state, eff_date), by = "state") %>%
  mutate(type = "lossalae") %>%
  select(state, eff_date, hg, type, limit, elf)

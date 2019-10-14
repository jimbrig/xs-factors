
#  ------------------------------------------------------------------------
#
# Title : NCCI Raw Data Extraction
#    By : Jimmy Briggs
#  Date : 2019-10-10
#
#  ------------------------------------------------------------------------

files_in <- files_meta$local_path_xlsx %>% set_names(files_meta$state)

ncci_loss_data <- pmap_dfr(list(xlsxFile = files_in,
                                sheet = files_meta$loss_tab,
                                startRow = files_meta$start_row,
                                rows = files_meta$rows,
                                cols = files_meta$cols),
                           openxlsx::read.xlsx,
                           colNames = FALSE,
                           na.strings = c("NA", "n/a"),
                           .id = "state") %>%
  set_names(c("state", "limit", "elf")) %>%
  tibble::add_column(hg = rep(rep(LETTERS[1:7], each = 40), length(files_in))) %>%
  left_join(select(files_meta, state, eff_date), by = "state") %>%
  mutate(type = "loss") %>%
  select(state, eff_date, hg, type, limit, elf)






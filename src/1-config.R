# packages ----------------------------------------------------------------
if (!require(pacman)) install.packages("pacman")
pacman::p_load(openxlsx,
               RDCOMClient,
               readxl,
               tidyr,
               lubridate,
               fs,
               fst,
               tidyverse)

# load data from network --------------------------------------------------
netpath <- fs::path("H:\\ATLRFI", "INDUSTRY", "Rates - Loss Costs",
                    "Exposure_Adjustment_Template", "2019", "Backup",
                    "Excess Loss Factors", "Raw Data", "NCCI")

# make dirs
dirs <- c("data", "data/ncci", "data/ncci/xls", "data/ncci/xlsx", "output", "cache")
fs::dir_create(dirs)

# copy to local project
fs::dir_copy(netpath, "data/ncci/xls", overwrite = TRUE)

# convert to xlsx
source("lib/xls_to_xlsx.R")
convert_xls_to_xlsx(in_folder = "data/ncci/xls",
                    out_folder = "data/ncci/xlsx")

# create metadata table with info on files
raw_meta <- fs::dir_info(netpath) %>%
  tibble::as_tibble() %>%
  select(path, size, modification_time, change_time, birth_time, access_time) %>%
  mutate(netpath = dirname(path),
         file = basename(path),
         state = stringr::str_sub(file, 1, 2),
         eff_date = stringr::str_sub(file, 3, 10) %>% lubridate::ymd(.),
         local_path_xls = paste0("data/ncci/xls/", file),
         local_path_xlsx = paste0("data/ncci/xlsx/", file, "x"),
         sheets = map(local_path_xlsx, ~ readxl::excel_sheets(.)),
         has_e9 = map_lgl(sheets, ~ if_else("Exhibit 9" %in% ., TRUE, FALSE)),
         has_e10 = map_lgl(sheets, ~ if_else("Exhibit 10" %in% ., TRUE, FALSE)),
         loss_tab = case_when(state == "TX" ~ "Exhibit VIII",
                              TRUE ~ "Exhibit 9"),
         alae_tab = case_when(state == "TX" ~ "Exhibit IX",
                              state %in% c("KY", "GA", "LA", "MD",
                                           "OR", "SD", "VA") ~ NA_character_,
                              TRUE ~ "Exhibit 10"),
         cols = case_when(state == "TX" ~ list(c(2, 19)),
                          state %in% c("DC", "KY", "WV") ~ list(c(2, 11)),
                          TRUE ~ list(c(2, 17))),
         rows = case_when(state == "TX" ~ list(c(13:52, 71:110, 129:168, 187:226,
                                                 245:284, 303:342, 361:400)),
                          state %in% c("DC", "KY", "WV") ~ list(c(
                            12:51, 75:114, 138:177, 201:240, 264:303, 327:366,
                            390:429
                          )),
                          TRUE ~ list(c(
                            12:51, 69:108, 126:165, 183:222, 240:279, 297:336,
                            354:393
                          ))),
         start_row = case_when(state == "TX" ~ 13,
                               TRUE ~ 12))

saveRDS(files_meta, "cache/files_meta-2019.RDS")

# ============================================================
# OpenMusE Statistical Yearbook – Prices (HICP) Module
#
# Tables produced (EU27 + by country):
#  - Total annual average index for musical goods and services
#              (Eurostat unit: INX_A_AVG)
#  - Total annual average rate of change for musical goods and services
#              (Eurostat unit: RCH_A_AVG)
#
# Dataset: Eurostat prc_hicp_aind
# Dimensions: freq, unit, coicop, geo, time, values
# Frequency used: A (annual)
#
# Years shown: 2021–2023 (column order: 2023, 2022, 2021)
#
# Naming convention:
#  - Tables:  HICP_<EU27|CTY>_<T16|T17>_<INDEX|RATE>
#  - Legends: HICP_LEGEND_<EU27|CTY>_<T16|T17>_<INDEX|RATE>
# ============================================================

suppressPackageStartupMessages({
  library(eurostat)
  library(dplyr)
  library(tidyr)
  library(readr)
  library(tibble)
})

# ------------------------------------------------------------
# 0) Run switchboard
# ------------------------------------------------------------
RUN <- list(
  eu_tables      = TRUE,
  country_tables = TRUE,
  write_csv      = TRUE,
  write_excel    = TRUE
)

# ------------------------------------------------------------
# 1) Configuration
# ------------------------------------------------------------
cfg <- list(
  outdir = "OpenMusE_YEARLY_STATISTICAL_BOOK",
  dataset_id = "prc_hicp_aind",
  freq = "A",
  
  years_out = c(2023, 2022, 2021),
  geo_eu27  = "EU27_2020",
  
  countries_eu27 = c(
    "AT","BE","BG","CY","CZ","DE","DK","EE","EL","ES","FI","FR","HR","HU",
    "IE","IT","LT","LU","LV","MT","NL","PL","PT","RO","SE","SI","SK"
  ),
  
  # Units confirmed in your data:
  unit_index = "INX_A_AVG",
  unit_rate  = "RCH_A_AVG",
  
  coicop = c("CP09111","CP09113","CP09119","CP09141","CP0915","CP09221","CP09421","CP09423")
)

if (!dir.exists(cfg$outdir)) dir.create(cfg$outdir, recursive = TRUE, showWarnings = FALSE)
stopifnot(dir.exists(cfg$outdir))

if (isTRUE(RUN$write_excel)) {
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Package 'openxlsx' is required for Excel export. Install with: install.packages('openxlsx')")
  }
}

# ------------------------------------------------------------
# 2) Local labels (NO label_eurostat to avoid codelist 404)
# ------------------------------------------------------------
coicop_labels <- tibble::tibble(
  coicop = cfg$coicop,
  coicop_label = c(
    "Equipment for the reception, recording and reproduction of sound",
    "Portable sound and vision devices",
    "Other equipment for the reception, recording and reproduction of sound and picture",
    "Pre-recorded recording media",
    "Repair of audio-visual, photographic and information processing equipment",
    "Musical instruments",
    "Cinemas, theatres, concerts",
    "Television and radio licence fees, subscriptions"
  )
)

# ------------------------------------------------------------
# 3) Helpers
# ------------------------------------------------------------
add_year <- function(df, time_col = "time") {
  df %>% mutate(year = suppressWarnings(as.integer(substr(as.character(.data[[time_col]]), 1, 4))))
}

format_table_values <- function(df, id_cols, digits = 2) {
  out <- df
  num_cols <- setdiff(names(out), id_cols)
  for (cc in num_cols) if (is.numeric(out[[cc]])) out[[cc]] <- round(out[[cc]], digits)
  out
}

export_csv <- function(x, filename) {
  path <- file.path(cfg$outdir, filename)
  readr::write_csv(x, path, na = "")
  invisible(path)
}

# Excel sheet names must be unique (case-insensitive) and <= 31 chars
make_unique_sheet_names <- function(x) {
  out <- character(length(x))
  used <- character(0)
  for (i in seq_along(x)) {
    base <- substr(x[i], 1, 31)
    nm <- base
    k <- 1
    while (tolower(nm) %in% tolower(used)) {
      suffix <- paste0("_", k)
      nm <- substr(base, 1, max(0, 31 - nchar(suffix)))
      nm <- paste0(nm, suffix)
      k <- k + 1
    }
    out[i] <- nm
    used <- c(used, nm)
  }
  out
}

safe_get <- function(filters) {
  tryCatch(
    eurostat::get_eurostat(cfg$dataset_id, filters = filters, time_format = "num", cache = TRUE),
    error = function(e) tibble::tibble()
  )
}

make_legend <- function(df_raw, table_name, geo_coverage, metric, unit_code, details) {
  tibble::tibble(
    table        = table_name,
    dataset_id   = cfg$dataset_id,
    geo_coverage = geo_coverage,
    freq         = cfg$freq,
    unit         = unit_code,
    years_shown  = paste(cfg$years_out, collapse = ", "),
    coicop       = paste(cfg$coicop, collapse = ", "),
    metric       = metric,
    details      = details
  )
}

# ------------------------------------------------------------
# 4) Core pull + builders
# ------------------------------------------------------------
get_hicp <- function(geo_vec, unit_code) {
  df <- safe_get(list(
    geo    = geo_vec,
    unit   = unit_code,
    coicop = cfg$coicop,
    freq   = cfg$freq
  ))
  if (nrow(df) == 0) {
    stop("No data returned for unit=", unit_code, ". Check geo/coicop/freq/years availability.")
  }
  df %>%
    add_year() %>%
    filter(year %in% cfg$years_out)
}

build_eu_table <- function(df, table_years = cfg$years_out) {
  df %>%
    filter(geo == cfg$geo_eu27, year %in% table_years) %>%
    left_join(coicop_labels, by = "coicop") %>%
    mutate(
      row  = paste0(coicop, " ", coicop_label),
      year = as.character(year)
    ) %>%
    select(row, year, value = values) %>%
    distinct() %>%
    pivot_wider(names_from = year, values_from = value) %>%
    select(row, all_of(as.character(table_years)))
}

build_cty_table <- function(df, table_years = cfg$years_out) {
  df %>%
    filter(geo %in% cfg$countries_eu27, year %in% table_years) %>%
    left_join(coicop_labels, by = "coicop") %>%
    mutate(
      row  = paste0(coicop, " ", coicop_label),
      year = as.character(year)
    ) %>%
    select(country = geo, row, year, value = values) %>%
    distinct() %>%
    pivot_wider(names_from = year, values_from = value) %>%
    select(country, row, all_of(as.character(table_years))) %>%
    arrange(country, row)
}

# ------------------------------------------------------------
# 5) Pipeline
# ------------------------------------------------------------
tables  <- list()
legends <- list()

# --- Table 16: INDEX (INX_A_AVG) ---
df_index <- get_hicp(
  geo_vec = c(cfg$geo_eu27, cfg$countries_eu27),
  unit_code = cfg$unit_index
)

if (isTRUE(RUN$eu_tables)) {
  t16_eu <- build_eu_table(df_index)
  tables[["HICP_EU27_T16_INDEX"]] <- format_table_values(t16_eu, "row", digits = 2)
  
  legends[["HICP_LEGEND_EU27_T16_INDEX"]] <- make_legend(
    df_index,
    "HICP_EU27_T16_INDEX",
    "EU27_2020",
    "Annual average HICP index (level)",
    cfg$unit_index,
    "Direct from Eurostat prc_hicp_aind with unit=INX_A_AVG for the specified COICOP categories."
  )
}

if (isTRUE(RUN$country_tables)) {
  t16_cty <- build_cty_table(df_index)
  tables[["HICP_CTY_T16_INDEX"]] <- format_table_values(t16_cty, c("country","row"), digits = 2)
  
  legends[["HICP_LEGEND_CTY_T16_INDEX"]] <- make_legend(
    df_index,
    "HICP_CTY_T16_INDEX",
    "EU27 countries",
    "Annual average HICP index (level)",
    cfg$unit_index,
    "Direct from Eurostat prc_hicp_aind with unit=INX_A_AVG for the specified COICOP categories, by country."
  )
}

# --- Table 17: RATE OF CHANGE (RCH_A_AVG) ---
df_rate <- get_hicp(
  geo_vec = c(cfg$geo_eu27, cfg$countries_eu27),
  unit_code = cfg$unit_rate
)

if (isTRUE(RUN$eu_tables)) {
  t17_eu <- build_eu_table(df_rate)
  tables[["HICP_EU27_T17_RATE"]] <- format_table_values(t17_eu, "row", digits = 2)
  
  legends[["HICP_LEGEND_EU27_T17_RATE"]] <- make_legend(
    df_rate,
    "HICP_EU27_T17_RATE",
    "EU27_2020",
    "Annual average rate of change (%, YoY)",
    cfg$unit_rate,
    "Direct from Eurostat prc_hicp_aind with unit=RCH_A_AVG (annual average rate of change) for the specified COICOP categories."
  )
}

if (isTRUE(RUN$country_tables)) {
  t17_cty <- build_cty_table(df_rate)
  tables[["HICP_CTY_T17_RATE"]] <- format_table_values(t17_cty, c("country","row"), digits = 2)
  
  legends[["HICP_LEGEND_CTY_T17_RATE"]] <- make_legend(
    df_rate,
    "HICP_CTY_T17_RATE",
    "EU27 countries",
    "Annual average rate of change (%, YoY)",
    cfg$unit_rate,
    "Direct from Eurostat prc_hicp_aind with unit=RCH_A_AVG (annual average rate of change) for the specified COICOP categories, by country."
  )
}

# ------------------------------------------------------------
# 6) Exports
# ------------------------------------------------------------
if (isTRUE(RUN$write_csv)) {
  for (nm in names(tables))  export_csv(tables[[nm]],  paste0(nm, ".csv"))
  for (nm in names(legends)) export_csv(legends[[nm]], paste0(nm, ".csv"))
}

if (isTRUE(RUN$write_excel)) {
  wb <- openxlsx::createWorkbook()
  
  t_names  <- names(tables)
  t_sheets <- make_unique_sheet_names(paste0("T_", t_names))
  for (i in seq_along(t_names)) {
    openxlsx::addWorksheet(wb, t_sheets[i])
    openxlsx::writeData(wb, t_sheets[i], tables[[t_names[i]]])
  }
  
  l_names  <- names(legends)
  l_sheets <- make_unique_sheet_names(paste0("L_", l_names))
  for (i in seq_along(l_names)) {
    openxlsx::addWorksheet(wb, l_sheets[i])
    openxlsx::writeData(wb, l_sheets[i], legends[[l_names[i]]])
  }
  
  xlsx_path <- file.path(cfg$outdir, "OpenMusE_HICP_Musical_Goods_and_Services.xlsx")
  openxlsx::saveWorkbook(wb, xlsx_path, overwrite = TRUE)
  cat("Excel workbook written to:\n", xlsx_path, "\n", sep = "")
}

cat("HICP module completed. Output folder:\n", cfg$outdir, "\n", sep = "")

# ============================================================
# OpenMusE Statistical Yearbook – Structural Business Statistics (SBS) Module
#
# Tables produced (EU27 + by country):
#  - Number of enterprises
#  - Net turnover (million euro)
#  - Value added (million euro)
#  - Value added per employee (thousand euro)
#
# Coverage:
#  - Cultural sector: CLT aggregate
#  - Music sector (minimum): selected NACE Rev.2 codes
#  - Music sector (maximum): extended set of NACE Rev.2 codes
#
# Datasets:
#  - sbs_ovw_act (CLT aggregate)
#  - sbs_sc_ovw  (NACE-level aggregation)
#
# Dimensions: geo, nace_r2, indic_sbs, size_emp, time, values
# Frequency used: A (annual)
#
# Years shown: 2021–2023 (column order: 2023, 2022, 2021)
#
# Naming convention:
#  - Tables:  SBS_<EU27|CTY>_<TOPIC>_<CULT|MUSIC>
#  - Legends: SBS_LEGEND_<EU27|CTY>_<TOPIC>_<CULT|MUSIC>
# ============================================================


suppressPackageStartupMessages({
  library(eurostat)
  library(dplyr)
  library(tidyr)
  library(readr)
})

# ------------------------------------------------------------
# 0) Run switchboard
# ------------------------------------------------------------
RUN <- list(
  eu27_tables     = TRUE,
  country_tables  = TRUE,
  write_csv       = TRUE,
  write_excel     = TRUE
)

# ------------------------------------------------------------
# 1) Configuration
# ------------------------------------------------------------
cfg <- list(
  outdir = "OpenMusE_YEARLY_STATISTICAL_BOOK",
  years  = c(2023, 2022, 2021),  # required column order
  geo_eu27 = "EU27_2020",
  countries_eu27 = c(
    "AT","BE","BG","CY","CZ","DE","DK","EE","EL","ES","FI","FR","HR","HU",
    "IE","IT","LT","LU","LV","MT","NL","PL","PT","RO","SE","SI","SK"
  ),
  music_min = c("J592","C322"),
  music_max = c("J592","C322","C182","G476","J601","J602","M731","R90")
)

if (!dir.exists(cfg$outdir)) dir.create(cfg$outdir, recursive = TRUE, showWarnings = FALSE)
stopifnot(dir.exists(cfg$outdir))

if (isTRUE(RUN$write_excel)) {
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Package 'openxlsx' is required for Excel export. Install with: install.packages('openxlsx')")
  }
}

# ------------------------------------------------------------
# 2) Indicator codes
# ------------------------------------------------------------
ind_codes <- list(
  enterprises     = "ENT_NR",
  net_turnover    = "NETTUR_MEUR",  # million euro
  value_added     = "AV_MEUR",      # million euro
  employment      = "EMP_NR"
)

# ------------------------------------------------------------
# 3) Helpers
# ------------------------------------------------------------
add_year <- function(df, time_col = "time") {
  df %>% mutate(year = suppressWarnings(as.integer(substr(as.character(.data[[time_col]]), 1, 4))))
}

safe_group_sum <- function(df, keys) {
  grp <- intersect(keys, names(df))
  if (length(grp) == 0) stop("No valid grouping keys found. Requested: ", paste(keys, collapse = ", "))
  df %>%
    group_by(across(all_of(grp))) %>%
    summarise(values = sum(values, na.rm = TRUE), .groups = "drop")
}

keep_first_combo <- function(df) {
  if (all(c("freq","unit") %in% names(df))) {
    combos <- distinct(df, freq, unit)
    if (nrow(combos) > 1) df <- df %>% filter(freq == combos$freq[1], unit == combos$unit[1])
  } else if ("freq" %in% names(df)) {
    freqs <- distinct(df, freq)
    if (nrow(freqs) > 1) df <- df %>% filter(freq == freqs$freq[1])
  }
  df
}

make_legend <- function(df, table_name, concept, dataset_id, transformation, details) {
  tibble::tibble(
    table          = table_name,
    concept        = concept,
    dataset_id     = dataset_id,
    freq           = if ("freq" %in% names(df)) paste(sort(unique(df$freq)), collapse = ", ") else NA_character_,
    unit           = if ("unit" %in% names(df)) paste(sort(unique(df$unit)), collapse = ", ") else NA_character_,
    transformation = transformation,
    details        = details
  )
}

format_table_values <- function(df, id_cols, digits = 3) {
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

# ------------------------------------------------------------
# 4) Core data pulls
# ------------------------------------------------------------
get_culture_clt <- function(geo, indic_vec) {
  get_eurostat(
    id = "sbs_ovw_act",
    filters = list(geo = geo, nace_r2 = "CLT", indic_sbs = indic_vec),
    cache = TRUE
  ) %>%
    add_year() %>%
    filter(year %in% cfg$years) %>%
    keep_first_combo()
}

get_music_sbs <- function(geo, nace_codes, indic_vec) {
  get_eurostat(
    id = "sbs_sc_ovw",
    filters = list(geo = geo, size_emp = "TOTAL", nace_r2 = nace_codes, indic_sbs = indic_vec),
    time_format = "num",
    cache = TRUE
  ) %>%
    add_year() %>%
    filter(year %in% cfg$years) %>%
    keep_first_combo()
}

# ------------------------------------------------------------
# 5) Builders – EU27 3×3
# ------------------------------------------------------------
build_eu27_level_table <- function(table_key, row_titles, indic_code) {

  cult_raw <- get_culture_clt(cfg$geo_eu27, indic_code)
  cult_row <- cult_raw %>%
    distinct(year, values) %>%
    transmute(row = row_titles$culture, year = as.character(year), value = values) %>%
    pivot_wider(names_from = year, values_from = value)

  min_raw <- get_music_sbs(cfg$geo_eu27, cfg$music_min, indic_code)
  min_row <- min_raw %>%
    safe_group_sum(c("year","freq","unit")) %>%
    transmute(row = row_titles$music_min, year = as.character(year), value = values) %>%
    pivot_wider(names_from = year, values_from = value)

  max_raw <- get_music_sbs(cfg$geo_eu27, cfg$music_max, indic_code)
  max_row <- max_raw %>%
    safe_group_sum(c("year","freq","unit")) %>%
    transmute(row = row_titles$music_max, year = as.character(year), value = values) %>%
    pivot_wider(names_from = year, values_from = value)

  tab <- bind_rows(cult_row, min_row, max_row) %>%
    select(row, `2023`, `2022`, `2021`) %>%
    format_table_values(id_cols = "row", digits = 3)

  leg <- bind_rows(
    make_legend(
      cult_raw, paste0("SBS_EU27_", table_key, "_3x3"), row_titles$culture, "sbs_ovw_act",
      "Cultural sector uses CLT aggregate (nace_r2=CLT) from sbs_ovw_act",
      paste0("geo=", cfg$geo_eu27, "; nace_r2=CLT; indic_sbs=", indic_code, "; years=", paste(cfg$years, collapse = ","))
    ),
    make_legend(
      min_raw, paste0("SBS_EU27_", table_key, "_3x3"), row_titles$music_min, "sbs_sc_ovw",
      "Aggregated over NACE codes (music minimum) from sbs_sc_ovw; size_emp=TOTAL",
      paste0("geo=", cfg$geo_eu27, "; size_emp=TOTAL; indic_sbs=", indic_code, "; nace_r2=", paste(cfg$music_min, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    ),
    make_legend(
      max_raw, paste0("SBS_EU27_", table_key, "_3x3"), row_titles$music_max, "sbs_sc_ovw",
      "Aggregated over NACE codes (music maximum) from sbs_sc_ovw; size_emp=TOTAL",
      paste0("geo=", cfg$geo_eu27, "; size_emp=TOTAL; indic_sbs=", indic_code, "; nace_r2=", paste(cfg$music_max, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    )
  )

  list(table = tab, legend = leg)
}

build_eu27_va_per_emp <- function() {

  # Culture CLT: compute (AV_MEUR * 1000) / EMP_NR
  cult_raw <- get_culture_clt(cfg$geo_eu27, c(ind_codes$value_added, ind_codes$employment))
  cult_tot <- cult_raw %>%
    safe_group_sum(c("year","indic_sbs","freq","unit")) %>%
    select(year, indic_sbs, values) %>%
    pivot_wider(names_from = indic_sbs, values_from = values) %>%
    mutate(value = (`AV_MEUR` * 1000) / `EMP_NR`)

  # Music min/max: same computation from sums over NACE
  min_raw <- get_music_sbs(cfg$geo_eu27, cfg$music_min, c(ind_codes$value_added, ind_codes$employment))
  min_tot <- min_raw %>%
    safe_group_sum(c("year","indic_sbs","freq","unit")) %>%
    select(year, indic_sbs, values) %>%
    pivot_wider(names_from = indic_sbs, values_from = values) %>%
    mutate(value = (`AV_MEUR` * 1000) / `EMP_NR`)

  max_raw <- get_music_sbs(cfg$geo_eu27, cfg$music_max, c(ind_codes$value_added, ind_codes$employment))
  max_tot <- max_raw %>%
    safe_group_sum(c("year","indic_sbs","freq","unit")) %>%
    select(year, indic_sbs, values) %>%
    pivot_wider(names_from = indic_sbs, values_from = values) %>%
    mutate(value = (`AV_MEUR` * 1000) / `EMP_NR`)

  cult_row <- cult_tot %>% transmute(row = "Value added per employee, cultural sector (CLT, thousand euro)", year = as.character(year), value = value) %>%
    pivot_wider(names_from = year, values_from = value)
  min_row  <- min_tot  %>% transmute(row = "Value added per employee, music sector minimum (thousand euro)", year = as.character(year), value = value) %>%
    pivot_wider(names_from = year, values_from = value)
  max_row  <- max_tot  %>% transmute(row = "Value added per employee, music sector maximum (thousand euro)", year = as.character(year), value = value) %>%
    pivot_wider(names_from = year, values_from = value)

  tab <- bind_rows(cult_row, min_row, max_row) %>%
    select(row, `2023`, `2022`, `2021`) %>%
    format_table_values(id_cols = "row", digits = 3)

  leg <- bind_rows(
    make_legend(
      cult_raw, "SBS_EU27_va_per_employee_3x3", "Value added per employee, cultural sector (CLT, thousand euro)", "sbs_ovw_act",
      "Computed as (AV_MEUR * 1000) / EMP_NR using CLT aggregate (nace_r2=CLT)",
      paste0("geo=", cfg$geo_eu27, "; nace_r2=CLT; indic_sbs=AV_MEUR,EMP_NR; years=", paste(cfg$years, collapse = ","))
    ),
    make_legend(
      min_raw, "SBS_EU27_va_per_employee_3x3", "Value added per employee, music sector minimum (thousand euro)", "sbs_sc_ovw",
      "Computed as (sum AV_MEUR * 1000) / (sum EMP_NR) over NACE codes; size_emp=TOTAL",
      paste0("geo=", cfg$geo_eu27, "; size_emp=TOTAL; nace_r2=", paste(cfg$music_min, collapse = ","), "; indic_sbs=AV_MEUR,EMP_NR; years=", paste(cfg$years, collapse = ","))
    ),
    make_legend(
      max_raw, "SBS_EU27_va_per_employee_3x3", "Value added per employee, music sector maximum (thousand euro)", "sbs_sc_ovw",
      "Computed as (sum AV_MEUR * 1000) / (sum EMP_NR) over NACE codes; size_emp=TOTAL",
      paste0("geo=", cfg$geo_eu27, "; size_emp=TOTAL; nace_r2=", paste(cfg$music_max, collapse = ","), "; indic_sbs=AV_MEUR,EMP_NR; years=", paste(cfg$years, collapse = ","))
    )
  )

  list(table = tab, legend = leg)
}

# ------------------------------------------------------------
# 6) Builders – Country level
# ------------------------------------------------------------
build_country_level_table <- function(table_key, concept_culture, indic_code) {

  cult_raw <- get_culture_clt(cfg$countries_eu27, indic_code)
  cult_tab <- cult_raw %>%
    distinct(geo, year, values) %>%
    transmute(country = geo, year = as.character(year), value = values) %>%
    pivot_wider(names_from = year, values_from = value) %>%
    select(country, `2023`, `2022`, `2021`) %>%
    arrange(country) %>%
    format_table_values(id_cols = "country", digits = 3)

  cult_leg <- make_legend(
    cult_raw, paste0("SBS_CTY_", table_key, "_culture"), concept_culture, "sbs_ovw_act",
    "Cultural sector uses CLT aggregate (nace_r2=CLT) from sbs_ovw_act",
    paste0("geo=EU27 countries; nace_r2=CLT; indic_sbs=", indic_code, "; years=", paste(cfg$years, collapse = ","))
  )

  get_music_long <- function(nace_codes, group_label) {
    raw <- get_music_sbs(cfg$countries_eu27, nace_codes, indic_code)
    out <- raw %>%
      safe_group_sum(c("geo","year","freq","unit")) %>%
      transmute(country = geo, group = group_label, year = as.character(year), value = values)
    list(out = out, raw = raw)
  }

  minm <- get_music_long(cfg$music_min, "music sector minimum")
  maxm <- get_music_long(cfg$music_max, "music sector maximum")

  music_tab <- bind_rows(minm$out, maxm$out) %>%
    pivot_wider(names_from = year, values_from = value) %>%
    select(country, group, `2023`, `2022`, `2021`) %>%
    arrange(country, group) %>%
    format_table_values(id_cols = c("country","group"), digits = 3)

  music_leg <- bind_rows(
    make_legend(
      minm$raw, paste0("SBS_CTY_", table_key, "_music"), paste0(concept_culture, " (music sector minimum)"), "sbs_sc_ovw",
      "Aggregated over NACE codes (minimum); size_emp=TOTAL",
      paste0("geo=EU27 countries; size_emp=TOTAL; indic_sbs=", indic_code, "; nace_r2=", paste(cfg$music_min, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    ),
    make_legend(
      maxm$raw, paste0("SBS_CTY_", table_key, "_music"), paste0(concept_culture, " (music sector maximum)"), "sbs_sc_ovw",
      "Aggregated over NACE codes (maximum); size_emp=TOTAL",
      paste0("geo=EU27 countries; size_emp=TOTAL; indic_sbs=", indic_code, "; nace_r2=", paste(cfg$music_max, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    )
  )

  list(culture_table = cult_tab, culture_legend = cult_leg,
       music_table = music_tab, music_legend = music_leg)
}

build_country_va_per_emp <- function() {

  cult_raw <- get_culture_clt(cfg$countries_eu27, c(ind_codes$value_added, ind_codes$employment))
  cult_tot <- cult_raw %>%
    safe_group_sum(c("geo","year","indic_sbs","freq","unit")) %>%
    select(geo, year, indic_sbs, values) %>%
    pivot_wider(names_from = indic_sbs, values_from = values) %>%
    mutate(value = (`AV_MEUR` * 1000) / `EMP_NR`) %>%
    transmute(country = geo, year = as.character(year), value = value)

  cult_tab <- cult_tot %>%
    pivot_wider(names_from = year, values_from = value) %>%
    select(country, `2023`, `2022`, `2021`) %>%
    arrange(country) %>%
    format_table_values(id_cols = "country", digits = 3)

  cult_leg <- make_legend(
    cult_raw, "SBS_CTY_va_per_employee_culture", "Value added per employee, cultural sector (CLT, thousand euro)", "sbs_ovw_act",
    "Computed as (AV_MEUR * 1000) / EMP_NR using CLT aggregate (nace_r2=CLT)",
    paste0("geo=EU27 countries; nace_r2=CLT; indic_sbs=AV_MEUR,EMP_NR; years=", paste(cfg$years, collapse = ","))
  )

  get_music_va_emp <- function(nace_codes, group_label) {
    raw <- get_music_sbs(cfg$countries_eu27, nace_codes, c(ind_codes$value_added, ind_codes$employment))
    totals <- raw %>%
      safe_group_sum(c("geo","year","indic_sbs","freq","unit")) %>%
      select(geo, year, indic_sbs, values) %>%
      pivot_wider(names_from = indic_sbs, values_from = values) %>%
      mutate(value = (`AV_MEUR` * 1000) / `EMP_NR`) %>%
      transmute(country = geo, group = group_label, year = as.character(year), value = value)
    list(out = totals, raw = raw)
  }

  minm <- get_music_va_emp(cfg$music_min, "music sector minimum")
  maxm <- get_music_va_emp(cfg$music_max, "music sector maximum")

  music_tab <- bind_rows(minm$out, maxm$out) %>%
    pivot_wider(names_from = year, values_from = value) %>%
    select(country, group, `2023`, `2022`, `2021`) %>%
    arrange(country, group) %>%
    format_table_values(id_cols = c("country","group"), digits = 3)

  music_leg <- bind_rows(
    make_legend(
      minm$raw, "SBS_CTY_va_per_employee_music", "Value added per employee, music sector minimum (thousand euro)", "sbs_sc_ovw",
      "Computed as (sum AV_MEUR * 1000) / (sum EMP_NR) over NACE codes; size_emp=TOTAL",
      paste0("geo=EU27 countries; size_emp=TOTAL; indic_sbs=AV_MEUR,EMP_NR; nace_r2=", paste(cfg$music_min, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    ),
    make_legend(
      maxm$raw, "SBS_CTY_va_per_employee_music", "Value added per employee, music sector maximum (thousand euro)", "sbs_sc_ovw",
      "Computed as (sum AV_MEUR * 1000) / (sum EMP_NR) over NACE codes; size_emp=TOTAL",
      paste0("geo=EU27 countries; size_emp=TOTAL; indic_sbs=AV_MEUR,EMP_NR; nace_r2=", paste(cfg$music_max, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    )
  )

  list(culture_table = cult_tab, culture_legend = cult_leg,
       music_table = music_tab, music_legend = music_leg)
}

# ------------------------------------------------------------
# 7) Pipeline
# ------------------------------------------------------------
tables  <- list()
legends <- list()

# EU27 3×3 tables
if (isTRUE(RUN$eu27_tables)) {

  res <- build_eu27_level_table(
    table_key = "enterprises",
    row_titles = list(
      culture   = "Total number of enterprises, cultural sector (CLT)",
      music_min = "Total number of enterprises, music sector minimum",
      music_max = "Total number of enterprises, music sector maximum"
    ),
    indic_code = ind_codes$enterprises
  )
  tables[["SBS_EU27_enterprises_3x3"]]         <- res$table
  legends[["SBS_LEGEND_EU27_enterprises_3x3"]] <- res$legend

  res <- build_eu27_level_table(
    table_key = "net_turnover",
    row_titles = list(
      culture   = "Total net turnover, cultural sector (CLT, million euro)",
      music_min = "Total net turnover, music sector minimum (million euro)",
      music_max = "Total net turnover, music sector maximum (million euro)"
    ),
    indic_code = ind_codes$net_turnover
  )
  tables[["SBS_EU27_net_turnover_3x3"]]         <- res$table
  legends[["SBS_LEGEND_EU27_net_turnover_3x3"]] <- res$legend

  res <- build_eu27_level_table(
    table_key = "value_added",
    row_titles = list(
      culture   = "Total value added, cultural sector (CLT, million euro)",
      music_min = "Total value added, music sector minimum (million euro)",
      music_max = "Total value added, music sector maximum (million euro)"
    ),
    indic_code = ind_codes$value_added
  )
  tables[["SBS_EU27_value_added_3x3"]]         <- res$table
  legends[["SBS_LEGEND_EU27_value_added_3x3"]] <- res$legend

  res <- build_eu27_va_per_emp()
  tables[["SBS_EU27_va_per_employee_3x3"]]         <- res$table
  legends[["SBS_LEGEND_EU27_va_per_employee_3x3"]] <- res$legend
}

# Country tables
if (isTRUE(RUN$country_tables)) {
  
  res <- build_country_level_table(
    table_key = "enterprises",
    concept_culture = "Total number of enterprises, cultural sector (CLT)",
    indic_code = ind_codes$enterprises
  )
  tables[["SBS_CTY_CULT_enterprises"]]         <- res$culture_table
  legends[["SBS_LEGEND_CTY_CULT_enterprises"]] <- res$culture_legend
  tables[["SBS_CTY_MUS_enterprises"]]          <- res$music_table
  legends[["SBS_LEGEND_CTY_MUS_enterprises"]]  <- res$music_legend
  
  res <- build_country_level_table(
    table_key = "net_turnover",
    concept_culture = "Total net turnover, cultural sector (CLT, million euro)",
    indic_code = ind_codes$net_turnover
  )
  tables[["SBS_CTY_CULT_net_turnover"]]         <- res$culture_table
  legends[["SBS_LEGEND_CTY_CULT_net_turnover"]] <- res$culture_legend
  tables[["SBS_CTY_MUS_net_turnover"]]          <- res$music_table
  legends[["SBS_LEGEND_CTY_MUS_net_turnover"]]  <- res$music_legend
  
  res <- build_country_level_table(
    table_key = "value_added",
    concept_culture = "Total value added, cultural sector (CLT, million euro)",
    indic_code = ind_codes$value_added
  )
  tables[["SBS_CTY_CULT_value_added"]]         <- res$culture_table
  legends[["SBS_LEGEND_CTY_CULT_value_added"]] <- res$culture_legend
  tables[["SBS_CTY_MUS_value_added"]]          <- res$music_table
  legends[["SBS_LEGEND_CTY_MUS_value_added"]]  <- res$music_legend
  
  res <- build_country_va_per_emp()
  tables[["SBS_CTY_CULT_va_per_employee"]]         <- res$culture_table
  legends[["SBS_LEGEND_CTY_CULT_va_per_employee"]] <- res$culture_legend
  tables[["SBS_CTY_MUS_va_per_employee"]]          <- res$music_table
  legends[["SBS_LEGEND_CTY_MUS_va_per_employee"]]  <- res$music_legend
}

# ------------------------------------------------------------
# 8) Exports
# ------------------------------------------------------------
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
  
  xlsx_path <- file.path(cfg$outdir, "OpenMusE_SBS_Core_Tables_and_Legends.xlsx")
  openxlsx::saveWorkbook(wb, xlsx_path, overwrite = TRUE)
  cat("Excel workbook written to:\n", xlsx_path, "\n", sep = "")
}

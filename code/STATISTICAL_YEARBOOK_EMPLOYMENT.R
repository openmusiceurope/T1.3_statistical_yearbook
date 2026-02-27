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
})

# -----------------------------
# 0) RUN SWITCHBOARD
# -----------------------------
RUN <- list(
  eu27_total_employment          = TRUE,  # 3×3: culture + music min/max (thousand persons)
  cty_total_employment           = TRUE,  # country tables: culture; music min/max (thousand persons)
  eu27_emp_per_enterprise        = TRUE,  # 3×3: culture CLT + music min/max (persons per enterprise)
  cty_emp_per_enterprise         = TRUE,  # country tables: culture; music min/max (persons per enterprise)
  write_csv                      = TRUE,
  write_excel                    = TRUE
)

# -----------------------------
# 1) CONFIGURATION
# -----------------------------
cfg <- list(
  outdir         = "OpenMusE_YEARLY_STATISTICAL_BOOK",
  years          = c(2023, 2022, 2021),   # output column order
  geo_eu27       = "EU27_2020",
  countries_eu27 = c(
    "AT","BE","BG","CY","CZ","DE","DK","EE","EL","ES","FI","FR","HR","HU",
    "IE","IT","LT","LU","LV","MT","NL","PL","PT","RO","SE","SI","SK"
  ),
  music_min      = c("J592","C322"),
  music_max      = c("J592","C322","C182","G476","J601","J602","M731","R90")
)

if (!dir.exists(cfg$outdir)) dir.create(cfg$outdir, recursive = TRUE, showWarnings = FALSE)
stopifnot(dir.exists(cfg$outdir))

if (isTRUE(RUN$write_excel)) {
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Package 'openxlsx' is required for Excel export. Install with: install.packages('openxlsx')")
  }
}

# -----------------------------
# 2) HELPERS
# -----------------------------
add_year <- function(df, time_col = "time") {
  stopifnot(time_col %in% names(df))
  df %>%
    mutate(year = suppressWarnings(as.integer(substr(as.character(.data[[time_col]]), 1, 4))))
}

safe_group_sum <- function(df, keys) {
  grp <- intersect(keys, names(df))
  if (length(grp) == 0) stop("No valid grouping keys found. Requested: ", paste(keys, collapse = ", "))
  df %>%
    group_by(across(all_of(grp))) %>%
    summarise(values = sum(values, na.rm = TRUE), .groups = "drop")
}

keep_first_combo <- function(df) {
  # For some datasets Eurostat may return multiple freq/unit combinations
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

format_table_values <- function(df, id_cols, digits) {
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

# Excel sheet names must be unique (case-insensitive) and <= 31 chars.
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

# -----------------------------
# 3) CORE DATA PULLS
# -----------------------------
# Cultural employment (levels): cult_emp_sex; unit THS_PER; sex=T
get_culture_levels <- function(geo) {
  get_eurostat(
    id      = "cult_emp_sex",
    filters = list(geo = geo, sex = "T"),
    cache   = TRUE
  ) %>%
    add_year() %>%
    filter(year %in% cfg$years, unit == "THS_PER") %>%
    keep_first_combo()
}

# Music employment (levels): SBS EMP_NR aggregated over NACE; divide by 1000 to get thousand persons
get_music_levels <- function(geo, nace_codes) {
  get_eurostat(
    id      = "sbs_sc_ovw",
    filters = list(
      geo       = geo,
      indic_sbs = "EMP_NR",
      size_emp  = "TOTAL",
      nace_r2   = nace_codes
    ),
    time_format = "num",
    cache       = TRUE
  ) %>%
    add_year() %>%
    filter(year %in% cfg$years) %>%
    keep_first_combo()
}

# Cultural persons employed per enterprise (CLT): SBS EMP_ENT_NR direct indicator
get_culture_empent <- function(geo) {
  get_eurostat(
    id      = "sbs_ovw_act",
    filters = list(
      geo       = geo,
      nace_r2   = "CLT",
      indic_sbs = "EMP_ENT_NR"
    ),
    cache = TRUE
  ) %>%
    add_year() %>%
    filter(year %in% cfg$years) %>%
    keep_first_combo()
}

# Music persons employed per enterprise: compute sum(EMP_NR)/sum(ENT_NR) over NACE
get_music_empent <- function(geo, nace_codes) {
  get_eurostat(
    id      = "sbs_sc_ovw",
    filters = list(
      geo       = geo,
      size_emp  = "TOTAL",
      nace_r2   = nace_codes,
      indic_sbs = c("EMP_NR", "ENT_NR")
    ),
    time_format = "num",
    cache       = TRUE
  ) %>%
    add_year() %>%
    filter(year %in% cfg$years) %>%
    keep_first_combo()
}

# -----------------------------
# 4) BUILDERS
# -----------------------------
# 4.1 EU27 – total employment (thousand persons), 3×3 table
build_eu27_total_employment_3x3 <- function() {

  cult_raw <- get_culture_levels(cfg$geo_eu27)
  cult_row <- cult_raw %>%
    safe_group_sum(c("year","freq","unit")) %>%
    transmute(
      row   = "Total cultural employment (thousand persons)",
      year  = as.character(year),
      value = values
    ) %>%
    pivot_wider(names_from = year, values_from = value)

  min_raw <- get_music_levels(cfg$geo_eu27, cfg$music_min)
  min_row <- min_raw %>%
    safe_group_sum(c("year","freq","unit")) %>%
    transmute(
      row   = "Total music employment (minimum, thousand persons)",
      year  = as.character(year),
      value = values / 1000
    ) %>%
    pivot_wider(names_from = year, values_from = value)

  max_raw <- get_music_levels(cfg$geo_eu27, cfg$music_max)
  max_row <- max_raw %>%
    safe_group_sum(c("year","freq","unit")) %>%
    transmute(
      row   = "Total music employment (maximum, thousand persons)",
      year  = as.character(year),
      value = values / 1000
    ) %>%
    pivot_wider(names_from = year, values_from = value)

  tab <- bind_rows(cult_row, min_row, max_row) %>%
    select(row, `2023`, `2022`, `2021`) %>%
    format_table_values(id_cols = "row", digits = 3)

  leg <- bind_rows(
    make_legend(
      cult_raw, "EMPL_EU27_TOTAL_EMPLOYMENT_3x3",
      "Total cultural employment (thousand persons)",
      "cult_emp_sex",
      "Filtered to unit=THS_PER; used as reported by Eurostat",
      paste0("geo=", cfg$geo_eu27, "; sex=T; unit=THS_PER; years=", paste(cfg$years, collapse = ","))
    ),
    make_legend(
      min_raw, "EMPL_EU27_TOTAL_EMPLOYMENT_3x3",
      "Total music employment (minimum, thousand persons)",
      "sbs_sc_ovw",
      "Aggregated over NACE codes; EMP_NR divided by 1,000",
      paste0("geo=", cfg$geo_eu27, "; size_emp=TOTAL; indic_sbs=EMP_NR; nace_r2=", paste(cfg$music_min, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    ),
    make_legend(
      max_raw, "EMPL_EU27_TOTAL_EMPLOYMENT_3x3",
      "Total music employment (maximum, thousand persons)",
      "sbs_sc_ovw",
      "Aggregated over NACE codes; EMP_NR divided by 1,000",
      paste0("geo=", cfg$geo_eu27, "; size_emp=TOTAL; indic_sbs=EMP_NR; nace_r2=", paste(cfg$music_max, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    )
  )

  list(tab = tab, leg = leg)
}

# 4.2 CTY – cultural employment (thousand persons)
build_cty_cult_total_employment <- function() {
  raw <- get_culture_levels(cfg$countries_eu27)

  tab <- raw %>%
    safe_group_sum(c("geo","year","freq","unit")) %>%
    transmute(country = geo, year = as.character(year), value = values) %>%
    pivot_wider(names_from = year, values_from = value) %>%
    select(country, `2023`, `2022`, `2021`) %>%
    arrange(country) %>%
    format_table_values(id_cols = "country", digits = 3)

  leg <- make_legend(
    raw, "EMPL_CTY_CULT_TOTAL_EMPLOYMENT",
    "Total cultural employment (thousand persons)",
    "cult_emp_sex",
    "Filtered to unit=THS_PER; used as reported by Eurostat",
    paste0("sex=T; geo=EU27 countries; unit=THS_PER; years=", paste(cfg$years, collapse = ","))
  )

  list(tab = tab, leg = leg)
}

# 4.3 CTY – music employment (thousand persons): one table with group=min/max
build_cty_mus_total_employment <- function() {

  get_one <- function(nace_codes, group_label) {
    raw <- get_music_levels(cfg$countries_eu27, nace_codes)
    out <- raw %>%
      safe_group_sum(c("geo","year","freq","unit")) %>%
      transmute(country = geo, group = group_label, year = as.character(year), value = values / 1000)
    list(out = out, raw = raw)
  }

  minm <- get_one(cfg$music_min, "minimum")
  maxm <- get_one(cfg$music_max, "maximum")

  tab <- bind_rows(minm$out, maxm$out) %>%
    pivot_wider(names_from = year, values_from = value) %>%
    select(country, group, `2023`, `2022`, `2021`) %>%
    arrange(country, group) %>%
    format_table_values(id_cols = c("country","group"), digits = 3)

  leg <- bind_rows(
    make_legend(
      minm$raw, "EMPL_CTY_MUS_TOTAL_EMPLOYMENT",
      "Total music employment (minimum, thousand persons)",
      "sbs_sc_ovw",
      "Aggregated over NACE codes; EMP_NR divided by 1,000",
      paste0("geo=EU27 countries; size_emp=TOTAL; indic_sbs=EMP_NR; nace_r2=", paste(cfg$music_min, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    ),
    make_legend(
      maxm$raw, "EMPL_CTY_MUS_TOTAL_EMPLOYMENT",
      "Total music employment (maximum, thousand persons)",
      "sbs_sc_ovw",
      "Aggregated over NACE codes; EMP_NR divided by 1,000",
      paste0("geo=EU27 countries; size_emp=TOTAL; indic_sbs=EMP_NR; nace_r2=", paste(cfg$music_max, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    )
  )

  list(tab = tab, leg = leg)
}

# 4.4 EU27 – persons employed per enterprise, 3×3 table (ratio)
build_eu27_emp_per_enterprise_3x3 <- function() {

  cult_raw <- get_culture_empent(cfg$geo_eu27)
  cult_row <- cult_raw %>%
    distinct(year, values) %>%
    transmute(
      row   = "Persons employed per enterprise, cultural sector (CLT)",
      year  = as.character(year),
      value = values
    ) %>%
    pivot_wider(names_from = year, values_from = value)

  music_one <- function(nace_codes, row_label) {
    raw <- get_music_empent(cfg$geo_eu27, nace_codes)

    totals <- raw %>%
      safe_group_sum(c("year","indic_sbs","freq","unit")) %>%
      select(year, indic_sbs, values) %>%
      pivot_wider(names_from = indic_sbs, values_from = values)

    tab <- totals %>%
      mutate(value = ifelse(is.na(ENT_NR) | ENT_NR == 0, NA_real_, EMP_NR / ENT_NR)) %>%
      transmute(row = row_label, year = as.character(year), value = value) %>%
      pivot_wider(names_from = year, values_from = value)

    list(tab = tab, raw = raw)
  }

  minm <- music_one(cfg$music_min, "Persons employed per enterprise, music sector (minimum)")
  maxm <- music_one(cfg$music_max, "Persons employed per enterprise, music sector (maximum)")

  tab <- bind_rows(cult_row, minm$tab, maxm$tab) %>%
    select(row, `2023`, `2022`, `2021`) %>%
    format_table_values(id_cols = "row", digits = 3)

  leg <- bind_rows(
    make_legend(
      cult_raw, "EMPL_EU27_EMP_PER_ENTERPRISE_3x3",
      "Persons employed per enterprise, cultural sector (CLT)",
      "sbs_ovw_act",
      "Direct Eurostat indicator EMP_ENT_NR (no recomputation)",
      paste0("geo=", cfg$geo_eu27, "; nace_r2=CLT; indic_sbs=EMP_ENT_NR; years=", paste(cfg$years, collapse = ","))
    ),
    make_legend(
      minm$raw, "EMPL_EU27_EMP_PER_ENTERPRISE_3x3",
      "Persons employed per enterprise, music sector (minimum)",
      "sbs_sc_ovw",
      "Computed as sum(EMP_NR) / sum(ENT_NR) over NACE codes",
      paste0("geo=", cfg$geo_eu27, "; size_emp=TOTAL; indic_sbs=EMP_NR,ENT_NR; nace_r2=", paste(cfg$music_min, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    ),
    make_legend(
      maxm$raw, "EMPL_EU27_EMP_PER_ENTERPRISE_3x3",
      "Persons employed per enterprise, music sector (maximum)",
      "sbs_sc_ovw",
      "Computed as sum(EMP_NR) / sum(ENT_NR) over NACE codes",
      paste0("geo=", cfg$geo_eu27, "; size_emp=TOTAL; indic_sbs=EMP_NR,ENT_NR; nace_r2=", paste(cfg$music_max, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    )
  )

  list(tab = tab, leg = leg)
}

# 4.5 CTY – persons employed per enterprise (culture, CLT)
build_cty_cult_emp_per_enterprise <- function() {
  raw <- get_culture_empent(cfg$countries_eu27)

  tab <- raw %>%
    distinct(geo, year, values) %>%
    transmute(country = geo, year = as.character(year), value = values) %>%
    pivot_wider(names_from = year, values_from = value) %>%
    select(country, `2023`, `2022`, `2021`) %>%
    arrange(country) %>%
    format_table_values(id_cols = "country", digits = 2)

  leg <- make_legend(
    raw, "EMPL_CTY_CULT_EMP_PER_ENTERPRISE",
    "Persons employed per enterprise, cultural sector (CLT)",
    "sbs_ovw_act",
    "Direct Eurostat indicator EMP_ENT_NR (no recomputation)",
    paste0("geo=EU27 countries; nace_r2=CLT; indic_sbs=EMP_ENT_NR; years=", paste(cfg$years, collapse = ","))
  )

  list(tab = tab, leg = leg)
}

# 4.6 CTY – persons employed per enterprise (music, computed), one table with group=min/max
build_cty_mus_emp_per_enterprise <- function() {

  get_one <- function(nace_codes, group_label) {
    raw <- get_music_empent(cfg$countries_eu27, nace_codes)

    totals <- raw %>%
      safe_group_sum(c("geo","year","indic_sbs","freq","unit")) %>%
      select(geo, year, indic_sbs, values) %>%
      pivot_wider(names_from = indic_sbs, values_from = values)

    out <- totals %>%
      mutate(value = ifelse(is.na(ENT_NR) | ENT_NR == 0, NA_real_, EMP_NR / ENT_NR)) %>%
      transmute(country = geo, group = group_label, year = as.character(year), value = value)

    list(out = out, raw = raw)
  }

  minm <- get_one(cfg$music_min, "minimum")
  maxm <- get_one(cfg$music_max, "maximum")

  tab <- bind_rows(minm$out, maxm$out) %>%
    pivot_wider(names_from = year, values_from = value) %>%
    select(country, group, `2023`, `2022`, `2021`) %>%
    arrange(country, group) %>%
    format_table_values(id_cols = c("country","group"), digits = 3)

  leg <- bind_rows(
    make_legend(
      minm$raw, "EMPL_CTY_MUS_EMP_PER_ENTERPRISE",
      "Persons employed per enterprise, music sector (minimum)",
      "sbs_sc_ovw",
      "Computed as sum(EMP_NR) / sum(ENT_NR) over NACE codes",
      paste0("geo=EU27 countries; size_emp=TOTAL; indic_sbs=EMP_NR,ENT_NR; nace_r2=", paste(cfg$music_min, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    ),
    make_legend(
      maxm$raw, "EMPL_CTY_MUS_EMP_PER_ENTERPRISE",
      "Persons employed per enterprise, music sector (maximum)",
      "sbs_sc_ovw",
      "Computed as sum(EMP_NR) / sum(ENT_NR) over NACE codes",
      paste0("geo=EU27 countries; size_emp=TOTAL; indic_sbs=EMP_NR,ENT_NR; nace_r2=", paste(cfg$music_max, collapse = ","), "; years=", paste(cfg$years, collapse = ","))
    )
  )

  list(tab = tab, leg = leg)
}

# -----------------------------
# 5) RUN PIPELINE + COLLECT OUTPUTS
# -----------------------------
tables  <- list()
legends <- list()

if (isTRUE(RUN$eu27_total_employment)) {
  res <- build_eu27_total_employment_3x3()
  tables[["EMPL_EU27_TOTAL_EMPLOYMENT_3x3"]]          <- res$tab
  legends[["EMPL_LEGEND_EU27_TOTAL_EMPLOYMENT_3x3"]]  <- res$leg
}

if (isTRUE(RUN$cty_total_employment)) {
  cult <- build_cty_cult_total_employment()
  mus  <- build_cty_mus_total_employment()

  tables[["EMPL_CTY_CULT_TOTAL_EMPLOYMENT"]]          <- cult$tab
  legends[["EMPL_LEGEND_CTY_CULT_TOTAL_EMPLOYMENT"]]  <- cult$leg

  tables[["EMPL_CTY_MUS_TOTAL_EMPLOYMENT"]]           <- mus$tab
  legends[["EMPL_LEGEND_CTY_MUS_TOTAL_EMPLOYMENT"]]   <- mus$leg
}

if (isTRUE(RUN$eu27_emp_per_enterprise)) {
  res <- build_eu27_emp_per_enterprise_3x3()
  tables[["EMPL_EU27_EMP_PER_ENTERPRISE_3x3"]]         <- res$tab
  legends[["EMPL_LEGEND_EU27_EMP_PER_ENTERPRISE_3x3"]] <- res$leg
}

if (isTRUE(RUN$cty_emp_per_enterprise)) {
  cult <- build_cty_cult_emp_per_enterprise()
  mus  <- build_cty_mus_emp_per_enterprise()

  tables[["EMPL_CTY_CULT_EMP_PER_ENTERPRISE"]]         <- cult$tab
  legends[["EMPL_LEGEND_CTY_CULT_EMP_PER_ENTERPRISE"]] <- cult$leg

  tables[["EMPL_CTY_MUS_EMP_PER_ENTERPRISE"]]          <- mus$tab
  legends[["EMPL_LEGEND_CTY_MUS_EMP_PER_ENTERPRISE"]]  <- mus$leg
}

# -----------------------------
# 6) EXPORTS
# -----------------------------
if (isTRUE(RUN$write_csv)) {
  for (nm in names(tables))  export_csv(tables[[nm]],  paste0(nm, ".csv"))
  for (nm in names(legends)) export_csv(legends[[nm]], paste0(nm, ".csv"))
}

if (isTRUE(RUN$write_excel)) {
  wb <- openxlsx::createWorkbook()

  # Tables
  t_names  <- names(tables)
  t_sheets <- make_unique_sheet_names(substr(paste0("T_", t_names), 1, 31))
  for (i in seq_along(t_names)) {
    openxlsx::addWorksheet(wb, t_sheets[i])
    openxlsx::writeData(wb, t_sheets[i], tables[[t_names[i]]])
  }

  # Legends
  l_names  <- names(legends)
  l_sheets <- make_unique_sheet_names(substr(paste0("L_", l_names), 1, 31))
  for (i in seq_along(l_names)) {
    openxlsx::addWorksheet(wb, l_sheets[i])
    openxlsx::writeData(wb, l_sheets[i], legends[[l_names[i]]])
  }

  # Provenance
  prov <- tibble::tibble(
    item  = c(
      "generated_at",
      "outdir",
      "years (column order)",
      "EU27 aggregate code",
      "EU27 country list (codes)",
      "music_min NACE (codes)",
      "music_max NACE (codes)",
      "numeric formatting",
      "eurostat package version",
      "dplyr package version",
      "tidyr package version",
      "readr package version",
      "openxlsx package version",
      "RUN settings",
      "tables exported (names)",
      "legends exported (names)"
    ),
    value = c(
      as.character(Sys.time()),
      cfg$outdir,
      paste(cfg$years, collapse = ", "),
      cfg$geo_eu27,
      paste(cfg$countries_eu27, collapse = ", "),
      paste(cfg$music_min, collapse = ", "),
      paste(cfg$music_max, collapse = ", "),
      "Employment levels: rounded to 3 decimals (thousand persons). Ratios: rounded to 2–3 decimals (see builders).",
      as.character(utils::packageVersion("eurostat")),
      as.character(utils::packageVersion("dplyr")),
      as.character(utils::packageVersion("tidyr")),
      as.character(utils::packageVersion("readr")),
      as.character(utils::packageVersion("openxlsx")),
      paste0(names(RUN), "=", unlist(RUN), collapse = "; "),
      paste(names(tables), collapse = ", "),
      paste(names(legends), collapse = ", ")
    )
  )

  datasets_used <- tibble::tibble(
    dataset_id = c("cult_emp_sex", "sbs_sc_ovw", "sbs_ovw_act"),
    purpose = c(
      "Cultural employment (levels) – unit THS_PER; sex=T; EU27 and country tables.",
      "Music employment (levels) and computed ratios from EMP_NR and ENT_NR aggregated over NACE codes.",
      "Cultural CLT persons employed per enterprise (EMP_ENT_NR; nace_r2=CLT)."
    )
  )

  openxlsx::addWorksheet(wb, "Provenance")
  openxlsx::writeData(wb, "Provenance", prov, startRow = 1, startCol = 1)
  openxlsx::writeData(wb, "Provenance", datasets_used, startRow = nrow(prov) + 3, startCol = 1)
  openxlsx::setColWidths(wb, "Provenance", cols = 1:2, widths = c(34, 110))

  xlsx_path <- file.path(cfg$outdir, "OpenMusE_Employment_Tables_and_Legends.xlsx")
  openxlsx::saveWorkbook(wb, xlsx_path, overwrite = TRUE)

  cat("Excel workbook written to:\n", xlsx_path, "\n", sep = "")
}

cat("Employment module completed. Output folder:\n", cfg$outdir, "\n", sep = "")

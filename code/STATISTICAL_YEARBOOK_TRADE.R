# ============================================================
# OpenMusE Statistical Yearbook – Cultural Trade (TRADE) Module
#
# Tables produced (EU27 + by country):
#  - Total intra-EU trade in cultural goods by product (thousand euro)
#  - Total intra-EU trade in cultural goods by product (% of total)
#  - Total extra-EU trade in cultural goods by product (thousand euro)
#  - Total extra-EU trade in cultural goods by product (% of total)
#
# Products covered:
#  - TOTAL (Eurostat product aggregate)
#  - MUSI  (Musical instruments)
#  - RECMED_FILMVG_XVC (Recorded media, films and video games)
#
# Dataset: Eurostat cult_trd_prd
# Dimensions: freq, geo, partner, prod_ct, stk_flow, unit, time, values
#
# Definition:
#  - Total trade = imports + exports (sum over stk_flow)
#  - Intra-EU: partner = INT_EU27_2020
#  - Extra-EU: partner = EXT_EU27_2020
#
# Years shown: 2021–2023 (column order: 2023, 2022, 2021)
# Unit used: THS_EUR (thousand euro)
#
# Naming convention:
#  - Tables:  TRADE_<EU27|CTY>_<INTRA|EXTRA>_<LEVEL|SHARE>
#  - Legends: TRADE_LEGEND_<EU27|CTY>_<INTRA|EXTRA>_<LEVEL|SHARE>
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
  years  = c(2023, 2022, 2021),  # required column order

  countries_eu27 = c(
    "AT","BE","BG","CY","CZ","DE","DK","EE","EL","ES","FI","FR","HR","HU",
    "IE","IT","LT","LU","LV","MT","NL","PL","PT","RO","SE","SI","SK"
  ),

  # Requested product rows (prod_ct codes)
  products_focus = c("TOTAL", "MUSI", "RECMED_FILMVG_XVC"),

  unit_target    = "THS_EUR",
  dataset_id     = "cult_trd_prd",
  partner_intra  = "INT_EU27_2020",
  partner_extra  = "EXT_EU27_2020"
)

if (!dir.exists(cfg$outdir)) dir.create(cfg$outdir, recursive = TRUE, showWarnings = FALSE)
stopifnot(dir.exists(cfg$outdir))

if (isTRUE(RUN$write_excel)) {
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Package 'openxlsx' is required for Excel export. Install with: install.packages('openxlsx')")
  }
}

# ------------------------------------------------------------
# 2) Helpers
# ------------------------------------------------------------
add_year <- function(df, time_col = "time") {
  df %>%
    mutate(year = suppressWarnings(as.integer(substr(as.character(.data[[time_col]]), 1, 4))))
}

keep_years <- function(df) df %>% filter(year %in% cfg$years)

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

# Safe pull: return empty tibble instead of error
safe_get <- function(filters) {
  tryCatch(
    eurostat::get_eurostat(cfg$dataset_id, filters = filters, time_format = "num", cache = TRUE),
    error = function(e) tibble::tibble()
  )
}

# ------------------------------------------------------------
# 3) Core pull (partner scope)
# ------------------------------------------------------------
get_trade_scope <- function(geo_vec, scope = c("intra", "extra")) {
  scope <- match.arg(scope)
  partner_val <- if (scope == "intra") cfg$partner_intra else cfg$partner_extra

  df <- safe_get(list(
    geo     = geo_vec,
    partner = partner_val,
    unit    = cfg$unit_target
  ))

  if (nrow(df) == 0) {
    stop(
      "No data returned for scope=", scope,
      ". Check geo codes, partner codes, unit, and years availability.\n",
      "Using partner=", partner_val, ", unit=", cfg$unit_target
    )
  }

  df %>%
    add_year() %>%
    keep_years()
}

# Total trade = imports + exports (sum over stk_flow)
sum_over_flows <- function(df) {
  stopifnot(all(c("geo","prod_ct","year","unit","values") %in% names(df)))
  df %>%
    group_by(geo, prod_ct, year, unit) %>%
    summarise(values = sum(values, na.rm = TRUE), .groups = "drop")
}

# EU totals = sum of EU27 countries (dataset often has no EU aggregate geo)
sum_countries_to_eu <- function(df_sum_cty) {
  df_sum_cty %>%
    group_by(prod_ct, year) %>%
    summarise(values = sum(values, na.rm = TRUE), .groups = "drop")
}

# ------------------------------------------------------------
# 4) Builders – labels and tables
# ------------------------------------------------------------
label_row <- function(prod_ct) {
  dplyr::case_when(
    prod_ct == "TOTAL"             ~ "Total (TOTAL) (thousand euro)",
    prod_ct == "MUSI"              ~ "Musical instruments; parts and accessories thereof (MUSI) (thousand euro)",
    prod_ct == "RECMED_FILMVG_XVC" ~ "Music in manuscript, gramophone records, recorded magnetic tapes and optical media (CDs); audio-visual and interactive media (RECMED_FILMVG_XVC) (thousand euro)",
    TRUE                           ~ paste0(prod_ct, " (thousand euro)")
  )
}

# EU: keep ONLY the official product categories requested (incl. TOTAL)
build_level_table_eu <- function(df_eu_sum, years = cfg$years) {
  df_eu_sum %>%
    filter(prod_ct %in% cfg$products_focus) %>%
    group_by(prod_ct, year) %>%
    summarise(value = sum(values, na.rm = TRUE), .groups = "drop") %>%
    mutate(
      row  = label_row(prod_ct),
      year = as.character(year)
    ) %>%
    select(row, year, value) %>%
    pivot_wider(names_from = year, values_from = value) %>%
    select(row, all_of(as.character(years)))
}

# Shares: 100 * product / TOTAL, by year
build_share_table <- function(level_tab) {
  yrs <- setdiff(names(level_tab), "row")
  out <- level_tab

  denom <- out %>% filter(grepl("^Total", row)) %>% select(all_of(yrs))
  if (nrow(denom) != 1) stop("Expected exactly one TOTAL row for share calculation.")

  for (y in yrs) out[[y]] <- 100 * out[[y]] / as.numeric(denom[[y]])

  out %>%
    mutate(row = sub("\\(thousand euro\\)", "(% of total)", row))
}

# Country: keep ONLY the official product categories requested (incl. TOTAL)
build_country_level_table <- function(df_sum_cty, years = cfg$years) {
  df_sum_cty %>%
    filter(prod_ct %in% cfg$products_focus) %>%
    group_by(geo, prod_ct, year) %>%
    summarise(value = sum(values, na.rm = TRUE), .groups = "drop") %>%
    mutate(
      row  = label_row(prod_ct),
      year = as.character(year)
    ) %>%
    select(country = geo, row, year, value) %>%
    pivot_wider(names_from = year, values_from = value) %>%
    select(country, row, all_of(as.character(years))) %>%
    arrange(country, row)
}

# Shares by country: 100 * product / TOTAL within each country-year
build_country_share_table <- function(level_cty_tab) {
  yrs <- setdiff(names(level_cty_tab), c("country","row"))

  level_cty_tab %>%
    group_by(country) %>%
    group_modify(~{
      denom <- .x %>% filter(grepl("^Total", row)) %>% select(all_of(yrs))
      if (nrow(denom) != 1) stop("Expected exactly one TOTAL row per country for share calculation.")
      for (y in yrs) .x[[y]] <- 100 * .x[[y]] / as.numeric(denom[[y]])
      .x %>% mutate(row = sub("\\(thousand euro\\)", "(% of total)", row))
    }) %>%
    ungroup()
}

make_legend <- function(df_raw, table_name, scope_label, metric_label, details) {
  tibble::tibble(
    table   = table_name,
    scope   = scope_label,
    metric  = metric_label,
    dataset = cfg$dataset_id,
    unit    = cfg$unit_target,
    years   = paste(cfg$years, collapse = ", "),
    details = details
  )
}

# ------------------------------------------------------------
# 5) Pipeline
# ------------------------------------------------------------
tables  <- list()
legends <- list()

if (isTRUE(RUN$eu_tables)) {

  # -------- Intra-EU (EU27 total) --------
  raw_intra_cty <- get_trade_scope(cfg$countries_eu27, scope = "intra")
  sum_intra_cty <- sum_over_flows(raw_intra_cty)
  sum_intra_eu  <- sum_countries_to_eu(sum_intra_cty)

  eu_intra_level <- build_level_table_eu(sum_intra_eu)
  eu_intra_share <- build_share_table(eu_intra_level)

  tables[["TRADE_EU27_INTRA_LEVEL_THS_EUR"]] <- format_table_values(eu_intra_level, "row", digits = 2)
  tables[["TRADE_EU27_INTRA_SHARE_PCT"]]     <- format_table_values(eu_intra_share, "row", digits = 2)

  legends[["TRADE_LEGEND_EU27_INTRA_LEVEL_THS_EUR"]] <- make_legend(
    raw_intra_cty, "TRADE_EU27_INTRA_LEVEL_THS_EUR", "Intra-EU", "Level (thousand euro)",
    paste0(
      "Computed from Eurostat cult_trd_prd. Partner=", cfg$partner_intra, ". ",
      "Total trade is imports + exports (sum over stk_flow). ",
      "EU27 totals are computed as the sum of EU27 Member States. ",
      "Rows correspond to official prod_ct categories (TOTAL, MUSI, RECMED_FILMVG_XVC)."
    )
  )

  legends[["TRADE_LEGEND_EU27_INTRA_SHARE_PCT"]] <- make_legend(
    raw_intra_cty, "TRADE_EU27_INTRA_SHARE_PCT", "Intra-EU", "Share (% of total)",
    paste0(
      "Shares computed within each year as 100 * product / TOTAL using intra-EU totals. ",
      "TOTAL is the official prod_ct category (not computed as a sum). ",
      "Total trade is imports + exports (sum over stk_flow). EU27 totals are sum of Member States."
    )
  )

  # -------- Extra-EU (EU27 total) --------
  raw_extra_cty <- get_trade_scope(cfg$countries_eu27, scope = "extra")
  sum_extra_cty <- sum_over_flows(raw_extra_cty)
  sum_extra_eu  <- sum_countries_to_eu(sum_extra_cty)

  eu_extra_level <- build_level_table_eu(sum_extra_eu)
  eu_extra_share <- build_share_table(eu_extra_level)

  tables[["TRADE_EU27_EXTRA_LEVEL_THS_EUR"]] <- format_table_values(eu_extra_level, "row", digits = 2)
  tables[["TRADE_EU27_EXTRA_SHARE_PCT"]]     <- format_table_values(eu_extra_share, "row", digits = 2)

  legends[["TRADE_LEGEND_EU27_EXTRA_LEVEL_THS_EUR"]] <- make_legend(
    raw_extra_cty, "TRADE_EU27_EXTRA_LEVEL_THS_EUR", "Extra-EU", "Level (thousand euro)",
    paste0(
      "Computed from Eurostat cult_trd_prd. Partner=", cfg$partner_extra, ". ",
      "Total trade is imports + exports (sum over stk_flow). ",
      "EU27 totals are computed as the sum of EU27 Member States. ",
      "Rows correspond to official prod_ct categories (TOTAL, MUSI, RECMED_FILMVG_XVC)."
    )
  )

  legends[["TRADE_LEGEND_EU27_EXTRA_SHARE_PCT"]] <- make_legend(
    raw_extra_cty, "TRADE_EU27_EXTRA_SHARE_PCT", "Extra-EU", "Share (% of total)",
    paste0(
      "Shares computed within each year as 100 * product / TOTAL using extra-EU totals. ",
      "TOTAL is the official prod_ct category (not computed as a sum). ",
      "Total trade is imports + exports (sum over stk_flow). EU27 totals are sum of Member States."
    )
  )
}

if (isTRUE(RUN$country_tables)) {

  # -------- Intra-EU (country) --------
  raw_intra_cty <- get_trade_scope(cfg$countries_eu27, scope = "intra")
  sum_intra_cty <- sum_over_flows(raw_intra_cty)

  cty_intra_level <- build_country_level_table(sum_intra_cty)
  cty_intra_share <- build_country_share_table(cty_intra_level)

  tables[["TRADE_CTY_INTRA_LEVEL_THS_EUR"]] <- format_table_values(cty_intra_level, c("country","row"), digits = 2)
  tables[["TRADE_CTY_INTRA_SHARE_PCT"]]     <- format_table_values(cty_intra_share, c("country","row"), digits = 2)

  legends[["TRADE_LEGEND_CTY_INTRA_LEVEL_THS_EUR"]] <- make_legend(
    raw_intra_cty, "TRADE_CTY_INTRA_LEVEL_THS_EUR", "Intra-EU", "Level (thousand euro)",
    paste0(
      "Country-level totals from Eurostat cult_trd_prd. Partner=", cfg$partner_intra, ". ",
      "Total trade is imports + exports (sum over stk_flow). ",
      "Rows correspond to official prod_ct categories (TOTAL, MUSI, RECMED_FILMVG_XVC)."
    )
  )

  legends[["TRADE_LEGEND_CTY_INTRA_SHARE_PCT"]] <- make_legend(
    raw_intra_cty, "TRADE_CTY_INTRA_SHARE_PCT", "Intra-EU", "Share (% of total)",
    paste0(
      "Country-level shares computed within each year as 100 * product / TOTAL (intra-EU). ",
      "TOTAL is the official prod_ct category (not computed as a sum). ",
      "Total trade is imports + exports (sum over stk_flow)."
    )
  )

  # -------- Extra-EU (country) --------
  raw_extra_cty <- get_trade_scope(cfg$countries_eu27, scope = "extra")
  sum_extra_cty <- sum_over_flows(raw_extra_cty)

  cty_extra_level <- build_country_level_table(sum_extra_cty)
  cty_extra_share <- build_country_share_table(cty_extra_level)

  tables[["TRADE_CTY_EXTRA_LEVEL_THS_EUR"]] <- format_table_values(cty_extra_level, c("country","row"), digits = 2)
  tables[["TRADE_CTY_EXTRA_SHARE_PCT"]]     <- format_table_values(cty_extra_share, c("country","row"), digits = 2)

  legends[["TRADE_LEGEND_CTY_EXTRA_LEVEL_THS_EUR"]] <- make_legend(
    raw_extra_cty, "TRADE_CTY_EXTRA_LEVEL_THS_EUR", "Extra-EU", "Level (thousand euro)",
    paste0(
      "Country-level totals from Eurostat cult_trd_prd. Partner=", cfg$partner_extra, ". ",
      "Total trade is imports + exports (sum over stk_flow). ",
      "Rows correspond to official prod_ct categories (TOTAL, MUSI, RECMED_FILMVG_XVC)."
    )
  )

  legends[["TRADE_LEGEND_CTY_EXTRA_SHARE_PCT"]] <- make_legend(
    raw_extra_cty, "TRADE_CTY_EXTRA_SHARE_PCT", "Extra-EU", "Share (% of total)",
    paste0(
      "Country-level shares computed within each year as 100 * product / TOTAL (extra-EU). ",
      "TOTAL is the official prod_ct category (not computed as a sum). ",
      "Total trade is imports + exports (sum over stk_flow)."
    )
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

  xlsx_path <- file.path(cfg$outdir, "OpenMusE_Cultural_Trade.xlsx")
  openxlsx::saveWorkbook(wb, xlsx_path, overwrite = TRUE)

  cat("Excel workbook written to:\n", xlsx_path, "\n", sep = "")
}

cat("TRADE module completed. Output folder:\n", cfg$outdir, "\n", sep = "")

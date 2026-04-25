#   fdi_NI.xls
#   API_FP_CPI_TOTL_ZG_DS2_en_excel_v2_207080.xls    (Inflation)
#   API_NY_GDP_MKTP_KD_ZG_DS2_en_excel_v2_29.xls    (GDP growth)
#   API_NE_TRD_GNFS_ZS_DS2_en_excel_v2_37.xls       (Trade openness)
#   ert_bil_eur_m__custom_20843966_spreadsheet.xlsx  (Bilateral EUR)
#   ert_eff_ic_m__custom_20843983_page_spreadsheet.xlsx (NEER)
# ================================================================

# install.packages(c("tidyverse","readxl","plm","lmtest","sandwich"))

library(tidyverse)
library(readxl)
library(plm)
library(lmtest)
library(sandwich)
library(ggplot2)

# ================================================================
# SETTINGS
# ================================================================

YEAR_START <- 2000
YEAR_END   <- 2022

eu27 <- c("AUT","BEL","BGR","HRV","CYP","CZE","DNK","EST","FIN",
          "FRA","DEU","GRC","HUN","IRL","ITA","LVA","LTU","LUX",
          "MLT","NLD","POL","PRT","ROU","SVK","SVN","ESP","SWE")

eurozone <- c("AUT","BEL","CYP","EST","FIN","FRA","DEU","GRC",
              "HRV","IRL","ITA","LVA","LTU","LUX","MLT","NLD",
              "PRT","SVK","SVN","ESP")

non_euro <- c("CZE","DNK","HUN","POL","ROU","SWE", "BGR")

theme_set(theme_minimal(base_size = 13))
col_euro    <- "#1D9E75"
col_noneuro <- "#BA7517"


# ================================================================
# HELPER: parse Eurostat wide-format export
# ================================================================
# Both Eurostat files share the same layout:
#   Rows 1-8  : metadata
#   Row 9     : dates in columns 2, 4, 6, ... (odd positions, 1-based)
#   Row 10    : label header
#   Row 11+   : data — col 1 = name, same odd columns = values

parse_eurostat_wide <- function(path, sheet = "Sheet 1") {

  raw <- read_excel(path, sheet = sheet,
                    col_names = FALSE, col_types = "text")

  # Dates are in odd-indexed columns of row 9 (YYYY-MM format)
  date_row <- as.character(unlist(raw[9, ]))
  date_idx <- which(grepl("^\\d{4}-\\d{2}$", date_row))
  dates    <- date_row[date_idx]

  # Data starts at row 11; col 1 = name, value cols = same as date_idx
  data_rows <- raw |>
    slice(-(1:10)) |>
    filter(!is.na(...1))

  name_col  <- data_rows[[1]]
  value_mat <- data_rows[, date_idx]
  colnames(value_mat) <- dates

  bind_cols(name = name_col, value_mat) |>
    filter(!name %in% c("Special value", ":", "")) |>
    pivot_longer(cols      = -name,
                 names_to  = "month",
                 values_to = "value") |>
    mutate(month = as.Date(paste0(month, "-01")),
           year  = as.integer(format(month, "%Y")),
           value = suppressWarnings(as.numeric(value))) |>
    filter(!is.na(value),
           year >= YEAR_START,
           year <= YEAR_END)
}


# ================================================================
# PART 1: FDI DATA (World Bank)
# ================================================================

fdi_raw <- read_excel("fdi_NI.xls", sheet = "Data", skip = 3)
names(fdi_raw)[1:4] <- c("country_name","country_code",
                          "indicator_name","indicator_code")

fdi <- fdi_raw |>
  filter(country_code %in% eu27) |>
  select(country_name, country_code,
         as.character(YEAR_START:YEAR_END)) |>
  pivot_longer(cols      = -c(country_name, country_code),
               names_to  = "year",
               values_to = "fdi") |>
  mutate(year = as.integer(year),
         fdi  = as.numeric(fdi))

cat("FDI:", n_distinct(fdi$country_code), "countries |",
    n_distinct(fdi$year), "years |",
    sum(!is.na(fdi$fdi)), "non-missing obs\n")


# ================================================================
# PART 2: CONTROL VARIABLES (World Bank)
# ================================================================

load_wb <- function(path, var_name) {
  df <- read_excel(path, sheet = "Data", skip = 3)
  names(df)[1:4] <- c("country_name","country_code",
                       "indicator_name","indicator_code")
  df |>
    filter(country_code %in% eu27) |>
    select(country_code, as.character(YEAR_START:YEAR_END)) |>
    pivot_longer(cols      = -country_code,
                 names_to  = "year",
                 values_to = var_name) |>
    mutate(year       = as.integer(year),
           !!var_name := as.numeric(.data[[var_name]]))
}

inflation  <- load_wb(
  "Inf_CP.xls", "inflation")
gdp_growth <- load_wb(
  "GDP_Growth.xls",  "gdp_growth")
trade_open <- load_wb(
  "Trade_GDP.xls",     "trade_open")

cat("Controls loaded: inflation, gdp_growth, trade_open\n")


# ================================================================
# PART 3: BILATERAL EUR EXCHANGE RATES (Eurostat)
# ================================================================
# Sheet 1 = monthly average (national currency per EUR)
# We keep the 6 non-euro EU currencies only

# Read file as a plain data.frame — this fixes all column indexing issues
raw_bil <- as.data.frame(
  read_excel("Bilat_EUR.xlsx",
             sheet = "Sheet 1", col_names = FALSE, skip = 8)
)
# After skip=8:
#   Row 1 = dates row
#   Row 2 = "CURRENCY (Labels)" header
#   Row 3 = Czech koruna
#   Row 4 = Danish krone
#   Row 5 = Hungarian forint
#   Row 6 = Polish zloty
#   Row 7 = Romanian leu
#   Row 8 = Swedish krona

# Find which columns contain dates
dates_row <- as.character(raw_bil[1, ])
keep_cols <- which(grepl("^\\d{4}-\\d{2}$", dates_row))
dates     <- dates_row[keep_cols]

# Extract the 6 non-euro EU currencies (rows 3 to 8)
er <- raw_bil[3:8, c(1, keep_cols)]
colnames(er) <- c("currency", dates)
er$country_code <- c("CZE", "DNK", "HUN", "POL", "ROU", "SWE")

# Pivot to long format
er_long <- er |>
  select(-currency) |>
  pivot_longer(-country_code,
               names_to  = "month",
               values_to = "eur_rate") |>
  mutate(month    = as.Date(paste0(month, "-01")),
         year     = as.integer(format(month, "%Y")),
         eur_rate = suppressWarnings(as.numeric(eur_rate))) |>
  filter(!is.na(eur_rate), year >= YEAR_START, year <= YEAR_END)

cat("Bilateral EUR loaded:", nrow(er_long), "rows\n")
print(er_long |> filter(country_code == "CZE") |> head(6))


# ================================================================
# PART 4: COMPUTE ANNUAL EXCHANGE RATE VOLATILITY
# ================================================================
# Annual ERV = SD of monthly % changes in the bilateral EUR rate
# Follows Morina & Mera (2024) and Hanusch et al. (2018)


er_vol <- er_long |>
  group_by(country_code, year) |>
  summarise(er_vol = sd(100 * diff(log(eur_rate)), na.rm = TRUE),
            .groups = "drop")
cat("ERV computed:", nrow(er_vol), "country-year rows\n")
print(er_vol |> head(9))
er_vol |>
  filter(country_code %in% non_euro) |>
  summarise(mean = round(mean(er_vol, na.rm = TRUE), 4),
            sd   = round(sd(er_vol,   na.rm = TRUE), 4),
            min  = round(min(er_vol,  na.rm = TRUE), 4),
            max  = round(max(er_vol,  na.rm = TRUE), 4)) |>
  print()

# ================================================================
# PART 5: NEER (Eurostat) — all EU-27
# ================================================================
# Nominal effective exchange rate vs 19 trading partners
# Index 2015 = 100. Covers all 27 EU members.

raw_neer <- as.data.frame(
  read_excel("NEER.xlsx",
             sheet = "Sheet 1", col_names = FALSE, skip = 8)
)

dates_neer <- as.character(raw_neer[1, ])
keep_neer  <- which(grepl("^\\d{4}-\\d{2}$", dates_neer))

neer_raw <- raw_neer[, c(1, keep_neer)]
colnames(neer_raw) <- c("country", dates_neer[keep_neer])

neer_country_map <- c(
  "Belgium"     = "BEL", "Bulgaria"    = "BGR",
  "Czechia"     = "CZE", "Denmark"     = "DNK",
  "Germany"     = "DEU", "Estonia"     = "EST",
  "Ireland"     = "IRL", "Greece"      = "GRC",
  "Spain"       = "ESP", "France"      = "FRA",
  "Croatia"     = "HRV", "Italy"       = "ITA",
  "Cyprus"      = "CYP", "Latvia"      = "LVA",
  "Lithuania"   = "LTU", "Luxembourg"  = "LUX",
  "Hungary"     = "HUN", "Malta"       = "MLT",
  "Netherlands" = "NLD", "Austria"     = "AUT",
  "Poland"      = "POL", "Portugal"    = "PRT",
  "Romania"     = "ROU", "Slovenia"    = "SVN",
  "Slovakia"    = "SVK", "Finland"     = "FIN",
  "Sweden"      = "SWE"
)

neer <- neer_raw |>
  filter(country %in% names(neer_country_map)) |>
  mutate(country_code = neer_country_map[country]) |>
  select(country_code, all_of(dates_neer[keep_neer])) |>
  pivot_longer(-country_code,
               names_to  = "month",
               values_to = "neer") |>
  mutate(month = as.Date(paste0(month, "-01")),
         year  = as.integer(format(month, "%Y")),
         neer  = suppressWarnings(as.numeric(neer))) |>
  filter(!is.na(neer), year >= YEAR_START, year <= YEAR_END) |>
  group_by(country_code, year) |>
  summarise(neer_avg = mean(neer, na.rm = TRUE), .groups = "drop")

cat("NEER:", n_distinct(neer$country_code), "countries |",
    n_distinct(neer$year), "years\n")
print(neer |> head(6))

# ================================================================
# PART 6: FULL PANEL
# ================================================================

panel <- fdi |>
  mutate(
    euro_member = ifelse(country_code %in% eurozone, 1, 0),
    group       = ifelse(country_code %in% eurozone,
                         "Eurozone", "Non-euro EU"),
    period      = case_when(
      year <= 2007 ~ "Pre-crisis (2000-2007)",
      year <= 2012 ~ "Crisis (2008-2012)",
      TRUE         ~ "Post-crisis (2013-2022)"
    ),
    period = factor(period,
                    levels = c("Pre-crisis (2000-2007)",
                               "Crisis (2008-2012)",
                               "Post-crisis (2013-2022)"))
  ) |>
  left_join(gdp_growth, by = c("country_code","year")) |>
  left_join(inflation,  by = c("country_code","year")) |>
  left_join(trade_open, by = c("country_code","year")) |>
  left_join(er_vol,     by = c("country_code","year")) |>
  left_join(neer,       by = c("country_code","year"))

cat("\nPanel:", nrow(panel), "rows |",
    n_distinct(panel$country_code), "countries\n")
cat("er_vol present for:", sum(!is.na(panel$er_vol)), "rows",
    "(non-euro members only)\n")


# ================================================================
# PART 7: DESCRIPTIVE STATISTICS
# ================================================================

cat("\n--- FDI summary by group ---\n")
panel |>
  group_by(group) |>
  summarise(
    n      = sum(!is.na(fdi)),
    mean   = round(mean(fdi,   na.rm = TRUE), 2),
    median = round(median(fdi, na.rm = TRUE), 2),
    sd     = round(sd(fdi,     na.rm = TRUE), 2),
    min    = round(min(fdi,    na.rm = TRUE), 2),
    max    = round(max(fdi,    na.rm = TRUE), 2),
    .groups = "drop"
  ) |> print()

cat("\n--- FDI mean by period and group ---\n")
panel |>
  group_by(period, group) |>
  summarise(mean_fdi = round(mean(fdi, na.rm = TRUE), 2),
            sd_fdi   = round(sd(fdi,   na.rm = TRUE), 2),
            .groups  = "drop") |>
  print()

cat("\n--- ERV summary (non-euro members) ---\n")
panel |>
  filter(!is.na(er_vol)) |>
  group_by(country_code) |>
  summarise(mean = round(mean(er_vol, na.rm = TRUE), 4),
            sd   = round(sd(er_vol,   na.rm = TRUE), 4),
            min  = round(min(er_vol,  na.rm = TRUE), 4),
            max  = round(max(er_vol,  na.rm = TRUE), 4),
            .groups = "drop") |>
  print()

cat("\n--- All variables summary ---\n")
panel |>
  select(fdi, er_vol, gdp_growth, inflation, trade_open, neer_avg) |>
  summary() |> print()

cat("\n--- Correlation matrix (non-euro members only) ---\n")
panel |>
  filter(!is.na(er_vol)) |>
  select(fdi, er_vol, gdp_growth, inflation, trade_open) |>
  cor(use = "complete.obs") |>
  round(3) |> print()


# ================================================================
# PART 8: VISUALISATIONS
# ================================================================

# --- Plot 1: FDI trend by group ---------------------------------
p1 <- panel |>
  group_by(year, group) |>
  summarise(mean_fdi = mean(fdi, na.rm = TRUE), .groups = "drop") |>
  ggplot(aes(x = year, y = mean_fdi, colour = group)) +
  geom_line(linewidth = 1) + geom_point(size = 2) +
  geom_vline(xintercept = 2008, linetype = "dashed",
             colour = "grey50") +
  annotate("text", x = 2008.5, y = 10,
           label = "GFC", colour = "grey40", size = 3.5) +
  scale_colour_manual(
    values = c("Eurozone" = col_euro, "Non-euro EU" = col_noneuro),
    name = NULL) +
  scale_x_continuous(breaks = seq(2000, 2022, 4)) +
  labs(title    = "FDI net inflows: eurozone vs non-euro EU (2000-2022)",
       subtitle = "Mean across countries (% of GDP)",
       x = NULL, y = "FDI (% of GDP)") +
  theme(legend.position = "bottom")
print(p1)
ggsave("plot1_fdi_trend.png", p1, width = 10, height = 5.5, dpi = 150)


# --- Plot 2: Average FDI per country ----------------------------
p2 <- panel |>
  group_by(country_code, group) |>
  summarise(mean_fdi = mean(fdi, na.rm = TRUE), .groups = "drop") |>
  ggplot(aes(x = reorder(country_code, mean_fdi),
             y = mean_fdi, fill = group)) +
  geom_col() +
  scale_fill_manual(
    values = c("Eurozone" = col_euro, "Non-euro EU" = col_noneuro),
    name = NULL) +
  coord_flip() +
  labs(title = "Average FDI by EU country (2000-2022)",
       x = NULL, y = "Mean FDI (% of GDP)") +
  theme(legend.position = "bottom")
print(p2)
ggsave("plot2_fdi_country.png", p2, width = 9, height = 7, dpi = 150)

# --- Plot 3: FDI over time — non-euro members -------------------
p3 <- panel |>
  filter(country_code %in% non_euro) |>
  ggplot(aes(x = year, y = fdi)) +
  geom_line(colour = col_noneuro, linewidth = 0.9) +
  geom_hline(yintercept = 0, linetype = "dashed", colour = "grey60") +
  facet_wrap(~country_name, scales = "free_y", ncol = 3) +
  scale_x_continuous(breaks = c(2000, 2010, 2020)) +
  labs(title = "FDI — non-eurozone EU members (2000-2022)",
       x = NULL, y = "FDI (% of GDP)") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))
print(p3)
ggsave("plot3_noneuro_fdi.png", p3, width = 10, height = 6, dpi = 150)

# --- Plot 4: FDI boxplot by period and group --------------------
p4 <- panel |>
  ggplot(aes(x = period, y = fdi, fill = group)) +
  geom_boxplot(outlier.alpha = 0.3, width = 0.55) +
  scale_fill_manual(
    values = c("Eurozone" = col_euro, "Non-euro EU" = col_noneuro),
    name = NULL) +
  coord_cartesian(ylim = c(-15, 30)) +
  labs(title = "FDI distribution by period and group",
       x = NULL, y = "FDI (% of GDP)") +
  theme(legend.position = "bottom")
print(p4)
ggsave("plot4_boxplot.png", p4, width = 10, height = 5.5, dpi = 150)

# --- Plot 5: ERV over time — non-euro members -------------------
p5 <- panel |>
  filter(!is.na(er_vol)) |>
  ggplot(aes(x = year, y = er_vol, colour = country_name)) +
  geom_line(linewidth = 0.9) + geom_point(size = 1.5) +
  geom_vline(xintercept = 2008, linetype = "dashed",
             colour = "grey50") +
  scale_x_continuous(breaks = seq(2000, 2022, 4)) +
  labs(title    = "Annual exchange rate volatility vs EUR (2000-2022)",
       subtitle = "SD of monthly % changes — non-euro EU members",
       x = NULL, y = "ERV (SD of monthly % changes)",
       colour = NULL) +
  theme(legend.position = "bottom")
print(p5)
ggsave("plot5_erv_trend.png", p5, width = 10, height = 5.5, dpi = 150)

# --- Plot 6: ERV vs FDI scatter (non-euro members) --------------
p6 <- panel |>
  filter(!is.na(er_vol), !is.na(fdi)) |>
  ggplot(aes(x = er_vol, y = fdi, colour = country_name)) +
  geom_point(size = 2, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE,
              colour = "grey30", linewidth = 0.8) +
  labs(title    = "Exchange rate volatility vs FDI (non-euro EU members)",
       subtitle = "Each point = one country-year observation",
       x = "ERV (SD of monthly % changes)",
       y = "FDI (% of GDP)", colour = NULL) +
  theme(legend.position = "bottom")
print(p6)
ggsave("plot6_erv_fdi_scatter.png", p6, width = 9, height = 6, dpi = 150)

# --- Plot 7: NEER — non-euro members ----------------------------
p7 <- panel |>
  filter(country_code %in% non_euro) |>
  ggplot(aes(x = year, y = neer_avg, colour = country_name)) +
  geom_line(linewidth = 0.9) +
  geom_hline(yintercept = 100, linetype = "dashed",
             colour = "grey50") +
  scale_x_continuous(breaks = seq(2000, 2022, 4)) +
  labs(title    = "NEER — non-eurozone EU members (2000-2022)",
       subtitle = "Nominal effective exchange rate, index 2015=100",
       x = NULL, y = "NEER (2015=100)", colour = NULL) +
  theme(legend.position = "bottom")
print(p7)
ggsave("plot7_neer.png", p7, width = 10, height = 5.5, dpi = 150)

# --- Plot 8: GDP growth vs FDI scatter --------------------------
p8 <- panel |>
  filter(!is.na(fdi), !is.na(gdp_growth)) |>
  ggplot(aes(x = gdp_growth, y = fdi, colour = group)) +
  geom_point(alpha = 0.4, size = 1.5) +
  geom_smooth(method = "lm", se = FALSE, linewidth = 1) +
  scale_colour_manual(
    values = c("Eurozone" = col_euro, "Non-euro EU" = col_noneuro),
    name = NULL) +
  coord_cartesian(xlim = c(-15, 15), ylim = c(-20, 40)) +
  labs(title    = "GDP growth vs FDI inflows (2000-2022)",
       subtitle = "Each point = one country-year observation",
       x = "GDP growth (%)", y = "FDI (% of GDP)") +
  theme(legend.position = "bottom")
print(p8)
ggsave("plot8_gdp_fdi.png", p8, width = 9, height = 6, dpi = 150)


# ================================================================
# PART 9: PANEL REGRESSION — HANUSCH ET AL. (2018) MODEL
# ================================================================
# Equation 1: FDI ~ current ERV + controls
# Equation 2: FDI ~ lagged ERV + controls
# Both use country + year fixed effects (two-way within estimator)
# Robust standard errors (HC3) applied to both

library(plm)
library(lmtest)
library(sandwich)

# Prepare panel object
panel_plm <- pdata.frame(panel, index = c("country_code", "year"))

# Create lagged ERV (Equation 2)
panel_plm$er_vol_lag <- lag(panel_plm$er_vol, 1)

# --- Hausman test: fixed vs random effects ----------------------
fe <- plm(fdi ~ er_vol + gdp_growth + inflation + trade_open,
          data = panel_plm, model = "within", effect = "twoways")

re <- plm(fdi ~ er_vol + gdp_growth + inflation + trade_open,
          data = panel_plm, model = "random", effect = "twoways")

cat("--- Hausman test ---\n")
print(phtest(fe, re))

# --- Equation 1: contemporaneous ERV ----------------------------
cat("\n--- Equation 1: current ERV ---\n")
eq1 <- plm(fdi ~ er_vol + gdp_growth + inflation + trade_open,
           data   = panel_plm,
           model  = "within",
           effect = "twoways")
print(coeftest(eq1, vcov = vcovHC(eq1, type = "HC3")))
summary(eq1)

# --- Equation 2: lagged ERV -------------------------------------
cat("\n--- Equation 2: lagged ERV ---\n")
eq2 <- plm(fdi ~ er_vol_lag + gdp_growth + inflation + trade_open,
           data   = panel_plm,
           model  = "within",
           effect = "twoways")
print(coeftest(eq2, vcov = vcovHC(eq2, type = "HC3")))

# --- Summary comparison -----------------------------------------
cat("\n--- R-squared comparison ---\n")
cat("Equation 1 R²:", round(summary(eq1)$r.squared[1], 4), "\n")
cat("Equation 2 R²:", round(summary(eq2)$r.squared[1], 4), "\n")


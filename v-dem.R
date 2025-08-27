## hello, this is me trying some things.
# install.packages("tidyr")
# install.packages("tidyverse")
# install.packages("dplyr")
# install.packages("ggplot2")
# install.packages("haven")
# install.packages("foreign")
# install.packages("here")
# install.packages("xlsx")
# install.packages("usethis")
# install.packages("haven")
# install.packages("labelled")
# install.packages("writexl")
# install.packages("openxlsx")
# install.packages("devtools")



library(usethis)
library("tidyr")
library("tidyverse")
library("dplyr")
library("ggplot2")
library("haven")
library("foreign")
library("here")
library("usethis")
library("haven")
library("labelled")
library("writexl")
library(readxl)
library(openxlsx)
library(devtools)
library(vdemdata)
#### V-DEM data: 
# show temporal coverage

rm(list=ls())

# devtools::install_github("vdeminstitute/vdemdata")
library(vdemdata)
df_vdem <- vdemdata::vdem

find_var("equality")
var_info("v2x_polyarchy")


  # Quick plot
  countries <- c("Czechia", "France", "Netherlands", "Hungary")
  
  df2 <- df_vdem %>%
    filter(country_name %in% countries, year >= 2000, year <= 2024) %>%
    select(country_name, year, v2x_polyarchy)
  
  
  library(ggplot2)
  ggplot(df2, aes(x = year, y = v2x_polyarchy, color = country_name)) +
    geom_line() +
    labs(title = "Polyarchy (Electoral democracy) 2000–2024")

## V-DEM -----------------------------------------------------------------------
#### automating across all excel sheets
  # --- define your variable groups ---
  variable_groups <- list(
    Executive = list(prefix = "^v2ex", exclude = "v2exl", sheet = "Executive"),
    Regime = list(prefix = "^v2reg", sheet = "Regime"),
    Legislature = list(prefix = "^v2lg", sheet = "Legislature"),
    Judiciary = list(prefix = "^v2ju", sheet = "Judiciary"),
    `Democracy Indices (V-Dem)` = list(prefix = "^v2x_", sheet = "Democracy Indices (V-Dem)")
  )
  
  # --- load data and labels ---
  data <- vdemdata::vdem
  labels_df <- var_info(names(data)) %>%
    tibble::as_tibble() %>%
    transmute(variable = tag, label = as.character(name), question)
  
  country_lookup <- data.frame(
    country_id = c(157, 76, 91, 210),
    country = c("Czech Republic", "France", "Netherlands", "Hungary")
  )
  
  # --- function to process any variable set ---
  process_vars <- function(variables, data, labels_df, country_lookup) {
    filtered <- data %>%
      filter(country_id %in% country_lookup$country_id, year >= 2000, year <= 2024) %>%
      select(country_id, country_name, year, all_of(variables)) %>%
      left_join(country_lookup, by = "country_id") %>%
      select(country, year, all_of(variables)) %>%
      mutate(across(all_of(variables), as.character))
    
    long <- filtered %>%
      pivot_longer(cols = all_of(variables), names_to = "variable", values_to = "value") %>%
      filter(!is.na(value) & value != "") %>%
      arrange(country, variable, year)
    
    summary <- long %>%
      group_by(country, variable) %>%
      summarise(years_available = paste(sort(unique(year)), collapse = ", "), .groups = "drop") %>%
      left_join(labels_df, by = "variable") %>%
      select(country, variable, label, question, years_available)
    
    summary_wide <- summary %>%
      pivot_wider(names_from = country, values_from = years_available) %>%
      arrange(variable) %>%
      mutate(across(
        c("Czech Republic", "France", "Netherlands", "Hungary"),
        ~ case_when(
          . == paste(2000:2024, collapse = ", ") ~ "2000-2024",
          . == paste(2000:2023, collapse = ", ") ~ "2000-2023",
          TRUE ~ .
        )
      ))
    
    return(summary_wide)
  }
  
  # --- define suffixes to exclude derived variables ---
  derived_suffixes <- c("_codehigh", "_codelow", "_mean", "_nr", "_ord", "_sd", 
                        "_osp", "_0", "_1", "_2", "_3", "_4", "_5", "_6", "_7", 
                        "8", "_9", "_10", "_11", "_12", "_13", "_14", "_15", 
                        "_16", "_17", "_18", "_19", "_20")
  
  # --- helper function to keep only original variables ---
  keep_original_vars <- function(vars) {
    vars[!grepl(paste0(derived_suffixes, collapse = "|"), vars)]
  }
  
  # --- compute main sheets ---
  summary_tables <- list()
  for(group in variable_groups){
    vars <- names(data)[grepl(group$prefix, names(data))]
    if(!is.null(group$exclude)) vars <- setdiff(vars, group$exclude)
    vars <- keep_original_vars(vars)         # <--- remove derived variables
    summary_tables[[group$sheet]] <- process_vars(vars, data, labels_df, country_lookup)
  }
  
  # --- compute "Other" variables ---
  all_vars <- names(data)[grepl("^v", names(data))]
  grouped_vars <- unlist(lapply(variable_groups, function(g){
    vars <- names(data)[grepl(g$prefix, names(data))]
    if(!is.null(g$exclude)) vars <- setdiff(vars, g$exclude)
    vars
  }))
  other_vars <- setdiff(all_vars, grouped_vars)
  other_vars <- keep_original_vars(other_vars)  # <--- remove derived variables
  
  summary_tables$Other <- process_vars(other_vars, data, labels_df, country_lookup)
  
  # --- create workbook and add all sheets at once ---
  wb <- createWorkbook()
  for(sheet_name in names(summary_tables)){
    addWorksheet(wb, sheet_name)
    writeData(wb, sheet_name, summary_tables[[sheet_name]])
    setColWidths(wb, sheet_name, cols = 1:ncol(summary_tables[[sheet_name]]), widths = "auto")
  }
  
  # --- save workbook once ---
  saveWorkbook(wb, "H:/TWIN4DEM/v-dem.xlsx", overwrite = TRUE)

## V-PARTY ---------------------------------------------------------------------
df_vparty <- vdemdata::vparty

  # Checking labels, only for identifier variables given.
  meta_vparty <- var_info(names(vdemdata::vparty))
  View(meta_vparty)

# Get variable metadata
labels_df_vparty <- var_info(names(df_vparty)) %>%
  tibble::as_tibble() %>%
  transmute(
    variable = tag,
    label = as.character(name), 
    question
  )

# Define the countries of interest (same as before, if they exist in vparty)
country_lookup_vparty <- data.frame(
  country_id = c(157, 76, 91, 210),
  country = c("Czech Republic", "France", "Netherlands", "Hungary")
)

# Function for processing V-Party data
process_vparty <- function(data, labels_df, country_lookup) {
  
  # Use all variables
  relevant_vars <- names(data)
  
  # Filter, keep years 2000-2024
  filtered <- data %>%
    filter(country_id %in% country_lookup$country_id, year >= 2000, year <= 2024) %>%
    select(all_of(relevant_vars)) %>%
    left_join(country_lookup, by = "country_id") %>%
    relocate(country, .before = 1) %>%
    mutate(across(where(is.numeric), as.character))  # to avoid NA dropping
  
  # Pivot longer
  long <- filtered %>%
    pivot_longer(
      cols = setdiff(names(filtered), c("country", "year")),
      names_to = "variable", 
      values_to = "value",
      values_transform = list(value = as.character)   
    ) %>%
    filter(!is.na(value) & value != "") %>%
    arrange(country, variable, year)
  
  # Summarize availability
  summary <- long %>%
    group_by(country, variable) %>%
    summarise(years_available = paste(sort(unique(year)), collapse = ", "), .groups = "drop") %>%
    arrange(country, variable) %>%
    left_join(labels_df, by = "variable") %>%
    select(country, variable, label, question, years_available)
  
  # Pivot wider and shorten full ranges
  summary_wide <- summary %>%
    pivot_wider(names_from = country, values_from = years_available) %>%
    arrange(variable) %>%
    mutate(across(
      c("Czech Republic", "France", "Netherlands", "Hungary"),
      ~ case_when(
        . == paste(2000:2024, collapse = ", ") ~ "2000-2024",
        . == paste(2000:2023, collapse = ", ") ~ "2000-2023",
        TRUE ~ .
      )
    ))
  
  return(summary_wide)
}

# Define suffixes to exclude derived variables
derived_suffixes <- c("_codehigh", "_codelow", "_mean", "_nr", "_ord", "_sd", 
                      "_osp", "_0", "_1", "_2", "_3", "_4", "_5", "_6", "_7", 
                      "8", "_9", "_10", "_11", "_12", "_13", "_14", "_15", 
                      "_16", "_17", "_18", "_19", "_20")

# Helper function to keep only original variables
keep_original_vars <- function(vars) {
  vars[!grepl(paste0(derived_suffixes, collapse = "|"), vars)]
}

# Keep only original variables in df_vparty
original_vars <- keep_original_vars(names(df_vparty))

# Process only original variables
summary_vparty <- process_vparty(
  data = df_vparty %>% select(all_of(original_vars), country_id, year),
  labels_df = labels_df_vparty,
  country_lookup = country_lookup_vparty
)

# Load curated questions Excel
questions <- read_excel("H:/TWIN4DEM/v-party_question_texts.xlsx")

# Merge on variable
merged <- summary_vparty %>%
  left_join(questions, by = "variable", suffix = c("", "_curated")) %>%
  mutate(
    label   = coalesce(label_curated, label),       # prefer curated, else original
    question = coalesce(question_curated, question) # prefer curated, else original
  ) %>%
  select(-label_curated, -question_curated) %>%     # clean up
  relocate(label, question, .after = variable)

# Write to Excel (single sheet)
wb_vparty <- createWorkbook()
addWorksheet(wb_vparty, "V-Party")
writeData(wb_vparty, sheet = "V-Party", merged)
setColWidths(wb_vparty, sheet = "V-Party", cols = 1:ncol(merged), widths = "auto")

# Save file
saveWorkbook(wb_vparty, "H:/TWIN4DEM/v-party.xlsx", overwrite = TRUE)




##### MANIFESTO ################################################################
#manifesto <- read_excel("H:/TWIN4DEM/MPDataset_MPDS2024a.xlsx")
manifesto_df <- read_dta("H:/TWIN4DEM/MPDataset_MPDS2024a_stata14.dta")

var_labels <- tibble(
  variable = names(manifesto_df),
  label = sapply(manifesto_df, function(x) attr(x, "label"))
)

# Define countries of interest
country_lookup <- data.frame(
  countryname = c("Czech Republic", "France", "Netherlands", "Hungary")
)

process_manifesto <- function(data, var_labels, country_lookup) {
  
  # Create numeric year from 'date' (YYYYMM)
  data <- data %>%
    mutate(year = floor(date / 100)) %>%
    filter(countryname %in% country_lookup$country, year >= 2000, year <= 2024)
  
  # Identify all variables to summarize
  exclude_cols <- c("countryname", "party", "partyname", "edate", "date", "year", "partyid")
  relevant_vars <- setdiff(names(data), exclude_cols)
  
  # Pivot longer
  long <- data %>%
    pivot_longer(
      cols = all_of(relevant_vars),
      names_to = "variable",
      values_to = "value",
      values_transform = list(value = as.character)
    ) %>%
    filter(!is.na(value) & value != "") %>%
    arrange(countryname, variable, year)
  
  # Summarize years available
  summary <- long %>%
    group_by(countryname, variable) %>%
    summarise(years_available = paste(sort(unique(year)), collapse = ", "), .groups = "drop") %>%
    left_join(var_labels, by = "variable") %>%
    select(countryname, variable, label, years_available)
  
  # Pivot wider so each country is a column
  summary_wide <- summary %>%
    pivot_wider(names_from = countryname, values_from = years_available) %>%
    arrange(variable)
  
  return(summary_wide)
}
# -----------------------------
# Run processing
# -----------------------------
summary_manifesto <- process_manifesto(manifesto_df, var_labels, country_lookup)

# -----------------------------
# Write to Excel
# -----------------------------
wb <- createWorkbook()
addWorksheet(wb, "Manifesto")
writeData(wb, sheet = "Manifesto", summary_manifesto)
setColWidths(wb, sheet = "Manifesto", cols = 1:ncol(summary_manifesto), widths = "auto")
saveWorkbook(wb, "H:/TWIN4DEM/manifesto_coverage.xlsx", overwrite = TRUE)


### PARTY FACTS ################################################################

### 1. Overview: which datasets in which years
core_parties <- read_csv("H:/TWIN4DEM/partyfacts-core-parties.csv")

external_parties <- read_csv("H:/TWIN4DEM/partyfacts-external-parties.csv")

# download and read Party Facts mapping table
file_name <- "partyfacts-mapping.csv"
if( ! file_name %in% list.files()) {
  url <- "https://partyfacts.herokuapp.com/download/external-parties-csv/"
  download.file(url, file_name)
}
partyfacts_raw <- read_csv(file_name, guess_max = 50000)
partyfacts <- partyfacts_raw |> filter(! is.na(partyfacts_id))

# link datasets (select only linked parties)
cat(paste(unique(partyfacts$dataset_key), collapse='\n'))
dataset_1 <- partyfacts |> filter(dataset_key == "manifesto")
dataset_2 <- partyfacts |> filter(dataset_key == "parlgov")
link_table <-
  dataset_1 |>
  inner_join(dataset_2, by = c("partyfacts_id" = "partyfacts_id"))

# write results into file with dataset names in file name
file_out <- "partyfacts-linked.csv"
write_csv(link_table, file_out)


## connecting metadata
# 1. Load PartyFacts mapping table
file_name <- 'partyfacts-mapping.csv'
if (!file_name %in% list.files()) {
  download.file('https://partyfacts.herokuapp.com/download/external-parties-csv/', file_name)
}
pf <- read_csv(file_name, show_col_types = FALSE)

# 2. Filter for your 4 countries (use ISO codes as in PartyFacts)
countries <- c("CZE", "FRA", "NLD", "HUN")
pf_filt <- pf %>% filter(country %in% countries)

# 3. Keep only relevant columns for dataset-year mapping
pf_dt_years <- pf_filt %>%
  filter(!is.na(share_year) & share_year >= 2000 & share_year <= 2024) %>%
  distinct(country, dataset_key, share_year)

# 4. Create coverage table: dataset × country × years
coverage <- pf_dt_years %>%
  group_by(dataset_key, country) %>%
  summarise(years = paste(sort(unique(share_year)), collapse = ", "), .groups = "drop") %>%
  pivot_wider(names_from = country, values_from = years)

# 5. Write to Excel
wb <- createWorkbook()
addWorksheet(wb, "DatasetCoverage")
writeData(wb, "DatasetCoverage", coverage)
setColWidths(wb, "DatasetCoverage", cols = 1:ncol(coverage), widths = "auto")
saveWorkbook(wb, "H:/TWIN4DEM/partyfacts_dataset_coverage.xlsx", overwrite = TRUE)


### 2.meta-data tables for each dataset in partyfacts
library(DBI)
library(RSQLite)
# CHES

# --- Define countries ---
countries <- c("Czech Republic", "France", "Netherlands", "Hungary")

# --- Function to process a generic dataset ---
process_dataset <- function(data, country_col="country", year_col="year", labels_df=NULL) {
  
  # Identify variables (exclude country and year)
  vars <- setdiff(names(data), c(country_col, year_col))
  
  
  # Filter for countries and years
  filtered <- data %>%
    filter(.data[[country_col]] %in% countries,
           .data[[year_col]] >= 2000,
           .data[[year_col]] <= 2024)
  
  # Pivot longer
  long <- filtered %>%
    pivot_longer(cols = all_of(vars), names_to = "variable", values_to = "value") %>%
    filter(!is.na(value))
  
  # Summarize years available per country and variable
  summary <- long %>%
    group_by(.data[[country_col]], variable) %>%
    summarise(years_available = paste(sort(unique(.data[[year_col]])), collapse = ", "), .groups="drop")
  
  # Pivot wider: columns = countries
  summary_wide <- summary %>%
    pivot_wider(names_from = .data[[country_col]], values_from = years_available) %>%
    arrange(variable)
  
  # Add labels if provided
  if(!is.null(labels_df)){
    summary_wide <- summary_wide %>%
      left_join(labels_df, by="variable") %>%
      relocate(label, question, .after=variable)
  }
  
  return(summary_wide)
}

# --- Load CHES ---
ches <- read_dta("H:/TWIN4DEM/1999-2019_CHES_dataset_means(v3).dta")
ches_country_codes <- c(CZ = 21, FR = 6, HU = 23, NL = 10)

# Extract labels from CHES
labels_ches <- tibble(
  variable = names(ches),
  label = sapply(ches, function(x) attr(x, "label")),
  question = NA_character_
)

ches_clean <- ches %>%
  mutate(across(everything(), as.character)) %>%       # convert all columns to character to avoid pivot errors
  mutate(country = case_when(
    country == 21 ~ "Czech Republic",
    country == 6  ~ "France",
    country == 23 ~ "Hungary",
    country == 10 ~ "Netherlands",
    TRUE ~ NA_character_
  )) %>%
  filter(!is.na(country))   # keep only the four countries

# Filter and process only selected countries
summary_ches <- process_dataset(
  ches_clean %>% mutate(country = country),  # optional: add numeric country code column
  country_col = "country", 
  year_col = "year", 
  labels_df = labels_ches
)


# --- Load CLEA ---
clea <- read_dta("H:/TWIN4DEM/clea_lc_20240419.dta")  
labels_clea <- tibble(
  variable = names(clea),
  label = sapply(clea, function(x) attr(x, "label")),
  question = NA_character_
)
# --- Clean dataset (convert everything to character to avoid pivot issues) ---
clea_clean <- clea %>%
  mutate(across(everything(), as.character)) %>%
  filter(ctr_n %in% countries)

# --- Reuse the process_dataset function ---
process_dataset <- function(data, country_col="country", year_col="year", labels_df=NULL) {
  
  # Identify variables (exclude country and year)
  vars <- setdiff(names(data), c(country_col, year_col))
  
  
  # Filter for countries and years
  filtered <- data %>%
    filter(.data[[country_col]] %in% countries,
           as.numeric(.data[[year_col]]) >= 2000,
           as.numeric(.data[[year_col]]) <= 2024)
  
  # Pivot longer
  long <- filtered %>%
    pivot_longer(cols = all_of(vars), names_to = "variable", values_to = "value") %>%
    filter(!is.na(value) & value != "")
  
  # Summarize years available per country and variable
  summary <- long %>%
    group_by(.data[[country_col]], variable) %>%
    summarise(years_available = paste(sort(unique(.data[[year_col]])), collapse = ", "), .groups="drop")
  
  # Pivot wider: columns = countries
  summary_wide <- summary %>%
    pivot_wider(names_from = .data[[country_col]], values_from = years_available) %>%
    arrange(variable)
  
  # Add labels if provided
  if(!is.null(labels_df)){
    summary_wide <- summary_wide %>%
      left_join(labels_df, by="variable") %>%
      relocate(label, question, .after=variable)
  }
  
  return(summary_wide)
}

# --- Process CLEA dataset ---
summary_clea <- process_dataset(
  clea_clean,
  country_col = "ctr_n",
  year_col = "yr",
  labels_df = labels_clea
)

# --- Load DPI dataset ---
dpi <- read_dta("H:/TWIN4DEM/dpi2015-stata13.dta")  
labels_dpi <- tibble(
  variable = names(dpi),
  label = sapply(dpi, function(x) attr(x, "label")),
  question = NA_character_
)

# --- Clean dataset (convert everything to character to avoid pivot issues) ---
dpi_clean <- dpi %>%
  mutate(across(everything(), as.character)) %>%
  filter(countryname %in% countries)

# --- Reuse the process_dataset function ---
process_dataset <- function(data, country_col="countryname", year_col="year", labels_df=NULL) {
  
  # Identify variables (exclude country and year)
  vars <- setdiff(names(data), c(country_col, year_col))
  
  
  # Filter for countries and years
  filtered <- data %>%
    filter(.data[[country_col]] %in% countries,
           as.numeric(.data[[year_col]]) >= 2000,
           as.numeric(.data[[year_col]]) <= 2024)
  
  # Pivot longer
  long <- filtered %>%
    pivot_longer(cols = all_of(vars), names_to = "variable", values_to = "value") %>%
    filter(!is.na(value) & value != "")
  
  # Summarize years available per country and variable
  summary <- long %>%
    group_by(.data[[country_col]], variable) %>%
    summarise(years_available = paste(sort(unique(.data[[year_col]])), collapse = ", "), .groups="drop")
  
  # Pivot wider: columns = countries
  summary_wide <- summary %>%
    pivot_wider(names_from = .data[[country_col]], values_from = years_available) %>%
    arrange(variable)
  
  # Add labels if provided
  if(!is.null(labels_df)){
    summary_wide <- summary_wide %>%
      left_join(labels_df, by="variable") %>%
      relocate(label, question, .after=variable)
  }
  
  return(summary_wide)
}

# --- Process DPI dataset ---
summary_dpi <- process_dataset(
  dpi_clean,
  country_col = "countryname",
  year_col = "year",
  labels_df = labels_dpi
)


# --- Load PAGED dataset ---

paged_country_codes <- c("Netherlands" , "France" , "Czechia", "Hungary")

paged <- read_dta("H:/TWIN4DEM/PAGED-basic.dta")  
labels_paged <- tibble(
  variable = names(paged),
  label = sapply(paged, function(x) attr(x, "label")),
  question = NA_character_
)

# --- Clean dataset ---
paged_clean <- paged %>%
  mutate(
    year = year(as.Date(elecdate)),   
    across(everything(), as.character)
  ) %>%
  filter(country_name %in% paged_country_codes)



# --- Process PAGED dataset ---
summary_paged <- paged %>%
  filter(country_name %in% paged_country_codes) %>%
  mutate(year_in = as.numeric(year_in)) %>%
  filter(year_in >= 2000, year_in <= 2024) %>%
  group_by(country_name) %>%
  summarise(years_available = paste(sort(unique(year_in)), collapse = ", "),
            .groups = "drop") %>%
  pivot_wider(names_from = country_name, values_from = years_available)

# Create a row for each variable in REPDEM
summary_paged <- tibble(variable = setdiff(names(paged), c("country_name","year_in","year_out"))) %>%
  left_join(labels_paged, by = "variable") %>%
  bind_cols(summary_paged[rep(1, nrow(.)), ]) %>%
  relocate(label, question, .after = variable) %>%
  arrange(variable)


# --- Create workbook ---
wb <- createWorkbook()

# Add CHES sheet
addWorksheet(wb, "CHES")
writeData(wb, "CHES", summary_ches)
setColWidths(wb, "CHES", cols = 1:ncol(summary_ches), widths = "auto")

# Add CLEA sheet
addWorksheet(wb, "CLEA")
writeData(wb, "CLEA", "Note: Years shown correspond to election years (yr)", startRow = 1)
writeData(wb, "CLEA", summary_clea)
setColWidths(wb, "CLEA", cols = 1:ncol(summary_clea), widths = "auto")

# Add DPI sheet
addWorksheet(wb, "DPI")
writeData(wb, "DPI", summary_dpi)
setColWidths(wb, "DPI", cols = 1:ncol(summary_dpi), widths = "auto")

# Add PAGED sheet
addWorksheet(wb, "PAGED")
writeData(wb, "PAGED", "Note: Years shown correspond to cabinet formation years (year_in)", startRow = 1)
writeData(wb, "PAGED", summary_paged, startRow = 3)  # table starts below note
setColWidths(wb, "PAGED", cols = 1:ncol(summary_repdem), widths = "auto")

# Save to file
saveWorkbook(wb, "H:/TWIN4DEM/partyfacts_prep_temporal_coverage.xlsx", overwrite = TRUE)



 


# --- Load ParlGov ---
pg <- dbConnect(RSQLite::SQLite(), "H:/TWIN4DEM/parlgov-stable.db")
# Example: party-level table
dbListTables(pg)

country <- dbReadTable(pg, "country")
party   <- dbReadTable(pg, "party")
election <- dbReadTable(pg, "election")
election_result <- dbReadTable(pg, "election_result")
cabinet_party <- dbReadTable(pg, "cabinet_party")
cabinet <- dbReadTable(pg, "cabinet")

  names(party)
  names(election)
  names(election_result)

  # --- Focus countries ---
  countries <- c("Czech Republic", "France", "Netherlands", "Hungary")
  # build lookup 
  country_lookup <- country %>%
    transmute(country_id = id, country = name) %>%
    filter(country %in% countries)
  
  summary_election <- election %>%
    mutate(year = year(as.Date(date))) %>%
    filter(year >= 2000, year <= 2024,
           country_id %in% country_lookup$country_id) %>%
    left_join(country_lookup, by = "country_id") %>%
    # pivot *after* you have 'country' available
    pivot_longer(
      cols = -c(country_id, country, year, id),
      names_to = "variable",
      values_to = "value",
      values_transform = list(value = as.character)  # unify types for the value column
    ) %>%
    filter(!is.na(value) & value != "") %>%
    group_by(country, variable) %>%
    summarise(
      years_available = paste(sort(unique(year)), collapse = ", "),
      .groups = "drop"
    ) %>%
    pivot_wider(names_from = country, values_from = years_available) %>%
    arrange(variable)
  
 
  
  # --- Party table (time-invariant means) ---
  summary_party <- tibble(variable = setdiff(names(party), "country_id")) %>%
    mutate(`Czech Republic` = "time-invariant",
           France = "time-invariant",
           Hungary = "time-invariant",
           Netherlands = "time-invariant") %>%
    arrange(variable)
  
  # --- Cabinet (complete 2000-2023) ---
  summary_cabinet <- tibble(variable = setdiff(names(cabinet), "country_id")) %>%
    mutate(`Czech Republic` = "2000-2023",
           France = "2000-2023",
           Hungary = "2000-2023",
           Netherlands = "2000-2023") %>%
    arrange(variable)
  
  
  # --- Write to Excel ---
  wb <- createWorkbook()
  addWorksheet(wb, "Election")
  addWorksheet(wb, "Party")
  addWorksheet(wb, "Cabinet")
  
  writeData(wb, "Election", "Note: Years shown correspond to election years (date), including EU elections", startRow = 1)
  writeData(wb, "Election", summary_election)
  writeData(wb, "Party", summary_party)
  writeData(wb, "Cabinet", summary_cabinet)
  
  setColWidths(wb, "Election", cols = 1:ncol(summary_election), widths = "auto")
  setColWidths(wb, "Party", cols = 1:ncol(summary_party), widths = "auto")
  setColWidths(wb, "Cabinet", cols = 1:ncol(summary_cabinet), widths = "auto")
  
  saveWorkbook(wb, "H:/TWIN4DEM/parlgov_summary.xlsx", overwrite=TRUE)
  
  
  
  ##--- merging into one overview ----------------
# ----------------------------
# Load PartyFacts external IDs
# ----------------------------
pf_links <- read_csv("https://partyfacts.herokuapp.com/download/external-parties-csv/",
                     na = "", guess_max = 50000)

# ----------------------------
# CHES - party-level
# ----------------------------
ches <- read_dta("H:/TWIN4DEM/1999-2019_CHES_dataset_means(v3).dta")

# load CHES-PartyFacts mapping
ches_map <- pf_links %>%
  filter(dataset_key == "ches") %>%
  select(dataset_party_id, partyfacts_id) %>%
  mutate(dataset_party_id = as.character(dataset_party_id))

ches_linked <- ches %>%
  mutate(dataset_party_id = as.character(party_id)) %>%  
  left_join(ches_map, by = "dataset_party_id")

# ----------------------------
# CLEA - party-level
# ----------------------------
clea <- read_dta("H:/TWIN4DEM/clea_lc_20240419.dta")

clea_map <- pf_links %>%
  filter(dataset_key == "clea") %>%
  select(dataset_party_id, partyfacts_id)  %>%
  mutate(dataset_party_id = as.character(dataset_party_id))

clea_prepped <- clea %>%
  mutate(dataset_party_id = ctr * 1e6 + pty) %>%   
  mutate(dataset_party_id = as.character(dataset_party_id))

clea_linked <- clea_prepped %>%
  left_join(clea_map, by = "dataset_party_id")

# ----------------------------
# DPI - country-level
# ----------------------------
dpi <- read_dta("H:/TWIN4DEM/dpi2015-stata13.dta")

# PartyFacts DPI mapping
dpi_map <- pf_links %>%
  filter(dataset_key == "dpi") %>%
  select(dataset_party_id, partyfacts_id) %>%
  mutate(dataset_party_id = as.numeric(dataset_party_id))

# Prepare DPI: keep party identifiers (like prtyin)
dpi_prepped <- dpi %>%
  mutate(dataset_party_id = prtyin)   # or gov1party etc. depending on your needs

# Link to PartyFacts
dpi_linked <- dpi_prepped %>%
  left_join(dpi_map, by = "dataset_party_id")

# ----------------------------
# PAGED
# ----------------------------
paged_linked <- read_dta("H:/TWIN4DEM/PAGED-basic-party-dataset.dta")

paged_map <- pf_links %>%
  filter(dataset_key == "paged") %>%
  select(dataset_party_id, partyfacts_id)



# ----------------------------
# Combine all datasets
# ----------------------------


# Standardize country variables first
ches_clean <- ches_linked %>% mutate(country_std = country)
clea_clean <- clea_linked %>% mutate(country_std = ctr_n)
dpi_clean  <- dpi_linked  %>% mutate(country_std = countryname)
paged_clean<- paged_linked %>% mutate(country_std = country_name)

# Filter CHES for the 4 countries and convert to character names
ches_sel <- ches_sel %>%
  filter(country %in% c(21, 6, 23, 10)) %>%   # numeric codes
  mutate(
    country_std = case_when(
      country == 21 ~ "Czech Republic",
      country == 6  ~ "France",
      country == 23 ~ "Hungary",
      country == 10 ~ "Netherlands"
    )
  ) %>%
  mutate(country_std = as.character(country_std))

# Select relevant columns and rename year variables consistently
ches_sel  <- ches_sel %>% rename(year = year)
clea_sel  <- clea_clean %>% rename(year = yr)
dpi_sel   <- dpi_clean  %>% rename(year = year)
paged_sel <- paged_clean %>%
  mutate(
    year = year(date_in)  # extracts only the year from a Date
  )

# Merge datasets using left_join on partyfacts_id
integrated <- ches_sel %>%
  left_join(clea_sel, by = c("partyfacts_id", "country_std", "year")) %>%
  left_join(dpi_sel,  by = c("partyfacts_id", "country_std", "year")) %>%
  left_join(paged_sel,by = c("partyfacts_id", "country_std", "year"))


countries_of_interest <- c("Czech Republic", "France", "Netherlands", "Hungary")

integrated_filtered <- integrated %>%
  filter(country_std %in% countries_of_interest,
         year >= 2000,
         year <= 2024)

# Identify variables (exclude key columns)
key_vars <- setdiff(names(integrated_filtered), c("partyfacts_id", "country_std", "year"))

# Convert all variables to character
integrated_filtered <- integrated_filtered %>%
  mutate(across(all_of(key_vars), as.character))

# Then pivot longer
integrated_long <- integrated_filtered %>%
  pivot_longer(
    cols = all_of(key_vars),
    names_to = "variable",
    values_to = "value"
  ) %>%
  filter(!is.na(value) & value != "")


temporal_coverage <- integrated_long %>%
  group_by(variable, country_std) %>%
  summarise(years_available = paste(sort(unique(year)), collapse = ", "), .groups="drop") %>%
  pivot_wider(names_from = country_std, values_from = years_available) %>%
  arrange(variable)

# ----------------------------
# Write to Excel 
# ----------------------------
wb <- createWorkbook()

for(nm in names(integrated)){
  addWorksheet(wb, nm)
  writeData(wb, nm, integrated[[nm]])
  setColWidths(wb, nm, cols = 1:ncol(integrated[[nm]]), widths = "auto")
}

saveWorkbook(wb, "H:/TWIN4DEM/integrated_partyfacts_dataset.xlsx", overwrite = TRUE)
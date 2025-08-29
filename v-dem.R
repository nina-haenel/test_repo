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
  
  
  

  
  
# ## # === function to process any variable family and write Excel sheet ===
  # --- load data and labels ---
  data <- vdemdata::vdem
  labels_df <- var_info(names(data)) %>%
    tibble::as_tibble() %>%
    transmute(variable = tag, label = as.character(name), question)
  
  country_lookup <- data.frame(
    country_id = c(157, 76, 91, 210),
    country = c("Czech Republic", "France", "Netherlands", "Hungary")
  )
  # === function to process any variable family and write Excel sheet ===
  make_vdem_sheet <- function(data, var_prefix, sheet_name, 
                              country_lookup, years = 2000:2024, 
                              labels_df = NULL, outfile = "VDEM_output.xlsx") {
    
    # Exclude suffixes for derived variables
    derived_suffixes <- c("_codehigh","_codelow","_mean","_nr","_ord","_sd","_osp",
                          "_0","_1","_2","_3","_4","_5","_6","_7","_8","_9",
                          "_10","_11","_12","_13","_14","_15","_16","_17","_18","_19","_20")
    suf_pat <- paste0("(", paste(derived_suffixes, collapse="|"), ")$")
    
    # --- find relevant variables ---
    vars <- names(data)[grepl(paste0("^", var_prefix), names(data))]
    vars <- vars[!grepl(suf_pat, vars)]
    
    # --- build coverage table ---
    filtered <- data %>%
      filter(country_id %in% country_lookup$country_id, year %in% years) %>%
      left_join(country_lookup, by = "country_id") %>%
      select(country, year, all_of(vars))
    
    covered <- filtered %>%
      mutate(across(all_of(vars), ~ {
        x <- .
        if (is.character(x)) x <- na_if(trimws(x), "")
        as.integer(!is.na(x))
      }))
    
    long <- covered %>%
      pivot_longer(cols = all_of(vars), names_to = "variable", values_to = "coverage") %>%
      complete(variable, year = years, country = country_lookup$country, fill = list(coverage = 0))
    
    wide <- long %>%
      pivot_wider(names_from = c(year, country), names_sep = "_", values_from = coverage)
    
    if (!is.null(labels_df)) {
      wide <- wide %>%
        left_join(labels_df, by = "variable") %>%
        relocate(variable, label, question)
      base_cols <- c("variable","label","question")
    } else {
      base_cols <- "variable"
    }
    
    # --- parse column names to reorder ---
    yc_cols <- grep("\\d{4}_.+", names(wide), value = TRUE)
    col_info <- tibble(orig = yc_cols) %>%
      mutate(
        is_year_first = grepl("^\\d{4}_", orig),
        year = ifelse(is_year_first, sub("^([0-9]{4})_(.*)$", "\\1", orig),
                      ifelse(grepl("_(\\d{4})$", orig), sub("^(.*)_(\\d{4})$", "\\2", orig), NA)),
        country = ifelse(is_year_first, sub("^([0-9]{4})_(.*)$", "\\2", orig),
                         ifelse(grepl("_(\\d{4})$", orig), sub("^(.*)_(\\d{4})$", "\\1", orig), NA))
      ) %>%
      mutate(year = as.integer(year),
             country = trimws(country),
             country_factor = factor(country, levels = country_lookup$country)) %>%
      arrange(year, country_factor)
    
    ordered_cols <- col_info$orig
    df_for_excel <- wide %>% select(all_of(base_cols), all_of(ordered_cols))
    
    # --- create sheet ---
    if (!exists("wb_global", envir = .GlobalEnv)) {
      assign("wb_global", createWorkbook(), envir = .GlobalEnv)
    }
    wb <- get("wb_global", envir = .GlobalEnv)
    
    addWorksheet(wb, sheet_name)
    
    # headers
    top_header <- c(base_cols, as.character(col_info$year))
    bottom_header <- c(rep("", length(base_cols)), col_info$country)
    out <- rbind(top_header, bottom_header, as.matrix(df_for_excel))
    writeData(wb, sheet_name, out, colNames = FALSE)
    
    # merge year headers
    start_col <- length(base_cols) + 1
    for (yr in unique(col_info$year)) {
      cols <- which(col_info$year == yr) + length(base_cols)
      mergeCells(wb, sheet = sheet_name, cols = cols, rows = 1)
    }
    
    # style 1's green
    greenStyle <- createStyle(fgFill = "#C6EFCE")
    num_mat <- df_for_excel %>% select(all_of(ordered_cols)) %>% mutate(across(everything(), as.integer))
    one_positions <- which(num_mat == 1, arr.ind = TRUE)
    if (nrow(one_positions) > 0) {
      rows_cells <- one_positions[,1] + 2
      cols_cells <- one_positions[,2] + length(base_cols)
      addStyle(wb, sheet = sheet_name, style = greenStyle, rows = rows_cells, cols = cols_cells, gridExpand = FALSE)
    }
    
    # save
    saveWorkbook(wb, outfile, overwrite = TRUE)
    assign("wb_global", wb, envir = .GlobalEnv)
    message("Sheet ", sheet_name, " written to ", outfile)
  }
  
  
  
  # Example: Executive sheet
  make_vdem_sheet(data, var_prefix = "v2ex", sheet_name = "Executive",
                  country_lookup = country_lookup, labels_df = labels_df,
                  outfile = "H:/TWIN4DEM/VDEM_full.xlsx")
  
  # Legislature sheet
  make_vdem_sheet(data, var_prefix = "v2lg", sheet_name = "Legislature",
                  country_lookup = country_lookup, labels_df = labels_df,
                  outfile = "H:/TWIN4DEM/VDEM_full.xlsx")
  
  # Judiciary sheet
  make_vdem_sheet(data, var_prefix = "v2ju", sheet_name = "Judiciary",
                  country_lookup = country_lookup, labels_df = labels_df,
                  outfile = "H:/TWIN4DEM/VDEM_full.xlsx")
  
  # Regime sheet
  make_vdem_sheet(data, var_prefix = "^v2reg", sheet_name = "Regime",
                  country_lookup = country_lookup, labels_df = labels_df,
                  outfile = "H:/TWIN4DEM/VDEM_full.xlsx")
  
  # Democracy Indices sheet
  make_vdem_sheet(data, var_prefix = "^v2x_", sheet_name = "Democracy Indices",
                  country_lookup = country_lookup, labels_df = labels_df,
                  outfile = "H:/TWIN4DEM/VDEM_full.xlsx")
  
  # Others sheet
  # all variable names
  all_vars <- names(data)
  
  # variables already used in other sheets
  used_vars <- c(
    grep("^v2ex", all_vars, value = TRUE),
    grep("^v2lg", all_vars, value = TRUE),
    grep("^v2ju", all_vars, value = TRUE),
    grep("^v2reg", all_vars, value = TRUE),
    grep("^v2x_", all_vars, value = TRUE)
  )
  
  # exclude identifier columns and derived suffixes
  derived_suffixes <- c("_codehigh","_codelow","_mean","_nr","_ord","_sd","_osp",
                        "_0","_1","_2","_3","_4","_5","_6","_7","_8","_9",
                        "_10","_11","_12","_13","_14","_15","_16","_17","_18","_19","_20")
  suf_pat <- paste0("(", paste(derived_suffixes, collapse="|"), ")$")
  
  other_vars <- setdiff(all_vars, c(used_vars, "country_id", "country_name", "year"))
  other_vars <- other_vars[!grepl(suf_pat, other_vars)]
  
  # create a minimal wrapper for "Other"
  make_vdem_other <- function() {
    filtered <- data %>%
      filter(country_id %in% country_lookup$country_id, year %in% 2000:2024) %>%
      left_join(country_lookup, by = "country_id") %>%
      select(country, year, all_of(other_vars))
    
    covered <- filtered %>%
      mutate(across(all_of(other_vars), ~ as.integer(!is.na(.))))
    
    long <- covered %>%
      pivot_longer(cols = all_of(other_vars), names_to = "variable", values_to = "coverage") %>%
      complete(variable, year = 2000:2024, country = country_lookup$country, fill = list(coverage = 0))
    
    wide <- long %>%
      pivot_wider(names_from = c(year, country), names_sep = "_", values_from = coverage)
    
    if (!is.null(labels_df)) {
      wide <- wide %>%
        left_join(labels_df, by = "variable") %>%
        relocate(variable, label, question)
      base_cols <- c("variable","label","question")
    } else {
      base_cols <- "variable"
    }
    
    # reorder columns by year then country
    yc_cols <- grep("\\d{4}_.+", names(wide), value = TRUE)
    col_info <- tibble(orig = yc_cols) %>%
      mutate(
        year = as.integer(sub("^([0-9]{4})_.*$", "\\1", orig)),
        country = sub("^\\d{4}_(.*)$", "\\1", orig),
        country_factor = factor(country, levels = country_lookup$country)
      ) %>%
      arrange(year, country_factor)
    ordered_cols <- col_info$orig
    df_for_excel <- wide %>% select(all_of(base_cols), all_of(ordered_cols))
    
    # write to Excel
    wb <- get("wb_global", envir = .GlobalEnv)
    addWorksheet(wb, "Other")
    top_header <- c(base_cols, as.character(col_info$year))
    bottom_header <- c(rep("", length(base_cols)), col_info$country)
    out <- rbind(top_header, bottom_header, as.matrix(df_for_excel))
    writeData(wb, "Other", out, colNames = FALSE)
    
    # merge year headers
    start_col <- length(base_cols) + 1
    for (yr in unique(col_info$year)) {
      cols <- which(col_info$year == yr) + length(base_cols)
      mergeCells(wb, sheet = "Other", cols = cols, rows = 1)
    }
    
    # style 1's green
    greenStyle <- createStyle(fgFill = "#C6EFCE")
    num_mat <- df_for_excel %>% select(all_of(ordered_cols)) %>% mutate(across(everything(), as.integer))
    one_positions <- which(num_mat == 1, arr.ind = TRUE)
    if (nrow(one_positions) > 0) {
      rows_cells <- one_positions[,1] + 2
      cols_cells <- one_positions[,2] + length(base_cols)
      addStyle(wb, sheet = "Other", style = greenStyle, rows = rows_cells, cols = cols_cells, gridExpand = FALSE)
    }
    
    assign("wb_global", wb, envir = .GlobalEnv)
    saveWorkbook(wb, "H:/TWIN4DEM/VDEM_full.xlsx", overwrite = TRUE)
    message("Other sheet written successfully.")
  }
  
  # call it
  make_vdem_other()
 
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


# V-PARTY NEW FORMAT  -------------------------------
# --- Load V-Party data ---
df_vparty <- vdemdata::vparty

# Variable metadata
labels_df_vparty <- var_info(names(df_vparty)) %>%
  tibble::as_tibble() %>%
  transmute(variable = tag,
            label = as.character(name),
            question)

# Define countries of interest
country_lookup <- data.frame(
  country_id = c(157, 76, 91, 210),
  country = c("Czech Republic", "France", "Netherlands", "Hungary")
)

# --- Define derived suffixes to exclude ---
derived_suffixes <- c("_codehigh","_codelow","_mean","_nr","_ord","_sd",
                      "_osp","_0","_1","_2","_3","_4","_5","_6","_7",
                      "_8","_9","_10","_11","_12","_13","_14","_15",
                      "_16","_17","_18","_19","_20")

# Helper to keep only original variables
keep_original_vars <- function(vars) {
  vars[!grepl(paste0(derived_suffixes, collapse="|"), vars)]
}

# --- Prepare coverage table ---
vars <- setdiff(names(df_vparty), c("country_id", "year"))
original_vars <- keep_original_vars(vars)

filtered <- df_vparty %>%
  filter(country_id %in% country_lookup_vparty$country_id,
         year >= 2000, year <= 2024) %>%
  select(all_of(original_vars), country_id, year) %>%
  left_join(country_lookup_vparty, by = "country_id") %>%
  relocate(country, .before = 1)

# convert all variables to 0/1 coverage
coverage <- filtered %>%
  mutate(across(all_of(original_vars), ~ as.integer(!is.na(.)))) %>%
  pivot_longer(cols = all_of(original_vars), names_to = "variable", values_to = "coverage") %>%
  group_by(variable, year, country) %>%
  summarise(coverage = max(coverage), .groups = "drop") %>% # ensure single value
  complete(variable, year = 2000:2024, country = country_lookup_vparty$country, fill = list(coverage = 0)) %>%
  pivot_wider(names_from = c(year, country), names_sep = "_", values_from = coverage)

# --- Join labels and question texts ---
coverage_full <- labels_df_vparty %>%
  select(variable, label, question) %>%
  right_join(coverage, by = "variable") %>%   # keep all coverage rows
  relocate(label, question, .after = variable)

# --- Write to Excel ---
wb <- createWorkbook()
sheet_name <- "V-Party"
addWorksheet(wb, sheet_name)

writeData(wb, sheet_name, coverage_full, colNames = TRUE)

# --- Style 1's green ---
greenStyle <- createStyle(fgFill = "#C6EFCE")
num_mat <- coverage_full %>%
  select(-variable, -label, -question) %>%
  mutate(across(everything(), as.integer))

one_positions <- which(num_mat == 1, arr.ind = TRUE)
if (nrow(one_positions) > 0) {
  rows_cells <- one_positions[,1] + 1  # +1 for header row
  cols_cells <- one_positions[,2] + 3  # +3 for variable/label/question
  addStyle(wb, sheet = sheet_name, style = greenStyle, rows = rows_cells, cols = cols_cells, gridExpand = FALSE)
}

setColWidths(wb, sheet_name, cols = 1:ncol(coverage_full), widths = "auto")
saveWorkbook(wb, "H:/TWIN4DEM/v-party_coverage.xlsx", overwrite = TRUE)


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

# Run processing
summary_manifesto <- process_manifesto(manifesto_df, var_labels, country_lookup)


# Write to Excel

wb <- createWorkbook()
addWorksheet(wb, "Manifesto")
writeData(wb, sheet = "Manifesto", summary_manifesto)
setColWidths(wb, sheet = "Manifesto", cols = 1:ncol(summary_manifesto), widths = "auto")
saveWorkbook(wb, "H:/TWIN4DEM/manifesto_coverage.xlsx", overwrite = TRUE)

# MANIFESTO NEW FORMAT --------------------------------------------------------
# --- Load Manifesto data ---
manifesto_df <- read_dta("H:/TWIN4DEM/MPDataset_MPDS2024a_stata14.dta")

# --- Extract variable labels ---
var_labels <- tibble(
  variable = names(manifesto_df),
  label = sapply(manifesto_df, function(x) attr(x, "label")),
  question = NA_character_
)

# --- Define countries of interest ---
country_lookup <- tibble(
  countryname = c("Czech Republic", "France", "Netherlands", "Hungary")
)

# --- Identify variables to summarize ---
exclude_cols <- c("countryname", "party", "partyname", "edate", "date", "partyid")
manifesto_vars <- setdiff(names(manifesto_df), exclude_cols)

# --- Clean and filter ---
manifesto_clean <- manifesto_df %>%
  mutate(
    year = floor(as.numeric(date)/100),   # extract year from YYYYMM
    countryname = case_when(
      countryname == "Czech Rep." ~ "Czech Republic",
      TRUE ~ countryname
    )
  ) %>%
  filter(countryname %in% country_lookup$countryname,
         year >= 2000, year <= 2024) %>%
  mutate(across(all_of(manifesto_vars), as.character))  # unify types

# --- Pivot longer and create coverage ---
years <- 2000:2024
manifesto_long <- manifesto_clean %>%
  pivot_longer(
    cols = all_of(manifesto_vars),
    names_to = "variable",
    values_to = "value"
  ) %>%
  mutate(coverage = as.integer(!is.na(value) & value != "")) %>%
  group_by(variable, year, countryname) %>%
  summarise(coverage = as.integer(any(coverage == 1)), .groups = "drop") %>%
  complete(
    variable,
    year = years,
    countryname = country_lookup$countryname,
    fill = list(coverage = 0)
  )

# --- Pivot wider: columns = YEAR_COUNTRY ---
wide <- manifesto_long %>%
  mutate(colname = paste0(year, "_", countryname)) %>%
  select(variable, colname, coverage) %>%
  pivot_wider(names_from = colname, values_from = coverage)

# --- Add labels ---
wide <- wide %>%
  left_join(var_labels, by = "variable") %>%
  relocate(label, question, .after = variable)

# --- Identify base columns and numeric cols ---
base_cols <- "variable"
num_cols <- setdiff(names(wide), base_cols)

# --- Reorder columns by year and country ---
col_info <- tibble(orig = num_cols) %>%
  mutate(
    year = as.integer(sub("^(\\d{4})_.*$", "\\1", orig)),
    country = sub("^\\d{4}_(.*)$", "\\1", orig)
  ) %>%
  arrange(year, country)

ordered_cols <- col_info$orig
df_for_excel <- wide %>%
  select(all_of(base_cols), all_of(ordered_cols))

#  Write to Excel 
wb <- createWorkbook()
addWorksheet(wb, "Manifesto")
writeData(wb, "Manifesto", df_for_excel, startRow = 2)

#  Green style for 1's 
greenStyle <- createStyle(fgFill = "#C6EFCE")
num_mat <- df_for_excel %>% select(all_of(ordered_cols)) %>%
  mutate(across(everything(), ~as.integer(as.character(.))))
one_positions <- which(num_mat == 1, arr.ind = TRUE)
if (nrow(one_positions) > 0) {
  rows_cells <- one_positions[,1] + 2   # +2 because data starts at row 2
  cols_cells <- one_positions[,2] + length(base_cols)  # offset by base columns
  addStyle(wb, "Manifesto", style = greenStyle, rows = rows_cells, cols = cols_cells, gridExpand = FALSE)
}

#  Auto column width & save 
setColWidths(wb, "Manifesto", cols = 1:ncol(df_for_excel), widths = "auto")
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
  
  # Add labels if provided (without 'question')
  if(!is.null(labels_df)){
    summary_wide <- summary_wide %>%
      left_join(labels_df %>% select(-question), by="variable") %>%
      relocate(label, .after=variable)
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
  # Add labels if provided (without 'question')
  if(!is.null(labels_df)){
    summary_wide <- summary_wide %>%
      left_join(labels_df %>% select(-question), by="variable") %>%
      relocate(label, .after=variable)
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

# CLEA NEW FORMAT --------------
# --- Define countries ---
countries <- c("Czech Republic", "France", "Netherlands", "Hungary")

# --- Load CLEA ---
clea <- read_dta("H:/TWIN4DEM/clea_lc_20240419.dta")  

# --- Extract labels ---
labels_clea <- tibble(
  variable = names(clea),
  label = sapply(clea, function(x) attr(x, "label")),
  question = NA_character_
)

# --- Clean dataset ---
clea_clean <- clea %>%
  mutate(across(everything(), as.character)) %>%
  mutate(ctr_n = case_when(
    ctr_n %in% countries ~ ctr_n,
    TRUE ~ NA_character_
  )) %>%
  filter(!is.na(ctr_n)) %>%
  filter(as.numeric(yr) >= 2000, as.numeric(yr) <= 2024)

# --- Pivot longer and mark availability ---
vars <- setdiff(names(clea_clean), c("ctr_n","yr"))

clea_long <- clea_clean %>%
  pivot_longer(cols = all_of(vars), names_to = "variable", values_to = "value") %>%
  mutate(has_value = ifelse(!is.na(value) & value != "", 1L, 0L)) %>%
  group_by(variable, ctr_n, yr) %>%
  summarise(has_value = max(has_value), .groups = "drop")

# --- Complete all country-year-variable combinations ---
all_combinations <- expand_grid(
  variable = unique(clea_long$variable),
  ctr_n    = countries,
  yr       = 2000:2024
)

# Convert 'yr' in clea_long to integer to match all_combinations
clea_long <- clea_long %>% mutate(yr = as.integer(yr))

clea_long_complete <- all_combinations %>%
  left_join(clea_long, by = c("variable","ctr_n","yr")) %>%
  mutate(has_value = replace_na(has_value, 0L))

# --- Create country-year codes ---
clea_long_complete <- clea_long_complete %>%
  mutate(colname = paste0(
    ifelse(ctr_n=="Czech Republic","CZ",
           ifelse(ctr_n=="France","FR",
                  ifelse(ctr_n=="Hungary","HU",
                         ifelse(ctr_n=="Netherlands","NL",ctr_n)))),
    yr
  ))

# --- Pivot to binary matrix ---
clea_matrix <- clea_long_complete %>%
  select(variable, colname, has_value) %>%
  pivot_wider(
    names_from = colname,
    values_from = has_value,
    values_fill = 0L
  ) %>%
  arrange(variable)

# --- Add labels ---
clea_matrix <- clea_matrix %>%
  left_join(labels_clea, by="variable") %>%
  relocate(label, .after=variable)

# --- Save to Excel ---
wb <- createWorkbook()
addWorksheet(wb, "CLEA")
writeData(wb, "CLEA", clea_matrix, startRow = 2)

# --- Green style for 1s ---
base_cols <- c("variable","label")
ordered_cols <- setdiff(names(clea_matrix), base_cols)

greenStyle <- createStyle(fgFill = "#C6EFCE") 

num_mat <- clea_matrix %>% 
  select(all_of(ordered_cols)) %>% 
  mutate(across(everything(), as.integer))

one_positions <- which(num_mat == 1, arr.ind = TRUE) 
if (nrow(one_positions) > 0) { 
  rows_cells <- one_positions[,1] + 2  # offset for header
  cols_cells <- one_positions[,2] + length(base_cols)
  addStyle(
    wb, "CLEA", style = greenStyle, 
    rows = rows_cells, cols = cols_cells, 
    gridExpand = FALSE
  ) 
}

# --- Auto column width & save ---
setColWidths(wb, "CLEA", cols = 1:ncol(clea_matrix), widths = "auto")
saveWorkbook(wb, "H:/TWIN4DEM/clea_binary.xlsx", overwrite=TRUE)

#  DPI  --------------
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
      left_join(labels_df %>% select(-question), by="variable") %>%
      relocate(label, .after=variable)
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

paged_country_codes <- c("the Netherlands" , "France" , "Czechia", "Hungary")

paged <- read_dta("H:/TWIN4DEM/PAGED-basic-party-dataset.dta")  
labels_paged <- tibble(
  variable = names(paged),
  label = sapply(paged, function(x) attr(x, "label")),
  question = NA_character_
)

# --- Clean dataset ---
paged_clean <- paged %>%
  mutate(
    year_in = year(as.Date(date_in)),   
    across(everything(), as.character)
  ) %>%
  filter(country_name %in% paged_country_codes)



# --- Process PAGED dataset ---
summary_paged <- paged_clean %>%
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




# DPI new format --------------------
# --- Focus countries ---
countries <- c("Czech Republic", "France", "Netherlands", "Hungary")

# --- Load DPI dataset ---
dpi <- read_dta("H:/TWIN4DEM/dpi2015-stata13.dta")

# --- Extract labels ---
labels_dpi <- tibble(
  variable = names(dpi),
  label = sapply(dpi, function(x) attr(x, "label")),
  question = NA_character_
) %>% select(variable, label, question)  # ensure no 'year' column

# --- Clean dataset ---
dpi_clean <- dpi %>%
  mutate(across(everything(), as.character)) %>%
  mutate(countryname = case_when(
    countryname %in% c("Czech Rep.", "CSK") ~ "Czech Republic",
    countryname %in% c("France", "FRA") ~ "France",
    countryname %in% c("Netherlands", "NLD") ~ "Netherlands",
    countryname %in% c("Hungary", "HUN") ~ "Hungary",
    TRUE ~ countryname
  )) %>%
  filter(countryname %in% countries) %>%
  mutate(year = as.integer(year))  # ensure year is integer

# --- Select years and variables ---
years <- 2000:2024
vars <- setdiff(names(dpi_clean), c("countryname","year"))

# --- Build 0/1 coverage ---
dpi_covered <- dpi_clean %>%
  filter(year %in% years) %>%
  select(countryname, year, all_of(vars)) %>%
  mutate(across(all_of(vars), ~as.integer(!is.na(.))))  # 1 if non-NA, 0 if NA

# --- Pivot longer ---
dpi_long <- dpi_covered %>%
  pivot_longer(cols = all_of(vars), names_to = "variable", values_to = "coverage") %>%
  complete(variable, year = years, countryname = countries, fill = list(coverage = 0))

# --- Pivot wider: one column per country/year ---
dpi_wide <- dpi_long %>%
  unite("yr_country", year, countryname, sep = "_") %>%
  pivot_wider(names_from = yr_country, values_from = coverage)

# --- Add labels ---
dpi_wide <- dpi_wide %>%
  left_join(labels_dpi, by = "variable") %>%
  relocate(variable, label, question)

# --- Excel output ---
wb <- createWorkbook()
addWorksheet(wb, "DPI")

# --- Write data ---
writeData(wb, "DPI", dpi_wide, startRow = 2)

# --- Style 1's green ---
base_cols <- c("variable","label","question")
ordered_cols <- setdiff(names(dpi_wide), base_cols)
num_mat <- dpi_wide %>% select(all_of(ordered_cols)) %>% mutate(across(everything(), as.integer))

one_positions <- which(num_mat == 1, arr.ind = TRUE)
if(nrow(one_positions) > 0){
  rows_cells <- one_positions[,1] + 2
  cols_cells <- one_positions[,2] + length(base_cols)
  greenStyle <- createStyle(fgFill = "#C6EFCE")
  addStyle(wb, sheet = "DPI", style = greenStyle, rows = rows_cells, cols = cols_cells, gridExpand = FALSE)
}

# --- Auto width and save ---
setColWidths(wb, "DPI", cols = 1:ncol(dpi_wide), widths = "auto")
saveWorkbook(wb, "H:/TWIN4DEM/dpi_binary.xlsx", overwrite = TRUE)


# PAGED new format --------------------
library(dplyr)
library(tidyr)
library(readr)
library(lubridate)
library(openxlsx)
library(haven)

# --- Load PAGED dataset ---
paged <- read_dta("H:/TWIN4DEM/PAGED-basic-party-dataset.dta")  
labels_paged <- tibble(
  variable = names(paged),
  label = sapply(paged, function(x) attr(x, "label")),
  question = NA_character_
)

# --- Clean dataset ---
paged_clean <- paged %>%
  mutate(
    year_in = year(as.Date(date_in)),   
    across(everything(), as.character),
    country_name = case_when(
      country_name %in% c("the Netherlands") ~ "Netherlands",
      country_name == "Czechia" ~ "Czech Republic",
      TRUE ~ country_name
    )
  ) %>%
  filter(country_name %in% c("Czech Republic","France","Hungary","Netherlands"),
         year_in >= 2000, year_in <= 2024)

# --- Define variables ---
paged_vars <- setdiff(names(paged_clean), c("country_name","year_in","year_out","date_in","date_out"))


# Ensure year_in is numeric
paged_clean <- paged_clean %>%
  mutate(year_in = as.integer(year_in))  # convert to integer

# --- Pivot longer and create coverage ---
years <- 2000:2024
paged_long <- paged_clean %>%
  pivot_longer(cols = all_of(paged_vars), names_to = "variable", values_to = "value") %>%
  mutate(coverage = as.integer(!is.na(value) & value != "")) %>%
  group_by(variable, year_in, country_name) %>%
  summarise(coverage = as.integer(any(coverage == 1)), .groups = "drop") %>%
  complete(variable,
           year_in = years,
           country_name = c("Czech Republic","France","Hungary","Netherlands"),
           fill = list(coverage = 0))

# --- Pivot wider for Excel ---
paged_wide <- paged_long %>%
  unite(col = "year_country", year_in, country_name, sep = "_") %>%
  pivot_wider(names_from = year_country, values_from = coverage)

# --- Add labels ---
paged_wide <- paged_wide %>%
  left_join(labels_paged, by = "variable") %>%
  relocate(variable, label, question)

# --- Write to Excel with green 1s ---
wb <- createWorkbook()
addWorksheet(wb, "PAGED")
writeData(wb, "PAGED", paged_wide, startRow = 2)

base_cols <- c("variable","label","question")
yc_cols <- setdiff(names(paged_wide), base_cols)
num_mat <- paged_wide %>% select(all_of(yc_cols)) %>% mutate(across(everything(), as.integer))
one_positions <- which(num_mat == 1, arr.ind = TRUE)
greenStyle <- createStyle(fgFill = "#C6EFCE")
if (nrow(one_positions) > 0) {
  addStyle(wb, "PAGED", style = greenStyle,
           rows = one_positions[,1] + 2,   # account for startRow
           cols = one_positions[,2] + length(base_cols),
           gridExpand = FALSE)
}

setColWidths(wb, "PAGED", cols = 1:ncol(paged_wide), widths = "auto")
saveWorkbook(wb, "H:/TWIN4DEM/paged_binary.xlsx", overwrite = TRUE)



# --- TAP data ------------------
# Load the Excel file
tap_file <- ("H:/TWIN4DEM/API-2024.xlsx")

# List all sheet names
sheet_names <- excel_sheets(tap_file)
print(sheet_names)

# Read all sheets into a named list of data frames
tap_list <- lapply(sheet_names, function(s) {
  read_excel(tap_file, sheet = s) %>%
    mutate(sheet_name = s)  # optional: keep track of which sheet
})
names(tap_list) <- sheet_names

# Example: inspect the first few rows of each sheet
lapply(tap_list, head)


### --- Create workbook ---
wb <- createWorkbook()

# Add CHES sheet with note
addWorksheet(wb, "CHES")
writeData(wb, "CHES", "Chapel Hill expert survey Trend File 1999-2019", startRow = 1)
writeData(wb, "CHES", summary_ches, startRow = 3)
setColWidths(wb, "CHES", cols = 1:ncol(summary_ches), widths = "auto")

# Add CLEA sheet with note
addWorksheet(wb, "CLEA")
writeData(wb, "CLEA", "Constituency-Level Elections Archive", startRow = 1)
writeData(wb, "CLEA", "Note: Years shown correspond to election years (yr)", startRow = 2)
writeData(wb, "CLEA", summary_clea, startRow = 4)
setColWidths(wb, "CLEA", cols = 1:ncol(summary_clea), widths = "auto")

# Add DPI sheet with note
addWorksheet(wb, "DPI")
writeData(wb, "DPI", "The Database of Political Institutions 2015", startRow = 1)
writeData(wb, "DPI", summary_dpi, startRow = 3)
setColWidths(wb, "DPI", cols = 1:ncol(summary_dpi), widths = "auto")

# Add PAGED sheet with note
addWorksheet(wb, "PAGED")
writeData(wb, "PAGED", "Party Government in Europe Database (PAGED) – Basic dataset (Party dataset)", startRow = 1)
writeData(wb, "PAGED", "Note: Years shown correspond to cabinet formation years (year_in)", startRow = 2)
writeData(wb, "PAGED", summary_paged, startRow = 4)
setColWidths(wb, "PAGED", cols = 1:ncol(summary_paged), widths = "auto")

# Combine ParlGov summaries into one table
summary_parlgov <- bind_rows(
  summary_election %>% mutate(source = "Election"),
  summary_party    %>% mutate(source = "Party"),
  summary_cabinet  %>% mutate(source = "Cabinet")
) %>%
  relocate(source, .before = variable)  # put source as first column

# Add ParlGov sheet to the existing workbook
addWorksheet(wb, "ParlGov")
writeData(wb, "ParlGov", "Note: ParlGov data combined from Election, Party, Cabinet tables", startRow = 1)
writeData(wb, "ParlGov", summary_parlgov, startRow = 3)
setColWidths(wb, "ParlGov", cols = 1:ncol(summary_parlgov), widths = "auto")

# Save workbook
saveWorkbook(wb, "H:/TWIN4DEM/partyfacts_and_parlgov_temporal_coverage.xlsx", overwrite = TRUE)



# NEW format -------------------------------------------------------------
# Countries of interest
countries <- c("Czech Republic", "France", "Hungary", "Netherlands")

## CHES  NEW FORMAT ----- 
## CHES  NEW FORMAT ----- 
# Country mapping in CHES
ches_country_codes <- c(
  "Czech Republic" = 21,
  "France" = 6,
  "Hungary" = 23,
  "Netherlands" = 10
)

# --- Load CHES ---
ches <- read_dta("H:/TWIN4DEM/1999-2019_CHES_dataset_means(v3).dta")

ches_country_codes <- c(CZ = 21, FR = 6, HU = 23, NL = 10)

# Extract labels
labels_ches <- tibble(
  variable = names(ches),
  label = sapply(ches, function(x) attr(x, "label")),
  question = NA_character_
)

# --- Convert all relevant columns to character ---
ches_clean <- ches %>%
  mutate(
    country = case_when(
      country == 21 ~ "Czech Republic",
      country == 6  ~ "France",
      country == 23 ~ "Hungary",
      country == 10 ~ "Netherlands",
      TRUE ~ NA_character_
    )
  ) %>%
  filter(!is.na(country)) %>%
  filter(year >= 2000, year <= 2024) %>%
  mutate(across(-c(country, year), as.character))  # ensure consistent type

# --- Pivot longer and mark availability ---
vars <- setdiff(names(ches_clean), c("country","year"))
# --- Compute binary availability per variable-country-year ---
ches_long_clean <- ches_clean %>%
  pivot_longer(cols = all_of(vars), names_to = "variable", values_to = "value") %>%
  mutate(has_value = ifelse(!is.na(value) & value != "", 1L, 0L)) %>%
  group_by(variable, country, year) %>%
  summarise(has_value = max(has_value), .groups = "drop")  # <- ensures one row per variable-country-year

# --- Complete all country-year-variable combinations ---
all_combinations <- expand_grid(
  variable = unique(ches_long_clean$variable),
  country  = c("Czech Republic","France","Hungary","Netherlands"),
  year     = 2000:2024
)

ches_long_complete <- all_combinations %>%
  left_join(ches_long_clean, by = c("variable","country","year")) %>%
  mutate(has_value = replace_na(has_value, 0L))   # fill missing years with 0

# --- Create standard country-year column codes ---
ches_long_complete <- ches_long_complete %>%
  mutate(colname = paste0(
    ifelse(country=="Czech Republic","CZ",
           ifelse(country=="France","FR",
                  ifelse(country=="Hungary","HU",
                         ifelse(country=="Netherlands","NL",country)))),
    year
  ))

# --- Pivot to binary matrix with all years ---
ches_matrix <- ches_long_complete %>%
  select(variable, colname, has_value) %>%
  pivot_wider(names_from = colname, values_from = has_value) %>%
  arrange(variable)

# --- Pivot to binary matrix with all years ---
ches_matrix <- ches_long_complete %>%
  select(variable, colname, has_value) %>%
  pivot_wider(
    names_from = colname,
    values_from = has_value,
    values_fill = 0L  # fill missing years with 0
  ) %>%
  arrange(variable)

# --- Add labels ---
labels_ches <- tibble(
  variable = names(ches),
  label = sapply(ches, function(x) attr(x, "label"))
)
ches_matrix <- ches_matrix %>%
  left_join(labels_ches, by="variable") %>%
  relocate(label, .after=variable)

# --- Save to Excel ---
wb <- createWorkbook()
addWorksheet(wb, "CHES")
writeData(wb, "CHES", ches_matrix, startRow = 2)

# --- Green style for 1s ---
# --- Define base columns and numeric columns ---
base_cols <- c("variable","label")
ordered_cols <- setdiff(names(ches_matrix), base_cols)

# --- Green style for 1s ---
greenStyle <- createStyle(fgFill = "#C6EFCE") 

num_mat <- ches_matrix %>% 
  select(all_of(ordered_cols)) %>% 
  mutate(across(everything(), as.integer)) 

one_positions <- which(num_mat == 1, arr.ind = TRUE) 
if (nrow(one_positions) > 0) { 
  rows_cells <- one_positions[,1] + 2  # offset for header rows
  cols_cells <- one_positions[,2] + length(base_cols)
  addStyle(
    wb, "CHES", style = greenStyle, 
    rows = rows_cells, cols = cols_cells, 
    gridExpand = FALSE
  ) 
}


# --- Auto column width & save ---
setColWidths(wb, "CHES", cols = 1:ncol(ches_matrix), widths = "auto")
saveWorkbook(wb, "H:/TWIN4DEM/ches_binary.xlsx", overwrite=TRUE)

  
  
#   ##--- merging into one overview ----------------
# # ----------------------------
# # Load PartyFacts external IDs
# # ----------------------------
# pf_links <- read_csv("https://partyfacts.herokuapp.com/download/external-parties-csv/",
#                      na = "", guess_max = 50000)
#   
#   
#   # Function to harmonize country names
#   harmonize_country <- function(df, country_var) {
#     df %>%
#       mutate(
#         {{country_var}} := case_when(
#           {{country_var}} %in% c("Czechia", "Czech Republic") ~ "Czech Republic",
#           {{country_var}} %in% c("the Netherlands", "Netherlands") ~ "Netherlands",
#           {{country_var}} == "France" ~ "France",
#           {{country_var}} == "Hungary" ~ "Hungary",
#           TRUE ~ as.character({{country_var}})
#         )
#       )
#   }
# 
# # ----------------------------
# # CHES - party-level
# # ----------------------------
# ches <- read_dta("H:/TWIN4DEM/1999-2019_CHES_dataset_means(v3).dta")
# 
# # load CHES-PartyFacts mapping
# ches_map <- pf_links %>%
#   filter(dataset_key == "ches") %>%
#   select(dataset_party_id, partyfacts_id) %>%
#   mutate(dataset_party_id = as.character(dataset_party_id))
# 
# ches_linked <- ches %>%
#   mutate(dataset_party_id = as.character(party_id)) %>%  
#   left_join(ches_map, by = "dataset_party_id")
# 
# # ----------------------------
# # CLEA - party-level
# # ----------------------------
# clea <- read_dta("H:/TWIN4DEM/clea_lc_20240419.dta")
# 
# clea_map <- pf_links %>%
#   filter(dataset_key == "clea") %>%
#   select(dataset_party_id, partyfacts_id)  %>%
#   mutate(dataset_party_id = as.character(dataset_party_id))
# 
# clea_prepped <- clea %>%
#   mutate(dataset_party_id = ctr * 1e6 + pty) %>%   
#   mutate(dataset_party_id = as.character(dataset_party_id))
# 
# clea_linked <- clea_prepped %>%
#   left_join(clea_map, by = "dataset_party_id")
# 
# # ----------------------------
# # DPI - country-level
# # ----------------------------
# dpi <- read_dta("H:/TWIN4DEM/dpi2015-stata13.dta")
# 
# dpi_map1 <- read_csv("H:/TWIN4DEM/dpi.csv") %>%
#   select(country, countryname, party, first, last, party_id) %>%
#   mutate(party_id = as.character(party_id))
# dpi_pf <- pf_links %>%
#   filter(dataset_key == "dpi") %>%
#   select(dataset_party_id, partyfacts_id) %>%
#   mutate(dataset_party_id = as.character(dataset_party_id))
# dpi_map_full <- dpi_map1 %>%
#   left_join(dpi_pf, by = c("party_id" = "dataset_party_id"))
# 
# #---
# 
# dpi_prepped <- dpi %>%
#   mutate(dataset_party_id = as.character(prtyin))  # can also do gov1party etc.
# 
# dpi_linked <- dpi_prepped %>%
#   left_join(dpi_map_full, by = c("dataset_party_id" = "party_id")) %>%
#   mutate(country_std = countryname.y) %>%
#   select(-countryname.x, -countryname.y)   # cleanup
# 
# 
# 
# 
# # ----------------------------
# # PAGED
# # ----------------------------
# paged_linked <- read_dta("H:/TWIN4DEM/PAGED-basic-party-dataset.dta")
# 
# paged_map <- pf_links %>%
#   filter(dataset_key == "paged") %>%
#   select(dataset_party_id, partyfacts_id)
# 
# 
# 
# # ----------------------------
# # Combine all datasets
# # ----------------------------
# 
# 
# # Standardize country variables first
# ches_clean <- ches_linked %>% mutate(country_std = country)
# clea_clean <- clea_linked %>% mutate(country_std = ctr_n)
# dpi_clean  <- dpi_linked  
# paged_clean<- paged_linked %>% mutate(country_std = country_name)
# 
# 
# clea_clean  <- harmonize_country(clea_clean, country_std)
# dpi_clean   <- harmonize_country(dpi_clean, country_std)
# paged_clean <- harmonize_country(paged_clean, country_std)
# 
# # Filter CHES for the 4 countries and convert to character names
# ches_sel <- ches_clean %>%
#   filter(country %in% c(21, 6, 23, 10)) %>%   # numeric codes
#   mutate(
#     country_std = case_when(
#       country == 21 ~ "Czech Republic",
#       country == 6  ~ "France",
#       country == 23 ~ "Hungary",
#       country == 10 ~ "Netherlands"
#     )
#   ) %>%
#   mutate(country_std = as.character(country_std))
# 
# # Select relevant columns and rename year variables consistently
# ches_sel  <- ches_sel %>% rename(year = year)
# clea_sel  <- clea_clean %>% rename(year = yr)
# dpi_sel   <- dpi_clean  %>% rename(year = year)
# paged_sel <- paged_clean %>%
#   mutate(
#     year = year(date_in)  # extracts only the year from a Date
#   )
# 
# 
# dpi_sel <- dpi_sel %>%
#   mutate(
#     country_std = case_when(
#       ifs == "CZE" ~ "Czech Republic",
#       ifs == "FRA" ~ "France",
#       ifs == "HUN" ~ "Hungary",
#       ifs == "NLD" ~ "Netherlands",
#       TRUE ~ NA_character_
#     )
#   ) %>%
#   rename(year = year)   # keep year consistent
# 
# 
# # Merge datasets using left_join on partyfacts_id
# integrated <- paged_sel %>%
#   left_join(clea_sel, by = c("partyfacts_id", "country_std", "year")) %>%
#   left_join(ches_sel,by = c("partyfacts_id", "country_std", "year")) %>%
#   left_join(dpi_sel, by = c("country_std", "year"))  # note: not partyfacts_id
# 
# 
# countries_of_interest <- c("Czech Republic", "France", "Netherlands", "Hungary")
# 
# integrated_filtered <- integrated %>%
#   filter(country_std %in% countries_of_interest,
#          year >= 2000,
#          year <= 2024)
# 
# # Identify variables (exclude key columns)
# key_vars <- setdiff(names(integrated_filtered), c("partyfacts_id", "country_std", "year"))
# 
# # Convert all variables to character
# integrated_filtered <- integrated_filtered %>%
#   mutate(across(all_of(key_vars), as.character))
# 
# # Then pivot longer
# integrated_long <- integrated_filtered %>%
#   pivot_longer(
#     cols = all_of(key_vars),
#     names_to = "variable",
#     values_to = "value"
#   ) %>%
#   filter(!is.na(value) & value != "")
# 
# 
# temporal_coverage <- integrated_long %>%
#   group_by(variable, country_std) %>%
#   summarise(years_available = paste(sort(unique(year)), collapse = ", "), .groups="drop") %>%
#   pivot_wider(names_from = country_std, values_from = years_available) %>%
#   arrange(variable)
# 
# # ----------------------------
# # Write to Excel 
# # ----------------------------
# wb <- createWorkbook()
# addWorksheet(wb, "Temporal Coverage")
# writeData(wb, "Temporal Coverage", temporal_coverage)
# setColWidths(wb, "Temporal Coverage", cols = 1:ncol(temporal_coverage), widths = "auto")
# saveWorkbook(wb, "H:/TWIN4DEM/integrated_partyfacts_dataset.xlsx", overwrite = TRUE)
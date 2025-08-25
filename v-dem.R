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
#### V-DEM data: # show temporal coverage

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
variable_groups <- list(
  Executive = list(prefix = "^v2ex(?!l)", sheet = "Executive"),  # exclude "v2exl"
  Regime = list(prefix = "^v2reg", sheet = "Regime"),
  Legislature = list(prefix = "^v2lg", sheet = "Legislature"),
  Judiciary = list(prefix = "^v2ju", sheet = "Judiciary"),
  `Democracy Indices (V-Dem)` = list(prefix = "^v2x_", sheet = "Democracy Indices (V-Dem)")
)
# Function process_group
process_group <- function(prefix, sheet_name, data, labels_df, country_lookup) {
  
  # Select relevant variables
  relevant_vars <- names(data)[grepl(prefix, names(data), perl = TRUE)]
  
  # Filter and clean
  filtered <- data %>%
    filter(country_id %in% country_lookup$country_id, year >= 2000, year <= 2024) %>%
    select(country_id, country_name, year, all_of(relevant_vars)) %>%
    left_join(country_lookup, by = "country_id") %>%
    select(country, year, all_of(relevant_vars)) %>%
    mutate(across(all_of(relevant_vars), as.character))
  
  # Pivot longer
  long <- filtered %>%
    pivot_longer(cols = all_of(relevant_vars), names_to = "variable", values_to = "value") %>%
    filter(!is.na(value) & value != "") %>%
    arrange(country, variable, year)
  
  # Summarize
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

# Load original V-Dem data and labels
library(vdemdata)
data <- vdemdata::vdem

# var_meta <- var_info(names(vdemdata::vdem))
# str(var_meta)

labels_df <- var_info(names(vdemdata::vdem)) %>%
  tibble::as_tibble() %>%
  transmute(
    variable = tag,
    label = as.character(name), 
    question
  )
country_lookup <- data.frame(
  country_id = c(157, 76, 91, 210),
  country = c("Czech Republic", "France", "Netherlands", "Hungary")
)

   

# Apply function
summary_tables <- purrr::imap(
  variable_groups,
  ~ process_group(.x$prefix, .x$sheet, data, labels_df, country_lookup)
)

# Create workbook
wb <- createWorkbook()

for (sheet_name in names(summary_tables)) {
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet = sheet_name, summary_tables[[sheet_name]])
  setColWidths(wb, sheet = sheet_name, cols = 1:ncol(summary_tables[[sheet_name]]), widths = "auto")
}

# Save workbook into excel
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

# Run processing
summary_vparty <- process_vparty(df_vparty, labels_df_vparty, country_lookup_vparty)

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
## TO BE UPDATED! --------------------------------------------------------------
# download and read Party Facts mapping table
file_name <- 'partyfacts-mapping.csv'
if( ! file_name %in% list.files())
  url <- 'https://partyfacts.herokuapp.com/download/external-parties-csv/'
download.file(url, file_name)
partyfacts <- read.csv(file_name, as.is=TRUE)  # maybe conversion of character encoding

# -----------------------------
# Define countries of interest
# -----------------------------
# Note: use ISO codes as in PartyFacts ("CZE", "FRA", etc.)
countries <- c("CZE", "FRA", "NLD", "HUN")

partyfacts_filtered <- partyfacts %>% filter(country %in% countries)

# -----------------------------
# Identify variables to summarize
# -----------------------------
# Keep 'share_year' as temporal info
exclude_cols <- c("partyfacts_id", "name", "country", "dataset_key", "share_year")
relevant_vars <- setdiff(names(partyfacts_filtered), exclude_cols)

# -----------------------------
# Pivot longer
# -----------------------------
long <- partyfacts_filtered %>%
  pivot_longer(
    cols = all_of(relevant_vars),
    names_to = "variable",
    values_to = "value",
    values_transform = list(value = as.character)
  ) %>%
  filter(!is.na(value) & value != "" & !is.na(share_year) & share_year >= 2000 & share_year <= 2024) %>%
  arrange(country, variable)

# -----------------------------
# Summarize temporal coverage by variable and country
# -----------------------------
summary <- long %>%
  group_by(country, variable) %>%
  summarise(
    years_available = paste(sort(unique(share_year)), collapse = ", "),
    .groups = "drop"
  )

# -----------------------------
# Pivot wider so each country is a column
# -----------------------------
summary_wide <- summary %>%
  pivot_wider(names_from = country, values_from = years_available) %>%
  arrange(variable)

# -----------------------------
# Optional: add labels if you have a small lookup
# If not, just duplicate variable names
summary_wide <- summary_wide %>%
  mutate(label = variable, .before = 2)

# -----------------------------
# Write to Excel
# -----------------------------
wb <- createWorkbook()
addWorksheet(wb, "PartyFacts")
writeData(wb, sheet = "PartyFacts", summary_wide)
setColWidths(wb, sheet = "PartyFacts", cols = 1:ncol(summary_wide), widths = "auto")
saveWorkbook(wb, "H:/TWIN4DEM/partyfacts_coverage.xlsx", overwrite = TRUE)


###
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
dataset_1 <- partyfacts |> filter(dataset_key == "manifesto")
dataset_2 <- partyfacts |> filter(dataset_key == "parlgov")
link_table <-
  dataset_1 |>
  inner_join(dataset_2, by = c("partyfacts_id" = "partyfacts_id"))

# write results into file with dataset names in file name
file_out <- "partyfacts-linked.csv"
write_csv(link_table, file_out)
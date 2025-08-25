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
  
  # Use *all* variables
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
      values_transform = list(value = as.character)   # 🔑 here
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

# Write to Excel (single sheet)
wb_vparty <- createWorkbook()
addWorksheet(wb_vparty, "V-Party")
writeData(wb_vparty, sheet = "V-Party", summary_vparty)
setColWidths(wb_vparty, sheet = "V-Party", cols = 1:ncol(summary_vparty), widths = "auto")

# Save file
saveWorkbook(wb_vparty, "H:/TWIN4DEM/v-party.xlsx", overwrite = TRUE)

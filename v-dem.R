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
#### V-DEM data: # show temporal coverage
rm(list=ls())


#### trial to automate
variable_groups <- list(
  Executive = list(prefix = "^v2ex(?!l)", sheet = "Executive"),  # exclude "v2exl"
  Regime = list(prefix = "^v2reg", sheet = "Regime"),
  Legislature = list(prefix = "^v2lg", sheet = "Legislature"),
  Judiciary = list(prefix = "^v2ju", sheet = "Judiciary"),
  `Democracy Indices (V-Dem)` = list(prefix = "^v2x_", sheet = "Democracy Indices (V-Dem)")
)
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
    select(country, variable, label, years_available)
  
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
  
  # Merge with question text
  question_text <- read_excel("H:/TWIN4DEM/v-dem question texts.xlsx", sheet = sheet_name)
  
  summary_wide <- summary_wide %>%
    left_join(question_text %>% select(variable, question), by = "variable") %>%
    relocate(question, .after = label)
  
  return(summary_wide)
}

# Load original V-Dem data and labels
data <- read_dta("H:/TWIN4DEM/V-Dem-CY-Full+Others-v15.dta")
labels_df <- data.frame(
  variable = names(var_label(data)),
  label = unlist(var_label(data)),
  stringsAsFactors = FALSE
)
country_lookup <- data.frame(
  country_id = c(157, 76, 91, 210),
  country = c("Czech Republic", "France", "Netherlands", "Hungary")
)

# Process all groups
summary_tables <- purrr::imap(
  variable_groups,
  ~ process_group(.x$prefix, .x$sheet, data, labels_df, country_lookup)
)

wb <- createWorkbook()

for (sheet_name in names(summary_tables)) {
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet = sheet_name, summary_tables[[sheet_name]])
  setColWidths(wb, sheet = sheet_name, cols = 1:ncol(summary_tables[[sheet_name]]), widths = "auto")
}

saveWorkbook(wb, "H:/TWIN4DEM/v-dem.xlsx", overwrite = TRUE)


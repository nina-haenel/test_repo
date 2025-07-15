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

library(usethis)
library("tidyr")
library("tidyverse")
library("dplyr")
library("ggplot2")
library("haven")
library("foreign")
library("here")
library("xlsx")
library("usethis")
library("haven")
library("labelled")
library("writexl")
#### V-DEM data: # show temporal coverage
rm(list=ls())
`V-Dem-CY-Full+Others-v15` <- read_dta("H:/TWIN4DEM/V-Dem-CY-Full+Others-v15.dta")

# attr(`V-Dem-CY-Full+Others-v15`$v2exbribe, "label")

# Filtering CZ, FR, NL, and HU data only
filtered_vdem_data <- `V-Dem-CY-Full+Others-v15` %>%
  filter(country_id %in% c(157, 76, 91, 210))

# relevant variables: Executive
relevant_vars_ex <- names(filtered_vdem_data)[
  grepl("^v2ex", names(filtered_vdem_data)) &       # starts with "v2ex"
    !grepl("^v2exl", names(filtered_vdem_data))       # does NOT start with "v2exl"
]

# relevant variables: Regime
relevant_vars_reg <- names(filtered_vdem_data)[startsWith(names(filtered_vdem_data), "v2reg")]

# relevant variables: Legislature
relevant_vars_leg <- names(filtered_vdem_data)[startsWith(names(filtered_vdem_data), "v2lg")]

# relevant variables: Judiciary
relevant_vars_ju <- names(filtered_vdem_data)[startsWith(names(filtered_vdem_data), "v2ju")]


# Country lookup
country_lookup <- data.frame(
  country_id = c(157, 76, 91, 210),
  country = c("Czech Republic", "France", "Netherlands", "Hungary")
)

# Filter 
# EX
filtered_data_ex <- `V-Dem-CY-Full+Others-v15` %>%
  filter(country_id %in% country_lookup$country_id, year >= 2000, year <= 2024) %>%
  select(country_id, country_name, year, all_of(relevant_vars_ex)) %>%
  left_join(country_lookup, by = "country_id") %>%
  select(country, year, all_of(relevant_vars_ex))

# Regime
filtered_data_reg <- `V-Dem-CY-Full+Others-v15` %>%
  filter(country_id %in% country_lookup$country_id, year >= 2000, year <= 2024) %>%
  select(country_id, country_name, year, all_of(relevant_vars_reg)) %>%
  left_join(country_lookup, by = "country_id") %>%
  select(country, year, all_of(relevant_vars_reg))

# Legislature
filtered_data_leg <- `V-Dem-CY-Full+Others-v15` %>%
  filter(country_id %in% country_lookup$country_id, year >= 2000, year <= 2024) %>%
  select(country_id, country_name, year, all_of(relevant_vars_leg)) %>%
  left_join(country_lookup, by = "country_id") %>%
  select(country, year, all_of(relevant_vars_leg))

# judiciary
filtered_data_ju <- `V-Dem-CY-Full+Others-v15` %>%
  filter(country_id %in% country_lookup$country_id, year >= 2000, year <= 2024) %>%
  select(country_id, country_name, year, all_of(relevant_vars_ju)) %>%
  left_join(country_lookup, by = "country_id") %>%
  select(country, year, all_of(relevant_vars_ju))

# Clean column types for relevant variables
# Ex
cleaned_data_ex <- filtered_data_ex %>%
  mutate(across(all_of(relevant_vars_ex), as.character)) 
# Reg
cleaned_data_reg <- filtered_data_reg %>%
  mutate(across(all_of(relevant_vars_reg), as.character)) 
# Leg
cleaned_data_leg <- filtered_data_leg %>%
  mutate(across(all_of(relevant_vars_leg), as.character)) 

# Ju
cleaned_data_ju <- filtered_data_ju %>%
  mutate(across(all_of(relevant_vars_ju), as.character)) 

# Pivot longer
# Ex
long_data_ex <- cleaned_data_ex %>%
  pivot_longer(
    cols = all_of(relevant_vars_ex),
    names_to = "variable",
    values_to = "value"
  ) %>%
  arrange(country, variable, year)

# Reg
long_data_reg <- cleaned_data_reg %>%
  pivot_longer(
    cols = all_of(relevant_vars_reg),
    names_to = "variable",
    values_to = "value"
  ) %>%
  arrange(country, variable, year)

# Leg
long_data_leg <- cleaned_data_leg %>%
  pivot_longer(
    cols = all_of(relevant_vars_leg),
    names_to = "variable",
    values_to = "value"
  ) %>%
  arrange(country, variable, year)

# Ju
long_data_ju <- cleaned_data_ju %>%
  pivot_longer(
    cols = all_of(relevant_vars_ju),
    names_to = "variable",
    values_to = "value"
  ) %>%
  arrange(country, variable, year)

# Summary Table, variables as rows
# Ex
summary_table_ex <- long_data_ex %>%
  filter(!is.na(value) & value != "") %>%  # Keep only non-missing values
  group_by(country, variable) %>%
  summarise(
    years_available = paste(sort(unique(year)), collapse = ", "),
    .groups = "drop"
  ) %>%
  arrange(country, variable)

# Reg
summary_table_reg <- long_data_reg %>%
  filter(!is.na(value) & value != "") %>%  # Keep only non-missing values
  group_by(country, variable) %>%
  summarise(
    years_available = paste(sort(unique(year)), collapse = ", "),
    .groups = "drop"
  ) %>%
  arrange(country, variable)

# Leg
summary_table_leg <- long_data_leg %>%
  filter(!is.na(value) & value != "") %>%  # Keep only non-missing values
  group_by(country, variable) %>%
  summarise(
    years_available = paste(sort(unique(year)), collapse = ", "),
    .groups = "drop"
  ) %>%
  arrange(country, variable)

# Ju
summary_table_ju <- long_data_ju %>%
  filter(!is.na(value) & value != "") %>%  # Keep only non-missing values
  group_by(country, variable) %>%
  summarise(
    years_available = paste(sort(unique(year)), collapse = ", "),
    .groups = "drop"
  ) %>%
  arrange(country, variable)

# Adding labels
var_labels <- var_label(`V-Dem-CY-Full+Others-v15`)  # this keeps names and labels
labels_df <- data.frame(
  variable = names(var_labels),
  label = unlist(var_labels),
  stringsAsFactors = FALSE
)

# Ex
summary_table_ex <- summary_table_ex %>%
  left_join(labels_df, by = "variable") %>%
  select(country, variable, label, years_available)

# Reg
summary_table_reg <- summary_table_reg %>%
  left_join(labels_df, by = "variable") %>%
  select(country, variable, label, years_available)

# Reg
summary_table_leg <- summary_table_leg %>%
  left_join(labels_df, by = "variable") %>%
  select(country, variable, label, years_available)

# Ju
summary_table_ju <- summary_table_ju %>%
  left_join(labels_df, by = "variable") %>%
  select(country, variable, label, years_available)

#summary_table_wide <- summary_table %>%
#  select(country, variable, label, years_available) %>%
#  pivot_wider(
#    names_from = country,
#    values_from = years_available
#  ) %>%
#  arrange(variable)

# Ex
summary_table_wide_ex <- summary_table_ex %>%
  select(country, variable, label, years_available) %>%
  pivot_wider(
    names_from = country,
    values_from = years_available
  ) %>%
  arrange(variable) %>%
  mutate(across(
    c("Czech Republic", "France", "Netherlands", "Hungary"),
    ~ ifelse(
      . == paste(2000:2024, collapse = ", "),
      "2000-2024",
      .
    )
  ))  %>%
  mutate(across(
    c("Czech Republic", "France", "Netherlands", "Hungary"),
    ~ ifelse(
      . == paste(2000:2023, collapse = ", "),
      "2000-2023",
      .
    )
  ))


# Reg
summary_table_wide_reg <- summary_table_reg %>%
  select(country, variable, label, years_available) %>%
  pivot_wider(
    names_from = country,
    values_from = years_available
  ) %>%
  arrange(variable) %>%
  mutate(across(
    c("Czech Republic", "France", "Netherlands", "Hungary"),
    ~ ifelse(
      . == paste(2000:2024, collapse = ", "),
      "2000-2024",
      .
    )
  ))  %>%
  mutate(across(
    c("Czech Republic", "France", "Netherlands", "Hungary"),
    ~ ifelse(
      . == paste(2000:2023, collapse = ", "),
      "2000-2023",
      .
    )
  ))

# Leg
summary_table_wide_leg <- summary_table_leg %>%
  select(country, variable, label, years_available) %>%
  pivot_wider(
    names_from = country,
    values_from = years_available
  ) %>%
  arrange(variable) %>%
  mutate(across(
    c("Czech Republic", "France", "Netherlands", "Hungary"),
    ~ ifelse(
      . == paste(2000:2024, collapse = ", "),
      "2000-2024",
      .
    )
  ))  %>%
  mutate(across(
    c("Czech Republic", "France", "Netherlands", "Hungary"),
    ~ ifelse(
      . == paste(2000:2023, collapse = ", "),
      "2000-2023",
      .
    )
  ))

# Ju
summary_table_wide_ju <- summary_table_ju %>%
  select(country, variable, label, years_available) %>%
  pivot_wider(
    names_from = country,
    values_from = years_available
  ) %>%
  arrange(variable) %>%
  mutate(across(
    c("Czech Republic", "France", "Netherlands", "Hungary"),
    ~ ifelse(
      . == paste(2000:2024, collapse = ", "),
      "2000-2024",
      .
    )
  ))  %>%
  mutate(across(
    c("Czech Republic", "France", "Netherlands", "Hungary"),
    ~ ifelse(
      . == paste(2000:2023, collapse = ", "),
      "2000-2023",
      .
    )
  ))






# Writing into Excel Table
write_xlsx(
  list(Executive = summary_table_wide_ex, 
       Regime = summary_table_wide_reg,
       Legislature = summary_table_wide_leg,
       Judiciary = summary_table_wide_ju),
  path = "H:/TWIN4DEM/v-dem.xlsx"
)

## ich probiere hier unten mal was aus.
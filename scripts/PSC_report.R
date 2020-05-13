library(tidyverse)
library(readxl)
library(purrr)
library(lubridate)

# Cache CSV snapshot using readxl of raw datatables from PSC_JIRA and PSC_O2D report 
# workbooks. 


# JIRA REPORT READ AND CONVERSION TO CSV
# note: requires additional code to deal with absence of conventional Headers in spreadsheet

JIRA_assigned <- read_excel("data/PSC_JIRA_Mar20.xlsm", 
                            sheet = "Created",
                            col_names = c("class", "issue", "key_ident", "issue_id", "summary", "status", "resolution", "created", "resolved", "reporter", "assignee", "misc"),
                            col_type = c("text", "text", "numeric", "numeric", "text", "text", "text", "date", "date", "text", "text", "text")) %>% 
  write_csv("data/JIRA_assigned.csv")

JIRA_completed <- read_excel("data/PSC_JIRA_Mar20.xlsm", 
                             sheet = "Completed",
                             col_names = c("class", "issue", "key_ident", "issue_id", "summary", "status", "resolution", "created", "resolved", "reporter", "assignee", "misc", "misc"),
                             col_type = c("text", "text", "numeric", "numeric", "text", "text", "text", "date", "date", "text", "text", "text", "text")) %>% 
  write_csv("data/JIRA_completed.csv")


# O2D REPORT READ AND CONVERSION TO CSV

O2D_assigned <- read_excel("data/PSC_O2D_Mar20.xlsm",  #Q: possible to find/replace filenames, i.e. latest version?
                          sheet = "Assigned", 
                          col_names = TRUE) %>%  
  write_csv("data/O2D_assigned.csv")

O2D_completed <- read_excel("data/PSC_O2D_Mar20.xlsm",
                           sheet = "Completed", 
                           col_names = TRUE) %>%  
  write_csv("data/O2D_completed.csv")

O2D_prev_month <- read_excel("data/PSC_O2D_Mar20.xlsm",
                             sheet = "Open from previous mths") %>% 
  write_csv("data/O2D_prev_mths.csv")

## Include additional code to dispose sheets into .csv files with appropriate naming conventions.

# IMPORT CSV EXPORTS AS NEW VARIABLES

# JIRA REPORT DATAFRAMES
JIRA_open <- read_csv('data/JIRA_assigned.csv')             
JIRA_completed <- read_csv('data/JIRA_completed.csv')

# O2D REPORT DATAFRAMES
O2D_assigned <- read_csv('data/O2D_assigned.csv')
O2D_completed <- read_csv('data/O2D_completed.csv')
O2D_prev_mths <- read_csv('data/O2D_prev_mths.csv')

# JIRA & O2D DATA TIDY & TRANSFORM

JIRA_open <- 
  slice(JIRA_open, -1:-3, preserve = FALSE) %>% 
  select(-3) %>% 
  na.omit() # remove all rows containing 'NA'

JIRA_completed <- 
  slice(JIRA_completed, -1:-3, preserve = FALSE) %>% 
  select(-3, -10:-13) %>% 
  na.omit() # remove all rows containing 'NA'

JIRA_open_count <- group_by(JIRA_open, class) %>% 
  summarise(count = n())

JIRA_completed_count <- group_by(JIRA_completed, class) %>% 
  summarise(count = n())

## O2D DATA TIDY & TRANSFORM

O2D_assigned_count <- 
  rename(O2D_assigned, class = `Task Type`) %>% 
  group_by(class) %>% 
  summarise(count = n())

O2D_completed_count <- 
  rename(O2D_completed, class = `Task Type`) %>% 
  group_by(class) %>% 
  summarise(count = n())

O2D_prev_mths_count <- 
  rename(O2D_prev_mths, class = `Task Type`) %>% 
  group_by(class) %>% 
  summarise(count = n())

JIRA_open_count
JIRA_completed_count
O2D_assigned_count
O2D_completed_count
O2D_prev_mths_count

# NEXT STAGE: JOIN TOGETHER CLASS FOR OPEN AND COMPLETED

JIRA_open_close <- filter(JIRA_open_count, class == "Close Project")
JIRA_completed_close <- filter(JIRA_completed_count, class == "Close Project")



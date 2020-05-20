library(tidyverse)
library(readxl)
library(purrr)
library(lubridate)

## SET MONTH OF REPORT HERE ##

current_month <- ymd("2020-03-31") %>% 
  month(label = TRUE, abbr = TRUE, locale = Sys.getlocale("LC_COLLATE"))

##############################

# Cache CSV snapshot using readxl of raw datatables from PSC_JIRA and PSC_O2D report 
# workbooks. 


# JIRA REPORT READ AND CONVERSION TO CSV
# note: requires additional code to deal with absence of conventional Headers in spreadsheet

JIRA_assigned_import <- read_excel("data/PSC_JIRA_Mar20.xlsm", 
                            sheet = "Created",
                            col_names = c("type", "issue", "key_ident", "issue_id", "summary", "stat", "resolution", "created", "resolved", "reporter", "assignee", "misc"),
                            col_type = c("text", "text", "numeric", "numeric", "text", "text", "text", "date", "date", "text", "text", "text")) %>% 
  write_csv("data/JIRA_assigned.csv")

JIRA_completed_import <- read_excel("data/PSC_JIRA_Mar20.xlsm", 
                             sheet = "Completed",
                             col_names = c("type", "issue", "key_ident", "issue_id", "summary", "stat", "resolution", "created", "resolved", "reporter", "assignee", "misc", "misc"),
                             col_type = c("text", "text", "numeric", "numeric", "text", "text", "text", "date", "date", "text", "text", "text", "text")) %>% 
  write_csv("data/JIRA_completed.csv")


# O2D REPORT READ AND CONVERSION TO CSV

O2D_assigned_import <- read_excel("data/PSC_O2D_Mar20.xlsm",  #Q: possible to find/replace filenames, i.e. latest version?
                          sheet = "Assigned", 
                          col_names = TRUE) %>%  
  write_csv("data/O2D_assigned.csv")

O2D_completed_import <- read_excel("data/PSC_O2D_Mar20.xlsm",
                           sheet = "Completed", 
                           col_names = TRUE) %>%  
  write_csv("data/O2D_completed.csv")

O2D_prev_month_import <- read_excel("data/PSC_O2D_Mar20.xlsm",
                             sheet = "Open from previous mths") %>% 
  write_csv("data/O2D_prev_mths.csv")

# IMPORT CSV EXPORTS AS NEW VARIABLES

# JIRA REPORT DATAFRAMES
JIRA_open_init1 <- read_csv('data/JIRA_assigned.csv')             
JIRA_completed_init1 <- read_csv('data/JIRA_completed.csv')

# O2D REPORT DATAFRAMES
O2D_assigned_init1 <- read_csv('data/O2D_assigned.csv')
O2D_completed_init1 <- read_csv('data/O2D_completed.csv')
O2D_prev_mths_init1 <- read_csv('data/O2D_prev_mths.csv')

# JIRA & O2D DATA TIDY & TRANSFORM

JIRA_open_init1 <- 
  slice(JIRA_open_init1, -1:-3, preserve = FALSE) %>% 
  select(-3) %>% 
  na.omit() # remove all rows containing 'NA'

JIRA_completed_init1 <- 
  slice(JIRA_completed_init1, -1:-3, preserve = FALSE) %>% 
  select(-3, -10:-13) %>% 
  na.omit() # remove all rows containing 'NA'

JIRA_open_count <- group_by(JIRA_open_init1, type) %>% 
  summarise(count = n())

JIRA_completed_count <- group_by(JIRA_completed_init1, type) %>% 
  summarise(count = n())

## O2D DATA TIDY & TRANSFORM

O2D_assigned_count <- 
  rename(O2D_assigned_init1, type = `Task Type`) %>% 
  group_by(type) %>% 
  summarise(count = n())

O2D_completed_count <- 
  rename(O2D_completed_init1, type = `Task Type`) %>% 
  group_by(type) %>% 
  summarise(count = n())

O2D_prev_mths_count <- 
  rename(O2D_prev_mths_init1, type = `Task Type`) %>% 
  group_by(type) %>% 
  summarise(count = n())

JIRA_open_count
JIRA_completed_count
O2D_assigned_count
O2D_completed_count
O2D_prev_mths_count

# JOINING TOGETHER JIRA TASKS

JIRA_open <- JIRA_open_count %>% 
  rename("open" = "count")

JIRA_completed <- JIRA_completed_count %>% 
  rename("completed" = "count")

JIRA_all <- full_join(JIRA_open, JIRA_completed) %>% 
  mutate(month = current_month) %>% # set this to 'current month'
  select(month, everything())

JIRA_all

# JOINING TOGETHER O2D TASKS FOR JOBS OPEN & COMPLETE

O2D_open <- O2D_assigned_count %>% 
  rename("open" = "count")

O2D_completed <- O2D_completed_count %>% 
  rename("completed" = "count")

O2D_all <- full_join(O2D_open, O2D_completed) %>% 
  mutate(month = current_month) %>% # set this to 'current month'
  select(month, everything())

O2D_all

sum(O2D_prev_mths_count$count) #calculate the sum total of outstanding jobs from previous month.

####

# REORG DATAFRAMES - * consider using functions for repetitive process

JIRA_all

JIRA_CP <- filter(JIRA_all, type == 'Close Project') %>% 
  rename(CP_open = open, CP_comp = completed) %>% 
  select(3:4)

JIRA_GEN <- filter(JIRA_all, type == 'General') %>% 
  rename(GEN_open = open, GEN_comp = completed) %>% 
  select(3:4)

JIRA_PC <- filter(JIRA_all, type == 'Partners Change') %>% 
  rename(PC_open = open, PC_comp = completed) %>% 
  select(3:4)

JIRA_PLC <- filter(JIRA_all, type == 'Project Leader Change') %>% 
  rename(PLC_open = open, PLC_comp = completed) %>% 
  select(3:4)

JIRA_PLDC <- filter(JIRA_all, type == 'Project Leader Delegation Check') %>% 
  rename(PLDC_open = open, PLDC_comp = completed) %>% 
  select(3:4)

JIRA_PLDC <- filter(JIRA_all, type == 'Project Leader Delegation Check') %>% 
  rename(PLDC_open = open, PLDC_comp = completed) %>% 
  select(3:4)

Reduce(merge, list(JIRA_CP, 
                   JIRA_GEN, 
                   JIRA_PC, 
                   JIRA_PLC,
                   JIRA_PLDC,
                   JIRA_RRC)) %>%  # merges multiple data frames
  mutate(month = current_month) %>% # set this to 'current month'
  select(month, everything())
       
  
############## ATTEMPTED FUNCTION TO SIMPLIFY ABOVE CODE

merge_tasks <- function(df, task, abbr1, abbr2) {
  merging <- filter(df, type == task) %>% 
    rename(abbr1 = open, abbr2 = completed) %>% 
    select(3:4)
  
  merging
}


JIRA_CP <-merge_tasks(JIRA_all, 'Close Project', CP_open, CP_closed)
JIRA_GEN <- merge_tasks(JIRA_all, 'General', GEN_open, GEN_closed)

JIRA_CP
JIRA_GEN

##############

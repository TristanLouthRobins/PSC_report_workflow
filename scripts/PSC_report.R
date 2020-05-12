library(tidyverse)
library(readxl)
library(purrr)
library(lubridate)

remotes::install_github("sparce/DSreport") # installing Project Template.

# Cache CSV snapshot using readxl of raw datatables from PSC_JIRA and PSC_O2D report 
# workbooks. Using read_xl currently results in an error.

# PSC_JIRA_xl <- readxl_example(path = 'data/PSC_JIRA_Mar20.xlsm') %>% 
#    read_excel(sheet = "Created") %>% 
#    write_csv('data/JIRA_Mar20_created.csv')

## Include additional code to dispose sheets into .csv files with appropriate naming conventions.

# IMPORT CSV EXPORTS AS NEW VARIABLES

# JIRA REPORTS
JIRA_open <- read_csv('data/PSC_JIRA_Mar20_created.csv')              #Q: possible to find/replace filenames, i.e. latest version?
JIRA_completed <- read_csv('data/PSC_JIRA_Mar20_completed.csv')

# O2D REPORTS
O2D_assigned <- read_csv('data/PSC_O2D_Mar20_assigned.csv')
O2D_completed <- read_csv('data/PSC_O2D_Mar20_completed.csv')
O2D_open_prev_month <- read_csv('data/PSC_O2D_Mar20_prev_month.csv')

# JIRA & O2D DATA TIDY & TRANSFORM
# The JIRA reports (from xlsm files) originally has blank headers (X1, X2..) and the text formatting 
# had defaulted to 'chr'. This is why the rename function in called, with slice used to remove blank
# rows below the blank headers. At a later stage, it would be good to reformat the text format of given
# columns - i.e. issue_ident to dbl, created and resolved to dttm. Not necessary for the purposes of this 
# code but might be handy in future if data is to be manipulated in other ways.

JIRA_open <- 
  slice(JIRA_open, -0:-3, preserve = FALSE) %>% 
  select(1:11) %>% 
  rename(class = X1, issue_type = X2, Key = X3, issue_ident = X4, 
       summary = X5, status = X6, resolution = X7, created = X8, resolved = X9, 
       reporter = X10, assignee = X11) %>% 
  na.omit() # remove all rows containing 'NA'

JIRA_completed <- 
  slice(JIRA_completed, -0:-3, preserve = FALSE) %>% 
  select(1:11) %>% 
  rename(class = X1, issue_type = X2, Key = X3, issue_ident = X4, 
         summary = X5, status = X6, resolution = X7, created = X8, resolved = X9, 
         reporter = X10, assignee = X11)  %>% 
  na.omit() # remove all rows containing 'NA'

JIRA_open_class <- group_by(JIRA_open, class) %>% 
  summarise(count = n())

JIRA_completed_class <- group_by(JIRA_completed, class) %>% 
  summarise(count = n())

JIRA_open_class
JIRA_completed_class

## O2D DATA TIDY & TRANSFORM

O2D_assigned <- 
  rename(O2D_assigned, class = `Task Type`) %>% 
  group_by(class) %>% 
  summarise(count = n())

O2D_completed <- 
  rename(O2D_completed, class = `Task Type`) %>% 
  group_by(class) %>% 
  summarise(count = n())

O2D_open_prev_month <- 
  rename(O2D_open_prev_month, class = `Task Type`) %>% 
  group_by(class) %>% 
  summarise(count = n())

O2D_assigned
O2D_completed
O2D_open_prev_month

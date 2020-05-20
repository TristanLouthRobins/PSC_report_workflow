library(tidyverse)
library(readxl)
library(purrr)
library(lubridate)

## SET MONTH OF REPORT HERE! ##

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

#########################

JIRA_tasklist <- read_csv("data/JIRA_tasklist.csv") %>% 
  select(1)

#########################

JIRA_open_count <- group_by(JIRA_open_init1, type) %>% 
  summarise(count = n()) %>% 
  full_join(JIRA_tasklist)

JIRA_open_count[is.na(JIRA_open_count)] <- 0
  
JIRA_completed_count <- group_by(JIRA_completed_init1, type) %>% 
  summarise(count = n()) %>% 
  full_join(JIRA_tasklist)

JIRA_completed_count[is.na(JIRA_completed_count)] <- 0

## O2D DATA TIDY & TRANSFORM

#########################

O2D_tasklist <- read_csv("data/O2D_tasklist.csv") %>% 
  select(1)

#########################

O2D_assigned_count <- 
  rename(O2D_assigned_init1, type = `Task Type`) %>% 
  group_by(type) %>% 
  summarise(count = n()) %>% 
  full_join(O2D_tasklist)

O2D_assigned_count[is.na(O2D_assigned_count)] <- 0

O2D_completed_count <- 
  rename(O2D_completed_init1, type = `Task Type`) %>% 
  group_by(type) %>% 
  summarise(count = n()) %>% 
  full_join(O2D_tasklist)

O2D_completed_count[is.na(O2D_completed_count)] <- 0

O2D_prev_mths_count <- 
  rename(O2D_prev_mths_init1, type = `Task Type`) %>% 
  group_by(type) %>% 
  summarise(count = n()) %>% 
  full_join(O2D_tasklist)

O2D_prev_mths_count[is.na(O2D_prev_mths_count)] <- 0

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

view(O2D_all)

ALL_TIPV <- sum(O2D_prev_mths_count$count) #calculate the sum total of outstanding jobs from previous month.

# REORG DATAFRAMES - * consider using functions for repetitive process

############## TEST FUNCTION TO SIMPLIFY BELOW CODE

merge_tasks <- function(df, task) {  
  merging <- filter(df, type == task) %>%
    select(3:4)
  
  merging
}

JIRA_close <- merge_tasks(JIRA_all, "Close Project") %>% 
  rename(GEN_open = open, GEN_comp = completed)

JIRA_close

#############

# JIRA MERGING #

JIRA_all

JIRA_CP <- filter(JIRA_all, type == 'Close Project') %>% 
  rename(CP_open = open, CP_comp = completed) %>% 
  select(3:4)

JIRA_CWBS <- filter(JIRA_all, type == 'Close WBS Stage') %>% 
  rename(CWBS_open = open, CWBS_comp = completed) %>% 
  select(3:4)

JIRA_GEN <- filter(JIRA_all, type == 'General') %>% 
  rename(GEN_open = open, GEN_comp = completed) %>% 
  select(3:4)

JIRA_LEP <- filter(JIRA_all, type == 'Link To Existing Project') %>% 
  rename(LEP_open = open, LEP_comp = completed) %>% 
  select(3:4)

JIRA_NC_WBS_P <- filter(JIRA_all, type == 'New Capital (WBS) Project') %>% 
  rename(NC_WBS_P_open = open, NC_WBS_P_comp = completed) %>% 
  select(3:4)

JIRA_NS_WBS_P <- filter(JIRA_all, type == 'New Support (WBS) Project') %>% 
  rename(NS_WBS_P_open = open, NS_WBS_P_comp = completed) %>% 
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

JIRA_PTC <- filter(JIRA_all, type == 'Planning Tool Changes') %>% 
  rename(PTC_open = open, PTC_comp = completed) %>% 
  select(3:4)

JIRA_RRC <- filter(JIRA_all, type == 'Revenue Recognition Change') %>% 
  rename(RRC_open = open, RRC_comp = completed) %>% 
  select(3:4)

JIRA_full <- Reduce(merge, list(JIRA_CP,  # Reduce() and list() merges multiple data frames
                                JIRA_CWBS,
                                JIRA_GEN,
                                JIRA_LEP,
                                JIRA_NC_WBS_P,
                                JIRA_NS_WBS_P,
                                JIRA_PC,
                                JIRA_PLC,
                                JIRA_PLDC,
                                JIRA_PTC,
                                JIRA_RRC)) %>%  
  mutate(month = current_month) %>% # set this to 'current month'
  select(month, everything())

JIRA_full[is.na(JIRA_full)] <- 0    

JIRA_full

# O2D MERGING #

O2D_all
view(O2D_all)

O2D_BM_UPIT <- filter(O2D_all, type == 'Billing Milestone - Update Printed on Invoice Text') %>% 
  rename(BM_UPIT_open = open, BM_UPIT_comp = completed) %>% 
  select(3:4)

O2D_CP <- filter(O2D_all, type == 'Close Project') %>% 
  rename(CP_open = open, CP_comp = completed) %>% 
  select(3:4)

O2D_CWBS <- filter(O2D_all, type == 'Close WBS Stage') %>% 
  rename(CWBS_open = open, CWBS_comp = completed) %>% 
  select(3:4)

O2D_GEN <- filter(O2D_all, type == 'General') %>% 
  rename(GEN_open = open, GEN_comp = completed) %>% 
  select(3:4)

O2D_LEP <- filter(O2D_all, type == 'Link to Existing Project') %>% 
  rename(LEP_open = open, LEP_comp = completed) %>% 
  select(3:4)

O2D_NCP <- filter(O2D_all, type == 'New Capital Project') %>% 
  rename(NCP_open = open, NCP_comp = completed) %>% 
  select(3:4)

O2D_NCSP <- filter(O2D_all, type == 'New Complex Structure Project') %>% 
  rename(NCSP_open = open, NCSP_comp = completed) %>% 
  select(3:4)

O2D_NPC <- filter(O2D_all, type == 'New Project Created') %>% 
  rename(NPC_open = open, NPC_comp = completed) %>% 
  select(3:4)

O2D_NRRPMP <- filter(O2D_all, type == 'New RR Progress Milestones Project') %>% 
  rename(NRRPMP_open = open, NRRPMP_comp = completed) %>% 
  select(3:4)

O2D_RR_TPLP <- filter(O2D_all, type == 'New RR Time Proportional (Linear) Project') %>% 
  rename(RR_TPLP_open = open, RR_TPLP_comp = completed) %>% 
  select(3:4)

O2D_RR_TPPP <- filter(O2D_all, type == 'New RR Time Proportional (Phase) Project') %>% 
  rename(RR_TPPP_open = open, RR_TPPP_comp = completed) %>% 
  select(3:4)

O2D_NSP <- filter(O2D_all, type == 'New Support Project') %>% 
  rename(NSP_open = open, NSP_comp = completed) %>% 
  select(3:4)

O2D_WBS_RRP <- filter(O2D_all, type == 'New WBS for RR Project') %>% 
  rename(WBS_RRP_open = open, WBS_RRP_comp = completed) %>% 
  select(3:4)

O2D_PC <- filter(O2D_all, type == 'Partners Change') %>% 
  rename(PC_open = open, PC_comp = completed) %>% 
  select(3:4)

O2D_PLDC_PEDC <- filter(O2D_all, type == 'PL Delegation Check for Project End Date Change') %>% 
  rename(PLDC_PEDC_open = open, PLDC_PEDC_comp = completed) %>% 
  select(3:4)

O2D_PTC <- filter(O2D_all, type == 'Planning Tool Changes') %>% 
  rename(PTC_open = open, PTC_comp = completed) %>% 
  select(3:4)

O2D_PED <- filter(O2D_all, type == 'Project End Date') %>% 
  rename(PED_open = open, PED_comp = completed) %>% 
  select(3:4)

O2D_PLC <- filter(O2D_all, type == 'Project Leader Change') %>% 
  rename(PLC_open = open, PLC_comp = completed) %>% 
  select(3:4)

O2D_PLDC <- filter(O2D_all, type == 'Project Leader Delegation Check') %>% 
  rename(PLDC_open = open, PLDC_comp = completed) %>% 
  select(3:4)

O2D_RM_MC <- filter(O2D_all, type == 'Revenue Milestone - Milestone Completed') %>% 
  rename(RM_MC_open = open, RM_MC_comp = completed) %>% 
  select(3:4)

O2D_RRC <- filter(O2D_all, type == 'Revenue Recognition Change') %>% 
  rename(RRC_open = open, RRC_comp = completed) %>% 
  select(3:4)

O2D_RSDC <- filter(O2D_all, type == 'Review Stage Dates Changes') %>% 
  rename(RSDC_open = open, RSDC_comp = completed) %>% 
  select(3:4)

O2D_NCSP <- filter(O2D_all, type == 'New Complex Structure Project') %>% 
  rename(NCSP_open = open, NCSP_comp = completed) %>% 
  select(3:4)

O2D_NPC <- filter(O2D_all, type == 'New Project Created') %>% 
  rename(NPC_open = open, NPC_comp = completed) %>% 
  select(3:4)

O2D_full <- Reduce(merge, list(O2D_BM_UPIT, # Reduce() and list() merges multiple data frames
                               O2D_CP,
                               O2D_CWBS, 
                               O2D_GEN, 
                               O2D_LEP, 
                               O2D_NCP,
                               O2D_NCSP,
                               O2D_NPC,
                               O2D_NRRPMP,
                               O2D_RR_TPLP,
                               O2D_RR_TPPP,
                               O2D_NSP,
                               O2D_WBS_RRP,
                               O2D_PC,
                               O2D_PLDC_PEDC,
                               O2D_PTC,
                               O2D_PED,
                               O2D_PLC,
                               O2D_PLDC,
                               O2D_RM_MC,
                               O2D_RRC,
                               O2D_RSDC)) %>%  
  mutate(month = current_month) %>% # set this to 'current month'
  select(month, everything())

O2D_full[is.na(O2D_full)] <- 0

### O2D AND JIRA MERGE FOR MONTHLY SUMMARY ###

# Key for abbreviations and incorporated JIRA & O2D dataframes #

# PTC    - PLANNING TOOL CHANGES
## O2D_PTC, O2D_RSDC
## JIRA_PTC 

# NPC    - NEW PROJECT CREATION
## O2D_LEP, O2D_NPC, O2D_NCP, O2D_NCSP, O2D_NSP, O2D_WBS_RRP
## JIRA_LEP, JIRA_NC_WBS_P, JIRA_NS_WBS_P

# PLA    - PROJECT LEADER AUTHORISATIONS & CHANGES
## O2D_PLC, O2D_PLDC, O2D_PCDC_PEDC, O2D_PC*
## JIRA_PLDC, JIRA_PC

# OT     - OTHER TASKS
## O2D_GEN
## JIRA_GEN

# BPM_RR - BILLING & PROGRESS MS; REVENUE RECOGNITION
## O2D_RRC, O2D_RR_TPLP, O2D_RM_MC, O2D_NRRPMP, O2D_RR_TPPP, O2D_BM_UPIT
## JIRA_RRC

# CPS    - CLOSE PROJECT/STAGE
## O2D_CP, O2D_CWBS
## JIRA_CP, JIRA_CWBS

# PDC    - PROJECT DATE CHANGES
## O2D_PED
## JIRA_PED

# TCT    - TOTAL COMPLETED TASKS (AND %) open, completed, % completed
## SUM OF PTC, NPC, PLA, OT, BPM_RR, CPS, PDC

# TC     - TASKS CREATED
## TCT Open tasks

# TIPV   - TASKS INCOMPLETE FROM PREVIOUS MONTH
## ALL_TIPV

##

JIRA_full

O2D_full

monthly_summary <- 
  mutate(PTC_Open = sum())

  


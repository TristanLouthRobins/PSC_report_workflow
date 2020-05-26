library(tidyverse)
library(readxl)
library(gganimate)
library(gifski)
library(gridExtra)
library(purrr)
library(lubridate)

## BEFORE RUNNING SCRIPT, SET MONTH OF REPORT HERE! ##

current_month <- mdy("04-01-2020")

##############################

# Cache CSV snapshot using readxl of raw datatables from PSC_JIRA and PSC_O2D report 
# workbooks. 


# JIRA REPORT READ AND CONVERSION TO CSV
# note: requires additional code to deal with absence of conventional Headers in spreadsheet

JIRA_assigned_import <- read_excel("data/JIRA_10_19-20.xlsm", 
                            sheet = "Created",
                            col_names = c("type", "issue", "key_ident", "issue_id", "summary", "stat", "resolution", "created", "resolved", "reporter", "assignee", "misc"),
                            col_type = c("text", "text", "numeric", "numeric", "text", "text", "text", "date", "date", "text", "text", "text")) %>% 
  write_csv("data/JIRA_assigned.csv")

JIRA_completed_import <- read_excel("data/JIRA_10_19-20.xlsm", 
                             sheet = "Completed",
                             col_names = c("type", "issue", "key_ident", "issue_id", "summary", "stat", "resolution", "created", "resolved", "reporter", "assignee", "misc", "misc"),
                             col_type = c("text", "text", "numeric", "numeric", "text", "text", "text", "date", "date", "text", "text", "text", "text")) %>% 
  write_csv("data/JIRA_completed.csv")

JIRA_Project_WBS <- read_csv("data/JIRA_PWBS_10_19_20.csv")

# O2D REPORT READ AND CONVERSION TO CSV

O2D_assigned_import <- read_excel("data/O2D_10_19-20.xlsm",  #Q: possible to find/replace filenames, i.e. latest version?
                          sheet = "Assigned", 
                          col_names = TRUE) %>%  
  write_csv("data/O2D_assigned.csv")

O2D_completed_import <- read_excel("data/O2D_10_19-20.xlsm",
                           sheet = "Completed", 
                           col_names = TRUE) %>%  
  write_csv("data/O2D_completed.csv")

O2D_prev_month_import <- read_excel("data/O2D_10_19-20.xlsm",
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
view(O2D_assigned_init1)

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

JIRA_projectWBS <- 

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
  rename(JCP_open = open, JCP_comp = completed) %>% 
  select(3:4)

JIRA_CWBS <- filter(JIRA_all, type == 'Close WBS Stage') %>% 
  rename(JCWBS_open = open, JCWBS_comp = completed) %>% 
  select(3:4)

JIRA_GEN <- filter(JIRA_all, type == 'General') %>% 
  rename(JGEN_open = open, JGEN_comp = completed) %>% 
  select(3:4)

JIRA_PWBS <- JIRA_Project_WBS

JIRA_LEP <- filter(JIRA_all, type == 'Link To Existing Project') %>% 
  rename(JLEP_open = open, JLEP_comp = completed) %>% 
  select(3:4)

JIRA_NC_WBS_P <- filter(JIRA_all, type == 'New Capital (WBS) Project') %>% 
  rename(JNC_WBS_P_open = open, JNC_WBS_P_comp = completed) %>% 
  select(3:4)

JIRA_NS_WBS_P <- filter(JIRA_all, type == 'New Support (WBS) Project') %>% 
  rename(JNS_WBS_P_open = open, JNS_WBS_P_comp = completed) %>% 
  select(3:4)

JIRA_PC <- filter(JIRA_all, type == 'Capital/Support Partners Change') %>% 
  rename(JPC_open = open, JPC_comp = completed) %>% 
  select(3:4)

JIRA_PED <- filter(JIRA_all, type == 'Project End Date') %>% 
  rename(JPED_open = open, JPED_comp = completed) %>% 
  select(3:4)

JIRA_PLC <- filter(JIRA_all, type == 'Project Leader Change') %>% 
  rename(JPED_open = open, JPED_comp = completed) %>% 
  select(3:4)

JIRA_PLDC <- filter(JIRA_all, type == 'Project Leader Delegation Check') %>% 
  rename(JPLDC_open = open, JPLDC_comp = completed) %>% 
  select(3:4)

JIRA_PTC <- filter(JIRA_all, type == 'Planning Tool Changes') %>% 
  rename(JPTC_open = open, JPTC_comp = completed) %>% 
  select(3:4)

JIRA_RRC <- filter(JIRA_all, type == 'Revenue Recognition Change') %>% 
  rename(JRRC_open = open, JRRC_comp = completed) %>% 
  select(3:4)

# Reduce() and list() merges multiple data frames
JIRA_full <- Reduce(merge, list(JIRA_CP, JIRA_CWBS, JIRA_GEN, JIRA_LEP, JIRA_NC_WBS_P, JIRA_NS_WBS_P,
                                JIRA_PC, JIRA_PLC, JIRA_PLDC, JIRA_PTC, JIRA_RRC)) %>%  
  mutate(month = current_month) %>% # set this to 'current month'
  select(month, everything())

JIRA_full[is.na(JIRA_full)] <- 0    

JIRA_full

# O2D MERGING #

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

O2D_LEP <- filter(O2D_all, type == 'Link To Existing Project') %>% 
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

O2D_PC <- filter(O2D_all, type == 'Capital/Support Partners Change') %>% 
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

# Reduce() and list() merges multiple data frames
O2D_full <- Reduce(merge, list(O2D_BM_UPIT, O2D_CP, O2D_CWBS, O2D_GEN, O2D_LEP, O2D_NCP, O2D_NCSP,
                               O2D_NPC, O2D_NRRPMP, O2D_RR_TPLP, O2D_RR_TPPP, O2D_NSP, O2D_WBS_RRP,
                               O2D_PC, O2D_PLDC_PEDC, O2D_PTC, O2D_PED, O2D_PLC, O2D_PLDC, O2D_RM_MC,
                               O2D_RRC, O2D_RSDC)) %>%  
  mutate(month = current_month) %>% # set this to 'current month'
  select(month, everything()) # position new column in first position 

O2D_full[is.na(O2D_full)] <- 0

JIRA_full
O2D_full

## PLANNING TOOL CHANGES ##
PTC_month <- Reduce(merge, list(O2D_PTC,
                                O2D_RSDC,
                                JIRA_PTC)) %>% 
  mutate(MPTC_open = PTC_open + RSDC_open + JPTC_open) %>% 
  mutate(MPTC_comp = PTC_comp + RSDC_comp + JPTC_comp) %>% 
  select(MPTC_open, MPTC_comp)

## NEW PROJECT CREATION ##
NPC_month <- Reduce(merge, list(O2D_LEP, O2D_NPC, O2D_NCP, O2D_NCSP, O2D_NSP, O2D_WBS_RRP,
                                JIRA_LEP,JIRA_NC_WBS_P)) %>% 
  mutate(MNPC_open = LEP_open + NPC_open + NCP_open + NCSP_open + NSP_open + WBS_RRP_open +
           JLEP_open + JNC_WBS_P_open) %>% 
  mutate(MNPC_comp = LEP_comp + NPC_comp + NCP_comp + NCSP_comp + NSP_comp + WBS_RRP_comp +
           JLEP_comp + JNC_WBS_P_comp) %>% 
  select(MNPC_open, MNPC_comp)

## PROJECT LEADER AUTHORISATIONS & CHANGES ##
PLA_month <- Reduce(merge, list(O2D_PLC, O2D_PLDC, O2D_PLDC_PEDC, O2D_PC, 
                                JIRA_PLC, JIRA_PLDC, JIRA_PC)) %>% 
  mutate(MPLA_open = PLC_open + PLDC_open + PLDC_PEDC_open + PC_open + 
         JPLDC_open + JPC_open) %>% 
  mutate(MPLA_comp = PLC_comp + PLDC_comp + PLDC_PEDC_comp + PC_comp + 
          JPLDC_comp + JPC_comp) %>% 
  select(MPLA_open, MPLA_comp)

## OTHER TASKS ##

OT_month <- Reduce(merge, list(O2D_GEN, JIRA_GEN, JIRA_PWBS)) %>% 
  mutate(MOT_open = GEN_open + JGEN_open + PWBS_open) %>% 
    mutate(MOT_comp = GEN_comp + JGEN_comp + PWBS_comp) %>% 
    select(MOT_open, MOT_comp)
  
## BILLING & PROGRESS MS; REVENUE RECOGNITION ##
BPM_RR_month <- Reduce(merge, list(O2D_RRC, O2D_RR_TPLP, O2D_RM_MC, O2D_NRRPMP, O2D_RR_TPPP, O2D_BM_UPIT, 
                                   JIRA_RRC)) %>% 
  mutate(MBPM_RR_open = RRC_open + RR_TPLP_open + RM_MC_open + NRRPMP_open + RR_TPPP_open + BM_UPIT_open +
           JRRC_open) %>% 
  mutate(MBPM_RR_comp = RRC_comp + RR_TPLP_comp + RM_MC_comp + NRRPMP_comp + RR_TPPP_comp + BM_UPIT_comp +
           JRRC_comp) %>% 
  select(MBPM_RR_open, MBPM_RR_comp)

## CLOSE PROJECT/STAGE ##
CPS_month <- Reduce(merge, list(O2D_CP, O2D_CWBS, JIRA_CP, JIRA_CWBS)) %>% 
  mutate(MCPS_open = CP_open + CWBS_open + JCP_open + JCWBS_open) %>% 
    mutate(MCPS_comp = CP_comp + CWBS_comp + JCP_comp + JCWBS_comp) %>% 
    select(MCPS_open, MCPS_comp)
  
## PROJECT DATE CHANGES ##
PDC_month <- Reduce(merge, list(O2D_PED, JIRA_PED)) %>% 
  mutate(MPDC_open = PED_open + JPED_open) %>% 
  mutate(MPDC_comp = PED_comp + JPED_comp) %>% 
  select(MPDC_open, MPDC_comp)

## TOTAL COMPLETED TASKS ## 
TCT_month <- Reduce(merge, list(PTC_month, NPC_month, PLA_month, OT_month, BPM_RR_month, CPS_month, PDC_month)) %>% 
  mutate(TCT_open = MPTC_open + MNPC_open + MPLA_open + MOT_open + MBPM_RR_open + MCPS_open + MPDC_open) %>% 
  mutate(TCT_comp = MPTC_comp + MNPC_comp + MPLA_comp + MOT_comp + MBPM_RR_comp + MCPS_comp + MPDC_comp) %>% 
  mutate(TCT_perc = TCT_comp / TCT_open) %>% 
  select(TCT_open, TCT_comp, TCT_perc)

## FULL MONTHLY SUMMARY ##

full_month <- Reduce(merge, list(PTC_month, NPC_month, PLA_month, OT_month, BPM_RR_month, 
                                 CPS_month, PDC_month, TCT_month)) %>% 
  mutate(month = current_month) %>% # set this to 'current month'
  select(month, everything())


write_csv("data/outputs/full_month10.csv")

  
view(full_month)

summary1 <- read_csv("data/outputs/full_month1.csv")
summary2 <- read_csv("data/outputs/full_month2.csv")
summary3 <- read_csv("data/outputs/full_month3.csv")
summary4 <- read_csv("data/outputs/full_month4.csv")
summary5 <- read_csv("data/outputs/full_month5.csv")
summary6 <- read_csv("data/outputs/full_month6.csv")
summary7 <- read_csv("data/outputs/full_month7.csv")
summary8 <- read_csv("data/outputs/full_month8.csv")
summary9 <- read_csv("data/outputs/full_month9.csv")
summary10 <- read_csv("data/outputs/full_month10.csv")
summary11 <- read_csv("data/outputs/full_month11.csv")
summary12 <- read_csv("data/outputs/full_month12.csv")

PSC_bind <- bind_rows(summary1, summary2, summary3, summary4, summary5, summary6, summary7, summary8, summary9, summary10, summary11, summary12) %>% 
  write_csv("data/outputs/final.csv")

PSC_bind <- read_csv("data/outputs/final.csv")

## GATHERED DATA FOR PLOTTING CHARTS ##

PSC_open_gathered <-  
  select(PSC_bind, -3,-5,-7,-9,-11,-13,-15,-16:-18) %>% 
  rename("Planning Tool Changes" = MPTC_open, 
         "New Project Created" = MNPC_open,
         "Project Leader Authorisation" = MPLA_open,
         "Other Tasks" = MOT_open,
         "Billing & Progress MS; Revenue Recognition" = MBPM_RR_open,
         "Close Project/Stage" = MCPS_open,
         "Project Date Changes" = MPDC_open) %>% 
  gather(2:8 , key = 'Task', value = 'n') %>% 
  mutate(Status = "open")

PSC_comp_gathered <- 
  select(PSC_bind, -2,-4,-6,-8,-10,-12,-14,-16:-18) %>% 
  rename("Planning Tool Changes" = MPTC_comp, 
         "New Project Created" = MNPC_comp,
         "Project Leader Authorisation" = MPLA_comp,
         "Other Tasks" = MOT_comp,
         "Billing & Progress MS; Revenue Recognition" = MBPM_RR_comp,
         "Close Project/Stage" = MCPS_comp,
         "Project Date Changes" = MPDC_comp) %>% 
  gather(2:8, key = 'Task', value = 'n') %>% 
  mutate(Status = "completed")

PSC_combined <- full_join(PSC_open_gathered, PSC_comp_gathered)

## LONG DATA FOR PLOTTING TABLES ##

PSC_open_long <-  
  select(PSC_bind, -3,-5,-7,-9,-11,-13,-15,-16:-18) %>% 
  rename("Planning Tool Changes" = MPTC_open, 
         "New Project Created" = MNPC_open,
         "Project Leader Authorisation" = MPLA_open,
         "Other Tasks" = MOT_open,
         "Billing & Progress MS; Revenue Recognition" = MBPM_RR_open,
         "Close Project/Stage" = MCPS_open,
         "Project Date Changes" = MPDC_open) %>% 
  gather(2:8, key = 'Task', value = 'n') %>% 
  spread(1, key = 'month', value = 'n') %>% 
  arrange(Task) %>% 
  rename(" " = Task,
         "Jul" = `2019-07-01`,
         "Aug" = `2019-08-01`,
         "Sep" = `2019-09-01`,
         "Oct" = `2019-10-01`,
         "Nov" = `2019-11-01`,
         "Dec" = `2019-12-01`,
         "Jan" = `2020-01-01`,
         "Feb" = `2020-02-01`,
         "Mar" = `2020-03-01`,
         "Apr" = `2020-04-01`,
         "May" = `2020-05-01`,
         "Jun" = `2020-06-01`)


 

  


  



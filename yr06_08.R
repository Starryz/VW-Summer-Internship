library(readxl)
library(tidyverse)
library(stringi)
library(readr)
library(sqldf)

# help match --------------
# program_cohort__c
program_cohort <- data.frame(
  "year" = 2006:2008,
  "PROGRAM_COHORT__C" = c(
    "a2C39000002zYsyEAE",
    "a2C39000002zYt3EAE",
    "a2C39000002zYt8EAE"
  ), 
  "RECORDTYPEID" = "012390000009qIDAAY",
  "PROPOSAL_FUNDER__C" = "The Lemelson Foundation"
)

## for commit
extract_c <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/salesforce_examples/Organization_extract.csv") %>% 
  rename("GRANTED_INSTITUTION__C" = "NAME") %>% 
  select(-ORGANIZATION_ALIAS_NAME__C)

extract_alias_c <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/salesforce_examples/Organization_extract.csv") %>% 
  rename("GRANTED_INSTITUTION__C" = "ORGANIZATION_ALIAS_NAME__C") %>% 
  select(-NAME)

## for proposal
extract_p <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/salesforce_examples/Organization_extract.csv") %>% 
  rename("APPLYING_INSTITUTION_NAME__C" = "NAME") %>% 
  select(-ORGANIZATION_ALIAS_NAME__C)

extract_alias_p <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/salesforce_examples/Organization_extract.csv") %>% 
  rename("APPLYING_INSTITUTION_NAME__C" = "ORGANIZATION_ALIAS_NAME__C") %>% 
  select(-NAME) %>% 
  na.omit()

## match Zenn ID and team name
match_2006 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2006_proposals.xlsx", 
                         col_types = c("numeric", "text", "text", 
                                       "text", "text", "numeric", "text", 
                                       "text", "text", "text", "text", "text", 
                                       "text", "text", "numeric", "text", 
                                       "numeric", "text", "text", "text", 
                                       "numeric", "text", "text", "text", 
                                       "numeric", "text", "text", "numeric", 
                                       "text", "text", "numeric", "text", 
                                       "text", "text", "text", "text")) %>% 
  select(`Zenn ID`, `Grant Title`, `Institution Name`)

match_c_2006 <- match_2006 %>% 
  rename("GRANTED_INSTITUTION__C" = "Institution Name") %>% 
  select(-`Grant Title`)

match_p_2006 <- match_2006 %>% 
  rename("NAME" = "Grant Title") %>% 
  select(-`Institution Name`)

match_2007 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2007_proposals.xlsx", 
                         col_types = c("numeric", "text", "text", 
                                       "text", "text", "numeric", "text", 
                                       "text", "text", "text", "text", "text", 
                                       "text", "text", "numeric", "text", 
                                       "numeric", "text", "text", "text", 
                                       "numeric", "text", "text", "text", 
                                       "numeric", "text", "text", "numeric", 
                                       "text", "text", "numeric", "text", 
                                       "text", "text", "text", "text")) %>% 
  select(`Zenn ID`, `Grant Title`, `Institution Name`)

match_c_2007 <- match_2007 %>% 
  rename("GRANTED_INSTITUTION__C" = "Institution Name") %>% 
  select(-`Grant Title`)

match_p_2007 <- match_2007 %>% 
  rename("NAME" = "Grant Title") %>% 
  select(-`Institution Name`)

match_2008 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2008_proposals.xlsx", 
                         col_types = c("numeric", "text", "text", 
                                       "text", "text", "numeric", "text", 
                                       "text", "text", "text", "text", "text", 
                                       "text", "text", "numeric", "text", 
                                       "numeric", "text", "text", "text", 
                                       "numeric", "text", "text", "text", 
                                       "numeric", "text", "text", "numeric", 
                                       "text", "text", "numeric", "text", 
                                       "text", "text", "text", "text")) %>% 
  select(`Zenn ID`, `Grant Title`, `Institution Name`)

match_c_2008 <- match_2008 %>% 
  rename("GRANTED_INSTITUTION__C" = "Institution Name") %>% 
  select(-`Grant Title`)

match_p_2008 <- match_2008 %>% 
  rename("NAME" = "Grant Title") %>% 
  select(-`Institution Name`)
# for membership match
contacts <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/salesforce_examples/Contact_Extract.csv") %>% 
  select(ID, EMAIL, NPE01__ALTERNATEEMAIL__C, NPE01__HOMEEMAIL__C,
         NPE01__WORKEMAIL__C, PREVIOUS_EMAIL_ADDRESSES__C, BKUP_EMAIL_ADDRESS__C)

contacts_1 <- contacts %>% 
  select(ID, EMAIL)

# 2006  --------------

advisors_full_2006 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2006_advisors.xlsx") %>% 
  rename("ZENN_ID__C" = "Zenn ID",
         "ROLE__C" = "Team Role",
         "EMAIL" = "Email") %>% 
  mutate(EMAIL = tolower(EMAIL))

advisors_2006 <- advisors_full_2006 %>% 
  select(ZENN_ID__C, ROLE__C, EMAIL)

# proposal 
proposal_2006 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2006_proposals.xlsx") %>% 
  rename(
    "NAME" = "Grant Title",
    "PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C" = "Proposal Summary",
    "EXTERNAL_PROPOSAL_ID__C" = "External Proposal ID"
  ) %>% 
  mutate(
    "year" = as.numeric(format(as.Date(`Date Created`),'%Y')),
    "STATUS__C" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
    "PROPOSAL_NAME_LONG_VERSION__C" = as.character(NAME),
    "DATE_CREATED__C" = as.Date(`Date Created`),
    "DATE_SUBMITTED__C" = as.Date(`Date Application Submitted`),
    "GRANT_PERIOD_END__C" = as.Date(`Actual Period End`),
    "GRANT_PERIOD_START__C" = as.Date(`Actual Period Begin`),
    "AMOUNT_REQUESTED__C" = as.double(`Amount Requested`),
    "ZENN_ID__C" = as.double(`Zenn ID`),
    "AWARD_AMOUNT__C" = as.double(`Amount Approved`), 
    "APPLYING_INSTITUTION_NAME__C" = ifelse(`Institution Name` == "University of Maryland, College Park", "University of Maryland-College Park",
                                            ifelse(`Institution Name` == "Arizona State University at the Tempe Campus", "Arizona State University",
                                                   ifelse(`Institution Name` == "The City College of New York", "CUNY City College",
                                                          ifelse(`Institution Name` == "University of Oklahoma", "University of Oklahoma Norman Campus",
                                                   `Institution Name`))))
  ) %>% 
  select(
    year, NAME, AMOUNT_REQUESTED__C, PROPOSAL_NAME_LONG_VERSION__C, APPLYING_INSTITUTION_NAME__C,
    AWARD_AMOUNT__C, DATE_CREATED__C, DATE_SUBMITTED__C, GRANT_PERIOD_END__C, 
    GRANT_PERIOD_START__C, 
    PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, ZENN_ID__C, STATUS__C,
    EXTERNAL_PROPOSAL_ID__C
  ) %>% 
  filter(is.na(APPLYING_INSTITUTION_NAME__C) == FALSE) %>% 
  left_join(extract_p) %>% 
  left_join(extract_alias_p, by = "APPLYING_INSTITUTION_NAME__C") %>% 
  mutate(ID = coalesce(ID.x, ID.y)) %>% 
  select(-ID.x, -ID.y) %>% 
  rename("APPLYING_INSTITUTION__C" = "ID") %>% 
  left_join(match_p_2006) %>% 
  left_join(program_cohort) %>% 
  select( - `Zenn ID`, -year) %>% 
  unique()

proposal_2006$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C <- str_replace_all(proposal_2006$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, "[:cntrl:]", " ")

proposal_2006 <- sapply(proposal_2006, as.character)
proposal_2006[is.na(proposal_2006)] <- " "
proposal_2006 <- as.data.frame(proposal_2006) 

write_csv(proposal_2006, "new/2006/proposal_2006.csv")

proposal_2006_narrow <- proposal_2006 %>% 
  select(NAME, ZENN_ID__C, EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C, RECORDTYPEID) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") %>% 
  left_join(teamid_2006, by = "EXTERNAL_PROPOSAL_ID__C")

# team
team_2006 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2006_proposals.xlsx") %>% 
  rename(
    "NAME" = "Grant Title"
  ) %>% 
  mutate(
    "RECORDTYPEID" = "012390000009qKOAAY",
    "ALIAS__C" = ifelse(nchar(NAME)  > 80, NAME, "")
  ) %>% 
  select(
    NAME, RECORDTYPEID, ALIAS__C
  ) %>% 
  left_join(match_p_2006) %>% 
  write_csv("new/2006/team_2006.csv")

# note_task
task_2006 <- read_excel("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/sustainable_vision_grants_2006_post_award_notes.xlsx", 
                        col_types = c("numeric", "text", "text", 
                                      "text")) %>% 
  set_names(c("Zenn ID", "Created Date", "Created by",	"Note")) %>% 
  left_join(match_2006) %>% 
  rename("WHATID" = "Zenn ID",
         "DESCRIPTION" = "Note") %>% 
  mutate(STATUS = "Completed",
         PRIORITY = "Normal",
         TYPE = "Internal Note",
         TASKSUBTYPE = "Call",
         ACTIVITYDATE = as.Date(`Created Date`), 
         SUBJECT = "Post Award Note--",
         OWNER = ifelse(`Created by` == "Brenna Breeding", "00539000005UlQaAAK",
                        ifelse(`Created by` == "Michael Norton", "00539000004pukIAAQ",
                               ifelse(`Created by` == "Patricia Boynton", "00570000001K3bpAAC",
                                      ifelse(`Created by` == "Rachel Agoglia", "00570000003QASWAA4",
                                             "00570000004VlXPAA0"))
                        )
         )
  ) %>% 
  unite("SUBJECT", c(SUBJECT, `Created Date`), sep = "", remove = FALSE) %>% 
  unite("SUBJECT", c(SUBJECT, `Created by`), sep = " ", remove = FALSE) %>% 
  select(
    WHATID, ACTIVITYDATE, `Created by`, DESCRIPTION, TYPE, STATUS, PRIORITY, OWNER, SUBJECT
  ) %>% 
  write_csv("new/2006/note_task_2006.csv")

# memebrship 

teamid_2006 <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/new_dataset_migrate/2006/proposal_2006_extract.csv") %>% 
  select(ID, ZENN_ID__C, TEAM__C) %>% 
  rename("PROPOSAL__C" = "ID") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C))

proposal_2006_narrow <- proposal_2006 %>% 
  select(NAME, ZENN_ID__C, EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C, RECORDTYPEID) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C)) %>% 
  left_join(teamid_2006, by = "ZENN_ID__C")

membership_2006 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2006_proposals.xlsx") %>% 
  select(
    `Zenn ID`, `External Proposal ID`, `Application Status`,
    `Grant Title`, `Institution ID`, `Date Application Submitted`,
    `Actual Period Begin`, `Actual Period End`
  ) %>% 
  mutate("Application Status" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
         "Actual Period Begin" = ifelse(`Application Status` == "Funded", `Actual Period Begin`, `Date Application Submitted`),
         "Actual Period End" = ifelse(`Application Status` == "Funded", `Actual Period End`, `Date Application Submitted`), 
         "STATUS__C" = ifelse(`Application Status` == "Funded", "Completed", "Inactive")
  ) %>%
  rename("ZENN_ID__C" = "Zenn ID") %>% 
  left_join(proposal_2006_narrow, by = "ZENN_ID__C") %>% 
  rename(     
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2006) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(EMAIL = tolower(EMAIL),
         ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) %>%
  na.omit() %>% 
  left_join(contacts_1, by = "EMAIL") %>% 
  rename("MEMBER__C" = "ID") %>% 
  na.omit() %>% 
  write_csv("new/2006/member_2006.csv")

membership_2006_small <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2006_proposals.xlsx") %>% 
  select(
    `Zenn ID`, `External Proposal ID`, `Application Status`,
    `Grant Title`, `Institution ID`, `Date Application Submitted`,
    `Actual Period Begin`, `Actual Period End`
  ) %>% 
  mutate("Application Status" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
         "Actual Period Begin" = ifelse(`Application Status` == "Funded", `Actual Period Begin`, `Date Application Submitted`),
         "Actual Period End" = ifelse(`Application Status` == "Funded", `Actual Period End`, `Date Application Submitted`), 
         "STATUS__C" = ifelse(`Application Status` == "Funded", "Completed", "Inactive")
  ) %>%
  rename("ZENN_ID__C" = "Zenn ID") %>% 
  left_join(proposal_2006_narrow) %>% 
  rename(     
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2006) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) %>% 
  na.omit() %>% 
  left_join(contacts_1) %>% 
  rename("MEMBER__C" = "ID") %>% 
  select(-MEMBER__C)

membership_2006_big <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2006_proposals.xlsx") %>% 
  select(
    `Zenn ID`, `External Proposal ID`, `Application Status`,
    `Grant Title`, `Institution ID`, `Date Application Submitted`,
    `Actual Period Begin`, `Actual Period End`
  ) %>% 
  mutate("Application Status" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
         "Actual Period Begin" = ifelse(`Application Status` == "Funded", `Actual Period Begin`, `Date Application Submitted`),
         "Actual Period End" = ifelse(`Application Status` == "Funded", `Actual Period End`, `Date Application Submitted`), 
         "STATUS__C" = ifelse(`Application Status` == "Funded", "Completed", "Inactive")
  ) %>%
  rename("ZENN_ID__C" = "Zenn ID") %>% 
  left_join(proposal_2006_narrow) %>% 
  rename(     
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2006) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) 

no_id_2006 <- dplyr::setdiff(membership_2006_big, membership_2006_small) %>% 
  left_join(advisors_full_2006) %>% 
  write_csv("new/2006/no_id_2006.csv")

# 2007  --------------

advisors_full_2007 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2007_advisors.xlsx") %>% 
  rename("ZENN_ID__C" = "Zenn ID",
         "ROLE__C" = "Team Role",
         "EMAIL" = "Email") %>% 
  mutate(EMAIL = tolower(EMAIL))

advisors_2007 <- advisors_full_2007 %>% 
  select(ZENN_ID__C, ROLE__C, EMAIL)

# proposal 
proposal_2007 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2007_proposals.xlsx") %>% 
  rename(
    "NAME" = "Grant Title",
    "PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C" = "Proposal Summary",
    "EXTERNAL_PROPOSAL_ID__C" = "External Proposal ID"
  ) %>% 
  mutate(
    "year" = as.numeric(format(as.Date(`Date Created`),'%Y')),
    "STATUS__C" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
    "PROPOSAL_NAME_LONG_VERSION__C" = as.character(NAME),
    "DATE_CREATED__C" = as.Date(`Date Created`),
    "DATE_SUBMITTED__C" = as.Date(`Date Application Submitted`),
    "GRANT_PERIOD_END__C" = as.Date(`Actual Period End`),
    "GRANT_PERIOD_START__C" = as.Date(`Actual Period Begin`),
    "AMOUNT_REQUESTED__C" = as.double(`Amount Requested`),
    "ZENN_ID__C" = as.double(`Zenn ID`),
    "AWARD_AMOUNT__C" = as.double(`Amount Approved`), 
    "APPLYING_INSTITUTION_NAME__C" = ifelse(`Institution Name` == "University of Maryland, College Park", "University of Maryland-College Park",
                                            ifelse(`Institution Name` == "Arizona State University at the Tempe Campus", "Arizona State University",
                                                   ifelse(`Institution Name` == "The City College of New York", "CUNY City College",
                                                          ifelse(`Institution Name` == "University of Oklahoma", "University of Oklahoma Norman Campus",
                                                                 ifelse(`Institution Name` == "University of Texas at Arlington", "The University of Texas at Arlington",
                                                                        ifelse(`Institution Name` == "University of Tennessee, Knoxville", "The University of Tennessee",
                                                                 `Institution Name`))))))
  ) %>% 
  select(
    year, NAME, AMOUNT_REQUESTED__C, PROPOSAL_NAME_LONG_VERSION__C, APPLYING_INSTITUTION_NAME__C,
    AWARD_AMOUNT__C, DATE_CREATED__C, DATE_SUBMITTED__C, GRANT_PERIOD_END__C, 
    GRANT_PERIOD_START__C, 
    PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, ZENN_ID__C, STATUS__C,
    EXTERNAL_PROPOSAL_ID__C
  ) %>% 
  filter(is.na(APPLYING_INSTITUTION_NAME__C) == FALSE) %>% 
  left_join(extract_p) %>% 
  left_join(extract_alias_p, by = "APPLYING_INSTITUTION_NAME__C") %>% 
  mutate(ID = coalesce(ID.x, ID.y)) %>% 
  select(-ID.x, -ID.y) %>% 
  rename("APPLYING_INSTITUTION__C" = "ID") %>% 
  left_join(match_p_2007) %>% 
  left_join(program_cohort) %>% 
  select( - `Zenn ID`, -year) %>% 
  unique()

proposal_2007$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C <- str_replace_all(proposal_2007$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, "[:cntrl:]", " ")

proposal_2007 <- sapply(proposal_2007, as.character)
proposal_2007[is.na(proposal_2007)] <- " "
proposal_2007 <- as.data.frame(proposal_2007) 

write_csv(proposal_2007, "new/2007/proposal_2007.csv")

proposal_2007_narrow <- proposal_2007 %>% 
  select(NAME, ZENN_ID__C, EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C, RECORDTYPEID) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") %>% 
  left_join(teamid_2007, by = "EXTERNAL_PROPOSAL_ID__C")

# team
team_2007 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2007_proposals.xlsx") %>% 
  rename(
    "NAME" = "Grant Title"
  ) %>% 
  mutate(
    "RECORDTYPEID" = "012390000009qKOAAY",
    "ALIAS__C" = ifelse(nchar(NAME)  > 80, NAME, "")
  ) %>% 
  select(
    NAME, RECORDTYPEID, ALIAS__C
  ) %>% 
  left_join(match_p_2007) %>% 
  write_csv("new/2007/team_2007.csv")

# note_task
task_2007 <- read_excel("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/sustainable_vision_grants_2007_post_award_notes.xlsx", 
                        col_types = c("numeric", "text", "text", 
                                      "text")) %>% 
  set_names(c("Zenn ID", "Created Date", "Created by",	"Note")) %>% 
  left_join(match_2007) %>% 
  rename("WHATID" = "Zenn ID",
         "DESCRIPTION" = "Note") %>% 
  mutate(STATUS = "Completed",
         PRIORITY = "Normal",
         TYPE = "Internal Note",
         TASKSUBTYPE = "Call",
         ACTIVITYDATE = as.Date(`Created Date`), 
         SUBJECT = "Post Award Note--",
         OWNER = ifelse(`Created by` == "Brenna Breeding", "00539000005UlQaAAK",
                        ifelse(`Created by` == "Michael Norton", "00539000004pukIAAQ",
                               ifelse(`Created by` == "Patricia Boynton", "00570000001K3bpAAC",
                                      ifelse(`Created by` == "Rachel Agoglia", "00570000003QASWAA4",
                                             "00570000004VlXPAA0"))
                        )
         )
  ) %>% 
  unite("SUBJECT", c(SUBJECT, `Created Date`), sep = "", remove = FALSE) %>% 
  unite("SUBJECT", c(SUBJECT, `Created by`), sep = " ", remove = FALSE) %>% 
  select(
    WHATID, ACTIVITYDATE, `Created by`, DESCRIPTION, TYPE, STATUS, PRIORITY, OWNER, SUBJECT
  ) %>% 
  write_csv("new/2007/note_task_2007.csv")

# memebrship 

teamid_2007 <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/new_dataset_migrate/2007/proposal_2007_extract.csv") %>% 
  select(ID, ZENN_ID__C, TEAM__C) %>% 
  rename("PROPOSAL__C" = "ID") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C))

proposal_2007_narrow <- proposal_2007 %>% 
  select(NAME, ZENN_ID__C, EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C, RECORDTYPEID) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C)) %>% 
  left_join(teamid_2007, by = "ZENN_ID__C")

membership_2007 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2007_proposals.xlsx") %>% 
  select(
    `Zenn ID`, `External Proposal ID`, `Application Status`,
    `Grant Title`, `Institution ID`, `Date Application Submitted`,
    `Actual Period Begin`, `Actual Period End`
  ) %>% 
  mutate("Application Status" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
         "Actual Period Begin" = ifelse(`Application Status` == "Funded", `Actual Period Begin`, `Date Application Submitted`),
         "Actual Period End" = ifelse(`Application Status` == "Funded", `Actual Period End`, `Date Application Submitted`), 
         "STATUS__C" = ifelse(`Application Status` == "Funded", "Completed", "Inactive")
  ) %>%
  rename("ZENN_ID__C" = "Zenn ID") %>% 
  left_join(proposal_2007_narrow) %>% 
  rename(     
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2007) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(EMAIL = tolower(EMAIL),
         ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) %>%
  na.omit() %>% 
  left_join(contacts_1, by = "EMAIL") %>% 
  rename("MEMBER__C" = "ID") %>% 
  na.omit() %>% 
  write_csv("new/2007/member_2007.csv")

membership_2007_small <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2007_proposals.xlsx") %>% 
  select(
    `Zenn ID`, `External Proposal ID`, `Application Status`,
    `Grant Title`, `Institution ID`, `Date Application Submitted`,
    `Actual Period Begin`, `Actual Period End`
  ) %>% 
  mutate("Application Status" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
         "Actual Period Begin" = ifelse(`Application Status` == "Funded", `Actual Period Begin`, `Date Application Submitted`),
         "Actual Period End" = ifelse(`Application Status` == "Funded", `Actual Period End`, `Date Application Submitted`), 
         "STATUS__C" = ifelse(`Application Status` == "Funded", "Completed", "Inactive")
  ) %>%
  rename("ZENN_ID__C" = "Zenn ID") %>% 
  left_join(proposal_2007_narrow) %>% 
  rename(     
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2007) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) %>% 
  na.omit() %>% 
  left_join(contacts_1) %>% 
  rename("MEMBER__C" = "ID") %>% 
  select(-MEMBER__C)

membership_2007_big <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2007_proposals.xlsx") %>% 
  select(
    `Zenn ID`, `External Proposal ID`, `Application Status`,
    `Grant Title`, `Institution ID`, `Date Application Submitted`,
    `Actual Period Begin`, `Actual Period End`
  ) %>% 
  mutate("Application Status" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
         "Actual Period Begin" = ifelse(`Application Status` == "Funded", `Actual Period Begin`, `Date Application Submitted`),
         "Actual Period End" = ifelse(`Application Status` == "Funded", `Actual Period End`, `Date Application Submitted`), 
         "STATUS__C" = ifelse(`Application Status` == "Funded", "Completed", "Inactive")
  ) %>%
  rename("ZENN_ID__C" = "Zenn ID") %>% 
  left_join(proposal_2007_narrow) %>% 
  rename(     
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2007) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) 

no_id_2007 <- dplyr::setdiff(membership_2007_big, membership_2007_small) %>% 
  left_join(advisors_full_2007) %>% 
  write_csv("new/2007/no_id_2007.csv")

# 2008  --------------

advisors_full_2008 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2008_advisors.xlsx") %>% 
  rename("ZENN_ID__C" = "Zenn ID",
         "ROLE__C" = "Team Role",
         "EMAIL" = "Email") %>% 
  mutate(EMAIL = tolower(EMAIL))

advisors_2008 <- advisors_full_2008 %>% 
  select(ZENN_ID__C, ROLE__C, EMAIL)

# proposal 
proposal_2008 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2008_proposals.xlsx") %>% 
  rename(
    "NAME" = "Grant Title",
    "PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C" = "Proposal Summary",
    "EXTERNAL_PROPOSAL_ID__C" = "External Proposal ID"
  ) %>% 
  mutate(
    "year" = as.numeric(format(as.Date(`Date Created`),'%Y')),
    "STATUS__C" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
    "PROPOSAL_NAME_LONG_VERSION__C" = as.character(NAME),
    "DATE_CREATED__C" = as.Date(`Date Created`),
    "DATE_SUBMITTED__C" = as.Date(`Date Application Submitted`),
    "GRANT_PERIOD_END__C" = as.Date(`Actual Period End`),
    "GRANT_PERIOD_START__C" = as.Date(`Actual Period Begin`),
    "AMOUNT_REQUESTED__C" = as.double(`Amount Requested`),
    "ZENN_ID__C" = as.double(`Zenn ID`),
    "AWARD_AMOUNT__C" = as.double(`Amount Approved`), 
    "APPLYING_INSTITUTION_NAME__C" = ifelse(`Institution Name` == "University of Maryland, College Park", "University of Maryland-College Park",
                                            ifelse(`Institution Name` == "Arizona State University at the Tempe Campus", "Arizona State University",
                                                   ifelse(`Institution Name` == "The City College of New York", "CUNY City College",
                                                          ifelse(`Institution Name` == "University of Oklahoma", "University of Oklahoma Norman Campus",
                                                                 ifelse(`Institution Name` == "University of Texas at Arlington", "The University of Texas at Arlington",
                                                                        ifelse(`Institution Name` == "University of Tennessee, Knoxville", "The University of Tennessee",
                                                                               `Institution Name`))))))
  ) %>% 
  select(
    year, NAME, AMOUNT_REQUESTED__C, PROPOSAL_NAME_LONG_VERSION__C, APPLYING_INSTITUTION_NAME__C,
    AWARD_AMOUNT__C, DATE_CREATED__C, DATE_SUBMITTED__C, GRANT_PERIOD_END__C, 
    GRANT_PERIOD_START__C, 
    PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, ZENN_ID__C, STATUS__C,
    EXTERNAL_PROPOSAL_ID__C
  ) %>% 
  filter(is.na(APPLYING_INSTITUTION_NAME__C) == FALSE) %>% 
  left_join(extract_p) %>% 
  left_join(extract_alias_p, by = "APPLYING_INSTITUTION_NAME__C") %>% 
  mutate(ID = coalesce(ID.x, ID.y)) %>% 
  select(-ID.x, -ID.y) %>% 
  rename("APPLYING_INSTITUTION__C" = "ID") %>% 
  left_join(match_p_2008) %>% 
  left_join(program_cohort) %>% 
  select( - `Zenn ID`, -year) %>% 
  unique()

proposal_2008$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C <- str_replace_all(proposal_2008$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, "[:cntrl:]", " ")

proposal_2008 <- sapply(proposal_2008, as.character)
proposal_2008[is.na(proposal_2008)] <- " "
proposal_2008 <- as.data.frame(proposal_2008) 

write_csv(proposal_2008, "new/2008/proposal_2008.csv")

proposal_2008_narrow <- proposal_2008 %>% 
  select(NAME, ZENN_ID__C, EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C, RECORDTYPEID) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") %>% 
  left_join(teamid_2008, by = "EXTERNAL_PROPOSAL_ID__C")

# team
team_2008 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2008_proposals.xlsx") %>% 
  rename(
    "NAME" = "Grant Title"
  ) %>% 
  mutate(
    "RECORDTYPEID" = "012390000009qKOAAY",
    "ALIAS__C" = ifelse(nchar(NAME)  > 80, NAME, "")
  ) %>% 
  select(
    NAME, RECORDTYPEID, ALIAS__C
  ) %>% 
  left_join(match_p_2008) %>% 
  write_csv("new/2008/team_2008.csv")

# note_task
task_2008 <- read_excel("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/sustainable_vision_grants_2008_post_award_notes.xlsx", 
                        col_types = c("numeric", "text", "text", 
                                      "text")) %>% 
  set_names(c("Zenn ID", "Created Date", "Created by",	"Note")) %>% 
  left_join(match_2008) %>% 
  rename("WHATID" = "Zenn ID",
         "DESCRIPTION" = "Note") %>% 
  mutate(STATUS = "Completed",
         PRIORITY = "Normal",
         TYPE = "Internal Note",
         TASKSUBTYPE = "Call",
         ACTIVITYDATE = as.Date(`Created Date`), 
         SUBJECT = "Post Award Note--",
         OWNER = ifelse(`Created by` == "Brenna Breeding", "00539000005UlQaAAK",
                        ifelse(`Created by` == "Michael Norton", "00539000004pukIAAQ",
                               ifelse(`Created by` == "Patricia Boynton", "00570000001K3bpAAC",
                                      ifelse(`Created by` == "Rachel Agoglia", "00570000003QASWAA4",
                                             "00570000004VlXPAA0"))
                        )
         )
  ) %>% 
  unite("SUBJECT", c(SUBJECT, `Created Date`), sep = "", remove = FALSE) %>% 
  unite("SUBJECT", c(SUBJECT, `Created by`), sep = " ", remove = FALSE) %>% 
  select(
    WHATID, ACTIVITYDATE, `Created by`, DESCRIPTION, TYPE, STATUS, PRIORITY, OWNER, SUBJECT
  ) %>% 
  write_csv("new/2008/note_task_2008.csv")

# memebrship 

teamid_2008 <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/new_dataset_migrate/2008/proposal_2008_extract.csv") %>% 
  select(ID, ZENN_ID__C, TEAM__C) %>% 
  rename("PROPOSAL__C" = "ID") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C))

proposal_2008_narrow <- proposal_2008 %>% 
  select(NAME, ZENN_ID__C, EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C, RECORDTYPEID) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C)) %>% 
  left_join(teamid_2008, by = "ZENN_ID__C")

membership_2008 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2008_proposals.xlsx") %>% 
  select(
    `Zenn ID`, `External Proposal ID`, `Application Status`,
    `Grant Title`, `Institution ID`, `Date Application Submitted`,
    `Actual Period Begin`, `Actual Period End`
  ) %>% 
  mutate("Application Status" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
         "Actual Period Begin" = ifelse(`Application Status` == "Funded", `Actual Period Begin`, `Date Application Submitted`),
         "Actual Period End" = ifelse(`Application Status` == "Funded", `Actual Period End`, `Date Application Submitted`), 
         "STATUS__C" = ifelse(`Application Status` == "Funded", "Completed", "Inactive")
  ) %>%
  rename("ZENN_ID__C" = "Zenn ID") %>% 
  left_join(proposal_2008_narrow) %>% 
  rename(     
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2008) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(EMAIL = tolower(EMAIL),
         ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) %>%
  na.omit() %>% 
  left_join(contacts_1, by = "EMAIL") %>% 
  rename("MEMBER__C" = "ID") %>% 
  na.omit() %>% 
  write_csv("new/2008/member_2008.csv")

membership_2008_small <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2008_proposals.xlsx") %>% 
  select(
    `Zenn ID`, `External Proposal ID`, `Application Status`,
    `Grant Title`, `Institution ID`, `Date Application Submitted`,
    `Actual Period Begin`, `Actual Period End`
  ) %>% 
  mutate("Application Status" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
         "Actual Period Begin" = ifelse(`Application Status` == "Funded", `Actual Period Begin`, `Date Application Submitted`),
         "Actual Period End" = ifelse(`Application Status` == "Funded", `Actual Period End`, `Date Application Submitted`), 
         "STATUS__C" = ifelse(`Application Status` == "Funded", "Completed", "Inactive")
  ) %>%
  rename("ZENN_ID__C" = "Zenn ID") %>% 
  left_join(proposal_2008_narrow) %>% 
  rename(     
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2008) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) %>% 
  na.omit() %>% 
  left_join(contacts_1) %>% 
  rename("MEMBER__C" = "ID") %>% 
  select(-MEMBER__C)

membership_2008_big <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2008_proposals.xlsx") %>% 
  select(
    `Zenn ID`, `External Proposal ID`, `Application Status`,
    `Grant Title`, `Institution ID`, `Date Application Submitted`,
    `Actual Period Begin`, `Actual Period End`
  ) %>% 
  mutate("Application Status" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
         "Actual Period Begin" = ifelse(`Application Status` == "Funded", `Actual Period Begin`, `Date Application Submitted`),
         "Actual Period End" = ifelse(`Application Status` == "Funded", `Actual Period End`, `Date Application Submitted`), 
         "STATUS__C" = ifelse(`Application Status` == "Funded", "Completed", "Inactive")
  ) %>%
  rename("ZENN_ID__C" = "Zenn ID") %>% 
  left_join(proposal_2008_narrow) %>% 
  rename(     
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2008) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) 

no_id_2008 <- dplyr::setdiff(membership_2008_big, membership_2008_small) %>% 
  left_join(advisors_full_2008) %>% 
  write_csv("new/2008/no_id_2008.csv")
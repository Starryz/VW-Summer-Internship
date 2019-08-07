# setup --------------
library(readxl)
library(tidyverse)
library(stringi)
library(readr)
library(data.table)

# program_cohort__c
program_cohort <- data.frame(
  "year" = 2009:2013,
  "PROGRAM_COHORT__C" = c(
    "a2C39000002zYtDEAU",
    "a2C39000002zYt9EAE",
    "a2C39000002zYtIEAU",
    "a2C39000002zYtNEAU",
    "a2C39000002zYt4EAE"
  ), 
  "RECORDTYPEID" = "01239000000Ap02AAC",
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
match <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2012_proposals.xlsx", 
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

match_c <- match %>% 
  rename("GRANTED_INSTITUTION__C" = "Institution Name") %>% 
  select(-`Grant Title`)

match_p <- match %>% 
  rename("NAME" = "Grant Title") %>% 
  select(-`Institution Name`)

# for membership match
contacts <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/salesforce_examples/Contact_Extract.csv") %>% 
  select(ID, EMAIL, NPE01__ALTERNATEEMAIL__C, NPE01__HOMEEMAIL__C,
         NPE01__WORKEMAIL__C, PREVIOUS_EMAIL_ADDRESSES__C, BKUP_EMAIL_ADDRESS__C)

contacts_1 <- contacts %>% 
  select(ID, EMAIL)

advisors_full_2012 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2012_advisors.xlsx") %>% 
  rename("ZENN_ID__C" = "Zenn ID",
         "ROLE__C" = "Team Role",
         "EMAIL" = "Email") %>% 
  mutate(EMAIL = tolower(EMAIL))

advisors_2012 <- advisors_full_2012 %>% 
  select(ZENN_ID__C, ROLE__C, EMAIL)

# 2012  --------------
# proposal 
proposal_2012 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2012_proposals.xlsx") %>% 
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
    "APPLYING_INSTITUTION_NAME__C" = ifelse(`Institution Name` == "University of Oklahoma", "University of Oklahoma Norman Campus",`Institution Name`)
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
  left_join(match_p) %>% 
  left_join(program_cohort) %>% 
  select( - `Zenn ID`, -year) %>% 
  unique()

proposal_2012$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C <- str_replace_all(proposal_2012$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, "[:cntrl:]", " ")

proposal_2012 <- sapply(proposal_2012, as.character)
proposal_2012[is.na(proposal_2012)] <- " "
proposal_2012 <- as.data.frame(proposal_2012) 

write_csv(proposal_2012, "new/2012/proposal_2012.csv")

proposal_2012_narrow <- proposal_2012 %>% 
  select(NAME, ZENN_ID__C, EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C, RECORDTYPEID) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") %>% 
  left_join(teamid_2012, by = "EXTERNAL_PROPOSAL_ID__C")

# team
team_2012 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2012_proposals.xlsx") %>% 
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
  left_join(match_p) %>% 
  write_csv("new/2012/team_2012.csv")

# note_task
task_2012 <- read_excel("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/sustainable_vision_grants_2012_post_award_notes.xlsx", 
                        col_types = c("numeric", "text", "text", 
                                      "text")) %>% 
  set_names(c("Zenn ID", "Created Date", "Created by",	"Note")) %>% 
  left_join(match) %>% 
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
  write_csv("new/2012/note_task_2012.csv")

# memebrship 

teamid_2012 <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/2009-2012 to be uploaded/2012/proposal_2012_extract.csv") %>% 
  select(ID, ZENN_ID__C, TEAM__C) %>% 
  rename("TEAMID" = "ID", 
         "PROPOSAL__C" = "TEAM__C") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C))

proposal_2012_narrow <- proposal_2012 %>% 
  select(NAME, ZENN_ID__C, EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C, RECORDTYPEID) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C)) %>% 
  left_join(teamid_2012, by = "ZENN_ID__C")

membership_2012 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2012_proposals.xlsx") %>% 
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
  left_join(proposal_2012_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2012) %>% 
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
  write_csv("new/2012/member_2012.csv")

membership_2012_small <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2012_proposals.xlsx") %>% 
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
  left_join(proposal_2012_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2012) %>% 
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

membership_2012_big <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2012_proposals.xlsx") %>% 
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
  left_join(proposal_2012_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2012) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) 

no_id_2012 <- dplyr::setdiff(membership_2012_big, membership_2012_small) %>% 
  left_join(advisors_full_2012) %>% 
  write_csv("new/2012/no_id_2012.csv")

# 2011 -------------------------------
# proposal 
proposal_2011 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2011_proposals.xlsx") %>% 
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
    "APPLYING_INSTITUTION_NAME__C" = ifelse(`Institution Name` == "University of St. Thomas", "University of St Thomas (Saint Paul, Minnesota)",
                                            ifelse(`Institution Name` == "University of St Thomas", "University of St Thomas (Saint Paul, Minnesota)",
                                                   ifelse(`Institution Name` == "Indiana University Northwest", "Indiana University-Northwest",
                                                          ifelse(`Institution Name` == "University of Maryland, Baltimore County", "University of Maryland-Baltimore County", 
                                                                 `Institution Name`))))
  ) %>% 
  select(year, NAME, AMOUNT_REQUESTED__C, PROPOSAL_NAME_LONG_VERSION__C, APPLYING_INSTITUTION_NAME__C,
         AWARD_AMOUNT__C, DATE_CREATED__C, DATE_SUBMITTED__C, GRANT_PERIOD_END__C, 
         GRANT_PERIOD_START__C, 
         PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, ZENN_ID__C, STATUS__C,
         EXTERNAL_PROPOSAL_ID__C
  ) %>% 
  left_join(extract_p) %>% 
  left_join(extract_alias_p, by = "APPLYING_INSTITUTION_NAME__C") %>% 
  mutate(ID = coalesce(ID.x, ID.y)) %>% 
  select(-ID.x, -ID.y) %>% 
  rename("APPLYING_INSTITUTION__C" = "ID") %>% 
  left_join(match_p) %>% 
  left_join(program_cohort) %>% 
  select( - `Zenn ID`, -year) %>% 
  unique()

proposal_2011$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C <- str_replace_all(proposal_2011$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, "[:cntrl:]", " ")

proposal_2011 <- sapply(proposal_2011, as.character)
proposal_2011[is.na(proposal_2011)] <- " "
proposal_2011 <- as.data.frame(proposal_2011)

write_csv(proposal_2011, "new/2011/proposal_2011.csv")


# team 
team_2011 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2011_proposals.xlsx") %>% 
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
  left_join(match_p) %>% 
  write_csv("new/2011/team_2011.csv")

# note_task 
task_2011 <- read_excel("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/sustainable_vision_grants_2011_post_award_notes.xlsx", 
                        col_types = c("numeric", "text", "text", 
                                      "text")) %>% 
  set_names(c("Zenn ID", "Created Date", "Created by",	"Note")) %>% 
  left_join(match) %>% 
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
                                             ifelse(`Created by` == "Allyson Chabot", "0033900001x6eGUAAY",
                                                    "00570000004VlXPAA0")))
                        )
         )
  ) %>% 
  unite("SUBJECT", c(SUBJECT, `Created Date`), sep = "", remove = FALSE) %>% 
  unite("SUBJECT", c(SUBJECT, `Created by`), sep = " ", remove = FALSE) %>% 
  select(
    WHATID, ACTIVITYDATE, `Created by`, DESCRIPTION, TYPE, STATUS, PRIORITY, OWNER, SUBJECT
  ) %>% 
  write_csv("new/2011/note_task_2011.csv")

# memebrship 
advisors_full_2011 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2011_advisors.xlsx") %>% 
  rename("ZENN_ID__C" = "Zenn ID",
         "ROLE__C" = "Team Role",
         "EMAIL" = "Email") %>% 
  mutate(EMAIL = tolower(EMAIL))

advisors_2011 <- advisors_full_2011 %>% 
  select(ZENN_ID__C, ROLE__C, EMAIL)

teamid_2011 <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/2009-2012 to be uploaded/2011/proposal_2011_extract.csv") %>% 
  select(ID, ZENN_ID__C, TEAM__C) %>% 
  rename("TEAMID" = "ID", 
         "PROPOSAL__C" = "TEAM__C") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C))

proposal_2011_narrow <- proposal_2011 %>% 
    select(NAME, ZENN_ID__C, EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C, RECORDTYPEID) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C)) %>% 
  left_join(teamid_2011, by = "ZENN_ID__C")

membership_2011 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2011_proposals.xlsx") %>% 
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
  left_join(proposal_2011_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2011) %>% 
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
  write_csv("new/2011/member_2011.csv")

membership_2011_small <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2011_proposals.xlsx") %>% 
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
  left_join(proposal_2011_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2011) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) %>% 
  left_join(contacts_1) %>% 
  rename("MEMBER__C" = "ID") %>% 
  na.omit() %>% 
  select(-MEMBER__C)

membership_2011_big <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2011_proposals.xlsx") %>% 
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
  left_join(proposal_2011_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2011) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) 

no_id_2011 <- dplyr::setdiff(membership_2011_big, membership_2011_small) %>% 
  left_join(advisors_full_2011) %>% 
  write_csv("new/2011/no_id_2011.csv")

# 2010 --------------
# proposal 
proposal_2010 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2010_proposals.xlsx") %>% 
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
    "APPLYING_INSTITUTION_NAME__C" = ifelse(`Institution Name` == "Polytechnic Institute of New York University", "New York University",
                                            ifelse(`Institution Name` == "University of Minnesota", "University of Minnesota-Twin Cities",
                                                   ifelse(`Institution Name` == "University of Texas at San Antonio", "The University of Texas at San Antonio",
                                                          ifelse(`Institution Name` == "Arizona State University at the Tempe Campus", "Arizona State University",
                                                                 ifelse(`Institution Name` == "University of St. Thomas", "University of St Thomas (Saint Paul, Minnesota)",
                                                                        `Institution Name`)))))
  ) %>% 
  select(year, NAME, AMOUNT_REQUESTED__C, PROPOSAL_NAME_LONG_VERSION__C, APPLYING_INSTITUTION_NAME__C,
         AWARD_AMOUNT__C, DATE_CREATED__C, DATE_SUBMITTED__C, GRANT_PERIOD_END__C, 
         GRANT_PERIOD_START__C, 
         PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, ZENN_ID__C, STATUS__C,
         EXTERNAL_PROPOSAL_ID__C
  ) %>% 
  left_join(extract_p) %>% 
  left_join(extract_alias_p, by = "APPLYING_INSTITUTION_NAME__C") %>% 
  mutate(ID = coalesce(ID.x, ID.y)) %>% 
  select(-ID.x, -ID.y) %>% 
  rename("APPLYING_INSTITUTION__C" = "ID") %>% 
  left_join(match_p) %>% 
  left_join(program_cohort) %>% 
  select( - `Zenn ID`, -year) %>% 
  unique()

proposal_2010$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C <- str_replace_all(proposal_2010$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, "[:cntrl:]", " ")

proposal_2010 <- sapply(proposal_2010, as.character)
proposal_2010[is.na(proposal_2010)] <- " "
proposal_2010 <- as.data.frame(proposal_2010)

write_csv(proposal_2010, "new/2010/proposal_2010.csv")


# team 
team_2010 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2010_proposals.xlsx") %>% 
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
  left_join(match_p) %>% 
  write_csv("new/2010/team_2010.csv")

# note_task 
task_2010 <- read_excel("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/sustainable_vision_grants_2010_post_award_notes.xlsx", 
                        col_types = c("numeric", "text", "text", 
                                      "text")) %>% 
  set_names(c("Zenn ID", "Created Date", "Created by",	"Note")) %>% 
  left_join(match) %>% 
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
                                             ifelse(`Created by` == "Allyson Chabot", "0033900001x6eGUAAY",
                                                    "00570000004VlXPAA0")))
                        )
         )
  ) %>% 
  unite("SUBJECT", c(SUBJECT, `Created Date`), sep = "", remove = FALSE) %>% 
  unite("SUBJECT", c(SUBJECT, `Created by`), sep = " ", remove = FALSE) %>% 
  select(
    WHATID, ACTIVITYDATE, `Created by`, DESCRIPTION, TYPE, STATUS, PRIORITY, OWNER, SUBJECT
  ) %>% 
  write_csv("new/2010/note_task_2010.csv")

# memebrship 
advisors_full_2010 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2010_advisors.xlsx") %>% 
  rename("ZENN_ID__C" = "Zenn ID",
         "ROLE__C" = "Team Role",
         "EMAIL" = "Email") %>% 
  mutate(EMAIL = tolower(EMAIL))

advisors_2010 <- advisors_full_2010 %>% 
  select(ZENN_ID__C, ROLE__C, EMAIL)

teamid_2010 <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/2009-2012 to be uploaded/2010/proposal_2010_extract.csv") %>% 
  select(ID, ZENN_ID__C, TEAM__C) %>% 
  rename("TEAMID" = "ID", 
         "PROPOSAL__C" = "TEAM__C") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C))

proposal_2010_narrow <- proposal_2010 %>% 
  select(NAME, ZENN_ID__C, EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C, RECORDTYPEID) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C)) %>% 
  left_join(teamid_2010, by = "ZENN_ID__C")

membership_2010 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2010_proposals.xlsx") %>% 
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
  left_join(proposal_2010_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2010) %>% 
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
  write_csv("new/2010/member_2010.csv")

membership_2010_small <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2010_proposals.xlsx") %>% 
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
  left_join(proposal_2010_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2010) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) %>% 
  left_join(contacts_1) %>% 
  rename("MEMBER__C" = "ID") %>% 
  na.omit() %>% 
  select(-MEMBER__C)

membership_2010_big <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2010_proposals.xlsx") %>% 
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
  left_join(proposal_2010_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2010) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) 

no_id_2010 <- dplyr::setdiff(membership_2010_big, membership_2010_small) %>% 
  left_join(advisors_full) %>% 
  write_csv("new/2010/no_id_2010.csv")

# 2009 ------------------
# proposal 
proposal_2009 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2009_proposals.xlsx") %>% 
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
    "APPLYING_INSTITUTION_NAME__C" = ifelse(`Institution Name` == "University of Tennessee, Knoxville", "The University of Tennessee",
                                            ifelse(`Institution Name` == "Texas A&M University, Corpus Christi", "Texas A & M University-Corpus Christi",
                                                   ifelse(`Institution Name` == "University of Maryland, College Park", "University of Maryland-College Park",
                                                          ifelse(`Institution Name` == "Arizona State University at the Tempe Campus", "Arizona State University",
                                                                 `Institution Name`))))
  ) %>% 
  select(year, NAME, AMOUNT_REQUESTED__C, PROPOSAL_NAME_LONG_VERSION__C, APPLYING_INSTITUTION_NAME__C,
         AWARD_AMOUNT__C, DATE_CREATED__C, DATE_SUBMITTED__C, GRANT_PERIOD_END__C, 
         GRANT_PERIOD_START__C, 
         PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, ZENN_ID__C, STATUS__C,
         EXTERNAL_PROPOSAL_ID__C
  ) %>% 
  left_join(extract_p) %>% 
  left_join(extract_alias_p, by = "APPLYING_INSTITUTION_NAME__C") %>% 
  mutate(ID = coalesce(ID.x, ID.y)) %>% 
  select(-ID.x, -ID.y) %>% 
  rename("APPLYING_INSTITUTION__C" = "ID") %>% 
  left_join(match_p) %>% 
  left_join(program_cohort) %>% 
  select( - `Zenn ID`, -year) %>% 
  unique()

proposal_2009$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C <- str_replace_all(proposal_2009$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, "[:cntrl:]", " ")

proposal_2009 <- sapply(proposal_2009, as.character)
proposal_2009[is.na(proposal_2009)] <- " "
proposal_2009 <- as.data.frame(proposal_2009)

write_csv(proposal_2009, "new/2009/proposal_2009.csv")


# team 
team_2009 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2009_proposals.xlsx") %>% 
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
  left_join(match_p) %>% 
  write_csv("new/2009/team_2009.csv")

# note_task 
task_2009 <- read_excel("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/sustainable_vision_grants_2009_post_award_notes.xlsx", 
                        col_types = c("numeric", "text", "text", 
                                      "text")) %>% 
  set_names(c("Zenn ID", "Created Date", "Created by",	"Note")) %>% 
  left_join(match) %>% 
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
                                             ifelse(`Created by` == "Allyson Chabot", "0033900001x6eGUAAY",
                                                    "00570000004VlXPAA0")))
                        )
         )
  ) %>% 
  unite("SUBJECT", c(SUBJECT, `Created Date`), sep = "", remove = FALSE) %>% 
  unite("SUBJECT", c(SUBJECT, `Created by`), sep = " ", remove = FALSE) %>% 
  select(
    WHATID, ACTIVITYDATE, `Created by`, DESCRIPTION, TYPE, STATUS, PRIORITY, OWNER, SUBJECT
  ) %>% 
  write_csv("new/2009/note_task_2009.csv")

# memebrship 
advisors_full_2009 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2009_advisors.xlsx") %>% 
  rename("ZENN_ID__C" = "Zenn ID",
         "ROLE__C" = "Team Role",
         "EMAIL" = "Email")%>% 
  mutate(EMAIL = tolower(EMAIL))

advisors_2009 <- advisors_full_2009 %>% 
  select(ZENN_ID__C, ROLE__C, EMAIL)

teamid_2009 <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/2009-2012 to be uploaded/2009/proposal_2009_extract.csv") %>% 
  select(ID, ZENN_ID__C, TEAM__C) %>% 
  rename("TEAMID" = "ID", 
         "PROPOSAL__C" = "TEAM__C") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C))

proposal_2009_narrow <- proposal_2009 %>% 
  select(NAME, ZENN_ID__C, EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C, RECORDTYPEID) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") %>% 
  mutate(ZENN_ID__C = as.character(ZENN_ID__C)) %>% 
  left_join(teamid_2009, by = "ZENN_ID__C")

membership_2009 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2009_proposals.xlsx") %>% 
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
  left_join(proposal_2009_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2009) %>% 
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
  write_csv("new/2009/member_2009.csv")

membership_2009_small <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2009_proposals.xlsx") %>% 
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
  left_join(proposal_2009_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2009) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) %>% 
  left_join(contacts_1) %>% 
  rename("MEMBER__C" = "ID") %>% 
  na.omit() %>% 
  select(-MEMBER__C)

membership_2009_big <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2009_proposals.xlsx") %>% 
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
  left_join(proposal_2009_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors_2009) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) 

no_id_2009 <- dplyr::setdiff(membership_2009_big, membership_2009_small) %>% 
  left_join(advisors_full_2009) %>% 
  write_csv("new/2009/no_id_2009.csv")
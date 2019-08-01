library(readxl)
library(tidyverse)
library(stringi)
library(readr)

# help match --------------
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
  select(-NAME)

## match Zenn ID and team name
match <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx", 
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

# commits ------------------
commits_2013 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
  filter(`Application Status` == "funded") %>% 
  rename(
    "GRANT_ID__C" = "External Proposal ID", 
    "GRANTED_INSTITUTION__C" = "Institution Name",
    "AMOUNT_DISBURSED__C" = "Amount Disbursed",
    "PROGRAM__C" = "Type"
  ) %>% 
  mutate(
    "DISBURSEMENT_REQUEST_AMOUNT__C" = as.double(AMOUNT_DISBURSED__C),
    "GRANT_STATUS__C" = stri_trans_totitle(`Application Status`),
    "AWARD_LETTER_SENT__C" = as.Date(`Grant Letter Sent`),
    "AWARD_LETTER_SIGNED__C" = as.Date(`Grant Letter Signed`),
    "GRANT_START_DATE__C" = as.Date(`Actual Period Begin`),
    "GRANT_END_DATE__C" = as.Date(`Actual Period End`),
    "PAYMENT_STATUS__C" = "Paid", 
    "AMOUNT_APPROVED__C" = as.double(`Amount Approved`)
  ) %>% 
  select(
    AMOUNT_APPROVED__C, `GRANT_ID__C`, `AMOUNT_APPROVED__C`, `AWARD_LETTER_SENT__C`, `AWARD_LETTER_SIGNED__C`, 
    `PAYMENT_STATUS__C`, `GRANT_START_DATE__C`, `PROGRAM__C`, `GRANT_END_DATE__C`, 
    `GRANT_STATUS__C`, `GRANTED_INSTITUTION__C`, `DISBURSEMENT_REQUEST_AMOUNT__C`
  ) %>% 
  left_join(extract_c) %>% 
  left_join(extract_alias_c, by = "GRANTED_INSTITUTION__C") %>% 
  mutate(ID = coalesce(ID.x, ID.y)) %>% 
  select(-ID.x, -ID.y) %>% 
  left_join(match_c) %>% 
  write_csv("new/commits_2013.csv")

# proposal -------------------------------
proposal_2013 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
  rename(
    "NAME" = "Grant Title",
    #"APPLYING_INSTITUTION_NAME__C" = "Institution Name",
    "PROGRAM_COHORT_RECORD_TYPE__C" = "Type",
    "PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C" = "Proposal Summary",
    "EXTERNAL_PROPOSAL_ID__C" = "External Proposal ID"
  ) %>% 
  mutate(
    "RECORDTYPEID" = "01239000000Ap02AAC",
    "STATUS__C" = ifelse(`Application Status` == "invite resubmit", "Invited Resubmit", stri_trans_totitle(`Application Status`)),
    "PROPOSAL_NAME_LONG_VERSION__C" = as.character(NAME),
    "DATE_CREATED__C" = as.Date(`Date Created`),
    "DATE_SUBMITTED__C" = as.Date(`Date Application Submitted`),
    "GRANT_PERIOD_END__C" = as.Date(`Actual Period End`),
    "GRANT_PERIOD_START__C" = as.Date(`Actual Period Begin`),
    "AMOUNT_REQUESTED__C" = as.double(`Amount Requested`),
    "ZENN_ID__C" = as.double(`Zenn ID`),
    "AWARD_AMOUNT__C" = as.double(`Amount Approved`), 
    "PROGRAM_COHORT__C" = "a2C39000002zYt4EAE",
    "PROPOSAL_FUNDER__C" = "The Lemelson Foundation",
    "APPLYING_INSTITUTION_NAME__C" = ifelse(`Institution Name` == "University of Tennessee, Knoxville", "The University of Tennessee",
                                            ifelse(`Institution Name` == "Cogswell Polytechnical College", "Cogswell College",
                                                   ifelse(`Institution Name` == "Arizona State University at the Tempe Campus", "Arizona State University", `Institution Name`)))
  ) %>% 
  select(
    NAME, RECORDTYPEID, AMOUNT_REQUESTED__C, PROPOSAL_NAME_LONG_VERSION__C, APPLYING_INSTITUTION_NAME__C,
    AWARD_AMOUNT__C, DATE_CREATED__C, DATE_SUBMITTED__C, GRANT_PERIOD_END__C, 
    GRANT_PERIOD_START__C, PROGRAM_COHORT_RECORD_TYPE__C, 
    PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, ZENN_ID__C, STATUS__C,
    EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C
  ) %>% 
  left_join(extract_p) %>% 
  left_join(extract_alias_p, by = "APPLYING_INSTITUTION_NAME__C") %>% 
  mutate(ID = coalesce(ID.x, ID.y)) %>% 
  select(-ID.x, -ID.y) %>% 
  rename("APPLYING_INSTITUTION__C" = "ID") %>% 
  left_join(match_p) %>% 
  select( - `Zenn ID`)

proposal_2013$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C <- str_replace_all(proposal_2013$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, "[:cntrl:]", " ")

str_remove_all("\u0093systems\u0094", "[[\\[u]+[0-9]*]]")

proposal_2013 <- sapply(proposal_2013, as.character)
proposal_2013[is.na(proposal_2013)] <- " "
proposal_2013 <- as.data.frame(proposal_2013)

write_csv(proposal_2013, "new/proposal_2013.csv")

# team --------------------------
team_2013 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
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
  write_csv("new/team_2013.csv")

# membership -----------------------------
membership_2013_1a <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
  mutate(
    "START_DATE__C" = as.Date(`Actual Period Begin`),
    "END_DATE__C" = as.Date(`Actual Period End`)
  ) %>% 
  rename(
    "TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "Grant Title",
    "PROPOSAL_STATUS__C" ="Application Status"
  ) %>% 
  select(
    START_DATE__C, END_DATE__C, TEAM_NAME_TEXT_ONLY_HIDDEN__C, PROPOSAL_STATUS__C
  )

membership_2013_1b <- proposal_2013 %>% 
  select(NAME, ZENN_ID__C) %>% 
  rename(
    "TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME"
  ) 

membership_2013_2 <- merge(membership_2013_1a, membership_2013_1b) # add ZENN ID

advisors <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_advisors.xlsx") %>% 
  rename(
    "ZENN_ID__C" = "Zenn ID"
  )

membership_2013 <- merge(membership_2013_2, advisors) %>% 
  mutate(
    "START_DATE__C" = as.Date(`START_DATE__C`),
    "END_DATE__C" = as.Date(`END_DATE__C`),
    "PROGRAM_TYPE_FORMULA__C" = "Sustainable Vision"
  ) %>% 
  unite(
    "FULL_NAME__C", c(`First Name`, `Last Name`), sep = " ", remove = FALSE
  ) %>% 
  rename(
    "ROLE__C" = "Team Role",
    "EMAIL_FORMULA__C" = Email,
    "PHONE_FORMULA__C" = `Telephone 1`,
    "FIRST_NAME__C" = `First Name`,
    "LAST_NAME__C" = `Last Name`,
    "ORGANIZATION__C" = Organization
  ) %>% 
  select(
    ROLE__C, START_DATE__C, END_DATE__C, FULL_NAME__C, EMAIL_FORMULA__C, 
    PHONE_FORMULA__C, PROGRAM_TYPE_FORMULA__C, FIRST_NAME__C, ORGANIZATION__C,
    PROPOSAL_STATUS__C, LAST_NAME__C, TEAM_NAME_TEXT_ONLY_HIDDEN__C
  ) %>% 
  write_csv("new/member_2013.csv")
## note: status needs capitalization  - stri_trans_totitle()

# task -------------------------------------------------
task_2013 <- read_excel("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/sustainable_vision_grants_2013_post_award_notes.xlsx", 
                        col_types = c("numeric", "text", "text", 
                                      "text")) %>% 
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
  write_csv("new/note_task_2013.csv")
%>% 
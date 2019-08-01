library(readxl)
library(tidyverse)
library(stringi)
library(readr)
library(data.table)

# setup --------------
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
match <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2011_proposals.xlsx", 
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

# proposal -------------------------------
proposal_2011 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2011_proposals.xlsx") %>% 
  rename(
    "NAME" = "Grant Title",
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
    "APPLYING_INSTITUTION_NAME__C" = ifelse(`Institution Name` %in% c("University of St. Thomas", "University of St Thomas"), "University of St Thomas (Saint Paul, Minnesota)", 
                                            ifelse(`Institution Name` == "Indiana University Northwest", "Indiana University-Northwest",
                                                   ifelse(`Institution Name` == "University of Maryland, Baltimore County", "University of Maryland-Baltimore County",
                                                          `Institution Name`)))) %>% 
  select(
    NAME, RECORDTYPEID, AMOUNT_REQUESTED__C, PROPOSAL_NAME_LONG_VERSION__C, APPLYING_INSTITUTION_NAME__C,
    AWARD_AMOUNT__C, DATE_CREATED__C, DATE_SUBMITTED__C, GRANT_PERIOD_END__C, 
    GRANT_PERIOD_START__C, 
    PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, ZENN_ID__C, STATUS__C,
    EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C
  ) %>% 
  left_join(extract_p) %>% 
  left_join(extract_alias_p, by = "APPLYING_INSTITUTION_NAME__C") %>% 
  mutate(ID = coalesce(ID.x, ID.y)) %>% 
  select(-ID.x, -ID.y) %>% 
  rename("APPLYING_INSTITUTION__C" = "ID") %>% 
  left_join(match_p) %>% 
  select( - `Zenn ID`, APPLYING_INSTITUTION_NAME__C) %>% 
  unique()

proposal_2011$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C <- str_replace_all(proposal_2011$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, "[:cntrl:]", " ")

proposal_2011 <- sapply(proposal_2011, as.character)
proposal_2011[is.na(proposal_2011)] <- " "
proposal_2011 <- as.data.frame(proposal_2011)

write_csv(proposal_2011, "new/2011/proposal_2011.csv")


# team --------------------------
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

# note_task -------------------------------------------------
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

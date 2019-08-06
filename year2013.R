library(readxl)
library(tidyverse)
library(stringi)
library(readr)
library(sqldf)

# help match --------------
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

# team ID
teamid <- read_csv("teamid.csv") %>% 
  rename(
         "TEAMID" = "Team: ID",
         "EXTERNAL_PROPOSAL_ID__C" = "External Proposal ID", 
         "PROPOSAL__C" = "Proposal: ID")

# proposal -------------------------------
proposal_2013 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
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
                                            ifelse(`Institution Name` == "Cogswell Polytechnical College", "Cogswell College",
                                                   ifelse(`Institution Name` == "Arizona State University at the Tempe Campus", "Arizona State University", `Institution Name`)))
  ) %>% 
  select(
    year, NAME, AMOUNT_REQUESTED__C, PROPOSAL_NAME_LONG_VERSION__C, APPLYING_INSTITUTION_NAME__C,
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
  select( - `Zenn ID`, -year)

proposal_2013$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C <- str_replace_all(proposal_2013$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, "[:cntrl:]", " ")

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
  write_csv("new/team_2013.csv")

# membership -----------------------------
contacts <- read_csv("/Volumes/GoogleDrive/My Drive/Sustainable_Vision/salesforce_examples/Contact_Extract.csv") %>% 
  select(ID, EMAIL, NPE01__ALTERNATEEMAIL__C, NPE01__HOMEEMAIL__C,
         NPE01__WORKEMAIL__C, PREVIOUS_EMAIL_ADDRESSES__C, BKUP_EMAIL_ADDRESS__C)

contacts_1 <- contacts %>% 
  select(ID, EMAIL)

advisors <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_advisors.xlsx") %>% 
  select(`Zenn ID`, `Team Role`, Email) %>% 
  rename("ZENN_ID__C" = "Zenn ID",
         "ROLE__C" = "Team Role",
         "EMAIL" = "Email")

proposal_2013_narrow <- proposal_2013 %>% 
  select(NAME, ZENN_ID__C, EXTERNAL_PROPOSAL_ID__C, PROGRAM_COHORT__C, RECORDTYPEID) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") %>% 
  left_join(teamid, by = "EXTERNAL_PROPOSAL_ID__C")
  
membership_2013 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
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
  left_join(proposal_2013_narrow) %>% 
  rename(     
    "TEAM__C" = "TEAMID",
    "PROGRAM_COHORT_LOOKUP__C" = "PROGRAM_COHORT__C",
    "START_DATE__C" = "Actual Period Begin",
    "END_DATE__C" = "Actual Period End"
  ) %>% 
  right_join(advisors) %>% 
  select(
    EMAIL, ZENN_ID__C, 
    TEAM__C, PROPOSAL__C, PROGRAM_COHORT_LOOKUP__C, 
    ROLE__C, STATUS__C, START_DATE__C, END_DATE__C, RECORDTYPEID
  ) %>% 
  mutate(ROLE__C = ifelse(ROLE__C == "Dean of Faculty", "Dean", ROLE__C)) %>% 
  left_join(contacts_1) %>% 
  rename("MEMBER__C" = "ID") %>% 
  na.omit() %>% 
  write_csv("new/member_2013.csv")
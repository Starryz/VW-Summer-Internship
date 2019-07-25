library(readxl)
library(tidyverse)
library(readr)

# commits 
commits_2013 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
  filter(`Application Status` == "funded") %>% 
  rename(
    "GRANT_STATUS__C" = "Application Status",
    "GRANT_ID__C" = "External Proposal ID", 
    "GRANTED_INSTITUTION__C" = "Institution Name",
    "AMOUNT_DISBURSED__C" = "Amount Disbursed",
    "PROGRAM__C" = "Type"
         ) %>% 
  mutate(
    "DISBURSEMENT_REQUEST_AMOUNT__C" = as.double(AMOUNT_DISBURSED__C),
    "AWARD_LETTER_SENT__C" = as.Date(`Grant Letter Sent`),
    "AWARD_LETTER_SIGNED__C" = as.Date(`Grant Letter Signed`),
    "GRANT_START_DATE__C" = as.Date(`Actual Period Begin`),
    "GRANT_END_DATE__C" = as.Date(`Actual Period End`),
    "PAYMENT_STATUS__C" = "Paid", 
    "AMOUNT_APPROVED__C" = as.double(`Amount Approved`)
    ) %>% 
  select(
    `GRANT_ID__C`, `AMOUNT_APPROVED__C`, `AWARD_LETTER_SENT__C`, `AWARD_LETTER_SIGNED__C`, 
         `PAYMENT_STATUS__C`, `GRANT_START_DATE__C`, `PROGRAM__C`, `GRANT_END_DATE__C`, 
         `GRANT_STATUS__C`, `GRANTED_INSTITUTION__C`, `DISBURSEMENT_REQUEST_AMOUNT__C`
    ) 

# proposal 
proposal_2013 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
  rename(
    "NAME" = "Grant Title",
    "APPLYING_INSTITUTION_NAME__C" = "Institution Name",
    "PROGRAM_COHORT_RECORD_TYPE__C" = "Type",
    "PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C" = "Proposal Summary",
    "STATUS__C" = "Application Status", 
    "EXTERNAL_PROPOSAL_ID__C" = "External Proposal ID"
    ) %>% 
  mutate(
    "PROPOSAL_NAME_LONG_VERSION__C" = as.character(NAME),
    "DATE_CREATED__C" = as.Date(`Date Created`),
    "DATE_SUBMITTED__C" = as.Date(`Date Application Submitted`),
    "GRANT_PERIOD_END__C" = as.Date(`Actual Period End`),
    "GRANT_PERIOD_START__C" = as.Date(`Actual Period Begin`),
    "PROGRAM__C" = as.character(PROGRAM_COHORT_RECORD_TYPE__C),
    "AMOUNT_REQUESTED__C" = as.double(`Amount Requested`),
    "ZENN_ID__C" = as.double(`Zenn ID`),
    "AWARD_AMOUNT__C" = as.double(`Amount Approved`)
    ) %>% 
  select(
    NAME, AMOUNT_REQUESTED__C, PROPOSAL_NAME_LONG_VERSION__C, APPLYING_INSTITUTION_NAME__C,
    AWARD_AMOUNT__C, DATE_CREATED__C, DATE_SUBMITTED__C, GRANT_PERIOD_END__C, 
    GRANT_PERIOD_START__C, PROGRAM_COHORT_RECORD_TYPE__C, 
    PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, ZENN_ID__C, STATUS__C, PROGRAM__C,
    EXTERNAL_PROPOSAL_ID__C
    ) 
## needs to change proposal summary 

# team 
team_2013 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
  rename(
    "NAME" = "Grant Title"
    ) %>% 
  mutate(
    "END_DATE_OF_FIRST_PROGRAM_COMPLETED__C" = as.Date(`Actual Period End`),
    "PROPOSAL_TOTAL_FUNDED_AWARD_AMOUNT__C" = as.double(`Amount Approved`)
    ) %>% 
  select(
    NAME, PROPOSAL_TOTAL_FUNDED_AWARD_AMOUNT__C, END_DATE_OF_FIRST_PROGRAM_COMPLETED__C
    ) 

# team membership export 
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
    ) 
## note: status needs capitalization 
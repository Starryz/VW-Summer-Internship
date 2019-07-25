library(readxl)
library(tidyverse)

# commits export
commits_2013 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
  rename("GRANT_STATUS__C" = "Application Status") %>% 
  rename("GRANT_ID__C" = "External Proposal ID") %>% 
  filter(`GRANT_STATUS__C` == "funded") %>% 
  rename("GRANTED_INSTITUTION__C" = "Institution Name") %>% 
  rename("AMOUNT_APPROVED__C" = "Amount Approved") %>% 
  rename("AMOUNT_DISBURSED__C" = "Amount Disbursed") %>% 
  rename("PROGRAM__C" = "Type") %>% 
  mutate("DISBURSEMENT_REQUEST_AMOUNT__C" = as.numeric(AMOUNT_DISBURSED__C)) %>% 
  mutate("AWARD_LETTER_SENT__C" = as.Date(`Grant Letter Sent`)) %>% 
  mutate("AWARD_LETTER_SIGNED__C" = as.Date(`Grant Letter Signed`)) %>% 
  mutate("GRANT_START_DATE__C" = as.Date(`Actual Period Begin`)) %>% 
  mutate("GRANT_END_DATE__C" = as.Date(`Actual Period End`)) %>% 
  mutate("PAYMENT_STATUS__C" = "Paid") %>% 
  select(`GRANT_ID__C`, `AMOUNT_APPROVED__C`, `AWARD_LETTER_SENT__C`, `AWARD_LETTER_SIGNED__C`, 
         `PAYMENT_STATUS__C`, `GRANT_START_DATE__C`, `PROGRAM__C`, `GRANT_END_DATE__C`, 
         `GRANT_STATUS__C`, `GRANTED_INSTITUTION__C`, `DISBURSEMENT_REQUEST_AMOUNT__C`)

# proposal export
proposal_2013 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
  rename("NAME" = "Grant Title") %>% 
  rename("AMOUNT_REQUESTED__C" = "Amount Requested") %>% 
  mutate("PROPOSAL_NAME_LONG_VERSION__C" = as.character(NAME)) %>% 
  rename("APPLYING_INSTITUTION_NAME__C" = "Institution Name") %>% 
  rename("AWARD_AMOUNT__C" = "Amount Approved") %>% 
  mutate("DATE_CREATED__C" = as.Date(`Date Created`)) %>% 
  mutate("DATE_SUBMITTED__C" = as.Date(`Date Application Submitted`)) %>% 
  mutate("GRANT_PERIOD_END__C" = as.Date(`Actual Period End`)) %>% 
  mutate("GRANT_PERIOD_START__C" = as.Date(`Actual Period Begin`)) %>% 
  rename("PROGRAM_COHORT_RECORD_TYPE__C" = "Type") %>% 
  mutate("PROGRAM__C" = as.character(PROGRAM_COHORT_RECORD_TYPE__C)) %>% 
  rename("PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C" = "Proposal Summary") %>% 
  rename("ZENN_ID__C" = "Zenn ID") %>% 
  rename("STATUS__C" = "Application Status") %>% 
  select(NAME, AMOUNT_REQUESTED__C, PROPOSAL_NAME_LONG_VERSION__C, APPLYING_INSTITUTION_NAME__C,
         AWARD_AMOUNT__C, DATE_CREATED__C, DATE_SUBMITTED__C, GRANT_PERIOD_END__C, 
         GRANT_PERIOD_START__C, PROGRAM_COHORT_RECORD_TYPE__C, 
         PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, ZENN_ID__C, STATUS__C, PROGRAM__C)
  
# team export
team_2013 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
  rename("NAME" = "Grant Title") %>% 
  rename("PROPOSAL_TOTAL_FUNDED_AWARD_AMOUNT__C" = "Amount Approved") %>% 
  mutate("END_DATE_OF_FIRST_PROGRAM_COMPLETED__C" = as.Date(`Actual Period End`)) %>% 
  select(NAME, PROPOSAL_TOTAL_FUNDED_AWARD_AMOUNT__C, END_DATE_OF_FIRST_PROGRAM_COMPLETED__C)

# team membership export 
membership_2013_1a <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
  mutate("START_DATE__C" = as.Date(`Actual Period Begin`)) %>% 
  mutate("END_DATE__C" = as.Date(`Actual Period End`)) %>% 
  rename(
    "TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "Grant Title",
    "PROPOSAL_STATUS__C" ="Application Status") %>% 
  select(START_DATE__C, END_DATE__C, TEAM_NAME_TEXT_ONLY_HIDDEN__C, PROPOSAL_STATUS__C)

membership_2013_1b <- proposal_2013 %>% 
  select(NAME, ZENN_ID__C) %>% 
  rename("TEAM_NAME_TEXT_ONLY_HIDDEN__C" = "NAME") 

membership_2013_2 <- merge(membership_2013_1a, membership_2013_1b) # add ZENN ID

advisors <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_advisors.xlsx") %>% 
  rename("ZENN_ID__C" = "Zenn ID")

membership_2013 <- merge(membership_2013_2, advisors) %>% 
  mutate("START_DATE__C" = as.Date(`START_DATE__C`)) %>% 
  mutate("END_DATE__C" = as.Date(`END_DATE__C`)) %>% 
  unite("FULL_NAME__C", c(`First Name`, `Last Name`), sep = " ", remove = FALSE) %>% 
  mutate("PROGRAM_TYPE_FORMULA__C" = "Sustainable Vision") %>% 
  rename(
    "ROLE__C" = "Team Role",
    "EMAIL_FORMULA__C" = Email,
    "PHONE_FORMULA__C" = `Telephone 1`,
    "FIRST_NAME__C" = `First Name`,
    "LAST_NAME__C" = `Last Name`,
    "ORGANIZATION__C" = Organization
  ) %>% 
  select(ROLE__C, START_DATE__C, END_DATE__C, FULL_NAME__C, EMAIL_FORMULA__C, 
         PHONE_FORMULA__C, PROGRAM_TYPE_FORMULA__C, FIRST_NAME__C, ORGANIZATION__C,
         PROPOSAL_STATUS__C, LAST_NAME__C, TEAM_NAME_TEXT_ONLY_HIDDEN__C)
# note: status needs capitalization 
  

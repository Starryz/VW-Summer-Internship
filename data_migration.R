library(readxl)
library(tidyverse)

# proposal to commit
proposal_2013 <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2013_proposals.xlsx") %>% 
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
  select(`GRANT_ID__C`, `AMOUNT_APPROVED__C`, `AWARD_LETTER_SENT__C`, `AWARD_LETTER_SIGNED__C`, `PAYMENT_STATUS__C`,
          `GRANT_START_DATE__C`, `PROGRAM__C`, `GRANT_END_DATE__C`, `GRANT_STATUS__C`, 
         `GRANTED_INSTITUTION__C`, `DISBURSEMENT_REQUEST_AMOUNT__C`)

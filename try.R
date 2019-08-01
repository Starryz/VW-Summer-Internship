# proposal -------------------------------
goal <- read_excel("~/Desktop/Sustainable_Vision/sustainable_vision_grants_2012_proposals.xlsx") %>% 
  rename(
    "NAME" = "Grant Title",
    "PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C" = "Proposal Summary",
    "EXTERNAL_PROPOSAL_ID__C" = "External Proposal ID" 
  ) %>%  
  select(PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C) 


str_replace_all(string = goal$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C[1],
                  pattern  = "\\\\u",
                replacement = "GJAKLUTLIAUgautralkurlkauwelkrukla") %>% 
  write("try.txt")
  
  
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
    "APPLYING_INSTITUTION_NAME__C" = ifelse(`Institution Name` == "University of Oklahoma", "University of Oklahoma Norman Campus",`Institution Name`)
  ) %>% 
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

proposal_2012$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C <- str_replace_all(proposal_2012$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, "[:cntrl:]", " ")

#str_remove_all("\u0093systems\u0094", "[[\\[u]+[0-9]*]]")

proposal_2012 <- sapply(proposal_2012, as.character)
proposal_2012[is.na(proposal_2012)] <- " "
proposal_2012 <- as.data.frame(proposal_2012)

write_csv(proposal_2012, "new/2012/proposal_2012.csv")


# team --------------------------
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

# note_task -------------------------------------------------
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
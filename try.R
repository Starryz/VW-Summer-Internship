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

text <- "world\u0092s"
text1 <- text %>% str_replace_all(
  patter = "\\\\u", 
  replacement = "haha"
)

proposal_2012$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C <- 
  str_replace_all(proposal_2012$PROJECT_DESCRIPTION_PROPOSAL_ABSTRACT__C, "[:cntrl:]", " ")

library(tidyverse)
library(readxl)
library(writexl)

#manually setting the excel sheet numbers
sheets <- list('5001', '5007', '5010', '5013', '5016', '5019', '5022') 

#initializing the for loop for each sheet
for (named in sheets) {
  
  #creating another list to loop to build a nested forloop 
  final_df <- list()
  over_unders <- seq(.5, 30, by = .5)
  
  #initializing the for loop for each over/under
  for (phy in over_unders) {
    
    #reading in data and with minor column name edits
    data <- read_excel("/Users/marfi/Downloads/ALL_LN_DATA.xlsx", sheet = named) %>%
      rename(
        'prj_tm_twp' = 'prj_hm_tm_twp',
        'output' = 'stat') %>%
      select(-ou)
    
    #creating a team win percentage column, a spread column generated from the win percentage difference, as well as classifying when the desired statistic went over or under this iteration of over/under
    u_data <- data %>%
      mutate(
        opp_tm_twp = 1 - prj_tm_twp,
        spread =  opp_tm_twp - prj_tm_twp,
        ou = phy,
        stat = case_when(output > ou ~ 1,
                         output < ou ~ 0,
                         output == ou ~ NA)
      ) %>%
      select(opp_tm_twp,
             prj_tm_twp, 
             spread, 
             prj_gm_ou, 
             stat) %>%
      filter(!is.na(stat)) # filtering out pushes, only want observations where the statistic went cleanly over/under
    
    
    #generating a sequence for each previous iteration to then loop through all possible combinations of total over unders for the entire game and the project win percentage for the home team
    
    prj_gm_over_unders <- seq(20, 100, by = .5)
    temp <- list()
    for (x in prj_gm_over_unders) {
      prj_tm_twp_seq <- seq(0, 1, by = .01)
      
      current <- expand.grid(prj_gm_ou = x,
                             prj_tm_twp = prj_tm_twp_seq)
      
      temp[[length(temp) + 1]] <- current
    }
    totals_u <- bind_rows(temp)
    
    # creating all necessary variables to match to the u_data within the totals_u dataframe
    totals_u$opp_tm_twp <- 1 - totals_u$prj_tm_twp
    totals_u$spread <- totals_u$opp_tm_twp - totals_u$prj_tm_twp
    totals_u <- totals_u %>%
      select(prj_tm_twp, prj_gm_ou, spread, opp_tm_twp)
    
    #Logistic regression with the spread as an interactive term against all of the predictors trained on the u_data dataframe
    lm <- glm(stat ~ spread*prj_tm_twp + spread*prj_gm_ou +spread*opp_tm_twp, data = u_data, family = binomial)
    
    #using the trained model to predict the combinations we generated earlier for the totals_u dataframe
    predicted_probability <- predict(lm, newdata = totals_u, type = "response")
    predicted_probability<- format(predicted_probability, scientific = FALSE)
    all_probabilities <- cbind(totals_u, predicted_probability)
    
    #comparing this to the probabilities generated from a simple binomial model
    sta_model <- glm(stat ~ ., data = u_data, family = binomial)
    
    #run the simple model on the totals_u dataframe
    predicted_stat <- predict(sta_model, totals_u)
    
    #compare the probabilities of both models for the given over/under
    test <- bind_cols(all_probabilities,
                      predicted_stat = round(predicted_stat, digits = 4),
                      ou = phy) %>%
      select(predicted_probability, predicted_stat, prj_tm_twp, prj_gm_ou)
    
    #finalize dataframe
    test$cid <- named
    test$ou <- phy
    final_df[[length(final_df) + 1]] <- test
  }
  final_df <- bind_rows(final_df)
  
  #export dataframes for all sheets
  write_xlsx(final_df, paste0("/Users/marfi/Downloads/", named, "Win_Percentage.xlsx"))
}
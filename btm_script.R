rm(list = ls())

# Load required packages
library(foreign)
library(readxl)
library(rio)
library(dplyr)
library(tidyr)
library(reshape2)
library(QuantPsyc)
library(pastecs)
library(tidyverse)
library(semPLS)

##########################
### surveyCleaning_SEM ###
##########################
user <- Sys.getenv('username')
#raw <- read.spss(paste0("C:/Users/", user, "/OneWorkplace/Decisions Science - 3. Audience and Agility/Client Work/Barclays/Barclays BTM/2022/Q4 campaign/3. Raw data/survey/Final data/Barclays Q4 Final data SPSS.sav")) 
#raw <- data.frame(raw)

#raw <- read.csv(paste0("C:/Users/", user, "/OneWorkplace/Decisions Science - 3. Audience and Agility/Client Work/Barclays/Barclays BTM/2022/Q4 campaign/3. Raw data/survey/Final data/final_data.csv"))
raw <- read.csv(paste0("C:/Users/", user, "/OneWorkplace/Decisions Science - Documents/3. Audience and Agility/Client Work/Barclays/Barclays BTM/2022/Q4 campaign/3. Raw data/survey/Final data/final_data.csv"))


#setwd(paste0("C:/Users/", user, "/OneWorkplace/Decisions Science - 3. Audience and Agility/Client Work/Barclays/Barclays BTM/2022/Q4 campaign/5. Formatted data/"))
setwd(paste0("C:/Users/", user, "/OneWorkplace/Decisions Science - Documents/3. Audience and Agility/Client Work/Barclays/Barclays BTM/2022/Q4 campaign/5. Formatted data/"))

audiences <- c("comfortablyCommitted", "ActiveStrivers", "day2dayPlanners", "cherryPickers") # all adults and catalysts too?

survey <- raw %>% 
  mutate_at(vars("Q6_5"),~zoo::na.fill(., 0)) %>% # convert NA's to 0 for selected columns
  mutate(rejector = ifelse(Q3_3 == "I would not consider it" | Q3_3 == 3, 1, 0),# Rejector
         barclaysCustomer = ifelse((hBarclaysCustomer == 1) | (Q6_5 == 1), 1, 0), # Barclays Customer, Q6_5 is barclaycard
         allAdults = 1, # All Adults
         footballFan = ifelse((Q9 == 1) | (Q10 == 1), 1, 0), # Additional group
         id = raw[[1]]) %>% # ID
  rename_with(~ audiences, all_of(c("Available_1", "Available_2", "Available_3", "Available_4"))) %>%
  mutate_at(vars(audiences), funs(dig = ifelse((. == 1) & (S12 == 1), 1, 0)))

### Clean Columns ###
survey[survey == "Selected"] <- 1
survey[survey == "Not selected"] <- 0
survey[survey == "Yes"] <- 1
survey[survey == "No"] <- 0
survey[survey == "Barclays customer"] <- 1
survey[survey == "Not a Barclays customer"] <- 0
survey[survey == "It would be my first choice"] <- 1
survey[survey == "I would seriously consider it"] <- 2
survey[survey == "I might consider it"] <- 3
survey[survey == "I would not consider it"] <- 4
survey[survey == "I feel happiest doing my banking digitally via a mobile app or the bank website most of the time"] <- 1
survey[survey == "I feel happiest doing my banking with a bit of human contact via the branch or telephone banking most of the time"] <- 2
survey[survey == "I'm happy to bank digitally but sometimes appreciate the human contact of a branch or telephone banking"] <- 3
survey[survey == "10- Completely agree"] <- 10
survey[survey == "Don't know"] <- 11
survey[survey == "1- Completely disagree"] <- 1
survey[survey == "Can't remember"] <- 3

survey[survey == "East of England"] <- 1
survey[survey == "East Midlands"] <- 2
survey[survey == "London"] <- 3
survey[survey == "North East"] <- 4
survey[survey == "North West"] <- 5
survey[survey == "Northern Ireland"] <- 6
survey[survey == "Scotland"] <- 7
survey[survey == "South East"] <- 8
survey[survey == "South West"] <- 9
survey[survey == "Wales"] <- 10
survey[survey == "West Midlands"] <- 11
survey[survey == "Yorkshire and The Humber"] <- 12

survey[survey == "Never watch this channel/ Don't know"] <- 1
survey[survey == "Less than one hour"] <- 2
survey[survey == "1-2 hours"] <- 3
survey[survey == "3-4 hours"] <- 4
survey[survey == "5-6 hours"] <- 5
survey[survey == "7 hours or more"] <- 6

survey[survey == "6:00am - 9:30am"] <- 1
survey[survey == "9:30am-12:00 noon"] <- 1
survey[survey == "12:00 noon - 4:00pm"] <- 1
survey[survey == "4:00pm - 5:30pm"] <- 1
survey[survey == "5:30pm - 8:00pm"] <- 1
survey[survey == "8:00pm - 11:00pm"] <- 1
survey[survey == "11:00pm - 1:00am"] <- 1
survey[survey == "1:00am - 6:00am"] <- 1

survey[survey == "Never/NA"] <- 1
survey[survey == "Less than once a month"] <- 2
survey[survey == "A few times a month"] <- 3
survey[survey == "A few times a week"] <- 4
survey[survey == "About once a day"] <- 5
survey[survey == "Multiple times every day"] <- 6

survey[survey == "More than once a week"] <- 1
survey[survey == "Once a week"] <- 2
survey[survey == "Once every 2 weeks"] <- 3
survey[survey == "Once a month"] <- 4
survey[survey == "Once every 3 months"] <- 5
survey[survey == "Once every 6 months"] <- 6
survey[survey == "Once every 12 months"] <- 7
survey$M8[survey$M8 == "Less often"] <- 8  
survey[survey == "I never go to the cinema"] <- 9

survey[survey == "Less often"] <- 2
survey[survey == "At least once a month"] <- 3
survey[survey == "At least once a week"] <- 4
survey[survey == "At least once a day"] <- 5

survey[survey == "Today"] <- 1
survey[survey == "Yesterday"] <- 2
survey[survey == "2-3 days ago"] <- 3
survey[survey == "4-6 days ago"] <- 4
survey[survey == "A week ago"] <- 5
survey[survey == "2 weeks ago"] <- 6
survey[survey == "3 weeks ago"] <- 7
survey[survey == "A month ago"] <- 8
survey[survey == "More than a month ago"] <- 9

survey[survey == "Past 7 days"] <- 1
survey[survey == "Past 4 weeks"] <- 2
survey[survey == "Longer Ago"] <- 3
survey[survey == "Never / NA"] <- 4

survey2 <- survey %>% mutate_at(c('hBarclaysCustomer','Q3_3','Q6_5','S12','Q9','Q10'), as.numeric)

# perception questions to numeric
survey2[,319:446][is.na(survey2[,319:446])] <- 0
survey2[,c(319:446)] <- apply(survey2[,c(319:446)], 2,function(x) as.numeric(as.character(x)))



### Add brand perception questions ###
perceptions <- c("has_customer_interests_at_heart",
                 "has_good_reputation",
                 "recognises_and_rewards_customers",
                 "is_setting_the_lead",
                 "offers_better_value_for_money_than_others",
                 "offers_innovative_solutions_to_make_your_life_easier",
                 "is_more_trustworthy_than_others",
                 "is_easy_to_deal_with",
                 "is_flexible_when_you_need_them",
                 "has_competitive_offers",
                 "is_changing_banking_for_the_better",
                 "helps_you_to_stay_in_control_of_your_money",
                 "helps_you_to_make_the_most_of_your_money",
                 "listens_and_responds_to_their_customer_needs",
                 "makes_you_feel_secure_and_protected",
                 "helps_you_feel_more_positive_about_your_money",
                 "offers_consistently_high_quality_products_and_services",
                 "is_a_brand_you_see_and_hear_a_lot_about",
                 "is_a_brand_i_know_a_lot_about",
                 "is_a_socially_responsible_brand",
                 "is_an_honest_brand",
                 "is_a_brand_that_has_been_around_a_long_time",
                 "has_good_customer_service",
                 "is_more_helpful_than_other_banks",
                 "is_a_brand_i_would_consider_for_my_future_financial_needs",
                 "is_a_brand_i_would_consider_takingout_a_financial_product_with",
                 "is_a_brand_i_intend_to_take_out_a_financial_product_with",
                 "has_a_good_mobile_app_experience",
                 "provides_a_good_in_branch_experience",
                 "is_a_brand_i_would_consider_for_an_investment_product",
                 "is_a_brand_i_would_consider_for_a_savings_product",
                 "is_a_brand_i_would_consider_for_a_mortgage")

# Bind all B1 columns. B1 is the bank of perception questions about Barclays. 
# The letter refers to the audience and the number is the question.
survey2$B1_1 <-  (survey2$B1_A_1 + survey2$B1_B_1 + survey2$B1_C_1 + survey2$B1_D_1)
survey2$B1_2 <-  (survey2$B1_A_2 + survey2$B1_B_2 + survey2$B1_C_2 + survey2$B1_D_2)
survey2$B1_3 <-  (survey2$B1_A_3 + survey2$B1_B_3 + survey2$B1_C_3 + survey2$B1_D_3)
survey2$B1_4 <-  (survey2$B1_A_4 + survey2$B1_B_4 + survey2$B1_C_4 + survey2$B1_D_4)
survey2$B1_5 <-  (survey2$B1_A_5 + survey2$B1_B_5 + survey2$B1_C_5 + survey2$B1_D_5)
survey2$B1_6 <-  (survey2$B1_A_6 + survey2$B1_B_6 + survey2$B1_C_6 + survey2$B1_D_6)
survey2$B1_7 <-  (survey2$B1_A_7 + survey2$B1_B_7 + survey2$B1_C_7 + survey2$B1_D_7)
survey2$B1_8 <-  (survey2$B1_A_8 + survey2$B1_B_8 + survey2$B1_C_8 + survey2$B1_D_8)
survey2$B1_9 <-  (survey2$B1_A_9 + survey2$B1_B_9 + survey2$B1_C_9 + survey2$B1_D_9)

survey2$B1_10 <-  (survey2$B1_A_10 + survey2$B1_B_10 + survey2$B1_C_10 + survey2$B1_D_10)
survey2$B1_11 <-  (survey2$B1_A_11 + survey2$B1_B_11 + survey2$B1_C_11 + survey2$B1_D_11)
survey2$B1_12 <-  (survey2$B1_A_12 + survey2$B1_B_12 + survey2$B1_C_12 + survey2$B1_D_12)
survey2$B1_13 <-  (survey2$B1_A_13 + survey2$B1_B_13 + survey2$B1_C_13 + survey2$B1_D_13)
survey2$B1_14 <-  (survey2$B1_A_14 + survey2$B1_B_14 + survey2$B1_C_14 + survey2$B1_D_14)
survey2$B1_15 <-  (survey2$B1_A_15 + survey2$B1_B_15 + survey2$B1_C_15 + survey2$B1_D_15)
survey2$B1_16 <-  (survey2$B1_A_16 + survey2$B1_B_16 + survey2$B1_C_16 + survey2$B1_D_16)
survey2$B1_17 <-  (survey2$B1_A_17 + survey2$B1_B_17 + survey2$B1_C_17 + survey2$B1_D_17)
survey2$B1_18 <-  (survey2$B1_A_18 + survey2$B1_B_18 + survey2$B1_C_18 + survey2$B1_D_18)
survey2$B1_19 <-  (survey2$B1_A_19 + survey2$B1_B_19 + survey2$B1_C_19 + survey2$B1_D_19)

survey2$B1_20 <-  (survey2$B1_A_20 + survey2$B1_B_20 + survey2$B1_C_20 + survey2$B1_D_20)
survey2$B1_21 <-  (survey2$B1_A_21 + survey2$B1_B_21 + survey2$B1_C_21 + survey2$B1_D_21)
survey2$B1_22 <-  (survey2$B1_A_22 + survey2$B1_B_22 + survey2$B1_C_22 + survey2$B1_D_22)
survey2$B1_23 <-  (survey2$B1_A_23 + survey2$B1_B_23 + survey2$B1_C_23 + survey2$B1_D_23)
survey2$B1_24 <-  (survey2$B1_A_24 + survey2$B1_B_24 + survey2$B1_C_24 + survey2$B1_D_24)
survey2$B1_25 <-  (survey2$B1_A_25 + survey2$B1_B_25 + survey2$B1_C_25 + survey2$B1_D_25)
survey2$B1_26 <-  (survey2$B1_A_26 + survey2$B1_B_26 + survey2$B1_C_26 + survey2$B1_D_26)
survey2$B1_27 <-  (survey2$B1_A_27 + survey2$B1_B_27 + survey2$B1_C_27 + survey2$B1_D_27)
survey2$B1_28 <-  (survey2$B1_A_28 + survey2$B1_B_28 + survey2$B1_C_28 + survey2$B1_D_28)
survey2$B1_29 <-  (survey2$B1_A_29 + survey2$B1_B_29 + survey2$B1_C_29 + survey2$B1_D_29)

survey2$B1_30 <-  (survey2$B1_A_30 + survey2$B1_B_30 + survey2$B1_C_30 + survey2$B1_D_30)
survey2$B1_31 <-  (survey2$B1_A_31 + survey2$B1_B_31 + survey2$B1_C_31 + survey2$B1_D_31)
survey2$B1_32 <-  (survey2$B1_A_32 + survey2$B1_B_32 + survey2$B1_C_32 + survey2$B1_D_32)

# Remove old section split questions
survey2 <- survey2[,!(grepl("_A_", colnames(survey2))|grepl("_B_", colnames(survey2))|grepl("_C_", colnames(survey2))|grepl("_D_", colnames(survey2)))]

# Rename perception questions
colnames(survey2)[grepl("B1_", colnames(survey2))] <- perceptions



### Finalise columns ###
# Select only relevant columns
perceptionDF <- survey2[, c("id",
                            "Response_duration_seconds",
                            "rejector", "barclaysCustomer",
                            "allAdults",
                            audiences,
                            perceptions)]

# Run checks - NA's
nas <- sum(rowSums(is.na(survey2[perceptions])) == length(perceptions)) # should be 0
survey2 <- survey2[rowSums(is.na(survey2[perceptions])) != length(perceptions),]
if (nas > 0){
  print(paste0("There are ", nas, " NA's in the dataset. These have been removed //n"))
}


### Straightliners
straightlinerLimit = 0.8

freq <- apply(survey2[,perceptions], 1, table)
straightliners <- bind_rows(lapply(freq, as.data.frame.list))
straightliners <- straightliners[,order(names(straightliners))]
straightliners[is.na(straightliners)] <- 0
straightliners <- as.data.frame(straightliners)
straightliners <- straightliners/rowSums(straightliners)
straightliners['id'] <- survey2["id"]

straightliners <- straightliners %>%
  filter_at(vars(X1,X2,X3,X4,X5,X6,X7,X8,X9,X10,X11),any_vars(.>=straightlinerLimit))

if (nrow(straightliners) > 0){
  print(paste0("There are ", nrow(straightliners), " straightliners in the dataset. These have been removed"))
}

survey3 <- anti_join(survey2, straightliners, by = c("id"))


### Speedy Responders
speedingMins = 5
speeders <- survey3[survey3$Response_duration_seconds < speedingMins*60,]
#survey <- survey[survey$Response_duration_seconds >= speedingMins,]
survey3 <- survey3 %>% filter(Response_duration_seconds >= speedingMins*60)

if (nrow(speeders) > 0){
  print(paste0("There are ", nrow(speeders), " speedy responders in the dataset. These have been removed"))
}

# Format Don't know, completely disagree and Completely Agree answers
survey3[,perceptions] <- lapply(survey3[,perceptions], function(x){
  x[x==11] <- mean(as.numeric(x[x!=11]))
  x <- as.numeric(x)
})


### Save files ###
for (aud in c("rejector", "barclaysCustomer", "allAdults", audiences)){
  # Save out with ID column
  savFile <- survey3 %>%
    filter_(paste(aud, "== 1")) %>%
    dplyr::select(all_of(c("id", perceptions)))
  
  write.csv(savFile,
            paste0("SEM/", paste0("SEM_", aud,"_ID.csv")),
            row.names = F)
  
  # Save out without ID column
  savFile <- savFile %>%
    dplyr::select(all_of(perceptions))
  getwd
  write.csv(savFile,
            paste0("SEM/", paste0("SEM_", aud,".csv")),
            row.names = F)
}

### Save out sample sizes ###
sampleSizes <- survey3 %>%
  dplyr::select(c("rejector", "barclaysCustomer", "allAdults", "footballFan", audiences)) %>%
  mutate_if(is.character, as.numeric) %>%
  summarise_all(sum) %>%
  t() %>% 
  data.frame() %>%
  rename("sample_size" = ".")

         
write.csv(sampleSizes,paste0("SEM/sampleSizes.csv"),row.names = T)
#write.csv(perceptionDF, "perceptions.csv", row.names = F)
#write.csv(survey3, "survey_v4.csv", row.names = F)







###################
### Brand Cloud ###
###################
brandCloud <- function(audiences){
  
  # Bring in brand cloud structure
  # this is the outer network (so nodes to nodes)
  strucmod <-  read_excel("Brand Cloud/ModelFile.xlsx",
                         sheet = "Edges")
  
  # this is what makes up each latent (statement to latent) 
  measuremod <- read_excel("Brand Cloud/ModelFile.xlsx",
                          sheet = "MeasurementModel")
  
  structuralmodel <- as.matrix(strucmod)
  measurementmodel <- as.matrix(measuremod)
  
  
  # For each audience, create the brand cloud using the new data and outputs files
  for(aud in audiences){
    
    # read in your data
    data <- read.csv(paste0("SEM/SEM_", aud, ".csv"))
    
    SEM <-
      plsm(data = data,
           strucmod = structuralmodel,
           measuremod = measurementmodel)
    
    outputs <-
      sempls(model = SEM,
             data = data,
             wscheme = "centroid")
    
    
    # bootstrapping
    ecsiBoot <-
      bootsempls(
        outputs,
        nboot = nrow(data),
        start = "ones",
        verbose = FALSE
      )
    
    
    # total effects
    totaleffectsout <- as.data.frame(outputs[["total_effects"]])
    
    
    # path coefficients
    PCout <- as.data.frame(outputs[["path_coefficients"]])
    
    # r squared
    df <- matrix(c("outputs", rSquared2),
                 nrow = 1,
                 ncol = 2)
    df <- data.frame(df)
    testres <- df$X2[[1]](outputs)
    testres <- round(testres, 2)
    rownames(testres) <-
      gsub("%nonlatents", "", rownames(testres))
    
    
    # weights
    weightsout <- as.data.frame(outputs[["outer_weights"]])
    
    # loadings
    outload <- as.data.frame(outputs[["outer_loadings"]])
    
    
    ### create your p-value bootstrap estimates
    replicatestats <- as.data.frame(ecsiBoot$t)
    obsvalue <- as.data.frame(ecsiBoot$t0)
    
    rownames(obsvalue) <- colnames(replicatestats)
    
    #replicatestats <-
    #  replicatestats %>% select(starts_with("beta"))
    replicatestats <- replicatestats[,grepl("beta", colnames(replicatestats))]

    obsvalue <- as.data.frame(t(obsvalue))
    #obsvalue <- obsvalue %>% select(starts_with("beta"))
    obsvalue <- obsvalue[,grepl("beta", colnames(obsvalue))]
    
    relations <-
      as.data.frame(ecsiBoot[["fitted_model"]][["model"]][["strucmod"]])
    relations$path <- colnames(replicatestats)
    
    variables <- colnames(replicatestats)
    
    teststat <- NA
    
    for (i in (variables)) { 
      b3.under.H0 <- replicatestats[, i] - mean(replicatestats[, i])
      pvalue <- mean(abs(b3.under.H0) > abs(obsvalue[1, i]))
      #path <- i
      set <- cbind(i, pvalue)
      teststat <- rbind(teststat, set)
    }
    
    teststat <- as.data.frame(teststat)
    relations <- merge(relations, teststat)
    
    relations <- relations %>% dplyr::select(c('source', 'target', 'pvalue'))
    #relations <- select(relations, source, target, pvalue)
    relations$source <- str_replace_all(relations$source, "%nonlatents", "")
    relations$target <- str_replace_all(relations$target, "%nonlatents", "")
    relations$pvalue <- as.numeric(as.character(relations$pvalue))
    relations$path <- paste(relations$source, " -> ", relations$target)
    relations <- relations %>% dplyr::select(c('path', 'pvalue'))
    
    
    
    # create outputs list
    list_of_datasets <- list(
      "Total Effects" = totaleffectsout,
      "Bootstrap" = relations,
      "Path Coefficients" = PCout,
      "Outer Loadings" = outload,
      "Outer Weights" = weightsout,
      "R Squared" = testres
    )
    
    # write the output xlsx
    export(list_of_datasets, 
               paste0("/Brand Cloud/Brand Cloud Outputs/SEMResults_", aud, ".xlsx")
               , rowNames = TRUE)
    
  }
  
}
brandCloud(audiences)





###########################
###  surveyCleaning_exp ###
###########################
export(survey3, "survey3.csv")
survey <- read.csv("survey3.csv") 

## Select columns
columnsSem <- c("ResponseID",
                colnames(survey)[grepl("^B2_.*", colnames(survey))])


surveySem <- survey[, c("ResponseID",
                        colnames(survey)[grepl("^B2_.*", colnames(survey))])]

# Fill na's with 0 so that we can sum later
surveySem[is.na(surveySem)] <- 0

# Convert to numerics
surveySem[,2:ncol(surveySem)] <- lapply(surveySem[,2:ncol(surveySem)], as.numeric)

# Calculate barclays experience and wom scores
surveySem$barc_pos_wom <- surveySem$B2_3_1 + surveySem$B2_3_3 + surveySem$B2_3_4
surveySem$barc_neg_wom <- surveySem$B2_3_2 + surveySem$B2_3_5
surveySem$barc_neut_wom <- surveySem$B2_3_6 + surveySem$B2_3_7 + surveySem$B2_3_12 + surveySem$B2_3_13
surveySem$barc_pos_exp <- surveySem$B2_3_8 + surveySem$B2_3_10
surveySem$barc_neg_exp <- surveySem$B2_3_9 + surveySem$B2_3_11

# Calculate competitor experience and wom scores
pos_wom_cols <- colnames(surveySem)[(grepl("_1$", colnames(surveySem)) | grepl("_3$", colnames(surveySem)) | grepl("_4$", colnames(surveySem))) & !grepl("^B2_3", colnames(surveySem))]
neg_wom_cols <- colnames(surveySem)[(grepl("_2$", colnames(surveySem)) | grepl("_5$", colnames(surveySem))) & !grepl("^B2_3", colnames(surveySem))]
neut_wom_cols <- colnames(surveySem)[(grepl("_6$", colnames(surveySem)) | grepl("_7$", colnames(surveySem)) | grepl("_12$", colnames(surveySem)) | grepl("_13$", colnames(surveySem))) & !grepl("^B2_3", colnames(surveySem))]
pos_exp_cols <- colnames(surveySem)[(grepl("_8$", colnames(surveySem)) | grepl("_10$", colnames(surveySem))) & !grepl("^B2_3", colnames(surveySem))]
neg_exp_cols <- colnames(surveySem)[(grepl("_9$", colnames(surveySem)) | grepl("_11$", colnames(surveySem))) & !grepl("^B2_3", colnames(surveySem))]


surveySem$comp_pos_wom <- rowSums(surveySem[,pos_wom_cols])
surveySem$comp_neg_wom <- rowSums(surveySem[,neg_wom_cols])
surveySem$comp_neut_wom <- rowSums(surveySem[,neut_wom_cols])
surveySem$comp_pos_exp <- rowSums(surveySem[,pos_exp_cols])
surveySem$comp_neg_exp <- rowSums(surveySem[,neg_exp_cols])

# Remove unneeded columns
surveySem <- surveySem[,!grepl("^B2", colnames(surveySem))]

# normalise columns
surveySem[,2:ncol(surveySem)] <- lapply(surveySem[,2:ncol(surveySem)], function(col){(col - min(col)) / (max(col) - min(col))})

# change id column name to match others
colnames(surveySem)[1] <- "id"

# Add exp_ at the start of all columns
colnames(surveySem)[2:ncol(surveySem)] <- paste0("exp_", colnames(surveySem)[2:ncol(surveySem)])

# Save variables
write.csv(surveySem,'experienceScores.csv',row.names = F)
         


#########################
###  mediaScoresCalcs ###
#########################
# export(survey, "survey4.csv")
# survey <- read.csv("survey4.csv")
mediaCols <- read.csv("mediaColLookup.csv")
mrWeighting = 0.5
bottomOut = FALSE
trustRecognition = FALSE # trustRecognition  - determines how the media score is calculated. If true (media score to 0. If the recognition is 1 then 1 else its the OTS) if false (media score is the media reach + (1-weighting)*OTS)
dnValue = 0 # 'dnValue' - dont know value to set for seeing an ad. Can be 'reach', 'recognition' or a number eg 0
mediaTargeting = FALSE # if 'media targeting = True' it will run an extra bit of code which scales down the OTS for people not in the target audience
combineVars = FALSE
quantiles = NULL
logNotes = ''


# Selecting columns
mediaColsStart <- mediaCols$columnStart

# Format Media Score Columns BASICALLY CHANGE THE COLUMN NAMES TO THE NAMES IN THE MEDIACOLSTART 
surveySem <- survey[, c("ResponseID", audiences, mediaColsStart)] ##include cols with responce id, audience, and vector mediaColStart
surveySem <- melt(surveySem, id.vars = "ResponseID") ## pivot longer, leave response id, colnames go to variables, values go to value
surveySem$columnStart <- surveySem$variable ## duplicate variable of variable called column start
surveySem <- left_join(surveySem, mediaCols) ## leftjoin join colnumName to surveySem using columnStarts we have the mediaCol names for surveySems DF
surveySem$columnName <- ifelse(is.na(surveySem$columnName), paste0(surveySem$columnStart), paste0(surveySem$columnName)) ## if the newly joined column columnName is NA because consumer segments columns are not joined, copy the consumer segment name
surveySem <- surveySem[, c("ResponseID", "columnName", "value")] ## select responseid, columnname which is the new column header with the actual names and the value
surveySem <- pivot_wider(surveySem, id_cols = "ResponseID", names_from = "columnName", values_from = "value") ## change back to wider format
surveySem[is.na(surveySem)] <- 0


# Format media recognition
# Format Press and Digital
surveySem[, grepl("prs_", colnames(surveySem))] <- ifelse(surveySem[, grepl("prs_", colnames(surveySem))] == 3 | surveySem[, grepl("prs_", colnames(surveySem))] == 4, 0, 1) ## longer than 1 month, 0, within 1 month 1
surveySem[, grepl("dig_", colnames(surveySem))] <- ifelse(surveySem[, grepl("dig_", colnames(surveySem))] == 3 | surveySem[, grepl("dig_", colnames(surveySem))] == 4, 0, 1) ## same as above

# Distribute tv, vod and yt where unsure based on recognition for each channel
mr_tv <- mean(surveySem$mr_tv) ## proportion of people surveyed who's seen the ad on tv
mr_vod <- mean(surveySem$mr_vod) ## for video on demand
mr_vid <- mean(surveySem$mr_vid) ## on youtube
mr_soc <- mean(surveySem$mr_soc) ## social media
mr_cin <- mean(surveySem$mr_cin) ## cinema

## proportion of the different channels proportion 
mr_vid_tot <- mr_tv + mr_vod + mr_vid + mr_soc + mr_cin 
mr_tv <- mr_tv / mr_vid_tot
mr_vod <- mr_vod / mr_vid_tot
mr_vid <- mr_vid / mr_vid_tot
mr_soc <- mr_soc / mr_vid_tot
mr_cin <- mr_cin / mr_vid_tot





# Format other channels where don't know can be answered 
# can either be set to reach, recognition or a set value
dnValue <- 0

colnames(surveySem)[grepl("mr_", colnames(surveySem)) &
                      !colnames(surveySem)%in%c("mr_tv", "mr_vod", "mr_vid", "mr_soc", "mr_cin")]

## This function doesn't do anything 
if (dnValue == 'reach'){
  reachVals <- read.csv('Media/mediaReach.csv', stringsAsFactors = FALSE)
  mr_rad <- reachVals[reachVals$media == 'rad', 'reach']
  mr_ooh <- reachVals[reachVals$media == 'ooh', 'reach']
  mr_prs <- reachVals[reachVals$media == 'print', 'reach']
} else if (dnValue == 'recognition'){
  mr_rad <- nrow(surveySem[surveySem$mr_rad == 1, ]) / nrow(surveySem)
  mr_ooh <- nrow(surveySem[surveySem$mr_ooh == 1, ]) / nrow(surveySem)
  mr_prs <- nrow(surveySem[surveySem$mr_prs == 1, ]) / nrow(surveySem)
  
} else { 
  mr_prs <- dnValue
  mr_ooh <- dnValue
  mr_rad <- dnValue
  
}



## turning vector to boolean, 1 is they heard/seen the ad, 0 is they didnt

surveySem$mr_rad <- ifelse(surveySem$mr_radio == 1, 1,
                           ifelse(surveySem$mr_radio == 2, 0,
                                  mr_rad))

surveySem$mr_ooh <- ifelse(surveySem$mr_ooh == 1, 1,
                           ifelse(surveySem$mr_ooh == 2, 0,
                                  mr_ooh))

surveySem$mr_prs <- ifelse(surveySem$mr_print == 1, 1, 
                           ifelse(surveySem$mr_print == 3, 0,
                                  mr_prs))

surveySem$mr_radio <- NULL
surveySem$mr_print <- NULL




# Function for splitting by hour in to the hourly columns
hourSplit <- function(df, hours, oldColname, singleDigit = T, mediaType = "tv"){
  for (i in hours){
    newColname <- ifelse(singleDigit,
                         paste0(mediaType, "_0",i),
                         paste0(mediaType, "_",i)) 
    df[,newColname] <- ifelse(df[,oldColname]==1,1,0)
  }
  return(df)
}




# Compute one ots score for each media channel by combining an individuals usage (from survey) and weightings (from media plans)
mediaTypes <- c('tv', 'vid', 'soc', 'rad', 'vod', 'cin', 'ooh', 'prs') ## includes all the SEM variables except don't know variable

for (med in mediaTypes){
  
  # If cinema, base ots on only how often and recent they've been to the cinema
  if (med == 'cin'){
    surveySem <- surveySem %>%
      mutate(ots_cin = ((1/(cin_oft* cin_rec)) - min(1/(cin_oft* cin_rec))) / (max(1/(cin_oft* cin_rec)) - min(1/(cin_oft* cin_rec))) ) # Create ots then normalise it
  }
  
  # If ooh, base only on if live in london
  if (med == 'ooh'){
    surveySem <- surveySem %>%
      mutate(ots_ooh = if_else(ooh_lon == 3, 1, 0))
  }
  
  # Combine channel scores to create OTS for Print/Press
  if( med == "prs"){
    weighting <- read.csv(paste0("Media/", med, '_weightings.csv'), stringsAsFactors = FALSE)
    #Calculate a weighting for each channel
    for (channel in unique(weighting$channel)){
      surveySem[,paste(med, channel, "ots", sep = "_")] <- surveySem[,grepl(paste0(channel,"$"), colnames(surveySem))] * weighting$weighting[weighting$channel == channel] ## create columns by selecting columns of surveySem ends with channel * weight of channel
      
      surveySem[,paste0('ots_',med)] <- rowSums(surveySem[, grepl("ots$", colnames(surveySem))]) ## sum the values in the same row for the columns that were created above (ends in ots) and save it in new column 'ots_prs' values between 0 & 5
      surveySem[,paste0('ots_',med)] <- (surveySem[,paste0('ots_',med)] - min(surveySem[,paste0('ots_',med)])) / (max(surveySem[,paste0('ots_',med)]) - min(surveySem[,paste0('ots_',med)]))
      surveySem <- surveySem[,colnames(surveySem)[!grepl("^temp_", colnames(surveySem))]]
    }
  }
  
  # Otherwise base off the weightings files
  if (!med%in%c("cin", "ooh", "prs")){
    
    # Bring in media weightings file
    weighting <- read.csv(paste0("Media/", med, '_weightings.csv'), stringsAsFactors = FALSE)
    
    # if we have weightings by hour then need to run through a more complex process
    if (colnames(weighting)[1] == 'hour'){
      
      # 1. Split when people watch tv into singular hours ################### - can this be more efficient?
      surveySem <- hourSplit(surveySem, c("00","23"), paste0(med, "_2300_0100"), singleDigit = F, mediaType = med)
      surveySem <- hourSplit(surveySem, c(1:5), paste0(med, "_0100_0600"), singleDigit = T, mediaType = med)
      surveySem <- hourSplit(surveySem, c(6:8), paste0(med, "_0600_0930"), singleDigit = T, mediaType = med)
      surveySem <- hourSplit(surveySem, c(10:11), paste0(med, "_0930_1200"), singleDigit = F, mediaType = med)
      surveySem <- hourSplit(surveySem, c(12:15), paste0(med,"_1200_1600"), singleDigit = F, mediaType = med)
      surveySem <- hourSplit(surveySem, c(16), paste0(med, "_1600_1730"), singleDigit = F, mediaType = med)
      surveySem <- hourSplit(surveySem, c(18:19), paste0(med, "_1730_2000"), singleDigit = F, mediaType = med)
      surveySem <- hourSplit(surveySem, c(20:22), paste0(med, "_2000_2300"), singleDigit = F, mediaType = med)
      
      surveySem[,paste0(med,'_09')] <- (ifelse(surveySem[,paste0(med, "_0600_0930")]==1,1,0)+ifelse(surveySem[,paste0(med, "_0930_1200")]==1,1,0))/2
      surveySem[,paste0(med,'_17')] <- (ifelse(surveySem[,paste0(med, "_1600_1730")]==1,1,0)+ifelse(surveySem[,paste0(med, "_1730_2000")]==1,1,0))/2
      
      # remove initial columns
      surveySem[,paste0(med,'_0600_0930')] <- NULL
      surveySem[,paste0(med,'_0930_1200')] <- NULL
      surveySem[,paste0(med,'_1200_1600')] <- NULL
      surveySem[,paste0(med,'_1600_1730')] <- NULL
      surveySem[,paste0(med,'_1730_2000')] <- NULL
      surveySem[,paste0(med,'_2000_2300')] <- NULL
      surveySem[,paste0(med,'_2300_0100')] <- NULL
      surveySem[,paste0(med,'_0100_0600')] <- NULL
      
      # 2. Create temporary weighting column for each channel and hour
      channels <-  sub("_weighting", "", sub(".*?_", "", colnames(weighting)[2:ncol(weighting)]))
      
      
      for (channel in channels){
        
        # Create a temporary column for each channel and hour
        for (i in 0:23){
          hourColumn <- ifelse(i < 10,
                               paste0(med, "_0",i),
                               paste0(med, "_",i))
          channelColumn <- paste0(med, "_", channel)
          surveySem[,paste("temp", med, hourColumn, channel, sep = "_")] <- surveySem[,channelColumn] * surveySem[,hourColumn] * weighting[weighting$hour==i,paste(med, channel, "weighting", sep = "_")]
        }
        
        # Sum up hour columns to get score for each channel
        surveySem[,paste("temp", med, channel, "ots", sep = "_")] <- rowSums(surveySem[,(grepl(paste0(channel, "$"), colnames(surveySem)) & grepl(med, colnames(surveySem)))][2:ncol(surveySem[,(grepl(paste0(channel, "$"), colnames(surveySem)) & grepl(med, colnames(surveySem)))])])
        
      }
      
      # 3. Combine to create overall score and normalise
      surveySem[,paste0('ots_',med)] <- rowSums(surveySem[,grepl("ots$", colnames(surveySem))])
      surveySem[,paste0('ots_',med)] <- (surveySem[,paste0('ots_',med)] - min(surveySem[,paste0('ots_',med)])) / (max(surveySem[,paste0('ots_',med)]) - min(surveySem[,paste0('ots_',med)]))
      surveySem <- surveySem[,colnames(surveySem)[!grepl("^temp_", colnames(surveySem))]]
      
    } else {
      
      #Calculate a weighting for each channel
      for (channel in unique(weighting$channel)){
        surveySem[,paste("temp", med, channel, "ots", sep = "_")] <- surveySem[,(grepl(paste0(channel,"$"), colnames(surveySem)) & grepl(med, colnames(surveySem)))] * weighting$weighting[weighting$channel == channel]
      }
      
      # Combine channel scores to create one OTS score for vod
      surveySem[,paste0('ots_',med)] <- rowSums(surveySem[, grepl("ots$", colnames(surveySem))])
      surveySem[,paste0('ots_',med)] <- (surveySem[,paste0('ots_',med)] - min(surveySem[,paste0('ots_',med)])) / (max(surveySem[,paste0('ots_',med)]) - min(surveySem[,paste0('ots_',med)]))
      surveySem <- surveySem[,colnames(surveySem)[!grepl("^temp_", colnames(surveySem))]]
      
      
    }
    
  }
  
}



###Reduce OTS for targeted media ###
### Manually change when mediaTargeting is TRUE ###
mediaTargeting <- FALSE
if (mediaTargeting){
  
  # Reduce ots for digital when not the target audience
  surveySem <- surveySem %>%
    mutate(ots_vid = if_else(day2dayPlanners == 1, 0.25 * ots_vid, ots_vid),
           ots_soc = if_else(day2dayPlanners == 1, 0.25 * ots_soc, ots_soc),
           ots_vod = if_else(day2dayPlanners == 1, 0.25 * ots_vod, ots_vod))
  
  
  
}

### Combine MR and OTS ###
bottomOut = FALSE
combineVars = FALSE
trustRecognition <- 0.5

# Load media reach file if we're bottoming out the variables
# Take people who are most likely to have seen, e.g. keep 60%
if (bottomOut|combineVars){
  medReach <- read.csv('Media/mediaReach.csv', stringsAsFactors = FALSE)
} 

if (combineVars){
  # Combine tv and vod into av?
  surveySem$mr_av <- (0.5*surveySem$mr_tv) + (0.5*surveySem$mr_vod) # MR for AV
  surveySem$ots_av <- (0.5*surveySem$ots_tv) + (0.5*surveySem$ots_vod) # Opportunity for AV 
  surveySem$mr_tv <- NULL
  surveySem$mr_vod <- NULL
  surveySem$ots_tv <- NULL
  surveySem$ots_vod <- NULL
  mediaTypes <- c('av', 'vid', 'soc', 'rad', 'cin', 'ooh', 'prs')
}


for (med in mediaTypes){
  # Combine
  if (trustRecognition){
    surveySem[,paste0('medScore_',med)] <- 0
    surveySem[surveySem[paste0('mr_',med)]==1, paste0('medScore_',med)] <- 1
    surveySem[surveySem[paste0('mr_',med)]!=1, paste0('medScore_',med)] <- surveySem[surveySem[paste0('mr_',med)]!=1, paste0('ots_',med)]
  } else{
    surveySem[,paste0('medScore_',med)] <- (mrWeighting*surveySem[,paste0('mr_',med)] + (1-mrWeighting)*surveySem[,paste0('ots_',med)]) 
  }
  
  # Bottom Out
  if (bottomOut){
    # Order by variable
    surveySem <- surveySem[order(surveySem[,paste0('medScore_',med)]),]
    
    # % made zero
    zeroPerc <- ifelse((1-medReach[medReach$media==med, 'reach'])/2 < 0.2,
                       0.2,
                       (1-medReach[medReach$media==med, 'reach'])/2)
    
    maxPerc <- ifelse((1-medReach[medReach$media==med, 'reach'])/4 < 0.1,
                      0.1,
                      (1-medReach[medReach$media==med, 'reach'])/4)
    
    # Split into min and max
    surveySemMin <- surveySem[0:round(nrow(surveySem)*zeroPerc),]
    surveySemMid <- surveySem[(round(nrow(surveySem)*zeroPerc)+1):round(nrow(surveySem)*(1-maxPerc)),]
    surveySemMax <- surveySem[(round(nrow(surveySem)*(1-maxPerc))+1):nrow(surveySem),]
    
    # Set bottom to 0 and shift top so that minimum score is reach/2, max top % 1
    surveySemMin[,paste0('medScore_',med)] <- 0
    surveySemMid[,paste0('medScore_',med)] <- ifelse(max(surveySemMid[,paste0('medScore_',med)]) == 0, 0, (surveySemMid[,paste0('medScore_',med)] - min(surveySemMid[,paste0('medScore_',med)]))/(max(surveySemMid[,paste0('medScore_',med)])-min(surveySemMid[,paste0('medScore_',med)])))
    surveySemMid[,paste0('medScore_',med)] <- (1-(medReach[medReach$media==med, 'reach']/2)) * surveySemMid[,paste0('medScore_',med)] + medReach[medReach$media==med, 'reach']/2
    surveySemMax[,paste0('medScore_',med)] <- 1
    
    # Put back together
    surveySem <- surveySemMin %>%
      bind_rows(surveySemMid) %>%
      bind_rows(surveySemMax)
    
  }
  # Rescale column between 0 and 1
  surveySem[,paste0('medScore_',med)] <- (surveySem[,paste0('medScore_',med)] - min(surveySem[,paste0('medScore_',med)]))/(max(surveySem[,paste0('medScore_',med)])-min(surveySem[,paste0('medScore_',med)]))
  
}



# Convert media score variables into quantiles
# Number of quantiles set as argument
quantiles = NULL
if (!is.null(quantiles)){
  
  surveySem <- surveySem %>%
    mutate(across(colnames(surveySem)[grepl("medScore", colnames(surveySem))],
                  function(x){(ntile(x,(quantiles+1))-1)/quantiles}))
}

surveySem$id <- surveySem$ResponseID
# Save out a file including ots and mr
surveySem_total <- surveySem[,c('id',
                                colnames(surveySem)[grepl('ots_', colnames(surveySem))],
                                colnames(surveySem)[grepl('mr_', colnames(surveySem))],
                                colnames(surveySem)[grepl('medScore', colnames(surveySem))])]

##### Update for new run name
outputFilename <- "test"
write.csv(surveySem_total, paste0('./Media/Media Scores/mediaWorkings_', outputFilename, '.csv'), row.names = FALSE)


# Reduce columns to only media scores for each id
surveySem <- surveySem[,c('id',
                          colnames(surveySem)[grepl('medScore', colnames(surveySem))])]


# Save out
write.csv(surveySem, paste0('./Media/Media Scores/mediaScores_', outputFilename, '.csv'), row.names = FALSE)

logNotes = "Testing"
# Save out log 
write('###', file=paste0("./Media/Media Scores/mediaScores_log.txt"), append = TRUE)
write(paste0('Date and time: ', Sys.time()), file="./Media/Media Scores/mediaScores_log.txt", append = TRUE)
write(paste0('File name: mediaScores_', outputFilename, '.csv'), file="./Media/Media Scores/mediaScores_log.txt", append = TRUE)
write(paste0('mrWeighting: ', mrWeighting), file="./Media/Media Scores/mediaScores_log.txt", append = TRUE)
write(paste0('bottomOut: ', bottomOut), file="./Media/Media Scores/mediaScores_log.txt", append = TRUE)
write(paste0('trustRecognition: ', trustRecognition), file="./Media/Media Scores/mediaScores_log.txt", append = TRUE)
write(paste0('dnValue: ', dnValue), file="./Media/Media Scores/mediaScores_log.txt", append = TRUE)
write(paste0('mediaTargeting: ', mediaTargeting), file="./Media/Media Scores/mediaScores_log.txt", append = TRUE)
write(paste0('combineVars: ', combineVars), file="./Media/Media Scores/mediaScores_log.txt", append = TRUE)
write(paste0('quantiles: ', quantiles), file="./Media/Media Scores/mediaScores_log.txt", append = TRUE)
write(logNotes, file="./Media/Media Scores/mediaScores_log.txt", append = TRUE)
write('', file="./Media/Media Scores/mediaScores_log.txt", append = TRUE)








### Building the regression model ###
### This involves two functions   ###



##################
## step_nnMedia ##
##################
### This function is used to buld a stepwise regression model  ###
### while ensuring all media variables are non-negative        ###
### It also allows for specific media variables to be retested ###
### after all others have been removed                         ###
step_nnMedia <- function(formula, data, vars, latent, mediaRetesting, retestingThreshold){
  
  # Run stepwise regression
  model <- step(lm(formula = formula,
                   data = data))
  
  # Check for negative media variables
  medCoeffs <- model$coefficients[grepl("medScore_", names(model$coefficients))]
  
  removeVars <- c()
  
  if (length(medCoeffs)>0){
    # If there is a negative variable then find first and remove before rerunning
    while (min(medCoeffs)<0){
      
      # Find first negative coefficient and remove from input vars
      removeVars <- append(removeVars, names(model$coefficients)[grepl("medScore_", names(model$coefficients))&(model$coefficients<0)][1])
      vars <- vars[!vars%in%removeVars]
      
      # recreate formula
      formula <- paste(latent, " ~ ", paste(vars, collapse = " + "))
      
      # Build model again
      model <- step(lm(formula,
                       data = data))
      
      # Check for negative media variables
      medCoeffs <- model$coefficients[grepl("medScore_", names(model$coefficients))]
      
      # Add this so that while loop doesn't error
      if(length(medCoeffs)==0){
        medCoeffs <- 1
      }
    }
  }
  
  # Find the final model variables
  modelVars <- names(model$coefficients)[2:length(names(model$coefficients))]
  
  ### Media retargetting section ###
  # Retest the desired variables and check their p-values - keep if less than 0.15
  if (length(mediaRetesting)>0){
    # Add variables to retest - if already there, unique will get rid of the second
    vars <- unique(append(modelVars, mediaRetesting))
    
    # recreate formula
    formula <- paste(latent, " ~ ", paste(vars, collapse = " + "))
    
    # Build model again
    model <- lm(formula,
                data = data)
    
    # Remove if new variables have a negative coefficient or p-value is greater than 0.15
    removeRetesting <- names(model$coefficients)[names(model$coefficients)%in%mediaRetesting &
                                                   ((model$coefficients < 0) |
                                                      (summary(model)$coefficients[,4] > retestingThreshold))]
    
    # If we aren't keeping any of the retargetted variables, we need to remove them again and rebuild the model
    if (length(removeRetesting) > 0){
      
      # recreate formula
      vars <- vars[!vars%in%removeRetesting]
      formula <- paste(latent, " ~ ", paste(vars, collapse = " + "))
      
      # Build model again
      model <- lm(formula,
                  data = data)
      
    }
    
  }
  ###
  
  return(model)
}


#########################
## regressionModel_SEM ##
#########################
### This is the main function when building the regression           ###
### It brings in all of the experience and media variables           ###
### And runs a regression for each audience for all latents          ###
### These are aggregated to find the overall impact on consideration ###
### Multiple outputs are created with the overall                    ###

user <- Sys.getenv("USERNAME")
outputLoc = paste0("C:/Users/",user,"/OneWorkplace/OMD-Audience Analytics OMDUK - General/Client Work/Barclays/Barclays BTM/2022/Q4 Campaign/5. Formatted data/Regression/")
modelName = "test1"
#audiences = c("ActiveStrivers", "day2dayPlanners", "comfortablyCommitted", "cherryPickers", "allAdults")
mediaFilename = "test"
retestingThreshold = 0.15 # for p values
mediaRetesting = NULL
  
# Create empty dataframes to populate
all_output <- data.frame()
media_output <- data.frame()
media_drivers <- data.frame()
all_models <- data.frame()
correlations <- data.frame()
all_model_data <- data.frame()
  
# create folder for model files
file_list <- list.files(outputLoc)
if(!(modelName %in% file_list)){
  dir.create(paste0(outputLoc, modelName), showWarnings = FALSE)
}
  
# Cycle through audiences
for (aud in audiences){
    
    print(paste0("Running regressions for ", aud))
    
    # Bring in SEM values
    sem <- read.csv(paste0("SEM/SEM_",aud,"_ID.csv"))
    # note number of questions
    qs <- colnames(sem)[2:ncol(sem)]
    
    # combine statements into latents
    loadings <- read_excel(paste0("Brand Cloud/Brand Cloud Outputs/SEMResults_",aud,".xlsx"),
                           sheet = "Outer Loadings")
    #loadings <- read_excel(paste0("Brand Cloud/Brand Cloud Outputs/SEMResults_ActiveStrivers.xlsx"),
    #                       sheet = "Outer Loadings")
    
    # Replace & with _
    colnames(loadings) <- gsub("&", "_", colnames(loadings))
    
    for (latent in colnames(loadings)[2:ncol(loadings)]){
      sem[,latent] <- 0
      totalLoading <- 0
      
      # Cycle through the statements to fill in the score for the latent
      for (statement in loadings$...1){
        sem[, latent] <- sem[,latent] + (as.numeric(loadings[loadings$...1 == statement, latent]) * sem[,statement])
        totalLoading <- totalLoading + as.numeric(loadings[loadings$...1 == statement, latent])
      }
      
      # Scale latent score to between 1 and 10
      sem[,latent] <- sem[,latent] / totalLoading
      
    }
    
    # remove question vars
    sem <- sem[,!colnames(sem)%in%qs]
    
    
    
    ### Bring in variables
    # media
    med <- read.csv(paste0("Media/Media scores/mediaScores_",mediaFilename,".csv"))
    
    # experience
    exp <- read.csv("experienceScores.csv")
    
    # combine into final model file
    modelData <- merge(sem, med, by = "id", all.x = T)
    modelData <- merge(modelData, exp, by = "id", all.x = T)
    
    
    
    ### Build model
    # Select variables
    vars <- colnames(modelData)[c(grep("^exp", colnames(modelData)),
                                  grep("^med", colnames(modelData)))]
    
    
    # Run each media regression
    allModels <- list()
    mediaRetesting = NULL
    for (latent in colnames(loadings)[2:ncol(loadings)]){
      formula <- paste(latent, " ~ ", paste(vars, collapse = " + "))
      
      model <- step_nnMedia(formula = formula,
                            data = modelData,
                            vars = vars,
                            latent = latent,
                            mediaRetesting = mediaRetesting,
                            retestingThreshold = retestingThreshold)
      
      allModels[[latent]] <- model
    }
    
    # Pull out summaries of models
    summaryTotal <- data.frame()
    
    for (i in names(allModels)){
      summarytemp <- eval(parse(text = paste0("summary(allModels$", i,")")))
      summarytemp <- as.data.frame(summarytemp$coefficient)
      summarytemp$beta <- rownames(summarytemp) 
      summarytemp$model <- i
      summaryTotal <- rbind(summaryTotal,summarytemp)
    }
    
    summarytotal2 <- data.frame()
    # This section calculates the standardised coefficients
    for (i in names(allModels)){
      summarytemp2 <- eval(parse(text = paste0("lm.beta(allModels$", i,")")))
      summarytemp2 <- as.data.frame(summarytemp2)
      summarytemp2$beta <- rownames(summarytemp2) 
      summarytemp2$model <- i
      summarytotal2 <- rbind(summarytotal2,summarytemp2)
    }
    
    rownames(summaryTotal) <- c()
    rownames(summarytotal2) <- c()
    
    
    ### Calculate effects
    results <- summaryTotal
    stCoefficients <- summarytotal2  
    variableDesc <- stat.desc(modelData)
    
    # Bring in smartPls output and format
    smartPls <- read_excel(paste0("Brand Cloud/Brand Cloud Outputs/SEMResults_",aud,".xlsx"),
                           sheet = "Total Effects") 
    #smartPls <- read_excel("Brand Cloud/Brand Cloud Outputs/SEMResults_ActiveStrivers.xlsx",
    #                       sheet = "Total Effects") 
    
    # format smartPLS total effects output
    smartPls <- as.data.frame(smartPls)
    smartPls$X <- gsub("&", "_", smartPls[,1])
    smartPls$kpi <- smartPls[,'Consideration']
    
    smartPls <- dplyr::select(smartPls, X, kpi)
    smartPls <- transform(smartPls, kpi = as.numeric(kpi))
    smartPls$perc <- smartPls$kpi / sum(smartPls$kpi)
    colnames(smartPls) <- c("latent", "abs_value", "SEM_loadings")
    
    # Add mean and standard deviation
    i <- 1
    for(driver in smartPls$latent){
      smartPls$mean[i] <- variableDesc[9, which(names(variableDesc) %in% driver)]
      smartPls$SD[i] <- variableDesc[13, which(names(variableDesc) %in% driver)]
      i <- i + 1
    }
    
    # for results and SC create concat column of modelvariable
    results$model_variable <- paste0(results$model, results$beta)
    stCoefficients$model_variable <- paste0(stCoefficients$model, stCoefficients$beta)
    
    # bring in st coefficients to results
    all_variables <- merge(results, stCoefficients[ , c("summarytemp2", "model_variable")], all.x = T, by = "model_variable")
    
    # create model results without the intercept rows from results
    model_results <- na.omit(all_variables)
    
    # create latent name column 
    model_results$latent <- model_results$model
    
    # bring in variable info to model_results
    variableInfo <- read.csv("variableInfo.csv", stringsAsFactors=FALSE)            
    colnames(variableInfo) <- c("beta", "variable_group", "brand", "channel")      
    model_results <- merge(model_results, variableInfo[ , c("variable_group", "brand", "channel", "beta")], all.x = T, by = "beta")
    
    # SEM loadings by latent
    model_results <- merge(model_results, smartPls[ , c("latent", "abs_value", "SEM_loadings")], all.x = T, by = "latent")
    
    # calculations to match model results tab (could be condensed?)
    i <- 1
    for(var in model_results$beta){
      model_results$variable_mean[i] <- variableDesc[9, which(names(variableDesc) %in% var)]
      model_results$variable_SD[i] <- variableDesc[13, which(names(variableDesc) %in% var)]
      i <- i + 1
    }
    model_results$stC_abs <- model_results$summarytemp2 * model_results$abs_value
    model_results$unstandardising <- model_results$stC_abs/model_results$variable_SD
    KPI_SD <- smartPls[which(smartPls$latent == 'Consideration'), 5]
    model_results$convert_10pt <- model_results$unstandardising * KPI_SD
    model_results$contribution <- model_results$variable_mean * model_results$convert_10pt
    model_results$unstandardising_perceptions <- model_results$summarytemp2/model_results$variable_SD
    
    
    ############################################################################### 
    # Results - total impact
    ############################################################################### 
    print(paste0("Adding total impact results for ", aud))
    
    total_impact <- model_results[, c("brand", "variable_group", "contribution")] %>% 
      group_by(brand, variable_group) %>% 
      summarise(contribution = sum(contribution))
    
    total_impact <- total_impact[with(total_impact, order(brand, -contribution)), ]
    
    ############################################################################### 
    # Results - media impact 
    ############################################################################### 
    print(paste0("Adding media impact results for ", aud))
    media_impact <- model_results[, c("brand", "variable_group", "beta", "contribution")] %>%
      group_by(brand, variable_group, beta) %>%
      summarise(contribution = sum(contribution))
    
    media_impact <- filter(media_impact, brand == "Bar" & (grepl("Media", variable_group) | grepl("media", variable_group))) ### Change Brand
    media_impact <- media_impact[,-c(1,2)]
    media_impact <- media_impact[with(media_impact, order(-contribution)), ]
    
    media_impact$percent <- media_impact$contribution/sum(media_impact$contribution)
    
    media_impact$audience <- aud
    
    media_output <- rbind(media_output,media_impact)
    
    ############################################################################### 
    # Results - media impact on drivers 
    ############################################################################### 
    print(paste0("Adding media impact results on drivers for ", aud))
    media_variables <- filter(model_results, grepl("Media", variable_group) | grepl("media", variable_group))
    media_variables <- unique(media_variables$variable_group)
    
    media_driver_impact <- model_results[, c("latent", "beta", "contribution", "variable_group")] %>%
      group_by(variable_group, beta, latent) %>%
      summarise(contribution = sum(contribution))
    
    
    drivers_temp <- data.frame()
    
    #(loop for each media variable)
    for (m in media_variables){ 
      print(m)
      name <- paste0("media_driver_impact_", m)
      
      temp <- filter(media_driver_impact, variable_group == m)
      temp <- temp[with(temp, order(beta, -contribution)), ]
      temp <- temp[,-1]
      
      drivers_temp <- rbind(drivers_temp,temp)
      
      assign(name, temp)
    }
    
    drivers_temp$percent <- drivers_temp$contribution/sum(drivers_temp$contribution)
    
    drivers_temp$audience <- aud
    
    media_drivers <- rbind(drivers_temp,media_drivers)
    
    
    
    ###### Remaking results to calculate base and so they're a % of KPI
    
    all_impacts <- model_results[, c("brand", "variable_group", "beta", "contribution")] %>%
      group_by(brand, variable_group, beta) %>%
      summarise(contribution = sum(contribution))
    
    KPI_mean <- smartPls[which(smartPls$latent == 'Consideration'), 4]
    basenum <- KPI_mean - sum(total_impact$contribution)
    
    rownum <- nrow(total_impact)+1
    total_impact[rownum,1] <- "Bar"
    total_impact[rownum,2] <- "Base"
    total_impact[rownum,3] <- basenum
    
    total_impact[rownum+1,1] <- "Bar"
    total_impact[rownum+1,2] <- 'Consideration'
    total_impact[rownum+1,3] <- KPI_mean
    
    total_impact$perc <- total_impact$contribution/KPI_mean
    
    total_impact$audience <- aud
    
    total_impact <- as.data.frame(total_impact)
    
    all_output <- rbind(all_output,total_impact)
    
    ############################################################################### 
    # Results - the total model results stuff
    ############################################################################### 
    print(paste0("Adding full model results for ", aud))
    model_results$audience <- aud
    
    all_models <- rbind(all_models,model_results)
    
    
    ############################################################################### 
    # Correlations
    ###############################################################################
    corr_aud <- as.data.frame(cor(modelData[,22:ncol(modelData)]))#, use="complete.obs" for when some vars have missing values
    corr_aud$audience <- aud
    
    correlations <- bind_rows(correlations, corr_aud)
    
    ############################################################################### 
    # model data
    ###############################################################################
    all_model_data <- rbind(modelData, all_model_data)
    
    
    print(paste0("Regressions run for ", aud))
    
  }
  
  
### Write out output files
write.csv(all_model_data, 
          paste0(outputLoc, modelName, "/allaudienceModelData_", modelName, ".csv"), 
          row.names = F)
write.csv(all_output, 
          paste0(outputLoc, modelName, "/allaudiencetotalimpact_", modelName, ".csv"), 
          row.names = F)
write.csv(media_output, 
          paste0(outputLoc, modelName, "/allaudiencemediaimpact_", modelName, ".csv"), 
          row.names = F)
write.csv(media_drivers, 
          paste0(outputLoc, modelName, "/allaudiencemediadrivers_", modelName, ".csv"), 
          row.names = F)
write.csv(all_models, 
          paste0(outputLoc, modelName, "/allaudienceModelResults_", modelName, ".csv"), 
          row.names = F)
write.csv(correlations, 
          paste0(outputLoc, modelName, "/allaudienceCorrelations_", modelName, ".csv"), 
          row.names = T)
  
  
  
  # Save out log
write('###', file=paste0(outputLoc, "regression_log.txt"), append = TRUE)
write(paste0('Date and time: ', Sys.time()), file=paste0(outputLoc, "regression_log.txt"), append = TRUE)
write(paste0('Model name: ', modelName, '.csv'), file=paste0(outputLoc, "regression_log.txt"), append = TRUE)
write(paste0('mediaFilename: ', mediaFilename), file=paste0(outputLoc, "regression_log.txt"), append = TRUE)
write(paste0('mediaRetesting: ', mediaRetesting), file=paste0(outputLoc, "regression_log.txt"), append = TRUE)
write(paste0('retestingThreshold: ', retestingThreshold), file=paste0(outputLoc, "regression_log.txt"), append = TRUE)
write(logNotes, file=paste0(outputLoc, "regression_log.txt"), append = TRUE)
write('', file=paste0(outputLoc, "regression_log.txt"), append = TRUE)
  

### Check functions, load or install as necessary
packages <- c("qualtRics", "googlesheets4", "tidyverse", "gtools")
for (package in packages) {
    # If the package is not installed, install it
    if (!require(package, character.only = TRUE)) {
        install.packages(package, dependencies = TRUE)
    }
    # Load the package
    library(package, character.only = TRUE)
}



### Function 1: fetch & rename surveys from Qualtrics, using the qualtRics Package
fetch_and_rename_survey <- function(input_qid) {

    survey <- fetch_survey(
        input_qid,
        add_column_map = TRUE,
        include_metadata = c("ResponseId", "RecordedDate"),
        include_embedded = NA,
        force_request = TRUE,
        convert = FALSE,
        label = FALSE,
        verbose = FALSE
    )
    
    colnames(survey) <-
        attributes(survey)[["column_map"]][["description"]]
    
    #a function to rename columns if they exist
    rename_cols_if_exist <- function(df, name_dict) {
        intersecting_names <- intersect(names(df), names(name_dict))
        if (length(intersecting_names) > 0) {
            df <-
                rename(df, !!!setNames(intersecting_names, name_dict[intersecting_names]))
        }
        return(df)
    }
    
    #a function to select the columns of interest from the full survey.
    select_or_create_cols <- function(df, col_names) {
        # Select the existing columns
        df <- df[, names(df) %in% col_names, drop = FALSE]
        
        # Find out which columns are missing
        missing_cols <- names(df)[!names(df) %in% col_names]
        
        # Create the missing columns and fill them with NA
        df[missing_cols] <-
            lapply(missing_cols, function(x)
                rep(NA, nrow(df)))
        
        return(df)
    }
    
    #col_names we want in Google
    col_names <- c(
        "qid",
        "rid",
        "date",
        "overall_satisfaction",
        "content_relevant",
        "content_enriching",
        "overall_service",
        "deib_incorporate",
        "deib_represent",
        "community_partof",
        "community_shared",
        "community_network",
        "workload",
        "knowledge_skill_ready" ,
        "knowledge_skill_needs",
        "enrolled_with_colleagues" ,
        "in_us",
        "nps",
        "nps_score",
        "recommend_now",
        "current_sector",
        "modality_person",
        "modality_hybrid",
        "modality_online_asyn",
        "modality_online_syn",
        "length_day_1_2" ,
        "length_day_3_6" ,
        "length_week_1_2",
        "length_week_3_5" ,
        "length_week_6_plus",
        "modality_1st",
        "modality_2nd",
        "length_1st",
        "length_2nd"
    )
    
    #rename columns
    oldname_newname <- c(
        "Response ID" = "rid",
        "Recorded Date" = "date",
        "Overall, how satisfied were you with the program?" = "overall_satisfaction",
        "To what extent are the following statements about the program content true for you? - The program content was relevant to my professional context." = "content_relevant",
        "To what extent are the following statements about the program content true for you? - The program content was personally enriching." = "content_enriching",
        "Overall, how satisfied were you with the customer service you received from program staff (before and during the program)?" = "overall_service",
        "HGSE is committed to diversity, equity, inclusion, and belonging. To what extent do you think the following statements are true? - The program content incorporated diverse voices and identities." = "deib_incorporate",
        "HGSE is committed to diversity, equity, inclusion, and belonging. To what extent do you think the following statements are true? - The program participants represented a diversity of voices and identities." = "deib_represent",
        "We strive to create a community of practice in our programs. To what extent are the following statements true for you? - I felt like part of a learning community." = "community_partof",
        "We strive to create a community of practice in our programs. To what extent are the following statements true for you? - I shared reflections, feedback, and implementation strategies with my peers." = "community_shared",
        "We strive to create a community of practice in our programs. To what extent are the following statements true for you? - I grew my professional network." = "community_network",
        "How do you assess the workload required by this program?" = "workload",
        "To what extent do you feel ready to use the knowledge and skills you've gained from this program?" = "knowledge_skill_ready",
        "Did you enroll in this program with other colleagues from your institution?" = "enrolled_with_colleagues",
        "Do you currently reside in the United States?" = "in_us",
        "How likely are you to recommend this program to a friend or colleague? - Group" = "nps",
        "How likely are you to recommend this program to a friend or colleague?" = "nps_score",
        "Would you like to recommend this program to a colleague now, by entering their email(s) on the next screen?" = "recommend_now",
        "Which program delivery methods do you prefer when undertaking professional development? - Entirely in person" = "modality_person",
        "Which program delivery methods do you prefer when undertaking professional development? - Hybrid (in person and online)" = "modality_hybrid",
        "Which program delivery methods do you prefer when undertaking professional development? - Entirely online (mostly asynchronous)" =
            "modality_online_asyn",
        "Which program delivery methods do you prefer when undertaking professional development? - Entirely online (mostly live/synchronous)" = "modality_online_syn",
        "What sector do you currently work in?" = "current_sector",
        "Which program lengths do you prefer for your own professional development? - 1-2 days" = "length_day_1_2",
        "Which program lengths do you prefer for your own professional development? - 3-6 days" = "length_day_3_6",
        "Which program lengths do you prefer for your own professional development? - 1-2 weeks" = "length_week_1_2",
        "Which program lengths do you prefer for your own professional development? - 3-5 weeks" = "length_week_3_5" ,
        "Which program lengths do you prefer for your own professional development? - 6+ weeks, i.e. semester or year-long" =
            "length_week_6_plus",
        "Which program delivery method would be your first choice when undertaking professional development?" = "modality_1st",
        "Which program delivery method would be your second choice when undertaking professional development?" = "modality_2nd",
        "Which program length would be your first choice for your own professional development?" = "length_1st",
        "Which program length would be your second choice for your own professional development?" = "length_2nd"
        
    )
    
    #rename questions as column names, shorten to something readable
    survey <- rename_cols_if_exist(survey, oldname_newname)
    
    #add a column for qid
    survey <- cbind(qid = input_qid, survey)
    
    #select only the columns we want from the full survey
    survey <- select_or_create_cols(survey, col_names)
    
    return(survey)
    
}




### Function 2: update google sheets with the google sheet function
update_master <- function(master_sid = master_sid) {
    #read the 'master' sheet from google, all col_types set to character
    master <-
        read_sheet(ss = master_sid,
                   sheet = "master",
                   col_types = "c")
    
    #store program information in a separate table for later use
    program_info <-
        master %>% select(program_name, qid, modality, k12, higher_ed, early_childhood)
    
    #get all the qid's
    master.qid <- master$qid
    
    #fetch all responses in the update list and stack
    updates <- map_df(master.qid, fetch_and_rename_survey)
    
    #update progra info
    updates <-
        program_info |>  
        dplyr::right_join(updates, by = "qid")
    
    #remove old sheet
    sheet_delete(ss = master_sid, sheet = "stacked")
    
    #write the new sheet
    sheet_write(data = updates,
                ss = master_sid,
                sheet = "stacked")
    
    print("All surveys updated and new surveys added")
    
}




### Function 3 - summarize KPI by group

update_summaries <- function(master_sid = master_sid) {
    stacked <- read_sheet(ss = master_sid, sheet = "stacked")
    master <-
        read_sheet(ss = master_sid,
                   sheet = "master",
                   col_types = "c")
    
    # a custom function to handle NPS
    get_nps <- function(vec) {
        vec <- vec |>  na.omit()
        nps <- sum(vec == 3) / length(vec) - sum(vec == 1) / length(vec)
        return(round(nps * 100, 2))
    }
    
    #make sure all except ID, date, and modality columns are in numeric format
    stacked <-
        stacked |> mutate(across(
            .cols = -c("program_name","qid", "rid", "date", "modality"),
            .fns = as.numeric
        ))
    
    #scales and column names for scale mapping
    satisfaction_scale_7 <- c(
        "Very dissatisfied",
        "Moderately dissatisfied",
        "Slightly dissatisfied",
        "Neither satisfied nor dissatisfied",
        "Slightly satisfied",
        "Moderately satisfied",
        "Very satisfied"
    )
    amount_scale_5 <- c("Not at all",
                        "Slightly",
                        "Moderately",
                        "Quite",
                        "Completely")
    #side note - the value scales have a non-ordinal "did not participate"
    valueable_scale_5 <- c("Not at all",
                           "Slightly",
                           "Moderately",
                           "Quite",
                           "Completely")
    satisfaction_items <-
        c("overall_satisfaction", "overall_service")
    amount_items <-
        c(
            "content_relevant",
            "content_enriching",
            "deib_incorporate",
            "deib_represent",
            "community_partof",
            "community_shared" ,
            "community_network",
            "workload",
            "knowledge_skill_ready"
        )
    #Qualtrics recode these as 1 and NA
    indicator_items <- c(
        "modality_person",
        "modality_hybrid",
        "modality_online_asyn",
        "modality_online_syn",
        "length_day_1_2",
        "length_day_3_6",
        "length_week_1_2",
        "length_week_3_5",
        "length_week_6_plus"
    )
    #Qualtrics recode these as 2 and 1
    yes_no_items <- c("enrolled_with_colleagues", "in_us")
    
    #convert to factors
    stacked <- stacked %>%
        mutate(
            across(
                .cols = all_of(satisfaction_items),
                .fns = ~ factor(.x, levels = 1:7, labels = satisfaction_scale_7)
            ),
            across(
                .cols = all_of(amount_items),
                .fns = ~ factor(.x, levels = 1:5, labels = amount_scale_5)
            ),
            #yes/no - R codes these as 2/1. subtract 1.
            across(
                .cols = all_of(yes_no_items),
                .fns = ~ ifelse(is.na(.x), 0, .x - 1)
            ),
            #indicators - those are indicator variables converted categoricals, so we'll treat NA as 0
            across(
                .cols = all_of(indicator_items),
                .fns = ~ ifelse(is.na(.x), 0, .x)
            )
        )
    
    #convert datset to a dataframe for the likerk pagage
    stacked <- as.data.frame(stacked)
    
    #means kpis grouped by IDs
    kpi_summaries <- stacked |> group_by(qid) |> summarise(
        n_responses = n(),
        #counts total number of respones per qid
        nps = get_nps(nps),
        # calculates NPS from nps group scores
        across(all_of(
            c(
                satisfaction_items,
                amount_items,
                indicator_items,
                yes_no_items
            )
        ),
        ~ round(mean(
            as.numeric(.x), na.rm = TRUE
        ), 2),
        .names = "{.col}_mn")
    )
    
    #add meta data back into the spreadsheet
    kpi_summaries <-
        master |> 
        dplyr::right_join(kpi_summaries, by = "qid", keep = FALSE)
    
    #upload the updated sheet into google
    sheet_write(data = kpi_summaries,
                ss = master_sid,
                sheet = "kpi summaries")
    
    print("kpi summaries updated")
}

packages <- c("qualtRics", "dplyr", "readr", "tidyr", "gtools", "readxl", "openxlsx", "purrr", "Microsoft365R")
sapply(packages, function(pkg) {
    if (!requireNamespace(pkg, quietly = TRUE)) {
        message(paste("Installing package:", pkg))
        install.packages(pkg)
    }
    message(paste("Loading package:", pkg))
    library(pkg, character.only = TRUE)
})

update_master <- function(wb_path = local_wb_path, mapping_df) {
    
    # Initialize an empty list to store error messages
    error_messages <- list()
    
    # fetch survey
    fetch_and_rename_survey <- function(input_qid) {
        tryCatch({
            ## fetch survey----
            survey <-
                fetch_survey(
                    input_qid,
                    add_column_map = TRUE,
                    include_metadata = c("ResponseId", "RecordedDate"),
                    include_embedded = NA,
                    convert = FALSE,
                    label = FALSE,
                    verbose = FALSE
                )
            ## extract questions from column map and set it as column names----
            colnames(survey) <-
                attributes(survey)[["column_map"]][["description"]]
            
            ## make sure that survey names are unique----
            survey <- survey %>% select(unique(colnames(.)))
            
            
            ## a function to rename columns with a mapping document----
            rename_cols_with_map <- function(survey_df, mapping_df) {
                # Create a named vector where names are longnames and values are shortnames
                name_mapping <-
                    setNames(mapping_df$shortname, mapping_df$longname)
                
                # Vectorized matching and renaming
                # For each column name in survey_df, look for a match in the name_mapping vector
                # If a match is found, use the corresponding value (shortname) from name_mapping
                # If no match is found, keep the original column name
                colnames(survey_df) <-
                    ifelse(
                        colnames(survey_df) %in% names(name_mapping),
                        name_mapping[colnames(survey_df)],
                        colnames(survey_df)
                    )
                
                # Return the updated survey dataframe
                return(survey_df)
            }
            
            
            ## a function to select the columns of interest from the full survey.----
            select_or_create_cols <- function(df) {
                # Ensure col_names is a character vector
                col_names <- mapping_df |> pull(shortname) |> unique()
                
                # Existing column names in df
                existing_cols <- col_names[col_names %in% names(df)]
                
                # Select existing columns, using base R subsetting
                df_selected <- df[, existing_cols, drop = FALSE]
                
                # Identify which columns are missing from the desired col_names list
                missing_cols <- setdiff(col_names, existing_cols)
                
                # Create the missing columns and fill them with NA
                if (length(missing_cols) > 0) {
                    df_selected[missing_cols] <- NA
                }
                
                return(df_selected)
            }
            
            
            ## rename questions as column names, shorten to something readable----
            survey <- rename_cols_with_map(survey, mapping_df)
            
            
            ## select only the columns we want from the full survey
            s <- select_or_create_cols(survey)
            
            ## add a column for qid----
            s <- cbind(qid = input_qid, s)
            
            return(s)
        }, error = function(e) {
            # Append the error message to the error_messages list
            error_messages <<- append(error_messages, paste(input_qid, ":", e$message))
            return(NULL)
        })
    }
    
    wb_path <- wb_path
    
    # List all sheets in the workbook
    sheet_names <- excel_sheets(wb_path)
    
    # Read all sheets into a list of data frames
    sheets_list <- lapply(sheet_names, function(sheet) {
        read_excel(wb_path, sheet = sheet)
    })
    
    # Name the list elements
    names(sheets_list) <- sheet_names
    
    master <- sheets_list$master
    scales <- sheets_list$scales
    mapping_df <- sheets_list$mapping
    
    # store program information in a separate table for later use
    program_info <-
        master %>% select(
            program_name,
            qid,
            modality,
            k12,
            higher_ed,
            early_childhood,
            program_code,
            parent_program_code,
            portfolio,
            full_name,
            iteration_id,
            series,
            start_date,
            start_year,
            start_month,
            end_date,
            regular_individual_pricing,
            regular_team_pricing,
            lift,
            ceus,
            length,
            days_programming,
            response_count,
            final_enrollment,
            response_rate,
            response_rate
        )
    
    # get all the qid's
    master.qid <- master$qid
    
    # fetch all responses in the update list and stack
    updates <-
        purrr::map(master.qid, fetch_and_rename_survey) |> purrr::list_rbind()
    
    # Filter out any NULL values (surveys that failed to fetch)
    updates <- updates[!sapply(updates, is.null), ]
    
    # update program info
    updates <-
        program_info |>
        dplyr::right_join(updates, by = "qid")
    
    
    label_maps <- mapping_df |> select(shortname, scales) |> filter(!is.na(scales)) |> distinct()
    
    scaled_cols <- updates |> select(rid, any_of(label_maps$shortname))
    info_cols <- updates |> select(-any_of(label_maps$shortname))
    
    scaled_cols_long <- scaled_cols |> 
        pivot_longer(cols = label_maps$shortname, 
                     names_to = "metric",
                     values_to = "response")
    
    scaled_cols_long <- scaled_cols_long |> left_join(label_maps, by = c("metric"="shortname"))
    scaled_cols_long$response <- as.character(scaled_cols_long$response)
    scales$value <- as.character(scaled_cols$value)
    
    scaled_cols_long <- scaled_cols_long |> left_join(scales, by = c("scales"="scale","response"="value"))
    
    scaled_cols_long <- scaled_cols_long |> rename(responses_label = label)
    #create binomial responses
    scaled_cols_long <- scaled_cols_long |> mutate(
        response_bin = case_when(
            max == 7 & as.numeric(response) >= 6 ~ 1,
            max == 7 & as.numeric(response) <6 ~ 0,
            max == 5 & as.numeric(response) >= 4 ~ 1,
            max == 5 & as.numeric(response) <4 ~ 0,
            .default = NA_integer_
        )
    )
    
    scaled_cols_labels <- scaled_cols_long |> 
        select(-response_bin, -response, -scales, -min, -max) |> 
        pivot_wider(names_from = metric,
                    values_from = responses_label)
    
    scaled_cols_bin <- scaled_cols_long |> 
        select(-responses_label, -response, -scales, -min, -max) |> 
        pivot_wider(names_from = metric,
                    values_from = response_bin)
    
    update_label <- info_cols |> left_join(scaled_cols_labels, by = "rid")
    update_bin <- info_cols |> left_join(scaled_cols_bin, by = "rid")
    
    wb <- loadWorkbook(wb_path)
    
    removeWorksheet(wb, "stacked")
    removeWorksheet(wb, "stacked_labels")
    removeWorksheet(wb, "stacked_binomial")
    
    addWorksheet(wb, "stacked")
    addWorksheet(wb, "stacked_labels")
    addWorksheet(wb, "stacked_binomial")
    
    # Write the modified data frame to the new worksheet
    writeData(wb, sheet = "stacked", updates)
    writeData(wb, sheet = "stacked_labels", update_label)
    writeData(wb, sheet = "stacked_binomial", update_bin)
    # Save the workbook, overwriting the existing file
    
    saveWorkbook(wb, wb_path, overwrite = TRUE)
    
    # Print error messages, if any
    if (length(error_messages) > 0) {
        print("Errors occurred during the survey fetch:")
        print(error_messages)
    } else {
        print("All surveys fetched successfully!")
    }
}



# update the masters, stacked, and other spreadsheets

update_summaries <- function(wb_path=local_wb_path){
    
    wb_path <- wb_path
    
    # List all sheets in the workbook
    sheet_names <- excel_sheets(wb_path)
    
    # Read all sheets into a list of data frames
    sheets_list <- lapply(sheet_names, function(sheet) {
        read_excel(wb_path, sheet = sheet)
    })
    
    # Name the list elements
    names(sheets_list) <- sheet_names
    
    stacked <- sheets_list$stacked
    master <- sheets_list$master
    
    
    master <- master |> select(-missing_metrics)
    
    # a custom function to handle NPS
    get_nps <- function(vec) {
        vec <- vec |>  na.omit()
        nps <-
            sum(vec == 3) / length(vec) - sum(vec == 1) / length(vec)
        return(round(nps * 100, 2))
    }
    
    #make sure all except ID, date, and modality columns are in numeric format
    stacked <-
        stacked |> mutate(across(
            .cols = -c("program_name", "qid", "rid", "date", "modality"),
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
            )
            #indicators - those are indicator variables converted categoricals, so we'll treat NA as 0
            # across(
            #     .cols = all_of(indicator_items),
            #     .fns = ~ ifelse(is.na(.x), 0, .x)
            # )
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
            c(satisfaction_items,
              amount_items,
              # indicator_items,
              yes_no_items)
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
    
    #find missing KPI for each qid
    metric_cols <-
        c("nps","overall_satisfaction_mn","overall_service_mn","content_relevant_mn","content_enriching_mn","deib_incorporate_mn","deib_represent_mn","community_partof_mn","community_shared_mn","community_network_mn","workload_mn","knowledge_skill_ready_mn","enrolled_with_colleagues_mn","in_us_mn"
        )
    
    m <- kpi_summaries |> select(qid, all_of(metric_cols))
    
    m_long <- m %>%
        pivot_longer(cols = all_of(metric_cols), names_to = "metric", values_to = "value") %>%
        filter(is.na(value))
    
    # Create a new dataframe with id and missing metrics as a comma-separated string
    m_missing <- m_long %>%
        group_by(qid) %>%
        summarise(missing_metrics = toString(metric))
    
    # If you want to join this back to the original dataframe to see all data in one place
    master <- left_join(master, m_missing, by = "qid")
    
    wb <- loadWorkbook(wb_path)
    
    #upload the updated sheet into a preadsheet
    
    removeWorksheet(wb, "kpi_summaries")
    removeWorksheet(wb, "master")
    
    addWorksheet(wb, "kpi_summaries")
    addWorksheet(wb, "master")
    
    # Write the modified data frame to the new worksheet
    writeData(wb, sheet = "kpi_summaries", kpi_summaries)
    writeData(wb, sheet = "master", master)
    # Save the workbook, overwriting the existing file
    
    saveWorkbook(wb, wb_path, overwrite = TRUE)
    
    
    
    print("all updated!!")
    
}

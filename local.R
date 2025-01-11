# This function reads the source file location of this scrit, "local.R", and tries to # set it as the current active directory. Makes it easier to download file. 
set_working_directory_to_source <- function() {
    if (requireNamespace("rstudioapi", quietly = TRUE) && rstudioapi::isAvailable()) {
        script_path <- rstudioapi::getSourceEditorContext()$path
        if (nzchar(script_path)) {
            setwd(dirname(script_path))
            message("Working directory set to source file location: ", getwd())
        } else {
            message("Could not determine the script path.")
        }
    } else {
        message("rstudioapi package is not available or you are not using RStudio.")
    }
}

# Check if the required functions are installed. Install or loads them. 
# Set up api credentials for Qualtrics
setup_and_configure <- function(install=FALSE) {
    packages <- c("qualtRics", "dplyr", "readr", "tidyr", "gtools", "readxl", "openxlsx", "purrr","glue","rstudioapi")
    sapply(packages, function(pkg) {
        if (!requireNamespace(pkg, quietly = TRUE)) {
            message(paste("Installing package:", pkg))
            install.packages(pkg)
        }
        message(paste("Loading package:", pkg))
        library(pkg, character.only = TRUE)
    })
    
    qualtrics_api_credentials(
        api_key = "UacrIljM6AO37NF5f1uSkQgz2uarF3TXYgzybQAP",
        base_url = "sjc1.qualtrics.com",
        overwrite = TRUE,
        install = install
    )
    
    # creating a temporary directory to store survey exports
    temp_dir <<- paste0(getwd(),"/exports") #create and assign globally a temporary directory to store exported surveys
    
    # Check if temp_dir exists, and create it if not
    if (!dir.exists(temp_dir)) {
        dir.create(temp_dir, recursive = TRUE)  # Create directory, including parents if needed
        message(glue::glue("Directory {temp_dir} did not exist and was created."))
    } else {
        message(glue::glue("Directory {temp_dir} already exists."))
    }
    
    message(glue::glue("Downloaded Qualtrics responses will be saved to the temporary directory at {temp_dir}"))
}

# Wrapper around the qualtRics package to fetch surveys from Qualtrics
fetch_and_rename_survey <- function(input_qid, mapping_df, error_messages, testmode = FALSE) {
    
    # Wrapper around the Qualtrics::fetch_survey function to fetch a survey
    
    # # first, fetch survey
    lim <- if (testmode) {3} else (NULL) # conditional limit to extract less rows for testing
    
    tryCatch({
        survey <- fetch_survey(
            input_qid,
            add_column_map = TRUE,
            include_metadata = c("ResponseId", "RecordedDate"),
            include_embedded = NA,
            convert = FALSE,
            label = TRUE,
            verbose = FALSE,
            limit = lim,
            tmp_dir = temp_dir
        )
        
        # extract column metadata from survey, which contains the question text
        column_map <- extract_colmap(survey)
        
        # drop survey columns with all NA responses
        survey <- survey |> select(where(~ !all(is.na(.))))
        
        # drop metadata for non-existant columns
        column_map <- column_map[column_map$qname %in% names(survey), ]
        
        # keep qname as a survey attribute for backup
        attr(survey, "qname") <- column_map$qname
        
        # create named vector to map long names to short names
        name_mapping <- setNames(mapping_df$shortname, mapping_df$longname)
        
        # in metadata, create a shortname column
        column_map$shortname <- name_mapping[column_map$description]
        
        # if shortname doesn't exist, use original description
        column_map$shortname <- coalesce(column_map$shortname, column_map$description)
        
        # set survey column names using the metadata short name column
        colnames(survey) <- make.names(column_map$shortname, unique = TRUE)
        
        # Add missing short names as columns with NA
        missing_cols <- setdiff(col_names, colnames(survey))
        if (length(missing_cols) > 0) {
            survey[missing_cols] <- NA
        }
        
        # Add QID as the first column
        survey <- dplyr::mutate(survey, qid = input_qid, .before = 1)
        
        return(survey)
        
    }, error = function(e) {
        error_messages <<- append(error_messages, paste(input_qid, ":", e$message))
        return(NULL)
    })
}

# reads all sheets from the local workbook (wb) and store them in a list
read_sheets <- function(local_wb_path) {
    sheet_names <- excel_sheets(local_wb_path)
    sheets_list <- lapply(sheet_names, function(sheet) {
        read_excel(local_wb_path, sheet = sheet)
    })
    names(sheets_list) <- sheet_names
    return(sheets_list)
}


# pull out the program info columns from the master sheet so that they can be added to the responses
update_program_info <- function(master, updates) {
    program_info <- master %>% select(
        program_name, qid, modality, k12, higher_ed, early_childhood, program_code,
        parent_program_code, portfolio, full_name, iteration_id, series, start_date,
        start_year, start_month, end_date, regular_individual_pricing, regular_team_pricing,
        lift, ceus, length, days_programming, response_count, final_enrollment, response_rate
    )
    
    updates <- program_info %>% right_join(updates, by = "qid")
    return(updates)
}

process_scaled_columns <- function(updates, mapping_df, scales) {
    label_maps <- mapping_df %>% select(shortname, scales) %>% filter(!is.na(scales)) %>% distinct()
    
    scaled_cols <- updates %>% select(rid, any_of(label_maps$shortname))
    info_cols <- updates %>% select(-any_of(label_maps$shortname))
    
    scaled_cols_long <- scaled_cols %>% 
        pivot_longer(cols = label_maps$shortname, names_to = "metric", values_to = "response")
    
    scaled_cols_long <- scaled_cols_long %>% left_join(label_maps, by = c("metric" = "shortname"))
    scaled_cols_long$response <- as.character(scaled_cols_long$response)
    scales$value <- as.character(scales$value)
    
    scaled_cols_long <- scaled_cols_long %>% left_join(scales, by = c("scales" = "scale", "response" = "value"))
    scaled_cols_long <- scaled_cols_long %>% rename(responses_label = label)
    
    scaled_cols_long <- scaled_cols_long %>% mutate(
        response_bin = case_when(
            max == 7 & as.numeric(response) >= 6 ~ 1,
            max == 7 & as.numeric(response) < 6 ~ 0,
            max == 5 & as.numeric(response) >= 4 ~ 1,
            max == 5 & as.numeric(response) < 4 ~ 0,
            TRUE ~ NA_integer_
        )
    )
    
    scaled_cols_labels <- scaled_cols_long %>% 
        select(-response_bin, -response, -scales, -min, -max) %>% 
        pivot_wider(names_from = metric, values_from = responses_label)
    
    scaled_cols_bin <- scaled_cols_long %>% 
        select(-responses_label, -response, -scales, -min, -max) %>% 
        pivot_wider(names_from = metric, values_from = response_bin)
    
    update_label <- info_cols %>% left_join(scaled_cols_labels, by = "rid")
    update_bin <- info_cols %>% left_join(scaled_cols_bin, by = "rid")
    
    return(list(update_label = update_label, update_bin = update_bin))
}

write_to_workbook <- function(local_wb_path, updates, update_label, update_bin) {
    wb <- loadWorkbook(local_wb_path)
    
    removeWorksheet(wb, "stacked")
    removeWorksheet(wb, "stacked_labels")
    removeWorksheet(wb, "stacked_binomial")
    
    addWorksheet(wb, "stacked")
    addWorksheet(wb, "stacked_labels")
    addWorksheet(wb, "stacked_binomial")
    
    writeData(wb, sheet = "stacked", updates)
    writeData(wb, sheet = "stacked_labels", update_label)
    writeData(wb, sheet = "stacked_binomial", update_bin)
    
    saveWorkbook(wb, local_wb_path, overwrite = TRUE)
}

update_master <- function(wb_path = local_wb_path, mapping_df) {
    error_messages <- list()
    
    # Read Excel sheets
    sheets_list <- read_sheets(wb_path)
    master <- sheets_list$master
    scales <- sheets_list$scales
    mapping_df <- sheets_list$mapping
    
    # Fetch and rename surveys
    master.qid <- master$qid |> unique()
    updates <- purrr::map(master.qid, ~fetch_and_rename_survey(.x, mapping_df, error_messages)) %>% purrr::list_rbind()
    updates <- purrr::compact(updates)
    
    # Update program information
    updates <- update_program_info(master, updates)
    
    # Process scaled columns
    processed_cols <- process_scaled_columns(updates, mapping_df, scales)
    update_label <- processed_cols$update_label
    update_bin <- processed_cols$update_bin
    
    # Write to workbook
    write_to_workbook(wb_path, updates, update_label, update_bin)
    
    # Print error messages, if any
    if (length(error_messages) > 0) {
        print("Errors occurred during the survey fetch:")
        print(error_messages)
    } else {
        print("All surveys fetched successfully!")
    }
}

update_summaries <- function(wb_path = local_wb_path) {
    # Read Excel sheets
    sheets_list <- read_sheets(wb_path)
    stacked <- sheets_list$stacked
    master <- sheets_list$master
    
    master <- master %>% select(-missing_metrics)
    
    # get_nps <- function(vec) {
    #     vec <- vec %>% na.omit()
    #     nps <- sum(vec == 3) / length(vec) - sum(vec == 1) / length(vec)
    #     return(round(nps * 100, 2))
    # }
    
    
    # use a more flexibly nps function that takes both 1-10 and 1-3 as inputs
    get_nps <- function(vec) {
        # Remove NAs first
        vec <- na.omit(vec)
        
        # Determine scale based on max value
        max_val <- max(vec, na.rm = TRUE)
        
        # Validate input based on scale
        if(max_val <= 3) {
            valid_values <- 1:3
            if(!all(vec %in% valid_values)) {
                stop("For 3-point scale, values must be 1 (detractors), 2 (passives), or 3 (promoters)")
            }
            # Calculate NPS for 3-point scale
            promoters <- sum(vec == 3) / length(vec)
            detractors <- sum(vec == 1) / length(vec)
            
        } else if(max_val <= 10) {
            valid_values <- 0:10
            if(!all(vec %in% valid_values)) {
                stop("For 10-point scale, values must be between 0 and 10")
            }
            # Calculate NPS for 10-point scale
            promoters <- sum(vec >= 9) / length(vec)
            detractors <- sum(vec <= 6) / length(vec)
            
        } else {
            stop("Invalid scale detected. Values should be either 1-3 or 0-10")
        }
        
        # Calculate NPS
        nps <- (promoters - detractors) * 100
        return(round(nps, 2))
    }
    
    stacked <- stacked %>% mutate(across(
        .cols = -c("program_name", "qid", "rid", "date", "modality"),
        .fns = as.numeric
    ))
    
    satisfaction_scale_7 <- c(
        "Very dissatisfied", "Moderately dissatisfied", "Slightly dissatisfied",
        "Neither satisfied nor dissatisfied", "Slightly satisfied",
        "Moderately satisfied", "Very satisfied"
    )
    amount_scale_5 <- c("Not at all", "Slightly", "Moderately", "Quite", "Completely")
    valueable_scale_5 <- c("Not at all", "Slightly", "Moderately", "Quite", "Completely")
    
    satisfaction_items <- c("overall_satisfaction", "overall_service")
    amount_items <- c(
        "content_relevant", "content_enriching", "deib_incorporate", "deib_represent",
        "community_partof", "community_shared", "community_network", "workload", "knowledge_skill_ready"
    )
    indicator_items <- c(
        "modality_person", "modality_hybrid", "modality_online_asyn", "modality_online_syn",
        "length_day_1_2", "length_day_3_6", "length_week_1_2", "length_week_3_5", "length_week_6_plus"
    )
    yes_no_items <- c("enrolled_with_colleagues", "in_us")
    
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
            across(
                .cols = all_of(yes_no_items),
                .fns = ~ ifelse(is.na(.x), 0, .x - 1)
            )
        )
    
    stacked <- as.data.frame(stacked)
    
    kpi_summaries <- stacked %>% group_by(qid) %>% summarise(
        n_responses = n(),
        nps = get_nps(nps),
        across(all_of(c(satisfaction_items, amount_items, yes_no_items)),
               ~ round(mean(as.numeric(.x), na.rm = TRUE), 2),
               .names = "{.col}_mn")
    )
    
    kpi_summaries <- master %>% right_join(kpi_summaries, by = "qid", keep = FALSE)
    
    metric_cols <- c(
        "nps", "overall_satisfaction_mn", "overall_service_mn", "content_relevant_mn",
        "content_enriching_mn", "deib_incorporate_mn", "deib_represent_mn", "community_partof_mn",
        "community_shared_mn", "community_network_mn", "workload_mn", "knowledge_skill_ready_mn",
        "enrolled_with_colleagues_mn", "in_us_mn"
    )
    
    m <- kpi_summaries %>% select(qid, all_of(metric_cols))
    m_long <- m %>% pivot_longer(cols = all_of(metric_cols), names_to = "metric", values_to = "value") %>% filter(is.na(value))
    
    m_missing <- m_long %>% group_by(qid) %>% summarise(missing_metrics = toString(metric))
    master <- left_join(master, m_missing, by = "qid")
    
    wb <- loadWorkbook(wb_path)
    
    removeWorksheet(wb, "kpi_summaries")
    removeWorksheet(wb, "master")
    
    addWorksheet(wb, "kpi_summaries")
    addWorksheet(wb, "master")
    
    writeData(wb, sheet = "kpi_summaries", kpi_summaries)
    writeData(wb, sheet = "master", master)
    
    saveWorkbook(wb, wb_path, overwrite = TRUE)
    
    print("All updated!")
}

main <- function() {
    local_wb_path <- path_to_eval_aggregation_workbook
    
    setup_and_configure()
    
    sheets_list <- read_sheets(local_wb_path)
    mapping_df <- sheets_list$mapping
    
    update_master(local_wb_path, mapping_df)
    update_summaries(local_wb_path)
    
    print("Workbook updated!")
}

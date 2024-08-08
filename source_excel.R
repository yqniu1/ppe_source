setup_and_configure <- function() {
    packages <- c("qualtRics", "dplyr", "readr", "tidyr", "gtools", "readxl", "openxlsx", "purrr", "Microsoft365R")
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
        overwrite = TRUE
    )
    
    onedrive <- get_business_onedrive()
    wb_path <- "Documents/PPE Aggregation/eval_aggregation.xlsx"
    local_wb_path <- tempfile(fileext = ".xlsx")
    onedrive$download_file(wb_path, dest = local_wb_path)
    
    list(local_wb_path = local_wb_path, wb_path = wb_path, onedrive = onedrive)
}

fetch_and_rename_survey <- function(input_qid, mapping_df, error_messages) {
    tryCatch({
        survey <- fetch_survey(
            input_qid,
            add_column_map = TRUE,
            include_metadata = c("ResponseId", "RecordedDate"),
            include_embedded = NA,
            convert = FALSE,
            label = FALSE,
            verbose = FALSE
        )
        
        colnames(survey) <- attributes(survey)[["column_map"]][["description"]]
        survey <- survey %>% select(unique(colnames(.)))
        
        rename_cols_with_map <- function(survey_df, mapping_df) {
            name_mapping <- setNames(mapping_df$shortname, mapping_df$longname)
            colnames(survey_df) <- ifelse(
                colnames(survey_df) %in% names(name_mapping),
                name_mapping[colnames(survey_df)],
                colnames(survey_df)
            )
            return(survey_df)
        }
        
        select_or_create_cols <- function(df, col_names) {
            existing_cols <- col_names[col_names %in% names(df)]
            df_selected <- df[, existing_cols, drop = FALSE]
            missing_cols <- setdiff(col_names, existing_cols)
            if (length(missing_cols) > 0) {
                df_selected[missing_cols] <- NA
            }
            return(df_selected)
        }
        
        survey <- rename_cols_with_map(survey, mapping_df)
        col_names <- mapping_df %>% pull(shortname) %>% unique()
        s <- select_or_create_cols(survey, col_names)
        s <- cbind(qid = input_qid, s)
        
        return(s)
    }, error = function(e) {
        error_messages <<- append(error_messages, paste(input_qid, ":", e$message))
        return(NULL)
    })
}

read_sheets <- function(wb_path) {
    sheet_names <- excel_sheets(wb_path)
    sheets_list <- lapply(sheet_names, function(sheet) {
        read_excel(wb_path, sheet = sheet)
    })
    names(sheets_list) <- sheet_names
    return(sheets_list)
}

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

write_to_workbook <- function(wb_path, updates, update_label, update_bin) {
    wb <- loadWorkbook(wb_path)
    
    removeWorksheet(wb, "stacked")
    removeWorksheet(wb, "stacked_labels")
    removeWorksheet(wb, "stacked_binomial")
    
    addWorksheet(wb, "stacked")
    addWorksheet(wb, "stacked_labels")
    addWorksheet(wb, "stacked_binomial")
    
    writeData(wb, sheet = "stacked", updates)
    writeData(wb, sheet = "stacked_labels", update_label)
    writeData(wb, sheet = "stacked_binomial", update_bin)
    
    saveWorkbook(wb, wb_path, overwrite = TRUE)
}

update_master <- function(wb_path = local_wb_path, mapping_df) {
    error_messages <- list()
    
    # Read Excel sheets
    sheets_list <- read_sheets(wb_path)
    master <- sheets_list$master
    scales <- sheets_list$scales
    mapping_df <- sheets_list$mapping
    
    # Fetch and rename surveys
    master.qid <- master$qid
    updates <- purrr::map(master.qid, ~fetch_and_rename_survey(.x, mapping_df, error_messages)) %>% purrr::list_rbind()
    updates <- updates[!sapply(updates, is.null), ]
    
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
    
    get_nps <- function(vec) {
        vec <- vec %>% na.omit()
        nps <- sum(vec == 3) / length(vec) - sum(vec == 1) / length(vec)
        return(round(nps * 100, 2))
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
    config <- setup_and_configure()
    local_wb_path <- config$local_wb_path
    wb_path <- config$wb_path
    onedrive <- config$onedrive
    
    sheets_list <- read_sheets(local_wb_path)
    mapping_df <- sheets_list$mapping
    
    update_master(local_wb_path, mapping_df)
    update_summaries(local_wb_path)
    
    onedrive$upload_file(local_wb_path, dest = wb_path)
    print("Workbook updated and uploaded to OneDrive successfully!")
}

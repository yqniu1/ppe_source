

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
        
        # drop rows that can't be mapped
        column_map <- column_map |> filter(!is.na(shortname))
        
        # drop columns from the survey
        survey <- survey |> select(all_of(column_map[["qname"]]))
        
        # # if shortname doesn't exist, use original description
        # column_map$shortname <- coalesce(column_map$shortname, column_map$description)
        
        # set survey column names using the metadata short name column
        colnames(survey) <- make.names(column_map$shortname, unique = TRUE)
        
        # Add missing short names as columns with NA
        missing_cols <- setdiff(mapping_df$shortname, colnames(survey))
        
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

# reads all sheets from the local workbook (wb) and store them in a list; the masters sheet contains all metadata
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
        program_name, qid, custom, modality, k12, higher_ed, early_childhood, program_code,
        parent_program_code, portfolio, full_name, iteration_id, series, start_date,
        start_year, start_month, end_date, regular_individual_pricing, regular_team_pricing,
        lift, ceus, length, days_programming, response_count, final_enrollment, response_rate
    )
    
    updates <- program_info |>  right_join(updates, by = "qid")
    return(updates)
}

# quick helper to convert nps
get_nps <- function(vec) {
    vec <- vec %>% na.omit()
    nps <- sum(vec == 3) / length(vec) - sum(vec == 1) / length(vec)
    return(round(nps * 100, 2))
}

# This is the main workhorse that takes the raw export and converts it into numeric, binomial, and summarizes the numeric version into kpi
process_scaled_columns <- function(master=master, updates=updates, mapping_df=mapping_df, scales=scales) {
    
    #' Process scaled columns into numeric, binomial, and updates the summary sheet
    #'
    #' @param updates 
    #' @param mapping_df 
    #' @param scales 
    #'
    #' @returns a list with 3 elements 1-responses as numeric, 2-reponse as binomial, 3-response summary
    
    # Filter mapping_df for columns with a numeric scale
    label_maps <- mapping_df |> 
        select(shortname, scales) |>  
        filter(!is.na(scales)) |>  
        distinct()
    cols_to_update <- label_maps |> pull(shortname)
    
    # Create mappings for numeric and binomial conversions
    label_to_value <- scales |> 
        select(label, value) |> 
        na.omit() |> 
        distinct()
    label_to_value$value <- as.character(label_to_value$value)
    label_to_value <- setNames(label_to_value$value, label_to_value$label)
    
    label_to_bin <- scales |> 
        select(label, bin) |> 
        na.omit() |> 
        distinct()
    label_to_bin$bin <- as.character(label_to_bin$bin)
    label_to_bin <- setNames(label_to_bin$bin, label_to_bin$label)
    
    # Filter cols_to_update to only those present in updates
    cols_to_update <- intersect(cols_to_update, names(updates))

    # Convert to numeric
    update_numeric <- updates |> 
        select(all_of(c("qid", "rid", cols_to_update)))
    
    update_numeric[cols_to_update] <- lapply(update_numeric[cols_to_update], function(col) {
        unname(label_to_value[as.character(col)])
    })

    
    # Convert to binomial
    update_binomial <- updates |> 
        select(all_of(c("qid", "rid", setdiff(cols_to_update, "nps"))))
    
    # Apply binomial conversion
    cols_to_update_bin <- setdiff(cols_to_update, "nps")
    update_binomial[cols_to_update_bin] <- lapply(update_binomial[cols_to_update_bin], function(col) {
        # Convert column to character and map labels to binomial values
        unname(label_to_bin[as.character(col)])
    })

    
    # create summary using the update_numeric table
    kpi_summaries <- update_numeric |>  
        group_by(qid) |>  
        summarise(
            n_responses = n(),
            nps_score = get_nps(nps),
            across(
                setdiff(cols_to_update, "nps"),
                ~ round(mean(as.numeric(.x), na.rm = TRUE), 2),
                .names = "{.col}_mn"
            )
        )
    
    # Return results as a list
    update_list = list(update_label = updates, update_numeric = update_numeric, update_binomial = update_binomial, update_summaries = kpi_summaries)
    
    # Apply the function to each updates dataframe in the list
    update_outputs <- lapply(update_list, function(updates) {
        update_program_info(master, updates)
    })
    
    return(update_outputs)
    
}

# this is a helper that reads the existing eval_aggregation_workbook
write_to_workbook <- function(local_wb_path, update_numeric, update_label, update_bin, update_summaries ) {
    wb <- loadWorkbook(local_wb_path)
    
    removeWorksheet(wb, "stacked_label")
    removeWorksheet(wb, "stacked_numeric")
    removeWorksheet(wb, "stacked_binomial")
    removeWorksheet(wb, "kpi_summaries")
    
    
    addWorksheet(wb, "stacked_label")
    addWorksheet(wb, "stacked_numeric")
    addWorksheet(wb, "stacked_binomial")
    addWorksheet(wb, "kpi_summaries")
    
    writeData(wb, sheet = "stacked_label", update_label)
    writeData(wb, sheet = "stacked_numeric", update_numeric)
    writeData(wb, sheet = "stacked_binomial", update_bin)
    writeData(wb, sheet = "kpi_summaries", update_summaries)
    
    saveWorkbook(wb, local_wb_path, overwrite = TRUE)
}

update_master <- function(sheets = sheets_list, test=testmode) {
    
    # Read Excel sheets
    master <- sheets_list$master
    scales <- sheets_list$scales
    mapping_df <- sheets_list$mapping

    # create error log
    error_messages <- list()
    
    # Fetch and rename surveys
    master.qid <- master$qid |> unique()
    updates <- purrr::map(master.qid, ~fetch_and_rename_survey(.x, mapping_df, error_messages, test)) %>% purrr::list_rbind()
    updates <- purrr::compact(updates)
    
    # Process scaled columns
    output_list <- process_scaled_columns(master=master, updates=updates, mapping_df=mapping_df, scales=scales)
    update_numeric <- output_list$update_numeric
    update_bin <- output_list$update_bin
    update_label <- output_list$update_label
    update_summaries <- output_list$update_summaries
    
    # Write to workbook
    write_to_workbook(local_wb_path=local_wb_path, update_numeric=update_numeric, update_label=update_label, update_bin=update_bin, update_summaries=update_summaries)
    
    # Print error messages, if any
    if (length(error_messages) > 0) {
        print("Errors occurred during the survey fetch:")
        print(error_messages)
    } else {
        print("All surveys fetched successfully!")
    }
}

main <- function(testmode = FALSE) {
    local_wb_path <<- path_to_eval_aggregation_workbook
    
    sheets_list <<- read_sheets(local_wb_path)

    update_master(test = testmode)

    print("Workbook updated!")
}

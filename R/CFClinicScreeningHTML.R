# Install required packages if not already installed
required_packages <- c("shiny", "readr", "tidyr", "dplyr", "readxl", "openxlsx", "stringr", "purrr", "stringdist")
for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg)
    cat("Installed package:", pkg, "\n")
  }
}

# Load libraries
library(shiny)
library(readr)
library(tidyr)
library(dplyr)
library(readxl)
library(openxlsx)
library(stringr)
library(purrr)
library(stringdist)

# Define UI
ui <- fluidPage(
  titlePanel("CF Clinic Data Processing"),
  sidebarLayout(
    sidebarPanel(
      h3("Upload Excel File"),
      fileInput("inputFile", "Choose Excel File (will be saved as R.SLopez.xlsx)",
                accept = c(".xlsx", ".xls")
      ),
      actionButton("runScript", "Run Processing Script"),
      br(),
      br(),
      downloadButton("downloadOutput", "Download Output File")
    ),
    mainPanel(
      h4("Script Output"),
      verbatimTextOutput("scriptOutput"),
      h4("Status"),
      textOutput("statusMessage")
    )
  )
)

# Define server logic
server <- function(input, output, session) {
  # Reactive value to store script output
  script_result <- reactiveVal(NULL)
  output_file <- reactiveVal(NULL)
  
  # Handle file upload
  observeEvent(input$inputFile, {
    req(input$inputFile)
    tryCatch({
      # Define target path
      target_path <- file.path("~/Desktop/R Studio/CF Clinic", "R.SLopez.xlsx")
      # Copy uploaded file to R.SLopez.xlsx
      file.copy(input$inputFile$datapath, target_path, overwrite = TRUE)
      output$statusMessage <- renderText({
        paste("File uploaded and saved as", target_path)
      })
    }, error = function(e) {
      output$statusMessage <- renderText({
        paste("Error saving file:", e$message)
      })
    })
  })
  
  # Run script when button is clicked
  observeEvent(input$runScript, {
    output$statusMessage <- renderText({"Running script..."})
    
    # Define the processing script as a function
    process_data <- function() {
      setwd("~/Desktop/R Studio/CF Clinic")
      
      # Verify file existence
      if (!file.exists("R.SLopez.xlsx")) stop("File R.SLopez.xlsx not found in ", getwd())
      if (!file.exists("Barnes.CFMutationWork.xlsx")) stop("File Barnes.CFMutationWork.xlsx not found in ", getwd())
      if (!file.exists("Copy of NBSA Project Disorders and Variants of Interest_Jan 2025.xlsx")) {
        stop("File Copy of NBSA Project Disorders and Variants of Interest_Jan 2025.xlsx not found in ", getwd())
      }
      
      # Get all sheet names from R.SLopez.xlsx
      sheet_names <- excel_sheets("R.SLopez.xlsx")
      cat("Sheets found in R.SLopez.xlsx:", paste(sheet_names, collapse = ", "), "\n")
      
      # Input other Excel files
      BarnesCF <- read_excel("Barnes.CFMutationWork.xlsx")
      CDCList <- read_excel("Copy of NBSA Project Disorders and Variants of Interest_Jan 2025.xlsx", sheet = "CF Priority Groups")
      
      # Check column names in all datasets for debugging
      cat("\nColumn names in BarnesCF:\n")
      print(colnames(BarnesCF))
      cat("\nColumn names in CDCList (CF Priority Groups):\n")
      print(colnames(CDCList))
      
      # Define column names for BarnesCF
      barnes_firstname_col <- "Firstname"
      barnes_lastname_col <- "Lastname"
      barnes_mutation_col <- "Mutation 1"
      if (!all(c(barnes_firstname_col, barnes_lastname_col, barnes_mutation_col) %in% colnames(BarnesCF))) {
        stop("One or more columns (", barnes_firstname_col, ", ", barnes_lastname_col, ", ", 
             barnes_mutation_col, ") not found in BarnesCF.")
      }
      
      # Process CDCList
      cdc_values <- CDCList %>%
        as.data.frame() %>%
        unlist() %>%
        na.omit() %>%
        tolower() %>%
        trimws()
      cat("Sample cdc_values (first 5):\n")
      print(head(cdc_values, 5))
      
      # Extract first four characters of Mutation 1 from BarnesCF
      BarnesCF <- BarnesCF %>%
        mutate(
          Mutation_4char = tolower(str_sub(.[[barnes_mutation_col]], 1, 4)),
          Name_order1 = paste(tolower(.[[barnes_lastname_col]]), tolower(.[[barnes_firstname_col]])),
          Name_order2 = paste(tolower(.[[barnes_firstname_col]]), tolower(.[[barnes_lastname_col]]))
        )
      
      # Debug: Print sample BarnesCF data
      cat("Sample BarnesCF data (including Mutation_4char):\n")
      print(head(select(BarnesCF, all_of(c(barnes_firstname_col, barnes_lastname_col, "Mutation_4char", "Name_order1", "Name_order2"))), 5))
      cat("Number of rows in BarnesCF with NA Mutation_4char:", sum(is.na(BarnesCF$Mutation_4char)), "\n")
      
      # Initialize list for processed data
      all_new_og_data <- list()
      
      # Loop through each sheet
      for (sheet in sheet_names) {
        cat("\nProcessing sheet:", sheet, "\n")
        
        OGdata <- read_excel("R.SLopez.xlsx", sheet = sheet)
        
        cat("Column names in OGdata (sheet:", sheet, "):\n")
        print(colnames(OGdata))
        
        if (!"Name" %in% colnames(OGdata)) {
          cat("Warning: 'Name' column not found in sheet", sheet, ". Skipping this sheet.\n")
          next
        }
        
        if ("Chest Xray Q2" %in% colnames(OGdata)) {
          OGdata <- OGdata %>%
            mutate(`Chest Xray Q2` = as.character(`Chest Xray Q2`))
          cat("Converted 'Chest Xray Q2' to character in sheet:", sheet, "\n")
        }
        
        OGdata <- OGdata %>%
          mutate(Name_clean = tolower(trimws(gsub("[[:punct:]]", " ", Name)))) %>%
          filter(!is.na(Name_clean) & nzchar(Name_clean))
        
        cat("Sample Name_clean in OGdata (sheet:", sheet, "):\n")
        print(head(OGdata$Name_clean, 5))
        
        NewOGdata <- OGdata %>%
          mutate(`On Barnes List?` = case_when(
            mapply(function(name) {
              name_words <- str_split(trimws(name), "\\s+")[[1]]
              if (length(name_words) < 2 || any(is.na(name_words)) || any(!nzchar(name_words))) return(FALSE)
              word_pairs <- combn(name_words, 2, simplify = FALSE)
              any(
                sapply(word_pairs, function(pair) {
                  valid_names <- c(tolower(BarnesCF[[barnes_lastname_col]]), tolower(BarnesCF[[barnes_firstname_col]]))
                  valid_names <- valid_names[!is.na(valid_names) & nzchar(valid_names)]
                  if (length(valid_names) == 0) return(FALSE)
                  match1 <- any(stringdist(pair[1], valid_names, method = "lv") <= 1)
                  match2 <- any(stringdist(pair[2], valid_names, method = "lv") <= 1)
                  match1 && match2
                })
              )
            }, Name_clean) ~ "YES",
            TRUE ~ "NO"
          )) %>%
          left_join({
            BarnesCF_pairs <- BarnesCF %>%
              mutate(Name_clean_barnes = tolower(trimws(paste(.[[barnes_lastname_col]], .[[barnes_firstname_col]])))) %>%
              filter(!is.na(Name_clean_barnes) & nzchar(Name_clean_barnes)) %>%
              select(Name_clean_barnes, Mutation_4char)
            OGdata %>%
              mutate(
                Mutation_4char = map_chr(Name_clean, function(name) {
                  if (is.na(name) || !nzchar(name)) return(NA_character_)
                  distances <- stringdist(name, BarnesCF_pairs$Name_clean_barnes, method = "lv")
                  if (any(distances <= 2, na.rm = TRUE)) {
                    match_idx <- which.min(distances)
                    return(BarnesCF_pairs$Mutation_4char[match_idx])
                  }
                  NA_character_
                })
              ) %>%
              select(Name_clean, Mutation_4char)
          }, by = "Name_clean") %>%
          mutate(`CDC Eligible?` = if_else(
            !is.na(Mutation_4char) & nzchar(Mutation_4char) & sapply(Mutation_4char, function(x) {
              if (is.na(x) || !nzchar(x)) return(FALSE)
              matches <- str_detect(cdc_values, fixed(x))
              if (any(matches, na.rm = TRUE)) {
                cat("Mutation_4char:", x, "matches cdc_values:", cdc_values[matches], "\n")
              }
              any(matches, na.rm = TRUE)
            }),
            "YES",
            "NO"
          )) %>%
          select(-any_of("On barnes?")) %>%
          select(`On Barnes List?`, `CDC Eligible?`, Mutation_4char, everything(), -Name_clean) %>%
          mutate(Sheet = sheet)
        
        # Debug
        cat("Columns in NewOGdata after joins (sheet:", sheet, "):\n")
        print(colnames(NewOGdata))
        na_barnes_check <- NewOGdata %>% filter(is.na(`On Barnes List?`))
        if (nrow(na_barnes_check) > 0) {
          cat("Rows with NA in On Barnes List? in sheet:", sheet, "\n")
          print(na_barnes_check %>% select(Name, `On Barnes List?`, Mutation_4char))
        }
        cdc_eligible_check <- NewOGdata %>% filter(`CDC Eligible?` == "YES")
        if (nrow(cdc_eligible_check) > 0) {
          cat("Participants with CDC Eligible? = YES in sheet:", sheet, "\n")
          print(cdc_eligible_check %>% select(Name, `On Barnes List?`, `CDC Eligible?`, Mutation_4char))
        }
        cat("All Mutation_4char values in NewOGdata (sheet:", sheet, "):\n")
        print(table(NewOGdata$Mutation_4char, useNA = "always"))
        
        all_new_og_data[[sheet]] <- NewOGdata
        
        cat("\nUpdated NewOGdata for sheet:", sheet, "\n")
        print(head(NewOGdata, 3))
        
        matches_found <- NewOGdata %>% filter(`On Barnes List?` == "YES" | `CDC Eligible?` == "YES")
        if (nrow(matches_found) > 0) {
          cat("Rows with matches in BarnesCF or CDC Eligible for sheet:", sheet, "\n")
          print(head(matches_found, 3))
        } else {
          cat("No matches found in sheet:", sheet, "\n")
        }
      }
      
      # Check data types
      cat("\nChecking data types across sheets:\n")
      for (sheet in names(all_new_og_data)) {
        cat("Data types in sheet:", sheet, "\n")
        print(sapply(all_new_og_data[[sheet]], class))
      }
      
      # Standardize DOB column to Date format if present
      all_new_og_data <- lapply(all_new_og_data, function(df) {
        if ("DOB" %in% names(df)) {
          suppressWarnings(df$DOB <- as.Date(df$DOB))
        }
        return(df)
      })
      
      # Combine all sheets
      NewOGdata_combined <- bind_rows(all_new_og_data)
      
      # Check columns
      cat("\nChecking columns in NewOGdata_combined:\n")
      if ("On barnes?" %in% colnames(NewOGdata_combined)) {
        cat("Warning: 'On barnes?' is still present in NewOGdata_combined.\n")
      } else {
        cat("'On barnes?' was successfully removed or was not present.\n")
      }
      if ("On Barnes List?" %in% colnames(NewOGdata_combined)) {
        cat("'On Barnes List?' is present in NewOGdata_combined as expected.\n")
      } else {
        cat("Error: 'On Barnes List?' is missing from NewOGdata_combined.\n")
      }
      if ("Mutation_4char" %in% colnames(NewOGdata_combined)) {
        cat("Mutation_4char is present in NewOGdata_combined as expected.\n")
      } else {
        cat("Error: Mutation_4char is missing from NewOGdata_combined.\n")
      }
      cat("\nColumn names in NewOGdata_combined:\n")
      print(colnames(NewOGdata_combined))
      
      # Save combined data
      openxlsx::write.xlsx(NewOGdata_combined, file = "Modified_OGdata_AllSheets.xlsx", rowNames = FALSE)
      
      return(NewOGdata_combined)
    }
    
    # Execute script and capture output
    tryCatch({
      # Create a temporary file for logging
      log_file <- tempfile(fileext = ".txt")
      con <- file(log_file, open = "wt")
      
      # Sink both output and messages to the connection
      sink(con, type = "output")
      sink(con, type = "message", append = TRUE)
      
      # Run the script
      result <- process_data()
      
      # Close sinks
      sink(type = "message")
      sink(type = "output")
      close(con)
      
      # Read captured output
      log_output <- readLines(log_file)
      unlink(log_file)
      
      # Store results
      script_result(log_output)
      output_file(result)
      
      # Update UI
      output$scriptOutput <- renderText({
        paste(log_output, collapse = "\n")
      })
      output$statusMessage <- renderText({
        "Script completed successfully. Output file is ready for download."
      })
      
      # Provide download handler
      output$downloadOutput <- downloadHandler(
        filename = function() {
          "Modified_OGdata_AllSheets.xlsx"
        },
        content = function(file) {
          openxlsx::write.xlsx(result, file, rowNames = FALSE)
        }
      )
      
    }, error = function(e) {
      # Ensure sinks are closed on error
      if (sink.number(type = "message") > 0) sink(type = "message")
      if (sink.number(type = "output") > 0) sink(type = "output")
      if (exists("con")) try(close(con), silent = TRUE)
      if (file.exists(log_file)) unlink(log_file)
      
      output$statusMessage <- renderText({
        paste("Error running script:", e$message)
      })
      output$scriptOutput <- renderText({
        "Error occurred. See status message."
      })
    })
  })
}

# Run the Shiny app
shinyApp(ui = ui, server = server)
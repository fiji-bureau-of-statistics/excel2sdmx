#' Run the SDMX Converter Dashboard
#'
#' @export
run_fbosSDMX <- function() {

  # -------------------------------------------------
  # Load libraries
  # -------------------------------------------------
  library(shiny)
  library(shinydashboard)
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(stringr)
  library(zip)

  # -------------------------------------------------
  # UI
  # -------------------------------------------------
  ui <- dashboardPage(

    dashboardHeader(title = "SDMX Converter"),

    dashboardSidebar(
      sidebarMenu(
        menuItem("BOP Converter", tabName = "bop", icon = icon("balance-scale")),
        menuItem("CPI Converter", tabName = "cpi", icon = icon("chart-line")),
        menuItem("National Accoutns Converter", tabName = "national", icon = icon("landmark")),
        menuItem("Poverty Converter", tabName = "poverty", icon = icon("scale-balanced")),
        menuItem("Visitors Converter", tabName = "visitor", icon = icon("plane")),
        menuItem("About", tabName = "about", icon = icon("info-circle"))
      )
    ),

    dashboardBody(
      tabItems(

        # ---------------- BOP TAB ----------------
        tabItem(
          tabName = "bop",

          fluidRow(

            box(
              title = "Upload Files",
              status = "primary",
              solidHeader = TRUE,
              width = 4,

              fileInput("excel_file",
                        "Upload BOP Excel file (.xlsx)",
                        accept = ".xlsx"),

              fileInput("codelist_file",
                        "Upload bopAccounts.csv",
                        accept = ".csv"),

              actionButton("process", "Process File", icon = icon("play")),
              br(), br(),
              downloadButton("download_zip", "Download Output ZIP")
            ),

            box(
              title = "BOP Processing Log",
              status = "warning",
              solidHeader = TRUE,
              width = 8,
              verbatimTextOutput("log")
            )
          ),

          fluidRow(
            box(
              title = "BOP Preview (First 10 Rows)",
              status = "success",
              solidHeader = TRUE,
              width = 12,
              tableOutput("preview")
            )
          )
        ),

        # ---------------- CPI TAB ----------------
        tabItem(
          tabName = "cpi",

          fluidRow(

            box(
              title = "Upload CPI Excel File",
              status = "primary",
              solidHeader = TRUE,
              width = 4,

              fileInput("cpi_excel_file",
                        "Upload CPI Excel file (.xlsx)",
                        accept = ".xlsx"),

              actionButton("process_cpi", "Process CPI File", icon = icon("play")),
              br(), br(),
              downloadButton("download_cpi_zip", "Download CPI ZIP")
            ),

            box(
              title = "CPI Processing Log",
              status = "warning",
              solidHeader = TRUE,
              width = 8,
              verbatimTextOutput("cpi_log")
            )
          ),

          fluidRow(
            box(
              title = "CPI Preview (First 10 Rows)",
              status = "success",
              solidHeader = TRUE,
              width = 12,
              tableOutput("cpi_preview")
            )
          )
        ),

        # ---------------- NATIONAL ACCOUNTS TAB ----------------
        tabItem(
          tabName = "national",

          fluidRow(

            box(
              title = "Upload National Accounts Excel File",
              status = "primary",
              solidHeader = TRUE,
              width = 4,

              fileInput("na_excel_file",
                        "Upload NA Excel file (.xlsx)",
                        accept = ".xlsx"),

              actionButton("process_na", "Process NA File", icon = icon("play")),
              br(), br(),
              downloadButton("download_na_zip", "Download NA ZIP")
            ),

            box(
              title = "National Accounts Processing Log",
              status = "warning",
              solidHeader = TRUE,
              width = 8,
              verbatimTextOutput("na_log")
            )
          ),

          fluidRow(
            box(
              title = "National Accounts Preview (First 10 Rows)",
              status = "success",
              solidHeader = TRUE,
              width = 12,
              tableOutput("na_preview")
            )
          )
        ),

        # ---------------- POVERTY TAB ----------------
        tabItem(
          tabName = "poverty",

          fluidRow(

            box(
              title = "Upload Poverty Excel File",
              status = "primary",
              solidHeader = TRUE,
              width = 4,

              fileInput("poverty_excel_file",
                        "Upload Poverty Excel file (.xlsx)",
                        accept = ".xlsx"),

              actionButton("process_poverty", "Process Poverty File", icon = icon("play")),
              br(), br(),
              downloadButton("download_poverty_zip", "Download Poverty ZIP")
            ),

            box(
              title = "Poverty Processing Log",
              status = "warning",
              solidHeader = TRUE,
              width = 8,
              verbatimTextOutput("poverty_log")
            )
          ),

          fluidRow(
            box(
              title = "Poverty Preview (First 10 Rows)",
              status = "success",
              solidHeader = TRUE,
              width = 12,
              tableOutput("poverty_preview")
            )
          )
        ),

        # ---------------- VISITORS TAB ----------------
        tabItem(
          tabName = "visitor",

          fluidRow(

            box(
              title = "Upload Visitors Excel File",
              status = "primary",
              solidHeader = TRUE,
              width = 4,

              fileInput("visitor_excel_file",
                        "Upload Visitors Excel file (.xlsx)",
                        accept = ".xlsx"),

              actionButton("process_visitor", "Process Visitors File", icon = icon("play")),
              br(), br(),
              downloadButton("download_visitor_zip", "Download Visitors ZIP")
            ),

            box(
              title = "Visitors arrival Processing Log",
              status = "warning",
              solidHeader = TRUE,
              width = 8,
              verbatimTextOutput("visitor_log")
            )
          ),

          fluidRow(
            box(
              title = "Visitors arrival Preview (First 10 Rows)",
              status = "success",
              solidHeader = TRUE,
              width = 12,
              tableOutput("visitor_preview")
            )
          )
        ),

        # ---------------- ABOUT TAB ----------------
        tabItem(
          tabName = "about",
          fluidRow(
            box(
              title = "About This Application",
              status = "info",
              solidHeader = TRUE,
              width = 12,
              p("This application converts Excel files into SDMX-compliant CSV format."),
              p("Upload your excel file and any other needed csv files of codelist."),
              p("All sheets will be processed automatically."),
              p("After converting the excel file, you can download either a csv file or zip file with several csv files"),
              p("The excel files that can be processed includes the following:"),
              p("1. Balance of payment (BOP)"),
              p("2. Consumer Price Index (CPI)"),
              p("3. National Accounts (NA)"),
              p("4. Poverty Statistics"),
              p("5. Visitors Arrivals")
            )
          )
        )
      )
    )
  )

  # -------------------------------------------------
  # SERVER
  # -------------------------------------------------
  server <- function(input, output, session) {

    log_text <- reactiveVal("")
    output_files <- reactiveVal(NULL)
    preview_data <- reactiveVal(NULL)

    add_log <- function(msg) {
      log_text(paste(log_text(), msg, sep = "\n"))
    }

    # ---------------- BOP PROCESSING ----------------

    observeEvent(input$process, {

      req(input$excel_file)
      req(input$codelist_file)

      log_text("")
      add_log("Starting processing...")

      bopAcc <- read.csv(input$codelist_file$datapath)
      file_path <- input$excel_file$datapath
      sheet_names <- excel_sheets(file_path)

      temp_dir <- tempdir()
      created_files <- c()

      for (sheet in sheet_names) {

        add_log(paste("Processing sheet:", sheet))

        table <- read_excel(file_path, sheet = sheet)
        table <- merge(table, bopAcc, by = "label")

        table <- table |>
          relocate(ACCOUNT, .before = UNIT_MEASURE) |>
          arrange(order) |>
          select(-order, -label)

        table_long <- table |>
          pivot_longer(
            cols = -c(DATAFLOW:DECIMALS),
            names_to = "TIME_PERIOD",
            values_to = "OBS_VALUE"
          ) |>
          mutate(
            across(everything(), ~replace(., is.na(.), "")),

            FREQ = case_when(
              grepl("-Q[1-4]", TIME_PERIOD) ~ "Q",
              grepl("-0[1-9]|-1[0-2]", TIME_PERIOD) ~ "M",
              TRUE ~ "A"
            ),

            OBS_STATUS = case_when(
              str_detect(TIME_PERIOD, "\\(YTD\\)") ~ "YTD",
              str_detect(TIME_PERIOD, "\\(P\\)") ~ "P",
              str_detect(TIME_PERIOD, "\\(R\\)") ~ "R",
              TRUE ~ ""
            ),

            TIME_PERIOD = str_trim(
              str_remove_all(TIME_PERIOD,
                             "\\s*\\(YTD\\)|\\s*\\(P\\)|\\s*\\(R\\)")
            )
          ) |>
          relocate(OBS_VALUE, .before = UNIT_MEASURE) |>
          relocate(TIME_PERIOD, .before = OBS_VALUE)

        out_file <- file.path(temp_dir, paste0(sheet, ".csv"))
        write.csv(table_long, out_file, row.names = FALSE)

        created_files <- c(created_files, out_file)
        preview_data(head(table_long, 10))
      }

      output_files(created_files)
      add_log("Processing completed successfully ✔")
    })

    output$preview <- renderTable({
      req(preview_data())
      preview_data()
    })

    output$log <- renderText({
      log_text()
    })

    output$download_zip <- downloadHandler(
      filename = function() {
        paste0("BOP_output_", Sys.Date(), ".zip")
      },
      content = function(file) {

        req(output_files())

        temp_zip <- tempfile(fileext = ".zip")

        zip::zip(
          zipfile = temp_zip,
          files = output_files(),
          mode = "cherry-pick"
        )

        file.copy(temp_zip, file)
      },
      contentType = "application/zip"
    )

    # ---------------- CPI PROCESSING ----------------

    cpi_log_text <- reactiveVal("")
    cpi_output_files <- reactiveVal(NULL)
    cpi_preview_data <- reactiveVal(NULL)

    add_cpi_log <- function(msg) {
      cpi_log_text(paste(cpi_log_text(), msg, sep = "\n"))
    }

    observeEvent(input$process_cpi, {

      req(input$cpi_excel_file)

      cpi_log_text("")
      add_cpi_log("Starting CPI processing...")

      file_path <- input$cpi_excel_file$datapath
      sheet_names <- excel_sheets(file_path)

      temp_dir <- tempdir()
      created_files <- c()

      for (sheet in sheet_names) {

        add_cpi_log(paste("Processing sheet:", sheet))

        table <- read_excel(file_path, sheet = sheet)

        table_long <- table %>%
          pivot_longer(
            cols = -c(DATAFLOW:BASE_PER),
            names_to = "ITEM",
            values_to = "OBS_VALUE"
          )

        table_long <- table_long |>
          mutate(
            across(everything(), ~replace(., is.na(.), "")),
            SEASONAL_ADJUST = ifelse(ITEM == "S_T", "S", SEASONAL_ADJUST),
            ITEM = ifelse(ITEM == "S_T", "_T", ITEM)
          ) |>
          relocate(OBS_VALUE, .before = UNIT_MEASURE) |>
          relocate(ITEM, .before = TRANSFORMATION)

        out_file <- file.path(temp_dir, paste0(sheet, ".csv"))
        write.csv(table_long, out_file, row.names = FALSE)

        created_files <- c(created_files, out_file)
        cpi_preview_data(head(table_long, 10))
      }

      cpi_output_files(created_files)
      add_cpi_log("CPI processing completed successfully ✔")
    })

    output$cpi_preview <- renderTable({
      req(cpi_preview_data())
      cpi_preview_data()
    })

    output$cpi_log <- renderText({
      cpi_log_text()
    })

    output$download_cpi_zip <- downloadHandler(
      filename = function() {
        paste0("CPI_output_", Sys.Date(), ".zip")
      },
      content = function(file) {

        req(cpi_output_files())

        temp_zip <- tempfile(fileext = ".zip")

        zip::zip(
          zipfile = temp_zip,
          files = cpi_output_files(),
          mode = "cherry-pick"
        )

        file.copy(temp_zip, file)
      },
      contentType = "application/zip"
    )

    # ---------------- NATIONAL ACCOUNTS PROCESSING ----------------

    na_log_text <- reactiveVal("")
    na_output_files <- reactiveVal(NULL)
    na_preview_data <- reactiveVal(NULL)

    add_na_log <- function(msg) {
      na_log_text(paste(na_log_text(), msg, sep = "\n"))
    }

    observeEvent(input$process_na, {

      req(input$na_excel_file)

      na_log_text("")
      add_na_log("Starting National Accounts processing...")

      file_path <- input$na_excel_file$datapath
      sheet_names <- excel_sheets(file_path)

      temp_dir <- tempdir()
      created_files <- c()

      for (sheet in sheet_names) {

        add_na_log(paste("Processing sheet:", sheet))

        table <- read_excel(file_path, sheet = sheet)

        table_long <- table |>
          pivot_longer(
            cols = -c(DATAFLOW:DECIMALS),
            names_to = "TIME_PERIOD",
            values_to = "OBS_VALUE"
          ) |>
          mutate(
            across(everything(), ~replace(., is.na(.), "")),
            TRANSFORMATION = ifelse(TIME_PERIOD == "Weight", "WGT", TRANSFORMATION),
            UNIT_MEASURE = ifelse(TIME_PERIOD == "Weight", "PT", UNIT_MEASURE),
            UNIT_MULT = ifelse(TIME_PERIOD == "Weight", "", UNIT_MULT),
            INDUSTRY = ifelse(TIME_PERIOD == "Weight", "_T", INDUSTRY),
            INDICATOR = ifelse(TIME_PERIOD == "Weight", substr(INDICATOR, 1, 4), INDICATOR),
            GDP_BREAKDOWN = ifelse(TIME_PERIOD == "Weight", "_T", GDP_BREAKDOWN),
            TIME_PERIOD = ifelse(TIME_PERIOD == "Weight", "2014", TIME_PERIOD),
            OBS_STATUS = case_when(
              str_detect(TIME_PERIOD, "\\(YTD\\)") ~ "YTD",
              str_detect(TIME_PERIOD, "\\(P\\)") ~ "P",
              str_detect(TIME_PERIOD, "\\(R\\)") ~ "R",
              TRUE ~ ""),
            TIME_PERIOD = str_trim(str_remove_all(TIME_PERIOD, "\\s*\\(YTD\\)|\\s*\\(P\\)|\\s*\\(R\\)")),
            across(everything(), ~replace(., is.na(.), ""))
          ) |>
          relocate(OBS_VALUE, .before = UNIT_MEASURE) |>
          relocate(TIME_PERIOD, .before = TRANSFORMATION) |>
          filter(!is.na(OBS_VALUE) & OBS_VALUE != "")

        out_file <- file.path(temp_dir, paste0(sheet, ".csv"))
        write.csv(table_long, out_file, row.names = FALSE)

        created_files <- c(created_files, out_file)
        na_preview_data(head(table_long, 10))
      }

      na_output_files(created_files)
      add_na_log("National Accounts processing completed successfully ✔")
    })

    output$na_preview <- renderTable({
      req(na_preview_data())
      na_preview_data()
    })

    output$na_log <- renderText({
      na_log_text()
    })

    output$download_na_zip <- downloadHandler(
      filename = function() {
        paste0("NA_output_", Sys.Date(), ".zip")
      },
      content = function(file) {

        req(na_output_files())

        temp_zip <- tempfile(fileext = ".zip")

        zip::zip(
          zipfile = temp_zip,
          files = na_output_files(),
          mode = "cherry-pick"
        )

        file.copy(temp_zip, file)
      },
      contentType = "application/zip"
    )

    # ---------------- POVERTY PROCESSING ----------------

    poverty_log_text <- reactiveVal("")
    poverty_output_files <- reactiveVal(NULL)
    poverty_preview_data <- reactiveVal(NULL)

    add_poverty_log <- function(msg) {
      poverty_log_text(paste(poverty_log_text(), msg, sep = "\n"))
    }

    observeEvent(input$process_poverty, {

      req(input$poverty_excel_file)

      poverty_log_text("")
      add_poverty_log("Starting Poverty Statistics processing...")

      file_path <- input$poverty_excel_file$datapath
      sheet_names <- excel_sheets(file_path)

      temp_dir <- tempdir()
      created_files <- c()

      for (sheet in sheet_names) {

        add_poverty_log(paste("Processing sheet:", sheet))

        table <- read_excel(file_path, sheet = sheet) |> select(-Area)

        # Pivot longer
        table_long <- table |>
          pivot_longer(
            cols = -c(DATAFLOW:DECIMALS),
            names_to = "INDICATOR",
            values_to = "OBS_VALUE"
          ) |>
          mutate(
            across(everything(), ~replace(., is.na(.), "")),
            FREQ = case_when(
              grepl("-Q[1-4]", TIME_PERIOD) ~ "Q",
              grepl("-0[1-9]|-1[0-2]", TIME_PERIOD) ~ "M",
              TRUE ~ "A"
            ),
            OBS_STATUS = case_when(
              str_detect(TIME_PERIOD, "\\(YTD\\)") ~ "YTD",
              str_detect(TIME_PERIOD, "\\(P\\)") ~ "P",
              str_detect(TIME_PERIOD, "\\(R\\)") ~ "R",
              TRUE ~ ""
            ),
            TIME_PERIOD = str_trim(str_remove_all(TIME_PERIOD, "\\s*\\(YTD\\)|\\s*\\(P\\)|\\s*\\(R\\)"))
          ) |>
          relocate(OBS_VALUE, .before = UNIT_MEASURE) |>
          relocate(INDICATOR, .before = POVERTY_BREAKDOWN)

        # Save CSV in temp folder
        out_file <- file.path(temp_dir, paste0(sheet, ".csv"))
        write.csv(table_long, out_file, row.names = FALSE)

        created_files <- c(created_files, out_file)
        poverty_preview_data(head(table_long, 10))
      }

      poverty_output_files(created_files)
      add_poverty_log("Poverty Statistics processing completed successfully ✔")
    })

    # Preview table
    output$poverty_preview <- renderTable({
      req(poverty_preview_data())
      poverty_preview_data()
    })

    output$poverty_log <- renderText({
      poverty_log_text()
    })

    # Download ZIP
    output$download_poverty_zip <- downloadHandler(
      filename = function() {
        paste0("Poverty_output_", Sys.Date(), ".zip")
      },
      content = function(file) {

        req(poverty_output_files())

        temp_zip <- tempfile(fileext = ".zip")

        zip::zip(
          zipfile = temp_zip,
          files = poverty_output_files(),
          mode = "cherry-pick"
        )

        file.copy(temp_zip, file)
      },
      contentType = "application/zip"
    )

    # ---------------- VISITORS PROCESSING ----------------

    visitor_log_text <- reactiveVal("")
    visitor_output_files <- reactiveVal(NULL)
    visitor_preview_data <- reactiveVal(NULL)

    add_visitor_log <- function(msg) {
      visitor_log_text(paste(visitor_log_text(), msg, sep = "\n"))
    }

    observeEvent(input$process_visitor, {

      req(input$visitor_excel_file)

      visitor_log_text("")
      add_visitor_log("Starting Visitors processing...")

      file_path <- input$visitor_excel_file$datapath
      sheet_names <- excel_sheets(file_path)

      temp_dir <- tempdir()
      created_files <- c()

      i <- 1
      while (i <= length(sheet_names)) {

        add_visitor_log(paste("Processing sheet:", sheet_names[i]))

        if (sheet_names[i] == "DF_VISITORS_TABLE5") {
          table <- read_excel(file_path, sheet = sheet_names[i])

          # Define age columns
          age_cols <- c(
            "Y0","Y1T2","Y3T4","Y5T9","Y10T14","Y15T19",
            "Y20T24","Y25T29","Y30T34","Y35T39","Y40T44",
            "Y45T49","Y50T54","Y55T59","Y60T64","Y_GE65","_T"
          )

          totals_df <- table |>
            group_by(
              DATAFLOW, FREQ, REF_AREA, INDICATOR, DIRECTION,
              TYPE, PURPOSE, COUNTRY_RESIDENCE,
              COUNTRY_DESTINATION, TIME_PERIOD,
              UNIT_MEASURE, UNIT_MULT, OBS_STATUS,
              COMMENT, DECIMALS
            ) %>%
            summarise(
              across(all_of(age_cols), ~sum(as.numeric(.), na.rm = TRUE)),
              .groups = "drop"
            ) %>%
            mutate(SEX = "_T")

          # Combine original + totals
          df_combined <- bind_rows(table, totals_df) |>
            arrange(COUNTRY_RESIDENCE,
                    factor(SEX, levels = c("M", "F", "_T")))

          table_long <- df_combined |>
            pivot_longer(
              cols = all_of(age_cols),
              names_to = "AGE",
              values_to = "OBS_VALUE"
            ) %>%
            mutate(across(everything(), ~replace(., is.na(.), ""))) %>%
            relocate(OBS_VALUE, .before = UNIT_MEASURE) %>%
            relocate(AGE, .before = TIME_PERIOD)

        } else if (sheet_names[i] == "DF_VISITORS_TABLE6") {
          table <- read_excel(file_path, sheet = sheet_names[i])
          table_long <- table %>%
            pivot_longer(
              cols = -c(DATAFLOW:DECIMALS),
              names_to = "PURPOSE",
              values_to = "OBS_VALUE"
            ) %>%
            mutate(across(everything(), ~replace(., is.na(.), ""))) %>%
            relocate(OBS_VALUE, .before = UNIT_MEASURE) %>%
            relocate(PURPOSE, .before = COUNTRY_RESIDENCE)

        } else {
          table <- read_excel(file_path, sheet = sheet_names[i])
          table_long <- table %>%
            pivot_longer(
              cols = -c(DATAFLOW:DECIMALS),
              names_to = "TIME_PERIOD_ORIG",
              values_to = "OBS_VALUE"
            ) %>%
            mutate(
              FREQ = case_when(
                grepl("-Q[1-4]", TIME_PERIOD_ORIG) ~ "Q",
                grepl("-0[1-9]|-1[0-2]", TIME_PERIOD_ORIG) ~ "M",
                TRUE ~ "A"
              ),
              OBS_STATUS = ifelse(grepl("\\(P\\)", TIME_PERIOD_ORIG), "P", ""),
              TIME_PERIOD = gsub(" \\(P\\)", "", TIME_PERIOD_ORIG),
              across(everything(), ~replace(., is.na(.), ""))
            ) %>%
            select(-TIME_PERIOD_ORIG) %>%
            relocate(OBS_VALUE, .before = UNIT_MEASURE) %>%
            relocate(TIME_PERIOD, .before = OBS_VALUE) %>%
            relocate(FREQ, .before = REF_AREA) %>%
            relocate(OBS_STATUS, .before = COMMENT)
        }

        out_file <- file.path(temp_dir, paste0(sheet_names[i], ".csv"))
        write.csv(table_long, out_file, row.names = FALSE)

        created_files <- c(created_files, out_file)
        visitor_preview_data(head(table_long, 10))

        i <- i + 1
      }

      visitor_output_files(created_files)
      add_visitor_log("Visitors processing completed successfully ✔")
    })

    # Preview table
    output$visitor_preview <- renderTable({
      req(visitor_preview_data())
      visitor_preview_data()
    })

    output$visitor_log <- renderText({
      visitor_log_text()
    })

    # Download ZIP
    output$download_visitor_zip <- downloadHandler(
      filename = function() {
        paste0("Visitors_output_", Sys.Date(), ".zip")
      },
      content = function(file) {

        req(visitor_output_files())

        temp_zip <- tempfile(fileext = ".zip")

        zip::zip(
          zipfile = temp_zip,
          files = visitor_output_files(),
          mode = "cherry-pick"
        )

        file.copy(temp_zip, file)
      },
      contentType = "application/zip"
    )

  }

  # -------------------------------------------------
  # Run App
  # -------------------------------------------------
  shinyApp(ui, server)

}

# Load required packages
library(tesseract)
library(magick)
library(tidyverse)
library(lubridate)
library(ggplot2)
library(patchwork)
library(openxlsx)
library(stringr)
library(purrr)

# Initialize OCR engine
eng_engine <- tesseract("eng")

# Function to process image with OCR
process_image <- function(image_path) {
  img <- image_read(image_path) %>%
    image_convert(type = 'Grayscale') %>%
    image_contrast(sharpen = 1) %>%
    image_normalize()
  
  text <- ocr(img, engine = eng_engine)
  
  # Basic cleaning
  text <- str_replace_all(text, "\\s+", " ") %>%
    str_replace_all("(?i)rupees|rs\\.?", "₹") %>%
    str_replace_all("(?i)inr", "₹")
  
  return(text)
}

# Document classifier
classify_document <- function(text) {
  case_when(
    str_detect(text, "(?i)salary\\s*slip|payslip") ~ "SALARY_SLIP",
    str_detect(text, "(?i)form\\s*16|itr|income\\s*tax") ~ "ITR_16",
    str_detect(text, "(?i)cheque|check|drawee") ~ "CHEQUE",
    str_detect(text, "(?i)bank\\s*statement|account\\s*summary") ~ "BANK_STATEMENT",
    TRUE ~ "UNKNOWN"
  )
}

parse_salary_slip <- function(text) {
  # Initialize all fields with NA
  result <- list(
    document_type = "SALARY_SLIP",
    name = NA_character_,
    employee_id = NA_character_,
    payment_date = as.Date(NA),
    gross_earnings = NA_real_,
    total_deductions = NA_real_,
    net_pay = NA_real_
  )
  
  try({
    result$name <- str_extract(text, "(?i)name\\s*:\\s*([A-Za-z ]+)") %>%
      str_remove("(?i)name\\s*:\\s*") %>% na_if("")
    
    result$employee_id <- str_extract(text, "(?i)employee\\s*id\\s*:\\s*([A-Za-z0-9-]+)") %>%
      str_remove("(?i)employee\\s*id\\s*:\\s*") %>% na_if("")
    
    result$payment_date <- str_extract(text, "(?i)payment\\s*date\\s*:\\s*(\\d{1,2}/\\d{1,2}/\\d{4}|\\d{1,2}-[A-Za-z]{3}-\\d{4})") %>%
      str_remove("(?i)payment\\s*date\\s*:\\s*") %>%
      dmy() %>% as.Date()
    
    earnings <- str_extract_all(text, "(?i)(basic\\s*salary|meal\\s*allowance|transport\\s*allowance)[\\s:\\w]*\\d+[,\\.]\\d{2}")[[1]] %>%
      str_extract_all("\\d+[,\\.]\\d{2}") %>%
      unlist() %>%
      str_replace(",", ".") %>%
      as.numeric() %>%
      sum(na.rm = TRUE)
    
    deductions <- str_extract_all(text, "(?i)(tax|insurance|pf)[\\s:\\w]*\\d+[,\\.]\\d{2}")[[1]] %>%
      str_extract_all("\\d+[,\\.]\\d{2}") %>%
      unlist() %>%
      str_replace(",", ".") %>%
      as.numeric() %>%
      sum(na.rm = TRUE)
    
    net_pay <- str_extract(text, "(?i)net\\s*pay\\s*[₹$]?\\s*\\d+[,\\.]\\d{2}") %>%
      str_extract("\\d+[,\\.]\\d{2}") %>%
      str_replace(",", ".") %>%
      as.numeric()
    
    result$gross_earnings <- ifelse(length(earnings) > 0, earnings, NA_real_)
    result$total_deductions <- ifelse(length(deductions) > 0, deductions, NA_real_)
    result$net_pay <- ifelse(length(net_pay) > 0, net_pay, NA_real_)
  })
  
  return(as.data.frame(result))
}

# ITR-16 Parser
parse_itr_16 <- function(text) {
  pan <- str_extract(text, "[A-Z]{5}[0-9]{4}[A-Z]")
  
  assessment_year <- str_extract(text, "(?i)assessment\\s*year\\s*:\\s*\\d{4}-\\d{2}")
  
  gross_salary <- str_extract(text, "(?i)gross\\s*salary[\\s\\w]*\\d+[,\\.]\\d{2}") %>%
    str_extract("\\d+[,\\.]\\d{2}") %>%
    str_replace(",", ".") %>%
    as.numeric()
  
  total_income <- str_extract(text, "(?i)total\\s*income[\\s\\w]*\\d+[,\\.]\\d{2}") %>%
    str_extract("\\d+[,\\.]\\d{2}") %>%
    str_replace(",", ".") %>%
    as.numeric()
  
  tax_payable <- str_extract(text, "(?i)tax\\s*payable[\\s\\w]*\\d+[,\\.]\\d{2}") %>%
    str_extract("\\d+[,\\.]\\d{2}") %>%
    str_replace(",", ".") %>%
    as.numeric()
  
  list(
    document_type = "ITR_16",
    pan_number = pan,
    assessment_year = assessment_year,
    gross_salary = gross_salary,
    total_income = total_income,
    tax_payable = tax_payable
  )
}

# Cheque Parser
parse_cheque <- function(text) {
  cheque_no <- str_extract(text, "(?i)cheque\\s*no\\.?\\s*:\\s*\\d+") %>%
    str_remove("(?i)cheque\\s*no\\.?\\s*:\\s*")
  
  amount <- str_extract(text, "(?i)amount\\s*[₹$]?\\s*\\d+[,\\.]\\d{2}") %>%
    str_extract("\\d+[,\\.]\\d{2}") %>%
    str_replace(",", ".") %>%
    as.numeric()
  
  payee <- str_extract(text, "(?i)pay\\s*:\\s*([A-Za-z ]+)") %>%
    str_remove("(?i)pay\\s*:\\s*")
  
  date <- str_extract(text, "\\d{2}/\\d{2}/\\d{4}") %>%
    dmy()
  
  list(
    document_type = "CHEQUE",
    cheque_number = cheque_no,
    amount = amount,
    payee_name = payee,
    date = date
  )
}

# Bank Statement Parser
parse_bank_statement <- function(text) {
  account_no <- str_extract(text, "(?i)account\\s*no\\.?\\s*:\\s*[\\d-]+") %>%
    str_remove("(?i)account\\s*no\\.?\\s*:\\s*")
  
  transaction_lines <- str_extract_all(text, "\\d{2}/\\d{2}/\\d{4}\\s+[^\\n]+?\\s+[\\d,]+\\.[\\d]{2}\\s+[\\d,]+\\.[\\d]{2}")[[1]]
  
  if (length(transaction_lines) > 0) {
    transactions <- map_dfr(transaction_lines, function(line) {
      parts <- str_split(line, "\\s{2,}")[[1]]
      data.frame(
        date = dmy(parts[1]),
        description = parts[2],
        debit = ifelse(length(parts) > 2, as.numeric(str_remove_all(parts[3], ",")), NA),
        credit = ifelse(length(parts) > 3, as.numeric(str_remove_all(parts[4], ",")), NA),
        stringsAsFactors = FALSE
      )
    }) %>%
      mutate(
        amount = ifelse(is.na(debit), credit, -debit),
        category = case_when(
          str_detect(description, "(?i)salary|payroll") ~ "Salary",
          str_detect(description, "(?i)rent|lease") ~ "Housing",
          str_detect(description, "(?i)electric|water|utility") ~ "Utilities",
          str_detect(description, "(?i)grocery|food|restaurant") ~ "Food",
          str_detect(description, "(?i)tax") ~ "Taxes",
          str_detect(description, "(?i)insurance") ~ "Insurance",
          str_detect(description, "(?i)medical|health") ~ "Healthcare",
          str_detect(description, "(?i)transport|fuel") ~ "Transportation",
          str_detect(description, "(?i)phone|internet") ~ "Communication",
          str_detect(description, "(?i)shopping|purchase") ~ "Shopping",
          str_detect(description, "(?i)interest") ~ "Interest",
          amount > 0 ~ "Income",
          TRUE ~ "Other Expenses"
        )
      )
  } else {
    transactions <- data.frame(
      date = as.Date(character()),
      description = character(),
      debit = numeric(),
      credit = numeric(),
      amount = numeric(),
      category = character(),
      stringsAsFactors = FALSE
    )
  }
  
  list(
    document_type = "BANK_STATEMENT",
    account_number = account_no,
    transactions = transactions
  )
}

# Function to process all documents
process_financial_documents <- function(folder_path) {
  image_files <- list.files(folder_path, 
                            pattern = "\\.(jpg|jpeg|png)$", 
                            full.names = TRUE,
                            ignore.case = TRUE)
  
  all_results <- list(
    salary_slips = list(),
    itr_16_forms = list(),
    cheques = list(),
    bank_statements = list(),
    unknown = list()
  )
  
  for (file in image_files) {
    tryCatch({
      text <- process_image(file)
      doc_type <- classify_document(text)
      
      parsed_data <- switch(doc_type,
                            "SALARY_SLIP" = parse_salary_slip(text),
                            "ITR_16" = parse_itr_16(text),
                            "CHEQUE" = parse_cheque(text),
                            "BANK_STATEMENT" = parse_bank_statement(text),
                            list(document_type = "UNKNOWN", raw_text = substr(text, 1, 500))
      )
      
      if (doc_type == "SALARY_SLIP") {
        all_results$salary_slips <- c(all_results$salary_slips, list(parsed_data))
      } else if (doc_type == "ITR_16") {
        all_results$itr_16_forms <- c(all_results$itr_16_forms, list(parsed_data))
      } else if (doc_type == "CHEQUE") {
        all_results$cheques <- c(all_results$cheques, list(parsed_data))
      } else if (doc_type == "BANK_STATEMENT") {
        all_results$bank_statements <- c(all_results$bank_statements, list(parsed_data))
      } else {
        all_results$unknown <- c(all_results$unknown, list(
          list(file = basename(file), type = doc_type, sample_text = substr(text, 1, 500))
        ))
      }
    }, error = function(e) {
      warning(paste("Error processing file:", file, "-", e$message))
      all_results$unknown <- c(all_results$unknown, list(
        list(file = basename(file), error = e$message)
      ))
    })
  }
  
  # Convert lists to data frames
  all_results$salary_slips <- tryCatch(bind_rows(all_results$salary_slips), error = function(e) NULL)
  all_results$itr_16_forms <- tryCatch(bind_rows(all_results$itr_16_forms), error = function(e) NULL)
  all_results$cheques <- tryCatch(bind_rows(all_results$cheques), error = function(e) NULL)
  
  # Combine bank transactions
  all_results$all_transactions <- tryCatch({
    bind_rows(lapply(all_results$bank_statements, function(x) x$transactions))
  }, error = function(e) {
    warning("Failed to combine bank transactions")
    data.frame()
  })
  
  return(all_results)
}

# Function to analyze financial data
analyze_finances <- function(processed_data) {
  # Salary analysis
  salary_summary <- if (!is.null(processed_data$salary_slips) {
    processed_data$salary_slips %>%
      mutate(
        month = ifelse(!is.na(payment_date), 
                       floor_date(payment_date, "month"), 
                       as.Date(NA))
      ) %>%
      filter(!is.na(month)) %>%  # Only include records with valid dates
      group_by(month) %>%
      summarize(
        avg_gross = mean(gross_earnings, na.rm = TRUE),
        avg_net = mean(net_pay, na.rm = TRUE),
        avg_tax = mean(total_deductions, na.rm = TRUE),
        count = n()
      ) %>%
      filter(count > 0)  # Only include months with data
  } else NULL
  

  
  # ITR analysis
  itr_summary <- if (!is.null(processed_data$itr_16_forms)) {
    processed_data$itr_16_forms %>%
      summarize(
        avg_gross_salary = mean(gross_salary, na.rm = TRUE),
        avg_tax_paid = mean(tax_payable, na.rm = TRUE),
        tax_rate = mean(tax_payable/gross_salary, na.rm = TRUE),
        count = n()
      )
  } else NULL
  
  # Transaction analysis
  transaction_summary <- if (!is.null(processed_data$all_transactions) && nrow(processed_data$all_transactions) > 0) {
    processed_data$all_transactions %>%
      mutate(month = floor_date(date, "month")) %>%
      group_by(month, category) %>%
      summarize(
        total_amount = sum(amount, na.rm = TRUE),
        avg_amount = mean(amount, na.rm = TRUE),
        count = n(),
        .groups = "drop"
      )
  } else NULL
  
  list(
    salary_summary = salary_summary,
    itr_summary = itr_summary,
    transaction_summary = transaction_summary
  )
}

# Visualization functions
plot_salary_trends <- function(salary_summary) {
  if (is.null(salary_summary)) return(NULL)
  
  ggplot(salary_summary, aes(x = month)) +
    geom_line(aes(y = avg_gross, color = "Gross Salary"), linewidth = 1) +
    geom_line(aes(y = avg_net, color = "Net Salary"), linewidth = 1) +
    labs(title = "Monthly Salary Trends", x = "Month", y = "Amount", color = "Metric") +
    scale_color_manual(values = c("Gross Salary" = "blue", "Net Salary" = "green")) +
    theme_minimal()
}

plot_expense_categories <- function(transaction_summary) {
  if (is.null(transaction_summary)) return(NULL)
  
  category_totals <- transaction_summary %>%
    filter(total_amount < 0) %>%
    group_by(category) %>%
    summarize(total = sum(abs(total_amount)))
  
  ggplot(category_totals, aes(x = reorder(category, total), y = total, fill = category)) +
    geom_col() +
    coord_flip() +
    labs(title = "Expenses by Category", x = "Category", y = "Total Amount") +
    theme_minimal() +
    theme(legend.position = "none")
}

plot_cash_flow <- function(transaction_summary) {
  if (is.null(transaction_summary)) return(NULL)
  
  monthly_flow <- transaction_summary %>%
    group_by(month) %>%
    summarize(
      income = sum(total_amount[total_amount > 0], na.rm = TRUE),
      expenses = sum(abs(total_amount[total_amount < 0]), na.rm = TRUE),
      net = income - expenses
    )
  
  ggplot(monthly_flow, aes(x = month)) +
    geom_col(aes(y = income, fill = "Income"), alpha = 0.7) +
    geom_col(aes(y = -expenses, fill = "Expenses"), alpha = 0.7) +
    geom_line(aes(y = net, color = "Net Flow"), linewidth = 1) +
    labs(title = "Monthly Cash Flow", x = "Month", y = "Amount", fill = "Type", color = "Metric") +
    scale_fill_manual(values = c("Income" = "darkgreen", "Expenses" = "darkred")) +
    scale_color_manual(values = c("Net Flow" = "blue")) +
    theme_minimal()
}

# Main function to run full analysis
run_full_analysis <- function(folder_path, output_dir = "financial_reports") {
  dir.create(output_dir, showWarnings = FALSE)
  
  message("Processing documents...")
  processed_data <- process_financial_documents(folder_path)
  
  message("Analyzing data...")
  analysis <- analyze_finances(processed_data)
  
  message("Generating visualizations...")
  if (!is.null(analysis$salary_summary)) {
    p <- plot_salary_trends(analysis$salary_summary)
    if (!is.null(p)) ggsave(file.path(output_dir, "salary_trends.png"), p, width = 10, height = 6)
  }
  
  if (!is.null(analysis$transaction_summary)) {
    p <- plot_expense_categories(analysis$transaction_summary)
    if (!is.null(p)) ggsave(file.path(output_dir, "expense_categories.png"), p, width = 10, height = 6)
    
    p <- plot_cash_flow(analysis$transaction_summary)
    if (!is.null(p)) ggsave(file.path(output_dir, "cash_flow.png"), p, width = 10, height = 6)
  }
  
  message("Generating Excel report...")
  wb <- createWorkbook()
  
  if (!is.null(processed_data$salary_slips)) {
    addWorksheet(wb, "Salary Summary")
    writeData(wb, "Salary Summary", processed_data$salary_slips)
  }
  
  if (!is.null(processed_data$itr_16_forms)) {
    addWorksheet(wb, "ITR Summary")
    writeData(wb, "ITR Summary", processed_data$itr_16_forms)
  }
  
  if (!is.null(processed_data$cheques)) {
    addWorksheet(wb, "Cheque Summary")
    writeData(wb, "Cheque Summary", processed_data$cheques)
  }
  
  if (!is.null(processed_data$all_transactions) && nrow(processed_data$all_transactions) > 0) {
    addWorksheet(wb, "Transactions")
    writeData(wb, "Transactions", processed_data$all_transactions)
  }
  
  addWorksheet(wb, "Analysis")
  analysis_data <- list()
  
  if (!is.null(analysis$salary_summary)) {
    analysis_data[["Salary Trends"]] <- analysis$salary_summary
  }
  
  if (!is.null(analysis$itr_summary)) {
    analysis_data[["Tax Summary"]] <- analysis$itr_summary
  }
  
  if (!is.null(analysis$transaction_summary)) {
    analysis_data[["Expense Analysis"]] <- analysis$transaction_summary %>% 
      filter(total_amount < 0) %>%
      group_by(category) %>%
      summarize(Total = sum(abs(total_amount)))
  }
  
  writeData(wb, "Analysis", analysis_data, startCol = 1, startRow = 1, colNames = TRUE)
  saveWorkbook(wb, file.path(output_dir, "financial_analysis.xlsx"), overwrite = TRUE)
  
  message("Saving raw data...")
  saveRDS(processed_data, file.path(output_dir, "processed_data.rds"))
  saveRDS(analysis, file.path(output_dir, "financial_analysis.rds"))
  
  message(paste("Analysis complete! Results saved to:", output_dir))
  return(list(processed_data = processed_data, analysis = analysis))
}

# Run the analysis
results <- run_full_analysis("C:/Users/ANKIT ANAND/Downloads/FLD/Utility")
# 加载必要的R包 / Load necessary R packages
library(dplyr)
library(tidyr)
library(readxl)
library(openxlsx)

# Step 1: 读取数据 / Read data
file_path <- "aligned_data.xlsx"
data <- read_excel(file_path)

# Step 2: 从 Material 列中提取表情类型 / Extract expression types from the 'Material' column
data <- data %>%
  mutate(Expression_Type = case_when(
    grepl("dis", Material, ignore.case = TRUE) ~ "Disgust",
    grepl("enj", Material, ignore.case = TRUE) ~ "Enjoyment",
    grepl("aff", Material, ignore.case = TRUE) ~ "Affiliation",
    grepl("dom", Material, ignore.case = TRUE) ~ "Dominance",
    grepl("neu", Material, ignore.case = TRUE) ~ "Neutral",
    TRUE ~ "Other"
  ))

# Step 3: 将 Categorizing_Expressions_Score 映射为 Chosen_Expression / Map 'Categorizing_Expressions_Score' to 'Chosen_Expression'
data <- data %>%
  mutate(Chosen_Expression = case_when(
    Categorizing_Expressions_Score == 1 ~ "Enjoyment", 
    Categorizing_Expressions_Score == 2 ~ "Affiliation", 
    Categorizing_Expressions_Score == 3 ~ "Dominance", 
    Categorizing_Expressions_Score == 4 ~ "Disgust", 
    Categorizing_Expressions_Score == 5 ~ "Neutral", 
    Categorizing_Expressions_Score == 6 ~ "Other",
    TRUE ~ NA_character_
  ))

# 初始化存储结果的数据框 / Initialize a dataframe to store results
results <- data.frame(
  CASE = integer(),
  Expression_Type = character(),
  UHR = numeric(),
  Chance_UHR = numeric(),
  Performance_Above_Chance = numeric(),
  stringsAsFactors = FALSE
)

# 初始化存储不完整数据的CASE列表 / Initialize a list to store incomplete CASEs
incomplete_cases <- list()

# 初始化Excel工作簿 / Initialize an Excel workbook
wb <- createWorkbook()

# 计算 UHR 和 Chance UHR / Define a function to calculate UHR and Chance UHR
calculate_uhr_chance_uhr <- function(conf_matrix) {
  uhr_results <- numeric(length(rownames(conf_matrix)))
  chance_uhr_results <- numeric(length(rownames(conf_matrix)))
  N <- sum(conf_matrix)  # 总和 N / Total sum N
  
  for (i in seq_along(rownames(conf_matrix))) {
    emotion <- rownames(conf_matrix)[i]
    a <- conf_matrix[emotion, emotion]  # 对角线元素 / Diagonal element
    b <- sum(conf_matrix[emotion, ]) - a  # 行中的非对角线元素之和 / Sum of non-diagonal elements in the row
    d <- sum(conf_matrix[, emotion]) - a  # 列中的非对角线元素之和 / Sum of non-diagonal elements in the column
    
    # 检查 b 和 d 是否为0，避免除以零 / Check if b and d are zero to avoid division by zero
    if ((a + b) > 0 & (a + d) > 0 & N > 0) {
      UHR <- (a / (a + b)) * (a / (a + d))
      Pc <- ((a + b) / N) * ((a + d) / N)
    } else {
      UHR <- 0
      Pc <- 0
    }
    
    uhr_results[i] <- UHR
    chance_uhr_results[i] <- Pc
  }
  
  return(list(UHR = uhr_results, Chance_UHR = chance_uhr_results))
}

# Step 4: 按 CASE 计算 Unbiased_Hit_Rate 和 Chance_Unbiased_Hit_Rate / Calculate UHR and Chance UHR for each CASE
cases <- unique(data$CASE)
expected_emotions <- c("Affiliation", "Disgust", "Dominance", "Enjoyment", "Neutral")

for (case in cases) {
  # 筛选出当前 CASE 的数据 / Filter data for the current CASE
  case_data <- data %>% filter(CASE == case)
  
  # 创建选择矩阵，排除 "Other" 类别 / Create choice matrix excluding "Other" category
  choice_matrix <- case_data %>%
    filter(Expression_Type != "Other", Chosen_Expression != "Other") %>%
    count(Expression_Type, Chosen_Expression) %>%
    spread(key = Chosen_Expression, value = n, fill = 0) %>%
    complete(Expression_Type = expected_emotions, fill = list(n = 0))
  
  # 转换为普通数据框以支持行名 / Convert to a regular dataframe to support row names
  choice_matrix <- as.data.frame(choice_matrix)
  
  # 检查是否包含所有预期的情绪类别 / Check if all expected emotions are included
  missing_emotions <- setdiff(expected_emotions, colnames(choice_matrix))
  
  if (length(missing_emotions) > 0) {
    # 如果有缺失情绪类别，记录 CASE 并跳过此 CASE 的计算 / Record the CASE and skip calculation if emotions are missing
    incomplete_cases[[length(incomplete_cases) + 1]] <- list(CASE = case, Missing_Emotions = missing_emotions)
    next
  }
  
  rownames(choice_matrix) <- choice_matrix$Expression_Type
  choice_matrix <- choice_matrix[, -1]
  
  # 将数据框转换为矩阵 / Convert the dataframe to a matrix
  choice_matrix <- as.matrix(choice_matrix)
  
  # 将混淆矩阵保存到 Excel 工作簿 / Save the confusion matrix to the Excel workbook
  addWorksheet(wb, paste("CASE", case))
  writeData(wb, paste("CASE", case), choice_matrix, rowNames = TRUE)
  
  # 计算 UHR 和 Chance UHR / Calculate UHR and Chance UHR
  uhr_chance_results <- calculate_uhr_chance_uhr(conf_matrix = choice_matrix)
  
  for (i in seq_along(expected_emotions)) {
    results <- rbind(results, data.frame(
      CASE = case,
      Expression_Type = expected_emotions[i],
      UHR = uhr_chance_results$UHR[i],
      Chance_UHR = uhr_chance_results$Chance_UHR[i],
      Performance_Above_Chance = uhr_chance_results$UHR[i] - uhr_chance_results$Chance_UHR[i]
    ))
  }
}

# 保存混淆矩阵到文件 / Save confusion matrices to file
saveWorkbook(wb, "confusion_matrices_all_cases.xlsx", overwrite = TRUE)

# 打印缺失情绪类别的 CASE / Print CASEs with missing emotion categories
if (length(incomplete_cases) > 0) {
  cat("Incomplete CASEs:\n")
  for (case_info in incomplete_cases) {
    cat("CASE:", case_info$CASE, "- Missing Emotions:", paste(case_info$Missing_Emotions, collapse = ", "), "\n")
  }
}

# 保存 UHR 和 Chance UHR 结果到 Excel 文件 / Save UHR and Chance UHR results to Excel file
write.xlsx(results, "uhr_results_all_cases.xlsx", rowNames = FALSE)

# 打印最终结果 / Print final results
print(results)
